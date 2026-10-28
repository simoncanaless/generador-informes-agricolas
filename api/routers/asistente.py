"""Flujo asistido en dos pasos, pensado para el tasador y no para el programador.

    1. POST /asistente/analizar  -> sube papeles y fotos, devuelve un borrador
    2. POST /asistente/generar   -> con el borrador ya revisado, saca el .xlsx

El paso intermedio (que el humano confirme o corrija) es deliberado: un informe
que va al banco no se firma sobre datos que nadie miro. Por eso el borrador
viaja con la cita del documento y el nivel de confianza de cada bloque.

Entre ambos pasos las fotos quedan en ``data/uploads/<sesion>/``, porque el
navegador no puede reenviar 20 imagenes cada vez que se corrige una letra.
"""

from __future__ import annotations

import base64
import hashlib
import json
import re
import secrets
import tempfile
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import List, Optional, Tuple

from fastapi import Depends, APIRouter, File, Form, HTTPException, UploadFile, status
from pydantic import BaseModel, Field, ValidationError

from api.dependencias import UsuarioActual, obtener_usuario_actual
from api.dependencias import (
    Almacenamiento, Config, Expedientes, ExpedientesOpcional, IAOpcional, Motor,
    ParserOpcional, leer_acotado,
)
from api.routers.informes import RespuestaGeneracion, _nombre_seguro
from schemas.curvas import LecturaIA
from schemas.extraccion import ExtraccionLegal, TituloFotografia
from schemas.informe import Antecedentes, Fotografia, InformeFitogreen
from services.ensamblador import ensamblar_informe, fusionar_extracciones
from services.renderizado import (
    PPP_VISTA_PREVIA, RenderNoDisponible, renderizar_a_paginas,
)

# Todo lo de aqui exige sesion iniciada. La dependencia va en el router
# entero y no en cada endpoint: asi una ruta nueva nace protegida.
router = APIRouter(prefix="/asistente", tags=["asistente"],
                   dependencies=[Depends(obtener_usuario_actual)])

RE_SESION = re.compile(r"^[0-9a-f]{32}$")
MANIFIESTO = "sesion.json"


class DatosBasicos(BaseModel):
    """Lo unico que el tasador teclea a mano."""

    cliente: str = Field(min_length=2)
    rut: str
    ejecutivo: Optional[str] = None
    sucursal: Optional[str] = None
    fecha_visita: Optional[date] = None
    agronomo_informante: Optional[str] = None
    n_solicitud: Optional[str] = None

    def a_antecedentes(self) -> Antecedentes:
        return Antecedentes(
            cliente=self.cliente,
            rut=self.rut,
            ejecutivo=self.ejecutivo,
            sucursal=self.sucursal,
            fecha_visita=self.fecha_visita,
            agronomo_informante=self.agronomo_informante,
            n_solicitud=self.n_solicitud,
            fecha_emision_informe=date.today(),
        )


class FotoRevisable(BaseModel):
    """Una imagen en la revision: su titulo (editable) y donde va.

    Sin descripcion tecnica desde el 2026-10-05 (el agrónomo: la ve mejor el
    agronomo). Un borrador guardado antes trae el campo `descripcion`: se
    ignora al leerlo, no rompe nada.
    """

    archivo: str
    titulo: Optional[str] = None
    categoria: str = "otra"
    incluir: bool = True
    alta_resolucion: bool = Field(
        default=False,
        description="se subio por el apartado de satelitales, asi que esta "
                    "guardada a 2000 px. Importa porque la categoria SI se "
                    "puede cambiar en la revision: marcar como satelital una "
                    "foto que entro a 1024 px la manda al hueco grande de la "
                    "caratula, donde se va a ver blanda. Con esto la pantalla "
                    "lo puede avisar en vez de dejar que se descubra impreso",
    )


class TrazaPaso(BaseModel):
    paso: str
    modelo: Optional[str] = None
    degradado: bool = False
    duracion_s: float = 0.0
    incidencias: List[str] = Field(default_factory=list)


class RespuestaAnalisis(BaseModel):
    sesion: str
    basicos: DatosBasicos
    legal: ExtraccionLegal
    fotos: List[FotoRevisable] = Field(default_factory=list)
    conflictos: List[str] = Field(default_factory=list)
    avisos: List[str] = Field(default_factory=list)
    trazabilidad: List[TrazaPaso] = Field(default_factory=list)
    expediente_id: Optional[str] = Field(
        default=None,
        description="expediente donde quedo guardado el avance; null si no hay "
                    "base de datos y el trabajo solo vive en el navegador",
    )
    guardado: bool = False


class PeticionGeneracion(BaseModel):
    sesion: str = Field(
        default="",
        description="carpeta de los archivos subidos. Vacia cuando el informe "
                    "no nacio de una subida -se armo desde el catastro y la "
                    "visita-, que es un caso normal y no un error",
    )
    basicos: DatosBasicos
    legal: ExtraccionLegal = ExtraccionLegal()
    fotos: List[FotoRevisable] = Field(default_factory=list)
    expediente_id: Optional[str] = Field(
        default=None, description="expediente al que pertenece este informe"
    )
    galeria_fotos: bool = Field(
        default=True,
        description="anade la hoja de fotografias; los informes reales de "
                    "Fitogreen la usan, aunque la plantilla base no la traiga",
    )
    excluir: List[str] = Field(
        default_factory=list,
        description="ids de config/mapping_santander.json:unidades_excluibles "
                    "-hojas o tablas que el tasador pidio no llevar en este "
                    "informe (GET /informes/unidades-excluibles). El servidor "
                    "vuelve a resolver la cascada de dependencias, nunca "
                    "confia en lo que calculo la pantalla.",
    )


def _huella(datos: bytes) -> str:
    """SHA-256 del CONTENIDO, no del nombre.

    La misma escritura renombrada es la misma escritura, y dos fotos identicas
    se pagan una sola vez. Es lo que permite que reanalizar un informe no
    vuelva a gastar cuota.
    """
    return hashlib.sha256(datos).hexdigest()


async def _recordado(almacen, huella: str, tipo: str):
    """Lo que el modelo ya leyo de este archivo, o None.

    Falla en blando a proposito: si el almacen esta caido, se vuelve a llamar
    al modelo -se paga de mas- pero el informe sale igual. Al reves seria peor.
    """
    if almacen is None:
        return None
    try:
        lectura = await almacen.buscar_lectura(huella)
    except Exception:                                     # noqa: BLE001
        return None
    return lectura if lectura and lectura.tipo == tipo else None


async def _recordar(almacen, huella: str, tipo: str, datos: dict,
                    modelo: Optional[str], nombre: Optional[str]) -> None:
    if almacen is None:
        return
    try:
        await almacen.guardar_lectura(LecturaIA(
            huella=huella, tipo=tipo, datos=datos, modelo=modelo, nombre=nombre))
    except Exception:                                     # noqa: BLE001
        pass   # no poder recordar no puede impedir entregar el informe


def _carpeta_sesion(config, sesion: str) -> Optional[Path]:
    """Donde quedaron los archivos que se subieron, o None si no hubo subida.

    La sesion existe SOLO para localizar las fotos en disco. Un informe puede
    no tener ninguna: los que se arman desde el catastro y la visita, y los que
    se reabren con «Continuar» cuando el expediente nacio sin analizar
    documentos. Esos traen la sesion vacia, y exigirla hacia que el boton de
    generar respondiera 400 'identificador de sesion invalido' -o sea que un
    informe perfectamente valido no se podia emitir-.

    Lo que si se sigue rechazando es una sesion con forma rara: ese guardia es
    contra el salto de directorio, no contra la ausencia.
    """
    if not sesion:
        return None
    if not RE_SESION.match(sesion):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="identificador de sesión inválida",
        )
    carpeta = (config.ruta_subidas / sesion).resolve()
    if carpeta.parent != config.ruta_subidas.resolve():
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST,
                            detail="sesión inválida")
    return carpeta


@router.post(
    "/analizar",
    response_model=RespuestaAnalisis,
    summary="Sube papeles y fotos y devuelve el borrador para revisar",
)
async def analizar(
    parser: ParserOpcional,
    ia: IAOpcional,
    config: Config,
    repo: ExpedientesOpcional,
    almacen: Almacenamiento,
    usuario: UsuarioActual,
    cliente: str = Form(...),
    rut: str = Form(...),
    ejecutivo: Optional[str] = Form(None),
    sucursal: Optional[str] = Form(None),
    fecha_visita: Optional[str] = Form(None),
    agronomo_informante: Optional[str] = Form(None),
    n_solicitud: Optional[str] = Form(None),
    expediente_id: Optional[str] = Form(
        None, description="si se indica, se agrega a ese expediente en vez de crear uno"
    ),
    usar_ia: bool = Form(
        True,
        description="leer las escrituras y proponer el titulo de las fotos con el modelo. "
                    "En false se suben igual y los campos quedan en blanco para "
                    "escribirlos a mano: sirve para no gastar cuota en un "
                    "informe que no la necesita",
    ),
    documentos: List[UploadFile] = File(default_factory=list),
    fotos: List[UploadFile] = File(default_factory=list),
    satelitales: List[UploadFile] = File(
        default_factory=list,
        description="imagenes satelitales del predio. Van APARTE de las fotos "
                    "porque se guardan mucho mas grandes -2000 px contra 1024- "
                    "y porque la categoria la fija el apartado por el que se "
                    "suben, que es mejor dato que el que adivine el modelo",
    ),
) -> RespuestaAnalisis:
    basicos = DatosBasicos(
        cliente=cliente, rut=rut, ejecutivo=ejecutivo, sucursal=sucursal,
        fecha_visita=date.fromisoformat(fecha_visita) if fecha_visita else None,
        agronomo_informante=agronomo_informante, n_solicitud=n_solicitud,
    )

    # Agregar a un informe que ya existe es AGREGAR, no empezar de nuevo. Se
    # lee lo guardado antes de tocar nada: la carpeta de sus archivos, su
    # lectura legal -con lo que el tasador corrigio- y sus fotos. Sin esto,
    # subir el certificado que llego dos dias despues dejaba el informe con
    # ese papel solo: se perdian las escrituras leidas y las fotos anteriores
    # (verificacion del 2026-09-11).
    previo = {}
    if expediente_id and repo is not None:
        # Mismo guardia que generar: escribir en un informe ajeno es leerlo.
        if not await repo.puede_verlo(expediente_id, str(usuario.id)):
            raise HTTPException(
                status.HTTP_404_NOT_FOUND,
                "no existe el expediente {}".format(expediente_id))
        previo = dict((await repo.obtener(expediente_id)).datos or {})

    try:
        carpeta = _carpeta_sesion(config, previo.get("sesion") or "")
    except HTTPException:
        carpeta = None          # una sesion con forma rara no se reusa
    if carpeta is not None and carpeta.is_dir():
        sesion = carpeta.name   # lo nuevo, junto a lo de antes
    else:
        sesion = uuid.uuid4().hex
        carpeta = config.ruta_subidas / sesion
    carpeta.mkdir(parents=True, exist_ok=True)

    trazas: List[TrazaPaso] = []
    avisos: List[str] = []
    ahorradas: List[str] = []   # lo que no hubo que volver a pagar

    # Las fotos que ya tenia el informe siguen siendo suyas. Solo las que
    # siguen en la carpeta: una ficha sin su archivo no llega al Excel, y
    # generar la saltaria en silencio.
    fotos_previas: List[FotoRevisable] = []
    perdidas = []
    for guardada in previo.get("fotos") or []:
        foto_previa = FotoRevisable.model_validate(guardada)
        if (carpeta / foto_previa.archivo).is_file():
            fotos_previas.append(foto_previa)
        else:
            perdidas.append(foto_previa.archivo)
    if perdidas:
        avisos.append(
            "{} imagen(es) guardada(s) antes ya no están en el servidor y no "
            "van a llegar al informe: {}".format(
                len(perdidas), ", ".join(perdidas[:5])))

    # La lectura automatica es lo UNICO que necesita el modelo. Si no hay clave
    # -o si el tasador la apago para no gastar cuota- el informe se hace igual:
    # los papeles se suben, se guardan, y los campos legales quedan en blanco
    # para escribirlos a mano. El resto del informe no toca la IA en ningun
    # momento, asi que quedarse sin tokens no impide emitir nada.
    con_ia = bool(usar_ia and parser is not None and ia is not None)
    if documentos and not con_ia:
        avisos.append(
            "los documentos no se leyeron automáticamente ({}): escriba a mano "
            "el propietario, el predio, el rol y la inscripción".format(
                "lectura desactivada para este informe" if not usar_ia
                else "no hay IA configurada"))

    # --- documentos legales, uno a uno para no castigar la cuota de golpe
    extracciones, nombres = [], []
    for documento in documentos:
        if not documento.filename:
            continue
        contenido = await leer_acotado(
            documento, config.max_bytes_pdf, documento.filename
        )
        # Se guarda SIEMPRE, se lea o no: es el respaldo del informe.
        (carpeta / _nombre_archivo(documento.filename)).write_bytes(contenido)
        if not con_ia:
            continue

        # Si esta escritura ya se leyo alguna vez -aunque haya sido en otro
        # informe, o desde otro computador- se reusa y no se paga de nuevo.
        huella = _huella(contenido)
        recordada = await _recordado(almacen, huella, "documento")
        if recordada is not None:
            extracciones.append(ExtraccionLegal.model_validate(recordada.datos))
            nombres.append(documento.filename)
            ahorradas.append(documento.filename)
            continue

        resultado = await parser.extraer(contenido)
        extracciones.append(resultado.datos)
        nombres.append(documento.filename)
        await _recordar(almacen, huella, "documento",
                        resultado.datos.model_dump(mode="json"),
                        resultado.modelo, documento.filename)
        avisos.extend(
            "{}: {}".format(documento.filename, a)
            for a in resultado.conversion.advertencias
        )
        trazas.append(TrazaPaso(
            paso="documento: {}".format(documento.filename),
            modelo=resultado.modelo, degradado=resultado.degradado,
            duracion_s=resultado.duracion_s, incidencias=resultado.incidencias,
        ))

    # Lo ya guardado va PRIMERO: la fusion se queda con el primer valor no
    # vacio de cada campo, asi que un rol que el tasador corrigio no lo pisa
    # la lectura de un papel nuevo -y si el papel dice otra cosa, se avisa-.
    # Sin papeles nuevos no se fusiona nada: lo guardado queda tal cual.
    legal_previo = ExtraccionLegal.model_validate(previo.get("legal") or {})
    conflictos_nuevos: List[str] = []
    if extracciones:
        if legal_previo != ExtraccionLegal():
            extracciones = [legal_previo] + extracciones
            nombres = ["lo ya guardado del informe"] + nombres
        fusion = fusionar_extracciones(extracciones, nombres)
        avisos.extend(fusion.avisos)
        legal_final = fusion.datos
        conflictos_nuevos = [str(c) for c in fusion.conflictos]
    else:
        legal_final = legal_previo
    conflictos = [str(c) for c in previo.get("conflictos") or []]
    conflictos += [c for c in conflictos_nuevos if c not in conflictos]

    # --- fotografias e imagenes satelitales, en una sola llamada
    #
    # Van juntas al modelo pero se GUARDAN distinto (2026-09-09). Una foto de
    # terreno se mira entera; una satelital se mira por dentro, buscando los
    # limites del predio, asi que se conserva a 2000 px contra los 1024 de la
    # foto. Lo que viaja a la IA es la version chica en los dos casos: el
    # modelo trocea en mosaicos de 768 px y todo lo que sobra lo descarta al
    # otro lado, asi que mandarla grande seria pagar por nada.
    revisables: List[FotoRevisable] = []
    subidas = ([(f, False) for f in fotos] +
               [(f, True) for f in satelitales])
    # Se numeran a continuacion de las que ya hay en la carpeta: empezar otra
    # vez en 000 pisaria una foto anterior que se llamara igual.
    primer_numero = 1 + max(
        (int(p.name[:3]) for p in carpeta.iterdir()
         if re.match(r"^\d{3}_", p.name)),
        default=-1)
    if subidas:
        from services.imagenes import para_analisis, para_satelital
        imagenes, guardadas, categorias_forzadas = [], [], []
        for indice, (foto, es_satelital) in enumerate(subidas):
            if not foto.filename:
                continue
            datos = await leer_acotado(foto, config.max_bytes_imagen, foto.filename)

            # Se comprime UNA vez, al recibirla, y esa version es la que se
            # guarda y la que termina en el Excel. Una foto de camara son 3-5
            # MB; catorce de esas hacen un informe que no pasa por correo y un
            # analisis que tarda minutos de mas.
            perfil = para_satelital if es_satelital else para_analisis
            comprimida, aviso_imagen = perfil(datos)
            if aviso_imagen:
                avisos.append("{}: {}".format(foto.filename, aviso_imagen))
            if comprimida:
                datos = comprimida

            destino = carpeta / "{:03d}_{}".format(
                primer_numero + indice, _nombre_archivo(foto.filename))
            destino.write_bytes(datos)

            # La satelital se guarda grande pero al modelo va la copia chica:
            # sin esto, una imagen de 2000 px cuesta el triple de tokens y no
            # se describe ni un poco mejor.
            if es_satelital:
                chica, _ = para_analisis(datos)
                imagenes.append(chica or datos)
            else:
                imagenes.append(datos)
            guardadas.append(destino)
            categorias_forzadas.append("satelital" if es_satelital else None)

        # Titular fotos es lo mas caro del analisis despues de las escrituras:
        # cada imagen son tokens, y una visita trae catorce. Por eso el
        # interruptor apaga esto tambien. Sin titulo automatico las fotos NO
        # se pierden: se guardan, salen en la revision con su nombre de
        # archivo por titulo y el tasador las rotula. Al Excel llegan igual.
        # Solo se pide titulo y categoria, nunca una descripcion tecnica
        # (schemas/extraccion.py:TituloFotografia).
        # Cada foto se recuerda POR SEPARADO, y al modelo solo van las que no
        # se conocen. Es donde mas se ahorra: agregar una foto a un informe de
        # catorce cuesta una, no quince.
        titulos: List[Optional[TituloFotografia]] = [None] * len(imagenes)
        if imagenes and con_ia:
            huellas = [_huella(d) for d in imagenes]
            faltan = []
            for i, h in enumerate(huellas):
                recordada = await _recordado(almacen, h, "fotografia")
                if recordada is not None:
                    titulos[i] = TituloFotografia.model_validate(recordada.datos)
                    ahorradas.append(guardadas[i].name)
                else:
                    faltan.append(i)

            if faltan:
                mime = subidas[0][0].content_type or "image/jpeg"
                respuesta = await ia.titular_fotografias(
                    [imagenes[i] for i in faltan], mime)
                nuevas = respuesta.valor.titulos
                trazas.append(TrazaPaso(
                    paso="fotografias ({} de {})".format(len(faltan), len(imagenes)),
                    modelo=respuesta.modelo, degradado=respuesta.degradado,
                    duracion_s=round(respuesta.duracion_s, 2),
                    incidencias=respuesta.incidencias,
                ))
                if len(nuevas) != len(faltan):
                    avisos.append(
                        "se mandaron {} fotos y llegaron {} títulos: revise que "
                        "cada foto tenga el suyo".format(len(faltan), len(nuevas)))
                for orden, i in enumerate(faltan):
                    if orden >= len(nuevas):
                        break
                    titulos[i] = nuevas[orden]
                    await _recordar(almacen, huellas[i], "fotografia",
                                    nuevas[orden].model_dump(mode="json"),
                                    respuesta.modelo, guardadas[i].name)
        elif imagenes:
            avisos.append(
                "las {} imágenes se subieron sin título automático: escriba "
                "el de cada una y elija su categoría".format(len(imagenes)))

        for indice, ruta in enumerate(guardadas):
            desc = titulos[indice] if indice < len(titulos) else None
            # Lo que subio por el apartado de satelitales ES satelital, lo diga
            # o no el modelo: el tasador ya contesto esa pregunta al elegir el
            # apartado, y eso vale mas que una categoria adivinada.
            categoria = categorias_forzadas[indice] or (
                (desc.categoria_sugerida or "otra") if desc else "otra")
            revisables.append(FotoRevisable(
                archivo=ruta.name,
                # Sin titulo automatico, el nombre del archivo SIN el prefijo
                # «000_» con que se guarda para conservar el orden: ese
                # prefijo es interno y salia como titulo de la foto.
                titulo=desc.titulo if desc else re.sub(r"^\d{3}_", "", ruta.stem),
                categoria=categoria,
                alta_resolucion=categorias_forzadas[indice] is not None,
            ))

    (carpeta / MANIFIESTO).write_text(
        json.dumps({"creada": datetime.now().isoformat(), "cliente": cliente},
                   ensure_ascii=False),
        encoding="utf-8",
    )

    # Se dice cuanto NO se pago: es lo que hace visible que la memoria sirve.
    if ahorradas:
        avisos.append(
            "{} archivo(s) ya se habían leído antes y no se volvieron a "
            "consultar al modelo: {}".format(
                len(ahorradas), ", ".join(ahorradas[:4]) +
                ("..." if len(ahorradas) > 4 else "")))

    respuesta = RespuestaAnalisis(
        sesion=sesion,
        basicos=basicos,
        legal=legal_final,
        # Todas las fotos del informe y no solo las de esta subida: la
        # pantalla reemplaza su borrador con esta respuesta.
        fotos=fotos_previas + revisables,
        conflictos=conflictos,
        avisos=avisos,
        trazabilidad=trazas,
    )

    # Guardar es un extra: si la base falla, el tasador conserva su analisis en
    # pantalla y solo se le avisa de que no quedo respaldado.
    if repo is not None:
        try:
            respuesta.expediente_id = await _persistir(
                repo, expediente_id, basicos, respuesta, documentos, revisables,
                carpeta, trazas, propietario_id=usuario.id,
            )
            respuesta.guardado = True
        except Exception as error:  # noqa: BLE001
            respuesta.avisos.append(
                "no se pudo guardar el informe ({}): el análisis sigue en "
                "pantalla pero se perderá al recargar".format(type(error).__name__)
            )
    else:
        respuesta.avisos.append(
            "sin base de datos: este avance no queda guardado"
        )
    return respuesta


async def _persistir(
    repo, expediente_id, basicos, respuesta, documentos, revisables, carpeta,
    trazas, propietario_id=None
) -> str:
    """Crea o actualiza el expediente con todo lo analizado."""
    propiedad = respuesta.legal.propiedad

    campos = {
        "cliente": basicos.cliente,
        "rut": basicos.rut,
        "ejecutivo": basicos.ejecutivo,
        "sucursal": basicos.sucursal,
        "n_solicitud": basicos.n_solicitud,
        "agronomo_informante": basicos.agronomo_informante,
        "comuna": propiedad.comuna,
        "region": propiedad.region,
    }
    if basicos.fecha_visita:
        campos["fecha_visita"] = datetime.combine(
            basicos.fecha_visita, datetime.min.time()
        )

    if expediente_id:
        expediente = await repo.actualizar(expediente_id, **campos)
    else:
        # El informe queda a nombre de quien lo analiza. Sin esto naceria sin
        # dueno y lo veria todo el mundo.
        expediente = await repo.crear(propietario_id=propietario_id, **campos)
    id_expediente = str(expediente.id)

    # El estado refleja donde esta el trabajo: hay datos extraidos por revisar.
    # Solo se nombran las claves de la revision: el repositorio fusiona, asi
    # que la visita capturada -que vive en la misma columna- no se toca.
    await repo.actualizar(
        id_expediente,
        estado="en_revision",
        datos={
            "sesion": carpeta.name,
            "legal": respuesta.legal.model_dump(mode="json"),
            "conflictos": respuesta.conflictos,
            "fotos": [f.model_dump(mode="json") for f in respuesta.fotos],
        },
        avisos=respuesta.avisos,
    )

    for indice, documento in enumerate(documentos):
        if not documento.filename:
            continue
        traza = trazas[indice] if indice < len(trazas) else None
        await repo.agregar_documento(
            id_expediente,
            nombre_archivo=documento.filename,
            modelo_ia=traza.modelo if traza else None,
            degradado=bool(traza.degradado) if traza else False,
        )

    # Filas solo para las fotos NUEVAS: las anteriores ya tienen la suya, y
    # el orden sigue despues del de ellas.
    ya_habia = len(respuesta.fotos) - len(revisables)
    for orden, foto in enumerate(revisables):
        await repo.agregar_fotografia(
            id_expediente,
            nombre_archivo=foto.archivo,
            ruta=str(carpeta / foto.archivo),
            titulo=foto.titulo,
            categoria=foto.categoria,
            orden=ya_habia + orden,
        )

    return id_expediente


@router.get(
    "/expediente/{id_expediente}",
    response_model=RespuestaAnalisis,
    summary="Recupera un expediente guardado para seguir donde se dejo",
)
async def recuperar(id_expediente: str, repo: Expedientes,
                    usuario: UsuarioActual) -> RespuestaAnalisis:
    if not await repo.puede_verlo(id_expediente, str(usuario.id)):
        raise HTTPException(status.HTTP_404_NOT_FOUND,
                            "no existe el expediente {}".format(id_expediente))
    """Reconstruye el borrador tal como estaba, para retomar el trabajo."""
    from services.almacen import NoEncontrado

    try:
        expediente = await repo.obtener(id_expediente)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error

    datos = expediente.datos or {}
    basicos = DatosBasicos(
        cliente=expediente.cliente,
        rut=expediente.rut,
        ejecutivo=expediente.ejecutivo,
        sucursal=expediente.sucursal,
        fecha_visita=expediente.fecha_visita.date() if expediente.fecha_visita else None,
        agronomo_informante=expediente.agronomo_informante,
        n_solicitud=expediente.n_solicitud,
    )

    fotos = [FotoRevisable.model_validate(f) for f in datos.get("fotos", [])]
    if not fotos:
        # Respaldo: si el bloque JSON se perdio, se rearma desde las filas.
        fotos = [
            FotoRevisable(
                archivo=f.nombre_archivo, titulo=f.titulo,
                categoria=f.categoria, incluir=f.incluir,
            )
            for f in expediente.fotografias
        ]

    return RespuestaAnalisis(
        sesion=datos.get("sesion", ""),
        basicos=basicos,
        legal=ExtraccionLegal.model_validate(datos.get("legal", {})),
        fotos=fotos,
        conflictos=datos.get("conflictos", []),
        avisos=list(expediente.avisos or []),
        expediente_id=str(expediente.id),
        guardado=True,
    )


async def _verificar_permiso(peticion: PeticionGeneracion, repo, usuario) -> None:
    """Generar un informe ajeno seria leerlo entero: mismo guardia que verlo.

    La comparten /generar y /vista-previa: una vista previa de un expediente
    ajeno seria la misma fuga que descargarlo.
    """
    if peticion.expediente_id and repo is not None:
        if not await repo.puede_verlo(peticion.expediente_id, str(usuario.id)):
            raise HTTPException(
                status.HTTP_404_NOT_FOUND,
                "no existe el expediente {}".format(peticion.expediente_id))


async def _construir_informe(
    peticion: PeticionGeneracion, config, repo, almacen,
) -> Tuple[InformeFitogreen, List[str], List[str]]:
    """Arma el InformeFitogreen del borrador -fotos, y catastro/visita/curvas/
    riego guardados si hay expediente_id.

    La comparten /generar y /vista-previa A PROPOSITO: la vista previa tiene
    que mostrar EXACTAMENTE lo mismo que despues se descarga, nunca una
    version mas pobre del informe.

    Devuelve (informe, avisos_de_armado, fotos_sin_carpeta).
    """
    carpeta = _carpeta_sesion(config, peticion.sesion)

    rutas, categorias, titulos = [], [], []
    sin_carpeta = []
    for foto in peticion.fotos:
        if not foto.incluir:
            continue
        if carpeta is None:
            # No se pierde en silencio: si el borrador declara fotos y no hay
            # de donde sacarlas, el tasador tiene que enterarse.
            sin_carpeta.append(foto.archivo)
            continue
        ruta = (carpeta / foto.archivo).resolve()
        if ruta.parent != carpeta or not ruta.is_file():
            continue  # la sesion caduco o el nombre no cuadra: se omite
        rutas.append(str(ruta))
        categorias.append(foto.categoria)
        titulos.append(foto.titulo)

    fotografias: List[Fotografia] = [
        Fotografia(
            ruta=ruta,
            titulo=titulo,
            categoria=_categoria_valida(categoria),
        )
        for ruta, categoria, titulo in zip(rutas, categorias, titulos)
    ]

    try:
        informe = ensamblar_informe(
            antecedentes=peticion.basicos.a_antecedentes(),
            legal=peticion.legal,
            fotografias=fotografias,
        )
    except ValidationError as error:
        # Un dato mal tipeado (ej. el RUT sin el guion del digito
        # verificador) no puede tumbar el servidor con un 500 -eso es lo
        # que le paso a un usuario real: el mensaje del validador YA es
        # claro ("el RUT debe incluir..."), solo hacia falta no perderlo
        # en una excepcion sin capturar.
        raise HTTPException(
            status.HTTP_422_UNPROCESSABLE_ENTITY,
            "; ".join(e["msg"].removeprefix("Value error, ") for e in error.errors()),
        ) from error

    # Todo lo demas que el sistema ya sabe de este expediente: el catastro
    # importado, la visita capturada, las curvas del catalogo y la biblioteca
    # de riego. Sin esto el boton de generar solo usaba el formulario, las
    # escrituras y las fotos, y el resto habia que pasarlo por un script.
    avisos_armado: List[str] = []
    if repo is not None and peticion.expediente_id:
        armado = await _completar_con_lo_guardado(
            informe, peticion.expediente_id, repo, almacen, config)
        if armado:
            avisos_armado.extend(armado.get("avisos") or [])

    return informe, avisos_armado, sin_carpeta


@router.post(
    "/generar",
    response_model=RespuestaGeneracion,
    summary="Genera el .xlsx con el borrador ya revisado por el tasador",
)
async def generar(
    peticion: PeticionGeneracion,
    motor: Motor,
    config: Config,
    repo: ExpedientesOpcional,
    almacen: Almacenamiento,
    usuario: UsuarioActual,
) -> RespuestaGeneracion:
    await _verificar_permiso(peticion, repo, usuario)
    informe, avisos_armado, sin_carpeta = await _construir_informe(
        peticion, config, repo, almacen)

    marca = datetime.now()
    # El sufijo al azar hace que el enlace de descarga no se pueda adivinar
    # con el nombre del cliente y la hora: es la unica llave del archivo.
    nombre = "{}_{}_{}.xlsx".format(
        _nombre_seguro(peticion.basicos.cliente), marca.strftime("%Y%m%d_%H%M%S"),
        secrets.token_hex(8),
    )
    resultado = await motor.generar_async(
        informe, config.ruta_salida / nombre, peticion.galeria_fotos,
        peticion.excluir,
    )
    # Lo que quedo fuera del catastro/visita/riego se le dice al tasador, no
    # se esconde.
    resultado.avisos.extend(avisos_armado)
    if sin_carpeta:
        resultado.avisos.append(
            "{} fotografía(s) no se pudieron adjuntar porque el informe no "
            "tiene archivos subidos: {}".format(
                len(sin_carpeta), ", ".join(sin_carpeta[:5])))

    # Traza del Excel emitido. Un informe bancario tiene que poder auditarse:
    # que se genero, cuando, con que plantilla y con que observaciones.
    if repo is not None and peticion.expediente_id:
        try:
            await repo.registrar_generado(
                peticion.expediente_id,
                archivo=nombre,
                plantilla=config.plantilla_predeterminada,
                celdas_escritas=resultado.celdas_escritas,
                filas_insertadas=resultado.filas_insertadas,
                imagenes_insertadas=resultado.imagenes_insertadas,
                avisos=resultado.avisos,
            )
            await repo.actualizar(
                peticion.expediente_id, estado="completo", fecha_emision=marca
            )
        except Exception as error:  # noqa: BLE001
            resultado.avisos.append(
                "el Excel se generó pero no se pudo registrar en el informe "
                "({})".format(type(error).__name__)
            )

    return RespuestaGeneracion(
        archivo=nombre,
        descarga="/informes/{}/descargar".format(nombre),
        celdas_escritas=resultado.celdas_escritas,
        filas_insertadas=resultado.filas_insertadas,
        imagenes_insertadas=resultado.imagenes_insertadas,
        avisos=resultado.avisos,
        generado=marca,
    )


class RespuestaVistaPrevia(BaseModel):
    paginas: List[str] = Field(
        default_factory=list,
        description="cada pagina del informe como imagen en base64 (data URL, "
                    "WebP sin perdida o PNG), en el mismo orden en que las "
                    "veria el tasador",
    )
    tamanos: List[List[int]] = Field(
        default_factory=list,
        description="[ancho, alto] en pixeles de cada pagina, en el mismo "
                    "orden: la pantalla reserva el espacio antes de que la "
                    "imagen termine de cargar, y asi el visor no salta al "
                    "acercar o alejar",
    )
    ppp: int = Field(default=0, description="resolucion del render, en puntos por pulgada")
    motor: str = Field(description="'Excel' o 'LibreOffice', con lo que se renderizo")
    avisos: List[str] = Field(default_factory=list)


@router.post(
    "/vista-previa",
    response_model=RespuestaVistaPrevia,
    summary="Renderiza el .xlsx tal como quedaria con la seleccion actual, sin descargarlo",
)
async def vista_previa(
    peticion: PeticionGeneracion,
    motor: Motor,
    config: Config,
    repo: ExpedientesOpcional,
    almacen: Almacenamiento,
    usuario: UsuarioActual,
) -> RespuestaVistaPrevia:
    """El mismo informe que armaria /generar, pero renderizado con el Excel
    real (services/renderizado.py) y devuelto como imagenes -nada se guarda
    en ``output/`` ni se registra en el expediente, porque no es un informe
    entregado: es una prueba visual que el tasador puede pedir muchas veces
    por sesion mientras prueba que incluir.
    """
    await _verificar_permiso(peticion, repo, usuario)
    informe, avisos_armado, sin_carpeta = await _construir_informe(
        peticion, config, repo, almacen)

    with tempfile.TemporaryDirectory(prefix="fito_vista_previa_") as carpeta_tmp:
        destino = Path(carpeta_tmp) / "vista_previa.xlsx"
        resultado = await motor.generar_async(
            informe, destino, peticion.galeria_fotos, peticion.excluir,
        )
        try:
            paginas, motor_usado = await renderizar_a_paginas(destino)
        except RenderNoDisponible as error:
            raise HTTPException(
                status.HTTP_503_SERVICE_UNAVAILABLE,
                "no se pudo generar la vista previa: instale Excel o "
                "LibreOffice en este equipo ({})".format(error),
            ) from error
        # destino y cualquier PDF intermedio se borran solos al salir de
        # este `with` -una vista previa no debe dejar rastro en el disco.

    avisos = list(resultado.avisos) + avisos_armado
    if sin_carpeta:
        avisos.append(
            "{} fotografía(s) no se pudieron adjuntar porque el informe no "
            "tiene archivos subidos: {}".format(
                len(sin_carpeta), ", ".join(sin_carpeta[:5])))
    if motor_usado == "LibreOffice":
        avisos.append(
            "vista previa generada con LibreOffice: la apariencia puede "
            "diferir levemente de como se ve en Excel")

    return RespuestaVistaPrevia(
        paginas=[
            "data:{};base64,{}".format(
                pagina.tipo, base64.b64encode(pagina.datos).decode("ascii"))
            for pagina in paginas
        ],
        tamanos=[[pagina.ancho, pagina.alto] for pagina in paginas],
        ppp=PPP_VISTA_PREVIA,
        motor=motor_usado,
        avisos=avisos,
    )


async def _completar_con_lo_guardado(informe, id_expediente, repo, almacen,
                                     config):
    """Junta catastro, visita, curvas y riego sobre el informe ya ensamblado.

    Falla en silencio a proposito: si la base no responde o falta una
    biblioteca, el informe se genera igual con lo que haya -las escrituras y
    las fotos ya estan dentro- y el problema se cuenta como aviso. Que un
    catalogo de curvas caido impida entregar el informe legal seria peor que
    el hueco que deja.
    """
    from services.armado import completar_informe

    try:
        plantaciones = await repo.plantaciones(id_expediente)
    except Exception as error:                       # noqa: BLE001
        informe_avisos = "no se pudo leer el catastro guardado ({})".format(
            type(error).__name__)
        return {"avisos": [informe_avisos]}

    problemas = []
    try:
        curvas = await almacen.listar_curvas()
    except Exception as error:                       # noqa: BLE001
        curvas = []
        problemas.append(
            "no se pudo leer el catálogo de curvas ({})".format(
                type(error).__name__))

    visita, formulario, comuna = None, None, None
    try:
        fila = await repo.obtener(id_expediente)
        comuna = fila.comuna
        visita = (fila.datos or {}).get("visita")
        if visita:
            from services.visita import cargar_formulario
            formulario = cargar_formulario(config.ruta_formulario_visita)
    except Exception as error:                       # noqa: BLE001
        problemas.append(
            "no se pudo leer la visita guardada ({})".format(
                type(error).__name__))

    riego = None
    try:
        from services.riego import cargar_parametros
        riego = cargar_parametros(config.ruta_parametros_riego)
    except Exception as error:                       # noqa: BLE001
        problemas.append(
            "no se pudo leer la biblioteca de riego ({})".format(
                type(error).__name__))

    resultado = completar_informe(
        informe, plantaciones=plantaciones, curvas=curvas,
        visita=visita, formulario=formulario, parametros_riego=riego,
        comuna=comuna,
        formato=config.motor_de_plantilla(config.plantilla_predeterminada),
    )
    for problema in problemas:
        resultado.avisar(problema)
    return resultado.como_dict()


def _nombre_archivo(nombre: str) -> str:
    """Nombre de archivo sin sorpresas: el del navegador no es de fiar."""
    limpio = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(nombre).name)
    return limpio[:80] or "imagen.jpg"


def _categoria_valida(valor: str) -> str:
    from schemas.informe import CategoriaFoto

    try:
        return CategoriaFoto(valor.lower().strip())
    except ValueError:
        return CategoriaFoto.OTRA
