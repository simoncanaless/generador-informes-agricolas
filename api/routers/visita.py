"""Captura de la visita a terreno.

Sirve la definicion del formulario y guarda lo que el tasador va llenando.
El formulario vive en 'config/formulario_visita.json' y no en el navegador:
agregar una pregunta es agregar una entrada al JSON, sin tocar codigo.
"""

from __future__ import annotations

from datetime import datetime
from functools import lru_cache
from typing import Any, Dict, List, Optional

from io import BytesIO

from fastapi import Depends, APIRouter, File, HTTPException, UploadFile, status
from pydantic import BaseModel, Field

from api.dependencias import UsuarioActual, obtener_usuario_actual
from api.dependencias import Config, Expedientes, leer_acotado
from config.settings import obtener_ajustes
from services.almacen import NoEncontrado
from services.importar_tabla import (
    TablaIlegible, claves_sin_revisar, importar_seccion,
)
from services.visita import cargar_formulario, resumen_de_avance

# Todo lo de aqui exige sesion iniciada. La dependencia va en el router
# entero y no en cada endpoint: asi una ruta nueva nace protegida.
router = APIRouter(prefix="/visita", tags=["visita"],
                   dependencies=[Depends(obtener_usuario_actual)])


@lru_cache(maxsize=1)
def _formulario():
    return cargar_formulario(obtener_ajustes().ruta_formulario_visita)


class CapturaVisita(BaseModel):
    """Lo que el tasador lleva escrito.

    Dos formas de mandarlo, y la que importa es la segunda:

    - ``secciones``: todo el formulario. Reemplaza lo guardado. Es como
      funcionaba antes y sirve para los scripts, pero si dos personas tienen
      abierto el mismo informe, el que guarda segundo borra lo del primero.
    - ``cambios``: SOLO los campos que esa persona toco, con la clave de la
      pantalla (``seccion|fila|campo``). Se fusiona sobre lo guardado, asi que
      lo que llenó el otro queda intacto. Es lo que manda el navegador.
    """

    secciones: Dict[str, Any] = Field(default_factory=dict)
    cambios: Optional[Dict[str, Any]] = Field(
        default=None,
        description="campos tocados, por su clave 'seccion|fila|campo'. Si "
                    "viene, se fusiona y se ignora 'secciones'",
    )


def _fusionar(guardado: Dict[str, Any], cambios: Dict[str, Any]) -> Dict[str, Any]:
    """Aplica campo a campo sobre lo que ya estaba.

    La clave llega tal como la pantalla nombra cada control: 'tipo_cambio||uf'
    para un campo suelto y 'maquinaria|0|nombre' para una fila de tabla. El
    tramo del medio vacio significa "no es una tabla".

    Nunca borra una seccion entera: lo que no venga en 'cambios' se queda como
    estaba, que es justo lo que evita pisar el trabajo del otro.
    """
    salida = dict(guardado or {})
    for clave, valor in (cambios or {}).items():
        partes = str(clave).split("|")
        if len(partes) != 3:
            continue                      # clave que no reconocemos: se ignora
        seccion, fila, campo = partes
        if not seccion or not campo:
            continue
        if fila == "":
            actual = dict(salida.get(seccion) or {})
            actual[campo] = valor
            salida[seccion] = actual
        else:
            try:
                indice = int(fila)
            except ValueError:
                continue
            tabla = list(salida.get(seccion) or [])
            while len(tabla) <= indice:
                tabla.append({})
            renglon = dict(tabla[indice] or {})
            renglon[campo] = valor
            tabla[indice] = renglon
            salida[seccion] = tabla
    return salida


@router.get("/formulario", summary="Que preguntas tiene la pantalla de visita")
async def formulario() -> Dict[str, Any]:
    return _formulario().para_navegador()


@router.get("/{id_expediente}", summary="Lo capturado hasta ahora")
async def obtener(id_expediente: str, repo: Expedientes,
                  usuario: UsuarioActual) -> Dict[str, Any]:
    if not await repo.puede_verlo(id_expediente, str(usuario.id)):
        raise HTTPException(status.HTTP_404_NOT_FOUND,
                            "no existe el expediente {}".format(id_expediente))
    try:
        fila = await repo.obtener(id_expediente)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error

    capturado = (fila.datos or {}).get("visita", {})
    return {
        "secciones": capturado,
        "avance": resumen_de_avance(capturado, _formulario()),
        # La version que se lleva puesta: con ella el navegador puede
        # preguntar despues "cambio algo?" sin traerse todo de vuelta.
        "version": (fila.datos or {}).get("visita_version"),
        # Lo que entro desde una planilla y todavia no miro nadie. La pantalla
        # lo pinta distinto y el paso de revision lo cuenta.
        "sin_revisar": (fila.datos or {}).get("visita_sin_revisar") or [],
    }


@router.put("/{id_expediente}", summary="Guarda lo que se lleva llenado")
async def guardar(
    id_expediente: str, captura: CapturaVisita, repo: Expedientes,
    usuario: UsuarioActual,
) -> Dict[str, Any]:
    if not await repo.puede_verlo(id_expediente, str(usuario.id)):
        raise HTTPException(status.HTTP_404_NOT_FOUND,
                            "no existe el expediente {}".format(id_expediente))
    """Guarda a medias sin quejarse.

    El tasador vuelve varias veces a esta pantalla: consigue un dato del
    cliente, lo anota, y sigue. Exigir el formulario completo para poder
    guardar seria devolverlo al Excel.
    """
    try:
        fila = await repo.obtener(id_expediente)
        datos = dict(fila.datos or {})
        if captura.cambios is not None:
            datos["visita"] = _fusionar(datos.get("visita") or {}, captura.cambios)
            # Un campo que una persona toco deja de estar «sin revisar»: eso
            # es exactamente lo que significa revisarlo. Se reusa la lista de
            # claves que ya viaja para fusionar, asi que no hace falta que la
            # pantalla mande nada mas.
            sin_revisar = set(datos.get("visita_sin_revisar") or [])
            if sin_revisar:
                datos["visita_sin_revisar"] = sorted(
                    sin_revisar - set(captura.cambios))
        else:
            datos["visita"] = captura.secciones
        # Marca de version: con esto el navegador del otro sabe que hay algo
        # nuevo sin tener que comparar el formulario entero.
        #
        # Con precision de SEGUNDOS no servia: dos personas guardando dentro
        # del mismo segundo producian la misma marca y el segundo cambio
        # pasaba desapercibido. Con microsegundos no hay empate posible.
        datos["visita_version"] = datetime.now().isoformat()
        datos["visita_por"] = usuario.nombre
        actualizada = await repo.actualizar(id_expediente, datos=datos)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error

    capturado = (actualizada.datos or {}).get("visita", {})
    return {
        "guardado": True,
        "avance": resumen_de_avance(capturado, _formulario()),
        "version": (actualizada.datos or {}).get("visita_version"),
        "sin_revisar": (actualizada.datos or {}).get("visita_sin_revisar") or [],
    }


# ------------------------------------------------- trabajar de a dos

class Senal(BaseModel):
    """Lo que el navegador cuenta cada pocos segundos."""

    seccion: Optional[str] = Field(
        default=None, description="en que seccion esta parado, para poder decirlo")
    version: Optional[str] = Field(
        default=None,
        description="la version de la visita que tiene en pantalla. Si el "
                    "servidor tiene otra, devuelve los datos; si es la misma, "
                    "no devuelve nada y la respuesta pesa cuatro lineas")
    me_voy: bool = False


@router.post("/{id_expediente}/presencia",
             summary="Avisa que sigo aqui y me entero de los demas")
async def presencia(
    id_expediente: str, senal: Senal, repo: Expedientes, usuario: UsuarioActual
) -> Dict[str, Any]:
    """Una sola llamada resuelve las dos cosas: quien mas esta, y si hay algo
    nuevo que mostrar.

    Se hace en un endpoint y no en dos porque el navegador pregunta cada cinco
    segundos: dos peticiones serian el doble de ruido para el mismo dato.
    """
    from services import presencia as sala

    if not await repo.puede_verlo(id_expediente, str(usuario.id)):
        raise HTTPException(status.HTTP_404_NOT_FOUND,
                            "no existe el expediente {}".format(id_expediente))

    if senal.me_voy:
        sala.olvidar(id_expediente, str(usuario.id))
        return {"listo": True}

    sala.marcar(id_expediente, str(usuario.id), usuario.nombre, senal.seccion)

    fila = await repo.obtener(id_expediente)
    datos = fila.datos or {}
    version = datos.get("visita_version")

    respuesta: Dict[str, Any] = {
        "otros": sala.quienes_estan(id_expediente, excepto=str(usuario.id)),
        "version": version,
        "por": datos.get("visita_por"),
    }
    # Los datos SOLO viajan cuando cambiaron. Con tres personas preguntando
    # cada cinco segundos, mandar el formulario entero cada vez seria mandar
    # 284 campos por nada.
    if version and version != senal.version:
        respuesta["secciones"] = datos.get("visita", {})
        respuesta["avance"] = resumen_de_avance(
            respuesta["secciones"], _formulario())
    return respuesta



# ------------------------------------------- llenar desde la planilla
#
# El cliente ya tiene escrito lo que el tasador tipea: la maquinaria, los
# resultados de la temporada, las inversiones. Esto lo copia por el, y lo deja
# marcado SIN REVISAR hasta que una persona lo mire: el dato del cliente no es
# un dato verificado, y el que firma el informe responde por el.

@router.post("/{id_expediente}/importar/{seccion_id}",
             summary="Llena una seccion desde la planilla del cliente")
async def importar(
    id_expediente: str, seccion_id: str, repo: Expedientes,
    usuario: UsuarioActual, config: Config,
    archivo: UploadFile = File(..., description="la planilla del cliente"),
) -> Dict[str, Any]:
    if not await repo.puede_verlo(id_expediente, str(usuario.id)):
        raise HTTPException(status.HTTP_404_NOT_FOUND,
                            "no existe el expediente {}".format(id_expediente))

    seccion = next((s for s in _formulario().para_navegador()["secciones"]
                    if s.get("id") == seccion_id), None)
    if seccion is None:
        raise HTTPException(status.HTTP_404_NOT_FOUND,
                            "no hay una seccion '{}'".format(seccion_id))

    # Con tope, como las demas subidas (ver dependencias.leer_acotado).
    contenido = await leer_acotado(archivo, config.max_bytes_pdf, "la planilla")
    try:
        leido = importar_seccion(BytesIO(contenido), seccion)
    except TablaIlegible as error:
        raise HTTPException(status.HTTP_422_UNPROCESSABLE_ENTITY,
                            str(error)) from error

    try:
        fila = await repo.obtener(id_expediente)
        datos = dict(fila.datos or {})
        visita = dict(datos.get("visita") or {})
        visita[seccion_id] = leido.filas
        datos["visita"] = visita

        # Se marca lo importado y se DESMARCA lo que esta seccion tuviera de
        # antes: la tabla se reemplaza entera, asi que una clave vieja apuntaria
        # a un dato que ya no esta.
        sin_revisar = {c for c in (datos.get("visita_sin_revisar") or [])
                       if not str(c).startswith(seccion_id + "|")}
        sin_revisar |= set(claves_sin_revisar(seccion_id, leido.filas))
        datos["visita_sin_revisar"] = sorted(sin_revisar)

        datos["visita_version"] = datetime.now().isoformat()
        datos["visita_por"] = usuario.nombre
        actualizada = await repo.actualizar(id_expediente, datos=datos)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error

    return {
        "seccion": seccion_id,
        "titulo": seccion.get("titulo", seccion_id),
        "hoja": leido.hoja,
        "fila_encabezado": leido.fila_encabezado,
        "filas": leido.filas,
        "cuantas": leido.cuantas,
        "reconocidas": [
            {"campo": c.campo, "esperado": c.etiqueta_esperada,
             "en_el_archivo": c.encabezado_en_el_archivo, "columna": c.columna}
            for c in leido.reconocidas],
        "avisos": leido.avisos,
        "sin_revisar": (actualizada.datos or {}).get("visita_sin_revisar") or [],
        "version": (actualizada.datos or {}).get("visita_version"),
    }
