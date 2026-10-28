"""Expedientes: la tasacion en curso, guardada.

Es lo que hoy se pierde al cerrar el navegador. Un informe de tasacion no se
hace de una sentada: se visita el predio un martes, llega el certificado de
aguas el jueves y se firma la semana siguiente. Sin esto, cada interrupcion
obliga a rehacer la captura.

Requiere PostgreSQL. Sin ``DATABASE_URL`` estos endpoints responden 503 y el
resto del sistema sigue funcionando igual.
"""

from __future__ import annotations

import uuid
from datetime import datetime
from typing import Any, Dict, List, Optional

from fastapi import Depends, APIRouter, HTTPException, Query, status
from pydantic import BaseModel, Field

from api.dependencias import UsuarioActual, obtener_usuario_actual
from api.dependencias import Almacenamiento, Expedientes
from schemas.informe import Plantacion
from services.almacen import NoEncontrado
from services.usuarios import formatear_rut

# Todo lo de aqui exige sesion iniciada. La dependencia va en el router
# entero y no en cada endpoint: asi una ruta nueva nace protegida.
router = APIRouter(prefix="/expedientes", tags=["expedientes"],
                   dependencies=[Depends(obtener_usuario_actual)])

ESTADOS = ("borrador", "en_revision", "completo", "entregado", "archivado")


class NuevoExpediente(BaseModel):
    cliente: str = Field(min_length=2, max_length=300)
    rut: str = Field(min_length=3, max_length=20)
    ejecutivo: Optional[str] = None
    sucursal: Optional[str] = None
    n_solicitud: Optional[str] = None
    agronomo_informante: Optional[str] = None
    banco: Optional[str] = None
    fecha_visita: Optional[datetime] = None
    comuna: Optional[str] = None
    region: Optional[str] = None
    notas: Optional[str] = None


class CambioExpediente(BaseModel):
    """Todo opcional: se actualiza solo lo que llegue."""

    cliente: Optional[str] = None
    rut: Optional[str] = None
    ejecutivo: Optional[str] = None
    sucursal: Optional[str] = None
    n_solicitud: Optional[str] = None
    agronomo_informante: Optional[str] = None
    banco: Optional[str] = None
    fecha_visita: Optional[datetime] = None
    fecha_emision: Optional[datetime] = None
    comuna: Optional[str] = None
    region: Optional[str] = None
    estado: Optional[str] = None
    datos: Optional[Dict[str, Any]] = None
    avisos: Optional[List[Any]] = None
    notas: Optional[str] = None


class ResumenExpediente(BaseModel):
    """Lo justo para la pantalla de 'mis expedientes'."""

    id: uuid.UUID
    cliente: str
    rut: str
    estado: str
    comuna: Optional[str] = None
    superficie_total_ha: Optional[float] = None
    n_solicitud: Optional[str] = None
    documentos: int = 0
    fotografias: int = 0
    plantaciones: int = 0
    creado: datetime
    actualizado: datetime
    #: Falso cuando el informe es de otra persona y a uno se lo compartieron.
    #: Nulo en los informes anteriores a que hubiera cuentas, que no tienen
    #: dueno: la pantalla los muestra sin etiqueta en vez de inventarle una.
    soy_el_dueno: Optional[bool] = None


class DetalleExpediente(ResumenExpediente):
    ejecutivo: Optional[str] = None
    sucursal: Optional[str] = None
    agronomo_informante: Optional[str] = None
    banco: Optional[str] = None
    fecha_visita: Optional[datetime] = None
    fecha_emision: Optional[datetime] = None
    region: Optional[str] = None
    datos: Dict[str, Any] = Field(default_factory=dict)
    avisos: List[Any] = Field(default_factory=list)
    notas: Optional[str] = None


def _a_resumen(fila, usuario_id: Optional[str] = None) -> ResumenExpediente:
    return ResumenExpediente(
        id=fila.id, cliente=fila.cliente, rut=fila.rut, estado=fila.estado,
        comuna=fila.comuna, superficie_total_ha=fila.superficie_total_ha,
        n_solicitud=fila.n_solicitud,
        documentos=len(fila.documentos), fotografias=len(fila.fotografias),
        plantaciones=len(fila.plantaciones),
        creado=fila.creado, actualizado=fila.actualizado,
        soy_el_dueno=(None if fila.propietario_id is None or usuario_id is None
                      else str(fila.propietario_id) == str(usuario_id)),
    )


def _a_detalle(fila) -> DetalleExpediente:
    return DetalleExpediente(
        **_a_resumen(fila).model_dump(),
        ejecutivo=fila.ejecutivo, sucursal=fila.sucursal,
        agronomo_informante=fila.agronomo_informante, banco=fila.banco,
        fecha_visita=fila.fecha_visita, fecha_emision=fila.fecha_emision,
        region=fila.region, datos=fila.datos or {}, avisos=fila.avisos or [],
        notas=fila.notas,
    )


async def _mio_o_404(repo, id_expediente: str, usuario) -> None:
    """Corta el paso si el informe no es de quien pregunta.

    Devuelve 404 y no 403 a proposito: un 403 confirmaria que ese informe
    existe, y eso ya dice algo del trabajo de otra persona.
    """
    if not await repo.puede_verlo(id_expediente, str(usuario.id)):
        raise HTTPException(
            status.HTTP_404_NOT_FOUND,
            "no existe el expediente {}".format(id_expediente))


async def _soy_dueno_o_no(repo, id_expediente: str, usuario) -> None:
    """Para repartir accesos hay que ser el dueno, no bastar con verlo.

    `_mio_o_404` deja pasar tambien a quien tiene el informe COMPARTIDO, que
    esta bien para leerlo y seguirlo. Pero repartir accesos -compartir, quitar,
    traspasar- es del dueno: si no, a quien le prestaron un informe podria
    quitarselo a su dueno, o pasarselo a un tercero sin que el dueno se entere.

    Un informe SIN dueno -de antes de que hubiera cuentas- lo puede repartir
    cualquiera que lo vea: no hay a quien pedirle permiso, y bloquearlo dejaria
    ese trabajo sin forma de compartirse nunca.
    """
    await _mio_o_404(repo, id_expediente, usuario)
    fila = await repo.obtener(id_expediente)
    if fila.propietario_id is None:
        return
    if str(fila.propietario_id) != str(usuario.id):
        raise HTTPException(
            status.HTTP_403_FORBIDDEN,
            "este informe es de otra persona: se lo compartieron para verlo y "
            "seguirlo, pero repartir accesos le toca a su dueño.")


@router.post(
    "", response_model=DetalleExpediente, status_code=status.HTTP_201_CREATED,
    summary="Abre un expediente nuevo",
)
async def crear(nuevo: NuevoExpediente, repo: Expedientes,
                usuario: UsuarioActual) -> DetalleExpediente:
    fila = await repo.crear(propietario_id=usuario.id,
                            **nuevo.model_dump(exclude_none=True))
    return _a_detalle(fila)


@router.get("", response_model=List[ResumenExpediente],
            summary="Expedientes, del mas reciente al mas antiguo")
async def listar(
    repo: Expedientes,
    usuario: UsuarioActual,
    estado: Optional[str] = Query(None, description="borrador, en_revision, ..."),
    cliente: Optional[str] = Query(None, description="busqueda parcial"),
    limite: int = Query(50, ge=1, le=200),
) -> List[ResumenExpediente]:
    if estado and estado not in ESTADOS:
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST,
            "estado desconocido: {}. Validos: {}".format(estado, ", ".join(ESTADOS)),
        )
    filas = await repo.listar(estado=estado, cliente=cliente, limite=limite,
                              usuario_id=str(usuario.id))
    return [_a_resumen(f, str(usuario.id)) for f in filas]


@router.get("/estadisticas", summary="Cuantos expedientes hay y en que estado")
async def estadisticas(repo: Expedientes) -> Dict[str, Any]:
    return await repo.estadisticas()


@router.get("/{id_expediente}", response_model=DetalleExpediente,
            summary="Un expediente completo")
async def obtener(id_expediente: str, repo: Expedientes,
                  usuario: UsuarioActual) -> DetalleExpediente:
    await _mio_o_404(repo, id_expediente, usuario)
    try:
        return _a_detalle(await repo.obtener(id_expediente))
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error


@router.patch("/{id_expediente}", response_model=DetalleExpediente,
              summary="Actualiza lo que haya cambiado")
async def actualizar(
    id_expediente: str, cambios: CambioExpediente, repo: Expedientes,
    usuario: UsuarioActual,
) -> DetalleExpediente:
    await _mio_o_404(repo, id_expediente, usuario)
    if cambios.estado and cambios.estado not in ESTADOS:
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST,
            "estado desconocido: {}".format(cambios.estado),
        )
    try:
        fila = await repo.actualizar(
            id_expediente, **cambios.model_dump(exclude_none=True)
        )
        return _a_detalle(fila)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error


@router.delete("/{id_expediente}", status_code=status.HTTP_204_NO_CONTENT)
async def borrar(id_expediente: str, repo: Expedientes,
                 usuario: UsuarioActual) -> None:
    await _mio_o_404(repo, id_expediente, usuario)
    try:
        await repo.borrar(id_expediente)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error


class Destinatario(BaseModel):
    usuario_id: str = Field(description="con quien se comparte o a quien se pasa")


@router.get("/{id_expediente}/compartido",
            summary="Con quien esta compartido este informe")
async def ver_compartido(
    id_expediente: str, repo: Expedientes, usuario: UsuarioActual
) -> Dict[str, Any]:
    await _mio_o_404(repo, id_expediente, usuario)
    fila = await repo.obtener(id_expediente)
    con = await repo.compartido_con(id_expediente)
    return {
        "propietario_id": str(fila.propietario_id) if fila.propietario_id else None,
        "soy_el_dueno": str(fila.propietario_id) == str(usuario.id),
        # Con puntos y guion, igual que en el resto de la pantalla: el RUT se
        # guarda normalizado, pero nadie lo lee asi.
        "compartido_con": [{"id": str(u.id), "nombre": u.nombre,
                            "rut": formatear_rut(u.rut)} for u in con],
    }


@router.post("/{id_expediente}/compartir",
             summary="Deja que otra persona vea y siga este informe")
async def compartir(
    id_expediente: str, destino: Destinatario, repo: Expedientes,
    usuario: UsuarioActual,
) -> Dict[str, Any]:
    """Compartir NO quita el informe a nadie: los dos pueden seguir trabajando.

    Para pasarselo de verdad -que deje de ser mio- esta 'traspasar'.
    """
    await _soy_dueno_o_no(repo, id_expediente, usuario)
    await repo.compartir(id_expediente, destino.usuario_id)
    return {"compartido": True}


@router.post("/{id_expediente}/dejar-de-compartir",
             summary="Quita el acceso a esa persona")
async def dejar_de_compartir(
    id_expediente: str, destino: Destinatario, repo: Expedientes,
    usuario: UsuarioActual,
) -> Dict[str, Any]:
    await _soy_dueno_o_no(repo, id_expediente, usuario)
    await repo.dejar_de_compartir(id_expediente, destino.usuario_id)
    return {"compartido": False}


@router.post("/{id_expediente}/traspasar",
             summary="Le pasa el informe a otra persona")
async def traspasar(
    id_expediente: str, destino: Destinatario, repo: Expedientes,
    usuario: UsuarioActual,
) -> Dict[str, Any]:
    """Cambia el dueno. Al anterior se le deja compartido a proposito: pasarle
    un informe a alguien no deberia hacerlo desaparecer de tu lista sin aviso,
    y a veces hay que volver a mirarlo."""
    await _soy_dueno_o_no(repo, id_expediente, usuario)
    try:
        await repo.traspasar(id_expediente, destino.usuario_id)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error
    return {"traspasado": True}


class CatastroDelExpediente(BaseModel):
    plantaciones: List[Plantacion] = Field(default_factory=list)


@router.put("/{id_expediente}/catastro",
            summary="Guarda el catastro del expediente (reemplaza el anterior)")
async def guardar_catastro(
    id_expediente: str, catastro: CatastroDelExpediente, repo: Expedientes,
    usuario: UsuarioActual,
) -> Dict[str, Any]:
    """Reemplaza, no acumula.

    Si el tasador reimporta el catastro corregido, no pueden quedar hileras de
    la version anterior mezcladas con las nuevas.

    Hasta el 2026-10-05 esta ruta no preguntaba de quien era el informe:
    cualquiera con sesion podia reemplazar el catastro de un informe ajeno
    sabiendo su id. Ahora pasa por la misma puerta que leerlo.
    """
    await _mio_o_404(repo, id_expediente, usuario)
    try:
        await repo.obtener(id_expediente)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error

    guardadas = await repo.reemplazar_plantaciones(
        id_expediente, catastro.plantaciones
    )
    return {
        "plantaciones_guardadas": guardadas,
        "resumen": await repo.resumen_por_especie(id_expediente),
    }


@router.get(
    "/{id_expediente}/curvas-faltantes",
    summary="Que curvas de produccion le faltan a ESTE expediente",
)
async def curvas_faltantes(
    id_expediente: str, repo: Expedientes, almacen: Almacenamiento,
    usuario: UsuarioActual,
) -> Dict[str, Any]:
    """Lo que falta para poder proyectar el flujo de este cliente.

    Antes esta lista vivia en el navegador (``localStorage``) y se llenaba al
    importar el catastro. Tenia dos problemas que aparecieron al usarla:

    1. **No decia de que informe era.** El tasador la dejaba a medias, volvia
       dias despues y veia "faltan estas especies" sin saber para que cliente.
    2. **No salia del navegador.** Ahora que se entra desde otro computador
       por Tailscale, el navegador de el agrónomo tiene su propio almacenamiento:
       la lista simplemente no existiria para el.

    Colgada del expediente se arregla sola: se calcula del catastro guardado
    cada vez que se pide, asi que siempre esta al dia -si alguien creo la curva
    desde otra maquina, aqui ya aparece hecha- y siempre se sabe de quien es.
    """
    from services.envejecimiento import cobertura_de_curvas

    # Dice el cliente, sus hectareas y sus especies: lo mismo que verlo.
    await _mio_o_404(repo, id_expediente, usuario)
    try:
        fila = await repo.obtener(id_expediente)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error

    plantaciones = await repo.plantaciones(id_expediente)
    curvas = await almacen.listar_curvas()
    cobertura = cobertura_de_curvas(plantaciones, curvas)

    # Una linea por lo que hay que resolver: por variedad cuando la especie
    # trae varias, porque una curva de Gala no es la de Fuji aunque el
    # emparejador se la preste (ver services/envejecimiento.py).
    from schemas.informe import Plantacion
    from services.catastro_import import normalizar
    from services.envejecimiento import emparejar_curva

    items: List[Dict[str, Any]] = []
    for entrada in cobertura:
        variedades = entrada.get("variedades") or []
        candidatas = ([{"variedad": v["variedad"],
                        "superficie_ha": v["superficie_ha"],
                        "plantaciones": v["plantaciones"]} for v in variedades]
                      if len(variedades) > 1 else
                      [{"variedad": variedades[0]["variedad"] if variedades else None,
                        "superficie_ha": entrada["superficie_ha"],
                        "plantaciones": entrada["plantaciones"]}])
        for candidata in candidatas:
            curva = emparejar_curva(
                Plantacion(especie=entrada["especie"],
                           variedad=candidata["variedad"],
                           superficie_ha=0.1),
                curvas)
            if curva is None:
                estado = "falta"
            elif normalizar(curva.variedad or "") == normalizar(
                    candidata["variedad"] or ""):
                estado = "propia"
            elif not (curva.variedad or "").strip():
                estado = "de_la_especie"
            else:
                estado = "de_otra_variedad"
            items.append({
                "especie": entrada["especie"],
                "variedad": candidata["variedad"],
                "superficie_ha": candidata["superficie_ha"],
                "plantaciones": candidata["plantaciones"],
                "estado": estado,
                "curva": curva.nombre if curva else None,
                "variedad_de_la_curva": curva.variedad if curva else None,
            })

    resueltas = sum(1 for i in items
                    if i["estado"] in ("propia", "de_la_especie"))
    sin_curva = sum(i["superficie_ha"] for i in items if i["estado"] == "falta")
    return {
        "expediente": id_expediente,
        "cliente": fila.cliente,
        "plantaciones": len(plantaciones),
        "superficie_total_ha": round(
            sum(p.superficie_ha for p in plantaciones), 2),
        "superficie_sin_curva_ha": round(sin_curva, 2),
        "resueltas": resueltas,
        "total": len(items),
        "items": items,
    }


@router.get("/{id_expediente}/resumen",
            summary="Superficie por especie, agregada por la base de datos")
async def resumen(id_expediente: str, repo: Expedientes,
                  usuario: UsuarioActual) -> List[Dict[str, Any]]:
    await _mio_o_404(repo, id_expediente, usuario)
    return await repo.resumen_por_especie(id_expediente)
