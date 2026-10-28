"""Generacion de retasaciones Rabobank. 100% determinista, sin IA de por medio.

Mismo patron que api/routers/informes.py (recibe el dominio ya armado y
genera el .xlsx), pero con su propio motor -no comparte MotorExcel con
Santander: es un dominio distinto, "no mezclar los motores".
"""

from __future__ import annotations

import re
import secrets
import unicodedata
from io import BytesIO
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

from fastapi import APIRouter, Body, Depends, File, HTTPException, UploadFile, status
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from api.dependencias import Config, MotorRabo, leer_acotado, obtener_usuario_actual
from services.importar_retasacion import leer_retasacion_excel
from services.importar_retasacion_pdf import leer_retasacion_pdf
from services.retasacion_pendientes import leer_borrador

router = APIRouter(prefix="/retasaciones", tags=["retasaciones"],
                   dependencies=[Depends(obtener_usuario_actual)])

TIPO_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class PendienteRetasacion(BaseModel):
    """Algo que el informe necesita y no vino ("falta"), o que vino y no se
    pudo usar ("aviso") -ver services/retasacion_pendientes.py-."""

    seccion: str
    detalle: str
    tipo: str


class RespuestaRevision(BaseModel):
    pendientes: List[PendienteRetasacion] = Field(default_factory=list)


class RespuestaGeneracionRabobank(BaseModel):
    archivo: str
    descarga: str
    celdas_escritas: int
    filas_insertadas: int
    avisos: List[str] = Field(default_factory=list)
    #: Lo que quedo sin llenar -el informe se genera igual-. Distinto de
    #: `avisos`, que son observaciones del motor al escribir.
    pendientes: List[PendienteRetasacion] = Field(default_factory=list)
    #: Los huecos de imagen del formato (foto del predio, certificados,
    #: mapas, fotografias): en el Excel van con una imagen de muestra gris
    #: que se reemplaza con clic derecho > Cambiar imagen.
    imagenes_por_pegar: List[str] = Field(default_factory=list)
    generado: datetime


class UnidadPlant(BaseModel):
    """Las claves que aceptan las plantaciones sin royalty -especie contra
    PLANT!Tabla12, estructura contra PLANT!Tabla5-. Sirve para que la
    pantalla ofrezca un desplegable en vez de texto libre: una especie que
    no calza no da error, da 0 en silencio (ver services/excel_engine_rabobank.py)."""

    especies: List[str]
    estructuras: List[str]


def _nombre_seguro(texto: str) -> str:
    sin_tildes = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    limpio = re.sub(r"[^A-Za-z0-9]+", "_", sin_tildes).strip("_")
    return (limpio[:60] or "retasacion").lower()


@router.get(
    "/plant-biblioteca",
    response_model=UnidadPlant,
    summary="Especies y estructuras válidas para el costo de reposición de plantaciones",
)
async def plant_biblioteca(motor: MotorRabo) -> UnidadPlant:
    validacion = getattr(motor.mapa, "plant_validacion", None) or {}
    return UnidadPlant(
        especies=validacion.get("especies", []),
        estructuras=validacion.get("estructuras", []),
    )


class RespuestaLecturaAnterior(BaseModel):
    """Un borrador SUELTO, no validado contra RetasacionRabobank todavía: un
    campo que el archivo antiguo no traía simplemente no aparece, y el
    tasador lo completa a mano en el Paso 2 -mismo criterio que el resto
    del proyecto para un dato que no se pudo leer."""

    datos: Dict[str, Any] = Field(default_factory=dict)
    avisos: List[str] = Field(default_factory=list)


@router.post(
    "/leer-anterior",
    response_model=RespuestaLecturaAnterior,
    summary="Lee un .xlsx o .pdf de una retasación anterior para precargar el Paso 2",
)
async def leer_anterior(
    motor: MotorRabo,
    config: Config,
    archivo: UploadFile = File(...),
) -> RespuestaLecturaAnterior:
    """100% determinista, sin IA -las dos vías (excel_engine_rabobank en
    reversa para el .xlsx; posición x/y de cada palabra para el .pdf,
    ver services/importar_retasacion_pdf.py). Si el archivo no calza con
    lo que cada lector espera, vuelve vacío con un aviso, no con datos
    inventados."""
    nombre = (archivo.filename or "").lower()
    if not nombre.endswith((".xlsx", ".pdf")):
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST, "se espera un archivo .xlsx o .pdf")
    # Con tope, como las demas subidas: un .xlsx es un zip, y uno armado para
    # inflarse al abrirlo puede ocupar toda la memoria del computador.
    contenido = BytesIO(await leer_acotado(
        archivo, config.max_bytes_pdf, "el archivo anterior"))
    if nombre.endswith(".xlsx"):
        try:
            datos, avisos = leer_retasacion_excel(contenido, motor.mapa)
        except Exception as error:  # noqa: BLE001
            raise HTTPException(
                status.HTTP_422_UNPROCESSABLE_ENTITY,
                "no se pudo leer el archivo: {}".format(error)) from error
    elif nombre.endswith(".pdf"):
        try:
            datos, avisos = leer_retasacion_pdf(contenido)
        except Exception as error:  # noqa: BLE001
            raise HTTPException(
                status.HTTP_422_UNPROCESSABLE_ENTITY,
                "no se pudo leer el archivo: {}".format(error)) from error
    return RespuestaLecturaAnterior(datos=datos, avisos=avisos)


@router.post(
    "/revisar",
    response_model=RespuestaRevision,
    summary="Lista lo que le falta a una retasación, sin generar nada",
)
async def revisar(datos: Dict[str, Any] = Body(...)) -> RespuestaRevision:
    """La pantalla lo consulta ANTES de generar, para preguntar «¿generar con
    esto pendiente?». Es la misma lectura que hace /generar, asi que lo que
    se muestra en esa pregunta es exactamente lo que va a quedar vacio."""
    _, pendientes = leer_borrador(datos)
    return RespuestaRevision(pendientes=[p.como_dict() for p in pendientes])


@router.post(
    "/generar",
    response_model=RespuestaGeneracionRabobank,
    summary="Genera el .xlsx de la retasación sobre la plantilla de Rabobank",
)
async def generar(
    motor: MotorRabo,
    config: Config,
    datos: Dict[str, Any] = Body(...),
) -> RespuestaGeneracionRabobank:
    """Acepta una retasacion INCOMPLETA (2026-10-02): lo que falta queda
    vacio en el Excel y vuelve en `pendientes`, en vez de rechazar el
    informe entero. Lo que vino mal (un RUT invalido) no se escribe."""
    retasacion, pendientes = leer_borrador(datos)

    marca = datetime.now()
    nombre = "{}_{}_{}.xlsx".format(
        _nombre_seguro(retasacion.cliente.nombre or "retasacion"),
        marca.strftime("%Y%m%d_%H%M%S"),
        # 16 caracteres al azar: el enlace de descarga es la unica llave del
        # archivo, asi que no se tiene que poder adivinar (antes eran 6).
        secrets.token_hex(8),
    )
    destino = config.ruta_salida / nombre

    resultado = await motor.generar_async(retasacion, destino)

    return RespuestaGeneracionRabobank(
        archivo=nombre,
        descarga="/retasaciones/{}/descargar".format(nombre),
        celdas_escritas=resultado.celdas_escritas,
        filas_insertadas=resultado.filas_insertadas,
        avisos=resultado.avisos,
        pendientes=[p.como_dict() for p in pendientes],
        imagenes_por_pegar=[spec["titulo"] for spec in
                            (getattr(motor.mapa, "huecos_de_imagen", None) or {}).values()
                            if isinstance(spec, dict)],
        generado=marca,
    )


@router.get("/{nombre}/descargar", summary="Descarga un .xlsx ya generado")
async def descargar(nombre: str, config: Config) -> FileResponse:
    # Mismo guardia que informes.py: nombre no puede escapar de ruta_salida.
    destino = (config.ruta_salida / nombre).resolve()
    if destino.parent != config.ruta_salida.resolve() or not destino.is_file():
        raise HTTPException(status.HTTP_404_NOT_FOUND, "no existe: {}".format(nombre))
    return FileResponse(destino, media_type=TIPO_XLSX, filename=destino.name)
