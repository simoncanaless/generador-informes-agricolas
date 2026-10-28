"""Generacion de informes Excel. 100% determinista, sin IA de por medio."""

from __future__ import annotations

import re
import secrets
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from fastapi import Depends, APIRouter, HTTPException, Query, status
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from api.dependencias import obtener_usuario_actual
from api.dependencias import Config, Motor
from schemas.informe import InformeFitogreen

# Todo lo de aqui exige sesion iniciada. La dependencia va en el router
# entero y no en cada endpoint: asi una ruta nueva nace protegida.
router = APIRouter(prefix="/informes", tags=["informes"],
                   dependencies=[Depends(obtener_usuario_actual)])

TIPO_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class RespuestaGeneracion(BaseModel):
    """Resultado de generar un informe.

    Se devuelve JSON y no el archivo directamente para no perder los avisos:
    una colision de celdas combinadas o una tabla truncada tiene que llegarle
    al tasador antes de que mande el informe al banco.
    """

    archivo: str
    descarga: str
    celdas_escritas: int
    filas_insertadas: int
    imagenes_insertadas: int
    avisos: List[str] = Field(default_factory=list)
    generado: datetime


class UnidadExcluible(BaseModel):
    """Una hoja o tabla que el tasador puede sacar del .xlsx antes de bajarlo.

    Ver ``config/mapping_santander.json:unidades_excluibles`` y
    ``services/exclusiones.py``: ``depende_de`` son los ids de OTRAS
    unidades cuyas celdas leen las formulas de esta -si esas se excluyen,
    esta se cae con ellas-. La pantalla arma la cascada con el mismo
    algoritmo que el servidor (``resolver_cascada``), asi que tildar una
    casilla se siente instantaneo sin ir y volver al servidor, pero el
    servidor SIEMPRE la vuelve a calcular antes de tocar el archivo.
    """

    id: str
    etiqueta: str
    depende_de: List[str] = Field(default_factory=list)


class RespuestaValidacion(BaseModel):
    valido: bool
    cliente: Optional[str] = None
    plantaciones: int = 0
    superficie_declarada_ha: float = 0.0
    avisos: List[str] = Field(default_factory=list)


def _nombre_seguro(texto: str) -> str:
    """Convierte el nombre del cliente en un nombre de archivo sin sorpresas."""
    sin_tildes = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    limpio = re.sub(r"[^A-Za-z0-9]+", "_", sin_tildes).strip("_")
    return (limpio[:60] or "informe").lower()


@router.post(
    "/generar",
    response_model=RespuestaGeneracion,
    summary="Genera el .xlsx del informe sobre la plantilla del banco",
)
async def generar(
    informe: InformeFitogreen,
    motor: Motor,
    config: Config,
    galeria_fotos: bool = Query(
        False,
        description="anade una hoja de fotografias que la plantilla del banco "
                    "no contempla; por defecto se respeta el formato original",
    ),
    excluir: List[str] = Query(
        default_factory=list,
        description="ids de config/mapping_santander.json:unidades_excluibles "
                    "-hojas o tablas que el tasador pidio no llevar en este "
                    "informe. El servidor vuelve a resolver la cascada de "
                    "dependencias (nunca confia en lo que calculo la pantalla).",
    ),
) -> RespuestaGeneracion:
    marca = datetime.now()
    nombre = "{}_{}_{}.xlsx".format(
        _nombre_seguro(informe.antecedentes.cliente),
        marca.strftime("%Y%m%d_%H%M%S"),
        secrets.token_hex(8),   # el enlace no se tiene que poder adivinar
    )
    destino = config.ruta_salida / nombre

    resultado = await motor.generar_async(informe, destino, galeria_fotos, excluir)

    return RespuestaGeneracion(
        archivo=nombre,
        descarga="/informes/{}/descargar".format(nombre),
        celdas_escritas=resultado.celdas_escritas,
        filas_insertadas=resultado.filas_insertadas,
        imagenes_insertadas=resultado.imagenes_insertadas,
        avisos=resultado.avisos,
        generado=marca,
    )


@router.get(
    "/unidades-excluibles",
    response_model=List[UnidadExcluible],
    summary="Hojas y tablas que se pueden sacar del .xlsx antes de generarlo",
)
async def unidades_excluibles(motor: Motor) -> List[UnidadExcluible]:
    """El registro completo, para que la pantalla pinte el checklist y arme
    la cascada de dependencias sin tener que preguntarle al servidor en
    cada click. Ver ``UnidadExcluible``."""
    # El formato final de Santander (MotorSantander, mapping en dict) no
    # declara unidades excluibles todavia: la lista sale vacia.
    excluibles = getattr(motor.mapa, "unidades_excluibles", None) or {}
    return [
        UnidadExcluible(id=id_, etiqueta=unidad.etiqueta, depende_de=unidad.depende_de)
        for id_, unidad in excluibles.items()
    ]


@router.post(
    "/validar",
    response_model=RespuestaValidacion,
    summary="Valida el informe sin generar archivo",
)
async def validar(informe: InformeFitogreen) -> RespuestaValidacion:
    """Comprueba coherencias que Pydantic no puede expresar por si solo.

    Ojo: aqui NO se calcula nada que vaya al informe. Solo se avisa de
    inconsistencias para que el tasador las revise.
    """
    avisos: List[str] = []
    plantaciones = informe.estructura_productiva.plantaciones

    suma_plantaciones = sum(p.superficie_ha for p in plantaciones)
    suma_declarada = sum(f.superficie_ha for f in informe.estructura_temporada)
    if suma_declarada and abs(suma_plantaciones - suma_declarada) > 0.5:
        avisos.append(
            "la superficie del catastro ({:.2f} ha) no cuadra con la escala "
            "productiva declarada ({:.2f} ha)".format(suma_plantaciones, suma_declarada)
        )

    for i, planta in enumerate(plantaciones, start=1):
        if planta.superficie_ha <= 0:
            avisos.append("plantacion {}: superficie en cero".format(i))

    for derecho in informe.dotacion_recursos.derechos_agua:
        nominal, real = derecho.caudal_nominal_ls, derecho.caudal_real_ls
        if nominal is not None and real is not None and real > nominal:
            avisos.append(
                "derecho de agua de '{}': el caudal real ({}) supera al nominal "
                "({})".format(derecho.predio, real, nominal)
            )

    if not informe.antecedentes.fecha_visita:
        avisos.append("falta la fecha de visita a terreno")

    return RespuestaValidacion(
        valido=True,
        cliente=informe.antecedentes.cliente,
        plantaciones=len(plantaciones),
        superficie_declarada_ha=round(suma_plantaciones, 2),
        avisos=avisos,
    )


@router.get(
    "/{nombre}/descargar",
    response_class=FileResponse,
    summary="Descarga un informe ya generado",
)
async def descargar(nombre: str, config: Config) -> FileResponse:
    # El nombre viene de la URL: se resuelve y se comprueba que no escapa de
    # la carpeta de salida.
    destino = (config.ruta_salida / nombre).resolve()
    if destino.parent != config.ruta_salida.resolve() or not destino.is_file():
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="no existe el informe '{}'".format(nombre),
        )
    return FileResponse(destino, media_type=TIPO_XLSX, filename=destino.name)
