"""Endpoints del 20% probabilistico.

Toda respuesta declara que modelo la atendio y si la cadena tuvo que degradarse,
porque un informe bancario debe poder auditarse dato a dato.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from fastapi import Depends, APIRouter, File, Form, UploadFile
from pydantic import BaseModel, Field

from api.dependencias import obtener_usuario_actual
from api.dependencias import Config, IA, Parser, leer_acotado
from schemas.extraccion import (
    ExtraccionLegal, LoteFotografias, NotaTerreno, RedaccionTecnica,
)

# Todo lo de aqui exige sesion iniciada. La dependencia va en el router
# entero y no en cada endpoint: asi una ruta nueva nace protegida.
router = APIRouter(prefix="/ia", tags=["ia"],
                   dependencies=[Depends(obtener_usuario_actual)])

MIMES_IMAGEN = {"image/jpeg", "image/png", "image/webp", "image/heic", "image/heif"}


class Trazabilidad(BaseModel):
    """De donde salio la respuesta. Se adjunta a todo resultado de IA."""

    modelo: str
    degradado: bool = Field(
        description="True si tuvo que usarse un modelo peor que el preferido"
    )
    intentos: int
    duracion_s: float
    incidencias: List[str] = Field(default_factory=list)


class RespuestaDocumento(BaseModel):
    datos: ExtraccionLegal
    via: str = Field(description="'texto' o 'vision' segun el PDF tuviera capa de texto")
    paginas: int
    caracteres_enviados: int
    tokens_estimados: int
    advertencias: List[str] = Field(default_factory=list)
    trazabilidad: Trazabilidad


class RespuestaFotografias(BaseModel):
    resultado: LoteFotografias
    trazabilidad: Trazabilidad


class RespuestaNota(BaseModel):
    resultado: NotaTerreno
    trazabilidad: Trazabilidad


class RespuestaOpinion(BaseModel):
    resultado: RedaccionTecnica
    trazabilidad: Trazabilidad


def _traza(respuesta: Any) -> Trazabilidad:
    return Trazabilidad(
        modelo=respuesta.modelo,
        degradado=respuesta.degradado,
        intentos=respuesta.intentos,
        duracion_s=round(respuesta.duracion_s, 2),
        incidencias=respuesta.incidencias,
    )


@router.post(
    "/documentos/extraer",
    response_model=RespuestaDocumento,
    summary="Extrae antecedentes legales de un PDF (escritura, certificado, CBR)",
)
async def extraer_documento(
    parser: Parser,
    config: Config,
    archivo: UploadFile = File(..., description="PDF de hasta MAX_MB_PDF"),
    instrucciones: Optional[str] = Form(
        None, description="indicaciones extra para este documento concreto"
    ),
) -> RespuestaDocumento:
    contenido = await leer_acotado(archivo, config.max_bytes_pdf, "el PDF")
    resultado = await parser.extraer(contenido, instrucciones=instrucciones)

    return RespuestaDocumento(
        datos=resultado.datos,
        via=resultado.via,
        paginas=resultado.conversion.paginas,
        caracteres_enviados=resultado.conversion.caracteres,
        tokens_estimados=resultado.conversion.tokens_estimados,
        advertencias=resultado.conversion.advertencias,
        trazabilidad=Trazabilidad(
            modelo=resultado.modelo,
            degradado=resultado.degradado,
            intentos=1,
            duracion_s=resultado.duracion_s,
            incidencias=resultado.incidencias,
        ),
    )


@router.post(
    "/fotografias/titulos",
    response_model=RespuestaFotografias,
    summary="Propone el titulo y la categoria de fotografias de terreno",
)
async def titular_fotografias(
    ia: IA,
    config: Config,
    archivos: List[UploadFile] = File(..., description="una o varias imagenes"),
    contexto: Optional[str] = Form(
        None, description="especie, sector o cualquier dato que ayude a titularlas"
    ),
) -> RespuestaFotografias:
    imagenes: List[bytes] = []
    mime = "image/jpeg"
    for archivo in archivos:
        if archivo.content_type in MIMES_IMAGEN:
            mime = archivo.content_type
        imagenes.append(
            await leer_acotado(
                archivo, config.max_bytes_imagen, archivo.filename or "la imagen"
            )
        )

    respuesta = await ia.titular_fotografias(imagenes, mime, contexto)
    return RespuestaFotografias(resultado=respuesta.valor, trazabilidad=_traza(respuesta))


@router.post(
    "/notas-voz/estructurar",
    response_model=RespuestaNota,
    summary="Transcribe y estructura una nota de voz tomada en terreno",
)
async def estructurar_nota(
    ia: IA,
    config: Config,
    archivo: UploadFile = File(..., description="audio ogg/mp3/m4a/wav"),
    contexto: Optional[str] = Form(None),
) -> RespuestaNota:
    audio = await leer_acotado(archivo, config.max_bytes_pdf, "el audio")
    mime = archivo.content_type or "audio/ogg"
    respuesta = await ia.estructurar_nota_voz(audio, mime, contexto)
    return RespuestaNota(resultado=respuesta.valor, trazabilidad=_traza(respuesta))


class PeticionOpinion(BaseModel):
    apartado: str = Field(
        description="gestion administrativa, canales comerciales, calidad de "
                    "activos o conclusion final",
    )
    hechos: str = Field(
        min_length=20,
        description="hechos verificados; la IA no puede agregar ninguno mas",
    )


@router.post(
    "/opinion/redactar",
    response_model=RespuestaOpinion,
    summary="Propone el borrador de un apartado de opinion profesional",
)
async def redactar_opinion(peticion: PeticionOpinion, ia: IA) -> RespuestaOpinion:
    respuesta = await ia.redactar_opinion(peticion.hechos, peticion.apartado)
    return RespuestaOpinion(resultado=respuesta.valor, trazabilidad=_traza(respuesta))


@router.get("/estado", summary="Estado detallado de la cadena de modelos")
def estado(ia: IA) -> Dict[str, Any]:
    return ia.estado()
