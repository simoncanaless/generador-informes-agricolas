"""Esquemas de salida de la capa de IA (el 20% probabilistico).

Regla de oro del proyecto: **la IA no calcula**. Estos modelos solo admiten
hechos transcritos del documento o de la fotografia. No hay totales, ni sumas,
ni superficies derivadas, ni valores de tasacion: todo eso lo produce el motor
determinista a partir de estos datos ya validados.

Cada bloque extraido lleva ``confianza`` y ``fuente`` para que el tasador pueda
verificar el dato contra el documento original antes de firmar. Un campo que no
aparece en el documento se devuelve como ``null``: inventarlo es peor que
dejarlo vacio.
"""

from __future__ import annotations

from enum import Enum
from typing import List, Optional

from pydantic import BaseModel, ConfigDict, Field, field_validator

from schemas.comun import sin_signo_de_formula


class Base(BaseModel):
    # 'ignore' y no 'forbid': si el modelo agrega un campo de mas preferimos
    # descartarlo a que reviente la peticion entera del tasador.
    model_config = ConfigDict(extra="ignore", str_strip_whitespace=True)

    # Ningun texto del dominio llega al Excel como formula (schemas/comun.py).
    @field_validator("*", mode="after")
    @classmethod
    def _sin_formulas(cls, valor):
        return sin_signo_de_formula(valor)


class Confianza(str, Enum):
    ALTA = "alta"        # el dato aparece literal y sin ambiguedad
    MEDIA = "media"      # aparece pero requiere interpretacion
    BAJA = "baja"        # deducido del contexto; revisar siempre


class Procedencia(Base):
    """De donde sale un dato, para poder auditarlo."""

    confianza: Confianza = Confianza.BAJA
    fuente: Optional[str] = Field(
        default=None,
        description="cita textual breve del documento que respalda el dato",
    )
    pagina: Optional[int] = None


# ------------------------------------------------- documentos legales

class InscripcionCBR(Base):
    """Inscripcion en el Conservador de Bienes Raices.

    Las escrituras chilenas escriben estos numeros con palabras. Van SIEMPRE en
    cifras: es transcripcion, no calculo, y asi entran en el informe.
    """

    fojas: Optional[str] = Field(
        default=None,
        description="numero de fojas EN CIFRAS. 'fojas mil doscientos cuarenta' "
                    "se devuelve como '1240', nunca en palabras",
    )
    numero: Optional[str] = Field(
        default=None,
        description="numero de inscripcion EN CIFRAS. 'numero ochocientos "
                    "noventa' se devuelve como '890'",
    )
    anio: Optional[int] = Field(
        default=None,
        description="ano en cifras. 'dos mil once' se devuelve como 2011",
    )
    conservador: Optional[str] = Field(
        default=None, description="nombre del CBR, ej. 'CBR de Rancagua'"
    )
    procedencia: Procedencia = Procedencia()


class DerechoAguaExtraido(Base):
    """Un derecho de aprovechamiento de aguas tal como figura en la escritura."""

    tipo_fuente: Optional[str] = Field(
        default=None, description="pozo, canal, rio, vertiente, embalse"
    )
    nombre_fuente: Optional[str] = None
    caudal_litros_segundo: Optional[float] = Field(default=None, ge=0)
    acciones: Optional[float] = Field(
        default=None, ge=0, description="numero de acciones o regadores"
    )
    naturaleza: Optional[str] = Field(
        default=None, description="consuntivo/no consuntivo, permanente/eventual"
    )
    inscripcion: Optional[InscripcionCBR] = None
    procedencia: Procedencia = Procedencia()


class PropiedadExtraida(Base):
    """Identificacion del inmueble segun la escritura o el certificado."""

    propietario: Optional[str] = None
    rut_propietario: Optional[str] = None
    nombre_predio: Optional[str] = None
    rol_avaluo: Optional[str] = Field(default=None, description="formato 'NNN-NN'")
    comuna: Optional[str] = None
    provincia: Optional[str] = None
    region: Optional[str] = None
    superficie_escritura_ha: Optional[float] = Field(
        default=None, ge=0,
        description="superficie tal como la declara el documento, sin recalcular",
    )
    deslindes: Optional[str] = Field(
        default=None, description="norte/sur/oriente/poniente en texto"
    )
    inscripcion: Optional[InscripcionCBR] = None
    procedencia: Procedencia = Procedencia()


class ExtraccionLegal(Base):
    """Resultado completo de leer un documento legal."""

    tipo_documento: Optional[str] = Field(
        default=None,
        description="escritura de compraventa, certificado de dominio vigente, "
                    "certificado de aguas, plano, otro",
    )
    propiedad: PropiedadExtraida = PropiedadExtraida()
    derechos_agua: List[DerechoAguaExtraido] = Field(default_factory=list)
    gravamenes: List[str] = Field(
        default_factory=list,
        description="hipotecas, prohibiciones, servidumbres, usufructos",
    )
    advertencias: List[str] = Field(
        default_factory=list,
        description="ilegible, pagina faltante, dato contradictorio, firma ausente",
    )


# ------------------------------------------------------- fotografias

class TituloFotografia(Base):
    """El titulo de una imagen del predio, y en que parte del informe va.

    Solo eso, a proposito. Hasta el 2026-10-05 el modelo redactaba ademas un
    parrafo tecnico, la especie, el estado aparente («Bueno, Regular o
    Malo») y los elementos que veia. el agrónomo: no hace falta una descripcion
    tecnica para ningun tipo de imagen, porque eso lo detecta mejor el
    agronomo que estuvo en el predio. Queda un titulo corto -que el tasador
    puede cambiar en la revision- y la categoria, que no es descripcion sino
    el lugar de la imagen en el informe (la de emplazamiento va al hueco de la
    caratula; el resto, a la galeria).

    Las lecturas guardadas antes de este cambio traen tambien los campos
    viejos: se leen igual, porque esta clase ignora lo que no conoce.
    """

    titulo: str = Field(description="rotulo corto, el que va debajo de la imagen")
    categoria_sugerida: Optional[str] = Field(
        default=None,
        description="plantacion, infraestructura, riego, emplazamiento, otra",
    )


class LoteFotografias(Base):
    titulos: List[TituloFotografia] = Field(default_factory=list)


# --------------------------------------------- notas de voz en terreno

class NotaTerreno(Base):
    """Estructuracion de una nota de voz del tasador en terreno."""

    transcripcion: str = Field(description="transcripcion literal de lo dicho")
    resumen: Optional[str] = None
    sector: Optional[str] = None
    especie: Optional[str] = None
    variedad: Optional[str] = None
    anio_plantacion: Optional[str] = None
    superficie_ha: Optional[float] = Field(default=None, ge=0)
    marco_plantacion: Optional[str] = Field(
        default=None, description="literal, ej. '5 x 2' o '4x1,5'"
    )
    sistema_riego: Optional[str] = None
    estado_fitosanitario: Optional[str] = None
    observaciones: List[str] = Field(default_factory=list)
    procedencia: Procedencia = Procedencia()


# ------------------------------------------------------ redaccion libre

class ColumnaMapeada(Base):
    """Una columna del catastro ajeno y a que campo del informe corresponde."""

    letra: str = Field(description="letra de la columna en la planilla, ej. 'C'")
    campo: Optional[str] = Field(
        default=None,
        description="nombre exacto del campo destino, o null si no corresponde "
                    "a ninguno de los ofrecidos",
    )
    confianza: Confianza = Confianza.BAJA
    motivo: Optional[str] = Field(
        default=None, description="por que se asigno asi, en pocas palabras"
    )


class MapeoColumnas(Base):
    """Propuesta de correspondencia entre un catastro ajeno y el formato propio.

    La IA solo opina sobre ENCABEZADOS. Los datos de las filas no los toca
    nadie mas que el codigo.
    """

    columnas: List[ColumnaMapeada] = Field(default_factory=list)
    observaciones: List[str] = Field(default_factory=list)


class RedaccionTecnica(Base):
    """Parrafo de opinion profesional propuesto por la IA.

    Es un BORRADOR: el agronomo firma, no el modelo.
    """

    texto: str
    calificacion_sugerida: Optional[str] = Field(
        default=None, description="Muy Bueno, Bueno, Regular o Malo"
    )
    fundamentos: List[str] = Field(
        default_factory=list,
        description="hechos concretos en los que se apoya el texto",
    )
