"""Esquema Pydantic del informe agroeconomico Fitogreen / formato Banco Santander.

Los campos salen de lo observado en los tres libros analizados:

  - Formato Fitogreen Bsan.xlsx        -> etiquetas y estructura de la plantilla
  - RER Cliente B Ago-2025.xlsx -> informe R.E.R. real (mismo dominio)
  - Cliente D May-2026.xlsx            -> informe de TASACION (otro documento)

Las secciones marcadas con "solo tasacion" no tienen destino en la plantilla
Bsan; se modelan porque aparecen en los ejemplos reales, pero el mapping
config/mapping_santander.json las deja sin celda de destino.
"""

from __future__ import annotations

import re
from datetime import date
from enum import Enum
from typing import Dict, List, Optional

from pydantic import BaseModel, ConfigDict, Field, field_validator

from services.rut import RutInvalido, formatear_rut, normalizar_rut

from schemas.comun import sin_signo_de_formula

TEMPORADA_RE = re.compile(r"^\d{4}/\d{2}$")

# Serie indexada por temporada agricola: {"2025/26": 12.5, "2026/27": 13.1}
SerieTemporada = Dict[str, Optional[float]]


class Base(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    # Ningun texto del dominio llega al Excel como formula (schemas/comun.py).
    @field_validator("*", mode="after")
    @classmethod
    def _sin_formulas(cls, valor):
        return sin_signo_de_formula(valor)


def _validar_temporadas(v: SerieTemporada) -> SerieTemporada:
    for k in v:
        if not TEMPORADA_RE.match(k):
            raise ValueError(
                "clave de temporada invalida: {!r} (formato esperado '2025/26')".format(k)
            )
    return v


# --------------------------------------------------------------------- enums

class Calificacion(str, Enum):
    MUY_BUENO = "Muy Bueno"
    BUENO = "Bueno"
    REGULAR = "Regular"
    MALO = "Malo"


class Estado(str, Enum):
    BUENO = "Bueno"
    REGULAR = "Regular"
    MALO = "Malo"


class Tenencia(str, Enum):
    PROPIETARIO = "Propietario"
    ARRENDATARIO = "Arrendatario"
    OTROS = "Otros"
    # Las que agrega la lista del formato final de Santander (hoja Opciones).
    SOCIEDAD_RELACIONADA = "Sociedad Relacionada"
    PROPIEDAD_FAMILIAR = "Propiedad Familiar"
    USUFRUCTUARIO = "Usufructuario"


class EstatusProductivo(str, Enum):
    """Lista desplegable de la hoja Plantaciones del formato final."""

    PLENA = "Plena"
    CRECIENTE = "Creciente"
    FORMACION = "Formación"


class ClaseUso(str, Enum):
    """Clases de capacidad de uso del suelo, filas fijas 13-16 de la plantilla."""

    I_II = "I-II"
    III_IV = "III-IV"
    V_VI = "V-VI"
    VII_VIII = "VII-VIII"


class CategoriaFoto(str, Enum):
    EMPLAZAMIENTO = "emplazamiento"      # unica soportada por la plantilla Bsan
    PLANTACION = "plantacion"
    INFRAESTRUCTURA = "infraestructura"
    RIEGO = "riego"
    SATELITAL = "satelital"
    LEGAL = "legal"
    OTRA = "otra"


# ------------------------------------------------- hoja "Cliente" (portada)

class Antecedentes(Base):
    """Bloque ANTECEDENTES GENERALES, filas 10-13 de la hoja Cliente."""

    cliente: str
    rut: str
    ejecutivo: Optional[str] = None
    fecha_visita: Optional[date] = None
    sucursal: Optional[str] = None
    fecha_emision_informe: Optional[date] = None
    agronomo_informante: Optional[str] = None
    n_solicitud: Optional[str] = None

    @field_validator("rut")
    @classmethod
    def _rut_con_dv(cls, v: str) -> str:
        # No solo exige el guion: CALCULA el digito verificador (modulo 11)
        # y compara -asi da lo mismo si lo tipean con puntos, sin puntos o
        # todo junto ('765432103'), y un digito transpuesto no pasa colado.
        try:
            normalizar_rut(v)
        except RutInvalido as error:
            raise ValueError(str(error)) from error
        return formatear_rut(v)


class FilaEstructuraTemporada(Base):
    """Filas 20-23 de Cliente: especie x superficie x tenencia."""

    especie: str
    superficie_ha: float = Field(ge=0)
    #: FRACCION de 0 a 1 (0,8 = 80%), igual que ratio_exportacion_pct: la
    #: celda Cliente!D20 tiene formato de porcentaje y multiplica por 100.
    #: Hasta el 2026-10-05 se declaraba de 0 a 100, y un 80 salia impreso
    #: como «8000,0%» (visto en la vista previa del informe de prueba).
    consolidacion_pct: Optional[float] = Field(default=None, ge=0, le=1)
    tenencia_propio_ha: Optional[float] = Field(default=None, ge=0)
    tenencia_arriendo_ha: Optional[float] = Field(default=None, ge=0)
    tenencia_otros_ha: Optional[float] = Field(default=None, ge=0)


class MontoComparado(Base):
    """Un item del RESUMEN RESULTADOS: ultima temporada vs proyeccion, M$ y UF."""

    ultima_m_pesos: Optional[float] = None
    ultima_uf: Optional[float] = None
    proyeccion_m_pesos: Optional[float] = None
    proyeccion_uf: Optional[float] = None


class ResumenResultados(Base):
    """Filas 30-34 de Cliente."""

    temporada_ultima: str = Field(description="ej. '25/26'")
    temporada_proyeccion: str = Field(description="ej. '26/27'")
    ingresos_operacionales: MontoComparado = MontoComparado()
    costos_directos: MontoComparado = MontoComparado()
    margen_bruto: MontoComparado = MontoComparado()
    costos_indirectos: MontoComparado = MontoComparado()
    resultado_operacional_neto: MontoComparado = MontoComparado()


# --------------------------------------- hoja "Dotacion de recursos"

class SuperficieClaseUso(Base):
    riego_ha: Optional[float] = Field(default=None, ge=0)
    secano_ha: Optional[float] = Field(default=None, ge=0)
    # 'total' es formula en la plantilla (=riego+secano); no se escribe.


class PredioSuelos(Base):
    """Un predio = un bloque de 3 columnas en la tabla SUELOS (filas 12-17).

    La plantilla admite exactamente 4 predios (columnas B-D, E-G, H-J, K-M).
    """

    nombre: str
    superficies: Dict[ClaseUso, SuperficieClaseUso] = Field(default_factory=dict)


class DatosSuelos(Base):
    cliente: Optional[str] = None
    predio: Optional[str] = None
    dominio: Optional[str] = None
    rol_avaluo: Optional[str] = None
    comuna: Optional[str] = None
    inscripcion_dominio: Optional[str] = None
    predios: List[PredioSuelos] = Field(default_factory=list, max_length=4)


class DerechoAgua(Base):
    """Tabla AGUAS DE RIEGO, filas 23-26."""

    predio: str
    comuna: Optional[str] = None
    tipo_fuente: Optional[str] = None
    nombre_fuente: Optional[str] = None
    dominio: Optional[str] = None
    cantidad: Optional[float] = Field(default=None, description="D° o acciones")
    caudal_nominal_ls: Optional[float] = Field(default=None, ge=0)
    caudal_real_ls: Optional[float] = Field(default=None, ge=0)
    cbr_fojas: Optional[str] = None
    cbr_numero: Optional[str] = None
    cbr_anio: Optional[int] = None
    cbr_conservador: Optional[str] = None
    propietario: Optional[str] = None
    ejercicio: Optional[str] = Field(
        default=None,
        description="ejercicio del derecho: 'Permanente y continuo', 'Eventual y "
                    "discontinuo'... (lista de la hoja Opciones del formato final)",
    )


class Lote(Base):
    """Un lote de los DATOS LEGALES TERRENOS del formato final de Santander:
    una fila por rol, agrupados por unidad predial."""

    unidad_predial: str
    nombre: Optional[str] = None
    dominio: Optional[str] = None
    rol: Optional[str] = None
    region: Optional[str] = Field(default=None, description="'V', 'RM', 'XVI'...")
    comuna: Optional[str] = None
    fojas: Optional[str] = None
    numero: Optional[str] = None
    anio: Optional[int] = None
    conservador: Optional[str] = None
    superficie_legal_ha: Optional[float] = Field(default=None, ge=0)
    superficie_real_ha: Optional[float] = Field(default=None, ge=0)


class UnidadPredial(Base):
    """Clasificacion de suelos por unidad predial (formato final): riego solo
    en las clases I-II y III-IV, secano en las cuatro. Y lo que el formato
    pide de la unidad en el requerimiento de riego: sus tranques."""

    nombre: str
    i_ii_riego_ha: Optional[float] = Field(default=None, ge=0)
    iii_iv_riego_ha: Optional[float] = Field(default=None, ge=0)
    i_ii_secano_ha: Optional[float] = Field(default=None, ge=0)
    iii_iv_secano_ha: Optional[float] = Field(default=None, ge=0)
    v_vi_secano_ha: Optional[float] = Field(default=None, ge=0)
    vii_viii_secano_ha: Optional[float] = Field(default=None, ge=0)
    topografia: Optional[str] = None
    capacidad_tranques_m3: Optional[float] = Field(default=None, ge=0)


class RequerimientoRiego(Base):
    """Tabla ESTIMACION REQUERIMIENTO AGUAS DE RIEGO, filas 33-36.

    Los tres parametros del medio -eficiencia, ETo y Kc- salen de la biblioteca
    (`config/parametros_riego.json`), no del tasador. Los dos requerimientos
    son derivados y viajan al Excel como formula, no como numero; se calculan
    igual en Python para poder contrastar.
    """

    especie_variedad: str
    anio_plantacion: Optional[str] = None
    superficie_ha: Optional[float] = Field(default=None, ge=0)
    sistema_riego: Optional[str] = None
    eficiencia_riego: Optional[float] = Field(default=None, ge=0, le=1)
    eto_maximo_mm_enero: Optional[float] = Field(default=None, ge=0)
    coef_cultivo_kc_enero: Optional[float] = Field(default=None, ge=0)
    cobertura: Optional[float] = Field(
        default=None, ge=0, le=1,
        description="fraccion de desarrollo segun los anios desde la plantacion; "
                    "la plantilla del banco no trae esta columna y el generador "
                    "la agrega para que se vea por que un huerto nuevo pide menos",
    )
    requerimiento_max_ls_ha: Optional[float] = Field(default=None, ge=0)
    requerimiento_total_plantacion_ls: Optional[float] = Field(default=None, ge=0)
    densidad_plantacion: Optional[str] = None


class ItemDotacion(Base):
    """Maquinaria o infraestructura, filas 42-45 (tres bloques de 3 columnas)."""

    nombre: str
    cantidad: Optional[float] = Field(default=None, ge=0)
    estado: Optional[Estado] = None


class DotacionRecursos(Base):
    suelos: DatosSuelos = DatosSuelos()
    derechos_agua: List[DerechoAgua] = Field(default_factory=list)
    requerimientos_riego: List[RequerimientoRiego] = Field(default_factory=list)
    maquinaria: List[ItemDotacion] = Field(default_factory=list)
    infraestructura: List[ItemDotacion] = Field(default_factory=list)
    comentarios: Optional[str] = None
    # Formato final de Santander (2026-10-13).
    lotes: List[Lote] = Field(default_factory=list)
    unidades_prediales: List[UnidadPredial] = Field(default_factory=list)
    notas_suelos: Optional[str] = Field(
        default=None, description="el cuadro de notas bajo la clasificacion de suelos")


# ------------------------------------------------------- hoja "Inversion"

class LineaInversion(Base):
    """Filas 5-14 de Inversion. Las columnas G, J son formulas."""

    proyecto: str
    anio: Optional[int] = None
    superficie_ha: Optional[float] = Field(default=None, ge=0)
    item: Optional[str] = None
    cantidad_ha: Optional[float] = None
    costo_unitario_usd: Optional[float] = Field(default=None, ge=0)
    costo_inversion_total_usd: Optional[float] = Field(default=None, ge=0)
    costo_inversion_pesos: Optional[float] = Field(default=None, ge=0)


class Inversiones(Base):
    lineas: List[LineaInversion] = Field(default_factory=list)
    comentarios: Optional[str] = None


# ------------------------------ hoja "Estructura Productiva" (catastro)

class Plantacion(Base):
    """Catastro de plantas. Filas 7-10 en la plantilla; en informes reales
    llega a decenas de filas (Cliente B: 35), asi que hay que insertar filas."""

    especie: str
    variedad: Optional[str] = None
    anio_plantacion: Optional[str] = Field(
        default=None, description="texto: admite rangos como '2012-2013'"
    )
    superficie_ha: float = Field(ge=0)
    sistema_conduccion: Optional[str] = None
    predio: Optional[str] = None
    dominio: Optional[Tenencia] = None
    riego: Optional[str] = None
    estado: Optional[Estado] = None
    densidad_plantacion: Optional[str] = Field(
        default=None, description="texto: '4x1,2', '6 x 2,5'"
    )
    # Formato final de Santander.
    royalty: Optional[bool] = None
    estatus_productivo: Optional[EstatusProductivo] = Field(
        default=None,
        description="si no se declara, el generador lo deduce de la edad con las "
                    "reglas de madurez (config/reglas_agronomicas.json)",
    )
    supuestos: Optional["SupuestoPlantacion"] = Field(
        default=None,
        description="los numeros de la hoja Supuestos. Los propone el sistema "
                    "desde la biblioteca al generar (el agrónomo, 2026-10-14: opcion "
                    "a) y el tasador los corrige en el Excel",
    )


class SupuestoPlantacion(Base):
    """Una fila de la hoja Supuestos del formato final: lo que mueve el bloque
    de flujo de esa plantacion. Rendimientos de los 8 anios de proyeccion."""

    pct_exportacion: Optional[float] = Field(default=None, ge=0, le=1)
    precio_exportacion_usd_kg: Optional[float] = Field(default=None, ge=0)
    precio_exportacion_siguientes_usd_kg: Optional[float] = Field(default=None, ge=0)
    precio_mercado_interno_usd_kg: Optional[float] = Field(default=None, ge=0)
    costo_fijo_usd_ha: Optional[float] = Field(default=None, ge=0)
    costo_variable_usd_kg: Optional[float] = Field(default=None, ge=0)
    rendimientos_kg_ha: List[Optional[float]] = Field(default_factory=list, max_length=8)
    costo_formacion_1_usd_ha: Optional[float] = Field(default=None, ge=0)
    costo_formacion_2_usd_ha: Optional[float] = Field(default=None, ge=0)
    costo_fijo_formacion_3_usd_ha: Optional[float] = Field(default=None, ge=0)
    costo_variable_formacion_3_usd_kg: Optional[float] = Field(default=None, ge=0)
    curva: Optional[str] = Field(default=None, description="de que curva salieron")


class ResumenSuperficieEspecie(Base):
    """Filas 16-19. Las columnas D y E son formulas en la plantilla."""

    especie: str
    pct_por_especie: Optional[float] = Field(
        default=None,
        description="lo escribe el motor como formula viva; el campo"
                    " existe solo para que la columna tenga nombre",
    )
    propietario_ha: Optional[float] = Field(default=None, ge=0)
    arrendatario_ha: Optional[float] = Field(default=None, ge=0)
    plena_produccion_ha: Optional[float] = Field(default=None, ge=0)
    rendimientos_crecientes_ha: Optional[float] = Field(default=None, ge=0)
    formacion_ha: Optional[float] = Field(default=None, ge=0)
    proyectos_plantacion_ha: Optional[float] = Field(default=None, ge=0)
    superficie_total_con_proyectos_ha: Optional[float] = Field(default=None, ge=0)


class SistemaProduccion(Base):
    """Filas 31-35."""

    especie: str
    sistema_produccion: Optional[str] = None
    densidad_plantacion: Optional[str] = None
    sistema_riego: Optional[str] = None
    variedades_con_royalty: Optional[str] = None
    considera_techos: Optional[str] = None
    sistema_contra_heladas: Optional[str] = None
    nivel_tecnologico: Optional[str] = None
    nivel_homogeneidad: Optional[str] = None
    nivel_productivo_promedio: Optional[str] = None
    estado_fitosanitario: Optional[str] = None


class SistemaComercializacion(Base):
    """Filas 39-42."""

    especie: str
    empresas_comercializadoras: Optional[str] = None
    comercializacion: Optional[str] = None
    mercado: Optional[str] = None
    periodo_cosecha: Optional[str] = None
    periodo_liquidaciones: Optional[str] = None
    proceso: Optional[str] = None
    gestion_exportacion: Optional[str] = None
    certificaciones: Optional[str] = None


class EstructuraProductiva(Base):
    anio_catastro: Optional[int] = None
    plantaciones: List[Plantacion] = Field(default_factory=list)
    resumen_superficie: List[ResumenSuperficieEspecie] = Field(default_factory=list)
    comentarios_plantaciones: Optional[str] = None
    variedades_con_royalty: Optional[str] = Field(
        default=None,
        description="variedades con royalty, separadas por coma: el formato final "
                    "de Santander marca Si/No en cada plantacion de esas variedades",
    )
    sistemas_produccion: List[SistemaProduccion] = Field(default_factory=list)
    sistemas_comercializacion: List[SistemaComercializacion] = Field(default_factory=list)


# -------------------------------------------- hoja "Flujo proyectado "

class BloqueEspecieFlujo(Base):
    """Un bloque de assumptions (10 filas). La plantilla tiene 4 bloques:
    filas 10-19, 20-29, 30-39 y 40-49.

    Las seis primeras filas son entrada (superficie, precios, rendimiento,
    ratio de exportacion, costos). Las tres ultimas -produccion e ingresos- son
    DERIVADAS: el Excel las recibe como formula viva, no como numero, porque el
    banco tiene que poder mover un precio y ver el efecto. Aqui se siguen
    calculando en Python, pero para consumo interno -el resumen de la portada
    las suma- y como contraste: si Python y la formula no coinciden, algo esta
    mal. Las formulas viven en ``formulas_de_fila`` del mapping, no aqui.
    """

    especie_variedad_anio: str
    superficie_ha: SerieTemporada = Field(default_factory=dict)
    precio_exportacion_usd_kg: SerieTemporada = Field(default_factory=dict)
    precio_mercado_interno_usd_kg: SerieTemporada = Field(default_factory=dict)
    rendimiento_kg_ha: SerieTemporada = Field(default_factory=dict)
    ratio_exportacion_pct: SerieTemporada = Field(default_factory=dict)
    costos_usd_ha: SerieTemporada = Field(default_factory=dict)
    produccion_total_kg: SerieTemporada = Field(default_factory=dict)
    ingresos_exportacion: SerieTemporada = Field(default_factory=dict)
    ingresos_mercado_interno: SerieTemporada = Field(default_factory=dict)

    @field_validator(
        "superficie_ha",
        "precio_exportacion_usd_kg",
        "precio_mercado_interno_usd_kg",
        "rendimiento_kg_ha",
        "ratio_exportacion_pct",
        "costos_usd_ha",
        "produccion_total_kg",
        "ingresos_exportacion",
        "ingresos_mercado_interno",
    )
    @classmethod
    def _claves_temporada(cls, v: SerieTemporada) -> SerieTemporada:
        return _validar_temporadas(v)


class AjusteDeCurva(Base):
    """Un ajuste que el tasador decide EN TERRENO sobre una curva concreta.

    Es dato de origen A puro: existe porque alguien fue al predio y vio que el
    riego era mediocre o que la poda estaba atrasada. No sale de ninguna tabla
    y no pasa en todos los informes, solo en los que el agronomo observa algo
    que la curva de manual no recoge.

    Va por especie y no por informe entero: que el palto este mal podado no
    dice nada del cerezo de al lado.
    """

    especie: str = Field(
        min_length=2,
        description="la especie o el nombre de la curva a la que se le aplica")
    ajuste_pct: float = Field(
        ge=-100, le=100,
        description="variacion sobre la curva: -2 baja un 2%, +5 sube un 5%")
    motivo: Optional[str] = Field(
        default=None,
        description="lo que se observo. Es lo que se le responde al banco si "
                    "pregunta por que este huerto rinde distinto")


class FilaMatriz(Base):
    """Una edad dentro de la matriz de envejecimiento de un grupo.

    ``superficie_por_temporada`` es lo que envejece: la misma hectarea aparece
    en la edad 7 una temporada y en la 8 la siguiente.
    """

    edad: int = Field(ge=0, le=120)
    rendimiento_curva_kg_ha: float = Field(ge=0)
    costo_base_usd_ha: float = Field(default=0.0, ge=0)
    superficie_por_temporada: SerieTemporada = Field(default_factory=dict)


class BloqueMatriz(Base):
    """La cartera de un bloque del flujo, abierta por edad.

    Va a una hoja anexa para que el flujo la referencie con SUMPRODUCT en vez
    de recibir un promedio ya calculado en Python: asi el banco ve de donde
    sale cada rendimiento y puede mover una superficie.
    """

    etiqueta: str
    curva: str
    factor_ajuste: float = Field(default=1.0, gt=0)
    costo_por_kg_usd: float = Field(default=0.0, ge=0)
    filas: List[FilaMatriz] = Field(default_factory=list)


class MatrizEnvejecimiento(Base):
    """Evidencia de como se armo el flujo, temporada a temporada."""

    temporadas: List[str] = Field(default_factory=list)
    bloques: List[BloqueMatriz] = Field(default_factory=list)
    avisos: List[str] = Field(default_factory=list)
    superficie_excluida_ha: float = 0.0


class FlujoProyectado(Base):
    temporadas: List[str] = Field(
        default_factory=list,
        description="orden de columnas D..N, ej. ['2023/24', ..., '2033/34']",
    )
    tipo_cambio_pesos_usd: SerieTemporada = Field(default_factory=dict)
    uf: Optional[float] = Field(
        default=None,
        description="valor de la UF del dia en que se emite el informe. Lo "
                    "declara el tasador: la caratula pide las cifras en M$ y "
                    "en UF, y sin este numero la columna UF no se puede "
                    "calcular. Es del DIA DEL INFORME, no de la temporada",
    )
    bloques: List[BloqueEspecieFlujo] = Field(default_factory=list, max_length=4)
    administracion: SerieTemporada = Field(default_factory=dict)
    gastos_generales: SerieTemporada = Field(default_factory=dict)
    ventas_respaldadas_pesos: SerieTemporada = Field(
        default_factory=dict,
        description="ventas segun las declaraciones de IVA, en PESOS, por "
                    "temporada. Es lo que el cliente entrega y con lo que el "
                    "banco contrasta lo declarado. En el Excel la fila pide "
                    "dolares, asi que se escribe como formula dividiendo por el "
                    "tipo de cambio: el numero en pesos queda a la vista y el "
                    "cambio sigue siendo una celda viva. Solo existe para las "
                    "temporadas declaradas: una proyeccion no tiene IVAs",
    )
    porcentaje_respaldo: SerieTemporada = Field(
        default_factory=dict,
        description="DERIVADO. No se teclea: el motor lo escribe como formula "
                    "viva (respaldado / ingresos totales) en las columnas donde "
                    "hay monto respaldado. Se conserva el campo para poder "
                    "contrastar el calculo de Python contra el de Excel",
    )
    comentarios: Optional[str] = None
    clase_de_temporada: Dict[str, str] = Field(
        default_factory=dict,
        description="como se rotula cada temporada: 'Real ', 'Preliminar' o "
                    "'Projection '. La plantilla las trae escritas a mano y "
                    "eso solo sirve para el ano en que se escribieron",
    )
    matriz: Optional[MatrizEnvejecimiento] = Field(
        default=None,
        description="cartera abierta por edad; alimenta la hoja anexa y las "
                    "formulas SUMPRODUCT del flujo",
    )

    @field_validator("temporadas")
    @classmethod
    def _formato_temporadas(cls, v: List[str]) -> List[str]:
        for t in v:
            if not TEMPORADA_RE.match(t):
                raise ValueError("temporada invalida: {!r}".format(t))
        return v


# ------------------------------ hoja "Comentarios generales Opinion"

class BloqueOpinion(Base):
    texto: Optional[str] = None
    calificacion: Optional[Calificacion] = None


class OpinionProfesional(Base):
    gestion_administrativa_productiva: BloqueOpinion = BloqueOpinion()
    canales_comerciales: BloqueOpinion = BloqueOpinion()
    calidad_activos_productivos: BloqueOpinion = BloqueOpinion()
    conclusion_final: BloqueOpinion = BloqueOpinion()


# ------------------------------------------------------------ hoja "Anexo "

class FilaAnexo(Base):
    """Las columnas F, I y M son formulas en la plantilla."""

    especie: str
    superficie_ha: Optional[float] = Field(
        default=None, ge=0,
        description="superficie que produjo estos resultados. La plantilla del "
                    "banco no la trae y la pidio el agrónomo, porque sin ella seis "
                    "filas del resumen del flujo no se pueden calcular: el "
                    "rendimiento, el costo por hectarea y el costo unitario "
                    "salen todos de dividir por la superficie",
    )
    produccion_exportacion_kg: Optional[float] = Field(default=None, ge=0)
    produccion_mercado_interno_kg: Optional[float] = Field(default=None, ge=0)
    ventas_exportacion_pesos: Optional[float] = None
    ventas_mercado_interno_pesos: Optional[float] = None
    total_ventas_usd: Optional[float] = None
    costos_directos_pesos: Optional[float] = None
    costos_directos_usd: Optional[float] = None


class ResultadoTemporada(Base):
    temporada: Optional[str] = Field(
        default=None,
        description="a que temporada corresponden estos resultados, ej "
                    "'2024/25'. Lo declara el tasador porque NO siempre se lo "
                    "dan: a veces el cliente entrega dos anios hacia atras, a "
                    "veces uno y a veces ninguno. Sin esto habria que suponer "
                    "cual es, y suponer el ano de un resultado real es "
                    "exactamente lo que no debe hacer un informe al banco",
    )
    filas: List[FilaAnexo] = Field(default_factory=list, max_length=4)
    administracion_pesos: Optional[float] = None
    gastos_generales_pesos: Optional[float] = None


class Anexo(Base):
    resultado_real: ResultadoTemporada = ResultadoTemporada()
    resultado_preliminar: ResultadoTemporada = ResultadoTemporada()
    # Formato final de Santander: Anexo 2, ventas del F29 de julio a junio.
    ventas_f29_pesos: List[Optional[float]] = Field(default_factory=list, max_length=12)
    nota_f29: Optional[str] = None


class Sensibilizacion(Base):
    """Los cuatro Delta % de una especie en la hoja Supuestos (0,1 = +10%)."""

    precio_exportacion: float = 0.0
    precio_mercado_interno: float = 0.0
    rendimiento: float = 0.0
    costos: float = 0.0


class Supuestos(Base):
    """Lo que la hoja Supuestos lleva por especie y para todo el informe; lo
    de cada plantacion va en ``Plantacion.supuestos``."""

    sensibilizacion: Dict[str, Sensibilizacion] = Field(default_factory=dict)
    administracion_usd_anio: Optional[float] = Field(default=None, ge=0)
    notas: Dict[str, str] = Field(
        default_factory=dict, description="la nota bajo la tabla de cada especie")


# ------------------------------------------------------------- fotografias

class Fotografia(Base):
    """Imagen flotante a insertar.

    'celda_destino' se resuelve normalmente desde config/mapping_santander.json;
    solo se fija aqui cuando se quiere forzar una posicion concreta.
    """

    ruta: str
    titulo: Optional[str] = None
    categoria: CategoriaFoto = CategoriaFoto.OTRA
    celda_destino: Optional[str] = None
    ancho_px: Optional[int] = Field(default=None, gt=0)
    alto_px: Optional[int] = Field(default=None, gt=0)

    @field_validator("celda_destino")
    @classmethod
    def _celda_valida(cls, v: Optional[str]) -> Optional[str]:
        if v is not None and not re.match(r"^[A-Z]{1,3}\d+$", v):
            raise ValueError("celda invalida: {!r} (formato esperado 'B12')".format(v))
        return v


# ----------------------------------------- solo tasacion (Cliente D)

class Construccion(Base):
    """Tabla 5.1 CONSTRUCCIONES del informe de tasacion.

    No existe destino para esta tabla en la plantilla Bsan: su unica tabla de
    infraestructura es (nombre, cantidad, estado). Se conserva para el flujo de
    tasacion, no para el informe agroeconomico.
    """

    numero: Optional[int] = None
    edificacion: str
    estado_conservacion: Optional[Estado] = None
    ubicacion_rol: Optional[str] = None
    superficie_m2: Optional[float] = Field(default=None, ge=0)
    estructura_vertical: Optional[str] = None
    data_aproximada: Optional[int] = None
    estructura_techumbre: Optional[str] = None
    pavimentos: Optional[str] = None
    cubierta: Optional[str] = None
    revestimiento_vertical: Optional[str] = None
    bien_asegurable: Optional[bool] = None


# ------------------------------------------------------------------ raiz

class InformeFitogreen(Base):
    """Documento completo. Solo 'antecedentes' es obligatorio."""

    antecedentes: Antecedentes
    estructura_temporada: List[FilaEstructuraTemporada] = Field(default_factory=list)
    resumen_resultados: Optional[ResumenResultados] = None
    dotacion_recursos: DotacionRecursos = DotacionRecursos()
    inversiones: Inversiones = Inversiones()
    estructura_productiva: EstructuraProductiva = EstructuraProductiva()
    flujo_proyectado: FlujoProyectado = FlujoProyectado()
    opinion: OpinionProfesional = OpinionProfesional()
    anexo: Anexo = Anexo()
    fotografias: List[Fotografia] = Field(default_factory=list)
    ajustes_de_curva: List[AjusteDeCurva] = Field(
        default_factory=list,
        description="lo que el tasador observo en terreno y corrige sobre las "
                    "curvas de manual; se captura en la pantalla de visita",
    )
    construcciones: List[Construccion] = Field(
        default_factory=list,
        description="solo tasacion; sin destino en la plantilla Bsan",
    )
    supuestos: Supuestos = Supuestos()


Plantacion.model_rebuild()
