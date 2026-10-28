"""Esquema Pydantic de la retasacion Rabobank/RaboFinance.

Dominio DISTINTO al de schemas/informe.py (Bsan mide capacidad de pago con
un flujo proyectado; esto mide valor de garantia con hipoteca+prenda) -
"no mezclar los motores", CLAUDE.md.

Los campos salen de medir celda por celda el archivo real
`Tasacion_Base_Rabo.xlsx` (informe de retasacion de RaboFinance, entregado),
no de una lectura general de la plantilla. Las capacidades de fila (5
construcciones hipotecables, 6 plantaciones sin royalty, etc.) son las que
trae la plantilla del banco tal como esta hoy -ampliarlas exige insertar
filas, igual que Bsan (services/excel_ops.py:insertar_filas).

Fase 1: **una sola propiedad por retasacion**. La plantilla real no repite
el bloque de RESUMEN VALOR para mas de un predio (medido, no supuesto);
varios predios queda para una fase futura.

**Ningun dato impide generar** (2026-10-02). el usuario: *"podemos generar el
informe sin pedir todo pero advirtiendo que cosas le faltan?"*. Los campos
que antes eran obligatorios en el sentido de Pydantic -sin ellos el
servidor rechazaba el informe entero- se declaran con `Obligatorio(...)`:
siguen siendo exactamente los mismos, pero su ausencia se AVISA en vez de
bloquear (services/retasacion_pendientes.py lee esta marca). La lista de
que es obligatorio vive aqui, una sola vez, y no repetida en la pantalla.
"""

from __future__ import annotations

from datetime import date
from enum import Enum
from typing import Any, List, Optional

from pydantic import BaseModel, ConfigDict, Field, field_validator

from schemas.informe import Estado
from schemas.comun import sin_signo_de_formula
from services.rut import RutInvalido, formatear_rut, normalizar_rut


class Base(BaseModel):
    # coerce_numbers_to_str: Excel guarda las fojas y el numero de una
    # inscripcion como NUMERO (1234), y el dominio los quiere como texto
    # -"siempre en cifras", pero texto-. Sin esto, leer un informe anterior
    # los descartaba como "mal escritos" (medido 2026-10-03 regenerando el
    # archivo real del banco).
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True,
                              coerce_numbers_to_str=True)

    # Ningun texto del dominio llega al Excel como formula (schemas/comun.py).
    @field_validator("*", mode="after")
    @classmethod
    def _sin_formulas(cls, valor):
        return sin_signo_de_formula(valor)


def Obligatorio(etiqueta: str, consecuencia: Optional[str] = None,
                **restricciones: Any) -> Any:
    """Un dato que el informe necesita, pero cuya ausencia no impide generarlo.

    Queda vacio en el Excel y aparece en la lista de pendientes con
    `etiqueta`. `consecuencia` es para los pocos datos cuya falta no deja
    solo una celda vacia sino que arrastra otras -la UF, el dolar, la
    fecha-: dice que se descalza y donde escribirlo en el Excel. Las
    `restricciones` (gt=0...) siguen valiendo para un valor que SI vino.
    """
    extra = {"obligatorio": etiqueta}
    if consecuencia:
        extra["consecuencia"] = consecuencia
    return Field(default=None, title=etiqueta, json_schema_extra=extra,
                 **restricciones)


def _validar_rut(v: Optional[str]) -> Optional[str]:
    # Calcula el digito verificador (modulo 11) en vez de solo pedir un
    # guion -mismo criterio que Antecedentes.rut en schemas/informe.py.
    if not v:
        return v
    try:
        normalizar_rut(v)
    except RutInvalido as error:
        raise ValueError(str(error)) from error
    return formatear_rut(v)


# --------------------------------------------------------------------- 1-2

class InscripcionCBR(Base):
    """Fojas/numero/ano/CBR — igual convencion que schemas.extraccion: SIEMPRE
    en cifras, nunca en palabras."""

    fojas: Optional[str] = Field(default=None, title="Fojas")
    numero: Optional[str] = Field(default=None, title="Número de inscripción")
    anio: Optional[int] = Field(default=None, title="Año de inscripción")
    conservador: Optional[str] = Field(default=None, title="Conservador")


class Cliente(Base):
    """PORTADA!C14/H14 (cliente) y C15/H15 (propietario)."""

    nombre: Optional[str] = Obligatorio("Nombre del cliente")
    rut: Optional[str] = Obligatorio("RUT del cliente")
    propietario: Optional[str] = Field(default=None, title="Propietario")
    rut_propietario: Optional[str] = Field(default=None, title="RUT del propietario")

    _v1 = field_validator("rut")(_validar_rut)
    _v2 = field_validator("rut_propietario")(_validar_rut)


class Predio(Base):
    """LEGAL 2.1 (fila unica en esta fase) + 2.2 antecedentes expropiatorios."""

    nombre: Optional[str] = Obligatorio("Nombre del predio")
    #: LEGAL!B24 (tabla de expropiacion): el formato real usa ahi un nombre
    #: corto ("Parcela Nº2") en vez del legal completo de B6.
    #: Si no se escribe, va el nombre completo.
    nombre_corto: Optional[str] = Field(default=None, title="Nombre corto del predio")
    rol_avaluo: Optional[str] = Obligatorio("Rol de avalúo")
    # No es Obligatorio aunque antes era requerida: la pantalla la llena con
    # la MISMA comuna de los parametros, y ninguna celda del mapping la
    # escribe -marcarla haria salir dos veces el mismo aviso-.
    comuna: Optional[str] = Field(default=None, title="Comuna del predio")
    superficie_legal_ha: Optional[float] = Obligatorio("Superficie legal (ha)", ge=0)
    superficie_real_ha: Optional[float] = Field(
        default=None, ge=0, title="Superficie real (ha)")
    inscripcion: Optional[InscripcionCBR] = Field(default=None, title="Inscripción CBR")
    # LEGAL 2.1, columnas J y L (2026-10-03: estaban en el formato real y el
    # actualizador no las escribia).
    hipoteca_rabofinance: Optional[bool] = Field(
        default=None, title="Hipoteca a favor de Rabofinance")
    deuda_contribuciones_clp: Optional[float] = Field(
        default=None, ge=0, title="Deuda de contribuciones ($)")
    # LEGAL 2.3 (C35): los cuatro van en una sola celda, "Norte: ...".
    deslinde_norte: Optional[str] = None
    deslinde_sur: Optional[str] = None
    deslinde_oriente: Optional[str] = None
    deslinde_poniente: Optional[str] = None
    afecto_expropiacion: Optional[bool] = None
    superficie_expropiada: Optional[str] = Field(
        default=None, description="texto libre: la plantilla real trae "
                                   "'No senala' en vez de un numero")
    descripcion_expropiacion: Optional[str] = None


class ParametrosRetasacion(Base):
    """PORTADA: cabecera de la tasacion."""

    # Las consecuencias estan medidas sobre las formulas de la plantilla
    # (2026-10-02) y comprobadas recalculando con Excel real
    # (scripts/test_retasacion.py, seccion de pendientes).
    fecha_retasacion: Optional[date] = Obligatorio(
        "Fecha de la retasación",
        consecuencia="sin ella, la edad de las plantaciones y la vida "
                     "remanente de las instalaciones se calculan mal. Se "
                     "puede escribir después en PORTADA!H7")
    uf_base: Optional[float] = Obligatorio(
        "Valor de la UF",
        consecuencia="sin ella, el valor de los suelos sale en $0 y las "
                     "columnas en UF salen en error. Se puede escribir "
                     "después en PORTADA!H8", gt=0)
    valor_usd: Optional[float] = Obligatorio(
        "Valor del dólar",
        consecuencia="sin él, las plantaciones sin royalty salen valoradas "
                     "en $0 y las columnas en dólares salen en error. Se "
                     "puede escribir después en PORTADA!H9", gt=0)
    n_solicitud: Optional[str] = Field(default=None, description="ej. VR-240097")
    ejecutivo: Optional[str] = None
    region: Optional[str] = Obligatorio("Región")
    comuna: Optional[str] = Obligatorio("Comuna")
    tipo_informe: str = "Retasación"


class Tasador(Base):
    """RESUMEN VALOR!B24/B25/B26."""

    nombre: Optional[str] = Obligatorio("Nombre del tasador")
    rut: Optional[str] = Obligatorio("RUT del tasador")
    empresa: str = "Fitogreen Consultora"

    _v = field_validator("rut")(_validar_rut)


# ------------------------------------------------------------------- suelos

class FilaSuelo(Base):
    """SUELOS!3.2, filas 46-53 (capacidad medida: 8 filas)."""

    #: Columna "Propiedad". Libre en el formato real ("Parcela Nº2 Santa
    #: Cecilia (ZEUES)", "Afectación a utilidad pública"); si no se escribe,
    #: va el nombre del predio.
    nombre: Optional[str] = Field(default=None, title="Propiedad")
    # texto libre en la plantilla real, ej. 'ZEUES' — no hay lista fija
    clase_uso: Optional[str] = Obligatorio("Clase de uso")
    grado_habilitacion: Optional[str] = Field(
        default=None, description="'Completa', 'parcial' o 'sin habilitar' "
                                   "(las 3 opciones que trae el desplegable real)")
    superficie_m2: Optional[float] = Obligatorio("Superficie (m²)", ge=0)
    valor_unitario_uf_m2: Optional[float] = Obligatorio("Valor unitario (UF/m²)", ge=0)


class DescripcionSuelos(Base):
    """SUELOS!3.1: coordenadas, acceso."""

    tipo_propiedad: str = "AGRÍCOLA"
    #: El cuadro de texto de 3.1 (ubicacion, acceso, topografia, suelos).
    descripcion: Optional[str] = Field(default=None, title="Descripción general")
    coordenada_e: Optional[str] = None
    coordenada_s: Optional[str] = None
    ciudad_cercana: Optional[str] = None
    distancia_km: Optional[str] = None
    ruta_acceso: Optional[str] = None
    tipo_via: Optional[str] = None


# --------------------------------------------------------------------- agua

class DerechoAgua(Base):
    """Une AGUA!4.1 (inscripcion) y 4.2 (valor): en el archivo real son la
    MISMA lista de derechos vista desde dos tablas -se modela una vez."""

    # canal, rio, pozo, vertiente
    fuente: Optional[str] = Obligatorio("Fuente")
    #: 4.1 (inscripcion) puede nombrar la fuente distinto que 4.2 (valor):
    #: en el formato real, "Canal La Ballica" -el derivado por donde llega el
    #: agua- y "Canal La Compañía" -la asociacion-. Si no se escribe, va la
    #: misma fuente en las dos tablas.
    fuente_inscripcion: Optional[str] = Field(default=None, title="Fuente (inscripción)")
    rol_avaluo: Optional[str] = None
    derechos_cantidad: Optional[float] = Obligatorio("Cantidad de derechos", gt=0)
    tipo: str = Field(default="acción", description="'acción' o 'regador', "
                                                      "el texto libre real")
    #: 4.1, columna F: "Inscrito en CPA" (Catastro Publico de Aguas).
    inscrito_cpa: Optional[bool] = Field(default=None, title="Inscrito en el CPA")
    hipotecado_rabofinance: Optional[bool] = None
    inscripcion: Optional[InscripcionCBR] = Field(default=None, title="Inscripción CBR")
    #: 4.2, columna F: "Derechos expresados en acción". Si no se declara y el
    #: tipo es accion, es la misma cantidad.
    derechos_en_acciones: Optional[float] = Field(
        default=None, ge=0, title="Derechos expresados en acción")
    #: Litros por segundo de cada accion segun la administracion del canal.
    #: Si se declara, el caudal se escribe como formula viva (=D25*1,5, como
    #: en el formato real) en vez de como numero.
    equivalencia_ls_accion: Optional[float] = Field(
        default=None, gt=0, title="Equivalencia (l/s por acción)")
    caudal_nominal_ls: Optional[float] = Field(
        default=None, ge=0, title="Caudal (l/s)",
        description="si no se declara, se calcula con la equivalencia")
    valor_unitario_clp: Optional[float] = Field(
        default=None, ge=0, title="Valor ($/l/s)",
        description="$ por l/s de caudal nominal")


# --------------------------------------------------------- construcciones

class Construccion(Base):
    """CONSTR!5.1 (hipotecables, cap. 5) y 5.2 (prendables, cap. 2) — mismas
    columnas en las dos sub-tablas."""

    edificacion: Optional[str] = Obligatorio("Edificación")
    estado_conservacion: Optional[Estado] = Field(default=None, title="Estado")
    ubicacion_rol: Optional[str] = None
    superficie_m2: Optional[float] = Obligatorio("Superficie (m²)", ge=0)
    estructura_vertical: Optional[str] = None
    anio_dato_aprox: Optional[int] = Field(
        default=None, title="Año aproximado",
        description="año de construcción aproximado -la "
                    "plantilla CALCULA la vida remanente a "
                    "partir de este dato, no se declara aparte")
    costo_reposicion_uf_m2: Optional[float] = Obligatorio(
        "Costo de reposición (UF/m²)", ge=0)
    vida_util_total_anios: Optional[int] = Obligatorio("Vida útil (años)", gt=0)
    vida_remanente_anios: Optional[int] = Field(
        default=None, ge=0,
        description="informativo: la plantilla real lo CALCULA sola "
                    "(anio_dato_aprox + vida_util_total_anios - año del "
                    "informe) — no tiene celda de entrada, es una fórmula")
    valor_residual_pct: float = Field(default=0, ge=0, le=1, title="Valor residual")
    asegurable: bool = True


class Instalacion(Base):
    """CONSTR!5.3 (cap. medida: 1 fila en el ejemplo real, sin tope fijo
    declarado por la plantilla — se trata como ampliable)."""

    # ej. 'Riego tecnificado por goteo', 'Tranque', 'Torre antiheladas', 'Parronal'
    item: Optional[str] = Obligatorio("Ítem")
    # Texto libre, no la lista Bueno/Regular/Malo de las construcciones: el
    # formato real escribe "Operativo" para un riego (y esa celda no tiene
    # lista desplegable en la plantilla).
    estado_conservacion: Optional[str] = Field(default=None, title="Estado")
    ubicacion_rol: Optional[str] = None
    cantidad: Optional[float] = Obligatorio("Cantidad", gt=0)
    # ha, m3, unidad, kva...
    unidad: Optional[str] = Obligatorio("Unidad")
    descripcion: Optional[str] = None
    anio_dato_aprox: Optional[int] = Field(default=None, title="Año aproximado")
    costo_reposicion_unitario_uf: Optional[float] = Obligatorio(
        "Costo de reposición (UF)", ge=0)
    vida_util_total_anios: Optional[int] = Obligatorio("Vida útil (años)", gt=0)
    vida_remanente_anios: Optional[int] = Field(
        default=None, ge=0, description="informativo: la plantilla lo calcula sola")
    valor_residual_pct: float = Field(default=0, ge=0, le=1, title="Valor residual")
    asegurable: bool = True


# ---------------------------------------------------------- plantaciones

class PlantacionSinRoyalty(Base):
    """PLANTACION!6.1, capacidad medida: 6 filas.

    El valor $/ha lo calcula PLANT (costo de reposicion depreciado por
    edad, medido celda por celda en services/excel_engine_rabobank.py) via
    VLOOKUP contra dos tablas de la propia plantilla:

    - ``especie`` tiene que calzar con PLANT!Tabla12 (Cerezo, Ciruelo,
      Avellano, Manzano, Peral, Nogal, Vid Vinifera, Kiwi, Palto
      convencional, Palto Clonal, Uva de mesa). VLOOKUP de Excel no
      distingue mayusculas/minusculas, pero SI la palabra exacta -una
      especie que no calce no da error, da 0 en silencio
      (``IFERROR(...,)`` en la formula del banco), asi que el motor valida
      contra esta lista antes de escribir.
    - ``estructura`` tiene que calzar con PLANT!Tabla5 (Parron estandar,
      Open Gable, Espaldera, Alambrado Tutor, No).
    """

    especie: Optional[str] = Obligatorio("Especie")
    variedad: Optional[str] = Obligatorio("Variedad")
    anio_plantacion: Optional[int] = Obligatorio("Año de plantación")
    densidad_plantas_ha: Optional[int] = Obligatorio("Plantas por hectárea", gt=0)
    estructura: str = Field(
        default="No",
        description="'Parron estandar', 'Open Gable', 'Espaldera', "
                    "'Alambrado Tutor' o 'No' -clave de PLANT!Tabla5-")
    conduccion: Optional[str] = None
    riego: Optional[str] = None
    rol_avaluo: Optional[str] = None
    superficie_ha: Optional[float] = Obligatorio("Superficie (ha)", gt=0)
    estado: Estado = Field(default=Estado.REGULAR, title="Estado")
    # PLANT calcula los años de OPEX y el % de castigo por edad; el formato
    # real los pisa a mano cuando el tasador lo decide (un Merlot de 1980
    # con 70% de castigo y 0 años de OPEX). Vacios, manda la formula.
    opex_anios: Optional[float] = Field(
        default=None, ge=0, title="Años de OPEX (si se fija a mano)")
    castigo_pct: Optional[float] = Field(
        default=None, ge=0, le=1, title="Castigo (si se fija a mano, 0 a 1)")


class PlantacionConRoyalty(Base):
    """PLANTACION!6.2, capacidad medida: 2 filas. Misma forma que sin
    royalty, salvo que declara MARCO de plantacion en vez de densidad.

    **No pasa por PLANT** -medido: en el archivo real, G15/G16 (valor
    unitario) no traen formula ninguna, a diferencia de G5:G10. El costo de
    reposicion generico no aplica a una plantacion con derechos de royalty
    (paga licencia aparte); el valor $/ha se declara directo."""

    especie: Optional[str] = Obligatorio("Especie")
    variedad: Optional[str] = Obligatorio("Variedad")
    anio_plantacion: Optional[int] = Obligatorio("Año de plantación")
    # ej. '4x2', literal como en Bsan
    marco_plantacion: Optional[str] = Obligatorio("Marco de plantación")
    valor_unitario_clp_ha: Optional[float] = Obligatorio("Valor ($/ha)", gt=0)
    rol_avaluo: Optional[str] = None
    superficie_ha: Optional[float] = Obligatorio("Superficie (ha)", gt=0)
    estado: Estado = Field(default=Estado.REGULAR, title="Estado")


# --------------------------------------------------------- testigos/mercado

class TestigoPropiedad(Base):
    """REF. PROPIEDADES, una fila de comparable de suelo."""

    ubicacion: Optional[str] = Obligatorio("Ubicación")
    region: Optional[str] = None
    comuna: Optional[str] = None
    superficie_m2: Optional[float] = Obligatorio("Superficie (m²)", gt=0)
    valor_total_uf: Optional[float] = Obligatorio("Valor total (UF)", gt=0)
    fecha: Optional[date] = Field(default=None, title="Fecha")
    tipo_referencia: Optional[str] = Field(
        default=None, description="'oferta' u otro — decide el ajuste 0.85 "
                                   "de oferta a venta que ya trae la plantilla")
    link: Optional[str] = None
    descripcion: Optional[str] = Field(default=None, title="Breve descripción")


class TestigoCanal(Base):
    """REF. Canal, una fila de comparable de derechos de agua."""

    region: Optional[str] = None
    comuna: Optional[str] = None
    sub_cuenca: Optional[str] = Field(default=None, title="Sub-cuenca")
    #: Canal, Pozo, Río, Estero, Embalse o Tranque (lista de la plantilla).
    tipo_fuente: Optional[str] = Field(default=None, title="Tipo de fuente")
    #: El NOMBRE de la fuente (columna "Nombre de la Fuente").
    fuente: Optional[str] = Obligatorio("Fuente")
    derechos: Optional[float] = Field(default=None, ge=0, title="Derechos")
    tipo: str = "acción"
    caudal_nominal_ls: Optional[float] = Obligatorio("Caudal (l/s)", gt=0)
    valor_total_m_pesos: Optional[float] = Field(default=None, ge=0, title="Valor (M$)")
    valor_unitario_uf_ls: Optional[float] = Field(
        default=None, ge=0, title="Valor unitario (UF/l/s)")
    fecha: Optional[date] = Field(default=None, title="Fecha")
    tipo_referencia: Optional[str] = Field(
        default=None, title="Tipo de referencia",
        description="'Oferta' o 'Venta': la plantilla ajusta las ofertas por 0,85")
    link: Optional[str] = None


# ---------------------------------------------------- diferencias vs. anterior

class DiferenciasTasacionAnterior(Base):
    """Un parrafo de texto libre por seccion — es literalmente lo que trae
    la plantilla real (cajas de texto, no una tabla estructurada). El
    ANTES/DESPUES numerico para armar este texto vive en la pantalla
    (Paso 2), comparando lo que trajo el informe de origen contra lo que
    el tasador tipea ahora; aca solo llega la redaccion final, igual que
    RedaccionTecnica en el flujo de Bsan."""

    suelos: Optional[str] = None
    aguas: Optional[str] = None
    construcciones: Optional[str] = None
    instalaciones: Optional[str] = None
    plantaciones: Optional[str] = None


# --------------------------------------------------------------------- raiz

class RetasacionRabobank(Base):
    """El informe completo. `informe_origen_id` liga a la retasacion anterior
    -del propio expediente en la base, o de un archivo subido (Excel o PDF)
    que quedo registrado por separado-, para poder auditar de donde salio
    cada diferencia."""

    informe_origen_id: Optional[str] = Field(
        default=None,
        description="expediente en la base, o null si el informe anterior "
                    "solo vive en el archivo subido (no hay base de datos "
                    "o el tasador no lo guardo como expediente propio)",
    )

    # Los `title` de cada seccion son los rotulos con que la lista de
    # pendientes nombra donde falta algo -los mismos de la pantalla-.
    parametros: ParametrosRetasacion = Field(
        default_factory=ParametrosRetasacion, title="Datos de la tasación")
    cliente: Cliente = Field(default_factory=Cliente, title="Cliente")
    predio: Predio = Field(default_factory=Predio, title="Predio")
    #: LEGAL!K6 "Tasación anterior": en una retasacion, por definicion, si.
    tiene_tasacion_anterior: bool = Field(default=True, title="Tasación anterior")

    # Los cuadros de texto del formato (2026-10-03). Son los que el banco lee
    # como opinion del tasador; antes no habia donde escribirlos.
    observaciones_portada: Optional[str] = Field(
        default=None, title="Observaciones y advertencias (portada)")
    legal_observaciones: Optional[str] = Field(
        default=None, title="Observaciones de los antecedentes legales")
    legal_observaciones_expropiacion: Optional[str] = Field(
        default=None, title="Observaciones de los antecedentes expropiatorios")

    suelos: List[FilaSuelo] = Field(
        default_factory=list, max_length=8, title="Suelos")
    suelos_descripcion: DescripcionSuelos = Field(
        default_factory=DescripcionSuelos, title="Ubicación y acceso")

    aguas: List[DerechoAgua] = Field(
        default_factory=list, max_length=5, title="Derechos de agua")
    aguas_observaciones: Optional[str] = None
    aguas_suficiencia: Optional[str] = None

    construcciones_hipotecables: List[Construccion] = Field(
        default_factory=list, max_length=5, title="Construcciones hipotecables")
    construcciones_prendables: List[Construccion] = Field(
        default_factory=list, max_length=2, title="Construcciones prendables")
    instalaciones: List[Instalacion] = Field(
        default_factory=list, title="Instalaciones")
    construcciones_observaciones: Optional[str] = None

    plantaciones_sin_royalty: List[PlantacionSinRoyalty] = Field(
        default_factory=list, max_length=6, title="Plantaciones sin royalty")
    plantaciones_con_royalty: List[PlantacionConRoyalty] = Field(
        default_factory=list, max_length=2, title="Plantaciones con royalty")
    plantaciones_observaciones: Optional[str] = None

    # El criterio de valoracion: cuando los suelos se valoran por su mejor
    # uso (extension urbana, por ejemplo), las plantaciones y las
    # instalaciones se tasan pero NO suman al valor de la propiedad. El
    # formato real lo resuelve dejando esas lineas del resumen en blanco
    # -y la §16.a.centies las tomo por un error del banco y las conecto
    # siempre-. Es una decision del tasador, informe por informe.
    valorar_plantaciones: bool = Field(
        default=True, title="Considerar las plantaciones en el valor de la propiedad")
    valorar_instalaciones: bool = Field(
        default=True, title="Considerar las instalaciones en el valor de la propiedad")

    # Celdas EDITABLES en la plantilla real (RESUMEN VALOR!D6..D16), no
    # constantes de negocio — se declaran con el default que trae el banco.
    factor_liquidacion_hipoteca: float = Field(
        default=0.7, gt=0, le=1, title="Factor de liquidación de hipoteca")
    factor_liquidacion_prenda: float = Field(
        default=0.5, gt=0, le=1, title="Factor de liquidación de prenda")

    tasador: Tasador = Field(default_factory=Tasador, title="Tasador")

    diferencias: DiferenciasTasacionAnterior = Field(
        default_factory=DiferenciasTasacionAnterior,
        title="Diferencias con la tasación anterior")

    testigos_propiedades: List[TestigoPropiedad] = Field(
        default_factory=list, title="Testigos de mercado: propiedades")
    testigos_derechos_agua: List[TestigoCanal] = Field(
        default_factory=list, title="Testigos de mercado: derechos de agua")
    #: El analisis escrito que acompaña a la estadistica de cada hoja de
    #: referencias ("I. Tamaño y caracteristicas de la muestra...").
    testigos_propiedades_analisis: Optional[str] = Field(
        default=None, title="Análisis de las referencias de propiedades")
    testigos_derechos_agua_analisis: Optional[str] = Field(
        default=None, title="Análisis de las referencias de derechos de agua")
