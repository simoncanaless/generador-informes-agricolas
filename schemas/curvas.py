"""Curvas de produccion y ajustes del tasador.

Una curva describe como rinde una especie a lo largo de su vida: cuanto produce
por hectarea a cada edad, a que precio vende y cuanto cuesta mantenerla. Es el
"manual" de la especie.

Pero un huerto concreto nunca calza con el manual. Esta mas a la costa y acumula
menos calor, lo administra alguien que no riega bien, el suelo es mas delgado.
Para eso estan los ajustes: factores con nombre que el tasador aplica sobre la
curva base, dejando por escrito por que lo hizo.

Estructura tomada de 'HdU - Plantation Dec-2025 vf Esp.xlsx', hoja
'CURVA PX-RNP', apartados 8.1 a 8.4.
"""

from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Any, Dict, List, Optional
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class Base(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)


class CategoriaAjuste(str, Enum):
    """De donde viene la desviacion respecto de la curva de manual."""

    CLIMA = "clima"              # exposicion costera, heladas, horas frio
    SUELO = "suelo"              # profundidad, salinidad, drenaje
    AGUA = "agua"                # seguridad de riego, calidad
    GESTION = "gestion"          # administracion, mano de obra, poda
    SANIDAD = "sanidad"          # plagas, virosis, decaimiento
    MERCADO = "mercado"          # variedad fuera de mercado, sin packing
    OTRO = "otro"


# ------------------------------------------------------------- la curva

class PuntoCurva(Base):
    """Rendimiento y costo a una edad dada."""

    edad: int = Field(ge=0, le=100, description="anios desde la plantacion")
    rendimiento_kg_ha: float = Field(ge=0)
    costo_directo_usd_ha: Optional[float] = Field(default=None, ge=0)


class TipoDescuento(str, Enum):
    """Como se calcula un descuento de la cadena exportadora."""

    POR_KG = "por_kg"          # USD/kg fijos: packing, materiales, flete
    PORCENTAJE = "porcentaje"  # fraccion: la comision de la exportadora


class BaseDescuento(str, Enum):
    """Sobre que se aplica un descuento porcentual.

    No es un detalle: un 8% sobre el FOB no da lo mismo que un 8% sobre lo que
    queda despues de packing y flete. Se declara por descuento en vez de fijar
    una convencion, porque la respuesta la tiene el tasador y puede variar.
    """

    FOB = "fob"      # sobre el precio FOB original
    SALDO = "saldo"  # sobre lo que queda tras los descuentos anteriores


class Descuento(Base):
    """Un eslabon entre el precio FOB y lo que recibe el agricultor."""

    nombre: str = Field(min_length=2, description="ej. 'Packing', 'Comision exportadora'")
    tipo: TipoDescuento = TipoDescuento.POR_KG
    valor: float = Field(ge=0, description="USD/kg, o fraccion si es porcentaje")
    base: BaseDescuento = BaseDescuento.FOB
    nota: Optional[str] = None


class CadenaExportacion(Base):
    """De FOB publico a retorno del productor, descuento por descuento.

    El FOB es **publico y verificable**; el retorno al productor no. Por eso el
    dato que se guarda es el FOB y la lista de descuentos, y el retorno se
    calcula: asi el banco puede contrastar el precio por su cuenta y auditar
    cada resta, en vez de recibir una cifra sin origen.

    Los descuentos son filas que define el tasador, no una lista fija en el
    codigo: el dia que aparezca uno nuevo se agrega como dato.

    RLP (retorno liquido a productor) y RNP (retorno neto a productor) son el
    mismo concepto; en la casa se dice RLP.
    """

    precio_fob_usd_kg: float = Field(gt=0)
    fuente_fob: Optional[str] = Field(
        default=None, description="de donde salio: ODEPA, exportadora, boletin"
    )
    referencia_fob: Optional[str] = Field(
        default=None, description="periodo del precio: 'promedio 2025', 'semana 40'"
    )
    descuentos: List[Descuento] = Field(default_factory=list)

    def desglose(self) -> List[Dict[str, Any]]:
        """Cada descuento con su monto y el saldo que deja. Es la tabla del informe."""
        saldo = self.precio_fob_usd_kg
        filas: List[Dict[str, Any]] = []
        for descuento in self.descuentos:
            if descuento.tipo is TipoDescuento.PORCENTAJE:
                base = (self.precio_fob_usd_kg
                        if descuento.base is BaseDescuento.FOB else saldo)
                monto = base * descuento.valor
            else:
                monto = descuento.valor
            saldo -= monto
            filas.append({
                "nombre": descuento.nombre,
                "monto_usd_kg": round(monto, 6),
                "saldo_usd_kg": round(saldo, 6),
            })
        return filas

    def rlp_usd_kg(self) -> float:
        """Lo que le queda al agricultor por kilo exportado.

        Puede dar negativo si los descuentos superan al FOB. No se corrige aqui
        -seria tapar un error de carga-: lo detecta el panel de validacion.
        """
        if not self.descuentos:
            return round(self.precio_fob_usd_kg, 6)
        return round(self.desglose()[-1]["saldo_usd_kg"], 6)


class ParametrosComerciales(Base):
    """Apartado 8.2: como se vende y a cuanto."""

    pct_exportacion: float = Field(
        default=0.0, ge=0, le=1,
        description="proporcion exportada, 0 a 1 (0,7 = 70%)",
    )
    rnp_exportacion_usd_kg: float = Field(
        default=0.0, ge=0, description="retorno neto a productor de exportacion"
    )
    rnp_mercado_interno_usd_kg: float = Field(default=0.0, ge=0)
    cadena_exportacion: Optional[CadenaExportacion] = Field(
        default=None,
        description="si esta, el retorno de exportacion se CALCULA desde el FOB "
                    "y sus descuentos, y manda sobre rnp_exportacion_usd_kg",
    )

    @property
    def pct_mercado_interno(self) -> float:
        """Se deduce: lo que no se exporta va al mercado interno."""
        return round(1.0 - self.pct_exportacion, 6)

    def retorno_exportacion_usd_kg(self) -> float:
        """El retorno de exportacion, calculado si se puede y declarado si no.

        Con cadena manda el calculo: guardar el FOB y los descuentos y ademas
        un retorno tecleado a mano invita a que se desincronicen, y el que
        tiene trazabilidad es el calculado.
        """
        if self.cadena_exportacion is not None:
            return self.cadena_exportacion.rlp_usd_kg()
        return self.rnp_exportacion_usd_kg

    def precio_ponderado_usd_kg(self) -> float:
        """Precio medio por kilo segun la mezcla exportacion/interno."""
        return round(
            self.pct_exportacion * self.retorno_exportacion_usd_kg()
            + self.pct_mercado_interno * self.rnp_mercado_interno_usd_kg,
            6,
        )


class EstructuraCosto(Base):
    """Apartado 8.4: costo = base fija por hectarea + parte variable por kilo.

    Verificado contra el archivo original: palto en plena produccion da
    6000 + 13000 x 0,5 = 12.500 USD/ha, que es exactamente lo que declara.
    """

    costo_base_usd_ha: float = Field(default=0.0, ge=0)
    costo_por_kg_usd: float = Field(default=0.0, ge=0)

    def costo_usd_ha(self, rendimiento_kg_ha: float) -> float:
        return round(
            self.costo_base_usd_ha + rendimiento_kg_ha * self.costo_por_kg_usd, 4
        )


class CurvaProduccion(Base):
    """Curva completa de una especie o variedad."""

    id: str = Field(default_factory=lambda: uuid4().hex[:12])
    nombre: str = Field(min_length=2, description="ej. 'Palto Hass clonal'")
    especie: str = Field(min_length=2)
    variedad: Optional[str] = None

    puntos: List[PuntoCurva] = Field(min_length=2)
    comerciales: ParametrosComerciales = ParametrosComerciales()
    costos: EstructuraCosto = EstructuraCosto()

    unidad: str = "kg/ha"
    fuente: Optional[str] = Field(
        default=None, description="de donde salio: informe, publicacion, experiencia"
    )
    #: Si esta curva se armo PARTIENDO de otra -porque la especie no tenia y
    #: el agronomo decidio que otra se le parecia-, aqui queda de cual. No es
    #: adorno: una proyeccion construida sobre una curva prestada dice algo
    #: distinta de una medida, y el banco no deberia descubrirlo por su cuenta.
    #: Misma regla que el estado `de_otra_variedad` de §16.a.sexies, un nivel
    #: mas arriba: alla se presta entre variedades, aqui entre ESPECIES.
    derivada_de: Optional[str] = Field(
        default=None,
        description="nombre de la curva que se uso de base, si se copio de otra",
    )
    notas: Optional[str] = None
    autor: Optional[str] = None
    creada: datetime = Field(default_factory=datetime.now)
    actualizada: datetime = Field(default_factory=datetime.now)

    @field_validator("puntos")
    @classmethod
    def _ordenar(cls, v: List[PuntoCurva]) -> List[PuntoCurva]:
        edades = [p.edad for p in v]
        if len(set(edades)) != len(edades):
            raise ValueError("hay edades repetidas en la curva")
        return sorted(v, key=lambda p: p.edad)

    @property
    def edad_maxima(self) -> int:
        return self.puntos[-1].edad

    @property
    def rendimiento_maximo(self) -> float:
        return max(p.rendimiento_kg_ha for p in self.puntos)

    @property
    def edad_plena_produccion(self) -> Optional[int]:
        """Primera edad en la que se alcanza el rendimiento maximo."""
        techo = self.rendimiento_maximo
        for punto in self.puntos:
            if punto.rendimiento_kg_ha >= techo:
                return punto.edad
        return None


# ------------------------------------------------------------- ajustes

class Ajuste(Base):
    """Un factor reutilizable que el tasador guarda para volver a usarlo.

    ``factor_pct`` es la variacion sobre la curva base: -10 baja un 10%, +5 sube
    un 5%. Se guarda con su motivo porque el banco puede preguntar por que esta
    plantacion rinde menos que la de al lado.
    """

    id: str = Field(default_factory=lambda: uuid4().hex[:12])
    nombre: str = Field(min_length=2, description="ej. 'Sector costero'")
    descripcion: Optional[str] = Field(
        default=None, description="por que este factor baja o sube el rendimiento"
    )
    categoria: CategoriaAjuste = CategoriaAjuste.OTRO
    factor_pct: float = Field(
        ge=-100, le=100, description="variacion porcentual sobre la curva base"
    )
    autor: Optional[str] = None
    creada: datetime = Field(default_factory=datetime.now)

    @property
    def multiplicador(self) -> float:
        return round(1.0 + self.factor_pct / 100.0, 6)


class AjusteAplicado(Base):
    """Un ajuste puesto sobre una curva en un informe concreto."""

    nombre: str
    factor_pct: float = Field(ge=-100, le=100)
    categoria: CategoriaAjuste = CategoriaAjuste.OTRO
    justificacion: Optional[str] = Field(
        default=None, description="lo que observo el tasador en ESTE predio"
    )

    @property
    def multiplicador(self) -> float:
        return round(1.0 + self.factor_pct / 100.0, 6)


class PuntoProyectado(Base):
    """Una fila del resultado: que pasa a esta edad, ya con los ajustes puestos."""

    edad: int
    rendimiento_base_kg_ha: float
    rendimiento_ajustado_kg_ha: float
    produccion_total_kg: float
    ingreso_usd: float
    costo_usd: float
    margen_usd: float


class Proyeccion(Base):
    """Resultado completo de aplicar una curva a una plantacion concreta."""

    curva: str
    superficie_ha: float
    ajustes: List[AjusteAplicado] = Field(default_factory=list)
    factor_total: float = Field(
        description="producto de todos los ajustes; 0,855 = 14,5% menos"
    )
    puntos: List[PuntoProyectado] = Field(default_factory=list)
    variacion_pct: float = Field(
        default=0.0,
        description="variacion respecto de la curva de manual; -19 = 19% menos",
    )

    @model_validator(mode="after")
    def _derivar_variacion(self) -> "Proyeccion":
        """Campo normal y no @property ni computed_field.

        FastAPI clona el modelo de respuesta y en ese clon los campos calculados
        se pierden (llegan como null al navegador). Como campo derivado aqui
        siempre viaja, y sigue siendo imposible que se desincronice de
        ``factor_total`` porque se recalcula en cada validacion.
        """
        object.__setattr__(
            self, "variacion_pct", round((self.factor_total - 1.0) * 100, 2)
        )
        return self


class SolicitudProyeccion(Base):
    """Lo que pide el frontend para simular."""

    curva_id: str
    superficie_ha: float = Field(gt=0)
    edad_actual: Optional[int] = Field(
        default=None, ge=0, description="si se indica, solo se devuelve esa edad"
    )
    horizonte: int = Field(default=25, ge=1, le=100)
    ajustes: List[AjusteAplicado] = Field(default_factory=list)

    @model_validator(mode="after")
    def _limite_ajustes(self) -> "SolicitudProyeccion":
        if len(self.ajustes) > 20:
            raise ValueError("demasiados ajustes: revise la seleccion")
        return self


class SolicitudCurvaParametrica(Base):
    """Los 4 numeros del agronomo para armar una curva por rampa proporcional.

    Pensado para especies o variedades sin tabla propia: en vez de tabular
    rendimiento anio por anio, se declara cuanto tarda en producir, con que
    nivel arranca y cuanto rinde en plena produccion.
    """

    nombre: str = Field(min_length=2)
    especie: str = Field(min_length=2)
    variedad: Optional[str] = None
    anios_improductivos: int = Field(ge=0, le=100)
    nivel_productivo_inicial_kg_ha: float = Field(ge=0)
    anios_a_plena_produccion: int = Field(ge=0, le=100)
    nivel_plena_produccion_kg_ha: float = Field(gt=0)
    # El declive es opcional; si va, van los tres. Un frutal no se queda en su
    # peak para siempre y sin esto se le proyecta a un huerto de 30 anios el
    # rendimiento de uno en plenitud.
    edad_inicio_declive: Optional[int] = Field(default=None, ge=0, le=100)
    edad_final: Optional[int] = Field(default=None, ge=0, le=120)
    nivel_final_kg_ha: Optional[float] = Field(default=None, ge=0)
    #: La otra forma de decir lo mismo, y la que usa el agronomo: en vez de la
    #: EDAD en que termina de bajar, cuantos anios aguanta el peak y a que
    #: ritmo cae. Nadie sabe de memoria a que edad un nogal llega a 8.000
    #: kg/ha, pero si sabe que cae "como un 4% al anio". Si viene, el servidor
    #: deduce `edad_final` con `declive_desde_porcentaje`.
    anios_de_plena_produccion: Optional[int] = Field(default=None, ge=0, le=100)
    decrecimiento_anual_pct: Optional[float] = Field(default=None, gt=0, lt=100)
    comerciales: Optional[ParametrosComerciales] = None
    costos: Optional[EstructuraCosto] = None
    #: La forma en que el agronomo SI sabe el costo: cuanto cuesta la hectarea
    #: parada y cuanto con el huerto a full. De los dos extremos sale el
    #: coeficiente variable del modelo de §8.2, que nadie tiene en la cabeza.
    #: Lo pidio el usuario. Si vienen los dos, mandan sobre `costos`.
    costo_base_usd_ha: Optional[float] = Field(default=None, ge=0)
    costo_en_plena_produccion_usd_ha: Optional[float] = Field(default=None, ge=0)
    fuente: Optional[str] = None
    notas: Optional[str] = None
    autor: Optional[str] = None


# --------------------------------------------- mapeos de catastro aprendidos

class LecturaIA(Base):
    """Lo que el modelo ya leyo de un archivo, guardado por su huella.

    Existe para no pagar dos veces lo mismo. La huella es el SHA-256 del
    CONTENIDO, no del nombre: la misma escritura renombrada es la misma
    escritura, y dos fotos identicas se cobran una sola vez.

    Es compartido, no por usuario: si el usuario lee una escritura desde su
    computador, otro tasador reanalizando ese mismo PDF desde el suyo tampoco paga.
    """

    id: str = Field(default_factory=lambda: uuid4().hex[:12])
    huella: str = Field(description="SHA-256 del contenido del archivo")
    tipo: str = Field(description="'documento' o 'fotografia'")
    datos: Dict[str, Any] = Field(
        default_factory=dict,
        description="lo que devolvio el modelo, tal cual, para reconstruirlo",
    )
    modelo: Optional[str] = Field(default=None, description="quien lo leyo")
    nombre: Optional[str] = Field(
        default=None, description="como se llamaba el archivo la primera vez")
    veces_usado: int = 1
    creada: datetime = Field(default_factory=datetime.now)


class MapeoCatastro(Base):
    """Correspondencia de columnas ya confirmada por el tasador.

    Se guarda con la huella de los encabezados del archivo: si el mismo
    agricultor manda su catastro el mes que viene, se reconoce solo y no hay
    que volver a mapear nada.
    """

    id: str = Field(default_factory=lambda: uuid4().hex[:12])
    huella: str = Field(description="hash de los encabezados normalizados")
    nombre: str = Field(description="a quien pertenece este formato")
    columnas: Dict[str, str] = Field(description="campo del informe -> letra")
    encabezados: Dict[str, str] = Field(default_factory=dict)
    veces_usado: int = 1
    creada: datetime = Field(default_factory=datetime.now)
