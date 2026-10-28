"""Todo lo que ODEPA sabe de una especie, en una sola respuesta.

Son tres archivos distintos -FOB de exportacion, precio de mercado interno y
superficie del Catastro Fruticola- y el tasador no quiere hacer tres consultas:
quiere escribir «avellano» y ver que hay. Este modulo los junta.

**Cada uno usa un vocabulario distinto**, y esa es la mitad del trabajo:

| Fuente | Como se llama el avellano |
|---|---|
| El catastro del cliente | Avellano europeo |
| Aduanas (FOB) | Avellana |
| Precios de mercado interno | Avellana |
| Catastro Fruticola CIREN | Avellano |

O sea que hay que cruzar el nombre del ARBOL con el del FRUTO. El puente vive
en `services/precios.py:SINONIMOS_DE_ESPECIE` -grupos de nombres que
significan la misma especie- y se reusa aqui: tener dos copias de esa tabla es
garantia de que se separen, y la que se separaria es la que el tasador ve.

**Ningun numero de aqui entra solo al informe.** Son referencias para que una
persona decida, y cada una viene con su fuente, su periodo y sus reparos:

- El **FOB no es lo que recibe el productor** (§10.2): falta la cadena de
  descuentos. Y el grupo puede mezclar productos de precio muy distinto -con
  cascara y sin cascara difieren 2,4 veces-.
- Del **mercado interno** solo sirve el precio mayorista; feria y supermercado
  llevan margen de retail.
- La **superficie de CIREN** es regional y estimada, no medida en el predio.
  Sirve para comparar el orden de magnitud de lo que declara el cliente.
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from services.precios import (
    EQUIVALENTES, BibliotecaFOB, cargar_fob, emparejar_por_nombre,
)

__all__ = [
    "PrecioInterno", "SuperficieEspecie", "Referencias",
    "BibliotecaReferencias", "cargar_referencias",
    "ResumenSerie", "resumen_de_serie",
    "TEMPORADAS_DEL_RESUMEN", "ANIOS_DE_LA_DECADA", "AJUSTE_AL_PRODUCTOR",
]

#: Cuantas temporadas de FOB resume el cuadro: **la ultima cerrada y cinco
#: hacia atras**. Lo pidio el agrónomo, y el criterio es agronomico: un precio
#: de exportacion se mira contra su propia historia reciente -una helada o un
#: peak de demanda mueven un anio entero- pero la de hace diez temporadas ya
#: es otro mercado.
TEMPORADAS_DEL_RESUMEN = 6

#: La superficie se mira a diez anios: lo que interesa ahi no es el ruido de
#: un anio sino si la especie esta en expansion o en retirada.
ANIOS_DE_LA_DECADA = 10

#: Que fraccion del precio de MERCADO MAYORISTA le queda al productor. Lo dio
#: el agrónomo: *«tiene que haber como un precio del productor que tiene que ser
#: como un ajuste del 50% o 60% del precio del mercado interno»*.
#:
#: Se guarda el RANGO y no un solo numero porque eso fue lo que dijo.
#: Promediarlo a 55% inventaria una precision que no existe: lo que queda al
#: productor depende del canal, del volumen y de quien transporte, y esa
#: eleccion es del tasador. Es el mismo criterio que la cadena FOB -> RLP
#: (§10.2): se muestra de donde sale cada numero en vez de entregar uno solo
#: sin respaldo.
AJUSTE_AL_PRODUCTOR = (0.50, 0.60)


def _clave(texto: Any) -> str:
    sin = unicodedata.normalize("NFKD", str(texto or ""))
    sin = "".join(c for c in sin if not unicodedata.combining(c))
    return " ".join(sin.lower().split())


#: La busqueda es UNA, en `services/precios.py`. Antes vivia dos veces «porque
#: las tablas tienen forma distinta», con una nota diciendo que convenia que no
#: se separaran — y se separaron igual: las dos recortaban la ultima letra para
#: absorber el plural, y con eso 'Cerezos', 'Nogales' y 'Limoneros' no
#: encontraban nada en ninguna de las tres fuentes. Una regla de idioma no
#: puede tener copias.
_buscar = emparejar_por_nombre


def _anio_de(clave: Any) -> Tuple[int, str]:
    """Ordena cronologicamente las dos formas de rotular un periodo.

    La superficie viene por anio («2016») y el FOB por temporada
    («Sep 2016 / Ago 2017»). En los dos casos manda el primer anio que
    aparece. Ordenar por el texto funcionaria de casualidad con los anios
    sueltos y no con las temporadas.
    """
    digitos = re.findall(r"\d{4}", str(clave))
    return (int(digitos[0]) if digitos else 0, str(clave))


@dataclass
class ResumenSerie:
    """Minimo, maximo y promedio de las ultimas N observaciones de una serie.

    Es lo que pidio el agrónomo para el FOB: un precio suelto no dice si el anio
    fue normal o excepcional, y para saberlo hay que verlo contra su propio
    rango reciente.

    **El promedio es simple, no ponderado por volumen**, y por eso se rotula
    asi en la pantalla. La serie del boletin trae el precio medio de cada
    temporada pero no sus toneladas, asi que ponderar exigiria un dato que no
    esta; llamarlo «promedio» a secas dejaria creer que si lo esta.
    """

    puntos: List[Dict[str, Any]] = field(default_factory=list)
    desde: str = ""
    hasta: str = ""
    minimo: float = 0.0
    maximo: float = 0.0
    promedio: float = 0.0
    primero: float = 0.0
    ultimo: float = 0.0
    variacion: Optional[float] = None
    n: int = 0
    de_cuantos: int = 0
    de_cuantos_se_pidieron: int = 0

    @property
    def parcial(self) -> bool:
        """True cuando la serie no alcanzaba para la ventana pedida.

        Hay que poder decirlo: «promedio de 6 temporadas» sobre 3 datos es un
        promedio distinto, y en pantalla se rotula con el numero real.
        """
        return self.n < self.de_cuantos_se_pidieron


def resumen_de_serie(serie: Optional[Dict[str, Any]],
                     ultimos: int) -> Optional[ResumenSerie]:
    """Recorta la serie a sus ultimas `ultimos` observaciones y la resume.

    Devuelve `None` si no hay nada que resumir: una especie sin serie no
    puede mostrar un rango, y fabricar uno con un solo punto diria que el
    precio no se mueve.
    """
    if not serie:
        return None
    ordenada = [(str(k), float(v)) for k, v in
                sorted(serie.items(), key=lambda kv: _anio_de(kv[0]))
                if v is not None]
    if not ordenada:
        return None

    ventana = ordenada[-ultimos:] if ultimos and ultimos > 0 else ordenada
    valores = [v for _clave, v in ventana]
    primero, ultimo = valores[0], valores[-1]
    return ResumenSerie(
        puntos=[{"etiqueta": k, "valor": round(v, 4)} for k, v in ventana],
        desde=ventana[0][0],
        hasta=ventana[-1][0],
        minimo=round(min(valores), 4),
        maximo=round(max(valores), 4),
        promedio=round(sum(valores) / len(valores), 4),
        primero=round(primero, 4),
        ultimo=round(ultimo, 4),
        # Cuanto crecio o cayo de punta a punta de la ventana. Es lo que
        # hace legible la serie de superficie: «+217% en la decada» dice de
        # una vez lo que la tabla de anios obliga a comparar a mano.
        variacion=(round((ultimo - primero) / primero, 4) if primero else None),
        n=len(ventana),
        de_cuantos=len(ordenada),
        de_cuantos_se_pidieron=(ultimos if ultimos and ultimos > 0
                                else len(ordenada)),
    )


@dataclass
class PrecioInterno:
    """El precio de mercado interno, del punto de venta que sirve."""

    nombre: str
    clp_kg: Optional[float]
    anio: Optional[str]
    meses: Optional[int]
    punto_de_venta: str
    por_mes: Dict[str, float] = field(default_factory=dict)
    otros_puntos: Dict[str, float] = field(default_factory=dict)
    advertencia: Optional[str] = None
    confianza: str = "media"

    def usd_kg(self, tipo_cambio: Optional[float]) -> Optional[float]:
        """En dolares, con el tipo de cambio que declaro el tasador.

        No se guarda convertido: el tipo de cambio es del informe, no del dato.
        Convertirlo al importar lo congelaria con el cambio de ese dia.
        """
        if self.clp_kg is None or not tipo_cambio:
            return None
        return round(self.clp_kg / float(tipo_cambio), 4)

    def al_productor(self) -> Optional[Dict[str, Any]]:
        """Lo que le queda al productor del precio mayorista.

        El precio del mercado mayorista **no es lo que recibe el agricultor**:
        entre medio hay flete, comision y merma. el agrónomo lo cifro en un 50 a
        60 por ciento. Es el mismo problema que el FOB (§10.2) del otro lado
        de la cadena, y se trata igual: se muestra el rango y de donde sale,
        no un numero solo.
        """
        if self.clp_kg is None:
            return None
        bajo, alto = AJUSTE_AL_PRODUCTOR
        return {
            "pct_bajo": bajo,
            "pct_alto": alto,
            "clp_kg_bajo": round(self.clp_kg * bajo, 2),
            "clp_kg_alto": round(self.clp_kg * alto, 2),
            "de": self.punto_de_venta,
            "criterio": ("Estimación: entre el {:.0f}% y el {:.0f}% del precio "
                         "mayorista. No es un dato publicado —es el ajuste que "
                         "usa Fitogreen— y depende del canal y del volumen."
                         .format(bajo * 100, alto * 100)),
        }


@dataclass
class SuperficieEspecie:
    """Cuanta superficie hay de esta especie en Chile, y donde."""

    nombre: str
    hectareas: float
    regiones: Dict[str, float] = field(default_factory=dict)
    variedades: List[Dict[str, Any]] = field(default_factory=list)
    variedades_en_el_catastro: int = 0
    serie_nacional_ha: Dict[str, float] = field(default_factory=dict)

    def en_la_region(self, region: Optional[str]) -> Optional[float]:
        if not region:
            return None
        buscada = _clave(region)
        for nombre, ha in self.regiones.items():
            if _clave(nombre) == buscada or buscada in _clave(nombre):
                return ha
        return None

    def peso_del_cliente(self, hectareas_cliente: float,
                         region: Optional[str] = None) -> Optional[float]:
        """Que fraccion de la superficie regional representa este cliente.

        Es la frase que el banco entiende: «760 ha donde la region tiene
        4.956» dice mas que cualquiera de los dos numeros solo.
        """
        base = self.en_la_region(region) if region else self.hectareas
        if not base or not hectareas_cliente:
            return None
        return round(hectareas_cliente / base, 4)

    def decada(self) -> Optional[ResumenSerie]:
        """Como se movio la superficie de esta especie en los ultimos 10 anios.

        La serie del catastro llega hasta 2004, pero veintidos anios de
        fruticultura chilena son otro pais: lo que el tasador necesita saber
        es si la especie viene creciendo o retirandose AHORA.
        """
        return resumen_de_serie(self.serie_nacional_ha, ANIOS_DE_LA_DECADA)


@dataclass
class Referencias:
    """Lo que hay de una especie, junto."""

    especie: str
    fob: Optional[Any] = None                 # PrecioReferencia
    fob_codigos: List[Any] = field(default_factory=list)
    fob_es_mezcla: bool = False
    fob_serie: Dict[str, float] = field(default_factory=dict)
    #: «Fresca» o «Frutos secos». Va a la pantalla porque el agrónomo pidio que
    #: los datos salgan de la fruta fresca y no de conserva: que ya sea asi no
    #: sirve de nada si no se ve.
    fob_procesamiento: Optional[str] = None
    interno: Optional[PrecioInterno] = None
    superficie: Optional[SuperficieEspecie] = None

    @property
    def hay_algo(self) -> bool:
        return any((self.fob, self.interno, self.superficie))

    def resumen_fob(self) -> Optional[ResumenSerie]:
        """El rango reciente del FOB: minimo, maximo y promedio.

        Un FOB solo -«4,55 USD/kg»- no dice si esa temporada fue buena o
        mala. Contra las seis ultimas, si.
        """
        return resumen_de_serie(self.fob_serie, TEMPORADAS_DEL_RESUMEN)


class BibliotecaReferencias:
    """Las tres fuentes de ODEPA, consultables por especie."""

    def __init__(self, fob: Optional[BibliotecaFOB],
                 interno: Optional[Dict[str, Any]],
                 catastro: Optional[Dict[str, Any]]):
        self._fob = fob
        self._interno = interno or {}
        self._catastro = catastro or {}

    # -- lo que falta se dice, no se finge -------------------------------

    @property
    def fuentes(self) -> Dict[str, Any]:
        """Que archivos hay cargados y de cuando. Sin esto, una biblioteca a
        medias se ve igual que una completa y el tasador no se entera."""
        return {
            "fob": {
                "cargado": self._fob is not None,
                "temporada": self._fob.temporada if self._fob else None,
                "fuente": self._fob.fuente if self._fob else None,
            },
            "mercado_interno": {
                "cargado": bool(self._interno),
                "anio": self._interno.get("anio_de_referencia"),
                "punto_de_venta": self._interno.get("punto_de_venta_utilizable"),
                "fuente": self._interno.get("fuente_declarada"),
            },
            "superficie": {
                "cargado": bool(self._catastro),
                "fuente": self._catastro.get("fuente_declarada"),
                "advertencia": self._catastro.get("advertencia"),
            },
        }

    # -- consultas --------------------------------------------------------

    def interno_de(self, especie: str) -> Optional[PrecioInterno]:
        registro = _buscar(self._interno.get("productos", {}), especie)
        if registro is None:
            return None
        punto = self._interno.get("punto_de_venta_utilizable", "")
        anio = registro.get("anio_de_referencia")
        del_punto = (registro.get("puntos_de_venta", {}).get(punto, {})
                     .get(str(anio), {}))
        # Los otros puntos van como contexto: ver que el supermercado paga seis
        # veces mas es justamente lo que evita confundirlos.
        otros = {}
        for nombre, anios in registro.get("puntos_de_venta", {}).items():
            if nombre == punto:
                continue
            valor = anios.get(str(anio)) or (
                anios[sorted(anios)[-1]] if anios else None)
            if valor:
                otros[nombre] = valor.get("promedio_clp_kg")
        return PrecioInterno(
            nombre=registro.get("nombre", especie),
            clp_kg=registro.get("promedio_clp_kg"),
            anio=str(anio) if anio else None,
            meses=registro.get("meses_del_promedio"),
            punto_de_venta=punto,
            por_mes=del_punto.get("por_mes", {}),
            otros_puntos=otros,
            advertencia=registro.get("advertencia"),
            confianza=registro.get("confianza", "media"),
        )

    def superficie_de(self, especie: str) -> Optional[SuperficieEspecie]:
        registro = _buscar(self._catastro.get("especies", {}), especie)
        if registro is None:
            return None
        return SuperficieEspecie(
            nombre=registro.get("nombre", especie),
            hectareas=registro.get("hectareas", 0.0),
            regiones=registro.get("regiones", {}),
            variedades=registro.get("variedades", []),
            variedades_en_el_catastro=registro.get(
                "variedades_en_el_catastro", 0),
            serie_nacional_ha=registro.get("serie_nacional_ha", {}),
        )

    def de(self, especie: str) -> Referencias:
        r = Referencias(especie=especie)
        if self._fob is not None:
            r.fob = self._fob.fob_de(especie)
            r.fob_codigos = self._fob.codigos_de(especie)
            r.fob_es_mezcla = self._fob.es_mezcla(especie)
            r.fob_serie = self._fob.serie_de(especie)
            r.fob_procesamiento = self._fob.procesamiento_de(especie)
        r.interno = self.interno_de(especie)
        r.superficie = self.superficie_de(especie)
        return r

    def parecidas(self, texto: str, cuantas: int = 6) -> List[str]:
        """Las especies de ODEPA mas parecidas a lo que se escribio.

        Es para PROPONER, no para decidir: si alguien escribe «ciruelo» y en
        el boletin de mayoristas figura como «Ciruela», el emparejador ya lo
        resuelve solo; esto es para cuando NO lo resuelve y hay que ofrecerle
        algo en vez de un «no hay nada».

        Se busca por partes del nombre primero -que es como escribe una
        persona: «avell» para avellano- y despues por parecido de letras, que
        cubre el error de tipeo. `difflib` viene con Python: una dependencia
        nueva para ordenar seis nombres seria desproporcionada.
        """
        import difflib

        buscado = _clave(texto)
        if not buscado:
            return []
        todas = self.especies
        indice = {_clave(n): n for n in todas}

        salida: List[str] = []

        def sumar(nombre: str) -> None:
            if nombre and nombre not in salida:
                salida.append(nombre)

        # Empieza por lo escrito, o lo escrito esta adentro. Lo primero pesa
        # mas: quien escribe «cere» busca «Cereza», no «Otros berries».
        for clave in sorted(indice):
            if clave.startswith(buscado):
                sumar(indice[clave])
        for clave in sorted(indice):
            if buscado in clave:
                sumar(indice[clave])
        # Y los sinonimos declarados: quien escribe «nogal» tiene que ver
        # «Nuez» aunque no compartan una sola letra.
        for igual in EQUIVALENTES.get(buscado, ()):
            sumar(indice.get(igual))
        # El parecido de letras va SOLO si lo anterior no dio casi nada. Es
        # para el error de tipeo -«kivi» por «Kiwi»-, y si se deja suelto
        # ensucia: con «arandano» ofrecia «Mandarino» y «Naranja», que se
        # parecen en las letras y en nada mas.
        if len(salida) < 3:
            for clave in difflib.get_close_matches(buscado, list(indice),
                                                   n=cuantas, cutoff=0.6):
                sumar(indice[clave])
        return salida[:cuantas]

    @property
    def especies(self) -> List[str]:
        """Todo lo consultable, sin repetir. Sirve para el buscador."""
        nombres = set()
        if self._fob is not None:
            nombres.update(self._fob.grupos)
        nombres.update(r.get("nombre", k)
                       for k, r in self._interno.get("productos", {}).items())
        nombres.update(r.get("nombre", k)
                       for k, r in self._catastro.get("especies", {}).items())
        return sorted(nombres)


def _leer(ruta: Path) -> Optional[Dict[str, Any]]:
    """Falla en blando: sin el archivo se dice que no esta, y las otras dos
    fuentes siguen sirviendo. Que falte el catastro no puede dejar sin precio
    al tasador."""
    try:
        with ruta.open(encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def cargar_referencias(carpeta: Path) -> BibliotecaReferencias:
    fob = None
    ruta_fob = carpeta / "precios_fob_odepa.json"
    if ruta_fob.exists():
        try:
            fob = cargar_fob(ruta_fob)
        except (OSError, ValueError):
            fob = None
    return BibliotecaReferencias(
        fob=fob,
        interno=_leer(carpeta / "precios_mercado_interno.json"),
        catastro=_leer(carpeta / "catastro_ciren.json"),
    )
