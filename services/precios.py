"""Precios de referencia: de donde sale cada cifra y cuanto vale confiar en ella.

Hasta ahora el catalogo de curvas guardaba un precio suelto -1,75 USD/kg para
el palto- sin decir de donde salia. Si el banco pregunta, no hay respuesta.
Este modulo pone detras de cada precio su procedencia, su periodo y su nivel de
confianza, y deja que el tasador decida si lo usa.

**No impone nada.** Es una sugerencia con respaldo; lo que el cliente declara
manda siempre. Un precio de referencia sirve para dos cosas legitimas: rellenar
cuando el cliente no tiene registros, y proyectar anios futuros, donde ningun
dato declarado existe todavia.

Dos tipos de precio que NO son intercambiables:

- **A productor**: lo que le llega al agricultor. ODEPA lo publica, pero solo
  de ganado, leche, vino y uva vinifera. Por eso la tabla del HdU es solo de
  vid: es lo unico que hay a ese nivel.
- **FOB de exportacion**: la fruta puesta en el barco. Es publico para todo lo
  que se exporta, y es el dato de entrada que pidio el agrónomo. Todavia tiene que
  pasar por la cadena de descuentos (``schemas/curvas.py:CadenaExportacion``)
  para convertirse en retorno al productor.

Confundirlos infla el ingreso al doble: en el avellano, el FOB de la temporada
2025/26 (6,90 USD/kg) y el retorno a productor citado para Giffoni (3,35) se
llevan un 51% de diferencia, que es justamente lo que cobra la cadena.
"""

from __future__ import annotations

import json
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from services.catastro_import import formas_de_nombre

__all__ = [
    "PrecioReferencia", "BibliotecaPrecios", "cargar_precios",
]

#: Un precio FOB no es comparable con uno a productor.
TIPO_PRODUCTOR = "precio a productor"
TIPO_FOB = "FOB exportacion"


def _clave(texto: Any) -> str:
    plano = "".join(
        c for c in unicodedata.normalize("NFD", str(texto or "").strip())
        if unicodedata.category(c) != "Mn"
    )
    return " ".join(plano.lower().split())


@dataclass
class PrecioReferencia:
    """Un precio con su respaldo. El respaldo es la mitad del valor."""

    nombre: str
    valor_usd_kg: float
    tipo: str = TIPO_PRODUCTOR
    fuente: Optional[str] = None
    periodo: Optional[str] = None
    unidad: Optional[str] = None
    confianza: str = "media"
    advertencia: Optional[str] = None
    serie_clp_kg: Optional[Dict[str, float]] = None
    compuesto_por: Optional[List[str]] = None

    @property
    def es_fob(self) -> bool:
        return _clave(self.tipo).startswith("fob")

    @property
    def listo_para_el_flujo(self) -> bool:
        """Si puede ir directo a la fila de precio del flujo.

        Un FOB no puede: primero hay que restarle la cadena de descuentos.
        Ponerlo tal cual duplicaria el ingreso del cliente.
        """
        return not self.es_fob

    def procedencia(self) -> str:
        """Una linea para el comentario de celda del Excel."""
        partes = [self.fuente or "sin fuente declarada"]
        if self.periodo:
            partes.append(self.periodo)
        if self.unidad:
            partes.append(self.unidad)
        return "{} ({}) - confianza {}".format(
            self.nombre, ", ".join(partes), self.confianza)


class BibliotecaPrecios:
    """Los precios de referencia, consultables por cultivar, grupo o especie."""

    def __init__(self, crudo: Dict[str, Any]):
        self.version = crudo.get("version")
        self.origen = crudo.get("origen")
        self.tipo_cambio_clp_usd = crudo.get("tipo_cambio_clp_usd")
        self.anios: List[str] = crudo.get("anios", [])
        self.advertencia = crudo.get("advertencia")
        self._cultivares = crudo.get("cultivares", {})
        self._grupos = crudo.get("grupos", {})
        self._otras = {
            k: v for k, v in crudo.get("otras_especies", {}).items()
            if not k.startswith("_")
        }

    # -- consultas ---------------------------------------------------------

    def _de_tabla(self, registro: Dict[str, Any]) -> Optional[PrecioReferencia]:
        valor = registro.get("promedio_usd_kg")
        if valor is None:
            return None
        return PrecioReferencia(
            nombre=registro.get("nombre", ""),
            valor_usd_kg=float(valor),
            tipo=registro.get("tipo", TIPO_PRODUCTOR),
            fuente=registro.get("fuente"),
            periodo="promedio {}-{}".format(self.anios[0], self.anios[-1])
                    if self.anios else None,
            confianza=registro.get("confianza", "media"),
            serie_clp_kg=registro.get("serie_clp_kg"),
            compuesto_por=registro.get("compuesto_por"),
        )

    def buscar(self, nombre: str) -> List[PrecioReferencia]:
        """Todas las referencias que hay para un nombre, de mejor a peor.

        Devuelve una LISTA y no un valor: para el avellano hay un FOB y dos
        precios a productor de distinta variedad, y elegir por el tasador
        seria decidir por el lo que no me corresponde.
        """
        objetivo = _clave(nombre)

        for tabla in (self._grupos, self._cultivares):
            registro = tabla.get(objetivo)
            if registro:
                precio = self._de_tabla(registro)
                return [precio] if precio else []

        registro = self._otras.get(objetivo)
        if not registro:
            # 'Avellano europeo' contra la clave 'avellano', y al reves.
            for clave_otra, candidato in self._otras.items():
                if objetivo.startswith(clave_otra) or clave_otra.startswith(objetivo):
                    registro = candidato
                    break
        if not registro:
            return []

        encontrados = [
            PrecioReferencia(
                nombre=registro.get("nombre", nombre),
                valor_usd_kg=float(r["valor_usd_kg"]),
                tipo=r.get("tipo", TIPO_PRODUCTOR),
                fuente=r.get("fuente"),
                periodo=r.get("periodo"),
                unidad=r.get("unidad"),
                confianza=r.get("confianza", "media"),
                advertencia=r.get("advertencia"),
            )
            for r in registro.get("referencias", [])
            if r.get("valor_usd_kg") is not None
        ]
        orden = {"alta": 0, "media": 1, "baja": 2}
        return sorted(encontrados, key=lambda p: orden.get(p.confianza, 3))

    def sugerir_para_flujo(self, nombre: str) -> Optional[PrecioReferencia]:
        """La mejor referencia que se puede usar TAL CUAL en la fila de precio.

        Descarta los FOB: necesitan la cadena de descuentos primero. Si lo unico
        que hay es un FOB, devuelve None -y el llamador debe decirlo- en vez de
        entregar un precio que duplica el ingreso.
        """
        for precio in self.buscar(nombre):
            if precio.listo_para_el_flujo:
                return precio
        return None

    def fob_de(self, nombre: str) -> Optional[PrecioReferencia]:
        """El FOB publico, para alimentar la cadena FOB -> descuentos -> RLP."""
        for precio in self.buscar(nombre):
            if precio.es_fob:
                return precio
        return None

    @property
    def especies_con_precio(self) -> List[str]:
        nombres = [r.get("nombre", k) for k, r in self._grupos.items()]
        nombres += [r.get("nombre", k) for k, r in self._otras.items()]
        return sorted(nombres)


def cargar_precios(ruta: Union[str, Path]) -> BibliotecaPrecios:
    with Path(ruta).open(encoding="utf-8") as fh:
        return BibliotecaPrecios(json.load(fh))


# ------------------------------------------------------- FOB de exportacion
#
# Viene del boletin de fruta de ODEPA, que publica los datos de Aduanas
# (`scripts/importar_fob_odepa.py`). Es una biblioteca APARTE de la anterior a
# proposito: son dos cosas distintas -precio a productor contra FOB- y
# mezclarlas en el mismo archivo es exactamente el error que §10.2 previene.


#: El catastro y CIREN nombran el ARBOL; Aduanas y los precios de mercado
#: nombran el FRUTO. Sin este puente, media biblioteca queda inalcanzable.
#:
#: **Esta declarado y no deducido a proposito.** La primera version recortaba
#: la ultima letra -palto/palta, cerezo/cereza- y parecia elegante, pero es un
#: truco de letras: fallaba con nogal/nuez, con olivo/aceituna y con los
#: plurales del catalogo de curvas ('Cerezos', 'Nogales'). Y peor, un recorte
#: puede acertar por casualidad con la especie equivocada, que en un informe
#: del banco no se nota hasta que alguien compara precios.
#:
#: Los nombres de la derecha salen de los archivos REALES: los grupos de
#: Aduanas y los productos del boletin de precios, no de lo que uno supondria.
#: Por eso «ciruelo» aparece a los dos lados -ODEPA nombra al arbol en el FOB y
#: a la fruta en los precios internos- y «arandano» apunta a «arandanos y
#: mirtilos», que es como lo llama el arancel.
#: Nombres que significan LA MISMA especie, agrupados.
#:
#: Antes esto era un diccionario arbol -> fruto, de UN solo valor, y por eso
#: el ciruelo no encontraba su precio de mercado interno: Aduanas lo llama
#: «Ciruelo» -el arbol- y el boletin de mayoristas «Ciruela» -la fruta-, asi
#: que la entrada tenia que elegir una de las dos y la otra quedaba
#: inalcanzable. Con grupos no hay que elegir: se prueban todos.
#:
#: **El orden dentro del grupo importa** y es el orden en que se prueban.
#: Guardarlos en un conjunto haria que con dos candidatos en la misma tabla
#: -CIREN tiene «Ciruelo Europeo» y «Ciruelo Japones»- ganara uno distinto en
#: cada corrida, porque Python revuelve el orden de los conjuntos de textos
#: entre procesos. Un emparejado que cambia solo es peor que uno que se
#: equivoca siempre igual.
#:
#: Los nombres salen de los archivos REALES, no de lo que uno supondria: el
#: nogal y la nuez no comparten ni la raiz, y el arancel dice «arandanos y
#: mirtilos».
SINONIMOS_DE_ESPECIE = (
    ("almendro", "almendra"),
    ("arandano", "arandanos y mirtilos", "arandano americano"),
    ("avellano", "avellana"),
    ("castano", "castana"),
    ("cerezo", "cereza", "guindo", "guindo agrio"),
    ("chirimoyo", "chirimoya"),
    ("ciruelo", "ciruela", "ciruelo europeo", "ciruelo japones"),
    # «Duraznero Consumo Fresco» va DECLARADO y antes que nada: el catastro de
    # CIREN separa el duraznero de consumo fresco (1.557 ha) del tipo
    # conservero (5.330 ha), y el agrónomo pidio expresamente que los datos
    # salgan de la fruta fresca. Sin esta linea acertaba de casualidad, porque
    # el orden alfabetico pone «consumo fresco» antes que «tipo conservero»:
    # un acierto que depende de como esta ordenado el archivo no es un acierto.
    ("duraznero", "durazno", "duraznero consumo fresco"),
    ("granado", "granada"),
    ("guayabo", "guayaba"),
    ("higuera", "higo"),
    ("limonero", "limon"),
    ("lucumo", "lucuma"),
    ("mandarino", "mandarina"),
    ("manzano", "manzana"),
    ("naranjo", "naranja"),
    ("nectarino", "nectarin"),
    ("nogal", "nuez"),
    ("olivo", "aceituna"),
    ("palto", "palta"),
    ("papayo", "papaya"),
    ("peral", "pera"),
    ("peral asiatico", "pera asiatica"),
    ("vid", "uva", "vid de mesa", "vid vinifera", "parronal"),
)


def _armar_equivalencias() -> Dict[str, tuple]:
    """De cada nombre, los otros del grupo, en el orden en que se declararon.

    Si dos grupos comparten un nombre quedan unidos, que es lo que uno espera
    de una tabla de sinonimos.
    """
    tabla: Dict[str, list] = {}
    for grupo in SINONIMOS_DE_ESPECIE:
        for nombre in grupo:
            iguales = tabla.setdefault(nombre, [])
            for otro in grupo:
                if otro != nombre and otro not in iguales:
                    iguales.append(otro)
    return {nombre: tuple(otros) for nombre, otros in tabla.items()}


EQUIVALENTES = _armar_equivalencias()


def _sinonimos_de(formas) -> tuple:
    """Los equivalentes de todas esas formas, sin repetir y en orden."""
    salida: list = []
    for forma in sorted(formas):
        for igual in EQUIVALENTES.get(forma, ()):
            if igual not in salida:
                salida.append(igual)
    return tuple(salida)


def emparejar_por_nombre(tabla: Dict[str, Any],
                         nombre: str) -> Optional[Dict[str, Any]]:
    """Busca una especie en una tabla indexada por nombre normalizado.

    La usan las TRES fuentes de ODEPA, y por eso vive una sola vez: cada una
    nombra distinto -el catalogo de curvas dice 'Cerezos', Aduanas 'Cereza' y
    CIREN 'Cerezo'- pero la escalera de intentos tiene que ser la misma, o la
    misma especie encontraria precio y no superficie sin que nada lo explique.

    De lo especifico a lo general, soltando calificativos de derecha a
    izquierda:

    1. Alguna forma del nombre -singular y plural- contra la clave.
    2. Los sinonimos del grupo: arbol y fruto son la misma especie y cada
       fuente elige uno. No hay «sentido»: se prueban todos.
    3. Que una clave empiece por alguna forma, o al reves: cubre 'Ciruelo' ->
       'Ciruelo Europeo' y 'Arandano' -> 'Arandano Americano'.
    """
    palabras = _clave(nombre).split()
    while palabras:
        parcial = " ".join(palabras)
        formas = formas_de_nombre(parcial)

        for forma in formas:
            if forma in tabla:
                return tabla[forma]

        for forma in formas:
            for igual in EQUIVALENTES.get(forma, ()):
                if igual in tabla:
                    return tabla[igual]

        # Prefijos en los dos sentidos, y sobre PALABRAS enteras: asi 'nuez' no
        # engancha con una clave que solo empiece con esas letras.
        #
        # Se prueban primero las formas escritas y DESPUES los sinonimos, en
        # ese orden: lo declarado tiene que ganarle a lo adivinado. El segundo
        # paso es el que hace que «Durazno» encuentre «Duraznero Consumo
        # Fresco» en CIREN, donde la clave no es ni una ni la otra.
        for candidatos in (formas, _sinonimos_de(formas)):
            for guardada in tabla:
                for forma in candidatos:
                    if len(forma) < 4:
                        continue
                    if (guardada.startswith(forma + " ")
                            or forma.startswith(guardada + " ")):
                        return tabla[guardada]

        palabras.pop()
    return None


@dataclass
class ProductoFOB:
    """Un codigo arancelario con su precio. El desglose ES el dato."""

    codigo: str
    producto: str
    toneladas: float
    usd_kg: float


class BibliotecaFOB:
    """Precios FOB por grupo, con su desglose por codigo arancelario.

    El desglose no es un detalle: el grupo «Avellana» mezcla con cascara a
    4,17 USD/kg y sin cascara a 9,99. Quien pregunte por el grupo recibe el
    promedio Y la advertencia; quien sepa que vende el cliente puede pedir el
    codigo.
    """

    def __init__(self, crudo: Dict[str, Any]):
        self.version = crudo.get("version")
        self.origen = crudo.get("origen")
        self.temporada = crudo.get("temporada_de_referencia")
        self.fuente = crudo.get("fuente_declarada")
        self.advertencia = crudo.get("advertencia")
        self._grupos = crudo.get("grupos", {})

    def _registro(self, nombre: str) -> Optional[Dict[str, Any]]:
        """El grupo de Aduanas que corresponde a este nombre.

        Se delega a `emparejar_por_nombre`, la MISMA busqueda que usan las
        otras dos fuentes. La primera version tenia aqui su propio recorte de
        la ultima letra y fallaba con 'Cerezos', 'Nogales' y 'Limoneros' -tres
        de las trece curvas del catalogo- porque el plural castellano no es
        una letra.
        """
        return emparejar_por_nombre(self._grupos, nombre)

    def fob_de(self, nombre: str) -> Optional[PrecioReferencia]:
        """El FOB del grupo. NUNCA listo para el flujo: falta la cadena."""
        registro = self._registro(nombre)
        if registro is None or registro.get("usd_kg") is None:
            return None
        return PrecioReferencia(
            nombre=registro.get("nombre", nombre),
            valor_usd_kg=float(registro["usd_kg"]),
            tipo=TIPO_FOB,
            fuente=registro.get("fuente"),
            periodo=registro.get("temporada"),
            unidad="USD/kg",
            confianza=registro.get("confianza", "media"),
            advertencia=registro.get("advertencia"),
        )

    def codigos_de(self, nombre: str) -> List[ProductoFOB]:
        """El desglose, de mayor a menor volumen."""
        registro = self._registro(nombre) or {}
        return [ProductoFOB(codigo=c["codigo"], producto=c["producto"],
                            toneladas=c["toneladas"], usd_kg=c["usd_kg"])
                for c in registro.get("codigos", []) if c.get("usd_kg")]

    def es_mezcla(self, nombre: str) -> bool:
        """Si el promedio del grupo no describe a ninguno de sus productos."""
        return bool((self._registro(nombre) or {}).get("es_mezcla"))

    def procesamiento_de(self, nombre: str) -> Optional[str]:
        """En que estado se exporto: «Fresca» o «Frutos secos».

        el agrónomo lo pidio explicito: el precio tiene que ser el de la fruta
        FRESCA, no el de conserva ni el de jugo, que son otro negocio y otro
        precio. El importador ya descarta lo demas, pero **eso no se veia en
        ninguna parte**: un dato correcto que no dice de donde sale obliga a
        confiar en vez de comprobar.
        """
        return (self._registro(nombre) or {}).get("procesamiento")

    def serie_de(self, nombre: str) -> Dict[str, float]:
        """El FOB temporada a temporada. Un precio sin su historia no dice si
        el ultimo anio fue normal o excepcional, y eso lo pregunta el banco."""
        return dict((self._registro(nombre) or {}).get("serie_usd_kg", {}))

    @property
    def grupos(self) -> List[str]:
        return sorted(r.get("nombre", k) for k, r in self._grupos.items())


def cargar_fob(ruta: Union[str, Path]) -> BibliotecaFOB:
    with Path(ruta).open(encoding="utf-8") as fh:
        return BibliotecaFOB(json.load(fh))
