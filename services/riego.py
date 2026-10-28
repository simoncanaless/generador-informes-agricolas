"""Requerimiento hidrico en el mes de maxima demanda (enero).

Aritmetica pura, como todo el 80% determinista del proyecto. La formula y los
parametros salen de la planilla propia de Fitogreen
('45. AGRICOLA EJEMPLO (RER HIDRICO-TASACION)'), no de un manual generico:

    demanda (L/s/ha) = ETo x Kc x 10 / eficiencia / 86,4 x cobertura
    demanda total    = demanda por hectarea x superficie

De donde sale cada pieza:

- **ETo**: evapotranspiracion de referencia de enero, en mm/dia, por COMUNA.
  La tabla ya viene con un 15% de margen incorporado (la planilla original
  calcula mm_del_mes / 31 x 1,15), asi que no hay que volver a mayorarla.
- **Kc**: coeficiente del cultivo, por especie.
- **eficiencia**: del sistema de riego (goteo 0,9, surco 0,6...). Divide porque
  un sistema que pierde agua necesita entregar mas para poner lo mismo en la
  raiz.
- **cobertura**: cuanto suelo sombrea la planta, segun los anios que lleva
  plantada. Un huerto de un anio consume una fraccion de lo que consume el
  mismo huerto adulto, y por eso la demanda de una plantacion nueva no se
  calcula con la superficie a pleno.
- **86,4**: convierte mm/dia sobre una hectarea a litros por segundo. Un
  milimetro sobre una hectarea son 10 m3, y el dia tiene 86.400 segundos.

Verificado contra el ejemplo resuelto de la propia planilla (hoja 'Hoja4'):
palto con ETo 7,02, Kc 0,8 y microaspersion 0,8 da 0,8125 L/s/ha.
"""

from __future__ import annotations

import json
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

__all__ = [
    "ParametrosRiego", "DemandaPlantacion", "BalanceHidrico",
    "cargar_parametros", "demanda_por_hectarea", "balance_de_predio",
    "filas_de_requerimiento",
]

#: mm/dia sobre 1 ha -> L/s. 1 mm x 1 ha = 10 m3; el dia tiene 86.400 s.
CONSTANTE_CONVERSION = 86.4

#: Desde este anio la plantacion se considera adulta (cobertura plena).
ANIO_ADULTO = 4


def _clave(texto: Any) -> str:
    plano = "".join(
        c for c in unicodedata.normalize("NFD", str(texto or "").strip())
        if unicodedata.category(c) != "Mn"
    )
    return " ".join(plano.lower().split())


@dataclass
class ParametrosRiego:
    """La biblioteca de riego: ETo por comuna, Kc por especie, eficiencias."""

    kc_por_especie: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    eficiencia_por_sistema: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    cobertura_por_anio: Dict[str, float] = field(default_factory=dict)
    eto_por_comuna: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    comunas_por_region: Dict[str, List[str]] = field(default_factory=dict)
    cuencas: List[str] = field(default_factory=list)
    version: Optional[str] = None
    origen: Optional[str] = None

    # -- consultas, todas tolerantes a tildes y mayusculas -----------------

    def eto(self, comuna: str) -> Optional[float]:
        """ETo diaria de enero (mm/dia) de una comuna. None si no esta."""
        registro = self.eto_por_comuna.get(_clave(comuna))
        return registro["diario_mm"] if registro else None

    def kc(self, especie: str) -> Optional[float]:
        """Kc de una especie, tolerando como la escribe el agricultor.

        La tabla de el agrónomo dice 'Avellano'; el catastro del cliente decia
        'Avellano europeo'. Con la busqueda exacta el Kc quedaba vacio, y con
        el Kc vacio TODA la fila de riego daba 0,00 L/s -sin que nada fallara,
        porque el dato si estaba en la biblioteca-. Por eso se va soltando el
        calificativo de derecha a izquierda hasta dar con la especie base:
        'Avellano europeo' -> 'avellano', 'Palto (rebajado)' -> 'palto'.

        Solo actua cuando la clave exacta NO esta, asi que una especie que la
        tabla distinga por su calificativo sigue mandando sobre la generica.
        """
        registro = self.kc_por_especie.get(_clave(especie))
        if registro:
            return registro["kc"]

        # El plural y el nombre del fruto, con las MISMAS reglas que usan el
        # emparejador de curvas y las fuentes de ODEPA. Un catastro que dice
        # 'Nogales' o 'Limonero' encontraba su curva -el flujo lo proyectaba-
        # pero no su Kc ('nogal', 'limon'), asi que en el mismo informe esas
        # hectareas pedian 0 L/s. Solo nombres exactos, sin prefijos: 'Vid' no
        # se resuelve sola entre 'uva de mesa' y 'uva vinifera', y adivinar
        # daria un Kc con pinta de dato.
        from services.catastro_import import formas_de_nombre
        from services.precios import EQUIVALENTES

        base = _clave(especie).split("(")[0].split(" - ")[0].strip()
        palabras = base.split()
        while palabras:
            formas = sorted(formas_de_nombre(" ".join(palabras)))
            candidatas = formas + [igual for forma in formas
                                   for igual in EQUIVALENTES.get(forma, ())]
            for candidata in candidatas:
                registro = self.kc_por_especie.get(candidata)
                if registro:
                    return registro["kc"]
            palabras.pop()
        return None

    def eficiencia(self, sistema: str) -> Optional[float]:
        registro = self.eficiencia_por_sistema.get(_clave(sistema))
        return registro["eficiencia"] if registro else None

    def cobertura(self, anios_desde_plantacion: Optional[int]) -> float:
        """Fraccion de desarrollo del cultivo. Sin edad se asume adulta.

        Asumir adulto es lo conservador para el banco: da la demanda maxima, y
        una plantacion sin anio declarado no puede justificar pedir menos agua.
        """
        if anios_desde_plantacion is None:
            return 1.0
        if anios_desde_plantacion >= ANIO_ADULTO:
            return 1.0
        clave_anio = str(max(int(anios_desde_plantacion), 1))
        return float(self.cobertura_por_anio.get(clave_anio, 1.0))

    def region_de(self, comuna: str) -> Optional[str]:
        objetivo = _clave(comuna)
        for region, comunas in self.comunas_por_region.items():
            if any(_clave(c) == objetivo for c in comunas):
                return region
        return None


def cargar_parametros(ruta: Union[str, Path]) -> ParametrosRiego:
    with Path(ruta).open(encoding="utf-8") as fh:
        crudo = json.load(fh)
    return ParametrosRiego(
        kc_por_especie=crudo.get("kc_por_especie", {}),
        eficiencia_por_sistema=crudo.get("eficiencia_por_sistema", {}),
        cobertura_por_anio=crudo.get("cobertura_por_anio", {}),
        eto_por_comuna=crudo.get("eto_por_comuna", {}),
        comunas_por_region=crudo.get("comunas_por_region", {}),
        cuencas=crudo.get("cuencas", []),
        version=crudo.get("version"),
        origen=crudo.get("origen"),
    )


def demanda_por_hectarea(
    eto_mm_dia: float, kc: float, eficiencia: float, cobertura: float = 1.0
) -> float:
    """L/s por hectarea en el mes de maxima demanda."""
    if eficiencia <= 0:
        raise ValueError("la eficiencia de riego debe ser mayor que 0")
    return round(
        eto_mm_dia * kc * 10.0 / eficiencia / CONSTANTE_CONVERSION * cobertura, 6
    )


@dataclass
class DemandaPlantacion:
    """Lo que pide una plantacion, con los parametros que se usaron a la vista.

    Se guarda cada insumo -no solo el resultado- porque el banco puede preguntar
    por que este huerto pide mas agua que el de al lado, y la respuesta tiene
    que estar a mano.
    """

    especie: str
    superficie_ha: float
    sistema_riego: Optional[str] = None
    anios_desde_plantacion: Optional[int] = None
    eto_mm_dia: Optional[float] = None
    kc: Optional[float] = None
    eficiencia: Optional[float] = None
    cobertura: float = 1.0
    demanda_l_s_ha: Optional[float] = None
    demanda_total_l_s: Optional[float] = None
    faltantes: List[str] = field(default_factory=list)

    @property
    def calculable(self) -> bool:
        return not self.faltantes


@dataclass
class BalanceHidrico:
    """Demanda de todas las plantaciones frente al caudal disponible."""

    comuna: Optional[str] = None
    plantaciones: List[DemandaPlantacion] = field(default_factory=list)
    caudal_disponible_l_s: Optional[float] = None
    avisos: List[str] = field(default_factory=list)

    @property
    def demanda_total_l_s(self) -> float:
        return round(sum(p.demanda_total_l_s or 0.0 for p in self.plantaciones), 4)

    @property
    def superficie_calculada_ha(self) -> float:
        return round(
            sum(p.superficie_ha for p in self.plantaciones if p.calculable), 4)

    @property
    def deficit_l_s(self) -> Optional[float]:
        """Cuanto caudal falta. Cero o negativo significa que alcanza."""
        if self.caudal_disponible_l_s is None:
            return None
        return round(self.demanda_total_l_s - self.caudal_disponible_l_s, 4)

    @property
    def suficiente(self) -> Optional[bool]:
        deficit = self.deficit_l_s
        return None if deficit is None else deficit <= 0

    @property
    def deficit_pct(self) -> Optional[float]:
        """Que fraccion de la demanda queda sin cubrir."""
        if self.caudal_disponible_l_s is None or self.demanda_total_l_s <= 0:
            return None
        return round(
            max(self.demanda_total_l_s - self.caudal_disponible_l_s, 0.0)
            / self.demanda_total_l_s * 100, 2
        )


def balance_de_predio(
    plantaciones,
    parametros: ParametrosRiego,
    comuna: Optional[str] = None,
    eto_mm_dia: Optional[float] = None,
    caudal_disponible_l_s: Optional[float] = None,
    anio_informe: Optional[int] = None,
) -> BalanceHidrico:
    """Calcula la demanda de una cartera y la compara con el agua disponible.

    ``plantaciones`` son objetos con especie, superficie_ha, riego y
    anio_plantacion: el ``Plantacion`` del dominio calza tal cual.

    Lo que no se puede calcular NO se estima: se deja en ``faltantes`` con el
    nombre del parametro que falto. Una demanda inventada por defecto es peor
    que un hueco, porque el hueco se ve.
    """
    from services.catastro_import import anio_inicial

    balance = BalanceHidrico(comuna=comuna,
                             caudal_disponible_l_s=caudal_disponible_l_s)

    eto = eto_mm_dia if eto_mm_dia is not None else (
        parametros.eto(comuna) if comuna else None)
    if eto is None:
        balance.avisos.append(
            "sin ETo: {}".format(
                "no hay comuna declarada" if not comuna
                else "la comuna '{}' no esta en la tabla".format(comuna))
        )

    for plantacion in plantaciones:
        especie = getattr(plantacion, "especie", "") or ""
        sistema = getattr(plantacion, "riego", None)
        plantado = anio_inicial(getattr(plantacion, "anio_plantacion", None))
        edad = (anio_informe - plantado
                if anio_informe is not None and plantado is not None else None)

        kc = parametros.kc(especie)
        eficiencia = parametros.eficiencia(sistema) if sistema else None
        cobertura = parametros.cobertura(edad)

        detalle = DemandaPlantacion(
            especie=especie,
            superficie_ha=float(getattr(plantacion, "superficie_ha", 0.0) or 0.0),
            sistema_riego=sistema,
            anios_desde_plantacion=edad,
            eto_mm_dia=eto, kc=kc, eficiencia=eficiencia, cobertura=cobertura,
        )
        if eto is None:
            detalle.faltantes.append("ETo")
        if kc is None:
            detalle.faltantes.append("Kc de '{}'".format(especie))
        if eficiencia is None:
            detalle.faltantes.append(
                "eficiencia de '{}'".format(sistema) if sistema
                else "sistema de riego")

        if detalle.calculable:
            detalle.demanda_l_s_ha = demanda_por_hectarea(
                eto, kc, eficiencia, cobertura)
            detalle.demanda_total_l_s = round(
                detalle.demanda_l_s_ha * detalle.superficie_ha, 4)
        else:
            balance.avisos.append(
                "'{}' ({} ha) sin calcular: falta {}".format(
                    especie, detalle.superficie_ha, ", ".join(detalle.faltantes))
            )
        balance.plantaciones.append(detalle)

    return balance


def filas_de_requerimiento(balance: "BalanceHidrico", plantaciones) -> list:
    """Traduce un balance a las filas de la hoja 'Dotacion de recursos'.

    Devuelve ``RequerimientoRiego`` del dominio, listos para que el motor los
    vuelque. Los dos requerimientos se calculan igual aqui, pero al Excel van
    como formula: los numeros de Python sirven de contraste, no de contenido.

    Las plantaciones sin parametros completos igual bajan a la tabla con lo que
    se sabe de ellas. Dejarlas fuera las haria desaparecer del informe sin que
    nadie lo note; con la fila escrita y las celdas de parametro vacias, el
    hueco se ve.
    """
    from schemas.informe import RequerimientoRiego

    filas = []
    for detalle, plantacion in zip(balance.plantaciones, plantaciones):
        variedad = getattr(plantacion, "variedad", None)
        etiqueta = "{} {}".format(detalle.especie, variedad).strip() \
            if variedad else detalle.especie
        filas.append(RequerimientoRiego(
            especie_variedad=etiqueta or "(sin especie)",
            anio_plantacion=getattr(plantacion, "anio_plantacion", None),
            superficie_ha=detalle.superficie_ha,
            sistema_riego=detalle.sistema_riego,
            eficiencia_riego=detalle.eficiencia,
            eto_maximo_mm_enero=detalle.eto_mm_dia,
            coef_cultivo_kc_enero=detalle.kc,
            cobertura=detalle.cobertura,
            requerimiento_max_ls_ha=detalle.demanda_l_s_ha,
            requerimiento_total_plantacion_ls=detalle.demanda_total_l_s,
            densidad_plantacion=getattr(plantacion, "densidad_plantacion", None),
        ))
    return filas
