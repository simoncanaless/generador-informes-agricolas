"""Los numeros de la hoja Supuestos, propuestos desde la biblioteca.

El formato final de Santander mueve el bloque de flujo de cada plantacion con
una fila de Supuestos: % exportado, precios, costo fijo y variable, y el
rendimiento de cada uno de los 8 anios proyectados. En el informe de origen
el agrónomo los escribio a mano, plantacion por plantacion. Su respuesta a como
debia hacerlo el programa (2026-10-14): **"opcion a"** -que el programa
proponga todos los numeros con la biblioteca y el los corrija en el Excel-.

Lo que se propone y de donde sale:

- **Rendimiento de cada anio**: la curva de la especie a la EDAD que la
  plantacion cumple esa temporada (anio de la temporada menos anio de
  plantacion, la misma regla que el motor de envejecimiento), con los ajustes
  que el tasador anoto en la visita. Redondeado al kilo.
- **% exportado y precios**: los parametros comerciales de la curva. El precio
  de exportacion es el retorno al productor (el FOB menos su cadena de
  descuentos, si la curva la trae). El formato distingue el primer anio de los
  siguientes; la biblioteca no, asi que van iguales y el tasador los separa.
- **Costo fijo y variable**: la estructura de costos de la curva. El fijo es el
  de la edad que la plantacion tiene al final de la proyeccion; y si la curva
  trae costos por edad y la plantacion esta en formacion en los primeros anios,
  se llenan las columnas de costo en formacion que el formato trae para eso.

Nada se inventa: sin curva para la especie, la fila queda vacia y se avisa.
"""

from __future__ import annotations

from typing import Dict, List, Optional, Sequence

from schemas.curvas import CurvaProduccion, SolicitudProyeccion
from schemas.informe import Plantacion, SupuestoPlantacion
from services.catastro_import import anio_inicial
from services.curvas import costo_base_a_edad, proyectar
from services.envejecimiento import _coincide, anio_de_temporada, emparejar_curva

__all__ = ["proponer_supuestos"]


def _ajustes_de(curva: CurvaProduccion, por_curva: Dict[str, list]) -> list:
    """Los ajustes que el tasador anoto para esta curva -por id, nombre o
    especie, como los busca el motor de envejecimiento-."""
    for clave in (curva.id, curva.nombre, curva.especie):
        if por_curva.get(clave):
            return list(por_curva[clave])
    for clave, lista in por_curva.items():
        if _coincide(clave, curva.especie) or _coincide(clave, curva.nombre):
            return list(lista)
    return []


def _redondo(valor: Optional[float], decimales: int = 0) -> Optional[float]:
    if valor is None:
        return None
    return round(valor, decimales) if decimales else float(round(valor))


def proponer_supuestos(
    plantaciones: Sequence[Plantacion],
    curvas: Sequence[CurvaProduccion],
    temporadas: Sequence[str],
    ajustes_por_curva: Optional[Dict[str, list]] = None,
    formacion_hasta=None,
) -> List[str]:
    """Llena ``plantacion.supuestos`` de cada plantacion que no los traiga.

    ``temporadas``: las 8 de la proyeccion, en orden. ``formacion_hasta``:
    funcion especie -> ultima edad en formacion (las reglas de madurez);
    sin ella no se llenan los costos en formacion. Devuelve los avisos.
    """
    avisos: List[str] = []
    por_curva = ajustes_por_curva or {}
    anios = [anio_de_temporada(t) for t in temporadas]

    for plantacion in plantaciones:
        if plantacion.supuestos is not None:
            continue
        curva = emparejar_curva(plantacion, curvas)
        if curva is None:
            avisos.append(
                "Supuestos: '{}' no tiene curva en la biblioteca; su fila queda "
                "para escribirla a mano".format(plantacion.especie))
            continue
        plantado = anio_inicial(plantacion.anio_plantacion)
        if plantado is None:
            avisos.append(
                "Supuestos: '{} {}' no trae año de plantación; sus rendimientos "
                "quedan para escribirlos a mano".format(
                    plantacion.especie, plantacion.variedad or ""))

        ajustes = _ajustes_de(curva, por_curva)
        comerciales = curva.comerciales
        variable = curva.costos.costo_por_kg_usd
        por_edad = any(p.costo_directo_usd_ha is not None for p in curva.puntos)

        rendimientos: List[Optional[float]] = []
        edades: List[Optional[int]] = []
        for anio in anios:
            if plantado is None or anio is None:
                rendimientos.append(None)
                edades.append(None)
                continue
            edad = anio - plantado
            edades.append(edad)
            if edad < 0:
                rendimientos.append(0.0)
                continue
            punto = proyectar(curva, SolicitudProyeccion(
                curva_id=curva.id, superficie_ha=1.0, edad_actual=edad,
                ajustes=ajustes)).puntos[0]
            rendimientos.append(_redondo(punto.rendimiento_ajustado_kg_ha))

        retorno = comerciales.retorno_exportacion_usd_kg()
        interno = comerciales.rnp_mercado_interno_usd_kg
        sin_precios = not retorno and not interno
        if sin_precios:
            avisos.append(
                "Supuestos: la curva '{}' no tiene precios; '{}' queda sin "
                "precios en la hoja".format(curva.nombre, plantacion.especie))

        ultima = next((e for e in reversed(edades) if e is not None and e >= 0), None)
        fijo = _redondo(costo_base_a_edad(curva, ultima)) if ultima is not None \
            else _redondo(curva.costos.costo_base_usd_ha)

        supuesto = SupuestoPlantacion(
            pct_exportacion=round(comerciales.pct_exportacion, 4),
            precio_exportacion_usd_kg=None if sin_precios else round(retorno, 2),
            precio_exportacion_siguientes_usd_kg=None if sin_precios else round(retorno, 2),
            precio_mercado_interno_usd_kg=None if sin_precios else round(interno, 2),
            costo_fijo_usd_ha=fijo or None,
            costo_variable_usd_kg=round(variable, 4) if variable else None,
            rendimientos_kg_ha=rendimientos,
            curva=curva.nombre,
        )

        # Los costos en formacion: solo si la curva trae costos por edad (si
        # no, la formula del flujo ya da lo mismo con el fijo y el variable).
        limite = formacion_hasta(plantacion.especie) if (formacion_hasta and por_edad) else None
        if limite is not None:
            for k, campo in ((0, "costo_formacion_1_usd_ha"), (1, "costo_formacion_2_usd_ha")):
                edad = edades[k] if k < len(edades) else None
                if edad is not None and 0 <= edad <= limite:
                    rend = rendimientos[k] or 0.0
                    setattr(supuesto, campo, _redondo(costo_base_a_edad(curva, edad) + rend * variable))
            edad = edades[2] if len(edades) > 2 else None
            if edad is not None and 0 <= edad <= limite:
                supuesto.costo_fijo_formacion_3_usd_ha = _redondo(costo_base_a_edad(curva, edad))
                supuesto.costo_variable_formacion_3_usd_kg = round(variable, 4) if variable else None

        plantacion.supuestos = supuesto
    return avisos
