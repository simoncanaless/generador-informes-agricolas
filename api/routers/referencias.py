"""Lo que ODEPA sabe de una especie: precio de exportacion, de mercado interno
y superficie plantada en Chile.

Es informacion de la casa, como las curvas: la ven todos los que entran. No
cuelga de ningun informe porque no es de nadie -es publica- y sirve tanto para
armar una curva como para escribir la opinion profesional.
"""

from __future__ import annotations

from dataclasses import asdict
from functools import lru_cache
from typing import Any, Dict, List, Optional

from fastapi import Depends, APIRouter, HTTPException, Query, status

from api.dependencias import obtener_usuario_actual
from config.settings import obtener_ajustes
from services.referencias import cargar_referencias

router = APIRouter(prefix="/referencias", tags=["referencias"],
                   dependencies=[Depends(obtener_usuario_actual)])


@lru_cache(maxsize=1)
def _biblioteca():
    """Se lee una vez: son tres JSON que solo cambian cuando alguien corre un
    importador, y releerlos en cada consulta seria trabajo para nada."""
    return cargar_referencias(obtener_ajustes().raiz / "config")


@router.get("", summary="Que especies hay y de que fuentes")
async def indice(
    parecidas: Optional[str] = Query(
        None, description="para proponer especies cuando no se encontro"),
    cuantas: int = Query(
        6, ge=1, le=12,
        description="cuantas proponer. El buscador del campo Especie pide 8; "
                    "el «¿es alguna de estas?» de ODEPA se queda en 6"),
) -> Dict[str, Any]:
    biblioteca = _biblioteca()
    return {
        "especies": biblioteca.especies,
        "fuentes": biblioteca.fuentes,
        # Se propone, no se decide: el que sabe si un «ciruelo japonés» es lo
        # mismo que un «ciruelo» es el agrónomo, no la tabla.
        "parecidas": (biblioteca.parecidas(parecidas, cuantas)
                       if parecidas else []),
    }


@router.get("/{especie}", summary="Todo lo que hay de una especie")
async def de_la_especie(
    especie: str,
    region: Optional[str] = Query(None, description="para el peso relativo"),
    hectareas: Optional[float] = Query(
        None, description="las del cliente, para decir que fraccion es"),
    tipo_cambio: Optional[float] = Query(
        None, description="para convertir el precio interno a USD"),
) -> Dict[str, Any]:
    referencias = _biblioteca().de(especie)
    if not referencias.hay_algo:
        raise HTTPException(
            status.HTTP_404_NOT_FOUND,
            "no hay referencias de «{}». Puede ser que ODEPA la nombre "
            "distinto: pruebe con el nombre del fruto.".format(especie))

    salida: Dict[str, Any] = {"especie": especie}

    def _resumen(r) -> Optional[Dict[str, Any]]:
        """El resumen viaja entero -sus puntos incluidos- porque la pantalla
        dibuja EXACTAMENTE la ventana que resume. Mandar la serie completa y
        recortarla en el navegador seria tener el criterio en dos sitios."""
        if r is None:
            return None
        salida_r = asdict(r)
        salida_r["parcial"] = r.parcial      # es property, asdict no la trae
        return salida_r

    if referencias.fob:
        salida["exportacion"] = {
            "nombre": referencias.fob.nombre,
            "usd_kg": referencias.fob.valor_usd_kg,
            "temporada": referencias.fob.periodo,
            "fuente": referencias.fob.fuente,
            "confianza": referencias.fob.confianza,
            "advertencia": referencias.fob.advertencia,
            # Que NO se pueda usar directo es parte del dato, no una nota al
            # pie: un FOB puesto en la fila del flujo duplica el ingreso.
            "listo_para_el_flujo": referencias.fob.listo_para_el_flujo,
            "es_mezcla": referencias.fob_es_mezcla,
            # «Fresca» o «Frutos secos», nunca conserva ni jugo: lo pidio
            # el agrónomo y hasta ahora era cierto pero invisible.
            "procesamiento": referencias.fob_procesamiento,
            "codigos": [
                {"codigo": c.codigo, "producto": c.producto,
                 "toneladas": c.toneladas, "usd_kg": c.usd_kg}
                for c in referencias.fob_codigos],
            "serie": referencias.fob_serie,
            # Minimo, maximo y promedio de las ultimas temporadas: un precio
            # suelto no dice si el anio fue normal o excepcional.
            "resumen": _resumen(referencias.resumen_fob()),
        }

    if referencias.interno:
        i = referencias.interno
        salida["mercado_interno"] = {
            "nombre": i.nombre,
            "clp_kg": i.clp_kg,
            "usd_kg": i.usd_kg(tipo_cambio),
            "anio": i.anio,
            "meses": i.meses,
            "punto_de_venta": i.punto_de_venta,
            "por_mes": i.por_mes,
            "otros_puntos": i.otros_puntos,
            "advertencia": i.advertencia,
            "confianza": i.confianza,
            # El mayorista NO es lo que recibe el agricultor: entre medio hay
            # flete, comision y merma. El ajuste lo dio el agrónomo.
            "al_productor": i.al_productor(),
        }

    if referencias.superficie:
        s = referencias.superficie
        salida["superficie"] = {
            "nombre": s.nombre,
            "hectareas_chile": s.hectareas,
            "regiones": s.regiones,
            "en_la_region": s.en_la_region(region),
            "peso_del_cliente": s.peso_del_cliente(hectareas or 0, region),
            "variedades": s.variedades[:12],
            "variedades_en_el_catastro": s.variedades_en_el_catastro,
            "serie_nacional_ha": s.serie_nacional_ha,
            # La ultima decada, que es lo que dice si la especie viene en
            # expansion o en retirada. La serie entera llega a 2004.
            "decada": _resumen(s.decada()),
        }

    return salida
