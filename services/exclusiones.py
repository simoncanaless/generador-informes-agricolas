"""Que pasa cuando el tasador saca una hoja o tabla del informe antes de
descargarlo.

el usuario: *"en vez de hacer distintos tipos de informes... que lo eliminen
antes de descargar el excel nomas"*. El riesgo real no es la pantalla, es
que una tabla que se saca puede dejar sin de donde leer a otra formula
-`config/mapping_santander.json:unidades_excluibles` declara ese grafo, y
este modulo lo recorre-.

El grafo se midio, no se adivino: escaneando las 573 formulas reales de la
plantilla Bsan mas las que el propio generador inyecta se encontraron
exactamente 8 dependencias en todo el sistema. Ver
`scripts/graficar_dependencias.py`, que vuelve a escanear cada vez que
corre -si algun dia se agrega una formula que cruza de tabla sin declararla
aqui, ese script avisa antes de que llegue a un informe real.
"""

from __future__ import annotations

from typing import Dict, List, Tuple

from services.mapping import EspecUnidadExcluible


def resolver_cascada(
    unidades: Dict[str, EspecUnidadExcluible], pedidas: List[str]
) -> Tuple[List[str], List[str]]:
    """El cierre transitivo de lo que hay que excluir, y por que.

    Si el tasador pide excluir X y otra unidad Y lee celdas de X en sus
    formulas (``Y.depende_de`` incluye a X), Y se agrega sola: dejarla con
    una formula que ya no tiene de donde leer seria peor que sacarla
    tambien -y es exactamente el error que el resto del proyecto evita en
    todos lados (ningun ``#REF!`` silencioso).

    Devuelve ``(a_excluir, avisos)``. Los avisos solo nombran lo que se
    agrego POR ARRASTRE -lo que el tasador pidio a proposito no necesita
    explicacion-, y una unidad que no existe en el registro se avisa y se
    ignora, nunca revienta la generacion.
    """
    dependientes: Dict[str, List[str]] = {}
    for id_, unidad in unidades.items():
        for otra in unidad.depende_de:
            dependientes.setdefault(otra, []).append(id_)

    resultado: List[str] = []
    vistos = set()
    avisos: List[str] = []
    pendientes = list(pedidas)

    while pendientes:
        actual = pendientes.pop(0)
        if actual in vistos:
            continue
        if actual not in unidades:
            avisos.append("no existe la unidad excluible '{}'".format(actual))
            continue
        vistos.add(actual)
        resultado.append(actual)
        for dependiente in dependientes.get(actual, []):
            if dependiente in vistos or dependiente in pendientes:
                continue
            if dependiente not in pedidas:
                avisos.append(
                    "'{}' se excluye también porque depende de '{}'".format(
                        unidades[dependiente].etiqueta, unidades[actual].etiqueta
                    )
                )
            pendientes.append(dependiente)

    return resultado, avisos
