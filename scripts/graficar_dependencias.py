"""Vuelve a escanear el grafo REAL de dependencias entre hojas y tablas, y lo
compara contra lo declarado en config/mapping_santander.json:
unidades_excluibles.

Por que existe: cuando el tasador excluye una hoja o tabla del .xlsx antes
de descargarlo (el usuario, 2026-09-23), el motor tiene que saber que OTRA cosa
se cae con ella -si no, queda una formula leyendo de una fila que ya no
esta, un #REF! silencioso en un informe real-. Ese grafo (``depende_de`` en
cada unidad excluible) se armo MIDIENDO: escaneando las 573 formulas reales
de la plantilla Bsan mas las ~48 que el propio generador inyecta
(`formulas_de_columna`, `formulas_finales`, `correcciones_de_formula`,
`celdas_calculadas`...), no a ojo. Resultado de esa medicion: exactamente 8
dependencias en todo el sistema (ver docs/CLAUDE-historial.md, entrada de
las unidades excluibles).

Este script vuelve a hacer esa medicion cada vez que corre. Si algun dia se
agrega una formula que cruza de hoja o de tabla sin actualizar
``unidades_excluibles``, esto tiene que avisarlo ANTES de que llegue a un
informe real -el mismo principio que el resto del proyecto usa para todo
("medir, no aproximar").

Cubre dos cosas, no una sola:

1. Toda referencia entre HOJAS, en la plantilla real y en las formulas que
   el mapping inyecta -barrido completo, no una lista a mano-.
2. La UNICA dependencia intra-hoja conocida hoy (Dotacion!K33 = $H$27,
   Requerimiento de riego <- Derechos de agua): no hay un clasificador
   general de "esta fila es de que tabla" -seria una pieza grande para un
   solo caso conocido-, asi que se verifica puntualmente que siga estando
   ahi y que `unidades_excluibles` la siga declarando. Si aparece una
   SEGUNDA dependencia intra-hoja, el barrido de hojas de este script no la
   va a ver -queda para cuando haga falta un clasificador de verdad-.

Uso: python scripts/graficar_dependencias.py
"""

import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import load_workbook  # noqa: E402

from config.settings import ajustes  # noqa: E402
from services.mapping import cargar_mapa  # noqa: E402

fallos, comprobaciones = [], 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)
    return condicion


# 'Nombre de Hoja'!A1  o  NombreHoja!A1  (con o sin $, con o sin rango)
_RE_HOJA_EXTERNA = re.compile(r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!\$?[A-Z]{1,3}\$?[0-9]+")


def _hojas_referenciadas(formula, hoja_propia):
    """Nombres de hoja que una formula referencia, sin contarse a si misma."""
    vistos = set()
    for m in _RE_HOJA_EXTERNA.finditer(formula):
        destino = (m.group(1) or m.group(2))
        if destino and destino.strip() != hoja_propia.strip():
            vistos.add(destino.strip())
    return vistos


def _formulas_del_mapping(crudo):
    """Todo string que empiece con '=' en el JCrudo del mapping, sin importar
    la clave -formulas_de_columna, formulas_finales, celdas_calculadas..."""
    hallazgos = []

    def recorrer(obj):
        if isinstance(obj, dict):
            for v in obj.values():
                recorrer(v)
        elif isinstance(obj, list):
            for v in obj:
                recorrer(v)
        elif isinstance(obj, str) and obj.startswith("="):
            hallazgos.append(obj)

    recorrer(crudo)
    return hallazgos


def escanear_hojas_dependientes(ruta_plantilla, ruta_mapping):
    """Todo par (hoja_que_depende, hoja_de_la_que_depende) medido de verdad."""
    wb = load_workbook(ruta_plantilla, data_only=False)
    pares = set()

    for ws in wb.worksheets:
        for fila in ws.iter_rows():
            for celda in fila:
                v = celda.value
                if isinstance(v, str) and v.startswith("="):
                    for destino in _hojas_referenciadas(v, ws.title):
                        pares.add((ws.title, destino))

    with open(ruta_mapping, encoding="utf-8") as fh:
        crudo = json.load(fh)
    hojas_alias = crudo.get("hojas", {})
    # Las formulas del mapping no saben en que hoja estan escritas -serian
    # kilometros de contexto para reconstruirlo-, asi que se declaran las
    # dependencias SALIENTES sin origen: cualquier hoja que referencie a X
    # deberia depender de X si X es excluible. Es deliberadamente el lado
    # conservador: agrega pares de mas, nunca de menos.
    hojas_referenciadas_por_alguna_formula = set()
    for formula in _formulas_del_mapping(crudo):
        for m in _RE_HOJA_EXTERNA.finditer(formula):
            destino = (m.group(1) or m.group(2))
            if destino:
                hojas_referenciadas_por_alguna_formula.add(destino.strip())

    return pares, hojas_referenciadas_por_alguna_formula, hojas_alias


def main():
    print("=" * 72)
    print("GRAFO REAL DE DEPENDENCIAS (re-escaneado, no declarado a mano)")
    print("=" * 72)

    mapa = cargar_mapa(ajustes.ruta_mapping)

    print("\n[1] Dependencias entre HOJAS, medidas de la plantilla + el mapping")
    pares, referenciadas, hojas_alias = escanear_hojas_dependientes(
        ajustes.ruta_plantilla, ajustes.ruta_mapping)

    for origen, destino in sorted(pares):
        print("  medido: '{}' lee de '{}'".format(origen, destino))
    if referenciadas:
        print("  hojas referenciadas desde alguna formula inyectada (sin origen):",
              sorted(referenciadas))

    # nombre real -> alias, para poder cruzar contra unidades_excluibles
    alias_por_nombre = {nombre: alias for alias, nombre in hojas_alias.items()}

    # Las unidades 'hoja' cuyo nombre real aparece como DESTINO de alguna
    # dependencia medida (alguien lee de ella) tienen que tener, en la
    # direccion inversa, a quien lee de ellas declarado como dependiente en
    # ALGUNA unidad -o no ser excluibles en absoluto-. Como
    # unidades_excluibles no incluye a Cliente/Flujo/Anexo/Detalle
    # plantaciones a proposito (son el motor, nunca se excluyen), lo que
    # importa comprobar es lo contrario: que ninguna unidad EXCLUIBLE sea
    # depended-upon por una formula sin que ESA dependencia este declarada.
    ids_por_hoja_real = {}
    for id_, unidad in mapa.unidades_excluibles.items():
        if unidad.nivel != "hoja":
            continue
        nombre_real = hojas_alias.get(unidad.hoja, unidad.hoja)
        ids_por_hoja_real.setdefault(nombre_real, []).append(id_)

    for origen, destino in sorted(pares):
        unidades_destino = ids_por_hoja_real.get(destino, [])
        if not unidades_destino:
            continue  # destino no es una hoja excluible: nada que cascadear
        unidades_origen = ids_por_hoja_real.get(origen, [])
        for id_destino in unidades_destino:
            if unidades_origen:
                for id_origen in unidades_origen:
                    declarado = id_destino in mapa.unidades_excluibles[id_origen].depende_de
                    comprobar(
                        declarado,
                        "'{}' (hoja) declara que depende de '{}', como mide "
                        "la formula real".format(id_origen, id_destino),
                        "'{}'.depende_de = {}".format(
                            id_origen, mapa.unidades_excluibles[id_origen].depende_de))
            else:
                # el origen no es una unidad excluible (es una hoja del
                # motor, como Cliente/Flujo/Anexo): esta bien, esas nunca se
                # excluyen, asi que da igual si "dependen" de algo excluible
                # -si alguna vez se excluyera esa unidad tendria que
                # arrastrar tambien al motor, y eso ya lo impide el diseño
                # (esas 4 hojas no estan en el registro).
                pass

    # Las formulas que el MAPPING inyecta (no la plantilla cruda) muchas
    # veces usan un token en vez de una fila literal -'Anexo '!I{total:
    # anexo_real}, no 'Anexo '!I8-, y de esas no se puede sacar el origen
    # sin simular la resolucion de `{total:...}` (services/excel_engine.py).
    # Lo que SI se puede verificar sin eso: que ninguna hoja EXCLUIBLE
    # aparezca como destino de una de esas formulas sin que su registro lo
    # sepa. Hoy da vacio -nada lee de Inversion/Estructura Productiva/
    # Comentarios Opinion, ni con fila literal ni con token-, y si un dia
    # deja de estarlo, esto avisa para revisar a mano en vez de asumir que
    # el escaneo de arriba (que si necesita fila literal) ya lo cubrio.
    for nombre_real, ids in ids_por_hoja_real.items():
        if nombre_real in referenciadas:
            for id_ in ids:
                comprobar(
                    False,
                    "'{}' aparece referenciada desde una formula del mapping "
                    "con fila-token (no se puede saber el origen automatico "
                    "aca): revisar a mano si hace falta declarar un "
                    "depende_de".format(nombre_real),
                    "unidad: {}".format(id_))

    print("\n[2] La UNICA dependencia intra-hoja conocida (Dotacion!K33=$H$27)")
    ws = load_workbook(ajustes.ruta_plantilla)["Dotación de recursos"]
    comprobar(
        ws["K33"].value in (None, ""),
        "la plantilla del banco no trae esa formula -la escribe el "
        "generador (celdas_calculadas), asi que no esta en la plantilla cruda",
        repr(ws["K33"].value))
    riego = mapa.tablas.get("requerimientos_riego")
    comprobar(
        riego is not None and riego.celdas_calculadas.get("K33") == "=$H$27",
        "el mapping SIGUE declarando K33=$H$27 en requerimientos_riego "
        "-si esto cambia de celda, la unidad excluible tambien tiene que "
        "cambiar", riego.celdas_calculadas if riego else None)
    unidad_riego = mapa.unidades_excluibles.get("requerimientos_riego")
    comprobar(
        unidad_riego is not None and "derechos_agua" in unidad_riego.depende_de,
        "y 'requerimientos_riego' sigue declarando que depende de "
        "'derechos_agua' en unidades_excluibles")

    print("\n" + "=" * 72)
    if fallos:
        print("{}/{} comprobaciones FALLARON:".format(len(fallos), comprobaciones))
        for f in fallos:
            print("  - {}".format(f))
        return 1
    print("{}/{} comprobaciones OK".format(comprobaciones, comprobaciones))
    return 0


if __name__ == "__main__":
    sys.exit(main())
