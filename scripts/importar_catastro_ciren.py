"""Superficie frutal de Chile, del Catastro Fruticola ODEPA-CIREN.

Es la base para poder DECIR ALGO del contexto del cliente, que hoy el informe
no dice. No es un dato del predio: es contra que se compara.

    «El cliente declara 760 ha de avellano en Biobio, donde el catastro
     registra 4.956 ha. Es el 15% de la region.»

Eso es una frase que el ejecutivo del banco entiende, y hasta ahora no habia de
donde sacarla. Sirve ademas para tres cosas mas:

- **Contrastar lo que declara el cliente.** Una superficie o una variedad que
  no existe en su region es una pregunta que hacer, no necesariamente un error.
- **El vocabulario real de las especies chilenas**, con sus variedades. 74
  especies con sus nombres tal como se usan aca, que es material para mejorar
  el emparejado de curvas y los sinonimos del importador de catastros.
- **Saber si una variedad esta en retirada o en expansion**, mirando la serie
  2004-2025.

Se baja del mismo boletin mensual que el FOB, archivo **«Base del catastro para
boletin»** (ver la cabecera de `importar_fob_odepa.py`).

    python scripts/importar_catastro_ciren.py "ruta/Base-catastro...xlsx"

**Ojo con la frescura.** CIREN levanta cada region por turnos, cada ~3 anios.
El archivo del boletin **no** expone eso: entrega la serie 2004-2025 con todas
las regiones al dia, o sea ya interpolada entre levantamientos. Asi que la
cifra de una region puede ser una proyeccion de un catastro de hace dos anios,
y el archivo no dice cual. Se guarda `anio_por_region` -que hoy da 2025 para
las catorce- para que si eso cambia se note, pero la advertencia de fondo va
igual: es superficie ESTIMADA a nivel regional, no medida en el predio.
"""

import argparse
import io
import json
import os
import sys
import unicodedata
from collections import defaultdict
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import load_workbook                                # noqa: E402

from config.settings import ajustes                               # noqa: E402

REGION, ESPECIE, VARIEDAD, ANIO, SUPERFICIE = 0, 1, 2, 3, 4

#: Debajo de esto una variedad no se guarda por separado. El nectarino tiene
#: 593 variedades y la mayoria son fracciones de hectarea: listarlas todas
#: hincharia el archivo sin decir nada. Se cuentan igual en el total.
MINIMO_HA_VARIEDAD = 5.0


def _clave(texto):
    sin = unicodedata.normalize("NFKD", str(texto or ""))
    sin = "".join(c for c in sin if not unicodedata.combining(c))
    return " ".join(sin.lower().split())


def leer(ruta):
    libro = load_workbook(ruta, data_only=True, read_only=True)
    hoja = libro.worksheets[0]
    filas = hoja.iter_rows(values_only=True)
    encabezado = next(filas)

    registros = []
    for f in filas:
        if not f or f[ESPECIE] is None or f[ANIO] is None:
            continue
        registros.append((
            str(f[REGION]).strip(), str(f[ESPECIE]).strip(),
            str(f[VARIEDAD]).strip(), int(f[ANIO]),
            float(f[SUPERFICIE] or 0.0)))
    libro.close()
    return registros, list(encabezado)


def armar(registros, origen, encabezado):
    # El ultimo anio levantado en CADA region, que no es el mismo para todas.
    ultimo_de_region = defaultdict(int)
    for region, _e, _v, anio, _s in registros:
        ultimo_de_region[region] = max(ultimo_de_region[region], anio)

    especies = {}
    serie_nacional = defaultdict(lambda: defaultdict(float))
    for region, especie, variedad, anio, superficie in registros:
        serie_nacional[especie][anio] += superficie
        if anio != ultimo_de_region[region]:
            continue
        clave = _clave(especie)
        e = especies.setdefault(clave, {
            "nombre": especie,
            "hectareas": 0.0,
            "regiones": {},
            "variedades": defaultdict(float),
            "variedades_totales": set(),
        })
        e["hectareas"] += superficie
        e["regiones"][region] = round(
            e["regiones"].get(region, 0.0) + superficie, 2)
        e["variedades"][variedad] += superficie
        e["variedades_totales"].add(variedad)

    salida = {}
    for clave, e in especies.items():
        variedades = sorted(
            ({"nombre": v, "hectareas": round(ha, 2)}
             for v, ha in e["variedades"].items() if ha >= MINIMO_HA_VARIEDAD),
            key=lambda x: -x["hectareas"])
        salida[clave] = {
            "nombre": e["nombre"],
            "hectareas": round(e["hectareas"], 2),
            "regiones": dict(sorted(e["regiones"].items(),
                                    key=lambda kv: -kv[1])),
            "variedades": variedades,
            "variedades_listadas": len(variedades),
            "variedades_en_el_catastro": len(e["variedades_totales"]),
            "serie_nacional_ha": {
                str(a): round(ha, 2)
                for a, ha in sorted(serie_nacional[e["nombre"]].items())},
            "fuente": "Catastro Fruticola ODEPA-CIREN",
        }

    return {
        "version": "1.0",
        "generado": datetime.now().isoformat(timespec="seconds"),
        "origen": os.path.basename(origen),
        "encabezado_del_archivo": encabezado,
        "fuente_declarada": (
            "Catastro Fruticola ODEPA-CIREN, via el Boletin de la fruta."),
        "minimo_ha_variedad": MINIMO_HA_VARIEDAD,
        "anio_por_region": dict(sorted(ultimo_de_region.items())),
        "advertencia": (
            "Superficie regional ESTIMADA, no medida en el predio. CIREN "
            "levanta cada region por turnos cada ~3 anios y el boletin entrega "
            "la serie ya interpolada, asi que la cifra de una region puede ser "
            "una proyeccion. Sirve para comparar el orden de magnitud de lo "
            "que declara el cliente, no para corregirlo."),
        "metodo": (
            "Por especie se suma la superficie comercial del ULTIMO anio "
            "levantado en cada region. Las variedades bajo {} ha no se listan "
            "por separado pero si se cuentan en el total.".format(
                MINIMO_HA_VARIEDAD)),
        "especies": dict(sorted(salida.items())),
    }


def main():
    p = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    p.add_argument("archivo", help="el «Base del catastro para boletin»")
    p.add_argument("--salida", default=None)
    args = p.parse_args()

    if not os.path.exists(args.archivo):
        raise SystemExit("  No existe el archivo: {}".format(args.archivo))

    print("=" * 74)
    print("CATASTRO FRUTICOLA ODEPA-CIREN")
    print("=" * 74)
    registros, encabezado = leer(args.archivo)
    print("  {:,} filas leidas".format(len(registros)))

    datos = armar(registros, args.archivo, encabezado)

    salida = args.salida or str(ajustes.raiz / "config" / "catastro_ciren.json")
    with io.open(salida, "w", encoding="utf-8") as fh:
        json.dump(datos, fh, ensure_ascii=False, indent=2)
        fh.write("\n")

    anios = sorted(set(datos["anio_por_region"].values()))
    print("  {} especies · {} regiones · levantamientos de {} a {}".format(
        len(datos["especies"]), len(datos["anio_por_region"]),
        anios[0], anios[-1]))
    print()
    print("  {:<26} {:>12} {:>8} {:>10}".format(
        "especie", "ha", "reg.", "variedades"))
    print("  " + "-" * 60)
    for e in sorted(datos["especies"].values(),
                    key=lambda x: -x["hectareas"])[:14]:
        print("  {:<26} {:>12,.0f} {:>8} {:>10}".format(
            e["nombre"][:26], e["hectareas"], len(e["regiones"]),
            e["variedades_en_el_catastro"]))

    print()
    print("  Anio del ultimo levantamiento por region:")
    for region, anio in sorted(datos["anio_por_region"].items(),
                               key=lambda kv: (-kv[1], kv[0])):
        print("    {:<34} {}".format(region[:34], anio))

    print()
    print("  Escrito en {} ({:.0f} KB)".format(
        salida, os.path.getsize(salida) / 1024))
    return 0


if __name__ == "__main__":
    sys.exit(main())
