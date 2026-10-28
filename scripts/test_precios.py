"""Prueba de la biblioteca de precios de referencia.

Lo que se protege aqui es sobre todo una distincion: un precio FOB y un precio
a productor NO son intercambiables. Confundirlos infla el ingreso del cliente
al doble y el error no se ve en ninguna parte del informe.

Uso: python scripts/test_precios.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.settings import ajustes as config  # noqa: E402
from services.precios import cargar_precios  # noqa: E402

fallos, comprobaciones = [], 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)


def casi(a, b, tol=0.0001):
    return a is not None and abs(float(a) - float(b)) < tol


def biblioteca():
    return cargar_precios(config.ruta_precios_referencia)


def probar_extraccion():
    print("\n[1] Lo extraido del HdU cuadra con el archivo original")
    b = biblioteca()

    comprobar(casi(b.tipo_cambio_clp_usd, 928.16),
              "el tipo de cambio sale del nombre definido 'USD' del HdU, "
              "no de un literal inventado", str(b.tipo_cambio_clp_usd))
    comprobar(b.anios == ["2021", "2022", "2023", "2024", "2025"],
              "la serie es de cinco anios", str(b.anios))

    # Cabernet Sauvignon: 375, 420, 300, 216, 297 -> promedio 321,6 CLP/kg
    cabernet = b.buscar("Cabernet Sauvignon")
    comprobar(len(cabernet) == 1, "encuentra el cultivar")
    if cabernet:
        precio = cabernet[0]
        comprobar(casi(precio.valor_usd_kg, 321.6 / 928.16),
                  "321,6 CLP/kg / 928,16 = 0,3465 USD/kg, como en el HdU",
                  str(precio.valor_usd_kg))
        comprobar(precio.fuente == "ODEPA",
                  "y viene con su fuente: el banco puede verificarla")
        comprobar(precio.serie_clp_kg and len(precio.serie_clp_kg) == 5,
                  "conserva la serie anio a anio, no solo el promedio")
        comprobar(casi(precio.serie_clp_kg["2022"], 420),
                  "con los valores originales", str(precio.serie_clp_kg))

    # Los grupos son los que nombran las curvas del catalogo.
    grupo = b.buscar("Traditional Reds")
    comprobar(grupo and casi(grupo[0].valor_usd_kg, 0.334138),
              "el grupo 'Traditional Reds' promedia sus tres cultivares",
              str(grupo[0].valor_usd_kg if grupo else None))
    comprobar(grupo and grupo[0].compuesto_por
              and "Merlot" in grupo[0].compuesto_por,
              "y deja por escrito de que cultivares se compone",
              str(grupo[0].compuesto_por if grupo else None))


def probar_fob_vs_productor():
    print("\n[2] Un FOB NO es un precio a productor")
    b = biblioteca()
    refs = b.buscar("Avellano")

    comprobar(len(refs) >= 2,
              "el avellano tiene varias referencias, no una sola verdad",
              str(len(refs)))
    fob = b.fob_de("Avellano")
    comprobar(fob is not None and casi(fob.valor_usd_kg, 6.90),
              "hay un FOB de exportacion de 6,90 USD/kg",
              str(fob.valor_usd_kg if fob else None))
    comprobar(fob is not None and not fob.listo_para_el_flujo,
              "y esta marcado como NO usable directo en el flujo: primero hay "
              "que restarle la cadena de descuentos")

    sugerido = b.sugerir_para_flujo("Avellano")
    comprobar(sugerido is not None and sugerido.listo_para_el_flujo,
              "lo que se sugiere para el flujo es un precio a productor")
    comprobar(sugerido is not None and sugerido.valor_usd_kg < fob.valor_usd_kg,
              "y es menor que el FOB, como tiene que ser",
              "{} vs {}".format(sugerido.valor_usd_kg if sugerido else None,
                                fob.valor_usd_kg if fob else None))

    # La diferencia entre ambos es lo que cobra la cadena exportadora: en el
    # avellano, mas de la mitad del FOB. Poner el FOB en la fila de precio
    # duplicaria el ingreso del cliente sin que nada lo delate.
    if fob and sugerido:
        comprobar(fob.valor_usd_kg / sugerido.valor_usd_kg > 1.5,
                  "la cadena se lleva mas de un tercio del FOB: por eso "
                  "confundirlos no es un detalle",
                  "{:.0f}%".format(100 * (1 - sugerido.valor_usd_kg / fob.valor_usd_kg)))


def probar_procedencia():
    print("\n[3] Cada precio dice de donde salio y cuanto fiarse")
    b = biblioteca()

    odepa = b.buscar("Coastal White")[0]
    comprobar(odepa.confianza == "alta",
              "una serie de ODEPA de cinco anios es de confianza alta")
    comprobar("ODEPA" in odepa.procedencia(),
              "y su procedencia se puede pegar en un comentario de celda",
              odepa.procedencia())

    giffoni = [r for r in b.buscar("Avellano")
               if r.unidad and "Giffoni" in r.unidad]
    comprobar(giffoni and giffoni[0].confianza == "baja",
              "un precio de prensa sin anio preciso es de confianza baja")
    comprobar(giffoni and giffoni[0].advertencia,
              "y lleva escrito por que hay que confirmarlo con el tasador",
              str(giffoni[0].advertencia[:60] if giffoni else None))

    # Las referencias vienen ordenadas: primero la mas confiable.
    refs = b.buscar("Avellano")
    orden = {"alta": 0, "media": 1, "baja": 2}
    comprobar(refs == sorted(refs, key=lambda r: orden.get(r.confianza, 3)),
              "las referencias llegan ordenadas de mas a menos confiable")


def probar_sin_dato():
    print("\n[4] Lo que no esta, no se inventa")
    b = biblioteca()

    comprobar(b.buscar("Palto") == [],
              "el palto no tiene precio de referencia: ODEPA no publica "
              "precio a productor de fruta de exportacion")
    comprobar(b.sugerir_para_flujo("Kiwi") is None,
              "y una especie desconocida no devuelve un valor por defecto")
    comprobar(b.buscar("avellano EUROPEO") == b.buscar("Avellano"),
              "la busqueda tolera mayusculas y el nombre largo o corto")


def main():
    print("=" * 72)
    print("PRUEBA DE LA BIBLIOTECA DE PRECIOS DE REFERENCIA")
    print("=" * 72)

    probar_extraccion()
    probar_fob_vs_productor()
    probar_procedencia()
    probar_sin_dato()

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
