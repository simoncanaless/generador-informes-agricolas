"""Los precios FOB de exportacion que publica ODEPA.

Lo que se comprueba no es que el archivo se lea: es que el dato sirva sin
enganar. Tres cosas, y las tres nacieron de mirar los numeros de verdad:

- Que un FOB **no pueda** entrar al flujo por descuido.
- Que un grupo que mezcla productos de precio muy distinto lo DIGA. El grupo
  «Avellana» promedia con cascara (4,17 USD/kg) y sin cascara (9,99): el
  promedio no describe a ninguno.
- Que la temporada de referencia sea una CERRADA. El archivo trae tambien la
  en curso, cortada donde alcanzo el boletin, y esa sesga el precio.

Uso: python scripts/test_fob_odepa.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.settings import ajustes                              # noqa: E402
from services.precios import cargar_fob                          # noqa: E402

fallos, comprobaciones = [], 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)


def main():
    print("=" * 72)
    print("PRECIOS FOB DE EXPORTACION (ODEPA / Aduanas)")
    print("=" * 72)

    ruta = ajustes.raiz / "config" / "precios_fob_odepa.json"
    if not ruta.exists():
        print("\n  Falta {}.".format(ruta.name))
        print("  Se genera con scripts/importar_fob_odepa.py; el Excel se baja")
        print("  del boletin mensual de ODEPA (ver la cabecera de ese script).")
        return 1

    fob = cargar_fob(ruta)
    print("  temporada de referencia: {}".format(fob.temporada))
    print("  {} grupos".format(len(fob.grupos)))

    print("\n[1] Un FOB no puede entrar al flujo por descuido")
    avellana = fob.fob_de("Avellana")
    comprobar(avellana is not None, "encuentra el avellano")
    if avellana:
        comprobar(avellana.es_fob, "queda marcado como FOB", avellana.tipo)
        comprobar(not avellana.listo_para_el_flujo,
                  "y por eso NO esta listo para la fila de precio")
        comprobar(bool(fob.advertencia)
                  and "productor" in fob.advertencia.lower(),
                  "el archivo lo dice en su advertencia",
                  (fob.advertencia or "")[:70])

    print("\n[2] El nombre del catastro encuentra el grupo de Aduanas")
    # El catastro dice «Avellano europeo» y Aduanas «Avellana». Sin esto, el
    # 83% de la superficie de Cliente A se quedaba sin precio.
    for escrito, esperado in (("Avellano europeo", "Avellana"),
                              ("Nogal", "Nuez"),
                              ("Palto", "Palta"),
                              ("Cerezo", "Cereza")):
        hallado = fob.fob_de(escrito)
        comprobar(hallado is not None and hallado.nombre == esperado,
                  "«{}» -> {}".format(escrito, esperado),
                  hallado.nombre if hallado else "no lo encontro")

    print("\n[3] Los grupos que mezclan productos lo advierten")
    comprobar(fob.es_mezcla("Avellana"),
              "el avellano esta marcado como mezcla")
    codigos = fob.codigos_de("Avellana")
    comprobar(len(codigos) >= 2, "y trae su desglose por codigo",
              "{} codigos".format(len(codigos)))
    if len(codigos) >= 2:
        precios = [c.usd_kg for c in codigos]
        comprobar(max(precios) > 2 * min(precios),
                  "con cascara y sin cascara difieren mas del doble",
                  "{:.2f} vs {:.2f}".format(min(precios), max(precios)))
        con = [c for c in codigos if "con c" in c.producto.lower()]
        comprobar(bool(con), "el desglose distingue «con cascara»",
                  ", ".join(c.producto[:30] for c in codigos))
    comprobar(avellana is not None and bool(avellana.advertencia),
              "y el precio del grupo lleva la advertencia encima")

    print("\n[4] El nogal cae en la MISMA trampa")
    comprobar(fob.es_mezcla("Nogal"), "tambien esta marcado como mezcla")
    nueces = fob.codigos_de("Nogal")
    if len(nueces) >= 2:
        precios = [c.usd_kg for c in nueces]
        comprobar(max(precios) > 2 * min(precios),
                  "con cascara y sin cascara difieren mas del doble",
                  "{:.2f} vs {:.2f}".format(min(precios), max(precios)))

    print("\n[5] La temporada de referencia esta CERRADA")
    comprobar(fob.temporada and "Ago" in fob.temporada,
              "va de septiembre a agosto, no cortada a mitad", str(fob.temporada))

    print("\n[5.bis] Ningun grupo se queda SIN SERIE historica")
    # El error que esto atrapa: el importador guardaba la serie con la clave
    # CON tilde y los grupos sin ella, asi que «arandanos y mirtilos»,
    # «castana», «limon» y «nectarin» quedaban con la serie vacia. No fallaba
    # nada: simplemente esas cuatro especies no mostraban su rango de precios,
    # y el arandano es una de las que mas se tasan. Tercera vez que muerde la
    # misma regla de normalizacion (§16.a.quadragies.quinquies).
    sin_serie = [g.nombre for g in
                 (fob.fob_de(n) for n in fob.grupos)
                 if g is not None and not fob.serie_de(g.nombre)]
    comprobar(not sin_serie,
              "los {} grupos traen su serie".format(len(fob.grupos)),
              "sin serie: " + ", ".join(sin_serie))
    if sin_serie:
        print("       El importador YA esta corregido; lo que esta mal es el")
        print("       archivo, generado con la version que tenia el error.")
        print("       Se arregla bajando de nuevo «Exportaciones por")
        print("       Temporada» del boletin y corriendo:")
        print("         python scripts/importar_fob_odepa.py <el .xlsx>")
    con_tilde = [n for n in fob.grupos
                 if any(c in n for c in "áéíóúñÁÉÍÓÚÑ")]
    comprobar(all(fob.serie_de(n) for n in con_tilde),
              "incluidos los que llevan tilde en el nombre",
              ", ".join(con_tilde) or "(ninguno tiene tilde)")

    print("\n[5.ter] Solo fruta FRESCA (o fruto seco), nunca conserva")
    # Lo pidio el agrónomo: el precio tiene que ser el de la fruta fresca, no el
    # de conserva ni el de jugo, que son otro negocio y otro precio.
    permitidos = {"Fresca", "Frutos secos"}
    estados = {fob.procesamiento_de(n) for n in fob.grupos}
    comprobar(estados <= permitidos,
              "ningun grupo viene de conserva, jugo ni deshidratado",
              ", ".join(sorted(str(e) for e in estados)))
    comprobar(all(fob.procesamiento_de(n) for n in fob.grupos),
              "y cada grupo dice en que estado se exporto, para poder "
              "comprobarlo en pantalla")

    print("\n[6] Cada precio puede decir de donde salio")
    if avellana:
        linea = avellana.procedencia()
        comprobar("ODEPA" in linea and "Ago" in linea,
                  "la procedencia nombra la fuente y la temporada", linea)
    serie = fob.serie_de("Avellana")
    comprobar(len(serie) >= 5, "y trae la serie historica para ver si el "
              "ultimo anio fue normal", "{} temporadas".format(len(serie)))
    if serie:
        valores = list(serie.values())
        print("      avellano: de {:.2f} a {:.2f} USD/kg en {} temporadas".format(
            min(valores), max(valores), len(valores)))

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
