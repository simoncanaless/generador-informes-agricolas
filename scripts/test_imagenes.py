"""Prueba de la compresion de fotografias.

Las fotos de una visita salen de la camara a 4000x2250 y 3-5 MB. Catorce de
esas hacen dos danos: un Excel que no pasa por correo y un analisis que tarda
minutos de mas porque cada foto viaja entera a la IA.

Uso: python scripts/test_imagenes.py
"""

import glob
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.imagenes import (  # noqa: E402
    CALIDAD_ANALISIS, LADO_ANALISIS, LADO_INFORME, comprimir, para_analisis,
    para_informe,
)

fallos, comprobaciones = [], 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)


def foto_sintetica(ancho=4000, alto=2250):
    """Una foto del tamano de las de verdad, sin depender de archivos."""
    from PIL import Image

    imagen = Image.new("RGB", (ancho, alto))
    pixeles = imagen.load()
    # Ruido suave: una imagen plana comprime a nada y no probaria nada.
    for x in range(0, ancho, 4):
        for y in range(0, alto, 4):
            tono = (x * 7 + y * 13) % 256
            for dx in range(4):
                for dy in range(4):
                    if x + dx < ancho and y + dy < alto:
                        pixeles[x + dx, y + dy] = (tono, (tono * 3) % 256, 90)
    buffer = io.BytesIO()
    imagen.save(buffer, format="JPEG", quality=95)
    return buffer.getvalue()


def probar_reduccion():
    print("\n[1] Reduce de verdad, y a la medida de cada destino")
    from PIL import Image

    original = foto_sintetica()
    ia, aviso_ia = para_analisis(original)
    informe, aviso_inf = para_informe(original)

    comprobar(not aviso_ia and not aviso_inf,
              "comprime sin avisos", "{} {}".format(aviso_ia, aviso_inf))
    comprobar(ia is not None and informe is not None,
              "devuelve bytes en los dos casos")
    comprobar(len(ia) < len(original) * 0.2,
              "la version para la IA pesa menos del 20% del original",
              "{:.0f} KB de {:.0f} KB".format(len(ia)/1024, len(original)/1024))
    comprobar(len(informe) < len(original) * 0.3,
              "la del informe, menos del 30%",
              "{:.0f} KB".format(len(informe)/1024))

    with Image.open(io.BytesIO(ia)) as im:
        comprobar(max(im.size) == LADO_ANALISIS,
                  "la de analisis queda en el lado declarado", str(im.size))
        comprobar(im.size[0] > im.size[1],
                  "y conserva la proporcion (no la deforma)", str(im.size))
    with Image.open(io.BytesIO(informe)) as im:
        comprobar(max(im.size) == LADO_INFORME,
                  "la del informe, en el suyo", str(im.size))

    # La del informe conserva mas detalle que la del analisis: son dos
    # criterios distintos a proposito, no un descuido.
    comprobar(len(informe) > len(ia),
              "el informe conserva mas detalle que lo que se le manda a la IA",
              "{:.0f} KB vs {:.0f} KB".format(len(informe)/1024, len(ia)/1024))


def probar_sigue_siendo_legible():
    print("\n[2] Sigue siendo una imagen valida y con contenido")
    from PIL import Image

    original = foto_sintetica()
    datos, _ = para_analisis(original)
    with Image.open(io.BytesIO(datos)) as im:
        im.load()                       # revienta si el JPEG quedo corrupto
        comprobar(im.format == "JPEG", "el resultado es un JPEG", str(im.format))
        colores = im.convert("RGB").getcolors(maxcolors=1000000)
        comprobar(colores is not None and len(colores) > 500,
                  "conserva variedad de color: no quedo un borron",
                  str(len(colores) if colores else 0))


def probar_no_pierde_la_foto():
    print("\n[3] Ante la duda, la foto original")
    # Una entrada que no es una imagen no puede tumbar el analisis de una
    # visita entera: se devuelve None y un aviso, y el llamador usa lo que
    # tenia. Perder la foto es peor que mandarla pesada.
    datos, aviso = para_analisis(b"esto no es una imagen")
    comprobar(datos is None, "no inventa bytes cuando no puede leer")
    comprobar(bool(aviso), "y lo dice en vez de fallar en silencio", str(aviso))


def probar_forzar():
    print("\n[4] Una foto ya chica igual se recomprime si se pide")
    # Al mandar a la IA lo que importa es el PESO, no el tamano: una foto de
    # 900 px puede pesar 2 MB si viene con calidad 100.
    from PIL import Image

    imagen = Image.new("RGB", (800, 600), (120, 200, 80))
    buffer = io.BytesIO()
    imagen.save(buffer, format="JPEG", quality=100)
    pesada = buffer.getvalue()

    suave, _ = comprimir(pesada, LADO_ANALISIS, CALIDAD_ANALISIS, forzar=True)
    comprobar(suave is not None and len(suave) < len(pesada),
              "con forzar=True baja el peso aunque no cambie el tamano",
              "{} -> {}".format(len(pesada), len(suave) if suave else None))

    intacta, _ = comprimir(pesada, LADO_ANALISIS, CALIDAD_ANALISIS, forzar=False)
    comprobar(intacta is None,
              "sin forzar, una imagen ya chica no se toca")


def probar_fotos_reales():
    print("\n[5] Sobre las fotos de verdad de la visita")
    fotos = glob.glob("archivos para pruebas/*.jpg")[:14]
    if not fotos:
        print("  (no hay fotos de prueba; se omite)")
        return

    original = sum(os.path.getsize(f) for f in fotos)
    total_ia = total_informe = 0
    for ruta in fotos:
        datos = open(ruta, "rb").read()
        ia, _ = para_analisis(datos)
        inf, _ = para_informe(datos)
        total_ia += len(ia) if ia else len(datos)
        total_informe += len(inf) if inf else len(datos)

    mb = lambda n: n / 1024 / 1024                              # noqa: E731
    print("    original {:.1f} MB | analisis {:.1f} MB | informe {:.1f} MB"
          .format(mb(original), mb(total_ia), mb(total_informe)))

    comprobar(total_ia < original * 0.10,
              "las {} fotos que viajan a la IA pesan menos del 10%".format(
                  len(fotos)),
              "{:.1f} MB de {:.1f} MB".format(mb(total_ia), mb(original)))
    comprobar(mb(total_informe) < 8,
              "el Excel con todas las fotos se queda bajo 8 MB: pasa por correo",
              "{:.1f} MB".format(mb(total_informe)))


def main():
    print("=" * 72)
    print("PRUEBA DE COMPRESION DE FOTOGRAFIAS")
    print("=" * 72)
    try:
        import PIL                                             # noqa: F401
    except ImportError:
        print("  (sin Pillow: se omite la prueba entera)")
        return 0

    probar_reduccion()
    probar_sigue_siendo_legible()
    probar_no_pierde_la_foto()
    probar_forzar()
    probar_fotos_reales()

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
