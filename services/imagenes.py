"""Compresion de imagenes, con tres destinos y tres criterios.

Las fotos de una visita salen de la camara a 4000x2250 y 3-5 MB cada una.
Catorce de esas son 50 MB, y ese peso hace dano en dos lugares distintos:

- **En el informe.** Un Excel de 50 MB no pasa por el correo de nadie, y el
  ejecutivo del banco no lo va a poder abrir. Es lo que pidio el agrónomo: bajarles
  la calidad, porque en el informe impreso o en PDF no se nota.
- **En el analisis.** Cada foto viaja entera a Gemini. Subir 3 MB en vez de
  115 KB no mejora la descripcion -el modelo trabaja en mosaicos de 768 px, asi
  que todo lo que sobra de esa medida se descarta al otro lado- pero si alarga
  la espera, y con catorce fotos eso son minutos.

Por eso hay dos ajustes y no uno. El del informe conserva mas detalle porque
alguien lo va a mirar; el del analisis es mas agresivo porque solo lo mira un
modelo que igual lo iba a reducir.

Medido sobre una foto real de la visita de Cliente C (3.465 KB):

    1600 px  q85    413 KB    -88%      (lo que se usaba para el informe)
    1280 px  q75    201 KB    -94%      <- INFORME
    1024 px  q70    115 KB    -97%      <- ANALISIS
     768 px  q70     62 KB    -98%

**Y un TERCER ajuste, para las imagenes satelitales** (2026-09-09). el agrónomo:
*«las fotos satelitales tienen que tener buena calidad y tienen que ser mas
grandes que las otras imagenes»*. No es un capricho: una foto de terreno se
mira entera -se ve un huerto- y una satelital se mira POR DENTRO, buscando los
limites del predio y las lineas de plantacion. Ahi el detalle fino es el dato,
y es justo lo primero que se lleva el JPEG.

El numero sale de medir el hueco donde cae, no de elegirlo a ojo. La imagen
satelital va al marcador de emplazamiento de la plantilla -``Cliente!A40:I60``,
que es el unico hueco de imagen que trae la Bsan- y ese rango mide **794 x 420
px** en pantalla. Como Excel cuenta 96 px por pulgada, el ancho impreso es
794/96 = 8,27 pulgadas, asi que:

    1280 px -> 155 ppp     lo que usa una foto normal: se ve blanda al ampliar
    1600 px -> 193 ppp
    2000 px -> 242 ppp     <- SATELITAL
    2400 px -> 290 ppp

Se toma **2000 px**, que pasa los 200 ppp con margen. Y calidad **88** y no 75
por lo mismo de arriba: a q75 los bordes de los cuarteles salen con el halo
tipico del JPEG, que en una imagen de referencia se lee como si el limite
estuviera difuso.

Cuesta peso y esta bien que lo cueste. Medido sobre una foto real de 3264x1836
(1.380 KB), los tres ajustes dan:

    1024 px  q70     80 KB      <- lo que se guarda de una foto de terreno
    1280 px  q75    124 KB
    2000 px  q88    381 KB      <- SATELITAL, casi 5 veces mas

Son una o dos satelitales por informe contra catorce fotos, asi que el Excel
sigue pasando por correo, que era el problema que resolvio la compresion.

Nunca se falla hacia el error: si PIL no esta o la imagen no se puede abrir, se
devuelve el original. Una foto pesada es un problema; una foto perdida es peor.
"""

from __future__ import annotations

from io import BytesIO
from typing import Optional, Tuple

__all__ = [
    "comprimir", "para_informe", "para_analisis", "para_satelital",
    "LADO_INFORME", "CALIDAD_INFORME", "LADO_ANALISIS", "CALIDAD_ANALISIS",
    "LADO_SATELITAL", "CALIDAD_SATELITAL",
]

#: Lo que se incrusta en el Excel. Con la foto ocupando media pagina, 1280 px
#: da mas de 200 ppp al imprimir: mas que suficiente y la mitad de peso que
#: los 1600 px que se usaban antes.
LADO_INFORME = 1280
CALIDAD_INFORME = 75

#: Lo que viaja a la IA. Gemini trocea en mosaicos de 768 px, asi que subir
#: mas resolucion que esta no agrega informacion, solo espera.
LADO_ANALISIS = 1024
CALIDAD_ANALISIS = 70

#: La imagen satelital, que ocupa media carilla del informe y se mira POR
#: DENTRO. 2000 px sobre el hueco de 794 px de la plantilla dan 242 ppp al
#: imprimir. Ver la cabecera: el numero sale de medir el rango, no de elegirlo.
LADO_SATELITAL = 2000
CALIDAD_SATELITAL = 88

#: Formatos que openpyxl sabe incrustar. El resto se convierte a JPEG. Las
#: camaras de los drones DJI guardan en MPO -un JPEG con varias imagenes
#: dentro- y openpyxl revienta con KeyError: '.mpo'.
FORMATOS_SEGUROS = {"png", "jpeg", "gif", "bmp"}


def comprimir(
    origen,
    lado_maximo: int,
    calidad: int,
    forzar: bool = False,
) -> Tuple[Optional[bytes], Optional[str]]:
    """Reduce una imagen. Devuelve ``(bytes, aviso)``; bytes es None si no toco.

    ``origen`` puede ser una ruta o los bytes de la imagen. Con ``forzar`` se
    recomprime aunque la imagen ya sea chica: es lo que hace falta cuando lo
    que importa no es el tamano sino el PESO del archivo.
    """
    try:
        from PIL import Image
    except ImportError:
        return None, "sin Pillow: la imagen va sin comprimir"

    try:
        entrada = BytesIO(origen) if isinstance(origen, (bytes, bytearray)) else origen
        with Image.open(entrada) as imagen:
            formato = (imagen.format or "").lower()
            grande = max(imagen.size) > lado_maximo
            if not (forzar or grande or formato not in FORMATOS_SEGUROS):
                return None, None

            copia = imagen.convert("RGB")
            if grande:
                copia.thumbnail((lado_maximo, lado_maximo), Image.LANCZOS)
            buffer = BytesIO()
            copia.save(buffer, format="JPEG", quality=calidad, optimize=True)
            return buffer.getvalue(), None
    except Exception as error:                        # noqa: BLE001
        # Antes pesada que perdida: si algo sale mal se usa la original.
        return None, "no se pudo comprimir ({})".format(type(error).__name__)


def para_informe(origen, forzar: bool = False):
    """La version que se incrusta en el Excel que va al banco."""
    return comprimir(origen, LADO_INFORME, CALIDAD_INFORME, forzar)


def para_analisis(origen, forzar: bool = True):
    """La version que se le manda a la IA. Se fuerza: aqui manda el peso."""
    return comprimir(origen, LADO_ANALISIS, CALIDAD_ANALISIS, forzar)


def para_satelital(origen, forzar: bool = False):
    """La imagen satelital que se incrusta grande en el informe.

    NO se fuerza por defecto, al reves que ``para_analisis``: si la satelital
    ya viene por debajo de los 2000 px, recomprimirla solo le quita calidad
    -que es justo lo que el agrónomo pidio conservar- y no ahorra nada que
    importe. Solo se toca si es mas grande, o si el formato no lo aguanta
    openpyxl (las capturas de dron vienen en MPO).
    """
    return comprimir(origen, LADO_SATELITAL, CALIDAD_SATELITAL, forzar)
