"""Dibuja curvas de produccion a PNG para verificarlas a simple vista.

Un assert numerico dice que ``rendimiento_a_edad(curva, 5)`` da 1.600; no dice
si la rampa se ve razonable, si dos curvas quedan pegadas, o si el punto que
declaro el generador quedo donde correspondia. Mismo motivo que
``ver_excel.py``: hay problemas que ninguna comprobacion que cuenta numeros
detecta, porque el numero esta bien y lo que falla es como se ve.

No agrega dependencias: usa Pillow, que ya es requisito de openpyxl.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Sequence

from PIL import Image, ImageDraw, ImageFont

from schemas.curvas import CurvaProduccion
from services.curvas import rendimiento_a_edad

__all__ = ["graficar_curvas"]

_PALETA = [
    (0, 122, 194), (214, 92, 0), (0, 158, 115), (204, 51, 111),
    (117, 76, 173), (166, 124, 0),
]

_ANCHO, _ALTO = 960, 600
_MARGEN_IZQ, _MARGEN_DER = 90, 40
_MARGEN_SUP, _MARGEN_INF = 60, 70


def _fuente(tam: int) -> ImageFont.ImageFont:
    try:
        return ImageFont.truetype("arial.ttf", tam)
    except Exception:  # noqa: BLE001 - sin esa fuente, se usa la de reserva
        return ImageFont.load_default()


def graficar_curvas(
    curvas: Sequence[CurvaProduccion],
    ruta: Path,
    titulo: str,
    horizonte: Optional[int] = None,
) -> Path:
    """Superpone una o mas curvas: rendimiento (kg/ha) vs edad (anios).

    Cada curva se traza anio a anio con ``rendimiento_a_edad`` -lo que
    realmente devuelve el motor, no los puntos crudos- y ademas marca con un
    punto solido las edades que declaro quien construyo la curva
    (``curva.puntos``). Esa distincion es el punto del grafico: en una curva
    parametrica, el tramo entre dos puntos solidos es la rampa que inventa
    ``curva_desde_parametros``, y hay que poder ver que sea una recta
    razonable y no un escalon ni un salto raro.
    """
    if not curvas:
        raise ValueError("no hay curvas para graficar")

    horizonte = horizonte or max(c.edad_maxima for c in curvas) + 2
    techo = max(c.rendimiento_maximo for c in curvas) or 1.0
    techo *= 1.1

    imagen = Image.new("RGB", (_ANCHO, _ALTO), "white")
    dibujo = ImageDraw.Draw(imagen)

    x0, x1 = _MARGEN_IZQ, _ANCHO - _MARGEN_DER
    y0, y1 = _ALTO - _MARGEN_INF, _MARGEN_SUP  # y0 abajo (0), y1 arriba (techo)

    def px(edad: float, valor: float):
        fx = x0 + (edad / horizonte) * (x1 - x0)
        fy = y0 - (valor / techo) * (y0 - y1)
        return (fx, fy)

    fuente_chica = _fuente(13)
    fuente_titulo = _fuente(18)

    dibujo.line([(x0, y1), (x0, y0), (x1, y0)], fill=(60, 60, 60), width=2)

    pasos_y = 5
    for i in range(pasos_y + 1):
        valor = techo * i / pasos_y
        _, y = px(0, valor)
        dibujo.line([(x0 - 5, y), (x1, y)], fill=(228, 228, 228), width=1)
        dibujo.text((8, y - 7), "{:,.0f}".format(valor),
                    font=fuente_chica, fill=(90, 90, 90))

    paso_x = max(1, horizonte // 10)
    for edad in range(0, horizonte + 1, paso_x):
        x, _ = px(edad, 0)
        dibujo.line([(x, y0), (x, y1)], fill=(240, 240, 240), width=1)
        dibujo.text((x - 6, y0 + 8), str(edad),
                    font=fuente_chica, fill=(90, 90, 90))

    dibujo.text((x0, 12), titulo, font=fuente_titulo, fill=(20, 20, 20))
    dibujo.text((x1 - 90, _ALTO - 24), "edad (anios)",
                font=fuente_chica, fill=(60, 60, 60))
    dibujo.text((8, 8), "kg/ha", font=fuente_chica, fill=(60, 60, 60))

    for indice, curva in enumerate(curvas):
        color = _PALETA[indice % len(_PALETA)]
        linea = [px(edad, rendimiento_a_edad(curva, edad))
                  for edad in range(0, horizonte + 1)]
        dibujo.line(linea, fill=color, width=3, joint="curve")

        for punto in curva.puntos:
            if punto.edad > horizonte:
                continue
            cx, cy = px(punto.edad, punto.rendimiento_kg_ha)
            dibujo.ellipse([cx - 4, cy - 4, cx + 4, cy + 4],
                           fill=color, outline=(255, 255, 255))

        ly = _MARGEN_SUP + indice * 20
        dibujo.line([(x0 + 10, ly + 7), (x0 + 40, ly + 7)], fill=color, width=3)
        etiqueta = curva.nombre
        if curva.variedad:
            etiqueta += " ({})".format(curva.variedad)
        dibujo.text((x0 + 46, ly), etiqueta, font=fuente_chica, fill=(20, 20, 20))

    ruta.parent.mkdir(parents=True, exist_ok=True)
    imagen.save(ruta)
    return ruta
