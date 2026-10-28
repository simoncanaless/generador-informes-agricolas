"""Genera `web/fitogreen.ico` desde el emblema de la pagina.

    python scripts/hacer_icono.py

El icono es **el mismo dibujo** que muestra la pantalla al arrancar: el anillo
con el degradado del logo y las dos hojas. Se reproduce aca en vez de copiarse
un PNG suelto para que no haya dos versiones del emblema; si algun dia cambia
el de la pagina, esto se vuelve a correr y quedan iguales.

A 16 px no se reconoce mas que el anillo y una mancha verde, que es justamente
por lo que el original no lleva la «F»: a ese tamano seria una mancha mas.

Se usa en dos sitios: el acceso directo del escritorio y la ventana del panel
(`scripts/panel.pyw`). Sin el, el panel sale con la pluma de Python en la barra
de tareas, que no dice nada de que programa es.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PIL import Image, ImageDraw                                # noqa: E402

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SALIDA = os.path.join(RAIZ, "web", "fitogreen.ico")

#: El dibujo esta en coordenadas de 40x40 -las del SVG de la pagina- y se
#: trabaja 16 veces mas grande para que al reducir quede con bordes suaves.
LADO = 40
ESCALA = 16

#: Los cuatro colores del anillo, con su posicion. Salen del logo: turquesa a
#: la izquierda, amarillo a la derecha (§16.a.undevicies).
PARADAS = [(0.00, (0x35, 0xbf, 0xce)),
           (0.34, (0x6c, 0xc7, 0x9f)),
           (0.66, (0xa6, 0xce, 0x39)),
           (1.00, (0xe9, 0xde, 0x77))]

#: Las dos hojas, tal como estan en el SVG: un punto de partida y tramos de
#: curva cubica en coordenadas RELATIVAS, de a seis numeros.
HOJAS = [
    ((0xa6, 0xce, 0x39), (20.2, 20.6), [
        (3.4, -4.6, 8.2, -4.9, 10.1, -4.6),
        (0.5, 2.6, -1.1, 8.6, -7, 9.2),
        (-2.1, 0.2, -3.4, -0.6, -3.1, -4.6)]),
    ((0x7f, 0xb0, 0x27), (9.6, 25.2), [
        (-1.4, -4.6, 1.2, -9.2, 6.1, -10.4),
        (2.4, -0.6, 4, 0.3, 3.7, 3.4),
        (-0.4, 4.3, -5, 8, -9.8, 7)]),
]


def color_en(fraccion):
    """El color del degradado a esa altura, interpolando entre paradas."""
    for i in range(len(PARADAS) - 1):
        desde, color_a = PARADAS[i]
        hasta, color_b = PARADAS[i + 1]
        if desde <= fraccion <= hasta:
            t = (fraccion - desde) / (hasta - desde) if hasta > desde else 0
            return tuple(int(round(a + (b - a) * t))
                         for a, b in zip(color_a, color_b))
    return PARADAS[-1][1]


def curva(desde, tramos, por_tramo=48):
    """Convierte los tramos cubicos en una lista de puntos.

    Pillow no entiende curvas, asi que se evaluan a mano. Con 48 puntos por
    tramo y el dibujo 16 veces mas grande, al reducir no se nota que es un
    poligono.
    """
    puntos = []
    x0, y0 = desde
    for tramo in tramos:
        x1, y1 = x0 + tramo[0], y0 + tramo[1]
        x2, y2 = x0 + tramo[2], y0 + tramo[3]
        x3, y3 = x0 + tramo[4], y0 + tramo[5]
        for paso in range(por_tramo + 1):
            t = paso / por_tramo
            u = 1 - t
            x = (u ** 3 * x0 + 3 * u * u * t * x1
                 + 3 * u * t * t * x2 + t ** 3 * x3)
            y = (u ** 3 * y0 + 3 * u * u * t * y1
                 + 3 * u * t * t * y2 + t ** 3 * y3)
            puntos.append((x * ESCALA, y * ESCALA))
        x0, y0 = x3, y3
    return puntos


def dibujar():
    lado = LADO * ESCALA
    lienzo = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))

    # El degradado va de izquierda a derecha: una columna por color.
    degradado = Image.new("RGB", (lado, lado))
    pinta = ImageDraw.Draw(degradado)
    for x in range(lado):
        pinta.line([(x, 0), (x, lado)], fill=color_en(x / (lado - 1)))

    # El anillo es una mascara: circunferencia de radio 16,4 y grosor 3,2,
    # los mismos numeros del SVG.
    mascara = Image.new("L", (lado, lado), 0)
    trazo = ImageDraw.Draw(mascara)
    radio, grosor = 16.4 * ESCALA, 3.2 * ESCALA
    centro = 20 * ESCALA
    trazo.ellipse([centro - radio, centro - radio,
                   centro + radio, centro + radio],
                  outline=255, width=int(round(grosor)))
    lienzo.paste(degradado, (0, 0), mascara)

    hojas = ImageDraw.Draw(lienzo)
    for color, desde, tramos in HOJAS:
        hojas.polygon(curva(desde, tramos), fill=color + (255,))

    return lienzo


def main():
    grande = dibujar()
    tamanos = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    icono = grande.resize((256, 256), Image.LANCZOS)
    icono.save(SALIDA, format="ICO", sizes=tamanos)
    print("{}  ({} tamaños, {:,} bytes)".format(
        SALIDA, len(tamanos), os.path.getsize(SALIDA)))
    vista = os.path.join(RAIZ, "output", "pruebas", "icono.png")
    os.makedirs(os.path.dirname(vista), exist_ok=True)
    grande.resize((256, 256), Image.LANCZOS).save(vista)
    print("{}  (para mirarlo)".format(vista))


if __name__ == "__main__":
    main()
