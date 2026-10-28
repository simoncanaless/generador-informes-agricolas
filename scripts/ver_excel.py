"""Renderiza un .xlsx a imagenes para poder revisarlo VISUALMENTE.

Existe por una leccion concreta: la simulacion de Cliente C pasaba todas las
comprobaciones automaticas y aun asi, al abrirla, se veia rota. Habia fechas
como '#########', fotos apiladas una encima de otra y cuatro hojas vacias.
Nada de eso lo detecta una prueba que cuenta celdas, porque las celdas estaban
bien: lo que fallaba era como se veian.

Este script usa el propio Excel para exportar el libro a PDF y despues lo pasa a
PNG. El resultado es lo que el tasador ve en pantalla, no lo que openpyxl cree
que escribio.

    python scripts/ver_excel.py "output/SIMULACION Cliente C.xlsx"
    python scripts/ver_excel.py archivo.xlsx --hojas Cliente "Flujo proyectado"

Si no hay Excel instalado intenta con LibreOffice.

El pipeline (COM/PowerShell -> PDF -> PNG) vive en services/renderizado.py,
compartido con el boton "Vista previa" de la pantalla: este script es solo
la puerta de entrada por consola.
"""

import argparse
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.renderizado import (  # noqa: E402
    PPP,
    exportar_con_excel,
    exportar_con_libreoffice,
    pdf_a_imagenes,
)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("archivo")
    parser.add_argument("--hojas", nargs="*", default=None,
                        help="solo estas hojas; por defecto el libro entero")
    parser.add_argument("--ppp", type=int, default=PPP)
    parser.add_argument("--salida", default=None,
                        help="carpeta de las imagenes (por defecto output/vistas)")
    args = parser.parse_args()

    origen = Path(args.archivo).resolve()
    if not origen.exists():
        print("No existe: {}".format(origen))
        return 1

    from config.settings import ajustes
    carpeta = Path(args.salida) if args.salida else ajustes.ruta_salida / "vistas"
    carpeta.mkdir(parents=True, exist_ok=True)
    for viejo in carpeta.glob("{}_pag*.png".format(origen.stem[:30])):
        viejo.unlink()

    print("renderizando {} ...".format(origen.name))
    with tempfile.TemporaryDirectory() as temporal:
        pdf = Path(temporal) / "vista.pdf"

        if exportar_con_excel(origen, pdf, args.hojas):
            motor = "Excel"
        elif exportar_con_libreoffice(origen, Path(temporal)):
            motor = "LibreOffice"
            pdf = next(Path(temporal).glob("*.pdf"))
        else:
            print("\nNo se pudo renderizar: hace falta Excel o LibreOffice.")
            print("Comprueba que el archivo no este abierto en Excel.")
            return 1

        imagenes = pdf_a_imagenes(pdf, carpeta, origen.stem[:30], args.ppp)

    print("motor: {} | {} pagina(s)".format(motor, len(imagenes)))
    for ruta in imagenes:
        print("  {}  ({:,.0f} KB)".format(ruta, ruta.stat().st_size / 1024))
    return 0


if __name__ == "__main__":
    sys.exit(main())
