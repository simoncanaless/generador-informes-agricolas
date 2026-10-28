"""Prueba de services/renderizado.py: el pipeline que abre un .xlsx con el
Excel real (o LibreOffice) y lo convierte en imagenes -WebP sin perdida a
200 ppp para la vista previa desde 2026-10-05- -el mismo mecanismo que ya
usaba scripts/ver_excel.py por consola, ahora reusable desde el endpoint
"Vista previa" de la pantalla (el usuario, 2026-09-23: "que se vea en la pagina...
sin tener que descargarlo").

Dos niveles, como el resto del proyecto distingue lo que corre en cualquier
entorno de lo que necesita Excel instalado de verdad:

- Sin Excel: deterministas, con dobles de prueba. Corren siempre.
- Con Excel real: best-effort -si no hay Excel ni LibreOffice en esta
  maquina, avisan y siguen, igual que scripts/test_exclusiones.py hace con
  su recalculo real.

Uso: python scripts/test_renderizado.py
"""

import asyncio
import os
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import services.renderizado as renderizado  # noqa: E402
from services.renderizado import (  # noqa: E402
    PPP_VISTA_PREVIA,
    RenderNoDisponible,
    _hay_webp,
    _pdf_a_paginas,
    renderizar_a_paginas,
)

FIRMA_PNG = b"\x89PNG\r\n\x1a\n"


def es_imagen_valida(datos, tipo):
    """La firma real del formato que dice ser -no bytes cualquiera-.

    WebP es un contenedor RIFF: 'RIFF', cuatro bytes de largo, y 'WEBP'."""
    if tipo == "image/webp":
        return datos[:4] == b"RIFF" and datos[8:12] == b"WEBP"
    if tipo == "image/png":
        return datos[:8] == FIRMA_PNG
    return False

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


def _pdf_sintetico(ruta):
    """Una pagina en blanco, sin pasar por Excel -solo para ejercitar
    _pdf_a_bytes_png en aislamiento."""
    import pymupdf

    documento = pymupdf.open()
    try:
        documento.new_page(width=300, height=200)
        documento.save(ruta)
    finally:
        documento.close()


def probar_render_no_disponible():
    print("\n[1] Ni Excel ni LibreOffice: RenderNoDisponible, no una excepcion muda")
    original_excel = renderizado.exportar_con_excel
    original_libre = renderizado.exportar_con_libreoffice
    renderizado.exportar_con_excel = lambda *a, **k: False
    renderizado.exportar_con_libreoffice = lambda *a, **k: False
    try:
        lanzo = False
        try:
            asyncio.run(renderizar_a_paginas(Path("no_existe.xlsx")))
        except RenderNoDisponible:
            lanzo = True
        comprobar(lanzo, "renderizar_a_paginas avisa con un tipo de error "
                          "propio, no deja pasar un resultado vacio en silencio")
    finally:
        renderizado.exportar_con_excel = original_excel
        renderizado.exportar_con_libreoffice = original_libre

    # Verificado en negativo: sin el monkeypatch de arriba (o sea, con las
    # funciones reales), esta MISMA prueba fallaria en una maquina sin
    # Excel -es la prueba correcta, simplemente exige forzar el escenario.


def probar_pdf_a_paginas_con_pdf_sintetico():
    print("\n[2] _pdf_a_paginas: imagenes validas y SIN PERDIDA, en memoria")
    import io

    import pymupdf
    from PIL import Image

    with tempfile.TemporaryDirectory() as tmp:
        pdf = Path(tmp) / "sintetico.pdf"
        _pdf_sintetico(pdf)

        paginas = _pdf_a_paginas(pdf, ppp=72)
        comprobar(len(paginas) == 1, "una pagina de entrada, una imagen de salida",
                  len(paginas))
        pagina = paginas[0]
        comprobar(es_imagen_valida(pagina.datos, pagina.tipo),
                  "y arranca con la firma real del formato que dice ser", pagina.tipo)
        comprobar(pagina.tipo == "image/webp" or not _hay_webp(),
                  "va en WebP cuando Pillow lo trae (pesa menos que PNG)", pagina.tipo)
        comprobar((pagina.ancho, pagina.alto) == (300, 200),
                  "declara su tamano: 300x200 pt a 72 ppp son 300x200 px",
                  (pagina.ancho, pagina.alto))

        # "Sin perdida" se comprueba, no se supone: los pixeles decodificados
        # tienen que ser los mismos que dibujo el PDF. Con perdida el borde de
        # las letras se ensucia, que es justo lo que la vista previa no puede.
        documento = pymupdf.open(pdf)
        try:
            original = documento[0].get_pixmap(dpi=72)
            muestra = pymupdf.open()
            muestra.new_page(width=300, height=200).insert_text(
                (20, 100), "Tasacion 1.234,56 USD", fontsize=14)
            ruta_texto = Path(tmp) / "texto.pdf"
            muestra.save(ruta_texto)
            muestra.close()
        finally:
            documento.close()
        decodificada = Image.open(io.BytesIO(pagina.datos)).convert("RGB")
        comprobar(decodificada.tobytes() == original.samples,
                  "los pixeles decodificados son EXACTAMENTE los del PDF")

        con_texto = _pdf_a_paginas(ruta_texto, ppp=150)[0]
        documento = pymupdf.open(ruta_texto)
        try:
            original_texto = documento[0].get_pixmap(dpi=150)
        finally:
            documento.close()
        comprobar(Image.open(io.BytesIO(con_texto.datos)).convert("RGB").tobytes()
                  == original_texto.samples,
                  "tambien con texto: ni un pixel de las letras cambia")

        antes = set(Path(tmp).iterdir())
        _pdf_a_paginas(pdf, ppp=72)
        despues = set(Path(tmp).iterdir())
        comprobar(antes == despues,
                  "correrlo de nuevo no deja ningun archivo nuevo en la carpeta")


def probar_cerrojo_serializa():
    print("\n[3] El cerrojo: dos renders a la vez no comparten Excel al mismo tiempo")
    # Cada llamada registra SU PROPIO intervalo (entra, sale) bajo su id -no
    # una lista plana compartida-, porque si las dos corren en paralelo el
    # orden de llegada de los prints se entrelaza (entra-A, entra-B, sale-A,
    # sale-B) y una lista plana lleva a comparar el par equivocado. Se
    # detecto exactamente asi: la primera version de esta prueba, con una
    # lista plana, seguia en verde SIN el cerrojo -media hora perdida hasta
    # notar que estaba comparando 'entra' de la segunda contra 'sale' de la
    # primera sin darse cuenta.
    intervalos = {}

    def _lento(origen, hojas, ppp):
        id_llamada = origen.stem
        inicio = time.monotonic()
        time.sleep(0.3)
        intervalos[id_llamada] = (inicio, time.monotonic())
        return [], "Excel"

    original = renderizado._renderizar_sync
    renderizado._renderizar_sync = _lento
    try:
        async def _dos_a_la_vez():
            await asyncio.gather(
                renderizar_a_paginas(Path("a.xlsx")),
                renderizar_a_paginas(Path("b.xlsx")),
            )
        asyncio.run(_dos_a_la_vez())
    finally:
        renderizado._renderizar_sync = original

    comprobar(set(intervalos) == {"a", "b"}, "las dos llamadas de verdad corrieron",
              intervalos)
    (inicio_a, fin_a), (inicio_b, fin_b) = intervalos["a"], intervalos["b"]
    sin_solape = fin_a <= inicio_b or fin_b <= inicio_a
    comprobar(sin_solape,
              "los dos intervalos no se solapan -no hubo dos Excel.exe "
              "corriendo al mismo tiempo compitiendo por la misma maquina",
              intervalos)

    # Verificado en negativo: sacando el `async with _cerrojo` de
    # renderizar_a_paginas, esta comprobacion cae -los dos intervalos
    # arrancan casi juntos y terminan casi juntos, se solapan.


def probar_renderizar_a_paginas_con_excel_real():
    print("\n[4] Con Excel real: el mismo pipeline que scripts/ver_excel.py, en memoria")
    from schemas.informe import Antecedentes, InformeFitogreen
    from services.excel_engine import MotorExcel, cargar_mapa
    from config.settings import ajustes

    with tempfile.TemporaryDirectory() as tmp:
        destino = Path(tmp) / "prueba_render.xlsx"
        motor = MotorExcel(ajustes.ruta_plantilla, cargar_mapa(ajustes.ruta_mapping))
        informe = InformeFitogreen(
            antecedentes=Antecedentes(cliente="Prueba Render", rut="76.543.210-3"))
        resultado = motor.generar(informe, destino)
        comprobar(resultado.ok, "el .xlsx de prueba se genero")

        try:
            paginas, motor_usado = asyncio.run(renderizar_a_paginas(destino))
        except RenderNoDisponible as error:
            print("  (no hay Excel ni LibreOffice instalados: {})".format(error))
            return

        comprobar(motor_usado in ("Excel", "LibreOffice"),
                  "se identifico con que motor se renderizo", motor_usado)
        comprobar(len(paginas) > 0, "genero al menos una pagina", len(paginas))
        comprobar(all(es_imagen_valida(p.datos, p.tipo) for p in paginas),
                  "todas las paginas son imagenes validas")
        # Lo que pidio el usuario: que se distinga. Una hoja A4 a la resolucion de
        # la vista previa mide al menos 1.600 px por su lado corto (8,27" x
        # 200); a los 110 ppp de antes medía 910.
        lado_corto = min(min(p.ancho, p.alto) for p in paginas)
        comprobar(lado_corto >= 8 * PPP_VISTA_PREVIA,
                  "cada pagina tiene resolucion para ampliarla ({} px el lado "
                  "corto)".format(lado_corto), lado_corto)


def probar_vista_previa_no_deja_archivos():
    print("\n[5] vista_previa() llamado directo: nada nuevo queda en output/")
    import base64

    from api.routers.asistente import DatosBasicos, PeticionGeneracion, vista_previa
    from services.excel_engine import MotorExcel, cargar_mapa
    from config.settings import ajustes

    motor = MotorExcel(ajustes.ruta_plantilla, cargar_mapa(ajustes.ruta_mapping))
    peticion = PeticionGeneracion(
        basicos=DatosBasicos(cliente="Prueba Vista Previa", rut="76.543.210-3"))

    # repo=None/almacen=None/usuario=None: sin expediente_id, _verificar_permiso
    # y _construir_informe nunca los tocan -mismo camino que un informe
    # armado desde cero sin base de datos.
    antes = set(ajustes.ruta_salida.glob("*.xlsx"))
    try:
        respuesta = asyncio.run(vista_previa(
            peticion, motor, ajustes, None, None, None))
    except RenderNoDisponible as error:
        print("  (no hay Excel ni LibreOffice instalados: {})".format(error))
        return
    despues = set(ajustes.ruta_salida.glob("*.xlsx"))

    comprobar(antes == despues,
              "generar la vista previa no dejo NINGUN .xlsx nuevo en output/ "
              "-no es un informe entregado, no debe ensuciar la carpeta real",
              despues - antes)
    comprobar(len(respuesta.paginas) > 0, "la respuesta trae paginas",
              len(respuesta.paginas))
    comprobar(len(respuesta.tamanos) == len(respuesta.paginas),
              "y el tamano de cada una, para reservar su lugar antes de cargarla",
              (len(respuesta.tamanos), len(respuesta.paginas)))
    comprobar(respuesta.ppp == PPP_VISTA_PREVIA, "y dice a que resolucion",
              respuesta.ppp)
    import io

    from PIL import Image
    for pagina, (ancho, alto) in zip(respuesta.paginas, respuesta.tamanos):
        cabecera, datos = pagina.split(",", 1)
        tipo = cabecera[len("data:"):].split(";", 1)[0]
        comprobar(cabecera in ("data:image/webp;base64", "data:image/png;base64"),
                  "cada pagina viaja como data URL de imagen", cabecera)
        crudo = base64.b64decode(datos)
        comprobar(es_imagen_valida(crudo, tipo),
                  "y decodifica a una imagen valida del tipo que declara")
        comprobar(Image.open(io.BytesIO(crudo)).size == (ancho, alto),
                  "y mide lo que dice medir", (ancho, alto))

    # Verificado en negativo: quitando el `with tempfile.TemporaryDirectory()`
    # de vista_previa() y escribiendo directo a ajustes.ruta_salida, esta
    # comprobacion cae -aparece el .xlsx de la vista previa en 'despues'.


def main():
    print("=" * 72)
    print("PRUEBA DEL RENDER DE VISTA PREVIA (services/renderizado.py)")
    print("=" * 72)

    probar_render_no_disponible()
    probar_pdf_a_paginas_con_pdf_sintetico()
    probar_cerrojo_serializa()
    probar_renderizar_a_paginas_con_excel_real()
    probar_vista_previa_no_deja_archivos()

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
