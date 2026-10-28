"""Renderiza un .xlsx a imagenes usando el Excel real (o LibreOffice), para
que lo que se ve sea EXACTAMENTE lo que el tasador veria al abrirlo -no una
aproximacion de openpyxl, que nunca evalua formulas (CLAUDE.md SS14: "la
unica prueba de verdad es abrir el libro con Excel real").

Nace de scripts/ver_excel.py -que sigue existiendo como CLI delgado sobre
este modulo-, para que el boton "Vista previa" de la pantalla use el mismo
pipeline que ya se usaba para revisar informes a mano.
"""
import asyncio
import io
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

#: Resolucion del render a disco (scripts/ver_excel.py). 110 ppp deja el texto
#: legible sin archivos enormes.
PPP = 110

#: Resolucion de la VISTA PREVIA de la pantalla. el usuario (2026-10-05): "hacerla
#: mas grande porque no se alcanza a distinguir muy bien". A 110 ppp una hoja
#: horizontal media 1.287 px de ancho, y el dialogo la achicaba a ~800: los
#: numeros de una tabla quedaban de 5-6 px de alto. A 200 ppp mide 2.339 px,
#: asi que se puede ver a todo el ancho de la pantalla -y ampliar al doble-
#: sin que el texto se deshaga.
#:
#: Y pesa MENOS que antes, no mas: se manda en WebP sin perdida en vez de PNG.
#: Medido sobre los informes de prueba: 21 paginas de Rabobank pasan de 0,8 MB
#: (PNG a 110 ppp) a 0,6 MB (WebP a 200 ppp); 11 de Santander, de 0,9 a 0,4.
#: Una hoja de calculo es casi toda blanco y texto, que es justo lo que WebP
#: sin perdida comprime mejor que PNG. "Sin perdida" importa: con perdida el
#: borde de las letras se ensucia, que es lo que se queria arreglar.
PPP_VISTA_PREVIA = 200


@dataclass
class PaginaRenderizada:
    """Una pagina del informe lista para mostrarse en la pantalla."""

    datos: bytes
    tipo: str        # tipo MIME: "image/webp", o "image/png" si no hay WebP
    ancho: int       # en pixeles, para reservar el espacio antes de cargarla
    alto: int

# --------------------------------------------------- Excel por COM, a salvo
#
# Fragmentos de PowerShell reusables por CUALQUIER script que automatice
# Excel via COM -no solo este modulo-. Nacieron aca (exportar_con_excel) pero
# scripts/test_retasacion.py, scripts/test_engine.py y
# scripts/test_exclusiones.py tienen SU PROPIA automatizacion de Excel
# independiente, para verificar celdas despues de generar un informe -y las
# tres tenian la MISMA trampa de la impresora sin el arreglo, encontrada
# cuando un test se quedo colgado de verdad corriendo esta misma sesion. Se
# declaran una sola vez aca y cada script los inserta con `.format()` -que
# NO vuelve a escanear el VALOR insertado en busca de mas '{...}', asi que
# un fragmento con llaves de PowerShell (`try {{ }}`) se puede pasar tal
# cual, sin doble escape-.
#
# Ver el docstring de exportar_con_excel para la explicacion completa de
# por que hace falta cada uno.

GUION_ARRANQUE_IMPRESORA_SEGURA = r"""    $impresoraOriginal = $null
    $impresoraCambiada = $false
    try {
        $actual = Get-CimInstance Win32_Printer -Filter "Default=True" -ErrorAction Stop
        if ($actual -and $actual.Name -ne 'Microsoft Print to PDF') {
            $pdf = Get-CimInstance Win32_Printer -Filter "Name='Microsoft Print to PDF'" -ErrorAction Stop
            if ($pdf) {
                Invoke-CimMethod -InputObject $pdf -MethodName SetDefaultPrinter -ErrorAction Stop | Out-Null
                $impresoraOriginal = $actual.Name
                $impresoraCambiada = $true
            }
        }
    } catch { }
"""

GUION_RESTAURAR_IMPRESORA = r"""    if ($impresoraCambiada -and $impresoraOriginal) {
        try {
            $nombreEscapado = $impresoraOriginal -replace "'", "''"
            $orig = Get-CimInstance Win32_Printer -Filter "Name='$nombreEscapado'" -ErrorAction Stop
            if ($orig) { Invoke-CimMethod -InputObject $orig -MethodName SetDefaultPrinter -ErrorAction Stop | Out-Null }
        } catch { }
    }
"""

GUION_LIBERAR_COM_EXCEL = r"""    if ($libro -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($libro) | Out-Null }
    if ($excel -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    $libro = $null
    $excel = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
"""


class RenderNoDisponible(RuntimeError):
    """Ni Excel ni LibreOffice pudieron exportar el libro a PDF."""


def exportar_con_excel(origen: Path, destino_pdf: Path, hojas=None) -> bool:
    """Exporta con el Excel instalado, via automatizacion COM.

    Se abre en solo lectura y sin alertas para no tocar el archivo del
    usuario ni quedarse esperando un dialogo invisible. Fuerza un recalculo
    completo antes de exportar (CalculateFullRebuild +
    CalculateUntilAsyncQueriesDone): sin esto, Workbooks.Open confia en el
    recalculo automatico de Excel, y un valor guardado en cache podria no
    ser el que corresponde a los datos de este informe.

    Dos trampas de Excel por COM que costaron encontrar (el usuario: "cuando
    apreto vista previa me abre algo de la impresora" -y el primer arreglo
    de este comentario NO la resolvio: el usuario volvio a mandar la captura del
    mismo dialogo. Se reprodujo de verdad -Get-Process y una enumeracion de
    ventanas reales mientras corria una exportacion- antes de intentar de
    nuevo):

    - **Excel consulta la impresora POR DEFECTO del sistema** al exportar,
      aunque el destino sea un PDF -es como arma el tamano de pagina y el
      area imprimible-. En una PC domestica compartida esa impresora por
      defecto es fisica (verificado en esta maquina: "Brother DCP-T420W"),
      y consultar su driver por COM sin nadie enfrente de la pantalla abre
      el dialogo "Configurar impresora" -reproducido dos veces seguidas en
      la misma ventana que muestra el usuario-, que se queda esperando una
      respuesta que nunca llega.

      **El primer intento -`$excel.ActivePrinter = 'Microsoft Print to
      PDF'`- no servia: tira excepcion** (HRESULT 0x800A03EC, verificado
      aislando la linea). La Microsoft Print to PDF vive en un puerto
      virtual (`PORTPROMPT:`, el que le pide al usuario donde guardar) que
      el setter de `ActivePrinter` no sabe nombrar en ningun formato que
      pruebe -ni con el puerto, ni sin el-, asi que el `try/catch` que lo
      envolvia tapaba el error y la impresora seguia siendo la fisica.
      Arreglado a otro nivel: se cambia la impresora **POR DEFECTO DEL
      SISTEMA** (`Win32_Printer.SetDefaultPrinter`, no una propiedad de
      Excel) antes de abrir el libro, y se restaura la que habia despues
      -en el `finally`, pase lo que pase-. Verificado con la reproduccion
      real: con la PDF virtual como default, la exportacion termina sola;
      vuelta a la fisica, vuelve a colgarse con el dialogo. Si la maquina
      no tiene "Microsoft Print to PDF" instalada, no se toca nada y sigue
      con lo que hubiera -degrada, no bloquea-.
    - **`$excel.Quit()` no basta**: PowerShell retiene una referencia COM
      viva hasta que el recolector de basura de .NET corre, y sin eso el
      proceso EXCEL.EXE queda huerfano -sin ventana, pero vivo-. Verificado
      en esta misma maquina: tras una vista previa exitosa, el proceso
      seguia corriendo. Con varias vistas previas seguidas eso acumula
      procesos invisibles compitiendo por el mismo Excel en una PC pensada
      para 2-3 tasadores. Se fuerza el recolector dos veces despues de
      soltar el COM, que es el arreglo ya documentado para este problema.

    Cambiar la impresora por defecto del sistema, aunque sea por los ~15-20
    segundos que tarda una exportacion, es una accion con alcance mas amplio
    que lo que este modulo controla -si algo mas en la PC imprime justo en
    esa ventana, sale por la PDF en vez de la fisica-. Se acepta a proposito:
    el costo de ese cruce raro es un papel que no salio y hay que reimprimir;
    el costo de NO hacerlo es el dialogo colgado, indefinidamente, delante de
    alguien que no sabe que es ni que apretar -que es lo que de verdad estaba
    pasando-.
    """
    # Resolver a absoluta: Excel abierto por COM no hereda el cwd de Python,
    # asi que una ruta relativa se busca en SU directorio de trabajo (no el
    # nuestro) y falla en silencio -Workbooks.Open no encuentra el archivo,
    # el script de PowerShell corta por $ErrorActionPreference, y esta
    # funcion devuelve False como si no hubiera Excel instalado en absoluto.
    origen = Path(origen).resolve()
    destino_pdf = Path(destino_pdf).resolve()
    guion = r"""
$ErrorActionPreference = 'Stop'
$excel = $null
$libro = $null
try {{
{arranque_impresora}
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $libro = $excel.Workbooks.Open('{origen}', $false, $true)
    $excel.CalculateFullRebuild()
    $excel.CalculateUntilAsyncQueriesDone()
    {seleccion}
    $libro.ExportAsFixedFormat(0, '{destino}')
    Write-Output 'OK'
}} finally {{
    if ($libro -ne $null) {{ $libro.Close($false) }}
    if ($excel -ne $null) {{ $excel.Quit() }}
{liberar_com}
{restaurar_impresora}
}}
"""
    seleccion = ""
    if hojas:
        nombres = ",".join("'{}'".format(h.replace("'", "''")) for h in hojas)
        seleccion = "$libro.Worksheets(@({})).Select()".format(nombres)

    completo = guion.format(
        origen=str(origen).replace("'", "''"),
        destino=str(destino_pdf).replace("'", "''"),
        seleccion=seleccion,
        arranque_impresora=GUION_ARRANQUE_IMPRESORA_SEGURA,
        restaurar_impresora=GUION_RESTAURAR_IMPRESORA,
        liberar_com=GUION_LIBERAR_COM_EXCEL,
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", completo],
            capture_output=True, text=True, timeout=180,
        )
        return destino_pdf.exists()
    except Exception:  # noqa: BLE001
        return False


def exportar_con_libreoffice(origen: Path, carpeta: Path) -> bool:
    """Alternativa si no hay Excel."""
    binario = shutil.which("soffice") or shutil.which("libreoffice")
    if not binario:
        return False
    try:
        subprocess.run(
            [binario, "--headless", "--convert-to", "pdf",
             "--outdir", str(carpeta), str(origen)],
            capture_output=True, timeout=180,
        )
        return any(carpeta.glob("*.pdf"))
    except Exception:  # noqa: BLE001
        return False


def pdf_a_imagenes(pdf: Path, carpeta: Path, prefijo: str, ppp: int = PPP):
    """Convierte cada pagina del PDF en un PNG, dejandolo en disco."""
    import pymupdf

    generadas = []
    documento = pymupdf.open(pdf)
    try:
        for numero in range(documento.page_count):
            imagen = documento[numero].get_pixmap(dpi=ppp)
            ruta = carpeta / "{}_pag{:02d}.png".format(prefijo, numero + 1)
            imagen.save(ruta)
            generadas.append(ruta)
    finally:
        documento.close()
    return generadas


def _hay_webp() -> bool:
    try:
        from PIL import features
        return bool(features.check("webp"))
    except Exception:                                    # noqa: BLE001
        return False


def _pdf_a_paginas(pdf: Path, ppp: int = PPP_VISTA_PREVIA) -> List[PaginaRenderizada]:
    """Igual que pdf_a_imagenes, pero en memoria: nada se escribe a disco.

    Para la vista previa, que se pide muchas veces por sesion y no debe
    dejar rastro en una carpeta compartida entre tasadores.

    WebP sin perdida si Pillow lo trae (lo trae en las ruedas oficiales), PNG
    si no: la vista previa tiene que salir igual, solo que mas pesada.
    `method=4, quality=80` es el punto medido: 2 s para 21 paginas a 200 ppp;
    `method=6` ahorra un 3% mas y tarda 60 s.
    """
    import pymupdf

    webp = _hay_webp()
    if webp:
        from PIL import Image

    generadas = []
    documento = pymupdf.open(pdf)
    try:
        for numero in range(documento.page_count):
            imagen = documento[numero].get_pixmap(dpi=ppp)
            if webp:
                lienzo = Image.frombytes(
                    "RGB", (imagen.width, imagen.height), imagen.samples)
                salida = io.BytesIO()
                lienzo.save(salida, format="WEBP", lossless=True, method=4,
                            quality=80)
                datos, tipo = salida.getvalue(), "image/webp"
            else:
                datos, tipo = imagen.tobytes("png"), "image/png"
            generadas.append(PaginaRenderizada(
                datos=datos, tipo=tipo, ancho=imagen.width, alto=imagen.height))
    finally:
        documento.close()
    return generadas


def _renderizar_sync(
    origen: Path, hojas: Optional[List[str]], ppp: int,
) -> Tuple[List[PaginaRenderizada], str]:
    with tempfile.TemporaryDirectory(prefix="fito_render_") as temporal:
        pdf = Path(temporal) / "vista.pdf"
        if exportar_con_excel(origen, pdf, hojas):
            return _pdf_a_paginas(pdf, ppp), "Excel"
        if exportar_con_libreoffice(origen, Path(temporal)):
            pdf = next(Path(temporal).glob("*.pdf"))
            return _pdf_a_paginas(pdf, ppp), "LibreOffice"
    raise RenderNoDisponible("no se encontro Excel ni LibreOffice instalado")


#: Un Excel.exe a la vez. Es una PC hogarena compartida por 2-3 tasadores via
#: Tailscale, no un servidor dimensionado para varias instancias de Excel a
#: la vez compitiendo por CPU/memoria. MotorExcel (openpyxl puro, sin COM) no
#: pasa por aca y no se ve afectado por este cerrojo.
#: OJO: solo protege llamadas DENTRO de este proceso de uvicorn -no protege
#: contra correr scripts/ver_excel.py a mano en paralelo (otro proceso), ni
#: contra varios workers de uvicorn (el proyecto corre uno solo).
_cerrojo = asyncio.Lock()


async def renderizar_a_paginas(
    origen: Path, hojas: Optional[List[str]] = None, ppp: int = PPP_VISTA_PREVIA,
) -> Tuple[List[PaginaRenderizada], str]:
    """Cada pagina del .xlsx como imagen, renderizada con Excel real (o
    LibreOffice si no hay Excel), en memoria -nada queda en disco despues de
    esta llamada, ni siquiera el PDF intermedio.

    Lanza RenderNoDisponible si ni Excel ni LibreOffice estan instalados.
    """
    async with _cerrojo:
        return await asyncio.to_thread(_renderizar_sync, origen, hojas, ppp)
