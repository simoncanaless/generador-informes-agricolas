# Repone la aplicacion si se cayo sola. **PARA CORRERLO A MANO.**
#
# La tarea de Windows YA NO usa este archivo: usa `scripts\vigilante.pyw` con
# `pythonw.exe`, que no tiene consola. Con powershell aparecia una ventana
# negra cada cinco minutos -Windows se la crea ANTES de que `-WindowStyle
# Hidden` alcance a aplicarse- y con la logica en Python la ronda paso a durar
# tres segundos, asi que dejo de ser un parpadeo (2026-09-09).
#
# Este se conserva porque a mano uno SI quiere ver la salida en la consola:
#
#   powershell -ExecutionPolicy Bypass -File scripts\vigilante.ps1
#
# **Aca no hay logica.** La decision -si levantarla, si dejarla apagada, que
# hacer despues de prender el computador- vive en `services/servicio.py`, que
# es tambien lo que usan el panel y la tarea. Tener dos versiones de «cuando
# se levanta» es garantia de que se separen.

param(
    [int]$Puerto = 8000
)

$ErrorActionPreference = "Stop"

$Raiz = Split-Path -Parent $PSScriptRoot

# El Python del usuario. Si no esta donde se espera, se prueba con el del PATH
# antes de rendirse: que el vigilante deje de andar por una ruta cambiada seria
# un fallo mudo, y de esos es de los que nadie se entera hasta que hace falta.
$Python = Join-Path $env:LOCALAPPDATA "Programs\Python\Python311\python.exe"
if (-not (Test-Path $Python)) {
    $delPath = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($delPath) {
        $Python = $delPath.Source
    } else {
        $linea = "{0}  NO se pudo vigilar: no se encontro Python." -f `
                 (Get-Date -Format "dd-MM-yyyy HH:mm:ss")
        $registro = Join-Path $Raiz "data\vigilante.log"
        try { Add-Content -Path $registro -Value $linea -Encoding utf8 } catch { }
        Write-Output $linea
        exit 1
    }
}

& $Python -c "import sys; sys.path.insert(0, r'$Raiz'); from services.servicio import vigilar; sys.exit(vigilar($Puerto))"
exit $LASTEXITCODE
