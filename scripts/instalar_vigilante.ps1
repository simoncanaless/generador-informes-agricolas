# Programa el vigilante en Windows: repone la aplicacion si se cayo sola.
#
#   powershell -ExecutionPolicy Bypass -File scripts\instalar_vigilante.ps1
#   powershell -ExecutionPolicy Bypass -File scripts\instalar_vigilante.ps1 -Quitar
#
# Crea UNA tarea con dos disparadores: al iniciar sesion, y cada 5 minutos.
# Corre con `pythonw.exe`, que no tiene consola: no aparece ninguna ventana.
#
# **La tarea NO decide nada.** Solo llama a `scripts\vigilante.ps1`, que a su
# vez pregunta a `services/servicio.py`. Y ahi la respuesta depende de lo que
# haya elegido la persona en el panel:
#
#   - la apagaron con el boton  -> no se toca
#   - se cayo sola              -> se repone (si esa preferencia esta puesta)
#   - se prendio el computador  -> NO se levanta, salvo que se haya pedido
#
# O sea que instalar esto NO hace que la pagina aparezca sola al prender el
# computador: eso es una casilla del panel, y viene apagada.
#
# SIN permisos de administrador la tarea queda atada a la sesion del usuario:
# funciona con el computador prendido y la sesion iniciada -aunque este
# bloqueado, que es el caso normal-, pero NO si alguien cierra sesion o el
# computador se reinicia y nadie entra. Para que ande tambien en ese caso hay
# que correr esto como administrador: ahi se agrega el disparador «al arrancar
# el computador».

param([switch]$Quitar)

$ErrorActionPreference = "Stop"

$Nombre = "Fitogreen - vigilante"
$Raiz = Split-Path -Parent $PSScriptRoot
$Script = Join-Path $Raiz "scripts\vigilante.ps1"

if ($Quitar) {
    if (Get-ScheduledTask -TaskName $Nombre -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $Nombre -Confirm:$false
        Write-Output "  Quitada la tarea '$Nombre'."
        Write-Output "  La aplicacion ya no se levanta sola: use el icono del escritorio."
    } else {
        Write-Output "  No estaba programada."
    }
    exit 0
}

$identidad = [Security.Principal.WindowsIdentity]::GetCurrent()
$esAdmin = (New-Object Security.Principal.WindowsPrincipal($identidad)
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# La tarea llama a `pythonw.exe`, NO a powershell.
#
# `-WindowStyle Hidden` no alcanza: Windows le crea la consola a powershell
# ANTES de que este alcance a procesar ese argumento, asi que la ventana negra
# aparece igual y despues se esconde. Mientras la ronda duraba medio segundo el
# parpadeo casi no se veia; con la logica en Python paso a durar tres segundos
# y se convirtio en «cada cierto tiempo se me abre un terminal» (2026-09-09).
#
# `pythonw.exe` es un binario de subsistema GRAFICO: no tiene consola que
# mostrar, asi que no hay nada que parpadee dure lo que dure la ronda. Y de
# paso se ahorra la capa de PowerShell entera.
$Entrada = Join-Path $Raiz "scripts\vigilante.pyw"
if (-not (Test-Path $Entrada)) { throw "no se encontro $Entrada" }

$Pythonw = Join-Path $env:LOCALAPPDATA "Programs\Python\Python311\pythonw.exe"
if (-not (Test-Path $Pythonw)) {
    $delPath = Get-Command pythonw.exe -ErrorAction SilentlyContinue
    if (-not $delPath) {
        throw ("no se encontro pythonw.exe. Es el que evita que aparezca una " +
               "ventana cada cinco minutos.")
    }
    $Pythonw = $delPath.Source
}

$accion = New-ScheduledTaskAction -Execute $Pythonw `
    -Argument "`"$Entrada`"" `
    -WorkingDirectory $Raiz

# Cada 5 minutos, para siempre. `Once` con repeticion es la forma de decir
# «cada tanto» en el Programador de tareas.
$cada5 = New-ScheduledTaskTrigger -Once -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Minutes 5)
$disparadores = @($cada5, (New-ScheduledTaskTrigger -AtLogOn -User $identidad.Name))
if ($esAdmin) {
    $disparadores += New-ScheduledTaskTrigger -AtStartup
}

# Que NO se detenga sola: por defecto Windows mata las tareas a las 3 dias y
# no las corre a bateria. Un vigilante que se apaga solo no vigila.
$ajustes = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5)

$principal = if ($esAdmin) {
    New-ScheduledTaskPrincipal -UserId $identidad.Name -LogonType S4U -RunLevel Highest
} else {
    New-ScheduledTaskPrincipal -UserId $identidad.Name -LogonType Interactive
}

Register-ScheduledTask -TaskName $Nombre -Action $accion -Trigger $disparadores `
    -Settings $ajustes -Principal $principal -Force `
    -Description ("Mantiene levantado el asistente de informes de Fitogreen. " +
                  "Comprueba cada 5 minutos si responde y lo levanta si no.") | Out-Null

Write-Output "  Tarea '$Nombre' programada."
Write-Output "    - al iniciar sesion"
if ($esAdmin) { Write-Output "    - al arrancar el computador" }
Write-Output "    - y cada 5 minutos"
Write-Output ""
if (-not $esAdmin) {
    Write-Output "  OJO: sin permisos de administrador la tarea corre solo con la"
    Write-Output "  sesion iniciada. Con el computador bloqueado funciona; si alguien"
    Write-Output "  cierra sesion, no. Para cubrir tambien ese caso, vuelva a correr"
    Write-Output "  esto haciendo clic derecho en PowerShell > Ejecutar como"
    Write-Output "  administrador."
    Write-Output ""
}
Write-Output "  Registro de lo que hace:  data\vigilante.log"
Write-Output "  Para quitarla:            -Quitar"
