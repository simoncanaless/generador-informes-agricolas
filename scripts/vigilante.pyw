"""Lo que corre la tarea de Windows cada cinco minutos. SIN VENTANA.

    pythonw scripts\\vigilante.pyw

**La extension .pyw y el uso de `pythonw.exe` son el punto de este archivo.**
Hasta el 2026-09-09 la tarea llamaba a `powershell.exe -WindowStyle Hidden`, y
Windows le crea la consola ANTES de que PowerShell alcance a procesar ese
argumento: la ventana negra aparece igual y despues se esconde. Mientras la
ronda duraba medio segundo el parpadeo casi no se veia; al mudar la logica a
Python paso a durar tres segundos y se convirtio en «cada cierto tiempo se me
abre un terminal».

`pythonw.exe` es un binario de subsistema GRAFICO: no tiene consola que
mostrar, asi que no hay nada que parpadee, dure lo que dure la ronda.

**Aca no hay logica**, igual que en `vigilante.ps1`: la decision -si levantarla,
si dejarla apagada, que hacer despues de prender el computador- vive en
`services/servicio.py`, que es tambien lo que usa el panel.

`vigilante.ps1` se conserva para correrlo a mano cuando uno QUIERE ver la
salida en la consola; la tarea ya no lo usa.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.servicio import PUERTO, vigilar                   # noqa: E402


def main() -> int:
    puerto = PUERTO
    if len(sys.argv) > 1:
        try:
            puerto = int(sys.argv[1])
        except ValueError:
            pass
    return vigilar(puerto)


if __name__ == "__main__":
    sys.exit(main())
