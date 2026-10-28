"""Alta y mantencion de las cuentas. Se corre a mano, desde el servidor.

No hay pantalla de registro a proposito: son tres personas de la misma oficina
y una pagina de alta abierta a la red seria una puerta que nadie necesita.

    python scripts/usuarios.py --listar
    python scripts/usuarios.py --crear 21.280.840-7 "el usuario Canales"
    python scripts/usuarios.py --clave 21.280.840-7
    python scripts/usuarios.py --desactivar 14.409.133-7

La contrasena NUNCA se pasa por la linea de comandos: quedaria en el historial
del terminal y a la vista de cualquiera que corra `history`. Se pide aparte,
sin eco en pantalla.
"""

import argparse
import asyncio
import getpass
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from db.repositorios import RepositorioUsuarios                  # noqa: E402
from services.usuarios import (                                  # noqa: E402
    RutInvalido, cifrar_contrasena, formatear_rut, normalizar_rut,
)


class SeArrepintio(Exception):
    """Corto el prompt con Ctrl+C. No es un error: no se cambia nada."""


def _sin_eco(prompt: str) -> str:
    """La clave, sin mostrarla — y sin colgarse si no hay consola de verdad.

    En Windows `getpass` lee del handle de la consola con `msvcrt.getwch()`,
    no de la entrada estandar. Si la clave viene por una tuberia
    (`echo x | ...`, una prueba, una terminal embebida que no le pasa las
    teclas) NO falla: **se queda esperando para siempre** una tecla que nunca
    llega, y la unica salida es Ctrl+C. Por eso se comprueba antes si hay
    consola, en vez de intentarlo y atrapar el error: no hay error que
    atrapar.
    """
    if not sys.stdin.isatty():
        linea = sys.stdin.readline()
        if not linea:
            raise SeArrepintio()
        return linea.rstrip("\r\n")
    try:
        return getpass.getpass(prompt)
    except KeyboardInterrupt:
        raise SeArrepintio()


def _pedir_clave(para: str) -> str:
    # Se avisa a proposito: no se ve NADA al teclear -- ni asteriscos -- y sin
    # este aviso la pantalla parece colgada y uno corta con Ctrl+C.
    print("  Escriba la contrasena y pulse Enter.")
    print("  No se vera nada mientras la escribe: es normal.")
    while True:
        una = _sin_eco("  Contrasena para {}: ".format(para))
        if len(una) < 8:
            print("  Muy corta: al menos 8 caracteres.")
            continue
        otra = _sin_eco("  Repitala: ")
        if una != otra:
            print("  No coinciden.")
            continue
        return una


async def listar() -> int:
    repo = RepositorioUsuarios()
    filas = await repo.listar()
    if not filas:
        print("  No hay cuentas todavia.")
        return 0
    print("  {:<16} {:<28} {:<8} {}".format("RUT", "NOMBRE", "ACTIVO", "ULTIMA ENTRADA"))
    for f in filas:
        print("  {:<16} {:<28} {:<8} {}".format(
            formatear_rut(f.rut), f.nombre[:28], "si" if f.activo else "NO",
            f.ultima_entrada.strftime("%d-%m-%Y %H:%M") if f.ultima_entrada
            else "nunca"))
    return 0


async def crear(rut: str, nombre: str, clave: str = None) -> int:
    repo = RepositorioUsuarios()
    try:
        canonico = normalizar_rut(rut)
    except RutInvalido as error:
        print("  RUT invalido: {}".format(error))
        return 1
    if await repo.buscar_por_rut(canonico):
        print("  Ya existe una cuenta con el RUT {}".format(formatear_rut(canonico)))
        return 1
    clave = clave or _pedir_clave(nombre)
    await repo.crear(canonico, nombre, cifrar_contrasena(clave))
    print("  Cuenta creada: {} · {}".format(formatear_rut(canonico), nombre))
    return 0


async def cambiar(rut: str, clave: str = None) -> int:
    repo = RepositorioUsuarios()
    try:
        canonico = normalizar_rut(rut)
    except RutInvalido as error:
        print("  RUT invalido: {}".format(error))
        return 1
    fila = await repo.buscar_por_rut(canonico)
    if fila is None:
        print("  No hay ninguna cuenta con ese RUT.")
        return 1
    clave = clave or _pedir_clave(fila.nombre)
    await repo.cambiar_clave(str(fila.id), cifrar_contrasena(clave))
    print("  Contrasena cambiada para {}".format(fila.nombre))
    return 0


async def renombrar(rut: str, nombre: str) -> int:
    """Cambia el nombre visible. El RUT no se toca: es la identidad."""
    from db.sesion import sesion
    import db.modelos as m
    repo = RepositorioUsuarios()
    try:
        canonico = normalizar_rut(rut)
    except RutInvalido as error:
        print("  RUT invalido: {}".format(error))
        return 1
    fila = await repo.buscar_por_rut(canonico)
    if fila is None:
        print("  No hay ninguna cuenta con ese RUT.")
        return 1
    antes = fila.nombre
    async with sesion() as s:
        actual = await s.get(m.Usuario, fila.id)
        actual.nombre = nombre
        await s.flush()
    print("  {} -> {}".format(antes, nombre))
    return 0


async def desactivar(rut: str, activar: bool = False) -> int:
    """Se desactiva, no se borra: los informes que hizo siguen siendo suyos y
    borrar la cuenta los dejaria sin dueno."""
    from db.sesion import sesion
    repo = RepositorioUsuarios()
    try:
        canonico = normalizar_rut(rut)
    except RutInvalido as error:
        print("  RUT invalido: {}".format(error))
        return 1
    fila = await repo.buscar_por_rut(canonico)
    if fila is None:
        print("  No hay ninguna cuenta con ese RUT.")
        return 1
    async with sesion() as s:
        import db.modelos as m
        actual = await s.get(m.Usuario, fila.id)
        actual.activo = activar
        await s.flush()
    print("  {} quedo {}".format(fila.nombre, "activa" if activar else "DESACTIVADA"))
    return 0


def main() -> int:
    p = argparse.ArgumentParser(description="Cuentas del asistente de informes")
    p.add_argument("--listar", action="store_true")
    p.add_argument("--crear", nargs=2, metavar=("RUT", "NOMBRE"))
    p.add_argument("--clave", metavar="RUT", help="cambia la contrasena")
    p.add_argument("--nombre", nargs=2, metavar=("RUT", "NOMBRE"),
                   help="cambia el nombre que se muestra")
    p.add_argument("--desactivar", metavar="RUT")
    p.add_argument("--activar", metavar="RUT")
    args = p.parse_args()

    print("=" * 66)
    print("CUENTAS")
    print("=" * 66)
    try:
        return _hacer(args)
    except (SeArrepintio, KeyboardInterrupt):
        # Cortar el prompt no es un fallo del programa: no se cambio nada y no
        # hace falta un traceback de cuarenta lineas para decirlo.
        print("\n  Cancelado. No se cambio ninguna contrasena.")
        return 1


def _hacer(args) -> int:
    if args.crear:
        return asyncio.run(crear(args.crear[0], args.crear[1]))
    if args.clave:
        return asyncio.run(cambiar(args.clave))
    if args.nombre:
        return asyncio.run(renombrar(args.nombre[0], args.nombre[1]))
    if args.desactivar:
        return asyncio.run(desactivar(args.desactivar, activar=False))
    if args.activar:
        return asyncio.run(desactivar(args.activar, activar=True))
    return asyncio.run(listar())


if __name__ == "__main__":
    sys.exit(main())
