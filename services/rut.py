"""El RUT chileno: normalizacion, formato de pantalla, y el digito
verificador -que es una formula (modulo 11), no un caracter cualquiera.

Vivia solo dentro de services/usuarios.py (para el login), pero el mismo
calculo hace falta en cualquier campo de RUT del proyecto -cliente,
propietario, tasador de una retasacion-, y hasta ahora esos campos solo
comprobaban que hubiera un guion en el texto, sin verificar que el digito
fuera el que corresponde. Un RUT mal tecleado (ej. un digito transpuesto)
pasaba sin que nadie se enterara. Se separa en su propio modulo para que
schemas/*.py lo use sin arrastrar el cifrado de contrasenas de
services/usuarios.py, que es un asunto aparte.
"""

from __future__ import annotations

import re

__all__ = ["normalizar_rut", "formatear_rut", "rut_valido", "digito_verificador",
          "RutInvalido"]


class RutInvalido(ValueError):
    """El RUT no tiene forma de RUT o su digito verificador no cuadra."""


def digito_verificador(cuerpo: str) -> str:
    """El digito que le corresponde a un cuerpo de RUT (modulo 11).

    Se calcula, no se confia: un RUT mal tecleado con un digito cualquiera
    seria un usuario -o un cliente- que despues nadie puede identificar.
    """
    suma, factor = 0, 2
    for caracter in reversed(cuerpo):
        suma += int(caracter) * factor
        factor = 2 if factor == 7 else factor + 1
    resto = 11 - (suma % 11)
    if resto == 11:
        return "0"
    if resto == 10:
        return "K"
    return str(resto)


def normalizar_rut(rut: str) -> str:
    """Deja el RUT en su forma canonica: '12345678-5'.

    Se guarda SIN puntos y con el digito en mayuscula, porque la misma
    persona lo escribe de cinco maneras distintas.
    """
    limpio = re.sub(r"[^0-9kK]", "", rut or "").upper()
    if len(limpio) < 2:
        raise RutInvalido("el RUT esta incompleto")
    cuerpo, digito = limpio[:-1], limpio[-1]
    if not cuerpo.isdigit():
        raise RutInvalido("el cuerpo del RUT tiene que ser solo numeros")
    esperado = digito_verificador(cuerpo)
    if digito != esperado:
        raise RutInvalido(
            "el digito verificador no corresponde: {}-{} deberia terminar en {}"
            .format(cuerpo, digito, esperado))
    return "{}-{}".format(int(cuerpo), digito)


def rut_valido(rut: str) -> bool:
    try:
        normalizar_rut(rut)
        return True
    except RutInvalido:
        return False


def formatear_rut(rut: str) -> str:
    """Como se muestra en pantalla: '21.280.840-7'."""
    try:
        canonico = normalizar_rut(rut)
    except RutInvalido:
        return rut
    cuerpo, digito = canonico.split("-")
    return "{}-{}".format("{:,}".format(int(cuerpo)).replace(",", "."), digito)
