"""Identidad de quien usa el programa: RUT, contrasena y sesion.

El RUT es la llave. No es capricho: es el identificador que ya usan en todo lo
demas -el cliente, el propietario de la escritura- y no hay dos personas con el
mismo. Si manana entran cinco tasadores mas, no hay que inventar nombres de
usuario ni resolver que pasa con dos "juan".

**Las contrasenas no se guardan.** Se guarda un `scrypt` con sal por usuario, y
comparar es volver a derivar y contrastar en tiempo constante. `scrypt` viene en
la biblioteca estandar de Python y es memory-hard, asi que no hace falta traer
`bcrypt` ni `argon2` para hacerlo bien.
"""

from __future__ import annotations

import hashlib
import hmac
import os
from typing import Optional, Tuple

from services.rut import (  # noqa: F401 -- reexportado, ver __all__
    RutInvalido, digito_verificador, formatear_rut, normalizar_rut, rut_valido,
)

__all__ = [
    "normalizar_rut", "formatear_rut", "rut_valido", "digito_verificador",
    "cifrar_contrasena", "contrasena_correcta", "comparar_con_senuelo",
    "huella_de_sesion", "RutInvalido",
]

#: Parametros de scrypt. n=2^14 tarda ~50 ms por intento en un equipo de
#: escritorio: imperceptible al entrar, y suficiente para que probar
#: contrasenas a lo bruto no sea practico.
_N, _R, _P, _SAL, _LARGO = 2 ** 14, 8, 1, 16, 32


def cifrar_contrasena(contrasena: str) -> str:
    """Devuelve 'scrypt$n$r$p$sal$hash', todo lo necesario para verificar.

    Los parametros van DENTRO del texto guardado: el dia que haya que subirlos
    -porque el hardware mejora- las contrasenas viejas siguen verificandose con
    los suyos, y se re-cifran al entrar.
    """
    if not contrasena or len(contrasena) < 8:
        raise ValueError("la contrasena tiene que tener al menos 8 caracteres")
    sal = os.urandom(_SAL)
    derivada = hashlib.scrypt(
        contrasena.encode("utf-8"), salt=sal, n=_N, r=_R, p=_P, dklen=_LARGO)
    return "scrypt${}${}${}${}${}".format(
        _N, _R, _P, sal.hex(), derivada.hex())


def huella_de_sesion(guardada: str, secreto: str) -> str:
    """Lo que la cookie lleva de la contrasena: una firma de su cifrado.

    Cuando la contrasena cambia, el cifrado guardado cambia y la firma deja de
    coincidir, asi que toda sesion abierta con la clave anterior -la de otro
    computador, o una cookie que alguien copio- deja de valer sola. Va firmada
    con el secreto del servidor (HMAC) y no como hash suelto: la cookie de
    sesion se puede leer en el navegador, y asi lo que se lee no le sirve a
    nadie fuera del servidor.
    """
    return hmac.new(secreto.encode("utf-8"), (guardada or "").encode("utf-8"),
                    hashlib.sha256).hexdigest()[:32]


#: Un cifrado de verdad, de una clave que nadie tiene, para comparar contra el
#: cuando el RUT no existe. Sin esto, un RUT inexistente contestaba en 1 ms y
#: uno real en 50 ms (lo que tarda scrypt): midiendo el tiempo se sabia que
#: RUT tienen cuenta, aunque el mensaje fuera el mismo.
_SENUELO: Optional[str] = None


def comparar_con_senuelo(contrasena: str) -> bool:
    global _SENUELO
    if _SENUELO is None:
        _SENUELO = cifrar_contrasena(os.urandom(18).hex())
    contrasena_correcta(contrasena, _SENUELO)
    return False


def contrasena_correcta(contrasena: str, guardada: str) -> bool:
    """Compara en tiempo constante. Nunca revienta: una guardada corrupta es
    'no coincide', no una excepcion que delate que el usuario existe."""
    try:
        algoritmo, n, r, p, sal, esperado = guardada.split("$")
        if algoritmo != "scrypt":
            return False
        derivada = hashlib.scrypt(
            (contrasena or "").encode("utf-8"), salt=bytes.fromhex(sal),
            n=int(n), r=int(r), p=int(p), dklen=len(esperado) // 2)
        return hmac.compare_digest(derivada.hex(), esperado)
    except Exception:                                     # noqa: BLE001
        return False
