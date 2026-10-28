"""Freno a quien prueba contrasenas.

Hasta el 2026-10-05 se podian probar claves sin limite: cada intento costaba
~50 ms de scrypt, o sea unas 70.000 por hora desde cualquier equipo que
llegara al servidor. Ahora, despues de 5 fallos seguidos, cada intento nuevo
tiene que esperar -30 s, despues 1 min, 2, 4... hasta 15 min-, y la espera se
cuenta por RUT y por equipo a la vez: cambiar de RUT no sirve para seguir
probando desde el mismo equipo, y cambiar de equipo no sirve para seguir con
el mismo RUT.

Vive en memoria a proposito. Reiniciar el servidor lo borra, y esta bien: el
servidor se reinicia muy poco, y una tabla nueva en la base para esto seria
mas codigo que el riesgo que cubre. Lo que importa es que probar a lo bruto
deje de ser practico, y eso se logra igual.
"""

from __future__ import annotations

import threading
import time
from typing import Dict, Iterable, List, Optional

#: Fallos que se perdonan antes de empezar a frenar: equivocarse dos o tres
#: veces escribiendo la clave es lo normal.
FALLOS_LIBRES = 5
#: La primera espera, y el tope.
ESPERA_INICIAL = 30.0
ESPERA_MAXIMA = 15 * 60.0
#: Un fallo viejo se olvida: quien se equivoco ayer no empieza hoy castigado.
OLVIDO = 60 * 60.0
#: Tope de claves recordadas, para que un barrido de RUT inventados no haga
#: crecer la memoria sin fin.
MAXIMO_DE_CLAVES = 5000


class Intentos:
    def __init__(self, reloj=time.monotonic):
        self._reloj = reloj
        self._fallos: Dict[str, List[float]] = {}
        self._cerrojo = threading.Lock()

    def _vigentes(self, clave: str, ahora: float) -> List[float]:
        marcas = [t for t in self._fallos.get(clave, []) if ahora - t < OLVIDO]
        if marcas:
            self._fallos[clave] = marcas
        else:
            self._fallos.pop(clave, None)
        return marcas

    def espera(self, claves: Iterable[str]) -> float:
        """Segundos que faltan para poder intentar de nuevo (0 = ya se puede).

        Con varias claves (RUT y equipo) manda la mas restrictiva."""
        ahora = self._reloj()
        mayor = 0.0
        with self._cerrojo:
            for clave in claves:
                marcas = self._vigentes(clave, ahora)
                exceso = len(marcas) - FALLOS_LIBRES
                if exceso < 0:
                    continue
                castigo = min(ESPERA_MAXIMA, ESPERA_INICIAL * (2 ** exceso))
                mayor = max(mayor, marcas[-1] + castigo - ahora)
        return max(0.0, mayor)

    def fallo(self, claves: Iterable[str]) -> None:
        ahora = self._reloj()
        with self._cerrojo:
            if len(self._fallos) > MAXIMO_DE_CLAVES:
                # Se descarta lo mas viejo, no lo mas nuevo: un ataque en curso
                # no debe poder limpiar su propio rastro inflando la tabla.
                for clave in sorted(self._fallos, key=lambda c: self._fallos[c][-1])[
                        :len(self._fallos) - MAXIMO_DE_CLAVES]:
                    self._fallos.pop(clave, None)
            for clave in claves:
                self._fallos.setdefault(clave, []).append(ahora)

    def acierto(self, claves: Iterable[str]) -> None:
        with self._cerrojo:
            for clave in claves:
                self._fallos.pop(clave, None)


def claves_de(rut: Optional[str], ip: Optional[str]) -> List[str]:
    claves = []
    if rut:
        claves.append("rut:" + rut[:32])
    if ip:
        claves.append("ip:" + ip)
    return claves


def texto_de_espera(segundos: float) -> str:
    if segundos < 60:
        return "{} segundos".format(max(1, int(round(segundos))))
    minutos = int(round(segundos / 60.0))
    return "1 minuto" if minutos == 1 else "{} minutos".format(minutos)
