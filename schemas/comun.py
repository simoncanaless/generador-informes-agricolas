"""Lo que comparten los esquemas de dominio (informe, retasacion, extraccion)."""

from __future__ import annotations

from enum import Enum
from typing import Any


def sin_signo_de_formula(valor: Any) -> Any:
    """Un texto que viene de una persona o de un documento NUNCA llega al
    Excel como formula.

    openpyxl escribe como formula todo texto que empieza con "=", y es justo
    lo que el motor aprovecha para las formulas vivas que declara el mapping.
    Pero el mismo camino toma un nombre, una observacion o lo que la IA leyo
    de una escritura: un "=HYPERLINK(...)" o un "=WEBSERVICE(...)" metido en
    un PDF habria salido como formula activa en el informe que se le manda al
    banco. Medido el 2026-10-05: ningun campo de texto del dominio necesita
    empezar con "=", asi que se le quita el signo y queda el texto.

    No toca enumeraciones (son str, pero sus valores los fija el codigo) ni
    numeros, fechas o diccionarios.
    """
    if isinstance(valor, Enum):
        return valor
    if isinstance(valor, str):
        sin_espacios = valor.lstrip()
        if sin_espacios.startswith("="):
            return sin_espacios.lstrip("= \t")
        return valor
    if isinstance(valor, list):
        return [sin_signo_de_formula(v) for v in valor]
    return valor
