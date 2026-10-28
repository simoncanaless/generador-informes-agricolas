"""Rotacion de modelos: que la aplicacion no se detenga porque un modelo se agote.

Mantiene una cadena ordenada de candidatos ``(clave, modelo)`` y va probando
hacia abajo hasta que uno responde. Cada desenlace tiene una politica distinta:

  ``CUOTA``       429/RESOURCE_EXHAUSTED. El modelo entra en cuarentena durante
                  un tiempo y se salta al siguiente. La cuarentena escala x4 en
                  cada reincidencia: si la cuota es diaria, no se malgasta una
                  llamada por minuto contra un modelo que no va a responder.
  ``TRANSITORIO`` 503/500/timeout. El servidor esta saturado, no agotado: se
                  reintenta el MISMO modelo con espera creciente antes de bajar.
  ``DESCARTAR``   404. El modelo no existe o el plan no lo incluye; se elimina
                  de la rotacion para siempre (caso real: gemini-2.5-flash-lite
                  dejo de servirse a claves nuevas).
  ``FATAL``       400/401/403. Es culpa nuestra o de la clave; cambiar de modelo
                  no arregla nada, asi que se propaga de inmediato.

Este modulo no importa google.genai: recibe una funcion que ejecuta la llamada y
otra que clasifica el error. Asi la politica se prueba entera sin red (DIP).
"""

from __future__ import annotations

import asyncio
import time
from dataclasses import dataclass, field
from enum import Enum
from typing import Awaitable, Callable, Dict, Generic, List, Optional, TypeVar

T = TypeVar("T")

#: Espera entre reintentos del mismo modelo ante un fallo transitorio.
ESPERA_TRANSITORIA_BASE = 1.5


class Desenlace(str, Enum):
    """Como interpretar un error para decidir el siguiente paso."""

    CUOTA = "cuota"
    TRANSITORIO = "transitorio"
    DESCARTAR = "descartar"
    FATAL = "fatal"


class EstadoCandidato(str, Enum):
    DISPONIBLE = "disponible"
    EN_CUARENTENA = "en_cuarentena"
    DESCARTADO = "descartado"


class SinModelosDisponibles(RuntimeError):
    """Ningun candidato de la cadena puede atender la peticion ahora mismo."""

    def __init__(self, mensaje: str, segundos_para_reintentar: Optional[float] = None):
        super().__init__(mensaje)
        self.segundos_para_reintentar = segundos_para_reintentar


@dataclass
class Candidato:
    """Un par (clave, modelo) con su historial."""

    modelo: str
    indice_clave: int
    prioridad: int

    estado: EstadoCandidato = EstadoCandidato.DISPONIBLE
    disponible_desde: float = 0.0
    reincidencias: int = 0
    exitos: int = 0
    fallos: int = 0
    ultimo_error: Optional[str] = None

    @property
    def etiqueta(self) -> str:
        return "{}#clave{}".format(self.modelo, self.indice_clave)

    def utilizable(self, ahora: float) -> bool:
        if self.estado is EstadoCandidato.DESCARTADO:
            return False
        if self.estado is EstadoCandidato.EN_CUARENTENA:
            return ahora >= self.disponible_desde
        return True


@dataclass
class Respuesta(Generic[T]):
    """Resultado de una llamada, con la trazabilidad de quien la atendio.

    En un informe bancario importa poder decir que modelo genero cada dato.
    """

    valor: T
    modelo: str
    indice_clave: int
    intentos: int
    duracion_s: float
    degradado: bool = False
    incidencias: List[str] = field(default_factory=list)


class RouterModelos:
    """Cadena de respaldo sobre modelos y claves.

    El orden es: por prioridad de modelo y, dentro de cada modelo, por clave.
    Asi se agota el mejor modelo en todas las claves antes de bajar de calidad.
    """

    def __init__(
        self,
        modelos: List[str],
        n_claves: int = 1,
        cooldown_base: float = 60.0,
        cooldown_maximo: float = 3600.0,
        reintentos_transitorios: int = 2,
        reloj: Callable[[], float] = time.monotonic,
        dormir: Optional[Callable[[float], Awaitable[None]]] = None,
    ):
        if not modelos:
            raise ValueError("la cadena de modelos esta vacia")
        if n_claves < 1:
            raise ValueError("hace falta al menos una clave de API")

        self.cooldown_base = cooldown_base
        self.cooldown_maximo = cooldown_maximo
        self.reintentos_transitorios = reintentos_transitorios
        self._reloj = reloj
        self._dormir = dormir or asyncio.sleep
        self._cerrojo = asyncio.Lock()

        self.candidatos: List[Candidato] = [
            Candidato(modelo=modelo, indice_clave=clave, prioridad=prioridad)
            for prioridad, modelo in enumerate(modelos)
            for clave in range(n_claves)
        ]

    # ------------------------------------------------------------ consultas

    def disponibles(self) -> List[Candidato]:
        ahora = self._reloj()
        return [c for c in self.candidatos if c.utilizable(ahora)]

    def proxima_disponibilidad(self) -> Optional[float]:
        """Segundos hasta que vuelva a haber algun candidato, o None si ninguno."""
        ahora = self._reloj()
        esperas = [
            c.disponible_desde - ahora
            for c in self.candidatos
            if c.estado is EstadoCandidato.EN_CUARENTENA
        ]
        return max(0.0, min(esperas)) if esperas else None

    def instantanea(self) -> List[Dict[str, object]]:
        """Estado legible de la cadena, para el endpoint de salud."""
        ahora = self._reloj()
        return [
            {
                "modelo": c.modelo,
                "clave": c.indice_clave,
                "prioridad": c.prioridad,
                "estado": c.estado.value,
                "espera_s": (
                    round(max(0.0, c.disponible_desde - ahora), 1)
                    if c.estado is EstadoCandidato.EN_CUARENTENA
                    else 0.0
                ),
                "exitos": c.exitos,
                "fallos": c.fallos,
                "ultimo_error": c.ultimo_error,
            }
            for c in self.candidatos
        ]

    # --------------------------------------------------------- transiciones

    def registrar_exito(self, candidato: Candidato) -> None:
        candidato.exitos += 1
        candidato.estado = EstadoCandidato.DISPONIBLE
        candidato.disponible_desde = 0.0
        candidato.reincidencias = 0  # se recupero: se perdona el historial
        candidato.ultimo_error = None

    def registrar_cuota(
        self, candidato: Candidato, espera: Optional[float] = None, detalle: str = ""
    ) -> float:
        """Cuarentena por cuota agotada. Devuelve los segundos aplicados.

        Si la API indico cuanto esperar (RetryInfo), se respeta. Si no, se usa
        una espera que crece x4 por reincidencia hasta el techo configurado.
        """
        candidato.fallos += 1
        if espera is None:
            espera = min(
                self.cooldown_base * (4 ** candidato.reincidencias),
                self.cooldown_maximo,
            )
        candidato.reincidencias += 1
        candidato.estado = EstadoCandidato.EN_CUARENTENA
        candidato.disponible_desde = self._reloj() + espera
        candidato.ultimo_error = detalle or "cuota agotada"
        return espera

    def registrar_descarte(self, candidato: Candidato, detalle: str = "") -> None:
        candidato.fallos += 1
        candidato.estado = EstadoCandidato.DESCARTADO
        candidato.ultimo_error = detalle or "modelo no disponible"

    def registrar_transitorio(self, candidato: Candidato, detalle: str = "") -> None:
        candidato.fallos += 1
        candidato.ultimo_error = detalle or "error transitorio"

    # ------------------------------------------------------------ ejecucion

    async def ejecutar(
        self,
        operacion: Callable[[str, int], Awaitable[T]],
        clasificar: Callable[[BaseException], Desenlace],
        descripcion: str = "llamada",
    ) -> Respuesta[T]:
        """Ejecuta ``operacion(modelo, indice_clave)`` bajando por la cadena.

        Devuelve la primera respuesta buena. Si la cadena se agota entera,
        lanza :class:`SinModelosDisponibles` diciendo cuando reintentar.
        """
        inicio = self._reloj()
        intentos = 0
        incidencias: List[str] = []
        mejor_prioridad: Optional[int] = None

        for candidato in sorted(
            self.candidatos, key=lambda c: (c.prioridad, c.indice_clave)
        ):
            if not candidato.utilizable(self._reloj()):
                continue
            if mejor_prioridad is None:
                mejor_prioridad = candidato.prioridad

            for reintento in range(self.reintentos_transitorios + 1):
                intentos += 1
                try:
                    valor = await operacion(candidato.modelo, candidato.indice_clave)
                except asyncio.CancelledError:
                    raise
                except BaseException as error:  # noqa: BLE001 - se reclasifica
                    desenlace = clasificar(error)
                    detalle = "{}: {}".format(type(error).__name__, error)[:200]

                    if desenlace is Desenlace.FATAL:
                        raise

                    if desenlace is Desenlace.DESCARTAR:
                        async with self._cerrojo:
                            self.registrar_descarte(candidato, detalle)
                        incidencias.append(
                            "{} descartado ({})".format(candidato.etiqueta, detalle)
                        )
                        break

                    if desenlace is Desenlace.CUOTA:
                        async with self._cerrojo:
                            espera = self.registrar_cuota(
                                candidato, _espera_sugerida(error), detalle
                            )
                        incidencias.append(
                            "{} sin cuota, en espera {:.0f}s".format(
                                candidato.etiqueta, espera
                            )
                        )
                        break

                    # TRANSITORIO
                    async with self._cerrojo:
                        self.registrar_transitorio(candidato, detalle)
                    if reintento < self.reintentos_transitorios:
                        await self._dormir(ESPERA_TRANSITORIA_BASE * (2 ** reintento))
                        continue
                    incidencias.append(
                        "{} inestable tras {} intentos ({})".format(
                            candidato.etiqueta, reintento + 1, detalle
                        )
                    )
                    break
                else:
                    async with self._cerrojo:
                        self.registrar_exito(candidato)
                    return Respuesta(
                        valor=valor,
                        modelo=candidato.modelo,
                        indice_clave=candidato.indice_clave,
                        intentos=intentos,
                        duracion_s=self._reloj() - inicio,
                        degradado=candidato.prioridad > (mejor_prioridad or 0),
                        incidencias=incidencias,
                    )

        espera = self.proxima_disponibilidad()
        raise SinModelosDisponibles(
            "{}: se agotaron los {} candidatos de la cadena. {}".format(
                descripcion,
                len(self.candidatos),
                "; ".join(incidencias) or "sin candidatos utilizables",
            ),
            segundos_para_reintentar=espera,
        )


def _espera_sugerida(error: BaseException) -> Optional[float]:
    """Lee el ``retryDelay`` que la API adjunta a algunos 429.

    Se busca de forma tolerante para no acoplar el router al formato exacto de
    un proveedor concreto.
    """
    detalles = getattr(error, "details", None)
    if not isinstance(detalles, dict):
        return None

    pila = [detalles]
    while pila:
        actual = pila.pop()
        if isinstance(actual, dict):
            bruto = actual.get("retryDelay")
            if isinstance(bruto, str) and bruto.endswith("s"):
                try:
                    return float(bruto[:-1])
                except ValueError:
                    pass
            pila.extend(v for v in actual.values() if isinstance(v, (dict, list)))
        elif isinstance(actual, list):
            pila.extend(v for v in actual if isinstance(v, (dict, list)))
    return None
