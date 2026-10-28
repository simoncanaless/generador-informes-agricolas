"""Prueba de la politica de rotacion de modelos.

Todo con errores simulados, reloj falso y sleep instantaneo: verifica la
politica completa sin tocar la red ni gastar un solo token.

Uso: python scripts/test_router.py
"""

import asyncio
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.model_router import (  # noqa: E402
    Desenlace, EstadoCandidato, RouterModelos, SinModelosDisponibles,
)

fallos = []
comprobaciones = 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)


# ------------------------------------------------------------- utilidades

class ErrorFalso(Exception):
    """Imita un google.genai APIError: codigo, estado y detalles anidados."""

    def __init__(self, code, status, details=None):
        super().__init__("{} {}".format(code, status))
        self.code = code
        self.status = status
        self.details = details or {}


def clasificar(error):
    codigo = getattr(error, "code", None)
    if codigo == 429:
        return Desenlace.CUOTA
    if codigo == 404:
        return Desenlace.DESCARTAR
    if codigo in (500, 503) or isinstance(error, asyncio.TimeoutError):
        return Desenlace.TRANSITORIO
    return Desenlace.FATAL


class Reloj:
    """Reloj manipulable para no esperar cuarentenas de verdad."""

    def __init__(self):
        self.t = 1000.0

    def __call__(self):
        return self.t

    def avanzar(self, segundos):
        self.t += segundos


async def sin_dormir(_segundos):
    return None


def construir(modelos, n_claves=1, **kw):
    reloj = Reloj()
    router = RouterModelos(
        modelos, n_claves=n_claves, reloj=reloj, dormir=sin_dormir, **kw
    )
    return router, reloj


def guion(respuestas):
    """Crea una operacion que devuelve o lanza segun el modelo que le toque.

    ``respuestas`` es {modelo: valor | excepcion | [secuencia]}.
    """
    llamadas = []

    async def operacion(modelo, indice_clave):
        llamadas.append((modelo, indice_clave))
        resultado = respuestas.get(modelo, "ok:{}".format(modelo))
        if isinstance(resultado, list):
            resultado = resultado.pop(0) if resultado else "ok:{}".format(modelo)
        if isinstance(resultado, BaseException):
            raise resultado
        return resultado

    return operacion, llamadas


# ---------------------------------------------------------------- pruebas

async def caso_camino_feliz():
    print("\n[1] El primer modelo responde")
    router, _ = construir(["a", "b", "c"])
    operacion, llamadas = guion({})
    r = await router.ejecutar(operacion, clasificar)
    comprobar(r.valor == "ok:a", "devuelve el valor del primer modelo")
    comprobar(r.modelo == "a", "informa que modelo atendio")
    comprobar(not r.degradado, "no marca degradacion")
    comprobar(len(llamadas) == 1, "no llama a los modelos de respaldo",
              "llamadas={}".format(llamadas))


async def caso_cuota_baja_al_siguiente():
    print("\n[2] 429: cuarentena y salto al siguiente")
    router, _ = construir(["a", "b", "c"])
    operacion, llamadas = guion({"a": ErrorFalso(429, "RESOURCE_EXHAUSTED")})
    r = await router.ejecutar(operacion, clasificar)
    comprobar(r.modelo == "b", "responde el segundo de la cadena")
    comprobar(r.degradado, "marca la respuesta como degradada")
    comprobar(llamadas == [("a", 0), ("b", 0)], "probo 'a' antes que 'b'",
              str(llamadas))
    estado = {c["modelo"]: c for c in router.instantanea()}
    comprobar(estado["a"]["estado"] == EstadoCandidato.EN_CUARENTENA.value,
              "'a' queda en cuarentena")
    comprobar(estado["a"]["espera_s"] == 60.0,
              "la cuarentena por defecto es de 60s",
              "espera={}".format(estado["a"]["espera_s"]))
    comprobar(any("sin cuota" in i for i in r.incidencias),
              "la incidencia queda registrada", str(r.incidencias))


async def caso_no_reintenta_en_cuarentena():
    print("\n[3] Un modelo en cuarentena ya no se vuelve a probar")
    router, reloj = construir(["a", "b"])
    operacion, llamadas = guion({"a": [ErrorFalso(429, "RESOURCE_EXHAUSTED")]})
    await router.ejecutar(operacion, clasificar)
    llamadas.clear()
    await router.ejecutar(operacion, clasificar)
    comprobar(llamadas == [("b", 0)], "la segunda llamada va directa a 'b'",
              str(llamadas))

    reloj.avanzar(61)
    llamadas.clear()
    await router.ejecutar(operacion, clasificar)
    comprobar(llamadas == [("a", 0)], "cumplida la espera, 'a' vuelve a la cabeza",
              str(llamadas))


async def caso_404_se_descarta_para_siempre():
    print("\n[4] 404: el modelo sale de la rotacion definitivamente")
    router, reloj = construir(["a", "b"])
    operacion, llamadas = guion({"a": ErrorFalso(404, "NOT_FOUND")})
    await router.ejecutar(operacion, clasificar)
    reloj.avanzar(100_000)
    llamadas.clear()
    await router.ejecutar(operacion, clasificar)
    comprobar(llamadas == [("b", 0)],
              "ni pasado el tiempo se reintenta un modelo inexistente", str(llamadas))
    estado = {c["modelo"]: c for c in router.instantanea()}
    comprobar(estado["a"]["estado"] == EstadoCandidato.DESCARTADO.value,
              "'a' figura como descartado")


async def caso_transitorio_reintenta_mismo_modelo():
    print("\n[5] 503: reintenta el mismo modelo antes de bajar")
    router, _ = construir(["a", "b"], reintentos_transitorios=2)
    operacion, llamadas = guion(
        {"a": [ErrorFalso(503, "UNAVAILABLE"), "ok:a-recuperado"]}
    )
    r = await router.ejecutar(operacion, clasificar)
    comprobar(r.valor == "ok:a-recuperado", "el reintento sobre 'a' tiene exito")
    comprobar(llamadas == [("a", 0), ("a", 0)], "reintenta 'a', no salta a 'b'",
              str(llamadas))
    comprobar(not r.degradado, "no cuenta como degradacion")


async def caso_transitorio_agota_reintentos():
    print("\n[6] 503 persistente: agota reintentos y baja de modelo")
    router, _ = construir(["a", "b"], reintentos_transitorios=2)
    operacion, llamadas = guion({"a": ErrorFalso(503, "UNAVAILABLE")})
    r = await router.ejecutar(operacion, clasificar)
    comprobar(r.modelo == "b", "acaba respondiendo 'b'")
    comprobar(llamadas.count(("a", 0)) == 3,
              "probo 'a' 3 veces (1 + 2 reintentos)",
              "intentos_a={}".format(llamadas.count(("a", 0))))


async def caso_fatal_se_propaga():
    print("\n[7] 400: error nuestro, se propaga sin tocar la cadena")
    router, _ = construir(["a", "b"])
    operacion, llamadas = guion({"a": ErrorFalso(400, "INVALID_ARGUMENT")})
    try:
        await router.ejecutar(operacion, clasificar)
        comprobar(False, "deberia haber lanzado la excepcion original")
    except ErrorFalso as e:
        comprobar(e.code == 400, "propaga el error original sin envolver")
        comprobar(llamadas == [("a", 0)], "no prueba los modelos de respaldo",
                  str(llamadas))


async def caso_cadena_agotada():
    print("\n[8] Cadena entera sin cuota")
    router, _ = construir(["a", "b"])
    operacion, _ = guion({
        "a": ErrorFalso(429, "RESOURCE_EXHAUSTED"),
        "b": ErrorFalso(429, "RESOURCE_EXHAUSTED"),
    })
    try:
        await router.ejecutar(operacion, clasificar, descripcion="extraccion")
        comprobar(False, "deberia haber lanzado SinModelosDisponibles")
    except SinModelosDisponibles as e:
        comprobar("extraccion" in str(e), "el mensaje identifica la operacion")
        comprobar(e.segundos_para_reintentar == 60.0,
                  "indica cuantos segundos esperar (para el Retry-After)",
                  "espera={}".format(e.segundos_para_reintentar))


async def caso_retry_delay_de_la_api():
    print("\n[9] Se respeta el retryDelay que manda la API")
    router, _ = construir(["a", "b"])
    error = ErrorFalso(429, "RESOURCE_EXHAUSTED", details={
        "error": {"details": [
            {"@type": "type.googleapis.com/google.rpc.QuotaFailure"},
            {"@type": "type.googleapis.com/google.rpc.RetryInfo", "retryDelay": "17s"},
        ]}
    })
    operacion, _ = guion({"a": error})
    await router.ejecutar(operacion, clasificar)
    estado = {c["modelo"]: c for c in router.instantanea()}
    comprobar(estado["a"]["espera_s"] == 17.0,
              "usa los 17s de la API en vez de los 60s por defecto",
              "espera={}".format(estado["a"]["espera_s"]))


async def caso_cuarentena_escala():
    print("\n[10] La cuarentena escala x4 en cada reincidencia")
    router, reloj = construir(["a", "b"], cooldown_base=10, cooldown_maximo=200)
    operacion, _ = guion({"a": ErrorFalso(429, "RESOURCE_EXHAUSTED")})

    esperas = []
    for _ in range(4):
        await router.ejecutar(operacion, clasificar)
        estado = {c["modelo"]: c for c in router.instantanea()}
        esperas.append(estado["a"]["espera_s"])
        reloj.avanzar(estado["a"]["espera_s"] + 1)

    comprobar(esperas == [10.0, 40.0, 160.0, 200.0],
              "10 -> 40 -> 160 -> tope 200", str(esperas))


async def caso_exito_perdona_historial():
    print("\n[11] Un exito reinicia el castigo acumulado")
    router, reloj = construir(["a", "b"], cooldown_base=10)
    operacion, _ = guion({"a": [
        ErrorFalso(429, "RESOURCE_EXHAUSTED"),
        "ok:a",
        ErrorFalso(429, "RESOURCE_EXHAUSTED"),
    ]})
    await router.ejecutar(operacion, clasificar)
    reloj.avanzar(11)
    await router.ejecutar(operacion, clasificar)     # exito: perdona
    await router.ejecutar(operacion, clasificar)     # vuelve a fallar
    estado = {c["modelo"]: c for c in router.instantanea()}
    comprobar(estado["a"]["espera_s"] == 10.0,
              "vuelve a la cuarentena base, no a la escalada",
              "espera={}".format(estado["a"]["espera_s"]))


async def caso_varias_claves():
    print("\n[12] Con varias claves se agota el mejor modelo antes de bajar")
    router, _ = construir(["a", "b"], n_claves=2)
    comprobar(len(router.candidatos) == 4, "4 candidatos = 2 modelos x 2 claves")

    llamadas = []

    async def operacion(modelo, clave):
        llamadas.append((modelo, clave))
        if modelo == "a":
            raise ErrorFalso(429, "RESOURCE_EXHAUSTED")
        return "ok"

    r = await router.ejecutar(operacion, clasificar)
    comprobar(llamadas == [("a", 0), ("a", 1), ("b", 0)],
              "prueba 'a' con las dos claves antes de pasar a 'b'", str(llamadas))
    comprobar(r.modelo == "b" and r.indice_clave == 0,
              "informa modelo y clave que respondieron")


async def caso_concurrencia():
    print("\n[13] Bajo concurrencia el estado no se corrompe")
    router, _ = construir(["a", "b"])

    async def operacion(modelo, _clave):
        await asyncio.sleep(0)
        if modelo == "a":
            raise ErrorFalso(429, "RESOURCE_EXHAUSTED")
        return "ok"

    resultados = await asyncio.gather(
        *(router.ejecutar(operacion, clasificar) for _ in range(25))
    )
    comprobar(all(r.modelo == "b" for r in resultados),
              "las 25 peticiones simultaneas acaban en 'b'")
    estado = {c["modelo"]: c for c in router.instantanea()}
    comprobar(estado["b"]["exitos"] == 25, "se contabilizan los 25 exitos",
              "exitos={}".format(estado["b"]["exitos"]))
    comprobar(estado["a"]["estado"] == EstadoCandidato.EN_CUARENTENA.value,
              "'a' quedo en cuarentena una sola vez")


async def main():
    print("=" * 70)
    print("PRUEBA DEL ROUTER DE MODELOS")
    print("=" * 70)

    for caso in (
        caso_camino_feliz, caso_cuota_baja_al_siguiente,
        caso_no_reintenta_en_cuarentena, caso_404_se_descarta_para_siempre,
        caso_transitorio_reintenta_mismo_modelo, caso_transitorio_agota_reintentos,
        caso_fatal_se_propaga, caso_cadena_agotada, caso_retry_delay_de_la_api,
        caso_cuarentena_escala, caso_exito_perdona_historial,
        caso_varias_claves, caso_concurrencia,
    ):
        await caso()

    print("\n" + "=" * 70)
    if fallos:
        print("{}/{} comprobaciones FALLARON:".format(len(fallos), comprobaciones))
        for f in fallos:
            print("  - {}".format(f))
        return 1
    print("{}/{} comprobaciones OK".format(comprobaciones, comprobaciones))
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
