"""Diagnostico de la cuenta de Gemini y de la cadena de respaldo.

Google agrega y retira modelos con frecuencia (gemini-2.5-flash-lite dejo de
servirse a claves nuevas, y los Pro exigen plan de pago). Este script comprueba
contra la API real que la cadena de GEMINI_MODELOS sigue siendo valida.

    python scripts/probar_ia.py             # lista modelos y revisa la cadena
    python scripts/probar_ia.py --cadena    # ademas prueba cada modelo en vivo
"""

import argparse
import asyncio
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from google import genai  # noqa: E402
from google.genai import errors, types  # noqa: E402
from pydantic import BaseModel  # noqa: E402

from config.settings import ajustes  # noqa: E402
from services.ai_service import clasificar_error_gemini  # noqa: E402


class Sonda(BaseModel):
    """Esquema minimo: valida de paso que el modelo admite salida estructurada."""

    comuna: str
    superficie_ha: float


SONDA = (
    "Extrae los datos: 'El predio se ubica en la comuna de Requinoa y "
    "tiene cuarenta y siete coma cinco hectareas.'"
)


async def probar_modelo(cliente, modelo):
    inicio = time.monotonic()
    try:
        respuesta = await asyncio.wait_for(
            cliente.aio.models.generate_content(
                model=modelo,
                contents=SONDA,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=Sonda,
                    temperature=0.0,
                ),
            ),
            timeout=90,
        )
    except errors.APIError as error:
        return False, "{} {} -> {}".format(
            error.code, error.status, clasificar_error_gemini(error).value
        )
    except asyncio.TimeoutError:
        return False, "timeout (>90s)"
    except Exception as error:  # noqa: BLE001
        return False, "{}: {}".format(type(error).__name__, str(error)[:70])

    ms = int((time.monotonic() - inicio) * 1000)
    uso = respuesta.usage_metadata
    correcto = (
        respuesta.parsed is not None
        and abs(getattr(respuesta.parsed, "superficie_ha", 0) - 47.5) < 0.01
    )
    return True, "{:>6}ms  in={} out={}  dato correcto: {}".format(
        ms, uso.prompt_token_count, uso.candidates_token_count,
        "si" if correcto else "NO",
    )


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--cadena", action="store_true",
                        help="prueba en vivo cada modelo de la cadena")
    args = parser.parse_args()

    if not ajustes.ia_configurada:
        print("No hay GEMINI_API_KEYS en .env")
        return 1

    claves = ajustes.claves_gemini
    cadena = ajustes.modelos_gemini
    print("claves configuradas : {}".format(len(claves)))
    print("cadena configurada  : {}".format(" -> ".join(cadena)))
    print()

    cliente = genai.Client(api_key=claves[0])
    try:
        catalogo = await asyncio.to_thread(lambda: list(cliente.models.list()))
    except errors.APIError as error:
        print("La clave no sirve para listar modelos: {} {}".format(
            error.code, error.status))
        return 1

    disponibles = sorted(
        m.name.removeprefix("models/")
        for m in catalogo
        if "generateContent" in (getattr(m, "supported_actions", None) or [])
    )
    print("modelos de texto en la cuenta: {}".format(len(disponibles)))

    ausentes = [m for m in cadena if m not in disponibles]
    if ausentes:
        print("\nAVISO: estos modelos de la cadena NO estan en la cuenta:")
        for m in ausentes:
            print("   - {}".format(m))
        print("   Corrige GEMINI_MODELOS en .env")
    else:
        print("todos los modelos de la cadena existen en la cuenta")

    if not args.cadena:
        print("\n(usa --cadena para probarlos en vivo)")
        return 0

    print("\nprueba en vivo (consume cuota):")
    vivos = 0
    for modelo in cadena:
        ok, detalle = await probar_modelo(cliente, modelo)
        print("  {} {:26} {}".format("OK  " if ok else "FALLA", modelo, detalle))
        vivos += ok

    print("\n{}/{} modelos de la cadena responden.".format(vivos, len(cadena)))
    if vivos == 0:
        print("Ningun modelo disponible: la API respondera 503 en /ia/*")
        return 1
    if vivos < 2:
        print("Con un solo modelo vivo no hay respaldo real. Agrega mas a "
              "GEMINI_MODELOS o anade otra clave a GEMINI_API_KEYS.")
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
