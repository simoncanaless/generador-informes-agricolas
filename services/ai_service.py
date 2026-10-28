"""Adaptador asincrono de IA. Todo lo probabilistico del sistema vive aqui.

Responsabilidad unica: convertir documentos, imagenes y audio en datos
estructurados y validados por Pydantic. **No calcula nada.** Los totales, las
depreciaciones y las densidades los produce el motor determinista a partir de
lo que este modulo extraiga.

La interfaz ``ProveedorIA`` existe para que el resto del backend dependa de una
abstraccion y no de Google (DIP): cambiar a otro proveedor es escribir otra
implementacion, sin tocar la API ni el motor Excel.

La resiliencia ante cuotas agotadas la aporta ``RouterModelos``: este modulo
solo aporta el clasificador de errores propio de Gemini.
"""

from __future__ import annotations

import asyncio
import json
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional, Sequence, Type, TypeVar

from google import genai
from google.genai import errors, types
from pydantic import BaseModel

from config.settings import Ajustes, ajustes as ajustes_por_defecto
from schemas.extraccion import (
    ExtraccionLegal, LoteFotografias, MapeoColumnas, NotaTerreno,
    RedaccionTecnica,
)
from services.model_router import (
    Desenlace, Respuesta, RouterModelos, SinModelosDisponibles,
)

M = TypeVar("M", bound=BaseModel)

__all__ = [
    "ProveedorIA", "ServicioGemini", "IANoConfigurada", "RespuestaVacia",
    "SinModelosDisponibles", "construir_servicio_ia",
]


#: Tono comun a todas las llamadas. Fija el registro del informe bancario y,
#: sobre todo, prohibe inventar y prohibe calcular.
INSTRUCCION_BASE = (
    "Eres un asistente de un equipo de tasacion agricola en Chile que prepara "
    "informes para la banca. Trabajas en espanol tecnico, sobrio y verificable.\n"
    "Reglas inviolables:\n"
    "1. NUNCA inventes un dato. Si no aparece en el material entregado, "
    "devuelve null y explicalo en 'advertencias'.\n"
    "2. NUNCA calcules ni agregues cifras (sumas, totales, promedios, "
    "conversiones de unidades, superficies derivadas). Transcribe solo lo que "
    "el documento declara de forma explicita.\n"
    "3. Respeta la ortografia original de nombres propios y roles de avaluo.\n"
    "4. Los numerales SI se normalizan a cifras, porque eso es transcribir y no "
    "calcular: 'fojas mil doscientos cuarenta' se devuelve como '1240' y "
    "'cuarenta y siete coma cinco hectareas' como 47.5. La coma decimal chilena "
    "('47,5') equivale al punto decimal (47.5). Nunca alteres el valor.\n"
    "5. Marca 'confianza' con honestidad y cita en 'fuente' el fragmento que "
    "respalda cada bloque.\n"
    "6. No uses lenguaje comercial ni valorativo: describes, no vendes."
)


#: Lado mayor al que se reduce una fotografia antes de mandarla al modelo.
#: Las camaras de terreno y los drones entregan 4-6 MB por foto; una visita de
#: 14 fotos son ~50 MB y la peticion no cabe (el limite de contenido en linea
#: ronda los 20 MB). A 1600 px se describe exactamente igual de bien: el modelo
#: no necesita resolucion de impresion para decir que la techumbre es de zinc.
LADO_MAXIMO_IMAGEN = 1600
CALIDAD_JPEG = 85


def reducir_imagen(datos: bytes, lado_maximo: int = LADO_MAXIMO_IMAGEN) -> bytes:
    """Reescala una foto conservando la proporcion. Devuelve JPEG.

    Si algo falla (formato raro, imagen corrupta) se devuelve el original: es
    preferible una peticion pesada a perder la fotografia.
    """
    try:
        from io import BytesIO

        from PIL import Image

        with Image.open(BytesIO(datos)) as imagen:
            if max(imagen.size) <= lado_maximo and len(datos) < 2_000_000:
                return datos
            imagen = imagen.convert("RGB")
            imagen.thumbnail((lado_maximo, lado_maximo), Image.LANCZOS)
            salida = BytesIO()
            imagen.save(salida, format="JPEG", quality=CALIDAD_JPEG, optimize=True)
            return salida.getvalue()
    except Exception:  # noqa: BLE001 - nunca por una miniatura se pierde una foto
        return datos


class IANoConfigurada(RuntimeError):
    """No hay clave de API: el 80% determinista sigue funcionando, el 20% no."""


class RespuestaVacia(RuntimeError):
    """El modelo no devolvio contenido utilizable (bloqueo de seguridad, corte)."""


# ------------------------------------------------------------- interfaz

class ProveedorIA(ABC):
    """Contrato que el backend consume. Ninguna firma menciona a Google."""

    @abstractmethod
    async def extraer_de_documento(
        self, texto: str, instrucciones: Optional[str] = None
    ) -> Respuesta[ExtraccionLegal]:
        """Lee un documento legal ya convertido a texto/Markdown."""

    @abstractmethod
    async def titular_fotografias(
        self, imagenes: Sequence[bytes], mime: str = "image/jpeg",
        contexto: Optional[str] = None,
    ) -> Respuesta[LoteFotografias]:
        """Propone un titulo corto y la categoria de cada imagen del predio.

        Sin descripcion tecnica (ver schemas/extraccion.py:TituloFotografia).
        """

    @abstractmethod
    async def estructurar_nota_voz(
        self, audio: bytes, mime: str = "audio/ogg",
        contexto: Optional[str] = None,
    ) -> Respuesta[NotaTerreno]:
        """Transcribe y estructura una nota de voz tomada en terreno."""

    @abstractmethod
    async def redactar_opinion(
        self, hechos: str, apartado: str
    ) -> Respuesta[RedaccionTecnica]:
        """Propone el borrador de un apartado de opinion profesional."""

    @abstractmethod
    async def mapear_columnas(
        self,
        encabezados: Dict[str, str],
        muestras: List[Dict[str, str]],
        campos: Dict[str, str],
    ) -> Respuesta[MapeoColumnas]:
        """Propone a que campo corresponde cada columna de un catastro ajeno."""

    @abstractmethod
    def estado(self) -> Dict[str, Any]:
        """Instantanea de la cadena de modelos, para el endpoint de salud."""


# --------------------------------------------------------- clasificador

def clasificar_error_gemini(error: BaseException) -> Desenlace:
    """Traduce un error del SDK a la politica del router.

    Es la unica funcion del proyecto que necesita conocer los codigos de Google.
    """
    if isinstance(error, (asyncio.TimeoutError, TimeoutError)):
        return Desenlace.TRANSITORIO

    # Bloqueo de seguridad o respuesta truncada: cambiar de modelo puede
    # resolverlo y no consume cuota, asi que se trata como transitorio.
    if isinstance(error, RespuestaVacia):
        return Desenlace.TRANSITORIO

    if isinstance(error, errors.APIError):
        codigo = getattr(error, "code", None)
        estado = (getattr(error, "status", "") or "").upper()

        if codigo == 429 or estado == "RESOURCE_EXHAUSTED":
            return Desenlace.CUOTA
        if codigo == 404 or estado == "NOT_FOUND":
            return Desenlace.DESCARTAR
        if codigo in (500, 502, 503, 504) or estado in (
            "UNAVAILABLE", "INTERNAL", "DEADLINE_EXCEEDED"
        ):
            return Desenlace.TRANSITORIO
        # 400/401/403: peticion o credencial mal formadas. Rotar no arregla nada.
        return Desenlace.FATAL

    # Fallos de red del transporte: merecen reintento.
    if isinstance(error, (ConnectionError, OSError)):
        return Desenlace.TRANSITORIO

    return Desenlace.FATAL


# --------------------------------------------------------- implementacion

class ServicioGemini(ProveedorIA):
    """Implementacion sobre google-genai con rotacion de modelos y claves."""

    def __init__(self, config: Optional[Ajustes] = None):
        self.config = config or ajustes_por_defecto
        claves = self.config.claves_gemini
        if not claves:
            raise IANoConfigurada(
                "define GEMINI_API_KEYS en .env para habilitar las funciones de IA"
            )

        self._clientes = [genai.Client(api_key=clave) for clave in claves]
        self.router = RouterModelos(
            modelos=self.config.modelos_gemini,
            n_claves=len(self._clientes),
            cooldown_base=self.config.gemini_cooldown_base,
            cooldown_maximo=self.config.gemini_cooldown_maximo,
            reintentos_transitorios=self.config.gemini_reintentos_transitorios,
        )

    # -- nucleo ----------------------------------------------------------

    async def _generar(
        self,
        contenido: Any,
        esquema: Type[M],
        instrucciones: str,
        descripcion: str,
        temperatura: float = 0.0,
    ) -> Respuesta[M]:
        """Una llamada estructurada, con toda la cadena de respaldo detras.

        ``temperatura=0`` por defecto: en extraccion documental queremos la
        salida mas reproducible posible, no la mas creativa.
        """
        config = types.GenerateContentConfig(
            system_instruction="{}\n\n{}".format(INSTRUCCION_BASE, instrucciones),
            response_mime_type="application/json",
            response_schema=esquema,
            temperature=temperatura,
        )

        async def operacion(modelo: str, indice_clave: int) -> M:
            cliente = self._clientes[indice_clave]
            respuesta = await asyncio.wait_for(
                cliente.aio.models.generate_content(
                    model=modelo, contents=contenido, config=config
                ),
                timeout=self.config.gemini_timeout_segundos,
            )
            return _interpretar(respuesta, esquema)

        return await self.router.ejecutar(
            operacion, clasificar_error_gemini, descripcion
        )

    # -- operaciones de dominio ------------------------------------------

    async def extraer_de_documento(
        self, texto: str, instrucciones: Optional[str] = None
    ) -> Respuesta[ExtraccionLegal]:
        peticion = (
            "Extrae los antecedentes legales del siguiente documento. "
            "Copia los numeros de fojas, inscripcion y rol exactamente como "
            "aparecen. Si el documento menciona varios derechos de agua, "
            "devuelvelos todos por separado.\n\n"
            "--- DOCUMENTO ---\n{}".format(texto)
        )
        return await self._generar(
            peticion, ExtraccionLegal,
            instrucciones or "Trabajas sobre escrituras y certificados chilenos.",
            "extraccion de documento legal",
        )

    async def titular_fotografias(
        self, imagenes: Sequence[bytes], mime: str = "image/jpeg",
        contexto: Optional[str] = None,
    ) -> Respuesta[LoteFotografias]:
        if not imagenes:
            raise ValueError("no se entrego ninguna imagen")

        # Reducir antes de enviar: sin esto una visita normal no cabe en una
        # peticion, y ademas se paga resolucion que no aporta nada al analisis.
        reducidas = await asyncio.gather(
            *(asyncio.to_thread(reducir_imagen, imagen) for imagen in imagenes)
        )

        # Solo titulo y categoria (el agrónomo, 2026-10-05): la descripcion
        # tecnica -estado, especie, elementos- la ve mejor el agronomo, y un
        # modelo que la escribe igual termina en el informe sin que nadie la
        # haya mirado. Se le pide explicitamente que no la escriba.
        partes: List[Any] = [
            "Ponle a cada fotografia de este predio agricola un titulo corto, "
            "como el rotulo que va debajo de la imagen en un informe (por "
            "ejemplo 'Huerto de cerezos, sector norte' o 'Bodega de insumos'), "
            "y sugiere su categoria, en el orden en que se entregan. Devuelve "
            "exactamente {} titulos. No describas la imagen ni evalues el "
            "estado de lo que se ve: eso lo hace el agronomo.".format(len(reducidas))
        ]
        if contexto:
            partes.append("Contexto aportado por el tasador: {}".format(contexto))
        partes.extend(
            types.Part.from_bytes(data=imagen, mime_type="image/jpeg")
            for imagen in reducidas
        )

        return await self._generar(
            partes, LoteFotografias,
            "Rotulas las fotografias de un informe de tasacion bancaria.",
            "titulos de {} fotografia(s)".format(len(imagenes)),
            temperatura=0.2,
        )

    async def estructurar_nota_voz(
        self, audio: bytes, mime: str = "audio/ogg",
        contexto: Optional[str] = None,
    ) -> Respuesta[NotaTerreno]:
        partes: List[Any] = [
            "Transcribe literalmente esta nota de voz tomada en terreno y "
            "estructura los datos agronomicos que menciona. Si el tasador dice "
            "una cifra dudosa, transcribela igual y marca confianza baja."
        ]
        if contexto:
            partes.append("Contexto: {}".format(contexto))
        partes.append(types.Part.from_bytes(data=audio, mime_type=mime))

        return await self._generar(
            partes, NotaTerreno,
            "Transcribes espanol de Chile con vocabulario agronomico.",
            "transcripcion de nota de voz",
        )

    async def redactar_opinion(
        self, hechos: str, apartado: str
    ) -> Respuesta[RedaccionTecnica]:
        peticion = (
            "Redacta el apartado '{}' de la opinion profesional a partir "
            "EXCLUSIVAMENTE de estos hechos verificados. No agregues ningun "
            "hecho que no figure aqui.\n\n--- HECHOS ---\n{}".format(apartado, hechos)
        )
        return await self._generar(
            peticion, RedaccionTecnica,
            "Escribes 4 a 8 lineas en tercera persona, sin adjetivos comerciales. "
            "El texto es un borrador que firmara un agronomo.",
            "redaccion de '{}'".format(apartado),
            temperatura=0.3,
        )

    async def mapear_columnas(
        self,
        encabezados: Dict[str, str],
        muestras: List[Dict[str, str]],
        campos: Dict[str, str],
    ) -> Respuesta[MapeoColumnas]:
        """Identifica las columnas de un catastro que no estan en el diccionario.

        Se le mandan SOLO los encabezados y dos o tres filas de muestra: unos
        cientos de tokens, aunque el catastro tenga mil hileras. La muestra es
        necesaria porque hay encabezados que no dicen nada por si solos ('C1',
        'Sup.') y se identifican por lo que contienen.
        """
        catalogo = "\n".join(
            "- {}: {}".format(nombre, descripcion) for nombre, descripcion in campos.items()
        )
        tabla = "\n".join(
            "  {}: encabezado={!r}  ejemplos={}".format(
                letra, encabezados.get(letra, ""),
                [m.get(letra) for m in muestras[:3]],
            )
            for letra in sorted(encabezados)
        )
        peticion = (
            "Un agricultor entrego su catastro de plantaciones en su propio "
            "formato. Indica a que campo de nuestro informe corresponde cada "
            "columna.\n\n"
            "CAMPOS DE DESTINO DISPONIBLES:\n{}\n\n"
            "COLUMNAS DEL ARCHIVO (encabezado y ejemplos de contenido):\n{}\n\n"
            "Devuelve una entrada por columna. Si una columna no corresponde a "
            "ningun campo de la lista (totales, observaciones, numeracion, "
            "precios), pon campo=null. No inventes campos que no esten en la "
            "lista y no asignes el mismo campo a dos columnas."
        ).format(catalogo, tabla)

        return await self._generar(
            peticion, MapeoColumnas,
            "Reconoces formatos de planillas agricolas chilenas.",
            "mapeo de {} columnas de catastro".format(len(encabezados)),
        )

    # -- diagnostico -----------------------------------------------------

    def estado(self) -> Dict[str, Any]:
        espera = self.router.proxima_disponibilidad()
        return {
            "proveedor": "gemini",
            "claves": len(self._clientes),
            "cadena": self.config.modelos_gemini,
            "disponibles_ahora": len(self.router.disponibles()),
            "segundos_hasta_reintento": round(espera, 1) if espera else None,
            "candidatos": self.router.instantanea(),
        }

    async def listar_modelos(self) -> List[str]:
        """Modelos que la clave puede usar de verdad, segun la propia API."""
        cliente = self._clientes[0]
        modelos = await asyncio.to_thread(lambda: list(cliente.models.list()))
        return sorted(
            m.name.removeprefix("models/")
            for m in modelos
            if "generateContent" in (getattr(m, "supported_actions", None) or [])
        )


# ------------------------------------------------------------- helpers

def _interpretar(respuesta: Any, esquema: Type[M]) -> M:
    """Convierte la respuesta del SDK en un modelo validado.

    Se prefiere ``.parsed`` (el SDK ya valido contra el esquema). Si no viene,
    se reintenta a mano sobre el texto, porque algunos modelos devuelven el JSON
    envuelto en un bloque de codigo.
    """
    parseada = getattr(respuesta, "parsed", None)
    if isinstance(parseada, esquema):
        return parseada
    if isinstance(parseada, dict):
        return esquema.model_validate(parseada)

    texto = getattr(respuesta, "text", None)
    if not texto or not texto.strip():
        raise RespuestaVacia(
            "el modelo no devolvio contenido ({})".format(_motivo(respuesta))
        )

    limpio = texto.strip()
    if limpio.startswith("```"):
        limpio = limpio.split("\n", 1)[-1].rsplit("```", 1)[0]
    try:
        return esquema.model_validate(json.loads(limpio))
    except (json.JSONDecodeError, ValueError) as error:
        raise RespuestaVacia(
            "respuesta no interpretable como {}: {}".format(esquema.__name__, error)
        ) from error


def _motivo(respuesta: Any) -> str:
    """Extrae por que vino vacia la respuesta, para que el aviso sea util."""
    realimentacion = getattr(respuesta, "prompt_feedback", None)
    if realimentacion is not None and getattr(realimentacion, "block_reason", None):
        return "bloqueo: {}".format(realimentacion.block_reason)
    candidatos = getattr(respuesta, "candidates", None) or []
    if candidatos and getattr(candidatos[0], "finish_reason", None):
        return "finish_reason: {}".format(candidatos[0].finish_reason)
    return "sin motivo declarado"


def construir_servicio_ia(config: Optional[Ajustes] = None) -> Optional[ProveedorIA]:
    """Crea el proveedor, o ``None`` si no hay clave configurada.

    Devolver ``None`` es deliberado: la API debe arrancar y servir todo el 80%
    determinista aunque nadie haya configurado la IA todavia.
    """
    config = config or ajustes_por_defecto
    if not config.ia_configurada:
        return None
    return ServicioGemini(config)
