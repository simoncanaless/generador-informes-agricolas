"""Conversion de PDFs legados a texto y extraccion de sus antecedentes.

Dos capas bien separadas:

  1. **Determinista** (``convertir_a_markdown``): PyMuPDF pasa el PDF a Markdown
     plano. Aqui no interviene la IA. Es donde se gana el ahorro de tokens: un
     PDF de escritura pesa megabytes, su texto util son unos pocos miles de
     caracteres, y mandar el binario a un modelo multimodal cuesta un orden de
     magnitud mas que mandar ese texto.

  2. **Probabilistica** (``ParserPDF.extraer``): el Markdown ya reducido se pasa
     al proveedor de IA para que devuelva ``ExtraccionLegal`` validado.

Los PDF escaneados no tienen capa de texto, asi que el paso 1 no produce nada
util. En ese caso se rasterizan las paginas y se resuelven por vision, que es
mas caro pero es la unica via. La deteccion es automatica.
"""

from __future__ import annotations

import asyncio
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Sequence, Union

import pymupdf
import pymupdf4llm

from schemas.extraccion import ExtraccionLegal
from services.ai_service import ProveedorIA
from services.model_router import Respuesta

__all__ = [
    "ResultadoConversion", "ParserPDF", "convertir_a_markdown",
    "PdfIlegible", "PdfProtegido",
]

#: Si una pagina tiene menos caracteres que esto, se considera sin capa de texto.
#: Las paginas escaneadas suelen traer igualmente el sello de validacion del
#: Conservador ("Certificado N... Verifique validez... Pag 3/9"), que ronda los
#: 70-80 caracteres: por eso el umbral no puede ser mas bajo.
MINIMO_CARACTERES_POR_PAGINA = 110

#: Una pagina se da por escaneada cuando tiene poco texto Y ademas lleva
#: imagenes encima. El matiz importa: una pagina corta (una firma, un anexo de
#: dos lineas) tambien tiene poco texto, pero no hay nada que rescatar de ella
#: por vision y mandarla como imagen solo encarece la llamada.

#: Tope de paginas que se rasterizan. Cada pagina es una imagen cara de procesar.
MAXIMO_PAGINAS_RASTERIZADAS = 12

#: Resolucion de rasterizado. 150 ppp bastaba para leer un texto mecanografiado,
#: pero se paga por MOSAICOS de 768 px, no por pixeles: a 150 ppp una hoja
#: carta sale 1.252x2.067 y ocupa 6 mosaicos (1.548 tokens); a 110 sale
#: 919x1.516 y ocupa 4 (1.032). **Un tercio menos por pagina, y son 23 paginas
#: por informe.** De 150 a 120 no cambia nada -siguen 6 mosaicos-, asi que
#: bajar a medias no sirve: hay que cruzar el escalon.
#:
#: Comprobado contra las 3 escrituras reales del CBR antes de bajarlo: a 110
#: ppp la extraccion coincide en **27/27 campos** con la de 150. Se probo
#: tambien a 90 -que ahorraria un 67%, porque el ancho cae bajo 768 y quedan 2
#: mosaicos- y en una corrida difirio en el nombre del predio; repetido dos
#: veces mas dio igual que 150, o sea que **aquello era variacion del modelo
#: entre corridas, no del DPI**. Aun asi se deja en 110: la muestra para
#: certificar 90 es corta, y §12 ya tiene el caso de un modelo leyendo `207-7`
#: donde decia `203-7`. Un rol mal leido no es un ahorro.
DPI_RASTERIZADO = 110

#: Aproximacion habitual para espanol: ~4 caracteres por token.
CARACTERES_POR_TOKEN = 4


class PdfProtegido(RuntimeError):
    """El PDF pide contrasena."""


class PdfIlegible(RuntimeError):
    """El archivo no es un PDF valido o esta corrupto."""


@dataclass
class ResultadoConversion:
    """Lo que produjo la conversion, con su diagnostico."""

    markdown: str
    paginas: int
    paginas_convertidas: int
    tiene_capa_texto: bool
    caracteres_originales: int = 0
    bytes_originales: int = 0
    advertencias: List[str] = field(default_factory=list)
    #: Indices 0-based de las paginas sin capa de texto utilizable.
    paginas_sin_texto: List[int] = field(default_factory=list)

    @property
    def es_hibrido(self) -> bool:
        """Parte con texto real y parte escaneada.

        Es el formato habitual de los certificados electronicos del CBR: la
        caratula la firma el Conservador (texto de verdad) y la inscripcion en
        si viene fotografiada.
        """
        return bool(self.paginas_sin_texto) and self.caracteres > 0

    @property
    def caracteres(self) -> int:
        return len(self.markdown)

    @property
    def tokens_estimados(self) -> int:
        return self.caracteres // CARACTERES_POR_TOKEN

    @property
    def reduccion_pct(self) -> float:
        """Cuanto se encogio respecto al archivo original."""
        if not self.bytes_originales:
            return 0.0
        return round(100 * (1 - self.caracteres / self.bytes_originales), 1)

    def resumen(self) -> str:
        return (
            "{}/{} paginas, {} caracteres (~{} tokens), {}% menos que el binario, "
            "capa de texto: {}".format(
                self.paginas_convertidas, self.paginas, self.caracteres,
                self.tokens_estimados, self.reduccion_pct,
                "si" if self.tiene_capa_texto else "NO (escaneado)",
            )
        )


# ---------------------------------------------------------- determinista

def _abrir(origen: Union[str, Path, bytes]) -> "pymupdf.Document":
    try:
        if isinstance(origen, (str, Path)):
            documento = pymupdf.open(origen)
        else:
            documento = pymupdf.open(stream=origen, filetype="pdf")
    except Exception as error:  # pymupdf lanza tipos variados
        raise PdfIlegible("no se pudo abrir el PDF: {}".format(error)) from error

    if documento.needs_pass:
        documento.close()
        raise PdfProtegido("el PDF esta protegido con contrasena")
    return documento


def limpiar_markdown(texto: str) -> str:
    """Quita el ruido que no aporta significado pero si cuesta tokens.

    Encabezados y pies repetidos en todas las paginas, lineas de guiones de
    formulario, saltos triples y espacios de relleno.
    """
    texto = re.sub(r"[ \t]+\n", "\n", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    # lineas compuestas solo de separadores decorativos
    texto = re.sub(r"^[\s_\-=\.\*]{6,}$", "", texto, flags=re.MULTILINE)
    # numeros de pagina sueltos en su propia linea
    texto = re.sub(r"^\s*(?:pag(?:ina)?\.?\s*)?\d{1,3}\s*$", "", texto,
                   flags=re.MULTILINE | re.IGNORECASE)
    return re.sub(r"\n{3,}", "\n\n", texto).strip()


def convertir_a_markdown(
    origen: Union[str, Path, bytes],
    paginas: Optional[Sequence[int]] = None,
    limpiar: bool = True,
) -> ResultadoConversion:
    """Convierte un PDF a Markdown sin intervencion de IA.

    ``paginas`` son indices base 0; ``None`` convierte el documento entero.
    """
    documento = _abrir(origen)
    advertencias: List[str] = []

    try:
        total = documento.page_count
        if total == 0:
            raise PdfIlegible("el PDF no tiene paginas")

        indices = list(paginas) if paginas is not None else list(range(total))
        indices = [i for i in indices if 0 <= i < total]
        if not indices:
            raise PdfIlegible("el rango de paginas solicitado no existe")

        # Diagnostico PAGINA A PAGINA. Sumar los caracteres de todo el documento
        # engana: una sola caratula firmada electronicamente basta para superar
        # cualquier umbral global mientras el cuerpo real esta escaneado.
        caracteres_crudos = 0
        sin_texto: List[int] = []
        cortas: List[int] = []
        for i in indices:
            pagina = documento[i]
            largo = len(pagina.get_text("text").strip())
            caracteres_crudos += largo
            if largo >= MINIMO_CARACTERES_POR_PAGINA:
                continue
            # Poco texto: la pregunta es si hay algo debajo que leer.
            if pagina.get_images(full=True):
                sin_texto.append(i)
            else:
                cortas.append(i)

        # Basta con que UNA pagina venga escaneada para no fiarse del texto
        # suelto: es justo el caso del certificado del CBR, donde la caratula
        # digital trae las fojas pero el cuerpo de la inscripcion es una foto.
        tiene_texto = not sin_texto
        if sin_texto:
            advertencias.append(
                "{} de {} paginas escaneadas sin capa de texto (pag. {})".format(
                    len(sin_texto), len(indices),
                    ", ".join(str(i + 1) for i in sin_texto[:10]),
                )
            )
        if cortas:
            advertencias.append(
                "{} pagina(s) con muy poco texto y sin imagenes; se toman tal "
                "cual".format(len(cortas))
            )

        # Se convierten todas menos las escaneadas: pedirle Markdown a una
        # pagina escaneada devuelve el sello de validacion y nada mas, que es
        # ruido puro y encima cuesta tokens.
        escaneadas = set(sin_texto)
        con_texto_indices = [i for i in indices if i not in escaneadas]
        markdown = ""
        if con_texto_indices:
            markdown = pymupdf4llm.to_markdown(
                documento,
                pages=con_texto_indices,
                ignore_images=True,     # las fotos del PDF no aportan al texto legal
                ignore_graphics=True,   # sellos y marcos decorativos
                table_strategy="lines_strict",
                show_progress=False,
            )
            if limpiar:
                markdown = limpiar_markdown(markdown)
        if not tiene_texto:
            advertencias.append(
                "el cuerpo del documento esta escaneado: se resuelve por vision"
            )

        tamano = (
            len(origen) if isinstance(origen, bytes)
            else Path(origen).stat().st_size
        )

        return ResultadoConversion(
            markdown=markdown,
            paginas=total,
            paginas_convertidas=len(indices),
            tiene_capa_texto=tiene_texto,
            caracteres_originales=caracteres_crudos,
            bytes_originales=tamano,
            advertencias=advertencias,
            paginas_sin_texto=sin_texto,
        )
    finally:
        documento.close()


def rasterizar_paginas(
    origen: Union[str, Path, bytes],
    maximo: int = MAXIMO_PAGINAS_RASTERIZADAS,
    dpi: int = DPI_RASTERIZADO,
    paginas: Optional[Sequence[int]] = None,
) -> List[bytes]:
    """Convierte paginas en PNG para poder leerlas por vision.

    ``paginas`` (indices 0-based) permite rasterizar solo las que hagan falta:
    en un certificado del CBR la caratula ya se leyo como texto y mandarla otra
    vez como imagen solo encarece la llamada.
    """
    documento = _abrir(origen)
    try:
        elegidas = (
            [i for i in paginas if 0 <= i < documento.page_count]
            if paginas is not None
            else list(range(documento.page_count))
        )
        return [
            documento[i].get_pixmap(dpi=dpi).tobytes("png")
            for i in elegidas[:maximo]
        ]
    finally:
        documento.close()


# -------------------------------------------------------- con IA detras

class ParserPDF:
    """Une la conversion determinista con la extraccion por IA."""

    def __init__(self, ia: ProveedorIA):
        self.ia = ia

    async def extraer(
        self,
        origen: Union[str, Path, bytes],
        paginas: Optional[Sequence[int]] = None,
        instrucciones: Optional[str] = None,
    ) -> "ResultadoExtraccion":
        """Devuelve los antecedentes legales del PDF.

        Elige la via automaticamente: texto si el PDF tiene capa de texto,
        vision si esta escaneado.
        """
        conversion = await asyncio.to_thread(convertir_a_markdown, origen, paginas)

        if conversion.tiene_capa_texto and conversion.markdown:
            respuesta = await self.ia.extraer_de_documento(
                conversion.markdown, instrucciones
            )
            via = "texto"
        else:
            # Solo se fotografian las paginas que no tienen texto. Si el PDF trae
            # caratula digital (el caso normal en los certificados del CBR), su
            # texto se manda igualmente como contexto: ahi vienen las fojas, el
            # numero y el anio, y leerlos del texto es mas fiable que de la
            # imagen.
            imagenes = await asyncio.to_thread(
                rasterizar_paginas, origen, MAXIMO_PAGINAS_RASTERIZADAS,
                DPI_RASTERIZADO, conversion.paginas_sin_texto or None,
            )
            if not imagenes:
                raise PdfIlegible("no se pudo obtener ni texto ni imagen del PDF")
            respuesta = await self._extraer_por_vision(
                imagenes, instrucciones,
                conversion.markdown if conversion.es_hibrido else None,
            )
            via = "hibrido" if conversion.es_hibrido else "vision"
            conversion.advertencias.append(
                "resuelto por {} sobre {} pagina(s) escaneada(s): revisar con "
                "especial atencion los numeros de fojas e inscripcion".format(
                    via, len(imagenes)
                )
            )

        datos = respuesta.valor
        if via != "texto":
            _degradar_confianza(datos, conversion)
            if respuesta.degradado:
                # Escaneado + modelo de respaldo es la peor combinacion posible.
                # Medido sobre un certificado real del CBR: el modelo de reserva
                # devolvio el rol '207-7' donde decia '203-7'. Un digito de mas
                # en un rol de avaluo invalida el informe entero.
                conversion.advertencias.insert(0, (
                    "ATENCION: documento escaneado leido por un modelo de "
                    "respaldo ({}) porque el principal no tenia cuota. En esta "
                    "combinacion los numeros fallan mas de lo habitual. "
                    "Coteje TODOS los datos numericos, o repita la extraccion "
                    "mas tarde cuando se libere la cuota."
                ).format(respuesta.modelo))

        return ResultadoExtraccion(
            datos=datos,
            conversion=conversion,
            via=via,
            modelo=respuesta.modelo,
            degradado=respuesta.degradado,
            incidencias=respuesta.incidencias,
            duracion_s=round(respuesta.duracion_s, 2),
        )

    async def _extraer_por_vision(
        self,
        imagenes: List[bytes],
        instrucciones: Optional[str],
        texto_disponible: Optional[str] = None,
    ) -> Respuesta[ExtraccionLegal]:
        """Ruta para escaneados: se apoya en el mismo nucleo del proveedor."""
        from google.genai import types  # local: solo esta rama necesita el SDK

        partes: List[object] = [
            "Este documento legal esta escaneado y no tiene capa de texto. "
            "Leelo pagina a pagina y extrae los antecedentes. Si un numero no "
            "se lee con seguridad, marca confianza baja y dilo en advertencias."
        ]
        if texto_disponible:
            partes.append(
                "La caratula del documento SI trae texto digital y se transcribe "
                "abajo. Prefiere SIEMPRE este texto para las fojas, el "
                "numero y el anio de inscripcion, porque es exacto; usa las "
                "imagenes para todo lo demas (propietario, predio, rol, "
                "superficie, deslindes, aguas y gravamenes).\n\n"
                "--- CARATULA (texto exacto) ---\n{}".format(texto_disponible)
            )
        partes.extend(
            types.Part.from_bytes(data=imagen, mime_type="image/png")
            for imagen in imagenes
        )
        return await self.ia._generar(  # noqa: SLF001 - misma capa de servicio
            partes, ExtraccionLegal,
            instrucciones or "Trabajas sobre escrituras chilenas escaneadas.",
            "extraccion por vision de {} pagina(s)".format(len(imagenes)),
        )


def _degradar_confianza(
    datos: ExtraccionLegal, conversion: ResultadoConversion
) -> None:
    """Baja la confianza de lo que se leyo de una imagen, diga lo que diga el modelo.

    Medido sobre un certificado real del CBR: en 4 lecturas del mismo documento
    el rol, las fojas y el numero salieron identicos y correctos las 4 veces,
    pero la superficie salio bien 1 de 4 ('78,5' en vez de '78,55')... y el
    modelo declaro confianza ALTA en las cuatro.

    La conclusion es que la autoevaluacion del modelo no sirve como garantia.
    La confianza tiene que reflejar COMO se obtuvo el dato: un caracter leido de
    la capa de texto es exacto; uno leido de un escaneo puede no serlo. Por eso
    se degrada aqui, de forma determinista, y no se le pregunta al modelo.
    """
    from schemas.extraccion import Confianza

    datos.propiedad.procedencia.confianza = Confianza.BAJA
    for derecho in datos.derechos_agua:
        derecho.procedencia.confianza = Confianza.BAJA

    numericos = []
    if datos.propiedad.superficie_escritura_ha is not None:
        numericos.append("superficie ({} ha)".format(
            datos.propiedad.superficie_escritura_ha))
    for derecho in datos.derechos_agua:
        if derecho.caudal_litros_segundo is not None:
            numericos.append("caudal ({} l/s)".format(derecho.caudal_litros_segundo))
        if derecho.acciones is not None:
            numericos.append("acciones ({})".format(derecho.acciones))

    if numericos:
        conversion.advertencias.append(
            "COTEJAR CONTRA EL DOCUMENTO: {} se leyeron de una pagina escaneada. "
            "Los decimales son la parte que mas se equivoca.".format(
                "; ".join(numericos)
            )
        )


@dataclass
class ResultadoExtraccion:
    """Datos extraidos mas la trazabilidad de como se obtuvieron."""

    datos: ExtraccionLegal
    conversion: ResultadoConversion
    via: str
    modelo: str
    degradado: bool
    incidencias: List[str]
    duracion_s: float

    def resumen(self) -> str:
        return "via {} | modelo {} | {}s | {}".format(
            self.via, self.modelo, self.duracion_s, self.conversion.resumen()
        )
