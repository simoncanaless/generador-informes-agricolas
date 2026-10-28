"""Lo que protege a la aplicacion ANTES de que una peticion llegue a una ruta.

Tres capas, cada una por un riesgo concreto medido el 2026-10-05:

1. **Quien puede llegar** (`RedesPermitidas`). El servidor escucha en
   0.0.0.0:8000 -es lo que deja entrar por Tailscale-, y eso abre la pagina
   tambien a todo equipo conectado al wifi de la casa: visitas, el televisor,
   cualquier aparato que alguien haya enchufado. Para entrar seguia haciendo
   falta la clave, pero nada impedia probar claves desde ahi. Ahora solo
   llegan este computador (127.0.0.1) y los equipos de Tailscale
   (100.64.0.0/10); el resto recibe un 403 que explica por donde entrar. Las
   redes salen de REDES_PERMITIDAS en el .env, por si algun dia hace falta
   abrir otra.

2. **Cabeceras** (`CabecerasDeSeguridad`): que la pagina no se pueda meter en
   un marco de otro sitio (clickjacking), que el navegador no adivine tipos de
   archivo, que no se filtre la direccion al salir a otro sitio, y que nada de
   la API -RUT, informes, fotos- quede guardado en la cache del disco del
   navegador.

3. **Politica de contenido** (`politica_de_contenido`) para la pagina: el
   navegador solo ejecuta el guion que trae `web/index.html` -identificado por
   su huella SHA-256-, nada inyectado. Es la segunda linea contra un dato
   malicioso que se cuele al HTML: aunque un `<img onerror=...>` llegara a la
   pagina, el navegador se niega a ejecutarlo. La primera linea sigue siendo
   escapar todo (`escapar()` en la pagina).

Las tres son ASGI puro y no `BaseHTTPMiddleware`: este ultimo envuelve el
cuerpo de la respuesta, y con las descargas de Excel no hace falta pagar eso.
"""

from __future__ import annotations

import base64
import hashlib
import ipaddress
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

#: Este computador y Tailscale (IPv4 100.64.0.0/10 e IPv6 fd7a:115c:a1e0::/48,
#: los rangos que Tailscale asigna a cada equipo de la red).
REDES_POR_DEFECTO = "127.0.0.0/8,::1/128,100.64.0.0/10,fd7a:115c:a1e0::/48"

#: Starlette pone esto como "cliente" en su TestClient: no es una IP. Se deja
#: pasar porque solo existe dentro de una prueba, nunca en un socket real.
CLIENTE_DE_PRUEBA = "testclient"


def leer_redes(texto: str) -> List[ipaddress._BaseNetwork]:
    """'127.0.0.0/8, 100.64.0.0/10' -> redes. Una entrada mal escrita se
    rechaza al arrancar con un error claro, no se ignora: una red ignorada en
    silencio es un equipo que de repente no puede entrar, sin explicacion."""
    redes = []
    for trozo in (texto or "").split(","):
        trozo = trozo.strip()
        if trozo:
            redes.append(ipaddress.ip_network(trozo, strict=False))
    return redes


def ip_permitida(cliente: Optional[str], redes: Iterable) -> bool:
    if cliente == CLIENTE_DE_PRUEBA:
        return True
    if not cliente:
        return False
    try:
        ip = ipaddress.ip_address(cliente.split("%", 1)[0])
    except ValueError:
        return False
    # Una IPv4 escrita como IPv6 (::ffff:100.64.1.2) es la misma direccion.
    if isinstance(ip, ipaddress.IPv6Address) and ip.ipv4_mapped:
        ip = ip.ipv4_mapped
    return any(ip.version == red.version and ip in red for red in redes)


_PAGINA_BLOQUEADA = """<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Fitogreen</title></head>
<body style="font:16px/1.5 system-ui,sans-serif;background:#0c0f11;color:#e7ecea;
display:grid;place-items:center;min-height:100vh;margin:0;padding:20px">
<div style="max-width:460px">
<h1 style="font-size:20px;margin:0 0 10px">Este equipo no puede entrar</h1>
<p>El programa de Fitogreen solo se abre desde el computador donde está
instalado o por <b>Tailscale</b>. Encienda Tailscale en este equipo y entre por
la dirección de Tailscale del computador de la oficina.</p>
</div></body></html>"""


class RedesPermitidas:
    """Corta en la puerta toda peticion que no venga de una red permitida."""

    def __init__(self, app, redes: Iterable):
        self.app = app
        self.redes = list(redes)

    async def __call__(self, scope, receive, send):
        if scope["type"] in ("http", "websocket"):
            cliente = (scope.get("client") or (None, None))[0]
            if not ip_permitida(cliente, self.redes):
                if scope["type"] == "websocket":
                    await send({"type": "websocket.close", "code": 1008})
                    return
                cuerpo = _PAGINA_BLOQUEADA.encode("utf-8")
                await send({
                    "type": "http.response.start", "status": 403,
                    "headers": [
                        (b"content-type", b"text/html; charset=utf-8"),
                        (b"content-length", str(len(cuerpo)).encode()),
                        (b"cache-control", b"no-store"),
                    ],
                })
                await send({"type": "http.response.body", "body": cuerpo})
                return
        await self.app(scope, receive, send)


#: Para todas las respuestas.
CABECERAS_COMUNES: List[Tuple[bytes, bytes]] = [
    (b"x-content-type-options", b"nosniff"),
    (b"x-frame-options", b"DENY"),
    (b"referrer-policy", b"same-origin"),
    (b"cross-origin-opener-policy", b"same-origin"),
    (b"permissions-policy",
     b"camera=(), microphone=(), geolocation=(), payment=(), usb=()"),
]


class CabecerasDeSeguridad:
    """Agrega las cabeceras de seguridad, y `no-store` a lo que no es la
    pagina: la API devuelve RUT, informes y fotos, y nada de eso tiene por que
    quedar en la cache del navegador de un computador compartido."""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        es_pagina = scope.get("path", "").startswith("/app")

        async def enviar(mensaje):
            if mensaje["type"] == "http.response.start":
                cabeceras = list(mensaje.get("headers", []))
                presentes = {nombre.lower() for nombre, _ in cabeceras}
                for nombre, valor in CABECERAS_COMUNES:
                    if nombre not in presentes:
                        cabeceras.append((nombre, valor))
                if not es_pagina and b"cache-control" not in presentes:
                    cabeceras.append((b"cache-control", b"no-store"))
                mensaje = dict(mensaje, headers=cabeceras)
            await send(mensaje)

        await self.app(scope, receive, enviar)


# ------------------------------------------------- politica de contenido

_RE_GUION = re.compile(r"<script(\s[^>]*)?>([\s\S]*?)</script>", re.IGNORECASE)
_cache_politica: Dict[Path, Tuple[Tuple[int, int], str]] = {}


def huellas_de_guiones(html: str) -> List[str]:
    """La huella SHA-256 de cada guion EN LINEA de la pagina, como la calcula
    el navegador.

    Dos detalles que deciden si la pagina funciona o queda muerta:
    - el navegador hashea el texto despues de normalizar los saltos de linea
      (CRLF y CR pasan a LF, regla del analizador de HTML), asi que se hace
      lo mismo: en Windows Git puede dejar el archivo con CRLF;
    - un guion con `src` no se hashea: lo autoriza 'self'.
    """
    normal = html.replace("\r\n", "\n").replace("\r", "\n")
    huellas = []
    for atributos, cuerpo in _RE_GUION.findall(normal):
        if atributos and re.search(r"\bsrc\s*=", atributos, re.IGNORECASE):
            continue
        digest = hashlib.sha256(cuerpo.encode("utf-8")).digest()
        huellas.append("'sha256-{}'".format(base64.b64encode(digest).decode()))
    return huellas


def politica_de_contenido(ruta_html: Path) -> str:
    """La cabecera Content-Security-Policy para esa pagina.

    Se recalcula solo si el archivo cambio (fecha y tamano): en la carpeta de
    trabajo index.html se edita a cada rato, y una huella vieja dejaria la
    pagina sin guion -o sea, muerta- hasta reiniciar.
    """
    try:
        estado = ruta_html.stat()
    except OSError:
        return ""
    clave = (estado.st_mtime_ns, estado.st_size)
    guardada = _cache_politica.get(ruta_html)
    if guardada and guardada[0] == clave:
        return guardada[1]
    html = ruta_html.read_text(encoding="utf-8", errors="replace")
    guiones = " ".join(["'self'"] + huellas_de_guiones(html))
    politica = "; ".join([
        "default-src 'self'",
        "script-src " + guiones,
        # Los estilos en linea se quedan: la pagina entera los usa, y un
        # estilo inyectado no ejecuta nada.
        "style-src 'self' 'unsafe-inline'",
        # data: y blob: son las fotos y las paginas de la vista previa.
        "img-src 'self' data: blob:",
        "font-src 'self'",
        "connect-src 'self'",
        "object-src 'none'",
        "base-uri 'none'",
        "form-action 'self'",
        "frame-ancestors 'none'",
    ])
    _cache_politica[ruta_html] = (clave, politica)
    return politica
