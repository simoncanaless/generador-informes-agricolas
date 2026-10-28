"""Punto de entrada de la API de tasaciones Fitogreen.

Arranque:  uvicorn main:app --reload
Docs:      http://127.0.0.1:8000/docs

La aplicacion arranca aunque no haya clave de IA configurada: la generacion de
informes Excel es 100% determinista y no depende de ningun modelo. Solo los
endpoints bajo /ia responden 503 en ese caso.
"""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager
from pathlib import Path
from typing import AsyncIterator

from fastapi import FastAPI, Request, status
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.sessions import SessionMiddleware
from fastapi.responses import JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from api.dependencias import obtener_ia_opcional, obtener_motor
from api.seguridad import (
    CabecerasDeSeguridad, RedesPermitidas, leer_redes, politica_de_contenido,
)
from api.routers import asistente as router_asistente
from api.routers import catastro as router_catastro
from api.routers import curvas as router_curvas
from api.routers import referencias as router_referencias
from api.routers import expedientes as router_expedientes
from api.routers import ia as router_ia
from api.routers import informes as router_informes
from api.routers import retasaciones as router_retasaciones
from api.routers import acceso as router_acceso
from api.routers import salud as router_salud
from api.routers import visita as router_visita
from config.settings import obtener_ajustes
from services.ai_service import IANoConfigurada
from services.model_router import SinModelosDisponibles
from services.pdf_parser import PdfIlegible, PdfProtegido

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
)
log = logging.getLogger("fitogreen")


@asynccontextmanager
async def ciclo_de_vida(app: FastAPI) -> AsyncIterator[None]:
    """Prepara y comprueba todo al arrancar, para fallar pronto y con claridad."""
    config = obtener_ajustes()
    config.preparar_directorios()

    if not config.ruta_plantilla.exists():
        log.error("NO se encontro la plantilla: %s", config.ruta_plantilla)
    else:
        motor = obtener_motor()
        log.info(
            "motor Excel listo | plantilla=%s mapping v%s (%s campos, %s tablas)",
            config.ruta_plantilla.name, motor.mapa.version,
            len(motor.mapa.campos), len(motor.mapa.tablas),
        )

    # --- base de datos: si no esta configurada, se sigue con archivos JSON
    if config.bd_configurada:
        from db.sesion import cerrar, comprobar_conexion, crear_base_si_falta, crear_tablas

        try:
            if config.db_crear_al_arrancar:
                if await crear_base_si_falta(config):
                    log.info("se creo la base '%s'", config.nombre_base_datos)
                await crear_tablas(config)
            estado = await comprobar_conexion(config)
            if estado["conectada"]:
                log.info("PostgreSQL listo | %s | %s tablas",
                         estado["servidor"], estado["tablas"])
            else:
                log.error("PostgreSQL NO disponible: %s", estado["motivo"])
        except Exception as error:  # noqa: BLE001
            log.error(
                "no se pudo preparar PostgreSQL (%s: %s). El sistema seguira "
                "guardando en archivos JSON.", type(error).__name__, error
            )
    else:
        log.warning(
            "sin DATABASE_URL: curvas y ajustes se guardan en data/almacen/*.json "
            "y los expedientes no se persisten"
        )

    servicio = obtener_ia_opcional()
    if servicio is None:
        log.warning(
            "IA no configurada (falta GEMINI_API_KEYS): los endpoints /ia "
            "responderan 503. La generacion de informes NO se ve afectada."
        )
    else:
        log.info(
            "IA lista | %s clave(s), cadena: %s",
            len(config.claves_gemini), " -> ".join(config.modelos_gemini),
        )

    yield

    if config.bd_configurada:
        from db.sesion import cerrar

        await cerrar()
    log.info("apagando")


app = FastAPI(
    title=obtener_ajustes().api_titulo,
    version=obtener_ajustes().api_version,
    lifespan=ciclo_de_vida,
    description=(
        "Backend de tasaciones agricolas. El 80% del sistema es deterministico "
        "(calculo y escritura de Excel con openpyxl); la IA solo interviene "
        "para leer documentos, fotografias y notas de voz, y nunca calcula."
    ),
)

# La sesion va en una cookie firmada, HttpOnly y con caducidad. Se monta ANTES
# que CORS a proposito: el orden de los middleware es inverso al de registro,
# asi que este queda por dentro y ve la peticion ya normalizada.
app.add_middleware(
    SessionMiddleware,
    secret_key=obtener_ajustes().secreto_de_sesion(),
    session_cookie="fitogreen_sesion",
    max_age=obtener_ajustes().horas_de_sesion * 3600,
    same_site="lax",
    https_only=False,   # entran por Tailscale en http; no hay TLS de por medio
)

# Sin origenes declarados no se monta CORS: la pagina sale de este mismo
# servidor y no necesita permiso para hablar con su propia API.
if obtener_ajustes().origenes_cors:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=obtener_ajustes().origenes_cors,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

# Los dos ultimos en registrarse son los primeros en ver cada peticion (ver
# api/seguridad.py): primero se mira DESDE DONDE llega, despues todo lo demas.
app.add_middleware(CabecerasDeSeguridad)
app.add_middleware(RedesPermitidas,
                   redes=leer_redes(obtener_ajustes().redes_permitidas))

app.include_router(router_acceso.router)
app.include_router(router_salud.router)
app.include_router(router_informes.router)
app.include_router(router_retasaciones.router)
app.include_router(router_ia.router)
app.include_router(router_asistente.router)
app.include_router(router_catastro.router)
app.include_router(router_curvas.router)
app.include_router(router_expedientes.router)
app.include_router(router_visita.router)
app.include_router(router_referencias.router)

class WebSinCache(StaticFiles):
    """La pagina del tasador, siempre revalidada contra el servidor.

    Toda la aplicacion es UN archivo con su JavaScript y su CSS adentro, asi
    que el navegador que lo guarda en cache sigue corriendo la version vieja
    despues de una actualizacion -sin decirlo, y sin forma de que el usuario lo
    note-. Paso de verdad: un arreglo estaba puesto, servido y probado, y la
    pantalla seguia comportandose como antes. Y le pasaria igual a los tasadores,
    que entran desde otro computador por Tailscale.

    ``no-cache`` no significa "no guardes": significa "pregunta antes de usar".
    El navegador conserva la copia y el servidor responde 304 si no cambio, asi
    que no se pierde velocidad; lo que se pierde es la posibilidad de quedarse
    con una version vieja.
    """

    def file_response(self, *args, **kwargs):            # type: ignore[override]
        respuesta = super().file_response(*args, **kwargs)
        respuesta.headers["Cache-Control"] = "no-cache, must-revalidate"
        # La pagina lleva su politica de contenido: solo se ejecuta SU guion.
        ruta = Path(args[0] if args else kwargs.get("full_path", ""))
        if ruta.suffix.lower() in (".html", ".htm"):
            politica = politica_de_contenido(ruta)
            if politica:
                respuesta.headers["Content-Security-Policy"] = politica
        return respuesta


# La pagina del tasador. Se monta al final para no tapar las rutas de la API.
DIRECTORIO_WEB = obtener_ajustes().raiz / "web"
if DIRECTORIO_WEB.is_dir():
    app.mount("/app", WebSinCache(directory=DIRECTORIO_WEB, html=True), name="web")


# ------------------------------------------------------ manejo de errores

@app.exception_handler(SinModelosDisponibles)
async def sin_modelos(_: Request, error: SinModelosDisponibles) -> JSONResponse:
    """503 con Retry-After: el cliente sabe exactamente cuando reintentar.

    Es un agotamiento temporal de cuota, no un fallo del servicio, y la
    distincion importa para quien consuma la API desde la app de terreno.
    """
    espera = error.segundos_para_reintentar
    cabeceras = {"Retry-After": str(int(espera))} if espera else {}
    log.warning("cadena de modelos agotada: %s", error)
    return JSONResponse(
        status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
        headers=cabeceras,
        content={
            "error": "cuota_agotada",
            "detalle": str(error),
            "reintentar_en_segundos": round(espera, 1) if espera else None,
            "sugerencia": (
                "amplia GEMINI_MODELOS con mas modelos de respaldo o anade otra "
                "clave a GEMINI_API_KEYS en el .env"
            ),
        },
    )


@app.exception_handler(IANoConfigurada)
async def ia_no_configurada(_: Request, error: IANoConfigurada) -> JSONResponse:
    return JSONResponse(
        status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
        content={"error": "ia_no_configurada", "detalle": str(error)},
    )


@app.exception_handler(PdfProtegido)
async def pdf_protegido(_: Request, error: PdfProtegido) -> JSONResponse:
    return JSONResponse(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        content={"error": "pdf_protegido", "detalle": str(error)},
    )


@app.exception_handler(PdfIlegible)
async def pdf_ilegible(_: Request, error: PdfIlegible) -> JSONResponse:
    return JSONResponse(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        content={"error": "pdf_ilegible", "detalle": str(error)},
    )


@app.get("/", include_in_schema=False)
async def raiz() -> RedirectResponse:
    """A la herramienta del tasador, no a la documentacion tecnica.

    Quien entra aqui es quien hace informes; /docs sigue disponible para quien
    integre la API.
    """
    destino = "/app/" if DIRECTORIO_WEB.is_dir() else "/docs"
    return RedirectResponse(url=destino)
