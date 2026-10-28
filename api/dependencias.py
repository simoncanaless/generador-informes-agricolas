"""Inyeccion de dependencias de la API.

Los servicios caros (leer el mapping, crear los clientes de Gemini) se
construyen una vez por proceso y se comparten. Los routers piden abstracciones,
nunca implementaciones concretas.
"""

from __future__ import annotations

from functools import lru_cache
from typing import Any, Optional, Union

from fastapi import Depends, HTTPException, Request, UploadFile, status
from typing_extensions import Annotated

from config.settings import Ajustes, obtener_ajustes
from services.ai_service import ProveedorIA, construir_servicio_ia
from services.almacen import Almacen, construir_almacen
from services.catastro_import import ReglasAgronomicas, cargar_reglas
from services.excel_engine import MotorExcel
from services.excel_engine_rabobank import MotorRabobank
from services.excel_engine_santander import MotorSantander
from services.mapping import cargar_mapa
from services.pdf_parser import ParserPDF

Config = Annotated[Ajustes, Depends(obtener_ajustes)]


@lru_cache(maxsize=8)
def _motor_de(alias: str):
    """Un motor por formato de salida. Se cachean porque leer el mapping cuesta.

    El formato final de Santander tiene su propio motor (repite las partes
    del molde segun el informe); los demas los llena MotorExcel."""
    config = obtener_ajustes()
    plantilla, mapping = config.rutas_de_plantilla(alias)
    if config.motor_de_plantilla(alias) == "santander2026":
        return MotorSantander(plantilla, mapping, reglas=_reglas())
    return MotorExcel(plantilla, cargar_mapa(mapping))


def _motor() -> MotorExcel:
    return _motor_de(obtener_ajustes().plantilla_predeterminada)


def motor_para(alias: str) -> MotorExcel:
    """Motor del formato pedido; 404 si el alias no existe."""
    try:
        return _motor_de(alias)
    except KeyError as error:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND, detail=str(error)
        ) from error


@lru_cache(maxsize=1)
def _motor_rabobank() -> MotorRabobank:
    """Motor de retasaciones Rabobank, aparte de MotorExcel a proposito: es
    un dominio distinto (garantia, no capacidad de pago) -"no mezclar los
    motores"-, asi que no comparte instancia con `_motor_de`."""
    plantilla, mapping = obtener_ajustes().rutas_de_plantilla("rabobank")
    return MotorRabobank(plantilla, cargar_mapa(mapping))


def obtener_motor_rabobank() -> MotorRabobank:
    return _motor_rabobank()


@lru_cache(maxsize=1)
def _ia() -> Optional[ProveedorIA]:
    return construir_servicio_ia(obtener_ajustes())


@lru_cache(maxsize=1)
def _almacen() -> Almacen:
    return construir_almacen()


@lru_cache(maxsize=1)
def _reglas() -> ReglasAgronomicas:
    config = obtener_ajustes()
    return cargar_reglas(config.raiz / "config" / "reglas_agronomicas.json")


def obtener_almacen() -> Almacen:
    """Persistencia de curvas, ajustes y formatos aprendidos."""
    return _almacen()


def obtener_reglas() -> ReglasAgronomicas:
    """Reglas agronomicas: sinonimos de columnas y umbrales de madurez."""
    return _reglas()


def obtener_motor() -> MotorExcel:
    """El motor Excel siempre esta disponible: no depende de la IA."""
    return _motor()


def obtener_ia() -> ProveedorIA:
    """Proveedor de IA, o 503 si nadie configuro una clave."""
    servicio = _ia()
    if servicio is None:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=(
                "las funciones de IA no estan configuradas: define "
                "GEMINI_API_KEYS en el archivo .env"
            ),
        )
    return servicio


def obtener_ia_opcional() -> Optional[ProveedorIA]:
    """Para el endpoint de salud, que debe responder aunque no haya IA."""
    return _ia()


Motor = Annotated[Union[MotorExcel, MotorSantander], Depends(obtener_motor)]
MotorRabo = Annotated[MotorRabobank, Depends(obtener_motor_rabobank)]
IA = Annotated[ProveedorIA, Depends(obtener_ia)]
IAOpcional = Annotated[Optional[ProveedorIA], Depends(obtener_ia_opcional)]
def obtener_repositorio_expedientes():
    """Repositorio de expedientes, o 503 si no hay base de datos.

    Es el unico bloque del sistema que PostgreSQL hace imprescindible: el resto
    (generar informes, leer documentos, curvas) funciona sin base.
    """
    if not obtener_ajustes().bd_configurada:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=(
                "los expedientes necesitan PostgreSQL: define DATABASE_URL en "
                "el archivo .env. El resto del sistema funciona sin base."
            ),
        )
    from db.repositorios import RepositorioExpedientes

    return _repositorio_expedientes()


@lru_cache(maxsize=1)
def _repositorio_expedientes():
    from db.repositorios import RepositorioExpedientes

    return RepositorioExpedientes()


def obtener_repositorio_opcional():
    """Repositorio si hay base, o ``None``.

    Lo usa el asistente: guardar el avance es deseable, pero que la base este
    caida no puede impedir que el tasador lea una escritura o genere su Excel.
    """
    if not obtener_ajustes().bd_configurada:
        return None
    try:
        return _repositorio_expedientes()
    except Exception:  # noqa: BLE001 - la persistencia es un extra, no un requisito
        return None


Almacenamiento = Annotated[Almacen, Depends(obtener_almacen)]
Reglas = Annotated[ReglasAgronomicas, Depends(obtener_reglas)]
Expedientes = Annotated[Any, Depends(obtener_repositorio_expedientes)]
ExpedientesOpcional = Annotated[Any, Depends(obtener_repositorio_opcional)]


# ------------------------------------------------------------ quien entra

@lru_cache(maxsize=1)
def _usuarios():
    from db.repositorios import RepositorioUsuarios
    return RepositorioUsuarios()


def obtener_repositorio_usuarios():
    """Las cuentas. Requiere base de datos: sin ella no hay a quien dejar
    entrar, y abrir el programa a cualquiera seria peor que no arrancar."""
    config = obtener_ajustes()
    if not config.bd_configurada:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="no hay base de datos: sin ella no se puede validar quien entra",
        )
    return _usuarios()


Usuarios = Annotated[Any, Depends(obtener_repositorio_usuarios)]


#: Clave de la cookie donde va la firma de la contrasena con que se entro.
CLAVE_HUELLA = "huella"


async def obtener_usuario_actual(request: Request, repo: Usuarios):
    """El usuario de la cookie de sesion, o 401.

    Es la unica puerta: cualquier router que la pida queda protegido, y el que
    no la pida queda abierto. Por eso conviene mirarla al agregar rutas
    (scripts/test_seguridad.py recorre todas y avisa si una nueva quedo
    abierta sin estar en la lista de las que deben estarlo).
    """
    from services.usuarios import huella_de_sesion

    id_usuario = request.session.get("usuario_id")
    if not id_usuario:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="hay que iniciar sesion",
        )
    fila = await repo.obtener(id_usuario)
    if fila is None or not fila.activo:
        # La cuenta se borro o se desactivo mientras la sesion seguia viva.
        request.session.clear()
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="la sesion ya no es valida",
        )
    actual = huella_de_sesion(fila.clave, obtener_ajustes().secreto_de_sesion())
    huella = request.session.get(CLAVE_HUELLA)
    if huella is None:
        # Sesion abierta antes de que existiera la firma: se completa en vez
        # de echar a todos el dia de la actualizacion.
        request.session[CLAVE_HUELLA] = actual
    elif huella != actual:
        # La contrasena cambio despues de abrir esta sesion.
        request.session.clear()
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="la contraseña cambió: vuelva a entrar",
        )
    return fila


UsuarioActual = Annotated[Any, Depends(obtener_usuario_actual)]


def obtener_parser_pdf(ia: IA) -> ParserPDF:
    """El parser es barato: se compone sobre el proveedor ya inyectado."""
    return ParserPDF(ia)


def obtener_parser_opcional(ia: IAOpcional) -> Optional[ParserPDF]:
    """El mismo parser, o None si no hay IA configurada.

    Leer una escritura es lo UNICO que el parser sabe hacer, y para eso hace
    falta el modelo. Sin el no se levanta un parser a medias: se devuelve None
    y quien lo pidio decide que hacer -en el asistente, dejar los campos
    legales en blanco para que el tasador los escriba-.
    """
    return ParserPDF(ia) if ia is not None else None


Parser = Annotated[ParserPDF, Depends(obtener_parser_pdf)]
ParserOpcional = Annotated[Optional[ParserPDF], Depends(obtener_parser_opcional)]


async def leer_acotado(archivo: UploadFile, maximo_bytes: int, etiqueta: str) -> bytes:
    """Lee un archivo subido rechazandolo si excede el limite.

    Se lee por trozos para no dejar que un archivo enorme agote la memoria
    antes de poder rechazarlo.
    """
    trozos = bytearray()
    while trozo := await archivo.read(1024 * 1024):
        trozos.extend(trozo)
        if len(trozos) > maximo_bytes:
            raise HTTPException(
                status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                detail="{} supera el limite de {} MB".format(
                    etiqueta, maximo_bytes // (1024 * 1024)
                ),
            )
    if not trozos:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="{}: el archivo llego vacio".format(etiqueta),
        )
    return bytes(trozos)
