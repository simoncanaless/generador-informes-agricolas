"""Motor asincrono y sesiones de PostgreSQL.

La conexion se crea una vez por proceso y se comparte. Si no hay
``DATABASE_URL`` configurada, todo esto queda inerte y el sistema sigue
guardando en archivos JSON: la aplicacion tiene que poder arrancar y generar
informes aunque nadie haya montado la base todavia.
"""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager
from typing import AsyncIterator, Optional

from sqlalchemy.ext.asyncio import (
    AsyncEngine, AsyncSession, async_sessionmaker, create_async_engine,
)

from config.settings import Ajustes, obtener_ajustes
from db.modelos import Base

log = logging.getLogger("fitogreen.db")

_motor: Optional[AsyncEngine] = None
_fabrica: Optional[async_sessionmaker[AsyncSession]] = None


def motor(config: Optional[Ajustes] = None) -> AsyncEngine:
    """Motor unico del proceso."""
    global _motor
    if _motor is None:
        config = config or obtener_ajustes()
        if not config.bd_configurada:
            raise RuntimeError(
                "no hay DATABASE_URL configurada: define la conexion en .env"
            )
        _motor = create_async_engine(
            config.database_url,
            echo=config.db_echo,
            pool_size=config.db_pool_max,
            pool_pre_ping=True,   # una conexion caida se detecta antes de usarla
        )
    return _motor


def fabrica_de_sesiones(config: Optional[Ajustes] = None) -> async_sessionmaker[AsyncSession]:
    global _fabrica
    if _fabrica is None:
        _fabrica = async_sessionmaker(
            motor(config), expire_on_commit=False, class_=AsyncSession
        )
    return _fabrica


@asynccontextmanager
async def sesion() -> AsyncIterator[AsyncSession]:
    """Sesion con commit al salir y rollback si algo falla."""
    async with fabrica_de_sesiones()() as sesion_activa:
        try:
            yield sesion_activa
            await sesion_activa.commit()
        except Exception:
            await sesion_activa.rollback()
            raise


async def crear_base_si_falta(config: Optional[Ajustes] = None) -> bool:
    """Crea la base del proyecto si el servidor la tiene aun sin crear.

    Conectarse a 'postgres' para poder emitir el CREATE DATABASE es el rodeo
    habitual: no se puede crear una base estando dentro de ella.
    """
    import asyncpg

    config = config or obtener_ajustes()
    nombre = config.nombre_base_datos
    if not nombre:
        return False

    url_admin = config.url_sin_base().replace("postgresql+asyncpg://", "postgresql://")
    conexion = await asyncpg.connect(url_admin)
    try:
        existe = await conexion.fetchval(
            "select 1 from pg_database where datname = $1", nombre
        )
        if existe:
            return False
        # El nombre no puede ir como parametro en un CREATE DATABASE; se cita.
        await conexion.execute('create database "{}"'.format(nombre.replace('"', '""')))
        log.info("base de datos '%s' creada", nombre)
        return True
    finally:
        await conexion.close()


async def crear_tablas(config: Optional[Ajustes] = None) -> None:
    """Crea las tablas que falten. En produccion esto lo hace Alembic."""
    async with motor(config).begin() as conexion:
        await conexion.run_sync(Base.metadata.create_all)


async def comprobar_conexion(config: Optional[Ajustes] = None) -> dict:
    """Diagnostico para el endpoint de salud."""
    from sqlalchemy import text

    config = config or obtener_ajustes()
    if not config.bd_configurada:
        return {"conectada": False, "motivo": "sin DATABASE_URL"}
    try:
        async with sesion() as s:
            version = await s.scalar(text("select version()"))
            tablas = await s.scalar(text(
                "select count(*) from information_schema.tables "
                "where table_schema = 'public'"
            ))
        return {
            "conectada": True,
            "servidor": str(version).split(",")[0],
            "base": config.nombre_base_datos,
            "tablas": int(tablas or 0),
        }
    except Exception as error:  # noqa: BLE001 - el diagnostico no debe tumbar la API
        return {
            "conectada": False,
            "motivo": "{}: {}".format(type(error).__name__, str(error)[:200]),
        }


async def cerrar() -> None:
    global _motor, _fabrica
    if _motor is not None:
        await _motor.dispose()
    _motor, _fabrica = None, None
