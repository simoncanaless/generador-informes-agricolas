"""Entorno de Alembic.

Toma la conexion de ``DATABASE_URL`` (.env) y los metadatos de ``db.modelos``,
de forma que ``alembic revision --autogenerate`` detecte solo los cambios de los
modelos y la contrasena no quede escrita en ningun archivo versionado.
"""

from __future__ import annotations

import asyncio
import os
import sys
from logging.config import fileConfig

from alembic import context
from sqlalchemy import pool
from sqlalchemy.engine import Connection
from sqlalchemy.ext.asyncio import async_engine_from_config

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.settings import obtener_ajustes  # noqa: E402
from db.modelos import Base  # noqa: E402

config = context.config

if config.config_file_name is not None:
    fileConfig(config.config_file_name)

ajustes = obtener_ajustes()
if ajustes.database_url:
    # set_main_option escapa los % del propio ConfigParser; una contrasena con
    # '%' rompe la carga si no se duplica.
    config.set_main_option("sqlalchemy.url", ajustes.database_url.replace("%", "%%"))

target_metadata = Base.metadata


def ejecutar_offline() -> None:
    """Genera el SQL sin conectarse. Util para revisar antes de aplicar."""
    context.configure(
        url=config.get_main_option("sqlalchemy.url"),
        target_metadata=target_metadata,
        literal_binds=True,
        dialect_opts={"paramstyle": "named"},
        compare_type=True,
    )
    with context.begin_transaction():
        context.run_migrations()


def _migrar(conexion: Connection) -> None:
    context.configure(
        connection=conexion,
        target_metadata=target_metadata,
        compare_type=True,          # detecta cambios de tipo de columna
        compare_server_default=True,
    )
    with context.begin_transaction():
        context.run_migrations()


async def ejecutar_online() -> None:
    motor = async_engine_from_config(
        config.get_section(config.config_ini_section, {}),
        prefix="sqlalchemy.",
        poolclass=pool.NullPool,
    )
    async with motor.connect() as conexion:
        await conexion.run_sync(_migrar)
    await motor.dispose()


if context.is_offline_mode():
    ejecutar_offline()
else:
    asyncio.run(ejecutar_online())
