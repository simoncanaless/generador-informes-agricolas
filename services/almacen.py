"""Persistencia de curvas, ajustes y mapeos aprendidos.

La interfaz es abstracta y ASINCRONA (DIP). Hay dos implementaciones:

  - :class:`AlmacenJSON`  archivos en disco. Arranca sin instalar nada y sirve
    para probar el proyecto en cualquier maquina.
  - :class:`~db.repositorios.AlmacenPostgres`  PostgreSQL, la de produccion.

El sistema elige una u otra segun haya ``DATABASE_URL`` en el .env, y ningun
otro modulo se entera de cual esta usando.

Los metodos son ``async`` aunque la version JSON no lo necesite: si la interfaz
fuera sincrona, la implementacion sobre asyncpg no podria existir sin bloquear
el event loop. Se define el contrato por la implementacion mas exigente.
"""

from __future__ import annotations

import asyncio
import json
import os
import threading
from abc import ABC, abstractmethod
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Type, TypeVar, Union

from pydantic import BaseModel

from schemas.curvas import (Ajuste, CurvaProduccion, LecturaIA,
                            MapeoCatastro)

M = TypeVar("M", bound=BaseModel)

__all__ = ["Almacen", "AlmacenJSON", "construir_almacen", "NoEncontrado"]


class NoEncontrado(KeyError):
    """El identificador pedido no existe en el almacen."""


class Almacen(ABC):
    """Contrato de persistencia. Ninguna firma menciona JSON ni SQL."""

    #: Etiqueta para el endpoint de salud.
    tipo: str = "abstracto"

    # -- curvas
    @abstractmethod
    async def listar_curvas(self) -> List[CurvaProduccion]: ...

    @abstractmethod
    async def obtener_curva(self, id_curva: str) -> CurvaProduccion: ...

    @abstractmethod
    async def guardar_curva(self, curva: CurvaProduccion) -> CurvaProduccion: ...

    @abstractmethod
    async def borrar_curva(self, id_curva: str) -> None: ...

    # -- ajustes
    @abstractmethod
    async def listar_ajustes(self) -> List[Ajuste]: ...

    @abstractmethod
    async def guardar_ajuste(self, ajuste: Ajuste) -> Ajuste: ...

    @abstractmethod
    async def borrar_ajuste(self, id_ajuste: str) -> None: ...

    # -- mapeos de catastro aprendidos
    @abstractmethod
    async def buscar_mapeo(self, huella: str) -> Optional[MapeoCatastro]: ...

    @abstractmethod
    async def guardar_mapeo(self, mapeo: MapeoCatastro) -> MapeoCatastro: ...

    @abstractmethod
    async def listar_mapeos(self) -> List[MapeoCatastro]: ...

    # -- lecturas de la IA ya pagadas

    @abstractmethod
    async def buscar_lectura(self, huella: str) -> Optional[LecturaIA]: ...

    @abstractmethod
    async def guardar_lectura(self, lectura: LecturaIA) -> LecturaIA: ...


class AlmacenJSON(Almacen):
    """Implementacion sobre archivos JSON, una coleccion por archivo.

    La escritura es atomica (archivo temporal + reemplazo) para que un corte a
    mitad de guardado no deje el catalogo de curvas corrupto. El acceso a disco
    se delega a un hilo para no bloquear el event loop.
    """

    tipo = "json"

    def __init__(self, carpeta: Union[str, Path]):
        self.carpeta = Path(carpeta)
        self.carpeta.mkdir(parents=True, exist_ok=True)
        self._cerrojo = threading.Lock()

    # -- primitivas ------------------------------------------------------

    def _ruta(self, coleccion: str) -> Path:
        return self.carpeta / "{}.json".format(coleccion)

    def _leer(self, coleccion: str, modelo: Type[M]) -> List[M]:
        ruta = self._ruta(coleccion)
        if not ruta.exists():
            return []
        try:
            crudo = json.loads(ruta.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return []
        elementos = []
        for registro in crudo:
            try:
                elementos.append(modelo.model_validate(registro))
            except Exception:  # noqa: BLE001 - un registro viejo no tumba el catalogo
                continue
        return elementos

    def _escribir(self, coleccion: str, elementos: List[BaseModel]) -> None:
        ruta = self._ruta(coleccion)
        temporal = ruta.with_suffix(".tmp")
        temporal.write_text(
            json.dumps(
                [e.model_dump(mode="json") for e in elementos],
                ensure_ascii=False, indent=2,
            ),
            encoding="utf-8",
        )
        os.replace(temporal, ruta)

    def _guardar_sync(self, coleccion: str, modelo: Type[M], elemento: M) -> M:
        with self._cerrojo:
            elementos = self._leer(coleccion, modelo)
            for indice, existente in enumerate(elementos):
                if existente.id == elemento.id:  # type: ignore[attr-defined]
                    elementos[indice] = elemento
                    break
            else:
                elementos.append(elemento)
            self._escribir(coleccion, elementos)
        return elemento

    def _borrar_sync(self, coleccion: str, modelo: Type[M], id_elemento: str) -> None:
        with self._cerrojo:
            elementos = self._leer(coleccion, modelo)
            quedan = [e for e in elementos
                      if e.id != id_elemento]  # type: ignore[attr-defined]
            if len(quedan) == len(elementos):
                raise NoEncontrado(id_elemento)
            self._escribir(coleccion, quedan)

    # -- curvas ----------------------------------------------------------

    async def listar_curvas(self) -> List[CurvaProduccion]:
        curvas = await asyncio.to_thread(self._leer, "curvas", CurvaProduccion)
        return sorted(curvas, key=lambda c: c.nombre.lower())

    async def obtener_curva(self, id_curva: str) -> CurvaProduccion:
        for curva in await self.listar_curvas():
            if curva.id == id_curva:
                return curva
        raise NoEncontrado("no existe la curva {}".format(id_curva))

    async def guardar_curva(self, curva: CurvaProduccion) -> CurvaProduccion:
        curva.actualizada = datetime.now()
        return await asyncio.to_thread(
            self._guardar_sync, "curvas", CurvaProduccion, curva
        )

    async def borrar_curva(self, id_curva: str) -> None:
        await asyncio.to_thread(self._borrar_sync, "curvas", CurvaProduccion, id_curva)

    # -- ajustes ---------------------------------------------------------

    async def listar_ajustes(self) -> List[Ajuste]:
        ajustes = await asyncio.to_thread(self._leer, "ajustes", Ajuste)
        return sorted(ajustes, key=lambda a: (a.categoria.value, a.nombre.lower()))

    async def guardar_ajuste(self, ajuste: Ajuste) -> Ajuste:
        return await asyncio.to_thread(self._guardar_sync, "ajustes", Ajuste, ajuste)

    async def borrar_ajuste(self, id_ajuste: str) -> None:
        await asyncio.to_thread(self._borrar_sync, "ajustes", Ajuste, id_ajuste)

    # -- mapeos ----------------------------------------------------------

    async def listar_mapeos(self) -> List[MapeoCatastro]:
        return await asyncio.to_thread(self._leer, "mapeos", MapeoCatastro)

    async def buscar_mapeo(self, huella: str) -> Optional[MapeoCatastro]:
        for mapeo in await self.listar_mapeos():
            if mapeo.huella == huella:
                return mapeo
        return None

    def _guardar_mapeo_sync(self, mapeo: MapeoCatastro) -> MapeoCatastro:
        with self._cerrojo:
            mapeos = self._leer("mapeos", MapeoCatastro)
            for indice, existente in enumerate(mapeos):
                if existente.huella == mapeo.huella:
                    mapeo.id = existente.id
                    mapeo.veces_usado = existente.veces_usado + 1
                    mapeos[indice] = mapeo
                    break
            else:
                mapeos.append(mapeo)
            self._escribir("mapeos", mapeos)
        return mapeo

    async def guardar_mapeo(self, mapeo: MapeoCatastro) -> MapeoCatastro:
        return await asyncio.to_thread(self._guardar_mapeo_sync, mapeo)

    # -- lecturas de la IA -------------------------------------------------

    async def buscar_lectura(self, huella: str) -> Optional[LecturaIA]:
        for lectura in await asyncio.to_thread(self._leer, "lecturas", LecturaIA):
            if lectura.huella == huella:
                return lectura
        return None

    def _guardar_lectura_sync(self, lectura: LecturaIA) -> LecturaIA:
        with self._cerrojo:
            lecturas = self._leer("lecturas", LecturaIA)
            for indice, existente in enumerate(lecturas):
                if existente.huella == lectura.huella:
                    lectura.id = existente.id
                    lectura.veces_usado = existente.veces_usado + 1
                    lecturas[indice] = lectura
                    break
            else:
                lecturas.append(lectura)
            self._escribir("lecturas", lecturas)
        return lectura

    async def guardar_lectura(self, lectura: LecturaIA) -> LecturaIA:
        return await asyncio.to_thread(self._guardar_lectura_sync, lectura)


_almacen: Optional[Almacen] = None


def construir_almacen(carpeta: Optional[Union[str, Path]] = None) -> Almacen:
    """Devuelve el almacen del proceso: PostgreSQL si esta configurado, si no JSON.

    Que la eleccion viva en una sola funcion es lo que permite cambiar de motor
    sin tocar routers ni servicios.
    """
    global _almacen

    # Una carpeta explicita siempre da un almacen JSON nuevo y sin cachear:
    # es lo que usan las pruebas para trabajar en un directorio temporal.
    if carpeta is not None:
        return AlmacenJSON(carpeta)

    if _almacen is not None:
        return _almacen

    from config.settings import obtener_ajustes

    config = obtener_ajustes()
    if config.bd_configurada:
        from db.repositorios import AlmacenPostgres

        _almacen = AlmacenPostgres()
    else:
        _almacen = AlmacenJSON(config.raiz / "data" / "almacen")
    return _almacen


def reiniciar_almacen() -> None:
    """Olvida la instancia cacheada. Solo para pruebas."""
    global _almacen
    _almacen = None
