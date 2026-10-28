"""Diagnostico del servicio.

Distingue a proposito dos cosas que suelen confundirse: que la API este viva y
que la IA tenga cuota. El 80% determinista funciona sin IA, asi que quedarse sin
modelos NO deja el servicio caido.
"""

from __future__ import annotations

from typing import Any, Dict

from fastapi import APIRouter

from api.dependencias import Almacenamiento, Config, IAOpcional, Motor

router = APIRouter(tags=["salud"])


def _marca_de_la_pagina(config) -> Dict[str, Any]:
    """Que version de `web/index.html` esta sirviendo el servidor.

    Toda la aplicacion es un solo archivo, asi que un navegador con la copia
    vieja corre el PROGRAMA viejo. Ya paso: un arreglo puesto, servido y
    probado, y la pantalla comportandose como antes (§16.a.septdecies). La
    pagina muestra esta marca; si la de la pantalla no coincide con la del
    servidor, el navegador se quedo atras y no hay que buscar el error en
    otro sitio.
    """
    import hashlib
    from datetime import datetime

    ruta = config.raiz / "web" / "index.html"
    if not ruta.is_file():
        return {"encontrada": False}
    datos = ruta.read_bytes()
    return {
        "encontrada": True,
        "huella": hashlib.sha256(datos).hexdigest()[:8],
        "modificada": datetime.fromtimestamp(
            ruta.stat().st_mtime).strftime("%d-%m-%Y %H:%M"),
    }


@router.get("/salud", summary="Estado del servicio, la base y la cadena de modelos")
async def salud(
    config: Config, motor: Motor, ia: IAOpcional, almacen: Almacenamiento
) -> Dict[str, Any]:
    plantilla = getattr(motor, "plantilla", config.ruta_plantilla)
    plantilla_ok = plantilla.exists()
    mapa = motor.mapa
    if isinstance(mapa, dict):
        # El formato final de Santander: su mapping es un dict (partes que se
        # repiten en vez de tablas fijas).
        excel_mapa = {"hojas_mapeadas": len(mapa.get("hojas", {})),
                      "campos_mapeados": len(mapa.get("campos", {})),
                      "tablas_mapeadas": len(mapa.get("partes", [])),
                      "version_mapping": mapa.get("version")}
    else:
        excel_mapa = {"hojas_mapeadas": len(mapa.hojas),
                      "campos_mapeados": len(mapa.campos),
                      "tablas_mapeadas": len(mapa.tablas),
                      "version_mapping": mapa.version}

    if config.bd_configurada:
        from db.sesion import comprobar_conexion

        base = await comprobar_conexion(config)
    else:
        base = {
            "conectada": False,
            "motivo": "sin DATABASE_URL: se usan archivos JSON",
        }
    base["almacen_en_uso"] = almacen.tipo

    return {
        "base_datos": base,
        "estado": "operativo" if plantilla_ok else "degradado",
        "version": config.api_version,
        "pagina": _marca_de_la_pagina(config),
        "entorno": config.api_entorno,
        "excel": dict({
            "plantilla": plantilla.name,
            "plantilla_encontrada": plantilla_ok,
        }, **excel_mapa),
        "ia": ia.estado() if ia is not None else {
            "proveedor": None,
            "detalle": "sin GEMINI_API_KEYS: el 80% determinista sigue operativo",
        },
    }


@router.get("/salud/modelos", summary="Modelos que la clave puede usar realmente")
async def modelos(ia: IAOpcional) -> Dict[str, Any]:
    if ia is None:
        return {"configurada": False, "modelos": []}
    disponibles = await ia.listar_modelos()
    cadena = ia.estado()["cadena"]
    return {
        "configurada": True,
        "cadena_configurada": cadena,
        "modelos_de_la_cuenta": disponibles,
        "cadena_invalida": [m for m in cadena if m not in disponibles],
    }
