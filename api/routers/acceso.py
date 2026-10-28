"""Entrar y salir. La identidad viaja en una cookie firmada.

Por que cookie de sesion y no un token en el navegador: la pagina es una sola,
servida por el mismo servidor que la API, asi que no hay nada que ganar con un
JWT y si algo que perder -un token en `localStorage` se lo lleva cualquier
script-. La cookie va `HttpOnly`, firmada por el servidor, y caduca.

No hay registro publico a proposito. Las cuentas las crea el administrador con
`scripts/usuarios.py`: son tres personas de la misma oficina, y una pantalla de
alta abierta a la red seria una puerta que nadie necesita.
"""

from __future__ import annotations

import asyncio
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, HTTPException, Request, status
from pydantic import BaseModel, Field

from api.dependencias import CLAVE_HUELLA, Config, Usuarios, UsuarioActual
from services.intentos import Intentos, claves_de, texto_de_espera
from services.usuarios import (
    RutInvalido, comparar_con_senuelo, contrasena_correcta, formatear_rut,
    huella_de_sesion, normalizar_rut,
)

router = APIRouter(prefix="/acceso", tags=["acceso"])

#: Clave con la que se guarda el usuario dentro de la cookie de sesion.
CLAVE_SESION = "usuario_id"

#: Los fallos recientes, por RUT y por equipo (ver services/intentos.py).
intentos = Intentos()


class PeticionEntrada(BaseModel):
    # Con tope: una "contrasena" de 50 MB costaria lo mismo de procesar que
    # mil intentos.
    rut: str = Field(max_length=40, description="con o sin puntos, da igual")
    contrasena: str = Field(max_length=200)


def _ip(request: Request) -> Optional[str]:
    return request.client.host if request.client else None


def _frenar_si_corresponde(claves: List[str]) -> None:
    espera = intentos.espera(claves)
    if espera > 0:
        raise HTTPException(
            status.HTTP_429_TOO_MANY_REQUESTS,
            "Demasiados intentos fallidos. Espere {} y vuelva a intentar."
            .format(texto_de_espera(espera)),
            headers={"Retry-After": str(int(espera) + 1)},
        )


class UsuarioVisible(BaseModel):
    """Lo que la pantalla necesita saber. La clave nunca sale de la base."""

    id: str
    rut: str
    nombre: str


def _visible(fila) -> UsuarioVisible:
    return UsuarioVisible(
        id=str(fila.id), rut=formatear_rut(fila.rut), nombre=fila.nombre)


@router.post("/entrar", response_model=UsuarioVisible, summary="Iniciar sesion")
async def entrar(
    peticion: PeticionEntrada, request: Request, repo: Usuarios, config: Config,
) -> UsuarioVisible:
    # El mismo mensaje para RUT inexistente, RUT mal escrito y contrasena
    # equivocada: decir cual de los tres fallo le regala a quien prueba la
    # mitad del trabajo.
    generico = "El RUT o la contraseña no son correctos."
    try:
        rut = normalizar_rut(peticion.rut)
    except RutInvalido:
        rut = None
    claves = claves_de(rut or peticion.rut.strip(), _ip(request))
    # Antes de mirar la clave: si no, quien prueba igual sabria cuando acerto.
    _frenar_si_corresponde(claves)

    fila = await repo.buscar_por_rut(rut) if rut else None
    if fila is None or not fila.activo:
        # Mismo tiempo de respuesta que una clave equivocada de un RUT real.
        await asyncio.to_thread(comparar_con_senuelo, peticion.contrasena)
        intentos.fallo(claves)
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, generico)
    # scrypt tarda ~50 ms a proposito: en un hilo aparte, para que mientras
    # tanto el servidor siga atendiendo a los demas.
    if not await asyncio.to_thread(contrasena_correcta, peticion.contrasena,
                                   fila.clave):
        intentos.fallo(claves)
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, generico)

    intentos.acierto(claves)
    request.session.clear()
    request.session[CLAVE_SESION] = str(fila.id)
    request.session[CLAVE_HUELLA] = huella_de_sesion(
        fila.clave, config.secreto_de_sesion())
    await repo.marcar_entrada(str(fila.id))
    return _visible(fila)


@router.post("/salir", summary="Cerrar sesion")
async def salir(request: Request) -> Dict[str, Any]:
    request.session.clear()
    return {"salio": True}


@router.get("/yo", response_model=UsuarioVisible, summary="Quien esta dentro")
async def yo(usuario: UsuarioActual) -> UsuarioVisible:
    return _visible(usuario)


@router.get("/personas", response_model=List[UsuarioVisible],
            summary="Con quien se puede compartir un informe")
async def personas(repo: Usuarios, usuario: UsuarioActual) -> List[UsuarioVisible]:
    """Las cuentas activas, para poder elegir a quien pasarle un informe.

    Solo el nombre y el RUT: es una lista para elegir, no un directorio.
    """
    return [_visible(f) for f in await repo.listar()
            if f.activo and str(f.id) != str(usuario.id)]


class CambioClave(BaseModel):
    actual: str = Field(max_length=200)
    nueva: str = Field(min_length=8, max_length=200)


@router.post("/clave", summary="Cambiar la propia contrasena")
async def cambiar_clave(
    peticion: CambioClave, request: Request, usuario: UsuarioActual,
    repo: Usuarios, config: Config,
) -> Dict[str, Any]:
    """Cada uno cambia la suya, nadie cambia la de otro.

    Existe porque las claves iniciales las escribio el administrador: hasta que
    cada persona pone la suya, hay alguien mas que la conoce.

    Cambiarla cierra las demas sesiones abiertas con la clave anterior (ver
    `huella_de_sesion`), menos esta: quien la cambio sigue dentro. Y pedir la
    actual tiene el mismo freno que entrar: si no, una sesion ajena olvidada
    abierta serviria para adivinarla sin limite.
    """
    from services.usuarios import cifrar_contrasena

    claves = ["clave:" + str(usuario.id)]
    _frenar_si_corresponde(claves)
    if not await asyncio.to_thread(contrasena_correcta, peticion.actual,
                                   usuario.clave):
        intentos.fallo(claves)
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST, "La contraseña actual no es correcta.")
    intentos.acierto(claves)
    nueva = await asyncio.to_thread(cifrar_contrasena, peticion.nueva)
    await repo.cambiar_clave(str(usuario.id), nueva)
    request.session[CLAVE_HUELLA] = huella_de_sesion(
        nueva, config.secreto_de_sesion())
    return {"cambiada": True}
