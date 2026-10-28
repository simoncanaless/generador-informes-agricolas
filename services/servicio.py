"""Encender, apagar y vigilar la aplicacion. Un solo sitio.

Hasta aqui habia DOS formas de levantarla -el icono del escritorio y
`scripts/vigilante.ps1`- y ninguna de apagarla: cerrar la ventana la bajaba,
pero el vigilante la volvia a subir a los cinco minutos. O sea que un boton
«Apagar» habria sido mentira.

Este modulo es la unica implementacion de las tres operaciones. El panel
(`scripts/panel.pyw`) y el vigilante llaman aca; tener dos copias es garantia
de que se separen, y la que se separaria es la que usa la persona.

**Solo biblioteca estandar, a proposito.** Lo corre una tarea de Windows cada
cinco minutos: importar pydantic y toda la aplicacion para preguntar si un
puerto responde seria pagar el arranque completo, cada vez, para nada.

## La diferencia entre «se cayo» y «la apagaron»

Un vigilante que no la sepa distinguir es inutil en un sentido o en el otro.
La distingue `data/servicio.json`:

    {"deseado": "encendida", "revivir": true, "al_prender": false}

- `deseado` lo escribe quien aprieta el boton. Si dice `apagada`, el vigilante
  no la toca: la apagaron a proposito.
- `revivir` es el caso que motivo el vigilante (§16.a.tricies.octies): el usuario
  no esta en la casa, la aplicacion se cae sola y hay que poder entrar igual
  desde el telefono.
- `al_prender` es lo que el usuario pidio apagar: que prender el computador NO
  levante la pagina. Por defecto va en falso, y entonces despues de cada
  arranque el estado deseado vuelve a «apagada» hasta que alguien apriete
  Encender.

## Lo que este modulo NO hace

**Nunca mata nada que no sea la aplicacion.** Antes de apagar comprueba que
quien tiene el puerto conteste `/salud`; si contesta otra cosa, se niega y
dice quien es. Es la misma regla del lanzador (§16.a.tricies.septies): un
puerto ocupado no es lo mismo que nuestra aplicacion corriendo.
"""

from __future__ import annotations

import ctypes
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

__all__ = [
    "PUERTO", "estado", "encender", "apagar", "reiniciar", "vigilar",
    "preferencias", "guardar_preferencias", "sembrar_si_falta", "responde",
    "direccion_tailscale",
]

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "data"
PREFERENCIAS = DATOS / "servicio.json"
REGISTRO = DATOS / "servidor.log"
REGISTRO_ERROR = DATOS / "servidor.log.err"
REGISTRO_VIGILANTE = DATOS / "vigilante.log"

PUERTO = 8000

#: Cuanto se espera a que conteste despues de levantarla. Arranca en 3-4
#: segundos, pero con la base de datos fria se ha visto tardar mas; anotar
#: «no se levanto» cuando en realidad estaba arrancando seria peor que
#: esperar.
SEGUNDOS_PARA_ARRANCAR = 30
SEGUNDOS_PARA_APAGAR = 12

#: Los registros se recortan solos. Son un registro de servicio, no un
#: historial: cuando crecen sin limite, lo que uno busca no se encuentra.
TOPE_REGISTRO = 2 * 1024 * 1024
LINEAS_QUE_SE_CONSERVAN = 2000

POR_DEFECTO: Dict[str, Any] = {
    "deseado": "apagada",
    "revivir": True,
    "al_prender": False,
    "arranque_visto": 0,
}


# --------------------------------------------------------------- preferencias

def preferencias() -> Dict[str, Any]:
    """Lo guardado, completado con los valores por defecto.

    Falla en blando: si el archivo esta roto se usan los valores por defecto
    en vez de reventar. Que no se pueda leer una preferencia no puede dejar
    sin arrancar a la aplicacion.
    """
    datos = dict(POR_DEFECTO)
    try:
        with PREFERENCIAS.open(encoding="utf-8") as fh:
            guardado = json.load(fh)
        if isinstance(guardado, dict):
            datos.update({k: v for k, v in guardado.items() if k in POR_DEFECTO})
    except (OSError, ValueError):
        pass
    return datos


def guardar_preferencias(**cambios: Any) -> Dict[str, Any]:
    datos = preferencias()
    datos.update({k: v for k, v in cambios.items() if k in POR_DEFECTO})
    try:
        DATOS.mkdir(parents=True, exist_ok=True)
        with PREFERENCIAS.open("w", encoding="utf-8") as fh:
            json.dump(datos, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
    except OSError:
        pass
    return datos


def sembrar_si_falta(puerto: int = PUERTO) -> None:
    """La primera vez, el estado deseado sale de la REALIDAD.

    Sin esto, el dia que se instala el panel el archivo nace diciendo
    «apagada» aunque la aplicacion este corriendo, y el vigilante deja de
    reponerla sin que nadie lo haya pedido. Un cambio de comportamiento
    silencioso es peor que uno visible: se adopta lo que esta pasando y de
    ahi en adelante mandan los botones.
    """
    if PREFERENCIAS.exists():
        return
    guardar_preferencias(
        deseado="encendida" if responde(puerto) else "apagada",
        arranque_visto=_arranque_del_computador())


def _arranque_del_computador() -> int:
    """Cuando se prendio el computador, en segundos desde la epoca.

    Sirve para saber si esta es la primera vuelta del vigilante despues de un
    arranque. Se redondea al minuto porque `GetTickCount64` no da exactamente
    lo mismo en dos llamadas y un valor que cambia solo haria creer que el
    computador se reinicio cada cinco minutos.
    """
    try:
        milis = ctypes.windll.kernel32.GetTickCount64()      # type: ignore
        return int((time.time() - milis / 1000.0) // 60) * 60
    except Exception:
        return 0


# ------------------------------------------------------------------- consultas

def responde(puerto: int = PUERTO, segundos: float = 4.0) -> bool:
    """Si contesta `/salud`, no solo si el puerto esta abierto.

    Un puerto abierto no significa que la aplicacion este sana, y ademas asi
    se distingue de cualquier otro programa escuchando ahi.
    """
    try:
        with urllib.request.urlopen(
                "http://127.0.0.1:{}/salud".format(puerto), timeout=segundos) as r:
            return r.status == 200
    except (urllib.error.URLError, OSError, ValueError):
        return False


def quien_tiene_el_puerto(puerto: int = PUERTO) -> Optional[Tuple[int, str]]:
    """El PID y el nombre del proceso que escucha en el puerto, si hay alguno.

    `netstat` trae una linea por familia de direcciones -IPv4 e IPv6-, asi que
    el mismo proceso aparece dos veces; se devuelve uno solo.
    """
    try:
        salida = subprocess.run(
            ["netstat", "-ano", "-p", "TCP"], capture_output=True, text=True,
            creationflags=_SIN_VENTANA).stdout
    except (OSError, subprocess.SubprocessError):
        return None
    for linea in salida.splitlines():
        if ":{} ".format(puerto) in linea and "LISTENING" in linea:
            try:
                pid = int(linea.split()[-1])
            except (ValueError, IndexError):
                continue
            return pid, _nombre_de_proceso(pid)
    return None


def _nombre_de_proceso(pid: int) -> str:
    try:
        salida = subprocess.run(
            ["tasklist", "/FI", "PID eq {}".format(pid), "/FO", "CSV", "/NH"],
            capture_output=True, text=True, creationflags=_SIN_VENTANA).stdout
        return salida.strip().split(",")[0].strip('"') or "?"
    except (OSError, subprocess.SubprocessError, IndexError):
        return "?"


def direccion_tailscale() -> Optional[str]:
    """La direccion por la que entran desde el telefono o desde otro equipo."""
    ruta = Path(r"C:\Program Files\Tailscale\tailscale.exe")
    if not ruta.exists():
        return None
    try:
        salida = subprocess.run([str(ruta), "ip", "-4"], capture_output=True,
                                text=True, timeout=6,
                                creationflags=_SIN_VENTANA).stdout
    except (OSError, subprocess.SubprocessError):
        return None
    primera = salida.strip().splitlines()
    return primera[0].strip() if primera else None


def estado(puerto: int = PUERTO) -> Dict[str, Any]:
    """Como esta el servicio AHORA, y por que.

    «Apagada» a secas no alcanza: no es lo mismo que la hayan apagado a
    proposito, que se haya caido, o que el puerto lo tenga otro programa. Las
    tres se ven igual desde fuera y necesitan respuestas distintas.
    """
    prefs = preferencias()
    viva = responde(puerto)
    ocupante = quien_tiene_el_puerto(puerto)

    if viva:
        situacion, motivo = "encendida", ""
    elif ocupante:
        situacion = "ocupado"
        motivo = ("el puerto {} lo tiene «{}» (PID {}), que no es la "
                  "aplicación.".format(puerto, ocupante[1], ocupante[0]))
    elif prefs["deseado"] == "apagada":
        situacion, motivo = "apagada", "apagada a propósito."
    else:
        situacion = "caida"
        motivo = "debería estar encendida y no responde."

    return {
        "situacion": situacion,
        "motivo": motivo,
        "encendida": viva,
        "puerto": puerto,
        "pid": ocupante[0] if ocupante else None,
        "url": "http://localhost:{}".format(puerto),
        "deseado": prefs["deseado"],
        "revivir": prefs["revivir"],
        "al_prender": prefs["al_prender"],
    }


# -------------------------------------------------------------------- acciones

#: Que no parpadee una ventana negra en cada comprobacion. Es exactamente el
#: tipo de cosa que hace que alguien desactive el vigilante sin saber que era.
_SIN_VENTANA = 0x08000000 if os.name == "nt" else 0            # CREATE_NO_WINDOW


def _python() -> str:
    """El interprete con consola.

    El panel corre bajo `pythonw.exe` para no arrastrar una ventana negra, y
    levantar uvicorn con ese mismo binario deja un proceso sin salida
    estandar: los registros quedarian vacios y la consola del panel, muda.
    """
    ejecutable = Path(sys.executable)
    if ejecutable.name.lower() == "pythonw.exe":
        con_consola = ejecutable.with_name("python.exe")
        if con_consola.exists():
            return str(con_consola)
    return str(ejecutable)


def _recortar(registro: Path) -> None:
    try:
        if not registro.exists() or registro.stat().st_size <= TOPE_REGISTRO:
            return
        with registro.open("r", encoding="utf-8", errors="replace") as fh:
            ultimas = fh.readlines()[-LINEAS_QUE_SE_CONSERVAN:]
        with registro.open("w", encoding="utf-8") as fh:
            fh.writelines(ultimas)
    except OSError:
        pass


def encender(puerto: int = PUERTO, esperar: int = SEGUNDOS_PARA_ARRANCAR,
             anotar=None) -> Tuple[bool, str]:
    """La levanta y **comprueba que quedo respondiendo**.

    Decir «la levante» sin mirar es justo el tipo de aviso que no sirve para
    nada: el proceso puede arrancar y morir a los dos segundos por un error de
    importacion, y desde fuera se ve igual.
    """
    decir = anotar or (lambda _t: None)
    if responde(puerto):
        guardar_preferencias(deseado="encendida")
        return True, "ya estaba encendida."

    ocupante = quien_tiene_el_puerto(puerto)
    if ocupante:
        return False, ("el puerto {} lo tiene «{}» (PID {}). Ciérrelo y "
                       "vuelva a intentar.".format(puerto, ocupante[1],
                                                   ocupante[0]))

    guardar_preferencias(deseado="encendida",
                         arranque_visto=_arranque_del_computador())
    DATOS.mkdir(parents=True, exist_ok=True)
    _recortar(REGISTRO)
    _recortar(REGISTRO_ERROR)

    # Se ABRE EN AGREGAR, no en truncar: la consola del panel tiene que poder
    # mostrar lo que paso antes del ultimo reinicio, que es justamente cuando
    # uno va a mirarla.
    try:
        salida = REGISTRO.open("ab")
        error = REGISTRO_ERROR.open("ab")
    except OSError as e:
        return False, "no se pudo escribir el registro: {}".format(e)

    marca = "\n===== {} · se levanta la aplicación =====\n".format(
        datetime.now().strftime("%d-%m-%Y %H:%M:%S")).encode("utf-8")
    salida.write(marca)
    salida.flush()
    error.write(marca)
    error.flush()

    try:
        # --host 0.0.0.0 es lo que permite entrar desde otro equipo por
        # Tailscale. Sin --reload a proposito: el recargador falla de forma
        # intermitente y aca nadie va a estar mirando la ventana.
        subprocess.Popen(
            [_python(), "-m", "uvicorn", "main:app",
             "--host", "0.0.0.0", "--port", str(puerto)],
            cwd=str(RAIZ), stdout=salida, stderr=error,
            stdin=subprocess.DEVNULL, creationflags=_SIN_VENTANA)
    except OSError as e:
        return False, "no se pudo lanzar: {}".format(e)
    finally:
        salida.close()
        error.close()

    decir("levantando…")
    for _ in range(max(1, esperar * 2)):
        time.sleep(0.5)
        if responde(puerto, segundos=2.0):
            return True, "encendida y respondiendo."
    return False, ("se lanzó y no responde después de {} s. El motivo está en "
                   "la consola, en rojo.".format(esperar))


def apagar(puerto: int = PUERTO, a_proposito: bool = True) -> Tuple[bool, str]:
    """La baja, y solo a ella.

    `a_proposito` es lo que distingue apagarla de reiniciarla: al reiniciar el
    estado deseado sigue siendo «encendida», asi que si algo sale mal en el
    medio el vigilante la recupera.
    """
    if a_proposito:
        guardar_preferencias(deseado="apagada")

    ocupante = quien_tiene_el_puerto(puerto)
    if not ocupante:
        return True, "ya estaba apagada."

    # No se mata nada sin confirmar que es la aplicacion. Si en el puerto hay
    # otro programa, matarlo seria un destrozo silencioso.
    if not responde(puerto):
        return False, ("el puerto {} lo tiene «{}» (PID {}) y no contesta como "
                       "la aplicación: no se toca.".format(
                           puerto, ocupante[1], ocupante[0]))

    pid = ocupante[0]
    try:
        # /T se lleva tambien los procesos hijos. uvicorn sin --reload no crea
        # ninguno, pero con el si, y un hijo huerfano deja el puerto tomado.
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                       capture_output=True, creationflags=_SIN_VENTANA)
    except (OSError, subprocess.SubprocessError) as e:
        return False, "no se pudo detener el proceso {}: {}".format(pid, e)

    for _ in range(SEGUNDOS_PARA_APAGAR * 2):
        time.sleep(0.5)
        if not responde(puerto, segundos=1.5):
            return True, "apagada."
    return False, "se pidió detenerla y sigue respondiendo."


def reiniciar(puerto: int = PUERTO, anotar=None) -> Tuple[bool, str]:
    ok, mensaje = apagar(puerto, a_proposito=False)
    if not ok:
        return False, mensaje
    return encender(puerto, anotar=anotar)


# ------------------------------------------------------------------ vigilancia

def _anotar_vigilante(texto: str) -> None:
    linea = "{}  {}".format(datetime.now().strftime("%d-%m-%Y %H:%M:%S"), texto)
    try:
        DATOS.mkdir(parents=True, exist_ok=True)
        _recortar(REGISTRO_VIGILANTE)
        with REGISTRO_VIGILANTE.open("a", encoding="utf-8") as fh:
            fh.write(linea + "\n")
    except OSError:
        pass
    # Bajo `pythonw.exe` -que es como lo corre la tarea, para que no aparezca
    # ninguna ventana- NO HAY salida estandar: `sys.stdout` es None y un
    # `print` revienta con AttributeError. El vigilante se caeria entero por
    # intentar escribir un mensaje que nadie iba a leer.
    if sys.stdout is not None:
        try:
            print(linea)
        except (OSError, ValueError):
            pass


def vigilar(puerto: int = PUERTO) -> int:
    """Lo que corre la tarea de Windows cada cinco minutos.

    Solo repone lo que se cayo solo. **No levanta nada que alguien haya
    apagado**, y despues de prender el computador no levanta nada salvo que
    esa preferencia este puesta: era lo que molestaba a el usuario, que la pagina
    apareciera sola sin haberla pedido.
    """
    sembrar_si_falta(puerto)
    prefs = preferencias()
    arranque = _arranque_del_computador()

    # Primera vuelta despues de prender el computador.
    if arranque and prefs["arranque_visto"] != arranque:
        guardar_preferencias(arranque_visto=arranque)
        if not prefs["al_prender"] and not responde(puerto):
            guardar_preferencias(deseado="apagada")
            _anotar_vigilante("el computador arrancó. No se levanta sola: "
                              "esa preferencia está apagada.")
            return 0
        prefs = preferencias()

    if responde(puerto):
        return 0            # todo bien: ni se anota, para no llenar el registro

    if prefs["deseado"] == "apagada":
        return 0            # la apagaron a proposito

    if not prefs["revivir"]:
        _anotar_vigilante("no responde, pero «volver a levantarla si se cae» "
                          "está apagado.")
        return 1

    ocupante = quien_tiene_el_puerto(puerto)
    if ocupante:
        _anotar_vigilante(
            "NO se levanta: el puerto {} lo tiene «{}» (PID {}), que no es la "
            "aplicación.".format(puerto, ocupante[1], ocupante[0]))
        return 1

    _anotar_vigilante("se había caído: levantándola...")
    ok, mensaje = encender(puerto)
    _anotar_vigilante(mensaje if ok else "NO SE PUDO LEVANTAR: " + mensaje)
    return 0 if ok else 1


# --------------------------------------------------------------- los registros

#: `INFO:     127.0.0.1:55687 - "GET /expedientes HTTP/1.1" 200 OK`
_ACCESO = re.compile(
    r'^INFO:\s+(?P<origen>\S+) - "(?P<metodo>[A-Z]+) (?P<ruta>\S+) [^"]*"'
    r'\s+(?P<codigo>\d{3})')


def leer_linea_de_acceso(linea: str) -> Optional[Dict[str, Any]]:
    """Descompone una linea de acceso de uvicorn, o devuelve None.

    **No trae hora.** El formato de uvicorn no la incluye, y ponerle la hora a
    la que uno la leyo seria inventar un dato: para las lineas viejas estaria
    mal. El panel solo le pone hora a las que llegan con el abierto.
    """
    hallado = _ACCESO.match(linea.strip())
    if not hallado:
        return None
    codigo = int(hallado.group("codigo"))
    return {
        "metodo": hallado.group("metodo"),
        "ruta": hallado.group("ruta"),
        "codigo": codigo,
        "origen": hallado.group("origen"),
        "clase": ("error" if codigo >= 500 else
                  "aviso" if codigo >= 400 else "bien"),
    }


def resumen_de_accesos(lineas: List[str]) -> Dict[str, int]:
    """Cuantas solicitudes hubo y cuantas fallaron."""
    total = fallidas = servidor = 0
    for linea in lineas:
        dato = leer_linea_de_acceso(linea)
        if not dato:
            continue
        total += 1
        if dato["codigo"] >= 400:
            fallidas += 1
        if dato["codigo"] >= 500:
            servidor += 1
    return {"solicitudes": total, "con_error": fallidas,
            "errores_del_servidor": servidor}


if __name__ == "__main__":
    sys.exit(vigilar())
