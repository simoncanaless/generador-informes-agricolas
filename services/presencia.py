"""Quien esta mirando cada informe ahora mismo.

Vive en memoria a proposito. Es un dato que caduca en segundos y que no tiene
ningun sentido conservar: si el servidor se reinicia, nadie esta trabajando en
nada hasta que vuelva a avisar. Guardarlo en PostgreSQL seria escribir en disco
tres veces por segundo para nada.

Funciona porque el servidor es UN proceso. El dia que se levante con varios
`--workers` cada uno tendria su propia lista y la presencia se veria a medias:
ahi habria que sacarla a un sitio compartido. Queda dicho aqui para que no
sorprenda.
"""

from __future__ import annotations

import threading
import time
from typing import Any, Dict, List, Optional

__all__ = ["marcar", "quienes_estan", "olvidar"]

#: Sin senal en este tiempo, se da por ido. El navegador avisa cada 5 s, asi
#: que 15 aguanta un par de avisos perdidos sin sacar a nadie que sigue ahi.
SEGUNDOS_VIVO = 15.0

_cerrojo = threading.Lock()
#: informe -> usuario -> {nombre, seccion, visto}
_gente: Dict[str, Dict[str, Dict[str, Any]]] = {}


def marcar(
    id_expediente: str,
    usuario_id: str,
    nombre: str,
    seccion: Optional[str] = None,
) -> None:
    """Avisa que esta persona sigue en este informe."""
    with _cerrojo:
        sala = _gente.setdefault(id_expediente, {})
        sala[usuario_id] = {
            "nombre": nombre,
            "seccion": seccion,
            "visto": time.time(),
        }


def quienes_estan(id_expediente: str, excepto: Optional[str] = None
                  ) -> List[Dict[str, Any]]:
    """Los demas que estan en este informe, del mas reciente al mas antiguo.

    De paso limpia a los que dejaron de avisar: no hace falta una tarea aparte
    para eso, y asi la lista nunca crece sin control.
    """
    ahora = time.time()
    with _cerrojo:
        sala = _gente.get(id_expediente, {})
        for uid in [u for u, d in sala.items()
                    if ahora - d["visto"] > SEGUNDOS_VIVO]:
            sala.pop(uid, None)
        if not sala:
            _gente.pop(id_expediente, None)
        return sorted(
            ({"usuario_id": uid, "nombre": d["nombre"], "seccion": d["seccion"],
              "hace_segundos": round(ahora - d["visto"], 1)}
             for uid, d in sala.items() if uid != excepto),
            key=lambda x: x["hace_segundos"])


def olvidar(id_expediente: str, usuario_id: str) -> None:
    """Se fue por su cuenta: cerro la pestana o cambio de informe."""
    with _cerrojo:
        sala = _gente.get(id_expediente)
        if sala:
            sala.pop(usuario_id, None)
            if not sala:
                _gente.pop(id_expediente, None)
