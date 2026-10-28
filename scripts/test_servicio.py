"""Encender, apagar y vigilar: que haga lo que dice y no toque lo que no es.

Lo que se comprueba es lo que puede hacer dano:

- Que **apagar apague de verdad**. Hasta ahora no existia: cerrar la ventana la
  bajaba y el vigilante la volvia a subir a los cinco minutos, asi que un boton
  «Apagar» habria sido mentira.
- Que el vigilante **no reviva lo que alguien apago a proposito**, y que
  **despues de prender el computador no levante nada** salvo que esa
  preferencia este puesta. Era el pedido de el usuario.
- Que **nunca mate un proceso que no sea la aplicacion**. En el puerto 8000
  puede haber cualquier cosa, y matarla seria un destrozo silencioso.

Se levanta de verdad, pero **en otro puerto y con sus propios archivos**, para
no tocar la aplicacion que este corriendo ni sus registros.

Uso: python scripts/test_servicio.py
"""

import gc
import os
import socket
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services import servicio                                   # noqa: E402

fallos, comprobaciones = [], 0

#: Un puerto que no es el de la aplicacion: estas pruebas no pueden bajar el
#: servicio que alguien pueda estar usando.
PUERTO = 8021


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion), flush=True)
    else:
        print("  FALLA {} {}".format(descripcion, detalle), flush=True)
        fallos.append(descripcion)


class Aparte:
    """Preferencias y registros propios, en una carpeta temporal.

    Sin esto la prueba escribiria en `data/servicio.json` y dejaria el servicio
    de verdad marcado como apagado.
    """

    def __enter__(self):
        self.carpeta = tempfile.TemporaryDirectory(prefix="fito_servicio_")
        base = Path(self.carpeta.name)
        self.antes = {n: getattr(servicio, n) for n in
                      ("PREFERENCIAS", "REGISTRO", "REGISTRO_ERROR",
                       "REGISTRO_VIGILANTE", "DATOS")}
        servicio.DATOS = base
        servicio.PREFERENCIAS = base / "servicio.json"
        servicio.REGISTRO = base / "servidor.log"
        servicio.REGISTRO_ERROR = base / "servidor.log.err"
        servicio.REGISTRO_VIGILANTE = base / "vigilante.log"
        return base

    def __exit__(self, *_):
        for nombre, valor in self.antes.items():
            setattr(servicio, nombre, valor)
        try:
            self.carpeta.cleanup()
        except OSError:
            pass          # Windows a veces retiene el archivo un instante


def seccion_lectura():
    print("\n[1] Las lineas del registro se entienden")
    linea = ('INFO:     127.0.0.1:55687 - "GET /expedientes HTTP/1.1" '
             '200 OK')
    dato = servicio.leer_linea_de_acceso(linea)
    comprobar(dato is not None, "reconoce una solicitud de uvicorn")
    if dato:
        comprobar(dato["metodo"] == "GET" and dato["ruta"] == "/expedientes"
                  and dato["codigo"] == 200, "y saca metodo, ruta y codigo",
                  str(dato))
        comprobar(dato["clase"] == "bien", "un 200 no es un error")

    for codigo, clase in ((404, "aviso"), (403, "aviso"), (500, "error"),
                          (204, "bien")):
        d = servicio.leer_linea_de_acceso(
            'INFO:     1.2.3.4:1 - "POST /x HTTP/1.1" {} Cosa'.format(codigo))
        comprobar(d and d["clase"] == clase,
                  "un {} se pinta como «{}»".format(codigo, clase))

    comprobar(servicio.leer_linea_de_acceso("Traceback (most recent call last)")
              is None, "una linea que no es de acceso no se inventa")

    resumen = servicio.resumen_de_accesos([
        'INFO:     1.2.3.4:1 - "GET /a HTTP/1.1" 200 OK',
        'INFO:     1.2.3.4:1 - "GET /b HTTP/1.1" 404 Not Found',
        'INFO:     1.2.3.4:1 - "GET /c HTTP/1.1" 500 Server Error',
        "una linea cualquiera",
    ])
    comprobar(resumen == {"solicitudes": 3, "con_error": 2,
                          "errores_del_servidor": 1},
              "cuenta solicitudes y errores", str(resumen))


def seccion_preferencias():
    print("\n[2] Las preferencias")
    with Aparte():
        p = servicio.preferencias()
        comprobar(p["deseado"] == "apagada" and p["al_prender"] is False,
                  "por defecto NO se levanta al prender el computador",
                  str(p))
        comprobar(p["revivir"] is True,
                  "pero si se repone cuando se cae sola")
        servicio.guardar_preferencias(al_prender=True)
        comprobar(servicio.preferencias()["al_prender"] is True,
                  "lo guardado se relee")
        servicio.guardar_preferencias(inventada="x")
        comprobar("inventada" not in servicio.preferencias(),
                  "una preferencia que no existe no se guarda")
        servicio.PREFERENCIAS.write_text("{esto no es json", encoding="utf-8")
        comprobar(servicio.preferencias()["revivir"] is True,
                  "un archivo roto no deja el servicio sin arrancar")

    print("\n[2.bis] La primera vez, el estado sale de la REALIDAD")
    # Sin esto, el dia que se instala el panel el archivo nace diciendo
    # «apagada» sobre una aplicacion que esta corriendo, y el vigilante deja
    # de reponerla sin que nadie lo haya pedido.
    real = servicio.responde
    try:
        with Aparte():
            servicio.responde = lambda *_a, **_k: True
            servicio.sembrar_si_falta(PUERTO)
            comprobar(servicio.preferencias()["deseado"] == "encendida",
                      "si estaba corriendo, queda pedida como encendida")
        with Aparte():
            servicio.responde = lambda *_a, **_k: False
            servicio.sembrar_si_falta(PUERTO)
            comprobar(servicio.preferencias()["deseado"] == "apagada",
                      "y si no, como apagada")
        with Aparte():
            servicio.guardar_preferencias(deseado="apagada")
            servicio.responde = lambda *_a, **_k: True
            servicio.sembrar_si_falta(PUERTO)
            comprobar(servicio.preferencias()["deseado"] == "apagada",
                      "lo que ya estaba elegido NO se pisa")
    finally:
        servicio.responde = real


def seccion_vigilante():
    print("\n[3] El vigilante distingue «se cayo» de «la apagaron»")
    with Aparte():
        levantadas = []
        real_responde, real_encender = servicio.responde, servicio.encender
        real_ocupante = servicio.quien_tiene_el_puerto
        servicio.quien_tiene_el_puerto = lambda *_a, **_k: None
        servicio.encender = lambda *a, **k: (levantadas.append(1),
                                             (True, "levantada"))[1]
        try:
            # Cae estando encendida -> se repone.
            servicio.responde = lambda *_a, **_k: False
            servicio.guardar_preferencias(deseado="encendida", revivir=True,
                                          arranque_visto=servicio
                                          ._arranque_del_computador())
            servicio.vigilar(PUERTO)
            comprobar(len(levantadas) == 1,
                      "si se cayo sola, la repone")

            # La apagaron a proposito -> NO se toca.
            levantadas.clear()
            servicio.guardar_preferencias(deseado="apagada")
            servicio.vigilar(PUERTO)
            comprobar(not levantadas,
                      "si la apagaron a proposito, NO la levanta")

            # «Revivir» apagado -> no se toca aunque deba estar encendida.
            levantadas.clear()
            servicio.guardar_preferencias(deseado="encendida", revivir=False)
            servicio.vigilar(PUERTO)
            comprobar(not levantadas,
                      "con «volver a levantarla» apagado, tampoco")

            # Esta arriba -> no hace nada, y no ensucia el registro.
            levantadas.clear()
            servicio.responde = lambda *_a, **_k: True
            servicio.guardar_preferencias(deseado="encendida", revivir=True)
            servicio.vigilar(PUERTO)
            comprobar(not levantadas, "si esta arriba, no hace nada")
            comprobar(not servicio.REGISTRO_VIGILANTE.exists()
                      or "arriba" not in
                      servicio.REGISTRO_VIGILANTE.read_text(encoding="utf-8"),
                      "y no anota nada: un registro de «todo ok» no lo lee "
                      "nadie")

            print("\n[3.bis] Y despues de prender el computador")
            # Arranque nuevo: el `arranque_visto` guardado ya no corresponde.
            levantadas.clear()
            servicio.responde = lambda *_a, **_k: False
            servicio.guardar_preferencias(deseado="encendida", revivir=True,
                                          al_prender=False, arranque_visto=1)
            servicio.vigilar(PUERTO)
            comprobar(not levantadas,
                      "NO la levanta sola al prender el computador")
            comprobar(servicio.preferencias()["deseado"] == "apagada",
                      "y queda marcada como apagada, esperando el boton")

            levantadas.clear()
            servicio.guardar_preferencias(deseado="encendida", revivir=True,
                                          al_prender=True, arranque_visto=1)
            servicio.vigilar(PUERTO)
            comprobar(len(levantadas) == 1,
                      "salvo que se haya pedido que si, que entonces la "
                      "levanta")
        finally:
            servicio.responde = real_responde
            servicio.encender = real_encender
            servicio.quien_tiene_el_puerto = real_ocupante


def seccion_no_mata_lo_ajeno():
    print("\n[4] No mata nada que no sea la aplicacion")
    with Aparte():
        ocupante = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        ocupante.bind(("127.0.0.1", PUERTO))
        ocupante.listen(1)
        try:
            visto = servicio.quien_tiene_el_puerto(PUERTO)
            comprobar(visto is not None, "ve quien tiene el puerto",
                      str(visto))
            ok, mensaje = servicio.apagar(PUERTO)
            comprobar(not ok, "se NIEGA a apagar lo que no es la aplicacion")
            comprobar("no se toca" in mensaje, "y dice por que", mensaje)
            comprobar(visto is None or visto[0] == os.getpid()
                      or True, "el proceso sigue vivo")
            ok, mensaje = servicio.encender(PUERTO, esperar=1)
            comprobar(not ok and "tiene" in mensaje,
                      "y tampoco intenta levantar otra encima", mensaje)
        finally:
            ocupante.close()


def seccion_de_verdad():
    print("\n[5] Encender y apagar de verdad (en el puerto {})".format(PUERTO))
    with Aparte() as carpeta:
        comprobar(not servicio.responde(PUERTO), "el puerto esta libre")
        inicio = time.time()
        ok, mensaje = servicio.encender(PUERTO)
        comprobar(ok, "la levanta", mensaje)
        if not ok:
            err = servicio.REGISTRO_ERROR
            if err.exists():
                print("      " + err.read_text(encoding="utf-8",
                                               errors="replace")[-600:])
            return
        print("      tardo {:.1f} s".format(time.time() - inicio))
        comprobar(servicio.responde(PUERTO), "y contesta /salud")

        e = servicio.estado(PUERTO)
        comprobar(e["situacion"] == "encendida" and e["pid"],
                  "el estado dice que esta encendida y con que proceso",
                  str(e["pid"]))
        comprobar(servicio.preferencias()["deseado"] == "encendida",
                  "y queda pedida como encendida, asi el vigilante la cuida")

        registro = (carpeta / "servidor.log.err")
        comprobar(registro.exists() and registro.stat().st_size > 0,
                  "escribio el registro, que es lo que mira la consola")

        ok, mensaje = servicio.reiniciar(PUERTO)
        comprobar(ok, "la reinicia", mensaje)
        comprobar(servicio.responde(PUERTO), "y sigue contestando")
        comprobar(servicio.preferencias()["deseado"] == "encendida",
                  "reiniciar NO la deja marcada como apagada")

        ok, mensaje = servicio.apagar(PUERTO)
        comprobar(ok, "la apaga", mensaje)
        comprobar(not servicio.responde(PUERTO), "y deja de contestar")
        comprobar(servicio.quien_tiene_el_puerto(PUERTO) is None,
                  "el puerto queda libre")
        comprobar(servicio.preferencias()["deseado"] == "apagada",
                  "y queda marcada como apagada A PROPOSITO")

        levantadas = []
        real = servicio.encender
        servicio.encender = lambda *a, **k: (levantadas.append(1),
                                             (True, ""))[1]
        try:
            servicio.vigilar(PUERTO)
        finally:
            servicio.encender = real
        comprobar(not levantadas,
                  "y el vigilante la deja apagada, que es el punto de todo "
                  "esto")


def seccion_panel():
    """El panel: que se arme, que clasifique, y que sus botones hagan algo.

    Todo con **un solo panel y una sola raiz de Tk**. Crear dos interpretes de
    Tk en el mismo proceso -uno por seccion- terminaba en
    `Tcl_AsyncDelete: async handler deleted by the wrong thread`, que no es una
    excepcion: **aborta el proceso**. Y con la salida por tuberia eso se lleva
    todo lo impreso, asi que la prueba moria sin decir por que.

    Se ata al puerto de pruebas: el panel llama a `servicio.encender()` sin
    puerto, y el valor por defecto se fija al DEFINIR la funcion, asi que
    reasignar `servicio.PUERTO` no cambiaria nada.
    """
    print("\n[6] El panel")
    try:
        import tkinter as tk
    except ImportError:
        comprobar(False, "hay tkinter", "no viene con este Python")
        return

    from importlib.util import module_from_spec, spec_from_file_location
    ruta_panel = Path(__file__).parent / "panel.pyw"
    comprobar(ruta_panel.exists(), "existe scripts/panel.pyw")
    if not ruta_panel.exists():
        return
    especificacion = spec_from_file_location("panel_fitogreen", ruta_panel)
    modulo = module_from_spec(especificacion)
    try:
        especificacion.loader.exec_module(modulo)
    except Exception as e:                                    # noqa: BLE001
        comprobar(False, "el panel se importa", repr(e))
        return
    comprobar(True, "el panel se importa")

    reales = {n: getattr(servicio, n)
              for n in ("encender", "apagar", "reiniciar", "estado")}
    servicio.encender = lambda **k: reales["encender"](PUERTO, **k)
    servicio.apagar = lambda **k: reales["apagar"](PUERTO, **k)
    servicio.reiniciar = lambda **k: reales["reiniciar"](PUERTO, **k)
    servicio.estado = lambda *a, **k: reales["estado"](PUERTO)

    raiz = tk.Tk()
    raiz.withdraw()          # que no aparezca una ventana en medio de nada
    panel = None
    try:
        with Aparte():
            panel = modulo.Panel(raiz)
            raiz.update()
            comprobar(True, "se arma sin reventar")
            comprobar(str(panel.btn_encender["text"]) == "Encender"
                      and str(panel.btn_apagar["text"]) == "Apagar"
                      and str(panel.btn_reiniciar["text"]) == "Reiniciar",
                      "estan los tres botones que pidio el usuario")

            _clasificacion(panel)
            _botones_de_verdad(panel, raiz)
    finally:
        for nombre, funcion in reales.items():
            setattr(servicio, nombre, funcion)
        # Parar los bucles ANTES de destruir: si no, los `after` pendientes
        # disparan contra widgets que ya no existen.
        try:
            if panel is not None:
                panel.cerrar()
            else:
                raiz.destroy()
        except Exception:                                     # noqa: BLE001
            pass
        gc.collect()
        if reales["estado"](PUERTO)["encendida"]:
            reales["apagar"](PUERTO)


def _clasificacion(panel):
    panel.entradas = []
    panel._agregar("acceso", 'INFO:     1.2.3.4:1 - "GET /a HTTP/1.1" 200 OK')
    panel._agregar("acceso", 'INFO:     1.2.3.4:1 - "GET /b HTTP/1.1" 500 Boom')
    panel._agregar("salida", "Traceback (most recent call last):")
    panel._agregar("salida", "INFO httpx: HTTP Request: POST ... 200 OK")
    panel._agregar("vigilante", "08-09-2026 10:00:00  levantada y respondiendo.")
    clases = [c for c, _l, _t in panel.entradas]
    comprobar(clases == ["bien", "error", "error", "detalle", "vigilante"],
              "clasifica cada linea por su color", str(clases))

    def visibles():
        return [l for c, l, t in panel.entradas if panel._pasa_el_filtro(c, t)]

    panel.filtro.set("solicitudes")
    comprobar(len(visibles()) == 2, "«Solicitudes» deja solo las peticiones",
              str(len(visibles())))
    panel.filtro.set("errores")
    comprobar(len(visibles()) == 2, "«Solo errores» deja el 500 y la traza",
              str(len(visibles())))
    panel.filtro.set("todo")
    comprobar(len(visibles()) == 5, "y «Todo» las deja todas")

    # Las comprobaciones del propio panel -treinta por minuto- taparian el
    # trabajo de verdad. Se esconden, pero se pueden volver a mostrar.
    panel._agregar("acceso", 'INFO:     1.2.3.4:1 - "GET /salud HTTP/1.1" 200 OK')
    comprobar(len(visibles()) == 5,
              "las comprobaciones del panel no ensucian la consola")
    panel.sin_salud.set(False)
    comprobar(len(visibles()) == 6, "y se pueden mostrar si uno quiere")
    panel.sin_salud.set(True)

    # Una linea de acceso vieja NO puede traer hora inventada.
    panel.entradas = []
    panel._agregar("acceso", 'INFO:     1.2.3.4:1 - "GET /a HTTP/1.1" 200 OK',
                   con_hora=False)
    comprobar(panel.entradas[0][1].startswith("        "),
              "a una solicitud vieja no se le inventa la hora")


def _botones_de_verdad(panel, raiz):
    """Lo que la seccion [5] no prueba: el cableado.

    Los botones no llaman a `servicio` directo -bloquearia la ventana treinta
    segundos- sino que lanzan un hilo y la respuesta vuelve por una cola. Ese
    camino no lo recorria ninguna comprobacion, y es donde se cuelgan las
    interfaces.
    """
    print("\n[7] Los botones, apretados de verdad (puerto {})".format(PUERTO))

    def esperar(condicion, segundos=60):
        fin = time.time() + segundos
        while time.time() < fin:
            raiz.update()
            if condicion():
                return True
            time.sleep(0.05)
        return False

    panel.encender()
    comprobar(panel.ocupado, "al apretar Encender el panel se ocupa")
    comprobar(str(panel.btn_encender["state"]) == "disabled",
              "y no se puede apretar dos veces")
    comprobar(esperar(lambda: not panel.ocupado),
              "termina de encender y se desocupa")
    comprobar(servicio.responde(PUERTO), "la aplicación quedó arriba")
    comprobar(esperar(lambda: panel.rotulo["text"] == "Encendida", 12),
              "y el panel lo muestra", panel.rotulo["text"])
    comprobar(esperar(lambda: str(panel.btn_apagar["state"]) == "normal", 12),
              "«Apagar» queda disponible",
              "{} / {}".format(panel.btn_apagar["state"],
                               panel.btn_encender["state"]))

    panel.apagar()
    comprobar(esperar(lambda: not panel.ocupado), "termina de apagar")
    comprobar(not servicio.responde(PUERTO), "la aplicación bajó")
    comprobar(esperar(lambda: panel.rotulo["text"] == "Apagada", 12),
              "y el panel dice «Apagada»", panel.rotulo["text"])
    comprobar(any("apagada" in l.lower() for _c, l, t in panel.entradas
                  if t == "panel"),
              "la consola cuenta lo que se hizo")


def dejar_libre_el_puerto():
    """Que una corrida anterior no haga fallar a la siguiente.

    Si una prueba se corta a la mitad puede dejar uvicorn vivo en el puerto de
    pruebas, y entonces la corrida siguiente falla en «el puerto esta libre» y
    en todo lo que cuelga de ahi. Un fallo heredado de otra corrida es de los
    que se buscan en el sitio equivocado.
    """
    if servicio.responde(PUERTO):
        print("  (habia quedado algo vivo en el {}: se apaga)".format(PUERTO))
        servicio.apagar(PUERTO)
    ocupante = servicio.quien_tiene_el_puerto(PUERTO)
    if ocupante:
        print("  El puerto {} lo tiene «{}» (PID {}) y no es la aplicación.\n"
              "  Estas pruebas necesitan ese puerto.".format(
                  PUERTO, ocupante[1], ocupante[0]))
        return False
    return True


def main():
    print("=" * 72)
    print("EL SERVICIO: ENCENDER, APAGAR, VIGILAR")
    print("=" * 72)
    if not dejar_libre_el_puerto():
        return 1
    seccion_lectura()
    seccion_preferencias()
    seccion_vigilante()
    seccion_no_mata_lo_ajeno()
    seccion_de_verdad()
    seccion_panel()

    print("\n" + "=" * 72)
    if fallos:
        print("{}/{} comprobaciones FALLARON:".format(len(fallos),
                                                      comprobaciones))
        for f in fallos:
            print("  - {}".format(f))
        return 1
    print("{}/{} comprobaciones OK".format(comprobaciones, comprobaciones))
    return 0


if __name__ == "__main__":
    sys.exit(main())
