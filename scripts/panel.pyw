"""Panel de control del servicio: encender, apagar, reiniciar y mirar.

Lo pidio el usuario: *«me gustaria una interfaz en el pc que yo apriete un boton
para prenderlo, para apagarlo y uno para reiniciarlo, tambien poder ver la
consola con el tema de los errores, las solicitudes que se lograron y todo el
detalle de lo que esta pasando»*.

    pythonw scripts\\panel.pyw          (o el icono «Panel Fitogreen»)

**La logica no vive aca.** Encender, apagar y vigilar estan en
`services/servicio.py`, que es tambien lo que corre la tarea de Windows: si el
panel tuviera su propia forma de levantar la aplicacion, tarde o temprano
serian dos formas distintas y la que fallaria es la que nadie mira.

## Por que hay tres registros y no uno

- `data/servidor.log` — las solicitudes que atendio uvicorn. **No traen hora**:
  el formato de uvicorn no la incluye. Se les pone la hora a las que llegan con
  el panel abierto, y a las viejas no: ponerles la hora de lectura seria
  inventar un dato.
- `data/servidor.log.err` — todo lo demas que escribe el proceso: el arranque,
  los avisos de la IA y **los errores con su traza**.
- `data/vigilante.log` — cuando se cayo y cuando la repusieron.

Se muestran juntos, cada uno con su color, porque para saber que paso hay que
poder ver la solicitud que fallo AL LADO del error que la hizo fallar.

**Solo tkinter**, que viene con Python: el panel tiene que abrir aunque la
aplicacion no arranque, y depender de algo instalado con pip seria justamente
lo que falla el dia que hace falta.
"""

import os
import queue
import sys
import threading
import tkinter as tk
import webbrowser
from datetime import datetime
from tkinter import font as tkfont
from tkinter import ttk

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services import servicio                                   # noqa: E402

# Los colores son los de la pantalla (§16.a.vicies.bis): el panel y la
# aplicacion son la misma cosa vista de dos lados.
FONDO = "#0c0f11"
PAPEL = "#1b2124"
HUECO = "#080a0c"
TINTA = "#e7ecea"
GRIS = "#94a19c"
BORDE = "#242b2d"
BORDE_FUERTE = "#333c3e"
LIMA = "#a6ce39"
VERDE_BOTON = "#367019"
VERDE_BOTON_HOVER = "#428424"
ALERTA = "#f0857a"
AVISO = "#e9c95a"

#: Cuantas lineas se conservan en la vista. Con mas, cambiar de filtro empieza
#: a notarse y ademas nadie va a leer dos mil lineas hacia arriba: para eso
#: estan los archivos.
TOPE_DE_VISTA = 1500

#: Cada cuanto se pregunta como esta. Dos segundos es lo que tarda en notarse
#: un cambio sin que la consulta -netstat mas una peticion- se note en la
#: maquina.
CADA_CUANTO_MIRA = 2000
CADA_CUANTO_LEE = 700


class Panel:

    def __init__(self, raiz):
        self.raiz = raiz
        self.mensajes = queue.Queue()
        self.entradas = []            # (clase, texto) para poder refiltrar
        self.posiciones = {}          # por archivo: hasta donde se leyo
        self.restos = {}              # media linea que quedo sin \n
        self.filtro = tk.StringVar(value="todo")
        self.seguir = tk.BooleanVar(value=True)
        # El propio panel pregunta `/salud` cada dos segundos, o sea treinta
        # solicitudes por minuto que NO son de nadie usando el programa. Sin
        # esconderlas, la consola es una lista de comprobaciones del panel con
        # el trabajo de verdad perdido en el medio. Se esconden por defecto,
        # pero con una casilla a la vista: esconder datos sin decirlo es peor.
        self.sin_salud = tk.BooleanVar(value=True)
        self.revivir = tk.BooleanVar()
        self.al_prender = tk.BooleanVar()
        self.ocupado = False
        self.direccion_fuera = None
        # Los tres bucles se reprograman solos. Sin apuntarlos no hay forma de
        # pararlos, y al cerrar la ventana siguen disparando contra widgets
        # que ya no existen: una lluvia de «invalid command name» al salir.
        self.vivo = True
        self.tareas = []

        raiz.title("Fitogreen · control del servicio")
        raiz.configure(bg=FONDO)
        raiz.geometry("920x660")
        raiz.minsize(700, 480)
        # El emblema de la casa en la barra de tareas. Sin esto sale la pluma
        # de Python, que no dice de que programa es. Si falta el archivo, se
        # sigue: un icono no puede impedir abrir el panel.
        try:
            raiz.iconbitmap(str(servicio.RAIZ / "web" / "fitogreen.ico"))
        except tk.TclError:
            pass

        self.titulos = tkfont.Font(family="Segoe UI", size=11)
        self.grandes = tkfont.Font(family="Segoe UI", size=17, weight="bold")
        self.chicos = tkfont.Font(family="Segoe UI", size=9)
        self.consola = tkfont.Font(family="Consolas", size=9)

        # Las barras de desplazamiento las dibuja Windows con SU tema: salian
        # BLANCAS al costado de una consola negra, que es lo primero que se ve
        # al abrir. El tema «clam» es el unico que deja pintarlas entero, y
        # solo afecta a los widgets de ttk: el resto del panel es tk normal.
        estilo = ttk.Style()
        estilo.theme_use("clam")
        for orientacion in ("Vertical", "Horizontal"):
            estilo.configure(
                "Fito.{}.TScrollbar".format(orientacion),
                background="#2b3438", troughcolor=HUECO, bordercolor=HUECO,
                arrowcolor=GRIS, darkcolor="#2b3438", lightcolor="#2b3438",
                relief="flat")
            estilo.map("Fito.{}.TScrollbar".format(orientacion),
                       background=[("active", BORDE_FUERTE)])

        self._armar_cabecera()
        self._armar_botones()
        # El pie ANTES que la consola, y anclado abajo: `pack` reparte por
        # orden, asi que con la consola quedandose con todo lo que sobra el
        # contador terminaba cortado contra el borde de la ventana.
        self._armar_pie()
        self._armar_consola()

        self._cargar_preferencias()
        self._cargar_registros_previos()

        raiz.protocol("WM_DELETE_WINDOW", self.cerrar)
        self._programar(120, self._drenar)
        self._mirar_estado()
        self._programar(CADA_CUANTO_LEE, self._leer_registros)
        threading.Thread(target=self._buscar_direccion, daemon=True).start()

    def _buscar_direccion(self):
        self.direccion_fuera = servicio.direccion_tailscale()

    def _programar(self, milis, funcion):
        if self.vivo:
            self.tareas.append(self.raiz.after(milis, funcion))

    def cerrar(self):
        """Parar los bucles ANTES de destruir la ventana."""
        self.vivo = False
        for tarea in self.tareas:
            try:
                self.raiz.after_cancel(tarea)
            except tk.TclError:
                pass
        self.tareas.clear()
        try:
            self.raiz.destroy()
        except tk.TclError:
            pass
        # Y se sueltan las variables de Tk AQUI, con el hilo principal todavia
        # vivo. Si quedan colgando, el recolector las destruye al apagarse el
        # interprete -cuando Tk ya no esta- y eso en Windows termina en
        # `Tcl_AsyncDelete`, que aborta el proceso: nada de lo que se hubiera
        # impreso llega a la pantalla.
        self.filtro = self.seguir = self.revivir = self.al_prender = None

    # ---------------------------------------------------------------- armado

    def _armar_cabecera(self):
        caja = tk.Frame(self.raiz, bg=PAPEL, highlightbackground=BORDE,
                        highlightthickness=1)
        caja.pack(fill="x", padx=14, pady=(14, 10))
        dentro = tk.Frame(caja, bg=PAPEL)
        dentro.pack(fill="x", padx=18, pady=14)

        linea = tk.Frame(dentro, bg=PAPEL)
        linea.pack(fill="x")
        self.punto = tk.Canvas(linea, width=14, height=14, bg=PAPEL,
                               highlightthickness=0)
        self.punto.pack(side="left", pady=(6, 0))
        self.bola = self.punto.create_oval(2, 2, 12, 12, fill=GRIS, outline="")
        self.rotulo = tk.Label(linea, text="Comprobando…", font=self.grandes,
                               bg=PAPEL, fg=TINTA)
        self.rotulo.pack(side="left", padx=(9, 0))

        # El motivo se muestra SOLO cuando hay algo que explicar. Una etiqueta
        # vacia igual ocupa su linea, y dejaba un hueco de cuarenta pixeles en
        # medio de la tarjeta que se leia como algo sin terminar.
        self.motivo = tk.Label(dentro, text="", font=self.titulos, bg=PAPEL,
                               fg=GRIS, anchor="w", justify="left")

        self.direcciones = tk.Label(dentro, text="", font=self.chicos,
                                    bg=PAPEL, fg=GRIS, anchor="w")
        self.direcciones.pack(fill="x", pady=(8, 0))

    def _boton(self, padre, texto, orden, principal=False):
        fondo = VERDE_BOTON if principal else "#1e2528"
        encima = VERDE_BOTON_HOVER if principal else "#2a3336"
        color = "#ffffff" if principal else TINTA
        b = tk.Button(padre, text=texto, command=orden, font=self.titulos,
                      bg=fondo, fg=color, activebackground=encima,
                      activeforeground=color, relief="flat", bd=0,
                      padx=18, pady=9, cursor="hand2",
                      highlightthickness=1, highlightbackground=BORDE_FUERTE,
                      # Apagado tiene que SEGUIR LEYENDOSE. Con el gris de
                      # antes el boton desaparecia contra su propio fondo, que
                      # es el «boton fantasma» de §16.a.sexies: quien mira no
                      # sabe si esta deshabilitado o si no existe.
                      disabledforeground="#7d8784")
        b.bind("<Enter>", lambda _e: b.configure(bg=encima)
               if str(b["state"]) != "disabled" else None)
        b.bind("<Leave>", lambda _e: b.configure(bg=fondo)
               if str(b["state"]) != "disabled" else None)
        b._fondo = fondo
        return b

    def _armar_botones(self):
        fila = tk.Frame(self.raiz, bg=FONDO)
        fila.pack(fill="x", padx=14)

        self.btn_encender = self._boton(fila, "Encender", self.encender, True)
        self.btn_encender.pack(side="left")
        self.btn_apagar = self._boton(fila, "Apagar", self.apagar)
        self.btn_apagar.pack(side="left", padx=8)
        self.btn_reiniciar = self._boton(fila, "Reiniciar", self.reiniciar)
        self.btn_reiniciar.pack(side="left")
        self.btn_abrir = self._boton(fila, "Abrir en el navegador", self.abrir)
        self.btn_abrir.pack(side="left", padx=8)

        opciones = tk.Frame(self.raiz, bg=FONDO)
        opciones.pack(fill="x", padx=14, pady=(10, 4))
        self._casilla(opciones,
                      "Volver a levantarla si se cae sola",
                      self.revivir, self.guardar_preferencias).pack(side="left")
        self._casilla(opciones,
                      "Levantarla al prender el computador",
                      self.al_prender,
                      self.guardar_preferencias).pack(side="left", padx=(18, 0))

    def _casilla(self, padre, texto, variable, orden):
        return tk.Checkbutton(
            padre, text=texto, variable=variable, command=orden,
            font=self.chicos, bg=FONDO, fg=GRIS, selectcolor=HUECO,
            activebackground=FONDO, activeforeground=TINTA, bd=0,
            highlightthickness=0, cursor="hand2")

    def _armar_consola(self):
        barra = tk.Frame(self.raiz, bg=FONDO)
        barra.pack(fill="x", padx=14, pady=(8, 4))
        tk.Label(barra, text="CONSOLA", font=self.chicos, bg=FONDO,
                 fg=GRIS).pack(side="left")
        for valor, texto in (("todo", "Todo"), ("solicitudes", "Solicitudes"),
                             ("errores", "Solo errores")):
            tk.Radiobutton(
                barra, text=texto, value=valor, variable=self.filtro,
                command=self._repintar, font=self.chicos, bg=FONDO, fg=GRIS,
                selectcolor=HUECO, activebackground=FONDO,
                activeforeground=TINTA, bd=0, highlightthickness=0,
                cursor="hand2").pack(side="left", padx=(10, 0))
        self._casilla(barra, "Seguir el final", self.seguir,
                      lambda: None).pack(side="right")
        self._casilla(barra, "Sin las comprobaciones del panel",
                      self.sin_salud, self._repintar).pack(side="right",
                                                           padx=(0, 16))

        caja = tk.Frame(self.raiz, bg=HUECO, highlightbackground=BORDE,
                        highlightthickness=1)
        caja.pack(fill="both", expand=True, padx=14)
        caja.rowconfigure(0, weight=1)
        caja.columnconfigure(0, weight=1)
        self.texto = tk.Text(caja, bg=HUECO, fg=GRIS, font=self.consola,
                             relief="flat", bd=0, padx=12, pady=10, wrap="none",
                             insertbackground=TINTA, state="disabled")
        barra_v = ttk.Scrollbar(caja, command=self.texto.yview,
                                style="Fito.Vertical.TScrollbar")
        # Y una barra HORIZONTAL: las lineas no se parten -partirlas mezcla una
        # traza con la solicitud de al lado- asi que sin ella el final de una
        # linea larga no se puede leer de ninguna forma. Se veia en el aviso
        # del puerto ocupado, que es justo el que hay que leer entero.
        barra_h = ttk.Scrollbar(caja, orient="horizontal",
                                command=self.texto.xview,
                                style="Fito.Horizontal.TScrollbar")
        self.texto.configure(yscrollcommand=barra_v.set,
                             xscrollcommand=barra_h.set)
        self.texto.grid(row=0, column=0, sticky="nsew")
        barra_v.grid(row=0, column=1, sticky="ns")
        barra_h.grid(row=1, column=0, sticky="ew")

        self.texto.tag_config("bien", foreground="#8fa08c")
        self.texto.tag_config("aviso", foreground=AVISO)
        self.texto.tag_config("error", foreground=ALERTA)
        self.texto.tag_config("detalle", foreground="#6e7a76")
        self.texto.tag_config("vigilante", foreground=LIMA)
        self.texto.tag_config("separador", foreground="#4b5654")
        self.texto.tag_config("panel", foreground=TINTA)

    def _armar_pie(self):
        self.pie = tk.Label(self.raiz, text="", font=self.chicos, bg=FONDO,
                            fg=GRIS, anchor="w")
        self.pie.pack(side="bottom", fill="x", padx=16, pady=(6, 12))

    # ----------------------------------------------------------- estado

    def _mirar_estado(self):
        """Se pregunta en un hilo: `estado()` llama a netstat y hace una
        peticion, y hacerlo en el hilo de la ventana la congela cada dos
        segundos."""
        if not self.vivo:
            return

        def trabajo():
            try:
                self.mensajes.put(("estado", servicio.estado()))
            except Exception as e:                       # noqa: BLE001
                self.mensajes.put(("estado", {"situacion": "?",
                                              "motivo": str(e)}))
        threading.Thread(target=trabajo, daemon=True).start()
        self._programar(CADA_CUANTO_MIRA, self._mirar_estado)

    def _pintar_estado(self, e):
        colores = {"encendida": (LIMA, "Encendida"),
                   "apagada": (GRIS, "Apagada"),
                   "caida": (ALERTA, "Caída"),
                   "ocupado": (AVISO, "Puerto ocupado")}
        color, rotulo = colores.get(e.get("situacion"), (AVISO, "?"))
        self.punto.itemconfigure(self.bola, fill=color)
        self.rotulo.configure(text=rotulo, fg=TINTA)
        motivo = e.get("motivo", "")
        self.motivo.configure(text=motivo)
        if motivo and not self.motivo.winfo_ismapped():
            self.motivo.pack(fill="x", pady=(2, 0), before=self.direcciones)
        elif not motivo and self.motivo.winfo_ismapped():
            self.motivo.pack_forget()

        if e.get("encendida"):
            # La direccion de Tailscale se pregunta UNA vez, en un hilo: es un
            # programa externo con seis segundos de espera, y llamarlo aca lo
            # haria dos veces por segundo desde el hilo de la ventana.
            fuera = self.direccion_fuera
            texto = "En este computador:  {}".format(e.get("url", ""))
            if fuera:
                texto += "        Desde el teléfono:  http://{}:{}".format(
                    fuera, e.get("puerto"))
            if e.get("pid"):
                texto += "        proceso {}".format(e["pid"])
        else:
            texto = "Nadie puede entrar mientras esté apagada."
        self.direcciones.configure(text=texto)

        if not self.ocupado:
            self._habilitar(self.btn_encender, not e.get("encendida"))
            self._habilitar(self.btn_apagar, bool(e.get("encendida")))
            self._habilitar(self.btn_reiniciar, bool(e.get("encendida")))
            self._habilitar(self.btn_abrir, bool(e.get("encendida")))

    def _habilitar(self, boton, si):
        boton.configure(state="normal" if si else "disabled",
                        bg=boton._fondo if si else "#171c1e",
                        cursor="hand2" if si else "arrow")

    # ---------------------------------------------------------- preferencias

    def _cargar_preferencias(self):
        # La primera vez adopta lo que este pasando, en vez de nacer diciendo
        # «apagada» sobre una aplicacion que esta corriendo.
        servicio.sembrar_si_falta()
        p = servicio.preferencias()
        self.revivir.set(bool(p["revivir"]))
        self.al_prender.set(bool(p["al_prender"]))

    def guardar_preferencias(self):
        servicio.guardar_preferencias(revivir=self.revivir.get(),
                                      al_prender=self.al_prender.get())
        self._anotar("panel", "preferencias guardadas: revivir si se cae = {} · "
                     "levantarla al prender el computador = {}".format(
                         "sí" if self.revivir.get() else "no",
                         "sí" if self.al_prender.get() else "no"))

    # -------------------------------------------------------------- acciones

    def _en_un_hilo(self, nombre, funcion):
        """Encender puede tardar treinta segundos. En el hilo de la ventana
        eso es la aplicacion colgada, sin dibujar y sin poder cerrarse."""
        if self.ocupado:
            return
        self.ocupado = True
        for b in (self.btn_encender, self.btn_apagar, self.btn_reiniciar):
            self._habilitar(b, False)
        self._anotar("panel", nombre + "…")

        def trabajo():
            try:
                ok, mensaje = funcion()
            except Exception as e:                        # noqa: BLE001
                ok, mensaje = False, "falló: {}".format(e)
            self.mensajes.put(("accion", (ok, mensaje)))
        threading.Thread(target=trabajo, daemon=True).start()

    def encender(self):
        self._en_un_hilo("Encendiendo", lambda: servicio.encender(
            anotar=lambda t: self.mensajes.put(("nota", t))))

    def apagar(self):
        self._en_un_hilo("Apagando", servicio.apagar)

    def reiniciar(self):
        self._en_un_hilo("Reiniciando", lambda: servicio.reiniciar(
            anotar=lambda t: self.mensajes.put(("nota", t))))

    def abrir(self):
        webbrowser.open(servicio.estado()["url"])

    # -------------------------------------------------------- los registros

    ARCHIVOS = (
        (servicio.REGISTRO, "acceso"),
        (servicio.REGISTRO_ERROR, "salida"),
        (servicio.REGISTRO_VIGILANTE, "vigilante"),
    )

    def _cargar_registros_previos(self, cuantas=140):
        for ruta, origen in self.ARCHIVOS:
            try:
                with ruta.open("rb") as fh:
                    crudo = fh.read()
                self.posiciones[ruta] = len(crudo)
            except OSError:
                self.posiciones[ruta] = 0
                continue
            lineas = crudo.decode("utf-8", "replace").splitlines()
            for linea in lineas[-cuantas:]:
                self._agregar(origen, linea, con_hora=False)
        self._anotar("separador",
                     "──── lo de arriba es de antes de abrir el panel ────")
        self._repintar()

    def _leer_registros(self):
        if not self.vivo:
            return
        hubo = False
        for ruta, origen in self.ARCHIVOS:
            try:
                tam = ruta.stat().st_size
            except OSError:
                continue
            desde = self.posiciones.get(ruta, 0)
            if tam < desde:                 # lo recortaron: se empieza de nuevo
                desde = 0
                self._agregar("panel", "— el registro se recortó —")
            if tam == desde:
                continue
            try:
                with ruta.open("rb") as fh:
                    fh.seek(desde)
                    crudo = fh.read()
            except OSError:
                continue
            self.posiciones[ruta] = tam
            texto = self.restos.pop(ruta, "") + crudo.decode("utf-8", "replace")
            partes = texto.split("\n")
            self.restos[ruta] = partes.pop()      # media linea sin terminar
            for linea in partes:
                if linea.strip():
                    self._agregar(origen, linea.rstrip("\r"))
                    hubo = True
        if hubo:
            self._repintar()
        self._programar(CADA_CUANTO_LEE, self._leer_registros)

    def _agregar(self, origen, linea, con_hora=True):
        """Clasifica una linea y la guarda. El color lo decide esto, no el
        pintado: asi el filtro y la vista no pueden discrepar."""
        hora = datetime.now().strftime("%H:%M:%S") if con_hora else "        "
        if origen == "vigilante":
            # Que la haya repuesto no es una noticia buena: significa que se
            # cayo. Va con el color de lo que hay que mirar.
            incidente = any(p in linea for p in
                            ("NO ", "no responde", "caído", "cayó", "FALL"))
            self.entradas.append(("error" if incidente else "vigilante",
                                  linea, "vigilante"))
        elif origen == "panel":
            self.entradas.append(("panel", linea, "panel"))
        elif origen == "separador":
            self.entradas.append(("separador", linea, "separador"))
        elif origen == "acceso":
            dato = servicio.leer_linea_de_acceso(linea)
            if dato:
                self.entradas.append((
                    dato["clase"],
                    "{}  {:<6} {:<52} {}".format(hora, dato["metodo"],
                                                 dato["ruta"][:52],
                                                 dato["codigo"]),
                    "salud" if dato["ruta"].startswith("/salud")
                    else "solicitud"))
            else:
                self.entradas.append(("detalle", linea, "otro"))
        else:
            grave = any(p in linea for p in
                        ("Traceback", "ERROR", "Error", "CRITICAL",
                         "Exception", "raise "))
            self.entradas.append(("error" if grave else "detalle", linea,
                                  "otro"))
        del self.entradas[:-TOPE_DE_VISTA]

    def _anotar(self, clase, texto):
        self.entradas.append((clase, "{}  {}".format(
            datetime.now().strftime("%H:%M:%S"), texto), "panel"))
        self._repintar()

    def _pasa_el_filtro(self, clase, tipo):
        if tipo == "salud" and self.sin_salud.get():
            return False
        filtro = self.filtro.get()
        if filtro == "todo":
            return True
        if filtro == "solicitudes":
            return tipo in ("solicitud", "salud")
        # «Solo errores» incluye el 4xx: una solicitud que devolvio 403 no es
        # una caida del servidor, pero es lo que uno viene a buscar cuando
        # algo «no funciona».
        return clase in ("error", "aviso")

    def _repintar(self):
        # Se repinta entero en vez de ir agregando: son mil quinientas lineas
        # como mucho y asi cambiar de filtro y recibir una linea nueva pasan
        # por el mismo camino. Dos caminos para pintar lo mismo se separan.
        estaba_al_final = self.seguir.get()
        self.texto.configure(state="normal")
        self.texto.delete("1.0", "end")
        for clase, linea, tipo in self.entradas:
            if self._pasa_el_filtro(clase, tipo):
                self.texto.insert("end", linea + "\n", clase)
        self.texto.configure(state="disabled")
        if estaba_al_final:
            self.texto.see("end")
        self._pintar_pie()

    def _pintar_pie(self):
        # Se cuenta lo que SE VE. Decir «69 solicitudes» cuando en pantalla hay
        # cuatro, porque las otras 65 son comprobaciones escondidas, es un
        # numero que no se puede comprobar mirando.
        visibles = [(c, t) for c, _l, t in self.entradas
                    if self._pasa_el_filtro(c, t) and t in ("solicitud",
                                                            "salud")]
        total = len(visibles)
        malas = sum(1 for c, _t in visibles if c in ("aviso", "error"))
        graves = sum(1 for c, _t in visibles if c == "error")
        self.pie.configure(text=(
            "{} solicitudes en la vista · {} devolvieron error · {} fallaron "
            "en el servidor      ·      los registros están en data\\".format(
                total, malas, graves)))

    # ------------------------------------------------------------- el bucle

    def _drenar(self):
        """Todo lo que llega de los hilos se pinta ACA, en el hilo de la
        ventana. tkinter no es seguro desde otro hilo: se ve como cuelgues
        raros y ventanas a medio dibujar."""
        if not self.vivo:
            return
        try:
            while True:
                que, dato = self.mensajes.get_nowait()
                if que == "estado":
                    self._pintar_estado(dato)
                elif que == "nota":
                    self._anotar("panel", dato)
                elif que == "accion":
                    ok, mensaje = dato
                    self.ocupado = False
                    self._anotar("panel" if ok else "error", mensaje)
        except queue.Empty:
            pass
        self._programar(120, self._drenar)


def main():
    raiz = tk.Tk()
    Panel(raiz)
    raiz.mainloop()


if __name__ == "__main__":
    main()
