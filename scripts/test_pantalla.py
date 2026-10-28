"""Prueba de la PANTALLA: que los controles esten conectados.

Existe por una leccion cara. La revision completa del 2026-09-03 encontro que
el buscador de curvas y la casilla "solo incompletas" no hacian NADA: la logica
de filtrado estaba escrita, se leia bien, y ningun listener la disparaba.
Ninguna de las 636 comprobaciones del backend podia verlo, porque el problema
no estaba en el servidor.

No reemplaza probar en el navegador (§14): esto solo lee el HTML. Pero atrapa
la clase de error mas boba y mas frecuente -un control sin cablear, un id que
se renombro en un sitio y no en el otro- sin levantar nada.

Uso: python scripts/test_pantalla.py
"""

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.settings import ajustes as config                    # noqa: E402

fallos, comprobaciones = [], 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)


def cargar():
    return (config.raiz / "web" / "index.html").read_text(encoding="utf-8")


def guion_de(html):
    return html.split("<script>", 1)[-1]


def escucha(guion, identificador):
    """Si al control se le puso un listener, directo o desde una lista."""
    directo = r'\$\("' + re.escape(identificador) + r'"\)\.addEventListener'
    if re.search(directo, guion):
        return True
    # La forma corta: ["a", "b"].forEach(... addEventListener ...)
    cerca = r'"' + re.escape(identificador) + r'"[\s\S]{0,300}?addEventListener'
    return bool(re.search(cerca, guion))


def probar_controles_conectados(html):
    """Todo control con id tiene que ser LEIDO o ESCUCHADO por el guion.

    Un control que nadie lee es un control que no hace nada. Es exactamente lo
    que pasaba con el buscador de curvas: existia, se veia, y era decorativo.
    """
    print("\n[1] Ningun control quedo suelto")
    ids = set(re.findall(
        r'<(?:input|select|textarea|button)[^>]*\bid="([^"]+)"', html))
    comprobar(len(ids) > 30, "se encontraron los controles de la pagina",
              "{} controles".format(len(ids)))

    guion = guion_de(html)
    huerfanos = [i for i in sorted(ids)
                 if not re.search(r'["\']' + re.escape(i) + r'["\']', guion)]
    comprobar(not huerfanos,
              "todos los controles se usan desde el guion",
              "sin usar: " + ", ".join(huerfanos))


def probar_filtros_de_lista(html):
    """Un buscador que no repinta su lista es un adorno."""
    print("\n[2] Los buscadores y filtros repintan su lista")
    guion = guion_de(html)
    for campo, pintor in (("buscarCurva", "pintarListaCurvas"),
                          ("soloIncompletas", "pintarListaCurvas"),
                          ("buscarExp", "cargarExpedientes"),
                          ("filtroEstado", "cargarExpedientes")):
        comprobar(escucha(guion, campo),
                  "'{}' dispara el repintado".format(campo),
                  "no tiene ningun addEventListener")
        comprobar(pintor in guion, "existe '{}'".format(pintor))


def probar_botones_con_manejador(html):
    print("\n[3] Cada boton hace algo")
    guion = guion_de(html)
    botones = re.findall(r'<button[^>]*\bid="(btn[^"]+)"', html)
    comprobar(len(botones) >= 10, "se encontraron los botones",
              "{} botones".format(len(botones)))
    # Un boton type="submit" NO lleva su propio manejador: lo correcto es
    # escuchar el 'submit' del formulario, que ademas hace funcionar Enter en
    # los campos. Se da por conectado si su formulario esta escuchado.
    envios = set(re.findall(r'<button[^>]*type="submit"[^>]*\bid="(btn[^"]+)"', html))
    formularios = set(re.findall(r'<form[^>]*\bid="([^"]+)"', html))
    con_submit = any(re.search(
        r'\$\("' + re.escape(f) + r'"\)\.addEventListener\("submit"', guion)
        for f in formularios)
    mudos = [b for b in botones
             if not escucha(guion, b) and not (b in envios and con_submit)]
    comprobar(not mudos, "ningun boton quedo sin manejador", ", ".join(mudos))


def probar_vocabulario(html):
    """En la pantalla se dice 'informe', que es la palabra que usan ellos.

    'Expediente' es el nombre del MODELO -guarda documentos, fotos y catastro-
    y adentro sigue llamandose asi. Lo que no debe pasar es que la jerga del
    modelo se le aparezca al tasador.
    """
    print("\n[4] La pantalla habla como Fitogreen")
    visible = re.sub(r"<script>[\s\S]*?</script>", "", html)
    visible = re.sub(r"<!--[\s\S]*?-->", "", visible)
    textos = re.findall(r">([^<>]{4,})<", visible)
    jerga = [t.strip() for t in textos if re.search(r"\bexpediente", t, re.I)]
    comprobar(not jerga, "no queda 'expediente' en un texto visible",
              " | ".join(jerga[:3]))


def probar_un_solo_manejador_por_input(html):
    """Dos manejadores de 'change' sobre el mismo <input type=file> se pisan.

    Paso de verdad: el catastro llamaba a `conectarZona` -que en su manejador
    hace `input.value = ""` para poder reelegir el mismo archivo- Y ademas
    tenia su propio manejador. Como corren por orden de registro, el primero
    vaciaba `input.files` y el segundo ya no encontraba nada: el explorador se
    abria, uno elegia el Excel y no pasaba nada.
    """
    print("\n[5] Ningun input de archivo con dos manejadores")
    guion = guion_de(html)
    inputs = re.findall(r'<input[^>]*type="file"[^>]*\bid="([^"]+)"', html)
    comprobar(len(inputs) >= 3, "se encontraron los campos de archivo",
              ", ".join(inputs))
    for identificador in inputs:
        propios = len(re.findall(
            r'\$\("{}"\)\.addEventListener\("change"'.format(identificador), guion))
        por_zona = len(re.findall(
            # '[^)]*' no sirve: el primer argumento -$("zonaCat")- ya trae
            # un parentesis y la busqueda se cortaba ahi.
            r'conectarZona\([\s\S]{{0,140}}?\$\("{}"\)'.format(identificador), guion))
        comprobar(propios + por_zona <= 1,
                  "'{}' tiene un solo manejador de 'change'".format(identificador),
                  "{} propio(s) + {} por conectarZona".format(propios, por_zona))


def probar_sesion_opcional(html):
    """El boton de generar no puede exigir una sesion de subida.

    Un informe armado desde el catastro y la visita no tiene archivos subidos,
    asi que su sesion viene vacia. Exigirla devolvia 400 y el informe no se
    podia emitir.
    """
    print("\n[6] Generar no depende de haber subido archivos")
    ruta = config.raiz / "api" / "routers" / "asistente.py"
    codigo = ruta.read_text(encoding="utf-8")
    comprobar("if not sesion:\n        return None" in codigo,
              "la sesion vacia es un caso valido, no un error")
    comprobar('sesion: str = Field(' in codigo,
              "y el modelo la declara opcional")


def probar_que_todo_archivo_se_pueda_quitar(html):
    """Toda zona de subida tiene que dejar sacar lo que se puso.

    El catastro no lo tenia: elegir la planilla equivocada solo se arreglaba
    RECARGANDO la pagina. Y de paso, el input nunca se vaciaba, asi que volver
    a elegir el MISMO archivo no disparaba 'change' y la zona parecia muerta.
    """
    print("\n[8] Lo que se sube se puede quitar")
    guion = guion_de(html)
    comprobar("quitarArchivoCat" in guion,
              "el catastro tiene forma de quitar el archivo")
    comprobar('quitar.className = "quitar"' in guion.split("ponerArchivoCat")[1][:1200]
              if "ponerArchivoCat" in guion else False,
              "y lo hace con el mismo boton que las otras zonas")
    # Vaciar el input es lo que permite volver a elegir el mismo archivo. Tiene
    # que ir DESPUES de leerlo: al reves deja `files` sin nada
    # (§16.a.vicies.quinquies).
    #
    # Se quitan los comentarios ANTES de mirar el orden: el propio comentario
    # que explica la trampa nombra `input.value = ""`, y sin esto la prueba se
    # encontraba a si misma y daba en rojo con el codigo correcto.
    trozo = guion.split('$("inputCat").addEventListener("change"')[1][:500]
    codigo = "\n".join(l for l in trozo.splitlines()
                       if not l.strip().startswith("//"))
    comprobar('.files[0]' in codigo and '.value = ""' in codigo
              and codigo.index('.files[0]') < codigo.index('.value = ""'),
              "lee el archivo ANTES de vaciar el input, no al reves")


def probar_sin_alert(html):
    """Ningun `alert()`: en el navegador integrado quedan SUPRIMIDOS.

    O sea que el mensaje desaparece sin que nadie lo vea, y desde fuera es
    indistinguible de «apreté y no pasó nada». Paso de verdad: analizar una
    planilla equivocada no mostraba ni resultado ni error. Ademas sale con la
    tipografia del sistema y el rotulo «localhost:8000 dice:».
    """
    print("\n[9] Ningun alert() del navegador")
    guion = guion_de(html)
    llamadas = [l.strip() for l in guion.splitlines()
                if re.search(r"(^|[^.\w])alert\(", l)
                and not l.strip().startswith("//")]
    comprobar(not llamadas, "los mensajes salen por la pagina, no por alert()",
              " | ".join(l[:60] for l in llamadas[:3]))
    comprobar("function avisar(" in guion,
              "existe el aviso flotante que los reemplaza")
    comprobar('id="avisos"' in html, "y su contenedor esta en la pagina")


def probar_cadena_de_exportacion(html):
    """El FOB elegido tiene que LLEGAR al guardado, con su procedencia.

    Elegir «con cascara» y que se guarde el promedio del grupo -o que se
    guarde el numero sin decir de que producto era- es peor que no ofrecer la
    eleccion: el informe diria un precio que nadie puede auditar. Y en el
    avellano la diferencia es 4,17 contra 9,99 USD/kg.
    """
    print("\n[10] La cadena de exportacion llega al guardado")
    guion = guion_de(html)

    comprobar("cadena_exportacion:" in guion,
              "el cuerpo que se guarda incluye la cadena")
    # Los tres tienen que viajar juntos: un precio sin de donde salio no se
    # puede citar en el Excel, que es la mitad del punto (§10.2).
    trozo = guion.split("cadena_exportacion:")[1][:600]
    for campo in ("precio_fob_usd_kg", "fuente_fob", "referencia_fob",
                  "descuentos"):
        comprobar(campo in trozo, "viaja «{}»".format(campo))

    comprobar("function usarEsteFob(" in guion,
              "elegir un codigo de ODEPA baja el precio al formulario")
    # Sin FOB no se manda cadena: el modelo exige precio mayor que cero y una
    # cadena vacia haria fallar el guardado entero.
    comprobar('$("cFob").value || 0) > 0' in guion,
              "y no se manda cadena si no hay FOB")


def _solo_codigo(guion):
    """Quita comentarios y textos: una palabra castellana seguida de parentesis
    dentro de un comentario no es una llamada a funcion."""
    sin = re.sub(r"/\*[\s\S]*?\*/", " ", guion)
    sin = re.sub(r"(?m)//.*$", " ", sin)
    sin = re.sub(r"`(?:[^`\\]|\\.)*`", '""', sin)
    sin = re.sub(r"'(?:[^'\\\n]|\\.)*'", '""', sin)
    sin = re.sub(r'"(?:[^"\\\n]|\\.)*"', '""', sin)
    return sin


#: Lo que provee el navegador y no esta definido en el guion.
DEL_NAVEGADOR = {
    "fetch", "setTimeout", "setInterval", "clearTimeout", "clearInterval",
    "parseInt", "parseFloat", "isNaN", "encodeURIComponent",
    "decodeURIComponent", "atob", "btoa", "getComputedStyle", "matchMedia",
    "requestAnimationFrame", "cancelAnimationFrame", "structuredClone",
    "queueMicrotask", "alert", "confirm", "prompt",
}


def probar_funciones_que_existen(html):
    """Ninguna llamada a una funcion que no esta definida.

    Paso el 2026-09-08: se llamo a `dibujarCurva()`, que no existe -la funcion
    real es `pintarGraficoCurva`-. La pantalla se rompia solo al apretar ESE
    boton, asi que ninguna de las otras comprobaciones lo veia: el control
    estaba cableado, el boton tenia manejador, y el manejador reventaba.

    **Alcance, para que no se lea como mas de lo que es:** solo mira las
    llamadas en camelCase con mayuscula -`pintarGraficoCurva`- que es como se
    llaman las funciones de este proyecto. Una funcion de una sola palabra
    (`nota`, `num`) mal escrita no la atrapa: filtrarlas produce falsos
    positivos con el castellano de los textos.
    """
    print("\n[11] Ninguna llamada a una funcion inexistente")
    guion = _solo_codigo(guion_de(html))

    definidas = set(re.findall(r"\bfunction\s+([A-Za-z_$][\w$]*)\s*\(", guion))
    definidas |= set(re.findall(
        r"\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=", guion))
    # Los parametros de una funcion flecha tambien pueden ser funciones.
    for firma in re.findall(r"\(([^()]*)\)\s*=>", guion):
        definidas |= set(re.findall(r"[A-Za-z_$][\w$]*", firma))
    definidas |= set(re.findall(r"([A-Za-z_$][\w$]*)\s*=>", guion))

    comprobar(len(definidas) > 100, "se encontraron las funciones del guion",
              "{} definidas".format(len(definidas)))

    llamadas = set(re.findall(r"(?<![.\w$])([a-z][a-z0-9$_]*[A-Z][\w$]*)\s*\(",
                              guion))
    huerfanas = sorted(llamadas - definidas - DEL_NAVEGADOR)
    comprobar(not huerfanas, "todas las funciones que se llaman existen",
              ", ".join(huerfanas))


def probar_que_ocultar_oculta(html):
    """`.oculto` tiene que esconder, gane quien gane en especificidad.

    Costo una tarde. `.oculto` es una CLASE (0,0,1,0) y `#acceso` un ID
    (0,1,0,0) con `display:grid`: al entrar bien, el guion le ponia la clase,
    el navegador la ignoraba, y la pantalla de acceso seguia tapando la
    aplicacion -fixed, z-index 300-. El sintoma era login correcto, sin ningun
    error, y la sensacion de que la contrasena estaba mal.

    Y la razon de que se me escapara en las pruebas del navegador: yo
    comprobaba `classList.contains("oculto")`, que era **true**. Hay que mirar
    el `display` que se aplica, no la clase que se puso.
    """
    print("\n[7] Ocultar oculta de verdad")
    css = html.split("</style>")[0]

    regla = re.search(r'\.oculto\s*\{([^}]*)\}', css)
    comprobar(regla is not None and "!important" in regla.group(1),
              "`.oculto` se impone sobre cualquier regla por id",
              regla.group(1).strip() if regla else "no esta declarada")

    # Y, por si algun dia se quita el !important: que no haya un id con
    # `display` propio entre los elementos a los que el guion se la pone.
    con_display = {}
    for selector, cuerpo in re.findall(r'(#[A-Za-z0-9_-]+[^{}]*)\{([^{}]*)\}', css):
        if re.search(r'\bdisplay\s*:', cuerpo):
            for ident in re.findall(r'#([A-Za-z0-9_-]+)', selector):
                con_display.setdefault(ident, selector.strip())

    patron = r'(?:\$\("|getElementById\(")([A-Za-z0-9_-]+)"\)\.classList\.(?:add|remove|toggle)\("oculto"'
    tocados = set(re.findall(patron, html))
    chocan = sorted(tocados & set(con_display))
    comprobar(
        not chocan or ("!important" in (regla.group(1) if regla else "")),
        "ningun elemento que se oculta tiene un `display` por id que lo anule",
        ", ".join("#" + i for i in chocan))


def probar_grafica_de_odepa(html):
    """El resumen de ODEPA se CALCULA en el servidor; la pantalla lo dibuja.

    Es la regla 80/20 aplicada aca: el minimo, el maximo y el promedio son
    aritmetica, y viven en `services/referencias.py`, donde hay pruebas que
    los contrastan. Recalcularlos en el navegador daria dos versiones del
    mismo numero y la que veria el tasador seria la que nadie prueba.

    Y la grafica se despliega al lado midiendo el borde REAL del cuadro. Un
    elemento escondido mide cero, y con cero la cuenta manda el panel al
    costado de la pantalla: paso al probarlo con la pestana de Curvas
    cerrada.
    """
    print("\n[12] La grafica de ODEPA")
    guion = guion_de(html)
    codigo = _solo_codigo(guion)

    comprobar("e.resumen" in codigo,
              "el rango del FOB sale de la respuesta del servidor")
    comprobar("s.decada" in codigo,
              "y el de la superficie tambien")
    # Solo dentro de `pintarReferencias`: el `Math.max` de `graficoDeSerie`
    # es legitimo -dimensiona el eje- y no es el numero que se muestra.
    pintar = re.search(r"function pintarReferencias\([\s\S]*?\n\}", codigo)
    comprobar(pintar is not None, "existe pintarReferencias")
    if pintar:
        comprobar("Math.min" not in pintar.group(0)
                  and "Math.max" not in pintar.group(0),
                  "el minimo y el maximo no se recalculan para mostrarlos")
    comprobar("resumen.puntos" in codigo or "r.puntos" in codigo,
              "se dibujan los puntos que manda el servidor")

    comprobar('data-serie="${cual}"' in guion,
              "cada serie tiene su boton de desplegar")
    # En el guion crudo: `_solo_codigo` borra los textos, y el selector ES un
    # texto. Buscarlo ahi daba en rojo con el cableado puesto.
    comprobar('querySelectorAll("[data-serie]")' in guion
              and "desplegarGrafica(" in codigo,
              "y el boton esta cableado")

    cuerpo = re.search(r"function acomodarGrafica\([\s\S]*?\n\}", codigo)
    comprobar(cuerpo is not None, "existe la funcion que la coloca")
    if cuerpo:
        comprobar("caja.width" in cuerpo.group(0)
                  and "return" in cuerpo.group(0).split("caja.width")[1][:40],
                  "que no coloca nada si el cuadro esta escondido y mide cero")


def probar_documentos_legales_opcionales(html):
    """Se tiene que poder hacer un informe SIN documentos legales.

    el agrónomo dijo que no los usan. El boton de analizar es ademas el que CREA
    el informe (§16.a.septdecies), y exigia «al menos un documento o una
    fotografia»: sin subir un PDF que no van a subir, el circuito entero
    quedaba bloqueado en el paso 2 y no se podia ni empezar.
    """
    print("\n[13] Los documentos legales son opcionales")
    guion = guion_de(html)
    codigo = _solo_codigo(guion)

    comprobar("No hay nada que analizar" not in guion,
              "el boton de analizar NO exige subir archivos")
    # En el guion CRUDO: `_solo_codigo` borra los textos, y el selector es un
    # texto -`$("btnAnalizar")` queda en `$("")`-. Buscarlo ahi daba en rojo
    # con el manejador puesto. Segunda vez que me pasa lo mismo.
    manejador = re.search(
        r'\$\("btnAnalizar"\)\.addEventListener\([\s\S]*?\n\}\);', guion)
    comprobar(manejador is not None, "se encontro el manejador del boton")
    if manejador:
        comprobar("!docs.length && !fotos.length" not in manejador.group(0),
                  "y no se devuelve temprano por no tener archivos")

    comprobar('class="opcional"' in html,
              "el paso 2 esta rotulado como opcional")
    comprobar("salte este paso" in html,
              "y dice que se puede saltar")
    comprobar("rotularBotonDeAnalizar" in codigo,
              "el boton dice lo que va a hacer segun lo que haya cargado")
    comprobar('opcional: "' in guion and 'f[0] !== "opcional"' in guion,
              "«opcional» no cuenta como algo que falta en la lista del paso 4")


def probar_apartado_satelital(html):
    """Las satelitales tienen su propio apartado, y viajan por su propio campo.

    el agrónomo: «las fotos satelitales tienen que tener buena calidad y tienen
    que ser mas grandes que las otras imagenes, asi que no deberian tener que
    subirse al mismo apartado de las otras fotografias».

    Lo que hay que vigilar es que NO se cuelen dentro de `fotos`: si viajaran
    por ese campo el servidor las comprimiria a 1024 px como cualquier foto, y
    el defecto seria invisible -la imagen sale, solo que blanda- hasta que
    alguien imprima el informe.
    """
    print("\n[14] El apartado de imagenes satelitales")
    guion = guion_de(html)
    codigo = _solo_codigo(guion)

    comprobar('id="zonaSat"' in html and 'id="inputSat"' in html
              and 'id="listaSat"' in html,
              "el apartado existe, con su zona, su input y su lista")
    comprobar("satelitales" in codigo and "const satelitales" in guion,
              "y tiene su propio deposito, separado de `fotos`")
    comprobar(re.search(r'conectarZona\(\$\("zonaSat"\)', guion) is not None,
              "la zona esta cableada")

    # El campo del formulario: `satelitales`, no `fotos`.
    manejador = re.search(
        r'\$\("btnAnalizar"\)\.addEventListener\([\s\S]*?\n\}\);', guion)
    comprobar(manejador is not None, "se encontro el manejador de analizar")
    if manejador:
        cuerpo = manejador.group(0)
        comprobar('cuerpo.append("satelitales"' in cuerpo,
                  "las satelitales viajan por su propio campo")
        comprobar('satelitales.forEach' in cuerpo,
                  "y se mandan todas, no solo la primera")
        # La trampa: que alguien las agregue a `fotos` "para simplificar".
        #
        # OJO con el `[^)]*` aqui: la primera version lo usaba y NO detectaba
        # nada, porque la flecha `(s) =>` ya trae un parentesis y la busqueda
        # se cortaba ahi. Es exactamente el error de §16.a.vicies.quinquies, y
        # lo destapo comprobar la seccion en negativo -no leerla-.
        comprobar(not re.search(r'satelitales\.forEach\([\s\S]{0,120}?'
                                r'append\("fotos"', cuerpo),
                  "y NO se cuelan dentro del campo de fotos")

    comprobar(html.count('class="opcional"') >= 3,
              "los tres apartados prescindibles estan rotulados como opcionales",
              "hay {}".format(html.count('class="opcional"')))
    comprobar("Imágenes satelitales" in html and "opcional" in html,
              "el paso 4 se llama por su nombre")

    # La miniatura de la revision: el servidor devuelve fotos + satelitales en
    # ese orden, asi que leer solo `fotos[n]` deja sin miniatura a las
    # satelitales. No revienta -queda un hueco- y por eso hace falta mirarlo.
    comprobar("[...fotos, ...satelitales][n]" in codigo,
              "la miniatura de la revision recorre las dos listas")

    # El aviso de nitidez: cambiar la categoria a satelital NO cambia la
    # resolucion con que ya se guardo.
    comprobar("alta_resolucion" in codigo,
              "se sabe cual entro en alta resolucion")
    comprobar("data-blanda" in html,
              "y se avisa si una foto de 1024 px se marca como satelital")


def probar_calendario_propio(html):
    """El calendario de la fecha es de la pagina, no el de Chrome.

    el usuario: *«cuando uno escoge la fecha de visita el calendario es un diseno
    bien normal nomas»*. El almanaque de un `<input type="date">` lo dibuja
    el navegador y NO hay CSS que lo toque, asi que se reemplaza igual que la
    lista de un `<select>` (§16.a.vicies.quater): el input real se queda vivo
    y encima va uno propio.

    Lo que hay que vigilar no es el aspecto -eso se mira- sino que el input
    REAL siga siendo la fuente de la verdad: todo el guion lee
    `$("fecha").value` y espera un ISO. Si el reemplazo se queda con el dato,
    la fecha del informe deja de llegar al Excel y no lo avisa nadie.
    """
    print("\n[15] El calendario propio de la fecha")
    guion = guion_de(html)
    codigo = _solo_codigo(guion)

    # Se busca la DEFINICION, no el nombre suelto: `noMejorarFecha` contiene
    # la cadena «mejorarFecha» y una comprobacion por subcadena la aceptaba.
    comprobar(re.search(r"\bfunction mejorarFecha\(", codigo)
              and re.search(r"\bfunction mejorarFechas\(", codigo),
              "existe el mejorador y su recorrido")
    comprobar('input[type="date"]:not([data-mejorado])' in guion,
              "busca los que faltan por el mismo criterio que los <select>")
    # La llamada del ARRANQUE, al margen izquierdo. Sin anclarla, la que hay
    # dentro del observador servia de coartada y la comprobacion no podia
    # fallar nunca.
    comprobar(re.search(r"(?m)^mejorarFechas\(\);", codigo) is not None,
              "se llama al arrancar, y no solo desde el observador")
    observador = re.search(r"new MutationObserver\(\(\) => \{[\s\S]{0,400}?"
                           r"\}\)\.observe\(document\.body", codigo)
    comprobar(observador is not None and "mejorarFechas" in observador.group(0),
              "y desde el observador del body, para los que se creen sobre la marcha")

    # El input real se ESCONDE, no se borra ni se cambia de tipo: es lo que
    # lee el resto del guion.
    comprobar('type="date"' in html,
              "el input real sigue en la pagina, con su type")
    # La regla CONCRETA, no cualquier `clip` de la hoja: el `.selector >
    # select` tiene el mismo, asi que buscarlo en todo el HTML daba verde
    # aunque a este se le hubiera puesto display:none.
    regla = re.search(r"\.fecha > input\[type=\"date\"\] \{([^}]*)\}", html)
    comprobar(regla is not None,
              "y se esconde por CSS, sin sacarlo del formulario")
    if regla:
        comprobar("clip:" in regla.group(1) and "display:none" not in regla.group(1),
                  "escondido con clip, no con display:none",
                  "display:none sacaria el campo del formulario")

    # Lo elegido tiene que llegar al input real Y avisar al resto. Se busca
    # DENTRO de `mejorarFecha`: el mejorador de <select> tambien dispara un
    # 'change', asi que mirar todo el guion daba verde con esto quitado.
    #
    # Y se busca en el guion CRUDO, no en `codigo`: `_solo_codigo` borra los
    # textos, asi que `new Event("change", ...)` queda en `new Event("", ...)`
    # y la comprobacion daba rojo con el codigo correcto. Tercera vez que me
    # pasa lo mismo (§16.a.quinquagies.semel). Lo que es un texto se busca en
    # el guion crudo.
    cuerpo = re.search(r"(?m)^function mejorarFecha\(entrada\) \{[\s\S]*?^\}",
                       guion)
    comprobar(cuerpo is not None, "se encontro el cuerpo de mejorarFecha")
    if cuerpo:
        comprobar("entrada.value = iso" in cuerpo.group(0),
                  "lo elegido se escribe en el input real")
        comprobar(re.search(r'entrada\.dispatchEvent\(\s*new Event\(\s*"change"'
                            r'[\s\S]{0,40}?bubbles: true', cuerpo.group(0)),
                  "y dispara un 'change' de verdad, que es lo que escucha la pagina")

    # La trampa de la zona horaria: `new Date("2026-03-15")` se interpreta en
    # UTC y en Chile devuelve el DIA ANTERIOR. El calendario marcaria el 14.
    comprobar("fechaDesdeIso" in codigo,
              "la fecha ISO se parte a mano y no con new Date(texto)")
    sospechosas = re.findall(r"new Date\(\s*(?:iso|entrada\.value|texto\.value)", codigo)
    comprobar(not sospechosas,
              "ningun new Date() sobre el texto ISO: en Chile daria el dia anterior",
              str(sospechosas))

    # Un `input.value = x` no dispara evento ni mutacion del atributo.
    comprobar("_repintar" in codigo and "sincronizarFechas" in codigo,
              "hay como repintar tras un `.value = x` suelto")
    valor = re.search(r"function valor\(id, v\) \{[\s\S]*?\n\}", codigo)
    comprobar(valor is not None, "se encontro `valor()`")
    if valor:
        comprobar("_repintar" in valor.group(0),
                  "y `valor()` repinta: no depende de que irAPanel venga despues")

    # Misma leccion que `acomodarGrafica`: un elemento escondido mide cero.
    comprobar("acomodarCalendario" in codigo, "el almanaque se coloca solo")
    acomodar = re.search(r"function acomodarCalendario[\s\S]*?\n\}", codigo)
    comprobar(acomodar is not None and "!marco.width" in acomodar.group(0),
              "y no coloca nada si el campo esta escondido y mide cero")
    comprobar(acomodar is not None
              and "clientWidth" in acomodar.group(0)
              and "clientHeight" in acomodar.group(0),
              "midiendo el borde real, no 100vw/100vh")


def _regla(html, selector):
    """El cuerpo de una regla CSS por su selector exacto. None si no esta."""
    m = re.search(re.escape(selector) + r"\s*\{([^}]*)\}", html)
    return m.group(1) if m else None


def probar_casillas_y_desplegables(html):
    """Los dos controles que todavia se veian «del sistema».

    el usuario: la casilla *«no tiene tanto diseno de nuestra pagina»*, y de los
    `<details>`: *«que cuando pase por encima el mouse sepa yo que tengo que
    apretarlo... hacerlo mas obvio que apretando ahi se desplegara algo»*.

    Una casilla SI se puede dibujar entera con CSS -al reves que la lista de
    un <select> o el almanaque de una fecha-, asi que aca no hace falta
    reemplazar nada: basta con apagar la del sistema con `appearance:none`.
    Sin eso, todo lo demas que se declare lo ignora Windows.
    """
    print("\n[16] Las casillas y los desplegables")

    casilla = _regla(html, 'input[type="checkbox"], input[type="radio"]')
    comprobar(casilla is not None, "hay una regla propia para las casillas")
    if casilla:
        comprobar("appearance:none" in casilla,
                  "se apaga la del sistema: sin esto Windows ignora el resto")
        comprobar("cursor:pointer" in casilla,
                  "y se ve que se puede apretar")

    marcada = _regla(html, 'input[type="checkbox"]:checked')
    comprobar(marcada is not None, "la casilla marcada se dibuja")
    if marcada:
        comprobar("var(--lima)" in marcada, "en el acento de la pagina")
        # El visto sobre lima tiene que ir en el oscuro: en blanco da 1,8:1.
        comprobar("stroke='%23ffffff'" not in marcada
                  and "stroke='%23fff'" not in marcada,
                  "con el visto en oscuro, no en blanco: sobre lima no se leeria")

    # Un `style="width:auto"` en linea pisaba el tamano de la casilla nueva.
    comprobar(not re.search(r'type="checkbox"[^>]*style="[^"]*width:', html),
              "ninguna casilla trae un ancho en linea que pise el tamano")

    resumen = _regla(html, "details > summary")
    comprobar(resumen is not None, "hay una regla propia para los desplegables")
    if resumen:
        comprobar("cursor:pointer" in resumen,
                  "el desplegable dice que se puede apretar")
        comprobar("display:flex" in resumen,
                  "y se ordena en fila para poder poner la pista a la derecha")
    comprobar(_regla(html, "details > summary:hover") is not None,
              "responde al pasar el mouse por encima")

    # El triangulito del navegador, en sus DOS formas: Chrome usa el
    # pseudo-elemento propio y Firefox el ::marker. Apagar solo uno deja el
    # otro dibujado al lado de la flecha nueva.
    comprobar("::-webkit-details-marker" in html and "summary::marker" in html,
              "el triangulo del navegador se apaga en sus dos formas")
    comprobar(_regla(html, "details[open] > summary::before") is not None,
              "la flecha gira al abrirse: dice hacia donde va")
    pista = _regla(html, "details > summary::after")
    comprobar(pista is not None and "Desplegar" in pista,
              "y lo dice con todas las letras")
    abierta = _regla(html, "details[open] > summary::after")
    comprobar(abierta is not None and "Ocultar" in abierta,
              "que cambia cuando ya esta abierto")

    # Lo que se rompe en SILENCIO: la grafica de ODEPA se despliega al lado
    # con `position:absolute; left:100%` DENTRO de un <details>. Un
    # `overflow:hidden` la haria desaparecer sin ningun aviso.
    caja = _regla(html, "details")
    comprobar(caja is not None, "el desplegable tiene su recuadro")
    if caja:
        comprobar("overflow:hidden" not in caja,
                  "y NO recorta: la grafica de ODEPA se despliega fuera de el")


def _solo_marcado(html):
    """El HTML sin estilos, sin guion y sin comentarios.

    Hace falta para contar estructura: el CSS de la pagina trae un comentario
    que EXPLICA los `<details>` y lo nombra, asi que contar sobre el archivo
    crudo daba un desplegable abierto que no existe. La primera version de
    `_dentro_de_details` daba rojo con la pagina correcta por eso.
    """
    sin = re.sub(r"<style[\s\S]*?</style>", " ", html)
    sin = re.sub(r"<script[\s\S]*?</script>", " ", sin)
    return re.sub(r"<!--[\s\S]*?-->", " ", sin)


def _dentro_de_details(html, marca):
    """Si `marca` cae dentro de un <details> sin cerrar. None si no aparece.

    Existe por un defecto concreto: el campo «Especie» vivia DENTRO del
    desplegable «Generar desde 4 numeros», y el bloque de recomendaciones -que
    era un desplegable hermano- lo leia de ahi. Abriendo solo el segundo no
    habia donde escribir la especie, asi que no salia ninguna recomendacion y
    no se entendia por que. el usuario: *«no lo encuentro muy intuitivo»*.
    """
    marcado = _solo_marcado(html)
    corte = marcado.find(marca)
    if corte < 0:
        return None
    delante = marcado[:corte]
    return delante.count("<details") > delante.count("</details>")


def probar_partir_de_otra_curva(html):
    """Crear una curva: todo cuelga de UNA pregunta, la especie.

    el agrónomo pidio que el programa recomendara curvas parecidas; el usuario acoto
    que solo tiene sentido **entre variedades de la misma especie** -«no
    cerezos con arandanos ni kiwi con paltos»- y despues que **no era
    intuitivo**. Las dos correcciones eran ciertas.

    La segunda tenia una causa concreta y no de gusto: el campo de la especie
    estaba escondido dentro de OTRO desplegable, asi que el bloque que lo lee
    podia quedar sin nada que recomendar y sin forma de arreglarlo. Es el mismo
    error de §16.a.quadragies.ter. La comprobacion que mas vale de esta seccion
    es justamente esa: que la especie este SIEMPRE a la vista.
    """
    print("\n[17] Crear una curva partiendo de otra variedad")
    guion = guion_de(html)
    codigo = _solo_codigo(guion)

    # LA comprobacion de esta seccion.
    comprobar(_dentro_de_details(html, 'id="gEspecie"') is False,
              "el campo Especie esta SIEMPRE a la vista, fuera de todo <details>",
              "escondido, el bloque que lo lee se queda sin nada que recomendar")
    comprobar(_dentro_de_details(html, 'id="curvasDeLaEspecie"') is False,
              "y el bloque de recomendaciones tambien: sale solo, sin abrir nada")
    # Ya no van plegados en un <details>: son el paso 3 del asistente, y se
    # ven cuando toca escribirlos (2026-09-14).
    marcado_17 = _solo_marcado(html)
    comprobar(0 < marcado_17.find('data-paso="3"')
              < marcado_17.find('id="gImproductivos"')
              < marcado_17.find('data-paso="4"'),
              "y los 7 numeros viven en el paso 3: se ven cuando toca "
              "escribirlos, no antes")

    comprobar(re.search(r"\bfunction cargarBases\(", codigo)
              and re.search(r"\bfunction partirDeCurva\(", codigo),
              "existen las dos funciones")
    # No basta con que el campo tenga un oyente: tiene que REHACER el bloque.
    # `escucha` solo mira que exista un addEventListener, y un oyente que no
    # llama a `cargarBases` deja la recomendacion congelada en la especie
    # anterior — que es peor que no tenerla.
    comprobar(
        re.search(r'\$\("gEspecie"\)\.addEventListener\("input"[\s\S]{0,300}?'
                  r"cargarBases", guion) is not None,
        "escribir la especie rehace el bloque solo")

    # Solo se recomiendan variedades de la MISMA especie (lo que puso el usuario).
    consulta = re.search(r"(?m)^async function cargarBases\(\) \{[\s\S]*?^\}", guion)
    comprobar(consulta is not None, "se encontro el cuerpo de cargarBases")
    if consulta:
        cuerpo_c = consulta.group(0)
        comprobar("especie: especie" in cuerpo_c or 'set("especie"' in cuerpo_c,
                  "la especie viaja al servidor: es lo que acota que se recomienda")
        comprobar(re.search(r"\brecomendadas\.map\(opcion\)", cuerpo_c) is not None,
                  "las recomendadas se pintan aparte")
        comprobar(re.search(r"<details[\s\S]{0,500}?\botras\.map\(opcion\)", cuerpo_c)
                  is not None,
                  "y el resto del catalogo va PLEGADO, no al mismo nivel",
                  "mezclarlos presentaria un cerezo como recomendacion "
                  "para un arandano")
        comprobar("aviso_otras" in cuerpo_c,
                  "rotulado como que NO es una recomendacion")
        # Se dice en castellano CUANTAS hay, que es lo que la persona quiere
        # saber. Un titulo sobre una lista vacia no dice nada.
        comprobar("cuantas === 0" in cuerpo_c and "cuantas === 1" in cuerpo_c,
                  "dice cuantas curvas hay de esa especie: ninguna, una o varias")
        # Los dos numeros de afinado solo cuando pueden cambiar algo.
        comprobar(re.search(r"cuantas > 1[\s\S]{0,400}?pEntra", cuerpo_c) is not None,
                  "los campos de afinado salen SOLO con varias candidatas",
                  "con una sola no hay nada que ordenar")

    # Copiar la FORMA, no los precios.
    cuerpo = re.search(r"(?m)^function partirDeCurva\(id\) \{[\s\S]*?^\}", guion)
    comprobar(cuerpo is not None, "se encontro el cuerpo de partirDeCurva")
    if cuerpo:
        copiar = cuerpo.group(0)
        comprobar("rendimiento_kg_ha: p.rendimiento_kg_ha" in copiar,
                  "se copia la forma: los kg/ha por edad")
        comprobar("base.comerciales" not in copiar,
                  "y NO se heredan los precios de la curva de la que se parte",
                  "una variedad no vale lo mismo que otra")
        comprobar("base.costos" not in copiar, "ni los costos")
        comprobar("costo_directo_usd_ha: null" in copiar,
                  "ni el costo por edad de cada punto")
        comprobar("derivada_de: base.nombre" in copiar,
                  "y queda anotado de que curva salio")

    guardar = re.search(
        r'\$\("btnGuardarCurva"\)\.addEventListener\([\s\S]*?\n\}\);', guion)
    comprobar(guardar is not None, "se encontro el guardado")
    if guardar:
        comprobar("derivada_de" in guardar.group(0),
                  "de que curva salio viaja al guardar")

    comprobar('id="avisoDerivada"' in html and "avisoDerivada" in guion,
              "y se avisa en pantalla que la curva es prestada")

    # Desde la lista de pendientes: la recomendacion sale sola y se abre el
    # generador solo si NO hay de que partir. El estado ya lo dice.
    preparar = re.search(
        r"(?m)^function prepararCurvaNueva\(it, cliente\) \{[\s\S]*?^\}", guion)
    comprobar(preparar is not None,
              "la lista de pendientes deja la pantalla lista para esa curva")
    if preparar:
        cuerpo_p = preparar.group(0)
        comprobar('it.estado === "de_otra_variedad"' in cuerpo_p,
                  "y decide con el estado, sin preguntarle nada a nadie")
        comprobar("cargarBases()" in cuerpo_p,
                  "el bloque de recomendaciones se rehace")
        comprobar("irAPasoDeCurva(hayHermana ? 2 : 3" in cuerpo_p,
                  "y entra por el paso 2 si hay de que partir y por el 3 si no")
    comprobar(re.search(r"data-crear[\s\S]{0,200}?prepararCurvaNueva", guion)
              is not None,
              "el boton de la lista de pendientes es el que lo dispara")

    # Dos consultas que se pisan: ya paso dos veces en este proyecto.
    comprobar(re.search(r"let basesPedidas\s*=", codigo) is not None,
              "hay contador para que dos consultas no se pisen")
    comprobar(
        re.search(r"await r\.json\(\);\s*if \(mia !== basesPedidas\) return;",
                  codigo) is not None,
        "y la respuesta que llega tarde no pinta encima")


def probar_pasos_de_creacion(html):
    """Crear una curva es UNA linea de pasos, y manda todo lo que hace falta.

    el usuario, con capturas: *«la creacion de las curvas no esta clara, no esta
    ordenada; quizas deberiamos juntar el generar desde 4 numeros con el crear
    curva nueva para que asi el agronomo pueda seguir una linea de creacion»*. Y el
    motivo de fondo: *«esta funcionando bastante mal y no esta mostrando datos
    que son fundamentales»* -la curva generada salia con costo cero, precio
    cero y 0% de exportacion, o sea proyectando kilos sin ingresos-.

    Lo que mas vale de esta seccion es lo ultimo: **que los dos formularios que
    arman una curva -el generador y el «rehacer desde sus hitos» del editor-
    pidan lo mismo con las mismas palabras y manden el mismo campo**. Pedian la
    EDAD en que decrece uno y la DURACION de la meseta el otro, en casillas
    gemelas.
    """
    print("\n[18] Los pasos para crear una curva")
    guion = guion_de(html)
    marcado = _solo_marcado(html)

    # --- 1. es un informe agricola, no una tasacion (lo primero que pidio)
    comprobar(re.search(r"Informe agr[ií]cola", marcado) is not None,
              "la cabecera dice «Informe agrícola»")
    tasacion = re.findall(r"[Tt]asaci[oó]n\s+agr[ií]cola", marcado)
    comprobar(not tasacion,
              "y no queda «tasación agrícola» en ninguna parte de la página",
              "aparece {} vez/veces".format(len(tasacion)))

    # --- 2. los pasos: numerados, en orden, y en ese orden en la pantalla
    pasos = re.findall(r'class="paso-n">(\d)<', marcado)
    comprobar(pasos == ["1", "2", "3", "4"],
              "los cuatro pasos salen numerados y en orden", str(pasos))
    orden = ['id="gEspecie"', 'id="curvasDeLaEspecie"', 'id="btnNuevaCurva"',
             'data-paso="3"', 'id="gCostoPlena"', 'id="btnGenerarCurva"']
    donde = [marcado.find(m) for m in orden]
    comprobar(all(d > 0 for d in donde) and donde == sorted(donde),
              "y la pantalla los presenta en esa misma linea, de arriba abajo",
              str(list(zip(orden, donde))))
    # «Empezar con la tabla en blanco» era una entrada suelta ARRIBA de la
    # lista de curvas. Es una opcion del paso 2, no una puerta aparte.
    comprobar(marcado.find('id="btnNuevaCurva"')
              > marcado.find('id="curvasDeLaEspecie"'),
              "la tabla en blanco es una opción del paso 2, no otra entrada")

    # --- 3. el asistente: un paso a la vez, con «Atras» y «Siguiente»
    # el usuario: *«que pueda apretar siguiente y devolverse; siento que el diseno
    # que tiene ahora sigue siendo muy complicado de entender»*. Los cuatro
    # pasos ya estaban numerados; lo que faltaba era verlos de a uno.
    vistas = re.findall(r'class="paso-vista([^"]*)" data-paso="(\d)"', marcado)
    comprobar([v[1] for v in vistas] == ["1", "2", "3", "4"],
              "cada paso es una vista aparte, en orden",
              str([v[1] for v in vistas]))
    comprobar([bool(v[0].strip()) for v in vistas] == [False, True, True, True],
              "y solo la primera nace visible: al resto se llega con «Siguiente»",
              str([v[0] for v in vistas]))

    cambio = re.search(
        r"(?m)^function irAPasoDeCurva\(n, opciones\) \{[\s\S]*?^\}", guion)
    comprobar(cambio is not None, "existe el que cambia de paso")
    ir = cambio.group(0) if cambio else ""
    comprobar('$("btnPasoAtras").disabled = paso === 1' in ir,
              "«Atras» queda APAGADO en el primer paso, no escondido")
    comprobar('$("btnPasoSiguiente").classList.toggle("oculto", ultimo)' in ir
              and '$("btnGenerarCurva").classList.toggle("oculto", !ultimo)' in ir,
              "y en el ultimo el boton es «Generar la curva», no «Siguiente»")
    comprobar('$("pasoDeCuantos").textContent' in ir,
              "la pantalla dice en que paso va y de cuantos")

    siguiente = re.search(r'\$\("btnPasoSiguiente"\)\.addEventListener'
                          r"[\s\S]{0,400}?\n\}\);", guion)
    tramo = siguiente.group(0) if siguiente else ""
    comprobar(siguiente is not None and "faltaAlgoDelPaso" in tramo
              and "irAPasoDeCurva(paso + 1)" in tramo,
              "«Siguiente» comprueba lo que falta antes de avanzar")
    falta = re.search(r"(?m)^function faltaAlgoDelPaso\(paso\) \{[\s\S]*?^\}",
                      guion)
    comprobar(falta is not None and '"gEspecie"' in falta.group(0),
              "y lo unico que no se puede dejar para despues es la especie")
    atras = re.search(r'\$\("btnPasoAtras"\)\.addEventListener'
                      r"[\s\S]{0,200}?\n\}\);", guion)
    comprobar(atras is not None and "pasoDeCurva() - 1" in atras.group(0),
              "«Atras» vuelve al paso anterior")
    volver = re.search(r'\$\("pasosCurva"\)\.addEventListener'
                       r"[\s\S]{0,400}?\n\}\);", guion)
    comprobar(volver is not None and "destino < pasoDeCurva()" in volver.group(0),
              "se vuelve a un paso hecho con un clic, pero no se salta ninguno")
    comprobar(re.search(r'FOCO_DEL_PASO = \{[^}]*3: "gImproductivos"', guion)
              is not None,
              "y al llegar a un paso el cursor queda en su primer campo")

    # --- 4. las frases formales que pidio el usuario, en los DOS formularios
    for etiqueta, cuantos in (("Año que entra en plena producción", 2),
                              ("Año en que decrece la producción", 2)):
        comprobar(marcado.count(etiqueta) == cuantos,
                  "«{}» está en los dos formularios".format(etiqueta),
                  "aparece {}".format(marcado.count(etiqueta)))
    for viejo in ("Años a plena producción", "que aguanta la plena producción",
                  "Año de plena producción"):
        comprobar(viejo not in marcado,
                  "y ya no se dice «{}»".format(viejo))

    # --- 5. el costo por kg no se llama igual en dos sitios
    # Los campos del editor contra las columnas de SU tabla de edades: es donde
    # estaba el choque -«Costo por kg» arriba y «Costo por kg» en la columna
    # nueva- y donde el agronomo los ve a los dos al mismo tiempo.
    desde = marcado.find('id="detalleCurva"')
    hasta = marcado.find('id="tablaPuntos"')
    campos = {t.strip() for t in
              re.findall(r"<label[^>]*>([^<]+)</label>", marcado[desde:hasta])}
    columnas = {t.strip() for t in re.findall(
        r"<th[^>]*>([^<]+)</th>",
        marcado[hasta:marcado.find("</table>", hasta)])}
    comprobar("Costo variable (USD/kg)" in campos,
              "el campo se llama «Costo variable (USD/kg)»")
    comprobar(len(columnas) >= 4 and not (campos & columnas),
              "ninguna columna de la tabla se llama igual que un campo",
              " | ".join(sorted(campos & columnas)))

    # La tabla muestra la formula de §8.2 abierta en sus terminos, en el orden
    # en que se encadenan. Lo pidio el agrónomo: con un solo numero al final, una
    # curva mal calibrada no dice en cual de los dos terminos esta el error.
    esperadas = ["Edad (años)", "Rendimiento (kg/ha)",
                 "Costo base a esta edad (USD/ha)", "Costo/kg de fruta (USD/kg)",
                 "Costo/ha calculado (USD/ha)", "Costo final/kg (USD/kg)"]
    orden = [t.strip() for t in re.findall(
        r"<th[^>]*>([^<]+)</th>",
        marcado[hasta:marcado.find("</table>", hasta)])]
    comprobar(orden == esperadas,
              "las seis columnas de la tabla de edades, en orden", str(orden))

    # --- 6. lo que VIAJA: era el «no muestra datos que son fundamentales»
    gen = re.search(r'\$\("btnGenerarCurva"\)\.addEventListener'
                    r'[\s\S]*?fetch\("/curvas/generar"', guion)
    comprobar(gen is not None, "se encontró lo que manda el generador")
    cuerpo = gen.group(0) if gen else ""
    for campo, porque in (
            ("costo_base_usd_ha", "el costo base"),
            ("costo_en_plena_produccion_usd_ha", "el costo a plena producción"),
            ("pct_exportacion", "el % de exportación"),
            ("rnp_exportacion_usd_kg", "el precio de exportación"),
            ("rnp_mercado_interno_usd_kg", "el precio interno")):
        comprobar(campo in cuerpo, "viaja {}".format(porque))
    comprobar('$("gExport").value || 0) / 100' in cuerpo,
              "y el porcentaje viaja de 0 a 1, que es como lo guarda el modelo")
    # El costo variable NO se manda: lo deduce el servidor de los dos extremos.
    # Mandarlo desde aqui seria tener la misma resta en dos sitios (§2).
    comprobar("costo_por_kg_usd" not in cuerpo,
              "y el costo variable NO: lo deduce el servidor de los extremos")

    # --- 7. los dos formularios mandan el MISMO campo de declive
    rehacer = re.search(r'\$\("btnRehacerCurva"\)\.addEventListener'
                        r'[\s\S]*?fetch\("/curvas/generar"', guion)
    comprobar(rehacer is not None, "se encontró lo que manda el editor")
    otro = rehacer.group(0) if rehacer else ""
    # Sin los comentarios: los dos EXPLICAN que antes se mandaba el otro campo,
    # y el nombre viejo escrito en una explicacion no es el nombre viejo
    # viajando. Se busca en el codigo, no en la prosa.
    sin_notas = lambda t: re.sub(r"//[^\n]*", "", t)                # noqa: E731
    for donde, texto in (("el generador", sin_notas(cuerpo)),
                         ("el editor", sin_notas(otro))):
        comprobar("edad_inicio_declive" in texto,
                  "{} manda la EDAD en que decrece".format(donde))
        comprobar("anios_de_plena_produccion" not in texto,
                  "y no la duración de la meseta, que es otro número ({})"
                  .format(donde))

    # --- 8. la pista del costo en plena produccion, en el editor
    comprobar('id="pistaCostoPlena"' in marcado,
              "el editor muestra a cuánto sale la hectárea a plena producción")
    comprobar("pintarCostoEnPlena(baseGlobal, porKg)" in guion,
              "y se repinta desde la misma función que la columna: no se "
              "queda con un número viejo")

    # Las tres columnas calculadas salen de UNA función, la misma que ya se
    # dispara al escribir en la tabla y en los dos costos globales. Cada una
    # por su lado se quedaría con un número viejo sin que nada lo delate.
    pinta = re.search(r"function pintarCostoPorKg\([\s\S]*?\n\}", guion)
    comprobar(pinta is not None, "se encontró el pintado de las tres columnas")
    cuenta = pinta.group(0) if pinta else ""
    # La celda se busca en el cuerpo de `pintarPuntos`, que es quien arma la
    # fila. En todo el guion no sirve: la funcion que las rellena nombra las
    # mismas marcas para buscarlas, asi que borrar la celda dejaba la
    # comprobacion en verde. Lo destapo el negativo, no leerla.
    arma = re.search(r"function pintarPuntos\([\s\S]*?\.join\(\"\"\);", guion)
    comprobar(arma is not None, "se encontró el armado de las filas")
    filas_html = arma.group(0) if arma else ""
    for marca, que in (("data-costovar", "el costo/kg de fruta"),
                       ("data-costoha", "el costo/ha calculado"),
                       ("data-costokg", "el costo final/kg")):
        comprobar('{}="${{n}}"'.format(marca) in filas_html,
                  "cada fila tiene su celda para {}".format(que))
        comprobar(marca in cuenta,
                  "y la rellena la misma función que las otras ({})".format(que))
    # El costo/ha se escribe ANTES del corte por rendimiento cero: un huerto en
    # formación no produce y cuesta igual, y ese costo va al flujo.
    comprobar(0 <= cuenta.find("data-costoha") < cuenta.find("if (!rendimiento)"),
              "el costo/ha se muestra también en los años improductivos: "
              "no producen y cuestan igual")

    # --- 9. los cuatro defectos que apareceron probandolo, para que no vuelvan
    #
    # Ninguno lo habria visto una lectura: los cuatro se vieron manejando la
    # pantalla contra los routers de verdad.
    lectura = re.search(r"function rellenarHitos\([\s\S]*?\n\}", guion)
    comprobar(lectura is not None, "se encontró la lectura de los hitos")
    hitos = lectura.group(0) if lectura else ""
    # Los anios improductivos son la EDAD del primer punto que produce: un
    # huerto que da su primera cosecha a los 3 tuvo tres anios sin producir.
    # Con el `- 1`, rehacer una curva la adelantaba un anio entero.
    comprobar('valor("hImproductivos", productivos.length'
              " ? productivos[0].edad : 0)" in re.sub(r"\s+", " ", hitos),
              "los años improductivos se leen sin restarle uno: rehacer una "
              "curva no la adelanta un año")
    # Los tres del declive van juntos o no va ninguno: con uno puesto y los
    # otros dos vacios, rehacer sin tocar nada devolvia 422.
    comprobar(re.search(r'valor\("hAniosDePlena", hayDeclive', hitos)
              is not None,
              "y el declive sólo se lee si de verdad lo hay")

    preparar = re.search(r"function prepararCurvaNueva\([\s\S]*?\n\}", guion)
    comprobar(preparar is not None, "se encontró la llegada desde pendientes")
    llegada = preparar.group(0) if preparar else ""
    # `campo.value = x` no dispara ningun evento: sin esto el cuadro de ODEPA
    # se quedaba en la especie anterior, o sea mostrando el precio de otra
    # fruta con esta escrita al lado.
    comprobar('$("gEspecie").dispatchEvent(new Event("input"' in llegada,
              "llegando desde la lista de pendientes se avisa que la especie "
              "cambió: si no, lo de ODEPA se queda en la anterior")

    nueva = re.search(r'\$\("btnNuevaCurva"\)\.addEventListener[\s\S]*?\n\}\);',
                      guion)
    comprobar(nueva is not None, "se encontró la tabla en blanco")
    blanco = nueva.group(0) if nueva else ""
    comprobar('$("gEspecie").value.trim()' in blanco
              and '$("gVariedad").value.trim()' in blanco,
              "la tabla en blanco se lleva la especie del paso 1, no abre "
              "vacía")
    # Un 70% que nadie declaro se lee como un dato, y ademas hace que la
    # tarjeta no se marque «sin precios».
    for donde, texto in (("la tabla en blanco", blanco),
                         ("partir de otra curva",
                          re.search(r"function partirDeCurva\([\s\S]*?\n\}",
                                    guion).group(0))):
        comprobar("pct_exportacion: 0," in texto,
                  "y el % de exportación arranca en cero ({})".format(donde))

    guardar = re.search(r'\$\("btnGuardarCurva"\)\.addEventListener'
                        r'[\s\S]*?fetch\("/curvas",', guion)
    comprobar(guardar is not None, "se encontró el guardado")
    salva = guardar.group(0) if guardar else ""
    comprobar('fetch("/curvas/cobertura"' in salva,
              "antes de guardar una curva nueva se pregunta si ya hay una de "
              "esa especie y variedad")
    comprobar('estado.estado === "propia"' in salva,
              "y lo decide el emparejador del informe, no comparando textos")


def probar_verificacion_completa(html):
    """19. Lo que encontro la verificacion completa del 2026-09-11.

    Cada comprobacion es un defecto que aparecio MANEJANDO la pagina contra el
    servidor de verdad, con capturas: ninguno se veia leyendo el codigo.
    """
    print("\n[19] Lo que encontro la verificacion completa")
    guion = guion_de(html)

    # --- las selecciones de la visita se guardan
    # El desplegable propio solo dispara «change». Con el oyente puesto nada
    # mas en «input», elegir una calificacion o un estado no quedaba entre lo
    # tocado, y el guardado lo ignoraba sin avisar: la opinion nunca llegaba.
    for evento in ("input", "change"):
        comprobar('document.addEventListener("{}", marcarTocado, true)'
                  .format(evento) in guion,
                  "lo tocado en la visita se anota también con «{}»".format(evento))
    marcar = re.search(r"function marcarTocado\([\s\S]*?\n\}", guion)
    comprobar(marcar is not None and "sinRevisar.delete(clave)" in marcar.group(0),
              "y elegir en un desplegable da por revisado el dato importado")

    # --- soltar el informe deja la fecha en blanco tambien a la vista
    empezar = re.search(r'\$\("btnEmpezarDeCero"\)\.addEventListener[\s\S]*?\n\}\);',
                        guion)
    comprobar(empezar is not None and 'valor(id, "")' in empezar.group(0),
              "empezar otro informe limpia los campos con valor(): el calendario "
              "de la fecha se entera y no sigue mostrando la del informe anterior")

    # --- «que falta» nombra las secciones por su titulo tambien tras recargar
    # (la logica vive en calcularFaltantesInforme desde que se separo de
    # avisarQueFaltaEnElInforme para que la confirmacion de "generar igual"
    # use la misma lista sin pintar nada dos veces)
    falta = re.search(r"async function calcularFaltantesInforme\([\s\S]*?\n\}", guion)
    comprobar(falta is not None and 'fetch("/visita/formulario")' in falta.group(0),
              "«qué falta» carga el formulario si todavía no está: sin eso "
              "nombraba las secciones por su id interno")

    # --- lo importado sin revisar se ve tambien en los desplegables
    comprobar(re.search(r"\.selector:has\(> select\.sin-revisar\) \.selector-boton"
                        r"\s*\{[^}]*--aviso", html) is not None,
              "un desplegable con un dato importado sin revisar se ve marcado")

    # --- los dialogos que no borran no preguntan si se quiere borrar
    llamadas = re.findall(r"confirmarBorrado\(\{[\s\S]*?\}\)", guion)
    for boton, no_es_borrar in (("Soltarlo", True), ("Traspasar", True),
                                ("Quitarle el acceso", False), ("Rehacer", False)):
        propia = [c for c in llamadas if 'boton: "{}"'.format(boton) in c]
        comprobar(len(propia) == 1 and "titulo:" in propia[0],
                  "el diálogo de «{}» hace su propia pregunta, no «¿borrar?»"
                  .format(boton))
        if no_es_borrar:
            comprobar(len(propia) == 1 and "peligro: false" in propia[0],
                      "y «{}» no se pinta con el tacho de basura".format(boton))

    # --- mensajes que decian algo que no era
    comprobar("escritura(s) leída(s)" not in guion
              and "documento(s) cargado(s)" in guion,
              "el resumen cuenta documentos cargados, no «leídos»: sin IA no se "
              "lee ninguno")
    comprobar("Ningún informe coincide con la búsqueda." in guion
              and '$("buscarExp").value.trim() || $("filtroEstado").value' in guion,
              "Mis informes distingue «todavía no hay» de «no coincide con la "
              "búsqueda»")
    comprobar("No se detectaron derechos de agua" not in guion,
              "sin documentos leídos no se afirma que «no se detectaron» derechos "
              "de agua")
    comprobar(re.search(r"buscado\s*\?\s*'<p class=\"pasos\">Ninguna curva calza",
                        guion) is not None,
              "con solo la casilla de incompletas no se habla de «búsqueda»")
    comprobar("«Armarla desde 4 números»" not in guion,
              "ningún aviso remite a un botón que ya no se llama así")

    # --- presentacion
    comprobar("#listaExpedientes .insignia { white-space:nowrap; }" in html,
              "las insignias de estado no se parten en dos renglones")
    comprobar(re.search(r"edadMax - edadesEje\[edadesEje\.length - 1\] < 3\)\s*\{"
                        r"\s*edadesEje\.pop\(\);", guion) is not None,
              "la última edad del eje no se monta sobre la marca anterior")
    comprobar("num(p.factor_total, 3)" in guion and "num(p.variacion_pct, 1)" in guion,
              "la simulación muestra el factor con coma decimal")
    comprobar("de manual" not in guion,
              "y ya no habla de la «curva de manual»")


def probar_buscador_de_especies(html):
    """El campo Especie ofrece nombres mientras se escribe, con el diseno de la
    pagina.

    el usuario: *«que cuando se ponga una especie sea como un buscador predictivo...
    como cuando uno busca algo en Mercado Libre y te ofrece cosas similares»*,
    y que se vea como el calendario y los botones.

    Lo que se rompe sin hacer ruido: que vuelva el `<datalist>` -Chrome dibuja
    SU lista encima de la nuestra-, que Enter elija una opcion que el usuario
    no habia visto -cambiaria lo escrito en silencio-, que apretar una opcion
    le quite el foco al campo antes del clic -la lista se cierra sin elegir-, y
    que una respuesta vieja pise a la nueva.
    """
    print("\n[20] El buscador del campo Especie")
    guion = guion_de(html)
    codigo = _solo_codigo(guion)
    marcado = _solo_marcado(html)

    # Se busca el cuerpo de la FUNCION, no el guion entero: `mio !== pedido`
    # o un `blur` cualquiera en otro sitio servian de coartada.
    m = re.search(r"function autocompletarEspecie\(id\) \{([\s\S]*?)\n\}\n", guion)
    comprobar(m is not None, "existe autocompletarEspecie")
    cuerpo = m.group(1) if m else ""
    cuerpo_codigo = _solo_codigo(cuerpo)

    # --- el buscador del sistema no puede volver -------------------------
    comprobar("<datalist" not in marcado,
              "no hay ningun <datalist> en la pagina")
    comprobar(not re.search(r"<input[^>]*\blist=", marcado),
              "y ningun campo apunta a uno con `list=`")
    comprobar("especiesOdepa" not in guion
              and "cargarEspeciesDeOdepa" not in guion,
              "ni queda el codigo que llenaba la lista vieja")

    # --- quien lo usa ----------------------------------------------------
    comprobar(re.search(r'(?m)^autocompletarEspecie\("gEspecie"\);', guion)
              is not None,
              "el campo del generador se arranca al cargar la pagina")
    comprobar(re.search(r'(?m)^autocompletarEspecie\("cEspecie"\);', guion)
              is not None,
              "y el del editor de la curva tambien")
    comprobar("entrada.dataset.autocompleta" in cuerpo,
              "arrancarlo dos veces no envuelve el campo dos veces")

    # --- de donde salen los nombres: del servidor ------------------------
    comprobar('"/referencias?cuantas=8&parecidas=" + encodeURIComponent(texto)'
              in cuerpo,
              "las sugerencias las pide al servidor, que ya ordena como un "
              "buscador, y pide 8")

    # --- las tres trampas de un desplegable propio ------------------------
    comprobar("mio !== pedido" in cuerpo,
              "solo pinta la ultima consulta")
    comprobar(re.search(r'lista\.addEventListener\("mousedown", \(e\) => '
                        r'e\.preventDefault\(\)\)', cuerpo) is not None,
              "apretar una opcion no le quita el foco al campo antes del clic")
    comprobar(re.search(r'entrada\.addEventListener\("blur", cerrar\)', cuerpo)
              is not None,
              "al salir del campo la lista se cierra")
    comprobar(cuerpo_codigo.count("document.activeElement !== entrada") >= 2,
              "un input disparado por codigo (sin foco) no abre la lista ni "
              "gasta una consulta")

    # --- el teclado ------------------------------------------------------
    for tecla in ("ArrowDown", "ArrowUp", "Escape", "Enter"):
        comprobar('"{}"'.format(tecla) in cuerpo,
                  "responde a {}".format(tecla))
    comprobar(re.search(r'e\.key === "Enter" && lista && activa >= 0', cuerpo)
              is not None,
              "Enter solo elige si el usuario se paro en una opcion")

    # --- elegir avisa al resto de la pagina ------------------------------
    comprobar(re.search(r'entrada\.dispatchEvent\(new Event\("input"', cuerpo)
              is not None,
              "elegir dispara `input`: de la especie cuelgan el cuadro de "
              "ODEPA y las curvas que ya hay")
    comprobar('new Event("change"' in cuerpo,
              "y `change`")
    comprobar("eligiendo" in cuerpo_codigo and cuerpo_codigo.count("eligiendo") >= 3,
              "sin que ese `input` reabra la lista")

    # --- el diseno: es la MISMA lista que la de los desplegables ---------
    comprobar('lista.className = "selector-lista"' in cuerpo,
              "reusa `.selector-lista`: se ve hermano de los desplegables")
    caja = _regla(html, ".autocompletar")
    comprobar(caja is not None and "position:relative" in caja.replace(" ", ""),
              "la caja que envuelve al campo ancla la lista")
    arriba = _regla(html, ".selector-lista.arriba")
    comprobar(arriba is not None and "bottom:" in arriba,
              "y se da vuelta cuando abajo no cabe")
    resaltado = _regla(html, ".selector-lista li .txt b")
    comprobar(resaltado is not None and "var(--lima)" in resaltado,
              "lo escrito va resaltado en el lima de la pagina")
    comprobar("aria-expanded" in cuerpo and "combobox" in cuerpo,
              "el campo se declara combobox, para quien no usa el raton")
    comprobar(re.search(r"(?m)^  input, select, textarea \{", html) is not None
              and 'class="lupa"' in guion,
              "las filas llevan la lupa, y el campo conserva el estilo de "
              "todos los campos")


def probar_consulta_de_anios(html):
    """Consultar produccion y costo de un rango de anios, sin tabular cada uno.

    el usuario: *«que pueda escoger cuantos anios quiere ver en la tabla... que sea
    opcional... que pueda consultar de que anio a que anio cuanto produce»*.

    Es de SOLO LECTURA: interpola con `/curvas/serie` -el MISMO endpoint del
    grafico (§16.a.septies)- en vez de reimplementar la regla, y sobre lo que
    hay EN PANTALLA en ese momento (`recogerPuntos()`), no lo ultimo guardado.
    """
    print("\n[21] Consultar produccion y costo año por año")
    guion = guion_de(html)
    codigo = _solo_codigo(guion)
    marcado = _solo_marcado(html)

    # --- opcional y plegado por defecto -----------------------------------
    bloque = re.search(
        r'<details id="detalleConsultaAnios"[^>]*>([\s\S]*?)</details>', marcado)
    comprobar(bloque is not None, "existe el bloque de la consulta")
    comprobar(bloque is not None and "open" not in re.search(
        r'<details id="detalleConsultaAnios"[^>]*>', marcado).group(0),
              "arranca PLEGADO: es opcional, no un paso del flujo")
    cuerpo = bloque.group(1) if bloque else ""
    comprobar('id="qAnioDesde"' in cuerpo and 'id="qAnioHasta"' in cuerpo,
              "tiene los dos campos: desde y hasta")
    comprobar('id="consultaAnios"' in cuerpo, "y donde pintar el resultado")

    # --- consultarAnios(): el cuerpo de la funcion, no el nombre suelto ---
    m = re.search(r"async function consultarAnios\(\) \{([\s\S]*?)\n\}\n", guion)
    comprobar(m is not None, "existe consultarAnios")
    adentro = m.group(1) if m else ""
    adentro_codigo = _solo_codigo(adentro)

    comprobar(re.search(r"if \(!detalle \|\| !detalle\.open\) return;", adentro)
              is not None,
              "plegado, no gasta ninguna consulta")
    comprobar('fetch("/curvas/serie"' in adentro,
              "interpola con el MISMO endpoint que el gráfico, no otro nuevo")
    comprobar("curvaParaConsultar()" in adentro,
              "arma la curva a consultar con una función propia")
    comprobar("recogerPuntos()" in _solo_codigo(
        re.search(r"function curvaParaConsultar\(\) \{([\s\S]*?)\n\}\n",
                  guion).group(1)),
              "y esa curva usa lo que hay EN PANTALLA, no `curvaActual.puntos` "
              "a secas -que puede estar desactualizado-")
    comprobar("loQueFaltaParaDibujar(curva)" in adentro,
              "reusa el mismo aviso que el gráfico -«escriba la especie»-, "
              "no otra version del mismo mensaje")
    # La tabla que arma esta función es la que ve el tasador: no debe traer
    # <input>, o se leería como otra tabla editable igual a la de arriba.
    comprobar("<input" not in adentro,
              "la tabla que pinta es de solo lectura, sin ningún campo")

    # --- solo pinta la ultima consulta: mismo patron que el resto de la pagina
    comprobar("++pedidoDeConsulta" in adentro_codigo
              and "mio !== pedidoDeConsulta" in adentro_codigo,
              "una respuesta que llega tarde no pisa a la que llegó después")

    # --- el rango: se acota a lo que la curva realmente tiene -------------
    comprobar("Math.min(desde, maxima)" in adentro_codigo
              and "Math.min(hasta, maxima)" in adentro_codigo,
              "el año pedido se recorta al que la curva realmente alcanza")
    comprobar(re.search(r"desdeClamp > hastaClamp", adentro) is not None,
              "y si el inicio queda después del término, avisa en vez de "
              "mostrar una tabla vacía sin explicación")

    # --- se dispara solo, con la misma disciplina que el resto de la pagina
    comprobar(re.search(r'\$\("detalleConsultaAnios"\)\.addEventListener\('
                        r'"toggle"', guion) is not None,
              "al abrirla por primera vez, consulta sola -no hay que apretar "
              "nada más-")
    comprobar(re.search(r'function pedirConsultaAnios\(\) \{\s*'
                        r'clearTimeout\(esperaDeConsulta\);\s*'
                        r'esperaDeConsulta = setTimeout\(consultarAnios',
                        guion) is not None,
              "espera a que se deje de escribir, no consulta en cada tecla")
    forq = re.search(
        r'\["qAnioDesde", "qAnioHasta"\]\.forEach\([\s\S]{0,150}?'
        r'addEventListener\("input", pedirConsultaAnios\)', guion)
    comprobar(forq is not None,
              "«qAnioDesde» y «qAnioHasta» disparan la consulta al escribir")

    # --- cubre los cuatro caminos que repintan la tabla de arriba ---------
    pintar = re.search(r"function pintarPuntos\(puntos\) \{([\s\S]*?)\n\}\n",
                       guion)
    comprobar(pintar is not None and "pedirConsultaAnios();" in pintar.group(1),
              "abrir la curva, agregar/quitar una edad y rehacer desde los "
              "hitos actualizan la consulta -los cuatro pasan por aquí-")

    # --- cambiar de curva no deja pegado el rango de la anterior ----------
    curva_pintada = re.search(r"function pintarCurva\(c\) \{([\s\S]*?)\n\}\n",
                              guion)
    comprobar(curva_pintada is not None, "se encontró pintarCurva")
    cuerpo_pc = curva_pintada.group(1) if curva_pintada else ""
    orden = [cuerpo_pc.find('valor("qAnioDesde"'),
             cuerpo_pc.find('valor("qAnioHasta"'),
             cuerpo_pc.find("pintarPuntos(c.puntos)")]
    comprobar(all(o >= 0 for o in orden) and orden == sorted(orden),
              "el rango se reseedea ANTES de repintar la tabla -si no, la "
              "consulta dispara con el año de la curva vieja todavía puesto-",
              str(orden))

    # --- el choque de nombres que ya costó una vuelta ----------------------
    # `.alta` ya la usa `.insignia.alta` para el verde de «confianza alta»;
    # una tabla con esa misma clase heredaba el color sin que ninguna regla
    # lo dijera a propósito -los números salían en lima, no en el color de
    # tinta normal-. Se verifica el nombre elegido Y que el viejo no vuelva.
    comprobar('"tabla-envoltura desplazable"' in guion,
              "la tabla usa una clase que no choca con `.insignia.alta`")
    comprobar('"tabla-envoltura alta"' not in guion,
              "y no volvió el nombre que se prestaba el color sin querer")
    regla = _regla(html, ".tabla-envoltura.desplazable")
    comprobar(regla is not None and "color" not in regla,
              "la regla de la tabla no declara color: el que se vea no es "
              "un accidente de cascada")


def probar_panel_de_exclusiones(html):
    """El checklist de «qué sacar del informe», dentro del diálogo de vista
    previa.

    el usuario: *«en vez de hacer distintos tipos de informes... que lo eliminen
    antes de descargar el excel nomas»*. Primera version: panel fijo al
    costado del formulario (`position:sticky`). Segunda version: un botón
    flotante que abría un cajón aparte, con `right` calculado para un ancho
    de pantalla especifico -en un monitor mas ancho, el usuario: *"el botón esta
    como para el otro lado... no armoniza con el diseño"*-, y encima
    desconectado de la vista previa: *"quería que en la vista previa la
    persona pudiera seleccionar qué eliminar... como un excel interactivo"*
    (2026-09-30).

    Tercera version, esta: el checklist se arma DENTRO del diálogo de
    «Vista previa» (`mostrarVistaPrevia`), al lado de las páginas reales del
    informe -un solo lugar, no dos-, así que el botón flotante y su cálculo
    de posición desaparecen enteros en vez de corregirse una vez más.

    Lo que importa verificar acá no es el registro (eso ya lo prueba
    scripts/test_exclusiones.py, contra el motor real): es que el botón
    flotante viejo no volvió, que la cascada de dependencias en JS haga
    EXACTAMENTE lo mismo que services/exclusiones.py:resolver_cascada
    -tildar una que tiene dependientes los marca y los bloquea, con el
    motivo a la vista-, y que lo que se manda a generar sea siempre el
    resultado YA con la cascada aplicada.
    """
    print("\n[22] El panel de «qué sacar del informe»")
    guion = guion_de(html)
    codigo = _solo_codigo(guion)
    marcado = _solo_marcado(html)

    # --- el boton/cajon flotante de antes no volvio ------------------------
    for viejo in ("gatillo-exclusiones", "btnAbrirExclusiones", "veloExclusiones",
                  "panelExclusiones", "btnCerrarExclusiones", "abrirExclusiones",
                  "cerrarExclusiones"):
        comprobar(viejo not in guion,
                  "no volvió '{}' -el checklist vive en el diálogo de vista "
                  "previa, no en un elemento aparte con su propio "
                  "abrir/cerrar".format(viejo))

    # --- ahora vive dentro de mostrarVistaPrevia, no en el marcado fijo ----
    comprobar('id="listaExclusiones"' in guion,
              "el contenedor del checklist existe en el HTML que arma el "
              "diálogo de vista previa")
    comprobar('id="listaExclusiones"' not in marcado,
              "y NO en el marcado estático de la página: si estuviera ahí "
              "existiría aunque nunca se abra la vista previa, sin nada que "
              "lo pinte")
    m = re.search(r"function mostrarVistaPrevia\(([\s\S]*?)\n\}\n", guion)
    comprobar(m is not None and 'id="listaExclusiones"' in m.group(1),
              "el checklist se arma como parte del marcado del diálogo")
    comprobar(m is not None and "pintarExclusiones()" in m.group(1),
              "y se pinta apenas se abre, con lo que ya se cargó al "
              "arrancar la página -no hace falta ir a buscarlo de nuevo")
    comprobar(m is not None and "btnActualizarVistaPrevia" in m.group(1),
              "el diálogo trae su propio botón para volver a renderizar "
              "las páginas con la selección al día")

    # --- se carga sola al entrar, no hay que ir a buscarla ----------------
    comprobar(re.search(r"arrancarLaPagina\(\) \{[\s\S]{0,400}?"
                        r"cargarUnidadesExcluibles\(\)", guion) is not None,
              "se pide la lista al arrancar la página, sin que el tasador "
              "tenga que hacer nada")
    comprobar('fetch("/informes/unidades-excluibles")' in guion,
              "contra el endpoint real del registro")

    # --- calcularCascada: el mismo algoritmo que resolver_cascada ---------
    m = re.search(r"function calcularCascada\(pedidas\) \{([\s\S]*?)\n\}\n", guion)
    comprobar(m is not None, "existe calcularCascada")
    adentro = _solo_codigo(m.group(1)) if m else ""
    comprobar("u.depende_de" in adentro,
              "arma el grafo desde 'depende_de', tal como lo declara el "
              "registro -no una copia de las dependencias a mano en la "
              "pantalla")
    comprobar("resultado.add(actual)" in adentro and "pendientes.push(dep)" in adentro,
              "recorre el cierre transitivo -si el arrastrado tiene a su vez "
              "otro dependiente, tambien se arrastra ese")

    # --- pintarExclusiones: lo arrastrado se ve, y se explica por que -----
    m = re.search(r"function pintarExclusiones\(\) \{([\s\S]*?)\n\}\n", guion)
    comprobar(m is not None, "existe pintarExclusiones")
    adentro_pintar = m.group(1) if m else ""
    comprobar('disabled' in adentro_pintar,
              "lo arrastrado queda BLOQUEADO -no se puede destildar por su "
              "cuenta sin destildar primero de que depende")
    comprobar("Se excluye junto con" in adentro_pintar,
              "y con una nota que dice de QUE depende, no solo que esta "
              "bloqueado")
    comprobar("marcarVistaPreviaDesactualizada()" in adentro_pintar,
              "tildar una casilla avisa que la vista previa ya abierta "
              "quedó vieja, en vez de fingir que sigue mostrando lo mismo")

    # --- lo que se manda a generar es la cascada, no lo tildado a mano ----
    # Vive en armarPeticionGeneracion(), no en el manejador de btnGenerar: la
    # comparte con btnVistaPrevia (ver probar_vista_previa_en_pantalla), asi
    # que buscarla dentro del click de btnGenerar ya no la encontraria.
    armar = re.search(r"function armarPeticionGeneracion\(\) \{([\s\S]*?)\n\}\n",
                      guion)
    comprobar(armar is not None, "existe armarPeticionGeneracion")
    cuerpo_armar = armar.group(1) if armar else ""
    comprobar("calcularCascada([...excluidasElegidas]).resultado" in cuerpo_armar,
              "el payload manda la cascada YA resuelta -si el servidor "
              "recibiera solo lo tildado a mano, tendria que adivinar el "
              "arrastre por su cuenta")
    comprobar("[...excluidasElegidas]" in cuerpo_armar,
              "no manda el Set tal cual -JSON.stringify de un Set da '{}'-")


def probar_vista_previa_en_pantalla(html):
    """El botón «Vista previa»: renderiza el informe con el Excel real y lo
    muestra en la página, sin descargar nada -y, desde el 2026-09-30, con el
    checklist de qué sacar al lado (ver probar_panel_de_exclusiones).

    el usuario: *"no se puede hacer que se vea como una preview del informe sin
    tener que descargarlo?... que se vea en la pagina... y que si pongo
    incluir tal y tal cosa poder ver como se veria"*, y más tarde: *"quería
    que en la vista previa la persona pudiera seleccionar qué eliminar"*.

    Lo que importa verificar acá no es el render -eso lo prueba
    scripts/test_renderizado.py contra el Excel real-: es que el botón
    exista y esté cableado, que TANTO abrir la vista previa COMO el botón
    «Actualizar» de adentro del diálogo pasen por el mismo `pedirVistaPrevia`
    (para que nunca puedan mostrar una selección distinta de la que después
    se descarga), y que el visor sea un dialogo modal correcto que no se
    trague los avisos del motor.
    """
    print("\n[23] El botón «Vista previa»")
    guion = guion_de(html)
    codigo = _solo_codigo(guion)
    marcado = _solo_marcado(html)

    comprobar('id="btnVistaPrevia"' in marcado
              and 'id="progresoVistaPrevia"' in marcado
              and 'id="barraVistaPrevia"' in marcado,
              "existen el botón y su barra de progreso propia")

    comprobar('$("btnVistaPrevia").addEventListener("click"' in guion,
              "el botón está cableado a un manejador")

    boton = re.search(
        r'\$\("btnVistaPrevia"\)\.addEventListener\("click",'
        r' async \(\) => \{([\s\S]*?)\n\}\);', guion)
    comprobar(boton is not None, "se encontró el manejador")
    cuerpo_boton = boton.group(1) if boton else ""
    comprobar("pedirVistaPrevia()" in cuerpo_boton,
              "pide la vista previa por el mismo camino que después usa "
              "«Actualizar vista previa» dentro del diálogo -no hay dos "
              "formas distintas de armar el mismo pedido")
    comprobar("mostrarVistaPrevia(" in cuerpo_boton,
              "y pinta el resultado en un visor, no solo lo descarta")

    # --- pedirVistaPrevia: el UNICO lugar que arma el payload y llama al
    # endpoint -tanto el botón de afuera como «Actualizar» de adentro del
    # diálogo pasan por acá, asi que nunca pueden mostrar selecciones
    # distintas entre si ni contra lo que después descarga «Generar»
    m = re.search(r"async function pedirVistaPrevia\(\) \{([\s\S]*?)\n\}\n", guion)
    comprobar(m is not None, "existe pedirVistaPrevia")
    cuerpo_pedir = m.group(1) if m else ""
    comprobar("armarPeticionGeneracion()" in cuerpo_pedir,
              "usa el MISMO armado de payload que «Generar» -no puede "
              "mostrar una selección distinta de la que después se descarga")
    comprobar('fetch("/asistente/vista-previa"' in cuerpo_pedir,
              "llama al endpoint correcto")

    # --- el visor: dialogo modal correcto, mismo patron que confirmarBorrado
    m = re.search(r"function mostrarVistaPrevia\(([\s\S]*?)\n\}\n", guion)
    comprobar(m is not None, "existe mostrarVistaPrevia")
    adentro = m.group(1) if m else ""
    comprobar('role="dialog"' in adentro and 'aria-modal="true"' in adentro,
              "el visor se anuncia como un dialogo modal, no un div suelto")
    comprobar('e.key === "Escape"' in adentro and "cerrar()" in adentro,
              "cierra con Escape, mismo patrón que el resto de los modales")
    comprobar("velo.addEventListener(\"click\"" in adentro,
              "y clickeando afuera, en el velo")

    # --- las paginas (y sus avisos) se pintan en su propia funcion, porque
    # tanto abrir como actualizar necesitan repintarlas sin rearmar todo el
    # dialogo -perderia el scroll del checklist al lado-
    m = re.search(r"function pintarPaginasVistaPrevia\(([\s\S]*?)\n\}\n", guion)
    comprobar(m is not None, "existe pintarPaginasVistaPrevia")
    adentro_paginas = m.group(1) if m else ""
    comprobar("nota(" in adentro_paginas,
              "los avisos del motor se pintan -una vista previa con, por "
              "ejemplo, una curva faltante tiene que avisarlo igual que el "
              "informe final")

    # --- actualizar sin cerrar el dialogo: mismo dato, sin perder contexto -
    m = re.search(
        r"async function actualizarVistaPreviaAbierta\(\) \{([\s\S]*?)\n\}\n",
        guion)
    comprobar(m is not None, "existe actualizarVistaPreviaAbierta")
    adentro_actualizar = m.group(1) if m else ""
    comprobar("pedirVistaPrevia()" in adentro_actualizar
              and "pintarPaginasVistaPrevia(" in adentro_actualizar,
              "el botón de adentro del diálogo vuelve a pedir y a pintar "
              "solo las páginas -no cierra ni reabre el diálogo entero, que "
              "perdería el scroll y el foco del tasador en el checklist")


def probar_actualizador_rabobank(html):
    """El actualizador de informes (retasaciones Rabobank): 9 listas
    editables con el mismo patron generico (tablaEditable), el armado del
    payload, y la estadistica de testigos que se adelanta en pantalla.

    el usuario: "quiero que sea un apartado nuevo como lo de visita de terreno".
    """
    print("\n[24] El actualizador de informes (Rabobank)")
    guion = guion_de(html)
    codigo = _solo_codigo(guion)
    marcado = _solo_marcado(html)

    comprobar('data-panel="rabobank"' in marcado, "existe el boton de navegacion")
    comprobar('id="panel-rabobank"' in marcado, "y su panel")

    # --- las nueve listas usan el mismo editor generico, no nueve copias --
    for tabla_id in ("tablaRSuelos", "tablaRAguas", "tablaRConstrH", "tablaRConstrP",
                     "tablaRInstalaciones", "tablaRPlantSin", "tablaRPlantCon",
                     "tablaRTestigosPropiedades", "tablaRTestigosAgua"):
        comprobar('id="{}"'.format(tabla_id) in marcado,
                  "existe el contenedor de {}".format(tabla_id))
        comprobar("tablaEditable(\"{}\"".format(tabla_id) in guion,
                  "{} usa el editor generico tablaEditable, no una copia propia"
                  .format(tabla_id))

    comprobar(guion.count("function tablaEditable(") == 1,
              "tablaEditable esta declarada UNA sola vez -reusada 9 veces, no "
              "copiada 9 veces")

    # --- el payload: nombres de campo tienen que calzar con el schema Pydantic
    armar = re.search(r"function armarRetasacion\(\) \{([\s\S]*?)\n\}\n", guion)
    comprobar(armar is not None, "existe armarRetasacion")
    cuerpo_armar = armar.group(1) if armar else ""
    for clave in ("parametros", "cliente", "predio", "suelos", "aguas",
                 "construcciones_hipotecables", "construcciones_prendables",
                 "instalaciones", "plantaciones_sin_royalty",
                 "plantaciones_con_royalty", "tasador", "diferencias",
                 "testigos_propiedades", "testigos_derechos_agua"):
        comprobar("{}:".format(clave) in cuerpo_armar or '"{}"'.format(clave) in cuerpo_armar,
                  "el payload manda '{}' -mismo nombre que RetasacionRabobank"
                  .format(clave))

    # --- el boton de generar: mismo patron que btnGenerar/btnVistaPrevia ---
    comprobar('$("btnGenerarRabobank").addEventListener("click"' in guion,
              "el boton de generar esta cableado")
    generar = re.search(
        r'\$\("btnGenerarRabobank"\)\.addEventListener\("click", async \(\) => \{'
        r'([\s\S]*?)\n\}\);', guion)
    comprobar(generar is not None, "se encontro el manejador")
    cuerpo_generar = generar.group(1) if generar else ""
    comprobar('fetch("/retasaciones/generar"' in cuerpo_generar,
              "llama al endpoint correcto")
    comprobar("armarRetasacion()" in cuerpo_generar,
              "manda el payload armado por armarRetasacion, no uno aparte")
    comprobar("avisarDeFallo(" in cuerpo_generar,
              "un fallo del servidor se avisa, no se pierde en silencio")

    # --- generar con datos faltantes (2026-10-02): se pregunta, no se rechaza
    i_revisar = cuerpo_generar.find("revisarRetasacion(")
    i_confirmar = cuerpo_generar.find("confirmarGenerarConFaltantes(")
    i_deshabilitar = cuerpo_generar.find("boton.disabled = true")
    comprobar(-1 < i_revisar < i_confirmar < i_deshabilitar,
              "antes de generar consulta qué falta y pregunta «¿generar con "
              "esto pendiente?», ANTES de deshabilitar el botón -cancelar no "
              "deja la pantalla a medias-")
    comprobar('fetch("/retasaciones/revisar"' in guion,
              "la lista de lo que falta la decide el servidor, no una copia "
              "de las reglas en la pantalla")
    comprobar("salida.pendientes" in cuerpo_generar,
              "después de generar sigue a la vista lo que quedó pendiente")
    comprobar('Number($("rUfBase").value || 0)' not in cuerpo_armar
              and 'numeroONulo("rUfBase")' in cuerpo_armar,
              "una UF en blanco viaja como null, no como 0 -el servidor "
              "distingue «falta» de «no vale»")

    # --- el formato completo del banco (2026-10-03): cada cuadro de texto y
    # cada dato del formato tiene su campo, y viaja con el nombre del esquema
    for clave in ("observaciones_portada", "legal_observaciones",
                  "legal_observaciones_expropiacion", "aguas_observaciones",
                  "construcciones_observaciones", "plantaciones_observaciones",
                  "testigos_propiedades_analisis", "testigos_derechos_agua_analisis",
                  "valorar_plantaciones", "valorar_instalaciones",
                  "tiene_tasacion_anterior", "nombre_corto", "deslinde_norte",
                  "hipoteca_rabofinance", "deuda_contribuciones_clp",
                  "afecto_expropiacion", "descripcion"):
        comprobar("{}:".format(clave) in cuerpo_armar,
                  "el payload manda '{}'".format(clave))
    poblar = re.search(r"function poblarFormularioRabobank\(datos\) \{([\s\S]*?)\n  \}\n", guion)
    cuerpo_poblar = poblar.group(1) if poblar else ""
    comprobar("datos.diferencias" in cuerpo_poblar,
              "la precarga muestra las diferencias que lee el Paso 1 -antes se leian "
              "y no aparecian en ninguna parte-")
    for campo in ("rObsPortada", "rLegalObservaciones", "rSuelosDescripcion",
                  "rConstrObservaciones", "rPlantObservaciones", "rDeslindeNorte"):
        comprobar('id="{}"'.format(campo) in html and '"{}"'.format(campo) in cuerpo_poblar,
                  "{} existe y se precarga".format(campo))
    comprobar("c.porcentaje" in guion and "col.porcentaje ? 100 : 1" in guion,
              "las columnas de porcentaje se muestran por 100 y se guardan entre 100 "
              "-'Valor residual %' rechazaba un 20-")
    comprobar("colPlantSin[4]" not in guion and 'c.clave === "estructura"' in guion,
              "las listas de PLANT se ubican por clave, no por posición")
    comprobar("salida.imagenes_por_pegar" in cuerpo_generar,
              "después de generar se listan las imágenes que hay que pegar en el Excel")

    # --- estadistica de testigos: grados de libertad, no tamano de muestra -
    # Bug real que encontro la propia verificacion de esta sesion: la tabla
    # de t de Student estaba indexada un lugar corrida (T95[n-1] leyendo una
    # tabla armada para 'n', no para 'grados de libertad'), y el intervalo de
    # confianza salia mucho mas ancho de lo que correspondia.
    stats = re.search(r"function pintarStatsTestigos\(([\s\S]*?)\n\}\n", guion)
    comprobar(stats is not None, "existe pintarStatsTestigos")
    adentro_stats = stats.group(1) if stats else ""
    comprobar("T95_POR_GRADOS_LIBERTAD[n - 1]" in adentro_stats,
              "la tabla de t de Student se indexa por GRADOS DE LIBERTAD "
              "(n-1), no por el tamaño de la muestra -el bug real que se "
              "encontro y corrigio")
    comprobar("Math.sqrt(varianza)" in adentro_stats and "n - 1" in adentro_stats,
              "la desviacion estandar es muestral (entre n-1), no poblacional")

    # --- Paso 1: lectura automatica del informe anterior (.xlsx) -----------
    comprobar('id="zonaRabobank"' in marcado, "existe la zona de arrastrar/soltar")
    comprobar('id="inputRabobank"' in marcado, "existe el input de archivo")

    comprobar(guion.count("function poblarFormularioRabobank(") == 1,
              "poblarFormularioRabobank esta declarada UNA sola vez")
    comprobar("return poblarFormularioRabobank;" in guion,
              "inicializarListasRabobank devuelve la funcion directo -sin "
              "destructuring ni asignacion a window, que ya causo un falso "
              "positivo del propio checker en esta sesion")
    comprobar(
        "const poblarFormularioRabobank = inicializarListasRabobank();" in guion,
        "se captura por closure con const, no una variable global")

    leer = re.search(
        r"async function leerArchivoRabobankAnterior\(archivo\) \{([\s\S]*?)\n  \}\n",
        guion)
    comprobar(leer is not None, "existe leerArchivoRabobankAnterior")
    cuerpo_leer = leer.group(1) if leer else ""
    comprobar('fetch("/retasaciones/leer-anterior"' in cuerpo_leer,
              "llama al endpoint de lectura -unico, sirve tanto .xlsx como .pdf")
    comprobar("poblarFormularioRabobank(salida.datos)" in cuerpo_leer,
              "precarga el Paso 2 con lo que devolvio el servidor")
    comprobar('avisar(salida.avisos.length ? "aviso" : "bien"' in cuerpo_leer,
              "avisa distinto si hubo observaciones o no -no un solo mensaje generico")
    comprobar("avisarDeFallo(" in cuerpo_leer,
              "un fallo de lectura se avisa, no se pierde en silencio")

    comprobar(
        re.search(r'\["dragenter", "dragover"\]\.forEach', guion) is not None,
        "arrastrar sobre la zona se cablea")
    comprobar(
        'zona.addEventListener("drop", (e) => {' in guion
        and "leerArchivoRabobankAnterior(archivo)" in guion,
        "soltar un archivo llama a la misma lectura que el selector manual")
    comprobar(
        'input.addEventListener("change", () => {' in guion
        and "leerArchivoRabobankAnterior(input.files[0])" in guion,
        "elegir el archivo por click cablea a la misma funcion")


def probar_formato_rut(html):
    """El RUT se autoformatea mientras se escribe y se valida con el
    digito verificador real (modulo 11), no solo pidiendo un guion.

    Incidente real (2026-09-23): un tasador tipeo el RUT sin puntos ni
    guion y /asistente/generar devolvio 500. El arreglo del servidor
    (services/rut.py) ya calcula el digito real; el usuario pidio que ademas
    la pantalla lo formatee sola mientras se escribe, para no depender de
    que el servidor avise despues de mandar el formulario."""
    print("\n[25] El RUT se autoformatea y se valida con el digito real")
    guion = guion_de(html)
    marcado = _solo_marcado(html)

    for nombre in ("digitoVerificadorRUT", "formatearRUT", "rutValido", "cablearFormatoRUT"):
        comprobar(guion.count("function {}(".format(nombre)) == 1,
                  "{} esta declarada UNA sola vez".format(nombre))

    comprobar(
        '["rut", "expRut", "rClienteRut", "rPropietarioRut", "rTasadorRut"]'
        ".forEach(cablearFormatoRUT)" in guion,
        "los cinco campos de RUT de la pantalla quedan cableados al arrancar")

    comprobar("input.rut-invalido" in guion or "input.rut-invalido" in html,
              "existe el estilo que marca un RUT con digito verificador incorrecto")

    # El calculo se corre de verdad (con Node, si esta disponible) y se
    # cruza contra el mismo calculo en Python -services/rut.py-: dos
    # implementaciones independientes tienen que dar el mismo digito.
    import shutil
    import subprocess
    import tempfile

    nodejs = shutil.which("node")
    if not nodejs:
        print("  (sin Node.js instalado: se omite el cruce JS vs Python)")
        return
    funciones = re.search(
        r"function digitoVerificadorRUT.*?\n}\n\nfunction formatearRUT.*?\n}\n"
        r"\nfunction rutValido.*?\n}\n", guion, re.S)
    comprobar(funciones is not None, "se pueden extraer las tres funciones")
    if funciones is None:
        return
    from services.rut import digito_verificador as digito_verificador_py

    cuerpos = ["76543210", "77533490", "1", "22222222", "76111222", "76123456"]
    script = funciones.group(0) + "\n" + "\n".join(
        'console.log("{}", digitoVerificadorRUT("{}"));'.format(c, c) for c in cuerpos)
    with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as f:
        f.write(script)
        ruta = f.name
    try:
        salida = subprocess.run([nodejs, ruta], capture_output=True, text=True,
                                timeout=15).stdout
    finally:
        os.unlink(ruta)
    for cuerpo in cuerpos:
        esperado = digito_verificador_py(cuerpo)
        comprobar("{} {}".format(cuerpo, esperado) in salida,
                  "digitoVerificadorRUT('{}') en JS da lo mismo que Python ({})"
                  .format(cuerpo, esperado), salida)


def probar_fotos_solo_con_titulo(html):
    """Las imagenes van solo con titulo, que el tasador puede cambiar.

    el agrónomo (2026-10-05): no hace falta una descripcion tecnica para ningun
    tipo de imagen, porque eso lo detecta mejor el agronomo; mejor solo el
    titulo, y que lo puedan cambiar. La ficha de cada foto tenia una caja de
    «Descripcion tecnica» que el modelo llenaba solo.
    """
    print("\n[27] Las imagenes, solo con titulo")
    guion = guion_de(html)
    m = re.search(r"function pintarFotos\(lista\) \{([\s\S]*?)\n\}\n", guion)
    comprobar(m is not None, "existe pintarFotos")
    ficha = m.group(1) if m else ""
    comprobar('data-campo="descripcion"' not in ficha and "Descripción técnica" not in ficha,
              "la ficha de la foto ya no tiene caja de descripcion tecnica")
    comprobar('data-campo="titulo"' in ficha and "<input" in ficha,
              "el titulo sigue siendo un campo que se puede editar")
    comprobar("Cámbielo como quiera" in ficha,
              "y dice que es lo que va bajo la imagen, y que se puede cambiar")
    comprobar("describir las imágenes" not in html and "Describiendo las imágenes" not in html,
              "ningun texto de la pagina promete describir las imagenes")


def probar_confirmar_antes_de_generar(html):
    """el usuario: "que te señale que cosas te faltan... y el usuario pueda
    escoger si generar el excel medio vacío en algunas partes o no
    hacerlo" -en vez de que "Generar el Excel" truene si falta algo, se
    pregunta antes, con la MISMA lista que ya pinta el panel "Qué falta en
    el informe" (no una segunda copia que se pueda desalinear)."""
    print("\n[26] Antes de generar se pregunta si falta algo, no se rompe")
    guion = guion_de(html)

    comprobar(guion.count("async function calcularFaltantesInforme(") == 1,
              "calcularFaltantesInforme esta separada de la pintura del panel "
              "-la reusa tambien la confirmacion")
    comprobar(guion.count("function confirmarGenerarConFaltantes(") == 1,
              "confirmarGenerarConFaltantes esta declarada UNA sola vez")

    generar = re.search(
        r'\$\("btnGenerar"\)\.addEventListener\("click", async \(\) => \{'
        r'([\s\S]*?)\n\}\);', guion)
    comprobar(generar is not None, "se encontro el manejador de btnGenerar")
    cuerpo_generar = generar.group(1) if generar else ""
    comprobar("calcularFaltantesInforme(borrador.expediente_id)" in cuerpo_generar,
              "btnGenerar consulta que falta antes de generar")
    comprobar(
        re.search(r"const seguir = await confirmarGenerarConFaltantes\(filas\);"
                  r"\s*\n\s*if \(!seguir\) return;", cuerpo_generar) is not None,
        "si el usuario cancela, no se manda la peticion de generar")
    # El chequeo va ANTES de deshabilitar el boton y mostrar el spinner: si
    # se cancela, el boton tiene que quedar tal cual, no en "Generando...".
    comprobar(
        cuerpo_generar.index("confirmarGenerarConFaltantes")
        < cuerpo_generar.index("boton.disabled = true"),
        "la confirmacion se pregunta ANTES de poner el boton en 'Generando...'")

    confirmar = re.search(
        r"function confirmarGenerarConFaltantes\(filas\) \{([\s\S]*?)\n\}\n",
        guion)
    comprobar(confirmar is not None, "existe confirmarGenerarConFaltantes")
    adentro = confirmar.group(1) if confirmar else ""
    comprobar(
        '"falta" || f[0] === "aviso"' in adentro or "'falta' || f[0] === 'aviso'" in adentro,
        "solo pregunta por lo que falta o tiene aviso -lo 'listo'/'opcional' "
        "no interrumpe")
    comprobar("if (!pendientes.length) return Promise.resolve(true)" in adentro,
              "si no falta nada, genera directo sin preguntar nada")
    comprobar('"Escape"' in adentro and "terminar(false)" in adentro,
              "Escape cancela, igual que los demas dialogos de la pantalla")


def main():
    print("=" * 70)
    print("PRUEBA DE LA PANTALLA")
    print("=" * 70)
    html = cargar()
    probar_controles_conectados(html)
    probar_filtros_de_lista(html)
    probar_botones_con_manejador(html)
    probar_vocabulario(html)
    probar_un_solo_manejador_por_input(html)
    probar_sesion_opcional(html)
    probar_que_ocultar_oculta(html)
    probar_que_todo_archivo_se_pueda_quitar(html)
    probar_sin_alert(html)
    probar_cadena_de_exportacion(html)
    probar_funciones_que_existen(html)
    probar_grafica_de_odepa(html)
    probar_documentos_legales_opcionales(html)
    probar_apartado_satelital(html)
    probar_calendario_propio(html)
    probar_casillas_y_desplegables(html)
    probar_partir_de_otra_curva(html)
    probar_pasos_de_creacion(html)
    probar_verificacion_completa(html)
    probar_buscador_de_especies(html)
    probar_consulta_de_anios(html)
    probar_panel_de_exclusiones(html)
    probar_vista_previa_en_pantalla(html)
    probar_actualizador_rabobank(html)
    probar_formato_rut(html)
    probar_confirmar_antes_de_generar(html)
    probar_fotos_solo_con_titulo(html)

    print("\n" + "=" * 70)
    if fallos:
        print("{}/{} comprobaciones FALLARON:".format(len(fallos), comprobaciones))
        for f in fallos:
            print("  - {}".format(f))
        return 1
    print("{}/{} comprobaciones OK".format(comprobaciones, comprobaciones))
    return 0


if __name__ == "__main__":
    sys.exit(main())
