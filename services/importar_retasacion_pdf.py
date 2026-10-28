"""Lee un PDF de retasacion Rabobank anterior (el informe ya entregado,
no la plantilla vigente) para precargar el Paso 2 del actualizador.

Distinto de importar_retasacion.py (que lee un .xlsx por coordenada de
celda, exacto): un PDF no tiene celdas. Se probo primero si el texto era
escaneado -no lo es, los dos PDF reales que uso el usuario como referencia
traen texto real, extraible- y despues si el orden de lectura alcanzaba
para reconstruir las tablas. **Se puede, anclando a la posicion (x,y) de
cada palabra en vez de al orden en que aparece en el flujo de texto**,
igual como si la pagina fuera, de nuevo, una hoja de calculo: los dos PDF
reales de el usuario salen de imprimir un informe muy parecido al Excel que ya
generamos, asi que la geometria de tabla se conserva -verificado
comparando dos clientes distintos (Fundo Ejemplo, un solo predio; Hijuela
Campos Nuevos, DOS predios): mismas etiquetas, columnas corridas apenas
unos puntos.

**100% determinista, sin IA** -misma decision que el lector de Excel: se
ancla al TEXTO de la etiqueta conocida (siempre en espanol, siempre la
misma, porque es la propia plantilla de Fitogreen), nunca a una coordenada
fija.

**Que se extrae en esta primera version, y por que NO todo:**

- cliente / predio (identificacion: nombre, rol, comuna, superficie,
  inscripcion CBR, deslindes) -> SI.
- tasador.nombre -> SI.
- diferencias.* (el parrafo de "Diferencias Tasacion Anterior" que ya trae
  cada seccion) -> SI, tal cual, para que el tasador lo edite en vez de
  redactarlo de cero.
- parametros.fecha_retasacion / uf_base / valor_usd -> **NO se precargan**.
  Son los datos de la tasacion NUEVA, no los de la vieja: traer el UF y la
  fecha de hace un año y dejarlos ahi silenciosamente seria el mismo error
  que CLAUDE.md §10.1 ya identifico como el mas caro del proyecto (varias
  UF distintas en el mismo libro). Se informan en los avisos, como
  referencia, nunca como campo editable ya lleno.
- **Fase 1: un solo predio**, igual que el resto del actualizador -si el
  PDF trae varios (como Hijuela Campos Nuevos, con "PREDIO 1" y "PREDIO
  2"), se toma el primero y se avisa que hay mas.
- Tablas de inventario: **plantaciones, suelos, derechos de agua,
  construcciones e instalaciones -> SI** (cada una con su propia geometria
  de columnas, medida y verificada contra los dos PDF reales). **Testigos
  de mercado (propiedades y derechos de agua) -> SI**, con un limite
  declarado: solo si la pagina trae la estructura de columnas limpia
  -verificado que Ejemplo no la trae para aguas (su referencia de aguas
  queda mezclada en la misma pagina que la de propiedades, sin una tabla
  propia), y esa pagina especifica devuelve vacio en vez de forzar un
  formato que no calza.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple, Union

import fitz

_MARCA_SECCION = {
    "suelos": ("SUELOS",),
    "aguas": ("Derechos", "de", "Aguas"),
    "construcciones": ("5.1-", "Construcciones"),
    "instalaciones": ("5.2.-", "Instalaciones"),
    "plantaciones": ("6.-", "Plantaciones"),
}

_ETIQUETA_DIFERENCIAS = "Diferencias Tasación Anterior"

#: palabras que no cuentan como "valor" -son sobras de la etiqueta (el
#: ":" que separa "EJECUTIVO :" de su valor, por ejemplo).
_PUNTUACION_SUELTA = {":", "-", ".", ""}


def _palabras(pagina) -> List[Tuple[float, float, float, float, str]]:
    return [(p[0], p[1], p[2], p[3], p[4]) for p in pagina.get_text("words")]


def _buscar(pagina, texto: str) -> List["fitz.Rect"]:
    """pagina.search_for, pero ordenado de arriba hacia abajo. **No dar por
    sentado el orden de search_for**: se probó en carne propia -devolvía
    la 'Fs' de la tabla de Aguas antes que la de Inmuebles, aunque esta
    última está más arriba en la página- y arruinaba todo lo que dependía
    de 'la primera coincidencia'."""
    return sorted(pagina.search_for(texto), key=lambda r: (r.y0, r.x0))


def _limpiar_valor(texto: str) -> Optional[str]:
    piezas = [t for t in texto.split() if t not in _PUNTUACION_SUELTA]
    limpio = " ".join(piezas).strip()
    return limpio or None


def _valor_a_la_derecha(pagina, etiqueta: str, indice: int = 0,
                         tolerancia_y: float = 6.0) -> Optional[str]:
    """Busca 'etiqueta' (texto exacto) y devuelve las palabras a su derecha
    en la misma fila -la etiqueta ancla, el valor esta en la columna de al
    lado. Tolerancia chica a propósito: las filas de la portada estan a
    unos 12-18pt una de otra, una tolerancia mayor mezcla la fila vecina."""
    rects = _buscar(pagina, etiqueta)
    if indice >= len(rects):
        return None
    rect = rects[indice]
    y_centro = (rect.y0 + rect.y1) / 2
    candidatas = [p for p in _palabras(pagina)
                  if p[0] > rect.x1 - 2 and abs((p[1] + p[3]) / 2 - y_centro) <= tolerancia_y]
    candidatas.sort(key=lambda p: p[0])
    return _limpiar_valor(" ".join(p[4] for p in candidatas))


def _primer_valor_por_columna(pagina, limites_x: List[float], columnas: List[str],
                               y0: float, y1: float, tolerancia_y: float = 3.0
                               ) -> Dict[str, str]:
    """Recorre la ventana [y0, y1] en orden de lectura y se queda, por cada
    columna, con la PRIMERA linea de texto que cae en su rango de X -asi
    da lo mismo si una columna vecina (ej. un nombre largo) ocupa varias
    lineas: la columna corta (Fs, Rol, Superficie) igual aparece en su
    propia linea, en algun punto de la ventana."""
    palabras = sorted((p for p in _palabras(pagina) if y0 < p[1] < y1),
                      key=lambda p: (p[1], p[0]))
    primera_y: Dict[str, float] = {}
    valores: Dict[str, List[str]] = {}
    for x0, y0p, _x1, _y1, texto in palabras:
        idx = 0
        while idx < len(limites_x) and x0 >= limites_x[idx]:
            idx += 1
        col = columnas[idx]
        if col not in primera_y:
            primera_y[col] = y0p
            valores[col] = [texto]
        elif abs(y0p - primera_y[col]) <= tolerancia_y:
            valores[col].append(texto)
    return {k: v for k, v in ((c, " ".join(t)) for c, t in valores.items()) if v}


def _bloque_de_texto(pagina, etiqueta_inicio: str, y_max: Optional[float] = None
                      ) -> Optional[str]:
    """Todo el texto de la pagina que queda DEBAJO de 'etiqueta_inicio',
    hasta 'y_max' (o el final de la pagina). Sirve para los parrafos libres
    (Diferencias Tasacion Anterior)."""
    rects = _buscar(pagina, etiqueta_inicio)
    if not rects:
        return None
    y0 = rects[0].y1
    limite = y_max if y_max is not None else pagina.rect.y1
    palabras = [p for p in _palabras(pagina) if y0 < p[1] < limite]
    if not palabras:
        return None
    filas: Dict[float, List[Tuple[float, str]]] = {}
    for p in palabras:
        banda = round(p[1] / 4) * 4
        filas.setdefault(banda, []).append((p[0], p[4]))
    lineas = [" ".join(t for _, t in sorted(filas[banda])) for banda in sorted(filas)]
    return " ".join(lineas).strip() or None


def _pagina_con(doc, tokens: Tuple[str, ...], y_max: float = 100.0):
    """La primera pagina cuyo titulo (arriba de y_max) contiene TODOS los
    tokens dados -asi no se confunde, por ejemplo, la pagina '3.- SUELOS'
    con la fila 'SUELOS' de la tabla de Resumen de Valores, que esta mas
    abajo en su pagina."""
    for pagina in doc:
        arriba = " ".join(p[4] for p in _palabras(pagina) if p[1] < y_max)
        if all(t in arriba for t in tokens):
            return pagina
    return None


def _numero(texto: Optional[str]) -> Optional[float]:
    if not texto:
        return None
    limpio = texto.replace("$", "").replace(",", "").strip()
    try:
        return float(limpio)
    except ValueError:
        return None


def _nombre_y_rol_predio(pagina_resumen) -> Dict[str, Any]:
    """'1) PREDIO DENOMINADO: <lo que sea>'. En un predio unico es su
    nombre ('El Ejemplo'); en un caso de varios predios sin nombre comun
    la propia plantilla pone 'Rol 139-853' -se toma tal cual, es lo que EL
    INFORME usa para identificarlo, mejor eso que dejarlo vacio."""
    valor = _valor_a_la_derecha(pagina_resumen, "DENOMINADO:", tolerancia_y=6)
    if not valor:
        return {}
    salida: Dict[str, Any] = {"nombre": valor}
    partes = valor.split()
    if len(partes) == 2 and partes[0].lower() == "rol":
        salida["rol_avaluo"] = partes[1]
    return salida


def _predio_comuna_superficie(pagina1) -> Dict[str, Any]:
    """Tabla 'Predios Tasados' de la portada. Solo se usa para comuna y
    superficie -nombre y rol ya salen, mas confiables, de la pagina de
    Resumen (una sola linea, sin ambigüedad de predios multiples)."""
    rects_comuna = _buscar(pagina1, "Comuna")
    if not rects_comuna:
        return {}
    ancla = rects_comuna[0]
    rects_rol = [r for r in _buscar(pagina1, "Rol") if abs(r.y0 - ancla.y0) <= 10]
    rects_sup = [r for r in _buscar(pagina1, "Superficie")
                if abs(r.y0 - ancla.y0) <= 10] or _buscar(pagina1, "Superficies")
    if not rects_rol or not rects_sup:
        return {}
    x_rol = min(rects_rol, key=lambda r: abs(r.y0 - ancla.y0)).x0
    x_sup = min((r for r in rects_sup if abs(r.y0 - ancla.y0) <= 10),
               key=lambda r: r.x0, default=None)
    if x_sup is None:
        return {}
    limites = [(x_rol + ancla.x0) / 2, (ancla.x0 + x_sup.x0) / 2]

    fin = _buscar(pagina1, "Superficie Total")
    y1 = min((r.y0 for r in fin if r.y0 > ancla.y1), default=ancla.y1 + 60)

    fila = _primer_valor_por_columna(
        pagina1, limites, ["_antes", "comuna", "_superficie"], ancla.y1, y1)
    salida: Dict[str, Any] = {}
    if fila.get("comuna"):
        salida["comuna"] = fila["comuna"]
    superficie = _numero(fila.get("_superficie"))
    if superficie is not None:
        salida["superficie_legal_ha"] = superficie
    return salida


def _bloque_predio_uno(pagina_inmuebles):
    """Ventana [y0, y1] del PRIMER predio en la página de Inmuebles:
    desde su encabezado hasta la palabra 'Inmueble' que siempre le sigue
    (empieza ahí la descripción legal en prosa)."""
    encabezado = _buscar(pagina_inmuebles, "Fs")
    if not encabezado:
        return None
    y0 = encabezado[0].y1
    fin = _buscar(pagina_inmuebles, "Inmueble")
    y1 = min((r.y0 for r in fin if r.y0 > y0), default=y0 + 80)
    return y0, y1


_NUMERO_RE = re.compile(r"^\d+$")


def _inscripcion_y_rol_predio(pagina_inmuebles) -> Dict[str, Any]:
    """Fs/Nº/Año/C.B.R. (izquierda de 'Rol SII') y Rol/Superficie (a la
    derecha) del PRIMER predio.

    No reparte por columna de X fija: el nombre del predio puede ocupar
    varias líneas y desplazar en cuál línea cae cada dato (verificado con
    los dos PDF reales -Fundo Ejemplo, una línea; Hijuela Campos Nuevos,
    el nombre solo ya ocupa tres-). En vez de eso, del lado izquierdo se
    separa por si la palabra es un NÚMERO (Fs/Nº/Año) o no (el nombre del
    predio, o el C.B.R.) -eso sí es estable venga en la línea que venga-,
    y el C.B.R. se toma como la palabra no-numérica más cercana a su
    propio encabezado, no la primera que aparezca."""
    rects_fs = _buscar(pagina_inmuebles, "Fs")
    rects_rol = _buscar(pagina_inmuebles, "Rol")
    rects_cbr = _buscar(pagina_inmuebles, "C.B.R.")
    rects_sup = _buscar(pagina_inmuebles, "Superficie")
    if not (rects_fs and rects_rol and rects_cbr and rects_sup):
        return {}
    x_cbr = rects_cbr[0].x0
    x_rol = rects_rol[0].x0
    x_sup = rects_sup[0].x0

    ventana = _bloque_predio_uno(pagina_inmuebles)
    if ventana is None:
        return {}
    y0, y1 = ventana

    salida: Dict[str, Any] = {}

    # -2pt de margen: un valor de columna a veces arranca un pelo antes que
    # su propio encabezado (medido: '139-95' arranca en x=304.75 con 'Rol'
    # en x=305.59) -sin el margen, ese valor caia en la columna de al lado
    # y el rol de avaluo se perdia en silencio.
    corte = x_rol - 2

    izquierda = sorted(
        (p for p in _palabras(pagina_inmuebles) if y0 < p[1] < y1 and p[0] < corte),
        key=lambda p: (p[1], p[0]))
    numeros = [p for p in izquierda if _NUMERO_RE.match(p[4])]
    alfa = [p for p in izquierda if not _NUMERO_RE.match(p[4])]
    if len(numeros) >= 3:
        salida["fojas"] = numeros[0][4]
        salida["numero"] = numeros[1][4]
        anio = _numero(numeros[2][4])
        if anio is not None:
            salida["anio"] = int(anio)
    if alfa:
        conservador = min(alfa, key=lambda p: abs(p[0] - x_cbr))
        salida["conservador"] = conservador[4]

    derecha = sorted(
        (p for p in _palabras(pagina_inmuebles) if y0 < p[1] < y1 and p[0] >= corte),
        key=lambda p: (p[1], p[0]))
    if derecha:
        y_rol = derecha[0][1]
        rol_tokens = [p[4] for p in derecha if p[0] < x_sup and abs(p[1] - y_rol) <= 3]
        if rol_tokens:
            salida["rol_avaluo"] = " ".join(rol_tokens)
        sup_tokens = [p for p in derecha if p[0] >= x_sup]
        if sup_tokens:
            y_sup = sup_tokens[0][1]
            valores_sup = [p[4] for p in sup_tokens if abs(p[1] - y_sup) <= 3]
            superficie = _numero(valores_sup[0]) if valores_sup else None
            if superficie is not None:
                salida["superficie_legal_ha"] = superficie
    return salida


def _deslindes(pagina_inmuebles) -> Dict[str, str]:
    """Deslindes Norte/Sur/Oriente/Poniente del PRIMER predio: cuatro
    columnas de texto que se reparten en varias lineas. Se ubica por rango
    de X (punto medio entre encabezados) y se junta todo el texto de cada
    columna, en orden de lectura."""
    rects = _buscar(pagina_inmuebles, "Norte")
    if not rects:
        return {}
    y0 = rects[0].y1
    encabezados = {}
    for nombre in ("Norte", "Sur", "Oriente", "Poniente"):
        r = _buscar(pagina_inmuebles, nombre)
        cercanas = [x for x in r if abs(x.y0 - rects[0].y0) <= 5]
        if cercanas:
            encabezados[nombre] = cercanas[0].x0
    if len(encabezados) < 4:
        return {}
    orden = ["Norte", "Sur", "Oriente", "Poniente"]
    xs = [encabezados[o] for o in orden]
    limites = [(xs[i] + xs[i + 1]) / 2 for i in range(3)]

    siguiente = _buscar(pagina_inmuebles, "Inscripción dominio")
    y_fin = min((r.y0 for r in siguiente if r.y0 > y0), default=pagina_inmuebles.rect.y1)

    columnas: Dict[str, List[Tuple[float, float, str]]] = {o: [] for o in orden}
    for x0, y0p, _x1, _y1, texto in _palabras(pagina_inmuebles):
        if not (y0 < y0p < y_fin):
            continue
        idx = 0
        while idx < len(limites) and x0 >= limites[idx]:
            idx += 1
        columnas[orden[idx]].append((y0p, x0, texto))

    salida = {}
    for o in orden:
        piezas = sorted(columnas[o])
        if piezas:
            salida[o] = " ".join(t for _, _, t in piezas)
    return salida


_NUM_SIMPLE_RE = re.compile(r"^\d+[.,]\d+$|^\d+$")
_MARCO_RE = re.compile(r"(\d+[.,]?\d*)\s*[xX]\s*(\d+[.,]?\d*)")
_ANIO_RE = re.compile(r"(19|20)\d{2}")
_ESTADOS_VALIDOS = {"bueno", "regular", "malo"}


def _pos_header(pagina, y_aprox: float, *alternativas: str, x_min: float = 0,
                tolerancia: float = 25.0) -> Optional[float]:
    """La palabra de 'alternativas' -se prueban en orden- mas cercana a
    'y_aprox', con x0 >= x_min. Hace falta probar varias porque el mismo
    encabezado cambia de texto entre PDF de años distintos ('Tipo de
    Cultivo' en uno, 'Especie' en otro; 'Distancia Plantación' vs
    'Marco') aunque la columna sea la misma."""
    for alt in alternativas:
        candidatas = [r for r in _buscar(pagina, alt)
                     if r.x0 >= x_min and abs(r.y0 - y_aprox) <= tolerancia]
        if candidatas:
            return min(candidatas, key=lambda r: abs(r.y0 - y_aprox)).x0
    return None


def _filas_por_ancla(pagina, y0: float, y1: float, x_ancla0: float, x_ancla1: float,
                     limites_x: List[float], columnas: List[str],
                     patron_ancla=_NUM_SIMPLE_RE) -> List[Dict[str, str]]:
    """Lee TODAS las filas de una tabla cuyo numero de renglones fisicos
    por fila logica VARIA -una variedad larga empuja columnas a la linea
    de abajo, medido en los dos PDF reales (plantaciones, ver
    _plantaciones())-. Cada fila nueva se ancla a una palabra que cae en
    [x_ancla0, x_ancla1) y matchea 'patron_ancla' -la superficie en
    hectareas: un numero que nunca se corta ni se mueve de linea en
    ninguno de los dos PDF medidos-. Entre un ancla y la siguiente se lee
    la ventana completa con _primer_valor_por_columna, que ya resuelve
    sola una columna que aparece en la linea de abajo: toma la PRIMERA
    ocurrencia de cada columna en toda la ventana, este arriba o abajo."""
    anclas = sorted(
        p[1] for p in _palabras(pagina)
        if y0 < p[1] < y1 and x_ancla0 <= p[0] < x_ancla1 and patron_ancla.match(p[4]))
    filas = []
    for i, y_ancla in enumerate(anclas):
        y_ini = y_ancla - 3
        y_fin = anclas[i + 1] - 3 if i + 1 < len(anclas) else y1
        fila = _primer_valor_por_columna(pagina, limites_x, columnas, y_ini, y_fin)
        if fila:
            filas.append(fila)
    return filas


def _marco_a_densidad(texto: Optional[str]) -> Optional[int]:
    """'3,5X2,0' o '4,5 x 1,5' (metros entre hileras x metros entre
    plantas) -> plantas por hectarea. El formato del token cambia entre
    PDF (todo junto sin espacios vs con espacios alrededor de la x), asi
    que se busca el patron en vez de partir por un separador fijo."""
    if not texto:
        return None
    m = _MARCO_RE.search(texto)
    if not m:
        return None
    a = float(m.group(1).replace(",", "."))
    b = float(m.group(2).replace(",", "."))
    if a <= 0 or b <= 0:
        return None
    return round(10000 / (a * b))


def _anio_de(texto: Optional[str]) -> Optional[int]:
    """El primer año de 4 digitos dentro del texto -la columna a veces
    arrastra un pedazo de la variedad de al lado ('Cabernet S.' empuja
    una 'S.' suelta a la columna de año), asi que no se puede tomar el
    texto entero como numero."""
    if not texto:
        return None
    m = _ANIO_RE.search(texto)
    return int(m.group(0)) if m else None


def _plantaciones(pagina, marca_royalty: str, encabezado_alt: Tuple[str, ...],
                  fin_alt: Tuple[str, ...]) -> List[Dict[str, Any]]:
    """Una sub-tabla de Plantaciones (6.1 sin royalty o 6.2 con royalty).

    'Variedad' tambien aparece SUELTO en el parrafo narrativo de
    antecedentes generales ("...uva de mesa variedad Crimson..."), antes
    de la tabla de verdad -asi que anclar solo al titulo de seccion no
    alcanza para encontrar el encabezado. Se ancla primero a
    'marca_royalty' (el texto 'Royalties' o 'con royalt...' segun la
    sub-tabla), y desde ahi recien se busca 'Variedad'."""
    marcas = _buscar(pagina, marca_royalty)
    if not marcas:
        return []
    y_marca = marcas[0].y0
    variedad_rects = [r for r in _buscar(pagina, "Variedad") if r.y0 > y_marca]
    if not variedad_rects:
        return []
    y_header = variedad_rects[0].y0
    x_variedad = variedad_rects[0].x0
    # Ancla real, no un margen izquierdo adivinado: un margen fijo probo
    # estar demasiado cerca de un dato real ('MESA' de 'VID DE MESA' cae
    # justo al lado equivocado del limite si el margen no viene del propio
    # encabezado) -"Tipo de Cultivo" en un PDF, "Especie" en el otro.
    x_especie = _pos_header(pagina, y_header, "Tipo", "Especie", x_min=0) or 45.0

    x_anio = _pos_header(pagina, y_header, "Año", x_min=x_variedad)
    x_marco = _pos_header(pagina, y_header, "Marco", "Distancia", x_min=x_anio or x_variedad)
    x_rol = _pos_header(pagina, y_header, "Rol", x_min=x_marco or x_anio or x_variedad)
    x_ha = _pos_header(pagina, y_header, "ha", "Hás", x_min=(x_rol or 0) + 5)
    if not (x_anio and x_marco and x_rol and x_ha):
        return []
    dolar_rects = [r for r in _buscar(pagina, "$")
                  if abs(r.y0 - y_header) <= 15 and r.x0 > x_ha]
    x_precio = min((r.x0 for r in dolar_rects), default=None)
    x_total = _pos_header(pagina, y_header, "Total", x_min=(x_precio or x_ha) + 5)
    x_uf = _pos_header(pagina, y_header, "UF", x_min=(x_total or x_ha) + 5)
    x_estado = _pos_header(pagina, y_header, "Estado", x_min=(x_uf or x_ha) + 5)

    xs = [x_especie, x_variedad, x_anio, x_marco, x_rol, x_ha]
    columnas = ["especie", "variedad", "anio", "marco", "rol", "ha"]
    if x_precio:
        xs.append(x_precio)
        columnas.append("precio")
    if x_total:
        xs.append(x_total)
        columnas.append("total")
    if x_uf:
        xs.append(x_uf)
        columnas.append("uf")
    if x_estado:
        xs.append(x_estado)
        columnas.append("estado")
    limites = [(xs[i] + xs[i + 1]) / 2 for i in range(len(xs) - 1)]

    fin_candidatos = [r for alt in fin_alt for r in _buscar(pagina, alt)
                      if r.y0 > y_header + 10]
    y_fin = min((r.y0 for r in fin_candidatos), default=pagina.rect.y1)

    # El margen (-6) es el mismo que hizo falta en _inscripcion_y_rol_predio:
    # el valor de superficie arranca un pelo antes que su propio encabezado.
    filas = _filas_por_ancla(pagina, y_header, y_fin, x_ha - 6,
                             (x_precio or x_ha + 60), limites, columnas)
    return [f for f in filas if f.get("especie") and f.get("variedad")]


def _plantaciones_sin_royalty(pagina) -> List[Dict[str, Any]]:
    salida = []
    for fila in _plantaciones(pagina, "Royalties", ("Año",), ("Total", "frutales")):
        densidad = _marco_a_densidad(fila.get("marco"))
        anio = _anio_de(fila.get("anio"))
        if densidad is None or anio is None:
            continue
        dato: Dict[str, Any] = {
            "especie": fila["especie"],
            "variedad": fila["variedad"],
            "anio_plantacion": anio,
            "densidad_plantas_ha": densidad,
            "superficie_ha": _numero(fila.get("ha")),
        }
        if fila.get("rol"):
            dato["rol_avaluo"] = fila["rol"].split()[0]
        estado = (fila.get("estado") or "").strip().lower()
        if estado in _ESTADOS_VALIDOS:
            dato["estado"] = estado.capitalize()
        salida.append(dato)
    return salida


def _plantaciones_con_royalty(pagina) -> List[Dict[str, Any]]:
    salida = []
    for fila in _plantaciones(pagina, "con royalt", ("Año",), ("Diferencias", "Observaciones")):
        anio = _anio_de(fila.get("anio"))
        precio = _numero((fila.get("precio") or "").replace("$", "").strip())
        if anio is None or not precio:
            continue
        dato: Dict[str, Any] = {
            "especie": fila["especie"],
            "variedad": fila["variedad"],
            "anio_plantacion": anio,
            "marco_plantacion": (fila.get("marco") or "").strip(),
            "valor_unitario_clp_ha": precio,
            "superficie_ha": _numero(fila.get("ha")),
        }
        if fila.get("rol"):
            dato["rol_avaluo"] = fila["rol"].split()[0]
        estado = (fila.get("estado") or "").strip().lower()
        if estado in _ESTADOS_VALIDOS:
            dato["estado"] = estado.capitalize()
        salida.append(dato)
    return salida


_ROL_RE = re.compile(r"^\d[\d.]*-\d+$")


def _suelos(pagina) -> List[Dict[str, Any]]:
    """La tabla '3.2 Valores' de la pagina SUELOS -clase de uso, rol,
    superficie, valor $/ha, total $, total UF-.

    Solo cubre el caso de un predio (el mismo alcance que Fase 1 del
    actualizador): en el PDF de Hijuela Campos Nuevos -DOS predios- esta
    tabla trae una columna extra ('Clase de uso de suelo' aparte del
    nombre del predio) que no calza con este mapeo de columnas, y en vez
    de forzarla se descarta cada fila cuyo 'rol' no tiene forma de rol
    (verificado: asi se queda sin datos en vez de inventar algo)."""
    y_valores = [r.y0 for r in _buscar(pagina, "Valores") if r.x0 < 150]
    if not y_valores:
        return []
    y_marca = y_valores[0]
    rol_rects = [r for r in _buscar(pagina, "Rol") if r.y0 > y_marca]
    if not rol_rects:
        return []
    y_header = rol_rects[0].y0
    x_rol = rol_rects[0].x0

    x_sup = _pos_header(pagina, y_header, "Superficie", x_min=x_rol)
    dolar_rects = [r for r in _buscar(pagina, "$") if abs(r.y0 - y_header) <= 15]
    dolares = sorted(r.x0 for r in dolar_rects if r.x0 > (x_sup or x_rol))
    x_valorha = dolares[0] if len(dolares) >= 1 else None
    x_totalpesos = dolares[1] if len(dolares) >= 2 else None
    x_uf = _pos_header(pagina, y_header, "UF", x_min=(x_totalpesos or x_rol) + 5)
    if not (x_sup and x_valorha and x_totalpesos and x_uf):
        return []

    # x_rol - 70: 'R Plantado' (parte del texto de clase_uso, sin
    # encabezado propio) arranca lejos de x_rol; un margen fijo de 70pt
    # midio bien el hueco real en el PDF de referencia.
    xs = [x_rol - 70, x_rol, x_sup, x_valorha, x_totalpesos, x_uf]
    columnas = ["clase_uso", "rol", "superficie", "valor_ha", "total_pesos", "valor_uf"]
    limites = [(xs[i] + xs[i + 1]) / 2 for i in range(len(xs) - 1)]

    fin_candidatos = [r.y0 for r in _buscar(pagina, "Total") if r.y0 > y_header + 5]
    y_fin = (min(fin_candidatos) + 5) if fin_candidatos else pagina.rect.y1

    filas = _filas_por_ancla(pagina, y_header, y_fin, x_sup - 6, x_valorha,
                             limites, columnas)
    salida = []
    for fila in filas:
        clase = (fila.get("clase_uso") or "").strip()
        rol = (fila.get("rol") or "").strip().split()[-1] if fila.get("rol") else ""
        superficie_ha = _numero(fila.get("superficie"))
        valor_uf = _numero((fila.get("valor_uf") or "").replace("UF", "").strip())
        if not clase or clase.lower() == "total" or not _ROL_RE.match(rol):
            continue
        if not superficie_ha or not valor_uf:
            continue
        salida.append({
            "clase_uso": clase,
            "superficie_m2": round(superficie_ha * 10000, 2),
            # UF total de la clase / superficie -> UF/ha -> UF/m2. Se
            # deriva de dos columnas del PDF (Valor UF y Superficie), no
            # del tipo de cambio: no hace falta saber el TC de la
            # tasacion vieja para este calculo.
            "valor_unitario_uf_m2": round((valor_uf / superficie_ha) / 10000, 6),
        })
    return salida


def _numero_coma_decimal(texto: Optional[str]) -> Optional[float]:
    """'4,175' -> 4.175, NO 4175 -la coma es el separador DECIMAL en esta
    columna, al reves que en las columnas de dinero (`_numero`, coma de
    miles). No se adivino: se verifico cruzando contra el propio texto del
    PDF -"cada accion equivale a 15 lt/seg" mas la fila de Lts/seg: 4,175 x
    15 = 62,625, contra 62.6 en la tabla (Ejemplo); 10,33 x 1,5 = 15,495,
    contra los "15,495 l/s" que el parrafo de antecedentes declara tal cual
    y los 15.5 redondeados de la tabla (Hijuela). Las dos maneras de
    calcularlo coinciden, que es la prueba que este proyecto usa para todo
    lo demas."""
    if not texto:
        return None
    try:
        return float(texto.replace(".", "").replace(",", "."))
    except ValueError:
        return None


_NUMLIDER_RE = re.compile(r"^[\d.,]+$")


_ROL_INICIO_RE = re.compile(r"^\d[\d.]*-$")
_ROL_FIN_RE = re.compile(r"^\d+$")


def _numero_despues(palabras_fila, palabra_ancla, tolerancia_y: float = 3.0
                    ) -> Optional[str]:
    """La primera palabra a la derecha de 'palabra_ancla', en su misma
    linea -para leer el numero que sigue a un '$' o a un 'UF' suelto."""
    candidatas = sorted(
        (p for p in palabras_fila
         if abs(p[1] - palabra_ancla[1]) <= tolerancia_y and p[0] > palabra_ancla[0]),
        key=lambda p: p[0])
    return candidatas[0][4] if candidatas else None


def _rol_en(palabras: List[Tuple[float, float, float, float, str]]
           ) -> Optional[str]:
    """Un rol de avaluo entre las palabras dadas, sin importar la columna
    -mismo criterio que el resto de _aguas/_construcciones/_instalaciones:
    una forma reconocible (`numero-numero`) importa mas que su posicion.

    Tambien reconstruye un rol partido en dos LINEAS distintas ('139-' /
    '95', verificado en los PDF reales cuando el nombre del predio u otra
    columna empuja la fila a envolver): busca la mitad que sigue por su
    COLUMNA de X, no por ser la vecina siguiente en orden de lectura -otra
    columna de la misma fila puede caer, por Y, justo entre las dos mitades
    (verificado: un '45' suelto de 'Fuerza electrica 45 KVA' quedaba en
    medio de '139-' y '95')."""
    for p in palabras:
        if _ROL_RE.match(p[4]):
            return p[4]
    inicios = [p for p in palabras if _ROL_INICIO_RE.match(p[4])]
    fines = [p for p in palabras if _ROL_FIN_RE.match(p[4])]
    for a in inicios:
        b = min((f for f in fines if abs(f[0] - a[0]) < 5),
               key=lambda f: abs(f[1] - a[1]), default=None)
        if b is not None:
            return a[4] + b[4]
    return None


def _aguas(pagina) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    """La tabla '4.- Derechos de Aguas': Derecho / Origen / Rol / Lts-seg /
    Valor $ lt-s / Total $ / Valor UF, mas las Observaciones debajo.

    Mismo numero de filas variable que plantaciones y suelos, pero acá NO
    se reparte por columnas de X: se probo primero (como las otras dos
    tablas) y la columna 'Derecho' -una descripcion libre que puede envolver
    en 2 o 3 lineas- se desborda hacia la derecha lo bastante como para
    invadir Origen y a veces hasta Rol (verificado en Hijuela: la palabra
    'La' de su propio envoltorio cae mas a la derecha que el encabezado
    'Rol'). En cambio cada dato que de verdad importa tiene una FORMA
    reconocible sin importar en que columna haya caido el texto de al lado:

    - el rol: un patron `numero-numero` (`_ROL_RE`, ya usado en suelos).
    - la cantidad y el tipo de derecho ('4,175 Regadores', '10,33 acciones'):
      la primera palabra-numero de la zona izquierda (antes de Origen) y la
      palabra que la sigue.
    - el caudal: es el propio ancla de la fila (Lts/seg, nunca se corta).
    - el valor por L/s y el total: el primer y el segundo '$' de la fila, EN
      ORDEN DE X -no de Y: los dos "$" de una misma fila pueden tener un
      jitter vertical menor a 1pt que invierte el orden si se ordenara por Y
      primero, y eso alcanza para cambiar cual es cual (verificado: Hijuela
      los traia literalmente al reves ordenando por Y).
    - el UF: la palabra 'UF' y el numero que la sigue.

    Solo 'origen' (la fuente: rio, canal, pozo) sigue necesitando una franja
    de X -no tiene una forma reconocible propia-, con el mismo riesgo de
    desborde que las otras dos columnas de texto libre ya tuvieron: puede
    salir incompleto en un caso limite (verificado en Hijuela, donde el
    nombre de la asociacion de canales se repite parcialmente en Derecho Y
    en Origen) en vez de con texto de otra columna mezclado -se prefiere
    corto y limpio a largo y contaminado, porque el tasador revisa este
    campo igual que el resto."""
    rects_origen = _buscar(pagina, "Origen")
    if not rects_origen:
        return [], None
    y_header = rects_origen[0].y0
    x_origen = rects_origen[0].x0
    x_rol = _pos_header(pagina, y_header, "Rol", x_min=x_origen + 5)
    x_ltsseg = _pos_header(pagina, y_header, "Lts/seg", "L/s", x_min=(x_rol or 0) + 5)
    if not (x_rol and x_ltsseg):
        return [], None

    y_fin_candidatos = [r.y0 for r in _buscar(pagina, "Total") if r.y0 > y_header + 5]
    y_fin_tabla = (min(y_fin_candidatos) - 2) if y_fin_candidatos else pagina.rect.y1

    anclas = sorted(
        (p[1], p[4]) for p in _palabras(pagina)
        if y_header < p[1] < y_fin_tabla and x_ltsseg - 8 <= p[0] < x_ltsseg + 40
        and _NUM_SIMPLE_RE.match(p[4]))

    filas = []
    for i, (y_ancla, texto_ancla) in enumerate(anclas):
        y_ini = (anclas[i - 1][0] + y_ancla) / 2 if i > 0 else y_header
        y_fin = (anclas[i + 1][0] + y_ancla) / 2 if i + 1 < len(anclas) else y_fin_tabla
        palabras_fila = sorted(
            (p for p in _palabras(pagina) if y_ini < p[1] < y_fin),
            key=lambda p: (p[1], p[0]))

        zona_izquierda = [p for p in palabras_fila if p[0] < x_origen - 10]
        idx_cantidad = next((i for i, p in enumerate(zona_izquierda)
                            if _NUMLIDER_RE.match(p[4])), None)
        cantidad = (zona_izquierda[idx_cantidad][4]
                   if idx_cantidad is not None else None)
        tipo = (zona_izquierda[idx_cantidad + 1][4]
               if idx_cantidad is not None and idx_cantidad + 1 < len(zona_izquierda)
               else None)

        rol = _rol_en(palabras_fila)

        dolares = sorted((p for p in palabras_fila if p[4] == "$"), key=lambda p: p[0])
        valor_ls = _numero_despues(palabras_fila, dolares[0]) if dolares else None

        uf_rect = next((p for p in palabras_fila if p[4] == "UF"), None)

        origen = " ".join(p[4] for p in palabras_fila
                          if (x_origen - 10) <= p[0] < x_rol - 15) or None

        if not (cantidad and rol and origen):
            continue
        dato: Dict[str, Any] = {
            "fuente": origen,
            "rol_avaluo": rol,
            "derechos_cantidad": _numero_coma_decimal(cantidad),
            "caudal_nominal_ls": _numero(texto_ancla),
        }
        if tipo:
            dato["tipo"] = tipo
        if valor_ls:
            valor = _numero(valor_ls)
            if valor:
                dato["valor_unitario_clp"] = valor
        # El total $ y el UF de cada derecho no tienen celda propia en
        # RetasacionRabobank.aguas -los vuelve a calcular el Excel a partir
        # del caudal y el valor unitario-, asi que no hace falta guardarlos.
        if dato["derechos_cantidad"] and dato["derechos_cantidad"] > 0:
            filas.append(dato)

    # El bloque libre de "Observaciones:" queda DESPUES de la tabla y ANTES
    # de "Diferencias Tasación Anterior" -sin ese limite, _bloque_de_texto
    # se traga tambien esa seccion siguiente, mismo tipo de trampa que ya
    # costo con una tabla antes de su etiqueta (§16.a.centies.sexies)-.
    fin_diferencias = _buscar(pagina, _ETIQUETA_DIFERENCIAS)
    y_max_obs = fin_diferencias[0].y0 - 5 if fin_diferencias else None
    observaciones = _bloque_de_texto(pagina, "Observaciones:", y_max=y_max_obs)
    return filas, observaciones


def _bloque_de_encabezado(pagina, y_referencia: float, *etiquetas: str,
                          margen_arriba: float = 10.0, margen_abajo: float = 40.0
                          ) -> float:
    """La Y donde termina un encabezado que envuelve en varias lineas,
    buscando la mas baja entre un grupo de ETIQUETAS QUE SON DEL
    ENCABEZADO -nunca de una fila de datos-. Cuanto envuelve el encabezado
    cambia de un PDF a otro (verificado: en Ejemplo hay 29pt entre
    'Superficie' y el primer dato; en Hijuela, solo 19pt), asi que un
    margen fijo sirve para uno y se come la primera fila de datos del
    otro."""
    y_fin = y_referencia
    for etiqueta in etiquetas:
        for r in _buscar(pagina, etiqueta):
            if y_referencia - margen_arriba <= r.y0 <= y_referencia + margen_abajo:
                y_fin = max(y_fin, r.y0)
    return y_fin + 3


def _fin_de_tabla(pagina, y_desde: float) -> float:
    """Donde termina una tabla de bienes (construcciones/instalaciones):
    en el primer 'Total' de la izquierda (la fila de suma -o, si el predio
    tiene mas de un rol, el primer SUBtotal: Fase 1 ya esta limitada a un
    solo predio, asi que pararse ahi es lo mismo que ya se avisa en otro
    lado) o, si la tabla no rotula su suma con la palabra 'Total' -pasa en
    los dos PDF reales, no es un caso raro-, en 'Observaciones:'."""
    candidatos = [r.y0 for r in _buscar(pagina, "Total") if r.y0 > y_desde and r.x0 < 100]
    candidatos += [r.y0 for r in _buscar(pagina, "Observaciones:") if r.y0 > y_desde]
    return (min(candidatos) - 2) if candidatos else pagina.rect.y1


def _construcciones(pagina) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """La tabla '5.1- Construcciones': Bien / Tipo Construcción / Material /
    Año / Descripción y ubicación (rol) / Tipo de Garantía / Superficie
    (m2) / Valor por m2 / Total $ / Total UF / Asegurable.

    Una sola tabla, no dos -el esquema pide listas separadas de
    hipotecables y prendables (`construcciones_hipotecables`,
    `construcciones_prendables`), pero en la plantilla real es UNA tabla
    con una columna 'Tipo de Garantía' que dice 'Hipoteca' o 'Prenda' por
    fila-. Se lee una vez y se reparte con esa columna como discriminador;
    una fila sin ese dato no se sabe a cual lista pertenece y se descarta
    -mismo criterio que el resto: no inventar, avisar-.

    `costo_reposicion_uf_m2` no viene directo: se deriva de Total UF /
    Superficie -mismo principio que suelos (§16.a.centies.sexies), nunca de
    un tipo de cambio supuesto-. `vida_util_total_anios` (obligatorio en el
    esquema) no tiene de donde salir en este PDF -no es una columna, es una
    tabla de parámetros por tipo de material que vive en el motor PLANT de
    la plantilla-, así que la fila queda sin ese campo: `leer-anterior`
    devuelve un borrador SUELTO (`Dict[str, Any]`, sin validar contra
    `RetasacionRabobank` todavía -api/routers/retasaciones.py), y el
    tasador lo completa en el Paso 2, igual que cualquier dato que el
    archivo antiguo no traía."""
    rects_sup = _buscar(pagina, "Superficie")
    if not rects_sup:
        return []
    y_header = rects_sup[0].y0
    x_superficie = rects_sup[0].x0
    rects_material = [r for r in _buscar(pagina, "Material") if abs(r.y0 - y_header) <= 15]
    x_material = rects_material[0].x0 if rects_material else None
    rects_anio = [r for r in _buscar(pagina, "Año") if abs(r.y0 - y_header) <= 15]
    x_anio = rects_anio[0].x0 if rects_anio else None

    y_header_fin = _bloque_de_encabezado(
        pagina, y_header, "Tipo", "Construcción", "Material", "Año",
        "Descripción", "Descripcion", "ubicación", "ubicacion", "Garantía", "Asegurable")
    y_fin_tabla = _fin_de_tabla(pagina, y_header_fin)

    anclas = sorted(
        (p[1], p[4]) for p in _palabras(pagina)
        if y_header_fin < p[1] < y_fin_tabla and x_superficie - 8 <= p[0] < x_superficie + 45
        and _NUM_SIMPLE_RE.match(p[4]))

    hipotecables: List[Dict[str, Any]] = []
    prendables: List[Dict[str, Any]] = []
    for i, (y_ancla, texto_ancla) in enumerate(anclas):
        y_ini = (anclas[i - 1][0] + y_ancla) / 2 if i > 0 else y_header_fin
        y_fin = (anclas[i + 1][0] + y_ancla) / 2 if i + 1 < len(anclas) else y_fin_tabla
        palabras_fila = sorted(
            (p for p in _palabras(pagina) if y_ini < p[1] < y_fin),
            key=lambda p: (p[1], p[0]))

        zona_izq = [p for p in palabras_fila if p[0] < (x_material or x_superficie) - 40]
        edificacion = None
        if zona_izq:
            y0 = zona_izq[0][1]
            edificacion = " ".join(p[4] for p in zona_izq if abs(p[1] - y0) <= 2)
        if not edificacion:
            continue  # la fila de suma no tiene nombre de bien: se descarta

        material = None
        if x_material:
            candidatas_material = [
                p for p in palabras_fila
                if (x_material - 5) <= p[0] < (x_anio or x_material + 60) - 5]
            if candidatas_material:
                y0 = candidatas_material[0][1]
                material = " ".join(p[4] for p in candidatas_material if abs(p[1] - y0) <= 2)

        anio = next((int(p[4]) for p in palabras_fila
                    if x_anio and abs(p[0] - x_anio) < 30 and re.match(r"^(19|20)\d{2}$", p[4])),
                   None)

        rol = _rol_en(palabras_fila)

        # Los campos numericos solo se buscan cerca de la linea del propio
        # ancla -la fila de suma sin rotulo "Total" de Ejemplo cae dentro
        # de la ventana de la ULTIMA fila real, y sin este limite su "$" se
        # lee como si fuera el total de esa fila (verificado).
        misma_linea = [p for p in palabras_fila if abs(p[1] - y_ancla) <= 5]
        tipo_garantia = next((p[4] for p in misma_linea if p[4] in ("Hipoteca", "Prenda")), None)
        if tipo_garantia is None:
            continue  # sin saber a cual lista pertenece, no se precarga

        superficie = _numero(texto_ancla)
        total_uf = None
        uf_rect = next((p for p in misma_linea if p[4] == "UF"), None)
        if uf_rect:
            total_uf = _numero(_numero_despues(misma_linea, uf_rect))
        asegurable = next((p[4] for p in misma_linea if p[4] in ("Sí", "Si", "No")), None)

        dato: Dict[str, Any] = {"edificacion": edificacion, "superficie_m2": superficie}
        if material:
            dato["estructura_vertical"] = material
        if anio:
            dato["anio_dato_aprox"] = anio
        if rol:
            dato["ubicacion_rol"] = rol
        if asegurable:
            dato["asegurable"] = asegurable in ("Sí", "Si")
        if total_uf is not None and superficie:
            dato["costo_reposicion_uf_m2"] = round(total_uf / superficie, 4)

        (hipotecables if tipo_garantia == "Hipoteca" else prendables).append(dato)
    return hipotecables, prendables


def _instalaciones(pagina) -> List[Dict[str, Any]]:
    """La tabla '5.2.- Instalaciones y/o Obras Complementarias': Instalación
    / Unidad / Descripción y ubicación (rol) / Tipo de Garantía / Valor por
    Unidad / Total $ / Total UF / Asegurable. Misma tecnica que
    _construcciones, con dos diferencias medidas en los PDF reales:

    - el '$' viene PEGADO al numero ('$2,500,000', un solo token), no
      suelto seguido del numero como en construcciones/aguas.
    - el 'Total UF' a veces no trae el rotulo 'UF' -es el numero que sigue
      pegado al 'Total $', sin mas (verificado en Ejemplo; Hijuela sí trae
      el rotulo).

    A diferencia de _construcciones, esta tabla NO se reparte en dos listas
    -el esquema solo tiene `instalaciones`, una-, así que la columna 'Tipo
    de Garantía' (que en los dos PDF reales dice 'Prenda' siempre: las
    instalaciones no son hipotecables) no hace falta como discriminador,
    solo se ignora. `unidad` (obligatorio en el esquema) tampoco tiene de
    donde salir -la tabla real no imprime la unidad por fila, solo 'ha.'/
    'gl.' como referencia en el encabezado-, así que queda para que el
    tasador la complete, mismo criterio que `vida_util_total_anios` en
    construcciones."""
    rects_unidad = _buscar(pagina, "Unidad")
    if not rects_unidad:
        return []
    y_header = rects_unidad[0].y0
    x_unidad = rects_unidad[0].x0

    y_header_fin = _bloque_de_encabezado(
        pagina, y_header, "Instalación", "Instalacion", "Obras", "Unidad",
        "Descripción", "Descripcion", "ubicación", "ubicacion", "Garantía", "Garantia",
        "Valor", "Asegurable")
    y_fin_tabla = _fin_de_tabla(pagina, y_header_fin)

    anclas = sorted(
        (p[1], p[4]) for p in _palabras(pagina)
        if y_header_fin < p[1] < y_fin_tabla and x_unidad - 8 <= p[0] < x_unidad + 45
        and _NUM_SIMPLE_RE.match(p[4]))

    filas = []
    for i, (y_ancla, texto_ancla) in enumerate(anclas):
        y_ini = (anclas[i - 1][0] + y_ancla) / 2 if i > 0 else y_header_fin
        y_fin = (anclas[i + 1][0] + y_ancla) / 2 if i + 1 < len(anclas) else y_fin_tabla
        palabras_fila = sorted(
            (p for p in _palabras(pagina) if y_ini < p[1] < y_fin),
            key=lambda p: (p[1], p[0]))

        zona_izq = [p for p in palabras_fila if p[0] < x_unidad - 30]
        item = None
        if zona_izq:
            y0 = zona_izq[0][1]
            item = " ".join(p[4] for p in zona_izq if abs(p[1] - y0) <= 2)
        if not item:
            continue

        rol = _rol_en(palabras_fila)
        misma_linea = [p for p in palabras_fila if abs(p[1] - y_ancla) <= 5]
        cantidad = _numero(texto_ancla)
        dolares = sorted((p for p in misma_linea if "$" in p[4]), key=lambda p: p[0])

        def _valor_de(rect_dolar):
            if rect_dolar[4] != "$":
                return _numero(rect_dolar[4])
            return _numero(_numero_despues(misma_linea, rect_dolar))

        valor_unitario = _valor_de(dolares[0]) if dolares else None
        total_uf = None
        uf_rect = next((p for p in misma_linea if p[4] == "UF"), None)
        if uf_rect:
            total_uf = _numero(_numero_despues(misma_linea, uf_rect))
        elif len(dolares) >= 2:
            total_uf = _numero(_numero_despues(misma_linea, dolares[-1]))
        asegurable = next((p[4] for p in misma_linea if p[4] in ("Sí", "Si", "No")), None)

        if not cantidad:
            continue
        dato: Dict[str, Any] = {"item": item, "cantidad": cantidad}
        if rol:
            dato["ubicacion_rol"] = rol
        if asegurable:
            dato["asegurable"] = asegurable in ("Sí", "Si")
        if total_uf is not None and cantidad:
            dato["costo_reposicion_unitario_uf"] = round(total_uf / cantidad, 4)
        filas.append(dato)
    return filas


def _testigos_propiedades(pagina) -> List[Dict[str, Any]]:
    """La pagina '8.- REFERENCIAS DE MERCADO (BIEN RAIZ)': Propiedad /
    Superficie / Valor Total / Región / Comuna / Observaciones (+ un link
    de la oferta/venta debajo de cada fila).

    **La unidad de 'Valor Total' cambia entre PDF -verificado, no
    adivinado-**: Ejemplo la trae en pesos ('Valor Total $'); Hijuela la
    trae en UF directo ('Valor Total UF'), y ademas declara su propia tasa
    ('UF: $37.291,74') arriba de la tabla. Se detecta buscando 'UF' pegado
    al encabezado 'Total'; si no esta, se busca la tasa declarada en la
    PROPIA pagina y se convierte con ESA -nunca con la tasa de la tasacion
    nueva, que todavia no existe en este punto del flujo-. Si ninguna de
    las dos esta (el caso real de Ejemplo, que no declara tasa), el dato
    en pesos se guarda en 'valor_total_clp_sin_convertir' -un campo FUERA
    del esquema, visible igual porque `leer-anterior` devuelve un borrador
    suelto sin validar (api/routers/retasaciones.py)- en vez de inventar
    una conversion.

    Mismo par de bugs ya encontrados en otras tablas de esta pagina, y las
    mismas correcciones:

    - **Dos numeros de la misma fila con un jitter de Y menor a 1pt**
      -'Valor Total' y 'Valor unit.' pueden invertirse si se toma el
      primero en orden (Y, X) en vez de ordenar por X-.
    - **Un numero mas largo arranca mas a la izquierda** porque la columna
      esta alineada a la derecha (verificado: '1.789.607.350', 13
      caracteres, arranca 9pt antes que '550.000.000', 11 caracteres, en
      la MISMA columna) -el margen de busqueda tiene que darle cabida.

    'region'/'comuna' pueden salir con texto de Observaciones mezclado -esa
    columna de texto libre, larga, esta pegada a la derecha de las dos-;
    se acota con la posicion real de 'Observaciones' en vez de un margen
    fijo, pero no queda perfecto en todos los casos. Son campos opcionales
    y el tasador los revisa junto con el resto."""
    rects_prop = _buscar(pagina, "Propiedad")
    if not rects_prop:
        return []
    y_header = rects_prop[0].y0
    x_propiedad = rects_prop[0].x0
    x_superficie = _pos_header(pagina, y_header, "Superficie", "Superf.", x_min=x_propiedad + 5)
    x_valortotal = _pos_header(pagina, y_header, "Total", x_min=(x_superficie or 0) + 5)
    x_region = _pos_header(pagina, y_header, "Región", "Region", x_min=(x_valortotal or 0) + 40)
    x_comuna = _pos_header(pagina, y_header, "Comuna", x_min=(x_region or 0) + 5)
    x_obs = _pos_header(pagina, y_header, "Observaciones", x_min=(x_comuna or 0) + 5)
    if not (x_superficie and x_valortotal and x_region and x_comuna):
        return []

    es_uf = any(abs(r.y0 - y_header) <= 15 for r in _buscar(pagina, "UF"))
    tasa_uf = None
    if not es_uf:
        rects_tasa = _buscar(pagina, "UF:")
        if rects_tasa:
            r = rects_tasa[0]
            candidatas_tasa = sorted(
                (p for p in _palabras(pagina)
                 if abs(p[1] - r.y0) <= 4 and p[0] > r.x0
                 and _numero(p[4].replace("$", "")) is not None),
                key=lambda p: p[0])
            if candidatas_tasa:
                tasa_uf = _numero(candidatas_tasa[0][4].replace("$", ""))

    es_oferta = bool(_buscar(pagina, "OFERTAS"))

    y_header_fin = _bloque_de_encabezado(
        pagina, y_header, "Superficie", "Superf.", "Valor", "Total", "Región",
        "Region", "Comuna", "Observaciones", "UF", "Há")
    y_fin_candidatos = [r.y0 for r in _buscar(pagina, "Muestra") if r.y0 > y_header]
    y_fin_tabla = (min(y_fin_candidatos) - 10) if y_fin_candidatos else pagina.rect.y1

    anclas = sorted(
        (p[1], p[4]) for p in _palabras(pagina)
        if y_header_fin < p[1] < y_fin_tabla and x_superficie - 8 <= p[0] < x_superficie + 45
        and _NUM_SIMPLE_RE.match(p[4]))

    filas = []
    for i, (y_ancla, texto_ancla) in enumerate(anclas):
        y_ini = (anclas[i - 1][0] + y_ancla) / 2 if i > 0 else y_header_fin
        y_fin = (anclas[i + 1][0] + y_ancla) / 2 if i + 1 < len(anclas) else y_fin_tabla
        palabras_fila = sorted(
            (p for p in _palabras(pagina) if y_ini < p[1] < y_fin),
            key=lambda p: (p[1], p[0]))
        misma_linea = [p for p in palabras_fila if abs(p[1] - y_ancla) <= 5]

        ubicacion = " ".join(p[4] for p in misma_linea if p[0] < x_superficie - 15) or None
        if not ubicacion:
            continue

        # -16: un numero mas largo arranca mas a la izquierda (columna
        # alineada a la derecha). Ordenado por X, no por (y,x): ver
        # docstring.
        candidatas_valor = sorted(
            (p for p in misma_linea if (x_valortotal - 16) <= p[0] < x_region - 30),
            key=lambda p: p[0])
        valor_texto = candidatas_valor[0][4] if candidatas_valor else None
        valor_num = _numero((valor_texto or "").replace("$", "").replace("UF", "").strip())

        region = " ".join(p[4] for p in misma_linea
                          if (x_region - 3) <= p[0] < x_comuna - 5) or None
        limite_comuna_fin = (x_obs - 20) if x_obs else (x_comuna + 80)
        comuna = " ".join(p[4] for p in misma_linea
                          if x_comuna <= p[0] < limite_comuna_fin) or None
        link = next((p[4] for p in palabras_fila if p[4].startswith("http")), None)
        superficie_ha = _numero(texto_ancla)

        dato: Dict[str, Any] = {"ubicacion": ubicacion}
        if superficie_ha:
            dato["superficie_m2"] = round(superficie_ha * 10000, 2)
        if region:
            dato["region"] = region
        if comuna:
            dato["comuna"] = comuna
        if link:
            dato["link"] = link
        if es_oferta:
            dato["tipo_referencia"] = "oferta"
        if valor_num:
            if es_uf:
                dato["valor_total_uf"] = valor_num
            elif tasa_uf:
                dato["valor_total_uf"] = round(valor_num / tasa_uf, 2)
            else:
                dato["valor_total_clp_sin_convertir"] = valor_num
        filas.append(dato)
    return filas


def _testigos_derechos_agua(pagina) -> List[Dict[str, Any]]:
    """La pagina '8.1- REFERENCIAS DE COMPRAVENTA DERECHOS AGUAS
    SUPERFICIALES': D° Aguas / Caudal L/s / Valor Total UF / Valor unit. $
    por 1 L/s / Región / Comuna / Observaciones.

    A diferencia de la tabla de propiedades, esta SI trae 'Valor Total' en
    UF directo en los dos PDF reales -no hizo falta el camino de
    conversion-, pero 'Valor unit.' sigue en pesos. `valor_unitario_uf_ls`
    y `valor_total_m_pesos` (los dos campos del esquema) se derivan de ahi
    con una division/multiplicacion simple -no de una tasa externa: los dos
    numeros de origen ya vienen del PDF-. Verificado cruzando los dos
    caminos: Valor Total UF / Caudal y Valor unit.$ × Caudal / tasa UF del
    propio PDF dan resultados consistentes entre si."""
    rects_caudal = _buscar(pagina, "Caudal")
    if not rects_caudal:
        return []
    y_header = rects_caudal[0].y0
    x_caudal = rects_caudal[0].x0
    x_valortotal = _pos_header(pagina, y_header, "Total", x_min=x_caudal + 5)
    x_valorunit = _pos_header(pagina, y_header, "unit.", x_min=(x_valortotal or 0) + 5)
    x_region = _pos_header(pagina, y_header, "Región", "Region", x_min=(x_valorunit or 0) + 5)
    x_comuna = _pos_header(pagina, y_header, "Comuna", x_min=(x_region or 0) + 5)
    x_obs = _pos_header(pagina, y_header, "Observaciones", x_min=(x_comuna or 0) + 5)
    if not (x_valortotal and x_valorunit and x_region and x_comuna):
        return []

    y_header_fin = _bloque_de_encabezado(
        pagina, y_header, "Valor", "Total", "unit.", "Región", "Region", "Comuna",
        "Observaciones", "UF", "L/s")
    y_fin_candidatos = [r.y0 for r in _buscar(pagina, "Muestra") if r.y0 > y_header]
    y_fin_tabla = (min(y_fin_candidatos) - 10) if y_fin_candidatos else pagina.rect.y1

    anclas = sorted(
        (p[1], p[4]) for p in _palabras(pagina)
        if y_header_fin < p[1] < y_fin_tabla and x_caudal - 8 <= p[0] < x_caudal + 45
        and _NUM_SIMPLE_RE.match(p[4]))

    filas = []
    for i, (y_ancla, texto_ancla) in enumerate(anclas):
        y_ini = (anclas[i - 1][0] + y_ancla) / 2 if i > 0 else y_header_fin
        y_fin = (anclas[i + 1][0] + y_ancla) / 2 if i + 1 < len(anclas) else y_fin_tabla
        palabras_fila = sorted(
            (p for p in _palabras(pagina) if y_ini < p[1] < y_fin),
            key=lambda p: (p[1], p[0]))
        misma_linea = [p for p in palabras_fila if abs(p[1] - y_ancla) <= 5]

        fuente = " ".join(p[4] for p in misma_linea if p[0] < x_caudal - 15) or None
        if not fuente:
            continue
        fuente = re.sub(r"^\d+\.\s*", "", fuente)  # saca el "1. " inicial

        candidatas_uf = sorted(
            (p for p in misma_linea if (x_valortotal - 16) <= p[0] < x_valorunit - 5),
            key=lambda p: p[0])
        valor_total_uf = _numero(candidatas_uf[0][4]) if candidatas_uf else None

        candidatas_clp = sorted(
            (p for p in misma_linea if (x_valorunit - 16) <= p[0] < x_region - 30),
            key=lambda p: p[0])
        valor_unit_clp = _numero(candidatas_clp[0][4]) if candidatas_clp else None

        region = " ".join(p[4] for p in misma_linea
                          if (x_region - 3) <= p[0] < x_comuna - 5) or None
        limite_comuna_fin = (x_obs - 20) if x_obs else (x_comuna + 80)
        comuna = " ".join(p[4] for p in misma_linea
                          if x_comuna <= p[0] < limite_comuna_fin) or None
        link = next((p[4] for p in palabras_fila if p[4].startswith("http")), None)

        caudal = _numero(texto_ancla)
        if not caudal:
            continue
        dato: Dict[str, Any] = {"fuente": fuente, "caudal_nominal_ls": caudal}
        if valor_total_uf:
            dato["valor_unitario_uf_ls"] = round(valor_total_uf / caudal, 4)
        if valor_unit_clp:
            dato["valor_total_m_pesos"] = round(caudal * valor_unit_clp / 1000, 2)
        if region:
            dato["region"] = region
        if comuna:
            dato["comuna"] = comuna
        if link:
            dato["link"] = link
        filas.append(dato)
    return filas


def _tasador(pagina_resumen) -> Dict[str, Any]:
    nombre = _valor_a_la_derecha(pagina_resumen, "TASADOR", tolerancia_y=6)
    return {"nombre": nombre} if nombre else {}


def _resumen_texto(pagina_resumen) -> Optional[str]:
    """Un resumen legible de los totales de la tasacion ANTERIOR, para
    avisar -nunca para precargar un campo: estos numeros no tienen una
    celda propia en RetasacionRabobank, los vuelve a calcular el Excel
    nuevo a partir del inventario que declare el tasador."""
    fecha = _valor_a_la_derecha(pagina_resumen, "FECHA", tolerancia_y=6)
    uf = _valor_a_la_derecha(pagina_resumen, "VALOR UF", tolerancia_y=6)
    usd = _valor_a_la_derecha(pagina_resumen, "VALOR DÓLAR", tolerancia_y=6)
    if not any([fecha, uf, usd]):
        return None
    partes = []
    if fecha:
        partes.append("tasada el {}".format(fecha))
    if uf:
        partes.append("UF base {}".format(uf))
    if usd:
        partes.append("USD {}".format(usd))
    return "Tasación anterior: " + ", ".join(partes) + "."


def _diferencias(doc) -> Dict[str, str]:
    salida = {}
    for campo, marca in _MARCA_SECCION.items():
        pagina = _pagina_con(doc, marca)
        if pagina is None:
            continue
        texto = _bloque_de_texto(pagina, _ETIQUETA_DIFERENCIAS)
        if texto:
            salida[campo] = texto
    return salida


def _hay_varios_predios(doc) -> bool:
    for pagina in doc:
        if _buscar(pagina, "PREDIO 2"):
            return True
    return False


def leer_retasacion_pdf(
    archivo: Union[str, Any],
) -> Tuple[Dict[str, Any], List[str]]:
    """Devuelve (datos, avisos) -mismo contrato que
    importar_retasacion.leer_retasacion_excel, para que el Paso 1 use el
    mismo camino en la pantalla sea cual sea el formato subido.

    'archivo' puede ser una ruta, o un objeto tipo-archivo (UploadFile.file
    de FastAPI trae .read()); PyMuPDF acepta ambos via 'stream'.
    """
    avisos: List[str] = []
    if hasattr(archivo, "read"):
        doc = fitz.open(stream=archivo.read(), filetype="pdf")
    else:
        doc = fitz.open(archivo)
    try:
        if doc.page_count == 0:
            return {}, ["el PDF no tiene páginas"]

        datos: Dict[str, Any] = {}
        pagina1 = doc[0]

        parametros: Dict[str, Any] = {}
        region = _valor_a_la_derecha(pagina1, "Región")
        if region:
            parametros["region"] = region
        ejecutivo = _valor_a_la_derecha(pagina1, "EJECUTIVO :")
        if ejecutivo:
            parametros["ejecutivo"] = ejecutivo
        n_solicitud = _valor_a_la_derecha(pagina1, "SOLICITUD :")
        if n_solicitud:
            parametros["n_solicitud"] = n_solicitud

        cliente: Dict[str, Any] = {}
        nombre_cliente = _valor_a_la_derecha(pagina1, "Cliente:")
        if nombre_cliente:
            cliente["nombre"] = nombre_cliente
        rut_cliente = _valor_a_la_derecha(pagina1, "Rut:", indice=0)
        if rut_cliente:
            cliente["rut"] = rut_cliente
        propietario = _valor_a_la_derecha(pagina1, "Propietario:")
        if propietario:
            cliente["propietario"] = propietario
        rut_propietario = _valor_a_la_derecha(pagina1, "Rut:", indice=1)
        if rut_propietario:
            cliente["rut_propietario"] = rut_propietario

        predio: Dict[str, Any] = {}
        predio.update(_predio_comuna_superficie(pagina1))
        if predio.get("comuna") and "comuna" not in parametros:
            parametros["comuna"] = predio["comuna"]

        pagina_resumen = _pagina_con(doc, ("RESUMEN", "DE", "VALORES"))
        if pagina_resumen is not None:
            predio.update(_nombre_y_rol_predio(pagina_resumen))
            tasador = _tasador(pagina_resumen)
            if tasador:
                datos["tasador"] = tasador
            resumen = _resumen_texto(pagina_resumen)
            if resumen:
                avisos.append(resumen)
        else:
            avisos.append("no se encontró la página de Resumen de Valores")

        pagina_inmuebles = _pagina_con(doc, ("INMUEBLES",))
        if pagina_inmuebles is not None:
            insc = _inscripcion_y_rol_predio(pagina_inmuebles)
            superficie_insc = insc.pop("superficie_legal_ha", None)
            if superficie_insc is not None and "superficie_legal_ha" not in predio:
                predio["superficie_legal_ha"] = superficie_insc
            if insc.get("rol_avaluo") and "rol_avaluo" not in predio:
                predio["rol_avaluo"] = insc["rol_avaluo"]
            campos_insc = {k: v for k, v in insc.items()
                          if k in ("fojas", "numero", "anio", "conservador")}
            if campos_insc:
                predio["inscripcion"] = campos_insc

            deslindes = _deslindes(pagina_inmuebles)
            for lado, campo in (("Norte", "deslinde_norte"), ("Sur", "deslinde_sur"),
                                ("Oriente", "deslinde_oriente"),
                                ("Poniente", "deslinde_poniente")):
                if deslindes.get(lado):
                    predio[campo] = deslindes[lado]
        else:
            avisos.append("no se encontró la página de Inmuebles -inscripción "
                          "CBR y deslindes quedan sin precargar")

        if _hay_varios_predios(doc):
            avisos.append(
                "el PDF tiene más de un predio -se precargó solo el primero "
                "({}); complete el resto a mano".format(predio.get("nombre", "?")))

        if parametros:
            datos["parametros"] = parametros
        if cliente:
            datos["cliente"] = cliente
        if predio:
            datos["predio"] = predio

        pagina_plantaciones = _pagina_con(doc, ("6.-", "Plantaciones"))
        if pagina_plantaciones is not None:
            sin_royalty = _plantaciones_sin_royalty(pagina_plantaciones)
            if sin_royalty:
                datos["plantaciones_sin_royalty"] = sin_royalty
            con_royalty = _plantaciones_con_royalty(pagina_plantaciones)
            if con_royalty:
                datos["plantaciones_con_royalty"] = con_royalty

        pagina_suelos = _pagina_con(doc, ("SUELOS",))
        if pagina_suelos is not None:
            suelos = _suelos(pagina_suelos)
            if suelos:
                datos["suelos"] = suelos

        pagina_aguas = _pagina_con(doc, _MARCA_SECCION["aguas"])
        if pagina_aguas is not None:
            aguas, observaciones_aguas = _aguas(pagina_aguas)
            if aguas:
                datos["aguas"] = aguas
            if observaciones_aguas:
                datos["aguas_observaciones"] = observaciones_aguas

        pagina_construcciones = _pagina_con(doc, _MARCA_SECCION["construcciones"])
        if pagina_construcciones is not None:
            hipotecables, prendables = _construcciones(pagina_construcciones)
            if hipotecables:
                datos["construcciones_hipotecables"] = hipotecables
            if prendables:
                datos["construcciones_prendables"] = prendables

        pagina_instalaciones = _pagina_con(doc, _MARCA_SECCION["instalaciones"])
        if pagina_instalaciones is not None:
            instalaciones = _instalaciones(pagina_instalaciones)
            if instalaciones:
                datos["instalaciones"] = instalaciones

        # El "8.1-" de las aguas no siempre existe -Ejemplo no lo trae, su
        # propia referencia de aguas queda mezclada en la misma pagina que
        # la de propiedades, sin la estructura limpia de columnas que tiene
        # Hijuela-: se degrada a lista vacia en vez de forzar un formato
        # que no calza, mismo criterio que Suelos con el predio multiple.
        pagina_testigos_prop = _pagina_con(doc, ("8.-", "REFERENCIAS", "DE", "MERCADO"))
        if pagina_testigos_prop is not None:
            testigos_prop = _testigos_propiedades(pagina_testigos_prop)
            if testigos_prop:
                datos["testigos_propiedades"] = testigos_prop

        pagina_testigos_agua = _pagina_con(
            doc, ("8.1-", "REFERENCIAS", "DE", "COMPRAVENTA"))
        if pagina_testigos_agua is not None:
            testigos_agua = _testigos_derechos_agua(pagina_testigos_agua)
            if testigos_agua:
                datos["testigos_derechos_agua"] = testigos_agua

        diferencias = _diferencias(doc)
        if diferencias:
            datos["diferencias"] = diferencias

        if not datos:
            avisos.append(
                "no se reconoció ningún dato -puede ser un PDF de otro "
                "formato de retasación: revise y complete a mano")
        else:
            avisos.append(
                "lectura de PDF: cubre identificación del cliente/predio, "
                "inscripción CBR, deslindes, plantaciones, suelos, derechos "
                "de agua, construcciones, instalaciones, testigos de "
                "mercado, y las diferencias con la tasación anterior -en un "
                "predio con formato de tabla poco usual, alguna de estas "
                "puede quedar vacía en vez de forzar un dato que no calza; "
                "revísela. La fuente de cada derecho de agua, y la región/"
                "comuna de cada testigo, pueden salir incompletas si el PDF "
                "repite texto en columnas vecinas -revíselas contra el "
                "original-. El valor de un testigo de propiedad que el PDF "
                "trae en pesos y no declara una tasa UF propia queda en "
                "'valor_total_clp_sin_convertir' en vez de un UF inventado. "
                "La inscripción CBR de las aguas, la vida útil de "
                "construcciones/instalaciones y la unidad de cada "
                "instalación no se precargan todavía -son datos que no "
                "están en el PDF como columna propia, complételos a mano.")

        return datos, avisos
    finally:
        doc.close()
