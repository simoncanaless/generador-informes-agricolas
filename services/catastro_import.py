"""Importacion de catastros de plantaciones ajenos al formato de la casa.

El agricultor entrega su catastro en la planilla que a el le acomoda; el tasador
lo transcribe al formato del informe. Este modulo automatiza esa transcripcion.

Reparto de trabajo, fiel a la regla 80/20:

  - **La IA, si acaso, mira solo los ENCABEZADOS.** Y solo los que el diccionario
    de sinonimos no supo reconocer. Su unica frase es del tipo "la columna que
    dice CUARTEL corresponde a Predio".
  - **Los datos los copia codigo puro.** Ni una superficie, ni un anio, ni una
    variedad pasa por un modelo. Si el catastro trae 400 hileras, las 400 las
    mueve Python. Es imposible que se invente una hectarea.

Esa separacion no es purismo: es lo que permite firmar el informe. Una cifra
alucinada en una superficie es un error que nadie detecta hasta que es tarde.
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from schemas.informe import Plantacion

__all__ = [
    "ReglasAgronomicas", "cargar_reglas", "LecturaCatastro", "CatastroImportado",
    "leer_catastro", "clasificar_madurez", "CatastroIlegible",
    "proponer_mapeo", "PropuestaMapeo", "ColumnaPropuesta", "huella_encabezados",
    "formas_de_nombre",
    "DESCRIPCION_CAMPOS",
]

#: Hasta que fila se busca el encabezado. Los catastros suelen traer titulo,
#: logo y un par de lineas en blanco antes de la tabla.
FILAS_BUSQUEDA_ENCABEZADO = 20

#: Minimo de columnas reconocidas para dar una fila por encabezado valido.
MINIMO_COLUMNAS_RECONOCIDAS = 3

#: Filas vacias seguidas tras las que se da por terminada la tabla.
MAXIMO_VACIAS_SEGUIDAS = 8


class CatastroIlegible(RuntimeError):
    """El archivo no contiene nada que se parezca a un catastro."""


# ------------------------------------------------------------ normalizado

def normalizar(texto: Any) -> str:
    """Deja un encabezado comparable: sin tildes, sin saltos, sin puntuacion.

    Los encabezados reales vienen con saltos de linea incrustados
    ('Año \\nPlantación'), dobles espacios y unidades entre parentesis.
    """
    if texto is None:
        return ""
    plano = str(texto).replace("\n", " ").replace("\r", " ")
    sin_tildes = "".join(
        c for c in unicodedata.normalize("NFD", plano)
        if unicodedata.category(c) != "Mn"
    )
    limpio = re.sub(r"[^a-z0-9 ]+", " ", sin_tildes.lower())
    return re.sub(r"\s+", " ", limpio).strip()


def formas_de_nombre(texto: Any) -> set:
    """Formas normalizadas plausibles de un nombre, para absorber el plural.

    El catalogo de curvas trae 'Cerezos' y 'Nogales'; los catastros y las
    tablas de ODEPA dicen 'Cerezo' y 'Nogal'. Y el plural castellano no es uno
    solo: las palabras terminadas en vocal suman 's' (cerezo -> cerezos) y las
    terminadas en consonante suman 'es' (nogal -> nogales). Por eso se comparan
    CONJUNTOS de formas y no una raiz recortada a ojo.

    Vive aqui y no en cada modulo que la necesita: la escribi otra vez a mano
    en `services/precios.py` -recortando la ultima letra- y fallaba con
    'Cerezos', 'Nogales' y 'Limoneros', o sea con tres de las trece curvas del
    catalogo. Una regla de idioma tiene que estar una sola vez.
    """
    plano = normalizar(texto)
    if not plano:
        return set()
    formas = {plano}
    if plano.endswith("es"):
        formas.add(plano[:-2])
    if plano.endswith("s"):
        formas.add(plano[:-1])
    return formas


def limpiar_texto(valor: Any) -> Optional[str]:
    """Colapsa espacios sobrantes conservando el texto tal cual lo escribieron.

    'Palto ' y 'Hass  ' llegan con espacios de cola; '2,5  x 2,5' con dobles
    espacios internos. Se normaliza el espaciado, NO el contenido: 'Palto
    (rebajado)' se conserva entero porque distingue una plantacion de otra.
    """
    if valor is None:
        return None
    texto = re.sub(r"\s+", " ", str(valor).replace("\n", " ")).strip()
    return texto or None


def limpiar_numero(valor: Any) -> Optional[float]:
    """Convierte a float tolerando coma decimal y ruido de coma flotante.

    Las celdas con formula devuelven su valor cacheado, que arrastra el ruido
    binario habitual (10.440000000000001). Se redondea a 4 decimales: mas
    precision de la que cualquier catastro agricola declara.
    """
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)):
        return round(float(valor), 4)

    texto = str(valor).strip().replace(" ", "")
    if not texto:
        return None
    # 1.234,56 -> 1234.56  |  1234,56 -> 1234.56
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    else:
        texto = texto.replace(",", ".")
    try:
        return round(float(texto), 4)
    except ValueError:
        return None


def limpiar_anio(valor: Any) -> Optional[str]:
    """El anio se guarda como TEXTO porque admite rangos.

    En los catastros reales conviven 2020 (entero) y '1993-2000' o '2017-2018'
    (rango, cuando la plantacion se hizo por etapas). Forzarlo a entero perderia
    informacion que el tasador escribio a proposito.
    """
    if valor is None or valor == "":
        return None
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    if isinstance(valor, int):
        return str(valor)
    texto = limpiar_texto(valor)
    if texto and re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}.*", texto):
        return texto[:4]  # una fecha completa: nos quedamos con el anio
    return texto


def anio_inicial(anio: Optional[str]) -> Optional[int]:
    """Primer anio de un rango: '2017-2018' -> 2017. Para calcular antiguedad."""
    if not anio:
        return None
    encontrado = re.search(r"(19|20)\d{2}", str(anio))
    return int(encontrado.group()) if encontrado else None


# ---------------------------------------------------------------- reglas

@dataclass
class ReglasAgronomicas:
    """Configuracion agronomica externa. Nada de esto vive en el codigo."""

    sinonimos: Dict[str, List[str]] = field(default_factory=dict)
    madurez_por_defecto: Dict[str, int] = field(default_factory=dict)
    madurez_por_especie: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    separadores_calificativo: List[str] = field(default_factory=list)
    patrones_ignorar: List[str] = field(default_factory=list)

    def umbrales(self, especie: str) -> Tuple[int, int]:
        """(formacion_hasta, crecientes_hasta) para una especie."""
        base = normalizar(especie_base(especie, self.separadores_calificativo))
        regla = self.madurez_por_especie.get(base) or self.madurez_por_defecto
        return (
            int(regla.get("formacion_hasta", 3)),
            int(regla.get("crecientes_hasta", 7)),
        )


def cargar_reglas(ruta: Union[str, Path]) -> ReglasAgronomicas:
    with Path(ruta).open(encoding="utf-8") as fh:
        crudo = json.load(fh)

    sinonimos = {
        campo: valores
        for campo, valores in crudo.get("sinonimos_columnas", {}).items()
        if isinstance(valores, list)
    }
    madurez = crudo.get("madurez", {})
    return ReglasAgronomicas(
        sinonimos=sinonimos,
        madurez_por_defecto=madurez.get("por_defecto", {}),
        madurez_por_especie=madurez.get("por_especie", {}),
        separadores_calificativo=crudo.get("especies", {}).get(
            "separadores_calificativo", ["(", " - "]
        ),
        patrones_ignorar=crudo.get("filas_a_ignorar", {}).get("patrones", ["total"]),
    )


def especie_base(especie: str, separadores: Sequence[str]) -> str:
    """'Palto (rebajado)' -> 'Palto'. Solo para agrupar; no altera el catastro."""
    texto = (especie or "").strip()
    for separador in separadores:
        if separador in texto:
            texto = texto.split(separador, 1)[0]
    return texto.strip()


# --------------------------------------------------- deteccion de columnas

@dataclass
class LecturaCatastro:
    """Diagnostico de como se interpreto la planilla."""

    hoja: str
    fila_encabezado: int
    primera_fila: int
    ultima_fila: int
    columnas: Dict[str, str] = field(default_factory=dict)      # campo -> letra
    encabezados: Dict[str, str] = field(default_factory=dict)   # campo -> texto original
    sin_reconocer: Dict[str, str] = field(default_factory=dict) # letra -> texto
    avisos: List[str] = field(default_factory=list)

    def resumen(self) -> str:
        return "hoja {!r}, encabezado en fila {}, datos {}-{}, {} columnas reconocidas".format(
            self.hoja, self.fila_encabezado, self.primera_fila,
            self.ultima_fila, len(self.columnas)
        )


@dataclass
class CatastroImportado:
    plantaciones: List[Plantacion]
    lectura: LecturaCatastro
    superficie_total_ha: float = 0.0
    filas_descartadas: List[str] = field(default_factory=list)


def _emparejar(encabezado: str, sinonimos: Dict[str, List[str]]) -> Optional[str]:
    """Decide a que campo corresponde un encabezado.

    Tres pasadas de menor a mayor tolerancia. La exacta primero para que
    'Superficie (ha)' no se lleve por delante a 'Superficie Total Plantada'.
    """
    clave = normalizar(encabezado)
    if not clave:
        return None

    for campo, variantes in sinonimos.items():
        if clave in {normalizar(v) for v in variantes}:
            return campo

    for campo, variantes in sinonimos.items():
        for variante in variantes:
            normalizada = normalizar(variante)
            if len(normalizada) >= 4 and clave.startswith(normalizada):
                return campo

    for campo, variantes in sinonimos.items():
        for variante in variantes:
            normalizada = normalizar(variante)
            if len(normalizada) >= 5 and normalizada in clave:
                return campo

    # Ultima pasada, al reves: el encabezado es una ABREVIATURA del sinonimo.
    # Las planillas de terreno estan llenas de 'Sup.', 'Var.', 'Año Plant.', y
    # una abreviatura es casi siempre un prefijo de la palabra completa.
    if len(clave) >= 6:
        for campo, variantes in sinonimos.items():
            for variante in variantes:
                if normalizar(variante).startswith(clave):
                    return campo
    return None


def _detectar_encabezado(
    ws, reglas: ReglasAgronomicas, maximo: int
) -> Tuple[int, Dict[str, int], Dict[str, str], Dict[str, str]]:
    """Busca la fila que mas encabezados conocidos contiene."""
    mejor = (0, 0, {}, {}, {})

    for fila in range(1, min(maximo, ws.max_row) + 1):
        columnas: Dict[str, int] = {}
        textos: Dict[str, str] = {}
        sueltas: Dict[str, str] = {}

        for columna in range(1, ws.max_column + 1):
            bruto = ws.cell(fila, columna).value
            etiqueta = limpiar_texto(bruto)
            if not etiqueta:
                continue
            campo = _emparejar(etiqueta, reglas.sinonimos)
            if campo and campo not in columnas:
                columnas[campo] = columna
                textos[campo] = etiqueta
            elif campo is None:
                sueltas[get_column_letter(columna)] = etiqueta

        if len(columnas) > mejor[1]:
            mejor = (fila, len(columnas), columnas, textos, sueltas)

    fila, cuantas, columnas, textos, sueltas = mejor
    if cuantas < MINIMO_COLUMNAS_RECONOCIDAS:
        raise CatastroIlegible(
            "no se encontró una tabla de plantaciones: solo {} columna(s) "
            "reconocida(s) en las primeras {} filas".format(cuantas, maximo)
        )

    # Encabezados a dos alturas: es habitual que una celda este combinada en
    # vertical ('ESPECIE' ocupando A4:A6) y la de al lado escriba su rotulo una
    # fila mas abajo ('VARIEDAD' en B5). Se rescatan de las filas vecinas las
    # columnas que quedaron sin identificar.
    ocupadas = set(columnas.values())
    for vecina in (fila + 1, fila + 2, fila - 1):
        if vecina < 1 or vecina > ws.max_row:
            continue
        for columna in range(1, ws.max_column + 1):
            if columna in ocupadas:
                continue
            etiqueta = limpiar_texto(ws.cell(vecina, columna).value)
            if not etiqueta:
                continue
            campo = _emparejar(etiqueta, reglas.sinonimos)
            if campo and campo not in columnas:
                columnas[campo] = columna
                textos[campo] = etiqueta
                ocupadas.add(columna)
                sueltas.pop(get_column_letter(columna), None)

    return fila, columnas, textos, sueltas


def _es_fila_a_ignorar(texto: Optional[str], patrones: Sequence[str]) -> bool:
    if not texto:
        return False
    clave = normalizar(texto)
    return any(clave.startswith(normalizar(p)) for p in patrones)


# ------------------------------------------------------------- la lectura

def leer_catastro(
    origen: Union[str, Path, bytes],
    reglas: ReglasAgronomicas,
    hoja: Optional[str] = None,
    columnas_forzadas: Optional[Dict[str, str]] = None,
) -> CatastroImportado:
    """Lee un catastro de plantaciones desde un .xlsx cualquiera.

    ``columnas_forzadas`` permite imponer el mapeo (campo -> letra) cuando el
    tasador lo corrige en pantalla o cuando lo propuso la IA.

    Se lee con ``data_only=True``: en los catastros reales hay celdas con
    formula ('=22.44-12' para descontar una superficie replantada) y lo que
    interesa es el numero, no la formula.
    """
    from io import BytesIO

    libro = load_workbook(
        BytesIO(origen) if isinstance(origen, bytes) else origen,
        data_only=True,
    )
    hojas = [hoja] if hoja else libro.sheetnames

    mejor_error: Optional[Exception] = None
    for nombre in hojas:
        if nombre not in libro.sheetnames:
            continue
        try:
            return _leer_hoja(libro[nombre], reglas, columnas_forzadas)
        except CatastroIlegible as error:
            mejor_error = error
            continue

    raise CatastroIlegible(
        "ninguna hoja del libro contiene un catastro reconocible ({})".format(
            mejor_error or "sin detalle"
        )
    )


def _leer_hoja(
    ws, reglas: ReglasAgronomicas, columnas_forzadas: Optional[Dict[str, str]]
) -> CatastroImportado:
    from openpyxl.utils import column_index_from_string

    fila_encabezado, columnas, textos, sueltas = _detectar_encabezado(
        ws, reglas, FILAS_BUSQUEDA_ENCABEZADO
    )
    if columnas_forzadas:
        for campo, letra in columnas_forzadas.items():
            columnas[campo] = column_index_from_string(letra)
            textos.setdefault(campo, "(indicado a mano)")

    if "especie" not in columnas:
        raise CatastroIlegible(
            "no se identifico la columna de especie; sin ella no hay catastro"
        )
    if "superficie_ha" not in columnas:
        raise CatastroIlegible("no se identifico la columna de superficie")

    columna_especie = columnas["especie"]
    plantaciones: List[Plantacion] = []
    descartadas: List[str] = []
    avisos: List[str] = []

    primera = fila_encabezado + 1
    ultima = primera - 1
    vacias = 0

    for fila in range(primera, ws.max_row + 1):
        especie = limpiar_texto(ws.cell(fila, columna_especie).value)
        superficie = limpiar_numero(
            ws.cell(fila, columnas["superficie_ha"]).value
        )

        if not especie and superficie is None:
            vacias += 1
            if vacias > MAXIMO_VACIAS_SEGUIDAS:
                break
            continue
        vacias = 0

        if _es_fila_a_ignorar(especie, reglas.patrones_ignorar):
            descartadas.append("fila {}: '{}' (fila de totales)".format(fila, especie))
            break  # los totales cierran la tabla

        if not especie:
            descartadas.append(
                "fila {}: superficie {} sin especie".format(fila, superficie)
            )
            continue
        if superficie is None:
            descartadas.append("fila {}: '{}' sin superficie".format(fila, especie))
            continue

        def leer(campo: str):
            columna = columnas.get(campo)
            return ws.cell(fila, columna).value if columna else None

        plantaciones.append(Plantacion(
            especie=especie,
            variedad=limpiar_texto(leer("variedad")),
            anio_plantacion=limpiar_anio(leer("anio_plantacion")),
            superficie_ha=superficie,
            sistema_conduccion=limpiar_texto(leer("sistema_conduccion")),
            predio=limpiar_texto(leer("predio")),
            dominio=_tenencia(limpiar_texto(leer("dominio"))),
            riego=limpiar_texto(leer("riego")),
            estado=_estado(limpiar_texto(leer("estado"))),
            densidad_plantacion=limpiar_texto(leer("densidad_plantacion")),
        ))
        ultima = fila

    if not plantaciones:
        raise CatastroIlegible("se reconocio el encabezado pero no hay filas de datos")

    faltantes = [c for c in reglas.sinonimos if c not in columnas]
    if faltantes:
        avisos.append(
            "columnas no encontradas en el origen: {}".format(", ".join(faltantes))
        )
    if sueltas:
        avisos.append(
            "columnas del origen sin equivalencia: {}".format(
                ", ".join("{}='{}'".format(k, v) for k, v in sueltas.items())
            )
        )

    lectura = LecturaCatastro(
        hoja=ws.title,
        fila_encabezado=fila_encabezado,
        primera_fila=primera,
        ultima_fila=ultima,
        columnas={c: get_column_letter(i) for c, i in columnas.items()},
        encabezados=textos,
        sin_reconocer=sueltas,
        avisos=avisos,
    )
    return CatastroImportado(
        plantaciones=plantaciones,
        lectura=lectura,
        superficie_total_ha=round(sum(p.superficie_ha for p in plantaciones), 4),
        filas_descartadas=descartadas,
    )


def _tenencia(valor: Optional[str]):
    """Mapea la tenencia al enum, tolerando como la escriba cada planilla."""
    from schemas.informe import Tenencia

    if not valor:
        return None
    clave = normalizar(valor)
    if clave.startswith("propi"):
        return Tenencia.PROPIETARIO
    if clave.startswith("arrend") or "arriendo" in clave:
        return Tenencia.ARRENDATARIO
    return Tenencia.OTROS


def _estado(valor: Optional[str]):
    from schemas.informe import Estado

    if not valor:
        return None
    clave = normalizar(valor)
    for estado in (Estado.BUENO, Estado.REGULAR, Estado.MALO):
        if clave.startswith(normalizar(estado.value)):
            return estado
    return None


# ------------------------------------------- propuesta de mapeo adaptable

#: Que significa cada campo destino. Se le manda a la IA y se muestra en la web.
DESCRIPCION_CAMPOS = {
    "especie": "especie o cultivo (palto, cerezo, avellano, vid...)",
    "variedad": "variedad o cultivar (Hass, Santina, Giffoni...)",
    "anio_plantacion": "anio de plantacion; admite rangos como '2017-2018'",
    "superficie_ha": "superficie en hectareas (numero)",
    "sistema_conduccion": "sistema de conduccion (copa, eje, espaldera, multieje)",
    "predio": "predio, campo, fundo, cuartel o sector donde esta la plantacion",
    "dominio": "tenencia: propietario o arrendatario",
    "riego": "sistema de riego (goteo, microaspersion, gravitacional)",
    "estado": "estado del huerto (bueno, regular, malo)",
    "densidad_plantacion": "marco o distanciamiento de plantacion, ej. '5 x 2'",
}


@dataclass
class ColumnaPropuesta:
    """Una columna del archivo y a que campo se propone llevarla."""

    letra: str
    encabezado: str
    campo: Optional[str] = None
    origen: str = "sin identificar"   # aprendido | diccionario | ia | sin identificar
    confianza: str = "baja"
    muestras: List[str] = field(default_factory=list)


@dataclass
class PropuestaMapeo:
    """Todo lo que la pantalla necesita para que el tasador confirme o corrija."""

    hoja: str
    fila_encabezado: int
    columnas: List[ColumnaPropuesta] = field(default_factory=list)
    huella: str = ""
    aprendido_de: Optional[str] = None
    filas_estimadas: int = 0
    avisos: List[str] = field(default_factory=list)

    @property
    def mapeo(self) -> Dict[str, str]:
        """campo -> letra, con las columnas que tienen destino asignado."""
        return {c.campo: c.letra for c in self.columnas if c.campo}

    @property
    def completo(self) -> bool:
        """Con especie y superficie ya se puede importar; el resto es opcional."""
        asignados = set(self.mapeo)
        return "especie" in asignados and "superficie_ha" in asignados


def huella_encabezados(encabezados: Sequence[str]) -> str:
    """Identificador estable del formato de un archivo.

    Sirve para recordar el mapeo: si el mismo agricultor manda su planilla el
    mes que viene, los encabezados seran los mismos y se reconoce sola.
    """
    import hashlib

    clave = "|".join(sorted(normalizar(e) for e in encabezados if normalizar(e)))
    return hashlib.sha256(clave.encode("utf-8")).hexdigest()[:16]


def _muestras_de(ws, fila_encabezado: int, columna: int, cuantas: int = 3) -> List[str]:
    """Primeros valores no vacios de una columna, para orientar el mapeo."""
    valores = []
    for fila in range(fila_encabezado + 1, min(fila_encabezado + 25, ws.max_row) + 1):
        valor = ws.cell(fila, columna).value
        if valor is None or not str(valor).strip():
            continue
        valores.append(limpiar_texto(valor) or "")
        if len(valores) >= cuantas:
            break
    return valores


async def proponer_mapeo(
    origen: Union[str, Path, bytes],
    reglas: ReglasAgronomicas,
    hoja: Optional[str] = None,
    ia: Optional[object] = None,
    almacen: Optional[object] = None,
) -> PropuestaMapeo:
    """Propone como leer un catastro cualquiera, en tres pasadas.

    1. **Memoria**: si ya se importo antes un archivo con estos mismos
       encabezados, se reutiliza el mapeo que el tasador confirmo. Gratis y
       exacto.
    2. **Diccionario**: sinonimos conocidos ('CUARTEL' -> predio). Gratis.
    3. **IA**: solo para las columnas que sigan sin identificar, mandandole
       encabezados y tres filas de muestra. Unos cientos de tokens.

    Lo que ninguna de las tres resuelva queda marcado como "sin identificar"
    para que lo asigne el tasador en pantalla. Por eso el sistema se adapta a
    formatos que nadie ha visto nunca: siempre hay una ultima instancia humana.
    """
    from io import BytesIO

    libro = load_workbook(
        BytesIO(origen) if isinstance(origen, bytes) else origen, data_only=True
    )
    hojas = [hoja] if hoja and hoja in libro.sheetnames else libro.sheetnames

    ws = None
    fila_encabezado = 0
    columnas_dic: Dict[str, int] = {}
    textos: Dict[str, str] = {}
    sueltas: Dict[str, str] = {}
    ultimo_error = None

    for nombre in hojas:
        try:
            candidata = libro[nombre]
            fila_encabezado, columnas_dic, textos, sueltas = _detectar_encabezado(
                candidata, reglas, FILAS_BUSQUEDA_ENCABEZADO
            )
            ws = candidata
            break
        except CatastroIlegible as error:
            ultimo_error = error
            continue

    if ws is None:
        raise CatastroIlegible(
            "ninguna hoja parece un catastro de plantaciones ({})".format(ultimo_error)
        )

    # Todas las columnas con encabezado, tengan destino o no.
    encabezados: Dict[str, str] = {}
    for columna in range(1, ws.max_column + 1):
        etiqueta = limpiar_texto(ws.cell(fila_encabezado, columna).value)
        if not etiqueta:
            for vecina in (fila_encabezado + 1, fila_encabezado - 1):
                if 1 <= vecina <= ws.max_row:
                    etiqueta = limpiar_texto(ws.cell(vecina, columna).value)
                    if etiqueta:
                        break
        if etiqueta:
            encabezados[get_column_letter(columna)] = etiqueta

    propuesta = PropuestaMapeo(
        hoja=ws.title,
        fila_encabezado=fila_encabezado,
        huella=huella_encabezados(list(encabezados.values())),
    )

    por_letra = {get_column_letter(i): campo for campo, i in columnas_dic.items()}

    # (1) memoria
    recordado: Dict[str, str] = {}
    if almacen is not None:
        guardado = await almacen.buscar_mapeo(propuesta.huella)
        if guardado is not None:
            recordado = {letra: campo for campo, letra in guardado.columnas.items()}
            propuesta.aprendido_de = guardado.nombre
            # Sin aviso aparte: la pantalla ya lo dice a partir de
            # `aprendido_de`, y los dos juntos repetian la misma frase en dos
            # cajas seguidas.

    for letra, etiqueta in sorted(encabezados.items()):
        columna = ColumnaPropuesta(
            letra=letra, encabezado=etiqueta,
            muestras=_muestras_de(ws, fila_encabezado,
                                  column_index_from_string(letra)),
        )
        if letra in recordado:
            columna.campo = recordado[letra]
            columna.origen = "aprendido"
            columna.confianza = "alta"
        elif letra in por_letra:
            columna.campo = por_letra[letra]
            columna.origen = "diccionario"
            columna.confianza = "alta"
        propuesta.columnas.append(columna)

    # (3) IA solo para lo que quedo sin identificar
    pendientes = [c for c in propuesta.columnas if c.campo is None]
    usados = {c.campo for c in propuesta.columnas if c.campo}
    disponibles = {k: v for k, v in DESCRIPCION_CAMPOS.items() if k not in usados}

    if pendientes and disponibles and ia is not None:
        try:
            respuesta = await ia.mapear_columnas(
                encabezados={c.letra: c.encabezado for c in pendientes},
                muestras=[
                    {c.letra: (c.muestras[i] if i < len(c.muestras) else "")
                     for c in pendientes}
                    for i in range(3)
                ],
                campos=disponibles,
            )
            sugeridos = {
                col.letra: col for col in respuesta.valor.columnas if col.campo
            }
            for columna in pendientes:
                sugerido = sugeridos.get(columna.letra)
                if sugerido and sugerido.campo in disponibles:
                    columna.campo = sugerido.campo
                    columna.origen = "ia"
                    columna.confianza = sugerido.confianza.value
                    disponibles.pop(sugerido.campo, None)
            propuesta.avisos.append(
                "la IA propuso {} columna(s) con el modelo {}; confirmelas".format(
                    sum(1 for c in pendientes if c.origen == "ia"), respuesta.modelo
                )
            )
        except Exception as error:  # noqa: BLE001 - sin IA se sigue a mano
            propuesta.avisos.append(
                "no se pudo consultar a la IA ({}): asigne las columnas a mano".format(
                    type(error).__name__
                )
            )

    # Cuantas hileras trae la tabla. Se corta en la fila de totales igual que
    # la importacion: mas abajo suele venir el cuadro resumen, y contar sus
    # especies daria un numero que no cuadra con lo que luego se importa.
    if "especie" in propuesta.mapeo:
        indice = column_index_from_string(propuesta.mapeo["especie"])
        contador, vacias = 0, 0
        for fila in range(fila_encabezado + 1, ws.max_row + 1):
            etiqueta = limpiar_texto(ws.cell(fila, indice).value)
            if not etiqueta:
                vacias += 1
                if vacias > MAXIMO_VACIAS_SEGUIDAS:
                    break
                continue
            if _es_fila_a_ignorar(etiqueta, reglas.patrones_ignorar):
                break
            vacias = 0
            contador += 1
        propuesta.filas_estimadas = contador

    if not propuesta.completo:
        propuesta.avisos.append(
            "faltan por asignar los campos imprescindibles (especie y "
            "superficie): indiquelos en la tabla"
        )
    return propuesta


# ------------------------------------------------------ cuadro resumen 5.2

@dataclass
class ResumenEspecie:
    """Una linea del cuadro resumen, con las superficies ya repartidas."""

    especie: str
    propietario_ha: float = 0.0
    arrendatario_ha: float = 0.0
    plena_produccion_ha: float = 0.0
    rendimientos_crecientes_ha: float = 0.0
    formacion_ha: float = 0.0

    @property
    def total_ha(self) -> float:
        return round(self.propietario_ha + self.arrendatario_ha, 4)


def clasificar_madurez(
    plantaciones: Sequence[Plantacion],
    reglas: ReglasAgronomicas,
    anio_informe: int,
) -> List[ResumenEspecie]:
    """Arma el cuadro resumen por especie a partir del catastro.

    Sustituye a las formulas escritas a mano del tipo ``=D5+D6+D12``, que
    obligan al tasador a decidir fila por fila que plantacion esta en plena
    produccion y a rehacerlas cada vez que se agrega una hilera.

    La antiguedad se cuenta desde el PRIMER anio de un rango: una plantacion
    '2017-2018' se trata como de 2017.
    """
    from schemas.informe import Tenencia

    resumenes: Dict[str, ResumenEspecie] = {}

    for plantacion in plantaciones:
        base = especie_base(plantacion.especie, reglas.separadores_calificativo)
        clave = normalizar(base)
        resumen = resumenes.setdefault(clave, ResumenEspecie(especie=base))

        superficie = plantacion.superficie_ha
        if plantacion.dominio is Tenencia.ARRENDATARIO:
            resumen.arrendatario_ha = round(resumen.arrendatario_ha + superficie, 4)
        else:
            resumen.propietario_ha = round(resumen.propietario_ha + superficie, 4)

        inicio = anio_inicial(plantacion.anio_plantacion)
        if inicio is None:
            continue  # sin anio no se puede clasificar; queda fuera del reparto

        edad = anio_informe - inicio
        formacion_hasta, crecientes_hasta = reglas.umbrales(plantacion.especie)
        if edad <= formacion_hasta:
            resumen.formacion_ha = round(resumen.formacion_ha + superficie, 4)
        elif edad <= crecientes_hasta:
            resumen.rendimientos_crecientes_ha = round(
                resumen.rendimientos_crecientes_ha + superficie, 4
            )
        else:
            resumen.plena_produccion_ha = round(
                resumen.plena_produccion_ha + superficie, 4
            )

    return list(resumenes.values())
