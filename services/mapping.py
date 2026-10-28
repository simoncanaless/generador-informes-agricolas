"""Carga tipada de config/mapping_santander.json.

Aisla al motor del formato del JSON: si manana el mapping cambia de forma, se
ajusta este modulo y el motor no se entera (DIP). El motor NUNCA lee el JSON
crudo ni conoce una coordenada de celda literal.

Resuelve ademas las referencias cruzadas del tipo
    "columnas_por_temporada": "ver flujo_assumptions.columnas_por_temporada"
para que el JSON no tenga que duplicar bloques.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from pydantic import (BaseModel, ConfigDict, Field, field_validator,
                      model_validator)

#: Prefijo que marca una referencia cruzada dentro del propio mapping.
PREFIJO_REFERENCIA = "ver "


class _Spec(BaseModel):
    """Base de los specs del mapping.

    ``extra='allow'`` es deliberado: el JSON lleva metadatos de documentacion
    (``nota``, ``confianza``, ``area_visual``...) que el motor ignora pero que
    no deben romper la carga ni perderse al releer el archivo (OCP).

    OJO: nada de ``str_strip_whitespace``. Dos pestanas del formato del banco se
    llaman "Flujo proyectado " y "Anexo ", con espacio final significativo;
    recortarlo deja el nombre sin correspondencia en el libro.
    """

    model_config = ConfigDict(extra="allow")


class EspecCampo(_Spec):
    """Un campo simple: una ruta de atributo -> una celda."""

    hoja: str
    celda: str
    tipo: str = "str"
    formato: Optional[str] = None
    area_visual: Optional[str] = Field(
        default=None,
        description="rango que ocupa visualmente el campo en la plantilla. "
                    "Para los bloques de texto largo (comentarios, opinion) el "
                    "motor lo combina y activa el ajuste de linea: sin eso el "
                    "texto se sale de su celda o queda recortado.",
    )
    clave_temporada: Optional[str] = Field(
        default=None,
        description="si el atributo es una serie por temporada, la clave a extraer",
    )
    sirve_a_clases: List[str] = Field(
        default_factory=list,
        description="que clases de temporada usan ESTA celda como referencia. "
                    "La plantilla trae dos celdas de tipo de cambio -una para "
                    "lo historico y otra para la proyeccion-; cual le toca a "
                    "cada temporada es una propiedad del formato del banco, no "
                    "una regla del motor, asi que se declara aqui y el token "
                    "{tc} de las formulas la resuelve.",
    )


class EspecFormulaDeSerie(_Spec):
    """Una fila de una tabla ``filas_fijas`` que es DERIVADA, no capturada.

    Mismo principio que ``formulas_de_fila`` en los bloques verticales: en vez
    de escribir un numero se escribe una formula viva, y se declara con nombres
    de campo en vez de coordenadas. Tokens disponibles:

    - ``<campo>``  -> la celda de esa otra fila de la tabla, MISMA columna.
    - ``{col}``    -> la letra de columna de la temporada.
    - ``{valor}``  -> lo que el tasador declaro para esa temporada. Sirve para
      dejar a la vista el dato de origen dentro de la formula: las ventas
      respaldadas se declaran en pesos y la fila del banco pide dolares.
    - ``{tc}``     -> la celda del tipo de cambio que le toca a esa temporada,
      resuelta desde los ``campos`` que llevan ``clave_temporada``. Es una
      referencia viva: cambiar el tipo de cambio mueve la fila.

    ``temporadas_de`` nombra la fila cuyas temporadas mandan. Una fila que solo
    es formula no tiene datos propios que digan en que columnas va; el
    porcentaje de respaldo existe exactamente donde existe el monto respaldado.
    """

    formula: str
    temporadas_de: Optional[str] = None


class EspecTabla(_Spec):
    """Una tabla. La forma util depende de ``orientacion``.

    - ``filas`` (por defecto): lista vertical con ``primera_fila`` + ``columnas``.
    - ``filas_fijas``: filas nombradas (``filas``) x ``columnas`` o
      ``columnas_por_temporada``.
    - ``bloques_de_columnas`` / ``dos_bloques_de_columnas``: N bloques laterales.
    - ``bloques_de_filas``: N bloques verticales con ``offsets_de_fila``.

    ``formulas_de_fila`` marca, dentro de un bloque vertical, las filas que son
    derivadas: en vez de escribir el valor que traiga el modelo se escribe una
    formula de Excel viva. La plantilla usa ``<campo>`` para referirse a otra
    fila del mismo bloque, y el motor lo resuelve a la celda concreta::

        "produccion_total_kg": "=<rendimiento_kg_ha>*<superficie_ha>"

    ``formulas_de_columna`` es lo mismo para una tabla vertical corriente: la
    columna declarada se escribe como formula y ``<campo>`` apunta a otra
    columna de LA MISMA FILA::

        "requerimiento_total_plantacion_ls": "=<requerimiento_max_ls_ha>*<superficie_ha>"

    En los dos casos se declara con nombres de campo y no con coordenadas, para
    que mover una fila o una columna sea cambiar el mapping y nada mas.

    ``formulas_de_serie`` es lo mismo para una tabla ``filas_fijas`` con
    ``columnas_por_temporada``: la fila declarada se escribe como formula en
    la columna de cada temporada. Ver ``EspecFormulaDeSerie``.

    ``encabezados`` escribe rotulos fijos en celdas concretas. Sirve cuando el
    generador agrega una columna que la plantilla del banco no trae y hay que
    titularla para que se entienda en el informe impreso.

    ``celdas_calculadas`` escribe una formula fija en una celda concreta, para
    los totales y referencias que la plantilla dejo vacios::

        "K33": "=$H$27"

    El motor no pisa una formula que ya exista, asi que declarar de mas es
    inofensivo: si el banco ya la trae, se respeta la suya.
    """

    hoja: str
    atributo: Optional[str] = None
    orientacion: str = "filas"

    primera_fila: Optional[int] = None
    ultima_fila_plantilla: Optional[int] = None
    fila_modelo: Optional[int] = None
    capacidad: Optional[int] = None
    fila_totales: Optional[int] = None
    columnas_agregadas: List[str] = Field(
        default_factory=list,
        description="columnas que el generador agrega y la plantilla del banco "
                    "no trae (la superficie del Anexo, la cobertura de riego). "
                    "Sin esto quedan FUERA del recuadro, flotando a la derecha "
                    "de la tabla: se les copia el estilo de la columna vecina "
                    "para que se vean parte de ella.",
    )
    columnas_sumables: Optional[List[str]] = Field(
        default=None,
        description="que columnas admiten total. Sin declarar, el motor suma "
                    "toda columna numerica cuya celda de total este vacia. Eso "
                    "esta bien para superficies y montos, pero NO para tasas: "
                    "sumar las eficiencias de riego de cuatro plantaciones da "
                    "'3,6', un numero que no significa nada y que el banco ve "
                    "en el informe. Una lista vacia desactiva el automatismo.",
    )
    insertar_filas: bool = False

    columnas: Dict[str, str] = Field(default_factory=dict)
    filas: Dict[str, int] = Field(default_factory=dict)
    bloques: List[Dict[str, Any]] = Field(default_factory=list)
    filas_clase_uso: Dict[str, int] = Field(default_factory=dict)
    columnas_por_temporada: Union[Dict[str, str], str, None] = None
    offsets_de_fila: Dict[str, int] = Field(default_factory=dict)
    formulas: List[str] = Field(default_factory=list)
    formulas_de_fila: Dict[str, str] = Field(default_factory=dict)
    formulas_de_serie: Dict[str, EspecFormulaDeSerie] = Field(
        default_factory=dict,
        description="filas derivadas de una tabla de series por temporada. "
                    "Ver EspecFormulaDeSerie.",
    )
    filas_espejo: Dict[str, Any] = Field(
        default_factory=dict,
        description="filas de OTRA parte de la hoja que resumen bloque por "
                    "bloque y que la plantilla no rotula. El 'RESUMEN DE "
                    "INDICADORES GENERALES' repite cuatro filas por indicador "
                    "-una por especie- todas sin nombre: el informe mostraba "
                    "'69,5' y '765.000 kg' sin decir de que. Lleva "
                    "'grupos' (la primera fila de cada indicador; el bloque i "
                    "va en grupo+i), 'columna_etiqueta' y 'campo_etiqueta' "
                    "-el campo del bloque de donde sale el rotulo, para que "
                    "diga lo mismo que su encabezado y no pueda desfasarse-.",
    )
    columnas_acomodadas: List[str] = Field(
        default_factory=list,
        description="columnas cuyo contenido se acomoda a su recuadro: ajuste "
                    "de linea y centrado vertical. Dos motivos, los dos "
                    "vistos en la hoja de Inversion. Un texto mas ancho que "
                    "su columna y CENTRADO se recorta por los dos lados: "
                    "'Plantacion nogal El Almendro' salia como 'ion nogal El "
                    "Al'. Y en una columna combinada a lo alto -el proyecto, "
                    "su anio, su superficie- el valor cae al fondo del "
                    "recuadro, cada uno a distinta altura que sus vecinos.",
    )
    campos_de_grupo: List[str] = Field(
        default_factory=list,
        description="en una tabla 'filas' cuya plantilla combina varias filas "
                    "por grupo -Inversion: 'A5:A9' es UN proyecto que abarca "
                    "sus cinco lineas de item-, los campos que identifican ese "
                    "grupo (la TUPLA de sus valores, no cada uno por separado: "
                    "asi dos proyectos con el mismo nombre pero anio distinto "
                    "no se confunden). Sin esto el motor escribe los items "
                    "seguidos sin saber donde termina un grupo y empieza el "
                    "siguiente, y un segundo proyecto que no arranca justo en "
                    "el limite del bloque cae DENTRO del bloque combinado del "
                    "primero: el motor se niega a pisarlo -good- pero ese "
                    "proyecto queda sin su nombre, anio o superficie escritos.",
    )
    tamano_bloque: Optional[int] = Field(
        default=None,
        description="cuantas filas fisicas reserva la plantilla para CADA "
                    "grupo de 'campos_de_grupo' (Inversion: 5). Los items de "
                    "un grupo con menos filas que esto se completan con filas "
                    "en blanco -mismos campos de grupo, el resto vacio- para "
                    "que el proximo grupo arranque justo en el limite "
                    "siguiente. Si la plantilla no tiene sitio para todos los "
                    "grupos, se avisa cuales quedaron fuera; nunca se trunca "
                    "en silencio.",
    )
    fila_modelo_de_estilo_totales: Optional[int] = Field(
        default=None,
        description="fila de la que se copian bordes y relleno a la fila de "
                    "totales. La plantilla de Inversion deja esa fila SIN "
                    "ningun borde, asi que el total -que es el numero mas "
                    "grande de la pagina- queda flotando fuera del recuadro. "
                    "Se copia solo el recuadro, no la alineacion: la de las "
                    "celdas de datos esta centrada y centraria los numeros.",
    )
    columna_modelo_de_estilo: Optional[str] = Field(
        default=None,
        description="columna cuyo aspecto -bordes, relleno, fuente, "
                    "alineacion- se copia a las columnas que el generador "
                    "llena. El bloque de respaldo de ventas trae estilada UNA "
                    "sola columna, porque el banco esperaba una sola temporada "
                    "respaldada; sin esto las columnas con datos salen fuera "
                    "del recuadro, como numeros sueltos al costado.",
    )
    filas_de_cabecera: List[int] = Field(
        default_factory=list,
        description="filas de rotulo que hay que vestir junto con las de datos "
                    "(la de las temporadas). No estan en 'filas' porque no son "
                    "un atributo del modelo.",
    )
    formulas_de_columna: Dict[str, str] = Field(default_factory=dict)
    encabezados: Dict[str, str] = Field(default_factory=dict)
    celdas_calculadas: Dict[str, str] = Field(default_factory=dict)

    @field_validator("formulas_de_serie", mode="before")
    @classmethod
    def _formula_suelta(cls, v):
        """Admite la forma corta: el string es la formula y nada mas.

        La mayoria de las filas derivadas se declaran solas; obligar a escribir
        ``{"formula": "..."}`` en todas para que una sola pueda decir
        ``temporadas_de`` seria ruido en el mapping.
        """
        if isinstance(v, dict):
            return {k: ({"formula": x} if isinstance(x, str) else x)
                    for k, x in v.items()}
        return v

    @property
    def filas_disponibles(self) -> int:
        """Filas que la plantilla trae hechas para esta tabla."""
        if self.primera_fila is None or self.ultima_fila_plantilla is None:
            return 0
        return self.ultima_fila_plantilla - self.primera_fila + 1

    @property
    def fila_modelo_efectiva(self) -> int:
        """Fila cuyo estilo y formulas se clonan al insertar filas nuevas.

        Por defecto la segunda fila del cuerpo: la primera suele tener el borde
        superior distinto y, en el Anexo, literales 0 en vez de formulas. La
        ultima fila lleva el borde de cierre de la tabla y tampoco sirve.
        """
        if self.fila_modelo is not None:
            return self.fila_modelo
        if self.primera_fila is None:
            raise ValueError("la tabla no define primera_fila")
        if self.filas_disponibles >= 2:
            return self.primera_fila + 1
        return self.primera_fila

    @property
    def fila_insercion(self) -> int:
        """Fila donde se inyectan las filas extra.

        Se inserta EN ``ultima_fila_plantilla``, no despues: asi las formulas de
        totales tipo ``=SUM(D7:D10)`` quedan dentro del rango desplazado y se
        expanden solas a ``=SUM(D7:D41)``, igual que haria Excel.
        """
        if self.ultima_fila_plantilla is None:
            raise ValueError("la tabla no define ultima_fila_plantilla")
        return self.ultima_fila_plantilla


class EspecFoto(_Spec):
    """Un hueco de imagen declarado por la plantilla."""

    hoja: str
    celda_ancla: str
    rango_sugerido: Optional[str] = None
    titulo_en: Optional[str] = None
    titulo_texto: Optional[str] = None
    capacidad: int = 1
    #: Quien gana cuando dos categorias apuntan al MISMO hueco. Menor gana.
    #: La satelital y el emplazamiento comparten el unico marcador de imagen
    #: que trae la Bsan, y sin esto ganaria la que llegue primero en la lista:
    #: una foto de terreno que el modelo rotulo «emplazamiento» le robaria el
    #: hueco grande a la satelital que el tasador subio a proposito.
    prioridad: int = 50


class EspecRejilla(_Spec):
    """Rejilla de la galeria de fotografias (hoja que no existe en la plantilla)."""

    columnas_ancla: List[str] = Field(default_factory=list)
    filas_ancla: List[int] = Field(default_factory=list)
    ancho_columnas: Optional[float] = None
    alto_filas: Optional[float] = None
    capacidad_por_hoja: Optional[int] = None


class EspecGaleria(_Spec):
    estado: Optional[str] = None
    requiere_hoja_nueva: bool = False
    nombre_hoja_propuesto: str = "Fotografias"
    rejilla: Optional[EspecRejilla] = None
    celdas: List[str] = Field(default_factory=list)


class EspecFotografias(_Spec):
    soportado_por_plantilla: Dict[str, EspecFoto] = Field(default_factory=dict)
    galeria_inspeccion: Optional[EspecGaleria] = None


class EspecFormatos(_Spec):
    """Formato numerico de las celdas, por rango.

    La plantilla del banco viene entera en formato 'General': los pesos salen
    como '1107021,94' y las hectareas como '67,69999999999999'. Los datos son
    correctos y el informe se ve mal igual, que para un entregable es lo mismo
    que estar mal.

    Se declara por RANGO y no por campo porque la mayoria de los numeros del
    informe no los escribe el generador: los calculan las 573 formulas que ya
    trae la plantilla, y esas tambien hay que formatearlas.

    ``catalogo`` da nombre a cada formato para no repetir la mascara de Excel
    en cincuenta sitios y para poder cambiarla en uno solo.
    """

    catalogo: Dict[str, str] = Field(default_factory=dict)
    rangos: Dict[str, Dict[str, str]] = Field(default_factory=dict)
    anchos_minimos: Dict[str, Any] = Field(default_factory=dict)

    @field_validator("rangos", mode="before")
    @classmethod
    def _sin_documentacion(cls, v):
        """Descarta las claves de documentacion de los mapas de rangos.

        Todo el mapping lleva notas al lado de lo que declara, con la
        convencion de un guion bajo delante. En un mapa de rangos esa nota se
        leia como un rango mas y el motor avisaba de un formato inexistente en
        cada informe.
        """
        if isinstance(v, dict):
            return {alias: ({k: x for k, x in rangos.items()
                             if not k.startswith("_")}
                            if isinstance(rangos, dict) else rangos)
                    for alias, rangos in v.items()}
        return v

    def mascara(self, nombre: str) -> Optional[str]:
        return self.catalogo.get(nombre)

    def ancho(self, nombre: str) -> Optional[float]:
        """Ancho minimo de columna para que el formato no salga en '####'.

        Formatear ALARGA el texto: donde decia '1107021,94' pasa a decir
        'USD 1.107.022'. Sin ensanchar, el informe se ve peor formateado que
        sin formatear, y encima parece un error de calculo.
        """
        valor = self.anchos_minimos.get(nombre)
        return float(valor) if isinstance(valor, (int, float)) else None


class EspecMatriz(_Spec):
    """Hoja anexa con la matriz edad x temporada, que la plantilla no trae.

    Mismo criterio que ``EspecGaleria``: la hoja no existe en el formato del
    banco, asi que su geometria se declara aqui y el motor la crea. Las
    plantillas de ``formulas`` llevan ``{superficie}``, ``{rendimiento}`` y
    ``{costo}``, que el motor sustituye por los rangos reales una vez que sabe
    en que filas quedo cada bloque.
    """

    estado: Optional[str] = None
    requiere_hoja_nueva: bool = True
    nombre_hoja_propuesto: str = "Detalle plantaciones"
    primera_fila: int = 4
    filas_entre_bloques: int = 3
    columnas: Dict[str, str] = Field(default_factory=dict)
    primera_columna_temporada: str = "G"
    formulas: Dict[str, str] = Field(default_factory=dict)


class EspecCorreccion(_Spec):
    """Una formula de la plantilla que hay que reemplazar, y por que.

    La plantilla del banco trae errores propios. El caso que motivo esto:
    ``Flujo!D112:N112`` calcula ``=(D110/D106)*100`` sin guardia, asi que toda
    temporada sin ingresos muestra ``#DIV/0!`` -y las primeras columnas de un
    informe recien armado no tienen ingresos casi nunca-. Eso viola la regla de
    integridad del proyecto ("cero celdas con #DIV/0!") y, sobre todo, se ve
    como un error del informe cuando es un error de la plantilla.

    Se declara aqui y no en el codigo por el mismo motivo que todo lo demas:
    son coordenadas de celda. Y se declara una por una, con motivo escrito, en
    vez de dar al motor un permiso general para pisar formulas: lo que hay que
    poder hacer es corregir ESTAS celdas, no cualquiera.

    En ``formula``, ``{col}`` se sustituye por la columna de cada celda del
    rango, que es lo que permite escribir la correccion una sola vez para las
    once temporadas.
    """

    hoja: str
    rango: str
    formula: str
    motivo: str


class EspecFormulaFinal(_Spec):
    """Una formula que solo se puede escribir cuando ya se escribio todo.

    Existe por las celdas que cruzan de una hoja a otra. El resumen de la
    caratula lee los totales del Anexo, y la fila de totales del Anexo **se
    mueve** si el cliente trae mas especies que huecos trae la plantilla. Al
    escribirse durante su propia tabla, la caratula todavia no podia saber
    donde iba a terminar esa fila.

    Por eso hay un token: ``{total:anexo_real}`` se resuelve a la fila de
    totales de esa tabla YA COMPENSADA por las filas insertadas, y admite
    desplazamiento -``{total:anexo_real+1}``- para las filas que cuelgan de
    ella, como administracion y gastos generales.

    Se escriben con ``escribir``, asi que no pisan nada: si la plantilla ya
    trae algo en esa celda, se respeta y se avisa.
    """

    hoja: str
    celda: Optional[str] = None
    rango: Optional[str] = Field(
        default=None,
        description="alternativa a 'celda': escribe la MISMA formula en todas "
                    "las celdas del rango, con '{col}' resuelto a la columna "
                    "de cada una. Sirve para las filas que se repiten temporada "
                    "a temporada.",
    )
    formula: str
    solo_si_vacia: bool = Field(
        default=False,
        description="no escribir si la celda ya trae algo. Es lo correcto para "
                    "las formulas que son un SUPUESTO de respaldo: si el "
                    "tasador declaro el dato temporada a temporada, lo suyo "
                    "manda sobre lo que el sistema deduce.",
    )
    nota: Optional[str] = None


class EspecLiteral(_Spec):
    """Un texto fijo que la plantilla no trae y el informe necesita.

    Hasta ahora todo texto salia de un campo del informe o de los encabezados
    de una tabla. Falta un tercer caso: la ETIQUETA de un dato que el formato
    del banco no contempla. El valor de la UF, por ejemplo, hay que declararlo
    en alguna celda, y una celda con un numero suelto y sin rotulo al lado no
    se entiende al abrir el archivo.

    Se escribe con ``escribir``, asi que nunca pisa lo que la plantilla ya
    traiga.
    """

    hoja: str
    celda: str
    texto: str
    nota: Optional[str] = None


class EspecCabeceraTemporadas(_Spec):
    """Las filas donde la hoja repite los nombres de las temporadas.

    La plantilla las trae escritas: '2023/24', '2024/25'... Eso solo sirve para
    el ano en que se escribieron; un informe hecho en 2027 saldria rotulado con
    las temporadas de 2026. Es el patron que §13 marca como a evitar.

    El generador las escribe segun lo que el informe realmente tiene, y la fila
    de abajo dice de que tipo es cada una ('Real ', 'Preliminar',
    'Projection '). Una temporada sin datos se deja EN BLANCO: es la forma de
    que la columna "no exista" cuando el cliente no entrego ese ano.
    """

    hoja: str
    filas_temporada: List[int] = Field(default_factory=list)
    filas_clase: List[int] = Field(default_factory=list)
    filas_solo_declaradas: List[int] = Field(
        default_factory=list,
        description="filas que se rotulan SOLO en las temporadas que el cliente "
                    "declaro, no en las proyectadas. Es la cabecera del nivel "
                    "de respaldo de ventas: una proyeccion no tiene IVAs que la "
                    "respalden, asi que rotular esas columnas mostraria un anio "
                    "con la celda vacia debajo, que se lee como un dato que "
                    "falta y no como uno que no puede existir.",
    )
    clases_declaradas: List[str] = Field(
        default_factory=list,
        description="que clases de temporada cuentan como declaradas por el "
                    "cliente ('Real ', 'Preliminar'), para filas_solo_declaradas.",
    )
    filas_a_ocultar: List[int] = Field(
        default_factory=list,
        description="rango [desde, hasta] de filas cuyos CEROS se ocultan en "
                    "las columnas sin temporada. La plantilla calcula igual "
                    "-son sus formulas sobre celdas vacias- y muestra '0,0' y "
                    "'0 kg' en una columna que no existe, lo que se lee como un "
                    "dato en cero. No se borran las formulas: si el tasador "
                    "escribe ahi despues, el numero aparece solo.",
    )
    hojas_extra: List[Dict[str, Any]] = Field(
        default_factory=list,
        description="otras filas que repiten las temporadas, en otra hoja y "
                    "empezando en otra columna. La caratula lista los 8 anios "
                    "del horizonte, que arrancan en el ano 1 y no en el primer "
                    "historico: por eso lleva 'desde_temporada'.",
    )


class EspecUnidadExcluible(_Spec):
    """Una hoja o tabla que el tasador puede sacar del .xlsx antes de bajarlo.

    ``nivel`` dice como se saca:

    - ``"hoja"``: se borra la pestana entera (``services.excel_ops.
      eliminar_hoja``). No lleva ``filas`` ni ``columnas``.
    - ``"tabla"``: la tabla es dueña de un rango de filas EXCLUSIVO en su
      hoja -nada mas vive ahi-, asi que se borran de verdad
      (``eliminar_filas``). ``filas`` es ``[primera, ultima]``.
    - ``"tabla_columnas"``: la tabla COMPARTE filas con otra tabla de la
      misma hoja -Maquinaria y Infraestructura ocupan las mismas filas
      42-45 en columnas distintas de 'Dotacion de recursos'-, asi que
      borrar la fila destruiria a la vecina. Solo se vacia el contenido
      (``vaciar_columnas``): la tabla queda vacia, no desaparece el hueco.
      ``filas`` y ``columnas`` (``[desde, hasta]``) van los dos.

    ``depende_de`` son los ids de OTRAS unidades cuyas celdas leen las
    formulas de ESTA: si esas se excluyen, esta se queda con una formula sin
    de donde leer, asi que tiene que excluirse tambien -es la cascada que
    resuelve ``services/exclusiones.py``, nunca a mano en el codigo-.

    El grafo completo (quien depende de quien) se verifico escaneando las
    573 formulas reales de la plantilla mas las que el propio generador
    inyecta -no se declaro a ojo-. Ver ``scripts/graficar_dependencias.py``,
    que lo vuelve a escanear cada vez que corre para que esta lista no se
    desalinee en silencio si un dia se agrega una formula nueva que cruza
    de tabla.
    """

    etiqueta: str
    nivel: str
    hoja: str
    filas: Optional[List[int]] = None
    columnas: Optional[List[str]] = None
    depende_de: List[str] = Field(default_factory=list)


class MapaSantander(_Spec):
    """El mapping completo, ya validado y con las referencias resueltas."""

    version: str
    plantilla: str
    hojas: Dict[str, str]
    campos: Dict[str, EspecCampo] = Field(default_factory=dict)
    tablas: Dict[str, EspecTabla] = Field(default_factory=dict)
    fotografias: EspecFotografias = Field(default_factory=EspecFotografias)
    matriz_envejecimiento: Optional[EspecMatriz] = None
    formatos: EspecFormatos = Field(default_factory=EspecFormatos)
    correcciones_de_formula: List[EspecCorreccion] = Field(default_factory=list)
    formulas_finales: List[EspecFormulaFinal] = Field(default_factory=list)
    literales: List[EspecLiteral] = Field(default_factory=list)
    cabecera_temporadas: Optional[EspecCabeceraTemporadas] = None
    unidades_excluibles: Dict[str, EspecUnidadExcluible] = Field(default_factory=dict)

    @model_validator(mode="after")
    def _resolver_referencias(self) -> "MapaSantander":
        """Sustituye los ``"ver <tabla>.<clave>"`` por su valor real."""
        for nombre, tabla in self.tablas.items():
            valor = tabla.columnas_por_temporada
            if isinstance(valor, str) and valor.startswith(PREFIJO_REFERENCIA):
                ruta = valor[len(PREFIJO_REFERENCIA):].strip()
                tabla.columnas_por_temporada = self._resolver_ruta(ruta, nombre)
        return self

    def _resolver_ruta(self, ruta: str, origen: str) -> Dict[str, str]:
        destino, _, clave = ruta.partition(".")
        if destino not in self.tablas:
            raise ValueError(
                "tabla '{}': referencia a tabla inexistente '{}'".format(origen, destino)
            )
        valor = getattr(self.tablas[destino], clave, None)
        if not isinstance(valor, dict):
            raise ValueError(
                "tabla '{}': la referencia '{}' no apunta a un mapa".format(origen, ruta)
            )
        return valor

    def nombre_hoja(self, alias: str) -> str:
        """Traduce el alias del mapping al nombre real de la pestana."""
        try:
            return self.hojas[alias]
        except KeyError:
            raise KeyError("alias de hoja desconocido: {!r}".format(alias)) from None


def cargar_mapa(ruta: Union[str, Path]) -> MapaSantander:
    """Lee y valida el mapping desde disco."""
    ruta = Path(ruta)
    with ruta.open(encoding="utf-8") as fh:
        crudo = json.load(fh)
    return MapaSantander.model_validate(crudo)
