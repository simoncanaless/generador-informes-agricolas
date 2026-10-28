"""Motor de escritura Excel: vuelca un InformeFitogreen sobre la plantilla.

Determinista de principio a fin. No hay IA, no hay heuristicas y no hay una
sola coordenada de celda escrita en este archivo: todas salen de
``config/mapping_santander.json`` a traves de ``services.mapping``.

Orden de trabajo, y el porque:

  1. Campos simples. Se escriben primero porque ``insertar_filas`` mueve los
     valores ya escritos junto con sus celdas, asi que un campo situado debajo
     de una tabla que crece acaba solo en su fila correcta.
  2. Tablas, hoja por hoja y de ABAJO HACIA ARRIBA. Procesar la tabla mas baja
     primero significa que, cuando una tabla superior inserta filas, todo lo de
     abajo ya esta escrito y simplemente baja con sus estilos, sus merges y sus
     formulas reescritas. Asi ninguna coordenada del mapping caduca a mitad de
     la generacion y no hace falta llevar contabilidad de desplazamientos.
  3. Fotografias. Van al final, con las coordenadas corregidas por las filas
     insertadas en su hoja (unico sitio donde si hace falta esa correccion).
"""

from __future__ import annotations

import asyncio
import datetime as dt
import math
import re
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.numbers import is_date_format
from openpyxl.drawing.image import Image as ImagenExcel
from openpyxl.utils import (
    column_index_from_string, get_column_letter, range_boundaries,
)
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from schemas.informe import CategoriaFoto, Fotografia, InformeFitogreen
from services import excel_ops
from services.exclusiones import resolver_cascada
from services.mapping import EspecTabla, MapaSantander, cargar_mapa

__all__ = ["MotorExcel", "ResultadoGeneracion", "cargar_mapa"]


# ------------------------------------------------------------- resultado

@dataclass
class ResultadoGeneracion:
    """Que se genero y con que reservas."""

    ruta: Path
    celdas_escritas: int = 0
    filas_insertadas: int = 0
    imagenes_insertadas: int = 0
    correcciones: List[str] = field(default_factory=list)
    avisos: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.ruta.exists()

    def resumen(self) -> str:
        return (
            "{} | {} celdas, {} filas insertadas, {} imagenes, {} avisos".format(
                self.ruta.name,
                self.celdas_escritas,
                self.filas_insertadas,
                self.imagenes_insertadas,
                len(self.avisos),
            )
        )


# --------------------------------------------------------------- contexto

class _Contexto:
    """Estado mutable de una generacion. Lo comparten las estrategias."""

    def __init__(self, wb: Workbook, mapa: MapaSantander, informe: InformeFitogreen):
        self.wb = wb
        self.mapa = mapa
        self.informe = informe
        self.avisos: List[str] = []
        self.correcciones: List[str] = []
        self.celdas = 0
        self.filas_insertadas = 0
        self.imagenes = 0
        # inserciones aplicadas por nombre real de hoja: [(fila, cantidad), ...]
        self.inserciones: Dict[str, List[Tuple[int, int]]] = {}
        # rangos de la hoja anexa por indice de bloque, para que el flujo la
        # referencie con SUMPRODUCT. Lo llena _escribir_matriz, que corre antes
        # que las tablas justo por esto.
        self.rangos_matriz: List[Dict[str, Any]] = []
        # (hoja, columna, fila) cuyos ceros hay que dejar en blanco. Se junta
        # aqui y no se aplica en el momento porque el formato numerico se
        # escribe al FINAL: hacerlo antes lo pisa y los ceros vuelven.
        self.ceros_a_ocultar: List[Tuple[str, str, int]] = []
        # nombres REALES de hoja que _excluir_unidades saco del libro a
        # pedido del tasador. Las fases de despues (formatos, impresion) la
        # buscan por su cuenta; sin esto, "no existe" se leia como un error
        # del mapping en vez de una exclusion a proposito.
        self.hojas_excluidas: set = set()

    def hoja(self, alias: str) -> Worksheet:
        return self.wb[self.mapa.nombre_hoja(alias)]

    def avisar(self, mensaje: str) -> None:
        if mensaje not in self.avisos:
            self.avisos.append(mensaje)

    def escribir(
        self,
        ws: Worksheet,
        columna: str,
        fila: int,
        valor: Any,
        formato: Optional[str] = None,
    ) -> None:
        escrito, aviso = excel_ops.escribir(
            ws, "{}{}".format(columna, fila), valor, formato
        )
        if escrito:
            self.celdas += 1
        if aviso:
            self.avisar(aviso)

    def insertar(self, ws: Worksheet, desde: int, cantidad: int, modelo: int) -> None:
        excel_ops.insertar_filas(ws, desde, cantidad, modelo)
        self.filas_insertadas += cantidad
        self.inserciones.setdefault(ws.title, []).append((desde, cantidad))

    def fila_final(self, ws: Worksheet, fila_mapa: int) -> int:
        """Traduce una fila del mapping a su posicion tras las inserciones."""
        desplazamiento = sum(
            cantidad
            for desde, cantidad in self.inserciones.get(ws.title, [])
            if desde <= fila_mapa
        )
        return fila_mapa + desplazamiento


# ------------------------------------------------------ acceso al dominio

def resolver(raiz: Any, ruta: Optional[str]) -> Any:
    """Recorre una ruta 'a.b.c' sobre el informe. Devuelve None si se corta."""
    if not ruta:
        return None
    actual = raiz
    for parte in ruta.split("."):
        if actual is None:
            return None
        actual = getattr(actual, parte, None)
    return actual


def _campos_del_modelo(item: Any) -> Sequence[str]:
    return tuple(getattr(type(item), "model_fields", {}).keys())


def _clave(valor: Any) -> str:
    """Normaliza una clave de diccionario que puede venir como Enum."""
    return valor.value if isinstance(valor, Enum) else str(valor)


def _como_lista(ctx: _Contexto, tabla_nombre: str, tabla: EspecTabla) -> List[Any]:
    valor = resolver(ctx.informe, tabla.atributo)
    if valor is None:
        return []
    if not isinstance(valor, list):
        ctx.avisar(
            "tabla '{}': el atributo '{}' no es una lista".format(
                tabla_nombre, tabla.atributo
            )
        )
        return []
    return valor


#: Referencia a una columna que la plantilla ya maneja y el mapping no mapea:
#: '{I}' es "la columna I de esta misma fila". Hace falta porque una formula
#: derivada suele apoyarse en otra que trae la propia plantilla -en el Anexo,
#: 'Total ventas USD' se calcula sobre 'Total ventas ($)', que es formula del
#: banco- y esa columna no tiene campo en el modelo al que referirse.
_REFERENCIA_COLUMNA = re.compile(r"\{([A-Z]{1,3})\}")


def _formula_por_columnas(
    plantilla: str, columnas: Dict[str, str], fila: int,
    fila_totales: Optional[int] = None,
    ws: Optional[Worksheet] = None,
) -> Tuple[Optional[str], List[str]]:
    """Traduce '=<eto>*<kc>' a '=F33*G33' con el mapa campo -> columna.

    Las referencias son a la misma fila: en una tabla vertical cada renglon se
    calcula con sus propios datos. Ademas de ``<campo>`` se admite ``{COL}``
    para apuntar a una columna por su letra.

    Con ``ws``, cada referencia se resuelve al ANCLA de su celda combinada. En
    la hoja de Inversion la superficie ocupa `C5:C9` -una por proyecto, cinco
    lineas de item-, asi que la segunda linea multiplicaria por `C6`, que en
    Excel vale CERO por estar dentro del combinado. La fila saldria en 0 sin
    ninguna celda en error y sin ningun aviso.
    """
    faltantes: List[str] = []

    def coordenada(columna: str) -> str:
        destino = "{}{}".format(columna, fila)
        return excel_ops.celda_ancla(ws, destino) if ws is not None else destino

    def resolver(coincidencia: "re.Match[str]") -> str:
        campo = coincidencia.group(1)
        columna = columnas.get(campo)
        if columna is None:
            faltantes.append(campo)
            return coincidencia.group(0)
        return coordenada(columna)

    formula = _REFERENCIA_DE_FILA.sub(resolver, plantilla)
    formula = _REFERENCIA_COLUMNA.sub(
        lambda c: coordenada(c.group(1)), formula
    )
    # '{total}' es la fila de totales YA compensada: un porcentaje que divide
    # por el total no puede quedar apuntando a la fila donde estaba antes de
    # que la tabla creciera.
    if fila_totales is not None:
        formula = formula.replace("{total}", str(fila_totales))
    return (None if faltantes else formula), faltantes


def _escribir_item(
    ctx: _Contexto,
    ws: Worksheet,
    tabla_nombre: str,
    item: Any,
    fila: int,
    columnas: Dict[str, str],
    formulas: Optional[Dict[str, str]] = None,
    fila_totales: Optional[int] = None,
) -> None:
    """Vuelca los campos de un objeto en una fila, columna a columna.

    Las columnas declaradas en ``formulas`` no se escriben con el valor del
    modelo sino con una formula viva. Es el mismo criterio que en el flujo
    proyectado: lo derivado se entrega como formula para que el banco pueda
    mover un supuesto y ver el efecto.
    """
    disponibles = _campos_del_modelo(item)
    formulas = formulas or {}
    for campo, columna in columnas.items():
        plantilla = formulas.get(campo)
        if plantilla is not None:
            formula, faltantes = _formula_por_columnas(
                plantilla, columnas, fila, fila_totales, ws)
            if formula is None:
                ctx.avisar(
                    "tabla '{}': la formula de '{}' referencia {}, que no tiene "
                    "columna; no se escribe".format(
                        tabla_nombre, campo, ", ".join(sorted(set(faltantes)))
                    )
                )
                continue
            ctx.escribir(ws, columna, fila, formula)
            continue
        if campo not in disponibles:
            ctx.avisar(
                "tabla '{}': el campo '{}' no existe en {}".format(
                    tabla_nombre, campo, type(item).__name__
                )
            )
            continue
        ctx.escribir(ws, columna, fila, getattr(item, campo))


def _agrupar_en_bloques(
    ctx: _Contexto, tabla_nombre: str, items: List[Any], tabla: "EspecTabla",
) -> List[Any]:
    """Reparte ITEMS en bloques de tamano fijo, uno por grupo.

    Existe para tablas como Inversion, donde la plantilla combina
    verticalmente ``tabla.tamano_bloque`` filas por grupo -el nombre del
    proyecto, su anio y su superficie abarcan varias lineas de item-. Sin
    esto el motor escribia los items seguidos sin saber donde termina un
    grupo y empieza el siguiente: un segundo proyecto que no arrancaba justo
    en el limite del bloque caia DENTRO del bloque combinado del primero. El
    motor se negaba a pisarlo -no corrompia nada- pero ese proyecto quedaba
    sin su nombre, anio o superficie escritos, y solo se enteraba quien leyera
    los avisos.

    El grupo se identifica por la TUPLA de ``tabla.campos_de_grupo`` -para
    Inversion, proyecto+anio+superficie-, no por un campo solo: asi dos
    proyectos con el mismo nombre pero anio distinto no se confunden. Los
    items de un grupo mas corto que el bloque se completan con filas en
    blanco -mismos campos de grupo, el resto vacio, que ``escribir()`` deja
    pasar sin aviso por ser el mismo valor que ya esta en el ancla- para que
    el proximo grupo arranque justo en el limite siguiente.
    """
    bloque = tabla.tamano_bloque
    claves = tabla.campos_de_grupo

    grupos: Dict[tuple, List[Any]] = {}
    orden: List[tuple] = []
    for item in items:
        clave = tuple(getattr(item, campo, None) for campo in claves)
        if clave not in grupos:
            grupos[clave] = []
            orden.append(clave)
        grupos[clave].append(item)

    capacidad = tabla.capacidad or tabla.filas_disponibles
    cuantos_caben = max(1, capacidad // bloque) if capacidad else len(orden)
    if len(orden) > cuantos_caben:
        descartados = ", ".join(str(clave[0]) for clave in orden[cuantos_caben:])
        ctx.avisar(
            "tabla '{}': la plantilla solo tiene sitio para {} grupo(s) de {} "
            "filas cada uno; quedaron fuera: {}. Hay que agregarlos a mano o "
            "ampliar la plantilla.".format(
                tabla_nombre, cuantos_caben, bloque, descartados))
        orden = orden[:cuantos_caben]

    campos_vacios = {campo for campo in tabla.columnas if campo not in claves}
    salida: List[Any] = []
    for clave in orden:
        del_grupo = grupos[clave]
        if len(del_grupo) > bloque:
            ctx.avisar(
                "tabla '{}': el grupo '{}' trae {} filas y el bloque admite "
                "{}; se usan las primeras {}.".format(
                    tabla_nombre, clave[0], len(del_grupo), bloque, bloque))
            del_grupo = del_grupo[:bloque]
        salida.extend(del_grupo)
        if len(del_grupo) < bloque:
            relleno = del_grupo[-1].model_copy(
                update={campo: None for campo in campos_vacios})
            salida.extend([relleno] * (bloque - len(del_grupo)))
    return salida


# ------------------------------------- el mismo lenguaje visual del banco
#
# La plantilla tiene su propia identidad: barra roja para el titulo de cada
# hoja ('COMENTARIOS GENERALES'), barra gris para cada sub-seccion
# ('ESTRUCTURA PRODUCTIVA (TEMPORADA)'), Aptos Narrow en negrita para los
# rotulos, tablas con marco. Las hojas que el motor CREA -Detalle
# plantaciones, Fotografias- no existen en el archivo del banco, asi que no
# heredaban nada de eso: salian en blanco y negro, sin un solo borde, como un
# volcado de datos pegado al final en vez de parte del mismo informe.
#
# Los colores se sacaron DEL ARCHIVO REAL -no inventados-, con openpyxl:
# 'Comentarios generales Opinion'!A2 para el rojo (solido FFFF0000, texto
# blanco), 'Cliente'!A16 para el gris (blanco con tinte -0.35, que es
# exactamente A6A6A6 resuelto). El gris claro (blanco con tinte -0.05, que
# resuelve a F2F2F2) es el mismo que trae el banco de fondo en sus propias
# sub-etiquetas -'Dotación de recursos'!A7, 'Flujo proyectado '!F7,
# 'Comentarios generales Opinion'!A5-: se reusa como "mate" detras de una
# foto (§16.a.sexagies.octies).
_ROJO_BANCO = "FFFF0000"
_GRIS_SECCION = "FFA6A6A6"
_GRIS_MATE = "FFF2F2F2"
_FUENTE_BANCO = "Aptos Narrow"
_BORDE_FINO = Side(style="thin", color="FF000000")
_BORDE_MEDIO = Side(style="medium", color="FF000000")


def _fuente_banco(**kw) -> Font:
    kw.setdefault("size", 11)
    return Font(name=_FUENTE_BANCO, **kw)


def _barra_titulo(ws: Worksheet, texto: str, fila: int,
                  col_ini: int, col_fin: int) -> None:
    """La barra roja de titulo de hoja, igual que la que ya trae el banco."""
    if col_fin > col_ini:
        ws.merge_cells(start_row=fila, start_column=col_ini,
                       end_row=fila, end_column=col_fin)
    celda = ws.cell(fila, col_ini)
    celda.value = texto
    celda.font = _fuente_banco(bold=True, color="FFFFFFFF")
    celda.fill = PatternFill(fill_type="solid", fgColor=_ROJO_BANCO)
    celda.alignment = Alignment(horizontal="center", vertical="center")


def _barra_seccion(ws: Worksheet, texto: str, fila: int,
                   col_ini: int, col_fin: int) -> None:
    """La barra gris de cada sub-seccion, igual que la del banco."""
    if col_fin > col_ini:
        ws.merge_cells(start_row=fila, start_column=col_ini,
                       end_row=fila, end_column=col_fin)
    celda = ws.cell(fila, col_ini)
    celda.value = texto
    celda.font = _fuente_banco(bold=True)
    celda.fill = PatternFill(fill_type="solid", fgColor=_GRIS_SECCION)
    celda.alignment = Alignment(horizontal="center", vertical="center")


def _bordear(ws: Worksheet, fila_ini: int, fila_fin: int,
            col_ini: int, col_fin: int) -> None:
    """Marco medio alrededor y lineas finas adentro, como cualquier tabla de
    datos de la plantilla. Sin esto una tabla nueva se lee como texto suelto,
    no como una tabla."""
    for fila in range(fila_ini, fila_fin + 1):
        for col in range(col_ini, col_fin + 1):
            celda = ws.cell(fila, col)
            arriba = _BORDE_MEDIO if fila == fila_ini else _BORDE_FINO
            abajo = _BORDE_MEDIO if fila == fila_fin else _BORDE_FINO
            izq = _BORDE_MEDIO if col == col_ini else _BORDE_FINO
            der = _BORDE_MEDIO if col == col_fin else _BORDE_FINO
            celda.border = Border(top=arriba, bottom=abajo, left=izq, right=der)


def _enmarcar(ws: Worksheet, fila_ini: int, fila_fin: int,
             col_ini: int, col_fin: int) -> None:
    """Marco simple alrededor de un area -una foto-, sin la rejilla interna
    de `_bordear` (esa es para tablas de datos; una foto no tiene columnas).
    Solo toca el perimetro: las celdas de adentro quedan sin tocar."""
    for fila in range(fila_ini, fila_fin + 1):
        for col in range(col_ini, col_fin + 1):
            if fila not in (fila_ini, fila_fin) and col not in (col_ini, col_fin):
                continue
            arriba = _BORDE_MEDIO if fila == fila_ini else None
            abajo = _BORDE_MEDIO if fila == fila_fin else None
            izq = _BORDE_MEDIO if col == col_ini else None
            der = _BORDE_MEDIO if col == col_fin else None
            ws.cell(fila, col).border = Border(top=arriba, bottom=abajo,
                                               left=izq, right=der)


# ----------------------------------------------------------- estrategias

class EstrategiaTabla:
    """Contrato de una forma de tabla.

    Anadir un formato nuevo (otro banco, otro rubro) es registrar otra
    estrategia; el motor no cambia (OCP).
    """

    def aplicar(self, ctx: _Contexto, nombre: str, tabla: EspecTabla) -> None:
        raise NotImplementedError

    def fila_referencia(self, tabla: EspecTabla) -> int:
        """Fila representativa, solo para ordenar las tablas de la hoja."""
        candidatas: List[int] = []
        if tabla.primera_fila is not None:
            candidatas.append(tabla.primera_fila)
        candidatas.extend(tabla.filas.values())
        candidatas.extend(tabla.filas_clase_uso.values())
        candidatas.extend(
            int(b["primera_fila"]) for b in tabla.bloques if "primera_fila" in b
        )
        return min(candidatas) if candidatas else 0


ESTRATEGIAS: Dict[str, EstrategiaTabla] = {}


def registrar(nombre: str):
    def decorador(cls):
        ESTRATEGIAS[nombre] = cls()
        return cls

    return decorador


@registrar("filas")
class TablaDeFilas(EstrategiaTabla):
    """Lista vertical clasica: una fila por elemento."""

    def aplicar(self, ctx: _Contexto, nombre: str, tabla: EspecTabla) -> None:
        items = _como_lista(ctx, nombre, tabla)
        if not items:
            return
        if tabla.primera_fila is None:
            ctx.avisar("tabla '{}': falta primera_fila".format(nombre))
            return

        ws = ctx.hoja(tabla.hoja)
        capacidad = tabla.capacidad or tabla.filas_disponibles

        # Inversion: la plantilla combina varias filas por proyecto. Sin
        # repartir por grupo antes, un segundo proyecto que no arranca justo
        # en el limite del bloque cae dentro del bloque del primero (ver
        # `_agrupar_en_bloques`).
        if tabla.campos_de_grupo and tabla.tamano_bloque:
            items = _agrupar_en_bloques(ctx, nombre, items, tabla)

        if len(items) > capacidad:
            if not tabla.insertar_filas:
                ctx.avisar(
                    "tabla '{}': {} elementos para {} filas y la plantilla no "
                    "admite insercion; se escriben los {} primeros".format(
                        nombre, len(items), capacidad, capacidad
                    )
                )
                items = items[:capacidad]
            else:
                ctx.insertar(
                    ws,
                    tabla.fila_insercion,
                    len(items) - capacidad,
                    tabla.fila_modelo_efectiva,
                )

        for celda, texto in tabla.encabezados.items():
            columna, fila = _partir_celda(celda)
            ctx.escribir(ws, columna, ctx.fila_final(ws, fila), texto)
        for celda, formula in tabla.celdas_calculadas.items():
            # Estas celdas suelen vivir en la fila de TOTALES, que se corre
            # hacia abajo cuando la tabla desborda. Hay que mover dos cosas: la
            # celda destino y las referencias de dentro de la formula, para lo
            # que se usa '{fila}'. Sin esto, con cinco especies el total en USD
            # se escribia sobre una fila de datos y dividia la fila equivocada.
            columna, fila = _partir_celda(celda)
            destino = ctx.fila_final(ws, fila)
            ctx.escribir(ws, columna, destino,
                         _sustituir_fila(formula, destino))

        for indice, item in enumerate(items):
            _escribir_item(
                ctx, ws, nombre, item, tabla.primera_fila + indice, tabla.columnas,
                tabla.formulas_de_columna,
                ctx.fila_final(ws, tabla.fila_totales)
                if tabla.fila_totales else None,
            )

        _completar_totales(
            ctx, ws, tabla, tabla.primera_fila, tabla.primera_fila + len(items) - 1
        )
        _vestir_columnas_agregadas(
            ctx, ws, tabla, tabla.primera_fila,
            tabla.primera_fila + len(items) - 1
        )
        _acomodar_columnas(
            ws, tabla, tabla.primera_fila,
            tabla.primera_fila + len(items) - 1
        )
        _vestir_fila_de_totales(ctx, ws, tabla)


@registrar("filas_fijas")
class TablaDeFilasFijas(EstrategiaTabla):
    """Filas nombradas: cada atributo tiene su fila reservada en la plantilla.

    Dos variantes segun el mapping:
      - con ``columnas``: cada fila es un objeto (campo -> columna).
      - con ``columnas_por_temporada``: cada fila es una serie temporal
        (temporada -> columna).
    """

    def aplicar(self, ctx: _Contexto, nombre: str, tabla: EspecTabla) -> None:
        contenedor = resolver(ctx.informe, tabla.atributo)
        if contenedor is None:
            return
        ws = ctx.hoja(tabla.hoja)
        # Igual que en los bloques: por POSICION, no por el ano escrito en el
        # mapping. Con las columnas historicas dinamicas, un informe con un
        # solo ano declarado corre toda la proyeccion una columna a la
        # izquierda, y las claves fijas dejarian de calzar en silencio.
        por_temporada = (_columnas_de_las_temporadas(ctx, tabla, nombre)
                         if isinstance(tabla.columnas_por_temporada, dict)
                         else tabla.columnas_por_temporada)

        for atributo, fila in tabla.filas.items():
            # Las filas derivadas se escriben mas abajo como formula. Aqui se
            # saltan a proposito: escribir primero el valor y pisarlo despues
            # contaria la celda dos veces y, si la formula fallara, dejaria un
            # numero en pesos en una fila rotulada en dolares.
            if atributo in tabla.formulas_de_serie:
                continue
            valor = getattr(contenedor, atributo, None)
            if valor is None:
                continue
            if isinstance(por_temporada, dict):
                if not isinstance(valor, dict):
                    ctx.avisar(
                        "tabla '{}': '{}' deberia ser una serie por temporada"
                        .format(nombre, atributo)
                    )
                    continue
                for temporada, dato in valor.items():
                    columna = por_temporada.get(_clave(temporada))
                    if columna is None:
                        ctx.avisar(
                            "tabla '{}': la temporada '{}' no tiene columna en la "
                            "plantilla".format(nombre, _clave(temporada))
                        )
                        continue
                    ctx.escribir(ws, columna, fila, dato)
            else:
                _escribir_item(ctx, ws, nombre, valor, fila, tabla.columnas)

        if isinstance(por_temporada, dict):
            self._escribir_formulas_de_serie(
                ctx, ws, nombre, tabla, por_temporada, contenedor)

    def _escribir_formulas_de_serie(
        self,
        ctx: _Contexto,
        ws: Worksheet,
        nombre: str,
        tabla: EspecTabla,
        por_temporada: Dict[str, str],
        contenedor: Any,
    ) -> None:
        """Las filas derivadas de la tabla, como formula viva.

        Ver ``EspecFormulaDeSerie``. A diferencia de un bloque de especie, aqui
        la fila NO se llena entera: el nivel de respaldo de ventas solo existe
        donde el cliente entrego IVAs, y una proyeccion no tiene con que estar
        respaldada. Escribir un cero ahi diria que no hay respaldo, que es una
        afirmacion distinta de que no puede haberlo.
        """
        vestidas = set()
        for atributo, espec in tabla.formulas_de_serie.items():
            fila = tabla.filas.get(atributo)
            if fila is None:
                ctx.avisar(
                    "tabla '{}': la formula de '{}' no tiene fila declarada en "
                    "'filas'".format(nombre, atributo))
                continue

            # Que temporadas llevan esta formula. Una fila que solo es formula
            # no tiene datos propios que lo digan, asi que se cuelga de la fila
            # que los tiene.
            fuente = getattr(contenedor, espec.temporadas_de or atributo, None)
            if not isinstance(fuente, dict):
                continue

            clases = getattr(
                getattr(ctx.informe, "flujo_proyectado", None),
                "clase_de_temporada", None) or {}

            for temporada, dato in fuente.items():
                columna = por_temporada.get(_clave(temporada))
                if columna is None:
                    # Una clave que no es una temporada es la clave simbolica
                    # con que se captura ('real', 'preliminar') y no significa
                    # que falte una columna. Una que SI es temporada, si.
                    if temporada in clases and espec.temporadas_de is None:
                        ctx.avisar(
                            "tabla '{}': hay respaldo declarado para {} pero esa "
                            "temporada no tiene columna en la plantilla".format(
                                nombre, temporada))
                    continue
                if dato is None:
                    continue
                formula, aviso = self._resolver_formula_de_serie(
                    ctx, tabla, espec.formula, columna, temporada, dato)
                if aviso:
                    ctx.avisar("tabla '{}', fila '{}': {}".format(
                        nombre, atributo, aviso))
                    continue
                ctx.escribir(ws, columna, fila, formula)
                vestidas.add(columna)

        modelo = tabla.columna_modelo_de_estilo
        if modelo:
            filas = ([f for a, f in tabla.filas.items()
                      if a in tabla.formulas_de_serie]
                     + list(tabla.filas_de_cabecera))
            for columna in sorted(vestidas):
                _vestir_como_la_columna_modelo(ws, modelo, columna, filas)

    def _resolver_formula_de_serie(
        self,
        ctx: _Contexto,
        tabla: EspecTabla,
        plantilla: str,
        columna: str,
        temporada: str,
        dato: Any,
    ) -> "tuple[Optional[str], Optional[str]]":
        """Sustituye ``<campo>``, ``{col}``, ``{valor}`` y ``{tc}``."""
        faltantes: List[str] = []

        def resolver(coincidencia: "re.Match[str]") -> str:
            campo = coincidencia.group(1)
            destino = tabla.filas.get(campo)
            if destino is None:
                faltantes.append(campo)
                return coincidencia.group(0)
            return "{}{}".format(columna, destino)

        formula = _REFERENCIA_DE_FILA.sub(resolver, plantilla)
        if faltantes:
            return None, ("referencia {}, que no esta en 'filas'".format(
                ", ".join(sorted(set(faltantes)))))

        if "{tc}" in formula:
            celda = _celda_de_tipo_cambio(ctx, temporada)
            if celda is None:
                return None, ("la temporada '{}' no tiene celda de tipo de "
                              "cambio declarada".format(temporada))
            formula = formula.replace("{tc}", celda)

        formula = formula.replace("{valor}", _literal_de_excel(dato))
        return _sustituir_columna(formula, columna), None


def _literal_de_excel(valor: Any) -> str:
    """Un numero como lo escribiria Excel dentro de una formula.

    Con punto decimal y sin separador de miles: en la formula no manda la
    configuracion regional del que la lee.
    """
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return str(valor)
    if float(valor).is_integer():
        return str(int(valor))
    return repr(float(valor))


def _celda_de_tipo_cambio(ctx: _Contexto, temporada: str) -> Optional[str]:
    """La celda del tipo de cambio que le toca a una temporada.

    Sale del propio mapping -de los ``campos`` que declaran ``sirve_a_clases``-
    y no de una regla escrita aqui: la plantilla del banco trae DOS celdas de
    tipo de cambio, una para lo historico y otra para la proyeccion, y cual
    sirve a cual es una propiedad del formato, no del motor.
    """
    flujo = getattr(ctx.informe, "flujo_proyectado", None)
    clase = (getattr(flujo, "clase_de_temporada", None) or {}).get(temporada)
    if not clase:
        return None
    for spec in ctx.mapa.campos.values():
        if clase in (getattr(spec, "sirve_a_clases", None) or []):
            columna, fila = coordinate_from_string(spec.celda)
            return "${}${}".format(columna, fila)
    return None


@registrar("bloques_de_columnas")
class TablaDeBloquesLaterales(EstrategiaTabla):
    """N bloques de columnas en paralelo, con filas fijas por categoria.

    Es la tabla de SUELOS: un predio por bloque, una clase de uso por fila.
    """

    def aplicar(self, ctx: _Contexto, nombre: str, tabla: EspecTabla) -> None:
        items = _como_lista(ctx, nombre, tabla)
        if not items:
            return
        ws = ctx.hoja(tabla.hoja)

        if len(items) > len(tabla.bloques):
            ctx.avisar(
                "tabla '{}': {} elementos para {} bloques; se escriben los {} "
                "primeros".format(nombre, len(items), len(tabla.bloques), len(tabla.bloques))
            )
            items = items[: len(tabla.bloques)]

        for item, bloque in zip(items, tabla.bloques):
            subtablas = getattr(item, "superficies", None)
            if not isinstance(subtablas, dict):
                continue
            for categoria, fila in tabla.filas_clase_uso.items():
                destino = None
                for clave, sub in subtablas.items():
                    if _clave(clave) == categoria:
                        destino = sub
                        break
                if destino is None:
                    continue
                columnas = {
                    k: v for k, v in bloque.items()
                    if k != "indice" and isinstance(v, str)
                }
                _escribir_item(ctx, ws, nombre, destino, fila, columnas)


@registrar("dos_bloques_de_columnas")
class TablaDeBloquesEnParalelo(EstrategiaTabla):
    """Una lista repartida en bloques laterales de igual altura.

    Es la tabla de infraestructura: 4 filas x 2 bloques = 8 items sin tocar
    la plantilla.
    """

    def aplicar(self, ctx: _Contexto, nombre: str, tabla: EspecTabla) -> None:
        items = _como_lista(ctx, nombre, tabla)
        if not items or tabla.primera_fila is None:
            return
        ws = ctx.hoja(tabla.hoja)

        por_bloque = tabla.filas_disponibles or 1
        capacidad = tabla.capacidad or por_bloque * len(tabla.bloques)
        if len(items) > capacidad:
            ctx.avisar(
                "tabla '{}': {} elementos para {} huecos; se escriben los {} "
                "primeros".format(nombre, len(items), capacidad, capacidad)
            )
            items = items[:capacidad]

        for indice, item in enumerate(items):
            bloque = tabla.bloques[min(indice // por_bloque, len(tabla.bloques) - 1)]
            fila = tabla.primera_fila + indice % por_bloque
            columnas = {
                k: v for k, v in bloque.items() if k != "indice" and isinstance(v, str)
            }
            _escribir_item(ctx, ws, nombre, item, fila, columnas)


#: Referencia a otra fila del mismo bloque dentro de una formula declarada.
_REFERENCIA_DE_FILA = re.compile(r"<(\w+)>")


@registrar("bloques_de_filas")
class TablaDeBloquesVerticales(EstrategiaTabla):
    """N bloques apilados, cada uno con las mismas filas de assumptions.

    Es el flujo proyectado: un bloque de 10 filas por especie, con una columna
    por temporada. La primera fila del bloque lleva la etiqueta en la columna A;
    el resto son series temporales.

    Las filas declaradas en ``formulas_de_fila`` no se escriben con el valor del
    modelo sino con una formula de Excel. Es la diferencia entre entregar un
    numero muerto y entregar un libro donde el banco cambia un precio y ve
    moverse el EBITDA; el modelo sigue calculando esas mismas cifras en Python,
    pero para consumo interno (el resumen de portada) y no para escribirlas.
    """

    #: Columna donde va la etiqueta que encabeza cada bloque.
    COLUMNA_ETIQUETA = "A"

    def aplicar(self, ctx: _Contexto, nombre: str, tabla: EspecTabla) -> None:
        items = _como_lista(ctx, nombre, tabla)
        if not items:
            return
        ws = ctx.hoja(tabla.hoja)
        por_temporada = _columnas_de_las_temporadas(ctx, tabla, nombre)
        if not por_temporada:
            return

        if len(items) > len(tabla.bloques):
            ctx.avisar(
                "tabla '{}': {} bloques de datos para {} huecos; se escriben los "
                "{} primeros".format(
                    nombre, len(items), len(tabla.bloques), len(tabla.bloques)
                )
            )
            items = items[: len(tabla.bloques)]

        espec_matriz = ctx.mapa.matriz_envejecimiento
        formulas_matriz = espec_matriz.formulas if espec_matriz else {}

        for indice, (item, bloque) in enumerate(zip(items, tabla.bloques)):
            base = int(bloque["primera_fila"])
            rangos = (ctx.rangos_matriz[indice]
                      if indice < len(ctx.rangos_matriz) else None)
            for campo, salto in tabla.offsets_de_fila.items():
                desde_matriz = formulas_matriz.get(campo) if rangos else None
                if desde_matriz:
                    self._escribir_desde_matriz(
                        ctx, ws, nombre, desde_matriz, base + salto, rangos,
                        por_temporada, item, campo,
                    )
                    continue
                plantilla = tabla.formulas_de_fila.get(campo)
                if plantilla is not None:
                    # Las FORMULAS van en todas las columnas que trae la
                    # plantilla, no solo en las temporadas que este informe
                    # llena: si el agronomo tipea despues la superficie de una
                    # temporada historica, la formula ya esta puesta y el
                    # numero sale solo. Los VALORES, en cambio, van solo donde
                    # hay dato -por eso las dos listas son distintas-.
                    todas = {t: c for t, c in zip(
                        tabla.columnas_por_temporada,
                        tabla.columnas_por_temporada.values())}                         if isinstance(tabla.columnas_por_temporada, dict)                         else por_temporada
                    self._escribir_formulas(
                        ctx, ws, nombre, tabla, plantilla, base, salto,
                        todas,
                    )
                    continue
                if campo not in _campos_del_modelo(item):
                    ctx.avisar(
                        "tabla '{}': el campo '{}' no existe en {}".format(
                            nombre, campo, type(item).__name__
                        )
                    )
                    continue
                valor = getattr(item, campo)
                fila = base + salto
                if isinstance(valor, dict):
                    for temporada, dato in valor.items():
                        columna = por_temporada.get(_clave(temporada))
                        if columna is None:
                            ctx.avisar(
                                "tabla '{}': la temporada '{}' no tiene columna en "
                                "la plantilla".format(nombre, _clave(temporada))
                            )
                            continue
                        ctx.escribir(ws, columna, fila, dato)
                else:
                    ctx.escribir(ws, self.COLUMNA_ETIQUETA, fila, valor)

        self._rotular_filas_espejo(ctx, ws, tabla, items, por_temporada)

    def _rotular_filas_espejo(
        self, ctx: _Contexto, ws: Worksheet, tabla: EspecTabla,
        items: List[Any], por_temporada: Dict[str, str],
    ) -> None:
        """Le pone nombre a las filas que resumen cada bloque.

        El 'RESUMEN DE INDICADORES GENERALES' y el 'ESTADO DE RESULTADOS'
        repiten cuatro filas por indicador -una por bloque de especie- y la
        plantilla del banco NO las rotula: el informe mostraba '69,5' y
        '765.000 kg' sin decir de que especie. Los numeros estaban bien y el
        cuadro era ilegible igual, que para un entregable es lo mismo que estar
        mal.

        La etiqueta se toma del propio bloque, asi que dice exactamente lo
        mismo que su encabezado y no puede quedar desfasada.

        Y el bloque que no existe se apaga: sus filas calculan igual -son
        formulas de la plantilla sobre celdas vacias- y mostraban '0,0' y
        '0 kg' en una fila sin nombre, que es lo que mas confunde. Se ocultan
        con una seccion vacia en el formato, NO borrando la formula: si manana
        el cliente tiene una cuarta especie, el numero aparece solo.
        """
        espec = tabla.filas_espejo
        if not espec:
            return
        columna = espec.get("columna_etiqueta", self.COLUMNA_ETIQUETA)
        campo = espec.get("campo_etiqueta") or ""
        # La plantilla del banco no es consistente consigo misma: pone negrita
        # en la columna A de unos grupos y no en otros, aunque todos son
        # subordinados de su encabezado. Se iguala a la de un grupo declarado
        # para que las ocho listas se lean igual.
        modelo = espec.get("fila_modelo_de_estilo")
        for grupo in espec.get("grupos", []):
            for indice in range(len(tabla.bloques)):
                fila = int(grupo) + indice
                if indice < len(items):
                    etiqueta = getattr(items[indice], campo, None)
                    if etiqueta:
                        ctx.escribir(ws, columna, fila, etiqueta)
                        _igualar_fuente(ws, columna, fila, modelo)
                else:
                    for letra in por_temporada.values():
                        ctx.ceros_a_ocultar.append((ws.title, letra, fila))

    def _escribir_desde_matriz(
        self,
        ctx: _Contexto,
        ws: Worksheet,
        nombre: str,
        plantilla: str,
        fila: int,
        rangos: Dict[str, Any],
        por_temporada: Dict[str, str],
        item: Any,
        campo: str,
    ) -> None:
        """Escribe una fila del flujo apuntando a la hoja anexa.

        Solo donde la matriz tiene esa temporada, que son las de proyeccion.
        Las columnas Real y Preliminar llevan lo que el cliente declara -asi se
        acordo con Fitogreen- y para esas se escribe el valor del modelo, no la
        curva: rotular de 'Real' un supuesto teorico seria enganoso.
        """
        superficies = rangos.get("superficie", {})
        declarado = getattr(item, campo, None)

        for temporada, columna in por_temporada.items():
            rango = superficies.get(_clave(temporada))
            if rango is None:
                if isinstance(declarado, dict):
                    valor = declarado.get(temporada) or declarado.get(_clave(temporada))
                    if valor is not None:
                        ctx.escribir(ws, columna, fila, valor)
                continue
            ctx.escribir(ws, columna, fila, (
                plantilla
                .replace("{superficie}", rango)
                .replace("{rendimiento}", rangos.get("rendimiento", ""))
                .replace("{costo}", rangos.get("costo", ""))
            ))

    def _escribir_formulas(
        self,
        ctx: _Contexto,
        ws: Worksheet,
        nombre: str,
        tabla: EspecTabla,
        plantilla: str,
        base: int,
        salto: int,
        por_temporada: Dict[str, str],
    ) -> None:
        """Escribe una fila derivada como formula, temporada por temporada.

        La fila se llena entera aunque alguna temporada no traiga datos: una
        fila de formula con huecos es justo el patron que rompe los rangos de
        suma de mas abajo. Una columna sin superficie da cero, que es la
        respuesta correcta.
        """
        fila = base + salto

        for columna in por_temporada.values():
            faltantes: List[str] = []

            def resolver(coincidencia: "re.Match[str]") -> str:
                campo = coincidencia.group(1)
                destino = tabla.offsets_de_fila.get(campo)
                if destino is None:
                    faltantes.append(campo)
                    return coincidencia.group(0)
                return "{}{}".format(columna, base + destino)

            formula = _REFERENCIA_DE_FILA.sub(resolver, plantilla)
            if faltantes:
                ctx.avisar(
                    "tabla '{}': la formula de '{}' referencia {}, que no esta "
                    "en offsets_de_fila; no se escribe".format(
                        nombre, plantilla, ", ".join(sorted(set(faltantes)))
                    )
                )
                return
            ctx.escribir(ws, columna, fila, formula)


# -------------------------------------------------------------- el motor

class MotorExcel:
    """Escribe informes sobre la plantilla del banco sin romperla.

    Uso::

        motor = MotorExcel(Path("Formato Fitogreen Bsan.xlsx"),
                           cargar_mapa(Path("config/mapping_santander.json")))
        resultado = motor.generar(informe, Path("output/informe.xlsx"))
    """

    def __init__(self, plantilla: Union[str, Path], mapa: MapaSantander):
        self.plantilla = Path(plantilla)
        self.mapa = mapa
        if not self.plantilla.exists():
            raise FileNotFoundError("no existe la plantilla: {}".format(self.plantilla))

    # -- API publica ------------------------------------------------------

    def generar(
        self,
        informe: InformeFitogreen,
        destino: Union[str, Path],
        crear_galeria_fotos: bool = False,
        excluir: Optional[Sequence[str]] = None,
    ) -> ResultadoGeneracion:
        """Genera el .xlsx completo.

        ``crear_galeria_fotos`` anade una hoja de fotografias que la plantilla
        del banco NO trae. Por defecto va desactivado para no alterar el
        formato entregable.

        ``excluir`` son ids de ``config/mapping_santander.json:
        unidades_excluibles`` -hojas o tablas que el tasador pidio no llevar
        en este informe (el usuario, 2026-09-23: en vez de mantener "distintos
        tipos de informe", sacar lo que no corresponda antes de descargar).
        Nunca se confia en la cascada que ya haya calculado el navegador:
        se vuelve a resolver aca (``_excluir_unidades``).
        """
        destino = Path(destino)
        destino.parent.mkdir(parents=True, exist_ok=True)

        wb = load_workbook(self.plantilla)
        ctx = _Contexto(wb, self.mapa, informe)

        self._escribir_cabecera_de_temporadas(ctx)
        self._escribir_literales(ctx)
        self._escribir_campos(ctx)
        # Antes que las tablas: el flujo necesita saber en que filas quedo cada
        # bloque de la matriz para poder apuntarle con SUMPRODUCT.
        self._escribir_matriz(ctx)
        self._escribir_tablas(ctx)
        self._insertar_fotografias(ctx, crear_galeria_fotos)
        self._escribir_formulas_finales(ctx)
        self._corregir_formulas(ctx)
        # Despues de que todo -datos, matriz, formulas cruzadas- ya esta
        # escrito: mas simple razonar sobre un libro completo que sobre uno a
        # medio llenar. Antes de los formatos: si una tabla se va, no hace
        # falta formatear celdas que ya no estan.
        self._excluir_unidades(ctx, excluir)
        self._aplicar_formatos(ctx)
        self._ocultar_ceros_sin_temporada(ctx)
        self._apagar_ceros_anotados(ctx)
        _preparar_impresion(wb)

        wb.save(destino)
        return ResultadoGeneracion(
            ruta=destino,
            celdas_escritas=ctx.celdas,
            filas_insertadas=ctx.filas_insertadas,
            imagenes_insertadas=ctx.imagenes,
            correcciones=ctx.correcciones,
            avisos=ctx.avisos,
        )

    async def generar_async(
        self,
        informe: InformeFitogreen,
        destino: Union[str, Path],
        crear_galeria_fotos: bool = False,
        excluir: Optional[Sequence[str]] = None,
    ) -> ResultadoGeneracion:
        """Version para FastAPI.

        openpyxl es sincrono y consume CPU, asi que se delega a un hilo en vez
        de fingir concurrencia: el event loop no se bloquea.
        """
        return await asyncio.to_thread(
            self.generar, informe, destino, crear_galeria_fotos, excluir
        )

    # -- fases ------------------------------------------------------------

    def _escribir_formulas_finales(self, ctx: _Contexto) -> None:
        """Formulas que cruzan de hoja y por eso van al final.

        Ver ``EspecFormulaFinal``: aqui ya se sabe cuantas filas inserto cada
        tabla, asi que ``{total:<tabla>}`` puede resolverse a la fila real.
        """
        for espec in self.mapa.formulas_finales:
            ws = ctx.hoja(espec.hoja)
            formula, faltante = self._resolver_totales(ctx, espec.formula)
            if faltante:
                ctx.avisar(
                    "formula final de {}!{}: no existe la tabla '{}'".format(
                        espec.hoja, espec.celda, faltante))
                continue
            objetivo = espec.rango or espec.celda
            if not objetivo:
                ctx.avisar("formula final de {}: sin celda ni rango".format(
                    espec.hoja))
                continue
            for columna, fila in _celdas_del_rango(objetivo):
                destino = ctx.fila_final(ws, fila)
                if espec.solo_si_vacia:
                    ancla = excel_ops.celda_ancla(
                        ws, "{}{}".format(columna, destino))
                    if ws[ancla].value not in (None, ""):
                        continue   # el tasador ya lo declaro: lo suyo manda
                ctx.escribir(ws, columna, destino,
                             _sustituir_columna(
                                 _sustituir_fila(formula, destino), columna))

    def _resolver_totales(
        self, ctx: _Contexto, formula: str
    ) -> Tuple[str, Optional[str]]:
        """Cambia '{total:anexo_real+1}' por el numero de fila que quedo."""
        faltante: Optional[str] = None

        def resolver(coincidencia: "re.Match[str]") -> str:
            nonlocal faltante
            nombre = coincidencia.group(1)
            desplazamiento = int(coincidencia.group(2) or 0)
            tabla = self.mapa.tablas.get(nombre)
            if tabla is None or tabla.fila_totales is None:
                faltante = nombre
                return coincidencia.group(0)
            hoja = ctx.hoja(tabla.hoja)
            return str(ctx.fila_final(hoja, tabla.fila_totales) + desplazamiento)

        return _TOKEN_TOTAL.sub(resolver, formula), faltante

    def _corregir_formulas(self, ctx: _Contexto) -> None:
        """Sustituye las formulas defectuosas que declara el mapping.

        Va al final: si una tabla desbordo e inserto filas, la celda a corregir
        se movio con ellas, y ``ctx.fila_final`` sabe adonde.
        """
        for correccion in self.mapa.correcciones_de_formula:
            ws = ctx.hoja(correccion.hoja)
            plantilla, faltante = self._resolver_totales(ctx, correccion.formula)
            if faltante:
                ctx.avisar(
                    "correccion de {}!{}: no existe la tabla '{}'".format(
                        correccion.hoja, correccion.rango, faltante))
                continue
            for columna, fila in _celdas_del_rango(correccion.rango):
                destino = ctx.fila_final(ws, fila)
                # '{fila}' es la fila YA compensada. Una correccion que apunta a
                # sus propias vecinas con numeros fijos se desalinea en cuanto
                # la tabla desborda, y el resultado no es un error visible sino
                # un numero que parece correcto (§16.a.quinquies).
                formula = _sustituir_fila(
                    _sustituir_columna(plantilla, columna), destino)
                escrito, aviso = excel_ops.reemplazar_formula(
                    ws, "{}{}".format(columna, destino), formula
                )
                if escrito:
                    ctx.celdas += 1
                    ctx.correcciones.append(
                        "{}!{}{}: {}".format(
                            ws.title, columna, destino, correccion.motivo)
                    )
                if aviso:
                    ctx.avisar(aviso)

    def _escribir_cabecera_de_temporadas(self, ctx: _Contexto) -> None:
        """Rotula las columnas con las temporadas que el informe tiene de verdad.

        Ver ``EspecCabeceraTemporadas``. Una temporada sin datos se deja en
        blanco para que la columna no aparezca: es lo que pidio el agrónomo, que
        no siempre les entregan los anios hacia atras.
        """
        espec = self.mapa.cabecera_temporadas
        flujo = getattr(ctx.informe, "flujo_proyectado", None)
        temporadas = list(getattr(flujo, "temporadas", None) or [])
        if espec is None or not temporadas:
            return

        ws = ctx.hoja(espec.hoja)
        tabla = self.mapa.tablas.get("flujo_assumptions")
        columnas = (list(tabla.columnas_por_temporada.values())
                    if tabla and isinstance(tabla.columnas_por_temporada, dict)
                    else [])
        clases = getattr(flujo, "clase_de_temporada", None) or {}

        todas_las_filas = (list(espec.filas_temporada)
                           + list(espec.filas_clase)
                           + list(espec.filas_solo_declaradas))

        for temporada, columna in zip(temporadas, columnas):
            if not temporada:
                # Hueco reservado: la columna queda limpia.
                for fila in todas_las_filas:
                    excel_ops.reemplazar_literal(
                        ws, "{}{}".format(columna, fila), None)
                continue
            for fila in espec.filas_temporada:
                excel_ops.reemplazar_literal(
                    ws, "{}{}".format(columna, fila), temporada)
            # Si el informe no dice de que tipo es la temporada, se deja el
            # rotulo de la plantilla: borrarlo seria destruir lo unico que hay
            # sin tener con que reemplazarlo.
            clase = clases.get(temporada)
            if clase:
                for fila in espec.filas_clase:
                    excel_ops.reemplazar_literal(
                        ws, "{}{}".format(columna, fila), clase)
            # Las tablas que solo aplican a lo que el cliente declaro: una
            # proyeccion no tiene IVAs que la respalden, asi que su columna se
            # deja SIN rotular. Un anio escrito con la celda vacia debajo se lee
            # como un dato que falta, y no es que falte: no puede existir.
            declarada = clase in espec.clases_declaradas
            for fila in espec.filas_solo_declaradas:
                excel_ops.reemplazar_literal(
                    ws, "{}{}".format(columna, fila),
                    temporada if declarada else None)
            ctx.celdas += len(espec.filas_temporada)

        # Otras hojas que repiten las mismas temporadas -la caratula lista
        # los 8 anios del horizonte-. Sin esto sus rotulos quedan con los anios
        # que traia la plantilla y muestran una temporada corrida respecto de
        # la columna que de verdad estan leyendo.
        for extra in espec.hojas_extra:
            try:
                hoja_extra = ctx.hoja(extra["hoja"])
            except KeyError:
                ctx.avisar("cabecera: la hoja '{}' no existe".format(
                    extra.get("hoja")))
                continue
            inicio = int(extra.get("desde_temporada", 0))
            primera = column_index_from_string(extra.get("desde_columna", "B"))
            fila_extra = int(extra["fila"])
            # Cuantas columnas tiene esa tabla. Sin el tope, el bucle escribe
            # una temporada de mas en la celda vacia de al lado y el rotulo
            # aparece FUERA del recuadro del informe.
            cuantas = int(extra.get("columnas", 0)) or None
            for salto, temporada in enumerate(temporadas[inicio:]):
                if cuantas is not None and salto >= cuantas:
                    break
                letra = get_column_letter(primera + salto)
                if not excel_ops.reemplazar_literal(
                        hoja_extra, "{}{}".format(letra, fila_extra),
                        temporada or None):
                    break
                ctx.celdas += 1

        # Las columnas que sobran quedan limpias: si el cliente no entrego ese
        # anio, la columna no debe aparecer rotulada como si existiera.
        for columna in columnas[len(temporadas):]:
            for fila in todas_las_filas:
                excel_ops.reemplazar_literal(
                    ws, "{}{}".format(columna, fila), None)

    def _ocultar_ceros_sin_temporada(self, ctx: _Contexto) -> None:
        """Los ceros de las columnas sin anio, en blanco.

        Va DESPUES de ``_aplicar_formatos``: si se hiciera antes, el formato
        general del rango lo pisaria y los ceros volverian a verse.
        """
        espec = self.mapa.cabecera_temporadas
        flujo = getattr(ctx.informe, "flujo_proyectado", None)
        temporadas = list(getattr(flujo, "temporadas", None) or [])
        if espec is None or not temporadas or not espec.filas_a_ocultar:
            return

        ws = ctx.hoja(espec.hoja)
        tabla = self.mapa.tablas.get("flujo_assumptions")
        columnas = (list(tabla.columnas_por_temporada.values())
                    if tabla and isinstance(tabla.columnas_por_temporada, dict)
                    else [])
        for temporada, columna in zip(temporadas, columnas):
            if not temporada:
                _ocultar_ceros(ws, columna, espec.filas_a_ocultar)
        for columna in columnas[len(temporadas):]:
            _ocultar_ceros(ws, columna, espec.filas_a_ocultar)

    def _apagar_ceros_anotados(self, ctx: _Contexto) -> None:
        """Las celdas sueltas que alguna estrategia pidio dejar en blanco.

        Hoy son las filas de resumen de un bloque de especie que el cliente no
        tiene: calculan cero sobre celdas vacias y lo muestran en una fila sin
        rotulo. Va aqui, con el resto, porque tiene que correr DESPUES de
        ``_aplicar_formatos``.
        """
        for hoja, columna, fila in ctx.ceros_a_ocultar:
            _ocultar_ceros(ctx.wb[hoja], columna, [fila, fila])

    def _escribir_literales(self, ctx: _Contexto) -> None:
        """Rotulos que la plantilla no trae. Ver ``EspecLiteral``."""
        for espec in self.mapa.literales:
            ws = ctx.hoja(espec.hoja)
            columna, fila = _partir_celda(espec.celda)
            ctx.escribir(ws, columna, fila, espec.texto)

    def _escribir_campos(self, ctx: _Contexto) -> None:
        for ruta, spec in self.mapa.campos.items():
            ws = ctx.hoja(spec.hoja)
            valor = self._valor_de_campo(ctx, ruta, spec)
            columna, fila = _partir_celda(spec.celda)
            ctx.escribir(ws, columna, fila, valor,
                         self._mascara(ctx, ruta, spec, valor))
            if valor is not None and spec.area_visual:
                _acomodar_bloque_de_texto(ws, spec.celda, spec.area_visual)

    def _mascara(self, ctx: _Contexto, ruta: str, spec, valor: Any) -> Optional[str]:
        """El formato de un campo: nombre del catalogo o mascara literal.

        Dos cosas, y las dos por el mismo error cometido dos veces.

        Primero, ``formato`` admite el NOMBRE de un formato del catalogo
        (``"decimal2"``), no solo la mascara de Excel. Escribir el nombre donde
        iba la mascara es facilisimo -son el mismo campo- y el resultado no se
        parece a un error.

        Segundo, el guardia. openpyxl decide si una celda lleva fecha mirando
        las LETRAS del formato: ``d`` es dia, ``m`` mes, ``s`` segundo. La
        mascara ``decimal2`` tiene d, m y varias mas, asi que al releer el
        archivo openpyxl convirtio 39.412,55 en el 26 de noviembre de 2007.
        Ninguna celda quedo en error, ninguna prueba fallo: el numero
        simplemente dejo de ser el numero. Es el mismo mecanismo que ya habia
        destruido valores con ``USD #,##0`` (§12), asi que aqui se avisa en vez
        de escribir la fecha en silencio.
        """
        declarado = spec.formato
        if not declarado:
            return None

        del_catalogo = ctx.mapa.formatos.mascara(declarado)
        if del_catalogo:
            return del_catalogo

        es_fecha = isinstance(valor, (dt.date, dt.datetime))
        numerico = isinstance(valor, (int, float))
        if numerico and not es_fecha and is_date_format(declarado):
            ctx.avisar(
                "campo '{}': el formato {!r} no esta en el catalogo y openpyxl "
                "lo lee como formato de FECHA; el numero se convertiria en una "
                "fecha. Se escribe sin formato".format(ruta, declarado))
            return None
        return declarado

    def _valor_de_campo(self, ctx: _Contexto, ruta: str, spec) -> Any:
        """Resuelve el valor de un campo simple.

        Un campo puede apuntar a una serie por temporada y quedarse con una
        sola clave (``clave_temporada`` en el mapping); en ese caso la ruta
        lleva un sufijo descriptivo que no existe en el modelo y se recorta.
        """
        if spec.clave_temporada:
            base = ruta.rsplit(".", 1)[0]
            serie = resolver(ctx.informe, base)
            if isinstance(serie, dict):
                return serie.get(spec.clave_temporada)
            return None
        return resolver(ctx.informe, ruta)

    def _nombre_real_de_hoja(self, ctx: _Contexto, alias_o_nombre: str) -> Optional[str]:
        """Resuelve un ``hoja`` de ``unidades_excluibles``.

        Puede ser un alias del mapping, como el resto del archivo
        (``ctx.mapa.nombre_hoja``), o -para una hoja que CREA el motor y por
        eso no tiene alias, como 'Fotografias'- el nombre real directo.
        """
        try:
            return ctx.mapa.nombre_hoja(alias_o_nombre)
        except KeyError:
            return alias_o_nombre if alias_o_nombre in ctx.wb.sheetnames else None

    def _excluir_unidades(self, ctx: _Contexto, excluir: Optional[Sequence[str]]) -> None:
        """Saca del libro lo que el tasador pidio no llevar en este informe.

        Nunca se confia en la lista que mando el navegador tal cual: la
        cascada de dependencias (``services.exclusiones.resolver_cascada``)
        se vuelve a resolver aca, que es la unica fuente de verdad sobre que
        formula lee celdas de donde.
        """
        if not excluir:
            return
        unidades = ctx.mapa.unidades_excluibles
        a_excluir, avisos = resolver_cascada(unidades, list(excluir))
        for aviso in avisos:
            ctx.avisar(aviso)
        if not a_excluir:
            return

        # 1) hojas completas primero: no reordenan nada DENTRO de otra hoja,
        # asi que el orden entre ellas no importa.
        hojas_sacadas = set()
        for id_ in a_excluir:
            unidad = unidades.get(id_)
            if unidad is None or unidad.nivel != "hoja":
                continue
            nombre = self._nombre_real_de_hoja(ctx, unidad.hoja)
            if nombre is None:
                # Un alias declarado en 'hojas' que no resuelve a ninguna
                # pestana real es un error del mapping: se avisa. Un nombre
                # LITERAL sin alias -como 'Fotografias'- es una hoja que crea
                # el motor solo si hace falta (fotos de sobra): que no
                # exista es el caso normal, no un error, y no se avisa.
                if unidad.hoja in ctx.mapa.hojas:
                    ctx.avisar("unidad excluible '{}': la hoja '{}' no existe".format(
                        unidad.etiqueta, unidad.hoja))
                continue
            excel_ops.eliminar_hoja(ctx.wb, nombre)
            hojas_sacadas.add(nombre)
        ctx.hojas_excluidas |= hojas_sacadas

        # 2) tablas de fila, agrupadas por hoja y de ABAJO HACIA ARRIBA
        # dentro de cada una -mismo criterio que _escribir_tablas al
        # insertar (linea ~1599): borrar la de mas abajo primero no
        # invalida el rango, todavia no tocado, de la de mas arriba.
        por_hoja: Dict[str, List] = {}
        columnas_pendientes = []
        for id_ in a_excluir:
            unidad = unidades.get(id_)
            if unidad is None or unidad.nivel == "hoja":
                continue
            nombre = self._nombre_real_de_hoja(ctx, unidad.hoja)
            if nombre is None or nombre in hojas_sacadas:
                continue  # la hoja entera ya se fue, o nunca existio
            if unidad.nivel == "tabla":
                por_hoja.setdefault(nombre, []).append(unidad)
            elif unidad.nivel == "tabla_columnas":
                columnas_pendientes.append((nombre, unidad))
            else:
                ctx.avisar("unidad excluible '{}': nivel desconocido '{}'".format(
                    unidad.etiqueta, unidad.nivel))

        for nombre, lista in por_hoja.items():
            ws = ctx.wb[nombre]
            # 'filas' en el mapping es la posicion en la plantilla ORIGINAL;
            # si otra tabla de la misma hoja desbordo antes (insertar_filas),
            # el rango real quedo mas abajo. `fila_final` es la misma
            # traduccion que ya usa `_aplicar_formatos` para lo mismo.
            for unidad in sorted(lista, key=lambda u: u.filas[0], reverse=True):
                desde = ctx.fila_final(ws, unidad.filas[0])
                hasta = ctx.fila_final(ws, unidad.filas[1])
                cantidad = hasta - desde + 1
                excel_ops.eliminar_filas(ws, desde, cantidad)
                # Se registra como una insercion NEGATIVA -mismo mecanismo
                # que ya usa `fila_final` para las filas que desbordaron al
                # escribir-, asi que una tabla que se borra DESPUES en la
                # misma hoja (mas arriba, procesada despues por el orden de
                # abajo-hacia-arriba) o el vaciado de columnas del paso 3
                # siguen viendo la posicion real. El ancla va en coordenadas
                # de la PLANTILLA (``unidad.filas``, no ``hasta``): todo lo
                # demas que hay en `ctx.inserciones` esta en esa misma
                # unidad, y mezclarla con la posicion ya desplazada
                # comparia manzanas con naranjas en la proxima consulta.
                ctx.inserciones.setdefault(nombre, []).append(
                    (unidad.filas[1] + 1, -cantidad))

        # 3) columnas vaciadas: no mueven ninguna fila, el orden no importa.
        for nombre, unidad in columnas_pendientes:
            ws = ctx.wb[nombre]
            desde = ctx.fila_final(ws, unidad.filas[0])
            hasta = ctx.fila_final(ws, unidad.filas[1])
            col_ini, col_fin = unidad.columnas
            excel_ops.vaciar_columnas(ws, desde, hasta, col_ini, col_fin)

    def _aplicar_formatos(self, ctx: _Contexto) -> None:
        """Pone el formato numerico declarado en el mapping, rango por rango.

        Se corre al final y compensa las filas insertadas: si una tabla
        desbordo, el rango declarado sobre la plantilla original ya no apunta
        donde apuntaba. Es la misma correccion que usa el resto del motor.

        No toca ningun valor, solo ``number_format``. Una celda vacia tambien
        se formatea: si el agronomo escribe ahi despues, sale bien puesta.
        """
        espec = ctx.mapa.formatos
        if not espec.rangos:
            return

        for alias, rangos in espec.rangos.items():
            try:
                ws = ctx.hoja(alias)
            except KeyError:
                # Si el tasador pidio sacar esta hoja, que no aparezca no es
                # un error del mapping: es lo que se le pidio.
                if self.mapa.hojas.get(alias) not in ctx.hojas_excluidas:
                    ctx.avisar("formatos: la hoja '{}' no existe".format(alias))
                continue

            for rango, nombre in rangos.items():
                mascara = espec.mascara(nombre)
                if mascara is None:
                    ctx.avisar(
                        "formatos: '{}' no esta en el catalogo".format(nombre))
                    continue
                try:
                    col_ini, fila_ini, col_fin, fila_fin = range_boundaries(rango)
                except ValueError:
                    ctx.avisar("formatos: rango invalido '{}'".format(rango))
                    continue

                # Se traduce el RANGO entero, no fila por fila. Con una fila
                # insertada dentro del rango, mapear cada fila por separado
                # produce destinos salteados (33,34,35,37) y la fila nueva -la
                # 36- se queda sin formato: en el informe salia '12' entre
                # '24,5 ha' y '9,5 ha'. Tomando extremos y recorriendo seguido,
                # las filas insertadas heredan el formato de su tabla.
                destino_ini = ctx.fila_final(ws, fila_ini)
                destino_fin = ctx.fila_final(ws, fila_fin)
                for destino in range(destino_ini, destino_fin + 1):
                    for columna in range(col_ini, col_fin + 1):
                        ws.cell(destino, columna).number_format = mascara

                ancho = espec.ancho(nombre)
                if ancho:
                    objetivo = ancho + MARGEN_DE_ANCHO
                    for columna in range(col_ini, col_fin + 1):
                        letra = get_column_letter(columna)
                        actual = ws.column_dimensions[letra].width or 0
                        # Solo se ensancha: hay columnas que ya son anchas por
                        # el texto que llevan y encogerlas seria empeorarlas.
                        if actual < objetivo:
                            ws.column_dimensions[letra].width = objetivo

    def _escribir_matriz(self, ctx: _Contexto) -> None:
        """Vuelca la matriz de envejecimiento a una hoja anexa.

        La plantilla del banco no trae esta hoja. Existe porque el flujo solo
        admite 4 bloques con UNA serie de rendimiento cada uno, y ese numero es
        el promedio ponderado de plantaciones de distintas edades: sin abrirlo
        por edad, el banco recibe un promedio que no puede auditar ni mover.
        Es el patron del archivo HdU.

        Se escribe antes que las tablas porque el flujo necesita los rangos.
        """
        flujo = getattr(ctx.informe, "flujo_proyectado", None)
        matriz = getattr(flujo, "matriz", None) if flujo else None
        espec = ctx.mapa.matriz_envejecimiento
        if matriz is None or not matriz.bloques or espec is None:
            return

        nombre = espec.nombre_hoja_propuesto
        ws = ctx.wb[nombre] if nombre in ctx.wb.sheetnames \
            else ctx.wb.create_sheet(nombre)
        columnas = espec.columnas
        col_edad = columnas.get("edad", "A")
        col_rend = columnas.get("rendimiento_curva", "B")
        col_costo_base = columnas.get("costo_base", "C")
        col_rend_aj = columnas.get("rendimiento_ajustado", "D")
        col_costo = columnas.get("costo_total", "E")
        col_edad_i = column_index_from_string(col_edad)

        temporadas = list(matriz.temporadas)
        columnas_temporada = {
            temporada: get_column_letter(
                column_index_from_string(espec.primera_columna_temporada) + i
            )
            for i, temporada in enumerate(temporadas)
        }
        # La barra de titulo y cada barra de bloque abarcan hasta la ULTIMA
        # columna que de verdad se usa: la de costo total, o la ultima
        # temporada si hay mas alla de esa (siempre la hay).
        ultima_columna = column_index_from_string(col_costo)
        if columnas_temporada:
            ultima_columna = max(
                ultima_columna,
                column_index_from_string(list(columnas_temporada.values())[-1]))

        # Esta hoja tampoco viene en la plantilla del banco -mismo problema
        # que la galeria de fotos-: salia en texto plano, sin un solo borde
        # ni una fuente que coincidiera con el resto del informe. Ahora usa el
        # mismo lenguaje visual: barra roja de titulo, barra gris por bloque
        # de especie, encabezados en negrita y tabla con marco.
        _barra_titulo(ws, "DETALLE DE PLANTACIONES POR EDAD", 2,
                     col_edad_i, ultima_columna)
        ws.cell(3, col_edad_i,
               "Sustenta el rendimiento y el costo promedio del flujo "
               "proyectado: la superficie de cada edad, temporada a "
               "temporada.").font = _fuente_banco(italic=True, color="FF595959")

        fila = espec.primera_fila
        for bloque in matriz.bloques:
            _barra_seccion(ws, bloque.etiqueta, fila, col_edad_i, ultima_columna)
            fila_curva = fila + 1
            fila_factor = fila + 2
            fila_costo_kg = fila + 3
            ws["{}{}".format(col_edad, fila_curva)] = "Curva"
            ws["{}{}".format(col_rend, fila_curva)] = bloque.curva
            ws["{}{}".format(col_edad, fila_factor)] = "Factor de ajuste"
            ws["{}{}".format(col_rend, fila_factor)] = bloque.factor_ajuste
            ws["{}{}".format(col_edad, fila_costo_kg)] = "Costo variable (USD/kg)"
            ws["{}{}".format(col_rend, fila_costo_kg)] = bloque.costo_por_kg_usd
            # Los tres rotulos, en negrita: son la ficha del bloque, no una
            # fila de datos mas.
            col_rend_i = column_index_from_string(col_rend)
            for f in (fila_curva, fila_factor, fila_costo_kg):
                ws.cell(f, col_edad_i).font = _fuente_banco(bold=True)
                ws.cell(f, col_rend_i).number_format = "0.00"
            _bordear(ws, fila_curva, fila_costo_kg, col_edad_i, col_rend_i)

            encabezado = fila + 4
            ws["{}{}".format(col_edad, encabezado)] = "Edad"
            ws["{}{}".format(col_rend, encabezado)] = "Rend. curva (kg/ha)"
            ws["{}{}".format(col_costo_base, encabezado)] = "Costo base (USD/ha)"
            ws["{}{}".format(col_rend_aj, encabezado)] = "Rend. ajustado (kg/ha)"
            ws["{}{}".format(col_costo, encabezado)] = "Costo total (USD/ha)"
            for temporada, columna in columnas_temporada.items():
                # «(ha)»: es superficie, y el encabezado de esta tabla no lo
                # dice en ningun otro lado -al reves de 'Rend. curva (kg/ha)',
                # que ya lleva su unidad y por eso el dato de abajo va desnudo
                # (§16.a.duodecies)-.
                ws["{}{}".format(columna, encabezado)] = "{} (ha)".format(temporada)
            for col in range(col_edad_i, ultima_columna + 1):
                celda = ws.cell(encabezado, col)
                celda.font = _fuente_banco(bold=True)
                celda.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True)

            primera = encabezado + 1
            for indice, datos in enumerate(bloque.filas):
                actual = primera + indice
                ws["{}{}".format(col_edad, actual)] = datos.edad
                ws["{}{}".format(col_rend, actual)] = datos.rendimiento_curva_kg_ha
                ws["{}{}".format(col_costo_base, actual)] = datos.costo_base_usd_ha
                # El ajuste del tasador queda como celda visible y la columna lo
                # referencia: si el banco lo discute, se cambia un numero y se
                # mueve todo el flujo.
                ws["{}{}".format(col_rend_aj, actual)] = "={}{}*${}${}".format(
                    col_rend, actual, col_rend, fila_factor)
                ws["{}{}".format(col_costo, actual)] = "={}{}+{}{}*${}${}".format(
                    col_costo_base, actual, col_rend_aj, actual,
                    col_rend, fila_costo_kg)
                for temporada, columna in columnas_temporada.items():
                    ws["{}{}".format(columna, actual)] = (
                        datos.superficie_por_temporada.get(temporada, 0.0)
                    )
                # Desnudo -sin "kg/ha" ni "USD/ha" por celda-: el encabezado ya
                # lo dice, y repetirlo 30 filas seguidas es ruido (§16.a.
                # duodecies). Las temporadas llevan un decimal: son hectareas
                # de una plantacion real, no un conteo entero.
                for col in (col_rend, col_costo_base, col_rend_aj, col_costo):
                    ws["{}{}".format(col, actual)].number_format = "#,##0"
                ws["{}{}".format(col_edad, actual)].number_format = "0"
                for columna in columnas_temporada.values():
                    ws["{}{}".format(columna, actual)].number_format = "#,##0.0"
            ultima = primera + len(bloque.filas) - 1
            # Dos cajas, no una: entre 'Costo total' y la primera temporada
            # la plantilla deja una columna en blanco a proposito (separa la
            # economia por hectarea de la superficie por temporada). Bordear
            # TODO seguido convertia ese hueco en una celda vacia dentro de
            # la tabla, que se lee como un dato que falta.
            col_costo_i = column_index_from_string(col_costo)
            _bordear(ws, encabezado, ultima, col_edad_i, col_costo_i)
            if columnas_temporada:
                primera_col_temp = column_index_from_string(
                    espec.primera_columna_temporada)
                _bordear(ws, encabezado, ultima, primera_col_temp, ultima_columna)
            for col in range(col_edad_i, ultima_columna + 1):
                for f in range(primera, ultima + 1):
                    ws.cell(f, col).alignment = Alignment(horizontal="center")
                ws.cell(encabezado, col).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True)

            def rango(columna: str, fijar: bool = False) -> str:
                marca = "$" if fijar else ""
                return "'{}'!{}{}{}:{}{}{}".format(
                    nombre, marca, columna, primera, marca, columna, ultima
                )

            ctx.rangos_matriz.append({
                "rendimiento": rango(col_rend_aj, fijar=True),
                "costo": rango(col_costo, fijar=True),
                "superficie": {
                    _clave(temporada): rango(columna)
                    for temporada, columna in columnas_temporada.items()
                },
            })
            ctx.celdas += len(bloque.filas) * (5 + len(temporadas))
            fila = ultima + 1 + espec.filas_entre_bloques

        # La columna A hace doble papel -«Edad» abajo, «Costo variable
        # (USD/kg)» arriba, que es la etiqueta mas larga de la ficha del
        # bloque-: se ajusta a esa, o el rotulo sale cortado.
        ws.column_dimensions[col_edad].width = 24
        for columna in (col_rend, col_costo_base, col_rend_aj, col_costo):
            ws.column_dimensions[columna].width = 20
        for columna in columnas_temporada.values():
            ws.column_dimensions[columna].width = 13

        if matriz.avisos:
            fila += 1
            ws.cell(fila, col_edad_i, "Fuera del flujo:").font = (
                _fuente_banco(bold=True))
            for aviso in matriz.avisos:
                fila += 1
                ws.cell(fila, col_edad_i, aviso).font = _fuente_banco()

    def _escribir_tablas(self, ctx: _Contexto) -> None:
        """Agrupa por hoja y procesa de abajo hacia arriba."""
        por_hoja: Dict[str, List[Tuple[int, str, EspecTabla, EstrategiaTabla]]] = {}

        for nombre, tabla in self.mapa.tablas.items():
            estrategia = ESTRATEGIAS.get(tabla.orientacion)
            if estrategia is None:
                ctx.avisar(
                    "tabla '{}': orientacion desconocida '{}'; se omite".format(
                        nombre, tabla.orientacion
                    )
                )
                continue
            por_hoja.setdefault(tabla.hoja, []).append(
                (estrategia.fila_referencia(tabla), nombre, tabla, estrategia)
            )

        for entradas in por_hoja.values():
            # descendente: la tabla mas baja primero, para que las inserciones
            # de las de arriba arrastren lo ya escrito sin invalidar el mapping
            for _, nombre, tabla, estrategia in sorted(
                entradas, key=lambda e: e[0], reverse=True
            ):
                estrategia.aplicar(ctx, nombre, tabla)

    def _insertar_fotografias(self, ctx: _Contexto, crear_galeria: bool) -> None:
        fotos = list(ctx.informe.fotografias)
        if not fotos:
            return

        pendientes: List[Tuple[int, Fotografia]] = []
        # Cuantas caben en cada hueco de la plantilla. Sin llevar la cuenta,
        # tres fotos de emplazamiento acaban pegadas una encima de otra en la
        # misma celda: se ve una sola y las demas quedan invisibles.
        #
        # **La cuenta va por HUECO, no por categoria.** La Bsan trae un solo
        # marcador de imagen y lo comparten 'satelital' y 'emplazamiento'
        # (2026-09-09): contando por categoria, cada una creeria que el hueco
        # esta libre y la segunda se pegaria ENCIMA de la primera -sin aviso,
        # sin celda en error, y con la de abajo invisible-.
        ocupacion: Dict[Tuple[str, str], int] = {}

        # Se recorre por prioridad y no por orden de subida: si dos categorias
        # comparten hueco, gana la que el mapping declara antes, no la que
        # llegue primero en la lista. Ver `EspecFoto.prioridad`.
        def _turno(par: Tuple[int, Fotografia]) -> Tuple[int, int]:
            orden, foto = par
            spec = self.mapa.fotografias.soportado_por_plantilla.get(
                foto.categoria.value
            )
            return (spec.prioridad if spec else 99, orden)

        for orden, foto in sorted(enumerate(fotos), key=_turno):
            spec = self.mapa.fotografias.soportado_por_plantilla.get(
                foto.categoria.value
            )
            if spec is None:
                pendientes.append((orden, foto))
                continue

            hueco = (spec.hoja, spec.celda_ancla)
            usadas = ocupacion.get(hueco, 0)
            if usadas >= spec.capacidad:
                ctx.avisar(
                    "el hueco {}!{} de la plantilla admite {} imagen(es) y ya "
                    "esta ocupado: '{}' va a la hoja de fotografias".format(
                        spec.hoja, spec.celda_ancla, spec.capacidad,
                        foto.titulo or foto.categoria.value,
                    )
                )
                pendientes.append((orden, foto))
                continue
            ocupacion[hueco] = usadas + 1

            ws = ctx.hoja(spec.hoja)
            celda = foto.celda_destino or spec.celda_ancla
            columna, fila = _partir_celda(celda)
            destino = "{}{}".format(columna, ctx.fila_final(ws, fila))
            rango = self._rango_desplazado(ctx, ws, spec.rango_sugerido)
            self._pegar(ctx, ws, foto, destino, rango)

        if not pendientes:
            return
        # A la galeria vuelven en el orden en que las subieron: la prioridad
        # decide quien se queda con el hueco, no como se ordena el anexo.
        sobrantes = [foto for _, foto in sorted(pendientes, key=lambda p: p[0])]
        if crear_galeria:
            self._crear_galeria(ctx, sobrantes)
        else:
            ctx.avisar(
                "{} fotografia(s) sin hueco en la plantilla ({}); active "
                "crear_galeria_fotos para volcarlas en una hoja anexa".format(
                    len(sobrantes),
                    ", ".join(sorted({f.categoria.value for f in sobrantes})),
                )
            )

    def _rango_desplazado(
        self, ctx: _Contexto, ws: Worksheet, rango: Optional[str]
    ) -> Optional[str]:
        if not rango:
            return None
        min_col, min_fila, max_col, max_fila = range_boundaries(rango)
        return "{}{}:{}{}".format(
            get_column_letter(min_col), ctx.fila_final(ws, min_fila),
            get_column_letter(max_col), ctx.fila_final(ws, max_fila),
        )

    def _pegar(
        self,
        ctx: _Contexto,
        ws: Worksheet,
        foto: Fotografia,
        celda: str,
        rango: Optional[str],
    ) -> Optional[Tuple[int, int]]:
        """Incrusta la foto y devuelve el tamano final en pixeles -tal como
        quedo, ya encajada en el hueco-, para que quien la pego pueda dibujar
        un marco del mismo porte y no del porte del hueco entero."""
        ruta = Path(foto.ruta)
        if not ruta.exists():
            ctx.avisar("no existe la imagen: {}".format(ruta))
            return None

        # Normaliza formato y tamano antes de incrustar: las fotos de dron
        # vienen en MPO, que openpyxl no sabe empaquetar.
        imagen = ImagenExcel(excel_ops.preparar_imagen(str(ruta)))
        ancho, alto = foto.ancho_px, foto.alto_px
        if ancho is None or alto is None:
            if rango:
                ancho, alto = _encajar(
                    imagen.width, imagen.height, *excel_ops.dimensiones_rango_px(ws, rango)
                )
            else:
                ancho, alto = imagen.width, imagen.height
        imagen.width, imagen.height = ancho, alto
        ws.add_image(imagen, celda)
        ctx.imagenes += 1
        return ancho, alto

    def _crear_galeria(self, ctx: _Contexto, fotos: List[Fotografia]) -> None:
        """Vuelca las fotos sobrantes en una hoja anexa segun la rejilla del mapping."""
        galeria = self.mapa.fotografias.galeria_inspeccion
        if galeria is None or not galeria.celdas:
            ctx.avisar("el mapping no define rejilla de galeria de fotografias")
            return

        ws = ctx.wb.create_sheet(galeria.nombre_hoja_propuesto)

        # La hoja no viene en la plantilla del banco, asi que no heredaba su
        # barra de titulo: salia en blanco, sin nada que la distinguiera de
        # un volcado de fotos suelto. Misma barra roja que 'COMENTARIOS
        # GENERALES', con el cliente debajo para que se sepa de que informe es
        # aunque esta hoja se mire aparte.
        ultima_columna = max(range_boundaries(c)[2] for c in galeria.celdas)
        _barra_titulo(ws, "REGISTRO FOTOGRÁFICO", 2, 1, ultima_columna)
        antecedentes = getattr(ctx.informe, "antecedentes", None)
        cliente = getattr(antecedentes, "cliente", None) if antecedentes else None
        if cliente:
            rut = getattr(antecedentes, "rut", None)
            ws.cell(3, 1, cliente + (" · " + rut if rut else "")).font = (
                _fuente_banco(italic=True, color="FF595959"))

        rejilla = galeria.rejilla
        if rejilla is not None:
            if rejilla.ancho_columnas:
                columnas = {range_boundaries(c)[0] for c in galeria.celdas}
                anchos = {range_boundaries(c)[2] for c in galeria.celdas}
                for col in range(min(columnas), max(anchos) + 1):
                    ws.column_dimensions[get_column_letter(col)].width = (
                        rejilla.ancho_columnas
                    )
            if rejilla.alto_filas:
                filas = {range_boundaries(c)[1] for c in galeria.celdas}
                altos = {range_boundaries(c)[3] for c in galeria.celdas}
                for fila in range(min(filas), max(altos) + 1):
                    ws.row_dimensions[fila].height = rejilla.alto_filas

        if len(fotos) > len(galeria.celdas):
            ctx.avisar(
                "galeria: {} fotos para {} huecos; se pegan las {} primeras".format(
                    len(fotos), len(galeria.celdas), len(galeria.celdas)
                )
            )
            fotos = fotos[: len(galeria.celdas)]

        # El titulo va en la fila de DEBAJO de cada hueco, combinado a lo ancho
        # de la foto y numerado, como en los informes entregados («1.- Vista
        # general», Cliente D). En la celda del ancla lo tapaba la propia
        # foto: la hoja salia con las imagenes y ningun titulo a la vista.
        ancho_columna = (rejilla.ancho_columnas if rejilla and rejilla.ancho_columnas
                         else 8.43)
        alto_titulo: Dict[int, float] = {}
        for numero, (foto, rango) in enumerate(zip(fotos, galeria.celdas), start=1):
            celda = rango.split(":")[0]
            col_ini, fila_ini, col_fin, fila_fin = range_boundaries(rango)
            fila_titulo = fila_fin + 1
            titulo = ("{}.- {}".format(numero, foto.titulo) if foto.titulo
                      else "{}.-".format(numero))
            ctx.escribir(ws, get_column_letter(col_ini), fila_titulo, titulo)
            if col_fin > col_ini:
                ws.merge_cells(start_row=fila_titulo, start_column=col_ini,
                               end_row=fila_titulo, end_column=col_fin)
            # La misma tipografia del banco: sin esto el titulo sale en la
            # fuente por defecto de openpyxl, que no es la del resto del libro.
            ws.cell(fila_titulo, col_ini).font = _fuente_banco(size=10)
            # Sin wrap_text, un titulo mas largo que su foto se corta a media
            # palabra: Excel solo deja desbordar una celda sobre la vecina
            # cuando esa vecina esta vacia, y aqui nunca lo esta -la foto de al
            # lado tiene su propio titulo-. Con el ancho real de la rejilla se
            # estima cuantas lineas necesita el titulo mas largo del grupo,
            # para que tampoco quede cortado por abajo.
            ws.cell(fila_titulo, col_ini).alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True)
            ancho_merge = ancho_columna * (col_fin - col_ini + 1)
            # El factor 1,3 sale de medir con Excel real cuanto entra en una
            # linea con este ancho de rejilla y la fuente del banco (Aptos
            # Narrow 10): un primer intento con 1,7 sobreestimaba la
            # capacidad, calculaba una sola linea para un titulo que en
            # verdad necesitaba dos, y la segunda linea quedaba recortada por
            # el alto de fila -invisible, sin ningun error-. 1,3 deja margen
            # a proposito: mejor una fila un poco mas alta que una palabra
            # que desaparece.
            caracteres_por_linea = max(int(ancho_merge * 1.3), 15)
            lineas = -(-len(titulo) // caracteres_por_linea)  # division hacia arriba
            alto_titulo[fila_titulo] = max(alto_titulo.get(fila_titulo, 0.0),
                                           lineas * 13.0)

            # el usuario, mirando un informe real: *"sigue viendose mas o menos
            # mal, no podriamos poner algo justo por atras de las fotos que
            # se acomode segun el tamano de las fotos?"*. La foto se ENCAJA
            # en el hueco sin deformarse (`_encajar`), asi que casi nunca lo
            # llena entero -una foto panoramica de dron deja el hueco mas
            # bajo que ancho, y le queda una franja en blanco abajo, sin
            # marco ni fondo que la distinga del resto de la hoja-. Con el
            # tamano real que devuelve `_pegar` se calcula cuantas columnas
            # ocupa la foto DE VERDAD (las filas son uniformes por rejilla,
            # asi que el marco baja siempre hasta el titulo) y se le pone un
            # mate del mismo gris claro que ya usa el banco en sus propias
            # sub-etiquetas, mas un marco -no la rejilla de una tabla de
            # datos, que no pega para una foto-.
            ancho_hueco_px, _ = excel_ops.dimensiones_rango_px(ws, rango)
            tamano_foto = self._pegar(ctx, ws, foto, celda, rango)
            col_fin_foto = col_fin
            if tamano_foto and ancho_hueco_px:
                ancho_foto_px, _ = tamano_foto
                n_columnas = col_fin - col_ini + 1
                fraccion = min(1.0, ancho_foto_px / ancho_hueco_px)
                n_columnas_foto = max(1, math.ceil(fraccion * n_columnas))
                col_fin_foto = min(col_fin, col_ini + n_columnas_foto - 1)
            relleno = PatternFill(fill_type="solid", fgColor=_GRIS_MATE)
            for f in range(fila_ini, fila_fin + 1):
                for c in range(col_ini, col_fin_foto + 1):
                    ws.cell(f, c).fill = relleno
            _enmarcar(ws, fila_ini, fila_titulo, col_ini, col_fin_foto)

        for fila, alto in alto_titulo.items():
            actual = ws.row_dimensions[fila].height
            if actual is None or actual < alto:
                ws.row_dimensions[fila].height = alto


# --------------------------------------------------------------- helpers

def _vestir_columnas_agregadas(
    ctx: _Contexto, ws: Worksheet, tabla, desde: int, hasta: int
) -> None:
    """Le da a una columna agregada el aspecto de la tabla que la recibe.

    La plantilla del banco no trae la superficie del Anexo ni la cobertura de
    riego, asi que esas columnas caen justo fuera del recuadro y se ven como
    numeros sueltos al costado. Se copia el estilo -bordes, relleno, fuente-
    de la columna de al lado, que es la que define como se ve la tabla.
    """
    from copy import copy

    for letra in tabla.columnas_agregadas:
        columna = column_index_from_string(letra)
        vecina = columna - 1
        filas = list(range(desde, hasta + 1))
        for celda in tabla.encabezados:
            col_enc, fila_enc = _partir_celda(celda)
            if col_enc == letra:
                filas.append(ctx.fila_final(ws, fila_enc))
        if tabla.fila_totales:
            filas.append(ctx.fila_final(ws, tabla.fila_totales))

        for fila in filas:
            modelo = ws.cell(fila, vecina)
            destino = ws.cell(fila, columna)
            destino.border = copy(modelo.border)
            destino.fill = copy(modelo.fill)
            destino.font = copy(modelo.font)
            destino.alignment = copy(modelo.alignment)


def _vestir_como_la_columna_modelo(
    ws: Worksheet, modelo: str, columna: str, filas: List[int]
) -> None:
    """Le da a una columna el aspecto que la plantilla dejo hecho en otra.

    Distinto de ``_vestir_columnas_agregadas``, que copia de la vecina: aqui la
    columna de referencia se declara. El bloque de respaldo de ventas trae
    estilada UNA sola columna -el banco esperaba una sola temporada
    respaldada-, asi que las columnas que de verdad llevan los datos salen sin
    recuadro, flotando fuera de la tabla como numeros sueltos.

    Se visten SOLO las columnas con contenido. Estirar el recuadro sobre las
    columnas vacias devolveria el problema que se acaba de resolver: una celda
    dentro de la tabla y sin nada adentro se lee como un dato que falta.
    """
    from copy import copy

    origen = column_index_from_string(modelo)
    destino = column_index_from_string(columna)
    if origen == destino:
        return
    for fila in filas:
        molde = ws.cell(fila, origen)
        celda = ws.cell(fila, destino)
        celda.border = copy(molde.border)
        celda.fill = copy(molde.fill)
        celda.font = copy(molde.font)
        celda.alignment = copy(molde.alignment)


def _acomodar_columnas(
    ws: Worksheet, tabla, desde: int, hasta: int
) -> None:
    """Acomoda el contenido de una columna a su recuadro.

    Dos cosas, y las dos se vieron en la hoja de Inversion. Un texto mas ancho
    que su columna y CENTRADO se recorta por los dos lados -no se desborda-,
    asi que 'Plantacion nogal El Almendro' salia como 'ion nogal El Al' y nadie
    podia saber de que proyecto se trataba; es la misma trampa de §16.a.ter con
    el comentario de gravamenes. Y en una columna combinada a lo alto el valor
    cae al fondo del recuadro: el nombre del proyecto quedaba a una altura, su
    anio a otra y su superficie a otra.
    """
    from copy import copy

    for letra in tabla.columnas_acomodadas:
        columna = column_index_from_string(letra)
        for fila in range(desde, hasta + 1):
            celda = ws.cell(fila, columna)
            alineacion = copy(celda.alignment)
            alineacion.wrap_text = True
            alineacion.vertical = "center"
            celda.alignment = alineacion


def _vestir_fila_de_totales(ctx: _Contexto, ws: Worksheet, tabla) -> None:
    """Mete la fila de totales dentro del recuadro de su tabla.

    Se copian bordes y relleno, NO la alineacion: la de las celdas de datos
    viene centrada y centraria los numeros del total, que van a la derecha
    como el resto de la columna.
    """
    from copy import copy

    modelo = tabla.fila_modelo_de_estilo_totales
    if modelo is None or tabla.fila_totales is None:
        return
    origen = ctx.fila_final(ws, modelo)
    destino = ctx.fila_final(ws, tabla.fila_totales)
    if origen == destino:
        return
    for columna in tabla.columnas.values():
        indice = column_index_from_string(columna)
        celda = ws.cell(destino, indice)
        molde = ws.cell(origen, indice)
        celda.border = copy(molde.border)
        celda.fill = copy(molde.fill)


def _completar_totales(
    ctx: _Contexto, ws: Worksheet, tabla: EspecTabla, desde: int, hasta: int
) -> None:
    """Escribe la suma en la fila de totales cuando la plantilla no la trae.

    Hay tablas del formato donde la fila 'Total (ha)' esta rotulada pero vacia:
    el banco espera que el tasador teclee el numero. Si la columna lleva datos
    numericos y la celda de total esta libre, se pone la formula.

    Es una FORMULA y no un valor a proposito: si despues alguien corrige una
    superficie a mano en Excel, el total se recalcula solo.
    """
    if tabla.fila_totales is None or hasta < desde:
        return

    # La fila de totales tambien se corre cuando la tabla desborda: con 5
    # especies en 4 huecos, 'Total (ha)' baja de la 24 a la 25. Sin esta
    # correccion el motor escribia sobre una fila de datos -que ya tenia
    # contenido, asi que no la pisaba- y la fila de totales real quedaba vacia
    # en el informe entregado.
    fila_totales = ctx.fila_final(ws, tabla.fila_totales)

    sumables = tabla.columnas_sumables
    for campo, columna in tabla.columnas.items():
        if sumables is not None and campo not in sumables:
            continue   # declarado como no sumable: es una tasa, no un monto
        coordenada = "{}{}".format(columna, fila_totales)
        # La fila de totales suele traer celdas combinadas ('Total (ha)' ocupando
        # A24:C24): hay que escribir en el ancla, no en una celda de solo lectura.
        ancla = excel_ops.celda_ancla(ws, coordenada)
        if ancla != coordenada:
            continue   # el hueco pertenece al rotulo de al lado; no es nuestro

        celda = ws[ancla]
        if celda.value is not None:
            continue   # la plantilla ya trae formula o rotulo: no se toca

        valores = [ws["{}{}".format(columna, f)].value for f in range(desde, hasta + 1)]
        if not any(isinstance(v, (int, float)) for v in valores):
            continue   # columna de texto: sumarla no significa nada

        celda.value = "=SUM({col}{desde}:{col}{hasta})".format(
            col=columna, desde=desde, hasta=hasta
        )
        ctx.celdas += 1


#: Lo que hay que sumarle a un ancho declarado para que Excel lo respete.
#:
#: La unidad de ancho de openpyxl y la de Excel no coinciden del todo: un
#: ancho de 18 escrito con openpyxl se lee como 17,3 al abrir el archivo.
#: Medido sobre la columna de USD del Anexo, la diferencia ronda 0,7 y va
#: siempre en contra -queda mas angosto de lo pedido-, que es justo el lado
#: que produce '#####'. Se compensa aqui, en un solo sitio, en vez de inflar
#: cada numero del catalogo de formatos y tener que acordarse cada vez.
MARGEN_DE_ANCHO = 1.5

#: A partir de esta anchura la hoja se imprime apaisada.
COLUMNAS_PARA_APAISAR = 8


def _preparar_impresion(wb: Workbook) -> None:
    """Ajusta cada hoja para que quepa a lo ancho al imprimir o exportar a PDF.

    Sin esto Excel parte las tablas por la mitad: una pagina con las primeras
    columnas y otra con las que sobran, con los encabezados cortados. El informe
    va al banco en PDF o impreso, asi que verlo partido es verlo mal, aunque los
    datos esten perfectos.

    Solo se toca la configuracion de pagina; ni una celda cambia.
    """
    for ws in wb.worksheets:
        ancho = ws.max_column
        ws.page_setup.orientation = (
            "landscape" if ancho > COLUMNAS_PARA_APAISAR else "portrait"
        )
        # Encajar a lo ancho y dejar el alto libre: cortar por filas se lee
        # bien, cortar por columnas no.
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_options.horizontalCentered = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.page_margins.top = ws.page_margins.bottom = 0.5


def _columnas_de_las_temporadas(ctx: _Contexto, tabla, nombre: str):
    """Que columna le toca a cada temporada del informe.

    El mapping declara las temporadas ESCRITAS ("2023/24": "D", ...), y eso
    funciona mientras el informe se haga el ano para el que se escribieron: uno
    hecho en 2027 arranca en 2024/25 y ninguna de sus once temporadas calzaria
    con esas claves, asi que el flujo entero quedaria vacio sin un solo aviso.
    Es el patron que §13 marca como a evitar -el ano base escrito a mano-.

    Asi que se emparejan POR POSICION: la primera temporada del informe va a la
    primera columna declarada, y asi. Las claves del mapping se siguen usando
    para saber el ORDEN de las columnas, no para reconocer el ano.
    """
    declaradas = tabla.columnas_por_temporada
    if not isinstance(declaradas, dict) or not declaradas:
        ctx.avisar("tabla '{}': falta columnas_por_temporada".format(nombre))
        return None

    columnas = list(declaradas.values())
    flujo = getattr(ctx.informe, "flujo_proyectado", None)
    temporadas = list(getattr(flujo, "temporadas", None) or [])
    if not temporadas:
        return declaradas          # sin temporadas propias, manda el mapping

    if len(temporadas) > len(columnas):
        ctx.avisar(
            "tabla '{}': el informe trae {} temporadas y la plantilla tiene {} "
            "columnas; se escriben las {} primeras".format(
                nombre, len(temporadas), len(columnas), len(columnas)))
    # Un hueco (None) reserva la columna sin usarla: es como queda un anio
    # historico que el cliente no entrego, y lo que mantiene alineada la
    # caratula, que lee columnas fijas.
    return {t: c for t, c in zip(temporadas, columnas) if t}


def _igualar_fuente(
    ws: Worksheet, columna: str, fila: int, modelo: Optional[int]
) -> None:
    """Le da a una celda la fuente de otra fila de su misma columna."""
    from copy import copy

    if modelo is None or int(modelo) == fila:
        return
    ws["{}{}".format(columna, fila)].font = copy(
        ws["{}{}".format(columna, modelo)].font)


def _ocultar_ceros(ws: Worksheet, columna: str, rango) -> None:
    """Deja en blanco los ceros de una columna que no corresponde a ningun anio.

    Las formulas de la plantilla calculan igual sobre celdas vacias y muestran
    '0,0' y '0 kg' en una columna sin rotulo, lo que se lee como un dato en
    cero y no como una columna que no existe.

    NO se borran las formulas: se les agrega una seccion vacia al formato -en
    Excel la tercera seccion es la del cero- asi que el dia que el tasador
    escriba ahi, el numero aparece solo. Borrarlas seria destruir la plantilla
    del banco para arreglar un problema de presentacion.
    """
    if not rango or len(rango) < 2:
        return
    for fila in range(int(rango[0]), int(rango[1]) + 1):
        celda = ws["{}{}".format(columna, fila)]
        if celda.value is None:
            continue
        mascara = celda.number_format
        if mascara in (None, "", "General") or ";" in mascara:
            continue
        celda.number_format = "{0};-{0};".format(mascara)


def _acomodar_bloque_de_texto(ws: Worksheet, celda: str, area: str) -> None:
    """Deja un texto largo legible dentro del recuadro que le da la plantilla.

    Sin esto, un comentario de doscientos caracteres en una celda suelta de
    trece de ancho se ve mal de tres formas a la vez, y las tres se dieron en
    el mismo informe:

    - La plantilla trae ``horizontal='center'`` en esa celda. Un texto mas
      ancho que su columna, centrado, se recorta por LOS DOS LADOS: en pantalla
      empieza a mitad de una palabra. Eso es lo que se veia en 'Dotacion!A47'.
    - Sin ``wrap_text`` no hay salto de linea, asi que da igual el alto.
    - La fila 47 mide 4,5 puntos en la plantilla. Aunque ajustara, no cabria
      ni un renglon.

    El recuadro ya esta dibujado con bordes en la plantilla (A47:M53); lo que
    falta es combinarlo para que el texto pueda usar todo ese ancho. Excel no
    autoajusta el alto de una celda combinada, asi que las filas se reparten a
    mano un alto minimo razonable.
    """
    ancla = excel_ops.celda_ancla(ws, celda)
    if ancla != celda:
        return   # ya viene combinada de la plantilla: no se toca

    col_ini, fila_ini, col_fin, fila_fin = range_boundaries(area)
    for rango in list(ws.merged_cells.ranges):
        if (rango.min_row <= fila_ini <= rango.max_row
                and rango.min_col <= col_ini <= rango.max_col):
            return   # ya hay una combinacion ahi: no se pisa

    ws.merge_cells(area)
    ws[ancla].alignment = Alignment(
        horizontal="left", vertical="top", wrap_text=True)

    # Alto minimo por fila del bloque. La plantilla deja algunas en 4,5 pt.
    for fila in range(fila_ini, fila_fin + 1):
        alto = ws.row_dimensions[fila].height
        if alto is None or alto < 15:
            ws.row_dimensions[fila].height = 15


#: '{fila}' y '{fila-2}': la fila donde la formula termina de verdad.
_TOKEN_FILA = re.compile(r"\{fila([+-]\d+)?\}")


def _sustituir_fila(formula: str, fila: int) -> str:
    """Resuelve '{fila}' y '{fila-2}' contra la fila FINAL de la celda.

    Sin esto, una formula que se refiere a sus vecinas queda apuntando a donde
    estaban antes de insertar. Costo un error de los caros: en la caratula, con
    dos filas insertadas, 'RESULTADO OPERACIONAL NETO' seguia restando las dos
    celdas de mas arriba y mostraba el MARGEN BRUTO -un numero mas alto, con
    pinta de correcto y sin ningun error a la vista-.
    """
    def resolver(coincidencia: "re.Match[str]") -> str:
        return str(fila + int(coincidencia.group(1) or 0))

    return _TOKEN_FILA.sub(resolver, formula)


#: '{total:anexo_real}' y '{total:anexo_real+1}' en una formula final.
_TOKEN_TOTAL = re.compile(r"\{total:(\w+)([+-]\d+)?\}")

#: '{col}' y '{col+4}' dentro de una formula declarada en el mapping.
_TOKEN_COLUMNA = re.compile(r"\{col([+-]\d+)?\}")


def _sustituir_columna(formula: str, columna: str) -> str:
    """Resuelve '{col}' a la columna de la celda, y '{col+4}' a la de al lado.

    El desplazamiento hace falta porque no todas las tablas del formato estan
    alineadas entre si: la fila de 8 anios de la caratula ocupa 'B37:I37' pero
    lee 'Flujo!F110:M110', cuatro columnas mas a la derecha. Con una sola
    linea declarada -y '{col+4}'- quedan cubiertas las ocho temporadas.
    """
    base = column_index_from_string(columna)

    def resolver(coincidencia: "re.Match[str]") -> str:
        desplazamiento = int(coincidencia.group(1) or 0)
        destino = base + desplazamiento
        if destino < 1:
            return coincidencia.group(0)
        return get_column_letter(destino)

    return _TOKEN_COLUMNA.sub(resolver, formula)


def _celdas_del_rango(rango: str) -> List[Tuple[str, int]]:
    """'D112:N112' -> [('D',112), ('E',112), ...]. Tambien admite 'D112'."""
    inicio, _, fin = rango.partition(":")
    col_i, fila_i = _partir_celda(inicio)
    if not fin:
        return [(col_i, fila_i)]
    col_f, fila_f = _partir_celda(fin)
    return [
        (get_column_letter(c), f)
        for f in range(fila_i, fila_f + 1)
        for c in range(column_index_from_string(col_i),
                       column_index_from_string(col_f) + 1)
    ]


def _partir_celda(celda: str) -> Tuple[str, int]:
    """'B12' -> ('B', 12)."""
    indice = 0
    while indice < len(celda) and celda[indice].isalpha():
        indice += 1
    return celda[:indice], int(celda[indice:])


def _encajar(
    ancho: int, alto: int, ancho_max: int, alto_max: int
) -> Tuple[int, int]:
    """Escala manteniendo la proporcion para que quepa en el hueco."""
    if ancho <= 0 or alto <= 0:
        return ancho_max, alto_max
    factor = min(ancho_max / ancho, alto_max / alto)
    return max(1, int(ancho * factor)), max(1, int(alto * factor))
