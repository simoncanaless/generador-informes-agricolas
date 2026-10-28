"""Generar una retasacion incompleta, avisando que le falta.

el usuario (2026-10-02): *"para actualizar el informe tiene que tener todo
rellenado tambien, podemos generar el informe sin pedir todo pero
advirtiendo que cosas le faltan?"*. Antes, un solo dato faltante hacia que
el servidor rechazara el informe entero con un error de validacion.

Ahora `leer_borrador()` acepta lo que venga y devuelve la retasacion que SI
se puede escribir, mas una lista de pendientes de dos clases:

- **falta**: un dato marcado `Obligatorio(...)` en schemas/retasacion.py que
  no vino. Su celda queda vacia en el Excel. La lista de que es obligatorio
  vive en el esquema, no aqui.
- **aviso**: un dato que vino pero no se puede usar -un RUT con el digito
  verificador malo, una superficie negativa, mas filas de las que caben en
  la plantilla-. No se escribe: ponerlo en el informe del banco seria
  escribir un numero que el propio sistema sabe que esta mal. Se avisa con
  lo que se escribio, para poder corregirlo.

Determinista, sin IA.
"""

from __future__ import annotations

import copy
from dataclasses import asdict, dataclass
from typing import Any, Dict, List, Optional, Set, Tuple, Type, get_args, get_origin

from pydantic import BaseModel, ValidationError

from schemas.retasacion import (
    PlantacionConRoyalty, PlantacionSinRoyalty, RetasacionRabobank,
)


@dataclass
class Pendiente:
    seccion: str
    detalle: str
    #: "falta" o "aviso" -los mismos estados que ya pinta la pantalla en el
    #: dialogo de "¿Generar con esto pendiente?" (web/index.html).
    tipo: str

    def como_dict(self) -> Dict[str, str]:
        return asdict(self)


#: Que campos nombran una fila en un aviso ("fila 2 (Bodega)"). Por defecto,
#: el primer texto obligatorio que tenga valor; las plantaciones se nombran
#: por especie Y variedad, que es como las distingue el tasador.
_ROTULO_DE_FILA = {
    PlantacionSinRoyalty: ("especie", "variedad"),
    PlantacionConRoyalty: ("especie", "variedad"),
}

#: El mensaje de Pydantic viene en ingles; estos son los que puede producir
#: un dato tipeado en la pantalla. Lo que no este aqui cae al mensaje original.
_MOTIVOS = {
    "greater_than": "tiene que ser mayor que {gt:g}",
    "greater_than_equal": "no puede ser menor que {ge:g}",
    "less_than": "tiene que ser menor que {lt:g}",
    "less_than_equal": "no puede ser mayor que {le:g}",
    "float_parsing": "no es un número",
    "float_type": "no es un número",
    "int_parsing": "no es un número entero",
    "int_type": "no es un número entero",
    "int_from_float": "tiene que ser un número entero",
    "date_parsing": "no es una fecha",
    "date_from_datetime_parsing": "no es una fecha",
    "date_from_datetime_inexact": "no es una fecha",
    "date_type": "no es una fecha",
    "bool_parsing": "tiene que ser sí o no",
    "bool_type": "tiene que ser sí o no",
    "string_type": "tiene que ser texto",
    "enum": "no es una de las opciones ({expected})",
}

#: El orden en que se listan los pendientes: el de las secciones del
#: esquema, que es el mismo de la pantalla.
_ORDEN_DE_SECCION = {
    (campo.title or nombre): i
    for i, (nombre, campo) in enumerate(RetasacionRabobank.model_fields.items())
}


# ---------------------------------------------------------------- limpieza

def _vacio(valor: Any) -> bool:
    if valor is None or valor is False or valor == "":
        return True
    if isinstance(valor, dict):
        return all(_vacio(v) for v in valor.values())
    if isinstance(valor, list):
        return all(_vacio(v) for v in valor)
    return False


def _limpiar(valor: Any) -> Any:
    """Texto en blanco cuenta como no escrito."""
    if isinstance(valor, str):
        valor = valor.strip()
        return valor or None
    if isinstance(valor, dict):
        return {k: _limpiar(v) for k, v in valor.items()}
    if isinstance(valor, list):
        return [_limpiar(v) for v in valor]
    return valor


def _sacar_filas_vacias(datos: Dict[str, Any]) -> Dict[str, List[int]]:
    """Una fila de tabla completamente vacia -el tasador apreto «Agregar» y no
    la lleno- se descarta: no hay nada que avisar de una fila que no dice
    nada, y escrita dejaria un renglon en blanco en medio de la tabla del
    banco. Devuelve, por tabla, la posicion ORIGINAL de cada fila que queda:
    el aviso tiene que decir «fila 3» si en la pantalla es la tercera, aunque
    la segunda se haya descartado."""
    posiciones: Dict[str, List[int]] = {}
    for nombre, valor in list(datos.items()):
        if isinstance(valor, list):
            quedan = [i for i, fila in enumerate(valor)
                      if not (isinstance(fila, dict) and _vacio(fila))]
            datos[nombre] = [valor[i] for i in quedan]
            posiciones[nombre] = quedan
    return posiciones


# ------------------------------------------------- recorrer el esquema

def _desenvolver(anotacion: Any) -> Any:
    """Optional[List[X]] -> X; deja pasar lo que no es generico."""
    for _ in range(3):
        if get_origin(anotacion) is None:
            break
        argumentos = [a for a in get_args(anotacion) if a is not type(None)]
        if not argumentos:
            break
        anotacion = argumentos[0]
    return anotacion


def _es_modelo(tipo: Any) -> bool:
    return isinstance(tipo, type) and issubclass(tipo, BaseModel)


def _etiqueta(campo) -> Optional[str]:
    extra = campo.json_schema_extra or {}
    return extra.get("obligatorio") or campo.title


def _rotulo_de_fila(modelo: Type[BaseModel], fila: Dict[str, Any]) -> str:
    campos = _ROTULO_DE_FILA.get(modelo)
    if campos is None:
        campos = [nombre for nombre, campo in modelo.model_fields.items()
                  if (campo.json_schema_extra or {}).get("obligatorio")
                  and _desenvolver(campo.annotation) is str][:1]
    partes = [str(fila[c]) for c in campos
              if isinstance(fila, dict) and fila.get(c) not in (None, "")]
    return " ".join(partes)


def _nombre_de_fila(numero: int, rotulo: str) -> str:
    return "fila {}{}".format(numero, " ({})".format(rotulo) if rotulo else "")


def _donde(loc: Tuple[Any, ...], datos: Dict[str, Any],
           posiciones: Dict[str, List[int]]) -> Tuple[str, str, Any]:
    """Traduce la ruta de un error de Pydantic a palabras del tasador:
    ("construcciones_hipotecables", 1, "vida_util_total_anios") ->
    ("Construcciones hipotecables", "fila 2 (Bodega), Vida útil (años)",
    <el campo>)."""
    modelo: Any = RetasacionRabobank
    nodo: Any = datos
    seccion, partes, campo = "", [], None
    for parte in loc:
        if isinstance(parte, int):
            fila = nodo[parte] if isinstance(nodo, list) and 0 <= parte < len(nodo) else {}
            original = posiciones.get(loc[0], [])
            numero = (original[parte] if parte < len(original) else parte) + 1
            partes.append(_nombre_de_fila(
                numero, _rotulo_de_fila(modelo, fila) if _es_modelo(modelo) else ""))
            nodo = fila
            continue
        campo = getattr(modelo, "model_fields", {}).get(parte)
        if campo is None:
            partes.append(str(parte))
            break
        nombre = _etiqueta(campo) or parte.replace("_", " ")
        if seccion:
            partes.append(nombre)
        else:
            seccion = nombre
        modelo = _desenvolver(campo.annotation)
        nodo = nodo.get(parte) if isinstance(nodo, dict) else None
    return seccion, ", ".join(partes), campo


def _motivo(error: Dict[str, Any]) -> str:
    tipo = error.get("type", "")
    contexto = error.get("ctx") or {}
    if tipo == "value_error":
        # Los validadores propios (el RUT) ya hablan en castellano.
        return str(contexto.get("error") or error.get("msg", ""))
    plantilla = _MOTIVOS.get(tipo)
    if plantilla:
        try:
            return plantilla.format(**contexto)
        except (KeyError, ValueError):
            pass
    return error.get("msg", "no es válido")


def _mostrar(valor: Any) -> str:
    texto = str(valor)
    return texto if len(texto) <= 40 else texto[:37] + "…"


# ------------------------------------------------------- quitar lo invalido

_BORRAR = object()


def _quitar(datos: Any, loc: Tuple[Any, ...]) -> None:
    nodo = datos
    for parte in loc[:-1]:
        try:
            nodo = nodo[parte]
        except (KeyError, IndexError, TypeError):
            return
    ultimo = loc[-1]
    if isinstance(nodo, dict):
        nodo.pop(ultimo, None)
    elif isinstance(nodo, list) and isinstance(ultimo, int) and ultimo < len(nodo):
        # Una fila entera que no es una fila (no deberia venir de la
        # pantalla): se marca y se saca despues, para no correr los indices
        # de los demas errores de la misma pasada.
        nodo[ultimo] = _BORRAR


def _purgar_borrados(datos: Dict[str, Any], posiciones: Dict[str, List[int]]) -> None:
    for nombre, valor in list(datos.items()):
        if isinstance(valor, list) and any(v is _BORRAR for v in valor):
            quedan = [i for i, v in enumerate(valor) if v is not _BORRAR]
            original = posiciones.get(nombre, list(range(len(valor))))
            datos[nombre] = [valor[i] for i in quedan]
            posiciones[nombre] = [original[i] for i in quedan]


# ---------------------------------------------------------------- publico

def leer_borrador(datos: Optional[Dict[str, Any]]) -> Tuple[RetasacionRabobank, List[Pendiente]]:
    """Lo que venga de la pantalla -o de un archivo anterior- convertido en
    una retasacion que el motor puede escribir, y la lista de lo pendiente.
    Nunca levanta un error de validacion."""
    datos = _limpiar(copy.deepcopy(datos or {}))
    if not isinstance(datos, dict):
        datos = {}
    posiciones = _sacar_filas_vacias(datos)

    avisos: List[Pendiente] = []
    invalidos: Set[Tuple[Any, ...]] = set()

    # Cada pasada quita lo que no se puede usar; como todo campo admite
    # quedar vacio, la segunda pasada ya valida. El tope es solo un seguro.
    for _ in range(6):
        try:
            retasacion = RetasacionRabobank.model_validate(datos)
            break
        except ValidationError as error:
            for e in error.errors():
                loc = tuple(e["loc"])
                if e["type"] == "extra_forbidden":
                    # Un campo que el esquema no conoce no lo tipeo nadie: lo
                    # agrega la lectura de un archivo anterior (ej.
                    # valor_total_clp_sin_convertir de un PDF sin tasa UF, que
                    # ya se aviso al leerlo). Se descarta sin mas.
                    _quitar(datos, loc)
                    continue
                if e["type"] == "too_long" and len(loc) == 1:
                    maximo = (e.get("ctx") or {}).get("max_length")
                    lista = datos.get(loc[0])
                    if isinstance(lista, list) and maximo is not None:
                        sobran = lista[maximo:]
                        del lista[maximo:]
                        posiciones[loc[0]] = posiciones.get(loc[0], [])[:maximo]
                        modelo = _desenvolver(
                            RetasacionRabobank.model_fields[loc[0]].annotation)
                        nombres = [r for r in (_rotulo_de_fila(modelo, f) for f in sobran) if r]
                        avisos.append(Pendiente(
                            seccion=_donde(loc, datos, posiciones)[0],
                            detalle="la plantilla del banco tiene espacio para {} "
                                    "filas; {}{}".format(
                                        maximo,
                                        "quedó fuera la última" if len(sobran) == 1
                                        else "quedaron fuera las últimas {}".format(len(sobran)),
                                        " ({})".format(", ".join(nombres)) if nombres else ""),
                            tipo="aviso"))
                        continue
                seccion, lugar, campo = _donde(loc, datos, posiciones)
                consecuencia = ((campo.json_schema_extra or {}).get("consecuencia")
                                if campo is not None else None)
                avisos.append(Pendiente(
                    seccion=seccion,
                    detalle="{}«{}»: {}; no se escribió{}".format(
                        lugar + " " if lugar else "", _mostrar(e.get("input")),
                        _motivo(e),
                        " — {}".format(consecuencia) if consecuencia else ""),
                    tipo="aviso"))
                invalidos.add(loc)
                _quitar(datos, loc)
            _purgar_borrados(datos, posiciones)
    else:  # pragma: no cover -no deberia ocurrir: todo campo admite vacio
        retasacion = RetasacionRabobank()
        avisos.append(Pendiente("Retasación", "no se pudo leer lo enviado", "aviso"))

    pendientes = faltantes(retasacion, invalidos, posiciones) + avisos
    # sorted es estable: dentro de una seccion queda primero lo que falta y
    # despues lo que vino mal, cada uno en su orden.
    return retasacion, sorted(
        pendientes, key=lambda p: _ORDEN_DE_SECCION.get(p.seccion, len(_ORDEN_DE_SECCION)))


def faltantes(retasacion: RetasacionRabobank,
              ya_avisados: Optional[Set[Tuple[Any, ...]]] = None,
              posiciones: Optional[Dict[str, List[int]]] = None) -> List[Pendiente]:
    """Los datos `Obligatorio(...)` que no estan. Uno por seccion -o por fila
    de tabla-, para que la lista se lea de un vistazo; los que tienen
    consecuencia (UF, dolar, fecha) van aparte, con lo que se descalza."""
    ya_avisados = ya_avisados or set()
    posiciones = posiciones or {}
    salida: List[Pendiente] = []

    def vacios(objeto: BaseModel, prefijo: Tuple[Any, ...]):
        simples, graves = [], []
        for nombre, campo in type(objeto).model_fields.items():
            extra = campo.json_schema_extra or {}
            if not extra.get("obligatorio") or getattr(objeto, nombre) is not None:
                continue
            if prefijo + (nombre,) in ya_avisados:
                continue  # ya dice «X no es válido; no se escribió»
            if extra.get("consecuencia"):
                graves.append((extra["obligatorio"], extra["consecuencia"]))
            else:
                simples.append(extra["obligatorio"])
        return simples, graves

    for nombre, campo in RetasacionRabobank.model_fields.items():
        valor = getattr(retasacion, nombre)
        seccion = campo.title or nombre
        if isinstance(valor, BaseModel):
            simples, graves = vacios(valor, (nombre,))
            for etiqueta, consecuencia in graves:
                salida.append(Pendiente(
                    seccion, "falta {}: {}".format(etiqueta, consecuencia), "falta"))
            if simples:
                salida.append(Pendiente(
                    seccion, "falta: {}".format(", ".join(simples)), "falta"))
        elif isinstance(valor, list):
            original = posiciones.get(nombre, [])
            for i, fila in enumerate(valor):
                if not isinstance(fila, BaseModel):
                    continue
                simples, graves = vacios(fila, (nombre, i))
                faltan = simples + [e for e, _ in graves]
                if faltan:
                    numero = (original[i] if i < len(original) else i) + 1
                    salida.append(Pendiente(
                        seccion, "{}: falta {}".format(
                            _nombre_de_fila(
                                numero, _rotulo_de_fila(type(fila), fila.model_dump())),
                            ", ".join(faltan)), "falta"))

    # La estadistica de las referencias de mercado son formulas del banco
    # (desviacion estandar, intervalo de confianza): con menos de dos
    # referencias Excel no puede calcularlas y las deja en #¡DIV/0!.
    for nombre in ("testigos_propiedades", "testigos_derechos_agua"):
        cantidad = len(getattr(retasacion, nombre))
        seccion = RetasacionRabobank.model_fields[nombre].title
        if cantidad == 0:
            salida.append(Pendiente(
                seccion, "falta: ninguna referencia. La estadística de esa hoja sale "
                         "en error hasta que se agreguen (hacen falta al menos 2)", "falta"))
        elif cantidad == 1:
            salida.append(Pendiente(
                seccion, "una sola referencia: la desviación estándar y el intervalo "
                         "de confianza necesitan al menos 2 y salen en error", "aviso"))
    return salida
