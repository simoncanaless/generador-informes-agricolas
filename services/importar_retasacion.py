"""Lee un .xlsx de retasacion anterior -mismo formato vigente de Rabobank-
para precargar el Paso 2 del actualizador de informes.

Es la lectura en reversa de `services/excel_engine_rabobank.py:MotorRabobank`,
con el MISMO mapping (config/mapping_rabobank.json) -no hay una segunda
tabla de coordenadas que mantener sincronizada.

**100% determinista, sin IA.** Si el archivo subido es de un año o una
version anterior de la plantilla y las celdas no calzan, el campo sale
vacio -nunca un numero inventado-. El tasador completa a mano lo que no se
pudo leer, igual que con cualquier extraccion de este proyecto: mejor un
hueco visible que un dato mudo.

Distinto de leer un PDF (que si necesita la IA para transcribir texto):
un .xlsx tiene celdas con coordenadas conocidas, así que leerlo es tan
determinista como escribirlo.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Tuple, Union

from openpyxl import load_workbook

from services.dibujos_xlsx import leer_textos
from services.mapping import MapaSantander


def _poner_atributo(destino: Dict[str, Any], ruta: str, valor: Any) -> None:
    """'predio.inscripcion.fojas' -> destino['predio']['inscripcion']['fojas']."""
    partes = ruta.split(".")
    actual = destino
    for parte in partes[:-1]:
        actual = actual.setdefault(parte, {})
    actual[partes[-1]] = valor


def _fila_vacia(valores: Dict[str, Any]) -> bool:
    return not any(v not in (None, "") for v in valores.values())


#: Campos que el dominio espera como booleano, pero que la plantilla real
#: escribe como texto ("Sí"/"No") -no son datos de cliente, son la misma
#: convencion de la plantilla en las tres tablas donde aparecen.
CAMPOS_BOOLEANOS = {"hipotecado_rabofinance", "afecto_expropiacion", "asegurable",
                    "hipoteca_rabofinance", "tiene_tasacion_anterior"}


def _normalizar_booleanos(valor: Any) -> Any:
    """Recorre el dict/list leido y convierte 'Sí'/'No' (como los escribe
    la plantilla) a True/False en los campos que el dominio espera booleanos.
    """
    if isinstance(valor, dict):
        salida = {}
        for clave, v in valor.items():
            if clave in CAMPOS_BOOLEANOS and isinstance(v, str):
                bajo = v.strip().lower()
                if bajo in ("sí", "si", "true", "x"):
                    v = True
                elif bajo in ("no", "false"):
                    v = False
            salida[clave] = _normalizar_booleanos(v)
        return salida
    if isinstance(valor, list):
        return [_normalizar_booleanos(v) for v in valor]
    return valor


def _es_rotulo_de_totales(ws, fila: int, rotulo) -> bool:
    valor = ws["{}{}".format(rotulo["columna"], fila)].value
    return isinstance(valor, str) and valor.strip().lower().startswith(rotulo["texto"].lower())


def _leer_tabla_simple(ws, tabla) -> List[Dict[str, Any]]:
    """Las filas de una tabla. Si la tabla admite insertar filas, un informe
    anterior puede traer MAS que la capacidad de la plantilla (tres
    instalaciones donde la plantilla trae una): se sigue leyendo hasta la
    fila de totales, que se reconoce por su rotulo -su posicion cambia con
    cada fila insertada-."""
    capacidad = tabla.capacidad or 1
    rotulo = getattr(tabla, "rotulo_de_totales", None)
    hasta = capacidad + (200 if getattr(tabla, "insertar_filas", False) and rotulo else 0)
    filas = []
    for i in range(hasta):
        f = tabla.primera_fila + i
        if rotulo and _es_rotulo_de_totales(ws, f, rotulo):
            break
        fila = {campo: ws["{}{}".format(col, f)].value
               for campo, col in tabla.columnas.items()}
        if not _fila_vacia(fila):
            filas.append(fila)
    return filas


def _leer_referencias(ws, tabla) -> List[Dict[str, Any]]:
    """REF. PROPIEDADES / REF. Canal son la BASE de referencias de la
    tasadora: las del informe son las que quedaron VISIBLES (las demas estan
    ocultas por el filtro). Se leen las filas visibles con datos dentro del
    rango del autofiltro -que crece con cada informe, asi que no hay una
    fila fija donde terminar-."""
    desde = getattr(tabla, "primera_fila_base", None) or 6
    hasta = desde
    if ws.auto_filter.ref:
        hasta = int("".join(c for c in ws.auto_filter.ref.split(":")[-1] if c.isdigit()))
    filas = []
    for f in range(desde, hasta + 1):
        if ws.row_dimensions[f].hidden:
            continue
        fila = {campo: ws["{}{}".format(col, f)].value
                for campo, col in tabla.columnas.items()}
        if not _fila_vacia(fila):
            filas.append(fila)
    return filas


def _leer_deslindes(ws, spec) -> Dict[str, str]:
    """'Norte: con ...\\nSur: con ...' -> {'deslinde_norte': 'con ...'}."""
    texto = ws[spec["celda"]].value
    if not isinstance(texto, str):
        return {}
    salida = {}
    for rotulo, ruta in spec["partes"].items():
        for linea in texto.splitlines():
            limpia = linea.strip()
            if limpia.lower().startswith(rotulo.lower()):
                salida[ruta] = limpia[len(rotulo):].lstrip(" :").strip()
                break
    return salida


def _leer_ajustes_plant(formulas, plantaciones, mapa) -> None:
    """Los años de OPEX (T) y el castigo (X) que el tasador fijo A MANO en
    PLANT, pisando la formula. Se leen del libro SIN calcular (`formulas`):
    en el calculado no se distingue un 0,7 escrito de uno que dio la formula.
    La fila 37+k de PLANT es la plantacion de la fila 5+k de PLANTACION."""
    tabla = mapa.tablas.get("plantaciones_sin_royalty")
    if not plantaciones or tabla is None or "PLANT" not in formulas.sheetnames:
        return
    plant = formulas["PLANT"]
    hoja = formulas[mapa.nombre_hoja(tabla.hoja)]
    indice = 0
    for k in range(tabla.capacidad or 6):
        fila = {campo: hoja["{}{}".format(col, tabla.primera_fila + k)].value
                for campo, col in tabla.columnas.items()}
        if _fila_vacia(fila):
            continue
        if indice >= len(plantaciones):
            break
        for columna, campo in (("T", "opex_anios"), ("X", "castigo_pct")):
            valor = plant["{}{}".format(columna, 37 + k)].value
            if isinstance(valor, (int, float)) and not isinstance(valor, bool):
                plantaciones[indice][campo] = valor
        # Conduccion y riego solo viven en PLANT (no en la tabla de
        # PLANTACION): sin esto el Paso 1 los perdia.
        for columna, campo in (("E", "conduccion"), ("M", "riego")):
            valor = plant["{}{}".format(columna, 37 + k)].value
            if isinstance(valor, str) and valor.strip() and not valor.startswith("="):
                plantaciones[indice].setdefault(campo, valor.strip())
        indice += 1


def _leer_aguas(ws, tabla) -> List[Dict[str, Any]]:
    """AGUA!4.1 (inscripcion) y 4.2 (valor) se leen juntas: la fila i de una
    corresponde a la fila i de la otra (mismo orden en que las escribe
    MotorRabobank._escribir_aguas)."""
    cap_insc = (tabla.ultima_fila_inscripcion_plantilla
               - tabla.primera_fila_inscripcion + 1)
    cap_valor = (tabla.ultima_fila_valor_plantilla - tabla.primera_fila_valor + 1)
    capacidad = max(cap_insc, cap_valor)

    filas = []
    for i in range(capacidad):
        derecho: Dict[str, Any] = {}
        if i < cap_insc:
            f = tabla.primera_fila_inscripcion + i
            for campo, col in tabla.columnas_inscripcion.items():
                valor = ws["{}{}".format(col, f)].value
                if valor is not None:
                    _poner_atributo(derecho, campo, valor)
        if i < cap_valor:
            f = tabla.primera_fila_valor + i
            for campo, col in tabla.columnas_valor.items():
                valor = ws["{}{}".format(col, f)].value
                if valor is not None:
                    _poner_atributo(derecho, campo, valor)
        if derecho:
            filas.append(derecho)
    return filas


def leer_retasacion_excel(
    ruta: Union[str, Path], mapa: MapaSantander,
) -> Tuple[Dict[str, Any], List[str]]:
    """Devuelve (datos, avisos). ``datos`` es un dict suelto -no validado
    contra RetasacionRabobank todavia-, para que la pantalla lo use como
    borrador editable: un campo que el archivo antiguo no traia simplemente
    no aparece, y el tasador lo llena.

    ``data_only=True`` porque se lee un archivo YA CALCULADO por Excel -es
    la unica vez que este proyecto usa esa opcion para algo que despues no
    se vuelve a guardar (CLAUDE.md §12: cargar con data_only y guardar
    destruye las formulas; aca nunca se guarda, solo se lee).
    """
    avisos: List[str] = []
    wb = load_workbook(ruta, data_only=True)
    try:
        hojas_presentes = set(wb.sheetnames)
        datos: Dict[str, Any] = {}

        for ruta_campo, spec in mapa.campos.items():
            nombre_hoja = mapa.nombre_hoja(spec.hoja)
            if nombre_hoja not in hojas_presentes:
                continue
            valor = wb[nombre_hoja][spec.celda].value
            if valor is not None:
                _poner_atributo(datos, ruta_campo, valor)

        for nombre_tabla, tabla in mapa.tablas.items():
            atributo = getattr(tabla, "atributo", None)
            if not atributo:
                continue
            nombre_hoja = mapa.nombre_hoja(tabla.hoja)
            if nombre_hoja not in hojas_presentes:
                avisos.append(
                    "'{}': la hoja '{}' no existe en el archivo subido -se "
                    "omite esa tabla".format(nombre_tabla, tabla.hoja))
                continue
            ws = wb[nombre_hoja]
            if nombre_tabla == "aguas":
                filas = _leer_aguas(ws, tabla)
            elif getattr(tabla, "base_de_referencias", False):
                filas = _leer_referencias(ws, tabla)
            else:
                filas = _leer_tabla_simple(ws, tabla)
            if filas:
                datos[atributo] = filas

        # El libro SIN calcular, para lo que depende de si una celda es
        # formula o dato escrito. En el calculado (data_only) una formula que
        # Excel todavia no calculo -un archivo recien generado por el motor,
        # que nadie abrio- vale None, igual que una celda vacia.
        if hasattr(ruta, "seek"):
            ruta.seek(0)
        formulas = load_workbook(ruta, read_only=True)
        try:
            # El criterio de valor: si la linea del resumen esta en blanco,
            # el tasador no considero eso en el valor de la propiedad.
            for opcion, celdas in (getattr(mapa, "criterio_de_valor", None) or {}).items():
                if not isinstance(celdas, list):
                    continue
                resumen = next((c for c in celdas if c["valor"] is None), None)
                if resumen and mapa.nombre_hoja(resumen["hoja"]) in hojas_presentes:
                    celda = formulas[mapa.nombre_hoja(resumen["hoja"])][resumen["celda"]].value
                    datos[opcion] = celda not in (None, "")
            _leer_ajustes_plant(formulas, datos.get("plantaciones_sin_royalty"), mapa)
        finally:
            formulas.close()
            if hasattr(ruta, "seek"):
                ruta.seek(0)

        deslindes = getattr(mapa, "deslindes", None)
        if deslindes and mapa.nombre_hoja(deslindes["hoja"]) in hojas_presentes:
            for ruta_campo, texto in _leer_deslindes(
                    wb[mapa.nombre_hoja(deslindes["hoja"])], deslindes).items():
                _poner_atributo(datos, ruta_campo, texto)

        # Los cuadros de texto no son celdas: se leen del dibujo de la hoja.
        cuadros = {ruta_campo: (mapa.nombre_hoja(spec["hoja"]), spec["forma"])
                   for ruta_campo, spec in (getattr(mapa, "cuadros_de_texto", None) or {}).items()
                   if isinstance(spec, dict)}
        try:
            leidos = leer_textos(ruta, cuadros.values())
        except Exception as error:  # noqa: BLE001 -un dibujo raro no tumba la lectura
            leidos = {}
            avisos.append("no se pudieron leer los cuadros de texto: {}".format(error))
        for ruta_campo, hueco in cuadros.items():
            if hueco not in leidos:
                continue
            if ruta_campo == "diferencias":
                _poner_atributo(datos, "diferencias.plantaciones", leidos[hueco])
            else:
                _poner_atributo(datos, ruta_campo, leidos[hueco])

        if not datos:
            avisos.append(
                "no se reconoció ningún dato -puede ser un formato de "
                "Rabobank distinto al vigente (coordenadas de celda "
                "distintas): revise y complete a mano")

        return _normalizar_booleanos(datos), avisos
    finally:
        wb.close()
