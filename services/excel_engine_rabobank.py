"""Motor de escritura de retasaciones Rabobank/RaboFinance.

Paralelo a `services/excel_engine.py:MotorExcel`, a proposito: ese motor esta
tipado alrededor de `InformeFitogreen` (flujo proyectado, EBITDA, curvas) y
forzarlo a servir tambien hipoteca/prenda seria exactamente lo que
CLAUDE.md ya advertia evitar ("no mezclar los motores").

Lo que SI se reusa, porque ya es generico (no conoce a Bsan ni a Rabobank):
- `services/mapping.py:cargar_mapa` / `MapaSantander` -el nombre es
  historico, la clase no tiene nada especifico de Santander.
- `services/excel_ops.py:escribir/insertar_filas/celda_ancla` -las mismas
  primitivas que ya usa Bsan.

**El archivo sale IDENTICO al formato del banco** (2026-10-03). el usuario:
*"ese formato, identico a ese"*. openpyxl escribe las celdas, pero al
guardar pierde los cuadros de texto, los grupos y las formas del formato real
-15 cuadros de texto en el archivo de RaboFinance-. Despues de guardar,
`services/dibujos_xlsx.py:injertar_dibujos` vuelve a poner el dibujo
completo de cada hoja, tal cual lo trae la plantilla, con los textos del
informe adentro. Verificado contra el archivo real: abierto con openpyxl,
guardado e injertado, se ve pixel por pixel igual en sus 21 paginas.

Alcance de esta version: **una sola propiedad por retasacion** (medido
contra el archivo real: RESUMEN VALOR no repite su bloque para mas de un
predio).

El valor unitario de las plantaciones SIN royalty se calcula escribiendo
la fila espejo en PLANT (costo de reposicion depreciado por edad), con las
formulas VLOOKUP medidas celda por celda contra Tabla12/Tabla5 -ver
`_escribir_plant`-. Una correccion real sobre la plantilla: la edad se
calculaba con `2026` escrito a mano (`=2026-F37`); el motor la escribe
contra `YEAR(PORTADA!$H$7)`, la fecha de la retasacion.

Las plantaciones CON royalty no pasan por PLANT (medido: la plantilla real
no trae formula en esas celdas) y llevan su valor unitario como dato
directo del dominio.
"""

from __future__ import annotations

import asyncio
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.utils.cell import coordinate_from_string

from services import excel_ops
from services.dibujos_xlsx import injertar_dibujos
from services.mapping import MapaSantander


@dataclass
class ResultadoGeneracionRabobank:
    ruta: Path
    celdas_escritas: int = 0
    filas_insertadas: int = 0
    avisos: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.ruta.exists()


def _resolver_atributo(objeto: Any, ruta: str) -> Any:
    """'predio.inscripcion.fojas' -> objeto.predio.inscripcion.fojas.

    Si algun tramo del camino es None (ej. sin inscripcion declarada), el
    dato simplemente no esta -no es un error, es un campo que el tasador
    no lleno todavia-.
    """
    actual = objeto
    for tramo in ruta.split("."):
        if actual is None:
            return None
        actual = getattr(actual, tramo, None)
    return actual


#: Como escribe la plantilla un si/no cuando el mapping no dice otra cosa.
_SI_NO = ("Sí", "No")


def _sin_tildes(texto: str) -> str:
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", texto)
                   if unicodedata.category(c) != "Mn")


def _clave_de_lista(texto: Optional[str], validas: Sequence[str]) -> Optional[str]:
    """La clave de `validas` que corresponde a `texto`, sin importar
    mayusculas, tildes ni un plural: 'Vid vinífera' -> 'Vid Vinifera',
    'Cerezos' -> 'Cerezo'. None si ninguna calza."""
    if not texto:
        return None

    def normal(t):
        return " ".join(_sin_tildes(str(t)).lower().split())

    buscado = normal(texto)
    candidatos = {normal(v): v for v in validas}
    if buscado in candidatos:
        return candidatos[buscado]
    for sufijo in ("es", "s"):
        if buscado.endswith(sufijo) and buscado[: -len(sufijo)] in candidatos:
            return candidatos[buscado[: -len(sufijo)]]
    return None


def _valor_celda(valor: Any, si_no: Sequence[str]) -> Any:
    """Un booleano se escribe con la ortografia de la lista desplegable de
    esa columna ('Sí'/'No', o 'SÍ'/'NO' en AGUA): un 'Si' sin tilde no calza
    con la lista y Excel lo marca como invalido."""
    if isinstance(valor, bool):
        return si_no[0] if valor else si_no[1]
    if hasattr(valor, "value"):  # Enum
        return valor.value
    return valor


class MotorRabobank:
    def __init__(self, ruta_plantilla: Union[str, Path], mapa: MapaSantander):
        self.ruta_plantilla = Path(ruta_plantilla)
        self.mapa = mapa

    @property
    def _si_no(self) -> Sequence[str]:
        return getattr(self.mapa, "si_no", None) or _SI_NO

    def generar(
        self, retasacion, destino: Union[str, Path]
    ) -> ResultadoGeneracionRabobank:
        destino = Path(destino)
        # rich_text: la plantilla tiene celdas con formato dentro del texto
        # (los "m²" con el 2 en superindice de los encabezados, "Norte:" en
        # negrita). Sin esto openpyxl lo aplana a texto simple al guardar.
        wb = load_workbook(self.ruta_plantilla, rich_text=True)
        resultado = ResultadoGeneracionRabobank(ruta=destino)
        inserciones: Dict[str, List[Tuple[int, int]]] = {}

        self._escribir_campos(wb, retasacion, resultado)
        self._aplicar_criterio_de_valor(wb, retasacion, resultado)
        self._escribir_deslindes(wb, retasacion, resultado)
        self._escribir_factores_liquidacion(wb, retasacion, resultado)
        self._escribir_predio_en_suelos(wb, retasacion, resultado)

        self._escribir_tabla(wb, "suelos", retasacion.suelos, resultado, inserciones)
        self._escribir_aguas(wb, retasacion.aguas, resultado)
        self._escribir_tabla(wb, "construcciones_hipotecables",
                             retasacion.construcciones_hipotecables, resultado, inserciones)
        self._escribir_tabla(wb, "construcciones_prendables",
                             retasacion.construcciones_prendables, resultado, inserciones)
        self._escribir_tabla(wb, "instalaciones", retasacion.instalaciones,
                             resultado, inserciones)
        self._escribir_tabla(wb, "plantaciones_sin_royalty",
                             retasacion.plantaciones_sin_royalty, resultado, inserciones)
        self._escribir_plant(wb, retasacion, resultado)
        self._escribir_tabla(wb, "plantaciones_con_royalty",
                             retasacion.plantaciones_con_royalty, resultado, inserciones)
        self._escribir_referencias(wb, "testigos_propiedades",
                                   retasacion.testigos_propiedades, resultado, inserciones)
        self._escribir_referencias(wb, "testigos_derechos_agua",
                                   retasacion.testigos_derechos_agua, resultado, inserciones)

        destino.parent.mkdir(parents=True, exist_ok=True)
        wb.save(destino)

        resultado.avisos.extend(injertar_dibujos(
            self.ruta_plantilla, destino,
            textos=self._textos(retasacion),
            filas_insertadas=inserciones,
        ))
        return resultado

    async def generar_async(
        self, retasacion, destino: Union[str, Path]
    ) -> ResultadoGeneracionRabobank:
        """Version para FastAPI -mismo motivo que MotorExcel.generar_async:
        openpyxl es sincrono y consume CPU, no conviene bloquear el event loop."""
        return await asyncio.to_thread(self.generar, retasacion, destino)

    # ------------------------------------------------------------ campos

    def _escribir_celda(self, ws, coordenada: str, valor: Any, resultado,
                        si_no: Optional[Sequence[str]] = None,
                        formato: Optional[str] = None) -> None:
        valor = _valor_celda(valor, si_no or self._si_no)
        escrito, aviso = excel_ops.escribir(ws, coordenada, valor, formato)
        if escrito:
            resultado.celdas_escritas += 1
        if aviso:
            resultado.avisos.append(aviso)

    def _escribir_campos(self, wb, retasacion, resultado) -> None:
        for ruta, spec in self.mapa.campos.items():
            valor = _resolver_atributo(retasacion, ruta)
            respaldo = getattr(spec, "si_falta", None)
            if valor is None and respaldo:
                valor = _resolver_atributo(retasacion, respaldo)
            if valor is None:
                continue
            ws = wb[self.mapa.nombre_hoja(spec.hoja)]
            self._escribir_celda(ws, spec.celda, valor, resultado, formato=spec.formato)

    def _aplicar_criterio_de_valor(self, wb, retasacion, resultado) -> None:
        """Si el tasador NO considera las plantaciones o las instalaciones
        en el valor de la propiedad, sus lineas del resumen y de la portada
        quedan como en el formato real de ese caso (declarado en el mapping).
        Se pisan formulas de la plantilla a proposito: es la unica forma de
        sacarlas del total sin tocar el total."""
        criterio = getattr(self.mapa, "criterio_de_valor", None) or {}
        for opcion, celdas in criterio.items():
            if not isinstance(celdas, list) or getattr(retasacion, opcion, True):
                continue
            for spec in celdas:
                ws = wb[self.mapa.nombre_hoja(spec["hoja"])]
                ws[spec["celda"]].value = spec["valor"]
                resultado.celdas_escritas += 1

    def _escribir_deslindes(self, wb, retasacion, resultado) -> None:
        """LEGAL!C35: los cuatro deslindes en una sola celda, un renglon por
        punto cardinal y su nombre en negrita, como en el formato real."""
        spec = getattr(self.mapa, "deslindes", None)
        if not spec:
            return
        partes = []
        for rotulo, ruta in spec["partes"].items():
            texto = _resolver_atributo(retasacion, ruta)
            if texto:
                partes.append((rotulo, str(texto).strip()))
        if not partes:
            return
        normal = InlineFont(rFont="Calibri", sz=10)
        negrita = InlineFont(rFont="Calibri", sz=10, b=True)
        bloques: List[Any] = []
        for i, (rotulo, texto) in enumerate(partes):
            bloques.append(TextBlock(negrita, rotulo))
            fin = "" if i == len(partes) - 1 else "\n"
            bloques.append(TextBlock(normal, ": {}{}".format(texto.lstrip(":").strip(), fin)))
        ws = wb[self.mapa.nombre_hoja(spec["hoja"])]
        celda = ws[excel_ops.celda_ancla(ws, spec["celda"])]
        if celda.value in (None, ""):
            celda.value = CellRichText(*bloques)
            resultado.celdas_escritas += 1

    def _escribir_factores_liquidacion(self, wb, retasacion, resultado) -> None:
        factores = getattr(self.mapa, "factores_liquidacion", None)
        if not factores:
            return
        for clave, valor in (
            ("hipoteca", retasacion.factor_liquidacion_hipoteca),
            ("prenda", retasacion.factor_liquidacion_prenda),
        ):
            grupo = factores.get(clave)
            if not grupo:
                continue
            ws = wb[self.mapa.nombre_hoja(grupo["hoja"])]
            for celda in grupo["celdas"]:
                self._escribir_celda(ws, celda, valor, resultado)

    def _escribir_predio_en_suelos(self, wb, retasacion, resultado) -> None:
        """La columna A/B de SUELOS!3.2 (Propiedad/Rol) es la MISMA en las
        capacidad filas -una sola propiedad en esta fase-, no viene del
        modelo por fila."""
        tabla = self.mapa.tablas.get("suelos")
        if not tabla:
            return
        ws = wb[self.mapa.nombre_hoja(tabla.hoja)]
        for i, suelo in enumerate(retasacion.suelos):
            fila = tabla.primera_fila + i
            # La columna Propiedad es libre en el formato ("Afectación a
            # utilidad pública"): si la fila no la trae, el nombre del predio.
            for col, valor in (("A", getattr(suelo, "nombre", None) or retasacion.predio.nombre),
                              ("B", retasacion.predio.rol_avaluo)):
                self._escribir_celda(ws, "{}{}".format(col, fila), valor, resultado)

    # ------------------------------------------------------------- tablas

    def _escribir_tabla(self, wb, nombre_tabla: str, filas: List[Any], resultado,
                        inserciones: Dict[str, List[Tuple[int, int]]]) -> None:
        tabla = self.mapa.tablas.get(nombre_tabla)
        if tabla is None or not filas:
            return
        ws = wb[self.mapa.nombre_hoja(tabla.hoja)]
        capacidad = tabla.capacidad or len(filas)
        si_no = getattr(tabla, "si_no", None)

        if len(filas) > capacidad:
            faltan = len(filas) - capacidad
            if getattr(tabla, "insertar_filas", False):
                # Insertar en la PRIMERA fila corre entero todo rango que empieza
                # ahi: el total (=SUM(E22:E22)) terminaba sumando solo la ultima
                # instalacion. Con varias filas de plantilla se inserta en la
                # ultima (el rango crece solo, igual que en Santander); con una
                # sola, debajo de ella, estirando los rangos que terminan ahi.
                if capacidad >= 2 and tabla.ultima_fila_plantilla:
                    desde, extender = tabla.ultima_fila_plantilla, None
                else:
                    desde, extender = tabla.primera_fila + 1, tabla.primera_fila
                excel_ops.insertar_filas(ws, desde, faltan, tabla.primera_fila,
                                         extender_fin=extender)
                resultado.filas_insertadas += faltan
                inserciones.setdefault(ws.title, []).append((desde, faltan))
            else:
                resultado.avisos.append(
                    "{}: {} filas no caben en la capacidad de la plantilla "
                    "({}) y esta tabla no admite insertar filas -se "
                    "escribieron solo las primeras {}".format(
                        nombre_tabla, len(filas), capacidad, capacidad))
                filas = filas[:capacidad]

        for i, item in enumerate(filas):
            fila = tabla.primera_fila + i
            for campo, columna in tabla.columnas.items():
                valor = getattr(item, campo, None)
                if valor is None:
                    continue
                self._escribir_celda(ws, "{}{}".format(columna, fila), valor,
                                     resultado, si_no=si_no)

    def _escribir_aguas(self, wb, aguas: List[Any], resultado) -> None:
        """AGUA!4.1 (inscripcion) y 4.2 (valor): la MISMA lista de derechos,
        escrita dos veces con dos juegos de columnas distintos -asi separa
        la plantilla del banco la informacion legal de la de valor."""
        tabla = self.mapa.tablas.get("aguas")
        if tabla is None or not aguas:
            return
        ws = wb[self.mapa.nombre_hoja(tabla.hoja)]
        si_no = getattr(tabla, "si_no", None)
        cap_insc = tabla.ultima_fila_inscripcion_plantilla - tabla.primera_fila_inscripcion + 1
        cap_valor = tabla.ultima_fila_valor_plantilla - tabla.primera_fila_valor + 1
        if len(aguas) > cap_insc or len(aguas) > cap_valor:
            resultado.avisos.append(
                "aguas: {} derechos no caben en la capacidad de la plantilla "
                "(inscripcion {}, valor {}) -se escribieron solo los "
                "primeros".format(len(aguas), cap_insc, cap_valor))

        for base_fila, columnas, cap in (
            (tabla.primera_fila_inscripcion, tabla.columnas_inscripcion, cap_insc),
            (tabla.primera_fila_valor, tabla.columnas_valor, cap_valor),
        ):
            for i, derecho in enumerate(aguas[:cap]):
                fila = base_fila + i
                for campo, columna in columnas.items():
                    valor = _resolver_atributo(derecho, campo)
                    if campo == "fuente_inscripcion" and valor is None:
                        valor = derecho.fuente
                    if campo == "derechos_en_acciones" and valor is None and \
                            str(derecho.tipo or "").lower().startswith("acci"):
                        valor = derecho.derechos_cantidad
                    if campo == "caudal_nominal_ls" and derecho.equivalencia_ls_accion:
                        # Como en el formato real (=D25*1.5): la equivalencia
                        # queda a la vista y el caudal sigue a las acciones.
                        valor = "=D{}*{}".format(fila, repr(float(derecho.equivalencia_ls_accion)))
                    if valor is None:
                        continue
                    self._escribir_celda(ws, "{}{}".format(columna, fila), valor,
                                         resultado, si_no=si_no)

    def _escribir_referencias(self, wb, nombre_tabla: str, filas: List[Any], resultado,
                              inserciones: Dict[str, List[Tuple[int, int]]]) -> None:
        """REF. PROPIEDADES / REF. Canal: la hoja es la BASE de referencias
        que la tasadora va juntando, y en el informe quedan visibles solo las
        que usa -la estadistica (SUBTOTAL, AGREGAR) ignora lo oculto-. La
        plantilla trae la base oculta; aca se insertan las referencias del
        informe DENTRO del rango de la base, para que todas las formulas de
        la estadistica las cubran, con las formulas de una fila de la base
        (E, I, L, M en propiedades; J, K, L en canal). La asimetria y la
        curtosis nombran sus celdas una por una: se apuntan a estas filas."""
        tabla = self.mapa.tablas.get(nombre_tabla)
        if tabla is None or not filas:
            return
        ws = wb[self.mapa.nombre_hoja(tabla.hoja)]
        desde = tabla.insertar_en
        cantidad = len(filas)
        excel_ops.insertar_filas(ws, desde, cantidad, tabla.fila_modelo)
        resultado.filas_insertadas += cantidad
        inserciones.setdefault(ws.title, []).append((desde, cantidad))

        for i, item in enumerate(filas):
            fila = desde + i
            ws.row_dimensions[fila].hidden = False
            for campo, columna in tabla.columnas.items():
                valor = getattr(item, campo, None)
                if valor is None:
                    continue
                self._escribir_celda(ws, "{}{}".format(columna, fila), valor, resultado)
            for columna in getattr(tabla, "ceros", None) or []:
                ws["{}{}".format(columna, fila)].value = 0

        rango = "{{col}}{}:{{col}}{}".format(desde, desde + cantidad - 1)
        for clave, funcion in (("asimetria", "SKEW"), ("curtosis", "KURT")):
            spec = getattr(tabla, clave, None)
            if not spec:
                continue
            columna, numero = coordinate_from_string(spec["celda"])
            if numero >= desde:
                numero += cantidad
            # IFERROR: con menos de 3 referencias (4 para la curtosis) Excel no
            # puede calcularlas y dejaria un #¡DIV/0! en el informe.
            ws["{}{}".format(columna, numero)].value = '=IFERROR({}({}),"")'.format(
                funcion, rango.format(col=spec["columna"]))

    # ------------------------------------------------------- cuadros de texto

    def _textos(self, retasacion) -> Dict[Tuple[str, str], str]:
        cuadros = getattr(self.mapa, "cuadros_de_texto", None) or {}
        textos: Dict[Tuple[str, str], str] = {}
        for ruta, spec in cuadros.items():
            if not isinstance(spec, dict):
                continue  # la "nota"
            if ruta == "diferencias":
                texto = self._diferencias(retasacion)
            else:
                texto = _resolver_atributo(retasacion, ruta)
            if texto:
                textos[(self.mapa.nombre_hoja(spec["hoja"]), spec["forma"])] = str(texto)
        return textos

    @staticmethod
    def _diferencias(retasacion) -> Optional[str]:
        """El formato tiene UN cuadro de "Diferencias con tasacion anterior".
        Si solo se escribio el de plantaciones -el caso del formato- va tal
        cual; si hay de otras secciones, van todas, cada una con su seccion
        delante: antes se perdian en silencio."""
        d = retasacion.diferencias
        secciones = [("Suelos", d.suelos), ("Derechos de agua", d.aguas),
                     ("Construcciones", d.construcciones),
                     ("Instalaciones", d.instalaciones), ("Plantaciones", d.plantaciones)]
        llenas = [(n, t.strip()) for n, t in secciones if t and t.strip()]
        if not llenas:
            return None
        if len(llenas) == 1 and llenas[0][0] == "Plantaciones":
            return llenas[0][1]
        return "\n".join("{}: {}".format(n, t) for n, t in llenas)

    # -------------------------------------------------------------- PLANT

    #: Donde arranca el bloque de plantas en PLANT en la plantilla real
    #: (medido: filas 37-42 en el archivo de referencia, una por cada fila
    #: de PLANTACION!6.1 -misma capacidad, 6-).
    _PLANT_PRIMERA_FILA = 37

    def _escribir_plant(self, wb, retasacion, resultado) -> None:
        """Escribe la fila espejo en PLANT para cada plantacion SIN royalty
        -costo de reposicion depreciado por edad, via VLOOKUP contra
        Tabla12 (especie) y Tabla5 (estructura), medido formula por formula
        contra el archivo real (fila 37). PLANTACION!G ya trae
        `=PLANT!AA{fila}` en la plantilla, asi que no hace falta tocarlo:
        basta con completar PLANT para que ese VLOOKUP deje de dar 0.
        """
        filas = retasacion.plantaciones_sin_royalty
        if not filas:
            return
        pt = wb["PLANT"]

        validacion = getattr(self.mapa, "plant_validacion", None) or {}
        especies_validas = validacion.get("especies", [])
        estructuras_validas = validacion.get("estructuras", [])

        for i, planta in enumerate(filas):
            f = self._PLANT_PRIMERA_FILA + i
            # Una plantacion a medio llenar llega con huecos (2026-10-02,
            # services/retasacion_pendientes.py): la especie que falta ya
            # sale en la lista de pendientes, aqui solo no hay que caerse.
            nombre = planta.variedad or planta.especie or "fila {}".format(i + 1)

            # El formato real escribe la especie para el lector en PLANTACION
            # ("Vid vinífera", "Cerezos") y la CLAVE de la tabla de costos en
            # PLANT ("Vid Vinifera", "Cerezo"): son textos distintos, y el
            # VLOOKUP no perdona una tilde ni un plural -da 0 en silencio-.
            especie = _clave_de_lista(planta.especie, especies_validas)
            estructura = _clave_de_lista(planta.estructura, estructuras_validas)
            if especies_validas and planta.especie and especie is None:
                resultado.avisos.append(
                    "plantacion '{}': la especie '{}' no está en la biblioteca "
                    "de costos de PLANT -el valor unitario va a salir en 0, no "
                    "en error. Especies válidas: {}".format(
                        nombre, planta.especie, ", ".join(especies_validas)))
            if estructuras_validas and planta.estructura and estructura is None:
                resultado.avisos.append(
                    "plantacion '{}': la estructura '{}' no está en "
                    "PLANT!Tabla5 -el costo de estructura va a salir en "
                    "0".format(nombre, planta.estructura))

            valores = {
                # El formato real usa el nombre corto ("Parcela n° 2"): el
                # legal completo no cabe en la columna.
                "A": retasacion.predio.nombre_corto or retasacion.predio.nombre,
                "B": planta.rol_avaluo or retasacion.predio.rol_avaluo,
                "C": especie or planta.especie,
                "D": planta.variedad,
                "E": planta.conduccion,
                "F": planta.anio_plantacion,
                "I": planta.densidad_plantas_ha,
                "J": planta.superficie_ha,
                "L": estructura or planta.estructura,
                "M": planta.riego,
            }
            for columna, valor in valores.items():
                if valor is None:
                    continue
                pt["{}{}".format(columna, f)] = valor
                resultado.celdas_escritas += 1

            # Formulas medidas contra la fila 37 del archivo real, con la
            # UNICA correccion real encontrada: la edad usaba '2026' escrito
            # a mano en vez de la fecha de la retasacion.
            formulas = {
                "K": "=+J{f}",
                "N": "=IFERROR(VLOOKUP(C{f},Tabla12[],3,FALSE),)",
                "O": "=IFERROR(VLOOKUP(C{f},Tabla12[#All],5,FALSE),)",
                "P": "=IFERROR(VLOOKUP(C{f},Tabla12[#All],6,FALSE),)",
                "Q": "=IFERROR(VLOOKUP(L{f},Tabla5[#All],2,FALSE),)",
                "R": "=IF(F{f}>0,YEAR(PORTADA!$H$7)-F{f},0)",
                "S": "=IFERROR(VLOOKUP(C{f},Tabla12[#All],4,FALSE),)",
                "T": "=IF(S{f}>R{f},R{f},S{f})",
                "U": "=IFERROR(VLOOKUP(C{f},Tabla12[#All],2,FALSE),)",
                "V": "=U{f}*T{f}",
                "W": '=IFERROR(ROUND((N{f}*I{f}+O{f}*I{f}+Q{f}+V{f}+P{f}),0),"")',
                "X": "=IF(R{f}<20,0%,IF(R{f}<30,20%,IF(R{f}<40,30%,"
                     "IF(R{f}<50,40%,IF(R{f}>50,50%,0)))))",
                "Y": "=W{f}*(1-X{f})",
                "Z": "=Y{f}*J{f}",
                "AA": "=ROUND(Y{f}*$D$1,-5)",
                "AB": "=AA{f}*J{f}",
                "AC": "=ROUND(W{f}*$D$1,-5)",
            }
            # Lo que el tasador fija a mano pisa la formula, como en el
            # formato real (T38 = 0, X38 = 0,7 para un Merlot de 1980).
            if planta.opex_anios is not None:
                formulas["T"] = planta.opex_anios
            if planta.castigo_pct is not None:
                formulas["X"] = planta.castigo_pct
            for columna, plantilla in formulas.items():
                pt["{}{}".format(columna, f)] = (
                    plantilla.format(f=f) if isinstance(plantilla, str) else plantilla)
                resultado.celdas_escritas += 1
