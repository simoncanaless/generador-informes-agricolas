"""Captura de la visita a terreno: del formulario al modelo del informe.

Es el frente de los datos de ORIGEN A, los que solo existen porque alguien fue
al predio: maquinaria, comercializacion, resultados que declara el cliente,
inversiones, opinion profesional. Medido con ``scripts/medir_cobertura.py``,
son unas 240 celdas del informe Santander y hoy el tasador las teclea dentro
del Excel, cazandolas entre siete hojas.

Dos decisiones de diseno:

- **El formulario se declara en datos**, en ``config/formulario_visita.json``,
  no en el HTML ni aqui. el agrónomo va a mandar una lista de las preguntas mas
  recurrentes de sus entrevistas; con esto, incorporarlas es agregar entradas
  al JSON.
- **No impone un orden.** La entrevista no es pauteada -cada agronomo la lleva
  a su manera-, asi que las secciones siguen la estructura del INFORME y se
  llenan en cualquier orden, guardando a medias las veces que haga falta.

Lo capturado se guarda tal cual en ``expediente.datos['visita']``; este modulo
lo traduce al ``InformeFitogreen`` en el momento de generar.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from schemas.comun import sin_signo_de_formula

__all__ = [
    "DefinicionFormulario", "cargar_formulario", "aplicar_visita",
    "resumen_de_avance", "ajustes_por_curva",
]


class DefinicionFormulario:
    """El formulario tal como se declaro, con sus escalas resueltas."""

    def __init__(self, crudo: Dict[str, Any]):
        self.version = crudo.get("version")
        self.escalas: Dict[str, List[str]] = crudo.get("escalas", {})
        self.secciones: List[Dict[str, Any]] = crudo.get("secciones", [])

    def para_navegador(self) -> Dict[str, Any]:
        """Secciones con las opciones ya expandidas, listas para pintar."""
        secciones = []
        for seccion in self.secciones:
            campos = []
            for campo in seccion.get("campos", []):
                copia = dict(campo)
                escala = copia.pop("escala", None)
                if escala:
                    copia["opciones"] = list(self.escalas.get(escala, []))
                campos.append(copia)
            copia_seccion = dict(seccion)
            copia_seccion["campos"] = campos
            secciones.append(copia_seccion)
        return {"version": self.version, "secciones": secciones}

    def seccion(self, id_seccion: str) -> Optional[Dict[str, Any]]:
        for seccion in self.secciones:
            if seccion.get("id") == id_seccion:
                return seccion
        return None


def cargar_formulario(ruta: Union[str, Path]) -> DefinicionFormulario:
    with Path(ruta).open(encoding="utf-8") as fh:
        return DefinicionFormulario(json.load(fh))


def _vacio(valor: Any) -> bool:
    return valor is None or (isinstance(valor, str) and not valor.strip())


def _asignar(destino: Any, ruta: str, valor: Any) -> bool:
    """Escribe ``valor`` en una ruta con puntos sobre el modelo. Devuelve si escribio.

    Camina la ruta creando nada: todos los contenedores intermedios del informe
    existen desde su construccion (``DotacionRecursos()``, ``Anexo()``...), asi
    que si un tramo no existe es que la ruta esta mal escrita y hay que verlo,
    no taparlo.
    """
    partes = ruta.split(".")
    actual = destino
    for parte in partes[:-1]:
        actual = getattr(actual, parte, None)
        if actual is None:
            return False

    ultimo = partes[-1]

    # Algunos destinos no son un atributo sino UNA CLAVE dentro de un
    # diccionario: el tipo de cambio vive en
    # ``flujo_proyectado.tipo_cambio_pesos_usd`` indexado por temporada, y sin
    # esto no habia forma de capturarlo desde el navegador. Sin tipo de cambio
    # declarado, las dos columnas USD del Anexo quedan vacias para siempre.
    if ultimo.endswith("]") and "[" in ultimo:
        nombre, _, clave = ultimo.partition("[")
        clave = clave[:-1]
        contenedor = getattr(actual, nombre, None)
        # Una LISTA con indice: las ventas del F29 son doce meses en orden
        # (julio a junio), cada uno un campo del formulario.
        if isinstance(contenedor, list) and clave.isdigit():
            indice = int(clave)
            while len(contenedor) <= indice:
                contenedor.append(None)
            contenedor[indice] = sin_signo_de_formula(valor)
            return True
        if not isinstance(contenedor, dict):
            return False
        contenedor[clave] = sin_signo_de_formula(valor)
        return True

    if not hasattr(actual, ultimo):
        return False
    # setattr no pasa por los validadores del esquema: el que quita el "="
    # de adelante se aplica aca a mano (ver schemas/comun.py).
    setattr(actual, ultimo, sin_signo_de_formula(valor))
    return True


def _fila_suelos(
    nombre: str,
    i_ii_riego_ha: Optional[float] = None, i_ii_secano_ha: Optional[float] = None,
    iii_iv_riego_ha: Optional[float] = None, iii_iv_secano_ha: Optional[float] = None,
    v_vi_riego_ha: Optional[float] = None, v_vi_secano_ha: Optional[float] = None,
    vii_viii_riego_ha: Optional[float] = None, vii_viii_secano_ha: Optional[float] = None,
) -> Any:
    """Arma un ``PredioSuelos`` desde las columnas planas de la visita.

    El modelo real anida la superficie en un diccionario por ``ClaseUso``
    -asi lo recorre `services/excel_engine.py` para escribir la tabla-, pero
    una fila con las 4 clases en columnas (una por predio, como en la
    plantilla del banco) es lo que se puede llenar a mano o importar desde
    la planilla del cliente; un editor de diccionarios no.
    """
    from schemas.informe import ClaseUso, PredioSuelos, SuperficieClaseUso

    pares = {
        ClaseUso.I_II: (i_ii_riego_ha, i_ii_secano_ha),
        ClaseUso.III_IV: (iii_iv_riego_ha, iii_iv_secano_ha),
        ClaseUso.V_VI: (v_vi_riego_ha, v_vi_secano_ha),
        ClaseUso.VII_VIII: (vii_viii_riego_ha, vii_viii_secano_ha),
    }
    superficies = {
        clase: SuperficieClaseUso(riego_ha=riego, secano_ha=secano)
        for clase, (riego, secano) in pares.items()
        if riego is not None or secano is not None
    }
    return PredioSuelos(nombre=nombre, superficies=superficies)


def _fila_sensibilizacion(especie: str, precio_exportacion=None, precio_mercado_interno=None,
                          rendimiento=None, costos=None):
    """(especie, Sensibilizacion) desde una fila de la pantalla, que pide los
    porcentajes como porcentajes (-5 es -5%); el informe los guarda como
    fraccion (-0,05), que es lo que multiplican las formulas."""
    from schemas.informe import Sensibilizacion

    def fraccion(valor):
        return round(float(valor) / 100.0, 6) if valor not in (None, "") else 0.0

    return especie, Sensibilizacion(
        precio_exportacion=fraccion(precio_exportacion),
        precio_mercado_interno=fraccion(precio_mercado_interno),
        rendimiento=fraccion(rendimiento), costos=fraccion(costos))


def aplicar_visita(
    informe: Any, capturado: Dict[str, Any], formulario: DefinicionFormulario
) -> List[str]:
    """Vuelca lo capturado en el informe. Devuelve los avisos de lo que no calzo.

    Lo que viene vacio no se escribe: una visita a medio llenar no debe borrar
    lo que ya habia en el informe -por ejemplo la opinion que redacto la IA a
    partir de las fotos-, solo completarlo o corregirlo.
    """
    from schemas.informe import (
        AjusteDeCurva, DerechoAgua, FilaAnexo, ItemDotacion, LineaInversion, Lote,
        SistemaComercializacion, SistemaProduccion, UnidadPredial,
    )

    constructores = {
        "ajustes_de_curva": AjusteDeCurva,
        "anexo.resultado_real.filas": FilaAnexo,
        "anexo.resultado_preliminar.filas": FilaAnexo,
        "dotacion_recursos.derechos_agua": DerechoAgua,
        "dotacion_recursos.maquinaria": ItemDotacion,
        "dotacion_recursos.infraestructura": ItemDotacion,
        "dotacion_recursos.suelos.predios": _fila_suelos,
        "dotacion_recursos.lotes": Lote,
        "dotacion_recursos.unidades_prediales": UnidadPredial,
        "estructura_productiva.sistemas_produccion": SistemaProduccion,
        "estructura_productiva.sistemas_comercializacion": SistemaComercializacion,
        "inversiones.lineas": LineaInversion,
        "supuestos.sensibilizacion": _fila_sensibilizacion,
    }
    avisos: List[str] = []

    for seccion in formulario.secciones:
        id_seccion = seccion.get("id")
        datos = capturado.get(id_seccion)
        if not datos:
            continue

        if seccion.get("tipo") == "tabla":
            destino = seccion.get("destino")
            constructor = constructores.get(destino)
            if constructor is None:
                avisos.append(
                    "seccion '{}': no hay modelo para '{}'".format(
                        id_seccion, destino))
                continue

            filas = []
            for cruda in datos:
                limpia = {k: v for k, v in (cruda or {}).items()
                          if not _vacio(v)}
                if not limpia:
                    continue   # fila en blanco: el tasador no la lleno
                try:
                    filas.append(constructor(**limpia))
                except Exception as error:            # noqa: BLE001
                    avisos.append(
                        "seccion '{}': una fila no se pudo leer ({})".format(
                            id_seccion, error))
            if filas and destino == "supuestos.sensibilizacion":
                # Una tabla en la pantalla, un diccionario por especie en el
                # informe: la hoja Supuestos lleva una columna por especie.
                filas = {especie: datos_especie for especie, datos_especie in filas}
            if filas and not _asignar(informe, destino, filas):
                avisos.append(
                    "seccion '{}': la ruta '{}' no existe en el informe".format(
                        id_seccion, destino))
        else:
            for campo in seccion.get("campos", []):
                destino = campo.get("destino")
                valor = datos.get(campo.get("id"))
                if not destino or _vacio(valor):
                    continue
                if not _asignar(informe, destino, valor):
                    avisos.append(
                        "campo '{}': la ruta '{}' no existe en el informe".format(
                            campo.get("id"), destino))

    return avisos


def resumen_de_avance(
    capturado: Dict[str, Any], formulario: DefinicionFormulario
) -> List[Dict[str, Any]]:
    """Cuanto lleva llenado cada seccion, para mostrarlo en la pantalla.

    El tasador vuelve a esta pantalla varias veces -consigue un dato, vuelve,
    consigue otro-, y necesita ver de un vistazo que le falta.
    """
    resumen = []
    for seccion in formulario.secciones:
        datos = capturado.get(seccion.get("id"))
        if seccion.get("tipo") == "tabla":
            llenas = sum(
                1 for fila in (datos or [])
                if any(not _vacio(v) for v in (fila or {}).values())
            )
            total = seccion.get("filas", 0)
        else:
            campos = seccion.get("campos", [])
            llenas = sum(
                1 for c in campos if not _vacio((datos or {}).get(c.get("id")))
            )
            total = len(campos)
        resumen.append({
            "id": seccion.get("id"),
            "titulo": seccion.get("titulo"),
            "llenos": llenas,
            "total": total,
            "completo": bool(total) and llenas >= total,
            "vacio": llenas == 0,
        })
    return resumen


def ajustes_por_curva(informe: Any) -> Dict[str, list]:
    """Traduce lo que el tasador observo en terreno a ajustes por curva.

    Devuelve el diccionario que espera
    ``services/envejecimiento.py:proyectar_cartera``, indexado por la especie
    o el nombre de curva que escribio el agronomo.

    Existe para cerrar el ultimo tramo del recorrido: hasta ahora el ajuste
    vivia en la simulacion -que es solo una vista previa- o en un argumento de
    linea de comandos. Ahora sale de la pantalla de visita, que es donde el
    tasador anota lo que vio, y llega al informe sin que nadie lo transcriba.
    """
    from schemas.curvas import AjusteAplicado, CategoriaAjuste

    por_curva: Dict[str, list] = {}
    for ajuste in getattr(informe, "ajustes_de_curva", None) or []:
        if not ajuste.ajuste_pct:
            continue   # un 0% no es un ajuste, es una fila que quedo a medias
        por_curva.setdefault(ajuste.especie.strip(), []).append(AjusteAplicado(
            nombre="Observado en terreno",
            factor_pct=ajuste.ajuste_pct,
            categoria=CategoriaAjuste.OTRO,
            justificacion=ajuste.motivo,
        ))
    return por_curva
