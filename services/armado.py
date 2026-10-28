"""Arma el informe completo con TODO lo que el sistema ya sabe del expediente.

El ultimo tramo que faltaba. Hasta ahora cada pieza funcionaba y se probaba por
separado -el catastro se importaba, la visita se capturaba, las curvas se
cargaban, el riego se calculaba- pero el boton de generar del navegador solo
usaba los datos del formulario, las escrituras y las fotos. Todo lo demas habia
que pasarlo a mano por un script.

Aqui se juntan:

    catastro guardado ---+
    visita capturada ----+--> informe completo --> MotorExcel --> .xlsx
    curvas del catalogo -+
    biblioteca de riego -+

Dos criterios que gobiernan todo el modulo:

- **Nada se inventa.** Si falta la curva de una especie, ese bloque no se
  proyecta y se dice; si falta la comuna, el riego no se calcula y se dice. Un
  hueco visible es mejor que un numero inventado en un informe que va al banco.
- **Lo declarado le gana a lo calculado.** Las columnas Real y Preliminar del
  flujo llevan lo que declara el cliente, nunca la curva.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence

__all__ = ["completar_informe", "ResultadoArmado"]


class ResultadoArmado:
    """Que se pudo armar y que no. Lo segundo es lo que hay que mostrar."""

    def __init__(self) -> None:
        self.avisos: List[str] = []
        self.plantaciones: int = 0
        self.bloques_de_flujo: int = 0
        self.superficie_proyectada_ha: float = 0.0
        self.superficie_sin_curva_ha: float = 0.0
        self.demanda_riego_l_s: Optional[float] = None

    def avisar(self, mensaje: str) -> None:
        if mensaje not in self.avisos:
            self.avisos.append(mensaje)

    def como_dict(self) -> Dict[str, Any]:
        return {
            "plantaciones": self.plantaciones,
            "bloques_de_flujo": self.bloques_de_flujo,
            "superficie_proyectada_ha": self.superficie_proyectada_ha,
            "superficie_sin_curva_ha": self.superficie_sin_curva_ha,
            "demanda_riego_l_s": self.demanda_riego_l_s,
            "avisos": self.avisos,
        }


def _anio_de(temporada: str) -> Optional[int]:
    """'2024/25' -> 2024. None si no se puede leer."""
    try:
        return int(str(temporada).split("/")[0])
    except (ValueError, AttributeError, IndexError):
        return None


def _temporadas_de_proyeccion(anio_base: int, cuantas: int = 8) -> List[str]:
    return ["{}/{}".format(a, str(a + 1)[2:])
            for a in range(anio_base, anio_base + cuantas)]


def _temporadas_declaradas(informe, anio: int) -> Dict[str, str]:
    """Que temporadas trae el Anexo y como se rotula cada una.

    No siempre les entregan los anios hacia atras -a veces dos, a veces uno,
    a veces ninguno-, asi que la temporada la DECLARA el tasador en el Anexo y
    no se supone aqui. Si no la declaro, la tabla no aporta columna.
    """
    anexo = informe.anexo
    salida: Dict[str, str] = {}
    for tabla, clase in ((anexo.resultado_real, "Real "),
                         (anexo.resultado_preliminar, "Preliminar")):
        temporada = (tabla.temporada or "").strip()
        if temporada and tabla.filas:
            salida[temporada] = clase
    return salida


def _volcar_anexo_en_bloques(bloques, informe, declaradas, resultado) -> None:
    """Lleva lo que el cliente declaro a las columnas historicas del flujo.

    Dos listas distintas: los bloques salen del catastro agrupado por CURVA
    ("Palto Hass 2012-2023") y las filas del Anexo salen de lo que declara el
    cliente ("Palto"), en otro orden. Emparejarlas por posicion pondria los
    resultados del nogal en la fila del cerezo, y nada lo delataria. Se empareja
    por especie, con el mismo criterio que usa el resto del sistema.

    Se escriben los INSUMOS -superficie, precios, rendimiento, ratio, costos- y
    no los resultados: las tres filas derivadas del bloque ya son formulas de la
    plantilla, asi que con los insumos puestos ellas reproducen exactamente lo
    declarado. Pelearse con esas formulas seria innecesario y fragil.
    """
    from services.catastro_import import especie_base, normalizar

    def clave(texto):
        return normalizar(especie_base(str(texto or ""), ["(", " - "]))

    anexo = informe.anexo
    tablas = {}
    for tabla in (anexo.resultado_real, anexo.resultado_preliminar):
        temporada = (tabla.temporada or "").strip()
        if temporada and tabla.filas:
            tablas[temporada] = tabla

    # El tipo de cambio se captura con claves simbolicas ('real',
    # 'preliminar') porque el ano lo declara el Anexo. Aqui se le agrega la
    # temporada concreta, sin borrar las simbolicas: el mapping escribe las
    # celdas D3/E3 leyendo justamente esas.
    flujo = informe.flujo_proyectado
    tipo_cambio = flujo.tipo_cambio_pesos_usd
    respaldo = flujo.ventas_respaldadas_pesos
    for simbolica, tabla in (("real", anexo.resultado_real),
                             ("preliminar", anexo.resultado_preliminar)):
        temporada = (tabla.temporada or "").strip()
        if not temporada:
            continue
        # Las ventas respaldadas por los IVAs viajan igual: se capturan con la
        # misma clave simbolica y aqui se les pone el anio, que es lo que el
        # motor necesita para dar con la columna.
        for serie in (tipo_cambio, respaldo):
            if simbolica in serie and serie[simbolica] is not None:
                serie.setdefault(temporada, serie[simbolica])

    for temporada, tabla in tablas.items():
        if temporada not in declaradas:
            continue
        cambio = tipo_cambio.get(temporada)
        if not cambio:
            resultado.avisar(
                "sin tipo de cambio para {}: sus resultados quedan en kilos "
                "pero no en dolares".format(temporada))

        por_especie = {clave(f.especie): f for f in tabla.filas if f.especie}
        for bloque in bloques:
            fila = por_especie.get(clave(bloque.especie_variedad_anio.split()[0]
                                         if bloque.especie_variedad_anio else ""))
            if fila is None:
                continue

            superficie = fila.superficie_ha
            exp = fila.produccion_exportacion_kg or 0.0
            interno = fila.produccion_mercado_interno_kg or 0.0
            produccion = exp + interno

            if superficie:
                bloque.superficie_ha[temporada] = superficie
                if produccion:
                    bloque.rendimiento_kg_ha[temporada] = round(
                        produccion / superficie, 2)
                if fila.costos_directos_pesos and cambio:
                    bloque.costos_usd_ha[temporada] = round(
                        fila.costos_directos_pesos / cambio / superficie, 2)
            elif produccion:
                resultado.avisar(
                    "'{}' declaro produccion en {} pero no la superficie: sin "
                    "ella no se puede calcular su rendimiento ni su costo por "
                    "hectarea".format(fila.especie, temporada))

            if produccion:
                bloque.ratio_exportacion_pct[temporada] = round(
                    exp / produccion, 4)
            if exp and fila.ventas_exportacion_pesos and cambio:
                bloque.precio_exportacion_usd_kg[temporada] = round(
                    fila.ventas_exportacion_pesos / cambio / exp, 4)
            if interno and fila.ventas_mercado_interno_pesos and cambio:
                bloque.precio_mercado_interno_usd_kg[temporada] = round(
                    fila.ventas_mercado_interno_pesos / cambio / interno, 4)


def _proyectar_costos_generales(informe, proyeccion, resultado) -> None:
    """Arrastra a los anios proyectados el gasto indirecto de la ultima temporada.

    Sin esto el EBITDA proyectado sale SIN costos indirectos y por eso inflado
    -medido: un 29% en el informe de demostracion-. Era el unico numero del
    informe que le decia al banco algo que no era.

    Es un SUPUESTO, no un dato, y queda dicho en los comentarios del flujo. Se
    escribe solo donde el tasador no declaro nada: si el puso la administracion
    temporada a temporada, lo suyo manda.
    """
    flujo = informe.flujo_proyectado
    preliminar = informe.anexo.resultado_preliminar
    temporada = (preliminar.temporada or "").strip()
    cambio = (flujo.tipo_cambio_pesos_usd or {}).get(temporada)         or (flujo.tipo_cambio_pesos_usd or {}).get("preliminar")

    for campo, monto in (("administracion", preliminar.administracion_pesos),
                         ("gastos_generales", preliminar.gastos_generales_pesos)):
        if not monto:
            continue
        if not cambio:
            resultado.avisar(
                "sin tipo de cambio: los costos generales proyectados quedan "
                "fuera y el EBITDA sale sin descontarlos")
            return
        serie = getattr(flujo, campo)
        for t in proyeccion:
            serie.setdefault(t, round(monto / cambio, 2))


def _proponer_supuestos(informe, plantaciones, curvas, resultado) -> None:
    """La hoja Supuestos del formato final, desde la biblioteca (el agrónomo,
    2026-10-14: "opcion a" -el programa propone, el tasador corrige-).

    Las 8 temporadas son las que siguen a la ultima temporada REAL: la que
    declaro el tasador en el Anexo 1, o la ultima cerrada a la fecha del
    informe si no la declaro."""
    from config.settings import ajustes as _config
    from services.catastro_import import cargar_reglas
    from services.excel_engine_santander import (
        temporada_real_por_defecto, temporadas_siguientes)
    from services.supuestos import proponer_supuestos
    from services.visita import ajustes_por_curva

    real = (informe.anexo.resultado_real.temporada or "").strip() or \
        temporada_real_por_defecto(informe.antecedentes.fecha_emision_informe)
    try:
        reglas = cargar_reglas(_config.raiz / "config" / "reglas_agronomicas.json")
        formacion_hasta = lambda especie: reglas.umbrales(especie)[0]  # noqa: E731
    except Exception:                                   # noqa: BLE001
        formacion_hasta = None
    resultado.superficie_proyectada_ha = round(sum(
        p.superficie_ha for p in plantaciones), 4)
    if not curvas:
        resultado.avisar(
            "no hay curvas en el catálogo: la hoja Supuestos queda para "
            "llenarla a mano")
        return
    for aviso in proponer_supuestos(
            plantaciones, curvas, temporadas_siguientes(real, 8),
            ajustes_por_curva=ajustes_por_curva(informe),
            formacion_hasta=formacion_hasta):
        resultado.avisar(aviso)
    sin = [p for p in plantaciones if p.supuestos is None]
    resultado.superficie_sin_curva_ha = round(sum(p.superficie_ha for p in sin), 4)


def completar_informe(
    informe: Any,
    plantaciones: Sequence[Any],
    curvas: Sequence[Any],
    visita: Optional[Dict[str, Any]] = None,
    formulario: Optional[Any] = None,
    parametros_riego: Optional[Any] = None,
    comuna: Optional[str] = None,
    caudal_disponible_l_s: Optional[float] = None,
    anio_informe: Optional[int] = None,
    formato: str = "bsan",
) -> ResultadoArmado:
    """Rellena el informe con el catastro, la visita, el flujo y el riego.

    Devuelve un ``ResultadoArmado`` con lo que quedo fuera. El informe se
    modifica en el sitio; lo que ya venia escrito -por ejemplo la opinion que
    redacto la IA desde las fotos- solo se pisa si la visita trae algo.

    ``formato="santander2026"`` es el formato final de Santander: ahi el
    flujo lo arman las formulas del Excel desde la hoja Supuestos, asi que en
    vez del flujo por curva (y del riego calculado, que el formato hace con
    sus propias formulas) se proponen los Supuestos de cada plantacion.
    """
    from datetime import date

    from schemas.informe import BloqueEspecieFlujo, FilaEstructuraTemporada
    from services.envejecimiento import a_matriz_de_dominio, proyectar_cartera

    resultado = ResultadoArmado()
    anio = anio_informe or date.today().year

    # --- 1. la visita: primero, porque de ahi salen los ajustes -----------
    if visita and formulario is not None:
        from services.visita import aplicar_visita
        for aviso in aplicar_visita(informe, visita, formulario):
            resultado.avisar("visita: {}".format(aviso))

    # --- 2. el catastro ---------------------------------------------------
    plantaciones = list(plantaciones or [])
    resultado.plantaciones = len(plantaciones)
    if not plantaciones:
        resultado.avisar(
            "el informe no tiene catastro cargado: sin plantaciones no hay "
            "flujo proyectado ni requerimiento de riego")
        return resultado

    informe.estructura_productiva.plantaciones = list(plantaciones)
    if informe.estructura_productiva.anio_catastro is None:
        informe.estructura_productiva.anio_catastro = anio

    if formato == "santander2026":
        _proponer_supuestos(informe, plantaciones, curvas, resultado)
        return resultado

    # Superficie por especie para la portada.
    por_especie: Dict[str, Dict[str, float]] = {}
    for planta in plantaciones:
        clave = planta.especie or "(sin especie)"
        acumulado = por_especie.setdefault(
            clave, {"total": 0.0, "propio": 0.0, "arriendo": 0.0})
        acumulado["total"] += planta.superficie_ha
        destino = ("arriendo"
                   if getattr(planta.dominio, "value", "") == "Arrendatario"
                   else "propio")
        acumulado[destino] += planta.superficie_ha

    # El cuadro "Resumen superficie plantada" salia vacio aunque el catastro
    # estuviera completo: nadie lo llenaba. Se calcula del mismo catastro con
    # las reglas de madurez del proyecto -plena produccion, rendimientos
    # crecientes, formacion- en vez de dejarselo al tasador fila por fila, que
    # es lo que obligaban las formulas escritas a mano del formato original.
    if not informe.estructura_productiva.resumen_superficie:
        try:
            from config.settings import ajustes as _config
            from schemas.informe import ResumenSuperficieEspecie
            from services.catastro_import import cargar_reglas, clasificar_madurez

            reglas = cargar_reglas(_config.raiz / "config" / "reglas_agronomicas.json")
            futuras = {}
            for planta in plantaciones:
                anio_plantado = _anio_de(planta.anio_plantacion) or 0
                if anio_plantado > anio:
                    futuras[planta.especie] = (
                        futuras.get(planta.especie, 0.0) + planta.superficie_ha)

            informe.estructura_productiva.resumen_superficie = [
                ResumenSuperficieEspecie(
                    especie=r.especie,
                    propietario_ha=round(r.propietario_ha, 2) or None,
                    arrendatario_ha=round(r.arrendatario_ha, 2) or None,
                    plena_produccion_ha=round(r.plena_produccion_ha, 2) or None,
                    rendimientos_crecientes_ha=round(
                        r.rendimientos_crecientes_ha, 2) or None,
                    formacion_ha=round(r.formacion_ha, 2) or None,
                    proyectos_plantacion_ha=round(
                        futuras.get(r.especie, 0.0), 2) or None,
                )
                for r in clasificar_madurez(plantaciones, reglas, anio)
            ]
        except Exception as error:                   # noqa: BLE001
            resultado.avisar(
                "no se pudo armar el resumen de superficie ({})".format(
                    type(error).__name__))

    if not informe.estructura_temporada:
        informe.estructura_temporada = [
            FilaEstructuraTemporada(
                especie=especie, superficie_ha=round(d["total"], 2),
                tenencia_propio_ha=round(d["propio"], 2) or None,
                tenencia_arriendo_ha=round(d["arriendo"], 2) or None)
            for especie, d in sorted(por_especie.items(),
                                     key=lambda kv: -kv[1]["total"])
        ]

    # --- 3. el flujo ------------------------------------------------------
    if not curvas:
        resultado.avisar(
            "no hay curvas en el catalogo: el flujo proyectado queda vacio")
    else:
        from services.visita import ajustes_por_curva

        # El horizonte arranca en la temporada PRELIMINAR: el agrónomo la cuenta
        # como el ano 1 de la proyeccion. Por eso los 8 anios de la caratula
        # (que lee F110:M110) son preliminar + 7 proyectados, y la proyeccion
        # propiamente tal empieza al ano siguiente.
        # El ano 1 es la temporada preliminar si el tasador la declaro; si no,
        # la temporada en curso. Siempre existe: es la columna F, la que abre
        # los 8 anios de la caratula.
        preliminar = (informe.anexo.resultado_preliminar.temporada or "").strip()
        anio_uno = _anio_de(preliminar) or anio
        ano_uno_temporada = preliminar or "{}/{}".format(
            anio_uno, str(anio_uno + 1)[2:])
        proyeccion = _temporadas_de_proyeccion(anio_uno + 1, 8)
        cartera = proyectar_cartera(
            plantaciones, curvas, proyeccion,
            ajustes_por_curva=ajustes_por_curva(informe),
            capacidad_bloques=4)
        for aviso in cartera.avisos:
            resultado.avisar(aviso)

        resultado.superficie_proyectada_ha = cartera.superficie_proyectada_ha
        resultado.superficie_sin_curva_ha = cartera.superficie_excluida_ha
        resultado.bloques_de_flujo = len(cartera.grupos)

        bloques = []
        for grupo in cartera.grupos[:4]:
            comerciales = grupo.curva.comerciales
            precio_exp = comerciales.retorno_exportacion_usd_kg()
            precio_int = comerciales.rnp_mercado_interno_usd_kg
            ratio = comerciales.pct_exportacion
            if not precio_exp and not precio_int:
                resultado.avisar(
                    "la curva '{}' no tiene precios: ese bloque proyecta kilos "
                    "pero no ingresos".format(grupo.curva.nombre))

            produccion, exportacion, interno = {}, {}, {}
            for temporada in proyeccion:
                kilos = grupo.produccion_total_kg.get(temporada, 0.0)
                produccion[temporada] = kilos
                exportacion[temporada] = round(kilos * ratio * precio_exp)
                interno[temporada] = round(kilos * (1 - ratio) * precio_int)

            bloques.append(BloqueEspecieFlujo(
                especie_variedad_anio=grupo.etiqueta,
                superficie_ha=dict(grupo.superficie_ha),
                rendimiento_kg_ha=dict(grupo.rendimiento_kg_ha),
                costos_usd_ha=dict(grupo.costos_usd_ha),
                precio_exportacion_usd_kg={t: precio_exp for t in proyeccion},
                precio_mercado_interno_usd_kg={t: precio_int for t in proyeccion},
                ratio_exportacion_pct={t: ratio for t in proyeccion},
                produccion_total_kg=produccion,
                ingresos_exportacion=exportacion,
                ingresos_mercado_interno=interno,
            ))

        # Las temporadas historicas son las que el tasador DECLARO en el
        # Anexo, no tres fijas hacia atras: no siempre se las entregan. Si no
        # hay ninguna, el informe arranca directo en la proyeccion y esas
        # columnas quedan sin rotular, que es como se ve que no existen.
        clases = _temporadas_declaradas(informe, anio)
        _volcar_anexo_en_bloques(bloques, informe, clases, resultado)
        clases.update({t: "Projection " for t in proyeccion})

        # La lista se alinea a la geometria de la plantilla: D y E para los
        # anios historicos, F para el ano 1 del horizonte y G..N para el resto.
        # Un hueco (None) deja la columna en blanco, que es como se ve que ese
        # ano no lo entregaron. Sin esta alineacion la caratula -que lee
        # F110:M110- mostraria los ocho anios corridos en uno.
        historicos = sorted(
            (t for t in clases if clases[t] == "Real " and t != preliminar),
            key=lambda t: _anio_de(t) or 0)
        huecos = [None] * max(0, 2 - len(historicos))
        declaradas = huecos + historicos[-2:] + [ano_uno_temporada]
        clases.setdefault(ano_uno_temporada, "Preliminar")

        # Los costos indirectos de los anios proyectados: se arrastra lo que el
        # cliente declaro para la temporada preliminar, convertido a USD. Va
        # por la serie POR TEMPORADA y no por un rango de columnas fijo, porque
        # las columnas se corren segun cuantos historicos haya entregado.
        _proyectar_costos_generales(informe, proyeccion, resultado)
        # Se COPIA el flujo previo y se le cambia lo calculado, en vez de
        # construir uno nuevo enumerando que conservar. La version anterior
        # listaba los campos a mano y por eso perdia en silencio cualquiera que
        # no estuviera en la lista: cuando se agrego el valor de la UF, el
        # tasador lo escribia en la visita, viajaba bien hasta aqui y
        # desaparecia sin un solo aviso. Con model_copy, lo que el tasador
        # declaro sobrevive por defecto y solo se pisa lo que este modulo
        # calcula de verdad.
        informe.flujo_proyectado = informe.flujo_proyectado.model_copy(update={
            "temporadas": declaradas + proyeccion,
            "clase_de_temporada": clases,
            "bloques": bloques,
            "matriz": a_matriz_de_dominio(cartera, proyeccion),
            # Este texto se imprime en el informe del banco: va con tildes.
            "comentarios": (
                "Proyección construida con la matriz de envejecimiento sobre "
                "{} plantaciones, agrupadas por curva. Las columnas Real y "
                "Preliminar llevan los resultados que declara el cliente. "
                "Los costos generales de los años proyectados arrastran los "
                "de la temporada en curso: es un supuesto, y está escrito como "
                "fórmula sobre esa celda para que cambiarla mueva la "
                "proyección entera."
                .format(len(plantaciones))),
        })

    # --- 4. el riego ------------------------------------------------------
    if parametros_riego is None:
        return resultado

    from services.riego import balance_de_predio, filas_de_requerimiento

    comuna = comuna or informe.dotacion_recursos.suelos.comuna
    if not comuna:
        resultado.avisar(
            "sin comuna declarada no se puede calcular el requerimiento de "
            "riego: la ETo se busca por comuna")
        return resultado

    # Van TODAS las plantaciones, no las cuatro mayores. La plantilla trae
    # cuatro huecos, pero la tabla declara ``insertar_filas`` y el motor la
    # amplia -igual que hace con el catastro-, asi que recortar aqui no era
    # necesario y salia caro: con cinco plantaciones, la fila 'Superficie Total
    # Plantada' sumaba 114,5 ha mientras la portada del mismo informe decia
    # 124,0. Dos superficies distintas para el mismo cliente en el mismo
    # archivo es exactamente el error que la validacion de consistencia de §11
    # existe para atrapar, y el banco lo ve antes que nosotros.
    #
    # Se ordenan de mayor a menor para que, si algun dia hay que recortar de
    # verdad, lo que quede arriba sea lo que manda en la demanda.
    ordenadas = sorted(plantaciones,
                       key=lambda p: p.superficie_ha, reverse=True)
    if not caudal_disponible_l_s:
        caudal_disponible_l_s = sum(
            (d.caudal_real_ls or d.caudal_nominal_ls or 0.0)
            for d in informe.dotacion_recursos.derechos_agua) or None

    balance = balance_de_predio(
        ordenadas, parametros_riego, comuna=comuna,
        caudal_disponible_l_s=caudal_disponible_l_s, anio_informe=anio)
    for aviso in balance.avisos:
        resultado.avisar("riego: {}".format(aviso))

    if not informe.dotacion_recursos.requerimientos_riego:
        informe.dotacion_recursos.requerimientos_riego = filas_de_requerimiento(
            balance, ordenadas)

    resultado.demanda_riego_l_s = balance.demanda_total_l_s
    if balance.suficiente is False:
        resultado.avisar(
            "el agua declarada no alcanza: faltan {:.1f} L/s ({:.0f}% de la "
            "demanda)".format(balance.deficit_l_s or 0,
                              balance.deficit_pct or 0))

    return resultado
