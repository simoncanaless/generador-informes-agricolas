"""Prueba de las curvas de produccion, los ajustes y el mapeo adaptable.

Las curvas se contrastan contra los valores del archivo original
'HdU - Plantation Dec-2025 vf Esp.xlsx' (hoja CURVA PX-RNP, 8.1 a 8.4).

El mapeo adaptable se prueba con un catastro inventado a proposito con
encabezados que NO estan en el diccionario ('CUARTEL', 'VAR.', 'SUP HÁS'),
para comprobar que el sistema se adapta a formatos que nunca ha visto.

Uso: python scripts/test_curvas.py [--sin-ia]
"""

import argparse
import asyncio
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.settings import ajustes  # noqa: E402
from graficar_curvas import graficar_curvas  # noqa: E402
from schemas.curvas import (  # noqa: E402
    Ajuste, AjusteAplicado, BaseDescuento, CadenaExportacion, CategoriaAjuste,
    CurvaProduccion, Descuento, EstructuraCosto, ParametrosComerciales,
    PuntoCurva, SolicitudProyeccion, TipoDescuento,
)
from services.almacen import AlmacenJSON, NoEncontrado  # noqa: E402
from services.envejecimiento import curvas_de_la_misma_especie  # noqa: E402
from services.catastro_import import (  # noqa: E402
    cargar_reglas, leer_catastro, proponer_mapeo,
)
from services.curvas import (  # noqa: E402
    curva_desde_parametros, factor_de_ajustes, forma_de_curva,
    parecido_de_forma, proyectar, rendimiento_a_edad, resumen_por_madurez,
)

fallos, comprobaciones = [], 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)


def casi(a, b, tol=0.01):
    return a is not None and abs(float(a) - float(b)) < tol


def curva_palto():
    """El palto del archivo original, con sus cifras exactas."""
    rendimientos = {0: 0, 1: 0, 2: 3000, 3: 6000, 4: 10000, 5: 13000}
    costos = {0: 0, 1: 1500, 2: 2500, 3: 4000, 4: 5000, 5: 6000}
    puntos = [
        PuntoCurva(edad=e, rendimiento_kg_ha=r,
                   costo_directo_usd_ha=costos.get(e, 6000))
        for e, r in rendimientos.items()
    ]
    puntos += [
        PuntoCurva(edad=e, rendimiento_kg_ha=13000, costo_directo_usd_ha=6000)
        for e in range(6, 15)
    ]
    return CurvaProduccion(
        nombre="Palto", especie="Palto", puntos=puntos,
        comerciales=ParametrosComerciales(
            pct_exportacion=0.7, rnp_exportacion_usd_kg=1.75,
            rnp_mercado_interno_usd_kg=1.25,
        ),
        costos=EstructuraCosto(costo_base_usd_ha=6000, costo_por_kg_usd=0.5),
    )


# ------------------------------------------------------------ la curva

def probar_curva():
    print("\n[1] Evaluacion de la curva (contrastada con el archivo original)")
    curva = curva_palto()

    comprobar(casi(rendimiento_a_edad(curva, 5), 13000),
              "palto a los 5 anios rinde 13.000 kg/ha")
    comprobar(casi(rendimiento_a_edad(curva, 2), 3000),
              "a los 2 anios, 3.000 kg/ha")
    comprobar(casi(rendimiento_a_edad(curva, 0), 0), "recien plantado no produce")
    comprobar(casi(rendimiento_a_edad(curva, 40), 13000),
              "mas alla de la tabla se mantiene el ultimo valor, no se extrapola")

    comprobar(curva.edad_plena_produccion == 5,
              "detecta que la plena produccion empieza a los 5",
              str(curva.edad_plena_produccion))
    comprobar(casi(curva.comerciales.precio_ponderado_usd_kg(), 1.6),
              "precio ponderado = 0,7x1,75 + 0,3x1,25 = 1,60 USD/kg",
              str(curva.comerciales.precio_ponderado_usd_kg()))
    comprobar(casi(curva.costos.costo_usd_ha(13000), 12500),
              "costo en plena produccion = 6000 + 13000x0,5 = 12.500 USD/ha",
              str(curva.costos.costo_usd_ha(13000)))


def probar_ajustes():
    print("\n[2] Ajustes del tasador")

    costero = AjusteAplicado(nombre="Sector costero", factor_pct=-10,
                             categoria=CategoriaAjuste.CLIMA)
    gestion = AjusteAplicado(nombre="Administracion deficiente", factor_pct=-10,
                             categoria=CategoriaAjuste.GESTION)

    comprobar(casi(factor_de_ajustes([costero]), 0.9), "un -10% deja el 0,90")
    comprobar(casi(factor_de_ajustes([costero, gestion]), 0.81),
              "dos -10% dejan 0,81 (se multiplican), no 0,80",
              str(factor_de_ajustes([costero, gestion])))
    comprobar(casi(factor_de_ajustes([]), 1.0), "sin ajustes no se altera nada")

    brutal = [AjusteAplicado(nombre="x", factor_pct=-60) for _ in range(3)]
    comprobar(factor_de_ajustes(brutal) > 0,
              "tres ajustes de -60% no producen rendimiento negativo",
              str(factor_de_ajustes(brutal)))

    mejora = AjusteAplicado(nombre="Manejo de excelencia", factor_pct=5,
                            categoria=CategoriaAjuste.GESTION)
    comprobar(casi(factor_de_ajustes([costero, mejora]), 0.945),
              "los ajustes al alza tambien cuentan: 0,9 x 1,05 = 0,945",
              str(factor_de_ajustes([costero, mejora])))


def probar_proyeccion():
    print("\n[3] Proyeccion sobre una plantacion concreta")
    curva = curva_palto()

    base = proyectar(curva, SolicitudProyeccion(
        curva_id="x", superficie_ha=20.0, edad_actual=8))
    punto = base.puntos[0]
    print("      sin ajustes: {:.0f} kg/ha x 20 ha = {:,.0f} kg | "
          "ingreso {:,.0f} USD | costo {:,.0f} | margen {:,.0f}".format(
              punto.rendimiento_ajustado_kg_ha, punto.produccion_total_kg,
              punto.ingreso_usd, punto.costo_usd, punto.margen_usd))

    comprobar(casi(punto.produccion_total_kg, 260000),
              "13.000 kg/ha x 20 ha = 260.000 kg")
    comprobar(casi(punto.ingreso_usd, 416000),
              "260.000 kg x 1,60 USD/kg = 416.000 USD", str(punto.ingreso_usd))
    comprobar(casi(punto.costo_usd, 250000),
              "12.500 USD/ha x 20 ha = 250.000 USD", str(punto.costo_usd))
    comprobar(casi(punto.margen_usd, 166000), "margen = 416.000 - 250.000")

    ajustada = proyectar(curva, SolicitudProyeccion(
        curva_id="x", superficie_ha=20.0, edad_actual=8,
        ajustes=[
            AjusteAplicado(nombre="Sector costero", factor_pct=-10,
                           categoria=CategoriaAjuste.CLIMA,
                           justificacion="predio a 18 km de la costa"),
            AjusteAplicado(nombre="Administracion deficiente", factor_pct=-10,
                           categoria=CategoriaAjuste.GESTION,
                           justificacion="poda atrasada, se observo en terreno"),
        ],
    ))
    p2 = ajustada.puntos[0]
    print("      con -10% y -10%: {:.0f} kg/ha | ingreso {:,.0f} | "
          "costo {:,.0f} | margen {:,.0f}".format(
              p2.rendimiento_ajustado_kg_ha, p2.ingreso_usd, p2.costo_usd,
              p2.margen_usd))

    comprobar(casi(ajustada.factor_total, 0.81), "el factor total es 0,81")
    comprobar(casi(ajustada.variacion_pct, -19.0), "informa -19% de variacion",
              str(ajustada.variacion_pct))
    comprobar(casi(p2.rendimiento_ajustado_kg_ha, 10530),
              "13.000 x 0,81 = 10.530 kg/ha")
    comprobar(casi(p2.costo_usd, 225300),
              "el costo variable BAJA con el rendimiento: (6000 + 10530x0,5) x 20",
              str(p2.costo_usd))
    comprobar(p2.costo_usd < punto.costo_usd,
              "cosechar menos fruta cuesta menos que cosechar mas")
    comprobar(p2.margen_usd < punto.margen_usd, "el margen cae al bajar la curva")

    horizonte = proyectar(curva, SolicitudProyeccion(
        curva_id="x", superficie_ha=10.0, horizonte=12))
    comprobar(len(horizonte.puntos) == 13, "proyecta el horizonte completo 0..12",
              str(len(horizonte.puntos)))
    # El anio 0 vale 0 en ingreso y en costo: el archivo original imputa la
    # plantacion como inversion aparte, no como costo directo de la temporada.
    comprobar(casi(horizonte.puntos[0].margen_usd, 0),
              "el anio de plantacion no suma ingreso ni costo directo",
              str(horizonte.puntos[0].margen_usd))
    comprobar(horizonte.puntos[1].margen_usd < 0,
              "desde el anio 1 el margen es negativo: hay costo y no hay fruta",
              str(horizonte.puntos[1].margen_usd))
    comprobar(horizonte.puntos[8].margen_usd > 0,
              "en plena produccion el margen es positivo",
              str(horizonte.puntos[8].margen_usd))


def probar_cadena_exportacion():
    print("\n[2b] De FOB publico a retorno del productor")
    # El FOB es el dato verificable; el retorno se calcula. Numeros redondos
    # para poder comprobar a mano.
    cadena = CadenaExportacion(
        precio_fob_usd_kg=3.00,
        fuente_fob="ODEPA",
        referencia_fob="promedio 2025",
        descuentos=[
            Descuento(nombre="Packing", tipo=TipoDescuento.POR_KG, valor=0.50),
            Descuento(nombre="Flete interno", tipo=TipoDescuento.POR_KG, valor=0.25),
            Descuento(nombre="Comision exportadora",
                      tipo=TipoDescuento.PORCENTAJE, valor=0.10,
                      base=BaseDescuento.FOB),
        ],
    )
    # 3,00 - 0,50 - 0,25 - (10% de 3,00 = 0,30) = 1,95
    comprobar(casi(cadena.rlp_usd_kg(), 1.95),
              "3,00 FOB - 0,50 packing - 0,25 flete - 10% comision = 1,95",
              str(cadena.rlp_usd_kg()))
    comprobar(len(cadena.desglose()) == 3,
              "el desglose trae una fila por descuento, para la tabla del informe")
    comprobar(casi(cadena.desglose()[-1]["monto_usd_kg"], 0.30),
              "la comision sobre FOB son 0,30 y no un porcentaje del saldo",
              str(cadena.desglose()[-1]["monto_usd_kg"]))

    # La base del porcentaje NO es un detalle: mismo 10%, resultado distinto.
    sobre_saldo = cadena.model_copy(deep=True)
    sobre_saldo.descuentos[-1].base = BaseDescuento.SALDO
    # saldo antes de la comision = 2,25; 10% = 0,225 -> 2,025
    comprobar(casi(sobre_saldo.rlp_usd_kg(), 2.025),
              "la misma comision sobre el saldo da 2,025, no 1,95: por eso se "
              "declara la base en vez de asumirla",
              str(sobre_saldo.rlp_usd_kg()))

    # Sin cadena, se usa el retorno declarado: los catalogos que ya existen
    # siguen funcionando igual.
    viejo = ParametrosComerciales(
        pct_exportacion=0.7, rnp_exportacion_usd_kg=1.75,
        rnp_mercado_interno_usd_kg=1.25)
    comprobar(casi(viejo.retorno_exportacion_usd_kg(), 1.75),
              "sin cadena manda el retorno declarado (compatibilidad)")
    comprobar(casi(viejo.precio_ponderado_usd_kg(), 1.6),
              "y el precio ponderado no cambia: 0,7x1,75 + 0,3x1,25 = 1,60")

    # Con cadena, el calculo manda sobre el numero tecleado.
    nuevo = ParametrosComerciales(
        pct_exportacion=0.7, rnp_exportacion_usd_kg=99.0,
        rnp_mercado_interno_usd_kg=1.25, cadena_exportacion=cadena)
    comprobar(casi(nuevo.retorno_exportacion_usd_kg(), 1.95),
              "con cadena manda el calculado, no el tecleado a mano",
              str(nuevo.retorno_exportacion_usd_kg()))
    comprobar(casi(nuevo.precio_ponderado_usd_kg(), 0.7 * 1.95 + 0.3 * 1.25),
              "y el precio ponderado del flujo usa ese retorno calculado",
              str(nuevo.precio_ponderado_usd_kg()))

    solo_fob = CadenaExportacion(precio_fob_usd_kg=2.40)
    comprobar(casi(solo_fob.rlp_usd_kg(), 2.40),
              "sin descuentos cargados todavia, el retorno es el propio FOB")

    # Descuentos mal cargados dan negativo. No se corrige en silencio: se ve.
    rota = CadenaExportacion(
        precio_fob_usd_kg=1.00,
        descuentos=[Descuento(nombre="Packing", valor=1.50)])
    comprobar(rota.rlp_usd_kg() < 0,
              "si los descuentos superan al FOB el retorno sale negativo y "
              "queda a la vista, en vez de recortarse a cero")


def probar_curva_parametrica():
    print("\n[3b] Curva generada por incremento proporcional (4 numeros)")

    # Uva de mesa: llega a plena produccion rapido, en 3 temporadas.
    uva = curva_desde_parametros(
        nombre="Uva de mesa", especie="Uva de mesa",
        anios_improductivos=1, nivel_productivo_inicial_kg_ha=5000,
        anios_a_plena_produccion=3, nivel_plena_produccion_kg_ha=15000,
    )
    comprobar(casi(rendimiento_a_edad(uva, 0), 0), "uva: anio de plantacion, 0")
    comprobar(casi(rendimiento_a_edad(uva, 1), 5000),
              "uva: primer anio productivo, nivel inicial 5.000")
    comprobar(casi(rendimiento_a_edad(uva, 2), 10000),
              "uva: a mitad de camino, interpola 10.000",
              str(rendimiento_a_edad(uva, 2)))
    comprobar(casi(rendimiento_a_edad(uva, 3), 15000),
              "uva: a los 3 anios, plena produccion")
    comprobar(casi(rendimiento_a_edad(uva, 10), 15000),
              "uva: mas alla de la tabla se sostiene, no se extrapola")

    # Nogal: demora bastante mas, 6-7 anios, con varios anios improductivos.
    nogal = curva_desde_parametros(
        nombre="Nogal Chandler", especie="Nogal", variedad="Chandler",
        anios_improductivos=3, nivel_productivo_inicial_kg_ha=200,
        anios_a_plena_produccion=7, nivel_plena_produccion_kg_ha=3000,
    )
    comprobar(casi(rendimiento_a_edad(nogal, 0), 0), "nogal: recien plantado, 0")
    comprobar(casi(rendimiento_a_edad(nogal, 2), 0),
              "nogal: sigue en 0 durante todos los anios improductivos",
              str(rendimiento_a_edad(nogal, 2)))
    comprobar(casi(rendimiento_a_edad(nogal, 3), 200),
              "nogal: primer anio productivo, nivel inicial 200")
    comprobar(casi(rendimiento_a_edad(nogal, 5), 1600),
              "nogal: a mitad de la rampa, interpola 1.600",
              str(rendimiento_a_edad(nogal, 5)))
    comprobar(casi(rendimiento_a_edad(nogal, 7), 3000),
              "nogal: a los 7 anios, plena produccion")

    # Se integra sin cambios con la proyeccion economica normal.
    proyeccion = proyectar(nogal, SolicitudProyeccion(
        curva_id="x", superficie_ha=5.0, edad_actual=7))
    comprobar(casi(proyeccion.puntos[0].rendimiento_ajustado_kg_ha, 3000),
              "la curva parametrica se proyecta igual que cualquier otra")

    # Validaciones: no debe aceptar parametros sin sentido.
    try:
        curva_desde_parametros(
            nombre="x", especie="x", anios_improductivos=5,
            nivel_productivo_inicial_kg_ha=100, anios_a_plena_produccion=3,
            nivel_plena_produccion_kg_ha=1000,
        )
        comprobar(False, "rechaza plena produccion antes que fin de improductivos")
    except ValueError:
        comprobar(True, "rechaza plena produccion antes que fin de improductivos")

    try:
        curva_desde_parametros(
            nombre="x", especie="x", anios_improductivos=0,
            nivel_productivo_inicial_kg_ha=2000, anios_a_plena_produccion=5,
            nivel_plena_produccion_kg_ha=1000,
        )
        comprobar(False, "rechaza nivel inicial mayor que el de plena produccion")
    except ValueError:
        comprobar(True, "rechaza nivel inicial mayor que el de plena produccion")

    try:
        curva_desde_parametros(
            nombre="x", especie="x", anios_improductivos=-1,
            nivel_productivo_inicial_kg_ha=100, anios_a_plena_produccion=5,
            nivel_plena_produccion_kg_ha=1000,
        )
        comprobar(False, "rechaza anios improductivos negativos")
    except ValueError:
        comprobar(True, "rechaza anios improductivos negativos")

    try:
        curva_desde_parametros(
            nombre="x", especie="x", anios_improductivos=0,
            nivel_productivo_inicial_kg_ha=100, anios_a_plena_produccion=5,
            nivel_plena_produccion_kg_ha=0,
        )
        comprobar(False, "rechaza plena produccion en 0")
    except ValueError:
        comprobar(True, "rechaza plena produccion en 0")

    # Sin anios improductivos: produce desde el mismo anio de plantacion.
    berries = curva_desde_parametros(
        nombre="Arandano", especie="Arandano",
        anios_improductivos=0, nivel_productivo_inicial_kg_ha=1000,
        anios_a_plena_produccion=4, nivel_plena_produccion_kg_ha=9000,
    )
    comprobar(casi(rendimiento_a_edad(berries, 0), 1000),
              "sin anios improductivos, produce desde la plantacion",
              str(rendimiento_a_edad(berries, 0)))
    comprobar(casi(rendimiento_a_edad(berries, 2), 5000),
              "arandano: a mitad de la rampa, interpola 5.000",
              str(rendimiento_a_edad(berries, 2)))
    comprobar(len(berries.puntos) == 2,
              "sin anios improductivos bastan 2 puntos, no se agregan de mas",
              str(len(berries.puntos)))

    # Un solo anio improductivo: no debe duplicar el punto edad=0.
    palto_joven = curva_desde_parametros(
        nombre="xx", especie="xx", anios_improductivos=1,
        nivel_productivo_inicial_kg_ha=500, anios_a_plena_produccion=6,
        nivel_plena_produccion_kg_ha=8000,
    )
    comprobar(len({p.edad for p in palto_joven.puntos}) == len(palto_joven.puntos),
              "no genera edades repetidas cuando hay un solo anio improductivo")

    # Metadatos: variedad, comerciales y costos viajan completos; sin
    # indicarlos, quedan en los valores por defecto (no None ni error).
    con_metadatos = curva_desde_parametros(
        nombre="Palto Hass costero", especie="Palto", variedad="Hass",
        anios_improductivos=2, nivel_productivo_inicial_kg_ha=1000,
        anios_a_plena_produccion=6, nivel_plena_produccion_kg_ha=13000,
        comerciales=ParametrosComerciales(
            pct_exportacion=0.6, rnp_exportacion_usd_kg=1.8,
            rnp_mercado_interno_usd_kg=1.1,
        ),
        costos=EstructuraCosto(costo_base_usd_ha=5500, costo_por_kg_usd=0.45),
        fuente="entrevista con el cliente", notas="huerto joven",
        autor="el agrónomo Canales",
    )
    comprobar(con_metadatos.variedad == "Hass", "conserva la variedad")
    comprobar(casi(con_metadatos.comerciales.precio_ponderado_usd_kg(), 1.52),
               "conserva los parametros comerciales indicados",
               str(con_metadatos.comerciales.precio_ponderado_usd_kg()))
    comprobar(casi(con_metadatos.costos.costo_base_usd_ha, 5500),
              "conserva la estructura de costos indicada")
    comprobar(con_metadatos.fuente == "entrevista con el cliente",
              "conserva la fuente")
    comprobar(con_metadatos.autor == "el agrónomo Canales", "conserva el autor")

    sin_metadatos = curva_desde_parametros(
        nombre="xx", especie="xx", anios_improductivos=1,
        nivel_productivo_inicial_kg_ha=100, anios_a_plena_produccion=4,
        nivel_plena_produccion_kg_ha=1000,
    )
    comprobar(sin_metadatos.variedad is None,
              "sin variedad indicada, queda None y no falla")
    comprobar(casi(sin_metadatos.comerciales.precio_ponderado_usd_kg(), 0.0),
              "sin comerciales indicados, usa el valor por defecto (0)")
    comprobar(casi(sin_metadatos.costos.costo_base_usd_ha, 0.0),
              "sin costos indicados, usa el valor por defecto (0)")

    # Verificacion visual: un assert dice que a los 5 anios da 1.600 kg/ha;
    # no dice si la rampa se ve razonable o si un punto quedo mal puesto.
    # Misma disciplina que ver_excel.py: mirar el PNG antes de dar por buena
    # la logica del generador.
    carpeta = ajustes.ruta_salida_pruebas / "curvas"

    comparacion = graficar_curvas(
        [uva, nogal, berries, con_metadatos],
        carpeta / "parametricas_comparacion.png",
        titulo="Curvas parametricas: 4 especies, 4 rampas distintas",
    )
    comprobar(comparacion.exists(), "genera el grafico comparativo")

    detalle_nogal = graficar_curvas(
        [nogal], carpeta / "parametricas_nogal_detalle.png",
        titulo="Nogal: puntos declarados (solidos) vs interpolacion lineal",
    )
    comprobar(detalle_nogal.exists(), "genera el grafico de detalle del nogal")

    # La prueba de fuego: comparar la curva parametrica contra la curva REAL
    # del palto (tabulada a mano desde el HdU, con su propia forma no
    # perfectamente lineal) usando los mismos hitos -1 anio improductivo,
    # 3.000 kg/ha el primer anio, 13.000 kg/ha a los 5-. Si la aproximacion
    # por rampa se aleja demasiado del original, tiene que notarse a simple
    # vista en este grafico, no solo en un numero.
    palto_manual = curva_palto()
    palto_manual.nombre = "Palto (HdU, tabulada a mano)"
    palto_parametrico = curva_desde_parametros(
        nombre="Palto (parametrico)", especie="Palto",
        anios_improductivos=1, nivel_productivo_inicial_kg_ha=3000,
        anios_a_plena_produccion=5, nivel_plena_produccion_kg_ha=13000,
    )
    comparacion_real = graficar_curvas(
        [palto_manual, palto_parametrico],
        carpeta / "parametrica_vs_real_palto.png",
        titulo="Palto: tabla real del HdU vs aproximacion por 4 numeros",
        horizonte=10,
    )
    comprobar(comparacion_real.exists(),
              "genera el grafico parametrica vs curva real")

    print("      graficos en: {}".format(carpeta))


def probar_declive():
    print("\n[3c] Declive por senescencia")
    # Un frutal no se queda en su peak para siempre. Forma del palto real del
    # HdU: sostiene 13.000 hasta los 20 y cae a 10.100 a los 25.
    palto = curva_desde_parametros(
        nombre="Palto con declive", especie="Palto",
        anios_improductivos=1, nivel_productivo_inicial_kg_ha=3000,
        anios_a_plena_produccion=5, nivel_plena_produccion_kg_ha=13000,
        edad_inicio_declive=20, edad_final=25, nivel_final_kg_ha=10100)

    comprobar(casi(rendimiento_a_edad(palto, 10), 13000),
              "en plena produccion sostiene el peak")
    comprobar(casi(rendimiento_a_edad(palto, 20), 13000),
              "y lo sostiene hasta la edad de inicio del declive")
    comprobar(casi(rendimiento_a_edad(palto, 22), 11840),
              "a mitad del declive interpola: 13.000 - 2/5 x 2.900 = 11.840",
              str(rendimiento_a_edad(palto, 22)))
    comprobar(casi(rendimiento_a_edad(palto, 25), 10100),
              "y llega al nivel final declarado")
    # Un huerto viejo produce poco, pero no deja de producir: no se extrapola
    # la caida hasta cero.
    comprobar(casi(rendimiento_a_edad(palto, 40), 10100),
              "pasada la edad final se sostiene, no sigue cayendo",
              str(rendimiento_a_edad(palto, 40)))
    comprobar(palto.rendimiento_maximo == 13000,
              "el techo de la curva sigue siendo el peak, no el nivel final")

    # Sin declive, el comportamiento de antes: meseta indefinida.
    plana = curva_desde_parametros(
        nombre="Sin declive", especie="Palto", anios_improductivos=1,
        nivel_productivo_inicial_kg_ha=3000, anios_a_plena_produccion=5,
        nivel_plena_produccion_kg_ha=13000)
    comprobar(casi(rendimiento_a_edad(plana, 40), 13000),
              "sin declive la curva no cambia: las curvas ya cargadas siguen "
              "funcionando igual")

    # Los tres datos del declive van juntos o no van.
    for parciales, motivo in (
        ({"edad_inicio_declive": 20}, "solo la edad de inicio"),
        ({"edad_inicio_declive": 20, "edad_final": 25}, "sin el nivel final"),
    ):
        try:
            curva_desde_parametros(
                nombre="xx", especie="xx", anios_improductivos=1,
                nivel_productivo_inicial_kg_ha=3000,
                anios_a_plena_produccion=5,
                nivel_plena_produccion_kg_ha=13000, **parciales)
            comprobar(False, "rechaza un declive a medias ({})".format(motivo))
        except ValueError:
            comprobar(True, "rechaza un declive a medias ({})".format(motivo))

    # Y tienen que tener sentido entre si.
    for parciales, motivo in (
        ({"edad_inicio_declive": 3, "edad_final": 25, "nivel_final_kg_ha": 9000},
         "el declive empieza antes de la plena produccion"),
        ({"edad_inicio_declive": 20, "edad_final": 18, "nivel_final_kg_ha": 9000},
         "termina antes de empezar"),
        ({"edad_inicio_declive": 20, "edad_final": 25,
          "nivel_final_kg_ha": 20000},
         "el nivel final supera al de plena produccion"),
    ):
        try:
            curva_desde_parametros(
                nombre="xx", especie="xx", anios_improductivos=1,
                nivel_productivo_inicial_kg_ha=3000,
                anios_a_plena_produccion=5,
                nivel_plena_produccion_kg_ha=13000, **parciales)
            comprobar(False, "rechaza que " + motivo)
        except ValueError:
            comprobar(True, "rechaza que " + motivo)

    # El declive se combina con los ajustes del tasador sin pelearse.
    ajustada = proyectar(palto, SolicitudProyeccion(
        curva_id="x", superficie_ha=10.0, edad_actual=25,
        ajustes=[AjusteAplicado(nombre="Manejo deficiente", factor_pct=-10)]))
    comprobar(casi(ajustada.puntos[0].rendimiento_ajustado_kg_ha, 9090),
              "un -10% sobre un huerto en declive: 10.100 x 0,9 = 9.090",
              str(ajustada.puntos[0].rendimiento_ajustado_kg_ha))


def probar_coherencia():
    print("\n[4] Coherencia entre la curva y los umbrales del informe")
    curva = curva_palto()
    reglas = cargar_reglas(ajustes.raiz / "config" / "reglas_agronomicas.json")
    formacion, crecientes = reglas.umbrales("Palto")

    resultado = resumen_por_madurez(curva, formacion, crecientes)
    print("      curva llega al techo a los {} anios; el informe declara plena "
          "produccion desde los {}".format(
              resultado["edad_techo_curva"], crecientes + 1))
    comprobar(resultado["edad_techo_curva"] == 5,
              "detecta que el palto llega al techo a los 5")
    comprobar(resultado["coherente"],
              "los umbrales configurados son compatibles con la curva")


async def probar_almacen():
    print("\n[5] Persistencia (misma interfaz para JSON y para PostgreSQL)")
    with tempfile.TemporaryDirectory() as tmp:
        almacen = AlmacenJSON(tmp)
        curva = curva_palto()
        curva.nombre = "Palto Hass costa"
        await almacen.guardar_curva(curva)

        comprobar(len(await almacen.listar_curvas()) == 1, "guarda una curva nueva")
        recuperada = await almacen.obtener_curva(curva.id)
        comprobar(recuperada.nombre == "Palto Hass costa", "la recupera por id")
        comprobar(len(recuperada.puntos) == len(curva.puntos),
                  "conserva todos los puntos de la curva")

        curva.notas = "ajustada tras la visita de julio"
        await almacen.guardar_curva(curva)
        comprobar(len(await almacen.listar_curvas()) == 1,
                  "guardar dos veces la misma curva no la duplica")
        comprobar((await almacen.obtener_curva(curva.id)).notas is not None,
                  "conserva la edicion")

        await almacen.guardar_ajuste(Ajuste(
            nombre="Sector costero", factor_pct=-10,
            categoria=CategoriaAjuste.CLIMA))
        comprobar(len(await almacen.listar_ajustes()) == 1, "guarda un ajuste")

        await almacen.borrar_curva(curva.id)
        comprobar(len(await almacen.listar_curvas()) == 0, "borra la curva")
        try:
            await almacen.obtener_curva(curva.id)
            comprobar(False, "buscar una curva borrada deberia fallar")
        except NoEncontrado:
            comprobar(True, "avisa con claridad si el id no existe")


# ------------------------------------------------- catastro en formato ajeno

def catastro_ajeno(ruta):
    """Planilla con encabezados que NO estan en el diccionario de sinonimos."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Hoja1"
    ws["A1"] = "AGRICOLA SANTA ELENA"
    ws["A2"] = "Catastro frutal temporada 2025/26"
    encabezados = ["CUARTEL", "ESPECIE FRUTAL", "VAR.", "SUP HÁS",
                   "AÑO PLANT.", "MARCO (m)", "RIEGO", "OBSERV."]
    for i, texto in enumerate(encabezados, start=1):
        ws.cell(4, i, texto)
    filas = [
        ["C-1", "Palto", "Hass", 12.5, 2015, "5 x 3", "Goteo", "sin novedad"],
        ["C-2", "Palto", "Hass", 8.3, 2018, "5 x 3", "Goteo", "replante parcial"],
        ["C-3", "Cerezo", "Santina", 6.0, 2020, "4 x 2", "Goteo", ""],
        ["C-4", "Cerezo", "Lapins", 4.2, 2021, "4 x 2", "Goteo", ""],
        ["C-5", "Nogal", "Chandler", 15.0, 2012, "7 x 7", "Goteo", "buen estado"],
    ]
    for f, fila in enumerate(filas, start=5):
        for c, valor in enumerate(fila, start=1):
            ws.cell(f, c, valor)
    ws.cell(11, 1, "TOTAL")
    ws.cell(11, 4, "=SUM(D5:D10)")
    wb.save(ruta)
    return ruta


async def probar_mapeo_adaptable(usar_ia):
    print("\n[6] Catastro en un formato nunca visto")
    reglas = cargar_reglas(ajustes.raiz / "config" / "reglas_agronomicas.json")

    with tempfile.TemporaryDirectory() as tmp:
        ruta = catastro_ajeno(Path(tmp) / "santa_elena.xlsx")
        almacen = AlmacenJSON(Path(tmp) / "almacen")

        ia = None
        if usar_ia:
            from services.ai_service import construir_servicio_ia
            ia = construir_servicio_ia()

        propuesta = await proponer_mapeo(ruta, reglas, ia=ia, almacen=almacen)
        print("      hoja {!r}, encabezado en fila {}, {} filas".format(
            propuesta.hoja, propuesta.fila_encabezado, propuesta.filas_estimadas))
        for c in propuesta.columnas:
            print("      {} {:16} -> {:20} [{}] {}".format(
                c.letra, c.encabezado[:16], str(c.campo), c.origen,
                c.muestras[:2]))

        comprobar(propuesta.fila_encabezado == 4,
                  "encuentra el encabezado saltandose el titulo",
                  str(propuesta.fila_encabezado))
        comprobar(propuesta.filas_estimadas == 5, "cuenta las 5 hileras",
                  str(propuesta.filas_estimadas))
        comprobar(propuesta.completo,
                  "identifica al menos especie y superficie", str(propuesta.mapeo))
        comprobar(propuesta.mapeo.get("especie") == "B",
                  "'ESPECIE FRUTAL' es la especie", str(propuesta.mapeo.get("especie")))
        comprobar(propuesta.mapeo.get("superficie_ha") == "D",
                  "'SUP HÁS' es la superficie", str(propuesta.mapeo.get("superficie_ha")))
        comprobar(propuesta.mapeo.get("anio_plantacion") == "E",
                  "'AÑO PLANT.' es el anio", str(propuesta.mapeo.get("anio_plantacion")))

        if usar_ia:
            comprobar(propuesta.mapeo.get("predio") == "A",
                      "la IA reconoce que 'CUARTEL' es el predio",
                      str(propuesta.mapeo.get("predio")))
            comprobar(propuesta.mapeo.get("variedad") == "C",
                      "y que 'VAR.' es la variedad",
                      str(propuesta.mapeo.get("variedad")))

        # Importacion con el mapeo confirmado
        importado = leer_catastro(ruta, reglas, columnas_forzadas=propuesta.mapeo)
        comprobar(len(importado.plantaciones) == 5, "importa las 5 plantaciones",
                  str(len(importado.plantaciones)))
        comprobar(casi(importado.superficie_total_ha, 46.0),
                  "suma 46,0 ha", str(importado.superficie_total_ha))
        # La palabra TOTAL esta en la columna A, pero la especie se mapeo a la B:
        # la fila queda vacia en las columnas que importan y se salta como tal.
        # Lo que hay que garantizar es que no entre como una plantacion mas.
        especies = {p.especie.upper() for p in importado.plantaciones}
        comprobar("TOTAL" not in especies,
                  "la fila TOTAL no se cuela como una plantacion", str(especies))
        comprobar(all(p.superficie_ha > 0 for p in importado.plantaciones),
                  "ninguna plantacion importada quedo sin superficie")

        # Memoria del formato
        from schemas.curvas import MapeoCatastro
        await almacen.guardar_mapeo(MapeoCatastro(
            huella=propuesta.huella, nombre="Agricola Santa Elena",
            columnas=propuesta.mapeo))
        segunda = await proponer_mapeo(ruta, reglas, ia=None, almacen=almacen)
        comprobar(segunda.aprendido_de == "Agricola Santa Elena",
                  "la segunda vez reconoce el formato sin preguntarle a nadie",
                  str(segunda.aprendido_de))
        comprobar(segunda.mapeo == propuesta.mapeo,
                  "y reproduce el mismo mapeo exacto")


def probar_serie_para_graficar():
    print("\n[N] La curva muestreada edad por edad, para el grafico")
    # El grafico de la pantalla NO interpola por su cuenta: pide la serie al
    # servidor, que la evalua con el mismo codigo que alimenta el informe.
    # Si esto se separara, el agronomo estaria mirando un dibujo que no es el
    # que despues se usa para proyectar.
    from services.curvas import curva_desde_parametros, rendimiento_a_edad

    curva = curva_desde_parametros(
        nombre="Manzano Gala", especie="Manzano", variedad="Gala",
        anios_improductivos=3, nivel_productivo_inicial_kg_ha=8000,
        anios_a_plena_produccion=7, nivel_plena_produccion_kg_ha=60000,
        edad_inicio_declive=20, edad_final=30, nivel_final_kg_ha=42000)

    edades = list(range(0, max(p.edad for p in curva.puntos) + 1))
    serie = [rendimiento_a_edad(curva, e) for e in edades]

    comprobar(len(edades) == 31, "se muestrean las 31 edades, no solo los "
                                 "6 puntos declarados", str(len(edades)))
    comprobar(all(serie[e] == 0 for e in (0, 1, 2)),
              "los anios improductivos van en cero", str(serie[:3]))
    comprobar(serie[3] == 8000, "arranca en el nivel inicial declarado",
              str(serie[3]))
    comprobar(serie[4] > serie[3] and serie[6] > serie[5],
              "la rampa sube edad a edad, no de golpe",
              str(serie[3:8]))
    comprobar(serie[7] == 60000, "llega al peak en el anio declarado",
              str(serie[7]))
    comprobar(all(serie[e] == 60000 for e in range(7, 21)),
              "y lo sostiene hasta que empieza el declive")
    comprobar(serie[21] < serie[20] and serie[25] < serie[21],
              "despues baja, edad a edad", str([serie[20], serie[21], serie[25]]))
    comprobar(serie[30] == 42000, "hasta el nivel final declarado",
              str(serie[30]))

    # Lo que el grafico anota: el peak y a que edad ocurre.
    comprobar(max(serie) == 60000 and serie.index(max(serie)) == 7,
              "el peak que rotula el grafico es el real",
              "{} a los {}".format(max(serie), serie.index(max(serie))))

    # Sin declive la curva se sostiene: es el comportamiento de las curvas que
    # no declaran los tres parametros opcionales.
    plana = curva_desde_parametros(
        nombre="Kiwi", especie="Kiwi", anios_improductivos=3,
        nivel_productivo_inicial_kg_ha=5000, anios_a_plena_produccion=6,
        nivel_plena_produccion_kg_ha=30000)
    ultima = max(p.edad for p in plana.puntos)
    comprobar(rendimiento_a_edad(plana, ultima) == 30000,
              "una curva sin declive termina en su peak, no cayendo",
              str(rendimiento_a_edad(plana, ultima)))


# ---------------------------------------------------------------------------
# El declive expresado como PORCENTAJE ANUAL (2026-09-08)
#
# el agrónomo: "cualquier curva de produccion se puede construir con ... anios de
# plena produccion (n), % de decrecimiento productivo a partir de anio n, nivel
# productivo base al final de la vida util". El generador pedia la EDAD en que
# termina de bajar; el porcentaje es el dato que el tasador tiene de verdad.

def probar_partir_de_otra_curva():
    """Partir de la curva de OTRA VARIEDAD DE LA MISMA ESPECIE.

    el agrónomo planteo que el programa podria proponer una curva parecida, y
    el usuario acoto el alcance, que era lo que faltaba: *«las especies de arandanos
    que le recomiende curvas de otros arandanos, no cerezos con arandanos o
    kiwi con paltos, eso no tiene sentido»*.

    Asi que se recomienda una sola cosa: **otras variedades de la misma
    especie**. Un arandano Emerald se parece a un arandano Duke y eso no hay
    que inventarlo; que un cerezo de numeros cercanos es casualidad
    aritmetica. El propio `emparejar_curva` ya lo dice: *«inventarle una curva
    a un avellano porque se parece a un nogal es exactamente lo que este
    proyecto no debe hacer»*.

    Las comprobaciones que mas valen aqui son negativas: que NO se recomiende
    otra especie, que sin datos declarados NO se invente un orden, y que al
    copiar NO se hereden los precios.
    """
    print("\n[13] Partir de la curva de otra variedad")

    cerezo = CurvaProduccion(
        nombre="Cerezo", especie="Cerezo",
        puntos=[PuntoCurva(edad=0, rendimiento_kg_ha=0),
                PuntoCurva(edad=2, rendimiento_kg_ha=0),
                PuntoCurva(edad=3, rendimiento_kg_ha=4000),
                PuntoCurva(edad=6, rendimiento_kg_ha=13000),
                PuntoCurva(edad=20, rendimiento_kg_ha=10100)],
        comerciales=ParametrosComerciales(rnp_exportacion_usd_kg=3.4),
        costos=EstructuraCosto(costo_base_usd_ha=6000, costo_por_kg_usd=0.5),
    )
    nogal = CurvaProduccion(
        nombre="Nogal", especie="Nogal",
        puntos=[PuntoCurva(edad=0, rendimiento_kg_ha=0),
                PuntoCurva(edad=3, rendimiento_kg_ha=800),
                PuntoCurva(edad=7, rendimiento_kg_ha=6000)],
    )
    parron = CurvaProduccion(
        nombre="Vid", especie="Vid",
        puntos=[PuntoCurva(edad=0, rendimiento_kg_ha=0),
                PuntoCurva(edad=2, rendimiento_kg_ha=6000),
                PuntoCurva(edad=4, rendimiento_kg_ha=14000)],
    )

    forma = forma_de_curva(cerezo)
    comprobar(forma["entra_en_produccion"] == 3,
              "la forma dice a que anio entra en produccion",
              "dio {}".format(forma["entra_en_produccion"]))
    comprobar(forma["peak_kg_ha"] == 13000,
              "y cual es su peak", "dio {}".format(forma["peak_kg_ha"]))
    comprobar(forma["edad_plena_produccion"] == 6,
              "y a que edad llega a plena produccion")
    comprobar(forma["tiene_declive"] is True,
              "reconoce que modela senescencia")
    comprobar(forma_de_curva(nogal)["tiene_declive"] is False,
              "y que una que termina en su techo NO la modela")

    # LA REGLA: solo se recomiendan otras variedades de la MISMA especie.
    arandano_duke = CurvaProduccion(
        nombre="Arandano Duke", especie="Arandano", variedad="Duke",
        puntos=[PuntoCurva(edad=0, rendimiento_kg_ha=0),
                PuntoCurva(edad=3, rendimiento_kg_ha=6000),
                PuntoCurva(edad=6, rendimiento_kg_ha=11000)])
    arandano_legacy = CurvaProduccion(
        nombre="Arandanos Legacy", especie="Arandanos", variedad="Legacy",
        puntos=[PuntoCurva(edad=0, rendimiento_kg_ha=0),
                PuntoCurva(edad=3, rendimiento_kg_ha=5000),
                PuntoCurva(edad=6, rendimiento_kg_ha=9000)])
    catalogo = [cerezo, nogal, parron, arandano_duke, arandano_legacy]

    mismas, otras = curvas_de_la_misma_especie("Arandano", catalogo)
    comprobar([c.nombre for c in mismas] == ["Arandano Duke", "Arandanos Legacy"],
              "para un arandano se ofrecen los OTROS arandanos",
              "salieron {}".format([c.nombre for c in mismas]))
    comprobar("Arandanos Legacy" in [c.nombre for c in mismas],
              "y el plural no lo despista: reusa la regla de `emparejar_curva`")
    comprobar(cerezo not in mismas and parron not in mismas,
              "el cerezo y la vid NO se recomiendan para un arandano",
              "que den numeros cercanos es casualidad, no agronomia")
    comprobar([c.nombre for c in otras] == ["Cerezo", "Nogal", "Vid"],
              "quedan aparte, para ofrecerlas rotuladas como lo que son")

    vacias, todas = curvas_de_la_misma_especie("Kiwi", catalogo)
    comprobar(not vacias,
              "una especie sin hermanas no recibe NINGUNA recomendacion",
              "ofrecerle un cerezo seria justo lo que hay que evitar")
    comprobar(len(todas) == len(catalogo), "y el catalogo entero queda en 'otras'")

    sin_especie, resto = curvas_de_la_misma_especie("", catalogo)
    comprobar(not sin_especie and len(resto) == len(catalogo),
              "sin especie no se recomienda nada: no hay contra que comparar")

    # Sin nada declarado no hay orden.
    comprobar(parecido_de_forma(forma) is None,
              "SIN datos declarados no se inventa ningun parecido",
              "un orden inventado es peor que no ordenar")

    # Los dos numeros declarados ordenan DENTRO del grupo recomendable, que es
    # donde valen: entre un Duke y un Legacy la diferencia es de rendimiento,
    # y esa la sabe el agronomo. Un Emerald que hace ~11.000 se parece al Duke.
    orden = sorted(
        (parecido_de_forma(forma_de_curva(c), 3, 11000), c.nombre)
        for c in mismas)
    comprobar(orden[0][1] == "Arandano Duke",
              "entre las variedades de la especie, ordena por lo que se declara",
              "salio {}".format(orden[0][1]))
    distancias = dict((n, d) for d, n in orden)
    comprobar(distancias["Arandano Duke"] < distancias["Arandanos Legacy"],
              "el Duke (11.000) le gana al Legacy (9.000) para un Emerald de 11.000",
              "{:.2f} contra {:.2f}".format(distancias["Arandano Duke"],
                                            distancias["Arandanos Legacy"]))

    # Copiar la FORMA, no los precios. Un precio heredado de otra especie
    # entra al informe del banco con pinta de hecho.
    derivada = CurvaProduccion(
        nombre="Arandano Duke", especie="Arandano", variedad="Duke",
        puntos=[PuntoCurva(edad=p.edad, rendimiento_kg_ha=p.rendimiento_kg_ha)
                for p in cerezo.puntos],
        derivada_de=cerezo.nombre,
    )
    comprobar(forma_de_curva(derivada)["peak_kg_ha"] == 13000,
              "la curva derivada hereda la forma")
    comprobar(derivada.comerciales.rnp_exportacion_usd_kg == 0,
              "y NO hereda el precio de la especie de la que partio",
              "el cerezo vale 3,4 USD/kg y el arandano no")
    comprobar(derivada.costos.costo_base_usd_ha == 0,
              "ni el costo")

    # Y queda anotado: una curva prestada no puede leerse como una medida.
    vuelta = CurvaProduccion.model_validate(derivada.model_dump(mode="json"))
    comprobar(vuelta.derivada_de == "Cerezo",
              "de que curva salio sobrevive a guardar y releer",
              "sin esto la curva prestada seria indistinguible de una medida")
    comprobar(cerezo.derivada_de is None,
              "y una curva normal no queda rotulada")


def probar_declive_por_porcentaje():
    from services.curvas import declive_desde_porcentaje

    print("\n[declive] El % anual se traduce a edad final")

    # El palto del HdU es el contraste: su tabla real va de 13.000 a 10.100
    # entre los 10 y los 20 anios. Eso es 2,5% anual compuesto, y la funcion
    # tiene que reproducirlo.
    r = declive_desde_porcentaje(13000, 10, 2.5, 10100)
    comprobar(r["edad_final"] == 20,
              "el palto del HdU: 13.000 -> 10.100 a 2,5%/anio termina a los 20",
              "dio {}".format(r["edad_final"]))
    comprobar(r["anios_de_declive"] == 10, "o sea 10 anios de declive")
    comprobar(r["nivel_final_kg_ha"] == 10100,
              "y el nivel final es el declarado, no uno calculado")

    # Compuesto, no lineal: cada anio se pierde el % de lo que QUEDABA.
    lento = declive_desde_porcentaje(13000, 10, 1, 10100)
    rapido = declive_desde_porcentaje(13000, 10, 5, 10100)
    comprobar(lento["anios_de_declive"] > r["anios_de_declive"] >
              rapido["anios_de_declive"],
              "a menor porcentaje, mas anios de declive",
              "{} > {} > {}".format(lento["anios_de_declive"],
                                    r["anios_de_declive"],
                                    rapido["anios_de_declive"]))

    print("\n[declive] Lo imposible se rechaza con el motivo escrito")
    for caso, argumentos in (
            ("el nivel final por encima del peak",
             (13000, 10, 2.5, 20000)),
            ("un decrecimiento de 0%", (13000, 10, 0, 10100)),
            ("un decrecimiento de 100%", (13000, 10, 100, 10100)),
            ("un nivel final de cero", (13000, 10, 2.5, 0))):
        try:
            declive_desde_porcentaje(*argumentos)
            comprobar(False, "rechaza {}".format(caso), "no fallo")
        except ValueError as error:
            comprobar(len(str(error)) > 30,
                      "rechaza {} y explica por que".format(caso),
                      str(error)[:60])


def probar_costo_desde_plena_produccion():
    """El costo como lo sabe el agronomo: los dos extremos de la hectarea.

    Pedido de el usuario: *«hay que agregar poder poner cual es el costo de cuando
    llega a plena produccion»*. Nadie tiene en la cabeza «0,50 USD por kilo»,
    pero si cuanto le cuesta la hectarea parada y cuanto con el huerto a full.

    El coeficiente variable de §8.2 sale de esos dos, y se deduce en el
    SERVIDOR: la pantalla lo muestra mientras se escribe, pero el que alimenta
    el informe es este (§2). El contraste que vale es que reproduzca los
    numeros ya tabulados del HdU, que se cargaron por otro camino.
    """
    from services.curvas import costos_desde_plena_produccion

    print("\n[costos] Del costo en plena produccion sale el costo variable")

    # El palto del HdU: base 6.000, en plena 12.500 sobre 13.000 kg/ha. La
    # tabla trae 0,50 USD/kg, y §8.2 lo usa de caso de prueba desde el
    # principio. Si la resta estuviera al reves, o dividiera por otra cosa,
    # este numero no saldria.
    palto = costos_desde_plena_produccion(6000, 12500, 13000)
    comprobar(casi(palto.costo_por_kg_usd, 0.50, 0.0001),
              "el palto del HdU: 6.000 -> 12.500 sobre 13.000 kg da 0,50 USD/kg",
              "dio {}".format(palto.costo_por_kg_usd))
    comprobar(palto.costo_base_usd_ha == 6000,
              "y el costo base se conserva tal cual se declaro")

    # Y la vuelta completa: con esa estructura, la hectarea a plena produccion
    # tiene que volver a costar los 12.500 que se escribieron. Es el mismo
    # numero por dos caminos, que es el contraste que usa todo el proyecto.
    comprobar(casi(palto.costo_usd_ha(13000), 12500, 0.01),
              "y la hectarea a plena produccion vuelve a costar 12.500",
              "dio {}".format(palto.costo_usd_ha(13000)))

    # Un huerto sin costo variable: cuesta lo mismo parado que produciendo.
    # Es raro pero no es un error, y tiene que dar cero, no fallar.
    plano = costos_desde_plena_produccion(5000, 5000, 10000)
    comprobar(plano.costo_por_kg_usd == 0,
              "un huerto que cuesta lo mismo parado que a full da 0 USD/kg")

    print("\n[costos] Lo que no puede ser se rechaza con el motivo escrito")
    for caso, argumentos in (
            ("que produciendo cueste MENOS que parado", (12000, 6000, 13000)),
            ("un nivel de plena produccion de cero", (6000, 12500, 0)),
            ("un nivel de plena produccion negativo", (6000, 12500, -1)),
            ("un costo negativo", (-1, 12500, 13000))):
        try:
            costos_desde_plena_produccion(*argumentos)
            comprobar(False, "rechaza {}".format(caso), "no fallo")
        except ValueError as error:
            comprobar(len(str(error)) > 30,
                      "rechaza {} y explica por que".format(caso),
                      str(error)[:70])

    # El limonero, para no probar con un solo caso: base 4.000, en plena 6.100
    # sobre 14.000 kg/ha -> 0,15 USD/kg, que es lo que trae su curva cargada.
    limon = costos_desde_plena_produccion(4000, 6100, 14000)
    comprobar(casi(limon.costo_por_kg_usd, 0.15, 0.0001),
              "el limonero: 4.000 -> 6.100 sobre 14.000 kg da 0,15 USD/kg",
              "dio {}".format(limon.costo_por_kg_usd))


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--sin-ia", action="store_true")
    args = parser.parse_args()

    print("=" * 72)
    print("PRUEBA DE CURVAS, AJUSTES Y MAPEO ADAPTABLE")
    print("=" * 72)

    probar_curva()
    probar_ajustes()
    probar_cadena_exportacion()
    probar_proyeccion()
    probar_curva_parametrica()
    probar_declive()
    probar_declive_por_porcentaje()
    probar_costo_desde_plena_produccion()
    probar_coherencia()
    probar_serie_para_graficar()
    probar_partir_de_otra_curva()
    await probar_almacen()
    await probar_mapeo_adaptable(not args.sin_ia and ajustes.ia_configurada)

    print("\n" + "=" * 72)
    if fallos:
        print("{}/{} comprobaciones FALLARON:".format(len(fallos), comprobaciones))
        for f in fallos:
            print("  - {}".format(f))
        return 1
    print("{}/{} comprobaciones OK".format(comprobaciones, comprobaciones))
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))

