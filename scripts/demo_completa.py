"""Genera un informe con TODAS las tablas llenas, entrando por la web.

Para que es: hasta ahora cada prueba llenaba una parte -el catastro, la
visita, el riego- y el resto quedaba en blanco. Un informe a medias no deja
ver como se ve el producto ni donde estan los huecos de verdad. Este script
llena todo lo que un tasador puede llenar y genera el .xlsx.

Y lo hace **por los mismos endpoints que usa el navegador**, no llamando a las
funciones por dentro. Esa distincion no es un detalle: las ~430 comprobaciones
del proyecto estaban todas en verde mientras el circuito de la web estaba roto,
justamente porque ninguna entraba por donde entra el tasador.

    python scripts/demo_completa.py

Deja el archivo en output/ y dice que quedo vacio y por que.

Los datos son inventados pero coherentes: un cliente de Quillota -comuna que SI
esta en la tabla de ETo- con tres especies que SI tienen precios cargados
(Palto, Cerezo, Nogal), para que el flujo llegue hasta el EBITDA en vez de
quedarse en kilos.
"""

import json
import os
import sys
import urllib.error
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE = "http://localhost:8000"
SESION = "0" * 32

CLIENTE = "Agricola Santa Lucia Ltda."
RUT = "76.123.456-0"
COMUNA = "Quillota"


def pedir(metodo, ruta, cuerpo=None):
    datos = json.dumps(cuerpo).encode() if cuerpo is not None else None
    req = urllib.request.Request(BASE + ruta, data=datos, method=metodo)
    if datos:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read() or b"null")
    except urllib.error.HTTPError as error:
        detalle = error.read().decode("utf-8", "replace")[:300]
        raise SystemExit(
            "\n{} {} -> HTTP {}\n{}\n\n"
            "Si es 'Connection refused', el servidor no esta arriba:\n"
            "    python -m uvicorn main:app --port 8000\n".format(
                metodo, ruta, error.code, detalle)
        )


# --------------------------------------------------------------- el catastro
# Especies con curva Y con precios cargados, para que el flujo llegue a pesos.
PLANTACIONES = [
    {"especie": "Palto", "variedad": "Hass", "anio_plantacion": "2012",
     "superficie_ha": 42.0, "predio": "Santa Lucia", "dominio": "Propietario",
     "riego": "Goteo", "densidad_plantacion": "5 x 3",
     "sistema_conduccion": "Copa", "estado": "Bueno"},
    {"especie": "Palto", "variedad": "Hass", "anio_plantacion": "2018",
     "superficie_ha": 18.5, "predio": "Santa Lucia", "dominio": "Propietario",
     "riego": "Goteo", "densidad_plantacion": "5 x 2",
     "sistema_conduccion": "Copa", "estado": "Bueno"},
    {"especie": "Palto", "variedad": "Hass", "anio_plantacion": "2023",
     "superficie_ha": 9.0, "predio": "El Almendro", "dominio": "Arrendatario",
     "riego": "Goteo", "densidad_plantacion": "4 x 2",
     "sistema_conduccion": "Eje", "estado": "Bueno"},
    {"especie": "Cerezo", "variedad": "Santina", "anio_plantacion": "2016",
     "superficie_ha": 24.0, "predio": "Santa Lucia", "dominio": "Propietario",
     "riego": "Goteo", "densidad_plantacion": "4 x 2",
     "sistema_conduccion": "Eje", "estado": "Bueno"},
    {"especie": "Cerezo", "variedad": "Regina", "anio_plantacion": "2020",
     "superficie_ha": 15.5, "predio": "El Almendro", "dominio": "Propietario",
     "riego": "Goteo", "densidad_plantacion": "4 x 2",
     "sistema_conduccion": "Eje", "estado": "Regular"},
    {"especie": "Nogal", "variedad": "Chandler", "anio_plantacion": "2014",
     "superficie_ha": 31.0, "predio": "Santa Lucia", "dominio": "Propietario",
     "riego": "Microaspersion", "densidad_plantacion": "7 x 6",
     "sistema_conduccion": "Eje", "estado": "Bueno"},
    {"especie": "Nogal", "variedad": "Chandler", "anio_plantacion": "2027",
     "superficie_ha": 12.0, "predio": "El Almendro", "dominio": "Propietario",
     "riego": "Goteo", "densidad_plantacion": "7 x 6",
     "sistema_conduccion": "Eje", "estado": "Bueno"},
]

# ----------------------------------------------------------------- la visita
VISITA = {
    "resultado_real": [
        {"especie": "Palto", "superficie_ha": 69.5, "produccion_exportacion_kg": 612000,
         "produccion_mercado_interno_kg": 153000,
         "ventas_exportacion_pesos": 1_020_000_000,
         "ventas_mercado_interno_pesos": 178_000_000,
         "costos_directos_pesos": 690_000_000},
        {"especie": "Cerezo", "superficie_ha": 39.5, "produccion_exportacion_kg": 385000,
         "produccion_mercado_interno_kg": 96000,
         "ventas_exportacion_pesos": 892_000_000,
         "ventas_mercado_interno_pesos": 42_000_000,
         "costos_directos_pesos": 520_000_000},
        {"especie": "Nogal", "superficie_ha": 43.0, "produccion_exportacion_kg": 118000,
         "produccion_mercado_interno_kg": 12000,
         "ventas_exportacion_pesos": 274_000_000,
         "ventas_mercado_interno_pesos": 9_000_000,
         "costos_directos_pesos": 168_000_000},
    ],
    "resultado_preliminar": [
        {"especie": "Palto", "superficie_ha": 69.5, "produccion_exportacion_kg": 648000,
         "produccion_mercado_interno_kg": 141000,
         "ventas_exportacion_pesos": 1_142_000_000,
         "ventas_mercado_interno_pesos": 165_000_000,
         "costos_directos_pesos": 724_000_000},
        {"especie": "Cerezo", "superficie_ha": 39.5, "produccion_exportacion_kg": 412000,
         "produccion_mercado_interno_kg": 88000,
         "ventas_exportacion_pesos": 1_004_000_000,
         "ventas_mercado_interno_pesos": 39_000_000,
         "costos_directos_pesos": 548_000_000},
        {"especie": "Nogal", "superficie_ha": 43.0, "produccion_exportacion_kg": 143000,
         "produccion_mercado_interno_kg": 15000,
         "ventas_exportacion_pesos": 331_000_000,
         "ventas_mercado_interno_pesos": 11_000_000,
         "costos_directos_pesos": 181_000_000},
    ],
    # De que temporada es cada tabla. No siempre se las entregan, asi que se
    # declara: la columna del flujo aparece solo si hay dato.
    "temporadas_anexo": {
        "temporada_real": "2024/25",
        "temporada_preliminar": "2025/26",
    },
    "gastos_indirectos": {
        "real_administracion_pesos": 148_000_000,
        "real_gastos_generales_pesos": 96_000_000,
        "prelim_administracion_pesos": 162_000_000,
        "prelim_gastos_generales_pesos": 104_000_000,
    },
    # Los IVAs respaldan buena parte de lo declarado, pero no todo: la fruta se
    # liquida con desfase y siempre queda algo sin facturar al cierre. Un 100%
    # clavado seria sospechoso, no tranquilizador.
    "respaldo_ventas": {
        "real_ventas_respaldadas_pesos": 1_620_000_000,
        "prelim_ventas_respaldadas_pesos": 1_780_000_000,
    },
    "tipo_cambio": {
        "tc_ultima_temporada": 942.0,
        "tc_temporada_en_curso": 968.0,
        "uf_del_informe": 39412.55,
    },
    # El caudal REAL no viene en ningun papel: lo estima el agronomo en
    # terreno, y es contra ese numero que se juzga la suficiencia hidrica.
    "derechos_agua": [
        {"predio": "Fundo Santa Lucia", "comuna": COMUNA,
         "tipo_fuente": "Canal", "nombre_fuente": "Canal Waddington",
         "dominio": "Propietario", "cantidad": 46, "caudal_nominal_ls": 92,
         "caudal_real_ls": 78, "cbr_fojas": "212", "cbr_numero": "154",
         "cbr_anio": 2019, "cbr_conservador": "CBR de Quillota",
         "propietario": CLIENTE},
        {"predio": "Fundo Santa Lucia", "comuna": COMUNA,
         "tipo_fuente": "Pozo profundo", "nombre_fuente": "Pozo N 1",
         "dominio": "Propietario", "caudal_nominal_ls": 35,
         "caudal_real_ls": 30, "cbr_fojas": "98", "cbr_numero": "61",
         "cbr_anio": 2021, "cbr_conservador": "CBR de Quillota",
         "propietario": CLIENTE},
        {"predio": "El Almendro", "comuna": COMUNA,
         "tipo_fuente": "Pozo profundo", "nombre_fuente": "Pozo N 2",
         "dominio": "Propietario", "caudal_nominal_ls": 22,
         "caudal_real_ls": 18, "cbr_fojas": "104", "cbr_numero": "67",
         "cbr_anio": 2022, "cbr_conservador": "CBR de Quillota",
         "propietario": CLIENTE},
    ],
    "maquinaria": [
        {"nombre": "Tractor Deutz 110 HP", "cantidad": 3, "estado": "Bueno"},
        {"nombre": "Pulverizadora de turbina 2.000 L", "cantidad": 2,
         "estado": "Bueno"},
        {"nombre": "Camion 3/4 con rampla", "cantidad": 1, "estado": "Regular"},
    ],
    "infraestructura": [
        {"nombre": "Packing propio 1.200 m2", "cantidad": 1, "estado": "Bueno"},
        {"nombre": "Camara de frio 400 pallets", "cantidad": 2,
         "estado": "Bueno"},
        {"nombre": "Bodega de insumos y taller", "cantidad": 1,
         "estado": "Bueno"},
        {"nombre": "Tranque acumulador 18.000 m3", "cantidad": 1,
         "estado": "Bueno"},
        {"nombre": "Casa patronal y casas de personal", "cantidad": 4,
         "estado": "Regular"},
        {"nombre": "Sala de bombas y filtros", "cantidad": 2,
         "estado": "Bueno"},
    ],
    "sistemas_produccion": [
        {"especie": "Palto", "sistema_produccion": "Convencional",
         "densidad_plantacion": "5 x 3 (666 pl/ha)", "sistema_riego": "Goteo",
         "variedades_con_royalty": "No", "considera_techos": "No",
         "sistema_contra_heladas": "Si", "nivel_tecnologico": "Alto",
         "nivel_homogeneidad": "Alto", "nivel_productivo_promedio": "Alto"},
        {"especie": "Cerezo", "sistema_produccion": "Convencional",
         "densidad_plantacion": "4 x 2 (1.250 pl/ha)",
         "sistema_riego": "Goteo", "variedades_con_royalty": "Si",
         "considera_techos": "Si", "sistema_contra_heladas": "Si",
         "nivel_tecnologico": "Alto", "nivel_homogeneidad": "Medio",
         "nivel_productivo_promedio": "Alto"},
        {"especie": "Nogal", "sistema_produccion": "Convencional",
         "densidad_plantacion": "7 x 6 (238 pl/ha)",
         "sistema_riego": "Microaspersion", "variedades_con_royalty": "No",
         "considera_techos": "No", "sistema_contra_heladas": "No",
         "nivel_tecnologico": "Medio", "nivel_homogeneidad": "Medio",
         "nivel_productivo_promedio": "Medio"},
    ],
    "ajustes_de_curva": [
        {"especie": "Cerezo", "ajuste_pct": -8,
         "motivo": "Sector bajo con problemas de drenaje; se observo asfixia "
                   "radicular en el cuartel 4 y arboles de menor vigor"},
        {"especie": "Nogal", "ajuste_pct": 5,
         "motivo": "Huerto muy bien manejado, poda al dia y riego "
                   "sectorizado con sondas de humedad"},
    ],
    "sistemas_comercializacion": [
        {"especie": "Palto", "empresas_comercializadoras": "Exportadora Andes",
         "comercializacion": "Consignacion", "mercado": "Exportación",
         "periodo_cosecha": "Sep-Ene", "periodo_liquidaciones": "60-90 dias",
         "proceso": "Packing propio", "gestion_exportacion": "Exportadora"},
        {"especie": "Cerezo", "empresas_comercializadoras": "Frutas del Maipo",
         "comercializacion": "Consignacion", "mercado": "Exportación",
         "periodo_cosecha": "Nov-Dic", "periodo_liquidaciones": "90 dias",
         "proceso": "Packing de terceros",
         "gestion_exportacion": "Exportadora"},
        {"especie": "Nogal", "empresas_comercializadoras": "Nueces del Sur",
         "comercializacion": "Precio fijo", "mercado": "Exportación",
         "periodo_cosecha": "Mar-Abr", "periodo_liquidaciones": "30 dias",
         "proceso": "Secado propio", "gestion_exportacion": "Corredor"},
    ],
    "inversiones": [
        {"proyecto": "Plantacion nogal El Almendro", "anio": 2027,
         "superficie_ha": 12.0, "item": "Plantas y plantacion",
         "cantidad_ha": 238, "costo_unitario_usd": 14.0},
        {"proyecto": "Plantacion nogal El Almendro", "anio": 2027,
         "superficie_ha": 12.0, "item": "Riego tecnificado",
         "cantidad_ha": 1, "costo_unitario_usd": 4200.0},
        {"proyecto": "Plantacion nogal El Almendro", "anio": 2027,
         "superficie_ha": 12.0, "item": "Preparacion de suelo",
         "cantidad_ha": 1, "costo_unitario_usd": 1800.0},
        {"proyecto": "Plantacion nogal El Almendro", "anio": 2027,
         "superficie_ha": 12.0, "item": "Manejo anos 1-3",
         "cantidad_ha": 1, "costo_unitario_usd": 3400.0},
    ],
    "opinion": {
        "gestion_texto":
            "La administracion esta profesionalizada: hay un administrador "
            "agronomo a tiempo completo, jefes de cuartel por especie y "
            "registros de labores y aplicaciones al dia. La contabilidad la "
            "lleva una oficina externa y se dispone de estados financieros "
            "auditados de las ultimas tres temporadas. Los presupuestos se "
            "elaboran por cuartel y se controlan mensualmente contra lo "
            "ejecutado.",
        "gestion_calificacion": "Muy Bueno",
        "canales_texto":
            "El cliente trabaja con exportadoras establecidas y con relacion "
            "de mas de ocho temporadas en palto y cerezo. Maneja packing "
            "propio para palto, lo que le da control sobre la condicion de "
            "llegada y le permite negociar mejores retornos. En nogal opera "
            "con precio fijo a traves de un corredor, lo que reduce la "
            "exposicion a la volatilidad pero tambien el retorno en anos "
            "buenos.",
        "canales_calificacion": "Bueno",
        "activos_texto":
            "Los huertos estan bien conformados y con buen estado sanitario "
            "general. El sistema de riego es de goteo en el 82% de la "
            "superficie, con tranque acumulador de 18.000 m3 que da "
            "autonomia. Se observa un sector de 15,5 ha de cerezo con "
            "problemas de drenaje que requiere intervencion. La "
            "infraestructura de proceso y frio esta sobredimensionada "
            "respecto de la produccion actual, lo que es favorable frente a "
            "las plantaciones en desarrollo.",
        "activos_calificacion": "Muy Bueno",
        "conclusion_texto":
            "Se trata de una explotacion consolidada, con superficie "
            "mayoritariamente en plena produccion, especies de buen retorno y "
            "una estructura de costos controlada. La capacidad de generacion "
            "es suficiente para atender las obligaciones proyectadas con "
            "holgura. Los puntos de atencion son el drenaje del sector bajo "
            "de cerezo y la concentracion del 100% de la produccion en "
            "mercados de exportacion.",
        "conclusion_calificacion": "Muy Bueno",
    },
}

# ------------------------------------------------------------------ lo legal
LEGAL = {
    "tipo_documento": "escritura de compraventa",
    "propiedad": {
        "propietario": CLIENTE,
        "rut_propietario": RUT,
        "nombre_predio": "Fundo Santa Lucia",
        "rol_avaluo": "412-18",
        "comuna": COMUNA,
        "provincia": "Quillota",
        "region": "Valparaiso",
        "superficie_escritura_ha": 168.5,
        "deslindes": "Norte: Estero Rautén; Sur: Fundo Los Nogales; "
                     "Oriente: Camino publico a La Cruz; Poniente: "
                     "Canal Waddington.",
        "inscripcion": {"fojas": "1244", "numero": "876", "anio": 2019,
                        "conservador": "CBR de Quillota"},
    },
    "derechos_agua": [
        {"tipo_fuente": "Canal", "nombre_fuente": "Canal Waddington",
         "caudal_litros_segundo": 78.0, "acciones": 46.0,
         "naturaleza": "Consuntivo, permanente y continuo",
         "inscripcion": {"fojas": "212", "numero": "154", "anio": 2019,
                         "conservador": "CBR de Quillota"}},
        {"tipo_fuente": "Pozo profundo", "nombre_fuente": "Pozo N 1",
         "caudal_litros_segundo": 30.0,
         "naturaleza": "Consuntivo, permanente y continuo",
         "inscripcion": {"fojas": "98", "numero": "61", "anio": 2021,
                         "conservador": "CBR de Quillota"}},
        {"tipo_fuente": "Pozo profundo", "nombre_fuente": "Pozo N 2",
         "caudal_litros_segundo": 18.0,
         "naturaleza": "Consuntivo, permanente y continuo",
         "inscripcion": {"fojas": "104", "numero": "67", "anio": 2022,
                         "conservador": "CBR de Quillota"}},
    ],
    "gravamenes": [
        "HIPOTECA a favor de Banco Santander, a Fojas 1290 numero 902 de "
        "2019, por UF 45.000.",
        "PROHIBICION de gravar y enajenar a favor de Banco Santander, a "
        "Fojas 1291 numero 903 de 2019.",
        "SERVIDUMBRE DE TRANSITO a Fojas 88 numero 51 de 2004, en favor de "
        "las demas hijuelas del predio matriz.",
    ],
}


def generar_sin_servidor():
    """Arma el informe llamando a las funciones por dentro, sin HTTP.

    Es el mismo recorrido que hace ``/asistente/generar``: ensamblar lo legal,
    juntarle el catastro, la visita, las curvas y el riego, y escribir el Excel.
    Existe porque el modo normal necesita el servidor levantado, y no siempre se
    puede levantar -por ejemplo si no se esta frente al computador-.

    Lo que se pierde: la garantia de que el camino del NAVEGADOR funciona. Esa
    solo la da entrar por los endpoints, asi que este modo no reemplaza al otro,
    lo acompana.
    """
    import asyncio

    from config.settings import ajustes as config
    from schemas.extraccion import ExtraccionLegal
    from schemas.informe import Antecedentes, Plantacion
    from services.almacen import construir_almacen
    from services.armado import completar_informe
    from services.ensamblador import ensamblar_informe
    from services.excel_engine import MotorExcel, cargar_mapa
    from services.riego import cargar_parametros
    from services.visita import cargar_formulario

    print("\n[1] Lo legal y los antecedentes")
    informe = ensamblar_informe(
        antecedentes=Antecedentes(
            cliente=CLIENTE, rut=RUT, ejecutivo="M. Fuentes",
            sucursal="Quillota", agronomo_informante="el agrónomo Canales",
            n_solicitud="SOL-2026-00982"),
        legal=ExtraccionLegal(**LEGAL),
    )
    print("    predio: {} | {} derecho(s) de agua en las escrituras".format(
        informe.dotacion_recursos.suelos.predio,
        len(informe.dotacion_recursos.derechos_agua)))

    print("\n[2] Catastro, visita, curvas y riego")
    plantaciones = [Plantacion(**p) for p in PLANTACIONES]
    curvas = asyncio.run(construir_almacen().listar_curvas())
    print("    {} plantaciones, {:.1f} ha | {} curvas en el catalogo".format(
        len(plantaciones), sum(p.superficie_ha for p in plantaciones),
        len(curvas)))

    armado = completar_informe(
        informe, plantaciones=plantaciones, curvas=curvas,
        visita=VISITA, formulario=cargar_formulario(config.ruta_formulario_visita),
        parametros_riego=cargar_parametros(config.ruta_parametros_riego),
        comuna=COMUNA,
    )
    print("    {} bloque(s) de flujo | {:.1f} ha proyectadas | riego {} L/s".format(
        armado.bloques_de_flujo, armado.superficie_proyectada_ha,
        armado.demanda_riego_l_s))
    for aviso in armado.avisos:
        print("    aviso: {}".format(aviso))

    print("\n[3] Generar")
    motor = MotorExcel(config.ruta_plantilla, cargar_mapa(config.ruta_mapping))
    destino = config.ruta_salida / "DEMO Santa Lucia.xlsx"
    resultado = motor.generar(informe, destino)
    print("    {} celdas, {} filas insertadas".format(
        resultado.celdas_escritas, resultado.filas_insertadas))
    for aviso in resultado.avisos:
        print("    aviso: {}".format(aviso))
    if resultado.correcciones:
        print("    {} correcciones a formulas de la plantilla".format(
            len(resultado.correcciones)))

    print("\n    archivo: {}".format(destino))
    print('\n    Revisarlo a ojo:  python scripts/ver_excel.py "{}"'.format(
        destino))
    return 0


def main():
    if "--sin-servidor" in sys.argv:
        print("=" * 72)
        print("INFORME DE DEMOSTRACION - TODAS LAS TABLAS LLENAS (sin servidor)")
        print("=" * 72)
        return generar_sin_servidor()

    print("=" * 72)
    print("INFORME DE DEMOSTRACION - TODAS LAS TABLAS LLENAS")
    print("=" * 72)

    print("\n[1] Expediente")
    exp = pedir("POST", "/expedientes", {
        "cliente": CLIENTE, "rut": RUT, "comuna": COMUNA})
    eid = exp["id"]
    print("    creado: {}".format(eid))

    print("\n[2] Catastro")
    r = pedir("PUT", "/expedientes/{}/catastro".format(eid),
              {"plantaciones": PLANTACIONES})
    total = sum(p["superficie_ha"] for p in PLANTACIONES)
    print("    {} plantaciones, {:.1f} ha".format(
        r["plantaciones_guardadas"], total))

    print("\n[3] Visita a terreno")
    v = pedir("PUT", "/visita/{}".format(eid), {"secciones": VISITA})
    for a in v.get("avance", []):
        marca = "lleno " if a.get("completo") else ("VACIO " if a.get("vacio")
                                                    else "parcial")
        print("    {} {:<34} {}/{}".format(
            marca, a.get("titulo", a["id"])[:34], a["llenos"], a["total"]))

    print("\n[4] Generar")
    salida = pedir("POST", "/asistente/generar", {
        "sesion": SESION, "expediente_id": eid,
        "basicos": {
            "cliente": CLIENTE, "rut": RUT, "ejecutivo": "M. Fuentes",
            "sucursal": "Quillota", "agronomo_informante": "el agrónomo Canales",
            "n_solicitud": "SOL-2026-00982",
        },
        "legal": LEGAL, "fotos": [], "galeria_fotos": False,
    })
    print("    {} celdas escritas".format(salida["celdas_escritas"]))
    for a in salida.get("avisos", []):
        print("    aviso: {}".format(a))
    print("\n    archivo: output/{}".format(salida["archivo"]))
    print("\n    Revisarlo a ojo antes de darlo por bueno:")
    print('      python scripts/ver_excel.py "output/{}"'.format(
        salida["archivo"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
