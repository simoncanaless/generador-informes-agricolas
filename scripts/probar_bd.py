"""Diagnostico y prueba de la base de datos.

    python scripts/probar_bd.py            # comprueba conexion y tablas
    python scripts/probar_bd.py --crear    # crea la base y las tablas si faltan
    python scripts/probar_bd.py --migrar   # mueve el catalogo JSON a PostgreSQL
    python scripts/probar_bd.py --probar   # ejercita expedientes de punta a punta

Si la conexion falla, explica que revisar en vez de soltar la traza cruda.
"""

import argparse
import asyncio
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.settings import obtener_ajustes  # noqa: E402

fallos, comprobaciones = [], 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)


def explicar(error: Exception) -> None:
    """Traduce los fallos tipicos de conexion a algo accionable."""
    texto = "{}: {}".format(type(error).__name__, error)
    print("\n  No se pudo conectar.\n  {}\n".format(texto[:300]))

    pistas = []
    bajo = texto.lower()
    if "password authentication failed" in bajo or "contrasena" in bajo:
        pistas.append("La contrasena de DATABASE_URL no es la correcta.")
        pistas.append("Si lleva @ : / o #, hay que codificarla: @ -> %40, # -> %23")
    if "does not exist" in bajo and "database" in bajo:
        pistas.append("La base no existe todavia: ejecuta con --crear")
    if "cambiame" in bajo or "tu_contrasena" in bajo:
        pistas.append("Quedo el marcador sin reemplazar en el .env")
    if "connection refused" in bajo or "no se puede establecer" in bajo:
        pistas.append("El servicio PostgreSQL no esta corriendo o usa otro puerto")
    if "closed in the middle" in bajo:
        pistas.append(
            "El servidor corto la conexion: suele ser el metodo de "
            "autenticacion en pg_hba.conf, o un puerto ocupado por otro programa"
        )
    if not pistas:
        pistas.append("Revisa host, puerto, usuario y base en DATABASE_URL")

    for pista in pistas:
        print("  - {}".format(pista))
    print("\n  DATABASE_URL actual (sin contrasena):")
    url = obtener_ajustes().database_url
    if "@" in url:
        antes, despues = url.rsplit("@", 1)
        usuario = antes.split("//")[-1].split(":")[0]
        print("    {}//{}:******@{}".format(url.split("//")[0], usuario, despues))
    else:
        print("    {}".format(url or "(vacia)"))


async def diagnosticar():
    from db.sesion import comprobar_conexion

    config = obtener_ajustes()
    print("=" * 70)
    print("BASE DE DATOS")
    print("=" * 70)
    print("  configurada : {}".format(config.bd_configurada))
    print("  base        : {}".format(config.nombre_base_datos or "(sin definir)"))

    if not config.bd_configurada:
        print("\n  Falta DATABASE_URL en el archivo .env.")
        print("  Pon ahi la contrasena que usas en pgAdmin 4 y vuelve a ejecutar.")
        return False

    estado = await comprobar_conexion(config)
    if estado["conectada"]:
        print("  servidor    : {}".format(estado["servidor"]))
        print("  tablas      : {}".format(estado["tablas"]))
        return True

    print("\n  NO conectada: {}".format(estado["motivo"]))
    return False


async def crear():
    from db.sesion import crear_base_si_falta, crear_tablas

    config = obtener_ajustes()
    try:
        if await crear_base_si_falta(config):
            print("  base '{}' creada".format(config.nombre_base_datos))
        else:
            print("  la base '{}' ya existia".format(config.nombre_base_datos))
        await crear_tablas(config)
        print("  tablas creadas o ya presentes")
        return True
    except Exception as error:  # noqa: BLE001
        explicar(error)
        return False


async def migrar_catalogo():
    """Mueve curvas, ajustes y formatos aprendidos de los JSON a PostgreSQL."""
    from db.repositorios import AlmacenPostgres
    from services.almacen import AlmacenJSON

    config = obtener_ajustes()
    origen = AlmacenJSON(config.raiz / "data" / "almacen")
    destino = AlmacenPostgres()

    curvas = await origen.listar_curvas()
    ajustes = await origen.listar_ajustes()
    mapeos = await origen.listar_mapeos()
    print("  en los JSON: {} curvas, {} ajustes, {} formatos".format(
        len(curvas), len(ajustes), len(mapeos)))

    for curva in curvas:
        await destino.guardar_curva(curva)
    for ajuste in ajustes:
        await destino.guardar_ajuste(ajuste)
    for mapeo in mapeos:
        await destino.guardar_mapeo(mapeo)

    print("  en PostgreSQL: {} curvas, {} ajustes, {} formatos".format(
        len(await destino.listar_curvas()),
        len(await destino.listar_ajustes()),
        len(await destino.listar_mapeos())))
    print("\n  Los archivos JSON se conservan como respaldo en data/almacen/")


async def probar_expedientes():
    """Ejercita el ciclo completo de un expediente."""
    from db.repositorios import AlmacenPostgres, RepositorioExpedientes
    from schemas.curvas import Ajuste, CategoriaAjuste
    from schemas.informe import Estado, Plantacion, Tenencia
    from services.almacen import NoEncontrado

    repo = RepositorioExpedientes()

    print("\n[1] Alta y recuperacion")
    expediente = await repo.crear(
        cliente="Agricola de Prueba SpA", rut="76.111.222-8",
        ejecutivo="L. San Emeterio", comuna="Santa Barbara", banco="santander",
    )
    id_exp = str(expediente.id)
    comprobar(expediente.estado == "borrador", "nace en estado borrador")
    recuperado = await repo.obtener(id_exp)
    comprobar(recuperado.cliente == "Agricola de Prueba SpA", "se recupera por id")

    print("\n[2] El catastro sobrevive")
    plantaciones = [
        Plantacion(especie="Palto", variedad="Hass", anio_plantacion="2015",
                   superficie_ha=16.6, predio="El Olivar",
                   dominio=Tenencia.PROPIETARIO, estado=Estado.BUENO),
        Plantacion(especie="Palto", variedad="Hass", anio_plantacion="2003-2017",
                   superficie_ha=50.5, predio="El Olivar",
                   dominio=Tenencia.PROPIETARIO, estado=Estado.BUENO),
        Plantacion(especie="Avellano", variedad="Giffoni", anio_plantacion="2011",
                   superficie_ha=89.0, predio="Las Marias",
                   dominio=Tenencia.ARRENDATARIO, estado=Estado.BUENO),
    ]
    guardadas = await repo.reemplazar_plantaciones(id_exp, plantaciones)
    comprobar(guardadas == 3, "guarda las 3 hileras")

    resumen = await repo.resumen_por_especie(id_exp)
    porespecie = {r["especie"]: r for r in resumen}
    print("      resumen calculado por PostgreSQL: {}".format(
        {k: v["superficie_ha"] for k, v in porespecie.items()}))
    comprobar(abs(porespecie["Palto"]["superficie_ha"] - 67.1) < 0.01,
              "PostgreSQL suma las dos hileras de palto (67,1 ha)",
              str(porespecie.get("Palto")))
    comprobar(porespecie["Palto"]["hileras"] == 2, "y cuenta 2 hileras")

    fresco = await repo.obtener(id_exp)
    comprobar(abs((fresco.superficie_total_ha or 0) - 156.1) < 0.01,
              "la superficie total del expediente se actualizo sola",
              str(fresco.superficie_total_ha))

    print("\n[3] Reimportar reemplaza, no acumula")
    await repo.reemplazar_plantaciones(id_exp, plantaciones[:1])
    comprobar(len(await repo.resumen_por_especie(id_exp)) == 1,
              "tras reimportar queda solo la especie nueva")

    print("\n[4] Datos flexibles en JSONB y cambio de estado")
    await repo.actualizar(id_exp, estado="en_revision", datos={
        "legal": {"rol_avaluo": "203-7", "fojas": "408v"},
        "conflictos": ["el rol no coincide entre documentos"],
    })
    fresco = await repo.obtener(id_exp)
    comprobar(fresco.estado == "en_revision", "cambia de estado")
    comprobar(fresco.datos["legal"]["rol_avaluo"] == "203-7",
              "el bloque JSONB conserva la estructura anidada",
              str(fresco.datos))

    print("\n[5] Documentos, fotos e informes generados")
    await repo.agregar_documento(
        id_exp, nombre_archivo="escritura.pdf", via_extraccion="hibrido",
        modelo_ia="gemini-3.6-flash", degradado=False,
        extraccion={"propiedad": {"rol_avaluo": "203-7"}},
    )
    await repo.agregar_fotografia(
        id_exp, nombre_archivo="huerto.jpg", titulo="Vista general",
        categoria="emplazamiento",
    )
    await repo.registrar_generado(
        id_exp, archivo="informe.xlsx", plantilla="santander",
        celdas_escritas=192, filas_insertadas=15,
    )
    fresco = await repo.obtener(id_exp)
    comprobar(len(fresco.documentos) == 1, "el documento quedo asociado")
    comprobar(fresco.documentos[0].via_extraccion == "hibrido",
              "se guarda por que via se extrajo (importa para auditar)")
    comprobar(len(fresco.fotografias) == 1, "la fotografia quedo asociada")
    comprobar(len(fresco.generados) == 1, "queda traza del Excel emitido")

    print("\n[6] Listado y estadisticas")
    listado = await repo.listar(estado="en_revision")
    comprobar(any(str(e.id) == id_exp for e in listado),
              "el expediente aparece filtrando por estado")
    stats = await repo.estadisticas()
    comprobar(stats["expedientes_por_estado"].get("en_revision", 0) >= 1,
              "las estadisticas lo cuentan", str(stats))

    print("\n[7] Borrado en cascada")
    await repo.borrar(id_exp)
    try:
        await repo.obtener(id_exp)
        comprobar(False, "el expediente borrado deberia desaparecer")
    except NoEncontrado:
        comprobar(True, "el expediente y todo lo suyo se borro")

    print("\n[8] El catalogo compartido sobre PostgreSQL")
    almacen = AlmacenPostgres()
    prueba = Ajuste(nombre="Prueba temporal de conexion", factor_pct=-7,
                    categoria=CategoriaAjuste.OTRO)
    guardado = await almacen.guardar_ajuste(prueba)
    comprobar(guardado.nombre == prueba.nombre, "guarda un ajuste en la base")
    await almacen.borrar_ajuste(guardado.id)
    comprobar(True, "y lo borra")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--crear", action="store_true")
    parser.add_argument("--migrar", action="store_true")
    parser.add_argument("--probar", action="store_true")
    args = parser.parse_args()

    if args.crear:
        print("=" * 70)
        print("CREACION DE LA BASE")
        print("=" * 70)
        if not await crear():
            return 1
        print()

    conectada = await diagnosticar()
    if not conectada:
        return 1

    if args.migrar:
        print("\n" + "=" * 70)
        print("MIGRACION DEL CATALOGO JSON -> POSTGRESQL")
        print("=" * 70)
        await migrar_catalogo()

    if args.probar:
        print("\n" + "=" * 70)
        print("PRUEBA DE EXPEDIENTES")
        print("=" * 70)
        try:
            await probar_expedientes()
        except Exception as error:  # noqa: BLE001
            explicar(error)
            return 1
        print("\n" + "=" * 70)
        if fallos:
            print("{}/{} comprobaciones FALLARON".format(len(fallos), comprobaciones))
            for f in fallos:
                print("  - {}".format(f))
            return 1
        print("{}/{} comprobaciones OK".format(comprobaciones, comprobaciones))

    from db.sesion import cerrar
    await cerrar()
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
