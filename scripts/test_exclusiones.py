"""Prueba de las unidades excluibles: hojas o tablas que el tasador saca del
.xlsx antes de descargarlo (el usuario, 2026-09-23).

"en vez de hacer distintos tipos de informes... que lo eliminen antes de
descargar el excel nomas". El riesgo real no es la pantalla: es que una
tabla que se saca puede dejar a otra formula sin de donde leer. El grafo de
dependencias se midio escaneando las 573 formulas reales de la plantilla
mas las que el generador inyecta -ver scripts/graficar_dependencias.py, que
vuelve a hacer esa medicion cada vez que corre-.

Uso: python scripts/test_exclusiones.py
"""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook  # noqa: E402

from config.settings import ajustes  # noqa: E402
from schemas.informe import Antecedentes, InformeFitogreen  # noqa: E402
from services.exclusiones import resolver_cascada  # noqa: E402
from services.excel_engine import MotorExcel, cargar_mapa  # noqa: E402
from test_engine import construir_informe, imagen_de_prueba  # noqa: E402

PLANTILLA = ajustes.ruta_plantilla
MAPPING = ajustes.ruta_mapping

fallos, comprobaciones = [], 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)
    return condicion


def probar_cascada():
    print("\n[1] resolver_cascada: el cierre transitivo, y solo en un sentido")
    mapa = cargar_mapa(MAPPING)
    unidades = mapa.unidades_excluibles

    a_excluir, avisos = resolver_cascada(unidades, ["derechos_agua"])
    comprobar(set(a_excluir) == {"derechos_agua", "requerimientos_riego"},
              "pedir 'derechos_agua' arrastra a 'requerimientos_riego'",
              str(a_excluir))
    comprobar(len(avisos) == 1 and "depende de" in avisos[0],
              "y avisa POR QUE se arrastro -no en silencio", avisos)

    a_excluir, avisos = resolver_cascada(unidades, ["requerimientos_riego"])
    comprobar(a_excluir == ["requerimientos_riego"],
              "pedir SOLO 'requerimientos_riego' no arrastra a 'derechos_agua' "
              "-la dependencia es de un solo sentido", str(a_excluir))
    comprobar(not avisos, "y no hay nada que avisar en ese caso", avisos)

    a_excluir, avisos = resolver_cascada(unidades, ["no_existe"])
    comprobar(a_excluir == [], "pedir algo que no esta en el registro no rompe")
    comprobar(len(avisos) == 1 and "no existe" in avisos[0],
              "y avisa que esa unidad no existe", avisos)

    a_excluir, _ = resolver_cascada(unidades, ["inversiones", "comentarios_opinion"])
    comprobar(set(a_excluir) == {"inversiones", "comentarios_opinion"},
              "dos unidades sin relacion entre si no se arrastran mutuamente",
              str(a_excluir))


def probar_hoja_completa():
    print("\n[2] Excluir una hoja completa: Inversion, sin avisos ruidosos")
    informe = InformeFitogreen(
        antecedentes=Antecedentes(cliente="Prueba Exclusion Hoja", rut="76.543.210-3"))
    motor = MotorExcel(PLANTILLA, cargar_mapa(MAPPING))
    with tempfile.TemporaryDirectory() as tmp:
        destino = Path(tmp) / "out.xlsx"
        resultado = motor.generar(informe, destino, excluir=["inversiones"])
        comprobar(not resultado.avisos,
                  "ningun aviso -ni de la exclusion ni de 'formatos: la hoja "
                  "no existe', que es la misma exclusion vista desde otra "
                  "fase del motor", resultado.avisos)
        hojas = load_workbook(destino).sheetnames
        comprobar("Inversión" not in hojas, "la hoja ya no esta", hojas)
        comprobar("Cliente" in hojas and "Flujo proyectado " in hojas,
                  "y el resto del libro sigue completo", hojas)


def probar_tabla_con_cascada():
    print("\n[3] Excluir una tabla con cascada, sobre el informe CON desbordes")
    with tempfile.TemporaryDirectory() as tmp:
        informe = construir_informe(imagen_de_prueba(tmp))
        motor = MotorExcel(PLANTILLA, cargar_mapa(MAPPING))
        destino = Path(tmp) / "out.xlsx"
        # derechos_agua y requerimientos_riego DESBORDAN (5 items para 4
        # huecos, insertar_filas de por medio): es el caso que de verdad
        # importa probar, no una tabla que nunca crecio.
        resultado = motor.generar(informe, destino, excluir=["derechos_agua"])

        arrastre = [a for a in resultado.avisos if "requerimientos_riego" in a
                   or "Requerimiento de riego" in a]
        comprobar(arrastre, "el aviso de arrastre llega hasta el resultado "
                            "final", resultado.avisos)

        ws = load_workbook(destino)["Dotación de recursos"]
        etiquetas = [ws.cell(f, 1).value for f in range(1, 46) if ws.cell(f, 1).value]
        comprobar("AGUAS DE RIEGO " not in etiquetas,
                  "'Derechos de agua' -la que se pidio- ya no esta", etiquetas)
        comprobar("ESTIMACIÓN REQUERIMIENTO AGUAS DE RIEGO" not in etiquetas,
                  "'Requerimiento de riego' -arrastrada- tampoco", etiquetas)
        comprobar("SUELOS" in etiquetas and
                  "DOTACIÓN DE MAQUINARIA E INFRESTRUCTURA" in etiquetas,
                  "lo que NO se pidio sigue ahi -Suelos arriba, Maquinaria/"
                  "Infraestructura abajo, sin nada de por medio", etiquetas)

        # Maquinaria SI desbordo (6 items, 4 huecos) y esta MAS ABAJO que
        # las dos tablas borradas: si el rango se calculo con la posicion de
        # la PLANTILLA en vez de la posicion real (services/excel_engine.py:
        # _excluir_unidades, ctx.fila_final), esto sale vacio o mezclado.
        maquinas = [ws.cell(f, 1).value for f in range(1, 46)
                   if isinstance(ws.cell(f, 1).value, str)
                   and ws.cell(f, 1).value.startswith("Tractor")]
        comprobar(len(maquinas) == 6,
                  "y una tabla mas abajo que SI desbordo sigue completa, en "
                  "su nueva posicion -no se perdio ni se corrompio con el "
                  "corrimiento de las de arriba", maquinas)


def _fila_con_rotulo(ws, rotulo, columna=1):
    for f in range(1, ws.max_row + 1):
        if ws.cell(f, columna).value == rotulo:
            return f
    return None


def probar_tabla_columnas_no_mueve_filas():
    print("\n[4] Vaciar Maquinaria no mueve ninguna fila -Infraestructura al lado")
    with tempfile.TemporaryDirectory() as tmp:
        ruta_foto = imagen_de_prueba(tmp)

        # Referencia SIN excluir nada: donde caen las cosas con los mismos
        # desbordes (derechos_agua/requerimientos_riego tambien desbordan en
        # este informe), para comparar que nada se corrio despues.
        motor = MotorExcel(PLANTILLA, cargar_mapa(MAPPING))
        base = Path(tmp) / "base.xlsx"
        motor.generar(construir_informe(ruta_foto), base, crear_galeria_fotos=True)
        ws_base = load_workbook(base)["Dotación de recursos"]
        fila_comentarios_base = _fila_con_rotulo(ws_base, "COMENTARIOS DOTACIÓN DE RECURSOS")

        destino = Path(tmp) / "out.xlsx"
        resultado = motor.generar(construir_informe(ruta_foto), destino,
                                  crear_galeria_fotos=True, excluir=["maquinaria"])
        comprobar(not resultado.avisos, "sin avisos", resultado.avisos)

        ws = load_workbook(destino)["Dotación de recursos"]
        fila_encabezado = _fila_con_rotulo(ws, "Tipo de Maquinaria ")
        comprobar(fila_encabezado is not None,
                  "el encabezado de Maquinaria sigue -se vacia el dato, no "
                  "la tabla entera")

        columna_a = [ws.cell(f, 1).value
                    for f in range(fila_encabezado + 1, fila_encabezado + 5)]
        comprobar(all(v is None for v in columna_a),
                  "y la columna A -sus 'Tractor N'- quedo vacia", columna_a)

        infra = [ws.cell(f, 6).value
                for f in range(fila_encabezado + 1, fila_encabezado + 5)
                if ws.cell(f, 6).value]
        comprobar("Bodega" in infra and "Packing" in infra,
                  "Infraestructura -columnas F en adelante, MISMA fila- "
                  "sigue intacta: vaciar Maquinaria no la toco", infra)

        # El punto central de esta prueba: a diferencia de una tabla
        # BORRADA (seccion 3, que corre todo lo de abajo hacia arriba),
        # vaciar columnas no mueve ninguna fila. Todo lo que esta despues de
        # Maquinaria en la hoja tiene que quedar en la MISMA fila que en el
        # informe de referencia sin ninguna exclusion.
        fila_comentarios = _fila_con_rotulo(ws, "COMENTARIOS DOTACIÓN DE RECURSOS")
        comprobar(fila_comentarios == fila_comentarios_base,
                  "el bloque de comentarios sigue en la MISMA fila que sin "
                  "excluir nada ({}), no se corrio".format(fila_comentarios_base),
                  "con exclusion: fila {}".format(fila_comentarios))


def probar_borrado_y_vaciado_combinados():
    """El caso que de verdad importa: una tabla se BORRA (mueve filas) y,
    en la MISMA hoja, mas abajo, otra se VACIA (services/excel_engine.py:
    _excluir_unidades, pasos 2 y 3). Si el vaciado no se entera de cuanto se
    movio todo por el borrado de arriba, limpia las filas equivocadas -ni
    error de Excel ni aviso, solo datos en el lugar que no es-.
    """
    print("\n[5b] Borrar arriba y vaciar abajo, en la MISMA hoja y de una vez")
    with tempfile.TemporaryDirectory() as tmp:
        ruta_foto = imagen_de_prueba(tmp)
        motor = MotorExcel(PLANTILLA, cargar_mapa(MAPPING))
        destino = Path(tmp) / "out.xlsx"
        resultado = motor.generar(
            construir_informe(ruta_foto), destino, crear_galeria_fotos=True,
            excluir=["derechos_agua", "maquinaria"])
        comprobar(len(resultado.avisos) == 1 and "depende de" in resultado.avisos[0],
                  "el unico aviso es el de arrastre -nada de 'la hoja no "
                  "existe' ni otro ruido", resultado.avisos)

        ws = load_workbook(destino)["Dotación de recursos"]
        comprobar(_fila_con_rotulo(ws, "AGUAS DE RIEGO ") is None,
                  "Derechos de agua -y su arrastre, Requerimiento de "
                  "riego- ya no estan, moviendo todo lo de abajo hacia arriba")

        fila_encabezado = _fila_con_rotulo(ws, "Tipo de Maquinaria ")
        comprobar(fila_encabezado is not None,
                  "el encabezado de Maquinaria sigue, en su fila YA CORRIDA "
                  "por el borrado de arriba")
        columna_a = [ws.cell(f, 1).value
                    for f in range(fila_encabezado + 1, fila_encabezado + 7)]
        comprobar(all(v is None for v in columna_a),
                  "y el vaciado de Maquinaria acerto la fila real -no la de "
                  "la plantilla original-, borrando los 'Tractor N' que "
                  "estan ahi ahora", columna_a)
        infra = [ws.cell(f, 6).value
                for f in range(fila_encabezado + 1, fila_encabezado + 7)
                if ws.cell(f, 6).value]
        comprobar("Bodega" in infra,
                  "e Infraestructura, en la misma fila ya corrida, sigue "
                  "intacta", infra)


def probar_api():
    print("\n[6] La API: el listado y que PeticionGeneracion acepte 'excluir'")
    import asyncio

    import api.routers.informes as ri
    from api.routers.asistente import DatosBasicos, PeticionGeneracion

    motor = MotorExcel(PLANTILLA, cargar_mapa(MAPPING))
    listado = asyncio.run(ri.unidades_excluibles(motor))
    comprobar(len(listado) == len(motor.mapa.unidades_excluibles),
              "el endpoint devuelve TODO el registro, no un subconjunto",
              len(listado))
    req = next((u for u in listado if u.id == "requerimientos_riego"), None)
    comprobar(req is not None and req.depende_de == ["derechos_agua"],
              "y trae 'depende_de' tal cual esta en el mapping -la pantalla "
              "arma la cascada con esto, sin volver a preguntarle nada mas "
              "al servidor", req)

    peticion = PeticionGeneracion(
        basicos=DatosBasicos(cliente="Prueba API", rut="76.543.210-3"),
        excluir=["inversiones", "comentarios_opinion"])
    comprobar(peticion.excluir == ["inversiones", "comentarios_opinion"],
              "'/asistente/generar' -la que usa la pantalla real- acepta "
              "la lista de exclusiones en el payload")


def probar_recalculo_con_excel_real():
    print("\n[5] El libro con varias exclusiones combinadas recalcula sin errores")
    import subprocess

    with tempfile.TemporaryDirectory() as tmp:
        informe = construir_informe(imagen_de_prueba(tmp))
        motor = MotorExcel(PLANTILLA, cargar_mapa(MAPPING))
        destino = ajustes.ruta_salida_pruebas / "test_exclusiones_combinadas.xlsx"
        resultado = motor.generar(
            informe, destino, crear_galeria_fotos=True,
            excluir=["derechos_agua", "maquinaria", "infraestructura",
                    "estructura_sistemas_produccion"])
        comprobar(resultado.ok, "el archivo se genero")

    # Fragmentos de impresora-a-salvo/liberar-COM: los mismos que
    # services/renderizado.py -esta automatizacion es SU PROPIA copia
    # independiente, no pasa por ese modulo, y tenia la misma trampa de la
    # impresora sin el arreglo (encontrada porque un script hermano de este
    # -test_retasacion.py- se quedo colgado de verdad, corriendo esta sesion).
    from services.renderizado import (
        GUION_ARRANQUE_IMPRESORA_SEGURA, GUION_LIBERAR_COM_EXCEL, GUION_RESTAURAR_IMPRESORA,
    )

    guion = r"""
$ErrorActionPreference = 'Stop'
$excel = $null
$libro = $null
try {{
{arranque_impresora}
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $libro = $excel.Workbooks.Open('{ruta}', $false, $false)
    $excel.CalculateFullRebuild()
    $excel.CalculateUntilAsyncQueriesDone()
    $errores = 0
    foreach ($hoja in $libro.Sheets) {{
        foreach ($c in $hoja.UsedRange.Cells) {{ if ($c.Text -match '^#') {{ $errores++ }} }}
    }}
    Write-Output "ERRORES=$errores"
    Write-Output "HOJAS=$($libro.Sheets.Count)"
    $libro.Close($false)
}} finally {{
    if ($excel -ne $null) {{ $excel.Quit() }}
{liberar_com}
{restaurar_impresora}
}}
""".format(
        ruta=str(destino.resolve()).replace("'", "''"),
        arranque_impresora=GUION_ARRANQUE_IMPRESORA_SEGURA,
        restaurar_impresora=GUION_RESTAURAR_IMPRESORA,
        liberar_com=GUION_LIBERAR_COM_EXCEL,
    )

    try:
        proceso = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", guion],
            capture_output=True, text=True, timeout=180)
    except (FileNotFoundError, subprocess.TimeoutExpired) as error:
        print("  (no se pudo recalcular con Excel real: {})".format(error))
        return

    salida_ps = proceso.stdout
    if proceso.returncode != 0 or "ERRORES=" not in salida_ps:
        print("  (Excel no esta instalado o fallo: {})".format(
            (proceso.stderr or salida_ps)[:200]))
        return

    valores = dict(l.split("=", 1) for l in salida_ps.splitlines() if "=" in l)
    errores = int(valores.get("ERRORES", "-1"))
    comprobar(errores == 0,
              "cero celdas en error en las {} hojas que quedaron, con "
              "cascada + fila-borrada + columnas-vaciadas a la vez".format(
                  valores.get("HOJAS", "?")),
              "{} en error".format(errores))


def main():
    print("=" * 72)
    print("PRUEBA DE LAS UNIDADES EXCLUIBLES")
    print("=" * 72)

    probar_cascada()
    probar_hoja_completa()
    probar_tabla_con_cascada()
    probar_tabla_columnas_no_mueve_filas()
    probar_borrado_y_vaciado_combinados()
    probar_api()
    probar_recalculo_con_excel_real()

    print("\n" + "=" * 72)
    if fallos:
        print("{}/{} comprobaciones FALLARON:".format(len(fallos), comprobaciones))
        for f in fallos:
            print("  - {}".format(f))
        return 1
    print("{}/{} comprobaciones OK".format(comprobaciones, comprobaciones))
    return 0


if __name__ == "__main__":
    sys.exit(main())
