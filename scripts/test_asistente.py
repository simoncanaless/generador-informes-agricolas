"""Prueba del flujo asistido completo: lo que haria el tasador de principio a fin.

Sube dos documentos que SE CONTRADICEN a proposito (distinto rol de avaluo) mas
un par de fotos, revisa el borrador, genera el Excel y comprueba que los datos
cayeron en las celdas correctas de la plantilla del banco.

Requiere el servidor levantado:  python -m uvicorn main:app --port 8000
Uso: python scripts/test_asistente.py
"""

import io
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import httpx  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

from config.settings import ajustes  # noqa: E402
from services.ensamblador import fusionar_extracciones  # noqa: E402
from schemas.extraccion import (  # noqa: E402
    ExtraccionLegal, InscripcionCBR, PropiedadExtraida,
)

# FITO_URL, igual que test_persistencia: en la carpeta de trabajo se prueba
# contra el servidor de dev (8001), nunca contra el de la casa.
URL = os.environ.get("FITO_URL", "http://127.0.0.1:8000")
fallos = []
comprobaciones = 0


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)


ESCRITURA = """
NOTARIA DE RANCAGUA - COMPRAVENTA

Don PEDRO ANTONIO ROJAS CACERES vende a AGRICOLA LAS ENCINAS SpA, RUT
76.543.210-K, el inmueble denominado FUNDO EL MAITEN, comuna de Requinoa,
Provincia de Cachapoal, de una superficie de CUARENTA Y SIETE COMA CINCO
HECTAREAS (47,5 has), Rol de Avaluo numero 123-45.

TITULOS: inscrito a fojas mil doscientos cuarenta numero ochocientos noventa
del Registro de Propiedad del Conservador de Bienes Raices de Rancagua del ano
dos mil once.

AGUAS: se transfieren doce coma cinco acciones del Canal Requinoa, equivalentes
a treinta litros por segundo, de naturaleza consuntiva y permanente, inscritas
a fojas trescientos veinte numero ciento quince del ano dos mil doce.

GRAVAMENES: hipoteca de primer grado a favor del Banco Santander Chile.
"""

# Mismo predio, pero el rol y la superficie NO coinciden con la escritura.
CERTIFICADO = """
SERVICIO DE IMPUESTOS INTERNOS
CERTIFICADO DE AVALUO FISCAL

Propietario: AGRICOLA LAS ENCINAS SpA
Predio: FUNDO EL MAITEN
Comuna: Requinoa
Rol de Avaluo: 123-46
Superficie: CUARENTA Y NUEVE COMA DOS HECTAREAS (49,2 has)

Se deja constancia de que el predio cuenta ademas con un derecho de
aprovechamiento sobre pozo profundo de ocho litros por segundo, de caracter
consuntivo permanente.
"""


def pdf(texto, ruta):
    import pymupdf

    doc = pymupdf.open()
    pagina = doc.new_page()
    pagina.insert_textbox(pymupdf.Rect(50, 50, 545, 780), texto, fontsize=10)
    doc.save(ruta)
    doc.close()
    return ruta


def foto(ruta, tono):
    from PIL import Image, ImageDraw

    imagen = Image.new("RGB", (900, 675), (150, 190, 235))
    dibujo = ImageDraw.Draw(imagen)
    dibujo.rectangle([0, 380, 900, 675], fill=(120, 95, 60))
    for fila in range(4):
        y = 400 + fila * 65
        for col in range(8):
            x = 30 + col * 115 - fila * 20
            dibujo.ellipse([x, y - 40, x + 55, y + 12], fill=tono)
    imagen.save(ruta)
    return ruta


# ------------------------------------------- unidad: fusion sin red

def probar_fusion_offline():
    print("\n[1] Cruce de documentos (determinista, sin IA)")

    a = ExtraccionLegal(propiedad=PropiedadExtraida(
        propietario="Agrícola Las Encinas SpA", nombre_predio="Fundo El Maiten",
        rol_avaluo="123-45", comuna="Requínoa", superficie_escritura_ha=47.5,
        inscripcion=InscripcionCBR(fojas="1240", numero="890", anio=2011),
    ))
    b = ExtraccionLegal(propiedad=PropiedadExtraida(
        propietario="AGRICOLA LAS ENCINAS SPA", nombre_predio="FUNDO EL MAITEN",
        rol_avaluo="123-46", comuna="REQUINOA", superficie_escritura_ha=49.2,
    ))

    fusion = fusionar_extracciones([a, b], ["escritura.pdf", "certificado.pdf"])
    campos = {c.campo for c in fusion.conflictos}

    comprobar("rol_avaluo" in campos, "detecta que el rol no coincide", str(campos))
    comprobar("superficie_escritura_ha" in campos,
              "detecta que la superficie no coincide", str(campos))
    comprobar("propietario" not in campos,
              "NO se queja de 'Agrícola' vs 'AGRICOLA' (mismo dato)", str(campos))
    comprobar("comuna" not in campos,
              "NO se queja de 'Requínoa' vs 'REQUINOA' (tildes y mayusculas)",
              str(campos))
    comprobar(fusion.datos.propiedad.rol_avaluo == "123-45",
              "conserva el valor del primer documento",
              fusion.datos.propiedad.rol_avaluo)
    comprobar(fusion.datos.propiedad.inscripcion is not None,
              "hereda la inscripcion del documento que si la traia")
    texto = str(fusion.conflictos[0])
    comprobar("escritura.pdf" in texto and "certificado.pdf" in texto,
              "el aviso dice que documento dice que", texto)
    print("      {}".format(texto))


def probar_rut_invalido_no_tumba_el_servidor():
    """Incidente real (2026-09-23): un tasador tipeo el RUT sin puntos ni
    guion ('765432103') y /asistente/generar le devolvio un 500 -el
    ValidationError de Antecedentes se escapaba sin capturar.

    Al investigar salio un segundo hallazgo, mejor que el primer arreglo:
    ese RUT ('765432103') es en realidad VALIDO -el digito verificador da
    exacto-, el validador viejo lo rechazaba solo por no tener el guion
    tipeado a mano. Ahora se CALCULA el digito (modulo 11, services/rut.py)
    en vez de solo pedir el guion: el RUT sin formato entra solo, y recien
    un digito verificador que de verdad no corresponde se rechaza."""
    print("\n[1.b] El calculo real del RUT (no solo pedir el guion)")
    import asyncio

    from fastapi import HTTPException

    from api.routers.asistente import DatosBasicos, PeticionGeneracion, _construir_informe

    # El caso real del incidente: sin puntos ni guion, pero MATEMATICAMENTE
    # correcto -tiene que aceptarse solo, sin que el tasador retipee nada.
    peticion_valida = PeticionGeneracion(
        basicos=DatosBasicos(cliente="Cliente de prueba", rut="765432103"))
    informe, _, _ = asyncio.run(_construir_informe(peticion_valida, None, None, None))
    comprobar(informe.antecedentes.rut == "76.543.210-3",
              "el RUT sin formato se acepta y queda formateado con puntos y guion",
              informe.antecedentes.rut)

    # Un digito verificador que de verdad no corresponde SI se rechaza -sin
    # tumbar el servidor.
    peticion_invalida = PeticionGeneracion(
        basicos=DatosBasicos(cliente="Cliente de prueba", rut="76543210-4"))
    try:
        asyncio.run(_construir_informe(peticion_invalida, None, None, None))
        comprobar(False, "debia rechazar un digito verificador que no corresponde")
    except HTTPException as error:
        comprobar(error.status_code == 422,
                  "responde 422, no 500", error.status_code)
        comprobar("digito verificador" in error.detail,
                  "el mensaje explica que el digito no corresponde", error.detail)
    except Exception as error:  # noqa: BLE001
        comprobar(False, "se escapó sin capturar: {}: {}"
                  .format(type(error).__name__, error))


# ------------------------------------------------ integracion completa

def probar_flujo(cli, tmp):
    print("\n[2] Analisis: dos documentos contradictorios + dos fotos")

    d1 = pdf(ESCRITURA, Path(tmp) / "escritura.pdf")
    d2 = pdf(CERTIFICADO, Path(tmp) / "certificado_avaluo.pdf")
    f1 = foto(Path(tmp) / "cerezos.png", (45, 115, 50))
    f2 = foto(Path(tmp) / "nogales.png", (80, 105, 45))

    with open(d1, "rb") as a, open(d2, "rb") as b, \
         open(f1, "rb") as c, open(f2, "rb") as d:
        respuesta = cli.post(
            "/asistente/analizar",
            data={
                "cliente": "Agrícola Las Encinas SpA",
                "rut": "76.543.210-3",
                "ejecutivo": "M. Fuentes",
                "sucursal": "Rancagua",
                "fecha_visita": "2026-05-14",
                "agronomo_informante": "P. Aravena",
            },
            files=[
                ("documentos", ("escritura.pdf", a, "application/pdf")),
                ("documentos", ("certificado_avaluo.pdf", b, "application/pdf")),
                ("fotos", ("cerezos.png", c, "image/png")),
                ("fotos", ("nogales.png", d, "image/png")),
            ],
        )

    if respuesta.status_code == 503:
        comprobar(False, "sin cuota de IA ahora mismo", respuesta.text[:200])
        return None
    comprobar(respuesta.status_code == 200, "el analisis responde 200",
              respuesta.text[:300])
    borrador = respuesta.json()

    prop = borrador["legal"]["propiedad"]
    print("      predio={!r} rol={!r} comuna={!r}".format(
        prop.get("nombre_predio"), prop.get("rol_avaluo"), prop.get("comuna")))
    for paso in borrador["trazabilidad"]:
        print("      {} -> {} ({}s)".format(paso["paso"], paso["modelo"],
                                            paso["duracion_s"]))

    comprobar(prop.get("nombre_predio") is not None, "identifica el predio")
    comprobar(len(borrador["legal"]["derechos_agua"]) >= 2,
              "junta los derechos de agua de AMBOS documentos (canal + pozo)",
              str(len(borrador["legal"]["derechos_agua"])))
    comprobar(len(borrador["fotos"]) == 2, "titula las dos fotos",
              str(len(borrador["fotos"])))
    if borrador["fotos"]:
        comprobar(all((f.get("titulo") or "").strip() for f in borrador["fotos"]),
                  "cada foto trae un titulo propuesto")
        comprobar(all("descripcion" not in f for f in borrador["fotos"]),
                  "y ninguna trae descripcion tecnica: la ve el agronomo")
        print("      foto 1: {}".format(borrador["fotos"][0]["titulo"]))

    print("\n[3] Deteccion de contradicciones entre los documentos")
    for conflicto in borrador["conflictos"]:
        print("      ! {}".format(conflicto))
    unidos = " | ".join(borrador["conflictos"]).lower()
    comprobar(any("rol" in c.lower() for c in borrador["conflictos"]),
              "avisa de que el rol de avaluo no coincide", unidos)
    return borrador


def probar_generacion(cli, borrador):
    print("\n[4] Generacion del Excel con el borrador revisado")

    # El tasador corrige: se queda con el rol de la escritura.
    borrador["legal"]["propiedad"]["rol_avaluo"] = "123-45"

    peticion = {
        "sesion": borrador["sesion"],
        "basicos": borrador["basicos"],
        "legal": borrador["legal"],
        "fotos": borrador["fotos"],
        "galeria_fotos": True,
    }
    respuesta = cli.post("/asistente/generar", json=peticion)
    comprobar(respuesta.status_code == 200, "la generacion responde 200",
              respuesta.text[:400])
    if respuesta.status_code != 200:
        return
    salida = respuesta.json()
    print("      {} celdas, {} imagenes, {} avisos".format(
        salida["celdas_escritas"], salida["imagenes_insertadas"],
        len(salida["avisos"])))
    for aviso in salida["avisos"]:
        print("      aviso: {}".format(aviso))

    comprobar(salida["imagenes_insertadas"] == 2,
              "las dos fotos entraron en el libro",
              str(salida["imagenes_insertadas"]))

    descarga = cli.get(salida["descarga"])
    comprobar(descarga.status_code == 200, "se descarga el archivo")

    print("\n[5] El .xlsx tiene los datos en las celdas del formato final de Santander")
    libro = load_workbook(io.BytesIO(descarga.content))
    portada = libro["Portada"]
    recursos = libro["Recursos"]

    comprobar(portada["B10"].value == "Agrícola Las Encinas SpA",
              "el cliente esta en Portada!B10", repr(portada["B10"].value))
    # Los datos legales van lote por lote: la propiedad de la escritura es el
    # primer lote de DATOS LEGALES TERRENOS, con la inscripcion en columnas.
    comprobar(recursos["E8"].value == "123-45",
              "el rol corregido por el tasador llego al lote (Recursos!E8)",
              repr(recursos["E8"].value))
    comprobar((recursos["G8"].value or "").lower().startswith("requ"),
              "la comuna del lote (Recursos!G8)", repr(recursos["G8"].value))
    comprobar("1240" in str(recursos["H8"].value) and "890" in str(recursos["I8"].value),
              "fojas y numero de la inscripcion, cada uno en su columna",
              repr((recursos["H8"].value, recursos["I8"].value)))
    aguas = next(c.row for c in recursos["A"] if c.value == "TOTAL AGUAS DISPONIBLES ")
    filas_agua = [f for f in range(aguas - 4, aguas)
                  if recursos.cell(f, 2).value or recursos.cell(f, 3).value]
    comprobar(len(filas_agua) >= 2, "los derechos de agua de los dos documentos ocupan la tabla",
              str(filas_agua))

    anexo4 = libro["Anexo 4 Imágenes"]
    comprobar(anexo4["A4"].value is not None,
              "las fotos van al Anexo 4 con su titulo debajo", repr(anexo4["A4"].value))

    # La portada lee el EBITDA de la proyeccion, en M$, del flujo: formula viva.
    b36 = str(portada["B36"].value or "")
    comprobar(b36.startswith("='Flujo proyectado '!F"),
              "la portada lee el EBITDA del flujo", repr(b36))

    # No dejar el archivo de prueba en la carpeta de entregables.
    generado = ajustes.ruta_salida / salida["archivo"]
    if generado.exists():
        try:
            generado.unlink()
        except OSError:
            pass


def main():
    print("=" * 70)
    print("PRUEBA DEL FLUJO ASISTIDO (lo que haria el tasador)")
    print("=" * 70)

    probar_fusion_offline()
    probar_rut_invalido_no_tumba_el_servidor()

    with httpx.Client(base_url=URL, timeout=300.0) as cli:
        try:
            cli.get("/salud")
        except httpx.ConnectError:
            print("\nNo hay servidor en {}. Levantalo con:".format(URL))
            print("  python -m uvicorn main:app --port 8000")
            return 1

        # El asistente exige sesion desde que hay cuentas. Las credenciales van
        # por el entorno, nunca escritas aqui: el archivo se versiona.
        rut = os.environ.get("FITO_RUT", "")
        clave = os.environ.get("FITO_CLAVE", "")
        if not rut or not clave:
            print("\nHace falta una cuenta: set FITO_RUT y FITO_CLAVE.")
            return 1
        entrada = cli.post("/acceso/entrar", json={"rut": rut, "contrasena": clave})
        if entrada.status_code != 200:
            print("\nNo se pudo entrar: {}".format(entrada.text[:90]))
            return 1
        print("  sesion iniciada como {}".format(entrada.json()["nombre"]))

        with tempfile.TemporaryDirectory() as tmp:
            borrador = probar_flujo(cli, tmp)
            if borrador:
                probar_generacion(cli, borrador)

    print("\n" + "=" * 70)
    if fallos:
        print("{}/{} comprobaciones FALLARON:".format(len(fallos), comprobaciones))
        for f in fallos:
            print("  - {}".format(f))
        return 1
    print("{}/{} comprobaciones OK".format(comprobaciones, comprobaciones))
    return 0


if __name__ == "__main__":
    sys.exit(main())
