"""Prueba de integracion contra la API en marcha.

Requiere el servidor levantado:
    python -m uvicorn main:app --port 8000

Uso:
    python scripts/test_api.py [--url http://127.0.0.1:8000] [--sin-ia]

Los endpoints /ia hacen llamadas reales a Gemini y consumen cuota. Con --sin-ia
se prueba solo el 80% determinista.
"""

import argparse
import io
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import httpx  # noqa: E402

fallos = []
comprobaciones = 0


def _limpiar(nombre):
    """Borra un archivo generado por la prueba, si sigue ahi."""
    from config.settings import ajustes

    ruta = ajustes.ruta_salida / nombre
    if ruta.exists():
        try:
            ruta.unlink()
        except OSError:
            pass


def comprobar(condicion, descripcion, detalle=""):
    global comprobaciones
    comprobaciones += 1
    if condicion:
        print("  OK   {}".format(descripcion))
    else:
        print("  FALLA {} {}".format(descripcion, detalle))
        fallos.append(descripcion)


# --------------------------------------------------------- material falso

ESCRITURA = """
REPUBLICA DE CHILE
NOTARIA DE RANCAGUA

COMPRAVENTA
AGRICOLA LAS ENCINAS SpA

En Rancagua, a doce de marzo de dos mil once, ante mi, JORGE ELIAS MUNOZ
VALDES, Notario Publico Titular de la Primera Notaria de Rancagua, comparece
don PEDRO ANTONIO ROJAS CACERES, cedula nacional de identidad numero
9.876.543-2, quien vende, cede y transfiere a AGRICOLA LAS ENCINAS SpA, RUT
76.543.210-K, el inmueble denominado FUNDO EL MAITEN, ubicado en la comuna de
Requinoa, Provincia de Cachapoal, Region del Libertador General Bernardo
O'Higgins, de una superficie aproximada de CUARENTA Y SIETE COMA CINCO
HECTAREAS (47,5 has), Rol de Avaluo numero 123-45 de la citada comuna.

DESLINDES: Norte, con Canal Requinoa; Sur, con propiedad de don Luis Herrera;
Oriente, con Camino Publico a Los Lirios; Poniente, con Estero Las Cabras.

TITULOS: El inmueble se encuentra inscrito a fojas mil doscientos cuarenta
numero ochocientos noventa del Registro de Propiedad del Conservador de Bienes
Raices de Rancagua correspondiente al ano dos mil once.

DERECHOS DE APROVECHAMIENTO DE AGUAS: Se transfieren conjuntamente doce coma
cinco acciones del Canal Requinoa, equivalentes a un caudal de treinta litros
por segundo, de naturaleza consuntiva y ejercicio permanente y continuo,
inscritas a fojas trescientos veinte numero ciento quince del Registro de
Propiedad de Aguas del Conservador de Bienes Raices de Rancagua del ano dos mil
doce. Adicionalmente se transfiere un derecho sobre pozo profundo ubicado en el
sector norte del predio, de ocho litros por segundo.

GRAVAMENES: El inmueble se encuentra afecto a hipoteca de primer grado a favor
del Banco Santander Chile, inscrita a fojas dos mil cien numero mil cuarenta del
Registro de Hipotecas del ano dos mil once. Existe ademas servidumbre de transito
a favor del predio colindante del sur.
"""


def pdf_con_texto(texto, ruta):
    """PDF con capa de texto: ejercita la via barata (Markdown)."""
    import pymupdf

    doc = pymupdf.open()
    for i in range(0, len(texto), 1800):
        pagina = doc.new_page()
        pagina.insert_textbox(
            pymupdf.Rect(50, 50, 545, 780), texto[i:i + 1800], fontsize=10
        )
    doc.save(ruta)
    doc.close()
    return ruta


def pdf_escaneado(texto, ruta):
    """PDF sin capa de texto: obliga a la via de vision."""
    import pymupdf
    from PIL import Image, ImageDraw

    imagen = Image.new("RGB", (1240, 1754), "white")
    dibujo = ImageDraw.Draw(imagen)
    y = 60
    for linea in texto.strip().splitlines()[:55]:
        dibujo.text((70, y), linea.strip(), fill="black")
        y += 28

    buffer = io.BytesIO()
    imagen.save(buffer, format="PNG")
    doc = pymupdf.open()
    pagina = doc.new_page(width=595, height=842)
    pagina.insert_image(pymupdf.Rect(0, 0, 595, 842), stream=buffer.getvalue())
    doc.save(ruta)
    doc.close()
    return ruta


def foto_huerto(ruta):
    """Imagen sintetica con forma de huerto en hilera."""
    from PIL import Image, ImageDraw

    imagen = Image.new("RGB", (1024, 768), (150, 190, 235))
    dibujo = ImageDraw.Draw(imagen)
    dibujo.rectangle([0, 430, 1024, 768], fill=(120, 95, 60))
    for fila in range(5):
        y = 450 + fila * 62
        escala = 1 + fila * 0.35
        for col in range(9):
            x = 40 + col * 118 - fila * 18
            dibujo.rectangle(
                [x + 18 * escala, y, x + 26 * escala, y + 40 * escala],
                fill=(90, 65, 40),
            )
            dibujo.ellipse(
                [x, y - 45 * escala, x + 46 * escala, y + 6 * escala],
                fill=(40 + fila * 8, 110 + fila * 10, 45),
            )
    imagen.save(ruta)
    return ruta


def informe_minimo():
    return {
        "antecedentes": {
            "cliente": "Agricola Las Encinas SpA",
            "rut": "76.543.210-3",
            "ejecutivo": "M. Fuentes",
            "fecha_visita": "2026-05-14",
            "sucursal": "Rancagua",
            "agronomo_informante": "P. Aravena",
        },
        "estructura_temporada": [
            {"especie": "Cerezo", "superficie_ha": 30.0},
            {"especie": "Nogal", "superficie_ha": 17.5},
        ],
        "estructura_productiva": {
            "plantaciones": [
                {"especie": "Cerezo", "variedad": "Santina",
                 "anio_plantacion": "2015", "superficie_ha": 30.0,
                 "estado": "Bueno", "riego": "Goteo"},
                {"especie": "Nogal", "variedad": "Chandler",
                 "anio_plantacion": "2012", "superficie_ha": 17.5,
                 "estado": "Regular", "riego": "Goteo"},
            ]
        },
        "dotacion_recursos": {
            "derechos_agua": [
                {"predio": "Fundo El Maiten", "caudal_nominal_ls": 30.0,
                 "caudal_real_ls": 45.0}
            ]
        },
    }


# ------------------------------------------------------------- las pruebas

def probar_salud(cli):
    print("\n[1] Salud del servicio")
    r = cli.get("/salud")
    comprobar(r.status_code == 200, "GET /salud responde 200", str(r.status_code))
    datos = r.json()
    comprobar(datos["estado"] == "operativo", "el servicio se declara operativo")
    comprobar(datos["excel"]["plantilla_encontrada"], "encuentra la plantilla")
    comprobar(datos["excel"]["tablas_mapeadas"] == 16, "16 tablas mapeadas",
              str(datos["excel"]["tablas_mapeadas"]))
    print("      IA: {}".format(datos["ia"].get("proveedor")))
    return datos


def probar_validacion(cli):
    print("\n[2] Validacion de informe (detecta incoherencias)")
    r = cli.post("/informes/validar", json=informe_minimo())
    comprobar(r.status_code == 200, "POST /informes/validar responde 200",
              r.text[:200])
    datos = r.json()
    comprobar(datos["plantaciones"] == 2, "cuenta las plantaciones")
    avisos = " | ".join(datos["avisos"])
    comprobar("caudal real" in avisos,
              "detecta que el caudal real supera al nominal", avisos)
    print("      avisos: {}".format(datos["avisos"]))


def probar_validacion_rechaza_basura(cli):
    print("\n[3] Pydantic rechaza datos invalidos")
    malo = informe_minimo()
    malo["antecedentes"]["rut"] = "76543210"   # sin digito verificador
    r = cli.post("/informes/validar", json=malo)
    comprobar(r.status_code == 422, "un RUT sin DV devuelve 422", str(r.status_code))
    comprobar("digito verificador" in r.text, "el mensaje explica el motivo",
              r.text[:150])


def probar_generacion(cli):
    print("\n[4] Generacion y descarga del .xlsx")
    r = cli.post("/informes/generar", json=informe_minimo())
    comprobar(r.status_code == 200, "POST /informes/generar responde 200",
              r.text[:300])
    datos = r.json()
    comprobar(datos["celdas_escritas"] > 20, "escribio celdas",
              str(datos["celdas_escritas"]))
    print("      {} celdas, {} avisos".format(
        datos["celdas_escritas"], len(datos["avisos"])))

    descarga = cli.get(datos["descarga"])
    comprobar(descarga.status_code == 200, "la descarga responde 200")
    comprobar(descarga.content[:2] == b"PK", "el archivo es un .xlsx (zip)")
    comprobar(len(descarga.content) > 20_000, "el archivo tiene cuerpo",
              "{} bytes".format(len(descarga.content)))

    fuga = cli.get("/informes/..%2F..%2F.env/descargar")
    comprobar(fuga.status_code == 404,
              "la descarga no deja escapar de la carpeta de salida",
              str(fuga.status_code))

    # La API escribe en output/ porque ahi van los entregables de verdad. Una
    # prueba no deja basura en la carpeta que el tasador revisa.
    _limpiar(datos["archivo"])


def probar_pdf_texto(cli, tmp):
    print("\n[5] IA: extraccion de PDF con capa de texto")
    ruta = pdf_con_texto(ESCRITURA, Path(tmp) / "escritura.pdf")
    with open(ruta, "rb") as fh:
        r = cli.post("/ia/documentos/extraer",
                     files={"archivo": ("escritura.pdf", fh, "application/pdf")})
    if r.status_code == 503:
        comprobar(False, "sin cuota disponible ahora mismo", r.text[:200])
        return
    comprobar(r.status_code == 200, "responde 200", r.text[:300])
    d = r.json()
    prop = d["datos"]["propiedad"]
    print("      via={} modelo={} tokens~{} ({}s)".format(
        d["via"], d["trazabilidad"]["modelo"], d["tokens_estimados"],
        d["trazabilidad"]["duracion_s"]))
    print("      propietario={!r} rol={!r} comuna={!r} sup={}".format(
        prop.get("propietario"), prop.get("rol_avaluo"),
        prop.get("comuna"), prop.get("superficie_escritura_ha")))

    comprobar(d["via"] == "texto", "uso la via barata de texto", d["via"])
    comprobar(d["tokens_estimados"] < 1200,
              "el Markdown cabe en pocos tokens", str(d["tokens_estimados"]))
    comprobar("ENCINAS" in (prop.get("propietario") or "").upper(),
              "identifica al comprador", repr(prop.get("propietario")))
    comprobar(prop.get("rol_avaluo") == "123-45", "extrae el rol de avaluo",
              repr(prop.get("rol_avaluo")))
    comprobar((prop.get("comuna") or "").lower().startswith("requ"),
              "extrae la comuna", repr(prop.get("comuna")))
    comprobar(prop.get("superficie_escritura_ha") == 47.5,
              "convierte '47,5' a 47.5 sin recalcular nada",
              repr(prop.get("superficie_escritura_ha")))
    insc = prop.get("inscripcion") or {}
    comprobar(str(insc.get("fojas") or "").strip() == "1240",
              "extrae las fojas escritas en palabras", repr(insc.get("fojas")))
    comprobar(len(d["datos"]["derechos_agua"]) >= 2,
              "encuentra los dos derechos de agua (canal y pozo)",
              str(len(d["datos"]["derechos_agua"])))
    comprobar(len(d["datos"]["gravamenes"]) >= 1, "detecta la hipoteca",
              str(d["datos"]["gravamenes"]))


def probar_pdf_escaneado(cli, tmp):
    print("\n[6] IA: PDF escaneado (sin capa de texto -> vision)")
    ruta = pdf_escaneado(ESCRITURA, Path(tmp) / "escaneado.pdf")
    with open(ruta, "rb") as fh:
        r = cli.post("/ia/documentos/extraer",
                     files={"archivo": ("escaneado.pdf", fh, "application/pdf")})
    if r.status_code == 503:
        comprobar(False, "sin cuota disponible ahora mismo", r.text[:200])
        return
    comprobar(r.status_code == 200, "responde 200", r.text[:300])
    d = r.json()
    print("      via={} modelo={}".format(d["via"], d["trazabilidad"]["modelo"]))
    comprobar(d["via"] == "vision", "detecta el escaneo y cambia a vision", d["via"])
    comprobar(any("vision" in a for a in d["advertencias"]),
              "avisa de que hay que revisar los numeros", str(d["advertencias"]))
    prop = d["datos"]["propiedad"]
    comprobar(prop.get("comuna") is not None, "aun asi extrae datos",
              repr(prop.get("comuna")))


def probar_fotografias(cli, tmp):
    print("\n[7] IA: titulo de fotografias (sin descripcion tecnica)")
    ruta = foto_huerto(Path(tmp) / "huerto.png")
    with open(ruta, "rb") as fh:
        r = cli.post(
            "/ia/fotografias/titulos",
            files={"archivos": ("huerto.png", fh, "image/png")},
            data={"contexto": "Cerezo variedad Santina, plantacion 2015, riego por goteo"},
        )
    if r.status_code == 503:
        comprobar(False, "sin cuota disponible ahora mismo", r.text[:200])
        return
    comprobar(r.status_code == 200, "responde 200", r.text[:300])
    d = r.json()
    titulos = d["resultado"]["titulos"]
    comprobar(len(titulos) == 1, "devuelve un titulo por imagen", str(len(titulos)))
    if titulos:
        print("      titulo: {}".format(titulos[0]["titulo"]))
        comprobar(bool(titulos[0]["titulo"].strip()), "el titulo tiene contenido")
        comprobar(len(titulos[0]["titulo"]) <= 120,
                  "y es un rotulo corto, no un parrafo", len(titulos[0]["titulo"]))
        comprobar("descripcion" not in titulos[0],
                  "sin descripcion tecnica (la ve el agronomo)", sorted(titulos[0]))


def probar_opinion(cli):
    print("\n[8] IA: redaccion de opinion profesional")
    r = cli.post("/ia/opinion/redactar", json={
        "apartado": "calidad de los activos productivos",
        "hechos": (
            "Huerto de cerezo Santina de 30 ha plantado en 2015, riego por goteo, "
            "estado fitosanitario bueno. Nogal Chandler de 17,5 ha plantado en "
            "2012, estado regular por falta de poda. Derechos de agua por 30 l/s "
            "inscritos. Bodega de 400 m2 de albanileria en buen estado."
        ),
    })
    if r.status_code == 503:
        comprobar(False, "sin cuota disponible ahora mismo", r.text[:200])
        return
    comprobar(r.status_code == 200, "responde 200", r.text[:300])
    d = r.json()
    texto = d["resultado"]["texto"]
    print("      calificacion: {}".format(d["resultado"]["calificacion_sugerida"]))
    print("      texto: {}...".format(texto[:130]))
    comprobar(len(texto) > 80, "produce un parrafo utilizable")
    comprobar(d["trazabilidad"]["modelo"], "informa el modelo que lo redacto")


def probar_estado_cadena(cli):
    print("\n[9] Estado de la cadena tras el trabajo")
    r = cli.get("/ia/estado")
    if r.status_code != 200:
        comprobar(False, "GET /ia/estado responde 200", r.text[:200])
        return
    d = r.json()
    comprobar(d["disponibles_ahora"] >= 1, "queda al menos un candidato vivo",
              str(d["disponibles_ahora"]))
    for c in d["candidatos"]:
        marca = {"disponible": "  ", "en_cuarentena": "??", "descartado": "XX"}
        print("      {} {:24} exitos={} fallos={} espera={}s".format(
            marca.get(c["estado"], "  "), c["modelo"], c["exitos"],
            c["fallos"], c["espera_s"]))


def entrar(cli, rut: str, clave: str) -> bool:
    """Inicia sesion antes de probar lo demas.

    Las credenciales llegan por argumento o por el entorno, nunca escritas
    aqui: este archivo se lee, se comparte y se versiona.
    """
    if not rut or not clave:
        print("\nHace falta una cuenta para probar la API protegida:")
        print("  set FITO_RUT=12.345.678-9  &  set FITO_CLAVE=...")
        print("  (o pasarlos con --rut y --clave)")
        return False
    r = cli.post("/acceso/entrar", json={"rut": rut, "contrasena": clave})
    if r.status_code != 200:
        print("\nNo se pudo entrar: {} {}".format(r.status_code, r.text[:90]))
        return False
    print("\n  sesion iniciada como {}".format(r.json().get("nombre")))
    return True


def main():
    parser = argparse.ArgumentParser()
    # FITO_URL, igual que test_persistencia: en la carpeta de trabajo se
    # prueba contra el servidor de dev (8001), nunca contra el de la casa.
    parser.add_argument("--url", default=os.environ.get("FITO_URL", "http://127.0.0.1:8000"))
    parser.add_argument("--rut", default=os.environ.get("FITO_RUT", ""),
                        help="RUT con que entrar (o la variable FITO_RUT)")
    parser.add_argument("--clave", default=os.environ.get("FITO_CLAVE", ""),
                        help="contrasena (o la variable FITO_CLAVE). Nunca va "
                             "escrita en el codigo: este archivo se versiona")
    parser.add_argument("--sin-ia", action="store_true")
    args = parser.parse_args()

    print("=" * 70)
    print("PRUEBA DE INTEGRACION DE LA API  ({})".format(args.url))
    print("=" * 70)

    with httpx.Client(base_url=args.url, timeout=300.0) as cli:
        try:
            salud = probar_salud(cli)
        except httpx.ConnectError:
            print("\nNo hay servidor en {}. Levantalo con:".format(args.url))
            print("  python -m uvicorn main:app --port 8000")
            return 1

        # Desde que hay cuentas, casi todo exige sesion. El cliente de httpx
        # guarda la cookie solo, asi que basta con entrar una vez aqui.
        if not entrar(cli, args.rut, args.clave):
            return 1

        probar_validacion(cli)
        probar_validacion_rechaza_basura(cli)
        probar_generacion(cli)

        hay_ia = salud["ia"].get("proveedor") is not None
        if args.sin_ia:
            print("\n(--sin-ia: se omiten los endpoints de IA)")
        elif not hay_ia:
            print("\n(sin GEMINI_API_KEYS: se omiten los endpoints de IA)")
        else:
            with tempfile.TemporaryDirectory() as tmp:
                probar_pdf_texto(cli, tmp)
                probar_pdf_escaneado(cli, tmp)
                probar_fotografias(cli, tmp)
            probar_opinion(cli)
            probar_estado_cadena(cli)

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
