"""Conservar los dibujos de una plantilla -cuadros de texto, formas, grupos,
imagenes- en un .xlsx que se escribio con openpyxl.

El problema (2026-10-03, el usuario: *"el actualizador de tasaciones no esta
usando este formato de rabo... identico a ese"*): openpyxl NO lee formas ni
cuadros de texto de un archivo existente, y al guardar los pierde sin avisar.
El formato real de RaboFinance tiene 15 cuadros de texto -observaciones,
descripcion del predio, el analisis estadistico de las referencias- y el
informe generado salia sin ninguno.

La solucion no toca la escritura de celdas (sigue siendo openpyxl, con todo
lo que el motor ya sabe hacer): despues de guardar, `injertar_dibujos()`
toma el dibujo COMPLETO de cada hoja de la plantilla -tal cual esta en su
XML, con sus formas, grupos, estilos y anclajes- y lo pone en el archivo
generado en lugar del dibujo pobre que escribio openpyxl. En el camino:

- llena el texto de los cuadros que se le pidan, conservando el formato del
  cuadro (letra, tamaño, color, borde);
- pone una imagen nueva en un hueco de imagen, recortada a la proporcion del
  hueco para que no se deforme, o saca el hueco si no hay imagen;
- corre hacia abajo los anclajes que quedaron bajo filas insertadas.

Trabaja sobre el paquete (zip + XML) con la biblioteca estandar, sin
dependencias nuevas.
"""

from __future__ import annotations

import copy
import io
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple, Union
from xml.etree import ElementTree as ET

NS = {
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
REL_DIBUJO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
REL_IMAGEN = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
CT_DIBUJO = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_IMAGEN = {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg",
             "gif": "image/gif", "emf": "image/x-emf", "wmf": "image/x-wmf"}

#: Lo que va DESPUES de <drawing> en una hoja (orden de CT_Worksheet): el
#: elemento nuevo se mete antes del primero de estos que exista.
_DESPUES_DE_DRAWING = ("legacyDrawing", "legacyDrawingHF", "drawingHF", "picture",
                       "oleObjects", "controls", "webPublishItems", "tableParts",
                       "extLst")

Hueco = Tuple[str, str]  # (nombre real de la hoja, nombre de la forma)


# ------------------------------------------------------------ el paquete

def _leer_paquete(ruta) -> Dict[str, bytes]:
    if hasattr(ruta, "seek"):
        ruta.seek(0)
    with zipfile.ZipFile(ruta) as z:
        return {nombre: z.read(nombre) for nombre in z.namelist()}


def _escribir_paquete(partes: Dict[str, bytes], destino: Union[str, Path]) -> None:
    orden = ["[Content_Types].xml"] + [n for n in partes if n != "[Content_Types].xml"]
    temporal = io.BytesIO()
    with zipfile.ZipFile(temporal, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre in orden:
            z.writestr(nombre, partes[nombre])
    Path(destino).write_bytes(temporal.getvalue())


def _ruta_rels(parte: str) -> str:
    carpeta, nombre = posixpath.split(parte)
    return posixpath.join(carpeta, "_rels", nombre + ".rels")


def _resolver(parte: str, destino: str) -> str:
    if destino.startswith("/"):
        return destino.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(parte), destino))


def _relativa(desde_parte: str, hacia_parte: str) -> str:
    return posixpath.relpath(hacia_parte, posixpath.dirname(desde_parte))


def _rels(partes: Dict[str, bytes], parte: str) -> List[Dict[str, str]]:
    datos = partes.get(_ruta_rels(parte))
    if not datos:
        return []
    raiz = ET.fromstring(datos)
    return [dict(r.attrib) for r in raiz.findall("rel:Relationship", NS)]


def _rels_xml(relaciones: Sequence[Dict[str, str]]) -> bytes:
    filas = "".join(
        '<Relationship {}/>'.format(" ".join(
            '{}="{}"'.format(k, _escapar(v)) for k, v in r.items()))
        for r in relaciones)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="{}">{}</Relationships>'.format(NS["rel"], filas)
            ).encode("utf-8")


def _escapar(texto: str) -> str:
    return (texto.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _hojas(partes: Dict[str, bytes]) -> Dict[str, str]:
    """nombre de la hoja -> parte del paquete (xl/worksheets/sheetN.xml)."""
    libro = ET.fromstring(partes["xl/workbook.xml"])
    ids = {r["Id"]: _resolver("xl/workbook.xml", r["Target"])
           for r in _rels(partes, "xl/workbook.xml")}
    hojas = libro.find("x:sheets", NS)
    return {h.get("name"): ids[h.get("{%s}id" % NS["r"])] for h in hojas}


def _dibujo_de(partes: Dict[str, bytes], hoja: str) -> Optional[Tuple[str, str]]:
    """(id de la relacion, parte del dibujo) de una hoja, o None."""
    for r in _rels(partes, hoja):
        if r.get("Type") == REL_DIBUJO:
            return r["Id"], _resolver(hoja, r["Target"])
    return None


# ------------------------------------------------------------ el XML

def _parsear(datos: bytes) -> ET.Element:
    """Registra los prefijos del documento antes de leerlo, para que al
    escribirlo de vuelta sigan siendo xdr:, a:, a16:... y no ns0:, ns1:."""
    for _, (prefijo, uri) in ET.iterparse(io.BytesIO(datos), events=("start-ns",)):
        if prefijo:
            ET.register_namespace(prefijo, uri)
    return ET.fromstring(datos)


def _serializar(raiz: ET.Element) -> bytes:
    return (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            + ET.tostring(raiz, encoding="utf-8"))


def _q(etiqueta: str) -> str:
    prefijo, local = etiqueta.split(":")
    return "{%s}%s" % (NS[prefijo], local)


def _nombre_forma(elemento: ET.Element) -> Optional[str]:
    for hijo in elemento:
        if hijo.tag.startswith("{%s}nv" % NS["xdr"]):
            cnv = hijo.find("xdr:cNvPr", NS)
            if cnv is not None:
                return cnv.get("name")
    return None


def _buscar_forma(raiz: ET.Element, nombre: str):
    """(ancla, padre, forma) de la forma con ese nombre -incluso dentro de un
    grupo-, o None."""
    def recorrer(ancla, padre):
        for hijo in list(padre):
            local = hijo.tag.split("}")[-1]
            if local in ("sp", "pic", "grpSp", "graphicFrame", "cxnSp"):
                if _nombre_forma(hijo) == nombre:
                    return ancla, padre, hijo
                if local == "grpSp":
                    encontrado = recorrer(ancla, hijo)
                    if encontrado:
                        return encontrado
        return None
    for ancla in list(raiz):
        encontrado = recorrer(ancla, ancla)
        if encontrado:
            return encontrado
    return None


def nombres_de_formas(ruta: Union[str, Path]) -> Dict[str, List[Tuple[str, str]]]:
    """Por hoja, las formas que trae su dibujo: [(tipo, nombre)]. Sirve para
    validar el mapping contra la plantilla sin abrir Excel."""
    partes = _leer_paquete(ruta)
    salida: Dict[str, List[Tuple[str, str]]] = {}
    for nombre_hoja, hoja in _hojas(partes).items():
        dibujo = _dibujo_de(partes, hoja)
        if not dibujo:
            continue
        raiz = _parsear(partes[dibujo[1]])
        lista = []
        for elemento in raiz.iter():
            local = elemento.tag.split("}")[-1]
            if local in ("sp", "pic", "grpSp", "graphicFrame", "cxnSp"):
                lista.append((local, _nombre_forma(elemento) or ""))
        salida[nombre_hoja] = lista
    return salida


def texto_de_forma(ruta: Union[str, Path], hoja: str, forma: str) -> Optional[str]:
    """El texto de un cuadro, con un salto de linea por parrafo."""
    return leer_textos(ruta, [(hoja, forma)]).get((hoja, forma))


def leer_textos(ruta, pedidos: Iterable[Hueco]) -> Dict[Hueco, str]:
    """Los textos de varios cuadros de una vez -leer el paquete una sola
    vez-. `ruta` puede ser un archivo abierto (lo que llega por la API). Un
    cuadro que no existe, o que esta vacio, no aparece."""
    partes = _leer_paquete(ruta)
    hojas = _hojas(partes)
    salida: Dict[Hueco, str] = {}
    raices: Dict[str, Optional[ET.Element]] = {}
    for hoja, forma in pedidos:
        if hoja not in raices:
            dibujo = _dibujo_de(partes, hojas[hoja]) if hoja in hojas else None
            raices[hoja] = _parsear(partes[dibujo[1]]) if dibujo else None
        raiz = raices[hoja]
        encontrado = _buscar_forma(raiz, forma) if raiz is not None else None
        texto = _texto(encontrado[2]) if encontrado else None
        if texto and texto.strip():
            salida[(hoja, forma)] = texto.strip()
    return salida


def _texto(forma: ET.Element) -> Optional[str]:
    cuerpo = forma.find("xdr:txBody", NS)
    if cuerpo is None:
        return None
    return "\n".join("".join(t.text or "" for t in p.iter(_q("a:t")))
                     for p in cuerpo.findall("a:p", NS))


def _poner_texto(forma: ET.Element, texto: str) -> bool:
    """Reemplaza los parrafos del cuadro por `texto` (un parrafo por linea),
    con el estilo del primer tramo de texto que traia el cuadro -o, si ya
    venia vacio, el que guarda su marca de fin de parrafo-."""
    cuerpo = forma.find("xdr:txBody", NS)
    if cuerpo is None:
        return False
    parrafos = cuerpo.findall("a:p", NS)
    estilo_parrafo = next((p.find("a:pPr", NS) for p in parrafos
                           if p.find("a:pPr", NS) is not None), None)
    estilo_texto = None
    for p in parrafos:
        r = p.find("a:r/a:rPr", NS)
        if r is not None:
            estilo_texto = r
            break
    if estilo_texto is None:
        estilo_texto = next((p.find("a:endParaRPr", NS) for p in parrafos
                             if p.find("a:endParaRPr", NS) is not None), None)
    for p in parrafos:
        cuerpo.remove(p)

    for linea in (texto or "").replace("\r\n", "\n").split("\n"):
        p = ET.SubElement(cuerpo, _q("a:p"))
        if estilo_parrafo is not None:
            p.append(copy.deepcopy(estilo_parrafo))
        if linea:
            r = ET.SubElement(p, _q("a:r"))
            if estilo_texto is not None:
                rpr = copy.deepcopy(estilo_texto)
                rpr.tag = _q("a:rPr")
                r.append(rpr)
            ET.SubElement(r, _q("a:t")).text = linea
        if estilo_texto is not None:
            fin = copy.deepcopy(estilo_texto)
            fin.tag = _q("a:endParaRPr")
            p.append(fin)
    return True


def vaciar_texto(forma: ET.Element) -> None:
    """Deja el cuadro sin texto pero con su estilo guardado en la marca de fin
    de parrafo, para que el proximo texto que se le ponga salga igual."""
    _poner_texto(forma, "")


_FORMAS = ("sp", "pic", "grpSp", "graphicFrame", "cxnSp")


def _sacar_grupos_vacios(raiz: ET.Element) -> None:
    """Un grupo que se quedo sin formas -se saco la unica imagen que tenia-
    no puede quedar en el dibujo: se saca, y con el su ancla si era lo unico
    que colgaba de ella. Se repite porque un grupo vacio puede vaciar al que
    lo contiene."""
    cambio = True
    while cambio:
        cambio = False
        padres = {hijo: padre for padre in raiz.iter() for hijo in padre}
        for grupo in list(raiz.iter(_q("xdr:grpSp"))):
            if any(h.tag.split("}")[-1] in _FORMAS for h in grupo):
                continue
            padre = padres.get(grupo)
            if padre is None:
                continue
            padre.remove(grupo)
            if padre in list(raiz) and not any(
                    h.tag.split("}")[-1] in _FORMAS for h in padre):
                raiz.remove(padre)
            cambio = True
            break


def _desplazar_anclas(raiz: ET.Element, desde: int, cantidad: int) -> None:
    """`desde` es la fila 1-based donde se insertaron `cantidad` filas: todo
    anclaje en esa fila o mas abajo baja con ellas (las filas del anclaje
    son 0-based)."""
    corte = desde - 1
    for ancla in raiz:
        for punta in ("xdr:from", "xdr:to"):
            nodo = ancla.find(punta, NS)
            fila = nodo.find("xdr:row", NS) if nodo is not None else None
            if fila is not None and int(fila.text) >= corte:
                fila.text = str(int(fila.text) + cantidad)


def _proporcion_del_hueco(forma: ET.Element) -> Optional[float]:
    ext = forma.find("xdr:spPr/a:xfrm/a:ext", NS)
    if ext is None:
        return None
    ancho, alto = int(ext.get("cx", "0")), int(ext.get("cy", "0"))
    return ancho / alto if ancho and alto else None


def preparar_imagen(datos: bytes, proporcion: Optional[float],
                    lado_mayor: int = 1600) -> Tuple[bytes, str]:
    """Recorta al centro la imagen a la proporcion del hueco -una foto
    estirada a otra proporcion se nota de inmediato en un informe- y la deja
    en JPEG (o PNG si trae transparencia), sin pasar de `lado_mayor`."""
    from PIL import Image, ImageOps

    imagen = Image.open(io.BytesIO(datos))
    imagen = ImageOps.exif_transpose(imagen)
    if proporcion:
        ancho, alto = imagen.size
        if ancho / alto > proporcion:
            nuevo = int(round(alto * proporcion))
            izquierda = (ancho - nuevo) // 2
            imagen = imagen.crop((izquierda, 0, izquierda + nuevo, alto))
        else:
            nuevo = int(round(ancho / proporcion))
            arriba = (alto - nuevo) // 2
            imagen = imagen.crop((0, arriba, ancho, arriba + nuevo))
    if max(imagen.size) > lado_mayor:
        imagen.thumbnail((lado_mayor, lado_mayor))
    salida = io.BytesIO()
    if imagen.mode in ("RGBA", "LA", "P") and "transparency" in imagen.info or imagen.mode == "RGBA":
        imagen.save(salida, "PNG", optimize=True)
        return salida.getvalue(), "png"
    imagen.convert("RGB").save(salida, "JPEG", quality=85, optimize=True)
    return salida.getvalue(), "jpeg"


def _ids_usados(raiz: ET.Element) -> set:
    usados = set()
    for elemento in raiz.iter():
        for clave, valor in elemento.attrib.items():
            if clave.startswith("{%s}" % NS["r"]):
                usados.add(valor)
    return usados


def _insertar_drawing(hoja_xml: str, rid: str) -> str:
    elemento = '<drawing xmlns:r="{}" r:id="{}"/>'.format(NS["r"], rid)
    for etiqueta in _DESPUES_DE_DRAWING:
        m = re.search(r"<(?:\w+:)?%s[\s/>]" % etiqueta, hoja_xml)
        if m:
            return hoja_xml[:m.start()] + elemento + hoja_xml[m.start():]
    cierre = hoja_xml.rfind("</")
    return hoja_xml[:cierre] + elemento + hoja_xml[cierre:]


def _nuevo_nombre(partes: Dict[str, bytes], carpeta: str, base: str, extension: str) -> str:
    n = 1
    while True:
        nombre = "{}/{}{}.{}".format(carpeta, base, n, extension)
        if nombre not in partes:
            return nombre
        n += 1


# ------------------------------------------------------------ publico

def injertar_dibujos(
    plantilla: Union[str, Path],
    generado: Union[str, Path],
    *,
    textos: Optional[Dict[Hueco, Optional[str]]] = None,
    imagenes: Optional[Dict[Hueco, Optional[bytes]]] = None,
    filas_insertadas: Optional[Dict[str, Iterable[Tuple[int, int]]]] = None,
    hojas_sin_dibujo: Iterable[str] = (),
    filas_de: Optional[Dict[str, Callable[[int], List[int]]]] = None,
) -> List[str]:
    """Pone en `generado` (sobrescribiendolo) el dibujo de cada hoja de
    `plantilla`, con los textos e imagenes indicados.

    - `textos`: {(hoja, forma): texto}. None o "" deja el cuadro vacio.
    - `imagenes`: {(hoja, forma): bytes o None}. Bytes reemplaza la imagen
      del hueco; None saca el hueco. Un hueco que no se nombra queda como
      esta en la plantilla (el logo, la firma).
    - `filas_insertadas`: {hoja: [(fila, cantidad), ...]} en el orden en que
      se insertaron, para correr los anclajes.
    - `hojas_sin_dibujo`: hojas que no deben recibir el dibujo (por ejemplo,
      una hoja que se saco del informe).
    - `filas_de`: {hoja: funcion fila_de_la_plantilla -> [filas finales]}
      para las hojas cuyas filas no solo se corrieron sino que se REPITIERON
      (services/disposicion.py): una fila del molde puede quedar en varias.
      Con ella se reponen el alto automatico y el estilo de las celdas
      combinadas en cada copia; sin ella, una nota combinada del molde se
      estampaba sobre la fila de datos que cayo en su numero (2026-10-14).

    Devuelve avisos: una forma pedida que no existe en la plantilla no es un
    error que deba tumbar la generacion, pero tiene que quedar dicho.
    """
    avisos: List[str] = []
    textos = dict(textos or {})
    imagenes = dict(imagenes or {})
    filas_insertadas = {k: list(v) for k, v in (filas_insertadas or {}).items()}
    excluidas = set(hojas_sin_dibujo)

    origen = _leer_paquete(plantilla)
    destino = _leer_paquete(generado)
    hojas_origen = _hojas(origen)
    hojas_destino = _hojas(destino)
    tipos = _parsear(destino["[Content_Types].xml"])
    nuevos_tipos: Dict[str, str] = {}
    partes_borradas: List[str] = []

    pedidos = {h for h, _ in textos} | {h for h, _ in imagenes}

    for nombre_hoja, hoja_origen in hojas_origen.items():
        dibujo_origen = _dibujo_de(origen, hoja_origen)
        if dibujo_origen is None or nombre_hoja in excluidas:
            continue
        hoja_destino = hojas_destino.get(nombre_hoja)
        if hoja_destino is None:
            continue  # la hoja se saco del informe
        pedidos.discard(nombre_hoja)

        raiz = _parsear(origen[dibujo_origen[1]])

        for (hoja, forma), texto in textos.items():
            if hoja != nombre_hoja:
                continue
            encontrado = _buscar_forma(raiz, forma)
            if not encontrado or not _poner_texto(encontrado[2], texto or ""):
                avisos.append("{}: no hay un cuadro de texto llamado '{}' en la "
                              "plantilla".format(nombre_hoja.strip(), forma))

        reemplazos: Dict[str, bytes] = {}
        # Los ids de relacion que ya trae el dibujo de la plantilla: uno nuevo
        # no puede repetir ninguno. Una plantilla que ya paso por aqui (las de
        # Santander y Rabobank se derivaron asi) trae sus propios 'rIdFg1'..;
        # numerar desde 1 de nuevo dejaba dos relaciones con el mismo id, y
        # Excel mostraba rotas TODAS las imagenes de la hoja (2026-10-14).
        tomados = {r["Id"] for r in _rels(origen, dibujo_origen[1])}
        for (hoja, forma), datos in imagenes.items():
            if hoja != nombre_hoja:
                continue
            encontrado = _buscar_forma(raiz, forma)
            if not encontrado:
                avisos.append("{}: no hay un hueco de imagen llamado '{}' en la "
                              "plantilla".format(nombre_hoja.strip(), forma))
                continue
            ancla, padre, figura = encontrado
            if datos is None:
                if padre is ancla:
                    raiz.remove(ancla)
                else:
                    padre.remove(figura)
                continue
            blip = figura.find("xdr:blipFill/a:blip", NS)
            rid = blip.get("{%s}embed" % NS["r"]) if blip is not None else None
            if not rid:
                avisos.append("{}: '{}' no es un hueco de imagen".format(
                    nombre_hoja.strip(), forma))
                continue
            contenido, extension = preparar_imagen(datos, _proporcion_del_hueco(figura))
            # Cada hueco con su propio archivo: dos huecos pueden compartir la
            # misma imagen de muestra en la plantilla.
            n = len(reemplazos) + 1
            while "rIdFg{}".format(n) in tomados:
                n += 1
            nuevo_rid = "rIdFg{}".format(n)
            tomados.add(nuevo_rid)
            blip.set("{%s}embed" % NS["r"], nuevo_rid)
            reemplazos[nuevo_rid] = (contenido, extension)
            # El recorte ya se hizo sobre la imagen: si el hueco traia uno
            # propio, aplicarlo de nuevo cortaria la foto dos veces.
            relleno = figura.find("xdr:blipFill", NS)
            src = relleno.find("a:srcRect", NS) if relleno is not None else None
            if src is not None:
                src.attrib.clear()

        _sacar_grupos_vacios(raiz)
        for desde, cantidad in filas_insertadas.get(nombre_hoja, []):
            _desplazar_anclas(raiz, desde, cantidad)

        # -- el dibujo nuevo, con sus imagenes
        parte_dibujo = _nuevo_nombre(destino, "xl/drawings", "drawing", "xml")
        destino[parte_dibujo] = b""  # reserva el nombre
        usados = _ids_usados(raiz)
        relaciones = []
        for r in _rels(origen, dibujo_origen[1]):
            if r["Id"] not in usados:
                continue
            if r.get("Type") == REL_IMAGEN and r.get("TargetMode") != "External":
                media_origen = _resolver(dibujo_origen[1], r["Target"])
                extension = media_origen.rsplit(".", 1)[-1].lower()
                media = _nuevo_nombre(destino, "xl/media", "imagen_fg", extension)
                destino[media] = origen[media_origen]
                r = dict(r, Target=_relativa(parte_dibujo, media))
                nuevos_tipos.setdefault(extension, CT_IMAGEN.get(extension, "image/" + extension))
            relaciones.append(r)
        for rid, (contenido, extension) in reemplazos.items():
            media = _nuevo_nombre(destino, "xl/media", "imagen_fg", extension)
            destino[media] = contenido
            relaciones.append({"Id": rid, "Type": REL_IMAGEN,
                               "Target": _relativa(parte_dibujo, media)})
            nuevos_tipos.setdefault(extension, CT_IMAGEN[extension])
        destino[parte_dibujo] = _serializar(raiz)
        destino[_ruta_rels(parte_dibujo)] = _rels_xml(relaciones)
        ET.SubElement(tipos, _q("ct:Override"), {
            "PartName": "/" + parte_dibujo, "ContentType": CT_DIBUJO})

        # -- la hoja apunta al dibujo nuevo
        rels_hoja = _rels(destino, hoja_destino)
        previo = next((r for r in rels_hoja if r.get("Type") == REL_DIBUJO), None)
        if previo is not None:
            viejo = _resolver(hoja_destino, previo["Target"])
            previo["Target"] = _relativa(hoja_destino, parte_dibujo)
            partes_borradas.append(viejo)
        else:
            usados_hoja = {r["Id"] for r in rels_hoja}
            n = 1
            while "rId{}".format(n) in usados_hoja:
                n += 1
            rid = "rId{}".format(n)
            rels_hoja.append({"Id": rid, "Type": REL_DIBUJO,
                              "Target": _relativa(hoja_destino, parte_dibujo)})
            texto_hoja = destino[hoja_destino].decode("utf-8")
            destino[hoja_destino] = _insertar_drawing(texto_hoja, rid).encode("utf-8")
        destino[_ruta_rels(hoja_destino)] = _rels_xml(rels_hoja)

    for hoja in sorted(pedidos):
        avisos.append("{}: la plantilla no trae dibujo en esa hoja".format(hoja.strip()))

    for nombre_hoja, hoja_origen in hojas_origen.items():
        hoja_destino = hojas_destino.get(nombre_hoja)
        if hoja_destino is not None:
            inserciones = filas_insertadas.get(nombre_hoja, [])
            mapeo = (filas_de or {}).get(nombre_hoja)
            generada = _altos_automaticos(
                origen[hoja_origen], destino[hoja_destino], inserciones, mapeo)
            destino[hoja_destino] = _estilos_de_combinadas(
                origen[hoja_origen], generada, inserciones, mapeo)

    # -- lo que escribio openpyxl y ya no se usa
    for viejo in partes_borradas:
        destino.pop(viejo, None)
        destino.pop(_ruta_rels(viejo), None)
    referenciadas = set()
    for nombre in list(destino):
        if nombre.endswith(".rels"):
            carpeta = posixpath.dirname(posixpath.dirname(nombre))
            fuente = posixpath.join(carpeta, posixpath.basename(nombre)[:-5])
            for r in _rels(destino, fuente):
                if r.get("TargetMode") != "External":
                    referenciadas.add(_resolver(fuente, r["Target"]))
    for nombre in list(destino):
        if nombre.startswith("xl/media/") and nombre not in referenciadas:
            del destino[nombre]
    for override in list(tipos.findall("ct:Override", NS)):
        if override.get("PartName", "").lstrip("/") not in destino:
            tipos.remove(override)
    por_defecto = [dict(d.attrib) for d in tipos.findall("ct:Default", NS)]
    existentes = {d.get("Extension", "").lower() for d in por_defecto}
    for extension, tipo in nuevos_tipos.items():
        if extension not in existentes:
            por_defecto.append({"Extension": extension, "ContentType": tipo})
            existentes.add(extension)
    especificos = [dict(o.attrib) for o in tipos.findall("ct:Override", NS)]
    destino["[Content_Types].xml"] = _tipos_xml(por_defecto, especificos)

    _escribir_paquete(destino, generado)
    return avisos


_FILA = re.compile(r"<row\b[^>]*>")
_FILA_COMPLETA = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.S)


def _destinos(fila: int, inserciones: Sequence[Tuple[int, int]],
              mapeo: Optional[Callable[[int], List[int]]]) -> List[int]:
    """Las filas finales de una fila de la plantilla: corridas por las
    inserciones, o las que diga el mapeo si la hoja repitio partes."""
    if mapeo is not None:
        return list(mapeo(fila))
    for desde, cantidad in inserciones:
        if fila >= desde:
            fila += cantidad
    return [fila]


def _altos_automaticos(hoja_plantilla: bytes, hoja_generada: bytes,
                       inserciones: Sequence[Tuple[int, int]],
                       mapeo: Optional[Callable[[int], List[int]]] = None) -> bytes:
    """openpyxl marca TODA fila con alto como alto fijo (customHeight). En
    la plantilla del banco muchas no lo son: Excel guardo el alto que calculo
    al ajustarlas a su contenido, y lo recalcula al abrir. Fijarlas corre la
    pagina un par de pixeles -medido: 1 a 2 px- y, peor, deja cortado un
    texto mas largo que el de la plantilla. Se les devuelve el alto
    automatico a las filas que lo tenian.

    Salvo a las que tienen formulas: Excel ajusta la fila al abrir ANTES de
    calcular, asi que una fila cuyo texto sale de una formula se ajusta a una
    celda vacia y queda de un renglon (medido en la hoja 'plantas': el nombre
    del cliente, que alla es formula, salia cortado). Esas quedan con el alto
    que traia la plantilla, fijo."""
    automaticas = set()
    for etiqueta in _FILA.findall(hoja_plantilla.decode("utf-8")):
        if ' ht="' in etiqueta and "customHeight" not in etiqueta:
            fila = int(re.search(r'\br="(\d+)"', etiqueta).group(1))
            automaticas.update(_destinos(fila, inserciones, mapeo))
    if not automaticas:
        return hoja_generada

    def arreglar(m):
        fila = m.group(0)
        apertura = _FILA.match(fila).group(0)
        numero = re.search(r'\br="(\d+)"', apertura)
        if numero and int(numero.group(1)) in automaticas and "<f>" not in fila \
                and "<f " not in fila:
            return apertura.replace(' customHeight="1"', "") + fila[len(apertura):]
        return fila
    return _FILA_COMPLETA.sub(arreglar, hoja_generada.decode("utf-8")).encode("utf-8")


_CELDA = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/?)>')


def _columna(letras: str) -> int:
    n = 0
    for letra in letras:
        n = n * 26 + ord(letra) - 64
    return n


def _letras(n: int) -> str:
    salida = ""
    while n:
        n, resto = divmod(n - 1, 26)
        salida = chr(65 + resto) + salida
    return salida


def _estilos_de_combinadas(hoja_plantilla: bytes, hoja_generada: bytes,
                           inserciones: Sequence[Tuple[int, int]],
                           mapeo: Optional[Callable[[int], List[int]]] = None) -> bytes:
    """Al LEER un rango combinado, openpyxl "rehace" el borde de las celdas
    internas a partir del de la esquina y les asigna un estilo nuevo. En el
    formato del banco eso cambia lineas que se ven: medido en SUELOS, el
    encabezado A44:A45. A cada celda interna de un combinado se le devuelve
    el estilo que traia la plantilla -los indices de estilo de la plantilla
    siguen valiendo en el archivo generado: openpyxl conserva la lista de
    estilos en su orden y solo agrega al final-. La esquina, que es donde se
    escribe, no se toca."""
    texto = hoja_plantilla.decode("utf-8")
    estilos = {(m.group(1), int(m.group(2))): (re.search(r'\bs="(\d+)"', m.group(3)) or [None, "0"])[1]
               for m in _CELDA.finditer(texto)}
    internas: Dict[Tuple[str, int], str] = {}
    for ref in re.findall(r'<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"', texto):
        (c1, f1), (c2, f2) = [(re.match(r"([A-Z]+)(\d+)", p).group(1),
                               int(re.match(r"([A-Z]+)(\d+)", p).group(2)))
                              for p in ref.split(":")]
        for col in range(_columna(c1), _columna(c2) + 1):
            for fila in range(f1, f2 + 1):
                if (col, fila) == (_columna(c1), f1):
                    continue
                clave = (_letras(col), fila)
                for destino_fila in _destinos(fila, inserciones, mapeo):
                    internas[(clave[0], destino_fila)] = estilos.get(clave, "0")
    if not internas:
        return hoja_generada

    def arreglar(m):
        clave = (m.group(1), int(m.group(2)))
        if clave not in internas:
            return m.group(0)
        resto = re.sub(r'\s+s="\d+"', "", m.group(3))
        return '<c r="{}{}" s="{}"{}{}>'.format(m.group(1), m.group(2), internas[clave],
                                                resto, m.group(4))
    return _CELDA.sub(arreglar, hoja_generada.decode("utf-8")).encode("utf-8")


def _tipos_xml(por_defecto: Sequence[Dict[str, str]],
               especificos: Sequence[Dict[str, str]]) -> bytes:
    """El indice del paquete, escrito a mano: sin prefijo de espacio de
    nombres -como lo escribe Excel; no todos los lectores toleran un 'ns0:'
    ahi- y con los Default antes que los Override, que es el orden que
    exige el esquema."""
    def fila(etiqueta, atributos):
        return "<{} {}/>".format(etiqueta, " ".join(
            '{}="{}"'.format(k, _escapar(v)) for k, v in atributos.items()))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="{}">{}{}</Types>'
            .format(NS["ct"], "".join(fila("Default", d) for d in por_defecto),
                    "".join(fila("Override", o) for o in especificos))).encode("utf-8")
