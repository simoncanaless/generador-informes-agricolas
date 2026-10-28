"""Ensamblado determinista: de lo extraido por la IA al informe del banco.

Aqui NO interviene ningun modelo. La IA ya hizo su parte (leer papeles y fotos);
este modulo solo traduce esos datos a la estructura ``InformeFitogreen`` y, de
paso, hace algo que a un tasador le importa mucho: **cruzar los documentos entre
si y avisar cuando se contradicen**.

Si la escritura dice que el rol es 123-45 y el certificado de avaluo dice 123-46,
uno de los dos esta mal y hay que mirarlo antes de firmar. Detectarlo es
comparacion de cadenas, no inteligencia artificial.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence

from schemas.extraccion import (
    Confianza, DerechoAguaExtraido, ExtraccionLegal, InscripcionCBR,
    PropiedadExtraida, TituloFotografia,
)
from schemas.informe import (
    Antecedentes, CategoriaFoto, DatosSuelos, DerechoAgua, DotacionRecursos,
    Fotografia, InformeFitogreen, Lote,
)

__all__ = ["fusionar_extracciones", "ensamblar_informe", "Conflicto", "Fusion"]


#: Campos escalares de la propiedad que se cruzan entre documentos.
CAMPOS_COTEJADOS = (
    "propietario", "rut_propietario", "nombre_predio", "rol_avaluo",
    "comuna", "provincia", "region", "superficie_escritura_ha",
)

#: Etiquetas legibles para los avisos, con articulo para que concuerden.
ETIQUETAS = {
    "propietario": "El propietario",
    "rut_propietario": "El RUT del propietario",
    "nombre_predio": "El nombre del predio",
    "rol_avaluo": "El rol de avaluo",
    "comuna": "La comuna",
    "provincia": "La provincia",
    "region": "La region",
    "superficie_escritura_ha": "La superficie segun escritura",
}

#: Traduccion de la categoria sugerida por la IA a la del informe.
CATEGORIAS = {
    "emplazamiento": CategoriaFoto.EMPLAZAMIENTO,
    "plantacion": CategoriaFoto.PLANTACION,
    "infraestructura": CategoriaFoto.INFRAESTRUCTURA,
    "riego": CategoriaFoto.RIEGO,
    "satelital": CategoriaFoto.SATELITAL,
    "legal": CategoriaFoto.LEGAL,
}


@dataclass
class Conflicto:
    """Dos documentos afirman cosas distintas sobre el mismo campo.

    ``valores`` va indexado por documento y guarda el valor TAL COMO aparece en
    el papel, no el normalizado: el tasador tiene que reconocer lo que lee.
    """

    campo: str
    valores: Dict[str, Any]

    def __str__(self) -> str:
        detalle = "; ".join(
            "{} dice '{}'".format(doc, val) for doc, val in self.valores.items()
        )
        return "{} no coincide entre documentos: {}".format(
            ETIQUETAS.get(self.campo, self.campo), detalle
        )


@dataclass
class Fusion:
    """Resultado de cruzar varios documentos.

    ``propiedades`` puede traer mas de una: una tasacion normal cubre varios
    roles a la vez (el caso Cliente C son dos lotes, 6,84 ha y 78,55 ha). Dos
    documentos que hablan de predios distintos NO se contradicen, y confundir
    ambas cosas llenaria la pantalla de alarmas falsas.
    """

    datos: ExtraccionLegal
    propiedades: List[PropiedadExtraida] = field(default_factory=list)
    conflictos: List[Conflicto] = field(default_factory=list)
    avisos: List[str] = field(default_factory=list)

    @property
    def hay_que_revisar(self) -> bool:
        return bool(self.conflictos) or bool(self.avisos)

    @property
    def superficie_total_ha(self) -> float:
        """Suma de las superficies declaradas por cada propiedad."""
        return round(sum(
            p.superficie_escritura_ha or 0.0 for p in self.propiedades
        ), 4)


# ------------------------------------------------------------ normalizado

def _normalizar(valor: Any) -> Any:
    """Deja un valor comparable: sin tildes, sin puntuacion suelta, en minusculas.

    'Requínoa' y 'REQUINOA' son la misma comuna; '123-45' y '123 - 45' el mismo
    rol. Comparar en crudo produciria conflictos falsos que solo generan ruido.
    """
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return round(float(valor), 2)

    texto = str(valor).strip()
    if not texto:
        return None
    sin_tildes = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"[^a-z0-9]+", "", sin_tildes.lower()) or None


def _preferir(actual: Any, nuevo: Any) -> Any:
    """Se queda con el primer valor no vacio. El orden de los documentos manda."""
    if actual is None or (isinstance(actual, str) and not actual.strip()):
        return nuevo
    return actual


def _mejor_confianza(a: Confianza, b: Confianza) -> Confianza:
    orden = {Confianza.BAJA: 0, Confianza.MEDIA: 1, Confianza.ALTA: 2}
    return a if orden[a] >= orden[b] else b


# --------------------------------------------------------------- fusion

def fusionar_extracciones(
    extracciones: Sequence[ExtraccionLegal],
    nombres: Optional[Sequence[str]] = None,
) -> Fusion:
    """Cruza varios documentos en un solo juego de antecedentes.

    Para cada campo escalar se conserva el primer valor no vacio y se avisa si
    algun otro documento dice algo distinto. Las listas (derechos de agua,
    gravamenes) se acumulan sin perder nada, quitando duplicados evidentes.
    """
    if not extracciones:
        return Fusion(datos=ExtraccionLegal())

    nombres = list(nombres or [])
    etiqueta = lambda i: (  # noqa: E731
        nombres[i] if i < len(nombres) else "documento {}".format(i + 1)
    )

    fusion = Fusion(datos=ExtraccionLegal())

    # 1) Agrupar los documentos por la propiedad de la que hablan.
    grupos: List[List[int]] = []
    for indice, extraccion in enumerate(extracciones):
        if not _identifica_propiedad(extraccion.propiedad):
            continue  # p.ej. un certificado de aguas: aporta derechos, no predio
        for grupo in grupos:
            if any(_misma_propiedad(extraccion.propiedad,
                                    extracciones[otro].propiedad)
                   for otro in grupo):
                grupo.append(indice)
                break
        else:
            grupos.append([indice])

    # 2) Dentro de cada grupo si tiene sentido cotejar: son el mismo predio.
    for grupo in grupos:
        propiedad = PropiedadExtraida()
        vistos: Dict[str, Dict[Any, tuple]] = {c: {} for c in CAMPOS_COTEJADOS}

        for indice in grupo:
            origen = etiqueta(indice)
            actual = extracciones[indice].propiedad

            for campo in CAMPOS_COTEJADOS:
                valor = getattr(actual, campo, None)
                clave = _normalizar(valor)
                if clave is None:
                    continue
                vistos[campo].setdefault(clave, (origen, valor))
                setattr(propiedad, campo,
                        _preferir(getattr(propiedad, campo), valor))

            propiedad.deslindes = _preferir(propiedad.deslindes, actual.deslindes)
            if actual.inscripcion and not propiedad.inscripcion:
                propiedad.inscripcion = actual.inscripcion
            propiedad.procedencia.confianza = _mejor_confianza(
                propiedad.procedencia.confianza, actual.procedencia.confianza
            )

        for campo, valores in vistos.items():
            if len(valores) > 1:
                fusion.conflictos.append(Conflicto(
                    campo=campo,
                    valores={doc: original for doc, original in valores.values()},
                ))
        fusion.propiedades.append(propiedad)

    # 3) Los derechos y gravamenes son del expediente, no de un predio concreto.
    for indice, extraccion in enumerate(extracciones):
        fusion.datos.derechos_agua.extend(extraccion.derechos_agua)
        fusion.datos.gravamenes.extend(extraccion.gravamenes)
        fusion.datos.tipo_documento = _preferir(
            fusion.datos.tipo_documento, extraccion.tipo_documento
        )
        for advertencia in extraccion.advertencias:
            fusion.avisos.append("{}: {}".format(etiqueta(indice), advertencia))

    if len(fusion.propiedades) > 1:
        roles = [p.rol_avaluo or "sin rol" for p in fusion.propiedades]
        fusion.avisos.append(
            "el expediente cubre {} propiedades distintas (roles {}), {} ha en "
            "total: deben ir como predios separados en el informe".format(
                len(fusion.propiedades), ", ".join(roles), fusion.superficie_total_ha
            )
        )

    fusion.datos.propiedad = (
        fusion.propiedades[0] if fusion.propiedades else PropiedadExtraida()
    )
    fusion.datos.derechos_agua = _sin_duplicados(fusion.datos.derechos_agua)
    fusion.datos.gravamenes = _lista_unica(fusion.datos.gravamenes)
    return fusion


def _identifica_propiedad(propiedad: PropiedadExtraida) -> bool:
    """Un documento aporta un predio si trae rol o nombre."""
    return bool(_normalizar(propiedad.rol_avaluo)
                or _normalizar(propiedad.nombre_predio))


def _misma_propiedad(una: PropiedadExtraida, otra: PropiedadExtraida) -> bool:
    """Decide si dos documentos hablan del mismo predio.

    Basta con que coincida el rol de avaluo O el nombre del predio. Asi, dos
    escrituras del mismo predio con roles distintos se detectan como
    contradiccion, mientras que dos lotes diferentes de la misma tasacion se
    tratan como lo que son: dos propiedades.
    """
    rol_una, rol_otra = _normalizar(una.rol_avaluo), _normalizar(otra.rol_avaluo)
    if rol_una and rol_otra and rol_una == rol_otra:
        return True

    nombre_una = _normalizar(una.nombre_predio)
    nombre_otra = _normalizar(otra.nombre_predio)
    return bool(nombre_una and nombre_otra and nombre_una == nombre_otra)


def _sin_duplicados(derechos: List[DerechoAguaExtraido]) -> List[DerechoAguaExtraido]:
    """Quita derechos repetidos por aparecer en dos documentos."""
    unicos: List[DerechoAguaExtraido] = []
    huellas = set()
    for derecho in derechos:
        huella = (
            _normalizar(derecho.nombre_fuente),
            _normalizar(derecho.tipo_fuente),
            derecho.caudal_litros_segundo,
        )
        if huella in huellas:
            continue
        huellas.add(huella)
        unicos.append(derecho)
    return unicos


def _lista_unica(textos: List[str]) -> List[str]:
    vistos, salida = set(), []
    for texto in textos:
        clave = _normalizar(texto)
        if clave and clave not in vistos:
            vistos.add(clave)
            salida.append(texto)
    return salida


# ------------------------------------------------------------- ensamblado

def formatear_inscripcion(inscripcion: Optional[InscripcionCBR]) -> Optional[str]:
    """'Fojas 1240 N 890 ano 2011, CBR de Rancagua', como lo escribe el tasador."""
    if inscripcion is None:
        return None
    piezas = []
    if inscripcion.fojas:
        piezas.append("Fojas {}".format(inscripcion.fojas))
    if inscripcion.numero:
        piezas.append("N {}".format(inscripcion.numero))
    if inscripcion.anio:
        piezas.append("ano {}".format(inscripcion.anio))
    texto = " ".join(piezas)
    if inscripcion.conservador:
        texto = "{}, {}".format(texto, inscripcion.conservador) if texto else inscripcion.conservador
    return texto or None


def _categoria(sugerida: Optional[str]) -> CategoriaFoto:
    clave = _normalizar(sugerida) or ""
    for nombre, categoria in CATEGORIAS.items():
        if clave == _normalizar(nombre):
            return categoria
    return CategoriaFoto.OTRA


def ensamblar_informe(
    antecedentes: Antecedentes,
    legal: Optional[ExtraccionLegal] = None,
    fotografias: Optional[Sequence[Fotografia]] = None,
) -> InformeFitogreen:
    """Construye el informe a partir de los antecedentes y lo extraido.

    Solo se rellena lo que se sabe. Los bloques que exigen trabajo de terreno
    (catastro de plantaciones, flujo proyectado, opinion) quedan vacios a
    proposito: el Excel sale con esas tablas listas para completar a mano.
    """
    legal = legal or ExtraccionLegal()
    propiedad = legal.propiedad

    dominio = None
    if propiedad.propietario:
        mismo = _normalizar(propiedad.propietario) == _normalizar(antecedentes.cliente)
        dominio = "Propietario" if mismo else propiedad.propietario

    suelos = DatosSuelos(
        cliente=antecedentes.cliente,
        predio=propiedad.nombre_predio,
        dominio=dominio,
        rol_avaluo=propiedad.rol_avaluo,
        comuna=propiedad.comuna,
        inscripcion_dominio=formatear_inscripcion(propiedad.inscripcion),
    )

    derechos = [
        DerechoAgua(
            predio=propiedad.nombre_predio or antecedentes.cliente,
            comuna=propiedad.comuna,
            tipo_fuente=derecho.tipo_fuente,
            nombre_fuente=derecho.nombre_fuente,
            dominio=derecho.naturaleza,
            cantidad=derecho.acciones,
            caudal_nominal_ls=derecho.caudal_litros_segundo,
            cbr_fojas=derecho.inscripcion.fojas if derecho.inscripcion else None,
            cbr_numero=derecho.inscripcion.numero if derecho.inscripcion else None,
            cbr_anio=derecho.inscripcion.anio if derecho.inscripcion else None,
            cbr_conservador=(
                derecho.inscripcion.conservador if derecho.inscripcion else None
            ),
            propietario=propiedad.propietario,
        )
        for derecho in legal.derechos_agua
    ]

    # El formato final de Santander lleva los datos legales lote por lote
    # (Recursos, DATOS LEGALES TERRENOS): la propiedad de la escritura es el
    # primer lote, con su unidad predial -el predio- y su inscripcion en
    # columnas separadas.
    lotes = []
    if propiedad.rol_avaluo or propiedad.nombre_predio:
        inscripcion = propiedad.inscripcion
        lotes.append(Lote(
            unidad_predial=propiedad.nombre_predio or antecedentes.cliente,
            nombre=propiedad.nombre_predio,
            dominio=dominio,
            rol=propiedad.rol_avaluo,
            region=codigo_region(propiedad.region),
            comuna=propiedad.comuna,
            fojas=inscripcion.fojas if inscripcion else None,
            numero=inscripcion.numero if inscripcion else None,
            anio=inscripcion.anio if inscripcion else None,
            conservador=inscripcion.conservador if inscripcion else None,
            superficie_legal_ha=propiedad.superficie_escritura_ha,
        ))

    return InformeFitogreen(
        antecedentes=antecedentes,
        dotacion_recursos=DotacionRecursos(
            suelos=suelos,
            derechos_agua=derechos,
            comentarios=_comentario_gravamenes(legal.gravamenes),
            lotes=lotes,
        ),
        fotografias=list(fotografias or []),
    )


#: La region como la pide la lista desplegable del formato final de Santander
#: (hoja 'Listas Validación'): en numeros romanos, salvo la Metropolitana. Las
#: claves van como las deja `_normalizar` (sin tildes, espacios ni signos).
_REGIONES = {
    "atacama": "III", "coquimbo": "IV", "valparaiso": "V", "metropolitana": "RM",
    "santiago": "RM", "ohiggins": "VI", "libertador": "VI", "maule": "VII",
    "nuble": "XVI", "biobio": "VIII", "araucania": "IX", "losrios": "XIV",
    "loslagos": "X",
}
_ROMANOS = ("XVI", "XIV", "XIII", "XII", "XI", "X", "IX", "VIII", "VII", "VI", "V", "IV", "III",
            "II", "I", "RM")


def codigo_region(texto: Optional[str]) -> Optional[str]:
    """'Región del Maule' -> 'VII'; 'VII Region' -> 'VII'. Lo que no se
    reconoce se devuelve tal cual: mejor un texto a revisar que un codigo
    inventado."""
    if not texto:
        return None
    plano = _normalizar(texto) or ""
    for nombre, codigo in _REGIONES.items():
        if nombre in plano:
            return codigo
    for romano in _ROMANOS:
        if re.search(r"\b{}\b".format(romano), str(texto)):
            return romano
    return texto


def _comentario_gravamenes(gravamenes: List[str]) -> Optional[str]:
    if not gravamenes:
        return None
    return "Gravamenes y prohibiciones detectados en los titulos: {}".format(
        "; ".join(gravamenes)
    )


def fotografias_desde_titulos(
    rutas: Sequence[str],
    titulos: Sequence[TituloFotografia],
) -> List[Fotografia]:
    """Empareja cada archivo con su titulo, en orden.

    La primera imagen de emplazamiento va al unico hueco que trae la plantilla;
    el resto necesita la hoja de galeria.
    """
    fotos: List[Fotografia] = []
    for indice, ruta in enumerate(rutas):
        titulo = titulos[indice] if indice < len(titulos) else None
        fotos.append(
            Fotografia(
                ruta=ruta,
                titulo=titulo.titulo if titulo else None,
                categoria=(
                    _categoria(titulo.categoria_sugerida)
                    if titulo else CategoriaFoto.OTRA
                ),
            )
        )
    return fotos
