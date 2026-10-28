"""Motor de envejecimiento: proyecta una cartera de plantaciones, no una sola.

El problema que resuelve. ``services/curvas.py:proyectar`` evalua UNA curva
contra UNA superficie a UNA edad. Un cliente real no es eso: Cliente A tiene 19
plantaciones de 3 especies plantadas en anios distintos, y en la temporada
2027/28 conviven un palto de 12 anios en plena produccion con otro de 3 en
formacion. El rendimiento del bloque es el promedio ponderado de los dos, y
cambia solo temporada a temporada porque cada plantacion envejece.

El patron es el del archivo HdU (matriz edad x anio de proyeccion), no el del
RER (un bloque por plantacion y sumas encadenadas a mano, que hay que reescribir
cada vez que cambia el numero de plantaciones). Ver CLAUDE.md seccion 6.

Decisiones tomadas con Fitogreen (2026-08-12):

- **Se agrupa por curva**, no por especie. Palto Hass y Palto clonal rinden
  13.000 y 15.000 kg/ha: promediarlos inventa un rendimiento que no es de
  ninguno de los dos.
- **Las plantaciones futuras entran solas.** Un "proyecto de plantacion" no es
  un concepto aparte: es una plantacion cuyo anio todavia no llega. Hasta
  entonces aporta 0 ha y no ensucia el promedio.
- **Aqui no se calculan las temporadas Real ni Preliminar.** Esas columnas
  llevan lo que el cliente declara; llenarlas con la curva seria presentarle al
  banco un supuesto teorico rotulado como hecho. El llamador pasa solo las
  temporadas de proyeccion.

La aritmetica no se reimplementa: cada plantacion se evalua con ``proyectar``,
que ya aplica los ajustes del tasador y la regla de que el costo variable baja
con el rendimiento ajustado. Aqui solo se agrega.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

from schemas.curvas import AjusteAplicado, CurvaProduccion, SolicitudProyeccion
from schemas.informe import (
    BloqueMatriz, FilaMatriz, MatrizEnvejecimiento, Plantacion,
)
from services.catastro_import import (
    anio_inicial, especie_base, formas_de_nombre, normalizar,
)
from services.curvas import (
    costo_base_a_edad, factor_de_ajustes, proyectar, rendimiento_a_edad,
)

#: Lo que separa una especie de su calificativo: 'Palto (rebajado)',
#: 'Palto - clonal'. Mismo criterio y mismo valor por defecto que
#: ``catastro_import.cargar_reglas``, para no tener dos convenciones.
_SEPARADORES = ("(", " - ")

#: Calificativos que a veces vienen pegados con un simple espacio, sin ningun
#: separador: 'Palto convencional'. Se recortan SOLO como ultimo recurso,
#: cuando ya fallo todo lo demas, y solo si estan en esta lista. Truncar a la
#: primera palabra a ciegas seria peor que el problema: 'Uva de mesa' es otro
#: cultivo que la uva vinifera y no debe heredar su curva.
_CALIFICATIVOS = ("convencional", "tradicional", "comun", "injertado",
                  "adulto", "joven", "nuevo", "viejo", "vinifera", "europeo")

__all__ = [
    "GrupoDeFlujo", "CarteraProyectada", "anio_de_temporada",
    "emparejar_curva", "proyectar_cartera", "a_matriz_de_dominio",
    "cobertura_de_curvas", "curvas_de_la_misma_especie",
]


def anio_de_temporada(temporada: str) -> Optional[int]:
    """'2027/28' -> 2027. La temporada se nombra por el anio en que empieza."""
    inicio = anio_inicial(temporada)
    return inicio


#: La regla del plural castellano vive en `catastro_import`: la usan tambien
#: la biblioteca de precios y la de referencias de ODEPA, y tres copias de una
#: regla de idioma se separan solas.
_formas = formas_de_nombre


def _coincide(uno: Optional[str], otro: Optional[str]) -> bool:
    """Dos nombres designan lo mismo. Vacio no coincide con nada."""
    formas_uno, formas_otro = _formas(uno), _formas(otro)
    return bool(formas_uno and formas_otro and (formas_uno & formas_otro))


def emparejar_curva(
    plantacion: Plantacion,
    curvas: Sequence[CurvaProduccion],
    separadores: Sequence[str] = _SEPARADORES,
) -> Optional[CurvaProduccion]:
    """Busca la curva de una plantacion, de lo especifico a lo general.

    El orden importa y no es cosmetico. Los catastros reales traen la especie
    calificada: Cliente A tiene 'Palto', 'Palto - clonal' y 'Palto (rebajado)'
    en la misma planilla. El calificativo se prueba ANTES de recortarlo, para
    que 'Palto - clonal' encuentre su curva clonal (15.000 kg/ha) en vez de
    caer al palto generico (13.000). Solo si nada calza con el nombre completo
    se recorta a la especie base, que es lo correcto para 'Palto (rebajado)':
    un palto rebajado es un palto, lo distinto es el anio -el del rebaje, no el
    de plantacion-, y de eso ya se encarga el catastro.

    Devuelve ``None`` si no hay ninguna. Es deliberado: una especie sin
    parametros se reporta como tal y se deja fuera del flujo. Inventarle una
    curva a un avellano porque 'se parece' a un nogal es exactamente lo que
    este proyecto no debe hacer.
    """
    especie = plantacion.especie
    variedad = plantacion.variedad
    base = especie_base(especie or "", list(separadores))

    # 1. Solo si la especie VIENE CALIFICADA, el nombre completo contra el
    #    nombre de la curva: 'Palto - clonal' -> curva 'Palto Clonal'. Si no se
    #    exigiera el calificativo, una plantacion 'Palto' con variedad
    #    'Clonal' calzaria de inmediato con la curva llamada 'Palto' y se
    #    perderia la variedad, que es mas especifica.
    plana = normalizar(especie)
    if plana and normalizar(base) != plana:
        for curva in curvas:
            if plana == normalizar(curva.nombre):
                return curva

    # 2. Con el calificativo, y despues sin el.
    for candidata in (especie, base):
        if not _formas(candidata):
            continue
        if _formas(variedad):
            for curva in curvas:
                if (_coincide(curva.especie, candidata)
                        and _coincide(curva.variedad, variedad)):
                    return curva
            plana_variedad = normalizar(variedad)
            for curva in curvas:
                if (_coincide(curva.especie, candidata)
                        and plana_variedad in normalizar(curva.nombre)):
                    return curva
        for curva in curvas:
            if _coincide(curva.especie, candidata) and not _formas(curva.variedad):
                return curva
        for curva in curvas:
            if _coincide(curva.especie, candidata):
                return curva

    # 3. Ultimo recurso: quitar un calificativo pegado con un simple espacio.
    #    'Palto convencional' es un palto; 'Uva de mesa' NO es una uva
    #    vinifera, y por eso solo se recortan las palabras de la lista.
    plano = normalizar(especie)
    for calificativo in _CALIFICATIVOS:
        if plano.endswith(" " + calificativo):
            recortada = plano[: -(len(calificativo) + 1)].strip()
            for curva in curvas:
                if _coincide(curva.especie, recortada):
                    return curva
            break
    return None


def curvas_de_la_misma_especie(
    especie: Optional[str],
    curvas: Sequence[CurvaProduccion],
    separadores: Sequence[str] = _SEPARADORES,
) -> Tuple[List[CurvaProduccion], List[CurvaProduccion]]:
    """Reparte el catalogo en «otras variedades de esta especie» y el resto.

    Para cuando una VARIEDAD no tiene curva propia: un arandano Emerald sin
    curva se parece a un arandano Duke, y eso no hay que inventarlo —es la
    misma especie—. Un cerezo, en cambio, no se le parece: que los numeros den
    cerca es casualidad, no agronomia. Este es el reparto que separa lo que se
    puede recomendar de lo que no.

    Reusa `_coincide`, que es el mismo criterio de `emparejar_curva`: resuelve
    plurales ('Cerezos' -> 'cerezo') y calificativos. Escribir aqui un
    `.lower()` propio seria una segunda regla de idioma, y este proyecto ya
    pago tres veces por tener copias de esa regla (§16.a.quadragies.quinquies).

    Devuelve ``(mismas, otras)``. Sin especie, todo cae en ``otras``: no hay
    contra que comparar y no se recomienda nada.
    """
    if not _formas(especie):
        return [], list(curvas)

    base = especie_base(especie or "", list(separadores))
    mismas, otras = [], []
    for curva in curvas:
        de_la_especie = (_coincide(curva.especie, especie)
                         or _coincide(curva.especie, base))
        (mismas if de_la_especie else otras).append(curva)
    return mismas, otras


@dataclass
class GrupoDeFlujo:
    """Un bloque del flujo: las plantaciones que comparten curva, ya agregadas.

    ``superficie_por_edad`` es la matriz de envejecimiento propiamente dicha:
    para cada temporada, cuantas hectareas hay a cada edad. Es lo que se vuelca
    a la hoja anexa para que el flujo la referencie con SUMPRODUCT en vez de
    recibir un promedio ya masticado.
    """

    curva: CurvaProduccion
    plantaciones: List[Plantacion]
    etiqueta: str
    #: Los ajustes que se le aplicaron A ESTE bloque. Se guardan por grupo y no
    #: sueltos porque un huerto mal manejado no contagia al de al lado: bajarle
    #: la curva al palto no tiene por que bajarsela al cerezo.
    ajustes: List[AjusteAplicado] = field(default_factory=list)
    superficie_por_edad: Dict[str, Dict[int, float]] = field(default_factory=dict)
    superficie_ha: Dict[str, float] = field(default_factory=dict)
    rendimiento_kg_ha: Dict[str, float] = field(default_factory=dict)
    costos_usd_ha: Dict[str, float] = field(default_factory=dict)
    produccion_total_kg: Dict[str, float] = field(default_factory=dict)

    @property
    def edades_presentes(self) -> List[int]:
        """Edades con superficie en alguna temporada, ordenadas."""
        edades = {
            edad
            for reparto in self.superficie_por_edad.values()
            for edad, ha in reparto.items()
            if ha > 0
        }
        return sorted(edades)


@dataclass
class CarteraProyectada:
    """Resultado completo, con lo que quedo fuera dicho en voz alta."""

    grupos: List[GrupoDeFlujo] = field(default_factory=list)
    sin_curva: List[Plantacion] = field(default_factory=list)
    sin_anio: List[Plantacion] = field(default_factory=list)
    avisos: List[str] = field(default_factory=list)

    @property
    def superficie_proyectada_ha(self) -> float:
        return round(
            sum(p.superficie_ha for g in self.grupos for p in g.plantaciones), 4
        )

    @property
    def superficie_excluida_ha(self) -> float:
        return round(
            sum(p.superficie_ha for p in self.sin_curva + self.sin_anio), 4
        )


def _etiqueta(plantaciones: Sequence[Plantacion], curva: CurvaProduccion) -> str:
    """Texto de la primera fila del bloque: 'Especie/Variedad/Anio'."""
    variedades = sorted({p.variedad for p in plantaciones if p.variedad})
    anios = sorted({
        a for a in (anio_inicial(p.anio_plantacion) for p in plantaciones)
        if a is not None
    })

    partes = [plantaciones[0].especie or curva.especie]
    if len(variedades) == 1:
        partes.append(variedades[0])
    elif len(variedades) > 1:
        partes.append("/".join(variedades[:3]))
    if anios:
        partes.append(
            str(anios[0]) if anios[0] == anios[-1]
            else "{}-{}".format(anios[0], anios[-1])
        )
    return " ".join(str(p) for p in partes if p)


def _agrupar(
    plantaciones: Sequence[Plantacion], curvas: Sequence[CurvaProduccion]
) -> Tuple[List[Tuple[CurvaProduccion, List[Plantacion]]], List[Plantacion],
           List[Plantacion]]:
    """Reparte las plantaciones por curva, conservando el orden de aparicion."""
    grupos: Dict[str, Tuple[CurvaProduccion, List[Plantacion]]] = {}
    sin_curva: List[Plantacion] = []
    sin_anio: List[Plantacion] = []

    for plantacion in plantaciones:
        if anio_inicial(plantacion.anio_plantacion) is None:
            sin_anio.append(plantacion)
            continue
        curva = emparejar_curva(plantacion, curvas)
        if curva is None:
            sin_curva.append(plantacion)
            continue
        grupos.setdefault(curva.id, (curva, []))[1].append(plantacion)

    return list(grupos.values()), sin_curva, sin_anio


def cobertura_de_curvas(
    plantaciones: Sequence[Plantacion],
    curvas: Sequence[CurvaProduccion],
) -> List[Dict[str, Any]]:
    """Que especies del cliente tienen curva y cuales no, con sus hectareas.

    Existe para poder avisarlo **al importar el catastro** y no al generar el
    informe. Cuando el aviso llega al final, el tasador ya se sento a hacer el
    trabajo; al importar todavia esta a tiempo de pedir los numeros.

    Se agrupa por la especie tal como viene escrita en la planilla del
    agricultor -no por la curva- porque es lo que el tasador reconoce y lo que
    tendria que ir a cargar. Las hectareas importan tanto como el nombre: que
    falte la curva de 2 ha no es lo mismo que de 55 de 95, y ese fue el caso
    real de Cliente A (83% de la superficie sin curva).
    """
    por_especie: Dict[str, Dict[str, Any]] = {}
    for plantacion in plantaciones:
        etiqueta = (plantacion.especie or "(sin especie)").strip()
        entrada = por_especie.setdefault(etiqueta, {
            "especie": etiqueta,
            "superficie_ha": 0.0,
            "plantaciones": 0,
            "curva": None,
            "sin_anio_ha": 0.0,
            "variedades": {},
        })
        entrada["superficie_ha"] += plantacion.superficie_ha
        entrada["plantaciones"] += 1

        # Las variedades del catastro, con su peso. El agronomo no tendria por
        # que volver al Excel a mirar cuales eran; y ademas decide algo: una
        # sola variedad se puede precargar sin preguntar, varias no -puede
        # querer una curva por variedad, o una sola para toda la especie-.
        nombre_var = (plantacion.variedad or "").strip()
        if nombre_var:
            variedad = entrada["variedades"].setdefault(
                nombre_var, {"variedad": nombre_var,
                             "superficie_ha": 0.0, "plantaciones": 0})
            variedad["superficie_ha"] += plantacion.superficie_ha
            variedad["plantaciones"] += 1
        if not plantacion.anio_plantacion:
            entrada["sin_anio_ha"] += plantacion.superficie_ha
        if entrada["curva"] is None:
            curva = emparejar_curva(plantacion, curvas)
            if curva is not None:
                entrada["curva"] = curva.nombre

    for entrada in por_especie.values():
        entrada["superficie_ha"] = round(entrada["superficie_ha"], 2)
        entrada["sin_anio_ha"] = round(entrada["sin_anio_ha"], 2)
        entrada["tiene_curva"] = entrada["curva"] is not None
        entrada["variedades"] = sorted(
            ({**v, "superficie_ha": round(v["superficie_ha"], 2)}
             for v in entrada["variedades"].values()),
            key=lambda v: -v["superficie_ha"],
        )

    # Lo que falta primero, y dentro de eso lo mas grande: es el orden en que
    # conviene resolverlo.
    return sorted(por_especie.values(),
                  key=lambda e: (e["tiene_curva"], -e["superficie_ha"]))


def proyectar_cartera(
    plantaciones: Sequence[Plantacion],
    curvas: Sequence[CurvaProduccion],
    temporadas: Sequence[str],
    ajustes: Optional[Sequence[AjusteAplicado]] = None,
    ajustes_por_curva: Optional[Dict[str, Sequence[AjusteAplicado]]] = None,
    capacidad_bloques: Optional[int] = None,
) -> CarteraProyectada:
    """Envejece la cartera completa y la agrega a bloques listos para el flujo.

    ``temporadas`` son las de proyeccion, en el formato '2027/28'. Las columnas
    Real y Preliminar no se calculan aqui (ver el docstring del modulo).

    ``capacidad_bloques`` avisa -no trunca- cuando la cartera necesita mas
    bloques de los que la plantilla trae. Truncar en silencio dejaria una
    especie del cliente fuera del informe del banco.

    Hay dos niveles de ajuste y se multiplican entre si:

    - ``ajustes`` vale para toda la cartera. Es lo correcto para algo del
      predio: la seguridad de riego, la zona.
    - ``ajustes_por_curva`` vale solo para un bloque, indexado por el id o el
      nombre de su curva. Es lo correcto para algo de ESE huerto: que el palto
      esta sin podar no dice nada del cerezo de al lado.
    """
    ajustes = list(ajustes or [])
    por_curva = dict(ajustes_por_curva or {})
    crudos, sin_curva, sin_anio = _agrupar(plantaciones, curvas)
    resultado = CarteraProyectada(sin_curva=sin_curva, sin_anio=sin_anio)

    # Estos avisos los lee el tasador en la pantalla y salen impresos en la
    # hoja anexa del informe: con tildes y con la coma decimal de aca, no
    # «5.0 ha».
    def ha(valor: float) -> str:
        return ("{:,.1f}".format(valor)
                .replace(",", "_").replace(".", ",").replace("_", "."))

    for plantacion in sin_curva:
        resultado.avisos.append(
            "sin curva para '{}{}' ({} ha): queda fuera del flujo".format(
                plantacion.especie,
                " " + plantacion.variedad if plantacion.variedad else "",
                ha(plantacion.superficie_ha),
            )
        )
    for plantacion in sin_anio:
        resultado.avisos.append(
            "'{}' no trae año de plantación ({} ha): no se puede envejecer"
            .format(plantacion.especie, ha(plantacion.superficie_ha))
        )

    # Grupo mas grande primero: si la plantilla no da abasto, los bloques que
    # se escriban son los que mas pesan en el resultado del cliente.
    crudos.sort(
        key=lambda par: sum(p.superficie_ha for p in par[1]), reverse=True
    )

    for curva, del_grupo in crudos:
        # El tasador escribe 'Palto' en la pantalla; la curva puede llamarse
        # 'Palto Hass'. Se busca por id, por nombre y por especie para que no
        # tenga que acordarse de como se llama la curva en el catalogo.
        propios = []
        for clave in (curva.id, curva.nombre, curva.especie):
            propios = list(por_curva.get(clave, []))
            if propios:
                break
        if not propios:
            for clave, lista in por_curva.items():
                if _coincide(clave, curva.especie) or _coincide(clave, curva.nombre):
                    propios = list(lista)
                    break
        grupo = GrupoDeFlujo(
            curva=curva,
            plantaciones=list(del_grupo),
            etiqueta=_etiqueta(del_grupo, curva),
            ajustes=ajustes + propios,
        )

        for temporada in temporadas:
            anio = anio_de_temporada(temporada)
            if anio is None:
                resultado.avisos.append(
                    "temporada '{}' no tiene anio reconocible; se omite".format(
                        temporada
                    )
                )
                continue

            reparto: Dict[int, float] = {}
            superficie = produccion = costo = 0.0

            for plantacion in del_grupo:
                plantado = anio_inicial(plantacion.anio_plantacion)
                edad = anio - plantado
                if edad < 0 or plantacion.superficie_ha <= 0:
                    continue   # todavia no entra, o no aporta superficie

                reparto[edad] = round(
                    reparto.get(edad, 0.0) + plantacion.superficie_ha, 4
                )
                superficie += plantacion.superficie_ha

                punto = proyectar(curva, SolicitudProyeccion(
                    curva_id=curva.id,
                    superficie_ha=plantacion.superficie_ha,
                    edad_actual=edad,
                    ajustes=grupo.ajustes,
                )).puntos[0]
                produccion += punto.produccion_total_kg
                costo += punto.costo_usd

            grupo.superficie_por_edad[temporada] = reparto
            grupo.superficie_ha[temporada] = round(superficie, 4)
            grupo.produccion_total_kg[temporada] = round(produccion, 4)
            # El promedio es ponderado por superficie: es lo que exige el
            # encaje de N plantaciones en un unico bloque de la plantilla.
            grupo.rendimiento_kg_ha[temporada] = (
                round(produccion / superficie, 4) if superficie else 0.0
            )
            grupo.costos_usd_ha[temporada] = (
                round(costo / superficie, 4) if superficie else 0.0
            )

        resultado.grupos.append(grupo)

    if capacidad_bloques is not None and len(resultado.grupos) > capacidad_bloques:
        resultado.avisos.append(
            "la cartera necesita {} bloques y la plantilla trae {}: hay que "
            "ampliarla o agrupar especies antes de generar".format(
                len(resultado.grupos), capacidad_bloques
            )
        )

    return resultado


def a_matriz_de_dominio(
    cartera: CarteraProyectada,
    temporadas: Sequence[str],
    ajustes: Optional[Sequence[AjusteAplicado]] = None,
) -> MatrizEnvejecimiento:
    """Traduce la cartera proyectada al modelo que viaja al Excel.

    El eje de edades **no** se corta en la edad maxima de la curva. Cliente A
    tiene un palto de 1993 que en la temporada 2033/34 cumple 40 anios, y la
    curva llega a 25: si la matriz parara en 25, esas hectareas no tendrian
    fila y el SUMPRODUCT de la hoja las perderia en silencio. Se extiende hasta
    la edad mas vieja presente, y el rendimiento de esas filas es el del ultimo
    punto de la curva, que es justo lo que hace ``rendimiento_a_edad``.
    """
    globales = list(ajustes or [])
    matriz = MatrizEnvejecimiento(
        temporadas=list(temporadas),
        avisos=list(cartera.avisos),
        superficie_excluida_ha=cartera.superficie_excluida_ha,
    )

    for grupo in cartera.grupos:
        presentes = grupo.edades_presentes
        tope = max([grupo.curva.edad_maxima] + presentes) if presentes \
            else grupo.curva.edad_maxima

        # El eje va de 0 a la edad mas vieja SIN saltarse ninguna, aunque
        # muchas queden en cero: el rango del SUMPRODUCT tiene que ser un
        # bloque continuo, y un agujero en el medio lo desalinearia.
        filas = []
        for edad in range(0, tope + 1):
            reparto = {
                temporada: grupo.superficie_por_edad.get(temporada, {}).get(edad, 0.0)
                for temporada in temporadas
            }
            filas.append(FilaMatriz(
                edad=edad,
                rendimiento_curva_kg_ha=rendimiento_a_edad(grupo.curva, edad),
                costo_base_usd_ha=costo_base_a_edad(grupo.curva, edad),
                superficie_por_temporada=reparto,
            ))

        # El factor de ESTE bloque, no el de la cartera: es el numero que
        # queda como celda viva en la hoja anexa, y de el cuelga todo el
        # rendimiento del bloque.
        factor = factor_de_ajustes(grupo.ajustes or globales)
        # Si la curva se armo PARTIENDO de otra especie, la hoja anexa lo dice.
        # Esa hoja es la evidencia de como se armo el flujo, y una curva
        # prestada escrita igual que una medida deja al banco creyendo que hay
        # una medicion del arandano donde en realidad hay la forma del cerezo.
        # Misma regla que `de_otra_variedad` (§16.a.sexies), un nivel mas
        # arriba: alla se presta entre variedades, aqui entre ESPECIES.
        nombre_curva = grupo.curva.nombre
        if grupo.curva.derivada_de:
            nombre_curva = "{} (derivada de {})".format(
                nombre_curva, grupo.curva.derivada_de)
        matriz.bloques.append(BloqueMatriz(
            etiqueta=grupo.etiqueta,
            curva=nombre_curva,
            factor_ajuste=factor,
            costo_por_kg_usd=grupo.curva.costos.costo_por_kg_usd,
            filas=filas,
        ))

    return matriz
