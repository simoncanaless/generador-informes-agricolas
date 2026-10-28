"""Curvas de produccion y ajustes del tasador.

El catalogo de curvas es del equipo, no del informe: se crea una vez por especie
o variedad y se reutiliza en todas las tasaciones. Los ajustes son las
"perillas" del apartado 5 de la especificacion, con nombre y motivo para poder
defenderlos ante el banco.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from fastapi import Depends, APIRouter, HTTPException, status
from pydantic import BaseModel, Field

from api.dependencias import obtener_usuario_actual
from api.dependencias import Almacenamiento, Reglas
from schemas.curvas import (
    Ajuste, CurvaProduccion, EstructuraCosto, Proyeccion,
    SolicitudCurvaParametrica, SolicitudProyeccion,
)
from services.almacen import NoEncontrado
from services.curvas import (
    costos_desde_plena_produccion, curva_desde_parametros,
    declive_desde_porcentaje, forma_de_curva, parecido_de_forma, proyectar,
    resumen_por_madurez,
)
from services.envejecimiento import curvas_de_la_misma_especie

# Las curvas las ve y las edita CUALQUIERA que haya entrado: las hace
# el agrónomo y son informacion de la casa, no de un informe. Lo que se
# separa por persona son los informes, no la biblioteca.
router = APIRouter(prefix="/curvas", tags=["curvas"],
                   dependencies=[Depends(obtener_usuario_actual)])


class ItemCobertura(BaseModel):
    """Una especie del catastro, con su variedad si la trae."""

    especie: str
    variedad: Optional[str] = None


class PeticionCobertura(BaseModel):
    """Lo que hay en el catastro, para saber que curvas faltan."""

    items: List[ItemCobertura] = Field(default_factory=list, max_length=300)


@router.post(
    "/cobertura",
    summary="Cuales de estas especies ya tienen curva y cuales faltan",
)
async def cobertura(
    peticion: PeticionCobertura, almacen: Almacenamiento
) -> List[Dict[str, Any]]:
    """El estado de la lista de pendientes del tasador, refrescado.

    Existe para que la pantalla pueda tachar las curvas a medida que se van
    creando, sin reimplementar el emparejador en JavaScript. Emparejar no es
    trivial -calificativos, plurales, tildes, y el caso de 'Uva de mesa' que NO
    debe heredar la curva de la vid- y tener dos versiones de esa regla es
    garantia de que se van a separar.
    """
    from schemas.informe import Plantacion
    from services.catastro_import import normalizar
    from services.envejecimiento import emparejar_curva

    curvas = await almacen.listar_curvas()
    salida = []
    for item in peticion.items:
        especie = (item.especie or "").strip()
        if not especie:
            continue
        variedad = (item.variedad or "").strip() or None
        # superficie 0.1: el emparejador solo mira especie y variedad, pero el
        # modelo exige una superficie valida.
        curva = emparejar_curva(
            Plantacion(especie=especie, variedad=variedad, superficie_ha=0.1),
            curvas)

        # Tres estados, no dos. Que EXISTA una curva no basta: si el catastro
        # trae Gala y Fuji y solo se creo la de Gala, el emparejador le da la
        # de Gala a la Fuji -misma especie- y con dos estados la fila quedaria
        # en verde. El agronomo nunca se enteraria de que su Fuji se esta
        # proyectando con los rendimientos de la Gala.
        if curva is None:
            estado = "falta"
        elif normalizar(curva.variedad or "") == normalizar(variedad or ""):
            estado = "propia"
        elif not (curva.variedad or "").strip():
            estado = "de_la_especie"      # curva generica: correcto
        else:
            estado = "de_otra_variedad"   # prestada: hay que mostrarlo

        salida.append({
            "especie": especie,
            "variedad": variedad,
            "estado": estado,
            "tiene_curva": curva is not None,
            "curva": curva.nombre if curva else None,
            "variedad_de_la_curva": curva.variedad if curva else None,
        })
    return salida


@router.post(
    "/serie",
    summary="La curva muestreada edad por edad, para graficarla",
)
async def serie(curva: CurvaProduccion) -> Dict[str, Any]:
    """Rendimiento y costo a CADA edad, no solo en los puntos de inflexion.

    Existe para que el grafico dibuje lo que el motor calcula. Una curva
    guardada trae solo sus quiebres -0, 3, 7, 20, 30- y lo que hay entremedio
    lo pone ``rendimiento_a_edad`` interpolando. Si el navegador interpolara
    por su cuenta habria dos versiones de esa regla, y el agronomo estaria
    mirando un dibujo que no es el que despues alimenta el informe.

    Se acepta la curva en el cuerpo -no un id- para poder graficar tambien la
    que se acaba de generar y todavia no se guarda.
    """
    from services.curvas import costo_base_a_edad, rendimiento_a_edad

    if not curva.puntos:
        return {"edades": [], "rendimiento_kg_ha": [], "costo_usd_ha": []}

    ultima = max(p.edad for p in curva.puntos)
    edades = list(range(0, int(ultima) + 1))
    rendimientos = [rendimiento_a_edad(curva, e) for e in edades]
    costos = [
        round(costo_base_a_edad(curva, e)
              + curva.costos.costo_por_kg_usd * rendimientos[i], 2)
        for i, e in enumerate(edades)
    ]
    return {
        "nombre": curva.nombre,
        "edades": edades,
        "rendimiento_kg_ha": rendimientos,
        "costo_usd_ha": costos,
        "puntos_declarados": [p.edad for p in curva.puntos],
        "peak_kg_ha": max(rendimientos) if rendimientos else 0,
        "edad_del_peak": edades[rendimientos.index(max(rendimientos))]
                         if rendimientos else 0,
    }


@router.get("", response_model=List[CurvaProduccion], summary="Catalogo de curvas")
async def listar(almacen: Almacenamiento) -> List[CurvaProduccion]:
    return await almacen.listar_curvas()


# OJO: va ANTES de `/{id_curva}`, que si no se traga «/bases» como si fuera un
# identificador. FastAPI resuelve por orden de declaracion.
@router.get(
    "/bases",
    summary="Curvas que pueden servir de punto de partida para una especie nueva",
)
async def bases(
    almacen: Almacenamiento,
    especie: Optional[str] = None,
    entra_en_produccion: Optional[int] = None,
    peak_kg_ha: Optional[float] = None,
    excluir: Optional[str] = None,
) -> Dict[str, Any]:
    """De que curva partir cuando una VARIEDAD no tiene la suya.

    el agrónomo planteo que el programa podria proponer una curva parecida, y
    el usuario acoto el alcance, que era lo que faltaba: *«las especies de arandanos
    que le recomiende curvas de otros arandanos, no cerezos con arandanos o
    kiwi con paltos, eso no tiene sentido»*. Tiene razon, y el propio
    `emparejar_curva` ya lo dice por escrito: *«inventarle una curva a un
    avellano porque se parece a un nogal es exactamente lo que este proyecto no
    debe hacer»*.

    Asi que se recomienda **una sola cosa: otras variedades de la MISMA
    especie**. Un arandano Emerald sin curva se parece a un arandano Duke, y
    eso no hay que inventarlo. Que un cerezo de numeros cercanos es casualidad
    aritmetica, no agronomia, y presentarlo como recomendacion seria justo el
    error que este proyecto evita con los precios (§10.2).

    El resto del catalogo se devuelve aparte y **rotulado como lo que es**: un
    esqueleto para no teclear veintiseis filas a mano, no una recomendacion.

    Los dos numeros declarados —a que anio entra en produccion y cuanto rinde—
    ordenan DENTRO de cada grupo. Ahi si valen: entre un Duke y un Legacy la
    diferencia es de rendimiento, y esa la sabe el agronomo.
    """
    curvas = await almacen.listar_curvas()
    fuera = (excluir or "").strip().lower()
    if fuera:
        curvas = [c for c in curvas
                  if c.id != excluir and c.nombre.strip().lower() != fuera]

    # El reparto usa el MISMO criterio que `emparejar_curva` -plurales,
    # calificativos-: una segunda regla de idioma aqui se separaria de la que
    # de verdad busca la curva en el informe.
    mismas, otras = curvas_de_la_misma_especie(especie, curvas)

    def describir(curva: CurvaProduccion) -> Dict[str, Any]:
        forma = forma_de_curva(curva)
        return {
            "id": curva.id,
            "nombre": curva.nombre,
            "especie": curva.especie,
            "variedad": curva.variedad,
            "derivada_de": curva.derivada_de,
            "forma": forma,
            "distancia": parecido_de_forma(
                forma, entra_en_produccion, peak_kg_ha),
        }

    def ordenar(grupo: List[CurvaProduccion]) -> List[Dict[str, Any]]:
        descritas = [describir(c) for c in grupo]
        if any(x["distancia"] is not None for x in descritas):
            # Las que no se pueden comparar van al final, no primero: una curva
            # sin forma legible no es «la mas parecida».
            descritas.sort(
                key=lambda x: (x["distancia"] is None, x["distancia"] or 0))
        else:
            descritas.sort(key=lambda x: x["nombre"].lower())
        return descritas

    hay_declarado = entra_en_produccion is not None or bool(peak_kg_ha)
    return {
        "especie": especie,
        # Lo unico que el programa se permite recomendar.
        "recomendadas": ordenar(mismas),
        "otras": ordenar(otras),
        "ordenado_por_parecido": hay_declarado,
        "titulo_recomendadas": (
            "Otras variedades de {}".format(especie) if especie
            else "Otras variedades de la misma especie"),
        "criterio": (
            "Son de la misma especie, así que el parecido es real. "
            + ("Ordenadas por cercanía a lo que usted declaró."
               if hay_declarado else
               "Escriba a qué año entra en producción y cuánto rinde para "
               "ordenarlas por cercanía a eso.")
        ),
        "sin_recomendadas": (
            "No hay ninguna otra curva de {}.".format(especie or "esta especie")
            if not mismas else None
        ),
        # Se dice en el mismo sitio donde se ofrecen: si el rotulo viviera solo
        # en la pantalla, el dia que alguien reuse este endpoint lo perderia.
        "aviso_otras": (
            "El resto del catálogo. NO es una recomendación: son otras "
            "especies y su curva no dice nada de ésta. Sirve sólo como "
            "esqueleto para no teclear las edades a mano — y hay que "
            "revisarlas todas."
        ),
    }


@router.get("/{id_curva}", response_model=CurvaProduccion, summary="Una curva")
async def obtener(id_curva: str, almacen: Almacenamiento) -> CurvaProduccion:
    try:
        return await almacen.obtener_curva(id_curva)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error


@router.post(
    "", response_model=CurvaProduccion, status_code=status.HTTP_201_CREATED,
    summary="Crea o reemplaza una curva",
)
async def guardar(curva: CurvaProduccion, almacen: Almacenamiento) -> CurvaProduccion:
    """Crear una curva nueva es lo que permite cubrir especies que no estaban.

    Si el tasador se encuentra con un tipo de palto distinto, o con una especie
    que nunca habia tasado, arma su curva aqui y queda disponible para siempre.
    """
    return await almacen.guardar_curva(curva)


@router.post(
    "/generar", response_model=CurvaProduccion,
    summary="Genera una curva por incremento proporcional a partir de 4 numeros",
)
async def generar(solicitud: SolicitudCurvaParametrica) -> CurvaProduccion:
    """No guarda nada: el tasador revisa la tabla de puntos resultante -y la
    completa con precios y costos- antes de confirmarla con POST /curvas.
    """
    # La forma del agronomo: «aguanta n anios y despues cae un X% al anio».
    # Se traduce a la que pide el generador -una edad final y un nivel- en el
    # SERVIDOR, no en el navegador: es aritmetica que alimenta el informe.
    inicio = solicitud.edad_inicio_declive
    final = solicitud.edad_final
    nivel_final = solicitud.nivel_final_kg_ha
    if solicitud.decrecimiento_anual_pct is not None:
        if inicio is None and solicitud.anios_de_plena_produccion is not None:
            inicio = (solicitud.anios_a_plena_produccion
                      + solicitud.anios_de_plena_produccion)
        if inicio is None:
            raise HTTPException(
                status.HTTP_422_UNPROCESSABLE_ENTITY,
                "con un % de decrecimiento hay que decir cuantos anios aguanta "
                "la plena produccion")
        if nivel_final is None:
            raise HTTPException(
                status.HTTP_422_UNPROCESSABLE_ENTITY,
                "con un % de decrecimiento hay que decir en que nivel queda al "
                "final de la vida util: si no, no se sabe cuando parar de bajar")
        try:
            deducido = declive_desde_porcentaje(
                solicitud.nivel_plena_produccion_kg_ha, inicio,
                solicitud.decrecimiento_anual_pct, nivel_final)
        except ValueError as error:
            raise HTTPException(
                status.HTTP_422_UNPROCESSABLE_ENTITY, str(error)) from error
        final = deducido["edad_final"]

    # Los costos, en la forma que el agronomo sabe: cuanto cuesta la hectarea
    # parada y cuanto con el huerto a full. El coeficiente variable del modelo
    # de §8.2 sale de esos dos, y se deduce ACA por lo mismo que el declive:
    # es aritmetica que alimenta el informe.
    #
    # Sin esto la curva generada salia con costo CERO -y precios cero, y 0% de
    # exportacion-, o sea proyectando produccion sin ingresos ni costos. Era el
    # «no esta mostrando datos que son fundamentales» que reporto el usuario.
    costos = solicitud.costos
    if solicitud.costo_en_plena_produccion_usd_ha is not None:
        try:
            costos = costos_desde_plena_produccion(
                solicitud.costo_base_usd_ha or 0.0,
                solicitud.costo_en_plena_produccion_usd_ha,
                solicitud.nivel_plena_produccion_kg_ha)
        except ValueError as error:
            raise HTTPException(
                status.HTTP_422_UNPROCESSABLE_ENTITY, str(error)) from error
    elif solicitud.costo_base_usd_ha is not None and costos is None:
        # Solo el costo base: es un huerto sin costo variable declarado, que
        # es raro pero valido. Mejor eso que perder el dato que si dio.
        costos = EstructuraCosto(costo_base_usd_ha=solicitud.costo_base_usd_ha)

    try:
        return curva_desde_parametros(
            nombre=solicitud.nombre,
            especie=solicitud.especie,
            variedad=solicitud.variedad,
            anios_improductivos=solicitud.anios_improductivos,
            nivel_productivo_inicial_kg_ha=solicitud.nivel_productivo_inicial_kg_ha,
            anios_a_plena_produccion=solicitud.anios_a_plena_produccion,
            nivel_plena_produccion_kg_ha=solicitud.nivel_plena_produccion_kg_ha,
            edad_inicio_declive=inicio,
            edad_final=final,
            nivel_final_kg_ha=nivel_final,
            comerciales=solicitud.comerciales,
            costos=costos,
            fuente=solicitud.fuente,
            notas=solicitud.notas,
            autor=solicitud.autor,
        )
    except ValueError as error:
        raise HTTPException(
            status.HTTP_422_UNPROCESSABLE_ENTITY, str(error)
        ) from error


@router.delete("/{id_curva}", status_code=status.HTTP_204_NO_CONTENT)
async def borrar(id_curva: str, almacen: Almacenamiento) -> None:
    try:
        await almacen.borrar_curva(id_curva)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error


@router.post(
    "/proyectar", response_model=Proyeccion,
    summary="Aplica una curva y sus ajustes a una superficie concreta",
)
async def proyeccion(
    solicitud: SolicitudProyeccion, almacen: Almacenamiento
) -> Proyeccion:
    try:
        curva = await almacen.obtener_curva(solicitud.curva_id)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error
    return proyectar(curva, solicitud)


@router.get(
    "/{id_curva}/coherencia",
    summary="Compara la curva con los umbrales de madurez configurados",
)
async def coherencia(
    id_curva: str, almacen: Almacenamiento, reglas: Reglas
) -> Dict[str, Any]:
    """Avisa si la curva y las reglas del informe se contradicen.

    Si la curva alcanza el techo a los 6 anios pero el informe declara plena
    produccion recien a los 8, uno de los dos esta mal.
    """
    try:
        curva = await almacen.obtener_curva(id_curva)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error

    formacion, crecientes = reglas.umbrales(curva.especie)
    resultado = resumen_por_madurez(curva, formacion, crecientes)
    resultado["curva"] = curva.nombre
    resultado["especie"] = curva.especie
    if not resultado["coherente"]:
        resultado["sugerencia"] = (
            "la curva llega a plena produccion a los {} anios, pero las reglas "
            "la declaran recien a los {}: revise reglas_agronomicas.json".format(
                resultado["edad_techo_curva"], crecientes + 1
            )
        )
    return resultado


# ------------------------------------------------------------- ajustes

@router.get("/ajustes/catalogo", response_model=List[Ajuste],
            summary="Ajustes guardados")
async def listar_ajustes(almacen: Almacenamiento) -> List[Ajuste]:
    return await almacen.listar_ajustes()


@router.post(
    "/ajustes", response_model=Ajuste, status_code=status.HTTP_201_CREATED,
    summary="Guarda un ajuste reutilizable",
)
async def guardar_ajuste(ajuste: Ajuste, almacen: Almacenamiento) -> Ajuste:
    return await almacen.guardar_ajuste(ajuste)


@router.delete("/ajustes/{id_ajuste}", status_code=status.HTTP_204_NO_CONTENT)
async def borrar_ajuste(id_ajuste: str, almacen: Almacenamiento) -> None:
    try:
        await almacen.borrar_ajuste(id_ajuste)
    except NoEncontrado as error:
        raise HTTPException(status.HTTP_404_NOT_FOUND, str(error)) from error
