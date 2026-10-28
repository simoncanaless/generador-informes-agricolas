"""Evaluacion de curvas de produccion y aplicacion de los ajustes del tasador.

Todo lo de aqui es aritmetica pura y reproducible. La IA no participa: una
proyeccion de ingresos que alimenta una tasacion bancaria tiene que dar el mismo
numero las mil veces que se recalcule, y tiene que poder explicarse linea a
linea si el banco pregunta.

El modelo economico sale de 'HdU - Plantation Dec-2025 vf Esp.xlsx' (8.1 a 8.4):

    rendimiento(edad)  -> curva de la especie, en kg/ha
    rendimiento_real    = rendimiento(edad) x factor de los ajustes
    ingreso/ha          = rendimiento_real x precio ponderado
    precio ponderado    = %export x RNP export + %interno x RNP interno
    costo/ha            = costo base(edad) + rendimiento_real x costo por kg
    margen/ha           = ingreso/ha - costo/ha

Que el costo variable dependa del rendimiento YA AJUSTADO no es un detalle: si
el tasador baja la curva un 20% porque el huerto esta mal manejado, tambien baja
lo que cuesta cosechar. Calcularlo sobre la curva de manual inflaria el costo.
"""

from __future__ import annotations

import math
from typing import Any, Dict, List, Optional, Sequence

from schemas.curvas import (
    AjusteAplicado, CurvaProduccion, EstructuraCosto, ParametrosComerciales,
    Proyeccion, PuntoCurva, PuntoProyectado, SolicitudProyeccion,
)

__all__ = [
    "rendimiento_a_edad", "costo_base_a_edad", "factor_de_ajustes",
    "proyectar", "resumen_por_madurez", "curva_desde_parametros",
    "declive_desde_porcentaje", "forma_de_curva", "parecido_de_forma",
    "costos_desde_plena_produccion",
]


def _interpolar(puntos: Sequence[PuntoCurva], edad: int, atributo: str) -> float:
    """Valor de la curva a una edad, interpolando entre los puntos declarados.

    Fuera de rango se prolonga con el extremo mas cercano: antes del primer
    punto, el primero; despues del ultimo, el ultimo. Extrapolar una tendencia
    a 40 anios cuando la tabla llega a 25 seria inventarse datos.
    """
    if not puntos:
        return 0.0

    if edad <= puntos[0].edad:
        return float(getattr(puntos[0], atributo) or 0.0)
    if edad >= puntos[-1].edad:
        return float(getattr(puntos[-1], atributo) or 0.0)

    for anterior, siguiente in zip(puntos, puntos[1:]):
        if anterior.edad <= edad <= siguiente.edad:
            valor_a = float(getattr(anterior, atributo) or 0.0)
            valor_b = float(getattr(siguiente, atributo) or 0.0)
            if siguiente.edad == anterior.edad:
                return valor_a
            avance = (edad - anterior.edad) / (siguiente.edad - anterior.edad)
            return valor_a + (valor_b - valor_a) * avance
    return 0.0


def rendimiento_a_edad(curva: CurvaProduccion, edad: int) -> float:
    """kg/ha de la curva base a la edad dada."""
    return round(_interpolar(curva.puntos, edad, "rendimiento_kg_ha"), 4)


def costo_base_a_edad(curva: CurvaProduccion, edad: int) -> float:
    """Parte fija del costo a la edad dada.

    Si la curva no trae costos por edad se usa el costo base de la estructura,
    que es lo mismo para todas las edades.
    """
    if any(p.costo_directo_usd_ha is not None for p in curva.puntos):
        return round(_interpolar(curva.puntos, edad, "costo_directo_usd_ha"), 4)
    return curva.costos.costo_base_usd_ha


def factor_de_ajustes(ajustes: Sequence[AjusteAplicado]) -> float:
    """Combina los ajustes multiplicando, no sumando.

    Dos ajustes de -10% dejan el 81% (0,9 x 0,9), no el 80%. Multiplicar es lo
    correcto porque cada factor actua sobre lo que queda tras el anterior, y
    ademas impide que tres ajustes de -40% den un rendimiento negativo.
    """
    factor = 1.0
    for ajuste in ajustes:
        factor *= ajuste.multiplicador
    return round(max(factor, 0.0), 6)


def proyectar(
    curva: CurvaProduccion, solicitud: SolicitudProyeccion
) -> Proyeccion:
    """Aplica la curva y los ajustes a una superficie concreta."""
    factor = factor_de_ajustes(solicitud.ajustes)
    precio = curva.comerciales.precio_ponderado_usd_kg()
    superficie = solicitud.superficie_ha

    edades = (
        [solicitud.edad_actual]
        if solicitud.edad_actual is not None
        else list(range(0, solicitud.horizonte + 1))
    )

    puntos: List[PuntoProyectado] = []
    for edad in edades:
        base = rendimiento_a_edad(curva, edad)
        ajustado = round(base * factor, 4)
        produccion = round(ajustado * superficie, 4)
        ingreso = round(produccion * precio, 2)

        costo_ha = round(
            costo_base_a_edad(curva, edad)
            + ajustado * curva.costos.costo_por_kg_usd,
            4,
        )
        costo = round(costo_ha * superficie, 2)

        puntos.append(PuntoProyectado(
            edad=edad,
            rendimiento_base_kg_ha=base,
            rendimiento_ajustado_kg_ha=ajustado,
            produccion_total_kg=produccion,
            ingreso_usd=ingreso,
            costo_usd=costo,
            margen_usd=round(ingreso - costo, 2),
        ))

    return Proyeccion(
        curva=curva.nombre,
        superficie_ha=superficie,
        ajustes=list(solicitud.ajustes),
        factor_total=factor,
        puntos=puntos,
    )


def resumen_por_madurez(
    curva: CurvaProduccion, formacion_hasta: int, crecientes_hasta: int
) -> Dict[str, Optional[int]]:
    """Traduce la curva a los tramos que usa el informe.

    Permite comprobar si los umbrales configurados en reglas_agronomicas.json
    tienen sentido para esta especie: si la curva llega al techo a los 6 anios,
    declarar plena produccion a partir de los 8 esta subestimando el huerto.
    """
    techo = curva.rendimiento_maximo or 1.0
    primera_plena = next(
        (p.edad for p in curva.puntos if p.rendimiento_kg_ha >= techo * 0.95), None
    )
    return {
        "edad_techo_curva": primera_plena,
        "umbral_formacion": formacion_hasta,
        "umbral_crecientes": crecientes_hasta,
        "coherente": (
            primera_plena is not None and primera_plena <= crecientes_hasta + 1
        ),
    }


def costos_desde_plena_produccion(
    costo_base_usd_ha: float,
    costo_en_plena_produccion_usd_ha: float,
    nivel_plena_produccion_kg_ha: float,
) -> EstructuraCosto:
    """Traduce lo que el agronomo SABE al modelo que usa el informe.

    El modelo de costo de Fitogreen es `base + variable x rendimiento` (§8.2),
    pero nadie tiene en la cabeza el «0,50 USD por kilo»: lo que si sabe un
    agronomo es **cuanto le cuesta la hectarea cuando el huerto esta a full**.
    Lo pidio el usuario: *«agregar poder poner cual es el costo de cuando llega a
    plena produccion»*.

    De los dos extremos sale el coeficiente::

        variable = (costo en plena produccion - costo base) / nivel de plena

    Contrastado con el palto del HdU, que es el ejemplo de §8.2: base 6.000,
    en plena 12.500, nivel 13.000 -> variable 0,50 USD/kg, que es exactamente
    el numero tabulado. O sea que la tabla que el proyecto ya usaba se puede
    expresar asi.
    """
    if nivel_plena_produccion_kg_ha <= 0:
        raise ValueError("el nivel de plena produccion debe ser mayor que 0")
    if costo_base_usd_ha < 0 or costo_en_plena_produccion_usd_ha < 0:
        raise ValueError("los costos no pueden ser negativos")
    if costo_en_plena_produccion_usd_ha < costo_base_usd_ha:
        # Con el huerto produciendo no puede costar MENOS que parado: seria un
        # costo variable negativo, y la curva bajaria de costo al subir de
        # rendimiento. Es un dato mal tecleado, no un caso raro.
        raise ValueError(
            "el costo en plena produccion no puede ser menor que el costo "
            "base: el huerto produciendo no cuesta menos que parado")
    return EstructuraCosto(
        costo_base_usd_ha=costo_base_usd_ha,
        costo_por_kg_usd=round(
            (costo_en_plena_produccion_usd_ha - costo_base_usd_ha)
            / nivel_plena_produccion_kg_ha, 6),
    )


def forma_de_curva(curva: CurvaProduccion) -> Dict[str, Any]:
    """Los cuatro numeros que describen una curva, para poder compararlas.

    Existe porque el catalogo se elige por NOMBRE y eso no dice nada: «Cerezos»
    y «Nogales» son dos palabras y detras hay 13.000 y 6.000 kg/ha con dos
    anios de diferencia en entrar a producir. Puesta la forma al lado del
    nombre, el agronomo elige mirando la curva y no la etiqueta.

    Va en el servidor y no en el navegador por la regla de §2: es aritmetica
    sobre la curva que alimenta el informe, y dos versiones del mismo «entra en
    produccion al anio 3» terminarian separandose.
    """
    puntos = sorted(curva.puntos, key=lambda p: p.edad)
    techo = max((p.rendimiento_kg_ha for p in puntos), default=0.0)

    # «Entra en produccion» es el primer anio con algo distinto de cero, no el
    # primero declarado: una curva puede empezar en la edad 0 con cero.
    entra = next((p.edad for p in puntos if p.rendimiento_kg_ha > 0), None)
    # Plena produccion: el 95% del techo, el mismo corte que `resumen_por_madurez`
    # usa para decir si los umbrales del informe son coherentes. Un segundo
    # criterio para lo mismo seria dos definiciones de plena produccion.
    plena = next(
        (p.edad for p in puntos if techo and p.rendimiento_kg_ha >= techo * 0.95),
        None,
    )
    ultimo = puntos[-1] if puntos else None
    return {
        "entra_en_produccion": entra,
        "edad_plena_produccion": plena,
        "peak_kg_ha": round(techo, 1),
        "edad_ultima": ultimo.edad if ultimo else None,
        "nivel_final_kg_ha": round(ultimo.rendimiento_kg_ha, 1) if ultimo else None,
        # Si la curva termina bajo su techo es que modela senescencia; si no,
        # se queda en la meseta para siempre y eso hay que poder verlo (§16.b).
        "tiene_declive": bool(
            ultimo is not None and techo and ultimo.rendimiento_kg_ha < techo * 0.98
        ),
        "puntos": len(puntos),
    }


def parecido_de_forma(
    forma: Dict[str, Any],
    entra_en_produccion: Optional[int] = None,
    peak_kg_ha: Optional[float] = None,
) -> Optional[float]:
    """Cuanto se aleja una curva de lo que el agronomo DECLARA esperar.

    **El programa no sabe que especies se parecen** —no hay ninguna tabla que
    lo diga, y Kc es demanda de agua, no rendimiento— asi que no inventa el
    parecido: lo ordena contra los numeros que pone la persona. El criterio
    agronomico lo pone el, el programa solo busca cual de las curvas que ya
    existen se acerca mas.

    Devuelve una distancia (menor = mas parecida) o None si no se declaro nada
    con que comparar. Los dos ejes se normalizan antes de sumarse: un anio de
    diferencia y mil kilos no son comparables en crudo.
    """
    partes = []
    if entra_en_produccion is not None and forma.get("entra_en_produccion") is not None:
        # Un anio de diferencia al entrar en produccion pesa como un 25% de
        # diferencia en el peak: entrar dos anios antes cambia el flujo mucho.
        partes.append(abs(forma["entra_en_produccion"] - entra_en_produccion) * 0.25)
    if peak_kg_ha and forma.get("peak_kg_ha"):
        partes.append(abs(forma["peak_kg_ha"] - peak_kg_ha) / max(peak_kg_ha, 1.0))
    return round(sum(partes), 4) if partes else None


def declive_desde_porcentaje(
    nivel_plena_kg_ha: float,
    edad_inicio_declive: int,
    decrecimiento_anual_pct: float,
    nivel_base_kg_ha: float,
) -> Dict[str, Any]:
    """Cuantos anios tarda en caer del peak al nivel base, a X% por anio.

    Es la forma en que lo piensa el agronomo. el agrónomo lo dijo asi el
    2026-09-08: *"cualquier curva de produccion se puede construir con: anio de
    la primera produccion comercial; volumen 1a produccion comercial; anio
    estimado de plena produccion; nivel de plena produccion; **anios de plena
    produccion (n)**; **% de decrecimiento productivo a partir del anio n**;
    nivel productivo base al final de la vida util"*.

    `curva_desde_parametros` pide la EDAD en que termina de bajar; esto la
    deduce del porcentaje, que es el dato que el tasador tiene. Nadie sabe de
    memoria a que edad un nogal llega a 8.000 kg/ha, pero si sabe que cae "como
    un 4% al anio".

    Decaimiento compuesto: cada anio se pierde el X% de lo que quedaba, no del
    peak. Es lo que hace una planta -el declive es proporcional al vigor que
    le queda- y ademas nunca cruza el cero, que la resta lineal si haria.
    """
    if nivel_plena_kg_ha <= 0:
        raise ValueError("el nivel de plena produccion tiene que ser mayor que cero")
    if not 0 < decrecimiento_anual_pct < 100:
        raise ValueError(
            "el decrecimiento anual va entre 0 y 100%, y no puede ser ninguno "
            "de los dos: con 0% no baja nunca y con 100% se muere en un anio")
    if nivel_base_kg_ha >= nivel_plena_kg_ha:
        raise ValueError(
            "el nivel al final de la vida util tiene que ser MENOR que el de "
            "plena produccion: si no, no hay declive que describir")
    if nivel_base_kg_ha <= 0:
        raise ValueError(
            "el nivel base tiene que ser mayor que cero: un huerto viejo "
            "produce poco, no nada")

    factor = 1.0 - decrecimiento_anual_pct / 100.0
    anios = math.ceil(
        math.log(nivel_base_kg_ha / nivel_plena_kg_ha) / math.log(factor))
    return {
        "edad_final": int(edad_inicio_declive + anios),
        "nivel_final_kg_ha": round(float(nivel_base_kg_ha), 4),
        "anios_de_declive": int(anios),
        "vida_util_anios": int(edad_inicio_declive + anios),
    }


def curva_desde_parametros(
    *,
    nombre: str,
    especie: str,
    nivel_plena_produccion_kg_ha: float,
    anios_a_plena_produccion: int,
    anios_improductivos: int,
    nivel_productivo_inicial_kg_ha: float,
    edad_inicio_declive: Optional[int] = None,
    edad_final: Optional[int] = None,
    nivel_final_kg_ha: Optional[float] = None,
    variedad: Optional[str] = None,
    comerciales: Optional[ParametrosComerciales] = None,
    costos: Optional[EstructuraCosto] = None,
    fuente: Optional[str] = None,
    notas: Optional[str] = None,
    autor: Optional[str] = None,
) -> CurvaProduccion:
    """Arma una curva por incremento proporcional a partir de 4 numeros.

    Propuesta del agronomo para no tener que tabular la curva punto a punto
    por cada especie y variedad (la uva de mesa llega a plena produccion en
    3 temporadas, el nogal en 6-7; una tabla fija no escala a eso). Da:
    edad 0 -> 0, se mantiene en 0 hasta el ultimo anio improductivo, sube al
    nivel inicial en el primer anio productivo, y desde ahi interpola lineal
    hasta el nivel de plena produccion. ``rendimiento_a_edad`` y ``proyectar``
    no cambian: usan la misma interpolacion y el mismo sostenimiento del
    ultimo valor fuera de rango que ya tenian.

    **El declive por senescencia es opcional, y de el depende que la curva sea
    honesta en huertos viejos.** Un frutal no se queda en su peak para siempre:
    sostiene la meseta unos anios y despues baja. Las curvas tabuladas del HdU
    si lo modelan -el palto aguanta 13.000 kg/ha hasta los 10, cae a 12.400 a
    los 15 y a 10.100 desde los 20-, y sin esto la version parametrica le
    proyectaba a un huerto de 30 anios el rendimiento de uno en plenitud.

    Se declara con tres datos, y o van los tres o no va ninguno::

        edad_inicio_declive   hasta que edad aguanta el peak
        edad_final            cuando termina de caer
        nivel_final_kg_ha     en cuanto queda

    Pasada ``edad_final`` la curva se sostiene en el nivel final -es lo que
    hace ``_interpolar`` fuera de rango-: no se extrapola la caida hasta cero,
    porque un huerto viejo produce poco pero no deja de producir.
    """
    if anios_improductivos < 0:
        raise ValueError("los anios improductivos no pueden ser negativos")
    if anios_a_plena_produccion <= anios_improductivos:
        raise ValueError(
            "los anios a plena produccion deben ser mayores que los "
            "anios improductivos"
        )
    if nivel_plena_produccion_kg_ha <= 0:
        raise ValueError("el nivel de plena produccion debe ser mayor que 0")
    if not (0 <= nivel_productivo_inicial_kg_ha <= nivel_plena_produccion_kg_ha):
        raise ValueError(
            "el nivel productivo inicial debe estar entre 0 y el nivel de "
            "plena produccion"
        )

    puntos_por_edad: Dict[int, float] = {}
    if anios_improductivos > 0:
        puntos_por_edad[0] = 0.0
        if anios_improductivos > 1:
            puntos_por_edad[anios_improductivos - 1] = 0.0
        puntos_por_edad[anios_improductivos] = nivel_productivo_inicial_kg_ha
    else:
        puntos_por_edad[0] = nivel_productivo_inicial_kg_ha

    puntos_por_edad[anios_a_plena_produccion] = nivel_plena_produccion_kg_ha

    declive = (edad_inicio_declive, edad_final, nivel_final_kg_ha)
    if any(v is not None for v in declive):
        if any(v is None for v in declive):
            raise ValueError(
                "el declive necesita los tres datos: hasta que edad aguanta el "
                "peak, cuando termina de caer y en cuanto queda"
            )
        if edad_inicio_declive < anios_a_plena_produccion:
            raise ValueError(
                "el declive no puede empezar antes de la plena produccion"
            )
        if edad_final <= edad_inicio_declive:
            raise ValueError(
                "la edad final del declive debe ser mayor que la de inicio"
            )
        if not (0 <= nivel_final_kg_ha <= nivel_plena_produccion_kg_ha):
            raise ValueError(
                "el nivel final debe estar entre 0 y el de plena produccion"
            )
        # La meseta se sostiene hasta el inicio del declive, y de ahi cae.
        puntos_por_edad[edad_inicio_declive] = nivel_plena_produccion_kg_ha
        puntos_por_edad[edad_final] = nivel_final_kg_ha

    puntos = [
        PuntoCurva(edad=edad, rendimiento_kg_ha=rendimiento)
        for edad, rendimiento in sorted(puntos_por_edad.items())
    ]

    return CurvaProduccion(
        nombre=nombre,
        especie=especie,
        variedad=variedad,
        puntos=puntos,
        comerciales=comerciales or ParametrosComerciales(),
        costos=costos or EstructuraCosto(),
        fuente=fuente,
        notas=notas,
        autor=autor,
    )
