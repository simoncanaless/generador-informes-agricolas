"""Importacion de catastros de plantaciones en cualquier formato.

Dos pasos, igual que el resto del asistente:

    1. POST /catastro/analizar  -> sube la planilla, devuelve la propuesta de
                                   correspondencia de columnas para revisar
    2. POST /catastro/importar  -> con las columnas ya confirmadas, transpone
                                   las filas y calcula el cuadro resumen

El paso 1 propone; el paso 2 obedece. Si el tasador corrige una columna, el
mapeo se guarda con la huella del archivo y la proxima vez que llegue ese mismo
formato se reconoce solo.
"""

from __future__ import annotations

from datetime import datetime
import secrets
from typing import Any, Dict, List, Optional

from fastapi import Depends, APIRouter, File, Form, HTTPException, UploadFile, status
from pydantic import BaseModel, Field

from api.dependencias import obtener_usuario_actual
from api.dependencias import (
    Almacenamiento, Config, IAOpcional, Reglas, leer_acotado, motor_para,
)
from api.routers.informes import RespuestaGeneracion, _nombre_seguro
from schemas.curvas import MapeoCatastro
from schemas.informe import (
    Antecedentes, EstructuraProductiva, InformeFitogreen, Plantacion,
)
from services.catastro_import import (
    CatastroIlegible, DESCRIPCION_CAMPOS, clasificar_madurez, huella_encabezados,
    leer_catastro, proponer_mapeo,
)

# Todo lo de aqui exige sesion iniciada. La dependencia va en el router
# entero y no en cada endpoint: asi una ruta nueva nace protegida.
router = APIRouter(prefix="/catastro", tags=["catastro"],
                   dependencies=[Depends(obtener_usuario_actual)])


class ColumnaSalida(BaseModel):
    letra: str
    encabezado: str
    campo: Optional[str] = None
    origen: str
    confianza: str
    muestras: List[str] = Field(default_factory=list)


class RespuestaAnalisis(BaseModel):
    hoja: str
    fila_encabezado: int
    huella: str
    aprendido_de: Optional[str] = None
    filas_estimadas: int
    completo: bool
    columnas: List[ColumnaSalida]
    campos_disponibles: Dict[str, str]
    avisos: List[str] = Field(default_factory=list)


class LineaResumen(BaseModel):
    especie: str
    propietario_ha: float
    arrendatario_ha: float
    total_ha: float
    plena_produccion_ha: float
    rendimientos_crecientes_ha: float
    formacion_ha: float


class RespuestaImportacion(BaseModel):
    plantaciones: List[Plantacion]
    superficie_total_ha: float
    resumen: List[LineaResumen] = Field(default_factory=list)
    filas_descartadas: List[str] = Field(default_factory=list)
    avisos: List[str] = Field(default_factory=list)
    cobertura_curvas: List[Dict[str, Any]] = Field(
        default_factory=list,
        description="que especies del cliente tienen curva de produccion y "
                    "cuales no, con sus hectareas. Se calcula AQUI, al "
                    "importar, y no al generar: cuando el aviso llega al final "
                    "el tasador ya se sento a hacer el informe; aqui todavia "
                    "esta a tiempo de pedir los numeros que faltan",
    )


@router.post(
    "/analizar",
    response_model=RespuestaAnalisis,
    summary="Lee una planilla de catastro y propone como interpretarla",
)
async def analizar(
    config: Config,
    reglas: Reglas,
    ia: IAOpcional,
    almacen: Almacenamiento,
    archivo: UploadFile = File(..., description="planilla .xlsx del agricultor"),
    hoja: Optional[str] = Form(None),
) -> RespuestaAnalisis:
    contenido = await leer_acotado(archivo, config.max_bytes_pdf, "la planilla")
    try:
        propuesta = await proponer_mapeo(
            contenido, reglas, hoja=hoja, ia=ia, almacen=almacen
        )
    except CatastroIlegible as error:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(error)
        ) from error

    usados = set(propuesta.mapeo)
    return RespuestaAnalisis(
        hoja=propuesta.hoja,
        fila_encabezado=propuesta.fila_encabezado,
        huella=propuesta.huella,
        aprendido_de=propuesta.aprendido_de,
        filas_estimadas=propuesta.filas_estimadas,
        completo=propuesta.completo,
        columnas=[ColumnaSalida(**vars(c)) for c in propuesta.columnas],
        campos_disponibles=DESCRIPCION_CAMPOS,
        avisos=propuesta.avisos + (
            [] if usados else ["no se reconocio ninguna columna"]
        ),
    )


class PeticionImportacion(BaseModel):
    columnas: Dict[str, str] = Field(
        description="campo del informe -> letra de columna, ya confirmado"
    )
    hoja: Optional[str] = None
    anio_informe: int = Field(
        default=2026, ge=1990, le=2100,
        description="anio con el que se calcula la antiguedad de cada plantacion",
    )
    recordar_como: Optional[str] = Field(
        default=None,
        description="si se indica, se guarda este mapeo con ese nombre para "
                    "reconocer el formato la proxima vez",
    )


@router.post(
    "/importar",
    response_model=RespuestaImportacion,
    summary="Transpone el catastro al formato propio con las columnas confirmadas",
)
async def importar(
    config: Config,
    reglas: Reglas,
    almacen: Almacenamiento,
    archivo: UploadFile = File(...),
    columnas: str = Form(..., description="JSON con {campo: letra}"),
    hoja: Optional[str] = Form(None),
    anio_informe: int = Form(2026),
    recordar_como: Optional[str] = Form(None),
) -> RespuestaImportacion:
    import json

    contenido = await leer_acotado(archivo, config.max_bytes_pdf, "la planilla")
    try:
        mapeo = json.loads(columnas)
    except json.JSONDecodeError as error:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="el mapeo de columnas no es un JSON valido",
        ) from error

    if "especie" not in mapeo or "superficie_ha" not in mapeo:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="hay que asignar al menos las columnas de especie y superficie",
        )

    try:
        importado = leer_catastro(
            contenido, reglas, hoja=hoja, columnas_forzadas=mapeo
        )
    except CatastroIlegible as error:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(error)
        ) from error

    resumen = clasificar_madurez(importado.plantaciones, reglas, anio_informe)

    if recordar_como:
        await almacen.guardar_mapeo(MapeoCatastro(
            huella=huella_encabezados(list(importado.lectura.encabezados.values()))
            if importado.lectura.encabezados else "",
            nombre=recordar_como,
            columnas=mapeo,
            encabezados=importado.lectura.encabezados,
        ))

    # Sin curva no hay flujo proyectado para esa especie. Se falla en blando
    # -el catastro se importa igual- pero se dice, que es todo el punto.
    cobertura: List[Dict[str, Any]] = []
    try:
        from services.envejecimiento import cobertura_de_curvas
        cobertura = cobertura_de_curvas(
            importado.plantaciones, await almacen.listar_curvas())
    except Exception as error:                       # noqa: BLE001
        cobertura = []

    return RespuestaImportacion(
        cobertura_curvas=cobertura,
        plantaciones=importado.plantaciones,
        superficie_total_ha=importado.superficie_total_ha,
        resumen=[
            LineaResumen(
                especie=r.especie,
                propietario_ha=r.propietario_ha,
                arrendatario_ha=r.arrendatario_ha,
                total_ha=r.total_ha,
                plena_produccion_ha=r.plena_produccion_ha,
                rendimientos_crecientes_ha=r.rendimientos_crecientes_ha,
                formacion_ha=r.formacion_ha,
            )
            for r in resumen
        ],
        filas_descartadas=importado.filas_descartadas,
        avisos=importado.lectura.avisos,
    )


class PeticionExportacion(BaseModel):
    """Vuelca un catastro ya revisado en la plantilla de un banco."""

    plantilla: Optional[str] = Field(
        default=None,
        description="alias del formato de salida; por defecto el de la casa "
                    "para catastros (ver GET /catastro/plantillas)",
    )
    cliente: str = Field(min_length=2)
    rut: str
    plantaciones: List[Plantacion] = Field(min_length=1)


@router.get(
    "/plantillas",
    summary="Formatos de salida disponibles para volcar el catastro",
)
def plantillas(config: Config) -> Dict[str, Any]:
    """Los formatos disponibles, con el de catastro primero por ser el habitual."""
    registro = config.registro_plantillas
    predeterminada = config.plantilla_catastro
    return {
        "predeterminada": predeterminada,
        "plantillas": sorted(
            ({"alias": alias, **datos} for alias, datos in registro.items()),
            key=lambda p: (p["alias"] != predeterminada, p.get("nombre", "")),
        ),
    }


@router.post(
    "/exportar",
    response_model=RespuestaGeneracion,
    summary="Escribe el catastro importado en la plantilla del banco elegido",
)
async def exportar(
    peticion: PeticionExportacion, config: Config
) -> RespuestaGeneracion:
    """Este es el paso que cierra el circulo.

    Da igual en que formato llegara el catastro del agricultor: aqui sale
    siempre en el formato de la casa, con las formulas y los estilos de la
    plantilla intactos.
    """
    alias = peticion.plantilla or config.plantilla_catastro
    motor = motor_para(alias)

    informe = InformeFitogreen(
        antecedentes=Antecedentes(cliente=peticion.cliente, rut=peticion.rut),
        estructura_productiva=EstructuraProductiva(
            plantaciones=peticion.plantaciones
        ),
    )

    marca = datetime.now()
    nombre = "catastro_{}_{}_{}_{}.xlsx".format(
        _nombre_seguro(peticion.cliente), alias, marca.strftime("%Y%m%d_%H%M%S"),
        secrets.token_hex(8),   # el enlace no se tiene que poder adivinar
    )
    resultado = await motor.generar_async(informe, config.ruta_salida / nombre)

    return RespuestaGeneracion(
        archivo=nombre,
        descarga="/informes/{}/descargar".format(nombre),
        celdas_escritas=resultado.celdas_escritas,
        filas_insertadas=resultado.filas_insertadas,
        imagenes_insertadas=resultado.imagenes_insertadas,
        avisos=resultado.avisos,
        generado=marca,
    )


@router.get("/formatos", summary="Formatos de catastro ya aprendidos")
async def formatos(almacen: Almacenamiento) -> List[Dict[str, object]]:
    return [
        {
            "nombre": mapeo.nombre,
            "huella": mapeo.huella,
            "columnas": mapeo.columnas,
            "veces_usado": mapeo.veces_usado,
            "creada": mapeo.creada.isoformat(),
        }
        for mapeo in await almacen.listar_mapeos()
    ]
