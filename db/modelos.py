"""Modelos SQLAlchemy 2.0 del sistema de tasaciones.

Diseno hibrido, como pide la especificacion del proyecto:

  - **Relacional** lo que se consulta, se filtra o se suma: expedientes,
    plantaciones, documentos. Un tasador va a querer "todas las tasaciones de
    palto de la V Region" o "cuantas hectareas lleva este cliente", y eso en un
    JSON no se hace.
  - **JSONB** lo que cambia de forma segun el caso: la extraccion de un
    documento legal, los puntos de una curva, las notas de terreno. Meter eso
    en columnas obligaria a migrar la base cada vez que un banco pide un campo
    nuevo.

Todas las fechas se guardan con zona horaria: un informe bancario lleva fecha de
emision y discutir si eran las 23:50 de un dia o las 00:10 del siguiente es una
discusion que no queremos tener.
"""

from __future__ import annotations

import uuid
from datetime import datetime, timezone
from enum import Enum as EnumPy
from typing import Any, Dict, List, Optional

from sqlalchemy import (
    Boolean, DateTime, Float, ForeignKey, Index, Integer, String, Text,
    UniqueConstraint, func,
)
from sqlalchemy.dialects.postgresql import JSONB, UUID
from sqlalchemy.orm import (
    DeclarativeBase, Mapped, configure_mappers, mapped_column, relationship,
)


def ahora() -> datetime:
    return datetime.now(timezone.utc)


def nuevo_id() -> uuid.UUID:
    return uuid.uuid4()


class Base(DeclarativeBase):
    """Base declarativa. Alembic descubre las tablas desde aqui."""

    type_annotation_map = {Dict[str, Any]: JSONB, List[Any]: JSONB}


class EstadoExpediente(str, EnumPy):
    """Ciclo de vida de una tasacion.

    Existe para poder retomar: la pantalla de 'mis expedientes' se ordena por
    esto y el tasador ve de un vistazo cual quedo a medias.
    """

    BORRADOR = "borrador"          # se esta capturando informacion
    EN_REVISION = "en_revision"    # datos extraidos, pendientes de confirmar
    COMPLETO = "completo"          # listo para generar
    ENTREGADO = "entregado"        # ya se mando al banco
    ARCHIVADO = "archivado"


class Mezcla:
    """Columnas comunes a todas las tablas."""

    id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), primary_key=True, default=nuevo_id
    )
    creado: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), default=ahora, server_default=func.now()
    )
    actualizado: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), default=ahora, onupdate=ahora,
        server_default=func.now(),
    )


# ---------------------------------------------------------- expedientes

class Usuario(Mezcla, Base):
    """Quien usa el programa. La llave es el RUT.

    No se guarda la contrasena, solo su derivacion scrypt con sal. Ver
    services/usuarios.py.
    """

    __tablename__ = "usuarios"

    rut: Mapped[str] = mapped_column(String(20), unique=True, index=True)
    nombre: Mapped[str] = mapped_column(String(200))
    clave: Mapped[str] = mapped_column(String(300))
    activo: Mapped[bool] = mapped_column(Boolean, default=True)
    ultima_entrada: Mapped[Optional[datetime]] = mapped_column(
        DateTime(timezone=True))

    __table_args__ = (UniqueConstraint("rut", name="uq_usuario_rut"),)


#: Con quien se comparte cada informe. Tabla aparte y no una lista dentro del
#: expediente: asi se puede preguntar "que informes puede ver esta persona" con
#: un indice, y no leyendo todos los expedientes.
class ExpedienteCompartido(Mezcla, Base):
    __tablename__ = "expedientes_compartidos"

    expediente_id: Mapped[uuid.UUID] = mapped_column(
        ForeignKey("expedientes.id", ondelete="CASCADE"), index=True)
    usuario_id: Mapped[uuid.UUID] = mapped_column(
        ForeignKey("usuarios.id", ondelete="CASCADE"), index=True)

    __table_args__ = (
        UniqueConstraint("expediente_id", "usuario_id", name="uq_compartido"),
    )


class Expediente(Mezcla, Base):
    """Una tasacion. Es la unidad de trabajo que hoy se pierde al cerrar el navegador."""

    __tablename__ = "expedientes"

    # Quien lo creo. Nulo se admite a proposito: los informes anteriores a que
    # hubiera usuarios no tienen dueno, y ponerles uno inventado seria peor.
    propietario_id: Mapped[Optional[uuid.UUID]] = mapped_column(
        ForeignKey("usuarios.id", ondelete="SET NULL"), index=True)

    # -- identificacion del encargo
    cliente: Mapped[str] = mapped_column(String(300), index=True)
    rut: Mapped[str] = mapped_column(String(20), index=True)
    ejecutivo: Mapped[Optional[str]] = mapped_column(String(200))
    sucursal: Mapped[Optional[str]] = mapped_column(String(200))
    n_solicitud: Mapped[Optional[str]] = mapped_column(String(80), index=True)
    agronomo_informante: Mapped[Optional[str]] = mapped_column(String(200))
    banco: Mapped[Optional[str]] = mapped_column(String(80))

    fecha_visita: Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True))
    fecha_emision: Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True))

    estado: Mapped[str] = mapped_column(
        String(20), default=EstadoExpediente.BORRADOR.value, index=True
    )

    # -- resumen desnormalizado, para listar sin abrir el expediente entero
    superficie_total_ha: Mapped[Optional[float]] = mapped_column(Float)
    comuna: Mapped[Optional[str]] = mapped_column(String(150), index=True)
    region: Mapped[Optional[str]] = mapped_column(String(100))

    # -- lo flexible: antecedentes legales, opinion, flujo, notas
    datos: Mapped[Dict[str, Any]] = mapped_column(JSONB, default=dict)
    #: Lo que el tasador debe mirar antes de firmar (conflictos, baja confianza).
    avisos: Mapped[List[Any]] = mapped_column(JSONB, default=list)

    notas: Mapped[Optional[str]] = mapped_column(Text)

    documentos: Mapped[List["Documento"]] = relationship(
        back_populates="expediente", cascade="all, delete-orphan", lazy="selectin"
    )
    fotografias: Mapped[List["Fotografia"]] = relationship(
        back_populates="expediente", cascade="all, delete-orphan", lazy="selectin"
    )
    plantaciones: Mapped[List["Plantacion"]] = relationship(
        back_populates="expediente", cascade="all, delete-orphan",
        lazy="selectin", order_by="Plantacion.orden",
    )
    generados: Mapped[List["InformeGenerado"]] = relationship(
        back_populates="expediente", cascade="all, delete-orphan", lazy="selectin",
        order_by="InformeGenerado.creado.desc()",
    )

    __table_args__ = (
        Index("ix_expedientes_cliente_estado", "cliente", "estado"),
        # Busqueda por cualquier clave del bloque flexible sin indexar campo a campo
        Index("ix_expedientes_datos", "datos", postgresql_using="gin"),
    )


class Documento(Mezcla, Base):
    """Un PDF subido y lo que se extrajo de el."""

    __tablename__ = "documentos"

    expediente_id: Mapped[uuid.UUID] = mapped_column(
        ForeignKey("expedientes.id", ondelete="CASCADE"), index=True
    )
    nombre_archivo: Mapped[str] = mapped_column(String(400))
    tipo: Mapped[Optional[str]] = mapped_column(String(80))
    ruta: Mapped[Optional[str]] = mapped_column(String(600))
    bytes: Mapped[Optional[int]] = mapped_column(Integer)
    paginas: Mapped[Optional[int]] = mapped_column(Integer)

    #: 'texto', 'hibrido' o 'vision'. Determina cuanto fiarse de los numeros.
    via_extraccion: Mapped[Optional[str]] = mapped_column(String(20))
    modelo_ia: Mapped[Optional[str]] = mapped_column(String(80))
    degradado: Mapped[bool] = mapped_column(Boolean, default=False)

    extraccion: Mapped[Dict[str, Any]] = mapped_column(JSONB, default=dict)
    advertencias: Mapped[List[Any]] = mapped_column(JSONB, default=list)
    #: Queda en false hasta que un humano revisa los datos de este documento.
    confirmado: Mapped[bool] = mapped_column(Boolean, default=False)

    expediente: Mapped[Expediente] = relationship(back_populates="documentos")


class Fotografia(Mezcla, Base):
    """Una foto de terreno con su pie tecnico."""

    __tablename__ = "fotografias"

    expediente_id: Mapped[uuid.UUID] = mapped_column(
        ForeignKey("expedientes.id", ondelete="CASCADE"), index=True
    )
    nombre_archivo: Mapped[str] = mapped_column(String(400))
    ruta: Mapped[Optional[str]] = mapped_column(String(600))
    titulo: Mapped[Optional[str]] = mapped_column(String(400))
    descripcion: Mapped[Optional[str]] = mapped_column(Text)
    categoria: Mapped[str] = mapped_column(String(40), default="otra")
    incluir: Mapped[bool] = mapped_column(Boolean, default=True)
    orden: Mapped[int] = mapped_column(Integer, default=0)
    modelo_ia: Mapped[Optional[str]] = mapped_column(String(80))

    expediente: Mapped[Expediente] = relationship(back_populates="fotografias")


class Plantacion(Mezcla, Base):
    """Una hilera del catastro.

    Relacional y no JSON a proposito: aqui se suman hectareas, se agrupa por
    especie y se clasifica por edad. Es la tabla que mas se consulta.
    """

    __tablename__ = "plantaciones"

    expediente_id: Mapped[uuid.UUID] = mapped_column(
        ForeignKey("expedientes.id", ondelete="CASCADE"), index=True
    )
    orden: Mapped[int] = mapped_column(Integer, default=0)

    especie: Mapped[str] = mapped_column(String(150), index=True)
    variedad: Mapped[Optional[str]] = mapped_column(String(150))
    #: Texto y no entero: los catastros reales traen rangos ('2017-2018').
    anio_plantacion: Mapped[Optional[str]] = mapped_column(String(40))
    superficie_ha: Mapped[float] = mapped_column(Float)
    sistema_conduccion: Mapped[Optional[str]] = mapped_column(String(150))
    predio: Mapped[Optional[str]] = mapped_column(String(300), index=True)
    dominio: Mapped[Optional[str]] = mapped_column(String(40))
    riego: Mapped[Optional[str]] = mapped_column(String(150))
    estado: Mapped[Optional[str]] = mapped_column(String(40))
    densidad_plantacion: Mapped[Optional[str]] = mapped_column(String(100))

    #: Curva aplicada y ajustes puestos sobre ella, si se proyecto.
    curva_id: Mapped[Optional[uuid.UUID]] = mapped_column(
        ForeignKey("curvas.id", ondelete="SET NULL")
    )
    ajustes: Mapped[List[Any]] = mapped_column(JSONB, default=list)

    expediente: Mapped[Expediente] = relationship(back_populates="plantaciones")


class InformeGenerado(Mezcla, Base):
    """Traza de cada Excel emitido. Un informe bancario tiene que ser auditable."""

    __tablename__ = "informes_generados"

    expediente_id: Mapped[uuid.UUID] = mapped_column(
        ForeignKey("expedientes.id", ondelete="CASCADE"), index=True
    )
    archivo: Mapped[str] = mapped_column(String(400))
    plantilla: Mapped[str] = mapped_column(String(80))
    celdas_escritas: Mapped[int] = mapped_column(Integer, default=0)
    filas_insertadas: Mapped[int] = mapped_column(Integer, default=0)
    imagenes_insertadas: Mapped[int] = mapped_column(Integer, default=0)
    avisos: Mapped[List[Any]] = mapped_column(JSONB, default=list)

    # El lado inverso de Expediente.generados. Sin el, SQLAlchemy no puede
    # resolver el back_populates y falla al configurar TODOS los mapeos, no
    # solo este: por eso el error saltaba en la primera consulta que se hiciera.
    expediente: Mapped[Expediente] = relationship(back_populates="generados")


# ------------------------------------------------- catalogo de la casa

class Curva(Mezcla, Base):
    """Curva de produccion de una especie. Del equipo, no de un expediente."""

    __tablename__ = "curvas"

    nombre: Mapped[str] = mapped_column(String(200), unique=True)
    especie: Mapped[str] = mapped_column(String(150), index=True)
    variedad: Mapped[Optional[str]] = mapped_column(String(150))

    #: [{edad, rendimiento_kg_ha, costo_directo_usd_ha}, ...]
    puntos: Mapped[List[Any]] = mapped_column(JSONB, default=list)
    comerciales: Mapped[Dict[str, Any]] = mapped_column(JSONB, default=dict)
    costos: Mapped[Dict[str, Any]] = mapped_column(JSONB, default=dict)

    unidad: Mapped[str] = mapped_column(String(20), default="kg/ha")
    fuente: Mapped[Optional[str]] = mapped_column(Text)
    notas: Mapped[Optional[str]] = mapped_column(Text)
    autor: Mapped[Optional[str]] = mapped_column(String(200))


class Ajuste(Mezcla, Base):
    """Factor reutilizable que sube o baja una curva segun lo visto en terreno."""

    __tablename__ = "ajustes"

    nombre: Mapped[str] = mapped_column(String(200), unique=True)
    descripcion: Mapped[Optional[str]] = mapped_column(Text)
    categoria: Mapped[str] = mapped_column(String(40), default="otro", index=True)
    factor_pct: Mapped[float] = mapped_column(Float)
    autor: Mapped[Optional[str]] = mapped_column(String(200))


class LecturaIA(Mezcla, Base):
    """Lo que el modelo ya leyo de un archivo. Ver schemas.curvas.LecturaIA.

    La huella lleva indice UNICO: es la clave por la que se busca antes de
    gastar cuota, y dos filas con la misma huella significarian pagar dos veces
    lo mismo sin saberlo.
    """

    __tablename__ = "lecturas_ia"

    huella: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    tipo: Mapped[str] = mapped_column(String(20), index=True)
    datos: Mapped[Dict[str, Any]] = mapped_column(JSONB, default=dict)
    modelo: Mapped[Optional[str]] = mapped_column(String(80))
    nombre: Mapped[Optional[str]] = mapped_column(String(300))
    veces_usado: Mapped[int] = mapped_column(Integer, default=1)

    __table_args__ = (UniqueConstraint("huella", name="uq_lectura_huella"),)


class MapeoCatastro(Mezcla, Base):
    """Formato de catastro ya reconocido, para no volver a mapearlo a mano."""

    __tablename__ = "mapeos_catastro"

    huella: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    nombre: Mapped[str] = mapped_column(String(300))
    columnas: Mapped[Dict[str, Any]] = mapped_column(JSONB, default=dict)
    encabezados: Mapped[Dict[str, Any]] = mapped_column(JSONB, default=dict)
    veces_usado: Mapped[int] = mapped_column(Integer, default=1)

    __table_args__ = (UniqueConstraint("huella", name="uq_mapeo_huella"),)


# SQLAlchemy resuelve las relaciones de forma perezosa: un back_populates sin su
# lado inverso no se detecta al importar, sino en la primera consulta que se
# haga... y entonces revienta TODOS los mapeos, no solo el mal declarado, con un
# error que no apunta al modelo culpable. Forzar la configuracion aqui convierte
# eso en un fallo inmediato y localizado al arrancar.
configure_mappers()
