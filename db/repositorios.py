"""Repositorios sobre PostgreSQL.

Dos piezas:

  - :class:`AlmacenPostgres` implementa el mismo contrato que ``AlmacenJSON``,
    asi que el catalogo de curvas, ajustes y formatos aprendidos pasa a la base
    sin que ningun router se entere.
  - :class:`RepositorioExpedientes` es nuevo: guarda la tasacion en curso. Es
    lo que hoy se pierde al cerrar el navegador y el motivo por el que hacia
    falta una base de datos.

La conversion entre filas y modelos Pydantic se hace aqui y solo aqui: el resto
del sistema sigue hablando en Pydantic.
"""

from __future__ import annotations

import uuid
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Sequence

from sqlalchemy import delete, func, or_, select

from db import modelos as m
from db.sesion import sesion
from schemas.curvas import (Ajuste, CurvaProduccion, LecturaIA,
                            MapeoCatastro)
from schemas.informe import Plantacion as PlantacionPydantic
from services.almacen import Almacen, NoEncontrado

__all__ = ["AlmacenPostgres", "RepositorioExpedientes"]


def _a_uuid(valor: str) -> uuid.UUID:
    """Los ids de Pydantic son hex de 12; los de la base, UUID completos."""
    try:
        return uuid.UUID(valor)
    except (ValueError, AttributeError):
        # Un id corto del catalogo JSON se estabiliza en un UUID reproducible,
        # para que migrar de JSON a Postgres no duplique las curvas.
        return uuid.uuid5(uuid.NAMESPACE_OID, "fitogreen:{}".format(valor))


# ------------------------------------------------------ catalogo de la casa

class AlmacenPostgres(Almacen):
    """Catalogo compartido del equipo, en la base."""

    tipo = "postgresql"

    # -- curvas ----------------------------------------------------------

    @staticmethod
    def _curva_a_pydantic(fila: m.Curva) -> CurvaProduccion:
        return CurvaProduccion.model_validate({
            "id": str(fila.id),
            "nombre": fila.nombre,
            "especie": fila.especie,
            "variedad": fila.variedad,
            "puntos": fila.puntos or [],
            "comerciales": fila.comerciales or {},
            "costos": fila.costos or {},
            "unidad": fila.unidad,
            "fuente": fila.fuente,
            "notas": fila.notas,
            "autor": fila.autor,
            "creada": fila.creado,
            "actualizada": fila.actualizado,
        })

    async def listar_curvas(self) -> List[CurvaProduccion]:
        async with sesion() as s:
            filas = await s.scalars(select(m.Curva).order_by(m.Curva.nombre))
            return [self._curva_a_pydantic(f) for f in filas]

    async def obtener_curva(self, id_curva: str) -> CurvaProduccion:
        async with sesion() as s:
            fila = await s.get(m.Curva, _a_uuid(id_curva))
            if fila is None:
                raise NoEncontrado("no existe la curva {}".format(id_curva))
            return self._curva_a_pydantic(fila)

    async def guardar_curva(self, curva: CurvaProduccion) -> CurvaProduccion:
        datos = curva.model_dump(mode="json")
        async with sesion() as s:
            fila = await s.get(m.Curva, _a_uuid(curva.id))
            if fila is None:
                # Puede ser un alta nueva o un nombre que ya existe: el nombre
                # es unico para que no convivan dos curvas 'Palto' distintas.
                fila = await s.scalar(
                    select(m.Curva).where(m.Curva.nombre == curva.nombre)
                )
            if fila is None:
                fila = m.Curva(id=_a_uuid(curva.id))
                s.add(fila)

            fila.nombre = datos["nombre"]
            fila.especie = datos["especie"]
            fila.variedad = datos.get("variedad")
            fila.puntos = datos["puntos"]
            fila.comerciales = datos["comerciales"]
            fila.costos = datos["costos"]
            fila.unidad = datos.get("unidad", "kg/ha")
            fila.fuente = datos.get("fuente")
            fila.notas = datos.get("notas")
            fila.autor = datos.get("autor")
            await s.flush()
            return self._curva_a_pydantic(fila)

    async def borrar_curva(self, id_curva: str) -> None:
        async with sesion() as s:
            fila = await s.get(m.Curva, _a_uuid(id_curva))
            if fila is None:
                raise NoEncontrado("no existe la curva {}".format(id_curva))
            await s.delete(fila)

    # -- ajustes ---------------------------------------------------------

    @staticmethod
    def _ajuste_a_pydantic(fila: m.Ajuste) -> Ajuste:
        return Ajuste.model_validate({
            "id": str(fila.id),
            "nombre": fila.nombre,
            "descripcion": fila.descripcion,
            "categoria": fila.categoria,
            "factor_pct": fila.factor_pct,
            "autor": fila.autor,
            "creada": fila.creado,
        })

    async def listar_ajustes(self) -> List[Ajuste]:
        async with sesion() as s:
            filas = await s.scalars(
                select(m.Ajuste).order_by(m.Ajuste.categoria, m.Ajuste.nombre)
            )
            return [self._ajuste_a_pydantic(f) for f in filas]

    async def guardar_ajuste(self, ajuste: Ajuste) -> Ajuste:
        async with sesion() as s:
            fila = await s.get(m.Ajuste, _a_uuid(ajuste.id))
            if fila is None:
                fila = await s.scalar(
                    select(m.Ajuste).where(m.Ajuste.nombre == ajuste.nombre)
                )
            if fila is None:
                fila = m.Ajuste(id=_a_uuid(ajuste.id))
                s.add(fila)

            fila.nombre = ajuste.nombre
            fila.descripcion = ajuste.descripcion
            fila.categoria = ajuste.categoria.value
            fila.factor_pct = ajuste.factor_pct
            fila.autor = ajuste.autor
            await s.flush()
            return self._ajuste_a_pydantic(fila)

    async def borrar_ajuste(self, id_ajuste: str) -> None:
        async with sesion() as s:
            fila = await s.get(m.Ajuste, _a_uuid(id_ajuste))
            if fila is None:
                raise NoEncontrado("no existe el ajuste {}".format(id_ajuste))
            await s.delete(fila)

    # -- mapeos ----------------------------------------------------------

    @staticmethod
    def _mapeo_a_pydantic(fila: m.MapeoCatastro) -> MapeoCatastro:
        return MapeoCatastro.model_validate({
            "id": str(fila.id),
            "huella": fila.huella,
            "nombre": fila.nombre,
            "columnas": fila.columnas or {},
            "encabezados": fila.encabezados or {},
            "veces_usado": fila.veces_usado,
            "creada": fila.creado,
        })

    async def listar_mapeos(self) -> List[MapeoCatastro]:
        async with sesion() as s:
            filas = await s.scalars(
                select(m.MapeoCatastro).order_by(m.MapeoCatastro.nombre)
            )
            return [self._mapeo_a_pydantic(f) for f in filas]

    async def buscar_mapeo(self, huella: str) -> Optional[MapeoCatastro]:
        async with sesion() as s:
            fila = await s.scalar(
                select(m.MapeoCatastro).where(m.MapeoCatastro.huella == huella)
            )
            return self._mapeo_a_pydantic(fila) if fila else None

    async def guardar_mapeo(self, mapeo: MapeoCatastro) -> MapeoCatastro:
        async with sesion() as s:
            fila = await s.scalar(
                select(m.MapeoCatastro).where(m.MapeoCatastro.huella == mapeo.huella)
            )
            if fila is None:
                fila = m.MapeoCatastro(id=_a_uuid(mapeo.id), huella=mapeo.huella)
                s.add(fila)
            else:
                fila.veces_usado = (fila.veces_usado or 0) + 1

            fila.nombre = mapeo.nombre
            fila.columnas = dict(mapeo.columnas)
            fila.encabezados = dict(mapeo.encabezados)
            await s.flush()
            return self._mapeo_a_pydantic(fila)


    # -- lecturas de la IA ya pagadas -------------------------------------

    @staticmethod
    def _lectura_a_pydantic(fila: m.LecturaIA) -> LecturaIA:
        return LecturaIA.model_validate({
            "id": fila.id.hex[:12] if hasattr(fila.id, "hex") else str(fila.id),
            "huella": fila.huella, "tipo": fila.tipo,
            "datos": fila.datos or {}, "modelo": fila.modelo,
            "nombre": fila.nombre, "veces_usado": fila.veces_usado or 1,
            "creada": fila.creado,
        })

    async def buscar_lectura(self, huella: str) -> Optional[LecturaIA]:
        async with sesion() as s:
            fila = await s.scalar(
                select(m.LecturaIA).where(m.LecturaIA.huella == huella)
            )
            if fila is None:
                return None
            # Se cuenta cada reuso: es lo que dice cuanta cuota se ahorro.
            fila.veces_usado = (fila.veces_usado or 0) + 1
            await s.flush()
            return self._lectura_a_pydantic(fila)

    async def guardar_lectura(self, lectura: LecturaIA) -> LecturaIA:
        async with sesion() as s:
            fila = await s.scalar(
                select(m.LecturaIA).where(m.LecturaIA.huella == lectura.huella)
            )
            if fila is None:
                fila = m.LecturaIA(id=_a_uuid(lectura.id), huella=lectura.huella)
                s.add(fila)
            fila.tipo = lectura.tipo
            fila.datos = dict(lectura.datos)
            fila.modelo = lectura.modelo
            fila.nombre = lectura.nombre
            await s.flush()
            return self._lectura_a_pydantic(fila)


# ------------------------------------------------------------ expedientes

class RepositorioUsuarios:
    """Quien puede entrar. La llave es el RUT normalizado."""

    async def buscar_por_rut(self, rut: str) -> Optional[m.Usuario]:
        async with sesion() as s:
            return await s.scalar(
                select(m.Usuario).where(m.Usuario.rut == rut))

    async def obtener(self, id_usuario: str) -> Optional[m.Usuario]:
        async with sesion() as s:
            return await s.get(m.Usuario, _a_uuid(id_usuario))

    async def listar(self) -> List[m.Usuario]:
        async with sesion() as s:
            filas = await s.scalars(
                select(m.Usuario).order_by(m.Usuario.nombre))
            return list(filas)

    async def crear(self, rut: str, nombre: str, clave: str) -> m.Usuario:
        async with sesion() as s:
            fila = m.Usuario(rut=rut, nombre=nombre, clave=clave)
            s.add(fila)
            await s.flush()
            await s.refresh(fila)
            return fila

    async def cambiar_clave(self, id_usuario: str, clave: str) -> None:
        async with sesion() as s:
            fila = await s.get(m.Usuario, _a_uuid(id_usuario))
            if fila is not None:
                fila.clave = clave
                await s.flush()

    async def marcar_entrada(self, id_usuario: str) -> None:
        async with sesion() as s:
            fila = await s.get(m.Usuario, _a_uuid(id_usuario))
            if fila is not None:
                fila.ultima_entrada = datetime.now()
                await s.flush()


class RepositorioExpedientes:
    """Alta, consulta y actualizacion de tasaciones en curso."""

    async def crear(
        self,
        cliente: str,
        rut: str,
        **campos: Any,
    ) -> m.Expediente:
        async with sesion() as s:
            expediente = m.Expediente(cliente=cliente, rut=rut, **campos)
            s.add(expediente)
            await s.flush()
            await s.refresh(expediente)
            return expediente

    async def obtener(self, id_expediente: str) -> m.Expediente:
        async with sesion() as s:
            fila = await s.get(m.Expediente, _a_uuid(id_expediente))
            if fila is None:
                raise NoEncontrado(
                    "no existe el expediente {}".format(id_expediente)
                )
            return fila

    async def listar(
        self,
        estado: Optional[str] = None,
        cliente: Optional[str] = None,
        limite: int = 50,
        usuario_id: Optional[str] = None,
    ) -> List[m.Expediente]:
        """Los mas recientes primero: lo normal es retomar lo ultimo.

        Con ``usuario_id`` se devuelve SOLO lo suyo y lo que le compartieron.
        Sin el se devuelve todo, que es lo que necesitan los scripts y las
        pruebas -pero nunca la pantalla: ahi el filtro es obligatorio-.
        """
        async with sesion() as s:
            consulta = select(m.Expediente).order_by(m.Expediente.actualizado.desc())
            if estado:
                consulta = consulta.where(m.Expediente.estado == estado)
            if cliente:
                consulta = consulta.where(m.Expediente.cliente.ilike("%{}%".format(cliente)))
            if usuario_id:
                mio = _a_uuid(usuario_id)
                compartidos = select(m.ExpedienteCompartido.expediente_id).where(
                    m.ExpedienteCompartido.usuario_id == mio)
                consulta = consulta.where(or_(
                    m.Expediente.propietario_id == mio,
                    m.Expediente.id.in_(compartidos),
                ))
            filas = await s.scalars(consulta.limit(limite))
            return list(filas)

    # -- quien puede ver que --------------------------------------------

    async def puede_verlo(self, id_expediente: str, usuario_id: str) -> bool:
        """Dueno o alguien con quien se compartio. Nada mas.

        Un informe sin dueno -de antes de que hubiera usuarios- lo ve
        cualquiera: negarselo a todos dejaria trabajo inaccesible.
        """
        async with sesion() as s:
            fila = await s.get(m.Expediente, _a_uuid(id_expediente))
            if fila is None:
                return False
            if fila.propietario_id is None:
                return True
            if str(fila.propietario_id) == str(_a_uuid(usuario_id)):
                return True
            compartido = await s.scalar(select(m.ExpedienteCompartido).where(
                m.ExpedienteCompartido.expediente_id == fila.id,
                m.ExpedienteCompartido.usuario_id == _a_uuid(usuario_id)))
            return compartido is not None

    async def compartir(self, id_expediente: str, usuario_id: str) -> None:
        async with sesion() as s:
            ya = await s.scalar(select(m.ExpedienteCompartido).where(
                m.ExpedienteCompartido.expediente_id == _a_uuid(id_expediente),
                m.ExpedienteCompartido.usuario_id == _a_uuid(usuario_id)))
            if ya is None:
                s.add(m.ExpedienteCompartido(
                    expediente_id=_a_uuid(id_expediente),
                    usuario_id=_a_uuid(usuario_id)))
                await s.flush()

    async def dejar_de_compartir(self, id_expediente: str, usuario_id: str) -> None:
        async with sesion() as s:
            fila = await s.scalar(select(m.ExpedienteCompartido).where(
                m.ExpedienteCompartido.expediente_id == _a_uuid(id_expediente),
                m.ExpedienteCompartido.usuario_id == _a_uuid(usuario_id)))
            if fila is not None:
                await s.delete(fila)

    async def traspasar(self, id_expediente: str, usuario_id: str) -> None:
        """Cambia el dueno. Al anterior se le deja compartido: pasarle un
        informe a otra persona no deberia hacerlo desaparecer de tu lista sin
        aviso."""
        async with sesion() as s:
            fila = await s.get(m.Expediente, _a_uuid(id_expediente))
            if fila is None:
                raise NoEncontrado(
                    "no existe el expediente {}".format(id_expediente))
            anterior = fila.propietario_id
            fila.propietario_id = _a_uuid(usuario_id)
            await s.flush()
        if anterior is not None and str(anterior) != str(_a_uuid(usuario_id)):
            await self.compartir(id_expediente, str(anterior))

    async def compartido_con(self, id_expediente: str) -> List[m.Usuario]:
        async with sesion() as s:
            filas = await s.scalars(
                select(m.Usuario).join(
                    m.ExpedienteCompartido,
                    m.ExpedienteCompartido.usuario_id == m.Usuario.id)
                .where(m.ExpedienteCompartido.expediente_id
                       == _a_uuid(id_expediente))
                .order_by(m.Usuario.nombre))
            return list(filas)

    async def actualizar(self, id_expediente: str, **campos: Any) -> m.Expediente:
        """Cambia lo que venga; en ``datos``, solo las claves que vengan.

        ``datos`` es UNA columna donde escriben dos duenos: la revision guarda
        ``sesion``, ``legal``, ``conflictos`` y ``fotos``, y la visita guarda
        ``visita`` y sus marcas. Reemplazarla entera hacia que cada correccion
        en la revision -y cada documento agregado despues- borrara la visita
        capturada, sin ningun aviso (verificacion del 2026-09-11). Por eso se
        fusiona: cada escritor pisa solo las claves que nombra. Para vaciar
        una, se manda con ``None``.

        La fila se lee con bloqueo, asi que la fusion se hace sobre lo que hay
        en ese momento y no sobre una lectura anterior.
        """
        async with sesion() as s:
            fila = await s.get(m.Expediente, _a_uuid(id_expediente),
                               with_for_update=True)
            if fila is None:
                raise NoEncontrado(
                    "no existe el expediente {}".format(id_expediente)
                )
            for nombre, valor in campos.items():
                if valor is None or not hasattr(fila, nombre):
                    continue
                if nombre == "datos" and isinstance(valor, dict):
                    valor = {**(fila.datos or {}), **valor}
                setattr(fila, nombre, valor)
            await s.flush()
            await s.refresh(fila)
            return fila

    async def borrar(self, id_expediente: str) -> None:
        async with sesion() as s:
            fila = await s.get(m.Expediente, _a_uuid(id_expediente))
            if fila is None:
                raise NoEncontrado(
                    "no existe el expediente {}".format(id_expediente)
                )
            await s.delete(fila)

    # -- catastro del expediente -----------------------------------------

    async def reemplazar_plantaciones(
        self, id_expediente: str, plantaciones: Sequence[PlantacionPydantic]
    ) -> int:
        """Sustituye el catastro completo del expediente.

        Se borra y se vuelve a insertar en la misma transaccion: si el tasador
        reimporta el catastro corregido, no pueden quedar filas de la version
        anterior mezcladas con las nuevas.
        """
        expediente_uuid = _a_uuid(id_expediente)
        async with sesion() as s:
            await s.execute(
                delete(m.Plantacion).where(
                    m.Plantacion.expediente_id == expediente_uuid
                )
            )
            for orden, planta in enumerate(plantaciones):
                datos = planta.model_dump(mode="json")
                s.add(m.Plantacion(
                    expediente_id=expediente_uuid,
                    orden=orden,
                    especie=datos["especie"],
                    variedad=datos.get("variedad"),
                    anio_plantacion=datos.get("anio_plantacion"),
                    superficie_ha=datos["superficie_ha"],
                    sistema_conduccion=datos.get("sistema_conduccion"),
                    predio=datos.get("predio"),
                    dominio=datos.get("dominio"),
                    riego=datos.get("riego"),
                    estado=datos.get("estado"),
                    densidad_plantacion=datos.get("densidad_plantacion"),
                ))

            total = sum(p.superficie_ha for p in plantaciones)
            expediente = await s.get(m.Expediente, expediente_uuid)
            if expediente is not None:
                expediente.superficie_total_ha = round(total, 4)
            return len(plantaciones)

    async def plantaciones(self, id_expediente: str) -> List[PlantacionPydantic]:
        """El catastro guardado, en el orden en que se importo.

        Hacia falta para cerrar el circuito: el catastro se podia guardar y
        resumir, pero no leer de vuelta, asi que el informe no podia armarse
        con el desde la web.
        """
        expediente_uuid = _a_uuid(id_expediente)
        async with sesion() as s:
            filas = (await s.execute(
                select(m.Plantacion)
                .where(m.Plantacion.expediente_id == expediente_uuid)
                .order_by(m.Plantacion.orden)
            )).scalars().all()
            return [
                PlantacionPydantic(
                    especie=f.especie, variedad=f.variedad,
                    anio_plantacion=f.anio_plantacion,
                    superficie_ha=f.superficie_ha,
                    sistema_conduccion=f.sistema_conduccion,
                    predio=f.predio, dominio=f.dominio, riego=f.riego,
                    estado=f.estado,
                    densidad_plantacion=f.densidad_plantacion,
                )
                for f in filas
            ]

    async def resumen_por_especie(self, id_expediente: str) -> List[Dict[str, Any]]:
        """Superficie agrupada por especie, calculada por la base.

        Es la consulta que justifica tener las plantaciones en tabla y no en un
        JSON: la suma la hace PostgreSQL, no Python.
        """
        async with sesion() as s:
            filas = await s.execute(
                select(
                    m.Plantacion.especie,
                    func.count().label("hileras"),
                    func.sum(m.Plantacion.superficie_ha).label("superficie_ha"),
                )
                .where(m.Plantacion.expediente_id == _a_uuid(id_expediente))
                .group_by(m.Plantacion.especie)
                .order_by(func.sum(m.Plantacion.superficie_ha).desc())
            )
            return [
                {
                    "especie": especie,
                    "hileras": hileras,
                    "superficie_ha": round(float(superficie or 0), 4),
                }
                for especie, hileras, superficie in filas
            ]

    # -- documentos y fotos ----------------------------------------------

    async def agregar_documento(self, id_expediente: str, **campos: Any) -> m.Documento:
        async with sesion() as s:
            documento = m.Documento(expediente_id=_a_uuid(id_expediente), **campos)
            s.add(documento)
            await s.flush()
            await s.refresh(documento)
            return documento

    async def agregar_fotografia(self, id_expediente: str, **campos: Any) -> m.Fotografia:
        async with sesion() as s:
            foto = m.Fotografia(expediente_id=_a_uuid(id_expediente), **campos)
            s.add(foto)
            await s.flush()
            await s.refresh(foto)
            return foto

    async def registrar_generado(self, id_expediente: str, **campos: Any) -> m.InformeGenerado:
        async with sesion() as s:
            registro = m.InformeGenerado(
                expediente_id=_a_uuid(id_expediente), **campos
            )
            s.add(registro)
            await s.flush()
            await s.refresh(registro)
            return registro

    # -- estadisticas ------------------------------------------------------

    async def estadisticas(self) -> Dict[str, Any]:
        async with sesion() as s:
            por_estado = await s.execute(
                select(m.Expediente.estado, func.count())
                .group_by(m.Expediente.estado)
            )
            total_ha = await s.scalar(
                select(func.sum(m.Expediente.superficie_total_ha))
            )
            return {
                "expedientes_por_estado": {e: c for e, c in por_estado},
                "hectareas_acumuladas": round(float(total_ha or 0), 2),
            }
