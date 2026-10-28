"""Ajustes centrales del backend.

Unica fuente de verdad para rutas, claves y parametros de ejecucion. Ningun
modulo vuelve a construir una ruta a mano ni lee os.environ por su cuenta.

Las listas se declaran como texto separado por comas y se exponen ya troceadas
por propiedades: pydantic-settings intenta interpretar los campos complejos
como JSON, y "a,b,c" no es JSON valido.
"""

from __future__ import annotations

import json
import unicodedata
from functools import lru_cache
from pathlib import Path
from typing import Dict, List, Tuple

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict

#: Raiz del proyecto (este archivo vive en <raiz>/config/).
RAIZ = Path(__file__).resolve().parent.parent


def _trocear(texto: str) -> List[str]:
    """'a, b , ,c' -> ['a', 'b', 'c']."""
    return [parte.strip() for parte in texto.split(",") if parte.strip()]


def resolver_ruta(ruta: Path) -> Path:
    """Encuentra el archivo aunque la tilde este codificada de otra forma.

    'Información' se puede guardar como un caracter 'ó' (NFC) o como una 'o'
    seguida de una tilde combinante (NFD). Son cadenas distintas para Python,
    apuntan al mismo archivo, y cual de las dos llega depende de si el archivo
    se bajo de un correo, se copio de un Mac o se escribio a mano. Comparar
    normalizado evita un 'archivo no encontrado' incomprensible.
    """
    if ruta.exists():
        return ruta

    carpeta = ruta.parent
    if not carpeta.is_dir():
        return ruta

    buscado = unicodedata.normalize("NFC", ruta.name)
    for candidato in carpeta.iterdir():
        if unicodedata.normalize("NFC", candidato.name) == buscado:
            return candidato
    return ruta


class Ajustes(BaseSettings):
    """Configuracion leida de .env y del entorno."""

    model_config = SettingsConfigDict(
        env_file=RAIZ / ".env",
        env_file_encoding="utf-8",
        extra="ignore",
        case_sensitive=False,
    )

    # ------------------------------------------------------------- rutas
    raiz: Path = RAIZ
    ruta_plantilla: Path = RAIZ / "templates" / "Formato Fitogreen Bsan.xlsx"
    ruta_mapping: Path = RAIZ / "config" / "mapping_santander.json"
    ruta_formulario_visita: Path = RAIZ / "config" / "formulario_visita.json"
    ruta_parametros_riego: Path = RAIZ / "config" / "parametros_riego.json"
    ruta_precios_referencia: Path = RAIZ / "config" / "precios_referencia.json"
    ruta_salida: Path = RAIZ / "output"
    #: Los scripts de prueba escriben AQUI y no en output/, para que la carpeta
    #: de entregables no se llene de archivos desechables. Se puede borrar
    #: entera en cualquier momento sin perder nada.
    ruta_salida_pruebas: Path = RAIZ / "output" / "pruebas"
    ruta_muestras: Path = RAIZ / "data" / "samples"
    ruta_subidas: Path = RAIZ / "data" / "uploads"

    # --------------------------------------------------------------- IA
    gemini_api_keys: str = Field(default="", description="claves separadas por coma")

    #: Con esto se FIRMA la cookie de sesion. Si cambia, todos vuelven a entrar
    #: -que es lo que hay que hacer si alguna vez se filtra-. Se genera sola y
    #: se guarda en data/ la primera vez, para que nadie tenga que inventarla.
    clave_sesion: str = Field(default="", description="firma de la cookie de sesion")
    horas_de_sesion: int = Field(
        default=12, description="cuanto dura la sesion sin volver a entrar")
    gemini_modelos: str = Field(
        default=(
            "gemini-3.6-flash,gemini-3.5-flash,gemini-2.5-flash,"
            "gemini-3.5-flash-lite,gemini-3.1-flash-lite,gemini-flash-latest"
        ),
        description="cadena de modelos en orden de preferencia",
    )
    gemini_cooldown_base: float = 60.0
    gemini_cooldown_maximo: float = 3600.0
    gemini_reintentos_transitorios: int = 2
    gemini_timeout_segundos: float = 120.0

    # ----------------------------------------------------- base de datos
    database_url: str = Field(
        default="",
        description="cadena asyncpg; vacia = persistencia en archivos JSON",
    )
    db_crear_al_arrancar: bool = True
    db_echo: bool = False
    db_pool_min: int = 1
    db_pool_max: int = 10

    # -------------------------------------------------------------- API
    api_titulo: str = "Fitogreen Tasaciones API"
    api_version: str = "0.2.0"
    api_entorno: str = "desarrollo"
    #: Otros sitios web que pueden llamar a la API con la sesion del usuario.
    #: Vacio a proposito: la pagina y la API salen del mismo servidor, asi que
    #: no hace falta ninguno. Antes traia "http://localhost:3000", un resto de
    #: plantilla: cualquier programa que levantara una pagina en ese puerto de
    #: este computador podia hablar con la API como si fuera el tasador.
    api_cors_origenes: str = ""

    #: Desde donde se puede entrar (ver api/seguridad.py). Por defecto, este
    #: computador y los equipos de Tailscale; el wifi de la casa, no.
    redes_permitidas: str = Field(
        default="127.0.0.0/8,::1/128,100.64.0.0/10,fd7a:115c:a1e0::/48",
        description="redes desde las que se acepta entrar, separadas por coma",
    )

    # ---------------------------------------------------------- limites
    max_mb_pdf: int = 25
    max_mb_imagen: int = 15

    # ------------------------------------------------------ propiedades

    @property
    def claves_gemini(self) -> List[str]:
        return _trocear(self.gemini_api_keys)

    @property
    def modelos_gemini(self) -> List[str]:
        return _trocear(self.gemini_modelos)

    @property
    def origenes_cors(self) -> List[str]:
        return _trocear(self.api_cors_origenes)

    @property
    def bd_configurada(self) -> bool:
        """Hay PostgreSQL utilizable.

        Se comprueba que no quedo el marcador de plantilla sin reemplazar: es
        el error mas facil de cometer al copiar el .env.example.
        """
        url = self.database_url.strip()
        return bool(url) and "CAMBIAME" not in url and "TU_CONTRASENA" not in url

    @property
    def nombre_base_datos(self) -> str:
        """Ultimo segmento de la URL, para poder crear la base si no existe."""
        return self.database_url.rsplit("/", 1)[-1].split("?")[0] if self.database_url else ""

    def url_sin_base(self) -> str:
        """La misma URL apuntando a 'postgres', para crear la base del proyecto."""
        if "/" not in self.database_url:
            return self.database_url
        return self.database_url.rsplit("/", 1)[0] + "/postgres"

    def secreto_de_sesion(self) -> str:
        """La clave con que se firma la cookie, generandola la primera vez.

        Se guarda en un archivo aparte y no en el .env: es un secreto que el
        programa se da a si mismo, no un ajuste que alguien deba escribir. Si
        el archivo no se puede escribir se usa una clave de proceso: las
        sesiones no sobreviven a un reinicio, pero el programa arranca.
        """
        if self.clave_sesion.strip():
            return self.clave_sesion.strip()
        import secrets
        ruta = self.raiz / "data" / "clave_sesion.txt"
        try:
            if ruta.exists():
                guardada = ruta.read_text(encoding="utf-8").strip()
                if guardada:
                    return guardada
            nueva = secrets.token_urlsafe(48)
            ruta.parent.mkdir(parents=True, exist_ok=True)
            ruta.write_text(nueva, encoding="utf-8")
            return nueva
        except OSError:
            return secrets.token_urlsafe(48)

    @property
    def ia_configurada(self) -> bool:
        """Permite arrancar la API sin clave: el 80% determinista no la necesita."""
        return bool(self.claves_gemini)

    @property
    def max_bytes_pdf(self) -> int:
        return self.max_mb_pdf * 1024 * 1024

    @property
    def max_bytes_imagen(self) -> int:
        return self.max_mb_imagen * 1024 * 1024

    @property
    def registro_plantillas(self) -> Dict[str, Dict[str, str]]:
        """Formatos de salida disponibles, leidos de config/plantillas.json.

        Agregar un banco nuevo es agregar una entrada ahi y su mapping. Si el
        archivo no existe se cae al formato Santander, que es el que trae el
        proyecto de fabrica.
        """
        ruta = self.raiz / "config" / "plantillas.json"
        if not ruta.exists():
            return {
                "santander": {
                    "nombre": "Banco Santander",
                    "plantilla": str(self.ruta_plantilla),
                    "mapping": str(self.ruta_mapping),
                }
            }
        with ruta.open(encoding="utf-8") as fh:
            return json.load(fh).get("plantillas", {})

    def _registro_crudo(self) -> Dict[str, object]:
        ruta = self.raiz / "config" / "plantillas.json"
        if not ruta.exists():
            return {}
        with ruta.open(encoding="utf-8") as fh:
            return json.load(fh)

    @property
    def plantilla_predeterminada(self) -> str:
        """Formato del informe completo que va al banco."""
        return str(self._registro_crudo().get("predeterminada_informe", "santander"))

    @property
    def plantilla_catastro(self) -> str:
        """Formato propio de la casa para volcar catastros.

        Es distinto del anterior a proposito: un catastro y un informe
        agroeconomico son documentos diferentes y no salen en la misma plantilla.
        """
        return str(self._registro_crudo().get("predeterminada_catastro", "fitogreen"))

    def rutas_de_plantilla(self, alias: str) -> Tuple[Path, Path]:
        """(archivo .xlsx, archivo de mapping) del formato pedido."""
        registro = self.registro_plantillas
        if alias not in registro:
            raise KeyError(
                "formato desconocido: {!r}. Disponibles: {}".format(
                    alias, ", ".join(sorted(registro))
                )
            )
        entrada = registro[alias]
        return (
            resolver_ruta(self.raiz / entrada["plantilla"]),
            resolver_ruta(self.raiz / entrada["mapping"]),
        )

    def motor_de_plantilla(self, alias: str) -> str:
        """'santander2026' para el formato final de Santander (su propio
        motor, services/excel_engine_santander.py); 'bsan' para los que llena
        MotorExcel."""
        return str(self.registro_plantillas.get(alias, {}).get("motor", "bsan"))

    def preparar_directorios(self) -> None:
        """Crea las carpetas de escritura. Idempotente."""
        for ruta in (self.ruta_salida, self.ruta_salida_pruebas, self.ruta_subidas):
            ruta.mkdir(parents=True, exist_ok=True)


@lru_cache(maxsize=1)
def obtener_ajustes() -> Ajustes:
    """Instancia unica; cacheada para no releer .env en cada peticion."""
    return Ajustes()


#: Atajo de importacion. Para tests que necesiten otro entorno, usar Ajustes().
ajustes = obtener_ajustes()
