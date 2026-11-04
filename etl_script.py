"""
ETL de ventas: sucursal A (CSV) + sucursal B (Excel) -> PostgreSQL.

Requisitos (pandas >= 2.0):
    pip install pandas sqlalchemy psycopg2-binary openpyxl python-dotenv

"""
from __future__ import annotations

import logging
import os
import re
import sys
from pathlib import Path

import pandas as pd
from sqlalchemy import PrimaryKeyConstraint, create_engine
from sqlalchemy.engine import URL, Engine
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.types import Date, Integer, Numeric, String

try:  # opcional: carga el archivo .env si python-dotenv está instalado
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass

# ----------------------------------------------------------------------------
# CONFIGURACIÓN
# ----------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("etl_ventas")

# Rutas relativas al script (no al directorio desde el que se ejecute VS Code)
BASE_DIR = Path(__file__).resolve().parent
RUTA_CSV = BASE_DIR / "muestra_sucursal_A.csv"
RUTA_EXCEL = BASE_DIR / "muestra_sucursal_B.xlsx"
TABLA_DESTINO = "clientes_procesados"

COLUMNAS_ESTANDAR = ["id_cliente", "nombre", "email", "fecha", "monto", "estado"]

# NOTA DE NEGOCIO: verifica que estos mapeos sean los correctos. En el CSV,
# 'precio_producto' -> monto y 'stock_disponible' -> estado parecen columnas de
# inventario, no de ventas; confirma que esa equivalencia sea intencional.
MAPEO_CSV = {
    "id_cliente": "id_cliente",
    "nombre_cliente": "nombre",
    "email": "email",
    "fecha_registro": "fecha",
    "precio_producto": "monto",
    "stock_disponible": "estado",
}
MAPEO_EXCEL = {
    "ID_CLIENTE": "id_cliente",
    "Nombre": "nombre",
    "Correo Electronico": "email",
    "Fecha": "fecha",
    "Monto Venta": "monto",
    "Estado": "estado",
}

# Textos que se tratan como "dato faltante" (comparados en minúsculas)
VALORES_NULOS = {"", "nan", "none", "null", "n/a", "nat", "-"}

DTYPES_SQL = {
    "id_cliente": Integer(),
    "nombre": String(255),
    "email": String(255),
    "fecha": Date(),
    "monto": Numeric(14, 2),
    "estado": String(100),
}


# ----------------------------------------------------------------------------
# EXTRACCIÓN Y HOMOLOGACIÓN
# ----------------------------------------------------------------------------
def _verificar_archivo(ruta: Path) -> Path:
    ruta = Path(ruta)
    if not ruta.is_file():
        raise FileNotFoundError(f"No se encontró el archivo: {ruta}")
    return ruta


def _homologar(df: pd.DataFrame, mapeo: dict[str, str], origen: str) -> pd.DataFrame:
    """Renombra a columnas estándar validando que existan todas las de origen."""
    df = df.rename(columns=lambda c: str(c).strip())  # encabezados con espacios
    faltantes = sorted(set(mapeo) - set(df.columns))
    if faltantes:
        raise KeyError(
            f"{origen}: faltan columnas {faltantes}. Columnas encontradas: {list(df.columns)}"
        )
    return df.rename(columns=mapeo)[COLUMNAS_ESTANDAR].copy()


def leer_csv_sucursal_a(ruta: Path) -> pd.DataFrame:
    """Lee el CSV de la sucursal A. Prueba UTF-8 (con/sin BOM) y luego Latin-1."""
    ruta = _verificar_archivo(ruta)
    ultimo_error: Exception | None = None
    for codificacion in ("utf-8-sig", "latin-1"):
        try:
            # dtype=str: nada se convierte "solo" (ids a float, precios, fechas)
            # sep=None: detecta automáticamente ',' o ';'
            df = pd.read_csv(
                ruta, dtype=str, sep=None, engine="python", encoding=codificacion
            )
            log.info("CSV leído (%s): %d filas", codificacion, len(df))
            return _homologar(df, MAPEO_CSV, "Sucursal A (CSV)")
        except UnicodeDecodeError as exc:
            ultimo_error = exc
    raise ValueError(f"No se pudo decodificar {ruta.name}") from ultimo_error


def leer_excel_sucursal_b(ruta: Path) -> pd.DataFrame:
    """Lee el Excel de la sucursal B (requiere openpyxl)."""
    ruta = _verificar_archivo(ruta)
    df = pd.read_excel(ruta, dtype=str, engine="openpyxl")
    log.info("Excel leído: %d filas", len(df))
    return _homologar(df, MAPEO_EXCEL, "Sucursal B (Excel)")


# ----------------------------------------------------------------------------
# TRANSFORMACIÓN
# ----------------------------------------------------------------------------
def limpiar_cadenas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in ("nombre", "email", "estado"):
        s = df[col].astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
        df[col] = s.mask(s.str.lower().isin(VALORES_NULOS))
        
    # Estandarizar nombre (ej. "JUAN pablo diaz" -> "Juan Pablo Diaz")
    df["nombre"] = df["nombre"].str.title()
    
    # Un email nunca lleva espacios y va en minúsculas
    df["email"] = df["email"].str.replace(" ", "", regex=False).str.lower()
    return df


def limpiar_ids(df: pd.DataFrame) -> pd.DataFrame:
    """Convierte id_cliente a entero anulable (Int64); lo no numérico queda NULL."""
    df = df.copy()
    s = df["id_cliente"].astype("string").str.strip().str.replace(r"\.0+$", "", regex=True)
    num = pd.to_numeric(s, errors="coerce")
    num = num.where(num % 1 == 0)  # descarta no enteros (p. ej. 12.5)
    df["id_cliente"] = num.astype("Int64")
    return df


_PATRON_MILES = re.compile(r"^\d{1,3}(\.\d{3})+$")


def _parsear_monto(valor) -> float | None:
    """
    Convierte '$ 3.450.000 COP' -> 3450000.0, respetando decimales.
    Convención colombiana: '.' separa miles y ',' decimales.
      '$ 3.450.000 COP' -> 3450000.0   '1.234,56' -> 1234.56
      '1,234.56'        -> 1234.56     '3450000.5' (Excel) -> 3450000.5
    """
    if valor is None or pd.isna(valor):
        return None
    texto = str(valor).strip()
    if texto.lower() in VALORES_NULOS:
        return None

    negativo = bool(re.search(r"(^|[\s$])-\s*\$?\s*\d", texto)) or texto.startswith("(")
    texto = re.sub(r"[^\d.,]", "", texto).strip(".,")  # quita $, COP, espacios
    if not texto:
        return None

    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):  # 1.234,56
            texto = texto.replace(".", "").replace(",", ".")
        else:  # 1,234.56
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", "") if texto.count(",") > 1 else texto.replace(",", ".")
    elif "." in texto and _PATRON_MILES.match(texto):  # 3.450.000 o 3.450
        texto = texto.replace(".", "")

    try:
        numero = float(texto)
    except ValueError:
        return None
    return -numero if negativo else numero


def limpiar_precios(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["monto"] = df["monto"].map(_parsear_monto).astype("Float64")
    return df


def estandarizar_fechas(df: pd.DataFrame) -> pd.DataFrame:
    """Fechas a YYYY-MM-DD (objeto date). Nulos, vacíos y 'N/A' -> NULL."""
    df = df.copy()
    s = df["fecha"].astype("string").str.strip()
    s = s.mask(s.str.lower().isin(VALORES_NULOS))

    # 1) Formato ISO (YYYY-MM-DD, con o sin hora): se interpreta tal cual.
    #    Hay que separarlo porque dayfirst=True invertiría día y mes en ISO.
    parsed = pd.to_datetime(s, errors="coerce", format="ISO8601")

    # 2) Lo demás (05/03/2024, 5-3-2024...): día primero, como en Colombia
    pendientes = parsed.isna() & s.notna()
    if pendientes.any():
        resto = pd.to_datetime(
            s[pendientes], errors="coerce", dayfirst=True, format="mixed"
        )
        parsed = parsed.fillna(resto)

    invalidas = int((s.notna() & parsed.isna()).sum())
    if invalidas:
        log.warning("%d fechas no interpretables quedaron como NULL", invalidas)

    df["fecha"] = parsed.dt.date
    return df


def _sin_duplicados_ignorando_nulos(df: pd.DataFrame, columna: str) -> pd.DataFrame:
    """drop_duplicates por columna, pero los NULL no cuentan como duplicados."""
    repetido = df.duplicated(subset=[columna], keep="first") & df[columna].notna()
    return df[~repetido]


def eliminar_duplicados(df: pd.DataFrame) -> pd.DataFrame:
    """
    Deduplica por id_cliente y por email conservando la fila más completa.
    Desempate: la fecha más reciente. Los NULL nunca se tratan como iguales.
    """
    df = df.copy()
    antes = len(df)

    df["_completitud"] = df[COLUMNAS_ESTANDAR].notna().sum(axis=1)
    df["_fecha_orden"] = pd.to_datetime(df["fecha"], errors="coerce")
    df = df.sort_values(
        ["_completitud", "_fecha_orden"],
        ascending=[False, False],
        kind="mergesort",  # estable: resultado reproducible
        na_position="last",
    )

    df = _sin_duplicados_ignorando_nulos(df, "id_cliente")
    df = _sin_duplicados_ignorando_nulos(df, "email")
    df = df.drop(columns=["_completitud", "_fecha_orden"]).sort_index()

    log.info("Duplicados eliminados: %d", antes - len(df))
    return df


def transformar(df_a: pd.DataFrame, df_b: pd.DataFrame) -> pd.DataFrame:
    """Orquesta toda la limpieza sobre los datos unificados."""
    df = pd.concat([df_a, df_b], ignore_index=True)
    df = limpiar_cadenas(df)
    df = limpiar_ids(df)
    df = limpiar_precios(df)
    df = estandarizar_fechas(df)

    # Sin id ni email no hay forma de identificar al cliente ni de deduplicarlo
    sin_identificador = df["id_cliente"].isna() & df["email"].isna()
    if sin_identificador.any():
        log.warning("%d filas sin id_cliente ni email fueron descartadas", sin_identificador.sum())
        df = df[~sin_identificador]

    return eliminar_duplicados(df).reset_index(drop=True)


# ----------------------------------------------------------------------------
# CARGA
# ----------------------------------------------------------------------------
def conectar_postgresql() -> Engine:
    """Engine de SQLAlchemy. URL.create escapa símbolos raros en la contraseña."""
    password = os.getenv("DB_PASSWORD")
    if not password:
        raise EnvironmentError("Falta la variable de entorno DB_PASSWORD (revisa tu .env).")

    url = URL.create(
        drivername="postgresql+psycopg2",
        username=os.getenv("DB_USER", "postgres"),
        password=password,
        host=os.getenv("DB_HOST", "localhost"),
        port=int(os.getenv("DB_PORT", "5432")),
        database=os.getenv("DB_NAME", "etl_ventas_db"),
    )
    return create_engine(url, pool_pre_ping=True)


def guardar_en_postgresql(df: pd.DataFrame, engine: Engine, tabla: str = TABLA_DESTINO) -> int:
    """Inserta en una sola transacción: o entra todo o no entra nada."""
    if df.empty:
        log.warning("No hay registros para insertar.")
        return 0
    try:
        with engine.begin() as conexion:  # commit/rollback automático
            df.to_sql(
                name=tabla,
                con=conexion,
                if_exists="replace",
                index=False,
                dtype=DTYPES_SQL,
                chunksize=1000,
                method="multi",
            )
    except SQLAlchemyError:
        log.exception("Falló la inserción en PostgreSQL (se hizo rollback)")
        raise
    log.info("Insertados %d registros en '%s'", len(df), tabla)
    return len(df)


# ----------------------------------------------------------------------------
# PUNTO DE ENTRADA
# ----------------------------------------------------------------------------
def main() -> int:
    try:
        log.info("Iniciando ETL...")
        df_a = leer_csv_sucursal_a(RUTA_CSV)
        df_b = leer_excel_sucursal_b(RUTA_EXCEL)

        df = transformar(df_a, df_b)
        log.info("Registros finales: %d", len(df))
        print(df.head(), "\n")
        print(df.dtypes)

        engine = conectar_postgresql()
        guardar_en_postgresql(df, engine)
        log.info("ETL completado.")
        return 0
    except (FileNotFoundError, KeyError, ValueError, EnvironmentError) as exc:
        log.error("Error de datos/configuración: %s", exc)
    except SQLAlchemyError:
        pass  # ya registrado con traceback en guardar_en_postgresql
    except Exception:
        log.exception("Error inesperado")
    return 1


if __name__ == "__main__":
    sys.exit(main())