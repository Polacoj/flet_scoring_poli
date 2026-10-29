"""Utilidades compartidas por la lógica de los tres módulos.

Portadas de ``diccionario_OFICIALES.pyw`` (PyQt6) y unificadas para que los
tres módulos usen exactamente las mismas reglas de normalización, lectura de
hojas y nombres de archivos de salida.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import pandas as pd

#: Callback de registro. Se ejecuta en el hilo que lo invoca; las vistas se
#: encargan de que llegar a la interfaz sea seguro (ver ui/comunes.py).
LogFn = Callable[[str], None]


def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Devuelve una copia con los encabezados en MAYÚSCULAS y sin espacios."""
    df = df.copy()
    df.columns = [str(col).strip().upper() for col in df.columns]
    return df


def normalizar_nombre(valor) -> str:
    """Normaliza un nombre: sin NaN, en MAYÚSCULAS y con espacios simples."""
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto == "" or texto.lower() == "nan":
        return ""
    return " ".join(texto.upper().split())


def serie_id(serie: pd.Series) -> pd.Series:
    """Normaliza una columna de IDs a texto, sin NaN ni espacios sobrantes."""
    return serie.fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})


def validar_columnas(df: pd.DataFrame, requeridas: list[str]) -> tuple[bool, str]:
    """Comprueba que el DataFrame tenga todas las columnas requeridas."""
    presentes = set(df.columns)
    faltantes = [col for col in requeridas if col not in presentes]
    if faltantes:
        return False, (
            f"Columnas faltantes: {faltantes}. "
            f"Columnas encontradas: {sorted(presentes)}"
        )
    return True, ""


def listar_hojas(ruta: Path) -> list[str]:
    """Devuelve los nombres de las hojas del Excel sin cargar los datos."""
    with pd.ExcelFile(ruta) as libro:
        return list(libro.sheet_names)


def leer_hoja(ruta: Path, hoja: str) -> pd.DataFrame:
    """Lee una hoja como texto (``dtype=str``), igual que el script original."""
    return pd.read_excel(ruta, sheet_name=hoja, dtype=str)


def leer_todas_las_hojas(ruta: Path) -> dict[str, pd.DataFrame]:
    """Lee todas las hojas como texto, indexadas por nombre."""
    return pd.read_excel(ruta, sheet_name=None, dtype=str)


def encabezados_por_hoja(ruta: Path) -> dict[str, list[str]]:
    """Encabezados normalizados de cada hoja, sin cargar los datos (``nrows=0``)."""
    hojas = pd.read_excel(ruta, sheet_name=None, nrows=0)
    return {
        nombre: [str(col).strip().upper() for col in df.columns]
        for nombre, df in hojas.items()
    }


def hojas_con_columnas(
    ruta: Path, requeridas: list[str]
) -> tuple[list[str], dict[str, list[str]]]:
    """Hojas que contienen todas las columnas requeridas.

    Devuelve ``(hojas_válidas, encabezados_por_hoja)`` para poder informar qué
    se encontró cuando ninguna hoja sirve.
    """
    encabezados = encabezados_por_hoja(ruta)
    buscadas = [col.strip().upper() for col in requeridas]
    validas = [
        nombre
        for nombre, columnas in encabezados.items()
        if all(col in columnas for col in buscadas)
    ]
    return validas, encabezados


def buscar_hoja(nombres: list[str], objetivo: str) -> str | None:
    """Busca una hoja ignorando mayúsculas/minúsculas y espacios externos."""
    buscado = objetivo.strip().upper()
    return next((n for n in nombres if n.strip().upper() == buscado), None)


def ruta_procesado(ruta: Path) -> Path:
    """``CMU.xlsx`` -> ``CMU_PROCESADO.xlsx`` (en la misma carpeta)."""
    return ruta.parent / f"{ruta.stem}_PROCESADO{ruta.suffix}"


def ruta_revertido(ruta_procesada: Path) -> Path:
    """``CMU_PROCESADO.xlsx`` -> ``CMU_REVERTIDO.xlsx`` (no pisa el procesado)."""
    base = ruta_procesada.stem.replace("_PROCESADO", "")
    return ruta_procesada.parent / f"{base}_REVERTIDO{ruta_procesada.suffix}"


def validar_ruta(texto: str, etiqueta: str) -> Path:
    """Valida que la ruta escrita/pegada exista y devuelve un ``Path``."""
    if not texto or not texto.strip():
        raise RuntimeError(f"Debés seleccionar el archivo: {etiqueta}")
    ruta = Path(texto.strip())
    if not ruta.exists():
        raise RuntimeError(f"No se encontró el archivo {etiqueta}: {ruta}")
    return ruta
