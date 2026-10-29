"""Eventos con imputados: contraventores, detenidos, integrados y colaboración.

Portado de ``productividad_imputados.pyw`` (Tkinter). Genera
``DATOS/productividad_imputados.xlsx`` con una fila por evento detectado.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from .comunes import LogFn

# --- Encabezados del archivo maestro (CMU) ---
CONT_COL = "CONTRAVENTORES"
DET_COL = "DETENIDOS"
TIPIFICACION_COL = "TIPIFICACION"
CIERRE_COL = "TIPO DE CIERRE"
BARRIO_COL = "BARRIO"
ORIGEN_COL = "ORIGEN"
GAP_COL = "GAP"
SAE_COL = "SAE"
ACL_COL = "ACLARACIONES"
JEFE_SERVICIO_COL = "JEFE DE SERVICIO"

#: Posición 0-based de la columna OFICIAL en el archivo maestro.
OFICIAL_INDEX = 5

#: Valores de ``TIPO DE CIERRE`` que habilitan las reglas de imputados.
#: El script original buscaba ``"FINALIZA CON IMPUTADO/S PROPIOS"`` (sin la
#: barra antes de la S final), un texto que no existe en los archivos reales,
#: por lo que la rama de contraventores nunca generaba filas. Se aceptan las
#: dos variantes.
CIERRES_IMPUTADOS = (
    "FINALIZA CON IMPUTADO/S PROPIO/S",
    "FINALIZA CON IMPUTADO/S PROPIOS",
)
CIERRE_COLABORACION = (
    "COLABORACION CON INTERVENCION PROPIA",
    "COLABORACION CON IMPUTADO/S POSITIVA",
    "COLABORACION CON IMPUTADO/S POSITIVO",
    "COLABORACION CON IMPUTADO POSITIVA",
    "COLABORACION CON IMPUTADO POSITIVO",
)

DATOS_DIR = Path(__file__).resolve().parent.parent / "DATOS"
ARCHIVO_SALIDA = "productividad_imputados.xlsx"

#: Orden y nombres exactos de las columnas del archivo de salida.
COLUMNAS_SALIDA = [
    "OFICIAL",
    "JEFE DE SERVICIO",
    "TIPO DE EVENTO",
    "BARRIO",
    "ORIGEN",
    "TURNO",
    "P.EVENTO",
    "% ZONA",
    "P. TOTAL",
    "GAP",
    "SAE",
]


@dataclass
class ResultadoImputados:
    """Resumen del procesamiento, para el registro y el diálogo final."""

    ruta_salida: Path
    total: int = 0
    por_tipo: dict[str, int] = field(default_factory=dict)

    def resumen(self) -> str:
        lineas = [
            "✅ Proceso completado exitosamente",
            "",
            f"📁 Archivo generado:\n{self.ruta_salida}",
            "",
            "📊 Eventos por tipo:",
        ]
        if self.por_tipo:
            lineas.extend(
                f"  • {tipo}: {cantidad}" for tipo, cantidad in self.por_tipo.items()
            )
        else:
            lineas.append("  • (sin eventos)")
        lineas.append(f"  • Total: {self.total}")
        return "\n".join(lineas)


def _texto(valor) -> str:
    """Texto normalizado para comparar: sin NaN, sin espacios externos, MAYÚSCULAS."""
    if pd.isna(valor):
        return ""
    return str(valor).strip().upper()


def _texto_sin_acentos(valor) -> str:
    """Como :func:`_texto` pero además sin diacríticos (para ACLARACIONES)."""
    return "".join(
        c
        for c in unicodedata.normalize("NFKD", _texto(valor))
        if not unicodedata.combining(c)
    )


def _cantidad(valor) -> float:
    """Cantidad numérica tolerante a celdas de texto o vacías."""
    numero = pd.to_numeric(valor, errors="coerce")
    return 0.0 if pd.isna(numero) else float(numero)


def _oficial(fila: pd.Series):
    return fila.iloc[OFICIAL_INDEX] if len(fila) > OFICIAL_INDEX else None


def _fila(tipo_evento: str, fila: pd.Series) -> dict[str, object]:
    """Fila de salida con todas las columnas, incluidas las vacías."""
    return {
        "OFICIAL": _oficial(fila),
        "JEFE DE SERVICIO": fila.get(JEFE_SERVICIO_COL),
        "TIPO DE EVENTO": tipo_evento,
        "BARRIO": fila.get(BARRIO_COL),
        "ORIGEN": fila.get(ORIGEN_COL),
        "TURNO": "",
        "P.EVENTO": "",
        "% ZONA": "",
        "P. TOTAL": "",
        "GAP": fila.get(GAP_COL),
        "SAE": fila.get(SAE_COL),
    }


def _dataframe(filas: list[dict[str, object]]) -> pd.DataFrame:
    return pd.DataFrame(filas, columns=COLUMNAS_SALIDA)


def extraer_eventos(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica las cuatro reglas y devuelve la tabla de eventos.

    El orden de las filas es el del script original: primero todos los
    contraventores, después los detenidos, los integrados/destacados y por
    último las colaboraciones.
    """
    filas_contraventor: list[dict[str, object]] = []
    filas_detenido: list[dict[str, object]] = []
    filas_integrados: list[dict[str, object]] = []
    filas_colaboracion: list[dict[str, object]] = []

    for _, fila in df.iterrows():
        if (
            _cantidad(fila.get(CONT_COL, 0)) >= 1
            and _texto(fila.get(CIERRE_COL, "")) in CIERRES_IMPUTADOS
        ):
            filas_contraventor.append(_fila("CONTRAVENCION", fila))

    for _, fila in df.iterrows():
        if (
            _cantidad(fila.get(DET_COL, 0)) >= 1
            and _texto(fila.get(CIERRE_COL, "")) in CIERRES_IMPUTADOS
        ):
            tipificacion = _texto(fila.get(TIPIFICACION_COL, ""))
            tipo = "ROBO/HURTO" if tipificacion == "ROBO/HURTO" else "DELITO"
            filas_detenido.append(_fila(tipo, fila))

    for _, fila in df.iterrows():
        aclaraciones = _texto_sin_acentos(fila.get(ACL_COL, "")).lower()
        if "integ" in aclaraciones or "dest" in aclaraciones:
            filas_integrados.append(_fila("INTERVENCION DESTACADA", fila))

    for _, fila in df.iterrows():
        if _texto(fila.get(CIERRE_COL, "")) in CIERRE_COLABORACION:
            filas_colaboracion.append(
                _fila("COLABORACION CON IMPUTADO/S POSITIVA", fila)
            )

    return pd.concat(
        [
            _dataframe(filas_contraventor),
            _dataframe(filas_detenido),
            _dataframe(filas_integrados),
            _dataframe(filas_colaboracion),
        ],
        ignore_index=True,
    )


def procesar(ruta_archivo: Path | str, log: LogFn) -> ResultadoImputados:
    """Lee el archivo maestro, extrae los eventos y guarda la salida en DATOS/."""
    ruta_archivo = Path(ruta_archivo)
    log("═══════════════════════════════════════════════════════════════════════")
    log("🚀 INICIANDO PROCESO DE EVENTOS CON IMPUTADOS")
    log("═══════════════════════════════════════════════════════════════════════")
    log(f"📂 Archivo: {ruta_archivo.name}")

    log("→ Cargando archivo...")
    df = pd.read_excel(ruta_archivo, engine="openpyxl")
    log(f"✓ Archivo cargado — {df.shape[0]} filas × {df.shape[1]} columnas")

    if BARRIO_COL in df.columns:
        try:
            df[BARRIO_COL] = df[BARRIO_COL].str.lower()
            log("✓ Columna BARRIO convertida a minúsculas")
        except AttributeError:
            log("⚠ La columna BARRIO no es de texto; se deja sin convertir")

    log("→ Aplicando reglas de extracción...")
    resultado = extraer_eventos(df)

    DATOS_DIR.mkdir(exist_ok=True)
    salida = DATOS_DIR / ARCHIVO_SALIDA
    resultado.to_excel(salida, index=False, engine="openpyxl")

    por_tipo = (
        {
            str(tipo): int(cantidad)
            for tipo, cantidad in resultado["TIPO DE EVENTO"].value_counts().items()
        }
        if not resultado.empty
        else {}
    )

    log("───────────────────────────────────────────────────────────────────────")
    for tipo, cantidad in por_tipo.items():
        log(f"  • {tipo}: {cantidad}")
    log(f"✓ Total de eventos: {len(resultado)}")
    log(f"✓ Archivo guardado en: {salida}")
    log("═══════════════════════════════════════════════════════════════════════")

    return ResultadoImputados(
        ruta_salida=salida, total=len(resultado), por_tipo=por_tipo
    )
