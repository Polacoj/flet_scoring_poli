"""Unificación de eventos del CMU.

Portado de ``productividad_general.pyw`` (Tkinter). Genera
``eventos_unificados.xlsx`` (hoja ``eventos``) en la misma carpeta del archivo
de origen, combinando:

* los eventos principales extraídos por posición de columna,
* los eventos de colaboración con intervención propia,
* los eventos con contraventores/detenidos,
* los eventos integrados/destacados.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from complementos.contraventores_detenidos import ContraventoresDetenidosProcessor
from complementos.integrados import IntegradosProcessor
from complementos.intervencion_directa import ColaboracionPropiaProcessor

from .comunes import LogFn

#: Columnas a extraer, numeradas desde 1 (igual que el script original). Las
#: cadenas vacías generan las columnas ``VACIO_8`` … ``VACIO_14``.
COLUMNAS_A_EXTRAER: list[int | str] = [
    1,
    2,
    4,
    6,
    16,
    20,
    32,
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    29,
    30,
    13,
    14,
]

#: GAP y SAE, también en numeración desde 1 (posiciones 12 y 13 en base 0).
COLS_GAP_SAE_1BASED: tuple[int, ...] = (13, 14)

#: Descripción de cada columna extraída, sólo para el registro.
CORRESPONDENCIA: dict[int, str] = {
    1: "FECHA",
    2: "HORA",
    4: "JEFE DE SERVICIO",
    6: "OFICIAL",
    16: "TIPIFICACION",
    20: "BARRIO",
    32: "ORIGEN",
    29: "CONTRAVENTORES",
    30: "DETENIDOS",
    13: "GAP",
    14: "SAE",
}

ARCHIVO_SALIDA = "eventos_unificados.xlsx"
HOJA_SALIDA = "eventos"


@dataclass
class ResultadoGeneral:
    """Resumen del procesamiento, para el registro y el diálogo final."""

    ruta_salida: Path
    filas_eventos: int = 0
    filas_colaboracion: int | None = None
    filas_contraventores: int | None = None
    filas_integrados: int | None = None
    total: int = 0
    avisos: list[str] = field(default_factory=list)

    def resumen(self) -> str:
        lineas = [
            "✅ Proceso completado exitosamente",
            "",
            f"📁 Archivo generado:\n{self.ruta_salida}",
            "",
            "📊 Resumen:",
            f"  • Eventos principales: {self.filas_eventos}",
        ]
        lineas.append(
            f"  • Colaboración con intervención propia: {self.filas_colaboracion}"
            if self.filas_colaboracion is not None
            else "  • Colaboración con intervención propia: No procesado"
        )
        lineas.append(
            f"  • Contraventores/Detenidos: {self.filas_contraventores}"
            if self.filas_contraventores is not None
            else "  • Contraventores/Detenidos: No procesado"
        )
        lineas.append(
            f"  • Integrados/Destacados: {self.filas_integrados}"
            if self.filas_integrados is not None
            else "  • Integrados/Destacados: No procesado"
        )
        lineas.append(f"  • Total general: {self.total}")
        if self.avisos:
            lineas.append("")
            lineas.append("⚠ Avisos:")
            lineas.extend(f"  • {aviso}" for aviso in self.avisos)
        return "\n".join(lineas)


def _truncar_filas_vacias(df: pd.DataFrame, log: LogFn) -> pd.DataFrame:
    """Corta el DataFrame en la primera fila completamente vacía."""
    log("→ Eliminando filas vacías al final del archivo...")
    df_temp = df.replace(r"^\s*$", pd.NA, regex=True)
    filas_vacias = df_temp.isna().all(axis=1)
    if filas_vacias.any():
        pos = int(filas_vacias.idxmax())
        log(f"✓ Filas truncadas: {pos} filas procesadas")
        return df.iloc[:pos].reset_index(drop=True)
    log("✓ No se encontraron filas vacías para truncar")
    return df.reset_index(drop=True)


def _limpiar_gap_sae(df: pd.DataFrame, log: LogFn) -> pd.DataFrame:
    """Corta ``GAP`` y ``SAE`` en el primer ``/`` y los convierte a numérico."""
    log("→ Limpiando columnas GAP y SAE...")
    for uno_based in COLS_GAP_SAE_1BASED:
        pos = uno_based - 1
        if pos >= len(df.columns):
            continue
        nombre = df.columns[pos]
        serie = (
            df.iloc[:, pos]
            .fillna("")
            .astype(str)
            .str.split("/", n=1)
            .str[0]
            .str.strip()
        )
        df[nombre] = pd.to_numeric(serie, errors="coerce")
        log(f"  ✓ Columna '{nombre}' (posición {pos} en base 0) limpiada")
    return df


def _extraer_columnas(df: pd.DataFrame, log: LogFn) -> pd.DataFrame:
    """Extrae las columnas indicadas por posición (1-based)."""
    log("→ Extrayendo columnas específicas...")
    log(f"  Columnas seleccionadas (numeración desde 1): {COLUMNAS_A_EXTRAER}")
    log("  Correspondencia:")
    for uno_based, nombre in CORRESPONDENCIA.items():
        log(f"    Col {uno_based} → {nombre}")
    log("  (las columnas no listadas no se extraen; los huecos quedan como VACIO_*)")

    seleccionadas: list[pd.Series] = []
    nombres: list[str] = []
    for valor in COLUMNAS_A_EXTRAER:
        if isinstance(valor, int):
            pos = valor - 1
            if pos < 0 or pos >= len(df.columns):
                raise RuntimeError(
                    "El archivo no tiene suficientes columnas.\n"
                    f"Se esperaban al menos {pos + 1}"
                )
            seleccionadas.append(df.iloc[:, pos].copy())
            nombres.append(df.columns[pos])
        else:
            seleccionadas.append(
                pd.Series([""] * len(df), index=df.index, dtype=object)
            )
            nombres.append(f"VACIO_{len(nombres) + 1}")

    df_eventos = pd.concat(seleccionadas, axis=1)
    df_eventos.columns = nombres

    if "TIPIFICACION" in df_eventos.columns:
        df_eventos = df_eventos.rename(columns={"TIPIFICACION": "TIPO DE EVENTO"})
        log("✓ Columna 'TIPIFICACION' renombrada a 'TIPO DE EVENTO'")
    log(f"✓ Extracción completada — filas extraídas: {len(df_eventos)}")
    return df_eventos


def _procesar_complemento(
    etiqueta: str,
    log: LogFn,
    avisos: list[str],
    fabrica,
    df: pd.DataFrame,
) -> pd.DataFrame | None:
    """Ejecuta un procesador de ``complementos`` sin abortar el proceso si falla."""
    try:
        log(f"→ Procesando: {etiqueta}")
        resultado = fabrica().process_df(df)
        log(f"✓ Procesado exitosamente — eventos encontrados: {len(resultado)}")
        return resultado
    except Exception as exc:  # noqa: BLE001 - el original continuaba ante errores
        aviso = f"{etiqueta}: {exc}"
        avisos.append(aviso)
        log(f"✗ Error: {exc}")
        return None


def procesar(ruta_archivo: Path | str, log: LogFn) -> ResultadoGeneral:
    """Procesa el archivo maestro del CMU y genera ``eventos_unificados.xlsx``."""
    ruta_archivo = Path(ruta_archivo)
    ruta_dir = ruta_archivo.parent

    log("═══════════════════════════════════════════════════════════════════════")
    log("🚀 INICIANDO PROCESO PRINCIPAL")
    log("═══════════════════════════════════════════════════════════════════════")
    log(f"📂 Archivo: {ruta_archivo.name}")
    log(f"📁 Carpeta: {ruta_dir}")

    log("───────────────────────────────────────────────────────────────────────")
    log("📊 PASO 1: CARGANDO Y PREPARANDO DATOS")
    log("───────────────────────────────────────────────────────────────────────")
    df = pd.read_excel(ruta_archivo, engine="openpyxl")
    log("✓ Archivo cargado exitosamente")
    log(f"  Dimensiones originales: {df.shape[0]} filas × {df.shape[1]} columnas")

    df = _truncar_filas_vacias(df, log)
    df = _limpiar_gap_sae(df, log)

    log("───────────────────────────────────────────────────────────────────────")
    log("📊 PASO 2: EXTRACCIÓN DE COLUMNAS")
    log("───────────────────────────────────────────────────────────────────────")
    df_eventos = _extraer_columnas(df, log)
    log(f"  Carpeta de salida: {ruta_dir}")

    log("───────────────────────────────────────────────────────────────────────")
    log("📊 PASO 3: PROCESAMIENTO ESPECIALIZADO")
    log("───────────────────────────────────────────────────────────────────────")
    avisos: list[str] = []

    df_colaboracion = _procesar_complemento(
        "Eventos Colaboración con Intervención Propia",
        log,
        avisos,
        ColaboracionPropiaProcessor,
        df,
    )
    df_contra_det = _procesar_complemento(
        "Eventos con CONTRAVENTORES y DETENIDOS",
        log,
        avisos,
        ContraventoresDetenidosProcessor,
        df,
    )
    df_integrados = _procesar_complemento(
        "Eventos INTEGRADOS y DESTACADOS",
        log,
        avisos,
        IntegradosProcessor,
        df,
    )

    log("───────────────────────────────────────────────────────────────────────")
    log("📊 PASO 4: CREACIÓN DE ARCHIVO UNIFICADO")
    log("───────────────────────────────────────────────────────────────────────")
    ruta_unificado = ruta_dir / ARCHIVO_SALIDA
    log("→ Combinando todos los datos procesados...")

    partes = [df_eventos]
    if df_colaboracion is not None:
        partes.append(df_colaboracion)
        log(f"  ✓ Añadidos {len(df_colaboracion)} eventos de colaboración")
    if df_contra_det is not None:
        partes.append(df_contra_det)
        log(f"  ✓ Añadidos {len(df_contra_det)} eventos de contraventores/detenidos")
    if df_integrados is not None:
        partes.append(df_integrados)
        log(f"  ✓ Añadidos {len(df_integrados)} eventos integrados/destacados")

    df_combined = pd.concat(partes, ignore_index=True, sort=False)
    log(f"✓ Datos combinados exitosamente — filas totales: {len(df_combined)}")

    with pd.ExcelWriter(ruta_unificado, engine="openpyxl") as writer:
        df_combined.to_excel(writer, sheet_name=HOJA_SALIDA, index=False)

    log("═══════════════════════════════════════════════════════════════════════")
    log("✅ PROCESO COMPLETADO EXITOSAMENTE")
    log("═══════════════════════════════════════════════════════════════════════")
    log(f"📁 ARCHIVO GENERADO: {ruta_unificado.name}")
    log(f"📂 UBICACIÓN: {ruta_dir}")
    log(f"📊 TOTAL REGISTROS: {len(df_combined)}")
    log("═══════════════════════════════════════════════════════════════════════")

    return ResultadoGeneral(
        ruta_salida=ruta_unificado,
        filas_eventos=len(df_eventos),
        filas_colaboracion=len(df_colaboracion) if df_colaboracion is not None else None,
        filas_contraventores=len(df_contra_det) if df_contra_det is not None else None,
        filas_integrados=len(df_integrados) if df_integrados is not None else None,
        total=len(df_combined),
        avisos=avisos,
    )
