"""Pestaña «Diccionario»: reconciliación de nombres → LP y proceso inverso."""

from __future__ import annotations

from pathlib import Path

import flet as ft

from core.comunes import leer_hoja, validar_ruta
from core.diccionario_oficiales import (
    COLUMNAS_ANALISIS,
    COLUMNAS_DICCIONARIO,
    HOJA_PREFERIDA_DICCIONARIO,
    ProcesoIncompleto,
    construir_mapa_inverso,
    leer_validacion,
    reconciliar,
    revertir,
)

from .comunes import VistaBase, bloque_titulo, mostrar_dialogo, resolver_hoja
from .seleccion import CampoArchivo

COLUMNA_A_REVERTIR = "OFICIAL"


class VistaDiccionario(VistaBase):
    """Reconcilia los nombres de la columna OFICIAL con su LP del diccionario."""

    titulo = "Diccionario"
    icono = ft.Image(
        src="/Users/alexisjankowicz/Python/flet_scoring_poli/src/icons/diccionario_color.png",  # Ruta relativa dentro de la carpeta assets
        width=32,
        height=32,
        fit=ft.BoxFit.CONTAIN,
    )
    etiqueta_registro = "Registro — Diccionario de Oficiales"

    def __init__(self, page: ft.Page) -> None:
        super().__init__(page)
        self.ruta_procesada: Path | None = None
        self.ruta_diccionario: Path | None = None

        self.archivo_diccionario = CampoArchivo(
            page,
            "Diccionario Maestro OFICIALES:",
            "Seleccionar Diccionario Maestro OFICIALES",
        )
        self.archivo_analisis = CampoArchivo(
            page, "Archivo para analizar:", "Seleccionar archivo para analizar"
        )

        self.btn_reconciliar = ft.Button(
            "Validar y reconciliar",
            icon=ft.Icons.PLAY_ARROW,
            bgcolor="#2196F3",
            color=ft.Colors.WHITE,
            on_click=self.on_reconciliar,
        )
        self.btn_revertir = ft.Button(
            "Revertir IDs → Nombres (hoja validacion)",
            icon=ft.Icons.UNDO,
            bgcolor=ft.Colors.GREY_500,
            color=ft.Colors.WHITE,
            disabled=True,
            on_click=self.on_revertir,
        )

    # ------------------------------------------------------------------ UI --
    def build(self) -> ft.Control:
        ayuda = ft.Container(
            bgcolor="#F1F5F9",
            border_radius=8,
            alignment=ft.Alignment.CENTER_LEFT,
            padding=4,
            content=ft.Column(
                [
                    ft.Text("Cómo funciona", size=13, weight=ft.FontWeight.BOLD),
                    ft.Text(
                        "1. El Diccionario Maestro aporta la hoja Sheet1 "
                        "(APELLIDO Y NOMBRE + LP).",
                        size=12,
                    ),
                    ft.Text(
                        "2. El archivo a analizar debe tener la columna OFICIAL.",
                        size=12,
                    ),
                    ft.Text(
                        "3. «Validar y reconciliar» reemplaza los nombres por LP y guarda "
                        "<archivo>_PROCESADO.xlsx.",
                        size=12,
                    ),
                    ft.Text(
                        "   Si falta algún nombre se agrega al diccionario "
                        "(backup ANTIGUO_DICCIONARIO.xlsx) y el proceso se detiene para "
                        "completar los IDs.",
                        size=12,
                    ),
                    ft.Text(
                        "4. «Revertir» lee la hoja validacion y genera "
                        "<archivo>_REVERTIDO.xlsx sin tocar el _PROCESADO.",
                        size=12,
                    ),
                    ft.Text(
                        "Para elegir archivos: «Examinar», «Diálogo del sistema» "
                        "o escribí/pegá la ruta en el campo.",
                        size=12,
                        italic=True,
                    ),
                ],
                spacing=3,
            ),
        )
        return ft.Column(
            [
                bloque_titulo(
                    "Diccionario de Oficiales",
                    "Reconcilia los nombres de la columna OFICIAL con su LP del "
                    "Diccionario Maestro y permite el camino inverso.",
                ),
                self.archivo_diccionario.build(),
                self.archivo_analisis.build(),
                ft.Row([self.btn_reconciliar, self.progreso], spacing=12),
                self.btn_revertir,
                ayuda,
                ft.Divider(),
                self.registro.campo,
            ],
            expand=True,
            spacing=10,
        )

    # ------------------------------------------------------------- estado --
    def _ocupado(self, ocupado: bool) -> None:
        self.progreso.visible = ocupado
        self.btn_reconciliar.disabled = ocupado
        if ocupado:
            self.btn_revertir.disabled = True
        else:
            self._sincronizar_reversion()
        self.page.update()

    def _sincronizar_reversion(self) -> None:
        hay_procesado = self.ruta_procesada is not None and self.ruta_procesada.exists()
        self.btn_revertir.disabled = not hay_procesado
        self.btn_revertir.bgcolor = "#4CAF50" if hay_procesado else ft.Colors.GREY_500

    # ------------------------------------------------------------ acciones --
    async def on_reconciliar(self, e: ft.Event[ft.Button]) -> None:
        self._ocupado(True)
        self.registro.limpiar()
        try:
            ruta_diccionario = validar_ruta(
                self.archivo_diccionario.valor, "Diccionario Maestro"
            )
            ruta_analisis = validar_ruta(
                self.archivo_analisis.valor, "archivo para analizar"
            )

            hoja_diccionario = await resolver_hoja(
                self.page,
                ruta_diccionario,
                COLUMNAS_DICCIONARIO,
                "Diccionario Maestro",
                self.log,
                preferida=HOJA_PREFERIDA_DICCIONARIO,
            )
            hoja_analisis = await resolver_hoja(
                self.page,
                ruta_analisis,
                COLUMNAS_ANALISIS,
                "archivo para analizar",
                self.log,
            )

            def tarea() -> Path:
                self.log(f"📂 Diccionario Maestro: {ruta_diccionario.name}")
                self.log(f"📂 Archivo para analizar: {ruta_analisis.name}\n")
                salida, _stats = reconciliar(
                    leer_hoja(ruta_analisis, hoja_analisis),
                    leer_hoja(ruta_diccionario, hoja_diccionario),
                    ruta_analisis,
                    ruta_diccionario,
                    self.log,
                )
                return salida

            salida = await self.registro.ejecutar(tarea)
            self.ruta_procesada = salida
            self.ruta_diccionario = ruta_diccionario
            self.log("\n✔ Proceso inverso disponible: «Revertir IDs → Nombres».")
            await mostrar_dialogo(self.page, "Finalizado", "Proceso completado correctamente.")
        except ProcesoIncompleto as exc:
            self.log(f"\n⚠ {exc}")
            await mostrar_dialogo(
                self.page, "Coincidencia incompleta", str(exc), es_error=True
            )
        except Exception as exc:  # noqa: BLE001 - se informa al usuario
            self.log(f"\n⛔ ERROR: {exc}")
            await mostrar_dialogo(self.page, "Error", str(exc), es_error=True)
        finally:
            self._ocupado(False)

    async def on_revertir(self, e: ft.Event[ft.Button]) -> None:
        self._ocupado(True)
        try:
            if self.ruta_procesada is None or not self.ruta_procesada.exists():
                raise RuntimeError(
                    "No se encontró el archivo _PROCESADO.\n"
                    "Ejecutá primero «Validar y reconciliar»."
                )
            if self.ruta_diccionario is None:
                raise RuntimeError("No hay un Diccionario Maestro asociado a este ciclo.")

            ruta_procesada = self.ruta_procesada
            ruta_diccionario = self.ruta_diccionario

            def tarea() -> Path:
                self.log("═" * 50)
                self.log("  PROCESO INVERSO: IDs → NOMBRES (hoja validacion)")
                self.log("═" * 50)
                self.log(f"\nLeyendo hoja 'VALIDACION' de: {ruta_diccionario.name}...")
                df_validacion = leer_validacion(ruta_diccionario, self.log)
                mapa_inverso = construir_mapa_inverso(df_validacion)
                self.log(f"  Mapa inverso construido: {len(mapa_inverso)} entradas.")
                salida = revertir(
                    ruta_procesada,
                    [COLUMNA_A_REVERTIR],
                    mapa_inverso,
                    "Archivo procesado",
                    self.log,
                )
                self.log("\n" + "═" * 50)
                self.log("  ✔ PROCESO INVERSO COMPLETADO")
                self.log("═" * 50)
                return salida

            await self.registro.ejecutar(tarea)
            await mostrar_dialogo(
                self.page, "Proceso inverso", "Reversión completada correctamente."
            )
        except Exception as exc:  # noqa: BLE001 - se informa al usuario
            self.log(f"\n⛔ ERROR en proceso inverso: {exc}")
            await mostrar_dialogo(self.page, "Error", str(exc), es_error=True)
        finally:
            self._ocupado(False)
