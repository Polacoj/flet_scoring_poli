"""Pestaña «Productividad General»: unificación de eventos del CMU."""

from __future__ import annotations

import flet as ft

from core import productividad_general as general
from core.comunes import validar_ruta

from .comunes import VistaBase, bloque_titulo, mostrar_dialogo
from .seleccion import CampoArchivo


class VistaGeneral(VistaBase):
    """Genera ``eventos_unificados.xlsx`` a partir del archivo maestro del CMU."""

    titulo = "Productividad General"
    icono = ft.Image(
        src="/Users/alexisjankowicz/Python/flet_scoring_poli/src/icons/productividad_color.png",  # Ruta relativa dentro de la carpeta assets
        width=32,
        height=32,
        fit=ft.BoxFit.CONTAIN,
        )
    
    etiqueta_registro = "Registro — Productividad General"

    def __init__(self, page: ft.Page) -> None:
        super().__init__(page)
        self.archivo = CampoArchivo(
            page, "Archivo maestro (CMU):", "Seleccionar archivo maestro (CMU)"
        )
        self.btn_procesar = ft.Button(
            f"Procesar y generar {general.ARCHIVO_SALIDA}",
            icon=ft.Icons.PLAY_ARROW,
            bgcolor="#2196F3",
            color=ft.Colors.WHITE,
            on_click=self.on_procesar,
        )

    def build(self) -> ft.Control:
        detalle_columnas = "\n".join(
            f"  • Col {numero}: {nombre}"
            for numero, nombre in general.CORRESPONDENCIA.items()
        )
        ayuda = ft.Container(
            bgcolor="#F1F5F9",
            border_radius=8,
            padding=4,
            alignment=ft.Alignment.CENTER_LEFT,
            expand=1,
            content=ft.Column(
                [   
                    ft.Text("Qué genera", size=13, weight=ft.FontWeight.BOLD),
                    ft.Text(
                        f"Un único archivo {general.ARCHIVO_SALIDA} (hoja "
                        f"«{general.HOJA_SALIDA}») en la misma carpeta del archivo "
                        "elegido, combinando:",
                        size=12,
                    ),
                    ft.Text("  • eventos principales extraídos por posición de columna,", size=12),
                    ft.Text("  • colaboración con intervención propia,", size=12),
                    ft.Text("  • contraventores y detenidos,", size=12),
                    ft.Text("  • eventos integrados/destacados.", size=12),
                    ft.Text(
                        "Columnas extraídas (numeración de Excel):",
                        size=12,
                        weight=ft.FontWeight.BOLD,
                    ),
                    ft.Text(detalle_columnas, size=12, font_family="monospace"),
                    ft.Text(
                        "Las posiciones intermedias se conservan como columnas VACIO_*.",
                        size=11,
                        italic=True,
                    ),
                ],
                spacing=3,
            ),
        )
        return ft.Column(
            [
                bloque_titulo(
                    "Productividad General",
                    "Procesa el archivo maestro del CMU y unifica todos los eventos "
                    "en un solo Excel.",
                ),
                self.archivo.build(),
                ft.Row([self.btn_procesar, self.progreso], spacing=12),
                ft.Row(
                    [ayuda, self.registro.campo],
                    expand=True,
                    spacing=10,
                    vertical_alignment=ft.CrossAxisAlignment.STRETCH,
                ),
            ],
            expand=True,
            spacing=10,
        )

    async def on_procesar(self, e: ft.Event[ft.Button]) -> None:
        self.progreso.visible = True
        self.btn_procesar.disabled = True
        self.registro.limpiar()
        self.page.update()
        try:
            ruta = validar_ruta(self.archivo.valor, "archivo maestro (CMU)")
            resultado = await self.registro.ejecutar(general.procesar, ruta, self.registro.log)
            self.log("")
            self.log(resultado.resumen())
            await mostrar_dialogo(self.page, "Proceso completado", resultado.resumen())
        except Exception as exc:  # noqa: BLE001 - se informa al usuario
            self.log(f"\n⛔ ERROR: {exc}")
            await mostrar_dialogo(self.page, "Error", str(exc), es_error=True)
        finally:
            self.progreso.visible = False
            self.btn_procesar.disabled = False
            self.page.update()
