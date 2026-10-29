"""Pestaña «Productividad Imputados»: eventos con imputados."""

from __future__ import annotations

import flet as ft

from core import calculo_formulas
from core.comunes import validar_ruta

from .comunes import VistaBase, bloque_titulo, mostrar_dialogo
from .seleccion import CampoArchivo


class VistaFormulas(VistaBase):
    """Genera ``DATOS/productividad_imputados_calculado.xlsx`` desde el archivo maestro."""

    titulo = "Calculo"
    icono = ft.Image(
        src="/Users/alexisjankowicz/Python/flet_scoring_poli/src/icons/calculadora.png",  # Ruta relativa dentro de la carpeta assets
        width=32,
        height=32,
        fit=ft.BoxFit.CONTAIN,
    )
    etiqueta_registro = "Registro — Calculo con formulas para Imputados"

    def __init__(self, page: ft.Page) -> None:
        super().__init__(page)
        self.archivo = CampoArchivo(
            page, "Listado de imputados:", "Seleccionar listado de imputados"
        )
        # Si ya existe el listado generado por Imputados, se propone solo.
        ruta_sugerida = calculo_formulas.DATOS_DIR / calculo_formulas.ARCHIVO_LISTADO
        if ruta_sugerida.exists():
            self.archivo.campo.value = str(ruta_sugerida)
        self.btn_procesar = ft.Button(
            "Procesar archivo",
            icon=ft.Icons.PLAY_ARROW,
            bgcolor="#2196F3",
            color=ft.Colors.WHITE,
            on_click=self.on_procesar,
        )

    def build(self) -> ft.Control:
        extra = "PUNTAJE EXTRA --> APLICA +25% POR ZONA"
        ayuda = ft.Container(
            bgcolor="#F1F5F9",
            border_radius=8,
            alignment=ft.Alignment.CENTER_LEFT,
            padding=4,
            expand=1,
            content=ft.Column(
                [
                    ft.Text("Reglas aplicadas", size=13, weight=ft.FontWeight.BOLD),
                    ft.Text(
                        "  • INTERVENCION DESTACADA -> PUNTOS BASE 1,50*\n",
                        size=12,
                    ),
                    ft.Text(
                        "  • ROBO/HURTO -> PUNTOS BASE 1,25*\n",
                        size=12,
                    ),
                    ft.Text(
                        "  • DELITO -> PUNTOS BASE 0,50*\n",
                        size=12,
                    ),
                    ft.Text(
                        "  • CONTRAVENTORES -> PUNTOS BASE 0,25\n",
                        size=12,
                    ),
                    ft.Text(
                        "  • COLABORACION CON INTERVENCION PROPIA -> PUNTOS BASE 0,25*\n",
                        size=12,
                    ),
                    ft.Text(
                        f"  * {extra}\n",
                        size=14,
                    ),
                    ft.Text(
                        f"Salida: {calculo_formulas.DATOS_DIR / calculo_formulas.ARCHIVO_SALIDA}",
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
                    "Calculo con formulas para Imputados",
                    "Del archivo generado por el módulo de Productividad Imputados, se calculan los puntajes de eventos con imputados.",
                ),
                self.archivo.build(),
                ft.Row([self.btn_procesar, self.progreso], spacing=12),
                ft.Row(
                    [ayuda, self.registro.campo],
                    spacing=10,
                    expand=True,
                    vertical_alignment=ft.CrossAxisAlignment.START,
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
            ruta = validar_ruta(self.archivo.valor, "listado de imputados")
            resultado = await self.registro.ejecutar(
                calculo_formulas.procesar, ruta, self.registro.log
            )
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
