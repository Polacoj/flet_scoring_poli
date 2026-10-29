"""Sistema de Scoring — Oficiales.

Aplicación Flet 1.0 que reúne los tres módulos en una sola ventana con
pestañas, cada una con su propio registro de ejecución:

* **Diccionario** — reconciliación de nombres → LP y proceso inverso.
* **Productividad General** — unificación de eventos del CMU.
* **Productividad Imputados** — eventos con imputados.

Ejecutar con::

    ./scoring-poli
    # o bien
    uv run python main.py
"""

from __future__ import annotations

import flet as ft

from ui.vista_diccionario import VistaDiccionario
from ui.vista_general import VistaGeneral
from ui.vista_imputados import VistaImputados
from ui.vista_formulas import VistaFormulas

TITULO = "Sistema de Scoring — Oficiales"


async def main(page: ft.Page) -> None:
    page.title = TITULO
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 16
    page.window.width = 1100
    page.window.height = 860
    page.window.min_width = 900
    page.window.min_height = 640

    # Cada campo de archivo crea su propio FilePicker y se auto-registra con la
    # página: así la acción de cliente y su resultado no se mezclan entre campos.
    vistas = [
        VistaGeneral(page),
        VistaImputados(page),
        VistaDiccionario(page),
        # VistaFormulas(page), TODO deshabilitar hasta que se resuelva si realmente se necesita o no, y si se mantiene el nombre de la pestaña.
    ]

    page.add(
        ft.Column(
            [
                ft.Text(TITULO, size=26, weight=ft.FontWeight.BOLD, color="#2c3e50"),
                ft.Tabs(
                    length=len(vistas),
                    selected_index=0,
                    expand=True,
                    content=ft.Column(
                        expand=True,
                        controls=[
                            ft.TabBar(
                                tabs=[
                                    ft.Tab(label=vista.titulo, icon=vista.icono)
                                    for vista in vistas
                                ]
                            ),
                            ft.TabBarView(
                                expand=True,
                                controls=[vista.build() for vista in vistas],
                            ),
                        ],
                    ),
                ),
            ],
            expand=True,
            spacing=10,
        )
    )


if __name__ == "__main__":
    ft.run(main)
