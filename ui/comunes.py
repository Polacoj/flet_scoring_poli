"""Widgets y utilidades compartidas por las tres pestañas.

Incluye el :class:`PanelRegistro`, que permite que la lógica (ejecutada en un
hilo aparte con ``asyncio.to_thread``) escriba en la interfaz sin bloquearla y
sin tocar controles desde otro hilo.
"""

from __future__ import annotations

import asyncio
import queue
import threading
from collections.abc import Callable
from pathlib import Path
from typing import Any

import flet as ft

from core.comunes import hojas_con_columnas

LIMITE_CARACTERES_REGISTRO = 200_000
_HILO_PRINCIPAL = threading.main_thread()


class PanelRegistro:
    """Registro de ejecución seguro para escribir desde un hilo de trabajo."""

    def __init__(self, page: ft.Page, titulo: str) -> None:
        self.page = page
        self.campo = ft.TextField(
            multiline=True,
            read_only=True,
            expand=True,
            min_lines=10,
            text_size=13,
            label=titulo,
        )
        self.cola: queue.Queue[str] = queue.Queue()

    def log(self, mensaje: Any) -> None:
        """Encola un mensaje; es seguro llamarlo desde cualquier hilo."""
        self.cola.put(str(mensaje))
        if threading.current_thread() is _HILO_PRINCIPAL:
            self.volcar()

    def volcar(self) -> None:
        """Pasa al panel los mensajes encolados (desde el hilo de eventos)."""
        lineas = self._drenar()
        if not lineas:
            return
        texto = (self.campo.value or "") + "\n".join(lineas) + "\n"
        if len(texto) > LIMITE_CARACTERES_REGISTRO:
            texto = texto[-LIMITE_CARACTERES_REGISTRO:]
        self.campo.value = texto
        self.page.update()

    def limpiar(self) -> None:
        self._drenar()
        self.campo.value = ""
        self.page.update()

    def _drenar(self) -> list[str]:
        lineas: list[str] = []
        while True:
            try:
                lineas.append(self.cola.get_nowait())
            except queue.Empty:
                return lineas

    async def ejecutar(self, funcion: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
        """Corre ``funcion`` en un hilo y va volcando el registro en vivo."""
        tarea = asyncio.ensure_future(asyncio.to_thread(funcion, *args, **kwargs))
        while not tarea.done():
            self.volcar()
            await asyncio.sleep(0.05)
        self.volcar()
        return await tarea


async def mostrar_dialogo(
    page: ft.Page, titulo: str, mensaje: str, *, es_error: bool = False
) -> None:
    """Diálogo modal de aviso/error, equivalente a los ``QMessageBox`` previos."""
    bucle = asyncio.get_running_loop()
    listo: asyncio.Future[None] = bucle.create_future()

    def cerrar(_e=None) -> None:
        if not listo.done():
            listo.set_result(None)
        page.pop_dialog()

    page.show_dialog(
        ft.AlertDialog(
            modal=True,
            title=ft.Text(f"{'⛔ ' if es_error else ''}{titulo}"),
            content=ft.Text(mensaje, selectable=True),
            actions=[ft.Button("Aceptar", on_click=cerrar)],
            on_dismiss=lambda e: cerrar(),
        )
    )
    await listo


async def elegir_hoja(page: ft.Page, titulos: list[str], etiqueta: str) -> str | None:
    """Diálogo para elegir hoja cuando hay más de una candidata."""
    bucle = asyncio.get_running_loop()
    listo: asyncio.Future[str | None] = bucle.create_future()
    desplegable = ft.Dropdown(
        label="Hoja",
        options=[ft.DropdownOption(key=t, text=t) for t in titulos],
        value=titulos[0],
        width=420,
    )

    def resolver(valor: str | None) -> None:
        if not listo.done():
            listo.set_result(valor)
        page.pop_dialog()

    page.show_dialog(
        ft.AlertDialog(
            modal=True,
            title=ft.Text("Seleccionar hoja"),
            content=ft.Column(
                [
                    ft.Text(f"{etiqueta} tiene varias hojas válidas. Seleccioná una:"),
                    desplegable,
                ],
                tight=True,
                width=440,
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda e: resolver(None)),
                ft.Button("Aceptar", on_click=lambda e: resolver(desplegable.value)),
            ],
            on_dismiss=lambda e: resolver(None) if not listo.done() else None,
        )
    )
    return await listo


async def resolver_hoja(
    page: ft.Page,
    ruta: Path,
    requeridas: list[str],
    etiqueta: str,
    log: Callable[[str], None],
    preferida: str | None = None,
) -> str:
    """Elige la hoja que contiene las columnas requeridas.

    Sólo pregunta cuando hay más de una hoja válida; si no hay ninguna, informa
    qué columnas se encontraron en cada hoja.
    """
    validas, encabezados = hojas_con_columnas(ruta, requeridas)
    if not validas:
        detalle = " | ".join(
            f"'{nombre}': {columnas}" for nombre, columnas in encabezados.items()
        )
        raise RuntimeError(
            f"Ninguna hoja de {ruta.name} contiene las columnas {requeridas}.\n"
            f"Hojas y columnas encontradas → {detalle}"
        )

    if preferida:
        objetivo = preferida.strip().upper()
        for nombre in validas:
            if nombre.strip().upper() == objetivo:
                log(f"Hoja seleccionada automáticamente: {nombre}")
                return nombre

    if len(validas) == 1:
        log(f"Hoja seleccionada automáticamente: {validas[0]}")
        return validas[0]

    elegida = await elegir_hoja(page, validas, etiqueta)
    if not elegida:
        raise RuntimeError(f"Se canceló la selección de hoja para {etiqueta}.")
    log(f"Hoja seleccionada: {elegida}")
    return elegida


def bloque_titulo(titulo: str, descripcion: str) -> ft.Column:
    """Encabezado común de cada pestaña."""
    return ft.Column(
        [
            ft.Text(titulo, size=20, weight=ft.FontWeight.BOLD, color="#2c3e50"),
            ft.Text(descripcion, size=13, color="#7f8c8d"),
        ],
        spacing=2,
    )


class VistaBase:
    """Base de las tres pestañas: título, ícono y registro propio."""

    titulo = "Vista"
    icono = ft.Icons.CIRCLE
    etiqueta_registro = "Registro de ejecución"

    def __init__(self, page: ft.Page) -> None:
        self.page = page
        self.registro = PanelRegistro(page, self.etiqueta_registro)
        self.progreso = ft.ProgressRing(visible=False, width=20, height=20, stroke_width=2)

    def log(self, mensaje: Any) -> None:
        self.registro.log(mensaje)

    def build(self) -> ft.Control:
        raise NotImplementedError
