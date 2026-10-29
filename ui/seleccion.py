"""Selección de archivos en capas, a prueba de fallos.

Contexto
--------
En macOS 26 el servicio ``FilePicker`` de Flet puede no abrir ningún diálogo y
no informar ningún error (issues #5700 y #5886 de ``flet-dev/flet``): el botón
parece no responder. Por eso cada campo de archivo ofrece tres caminos:

1. **Examinar** — ``ft.PickFiles``: acción de cliente que abre el diálogo dentro
   del gesto del clic, sin ida y vuelta a Python (camino recomendado en 1.0).
2. **Diálogo del sistema** — abre el selector nativo del sistema operativo
   (macOS: ``osascript``; Windows: PowerShell; Linux: zenity/kdialog) sin pasar
   por Flet ni por sus librerías.
3. **Ruta manual** — el campo es editable: la ruta se puede escribir o pegar.

Los tres caminos escriben en el mismo ``TextField``, así que el resto de la
aplicación no necesita saber cuál se usó.
"""

from __future__ import annotations

import asyncio
import shutil
import subprocess
import sys

import flet as ft

EXTENSIONES_EXCEL = ["xlsx", "xlsm", "xls"]

#: Segundos que se esperan antes de sugerir el diálogo del sistema.
ESPERA_SUGERENCIA = 6.0

COLOR_AVISO = "#b45309"


# ---------------------------------------------------------------------------
# Diálogo nativo del sistema operativo (bloqueante: usar en un hilo)
# ---------------------------------------------------------------------------
def _elegir_macos(titulo: str, extensiones: list[str]) -> str | None:
    prompt = titulo.replace('"', "'")
    if extensiones:
        tipos = ", ".join(f'"{ext}"' for ext in extensiones)
        script = (
            f'POSIX path of (choose file with prompt "{prompt}" of type {{{tipos}}})'
        )
    else:
        script = f'POSIX path of (choose file with prompt "{prompt}")'
    proceso = subprocess.run(
        ["osascript", "-e", script], capture_output=True, text=True, check=False
    )
    if proceso.returncode != 0:
        error = (proceso.stderr or "").strip()
        if "User canceled" in error or "-128" in error:
            return None  # el usuario canceló: no es un error
        raise RuntimeError(
            error or f"osascript terminó con código {proceso.returncode}"
        )
    return proceso.stdout.strip() or None


def _elegir_windows(titulo: str, extensiones: list[str]) -> str | None:
    filtro = ";".join(f"*.{ext}" for ext in extensiones) or "*.*"
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$d = New-Object System.Windows.Forms.OpenFileDialog;"
        f"$d.Title = '{titulo}';"
        f"$d.Filter = 'Archivos ({filtro})|{filtro}|Todos (*.*)|*.*';"
        "if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) "
        "{ Write-Output $d.FileName }"
    )
    proceso = subprocess.run(
        ["powershell", "-NoProfile", "-STA", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
    )
    if proceso.returncode != 0:
        raise RuntimeError(
            (proceso.stderr or "").strip() or "No se pudo abrir el diálogo"
        )
    return proceso.stdout.strip() or None


def _elegir_linux(titulo: str, extensiones: list[str]) -> str | None:
    if shutil.which("zenity"):
        comando = ["zenity", "--file-selection", f"--title={titulo}"]
        for ext in extensiones:
            comando.append(f"--file-filter=*.{ext}")
        proceso = subprocess.run(comando, capture_output=True, text=True, check=False)
        if proceso.returncode == 1 and not proceso.stdout.strip():
            return None  # cancelado
        if proceso.returncode != 0:
            raise RuntimeError((proceso.stderr or "").strip() or "zenity falló")
        return proceso.stdout.strip() or None

    if shutil.which("kdialog"):
        patrones = " ".join(f"*.{ext}" for ext in extensiones) or "*"
        proceso = subprocess.run(
            ["kdialog", "--title", titulo, "--getopenfilename", ".", patrones],
            capture_output=True,
            text=True,
            check=False,
        )
        if proceso.returncode == 1 and not proceso.stdout.strip():
            return None  # cancelado
        if proceso.returncode != 0:
            raise RuntimeError((proceso.stderr or "").strip() or "kdialog falló")
        return proceso.stdout.strip() or None

    raise RuntimeError(
        "No se encontró zenity ni kdialog para abrir el diálogo del sistema."
    )


def elegir_archivo_nativo(
    titulo: str, extensiones: list[str] | None = None
) -> str | None:
    """Abre el diálogo de archivos nativo del sistema.

    Devuelve la ruta elegida o ``None`` si se canceló. Es bloqueante: hay que
    invocarla con ``asyncio.to_thread``.
    """
    extensiones = list(extensiones or EXTENSIONES_EXCEL)
    if sys.platform == "darwin":
        return _elegir_macos(titulo, extensiones)
    if sys.platform.startswith("win"):
        return _elegir_windows(titulo, extensiones)
    return _elegir_linux(titulo, extensiones)


# ---------------------------------------------------------------------------
# Campo de archivo con los tres caminos
# ---------------------------------------------------------------------------
class CampoArchivo:
    """Etiqueta + ruta + «Examinar» (Flet) + «Diálogo del sistema» + aviso."""

    def __init__(
        self,
        page: ft.Page,
        etiqueta: str,
        titulo_dialogo: str,
        extensiones: list[str] | None = None,
        ancho_etiqueta: int = 250,
    ) -> None:
        self.page = page
        self.etiqueta = etiqueta
        self.titulo_dialogo = titulo_dialogo
        self.extensiones = list(extensiones or EXTENSIONES_EXCEL)
        self.ancho_etiqueta = ancho_etiqueta
        self._pendiente = False

        # Servicio de Flet para el camino 1 (se auto-registra con la página).
        self.picker = ft.FilePicker(on_result=self._on_result)

        self.campo = ft.TextField(
            expand=True,
            hint_text="Elegí un archivo, o escribí/pegá la ruta",
            on_change=self._al_escribir,
        )

        """self.btn_examinar = ft.Button(
            "Examinar",
            icon=ft.Icons.FOLDER_OPEN,
            tooltip="Abre el selector de Flet (acción de cliente, dentro del gesto)",
            action=ft.PickFiles(
                self.picker,
                dialog_title=titulo_dialogo,
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=self.extensiones,
            ),
            on_click=self._vigilar,
        )"""  # no se usa porque en macOS 26 puede no abrir el diálogo y no informar error

        self.btn_sistema = ft.Button(
            "Archivo",
            icon=ft.Image(
                src="/Users/alexisjankowicz/Python/flet_scoring_poli/src/icons/descarga.png",  # Ruta relativa dentro de la carpeta assets
                width=16,
                height=16,
                # color=ft.Colors.WHITE,
                color="#626b75",
                fit=ft.BoxFit.CONTAIN,
            ),
            tooltip=(
                "Abre el selector nativo del sistema operativo, sin usar las "
                "librerías de Flet. Puede aparecer detrás de esta ventana."
            ),
            on_click=self._usar_dialogo_sistema,
        )

        self.aviso = ft.Text("", size=11, color=COLOR_AVISO, visible=False)

    # ------------------------------------------------------------ API pública
    @property
    def valor(self) -> str:
        return (self.campo.value or "").strip()

    def establecer(self, ruta: str | None) -> None:
        """Escribe la ruta elegida y limpia el aviso."""
        if not isinstance(ruta, str) or not ruta.strip():
            self._avisar(
                "El diálogo no devolvió una ruta de archivo. "
                "Probá «Diálogo del sistema» o pegá la ruta a mano."
            )
            return
        self.campo.value = ruta.strip()
        self._limpiar_aviso()
        self.page.update()

    def build(self) -> ft.Control:
        return ft.Column(
            [
                ft.Row(
                    [
                        ft.Text(self.etiqueta, width=self.ancho_etiqueta),
                        self.campo,
                        # self.btn_examinar,
                        self.btn_sistema,
                    ],
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=8,
                ),
                self.aviso,
            ],
            spacing=2,
        )

    # ------------------------------------------------------------ callbacks
    def _al_escribir(self, _e: ft.Event[ft.TextField]) -> None:
        self._limpiar_aviso()

    async def _on_result(self, e: ft.Event[ft.FilePicker]) -> None:
        """Llega la selección hecha con la acción de cliente (camino 1)."""
        self._pendiente = False
        archivos = list(getattr(e, "files", None) or [])
        if not archivos:
            self._avisar("Se canceló la selección.", informativo=True)
            return
        self.establecer(getattr(archivos[0], "path", None))

    async def _vigilar(self, _e: ft.Event[ft.Button]) -> None:
        """Se ejecuta después de que el cliente lanzó la acción.

        Si el diálogo no devuelve nada en unos segundos, sugiere el camino 2.
        """
        self._pendiente = True
        self._limpiar_aviso()
        await asyncio.sleep(ESPERA_SUGERENCIA)
        if self._pendiente:
            self._avisar(
                "¿No se abrió el diálogo? Usá «Diálogo del sistema» "
                "o escribí/pegá la ruta en el campo."
            )

    async def _usar_dialogo_sistema(self, _e: ft.Event[ft.Button]) -> None:
        """Camino 2: diálogo nativo del sistema, sin Flet de por medio."""
        self.btn_sistema.disabled = True
        self.page.update()
        try:
            ruta = await asyncio.to_thread(
                elegir_archivo_nativo, self.titulo_dialogo, self.extensiones
            )
            if ruta:
                self.establecer(ruta)
            else:
                self._avisar("No se seleccionó ningún archivo.", informativo=True)
        except Exception as exc:  # noqa: BLE001 - se muestra al usuario
            self._avisar(f"No se pudo abrir el diálogo del sistema: {exc}")
        finally:
            self.btn_sistema.disabled = False
            self.page.update()

    # ------------------------------------------------------------ internos
    def _avisar(self, mensaje: str, informativo: bool = False) -> None:
        self.aviso.value = mensaje
        self.aviso.color = "#475569" if informativo else COLOR_AVISO
        self.aviso.visible = True
        self.page.update()

    def _limpiar_aviso(self) -> None:
        if self.aviso.visible:
            self.aviso.visible = False
            self.aviso.value = ""
            self.page.update()
