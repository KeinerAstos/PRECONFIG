# app.py
import flet as ft
import threading
from scraper import ejecutar_scraping


# ---------------------------------------------------
# FUNCIÓN UI PRINCIPAL
# ---------------------------------------------------
def main(page: ft.Page):

    page.title = "Scraper RR - Claro"
    page.window_width = 800
    page.window_height = 600
    page.theme_mode = ft.ThemeMode.LIGHT

    # Variables UI
    ruta_archivo = ft.Text("", size=13, color="#444")
    usuario_input = ft.TextField(label="Usuario", width=250)
    password_input = ft.TextField(label="Contraseña", width=250, password=True, can_reveal_password=True)

    logs = ft.TextField(
        value="",
        multiline=True,
        min_lines=18,
        max_lines=18,
        read_only=True,
        width=780,
        border=ft.InputBorder.OUTLINE
    )

    # Handler para recibir mensajes del scraper
    def agregar_log(texto):
        logs.value += texto + "\n"
        logs.update()

    # ---------------------------------------
    # FilePicker
    # ---------------------------------------
    def file_picker_result(e: ft.FilePickerResultEvent):
        if e.files:
            ruta_archivo.value = e.files[0].path
        else:
            ruta_archivo.value = "No se seleccionó archivo."
        ruta_archivo.update()

    file_picker = ft.FilePicker(on_result=file_picker_result)
    page.overlay.append(file_picker)

    # ---------------------------------------
    # ACCIÓN DEL BOTÓN INICIAR
    # ---------------------------------------
    def comenzar_scraping(e):

        if ruta_archivo.value == "":
            agregar_log("⚠ Seleccione un archivo excel primero.")
            return

        if usuario_input.value.strip() == "" or password_input.value.strip() == "":
            agregar_log("⚠ Ingrese usuario y contraseña.")
            return

        logs.value = ""
        logs.update()
        agregar_log("🚀 Iniciando proceso...\n")

        # Ejecutar en hilo separado
        threading.Thread(
            target=lambda: ejecutar_scraping(
                usuario_input.value,
                password_input.value,
                ruta_archivo.value,
                agregar_log
            ),
            daemon=True
        ).start()

    # ---------------------------------------
    # ESTRUCTURA VISUAL
    # ---------------------------------------
    page.add(
        ft.Text("Scraper RR - Claro", size=22, weight=ft.FontWeight.BOLD),
        ft.Divider(),

        ft.Row([
            usuario_input,
            password_input
        ], spacing=20),

        ft.Row([
            ft.ElevatedButton(
                "Seleccionar archivo Excel",
                icon=ft.Icons.UPLOAD_FILE,
                on_click=lambda _: file_picker.pick_files(
                    allow_multiple=False,
                    allowed_extensions=["xlsx"]
                )
            ),
            ruta_archivo
        ], spacing=20),

        ft.Divider(),

        ft.ElevatedButton(
            "Iniciar Scraping",
            icon=ft.Icons.PLAY_ARROW,
            bgcolor="#007BFF",
            color="white",
            width=200,
            on_click=comenzar_scraping
        ),

        ft.Divider(),

        ft.Text("Registro de Logs:", size=16, weight=ft.FontWeight.BOLD),
        logs
    )


# Lanzar app
ft.app(target=main)
