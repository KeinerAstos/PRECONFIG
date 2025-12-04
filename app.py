# app.py
import flet as ft
from scraper import iniciar_driver, login, procesar_rr

def main(page: ft.Page):
    page.title = "Automatización RR"
    page.vertical_alignment = ft.MainAxisAlignment.START

    usuario_input = ft.TextField(label="Usuario", width=300)
    password_input = ft.TextField(label="Contraseña", password=True, can_reveal_password=True, width=300)
    log_output = ft.Column()

    def ejecutar_scraper(e):
        page.update()
        log_output.controls.append(ft.Text("Iniciando navegador..."))
        page.update()

        driver = iniciar_driver()
        driver, wait = login(driver, usuario_input.value, password_input.value)
        log_output.controls.append(ft.Text("Login exitoso, procesando RR..."))
        page.update()

        procesar_rr(driver, wait)
        log_output.controls.append(ft.Text("Proceso finalizado."))
        page.update()

    ejecutar_btn = ft.ElevatedButton("Ejecutar", on_click=ejecutar_scraper)

    page.add(usuario_input, password_input, ejecutar_btn, log_output)

ft.app(target=main)
