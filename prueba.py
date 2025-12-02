import flet as ft
import threading
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time
import json
from datetime import datetime
import subprocess
import tempfile
import shutil


# ==========================================================
#                   FUNCIONES DE SEGURIDAD
# ==========================================================

def safe_get_url(driver, url, handler, intentos=5):
    """Carga una URL con reintentos automáticos en caso de 500/502."""
    for i in range(1, intentos + 1):
        try:
            driver.get(url)
            time.sleep(2)

            html = driver.page_source.lower()

            if any(err in html for err in [
                "error 500", "error 502", "bad gateway",
                "esta página no funciona", "internal server error"
            ]):
                handler(f"⚠ PAGE ERROR {i}/{intentos}: 500/502 detectado, reintentando...")
                time.sleep(3)
                continue

            return True

        except Exception as e:
            handler(f"⚠ URL ERROR {i}/{intentos}: {e}")
            time.sleep(3)

    handler("❌ No se pudo cargar la URL después de varios intentos.")
    return False


def safe_find_retry(driver, by, value, handler, timeout=10, retries=3):
    """Busca un elemento con reintentos."""
    for i in range(1, retries + 1):
        try:
            elem = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
            return elem
        except:
            handler(f"⚠ No se encontró elemento ({value}) intento {i}/{retries}")
            time.sleep(1)
    return None


def safe_click_retry(driver, element, handler, retries=3):
    """Clic robusto con JS fallback."""
    if not element:
        return False

    for i in range(1, retries + 1):
        try:
            element.click()
            return True
        except:
            try:
                driver.execute_script("arguments[0].click();", element)
                return True
            except:
                handler(f"⚠ Error al hacer clic intento {i}/{retries}")
                time.sleep(1)
    return False


def safe_switch_frame(driver, frame_id, handler, retries=3):
    """Cambio de iframe con reintentos."""
    for i in range(1, retries + 1):
        try:
            WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, frame_id))
            )
            return True
        except:
            handler(f"⚠ Iframe '{frame_id}' no cargó, intento {i}/{retries}")
            time.sleep(1)
    return False


def safe_get_text(element):
    try:
        return element.text.strip()
    except:
        return "N/A"



# ==========================================================
#                   APP FLET PRINCIPAL
# ==========================================================

def main(page: ft.Page):
    page.title = "Scraping Agenda FO"
    page.window_width = 900
    page.window_height = 600

    archivo_label = ft.Text("Ningún archivo seleccionado", size=14)
    log_output = ft.Column(scroll="auto", expand=True)
    usuario_input = ft.TextField(label="Usuario")
    contrasena_input = ft.TextField(label="Contraseña", password=True, can_reveal_password=True)

    # FILE PICKER
    def seleccionar_archivo_result(e):
        if e.files:
            archivo_label.value = e.files[0].path
            page.update()

    archivo_picker = ft.FilePicker(on_result=seleccionar_archivo_result)
    page.overlay.append(archivo_picker)

    def abrir_picker(e):
        archivo_picker.pick_files()

    # LOG
    def log_handler(msg):
        log_output.controls.append(ft.Text(msg))
        page.update()

    # ==========================================================
    #                   SCRAPING PRINCIPAL
    # ==========================================================

    def ejecutar_scraping(handler):
        temp_profile = None
        driver = None

        try:
            handler("🔹 Iniciando navegador (ultrarobusto)...")

            chrome_options = Options()
            chrome_options.add_argument("--headless=new")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--disable-software-rasterizer")
            chrome_options.add_argument("--disable-features=VizDisplayCompositor")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--window-size=1920,1080")

            # User-Agent realista
            chrome_options.add_argument(
                "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
            )

            # Perfil temporal
            temp_profile = tempfile.mkdtemp()
            chrome_options.add_argument(f"--user-data-dir={temp_profile}")

            driver_path = ChromeDriverManager().install()
            driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)

            # ======================================================
            #                 LOGIN SEGURO
            # ======================================================

            if not safe_get_url(driver, "https://moduloagenda.cable.net.co/index.php", handler):
                handler("❌ No se pudo cargar el login.")
                return

            handler("🔹 Página cargada")

            user_field = safe_find_retry(driver, By.XPATH, "//td[contains(.,'Usuario')]/input", handler)
            pass_field = safe_find_retry(driver, By.XPATH, "//td[contains(.,'Contraseña')]/input", handler)

            if user_field:
                user_field.send_keys(usuario_input.value)
            if pass_field:
                pass_field.send_keys(contrasena_input.value)

            time.sleep(2)

            boton = safe_find_retry(driver, By.XPATH, "//button[@type='submit'] | //input[@type='submit']", handler)
            safe_click_retry(driver, boton, handler)

            handler("🔹 Login completo")

            # ======================================================
            #     EXTRAER MENÚ: Atras → Agendamiento → Agendar WFM
            # ======================================================

            safe_click_retry(driver, safe_find_retry(driver, By.ID, "imgAtrasMenu", handler), handler)
            safe_click_retry(driver, safe_find_retry(driver, By.XPATH, "//a[@title='Agendamiento']", handler), handler)

            safe_click_retry(driver, safe_find_retry(
                driver,
                By.XPATH,
                "//a[contains(@href,'Agendamiento/index.php') and contains(@title,'Agendar WFM')]",
                handler),
                handler
            )

            handler("🔹 Navegación completa")

            # ======================================================
            # LEER ARCHIVO
            # ======================================================
            ruta_archivo = archivo_label.value
            if not ruta_archivo:
                handler("❌ No se seleccionó archivo")
                return

            df = pd.read_excel(ruta_archivo)

            lista_rr = [
                x[:9] for x, y in zip(
                    df["Orden de trabajo"].astype(str),
                    df["Subtipo de la Orden de Trabajo"].astype(str)
                ) if y == "Entrega De Servicios FO"
            ]

            # LISTAS DE RESULTADOS
            notas_orden, estado_programa, fecha_programa, franja_programa = [], [], [], []
            nota_ofsc, cliente, direccion = [], [], []

            # ======================================================
            #               PROCESAR CADA ORDEN
            # ======================================================

            for orden in lista_rr:
                handler(f"🔵 Procesando orden {orden}")

                # Reintentos al cargar orden
                for intento in range(1, 4):
                    try:
                        tb = safe_find_retry(driver, By.ID, "TBorden", handler)
                        if tb:
                            tb.clear()
                            tb.send_keys(orden)
                            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", tb)

                        safe_click_retry(driver, safe_find_retry(driver, By.ID, "Rbot-O", handler), handler)
                        safe_click_retry(driver, safe_find_retry(driver, By.ID, "button", handler), handler)

                        time.sleep(2)
                        break
                    except:
                        handler(f"⚠ Error cargando orden {orden}, intento {intento}/3")
                        time.sleep(2)

                # ABRIR TAB ORDEN
                menu = safe_find_retry(driver, By.ID, "menuh-info-agendamiento", handler)
                if menu:
                    ActionChains(driver).move_to_element(menu).perform()

                safe_click_retry(driver, safe_find_retry(driver, By.ID, "orden_menu", handler), handler)

                time.sleep(1)

                # NOTAS DE ORDEN
                notas_elem = safe_find_retry(driver,
                                             By.XPATH,
                                             "//tr[th[contains(text(),'Notas de la Orden')]]/td/p",
                                             handler)
                notas_orden.append(safe_get_text(notas_elem))

                # SUSCRIPTOR
                cliente.append(
                    safe_get_text(safe_find_retry(
                        driver,
                        By.XPATH,
                        "//div[contains(@class,'bloquecss')]//div[contains(text(),'Suscriptor')]/following-sibling::div",
                        handler
                    ))
                )

                # DIRECCIÓN
                direccion.append(
                    safe_get_text(safe_find_retry(
                        driver,
                        By.XPATH,
                        "//div[contains(text(),'Dirección')]/following-sibling::div",
                        handler
                    ))
                )

                # VISITA
                safe_click_retry(driver, safe_find_retry(driver, By.ID, "visita_menu", handler), handler)

                estado_programa.append(
                    safe_get_text(safe_find_retry(driver,
                                                  By.XPATH,
                                                  "//th[contains(text(),'Estado de la Visita')]/following-sibling::td",
                                                  handler))
                )

                fecha_programa.append(
                    safe_get_text(safe_find_retry(driver,
                                                  By.XPATH,
                                                  "//th[contains(text(),'Fecha Programada')]/following-sibling::td",
                                                  handler))
                )

                franja_programa.append(
                    safe_get_text(safe_find_retry(driver,
                                                  By.XPATH,
                                                  "//th[contains(text(),'Franja Suscriptor')]/following-sibling::td",
                                                  handler))
                )

                # =====================================
                #        OFSC (super robusto)
                # =====================================

                safe_click_retry(driver, safe_find_retry(driver, By.ID, "ofsc_menu", handler), handler)

                if safe_switch_frame(driver, "iframe-ofsc", handler):

                    nota_guardada = "N/A"
                    tbody = safe_find_retry(driver, By.ID, "tbodyMainList", handler)

                    if tbody:
                        filas = tbody.find_elements(By.TAG_NAME, "tr")

                        for fila in filas:
                            celdas = fila.find_elements(By.TAG_NAME, "td")
                            if len(celdas) < 7:
                                continue

                            if celdas[5].text.strip().lower() == "pendiente":
                                boton = celdas[-1].find_element(By.CSS_SELECTOR, "button.checkDetails")
                                safe_click_retry(driver, boton, handler)
                                time.sleep(1)

                                notas = safe_find_retry(driver, By.ID, "notasorden", handler)
                                texto = notas.text.strip() if notas else "N/A"

                                # Intentar parsear JSON
                                try:
                                    data = json.loads(texto)
                                    if data.get("error"):
                                        nota_guardada = "SIN DATOS"
                                    else:
                                        nota_guardada = texto
                                except:
                                    nota_guardada = texto

                    nota_ofsc.append(nota_guardada)

                    # Salir del iframe
                    driver.switch_to.default_content()

                # VOLVER A MENÚ PRINCIPAL
                safe_click_retry(driver, safe_find_retry(driver, By.ID, "return-suscriptor", handler), handler)
                time.sleep(1)

            # =============================================
            #    EXPORTAR RESULTADOS
            # =============================================
            resultado_dir = os.path.join(os.getcwd(), "RESULTADO")
            os.makedirs(resultado_dir, exist_ok=True)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(resultado_dir, f"resultado_{timestamp}.xlsx")

            tabla = [{
                "orden": o,
                "estado_programa": es,
                "fecha_programada": fp,
                "franja_programada": fr,
                "cliente": c,
                "direccion": d,
                "nota_orden": n,
                "nota_ofsc": nf
            } for o, es, fp, fr, c, d, n, nf in zip(
                lista_rr, estado_programa, fecha_programa, franja_programa,
                cliente, direccion, notas_orden, nota_ofsc
            )]

            pd.DataFrame(tabla).to_excel(output_path, index=False)

            handler(f"✅ RESULTADO GUARDADO: {output_path}")

        except Exception as e:
            handler(f"❌ Error crítico total: {e}")

        finally:
            try:
                if driver:
                    driver.quit()
            except:
                pass
            try:
                shutil.rmtree(temp_profile, ignore_errors=True)
            except:
                pass


    # BOTONES
    def iniciar_scraping(e):
        threading.Thread(target=ejecutar_scraping, args=(log_handler,), daemon=True).start()

    def abrir_resultado(e):
        ruta = os.path.join(os.getcwd(), "RESULTADO")
        if os.name == "nt":
            subprocess.Popen(f'explorer "{ruta}"')
        else:
            subprocess.Popen(["open", ruta])

    # UI
    page.add(
        ft.Row([
            ft.Column([
                ft.Text("Menú", weight="bold", size=18),
                usuario_input,
                contrasena_input,
                ft.ElevatedButton("Seleccionar archivo", on_click=abrir_picker),
                archivo_label,
                ft.ElevatedButton("Iniciar Scraping", on_click=iniciar_scraping),
                ft.ElevatedButton("Abrir carpeta RESULTADO", on_click=abrir_resultado)
            ], width=250, spacing=10),
            ft.Column([
                ft.Text("Registro de Logs", weight="bold", size=16),
                log_output
            ], expand=True)
        ], expand=True)
    )


ft.app(target=main)
