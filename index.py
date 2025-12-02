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
import json
from datetime import datetime
import subprocess
from concurrent.futures import ThreadPoolExecutor

def main(page: ft.Page):
    page.title = "Scraping Agenda FO"
    page.window_width = 900
    page.window_height = 600

    # UI Elements
    archivo_label = ft.Text("Ningún archivo seleccionado", size=14)
    log_output = ft.Column(scroll="auto", expand=True)
    usuario_input = ft.TextField(label="Usuario")
    contrasena_input = ft.TextField(label="Contraseña", password=True, can_reveal_password=True)

    # File Picker
    def seleccionar_archivo_result(e: ft.FilePickerResultEvent):
        if e.files:
            archivo_label.value = e.files[0].path
            page.update()

    archivo_picker = ft.FilePicker(on_result=seleccionar_archivo_result)
    page.overlay.append(archivo_picker)

    def abrir_picker(e):
        archivo_picker.pick_files()

    # Log handler
    def log_handler(mensaje):
        log_output.controls.append(ft.Text(mensaje))
        page.update()

    # Función de scraping por orden
    def procesar_orden(orden, usuario_val, contrasena_val):
        resultado = {"orden": orden, "nota_orden": "N/A", "estado_programa": "N/A",
                     "fecha_programada": "N/A", "franja_programada": "N/A", "nota_ofsc": "N/A"}
        try:
            chrome_options = Options()
            chrome_options.add_argument("--headless=new")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-infobars")
            chrome_options.add_argument("--blink-settings=imagesEnabled=false")

            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
            wait = WebDriverWait(driver, 20)

            driver.get("https://moduloagenda.cable.net.co/index.php")

            # Login
            wait.until(EC.presence_of_element_located((By.XPATH, "//td[contains(., 'Usuario')]/input"))).send_keys(usuario_val)
            wait.until(EC.presence_of_element_located((By.XPATH, "//td[contains(., 'Contraseña')]/input"))).send_keys(contrasena_val)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit'] | //input[@type='submit']"))).click()

            # Esperar página principal
            wait.until(EC.element_to_be_clickable((By.ID, "imgAtrasMenu")))

            # Navegación Agendamiento
            wait.until(EC.element_to_be_clickable((By.ID, "imgAtrasMenu"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@title='Agendamiento']"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'Agendamiento/index.php') and contains(@title, 'Agendar WFM')]"))).click()

            # Procesar orden
            tb_orden = wait.until(EC.presence_of_element_located((By.ID, "TBorden")))
            tb_orden.clear()
            tb_orden.send_keys(orden)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", tb_orden)

            wait.until(EC.element_to_be_clickable((By.ID, "Rbot-O"))).click()
            wait.until(EC.element_to_be_clickable((By.ID, "button"))).click()

            # Pestaña Orden
            menu_container = wait.until(EC.presence_of_element_located((By.ID, "menuh-info-agendamiento")))
            ActionChains(driver).move_to_element(menu_container).perform()
            orden_tab = wait.until(EC.element_to_be_clickable((By.ID, "orden_menu")))
            try:
                orden_tab.click()
            except:
                driver.execute_script("arguments[0].click();", orden_tab)

            # Notas de la orden
            try:
                resultado["nota_orden"] = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//tr[th[contains(text(), 'Notas de la Orden')]]/td/p"))
                ).text
            except:
                resultado["nota_orden"] = "N/A"

            # Pestaña Visita
            wait.until(EC.element_to_be_clickable((By.ID, "visita_menu"))).click()
            try:
                resultado["estado_programa"] = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'Estado de la Visita')]/following-sibling::td[@class='verderesaltado']"))
                ).text.strip()
            except:
                resultado["estado_programa"] = "N/A"
            try:
                resultado["fecha_programada"] = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'Fecha Programada')]/following-sibling::td[@class='verderesaltado']"))
                ).text.strip()
            except:
                resultado["fecha_programada"] = "N/A"
            try:
                resultado["franja_programada"] = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'Franja Suscriptor')]/following-sibling::td[@class='verderesaltado']"))
                ).text.strip()
            except:
                resultado["franja_programada"] = "N/A"

            # OFSC
            wait.until(EC.element_to_be_clickable((By.ID, "ofsc_menu"))).click()
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "iframe-ofsc")))
            try:
                tbody = wait.until(EC.presence_of_element_located((By.ID, "tbodyMainList")))
                filas = tbody.find_elements(By.TAG_NAME, "tr")
                for fila in filas:
                    celdas = fila.find_elements(By.TAG_NAME, "td")
                    if len(celdas) < 7: 
                        continue
                    if celdas[5].text.strip().lower() == "pendiente":
                        boton = celdas[-1].find_element(By.CSS_SELECTOR, "button.checkDetails")
                        try:
                            boton.click()
                        except:
                            driver.execute_script("arguments[0].click();", boton)
                        try:
                            notas_text_ofsc = wait.until(EC.presence_of_element_located((By.ID, "notasorden"))).text.strip()
                            try:
                                data_json = json.loads(notas_text_ofsc)
                                if data_json.get("error"):
                                    resultado["nota_ofsc"] = "SIN DATOS"
                                else:
                                    resultado["nota_ofsc"] = notas_text_ofsc
                            except:
                                resultado["nota_ofsc"] = notas_text_ofsc
                        except:
                            resultado["nota_ofsc"] = "N/A"
            except:
                resultado["nota_ofsc"] = "N/A"
            driver.switch_to.default_content()
            wait.until(EC.element_to_be_clickable((By.ID, "return-suscriptor"))).click()
            driver.quit()
        except Exception as e:
            resultado["nota_orden"] = f"ERROR: {e}"
        return resultado

    # Función principal que maneja todos los hilos
    def ejecutar_scraping(handler):
        ruta_archivo = archivo_label.value
        if not ruta_archivo:
            handler("❌ No se seleccionó ningún archivo")
            return
        df = pd.read_excel(ruta_archivo, usecols=["Orden de trabajo","Subtipo de la Orden de Trabajo"], dtype=str)
        lista_rr = df.loc[df["Subtipo de la Orden de Trabajo"]=="Entrega De Servicios FO", "Orden de trabajo"].str[:9].tolist()
        usuario_val = usuario_input.value
        contrasena_val = contrasena_input.value

        resultados_finales = []

        handler(f"🔹 Iniciando procesamiento de {len(lista_rr)} órdenes en paralelo...")

        with ThreadPoolExecutor(max_workers=3) as executor:
            futures = [executor.submit(procesar_orden, orden, usuario_val, contrasena_val) for orden in lista_rr]
            for future in futures:
                resultado = future.result()
                resultados_finales.append(resultado)
                handler(f"✅ Orden {resultado['orden']} procesada")

        # Guardar resultados
        resultado_dir = os.path.join(os.getcwd(), "RESULTADO")
        os.makedirs(resultado_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(resultado_dir, f"resultado_{timestamp}.xlsx")
        pd.DataFrame(resultados_finales).to_excel(output_path, index=False)
        handler(f"✅ Resultados guardados en: {output_path}")

    # Botones
    def iniciar_scraping(e):
        threading.Thread(target=ejecutar_scraping, args=(log_handler,), daemon=True).start()

    def abrir_resultado(e):
        resultado_dir = os.path.join(os.getcwd(), "RESULTADO")
        if os.name == "nt":
            subprocess.Popen(f'explorer "{resultado_dir}"')
        else:
            subprocess.Popen(["open", resultado_dir])

    # Layout
    page.add(
        ft.Row([
            ft.Column([
                ft.Text("Menu", weight="bold", size=18),
                usuario_input,
                contrasena_input,
                ft.ElevatedButton("Seleccionar archivo", on_click=abrir_picker),
                archivo_label,
                ft.ElevatedButton("Iniciar Scraping", on_click=iniciar_scraping),
                ft.ElevatedButton("Abrir carpeta RESULTADO", on_click=abrir_resultado)
            ], width=250, scroll="auto", spacing=10),
            ft.Column([
                ft.Text("Registro de Logs", weight="bold", size=16),
                log_output
            ], expand=True)
        ], expand=True)
    )

ft.app(target=main)
