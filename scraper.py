# scraper.py
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import time
import tempfile
import shutil
from datetime import datetime
import os

# ----------------------------------------------------
# FUNCIONES SEGURAS
# ----------------------------------------------------
def safe_find(driver, by, value, timeout=12):
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
    except:
        return None


def safe_click(driver, element):
    if not element:
        return False
    try:
        element.click()
    except:
        try:
            driver.execute_script("arguments[0].click();", element)
        except:
            return False
    return True


def safe_get_text(element):
    try:
        return element.text.strip()
    except:
        return "N/A"


# ----------------------------------------------------
# 🔥 MÉTODO ROBUSTO PARA REGRESAR AL FORMULARIO
# ----------------------------------------------------
def regresar_formulario(driver, handler):
    """
    Intenta volver al formulario principal de consulta.
    Primero intenta usar el botón 'return-suscriptor'.
    Si falla o no navega correctamente, usa driver.back().
    Valida que TBorden esté disponible antes de continuar.
    """

    # 1️⃣ Intento con botón normal
    btn = safe_find(driver, By.ID, "return-suscriptor", timeout=4)
    if btn:
        safe_click(driver, btn)
        time.sleep(1.2)

        # 2️⃣ Verificar si realmente regresó
        if safe_find(driver, By.ID, "TBorden", timeout=5):
            return True
        else:
            handler("⚠ Regresar con botón falló → usando driver.back()…")

    else:
        handler("⚠ Botón 'return-suscriptor' NO encontrado → usando driver.back()…")

    # 3️⃣ Forzar regreso con BACK
    try:
        driver.back()
        time.sleep(1.5)

        # 4️⃣ Validar regreso
        if safe_find(driver, By.ID, "TBorden", timeout=6):
            return True
        else:
            handler("❌ No regresó al formulario. La página puede estar caída.")
            return False

    except:
        handler("❌ Error inesperado ejecutando driver.back()")
        return False


# ----------------------------------------------------
# DETECTAR Y CERRAR POPUP
# ----------------------------------------------------
def cerrar_popup_rr(driver, handler=None):
    try:
        popup = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class,'jqmWindow')]")
            )
        )

        boton = popup.find_element(By.XPATH, ".//button[contains(text(),'Aceptar')]")
        boton.click()

        if handler:
            handler("⚠ Popup 'Orden cerrada en RR' detectado y cerrado.")

        return True

    except:
        return False


# ----------------------------------------------------
# ANTI-BLOQUEO + REINTENTO DE CARGA
# ----------------------------------------------------
def cargar_url_segura(driver, url, handler, intentos=10):
    for i in range(intentos):
        driver.get(url)
        time.sleep(2)

        if "502 Bad Gateway" in driver.page_source:
            handler(f"⚠ Página caída (502). Reintentando... {i+1}/{intentos}")
            time.sleep(3)
            continue

        return True

    handler("❌ La página no respondió después de varios intentos.")
    return False


# ----------------------------------------------------
# INICIAR NAVEGADOR (Modo visible + Anti-detección)
# ----------------------------------------------------
def iniciar_navegador(temp_profile):
    chrome_options = Options()

    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument(f"--user-data-dir={temp_profile}")

    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=chrome_options)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    driver.set_page_load_timeout(40)

    return driver


# ----------------------------------------------------
# FUNCIÓN PRINCIPAL
# ----------------------------------------------------
def ejecutar_scraping(usuario, password, ruta_archivo, handler):
    driver = None
    temp_profile = tempfile.mkdtemp()

    try:
        handler("🔹 Iniciando navegador en primer plano...")
        driver = iniciar_navegador(temp_profile)
        wait = WebDriverWait(driver, 40)

        # =====================================================
        # 🔁 LOGIN CON REINTENTOS AUTOMÁTICOS
        # =====================================================
        MAX_REINTENTOS = 5
        intentos = 0

        while intentos < MAX_REINTENTOS:
            intentos += 1
            handler(f"🔹 Intento de login #{intentos}...")

            if not cargar_url_segura(
                driver,
                "https://moduloagenda.cable.net.co/index.php",
                handler
            ):
                return

            try:
                user_field = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//td[contains(.,'Usuario')]/input")
                ))
                pass_field = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//td[contains(.,'Contraseña')]/input")
                ))

                user_field.clear()
                pass_field.clear()
                user_field.send_keys(usuario)
                pass_field.send_keys(password)

                boton_login = wait.until(
                    EC.element_to_be_clickable((By.ID, "Submit"))
                )
                safe_click(driver, boton_login)
                handler("🔹 Login enviado... verificando acceso...")

                time.sleep(2)

                # ¿Sigue apareciendo el campo usuario? → login fallido
                still_login = safe_find(driver,
                    By.XPATH, "//td[contains(.,'Usuario')]/input", timeout=3)

                if still_login:
                    handler("⚠ La página volvió al login. Reintentando...")
                    continue

                # ¿Ya estamos dentro? → aparece el menú
                menu_btn = safe_find(driver, By.ID, "imgAtrasMenu", timeout=10)
                if menu_btn:
                    handler("🔹 Login realizado correctamente ✔")
                    break

            except Exception as e:
                handler(f"⚠ Error durante login: {e}")
                continue

        else:
            handler("❌ No fue posible iniciar sesión después de varios intentos.")
            return

        # =====================================================
        # NAVEGACIÓN AL MÓDULO
        # =====================================================
        handler("🔹 Cargando menú principal...")

        safe_click(driver, safe_find(driver, By.ID, "imgAtrasMenu", timeout=20))
        safe_click(driver, safe_find(driver, By.XPATH, "//a[@title='Agendamiento']"))
        safe_click(driver, safe_find(driver, By.XPATH,
            "//a[contains(@href,'Agendamiento') and contains(@title,'Agendar')]"
        ))

        handler("🔹 Navegación completa ✔")

        # =====================================================
        # LEER ARCHIVO
        # =====================================================
        df = pd.read_excel(ruta_archivo)

        lista_rr = [
            x[:9] for x, y in zip(
                df["Orden de trabajo"].astype(str),
                df["Subtipo de la Orden de Trabajo"].astype(str)
            ) if y == "Entrega De Servicios FO"
        ]

        resultados = {
            "RR": [],
            "Notas Orden": [],
            "Estado Programa": [],
            "Fecha Programa": [],
            "Franja Programa": [],
            "Nota OFSC": [],
            "Cliente": [],
            "Dirección": []
        }

        # =====================================================
        # 🔁 PROCESAR CADA ORDEN
        # =====================================================
        for orden in lista_rr:

            handler(f"🔹 Procesando orden: {orden}")
            resultados["RR"].append(orden)

            driver.switch_to.default_content()

            tb = safe_find(driver, By.ID, "TBorden", timeout=15)
            if not tb:
                handler("⚠ TBorden no apareció, regresando al menú...")
                regresar_formulario(driver, handler)

                tb = safe_find(driver, By.ID, "TBorden", timeout=10)
                if not tb:
                    handler("❌ ERROR: No se pudo cargar TBorden. Orden saltada.")
                    continue

            tb.clear()
            tb.send_keys(orden)

            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", tb)

            safe_click(driver, safe_find(driver, By.ID, "Rbot-O"))
            safe_click(driver, safe_find(driver, By.ID, "button"))
            time.sleep(1)

            # POPUP RR cerrado
            if cerrar_popup_rr(driver, handler):
                for key in resultados.keys():
                    if key != "RR":
                        resultados[key].append("N/A")
                regresar_formulario(driver, handler)
                continue

            # Menú orden
            menu = safe_find(driver, By.ID, "menuh-info-agendamiento")
            if menu:
                ActionChains(driver).move_to_element(menu).perform()

            safe_click(driver, safe_find(driver, By.ID, "orden_menu"))

            notas = safe_find(driver, By.XPATH, "//tr[th[contains(text(),'Notas')]]/td/p")
            resultados["Notas Orden"].append(safe_get_text(notas))

            sus = safe_find(driver, By.XPATH, "//div[contains(text(),'Suscriptor')]/following-sibling::div")
            resultados["Cliente"].append(safe_get_text(sus))

            direccion = safe_find(driver, By.XPATH, "//div[contains(text(),'Dirección')]/following-sibling::div")
            resultados["Dirección"].append(safe_get_text(direccion))

            estado = safe_find(driver, By.ID, "textoEstadoProg")
            fecha = safe_find(driver, By.ID, "textofechaProg")
            franja = safe_find(driver, By.ID, "textofranjaProg")
            nota_ofsc = safe_find(driver, By.ID, "textoDescripServicio")

            resultados["Estado Programa"].append(safe_get_text(estado))
            resultados["Fecha Programa"].append(safe_get_text(fecha))
            resultados["Franja Programa"].append(safe_get_text(franja))
            resultados["Nota OFSC"].append(safe_get_text(nota_ofsc))

            # =====================================================
            # 🔙 REGRESAR AL FORMULARIO
            # =====================================================
            regresar_formulario(driver, handler)

        # =====================================================
        # GUARDAR RESULTADOS
        # =====================================================
        df_res = pd.DataFrame(resultados)

        nombre = f"resultado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        carpeta = os.path.join(os.getcwd(), "RESULTADO")
        os.makedirs(carpeta, exist_ok=True)
        ruta_salida = os.path.join(carpeta, nombre)

        df_res.to_excel(ruta_salida, index=False)

        handler(f"✅ Resultados guardados en: {ruta_salida}")

    except Exception as e:
        handler(f"❌ Error crítico: {e}")

    finally:
        if driver:
            driver.quit()
        shutil.rmtree(temp_profile, ignore_errors=True)
        handler("✔ Proceso finalizado.")
