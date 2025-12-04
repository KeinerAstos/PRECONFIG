# scraper.py
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def iniciar_driver():
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-features=VizDisplayCompositor")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

def login(driver, username, password):
    driver.get("https://amx-res-co.etadirect.com/")
    print("Ingresando usuario...")
    # driver.find_element(By.NAME, "username").send_keys("38100467")  # tu usuario original
    driver.find_element(By.NAME, "username").send_keys(username)
    print("Ingresando contraseña...")
    # driver.find_element(By.NAME, "password").send_keys("Cl4r0/962***")  # tu contraseña original
    driver.find_element(By.NAME, "password").send_keys(password)

    wait = WebDriverWait(driver, 30)
    print("Dando click en login...")
    btn_login = wait.until(EC.element_to_be_clickable((By.ID, "sign-in")))
    btn_login.click()

    # Segundo paso "olvidar contraseña"
    wait = WebDriverWait(driver, 30)
    print("Dar click a olvidar (JS)")
    btn_olvidar = wait.until(EC.presence_of_element_located((By.ID, "delsession")))
    driver.execute_script("arguments[0].click();", btn_olvidar)

    time.sleep(1)
    print("Ingresando contraseña nuevamente...")
    driver.find_element(By.NAME, "password").send_keys(password)
    btn_login_2 = wait.until(EC.element_to_be_clickable((By.ID, "sign-in")))
    btn_login_2.click()
    print("✔ Login exitoso")
    return driver, wait

def procesar_rr(driver, wait, excel_path="documentos/documento.xlsx"):
    df = pd.read_excel(excel_path)
    lista_rr = [
        x for x, y in zip(
            df["Orden de trabajo"].astype(str),
            df["Subtipo de la Orden de Trabajo"].astype(str)
        ) if y == "Entrega De Servicios FO"
    ]

    franja, notas, suscriptores, direcciones = [], [], [], []

    for orden in lista_rr:
        print("\n=============================")
        print("Procesando RR:", orden)

        # Esperar pantalla de carga
        try:
            WebDriverWait(driver, 20).until(
                EC.invisibility_of_element_located(
                    (By.CSS_SELECTOR, ".loading-mask, .loading-overlay, .oj-fwk-loading")
                )
            )
        except:
            pass

        # Buscar barra de búsqueda
        campo_buscar = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.search-bar-input")))
        driver.execute_script("arguments[0].focus();", campo_buscar)
        time.sleep(0.4)
        campo_buscar.click()
        campo_buscar.clear()
        time.sleep(0.5)

        # Escribir la orden
        for letra in orden:
            driver.execute_cdp_cmd("Input.insertText", {"text": letra})
            driver.execute_cdp_cmd("Input.dispatchKeyEvent", {"type": "keyUp", "key": letra})
            time.sleep(0.13)

        driver.execute_cdp_cmd("Input.dispatchKeyEvent", {"type": "keyUp", "key": " "})
        time.sleep(2)

        # Seleccionar el primer item
        try:
            primer_item = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='global-search-found-item'][1]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", primer_item)
            time.sleep(0.5)
            primer_item.click()
            time.sleep(2)

            franja_sus = wait.until(EC.presence_of_element_located((By.ID, "id_index_812")))
            franja.append(franja_sus.text)

            # Notas OT
            boton_notas = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'NOTAS OT')]")))
            boton_notas.click()
            nota_ot = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//span[contains(., 'Notas OT')]/following::div[@class='form-item'][1]")))
            notas.append(nota_ot.text)

            # Regresar
            boton_regresar = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(., 'Detalles de actividad')]/ancestor::button")))
            boton_regresar.click()

            # Suscriptor
            boton_suscriptor = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'suscriptor')]")))
            boton_suscriptor.click()
            suscriptor_elem = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//label[@aria-label='Suscriptor:']/ancestor::div[contains(@class,'form-text-field')]//div[@class='form-item']")
                )
            )
            suscriptores.append(suscriptor_elem.text)

            direccion_elem = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//label[@aria-label='Dirección Correspondencia']/ancestor::div[contains(@class,'form-text-field')]//div[@class='form-item']")
                )
            )
            direcciones.append(direccion_elem.text)

        except Exception as e:
            print("❌ Error:", e)
            franja.append("N/A")
            notas.append("N/A")
            suscriptores.append("N/A")
            direcciones.append("N/A")

    # Guardar resultado
    tabla = [{
        "orden": o,
        "franja": f,
        "notas": n,
        "suscriptor": s,
        "direccion": d
    } for o, f, n, s, d in zip(lista_rr, franja, notas, suscriptores, direcciones)]
    output_path = "RESULTADO/resultado.xlsx"
    pd.DataFrame(tabla).to_excel(output_path, index=False)
    print(f"✅ RESULTADO GUARDADO: {output_path}")
