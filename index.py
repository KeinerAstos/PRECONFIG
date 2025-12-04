import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

chrome_options = Options()
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_argument("--disable-features=VizDisplayCompositor")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

driver.get("https://amx-res-co.etadirect.com/")
print("Ingresando usuario...")
driver.find_element(By.NAME, "username").send_keys("38100467")
print("Ingresando contraseña...")
driver.find_element(By.NAME, "password").send_keys("Cl4r0/962***")


wait = WebDriverWait(driver, 30)

print("Dando click en login...")
btn_login = wait.until(EC.element_to_be_clickable((By.ID, "sign-in")))
btn_login.click()

wait = WebDriverWait(driver, 30)

print("Dar click a olvidar (JS)")
btn_olvidar = wait.until(EC.presence_of_element_located((By.ID, "delsession")))
driver.execute_script("arguments[0].click();", btn_olvidar)

wait = WebDriverWait(driver, 30)    

print("Ingresando contraseña...")
driver.find_element(By.NAME, "password").send_keys("Cl4r0/962***")
wait = WebDriverWait(driver, 30)  

print("Dando click en login...")
btn_login_2 = wait.until(EC.element_to_be_clickable((By.ID, "sign-in")))
btn_login_2.click()

print("entramos")

time.sleep(4.2)
df = pd.read_excel("documentos/documento.xlsx")

lista_rr = [
    x for x, y in zip(
        df["Orden de trabajo"].astype(str),
        df["Subtipo de la Orden de Trabajo"].astype(str)
    ) if y == "Entrega De Servicios FO"
]

franja,notas,suscriptores,direcciones = [],[],[],[]

for orden in lista_rr:

    print("\n=============================")
    print("Procesando RR:", orden)

    # --- Esperar pantalla de carga si existe ---
    try:
        WebDriverWait(driver, 20).until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".loading-mask, .loading-overlay, .oj-fwk-loading")
            )
        )
    except:
        print("⚠ No se pudo detectar pantalla de carga inicial, continuando...")

    time.sleep(1)

    print("Abriendo barra de búsqueda...")

    campo_buscar = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input.search-bar-input"))
    )

    # Focus real
    driver.execute_script("arguments[0].focus();", campo_buscar)
    time.sleep(0.4)

    # Limpiar el campo como humano (CTRL+A + backspace vía CDP)
    campo_buscar.click()
    time.sleep(0.2)
    driver.execute_cdp_cmd("Input.dispatchKeyEvent", {"type": "keyDown", "key": "Control"})
    driver.execute_cdp_cmd("Input.dispatchKeyEvent", {"type": "keyDown", "key": "a"})
    driver.execute_cdp_cmd("Input.dispatchKeyEvent", {"type": "keyUp", "key": "a"})
    driver.execute_cdp_cmd("Input.dispatchKeyEvent", {"type": "keyUp", "key": "Control"})
    time.sleep(0.15)
    driver.execute_cdp_cmd("Input.dispatchKeyEvent", {"type": "keyDown", "key": "Backspace"})
    driver.execute_cdp_cmd("Input.dispatchKeyEvent", {"type": "keyUp", "key": "Backspace"})
    time.sleep(0.5)

    print("Escribiendo RR como humano (CDP tecla por tecla)...")

    # --- ESCRIBIR LETRA POR LETRA ---
    for letra in orden:
        driver.execute_cdp_cmd("Input.insertText", {"text": letra})
        driver.execute_cdp_cmd("Input.dispatchKeyEvent", {
            "type": "keyUp",
            "key": letra
        })
        time.sleep(0.13)  # ritmo humano

    print("✔ Texto escrito. Disparando keyup final para activar sugerencias...")

    driver.execute_cdp_cmd("Input.dispatchKeyEvent", {
        "type": "keyUp",
        "key": " "
    })

    time.sleep(2)

    # Intentar seleccionar primera sugerencia
    try:
        sugerencia = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, ".oj-listview-focused-element, .oj-listview-cell-element")
            )
        )
        sugerencia.click()
        print("✔ Sugerencia seleccionada.")
    except:
        print("⚠ No apareció sugerencia.")

    print("Esperando resultados...")
    try:
        WebDriverWait(driver, 30).until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".loading-mask, .loading-overlay, .oj-fwk-loading")
            )
        )
    except:
        pass

    time.sleep(1.5)
    print("🔍 Buscando el item AMARILLO con XPath...")

    try:
    # Esperar el ícono amarillo con color #FFFF26
        
        # Buscar íconos amarillos o naranjas
        # icono = WebDriverWait(driver, 12).until(
        #     EC.presence_of_element_located(
        #         (By.XPATH, "//div[contains(@class,'activity-icon') and (contains(@style,'#FFFF26') or contains(@style,'#FFAC63'))]")
        #     )
        # )
        # print("✔ Icono encontrado (amarillo o naranja). Subiendo al contenedor padre...")

        # # Subir al contenedor clickable global-search-found-item
        # contenedor = icono.find_element(By.XPATH, "./ancestor::div[contains(@class,'global-search-found-item')]")
        # driver.execute_script("arguments[0].scrollIntoView(true);", contenedor)
        # time.sleep(0.5)

        # contenedor.click()
        # print("✔ Click realizado sobre el item seleccionado.")

        # Seleccionar el primer item de la lista sin importar el color
        primer_item = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@class='global-search-found-item'][1]")
            )
        )

        # Subir al contenedor y hacer scroll
        driver.execute_script("arguments[0].scrollIntoView(true);", primer_item)
        time.sleep(0.5)

        # Click en el primer item
        primer_item.click()
        print("✔ Click realizado sobre el primer item de la lista.")
        time.sleep(2)

        franja_sus = wait.until(EC.presence_of_element_located((By.ID, "id_index_812")))
        franja.append(franja_sus.text)

        # -------- BOTÓN NOTAS OT --------
        
        boton_notas = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(., 'NOTAS OT')]")
            )
        )
        boton_notas.click()

        nota_ot = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//span[contains(., 'Notas OT')]/following::div[@class='form-item'][1]")
            )
        )
        notas.append(nota_ot.text)

        # -------- REGRESAR --------
        boton_regresar = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(., 'Detalles de actividad')]/ancestor::button")
            )
        )
        boton_regresar.click()

        # -------- ABRIR SUSCRIPTOR --------

        boton_suscriptor = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'suscriptor')]")
            )
        )
        boton_suscriptor.click()

        suscriptor_elem = wait.until(
        EC.presence_of_element_located(
                (By.XPATH, "//label[@aria-label='Suscriptor:']/ancestor::div[contains(@class,'form-text-field')]//div[@class='form-item']")
            )
        )

        suscriptor = suscriptor_elem.text
        suscriptores.append(suscriptor)
        direccion_elem = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//label[@aria-label='Dirección Correspondencia']/ancestor::div[contains(@class,'form-text-field')]//div[@class='form-item']")
            )
        )

        direccion = direccion_elem.text
        direcciones.append(direccion)
        print("Dirección:", direccion)



    except Exception as e:
        print("❌ Error:", e)
        franja.append("N/A")
        notas.append("N/A")
        suscriptores.append("N/A")
        direcciones.append("N/A")

tabla = [{
                "orden": o,
                "franja": f,
                "notas":n,
                "suscriptor": s,
                "direccion": d
            } for o,f,n,s,d in zip(
                lista_rr,franja,notas, suscriptores,direcciones 
            )]
output_path = "RESULTADO/resultado.xlsx"
pd.DataFrame(tabla).to_excel(output_path, index=False)

print(f"✅ RESULTADO GUARDADO: ")
