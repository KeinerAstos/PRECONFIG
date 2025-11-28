from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
import time

# Configuración del navegador
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Descomenta si quieres modo headless
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 20)

# Abrir portal
print("Abriendo portal...")
driver.get("https://moduloagenda.cable.net.co/index.php")

# Ingresar usuario
print("Ingresando usuario...")
usuario_input = wait.until(
    EC.presence_of_element_located((By.XPATH, "//td[contains(., 'Usuario')]/input"))
)
usuario_input.send_keys("46231030")

# Ingresar contraseña
print("Ingresando contraseña...")
contrasena_input = wait.until(
    EC.presence_of_element_located((By.XPATH, "//td[contains(., 'Contraseña')]/input"))
)
contrasena_input.send_keys("Crami1*1*=")

time.sleep(2)

# Click en el botón de login
print("Click en INGRESAR...")
btn_login = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//button[@type='submit'] | //input[@type='submit']"))
)
btn_login.click()

# Esperar que cargue la página principal
time.sleep(3)

# Click en la foto/logo para desplegar menú
print("Click en logo para desplegar menú...")
logo_click = wait.until(
    EC.element_to_be_clickable((By.ID, "imgAtrasMenu"))
)
logo_click.click()

# Click en Agendamiento
print("Click en Agendamiento...")
agendamiento = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//a[@title='Agendamiento']"))
)
agendamiento.click()

# Click en Agendar WFM
print("Click en Agendar WFM...")
agendar_wfm = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'Agendamiento/index.php') and contains(@title, 'Agendar WFM')]"))
)
agendar_wfm.click()

# Lista de órdenes a procesar
# Ejemplo de arreglo
# Lista de órdenes a procesar
ordenes = ["459530118","459530118"]  # Ejemplo de arreglo

for orden in ordenes:
    # Ingresar número de orden
    print(f"Ingresando orden: {orden}")
    tb_orden = wait.until(
        EC.presence_of_element_located((By.ID, "TBorden"))
    )
    tb_orden.clear()
    tb_orden.send_keys(orden)

    # Disparar evento onchange
    driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", tb_orden)

    # Seleccionar radio "Orden de Trabajo"
    orden_trabajo = wait.until(
        EC.element_to_be_clickable((By.ID, "Rbot-O"))
    )
    orden_trabajo.click()

    # Click en Consultar
    btn_consultar = wait.until(
        EC.element_to_be_clickable((By.ID, "button"))
    )
    btn_consultar.click()

    # Esperar procesamiento
    print("Esperando a que la página procese la orden...")
    time.sleep(2)

    # Hover sobre el menú y click en pestaña Orden
    print("Intentando click en pestaña Orden...")
    menu_container = wait.until(
        EC.presence_of_element_located((By.ID, "menuh-info-agendamiento"))
    )
    ActionChains(driver).move_to_element(menu_container).perform()

    orden_tab = wait.until(
        EC.element_to_be_clickable((By.ID, "orden_menu"))
    )
    try:
        orden_tab.click()
    except:
        driver.execute_script("arguments[0].click();", orden_tab)

    # Esperar carga
    time.sleep(1)

    # Extraer contenido Notas
    try:
        notas_element = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//tr[th[contains(text(), 'Notas de la Orden')]]/td/p")
            )
        )
        notas_text = notas_element.text
        print(f"Notas de la orden {orden}: {notas_text}")
    except:
        print(f"No se encontraron notas para la orden {orden}")
    # -------------------------------------------------------------------
    # 🔥 NUEVA PARTE: Buscar fila Pendiente y dar click en Ver Detalle
    # -------------------------------------------------------------------
    print("Buscando filas con estado Pendiente...")

    ofsc_tab = wait.until(
        EC.element_to_be_clickable((By.ID,"ofsc_menu"))
    )
    try:
        ofsc_tab.click()
    except:
        driver.execute_script("arguments[0].click();", ofsc_tab)
    time.sleep(1)

    print("Cambiando al iframe 'iframe-ofsc'...")
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "iframe-ofsc")))
    print("Dentro del iframe correctamente.")

    try:
        tbody = wait.until(
            EC.presence_of_element_located((By.ID, "tbodyMainList"))
        )

        filas = tbody.find_elements(By.TAG_NAME, "tr")

        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")

            if len(celdas) < 7:
                continue

            estado = celdas[5].text.strip()
            print(" - Estado encontrado:", estado)

            if estado.lower() == "pendiente":
                print("   ✔ Estado es Pendiente → haciendo click en Ver detalle")

                boton = celdas[-1].find_element(By.CSS_SELECTOR, "button.checkDetails")

                try:
                    boton.click()
                    time.sleep(1)

                except:
                    driver.execute_script("arguments[0].click();", boton)

                # -------------------------------------------------------------------
                # ⭐ NUEVA PARTE: Capturar el texto de <span id="notasorden">
                # -------------------------------------------------------------------
                try:
                    print("Esperando contenido de notas de la orden...")

                    notas = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.ID, "notasorden"))
                    )

                    print("📌 Notas de la orden:", notas.text)

                except Exception as e:
                    print("❌ No se pudo extraer notas de la orden:", e)

                # Si solo quieres procesar la primera orden pendiente, descomenta:
                # break
            
    except Exception as e:
        print("Error al procesar la tabla:", e)
    print("Saliendo del iframe para continuar...")
    driver.switch_to.default_content()

    regresar_tab = wait.until(
        EC.element_to_be_clickable((By.ID,"return-suscriptor"))
    )
    try:
        regresar_tab.click()
    except:
        driver.execute_script("arguments[0].click();", regresar_tab)

    time.sleep(1)
    print(f"Procesada orden: {orden}")

print("Todas las órdenes procesadas correctamente.")
driver.quit()
