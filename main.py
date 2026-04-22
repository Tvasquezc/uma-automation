# ========================
# IMPORTS
# ========================
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import time
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ========================
# CONFIG
# ========================
download_path = os.path.join(os.getcwd(), "descargas")

if not os.path.exists(download_path):
    os.makedirs(download_path)

ruta_credenciales = "credenciales.json"
nombre_sheet = "Reporte UMA Bajaj"

# ⚠️ CAMBIA ESTO POR TU COLUMNA REAL ÚNICA
columna_id = None  # ejemplo: "ID" o "NumeroPedido"

# ========================
# FUNCIONES
# ========================

def limpiar_descargas():
    for f in os.listdir(download_path):
        os.remove(os.path.join(download_path, f))


def esperar_descarga(archivos_antes, timeout=180):
    inicio = time.time()

    while time.time() - inicio < timeout:
        archivos_despues = set(os.listdir(download_path))
        nuevos = archivos_despues - archivos_antes
        nuevos = [f for f in nuevos if not f.endswith(".crdownload")]

        if nuevos:
            archivo = nuevos[0]
            ruta = os.path.join(download_path, archivo)
            print("📥 Nuevo archivo:", archivo)
            return ruta

        time.sleep(1)

    raise Exception("⛔ Timeout descarga")


def renombrar_archivo(ruta_archivo, trimestre):
    destino = os.path.join(download_path, f"ventas_Q{trimestre}.xlsx")

    if os.path.exists(destino):
        os.remove(destino)

    os.rename(ruta_archivo, destino)
    print(f"📁 Guardado como: ventas_Q{trimestre}.xlsx")


def set_fecha(input_element, valor):
    driver.execute_script("arguments[0].value = arguments[1];", input_element, valor)
    driver.execute_script("""
        arguments[0].dispatchEvent(new Event('input',{bubbles:true}));
        arguments[0].dispatchEvent(new Event('change',{bubbles:true}));
        arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));
    """, input_element)


def combinar_excels():
    archivos = [f"ventas_Q{i}.xlsx" for i in range(1,5)]
    dfs = []

    for archivo in archivos:
        ruta = os.path.join(download_path, archivo)
        if os.path.exists(ruta):
            dfs.append(pd.read_excel(ruta))

    if not dfs:
        print("⚠️ No hay archivos")
        return None

    df_final = pd.concat(dfs, ignore_index=True).drop_duplicates()

    ruta_final = os.path.join(download_path, "ventas_anual.xlsx")
    df_final.to_excel(ruta_final, index=False)

    print("📊 ventas_anual.xlsx creado")
    return ruta_final


# ========================
# 🔥 SUBIR SOLO NUEVOS
# ========================
def subir_solo_nuevos(ruta_excel):
    if not ruta_excel or not os.path.exists(ruta_excel):
        print("❌ Excel no encontrado")
        return

    df_nuevo = pd.read_excel(ruta_excel)

    # evitar NaN solo para envío
    df_nuevo = df_nuevo.astype(object).where(pd.notnull(df_nuevo), "")

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        ruta_credenciales, scope
    )

    client = gspread.authorize(creds)
    sheet = client.open(nombre_sheet).sheet1

    data_actual = sheet.get_all_values()

    # 🔹 si está vacío
    if not data_actual:
        print("📄 Sheet vacío → subiendo todo")
        sheet.update([df_nuevo.columns.tolist()] + df_nuevo.values.tolist())
        return

    df_actual = pd.DataFrame(data_actual[1:], columns=data_actual[0])

    # 🔥 detectar columna ID automáticamente si no defines
    global columna_id
    if columna_id is None:
        columna_id = df_nuevo.columns[0]
        print(f"⚠️ Usando columna ID automática: {columna_id}")

    ids_existentes = set(df_actual[columna_id])

    df_filtrado = df_nuevo[~df_nuevo[columna_id].astype(str).isin(ids_existentes)]

    if df_filtrado.empty:
        print("✅ No hay datos nuevos")
        return

    print(f"➕ Filas nuevas: {len(df_filtrado)}")

    fila_inicio = len(data_actual) + 1

    data_subir = df_filtrado.values.tolist()

    # 🔥 subir solo nuevos
    sheet.update(range_name=f"A{fila_inicio}", values=data_subir)

    # actualizar fecha
    fecha = datetime.now().strftime("%d de %B de %Y")
    sheet.update_acell("Z1", f"Actualizado: {fecha}")

    print("☁️ Nuevos datos agregados")


# ========================
# DRIVER
# ========================
options = webdriver.ChromeOptions()

# 🔥 PARA GITHUB ACTIONS (HEADLESS)
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_setting_values.automatic_downloads": 1
})

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

driver.maximize_window()
wait = WebDriverWait(driver, 20)

# ========================
# LOGIN
# ========================
USER = "susana.vasquez@grupouma.com"
PASSWORD = "UmaOne.123"

driver.get("https://umaone.grupouma.services/login")

wait.until(EC.presence_of_element_located((
    By.CSS_SELECTOR, "input[formcontrolname='userName']"
))).send_keys(USER)

driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='password']").send_keys(PASSWORD)

login_btn = wait.until(EC.element_to_be_clickable((
    By.XPATH, "//button[@type='submit']"
)))

driver.execute_script("arguments[0].click();", login_btn)
wait.until(EC.presence_of_element_located((By.ID, "side-menu")))

print("✅ Login OK")

# ========================
# NAVEGACIÓN
# ========================
ventas = wait.until(EC.presence_of_element_located((
    By.XPATH, "//a[contains(@href,'reporteVentas')]"
)))

driver.execute_script("arguments[0].click();", ventas)
wait.until(EC.url_contains("reporteVentas"))

# ========================
# INPUTS
# ========================
inputs = wait.until(EC.presence_of_all_elements_located((
    By.XPATH, "//input[contains(@class,'form-control')]"
)))

fecha_inicio_input = inputs[0]
fecha_fin_input = inputs[1]

buscar_btn = wait.until(EC.element_to_be_clickable((
    By.XPATH, "//button[.//span[contains(text(),'Buscar')]]"
)))

time.sleep(3)

# ========================
# DESCARGAS
# ========================
fecha_base = datetime(2026, 1, 1)

limpiar_descargas()

for i in range(4):

    fecha_inicio = fecha_base + relativedelta(months=3*i)
    fecha_fin = fecha_inicio + relativedelta(months=3, days=-1)

    inicio_str = fecha_inicio.strftime("%Y-%m-%d")
    fin_str = fecha_fin.strftime("%Y-%m-%d")

    print(f"📅 Descargando: {inicio_str} → {fin_str}")

    set_fecha(fecha_inicio_input, inicio_str)
    time.sleep(0.5)

    set_fecha(fecha_fin_input, fin_str)
    time.sleep(0.5)

    time.sleep(2)

    archivos_antes = set(os.listdir(download_path))

    driver.execute_script("arguments[0].click();", buscar_btn)

    time.sleep(3)

    archivo = esperar_descarga(archivos_antes)
    renombrar_archivo(archivo, i + 1)

    print(f"✅ Q{i+1} correcto")

# ========================
# FINAL
# ========================
driver.quit()
print("🌐 Navegador cerrado")

ruta_excel = combinar_excels()

# 🔥 SOLO SUBE NUEVOS
subir_solo_nuevos(ruta_excel)

print("🎯 PROCESO COMPLETO OK")