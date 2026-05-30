from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from gspread_formatting import *
from datetime import datetime, timedelta

import pandas as pd
import glob
import os
import time
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# =========================
# FECHAS
# =========================

hoy = datetime.today()
desde = hoy - timedelta(days=7)

fecha_desde = desde.strftime("%d/%m/%Y")
fecha_hasta = hoy.strftime("%d/%m/%Y")

# =========================
# ABRIR CHROME
# =========================

# =========================
# CARPETA DESCARGAS
# =========================

download_dir = os.path.join(
    os.getcwd(),
    "downloads"
)

os.makedirs(
    download_dir,
    exist_ok=True
)

# =========================
# CHROME
# =========================

options = webdriver.ChromeOptions()

options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True
}

options.add_experimental_option(
    "prefs",
    prefs
)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)
# =========================
# LOGIN
# =========================

driver.get("https://human-log.com/BPlogin.php")

time.sleep(2)

usuario = driver.find_element(By.NAME, "usu")
password = driver.find_element(By.NAME, "pass")

usuario.send_keys(
    os.getenv("HUMANLOG_USER", "GMP")
)

password.send_keys(
    os.getenv("HUMANLOG_PASS", "GMP")
)

boton_login = driver.find_element(By.ID, "registrar")
boton_login.click()

time.sleep(3)

# =========================
# ENTRAR A CONSULTAS
# =========================

driver.get("https://human-log.com/BPcon.php")

time.sleep(3)

# =========================
# FECHA DESDE
# =========================

campo_desde = driver.find_element(By.NAME, "fechadesde")
campo_desde.clear()
campo_desde.send_keys(fecha_desde)

# =========================
# FECHA HASTA
# =========================

campo_hasta = driver.find_element(By.NAME, "fechahasta")
campo_hasta.clear()
campo_hasta.send_keys(fecha_hasta)

time.sleep(1)

# =========================
# ACTUALIZA CONSULTA
# =========================

boton_actualizar = driver.find_element(
    By.XPATH,
    "//input[@value='Actualiza Consulta']"
)

boton_actualizar.click()

time.sleep(5)

# =========================
# GENERAR CSV
# =========================

boton_csv = driver.find_element(By.ID, "btnGeneraExcel")
boton_csv.click()

print("CSV descargado 😄")

time.sleep(5)

# =========================
# BUSCAR ULTIMO CSV
# =========================

downloads_path = os.path.join(
    download_dir,
    "*.csv"
)

archivos = glob.glob(downloads_path)

ultimo_csv = max(archivos, key=os.path.getctime)

print("CSV encontrado:", ultimo_csv)

# =========================
# LEER CSV
# =========================

df = pd.read_csv(
    ultimo_csv,
    sep=";",
    encoding="latin1"
)

# =========================
# TOMAR COLUMNAS NECESARIAS
# =========================

df = df[
    [
        "Pedido",
        "Destinatario",
        "Provincia Destino",
        "Numero de Guia"
    ]
]

# =========================
# RENOMBRAR COLUMNAS
# =========================

df.columns = [
    "pedido",
    "destinatario",
    "provincia destino",
    "guia"
]
df["fecha carga"] = datetime.now().strftime("%Y-%m-%d")
# =========================
# LIMPIAR PEDIDOS
# =========================

df["pedido"] = (
    df["pedido"]
    .astype(str)
    .str.replace("'", "")
    .str.strip()
)

# =========================
# DEDUPLICAR
# =========================

df.drop_duplicates(
    subset=["pedido", "guia"],
    inplace=True
)

# =========================
# REEMPLAZAR NaN
# =========================

df = df.fillna("")

# =========================
# ELIMINAR REGISTROS INVALIDOS
# =========================

df = df[
    (df["pedido"] != "") &
    (df["pedido"] != "0") &
    (df["guia"] != "")
]

# =========================
# NORMALIZAR CAMPOS
# =========================

df["pedido"] = df["pedido"].astype(str).str.strip()
df["guia"] = df["guia"].astype(str).str.strip()
# =========================
# MOSTRAR RESULTADO
# =========================

print(df.head())

# =========================
# GOOGLE SHEETS
# =========================

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

if os.getenv("GOOGLE_CREDS"):

    google_creds = json.loads(
        os.environ["GOOGLE_CREDS"]
    )

    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        google_creds,
        scope
    )

else:

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        "credenciales.json",
        scope
    )
client = gspread.authorize(creds)

sheet = client.open_by_url(
    "https://docs.google.com/spreadsheets/d/1qE4-tQ6BfCPkkHmAUQZSy9pwAkBa4w3BlNPsVdCixTw/edit?usp=sharing"
).sheet1
# =========================
# LEER HISTORICO
# =========================

try:
    historico = pd.DataFrame(sheet.get_all_records())
    print(f"Historico encontrado: {len(historico)} registros")
except:
    historico = pd.DataFrame()
    print("No existe historico")

# =========================
# UNIR HISTORICO + NUEVO
# =========================

print("COLUMNAS HISTORICO:")
print(historico.columns.tolist())

if len(historico) > 0:

    # normalizar nombres de columnas
    historico.columns = [
        c.lower().strip()
        for c in historico.columns
    ]

    print("COLUMNAS NORMALIZADAS:")
    print(historico.columns.tolist())

    historico["pedido"] = historico["pedido"].astype(str).str.strip()
    historico["guia"] = historico["guia"].astype(str).str.strip()

    df = pd.concat(
        [historico, df],
        ignore_index=True
    )
# =========================
# DEDUPLICAR FINAL
# =========================

df["pedido"] = df["pedido"].astype(str).str.strip()
df["guia"] = df["guia"].astype(str).str.strip()

df.drop_duplicates(
    subset=["pedido", "guia"],
    keep="first",
    inplace=True
)

print(f"Registros finales: {len(df)}")  
# =========================
# LIMPIAR SHEET
# =========================

sheet.clear()

# =========================
# LIMPIAR NaN FINAL
# =========================

df = df.fillna("")
df = df.astype(str)
df = df.replace("nan", "")

# =========================
# SUBIR DATOS
# =========================

datos = [[
    "PEDIDO",
    "NOMBRE DESTINATARIO",
    "PROVINCIA DESTINO",
    "GUIA",
    "FECHA CARGA"
]] + df.values.tolist()


sheet.update(datos)

# =========================
# FORMATO ENCABEZADO
# =========================

fmt = CellFormat(
    backgroundColor=Color(0.35, 0.20, 0.60),
    textFormat=TextFormat(
        foregroundColor=Color(1, 1, 1),
        bold=True,
        fontSize=12
    ),
    horizontalAlignment="CENTER"
)

format_cell_range(
    sheet,
    "A1:E1",
    fmt
)

set_frozen(sheet, rows=1)

set_column_width(sheet, "A", 120)
set_column_width(sheet, "B", 320)
set_column_width(sheet, "C", 180)
set_column_width(sheet, "D", 220)
set_column_width(sheet, "E", 140)

print("Formato aplicado 😄")

print("Google Sheets actualizado 😄🔥")

input("ENTER para cerrar...")