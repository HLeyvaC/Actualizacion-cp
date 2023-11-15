import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time

# Obtiene la ruta del directorio actual del proyecto
project_directory = os.path.dirname(os.path.abspath(__file__))
# Configura las opciones del navegador Chrome para las descargas
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
    'download.default_directory': project_directory,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': True
})

# Inicializa el controlador del navegador con las opciones configuradas
driver = webdriver.Chrome(options = options)

# Abre la página web
driver.get("https://www.correosdemexico.gob.mx/SSLServicios/ConsultaCP/CodigoPostal_Exportar.aspx")
driver.maximize_window()
time.sleep(3)

# Localiza el selector de estados y sus opciones
select = driver.find_element(By.XPATH, "//*[@id='cboEdo']")
opciones = select.find_elements(By.TAG_NAME, "Option")
time.sleep(3)
seleccionar = Select(select)
seleccionar.select_by_value("26")
time.sleep(1)

# Localiza el botón de descarga y haz clic en él
boton = driver.find_element(By.XPATH, "//*[@id='btnDescarga']")
boton.click()
time.sleep(30)

if os.path.exists("Sonora.xls"):
# Cierra el navegador cuando hayas terminado
    driver.quit()
else:
    time.sleep(15)
if os.path.exists("Sonora.xls"):
    driver.quit()
else:
    time.sleep(15)
    driver.quit()