from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import secret
import time

client_list = []

class Browser:

    # Inicializando el webdriver con el path a chromedriver.exe

    def __init__(self):
        download_path = secret.carpeta_descargas
        prefs = {"download.default_directory":download_path,
                "directory_upgrade": True
                }
        c_options = Options()
        c_options.add_experimental_option("prefs", prefs)
        self.browser = webdriver.Chrome(options= c_options)

    def open_page(self, url:str):
        time.sleep(3)
        self.browser.get(url)

    def close_browser(self):
        time.sleep(1)
        self.browser.close()
    
    def add_input(self, by, value:str, text:str):
        field = self.browser.find_element(by=by, value=value)
        field.send_keys(text)
        time.sleep(1)

    def click_button(self, by, value:str):
        time.sleep(1)
        wait = WebDriverWait(self.browser, 10)
        element = wait.until(EC.element_to_be_clickable((by, value)))

        # Hace clic en el elemento
        element.click()
        time.sleep(1)
        
    def login_LuxPower(self, user:str, password:str):
        time.sleep(1)
        self.add_input(by=By.ID, value="account", text=user)
        self.add_input(by=By.ID, value="password", text=password)
        self.click_button(by=By.CSS_SELECTOR, value="button.btn.btn-lg.btn-success")
    
    def station_client(self):

        time.sleep(1)
        dropdown_list_btn = self.browser.find_element(by=By.XPATH, value="/html/body/div[1]/div[1]/div[2]/span/span/a")
        self.browser.execute_script("arguments[0].click();", dropdown_list_btn)
        time.sleep(1)
        station_names_tbody = self.browser.find_element(by=By.XPATH, value="/html/body/div[2]/div/div/div/div[1]/div[2]/div[2]/table/tbody")
        elem_stat_names = station_names_tbody.find_elements(By.TAG_NAME, "tr")
        for name in elem_stat_names:
            celdas = name.find_elements(By.TAG_NAME, "td")
            texto_celdas = [celda.text for celda in celdas][0]
            station_name = texto_celdas.split("-")[0].strip()
            client_list.append(station_name)
        
        print(f"Lista de Clientes: {client_list}")
            
    def download_data_tables(self):
        for index, client in enumerate(client_list):
            time.sleep(1)
            dropdown_list_btn = self.browser.find_element(by=By.XPATH, value="/html/body/div[1]/div[1]/div[2]/span/span/a")
            self.browser.execute_script("arguments[0].click();", dropdown_list_btn)
            time.sleep(1)
            station_id = "datagrid-row-r1-2-" + str(index)
            station_element = self.browser.find_element(by=By.ID, value=station_id)
            time.sleep(1)
            self.browser.execute_script("arguments[0].click();", station_element)
            time.sleep(1)
            button = self.browser.find_element(By.ID, "exportMoreDataButton")
            self.browser.execute_script("arguments[0].click();", button)
            time.sleep(18)
            secret.mover_archivos_a_subcarpeta(carpeta_origen=secret.carpeta_descargas, nombre_subcarpeta=client)
            print(f"Terminada la descarga con el cliente: {client}")
            
        



def DataDownload():
    browser = Browser()

    browser.open_page("https://na.luxpowertek.com/WManage/web/login")

    browser.login_LuxPower(secret.user, secret.password)

    browser.open_page("https://na.luxpowertek.com/WManage/web/analyze/data")
    
    browser.station_client()

    browser.download_data_tables()

    print("Completado el proceso de descargas. Saliendo...")

    browser.close_browser()

