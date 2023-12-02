from selenium import webdriver
from selenium.webdriver.common.by import By
import secret
import time

class Browser:

    # Inicializando el webdriver con el path a chromedriver.exe

    def __init__(self):
        self.browser = webdriver.Chrome()

    def open_page(self, url:str):
        self.browser.get(url)

    def close_browser(self):
        self.browser.close()
    
    def add_input(self, by, value:str, text:str):
        field = self.browser.find_element(by=by, value=value)
        field.send_keys(text)
        time.sleep(1)

    def click_button(self, by, value:str):
        button = self.browser.find_element(by=by, value=value)
        button.click()
        time.sleep(1)
        
    def login_LuxPower(self, user:str, password:str):
        self.add_input(by=By.ID, value="account", text=user)
        self.add_input(by=By.ID, value="password", text=password)
        self.click_button(by=By.CSS_SELECTOR, value="button.btn.btn-lg.btn-success")
    
    def data_history(self):
        button = self.browser.find_element(By.ID, "exportMoreDataButton")
        self.browser.execute_script("arguments[0].click();", button)

if __name__ == "__main__":
    browser = Browser()

    browser.open_page("https://na.luxpowertek.com/WManage/web/login")
    time.sleep(5)

    browser.login_LuxPower(secret.user, secret.password)
    time.sleep(5)

    browser.open_page("https://na.luxpowertek.com/WManage/web/analyze/data")
    time.sleep(5)

    browser.data_history()
    time.sleep(60)

    browser.close_browser()

"""
from selenium import webdriver
from selenium.webdriver.common.by import By

import secret
import time

class Browser:

    def __init__(self):
        self.browser = webdriver.Chrome()

    def open_page(self, url: str):
        self.browser.get(url)
        self.wait_for_page_load()

    def wait_for_page_load(self, timeout=10):
        WebDriverWait(self.browser, timeout).until(
            EC.presence_of_element_located((By.TAG_NAME, 'body'))
        )

    def close_browser(self):
        self.browser.quit()

    def find_element(self, by: By, value: str):
        return self.browser.find_element(by=by, value=value)

    def add_input(self, by: By, value: str, text: str):
        field = self.find_element(by, value)
        field.send_keys(text)
        time.sleep(1)

    def click_button(self, by: By, value: str):
        button = self.find_element(by, value)
        button.click()
        time.sleep(1)

    def login_LuxPower(self, user: str, password: str):
        self.add_input(by=By.ID, value="account", text=user)
        self.add_input(by=By.ID, value="password", text=password)
        self.click_button(by=By.CSS_SELECTOR, value="button.btn.btn-lg.btn-success")

    def data_history(self):
        button = self.find_element(By.ID, "exportMoreDataButton")
        self.browser.execute_script("arguments[0].click();", button)

if __name__ == "__main__":
    try:
        browser = Browser()

        browser.open_page("https://na.luxpowertek.com/WManage/web/login")

        browser.login_LuxPower(secret.user, secret.password)

        browser.open_page("https://na.luxpowertek.com/WManage/web/analyze/data")

        browser.data_history()

    finally:
        browser.close_browser()

"""
