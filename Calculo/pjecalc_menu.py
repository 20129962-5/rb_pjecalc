from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep


class PjecalcMenu:


    def __init__(self):
        self.delay = 10
        self.delayP = 1.5


    def acessar_calculo(self, driver):
        list_elements = WebDriverWait(driver, self.delay).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dock")))
        list_elements[0].click()
        sleep(self.delayP)


    def acessar_operacoes(self, driver):
        list_elements = WebDriverWait(driver, self.delay).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dock")))
        list_elements[1].click()
        sleep(self.delayP)