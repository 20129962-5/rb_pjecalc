from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


class PjecalcPaginaInicial:


    def __init__(self):
        self.delay = 10

    def criar_novo_calculo(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sprite-criar > a:nth-child(1)'))).click()
        time.sleep(3)

    def buscar_calculo(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sprite-abrir > a:nth-child(1)'))).click()

    def importar_calculo(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sprite-importar > a:nth-child(1)'))).click()

    def anexar_arquivo_pjc(self, driver, source_file):
        source_file = source_file.replace("/", "\\")
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:arquivo:file'))).send_keys(source_file)

    def confirmar_operacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:confirmarImportacao'))).click()

    def buscar_reclamante(self, driver, reclamante):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:reclamanteBusca'))).send_keys(reclamante)

    def buscar_btn(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:buscar'))).click()

    def abrir_calculo(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:listagem:0:j_id599'))).click()

    def clicar_paginaInicial_pjecalc(self, driver):
        try:
            field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//div[@id="logo"]//a[@accesskey="1"]')))
            field.click()
        except Exception as e:
            print(f"- [except][clicar_paginaInicial_pjecalc]: {e}")