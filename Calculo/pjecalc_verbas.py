from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import gc
#
from Tools.pjecalc_control import Control


class VerbasModel:


    def __init__(self):
        self.delay = 10
        self.delayG = 1.5
        self.objTools = Control()

    def verificar_verba_criada(self, driver):

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo} | {status}\n')
            return file_txt_log.close()

        try:
            mensagem = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Verbas: Nova verba criada', 'Ok')
            else:
                # print('* ERRO!', msg)
                gerar_relatorio('Verbas: Nova verba criada ', '---------- Erro! ----------')
        except TimeoutException:
            print('- [Except][Verbas][1] - Elemento não encontrado/A Página demorou para responder. Encerrando...')


    def verificacao(self, driver):

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo} | {status}\n')
            return file_txt_log.close()

        try:
            mensagem = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Verbas', 'Ok')
            else:
                # print('* ERRO!', msg)
                gerar_relatorio('Verbas', '---------- Erro! ----------')
        except TimeoutException:
            print('- [Except][Verbas][2] - Elemento não encontrado/A Página demorou para responder. Encerrando...')


    def entrar_verbas(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageVerba"))).click()


    def selecionar_todos_checkbox(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:listagem:selecionarTodos'))).click()


    def selecionar_sobrescrever(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoRegeracao:1'))).click()


    def click_regerar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.botao#formulario\:regerarOcorrencias'))).click()


    def click_confirmar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.botao#popup_ok'))).click()


    def selecionar_expresso(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".botao#formulario\:lancamentoExpresso"))).click()


    def selecionar_verba_principal(self, driver): # Ajuda de Custo
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:j_id82:16:j_id84:0:selecionada"))).click()


    def salvar_operacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".botao#formulario\:salvar"))).click()


    def acessar_parametros_da_verba(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:listagem:0:j_id558"))).click()


    def definir_nome(self, driver):
        campo_nome = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, "formulario:descricao")))
        campo_nome.send_keys(Keys.CONTROL, "a")
        campo_nome.send_keys("---------- AUTOMAÇÃO")


    def definir_compor_principal(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:comporPrincipal:1"))).click()


    def definir_incidencia(self, driver):

        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:inss"))).click()
        time.sleep(0.5)
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:irpf"))).click()
        time.sleep(0.5)
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:fgts"))).click()
        time.sleep(0.5)


    def definir_incidenciaINSS(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:inss"))).click()


    def definir_incidenciaIRPF(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:irpf"))).click()


    def definir_incidenciaFGTS(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:fgts"))).click()


    def definir_valor_devido(self, driver):
        field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:valorInformadoDoDevido")))
        field.click()
        field.send_keys(Keys.CONTROL, "a")
        field.send_keys("1")
        field.send_keys(Keys.TAB)
        time.sleep(0.5)


    def regerar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, "formulario:regerarOcorrencias"))).click()


    def confimar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "popup_ok"))).click()

    def main_verbas(self, driver):

        self.entrar_verbas(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.selecionar_expresso(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.selecionar_verba_principal(driver)
        # Tempo de controle
        time.sleep(0.5)
        self.salvar_operacao(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.verificar_verba_criada(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.acessar_parametros_da_verba(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.definir_nome(driver)
        # Tempo de controle
        time.sleep(0.5)
        self.definir_compor_principal(driver)
        time.sleep(0.5)
        # self.definir_incidencia(driver)
        self.definir_incidenciaINSS(driver)
        time.sleep(0.5)
        self.definir_incidenciaIRPF(driver)
        time.sleep(0.5)
        self.definir_incidenciaFGTS(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.definir_valor_devido(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.salvar_operacao(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        # self.regerar(driver)
        # Tempo de controle
        # time.sleep(0.5)
        # self.confimar(driver)
        # self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        # time.sleep(self.delayG)
        self.verificacao(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        # - Limpar Temp
        self.objTools.limparFilesTemp()