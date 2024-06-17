from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, date, timedelta
from calendar import monthrange
import time
import os
#
from Tools.pjecalc_control import Control
from Calculo.pjecalc_menu import PjecalcMenu


class Liquidar:
    
    
    def __init__(self):
        self.objTools = Control()
        self.objMenu = PjecalcMenu()
        self.delay = 10


    def verificacao(self, driver):

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            # file_txt_log.write('* ' + campo + ' | ' + status + '\n')
            file_txt_log.write(f'- {campo} | {status}\n')
            return file_txt_log.close()

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                print('* Operação realizada com sucesso.')
                gerar_relatorio('Liquidar', 'Ok')
            else:
                print('* ERRO!', msg)
                gerar_relatorio('Liquidar', '---------- Erro! ----------')
                # exit() -> Antes
                pass
        except TimeoutException:
            print(
                '* Exceção - Verificação -  A Página demorou para responder ou o elemento não foi encontrado. Encerrando...')
            exit()

        # Tempo de controle
        time.sleep(2)

    def acessar_guia_operacoes(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'li.header:nth-child(3)'))).click()

    def acessar_liquidar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageLiquidar"))).click()

    def acumular_indice(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:indicesAcumulados:2"))).click()

    def preencher_data(self, driver):

        def gerar_data_liquidacao():
            # - DATA ATUAL
            data_hoje = date.today()

            # - DIA ATUAL
            dia_verificacao = data_hoje.day

            # - CALCULAR DOIS MESES ANTES AO MÊS ATUAL
            dt_dois_meses_antes = date(data_hoje.year, data_hoje.month, data_hoje.day) - timedelta(days=60)
            # print(" # ", dt_dois_meses_antes)

            if dia_verificacao >= 16:
                dt_dois_meses_antes = dt_dois_meses_antes + timedelta(days=30)
                # print(" ## ", dt_dois_meses_antes)
                # mes += 1

            # - MÊS ATUAL
            mes = dt_dois_meses_antes.month

            # - ANO ATUAL
            ano = dt_dois_meses_antes.year

            # Função retorna uma tubla com a semana e último dia do mês
            ultimo_dia_mes = monthrange(ano, mes)

            # - ÚLTIMO DIA
            dia = ultimo_dia_mes[1]

            # - DATA LIQUIDAÇÃO
            data_liquidacao = datetime(ano, mes, dia).strftime("%d/%m/%Y")
            print(F"- DATA ATUAL: {data_hoje.strftime('%d/%m/%Y')} | - DATA HÁ DOIS MESES: {data_liquidacao}")

            return data_liquidacao

        data = gerar_data_liquidacao()

        campo_data = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:dataDeLiquidacaoInputDate')))
        campo_data.send_keys(data)

    def click_liquidar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:liquidar'))).click()

    def main_liquidar(self, driver):

        self.objMenu.acessar_operacoes(driver)
        # Tempo de controle
        time.sleep(2)
        self.acessar_liquidar(driver)
        self.objTools.aguardar_carregamento(driver)
        time.sleep(1.5)
        self.preencher_data(driver)
        # Tempo de controle
        time.sleep(1)
        self.acumular_indice(driver)
        # Tempo de controle
        time.sleep(1)
        self.click_liquidar(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(1)
        self.verificacao(driver)
        # Tempo de controle
        time.sleep(1)
        print('-- Fim - (Liquidar) --')