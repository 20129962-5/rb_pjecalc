from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
import pandas as pd
import xlrd
import time
import os
import gc
#
# from Calculo.pjecalc_dados_calculo import DadosCalculo
from Tools.pjecalc_control import Control


class Faltas:


    def __init__(self, source):
        self.source = source
        self.delay = 10
        self.delayG = 1.5
        self.var_type_int = 1
        self.objTools = Control()
        self.planilha_base_faltas = pd.read_excel(self.source, sheet_name='PJeFaltas', header=0)


    def verificar_conteudo_vazio_para_faltas(self):

        coluna_01 = self.planilha_base_faltas["INICIO"]
        contagem = 0
        # Verificar conteúdo.
        for i in coluna_01:
            if i != "Excluir Linha":
                contagem += 1
        # Contagem de valores diferentes de 'Excluir Linha'
        if contagem >= 1:
            print(f"- Há valores para Faltas | - {contagem} registros encontrados.")
            return True
        else:
            print(f"- Não há valores para Faltas | - {contagem} registros encontrados.")
            return False

    def acessar_faltas(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageFaltas"))).click()
        # WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:j_id46:0:j_id49:7:j_id54'))).click()

    def verificacao(self, driver):

        # script = "alert('O escopo do cálculo é superior ao limite da planilha base (30 anos). Irei ignorar esta operação.')"

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write('- ' + campo + ' | ' + status + '\n')
            return file_txt_log.close()

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Faltas', 'Ok')
            else:
                driver.execute_script("alert('Algum erro ocorreu! Favor, verifique se os valores estão corretos. Irei dar continuidade.')")
                WebDriverWait(driver, self.delay).until(EC.alert_is_present())
                alerta = Alert(driver)
                time.sleep(5)
                try:
                    alerta.accept()
                except NoAlertPresentException:
                    pass
                # print('* ERRO!', msg)
                gerar_relatorio('Faltas', '---------- Erro! ----------')
        except TimeoutException:
            print('# - [Except][Faltas] - Elemento não encontrado/A Página demorou no carregamento. Encerrando...')

        # Tempo de controle
        time.sleep(2)

    def gerar_arquivo_faltas(self):

        faltas = self.planilha_base_faltas.dropna(how='all')
        faltas_2 = faltas.drop(faltas[faltas['INICIO'] == 'Excluir Linha'].index)

        for i in range(len(faltas_2)):
            # 1- Percorrer todos os valores  das colunas 'INICIO' e 'FIM' e adicionar os atribuí-los a variável 'data_inicial'
            data_inicial = faltas_2['INICIO'][i]
            data_final = faltas_2['FIM'][i]

            if type(data_inicial) == type(self.var_type_int) and type(data_final) == type(self.var_type_int):
                # 2- Converter a data do formato 'int excel' para datetime
                data_inicial_convertida = xlrd.xldate_as_datetime(data_inicial, 0)
                data_final_convertida = xlrd.xldate_as_datetime(data_final, 0)
                # 3- Converter os valores de datetime para date
                data_inicial = data_inicial_convertida.date()
                data_final = data_final_convertida.date()
                faltas_2['INICIO'][i] = data_inicial.strftime('%d/%m/%Y')
                faltas_2['FIM'][i] = data_final.strftime('%d/%m/%Y')
            else:
                pass

        # GERAR CSV
        faltas_2.to_csv("faltas.csv", sep=';', index=False)

    def adicionar_arquivo_csv(self, driver):

        # Origem do arquivo
        source_file = os.getcwd() + '\\faltas.csv'
        # Aguardar o arquivo ser gerado
        ct = 0
        while not os.path.exists(source_file):
            print('...', end='')
            time.sleep(1)
            ct += 1
            if ct == 10:
                print('- Arquivo não existe!')
                exit()
        print('- Arquivo "Faltas.csv" disponível.')
        # Campo para anexar
        campo_escolher_arquivo = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:arquivo:file')))
        campo_escolher_arquivo.send_keys(source_file)
        # Aguardar Processamento
        self.objTools.aguardar_carregamento(driver)

    def confirmar_operacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:confirmarImportacao'))).click()

    def main_faltas(self, driver):

        # Verificar se há valores para Faltas
        checagem = self.verificar_conteudo_vazio_para_faltas()
        print("- Retorno da Verificação de Conteúdo: ", checagem)
        if not checagem:
            print('-- Fim - (Faltas) --')
        else:
            self.acessar_faltas(driver)
            self.objTools.aguardar_carregamento(driver)
            # Tempo de controle
            time.sleep(self.delayG)
            self.gerar_arquivo_faltas()
            # Tempo de controle
            time.sleep(self.delayG)
            self.adicionar_arquivo_csv(driver)
            # Tempo de controle
            time.sleep(self.delayG)
            self.confirmar_operacao(driver)
            self.objTools.aguardar_carregamento(driver)
            # Tempo de controle
            time.sleep(self.delayG)
            self.verificacao(driver)
            # Tempo de controle
            time.sleep(self.delayG)

            # - Limpar Temp
            self.objTools.limparFilesTemp()
            gc.collect(generation=0)
            gc.collect(generation=1)
            gc.collect(generation=2)

            print('-- Fim - (Faltas) --')