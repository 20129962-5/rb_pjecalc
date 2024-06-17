import gc

from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
import pandas as pd
import xlrd
import time
import os
# from Calculo.pjecalc_dados_calculo import DadosCalculo
from Tools.pjecalc_control import Control


class Ferias:


    def __init__(self, source):
        self.source = source
        self.delay = 10
        self.delayG = 1.5
        self.var_type_int = 1
        self.objTools = Control()
        self.planilha_base_ferias = pd.read_excel(self.source, sheet_name='PJeFérias', header=0)


    def verificar_conteudo_vazio_para_ferias(self):

        coluna_01 = self.planilha_base_ferias["RELATIVAS"]
        contagem = 0
        # Verificar conteúdo.
        for i in coluna_01:
            if i != "Excluir Linha":
                contagem += 1
        # Contagem de valores diferentes de 'Excluir Linha'
        if contagem >= 1:
            print(f"- Há valores para Férias | - {contagem} registros encontrados.")
            return True
        else:
            print(f"- Não há valores para Férias | - {contagem} registros encontrados.")
            return False

    def acessar_ferias(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageFerias"))).click()

    def verificar_importacao(self, driver):

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write('- ' + campo + ' | ' + status + '\n')
            return file_txt_log.close()

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            if 'Operação realizada com sucesso.' in mensagem.text:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Ferias: Importação', 'Ok')
            else:
                driver.execute_script("alert('Algum erro ocorreu! Favor, verifique se os valores estão corretos. Irei dar continuidade.')")
                WebDriverWait(driver, self.delay).until(EC.alert_is_present())
                alerta = Alert(driver)
                time.sleep(5)
                try:
                    alerta.accept()
                except NoAlertPresentException:
                    pass
                # print('* ERRO!', mensagem.text)
                gerar_relatorio('Ferias: Importação', '---------- Erro! ----------')
        except TimeoutException:
            print('# - [Except][Férias][1] - Elemento não encontrado/A Página demorou no carregamento. Encerrando...')

        # Tempo de controle
        time.sleep(2)

    def verificacao(self, driver):

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
                gerar_relatorio('Ferias', 'Ok')
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
                gerar_relatorio('Ferias', '---------- Erro! ----------')

        except TimeoutException:
            print('# - [Except][Férias][2] - Elemento não encontrado/A Página demorou no carregamento. Encerrando...')

        # Tempo de controle
        time.sleep(2)

    def gerar_arquivo_ferias(self):

        ferias = self.planilha_base_ferias.dropna(how='all') # (how='all', thresh=None)
        # ferias = ferias.dropna(thresh=None)
        ferias_2 = ferias.drop(ferias[ferias['RELATIVAS'] == 'Excluir Linha'].index)

        for k in range(len(ferias_2)):
            # Percorrer as colunas G1INI e G1FIM
            coluna_g1_ini = ferias_2['G1INI'][k]
            coluna_g1_fim = ferias_2['G1FIM'][k]
            # Percorrer as colunas G2INI e G2FIM
            coluna_g2_ini = ferias_2['G2INI'][k]
            coluna_g2_fim = ferias_2['G2FIM'][k]
            # Percorrer as colunas G3INI e G3FIM
            coluna_g3_ini = ferias_2['G3INI'][k]
            coluna_g3_fim = ferias_2['G3FIM'][k]
            # Condição para pegar somente as datas do typo inteito do Excel coletadas da planilha base
            if type(coluna_g1_ini) == type(self.var_type_int) and type(coluna_g1_fim) == type(self.var_type_int):
                # 1- Convertendo as datas do tipo inteiro para o tipo datetime
                g1_ini_datatime = xlrd.xldate_as_datetime(coluna_g1_ini, 0)
                g1_fim_datatime = xlrd.xldate_as_datetime(coluna_g1_fim, 0)
                # 2- Convertendo as datas do tipo datetime para o tipo date
                g1_ini_data = g1_ini_datatime.date()
                g1_fim_data = g1_fim_datatime.date()
                # 3- Convertendo as datas do tipo date para o tipo str
                ferias_2['G1INI'][k] = g1_ini_data.strftime('%d/%m/%Y')
                ferias_2['G1FIM'][k] = g1_fim_data.strftime('%d/%m/%Y')
            else:
                pass
            # Condição para pegar somente as datas do typo inteito do Excel coletadas da planilha base
            if type(coluna_g2_ini) == type(self.var_type_int) and type(coluna_g2_fim) == type(self.var_type_int):
                # 1- Convertendo as datas do tipo inteiro para o tipo datetime
                g2_ini_datatime = xlrd.xldate_as_datetime(coluna_g2_ini, 0)
                g2_fim_datatime = xlrd.xldate_as_datetime(coluna_g2_fim, 0)
                # 2- Convertendo as datas do tipo datetime para o tipo date
                g2_ini_data = g2_ini_datatime.date()
                g2_fim_data = g2_fim_datatime.date()
                # 3- Convertendo as datas do tipo date para o tipo str
                ferias_2['G2INI'][k] = g2_ini_data.strftime('%d/%m/%Y')
                ferias_2['G2FIM'][k] = g2_fim_data.strftime('%d/%m/%Y')
            else:
                pass
            # Condição para pegar somente as datas do typo inteito do Excel coletadas da planilha base
            if type(coluna_g3_ini) == type(self.var_type_int) and type(coluna_g3_fim) == type(self.var_type_int):
                # 1- Convertendo as datas do tipo inteiro para o tipo datetime
                g3_ini_datatime = xlrd.xldate_as_datetime(coluna_g3_ini, 0)
                g3_fim_datatime = xlrd.xldate_as_datetime(coluna_g3_fim, 0)
                # 2- Convertendo as datas do tipo datetime para o tipo date
                g3_ini_data = g3_ini_datatime.date()
                g3_fim_data = g3_fim_datatime.date()
                # 3- Convertendo as datas do tipo date para o tipo str
                ferias_2['G3INI'][k] = g3_ini_data.strftime('%d/%m/%Y')
                ferias_2['G3FIM'][k] = g3_fim_data.strftime('%d/%m/%Y')
            else:
                pass
        # GERAR CSV
        ferias_2.to_csv("ferias.csv", sep=';', index=False)

    def adicionar_arquivo_csv(self, driver):

        # Origem do arquivo
        source_file = os.getcwd() + '\\ferias.csv'
        ct = 0
        while not os.path.exists(source_file):
            print('...', end='')
            time.sleep(1)
            ct += 1
            if ct == 10:
                print('- Arquivo não existe!')
                break
        print('- Arquivo "Ferias.csv" disponível.')
        # Campo para anexar
        campo_escolher_arquivo = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:arquivo:file')))
        campo_escolher_arquivo.send_keys(source_file)
        time.sleep(1.5)
        # Aguardar Processamento
        self.objTools.aguardar_carregamento(driver)

    def confirmar_operacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:j_id96'))).click()

    def salvar_operacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.NAME, 'formulario:j_id215'))).click()

    def main_ferias(self, driver):

        # Verificar se há valores para Faltas
        checagem = self.verificar_conteudo_vazio_para_ferias()
        print("- Retorno da Verificação de Conteúdo: ", checagem)
        if not checagem:
            print('-- Fim - (Férias) --')
        else:
            self.acessar_ferias(driver)
            self.objTools.aguardar_carregamento(driver)
            # Tempo de controle
            time.sleep(self.delayG)
            self.gerar_arquivo_ferias()
            # Tempo de controle
            time.sleep(self.delayG)
            self.adicionar_arquivo_csv(driver)
            # Tempo de controle
            time.sleep(self.delayG)
            self.confirmar_operacao(driver)
            self.objTools.aguardar_carregamento(driver)
            # Tempo de controle
            time.sleep(self.delayG)
            self.verificar_importacao(driver)
            self.salvar_operacao(driver)
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

            print('-- Fim - (Férias) --')