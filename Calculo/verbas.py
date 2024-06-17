import math
import gc
import os
import glob
import xlrd
import shutil
import numpy as np
import pandas as pd
from time import sleep
from datetime import datetime
from Tools.pjecalc_control import Control
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from pandas._libs.tslibs.timestamps import Timestamp
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException



class Verbas:


    def __init__(self, source):
        self.sourceFileExcel = source
        self.delay = 10
        self.delayP = 1
        self.hrinicialLog = ""
        self.hrfinalLog = ""
        self.objTime = Control()
        self.planilha_base = pd.read_excel(source, sheet_name="TBVERBABD")
        self.planilha_basecalculo_exp_hs = pd.read_excel(source, sheet_name="DEVIDOBD_EXP")
        self.planilha_basecalculovalorpago_exp_hs = pd.read_excel(source, sheet_name="PAGOBD_EXP")


    # --------------------- PJeCalc - Página Inicial --------------------- #
    def importarCalculo(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.sprite-importar > a:nth-child(1)'))).click()

    def esolherArquivo(self, driver):
        arquivo_base = os.getcwd() + r"\automacao_verbas.PJC"
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:arquivo:file"))).send_keys(arquivo_base)

    def confirmarOperacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:confirmarImportacao"))).click()

    # --------------------- PJeCalc - [Dados do Cálculo] - Pegar Nome do Reclamante --------------------- #
    def get_nome_reclamente(self, driver):
        value_reclamente = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:reclamanteNome")))
        nome_reclamente = value_reclamente.get_attribute("textContent")
        print("- Nome do Reclamente: ", nome_reclamente)
        return nome_reclamente

    # --------------------- PJeCalc - [Dados do Cálculo] - Salvar --------------------- #
    def salvar_operacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".botao#formulario\:salvar"))).click()

    # --------------------- Planilha Base - [Reclamante/Numero do Processo] - Salvar --------------------- #
    def coletar_dados_plan(self, source):

        planilha_base_dados = pd.read_excel(source, sheet_name="PJE-BD", header=1)

        myDados = []

        for value in range(1, len(planilha_base_dados))[:10]:

            key = planilha_base_dados.loc[value, 'IDENTIFICADOR']
            valor = planilha_base_dados.loc[value, 'INFORMACAO']

            if key == 'numero_processo':
                num_processo = valor
                myDados.append(num_processo)
            if key == 'nome_parte1':
                reclamente = valor
                myDados.append(reclamente)

        return myDados


    def coletar_dados_planilha_verbas(self, source):

        planilha_verbas = pd.read_excel(source, sheet_name="TBCALCULO", header=0)

        myDados = []

        snmreclamante = planilha_verbas.loc[0, 'SNMRECLAMANTE']
        snmnumeroprocesso = planilha_verbas.loc[0, 'SNMNUMEROPROCESSO']
        zidprocesso = planilha_verbas.loc[0, 'ZIDPROCESSO']

        myDados.append(snmreclamante)
        myDados.append(snmnumeroprocesso)
        myDados.append(zidprocesso)

        return myDados


    def coletar_dados_plan_new(self, source):

        planilha_base_dados = pd.read_excel(source, sheet_name="PJE-BD", header=1)

        myDados = []

        for value in range(1, len(planilha_base_dados))[:10]:

            key = planilha_base_dados.loc[value, 'IDENTIFICADOR']
            valor = planilha_base_dados.loc[value, 'INFORMACAO']

            if key == 'numero_processo':
                num_processo = valor
                myDados.append(num_processo)
            if key == 'nome_parte1':
                reclamente = valor
                myDados.append(reclamente)

        return myDados

    # --------------------- PJeCalc - Verificação --------------------- #

    def verificar_verba_criada(self, driver):

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo} | {status}\n')
            return file_txt_log.close()

        try:
            mensagem = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id77')))
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
            mensagem = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69'))).text
            if 'Operação realizada com sucesso.' in mensagem:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Verbas', 'Ok')
            else:
                # print('* ERRO!', mensagem)
                gerar_relatorio('Verbas', '---------- Erro! ----------')
        except TimeoutException:
            print('- [Except][Verbas][2] - Elemento não encontrado/A Página demorou para responder. Encerrando...')


    def mensagem_alert_frontend(self, driver, conteudo):
        driver.execute_script(f"alert('{conteudo}')")
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alerta = Alert(driver)
        sleep(3)
        try:
            alerta.accept()
        except NoAlertPresentException:
            pass

    def verificacao_2(self, driver):

        def gerar_relatorio(campo, msg, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo}: {msg} | {status}\n')
            return file_txt_log.close()

        # formulario:painelMensagens:j_id4593
        try:
            barraMensagem = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "formulario:painelMensagens:j_id69")))
            alertaMensagem = barraMensagem.get_attribute("textContent")
            print("- 01. ", alertaMensagem)
            if 'sucesso' in alertaMensagem:
                print("- Alerta de verificação: ", alertaMensagem)
                gerar_relatorio("Verbas", "--", "Ok")
            else:
                gerar_relatorio("Verbas", f"{alertaMensagem}", "---------- Erro! ----------")
                self.mensagem_alert_frontend(driver, alertaMensagem)
                self.click_BtnCancelar(driver)
                self.objTime.aguardar_carregamento(driver)
                return 'erro'

        except TimeoutException:

            barraMensagem = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.CLASS_NAME, "erro")))
            alertaMensagem = barraMensagem.get_attribute("textContent")
            print("- 02. ", alertaMensagem)
            elementos_error = WebDriverWait(driver, 2).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, 'linkErro')))
            for erro in elementos_error:
                erroUtil = erro.get_attribute("textContent")
                erroUtil = erroUtil.split("//<![")
                print(erroUtil[0])
                # Pop-up na tela 01
                self.mensagem_alert_frontend(driver, erroUtil[0])
                # Registrar Log de erro
                gerar_relatorio("Verbas", erroUtil[0], "---------- Erro! ----------")

            self.click_BtnCancelar(driver)
            self.objTime.aguardar_carregamento(driver)

        except:

            barraMensagem = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "formulario:painelMensagens:j_id77")))
            alertaMensagem = barraMensagem.get_attribute("textContent")
            print("- 02. ", alertaMensagem)
            if 'Erro' in alertaMensagem:
                print("- Alerta de verificação: ", alertaMensagem)
                gerar_relatorio("Verbas", f"{alertaMensagem}", "---------- Erro! ----------")

        sleep(2)

    def verificacao_2_new(self, driver):

        def gerar_relatorio(campo, msg, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo}: {msg} | {status}\n')
            return file_txt_log.close()

        # formulario:painelMensagens:j_id4593

        try:
            elemento_caixa_mensagem = WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.ID, "divMensagem")))
            alertaMensagem = elemento_caixa_mensagem.get_attribute("textContent")
            alertaMensagem = alertaMensagem.split("//<![CDATA[")

            if 'sucesso' in alertaMensagem[0]:
                # print("- Alerta de verificação: ", alertaMensagem[0])
                gerar_relatorio("Verbas", "--", "Ok")
            elif 'erro' in alertaMensagem[0]:
                gerar_relatorio("Verbas", f"{alertaMensagem[0]}", "---------- Erro! ----------")
                self.mensagem_alert_frontend(driver, alertaMensagem[0])
                self.click_BtnCancelar(driver)
                self.objTime.aguardar_carregamento(driver)
                return 'erro'
            else:
                alertaMensagem = "Algum erro ocorreu!"
                gerar_relatorio("Verbas", f"{alertaMensagem}", "---------- Erro! ----------")
                self.mensagem_alert_frontend(driver, alertaMensagem)
                self.click_BtnCancelar(driver)
                self.objTime.aguardar_carregamento(driver)
                return 'erro'

        except TimeoutException as e:
            print("- [e1][Elemento não encontrado!]: ", e)

        sleep(2)



    def verificacao_2_new_parcelaReflexa(self, driver):

        def gerar_relatorio(campo, msg, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo}: {msg} | {status}\n')
            return file_txt_log.close()

        try:
            elemento_caixa_mensagem = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID, "divMensagem")))
            alertaMensagem = elemento_caixa_mensagem.get_attribute("textContent")
            alertaMensagem = alertaMensagem.split("//<![CDATA[")

            if 'sucesso' in alertaMensagem[0]:
                # print("- Alerta de verificação: ", alertaMensagem[0])
                gerar_relatorio("Verbas > Parcela Reflexa", "--", "Ok")
            elif 'erro' in alertaMensagem[0]:
                gerar_relatorio("Verbas > Parcela Reflexa", f"{alertaMensagem[0]}", "---------- Erro! ----------")
                self.mensagem_alert_frontend(driver, alertaMensagem[0])
                self.click_BtnCancelar(driver)
                self.objTime.aguardar_carregamento(driver)
                return 'erro'
            else:
                alertaMensagem = "Algum erro ocorreu!"
                gerar_relatorio("Verbas > Parcela Reflexa", f"{alertaMensagem}", "---------- Erro! ----------")
                self.mensagem_alert_frontend(driver, alertaMensagem)
                self.click_BtnCancelar(driver)
                self.objTime.aguardar_carregamento(driver)
                return 'erro'

        except TimeoutException as e:
            print("- [e1][Elemento não encontrado!]: ", e)

        sleep(2)

    # --------------------- PJeCalc - Verbas - Página Inicial --------------------- #

    def acessarVerbas(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageVerba"))).click()

    # --------------------- PJeCalc - [Verbas] - Pegar Número do Processo --------------------- #
    def get_num_processo(self, driver):
        value_processo = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "panelVerbaCalculo")))
        num_processo = value_processo.get_attribute("textContent")
        print("- Nº do Processo: ", num_processo)
        return num_processo

    def acessarExpresso(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:lancamentoExpresso"))).click()

    def selecionarVerbaPrincipal(self, driver, valor):

        # PRÊMIO PRODUÇÃO

        indice = 0
        statusWhile = True

        try:
            while statusWhile:
                listagem_verbas = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, f'formulario:j_id82:{indice}:j_id84:0:nome')))
                nome_verba = listagem_verbas.get_attribute("textContent")
                if valor == nome_verba:
                    WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, f'formulario:j_id82:{indice}:j_id84:0:selecionada'))).click()
                    statusWhile = False

                indice += 1
        except TimeoutException:

            indice = 0
            statusWhile = True
            try:
                while statusWhile:

                    listagem_verbas2 = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, f'formulario:j_id82:{indice}:j_id84:1:nome')))
                    nome_verba2 = listagem_verbas2.get_attribute("textContent")
                    if valor == nome_verba2:
                        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, f'formulario:j_id82:{indice}:j_id84:1:selecionada'))).click()
                        statusWhile = False

                    indice += 1
            except TimeoutException:

                indice = 0
                statusWhile = True

                try:
                    while statusWhile:

                        listagem_verbas3 = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, f'formulario:j_id82:{indice}:j_id84:2:nome')))
                        nome_verba3 = listagem_verbas3.get_attribute("textContent")

                        if valor == nome_verba3:
                            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, f'formulario:j_id82:{indice}:j_id84:2:selecionada'))).click()
                            statusWhile = False

                        indice += 1
                except TimeoutException:
                    exit()

        del valor
        del indice

    def salvarOperacao(self, driver):

        btn_salvar = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, "//*[@id='formulario:salvar']")))
        valor_btn = btn_salvar.get_attribute("value")
        if "Salvar" in valor_btn:
            btn_salvar.click()

    def acessarParametrosVerba(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:listagem:0:j_id558"))).click()

    def acessarParametrosVerbaPreechida(self, driver, valor):


        statusWhile = True
        indice = 0

        try:
            while statusWhile:
                coluna_verba = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, f'formulario:listagem:{str(indice)}:j_id569'))).text

                if valor == coluna_verba:
                    WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, f'formulario:listagem:{str(indice)}:j_id566'))).click()
                    statusWhile = False

                indice += 1
        except TimeoutException:
            pass

        del valor
        del indice

    # --------------------- PJECALC - VERBAS - DADOS DA VERBA --------------------- #

    # STPPARCELA
    def marcar_stpparcela(self, driver, valor):

        if valor == 'F':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:tipoVariacaoDaParcela:0"))).click()
        elif valor == 'V':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:tipoVariacaoDaParcela:1"))).click()

        del valor

    # STPVALOR
    def marcar_stpvalor(self, driver, valor):

        if valor == 'C':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:valor:0"))).click()
        elif valor == 'I':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:valor:1"))).click()

        del valor

    # SNMDESCRICAOVERBA
    def modificar_snmdescricaoverba(self, driver, nome):

        campo_nome = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, "formulario:descricao")))
        campo_nome.send_keys(Keys.CONTROL, "a")
        campo_nome.send_keys(nome)

        del nome

    # SFLINCIDENCIAIRPF
    def marcar_sflincidenciairpf(self, driver, valor):

        elementoCheckbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:irpf")))
        situacaoCheckbox = elementoCheckbox.is_selected()

        if valor == 'S':
            if not situacaoCheckbox:
                elementoCheckbox.click()
        elif valor == 'N':
            if situacaoCheckbox:
                elementoCheckbox.click()

        del valor

    # SFLINCIDENCIAFGTS
    def marcar_sflincidenciafgts(self, driver, valor):

        elementoCheckbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:fgts")))
        situacaoCheckbox = elementoCheckbox.is_selected()

        if valor == 'S':
            if not situacaoCheckbox:
                elementoCheckbox.click()

        elif valor == 'N':
            if situacaoCheckbox:
                elementoCheckbox.click()

        del valor

    # SFLINCIDENCIAINSS
    def marcar_sflincidenciainss(self, driver, valor):

        elementoCheckbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:inss")))
        situacaoCheckbox = elementoCheckbox.is_selected()

        if valor == 'S':
            if not situacaoCheckbox:
                elementoCheckbox.click()
        elif valor == 'N':
            if situacaoCheckbox:
                elementoCheckbox.click()

        del valor

    # SFLINCIDENCIAPREVPRIVADA
    def marcar_sflincidenciaprevprivada(self, driver, valor):

        elementoCheckbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:previdenciaPrivada")))
        situacaoCheckbox = elementoCheckbox.is_selected()

        if valor == 'S':
            if not situacaoCheckbox:
                elementoCheckbox.click()
        elif valor == 'N':
            if situacaoCheckbox:
                elementoCheckbox.click()

        del valor

    # SFLINCIDENCIAPENSAO
    def marcar_sflincidenciapensao(self, driver, valor):

        elementoCheckbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:pensaoAlimenticia")))
        situacaoCheckbox = elementoCheckbox.is_selected()

        if valor == 'S':
            if not situacaoCheckbox:
                elementoCheckbox.click()
        elif valor == 'N':
            if situacaoCheckbox:
                elementoCheckbox.click()

        del valor

    # SFLCOMPORPRINCIPAL
    def marcar_sflcomporprincipal(self, driver, valor):

        if valor == 'S':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:comporPrincipal:0"))).click()
        else:
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:comporPrincipal:1"))).click()

        del valor

    # SFLPADRAOZERARVALORNEGATIVO
    def marcar_sflpadraozerarvalornegativo(self, driver, valor):

        elementoCheckbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:zeraValorNegativo")))
        situacaoCheckbox = elementoCheckbox.is_selected()

        if valor == 'S':
            if not situacaoCheckbox:
                elementoCheckbox.click()
        elif valor == 'N':
            if situacaoCheckbox:
                elementoCheckbox.click()

        del valor

    # -- STPCARACTERISTICAVERBA
    def marcar_stpcaracteristicaverba(self, driver, valor):

        # - GABARITO
        # C: Comum
        # DT: Décimo Terceiro
        # AP: Aviso Prévio
        # F: Férias

        if valor == 'C':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:caracteristicaVerba:0'))).click()
        elif valor == 'DT':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:caracteristicaVerba:1'))).click()
        elif valor == 'AP':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:caracteristicaVerba:2'))).click()
        elif valor == 'F':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:caracteristicaVerba:3'))).click()

        del valor

    # -- STPOCORRENCIAPAGAMENTO
    def marcar_stpocorrenciapagamento(self, driver, valor):

        # - GABARITO
        # DL: Desligamento
        # DZ: Dezembro
        # M: Mensal
        # PA: Período Aquisitivo

        elemento = WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.NAME, 'formulario:ocorrenciaPagto')))
        if elemento.is_enabled():

            if valor == 'DL':
                WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:ocorrenciaPagto:0'))).click()
            elif valor == 'DZ':
                WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:ocorrenciaPagto:1'))).click()
            elif valor == 'M':
                WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:ocorrenciaPagto:2'))).click()
            elif valor == 'PA':
                WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:ocorrenciaPagto:3'))).click()

        del valor

    # -- STPJUROSSUMULA
    def marcar_stpjurossumula(self, driver, valor):

        if valor == 'S':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:ocorrenciaAjuizamento:0'))).click()
        else:
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:ocorrenciaAjuizamento:1'))).click()

        del valor

    # -- STPTIPOVERBA
    def marcar_stptipoverba(self, driver, valor):

        try:

            if WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.NAME, 'formulario:tipoDeVerba'))):
                elemento = WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.NAME, 'formulario:tipoDeVerba')))
                if elemento.is_enabled():
                    if valor == 'P':
                        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeVerba:0'))).click()
                    elif valor == 'R':
                        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeVerba:1'))).click()

        except TimeoutException:
            pass


        del valor

    # -- STPGERARVERBAREFLEXA
    def marcar_stpgerarverbareflexa(self, driver, valor):

        if valor == 'DV':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:geraReflexo:0'))).click()
        elif valor == 'DF':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:geraReflexo:1'))).click()

        del valor

    # -- STPGERARVERBAPRINCIPAL
    def marcar_stpgerarverbaprincipal(self, driver, valor):

        if valor == 'DV':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:gerarPrincipal:0'))).click()
        elif valor == 'DF':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:gerarPrincipal:1'))).click()

        del valor


    # --------------------- PJeCalc - Verbas - Dados de Verba - (Incompletas) --------------------- #

    def selecionarAssuntoCNJ(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:linkModalAssunto'))).click()

    def definirAssuntoCNJ(self, driver, valor):

        # assunto = valor.split(" ")
        # id_assunto = assunto[0]
        # print("- Assunto: ", assunto, "- Id Assunto: ", id_assunto)

        try:
            WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, f'formularioModalCNJ:arv:864:{valor}::_defaultNodeFaceOutput'))).click()
        except TimeoutException:
            # 864 - Direito do Trabalho -- formularioModalCNJ:arv:864::_defaultNodeFaceOutput
            WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, 'formularioModalCNJ:arv:864::_defaultNodeFaceOutput'))).click()

    def btnSelecionarAssuntoCNJ(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'btnSelecionarCNJ'))).click()


    # --------------------- PJECALC - VERBAS - VALOR DEVIDO - PERÍODO --------------------- #
    # DDTINICIO
    def modificar_ddtinicio(self, driver, valor):

        print(f"- V: {valor} - T: {type(valor)}")

        if type(valor) == pd._libs.tslibs.timestamps.Timestamp:
            print("- TIPO IDENTIFICADO!!")
            valor = valor.strftime('%d/%m/%Y')
            print("- DATA FORMATADA -- ", valor, type(valor))

        if type(valor) == int:
            valor = xlrd.xldate_as_datetime(valor, 0)
            valor = valor.strftime('%d/%m/%Y')
            print("- DT: ", valor)

        elif type(valor) == np.int64:
            valor = int(valor)
            valor = xlrd.xldate_as_datetime(valor, 0)
            valor = valor.strftime('%d/%m/%Y')
            print("- DT2: ", valor)

        elif type(valor) == datetime:
            valor = valor.strftime('%d/%m/%Y')
            print(f"- Data: {valor} - Tipo: {type(valor)} #")

        else:
            print(f"- Data: {valor} - Tipo: {type(valor)} #")

        campoPeriodoInicial = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:periodoInicialInputDate")))
        campoPeriodoInicial.click()
        campoPeriodoInicial.send_keys(valor)

        del valor

    # DDTTERMINO
    def modificar_ddttermino(self, driver, valor):

        if type(valor) == pd._libs.tslibs.timestamps.Timestamp:
            # print("- TIPO IDENTIFICADO!!")
            valor = valor.strftime('%d/%m/%Y')
            print("- DATA FORMATADA -- ", valor, type(valor))

        if type(valor) == int:
            valor = xlrd.xldate_as_datetime(valor, 0)
            valor = valor.strftime('%d/%m/%Y')
        elif type(valor) == np.int64:
            valor = int(valor)
            valor = xlrd.xldate_as_datetime(valor, 0)
            valor = valor.strftime('%d/%m/%Y')
            print("- DT2: ", valor)
        elif type(valor) == datetime:
            valor = valor.strftime('%d/%m/%Y')
            print(f"- Data: {valor} - Tipo: {type(valor)} #")
        else:
            print(f"- Data: {valor} - Tipo: {type(valor)} #")

        campoPeriodoFinal = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:periodoFinalInputDate")))
        campoPeriodoFinal.click()
        campoPeriodoFinal.send_keys(valor)

        del valor

    # --------------------- PJECALC - VERBAS - VALOR DEVIDO - EXCLUSÕES --------------------- #

    # SFLEXCLUIRFALTAJUSTIFICADA
    def marcar_sflexcluirfaltajustificada(self, driver, valor):

        try:
            elementoCheckbox = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "formulario:excluirFaltaJustificada")))
            situacaoCheckbox = elementoCheckbox.is_selected()
            if valor == 'S':
                if not situacaoCheckbox:
                    elementoCheckbox.click()
            elif valor == 'N':
                if situacaoCheckbox:
                    elementoCheckbox.click()
        except TimeoutException as e:
            print(f"- [FNJ]:  {e}")

        del valor

    # SFLEXCLUIRFALTANAOJUSTIFICADA
    def marcar_sflexcluirfaltanaojustificada(self, driver, valor):

        try:
            elementoCheckbox = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "formulario:excluirFaltaNaoJustificada")))
            situacaoCheckbox = elementoCheckbox.is_selected()
            if valor == 'S':
                if not situacaoCheckbox:
                    elementoCheckbox.click()
            elif valor == 'N':
                if situacaoCheckbox:
                    elementoCheckbox.click()
        except TimeoutException as e:
            print(f"- [EFNJ]: {e}")

        del valor

    # SFLEXCLUIRFERIASGOZADAS
    def marcar_sflexcluirferiasgozadas(self, driver, valor):

        try:
            elementoCheckbox = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "formulario:excluirFeriasGozadas")))
            situacaoCheckbox = elementoCheckbox.is_selected()
            if valor == 'S':
                if not situacaoCheckbox:
                    elementoCheckbox.click()
            elif valor == 'N':
                if situacaoCheckbox:
                    elementoCheckbox.click()
        except TimeoutException as e:
            print(f"- [EFEG]: {e}")

        del valor

    # SFLDOBRARVALOR
    def marcar_sfldobrarvalor(self, driver, valor):

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "formulario:dobraValorDevido"))):
                elementoCheckbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:dobraValorDevido")))
                situacaoCheckbox = elementoCheckbox.is_selected()
                if valor == 'S':
                    if not situacaoCheckbox:
                        elementoCheckbox.click()
                elif valor == 'N':
                    if situacaoCheckbox:
                        elementoCheckbox.click()

        except TimeoutException as e:
            print(f"- [DV]: {e}")

        del valor

    # RVLVALORDEVIDO
    def modificar_rvlvalordevido(self, driver, valor):

        try:

            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "formulario:valorInformadoDoDevido"))):
                val = f"{valor:.2f}"
                WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:valorInformadoDoDevido'))).send_keys(val)

        except TimeoutException:
            pass


        del valor

    # SFLPROPORCIONALIDADE - BASE DE CÁLCULO
    def marcar_sflproporcionalidade_basecalculo(self, driver, valor):

        print("- APLICAR PROPORCIONALIDADE A BASE: ", valor)
        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "formulario:aplicarProporcionalidadeABase"))):
                elementoCheckbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:aplicarProporcionalidadeABase")))
                situacaoCheckbox = elementoCheckbox.is_selected()
                if valor == 'S':
                    if not situacaoCheckbox:
                        elementoCheckbox.click()
                elif valor == 'N':
                    if situacaoCheckbox:
                        elementoCheckbox.click()

        except TimeoutException:
            print("- APLICAR PROPORCIONALIDADE A BASE: [Except]")
            pass

        del valor

    # SFLPROPORCIONALIDADE
    def marcar_sflproporcionalidade(self, driver, valor):

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "formulario:aplicarProporcionalidadeAoValorDevido"))):
                elementoCheckbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:aplicarProporcionalidadeAoValorDevido")))
                situacaoCheckbox = elementoCheckbox.is_selected()
                if valor == 'S':
                    if not situacaoCheckbox:
                        elementoCheckbox.click()
                elif valor == 'N':
                    if situacaoCheckbox:
                        elementoCheckbox.click()

        except TimeoutException:
            pass

        del valor

    # --------------------- PJECALC - VERBAS - VALOR DEVIDO - FORMULA - BASE DE CÁLCULO --------------------- #
    # STPBASEDEVIDO
    def modificar_stpbasedevido(self, driver, valor):

        # GABARITO
        # MR: Maior Remuneração
        # HS: Histórico Salarial
        # PS: Piso Salarial
        # SM: Salário Mínimo
        # VT: Vale Transporte

        campoSelecaoBases = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:tipoDaBaseTabelada')))
        selecaoBases = Select(campoSelecaoBases)

        if valor == 'MR':
            selecaoBases.select_by_visible_text('Maior Remuneração')
        elif valor == 'HS':
            selecaoBases.select_by_visible_text('Histórico Salarial')
        elif valor == 'PS':
            selecaoBases.select_by_visible_text('Piso Salarial')
        elif valor == 'SM':
            selecaoBases.select_by_visible_text('Salário Mínimo')
        elif valor == 'VT':
            selecaoBases.select_by_visible_text('Vale Transporte')


        del valor


    # ----

    def definirHistoricoSalarial(self, driver, valor):
        campoHistoricoSalarial = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:baseHistoricos')))
        selecaoHistorico = Select(campoHistoricoSalarial)
        selecaoHistorico.select_by_visible_text(valor)

        del valor

    def proporcionalizar(self, driver, valor):

        positivo = 'Sim'
        negativo = 'Não'

        campoProporcionalizar = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:proporcionalizaHistorico')))
        selecaoProporcionalizar = Select(campoProporcionalizar)

        if valor == positivo:
            selecaoProporcionalizar.select_by_visible_text(negativo)
            selecaoProporcionalizar.select_by_visible_text(valor)
        else:
            selecaoProporcionalizar.select_by_visible_text(positivo)
            selecaoProporcionalizar.select_by_visible_text(valor)

        del valor

    def adicionarBase(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:incluirBaseHistorico'))).click()

    def definirVerba(self, driver, valor):

        try:
            campoVerba = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:baseVerbaDeCalculo')))
            selecaoVerba = Select(campoVerba)
            selecaoVerba.select_by_visible_text(valor)
        except NoSuchElementException:
            print("- Verba não encontrada na lista!")

        del valor

    def integralizar(self, driver, valor):
        campoIntegralizar = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:integralizarBase')))
        selecaoIntegralizar = Select(campoIntegralizar)
        selecaoIntegralizar.select_by_visible_text(valor)

        del valor

    def adicionarVerba(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:incluirItemProp'))).click()

    # --------------------- PJECALC - VERBAS - VALOR DEVIDO - FORMULA - DIVIDOR --------------------- #

    # STPDIVISOR
    def modificar_stpdivisor(self, driver, valor):

        # GABARITO
        # IN: Informado
        # CH: Carga Horária
        # DU: Dias Úteis
        # ICP: Importada do Cartão de Ponto

        if valor == 'IN':
            try:
                WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeDivisor:0'))).click()
            except TimeoutException as e:
                print(f"- [DIVISOR INFORMADO]: {e}")
        elif valor == 'CH':
            try:
                WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeDivisor:1'))).click()
            except TimeoutException as e:
                print(f"- [DIVISOR CARGA HORÁRIA]: {e}")
        elif valor == 'DU':
            try:
                WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeDivisor:2'))).click()
            except TimeoutException as e:
                print(f"- [DIVISOR DIAS ÚTEIS]: {e}")
        elif valor == 'ICP':
            try:
                WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeDivisor:3'))).click()
            except TimeoutException as e:
                print(f"- [DIVISOR CARTÃO DE PONTO]: {e}")

        del valor

    # RVLOUTRODIVISOR
    def modificar_rvloutrodivisor(self, driver, valor):

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:outroValorDoDivisor'))):

                valor = f"{valor:.4f}".replace(".", ",")
                campoValor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:outroValorDoDivisor')))
                campoValor.click()
                campoValor.send_keys(valor)

        except TimeoutException:
            pass

        del valor

    # RVLMULTIPLICADOR
    def modificar_rvlmultiplicador(self, driver, valor):

        print("- PARÂMETRO: ", valor, type(valor))

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:outroValorDoMultiplicador'))):

                # valor = f"{valor:.8f}".replace(".", ",")
                # valor = f"{valor:.8f}"
                try:
                    valor = float(valor)
                    valor = f"{valor:.8f}"
                    valor = valor.replace(".", ",")
                    print("- ", valor)
                except ValueError:
                    valor = valor.replace(",", ".")
                    valor = float(valor)
                    valor = f"{valor:.8f}"
                    valor = valor.replace(".", ",")
                    print("- ", valor)

                print("- PARÂMETRO - ANTES DA ESCRITA: ", valor, type(valor))
                campoValor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:outroValorDoMultiplicador')))
                campoValor.click()
                campoValor.send_keys(valor)

        except TimeoutException:
            pass

        del valor

    # STPQUANTIDADE
    def marcar_stpquantidade(self, driver, valor):

        # GABARITO
        # IN: Informada
        # ICL: Importada do Calendário
        # ICP: Importada do Cartão de Ponto
        # AV: Avos

        if valor == 'IN':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDaQuantidade:0'))).click()
        elif valor == 'ICL':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDaQuantidade:1'))).click()
        elif valor == 'ICP':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDaQuantidade:2'))).click()
        elif valor == 'AV':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDaQuantidade:1'))).click()

        del valor

    # STPCALENDARIO
    def modificar_stpcalendario(self, driver, valor):

        # GABARITO
        # RF: Repousos e Feriados
        # R: Repousos
        # F: Feriados
        # DU: Dias Úteis

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.NAME, 'formulario:tipoImportadaCalendario'))):

                campoSelecao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:tipoImportadaCalendario')))
                elemento = Select(campoSelecao)

                if valor == 'RF':
                    elemento.select_by_visible_text('Repousos e Feriados/Pontos Facultativos')
                elif valor == 'R':
                    elemento.select_by_visible_text('Repousos')
                elif valor == 'F':
                    elemento.select_by_visible_text('Feriados/Pontos Facultativos')
                elif valor == 'DU':
                    elemento.select_by_visible_text('Dias Úteis')

        except TimeoutException:
            pass


        del valor

    # STPCARTAOPONTO
    # def modificar_stpcartaoponto(self, driver, valor):
    #     pass
    #     del valor

    # RVLOUTROVALORQTD
    def modificar_rvloutrovalorqtd(self, driver, valor):

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:valorInformadoDaQuantidade'))):

                try:
                    valor = float(valor)
                    valor = f"{valor:.4f}"
                    valor = valor.replace(".", ",")
                    print("- ", valor)
                except ValueError:
                    valor = valor.replace(",", ".")
                    valor = float(valor)
                    valor = f"{valor:.4f}"
                    valor = valor.replace(".", ",")
                    print("- ", valor)

                # valor = f"{valor:.4f}".replace(".", ",")
                campoValor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:valorInformadoDaQuantidade')))
                if campoValor.is_enabled():
                    campoValor.click()
                    campoValor.send_keys(valor)

        except TimeoutException:
            pass

        del valor

    # SFLQTDPROPORCIONALIDADE
    def marcar_sflqtdproporcionalidade(self, driver, valor):

        try:

            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "formulario:aplicarProporcionalidadeAQuantidade"))):
                elementoCheckbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:aplicarProporcionalidadeAQuantidade")))
                situacaoCheckbox = elementoCheckbox.is_selected()
                EC.presence_of_element_located((By.TAG_NAME, 'H1'))
                if valor == 'S':
                    if not situacaoCheckbox:
                        elementoCheckbox.click()
                elif valor == 'N':
                    if situacaoCheckbox:
                        elementoCheckbox.click()

        except TimeoutException:
            pass

        del valor


    # --------------------- PJECALC - VERBAS - VALOR PAGO --------------------- #

    # STPVALORPAGO
    def marcar_stpvalorpago(self, driver, valor):

        # GABARITO
        # I: Informado
        # C: Calculado

        # acaoDuploClique = ActionChains(driver)

        xpath_info = "//input[@id='formulario:tipoDoValorPago:0' and @value='INFORMADO']"
        xpath_calc = "//input[@id='formulario:tipoDoValorPago:1' and @value='CALCULADO']"

        if valor == 'I':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.XPATH, f'{xpath_info}'))).click()
        elif valor == 'C':
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.XPATH, f'{xpath_calc}'))).click()

        del valor

    # STPBASEPAGO
    def modificar_stpbasepago(self, driver, valor):

        # GABARITO
        # MR: Maior Remuneração
        # HS: Histórico Salarial
        # PS: Piso Salarial
        # SM: Salário Mínimo
        # VT: Vale Transporte

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:baseTabelada'))):

                campoSelecaoBases = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:baseTabelada')))
                selecaoBases = Select(campoSelecaoBases)

                if valor == 'MR':
                    selecaoBases.select_by_visible_text('Maior Remuneração')
                elif valor == 'HS':
                    selecaoBases.select_by_visible_text('Histórico Salarial')
                elif valor == 'PS':
                    selecaoBases.select_by_visible_text('Piso Salarial')
                elif valor == 'SM':
                    selecaoBases.select_by_visible_text('Salário Mínimo')
                elif valor == 'VT':
                    selecaoBases.select_by_visible_text('Vale Transporte')
        except TimeoutException:
            pass

        del valor

    # RVLINOUTROVALOR
    def modificar_rvlinoutrovalor(self, driver, valor):

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:valorInformadoPago'))):

                try:
                    if not "." in valor and not "," in valor:
                        valor = f"{float(valor):.2f}".replace(".", ",")
                    elif "." in valor:
                        valor = valor.replace(".", ",")
                except TypeError:
                    valor = f"{float(valor):.2f}".replace(".", ",")

                campoValor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:valorInformadoPago')))
                campoValor.click()
                campoValor.send_keys(valor)

        except TimeoutException:
            pass

        del valor

    # SFLVPGPROPORCIONALIDADE
    def marcar_sflvpgproporcionalidade(self, driver, valor):

        try:

            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "formulario:aplicarProporcionalidadeValorPago"))):
                elementoCheckbox = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "formulario:aplicarProporcionalidadeValorPago")))
                situacaoCheckbox = elementoCheckbox.is_selected()
                if valor == 'S':
                    if not situacaoCheckbox:
                        elementoCheckbox.click()
                elif valor == 'N':
                    if situacaoCheckbox:
                        elementoCheckbox.click()

        except TimeoutException:
            pass


        del valor

    # RVLQUANTIDADE
    def modificar_rvlquantidade(self, driver, valor):

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:valorPagoQuantidade'))):

                try:
                    valor = int(valor)
                    val = str(valor)
                except ValueError:
                    val = '1'

                campoQuantidade = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:valorPagoQuantidade')))
                campoQuantidade.click()
                campoQuantidade.send_keys(val)

        except TimeoutException:
            pass

        del valor

    # SDSCOMENTARIO
    def modificar_sdscomentario(self, driver, valor):

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:comentarios'))):
                campoComentatios = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:comentarios')))
                campoComentatios.click()
                campoComentatios.send_keys(valor)

        except TimeoutException:
            pass

        del valor

    # SFLREFLEXO13TERCEIRO
    def marcar_sflreflexo13terceiro(self, driver, valor, snmdescricaoverba):

        # indice = 0
        # var_loop = True

        verba_principal = '13º SALÁRIO SOBRE ' + snmdescricaoverba
        verba_principal = verba_principal.upper()
        print("- VERBA REFLEXA: ", verba_principal)

        # TABELA:
        # //tbody[@id="formulario:listagem:0:listaReflexo:tb"]//tr

        # TEXTO:
        # //td[@class="rich-table-cell colunaReflexo"]//span

        # CHECKBOX:
        # //td[@class="rich-table-cell colunaReflexo"]//input
        if valor == "S":
            try:
                elementos = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="formulario:listagem:0:listaReflexo:tb"]//tr')))
                print(f"- [QTD_ELEMENTOS_LOCALIZADOS_TB]: {len(elementos)}")
                for item in elementos:
                    try:
                        texto_item = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//span').get_attribute('textContent')
                        print(f"- [ITEM_LISTA]: {texto_item}")
                        if texto_item == verba_principal:
                            print(f"- [STATUS]: [ITEM_LOCALIZADO]")
                            field_checkbox = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//input')
                            field_checkbox.click()
                            return [True, '']
                        else:
                            continue
                    except Exception as e:
                        print(f"- [except]: {e}")
                        continue
            except Exception as e:
                msg = f"[except][marcar_sflreflexo13terceiro]: {e}"
                print(f"- {msg}")
                return [False, msg]

        # try:
        #
        #     while var_loop:
        #
        #         elemento_verba_principal = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:j_id574')))
        #
        #         # print("- # WEB: ", elemento_verba_principal.get_attribute("textContent"))
        #         # print("- # VAR: ", verba_principal)
        #
        #         if verba_principal == elemento_verba_principal.get_attribute("textContent"):
        #             elementoCheckbox = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:ativo')))
        #             situacaoCheckbox = elementoCheckbox.is_selected()
        #
        #             if valor == 'S':
        #                 if not situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #             elif valor == 'N':
        #                 if situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #
        #         indice += 1
        #
        # except TimeoutException:
        #     pass


        del valor

    # SFLREFLEXOFERIAS
    def marcar_sflreflexoferias(self, driver, valor, snmdescricaoverba):

        # indice = 0
        # var_loop = True
        indice_atual = 0

        # formulario:listagem:0:divDestinacoes
        verba_principal = 'FÉRIAS + 1/3 SOBRE ' + snmdescricaoverba
        verba_principal = verba_principal.upper()

        print("- VERBA REFLEXA: ", verba_principal)

        try:
            elementos = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="formulario:listagem:0:listaReflexo:tb"]//tr')))
            print(f"- [QTD_ELEMENTOS_LOCALIZADOS_TB]: {len(elementos)}")
            for indice, item in enumerate(elementos):
                try:
                    texto_item = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//span').get_attribute('textContent')
                    print(f"- [ITEM_LISTA]: {texto_item}")
                    if texto_item == verba_principal:
                        indice_atual = indice
                        print(f"- [STATUS]: [ITEM_LOCALIZADO]")
                        if valor == 'S':
                            field_checkbox = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//input')
                            field_checkbox.click()
                        break
                    else:
                        continue
                except Exception as e:
                    print(f"- [except]: {e}")
                    continue
        except Exception as e:
            msg = f"[except][marcar_sflreflexoferias]: {e}"
            print(f"- {msg}")
            return [False, msg]


        # try:
        #
        #     while var_loop:
        #
        #         elemento_verba_principal = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:j_id574')))
        #
        #         # print("- # WEB: ", elemento_verba_principal.get_attribute("textContent"))
        #         # print("- # VAR: ", verba_principal)
        #
        #         if verba_principal == elemento_verba_principal.get_attribute("textContent"):
        #             elementoCheckbox = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:ativo')))
        #             situacaoCheckbox = elementoCheckbox.is_selected()
        #
        #             indice_atual = indice
        #
        #             if valor == 'S':
        #                 if not situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #             elif valor == 'N':
        #                 if situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #
        #         indice += 1
        #
        # except TimeoutException:
        #     pass

        del valor

        return indice_atual

    # SFLREFLEXOAVISO
    def marcar_sflreflexoaviso(self, driver, valor, snmdescricaoverba):

        # indice = 0
        # var_loop = True

        # formulario:listagem:0:divDestinacoes
        verba_principal = 'AVISO PRÉVIO SOBRE ' + snmdescricaoverba
        verba_principal = verba_principal.upper()

        print("- VERBA REFLEXA: ", verba_principal)

        if valor == "S":
            try:
                elementos = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="formulario:listagem:0:listaReflexo:tb"]//tr')))
                print(f"- [QTD_ELEMENTOS_LOCALIZADOS_TB]: {len(elementos)}")
                for item in elementos:
                    try:
                        texto_item = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//span').get_attribute('textContent')
                        print(f"- [ITEM_LISTA]: {texto_item}")
                        if texto_item == verba_principal:
                            print(f"- [STATUS]: [ITEM_LOCALIZADO]")
                            field_checkbox = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//input')
                            field_checkbox.click()
                            break
                        else:
                            continue
                    except Exception as e:
                        print(f"- [except]: {e}")
                        continue
            except Exception as e:
                msg = f"[except][marcar_sflreflexoaviso]: {e}"
                print(f"- {msg}")
                return [False, msg]

        # try:
        #
        #     while var_loop:
        #
        #         elemento_verba_principal = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:j_id574')))
        #
        #         # print("- # WEB: ", elemento_verba_principal.get_attribute("textContent"))
        #         # print("- # VAR: ", verba_principal)
        #
        #         if verba_principal == elemento_verba_principal.get_attribute("textContent"):
        #             elementoCheckbox = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:ativo')))
        #             situacaoCheckbox = elementoCheckbox.is_selected()
        #
        #
        #             if valor == 'S':
        #                 if not situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #             elif valor == 'N':
        #                 if situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #
        #         indice += 1
        #
        # except TimeoutException:
        #     pass

        del valor

    # SFLREFLEXOREPOUSO
    def marcar_sflreflexorepouso(self, driver, valor, snmdescricaoverba):

        # indice = 0
        # var_loop = True

        # formulario:listagem:0:divDestinacoes
        verba_principal = 'REPOUSO SEMANAL REMUNERADO E FERIADO SOBRE ' + snmdescricaoverba
        verba_principal_2 = 'REPOUSO SEMANAL REMUNERADO SOBRE ' + snmdescricaoverba

        verba_principal = verba_principal.upper()
        verba_principal_2 = verba_principal_2.upper()

        print("- VERBA REFLEXA: ", verba_principal)

        if valor == "S":
            try:
                elementos = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="formulario:listagem:0:listaReflexo:tb"]//tr')))
                print(f"- [QTD_ELEMENTOS_LOCALIZADOS_TB]: {len(elementos)}")
                for item in elementos:
                    try:
                        texto_item = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//span').get_attribute('textContent')
                        print(f"- [ITEM_LISTA]: {texto_item}")
                        if texto_item == verba_principal or texto_item == verba_principal_2:
                            print(f"- [STATUS]: [ITEM_LOCALIZADO]")
                            field_checkbox = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//input')
                            field_checkbox.click()
                            break
                        else:
                            continue
                    except Exception as e:
                        print(f"- [except]: {e}")
                        continue
            except Exception as e:
                msg = f"[except][marcar_sflreflexorepouso]: {e}"
                print(f"- {msg}")
                del valor
                return -1
                # return [False, msg]

        # try:
        #
        #     while var_loop:
        #
        #         elemento_verba_principal = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:j_id574')))
        #
        #         # print("- # WEB: ", elemento_verba_principal.get_attribute("textContent"), type(elemento_verba_principal.get_attribute("textContent")))
        #         # print("- # VAR: ", verba_principal, type(verba_principal))
        #
        #         if verba_principal == elemento_verba_principal.get_attribute("textContent") or verba_principal_2 == elemento_verba_principal.get_attribute("textContent"):
        #             elementoCheckbox = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:ativo')))
        #             situacaoCheckbox = elementoCheckbox.is_selected()
        #
        #             if valor == 'S':
        #                 if not situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #             elif valor == 'N':
        #                 if situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #
        #         indice += 1
        #
        # except TimeoutException:
        #     return -1

        del valor

    # SFLREFLEXOART477CLT
    def marcar_sflreflexoart477clt(self, driver, valor, snmdescricaoverba):

        # indice = 0
        # var_loop = True

        verba_principal = 'MULTA DO ARTIGO 477 DA CLT SOBRE ' + snmdescricaoverba
        verba_principal = verba_principal.upper()

        print("- VERBA REFLEXA: ", verba_principal)

        if valor == "S":
            try:
                elementos = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="formulario:listagem:0:listaReflexo:tb"]//tr')))
                print(f"- [QTD_ELEMENTOS_LOCALIZADOS_TB]: {len(elementos)}")
                for item in elementos:
                    try:
                        texto_item = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//span').get_attribute('textContent')
                        print(f"- [ITEM_LISTA]: {texto_item}")
                        if texto_item == verba_principal:
                            print(f"- [STATUS]: [ITEM_LOCALIZADO]")
                            field_checkbox = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//input')
                            field_checkbox.click()
                            break
                        else:
                            continue
                    except Exception as e:
                        print(f"- [except]: {e}")
                        continue
            except Exception as e:
                msg = f"[except][marcar_sflreflexoart477clt]: {e}"
                print(f"- {msg}")
                del valor
                return -1
                # return [False, msg]

        # try:
        #
        #     while var_loop:
        #
        #         elemento_verba_principal = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:j_id574')))
        #
        #         if verba_principal == elemento_verba_principal.get_attribute("textContent"):
        #             elementoCheckbox = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:ativo')))
        #             situacaoCheckbox = elementoCheckbox.is_selected()
        #
        #             if valor == 'S':
        #                 if not situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #             elif valor == 'N':
        #                 if situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #
        #         indice += 1
        #
        # except TimeoutException:
        #     return -1

        del valor

    # SFLREFLEXOART467CLT - Parei Aqui
    def marcar_sflreflexoart467clt(self, driver, valor, snmdescricaoverba):

        # - MULTA DO ARTIGO 467 DA CLT SOBRE 13º SALÁRIO
        # indice = 0
        # var_loop = True

        verba_principal = 'MULTA DO ARTIGO 467 DA CLT SOBRE ' + snmdescricaoverba
        verba_principal = verba_principal.upper()
        print("- VERBA REFLEXA: ", verba_principal)

        if valor == "S":
            try:
                elementos = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="formulario:listagem:0:listaReflexo:tb"]//tr')))
                print(f"- [QTD_ELEMENTOS_LOCALIZADOS_TB]: {len(elementos)}")
                for item in elementos:
                    try:
                        texto_item = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//span').get_attribute('textContent')
                        print(f"- [ITEM_LISTA]: {texto_item}")
                        if texto_item == verba_principal:
                            print(f"- [STATUS]: [ITEM_LOCALIZADO]")
                            field_checkbox = item.find_element(By.XPATH, './/td[@class="rich-table-cell colunaReflexo"]//input')
                            field_checkbox.click()
                            break
                        else:
                            continue
                    except Exception as e:
                        print(f"- [except]: {e}")
                        continue
            except Exception as e:
                msg = f"[except][marcar_sflreflexoart467clt]: {e}"
                print(f"- {msg}")
                del valor
                return -1
                # return [False, msg]

        # try:
        #
        #     while var_loop:
        #
        #         elemento_verba_principal = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:j_id574')))
        #         # 'formulario:listagem:0:listaReflexo:0:j_id574'
        #
        #         if verba_principal == elemento_verba_principal.get_attribute("textContent"):
        #             elementoCheckbox = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:ativo')))
        #             situacaoCheckbox = elementoCheckbox.is_selected()
        #
        #             if valor == 'S':
        #                 if not situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #             elif valor == 'N':
        #                 if situacaoCheckbox:
        #                     elementoCheckbox.click()
        #                     var_loop = False
        #
        #         indice += 1
        #
        # except TimeoutException:
        #     return -1

        del valor


    # --- OUTRAS FUNCIONALIDADES
    # ADICIONAR BASES HISTÓRICO SALARIAL - DEVIDO
    def incluir_bases_HistoricoSalarial(self, driver, iidverbajrs):

        print("- TAMANHO DA PLANILHA BASE: ", len(self.planilha_basecalculo_exp_hs))

        if len(self.planilha_basecalculo_exp_hs) > 0:

            for i in range(len(self.planilha_basecalculo_exp_hs)):

                fk_col = self.planilha_basecalculo_exp_hs.loc[i, 'FKVERBA']
                tipo_base = self.planilha_basecalculo_exp_hs.loc[i, 'TIPO']
                if fk_col == iidverbajrs and tipo_base == "HS":
                    parcela = self.planilha_basecalculo_exp_hs.loc[i, 'PARCELA']
                    prop_int = self.planilha_basecalculo_exp_hs.loc[i, 'PROP/INT']
                    # print(f"|# {tipo_base} |# PARCELA: {parcela} |# PROP/INT: {prop_int} |# FK: {fk_col} |# ID: {iidverbajrs}")
                    self.selecionar_historicoSalarial(driver, parcela)
                    self.selecionar_historicoSalarial_proporcionalizar(driver, prop_int)
                    self.click_btnIncluir(driver)
                    self.objTime.aguardar_carregamento(driver)
                    sleep(1)

                    del parcela
                    del prop_int

                else:
                    continue

            del fk_col
            del tipo_base

    def selecionar_historicoSalarial(self, driver, valor):

        print("- PARÂMETRO: ", valor)
        base_hs = valor.upper()
        base_hs = base_hs.strip()

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:baseHistoricos'))):
                campoSelecao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:baseHistoricos')))
                elementoSelecao = Select(campoSelecao)
                elementoSelecao.select_by_visible_text(base_hs)

        except TimeoutException:
            print("- CAMPO DE SELEÇÃO 'HISTÓRIRO SALARIAL' NÃO LOCALIZADO!!")
            pass

        del valor
        del base_hs

    def selecionar_historicoSalarial_proporcionalizar(self, driver, valor):

        print("- PARÂMETRO: ", valor)

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:proporcionalizaHistorico'))):
                campoSelecao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:proporcionalizaHistorico')))
                elementoSelecao = Select(campoSelecao)

                if valor == 'S':
                    valor = 'Sim'
                elif valor == 'N':
                    valor = 'Não'

                elementoSelecao.select_by_visible_text(valor)

        except TimeoutException:
            print("- CAMPO DE SELEÇÃO 'PROPORCIONALIZAR' NÃO LOCALIZADO!!")
            pass

        del valor

    def click_btnIncluir(self, driver):
        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:incluirBaseHistorico'))):
                WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:incluirBaseHistorico'))).click()
        except TimeoutException:
            pass

    def checar_arquivo_pjc_diretorio(self, source):

        source_pjc = os.path.dirname(source)

        files = os.listdir(source_pjc)
        condicao_file = "*.PJC"

        condicao_file = glob.glob(rf"{source_pjc}\{condicao_file}")

        try:
            pjc = condicao_file[0].split("\\")[-1]

            if pjc in files:
                print(f"# - PJC Localizado: {condicao_file} - [Ok]")
                return condicao_file
            else:
                print(f"# - PJC Não Localizado.")
        except IndexError:
            print(f"# - PJC Não Localizado.")
            return False

    # ADICIONAR BASES - VERBAS
    def incluir_bases_Verba(self, driver, iidverbajrs):

        print("- TAMANHO DA PLANILHA BASE: ", len(self.planilha_basecalculo_exp_hs))

        if len(self.planilha_basecalculo_exp_hs) > 0:

            for i in range(len(self.planilha_basecalculo_exp_hs)):

                fk_col = self.planilha_basecalculo_exp_hs.loc[i, 'FKVERBA']
                tipo_base = self.planilha_basecalculo_exp_hs.loc[i, 'TIPO']

                if fk_col == iidverbajrs and tipo_base == "VB":
                    parcela = self.planilha_basecalculo_exp_hs.loc[i, 'PARCELA']
                    prop_int = self.planilha_basecalculo_exp_hs.loc[i, 'PROP/INT']
                    print(f"|# - {tipo_base} |# - PARCELA: {parcela} |# - PROP/INT: {prop_int} |# - FK: {fk_col} |# - ID: {iidverbajrs}")
                    self.selecionar_verba(driver, parcela)
                    self.selecionar_verba_integralizar(driver, prop_int)
                    self.click_btnIncluirVerba(driver)
                    self.objTime.aguardar_carregamento(driver)
                    sleep(1)

                    del parcela
                    del prop_int

                else:
                    continue

            del fk_col
            del tipo_base

    def selecionar_verba(self, driver, valor):

        print("- PARÂMETRO: ", valor)
        base_hs = valor.upper()

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:baseVerbaDeCalculo'))):
                campoSelecao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:baseVerbaDeCalculo')))
                elementoSelecao = Select(campoSelecao)
                elementoSelecao.select_by_visible_text(base_hs)

        except TimeoutException:
            print("- CAMPO DE SELEÇÃO 'VERBA' NÃO LOCALIZADO!!")
            pass
        except NoSuchElementException:
            pass

        del base_hs
        del valor

    def selecionar_verba_integralizar(self, driver, valor):

        print("- PARÂMETRO: ", valor)

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:proporcionalizaHistorico'))):
                campoSelecao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:proporcionalizaHistorico')))
                elementoSelecao = Select(campoSelecao)

                if valor == 'S':
                    valor = 'Sim'
                elif valor == 'N':
                    valor = 'Não'

                elementoSelecao.select_by_visible_text(valor)

        except TimeoutException:
            print("- CAMPO DE SELEÇÃO 'PROPORCIONALIZAR' NÃO LOCALIZADO!!")
            pass

        del valor

    def click_btnIncluirVerba(self, driver):
        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:incluirItemProp'))):
                WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:incluirItemProp'))).click()
        except TimeoutException:
            pass

    # ADICIONAR BASES HISTÓRICO SALARIAL - PAGO
    def incluir_bases_HistoricoSalarialValPago(self, driver, iidverbajrs):

        print("- TAMANHO DA PLANILHA BASE: ", len(self.planilha_basecalculovalorpago_exp_hs))

        if len(self.planilha_basecalculovalorpago_exp_hs) > 0:

            for i in range(len(self.planilha_basecalculovalorpago_exp_hs)):

                fk_col = self.planilha_basecalculovalorpago_exp_hs.loc[i, 'FKVERBA']
                if fk_col == iidverbajrs:
                    parcela = self.planilha_basecalculovalorpago_exp_hs.loc[i, 'PARCELA']
                    prop = self.planilha_basecalculovalorpago_exp_hs.loc[i, 'PROP']
                    # print(f"|# PARCELA: {parcela} |# PROP: {prop} |# FK: {fk_col} |# ID: {iidverbajrs}")
                    self.selecionar_historicoSalarialValPago(driver, parcela)
                    self.selecionar_historicoSalarialValPago_proporcionalizar(driver, prop)
                    self.click_btnIncluirValpago(driver)
                    self.objTime.aguardar_carregamento(driver)
                    sleep(1)

                    del parcela
                    del prop

                else:
                    continue

            del fk_col

    def selecionar_historicoSalarialValPago(self, driver, valor):

        print("- PARÂMETRO: ", valor)
        base_hs = valor.upper()
        base_hs = base_hs.strip()

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:baseHistoricosValorPago'))):
                campoSelecao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:baseHistoricosValorPago')))
                elementoSelecao = Select(campoSelecao)
                elementoSelecao.select_by_visible_text(base_hs)

        except TimeoutException:
            print("- CAMPO DE SELEÇÃO 'HISTÓRIRO SALARIAL' NÃO LOCALIZADO!!")
            pass

        del valor

    def selecionar_historicoSalarialValPago_proporcionalizar(self, driver, valor):

        print("- PARÂMETRO: ", valor)

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:proporcionalizaHistoricoDoValorPago'))):
                campoSelecao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:proporcionalizaHistoricoDoValorPago')))
                elementoSelecao = Select(campoSelecao)

                if valor == 'Sim':
                    elementoSelecao.select_by_visible_text("Não")
                # elif valor == 'Não':
                #     valor = 'Não'

                elementoSelecao.select_by_visible_text(valor)

        except TimeoutException:
            print("- CAMPO DE SELEÇÃO 'PROPORCIONALIZAR' NÃO LOCALIZADO!!")
            pass

        del valor

    def click_btnIncluirValpago(self, driver):

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:incluirBaseHistoricoValorPago'))):
                WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:incluirBaseHistoricoValorPago'))).click()
        except TimeoutException:
            print("- ELEMENTO 'BOTÃO ADICIONAR' NÃO LOCALIZADO!!")
            pass

    # VERBA REFLEXA
    def expandir_verba_reflexa(self, driver, valor):

        s_loop = True
        indice = 0

        # if valor.islower():
        #     valor.upper()
        valor = valor.strip().upper()

        while s_loop:

            try:

                if WebDriverWait(driver, 1).until(EC.visibility_of_element_located((By.ID, f'formulario:listagem:{indice}:j_id561'))):

                    elemento_verba_principal = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.ID, f'formulario:listagem:{indice}:j_id561')))
                    print("- ELEMENTO: ", elemento_verba_principal.text)
                    print("- PARÂMETRO: ", valor)
                    if elemento_verba_principal.text == valor:
                        WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.ID, f'formulario:listagem:{indice}:divDestinacoes'))).click()
                        s_loop = False

            except TimeoutException:
                # print("- ELEMENTO NÃO ENCONTRADO!")
                break

            indice += 1

        del valor

    # SALVAR
    def selecionar_btnSalvarOperacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.botao#formulario\:salvar'))).click()

    # - Confirmar Alteração se valor da verba for Informado
    def clicar_btnConfirmarOperacao(self, driver):
        xpath_btn = "//input[@id='popup_ok' and @value='Ok']"
        # xpath_calc = "//input[@id='formulario:tipoDoValorPago:1' and @value='CALCULADO']"
        try:
            WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, f'{xpath_btn}'))).click()
        except TimeoutException as e:
            print(f"- [except]: {e}")

    # CANCELAR
    def click_BtnCancelar(self, driver):
        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:cancelar'))):
                WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:cancelar'))).click()
        except TimeoutException:
            pass

    def definirValorPagoBase(self, driver, valor):

        control = True

        while control:

            try:
                if WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.ID, 'formulario:baseTabelada'))):

                    campoBase = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:baseTabelada')))
                    selecaoBase = Select(campoBase)
                    selecaoBase.select_by_visible_text(valor)

                    control = False
            except TimeoutException:
                print('- Campo não encontrado! Repetindo a operação...')
                continue

        del valor

    def definirValorPagoHistorico(self, driver, valor):

        #
        campoHistorico = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:baseHistoricosValorPago')))
        selecaoHistorico = Select(campoHistorico)
        selecaoHistorico.select_by_visible_text(valor)

        del valor

    def definirValorPagoProporcionalizar(self, driver, valor):

        positivo = 'Sim'
        negativo = 'Não'

        campoProporcionalizar = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:proporcionalizaHistoricoDoValorPago')))
        selecaoProporcionalizar = Select(campoProporcionalizar)

        if valor == positivo:
            selecaoProporcionalizar.select_by_visible_text(negativo)
            selecaoProporcionalizar.select_by_visible_text(valor)
        else:
            selecaoProporcionalizar.select_by_visible_text(positivo)
            selecaoProporcionalizar.select_by_visible_text(valor)


        del valor

    def definirValorPagoAdicionar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:incluirBaseHistoricoValorPago'))).click()

    def verificar_existencia_conteudo(self):

        for vl in self.planilha_base['IIDVERBAJRS']:
            print("- Pk: ", vl, type(vl))
            if vl >= 1:
                return True
            else:
                return False

    def selecionar_todos_checks(self, driver):
        element = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:listagem:selecionarTodos')))
        if not element.is_selected():
            element.click()

    def selecionar_sobrescrever(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoRegeracao:1'))).click()

    def click_regerar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:regerarOcorrencias'))).click()

    def click_confirmar_regeracao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'popup_ok'))).click()

    def registrar_horario_inicial(self):
        horario_inicial_full = datetime.today()
        self.hrinicialLog = horario_inicial_full.strftime('%H:%M:%S')
        # print(inicio)
        file_txt_log = open(os.getcwd() + '\log.txt', "a")
        file_txt_log.write('- Horário Inicial: ' + self.hrinicialLog + '\n\n')
        return file_txt_log.close()

    def registrar_horario_final(self):
        horario_final_full = datetime.now()
        self.hrfinalLog = horario_final_full.strftime('%H:%M:%S')
        # print(final)
        file_txt_log = open(os.getcwd() + '\log.txt', "a")
        file_txt_log.write('\n- Horário Final: ' + self.hrfinalLog + '\n')
        # file_txt_log.write(f'\n- Tempo: {"00:00:00"}\n')
        return file_txt_log.close()

    def encaminhar_log(self, id_processo, source):

        contador = 0
        source_log = os.getcwd() + '\log.txt'
        diretorio_destino = os.path.dirname(source)

        target_log = fr"{diretorio_destino}\ID_{id_processo} - Log Automação.txt"

        if os.path.exists(target_log):
            for item in os.listdir(diretorio_destino):
                if 'Log' in item:
                    contador += 1
            target_log = fr"{diretorio_destino}\ID_{id_processo} - Log Automação ({contador}).txt"
            contador += 1
            shutil.copy(source_log, target_log)
        else:
            shutil.copy(source_log, target_log)

        # Limpar
        f = open(source_log, 'w')
        f.close()


    def verificar_valorTb(self, value):

        try:
            if math.isnan(value):
                return False
        except TypeError:
            return True


    def clear_nullValuesSheets(self):

        linhas_em_branco = self.planilha_base['IIDVERBAJRS'].isna()
        for indice, valor in enumerate(linhas_em_branco):
            print(f"[{indice}] - {valor} '{type(valor)}'")
            if valor:
                self.planilha_base.dropna(axis=0, how='any', subset=['IIDVERBAJRS'], inplace=True)


    def functionMain(self, driver):

        from Calculo.verbas_reflexa import VerbaReflexa

        delayG = 2.5
        # delayM = 1.5
        delayP = 0.5

        #  -- PJECALC PAGINA INICIAL -- #
        sleep(delayG)

        # - Função para remover linhas vazias da Tabela TBVERBABD
        self.clear_nullValuesSheets()

        if self.verificar_existencia_conteudo():

            #  -- PJECALC - VERBAS -- #
            self.acessarVerbas(driver)
            self.objTime.aguardar_carregamento(driver)
            sleep(delayG)

            # -- VERBAS TBVERBABD
            for vl in self.planilha_base['IIDVERBAJRS'].index:

                iidverbajrs = ""
                snmdescricaoverba = ""
                snmverbaexpressopjecalc = ""
                stpvalor = ""
                sflincidenciainss = ""
                sflincidenciairpf = ""
                sflincidenciafgts = ""
                sflincidenciaprevprivada = ""
                sflincidenciapensao = ""
                sflreflexo13terceiro = ""
                sflreflexoferias = ""
                sflreflexoaviso = ""
                sflreflexorepouso = ""
                sflreflexorepousoColumn = ""
                sflreflexoart477clt = ""
                sflreflexoart477cltColumn = ""
                sflreflexoart467clt = ""
                sflreflexoart467cltColumn = ""
                sflproporcionalidade = ""

                # Reflexa
                sflsepararferias = ""
                sflreflexohextras50 = ""
                sflreflexohextras50Column = ""
                sflreflexohextras100 = ""
                sflreflexohextras100Column = ""
                sflreflexoganual = ""
                sflreflexoganualColumn = ""
                sflreflexogsemestral = ""
                sflreflexogsemestralColumn = ""
                sflreflexolpremio = ""
                sflreflexolpremioColumn = ""
                sflreflexoapip = ""
                sflreflexoapipColumn = ""
                sflreflexoplr = ""
                sflreflexoplrColumn = ""

                self.acessarExpresso(driver)
                self.objTime.aguardar_carregamento(driver)
                sleep(delayG)

                for vc in self.planilha_base.columns:

                    dado = self.planilha_base.loc[vl, vc]

                    result = self.verificar_valorTb(dado)
                    if result is False:
                        continue

                    if vc == 'IIDVERBAJRS':
                        iidverbajrs = int(dado)
                        print("- IIDVERBAJRS: ", iidverbajrs)
                        continue

                    if vc == 'SNMVERBAJRS':
                        print("- SNMVERBAJRS: ", dado)
                        continue

                    if vc == 'SNMVERBAEXPRESSOPJECALC':
                        snmverbaexpressopjecalc = dado
                        print("- SNMVERBAEXPRESSOPJECALC: ", dado)
                        self.selecionarVerbaPrincipal(driver, dado)
                        self.salvarOperacao(driver)
                        self.objTime.aguardar_carregamento(driver)
                        sleep(delayG)
                        self.acessarParametrosVerba(driver)
                        self.objTime.aguardar_carregamento(driver)
                        sleep(delayG)
                        continue

                    if vc == 'SNMDESCRICAOVERBA':
                        snmdescricaoverba = dado
                        snmdescricaoverba = snmdescricaoverba.strip()
                        print("- SNMDESCRICAOVERBA: ", snmdescricaoverba)
                        self.modificar_snmdescricaoverba(driver, snmdescricaoverba)
                        continue

                    if vc == 'STPPARCELA':
                        print("- STPPARCELA: ", dado)
                        self.marcar_stpparcela(driver, dado)
                        continue

                    if vc == 'STPVALOR':
                        print("- STPVALOR: ", dado)
                        stpvalor = dado
                        self.marcar_stpvalor(driver, dado)
                        continue

                    if vc == 'SFLINCIDENCIAINSS':
                        print("- SFLINCIDENCIAINSS: ", dado)
                        self.marcar_sflincidenciainss(driver, dado)
                        sflincidenciainss = dado
                        continue

                    if vc == 'SFLINCIDENCIAIRPF':
                        print("- SFLINCIDENCIAIRPF: ", dado)
                        self.marcar_sflincidenciairpf(driver, dado)
                        sflincidenciairpf = dado
                        continue

                    if vc == 'SFLINCIDENCIAFGTS':
                        print("- SFLINCIDENCIAFGTS: ", dado)
                        self.marcar_sflincidenciafgts(driver, dado)
                        sflincidenciafgts = dado
                        continue

                    if vc == 'SFLINCIDENCIAPREVPRIVADA':
                        print("- SFLINCIDENCIAPREVPRIVADA: ", dado)
                        self.marcar_sflincidenciaprevprivada(driver, dado)
                        sflincidenciaprevprivada = dado
                        continue

                    if vc == 'SFLINCIDENCIAPENSAO':
                        print("- SFLINCIDENCIAPENSAO: ", dado)
                        self.marcar_sflincidenciapensao(driver, dado)
                        sflincidenciapensao = dado
                        continue

                    if vc == 'STPCARACTERISTICAVERBA':
                        print("- STPCARACTERISTICAVERBA: ", dado)
                        self.marcar_stpcaracteristicaverba(driver, dado)
                        continue

                    if vc == 'STPOCORRENCIAPAGAMENTO':
                        print("- STPOCORRENCIAPAGAMENTO: ", dado)
                        self.marcar_stpocorrenciapagamento(driver, dado)
                        continue

                    if vc == 'STPJUROSSUMULA':
                        print("- STPJUROSSUMULA: ", dado)
                        self.marcar_stpjurossumula(driver, dado)
                        continue

                    if vc == 'STPTIPOVERBA':
                        print("- STPTIPOVERBA: ", dado)
                        self.marcar_stptipoverba(driver, dado)
                        continue

                    if vc == 'STPGERARVERBAREFLEXA':
                        print("- STPGERARVERBAREFLEXA: ", dado)
                        self.marcar_stpgerarverbareflexa(driver, dado)
                        continue

                    if vc == 'STPGERARVERBAPRINCIPAL':
                        print("- STPGERARVERBAPRINCIPAL: ", dado)
                        self.marcar_stpgerarverbaprincipal(driver, dado)
                        continue

                    if vc == 'SFLCOMPORPRINCIPAL':
                        print("- SFLCOMPORPRINCIPAL: ", dado)
                        self.marcar_sflcomporprincipal(driver, dado)
                        if dado == 'N':
                            # Incidência - Marcar Novamente os Checkboxs
                            self.marcar_sflincidenciainss(driver, sflincidenciainss)
                            self.marcar_sflincidenciairpf(driver, sflincidenciairpf)
                            self.marcar_sflincidenciafgts(driver, sflincidenciafgts)
                            self.marcar_sflincidenciaprevprivada(driver, sflincidenciaprevprivada)
                            self.marcar_sflincidenciapensao(driver, sflincidenciapensao)
                            sleep(delayP)
                        continue

                    if vc == 'SFLPADRAOZERARVALORNEGATIVO':
                        print("- SFLPADRAOZERARVALORNEGATIVO: ", dado)
                        self.marcar_sflpadraozerarvalornegativo(driver, dado)
                        continue

                    if vc == 'DDTINICIO':
                        print("- DDTINICIO: ", dado)
                        self.modificar_ddtinicio(driver, dado)
                        continue

                    if vc == 'DDTTERMINO':
                        print("- DDTTERMINO: ", dado)
                        self.modificar_ddttermino(driver, dado)
                        continue

                    if vc == 'SFLEXCLUIRFALTAJUSTIFICADA':
                        print("- SFLEXCLUIRFALTAJUSTIFICADA: ", dado)
                        self.marcar_sflexcluirfaltajustificada(driver, dado)
                        continue

                    if vc == 'SFLEXCLUIRFALTANAOJUSTIFICADA':
                        print("- SFLEXCLUIRFALTANAOJUSTIFICADA: ", dado)
                        self.marcar_sflexcluirfaltanaojustificada(driver, dado)
                        continue

                    if vc == 'SFLEXCLUIRFERIASGOZADAS':
                        print("- SFLEXCLUIRFERIASGOZADAS: ", dado)
                        self.marcar_sflexcluirferiasgozadas(driver, dado)
                        continue

                    if vc == 'SFLDOBRARVALOR':
                        print("- SFLDOBRARVALOR: ", dado)
                        self.marcar_sfldobrarvalor(driver, dado)
                        continue

                    if stpvalor == "I":

                        if vc == 'RVLVALORDEVIDO':
                            print("- RVLVALORDEVIDO: ", dado)
                            self.modificar_rvlvalordevido(driver, dado)
                            continue

                        if vc == 'SFLPROPORCIONALIDADE':
                            print("- SFLPROPORCIONALIDADE: ", dado)
                            self.marcar_sflproporcionalidade(driver, dado)
                            continue

                    elif stpvalor == "C":

                        if vc == 'SFLPROPORCIONALIDADE':
                            sflproporcionalidade = dado

                        if vc == 'STPBASEDEVIDO':
                            print("- STPBASEDEVIDO: ", dado)
                            self.modificar_stpbasedevido(driver, dado)
                            if dado == 'HS':
                                self.incluir_bases_HistoricoSalarial(driver, iidverbajrs)
                            else:
                                print("- SFLPROPORCIONALIDADE: ", sflproporcionalidade)
                                self.marcar_sflproporcionalidade_basecalculo(driver, sflproporcionalidade)

                            self.incluir_bases_Verba(driver, iidverbajrs)
                            # ADICIONAR HISTÓRICO
                            continue

                        if vc == 'STPDIVISOR':
                            print("- STPDIVISOR: ", dado)
                            self.modificar_stpdivisor(driver, dado)
                            continue

                        if vc == 'RVLOUTRODIVISOR':
                            print("- RVLOUTRODIVISOR: ", dado)
                            self.modificar_rvloutrodivisor(driver, dado)
                            continue

                        if vc == 'RVLMULTIPLICADOR':
                            print("- RVLMULTIPLICADOR: ", dado)
                            self.modificar_rvlmultiplicador(driver, dado)
                            continue

                        if vc == 'STPQUANTIDADE':
                            print("- STPQUANTIDADE: ", dado)
                            self.marcar_stpquantidade(driver, dado)
                            continue

                        if vc == 'STPCALENDARIO':
                            print("- STPCALENDARIO: ", dado)
                            self.modificar_stpcalendario(driver, dado)
                            continue

                        if vc == 'RVLOUTROVALORQTD':
                            print("- RVLOUTROVALORQTD: ", dado)
                            self.modificar_rvloutrovalorqtd(driver, dado)
                            continue

                        if vc == 'SFLQTDPROPORCIONALIDADE':
                            print("- SFLQTDPROPORCIONALIDADE: ", dado)
                            self.marcar_sflqtdproporcionalidade(driver, dado)
                            continue

                    if vc == 'STPVALORPAGO':
                        print("- STPVALORPAGO: ", dado)
                        self.marcar_stpvalorpago(driver, dado)
                        continue

                    if vc == 'STPBASEPAGO':
                        print("- STPBASEPAGO: ", dado)
                        self.modificar_stpbasepago(driver, dado)

                        if dado == 'HS':
                            self.incluir_bases_HistoricoSalarialValPago(driver, iidverbajrs)
                        continue

                    if vc == 'RVLINOUTROVALOR':
                        print("- RVLINOUTROVALOR: ", dado)
                        self.modificar_rvlinoutrovalor(driver, dado)
                        continue

                    if vc == 'SFLVPGPROPORCIONALIDADE':
                        print("- SFLVPGPROPORCIONALIDADE: ", dado)
                        self.marcar_sflvpgproporcionalidade(driver, dado)
                        continue

                    if vc == 'RVLQUANTIDADE':
                        print("- RVLQUANTIDADE: ", dado)
                        self.modificar_rvlquantidade(driver, dado)
                        continue

                    if vc == 'SDSCOMENTARIO':
                        print("- SDSCOMENTARIO: ", dado)
                        if type(dado) == str:
                            self.modificar_sdscomentario(driver, dado)
                        continue

                    if vc == 'SFLREFLEXO13TERCEIRO':
                        sflreflexo13terceiro = dado
                        print("- SFLREFLEXO13TERCEIRO: ", dado)
                        continue

                    if vc == 'SFLREFLEXOAVISO':
                        sflreflexoaviso = dado
                        print("- SFLREFLEXOAVISO: ", dado)
                        continue

                    if vc == 'SFLREFLEXOFERIAS':
                        sflreflexoferias = dado
                        print("- SFLREFLEXOFERIAS: ", dado)
                        continue

                    if vc == 'SFLREFLEXOART477CLT':
                        sflreflexoart477clt = dado
                        sflreflexoart477cltColumn = vc
                        print("- SFLREFLEXOART477CLT: ", sflreflexoart477clt)
                        continue

                    if vc == 'SFLREFLEXOART467CLT':
                        sflreflexoart467clt = dado
                        sflreflexoart467cltColumn = vc
                        print("- SFLREFLEXOART467CLT: ", sflreflexoart467clt)
                        continue

                    if vc == 'SFLREFLEXOREPOUSO':
                        sflreflexorepouso = dado
                        sflreflexorepousoColumn = vc
                        print("- SFLREFLEXOREPOUSO: ", sflreflexorepouso)
                        continue

                    if vc == 'SFLSEPARARFERIAS':
                        sflsepararferias = dado
                        print("- SFLSEPARARFERIAS: ", sflsepararferias)
                        continue

                    if vc == 'SFLREFLEXOHEXTRAS50':
                        sflreflexohextras50 = dado
                        sflreflexohextras50Column = vc
                        print("- SFLREFLEXOHEXTRAS50: ", sflreflexohextras50)
                        continue
                    if vc == 'SFLREFLEXOHEXTRAS100':
                        sflreflexohextras100 = dado
                        sflreflexohextras100Column = vc
                        print("- SFLREFLEXOHEXTRAS100: ", sflreflexohextras100)
                        continue
                    if vc == 'SFLREFLEXOGANUAL':
                        sflreflexoganual = dado
                        sflreflexoganualColumn = vc
                        print("- SFLREFLEXOGANUAL: ", sflreflexoganual)
                        continue
                    if vc == 'SFLREFLEXOGSEMESTRAL':
                        sflreflexogsemestral = dado
                        sflreflexogsemestralColumn = vc
                        print("- SFLREFLEXOGSEMESTRAL: ", sflreflexogsemestral)
                        continue
                    if vc == 'SFLREFLEXOLPREMIO':
                        sflreflexolpremio = dado
                        sflreflexolpremioColumn = vc
                        print("- SFLREFLEXOLPREMIO: ", sflreflexolpremio)
                        continue
                    if vc == 'SFLREFLEXOAPIP':
                        sflreflexoapip = dado
                        sflreflexoapipColumn = vc
                        print("- SFLREFLEXOAPIP: ", sflreflexoapip)
                        continue
                    if vc == 'SFLREFLEXOPLR':
                        sflreflexoplr = dado
                        sflreflexoplrColumn = vc
                        print("- SFLREFLEXOPLR: ", sflreflexoplr)
                        continue


                    self.objTime.limparFilesTemp()


                # Time
                sleep(delayG)
                # SALVAR
                self.selecionar_btnSalvarOperacao(driver)

                # - Confirmar Operação se valor da verba for Informado
                if stpvalor == "I":
                    sleep(delayG)
                    self.clicar_btnConfirmarOperacao(driver)

                self.objTime.aguardar_carregamento(driver)
                sleep(delayG)

                # VERIFICAÇÃO
                r = self.verificacao_2_new(driver) # [Teste]

                # print('# RETORNO V: ', r)
                if r != 'erro':
                    # EXPANDIR O CONTEUDO DA VERBA REFLEXA
                    if sflreflexo13terceiro == 'S' or sflreflexoaviso == 'S' or sflreflexoferias == 'S' or sflreflexoart477clt == 'S' or sflreflexorepouso == 'S' or sflreflexoart467clt == 'S' or sflsepararferias == 'S' or sflreflexoganual == 'S' or sflreflexogsemestral == 'S' or sflreflexolpremio == 'S' or sflreflexoapip == 'S' or sflreflexohextras50 == 'S' or sflreflexohextras100 == 'S' or sflreflexoplr == 'S':

                        self.expandir_verba_reflexa(driver, snmdescricaoverba)
                        self.objTime.aguardar_carregamento(driver)
                        sleep(delayG)

                        # SFLREFLEXO13TERCEIRO
                        self.marcar_sflreflexo13terceiro(driver, sflreflexo13terceiro, snmdescricaoverba)
                        self.objTime.aguardar_carregamento(driver)
                        sleep(0.5)

                        # SFLREFLEXOAVISO
                        self.marcar_sflreflexoaviso(driver, sflreflexoaviso, snmdescricaoverba)
                        self.objTime.aguardar_carregamento(driver)
                        sleep(0.5)

                        # SFLREFLEXOFERIAS
                        indice_ferias_reflexa = self.marcar_sflreflexoferias(driver, sflreflexoferias, snmdescricaoverba)
                        self.objTime.aguardar_carregamento(driver)
                        sleep(0.5)

                        # SFLREFLEXOART477CLT
                        result_477clt = self.marcar_sflreflexoart477clt(driver, sflreflexoart477clt, snmdescricaoverba)
                        if result_477clt == -1 and sflreflexoart477clt == "S":
                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa(sflreflexoart477cltColumn, snmdescricaoverba)
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del result_477clt
                            del sflreflexoart477cltColumn
                            del idGetRegistro
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)
                        else:
                            self.objTime.aguardar_carregamento(driver)
                            sleep(0.5)

                        # SFLREFLEXOART467CLT
                        result_467clt = self.marcar_sflreflexoart467clt(driver, sflreflexoart467clt, snmdescricaoverba)
                        if result_467clt == -1 and sflreflexoart467clt == "S":
                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa(sflreflexoart467cltColumn, "")
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del result_467clt
                            del sflreflexoart467cltColumn
                            del idGetRegistro
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)
                        else:
                            self.objTime.aguardar_carregamento(driver)
                            sleep(0.5)

                        # SFLREFLEXOREPOUSO
                        result_repouso = self.marcar_sflreflexorepouso(driver, sflreflexorepouso, snmdescricaoverba)
                        # break
                        if result_repouso == -1 and sflreflexorepouso == "S":
                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa(sflreflexorepousoColumn, "")
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del result_repouso
                            del sflreflexorepousoColumn
                            del idGetRegistro
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)
                        else:
                            self.objTime.aguardar_carregamento(driver)
                            sleep(0.5)

                        # 01 SFLSEPARARFERIAS - PARAMETRIZAR PARCELA REFLEXA (SEPARAR TERÇO DE FÉRIAS)
                        if sflsepararferias == "S":

                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            # - Etapa 01
                            objVerbaReflexa.click_parametrizarReflexo(driver, indice_ferias_reflexa)
                            self.objTime.aguardar_carregamento(driver)
                            sleep(0.5)
                            objVerbaReflexa.modificar_snmdescricaoverba(driver, "FÉRIAS")
                            objVerbaReflexa.modificar_rvlmultiplicador(driver, "1")
                            sleep(0.5)
                            # objVerbaReflexa.salvar_operacao(driver)
                            objVerbaReflexa.selecionar_btnSalvarOperacao(driver)
                            self.objTime.aguardar_carregamento(driver)
                            sleep(delayG)
                            # self.verificacao(driver)
                            self.verificacao_2_new_parcelaReflexa(driver)
                            # sleep(1)

                            # - Etapa 02
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa("SFLSEPARARFERIAS", snmverbaexpressopjecalc)
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del idGetRegistro
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)

                        # 02 SFLREFLEXOHEXTRAS50 - PARAMETRIZAR PARCELA REFLEXA
                        if sflreflexoganual == "S":
                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa(sflreflexoganualColumn, "")
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del sflreflexoganualColumn
                            del idGetRegistro
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)

                        # 03 SFLREFLEXOHEXTRAS100 - PARAMETRIZAR PARCELA REFLEXA
                        if sflreflexogsemestral == "S":
                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa(sflreflexogsemestralColumn, "")
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del sflreflexogsemestralColumn
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)

                        # 04 SFLREFLEXOGANUAL - PARAMETRIZAR PARCELA REFLEXA
                        if sflreflexolpremio == "S":
                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa(sflreflexolpremioColumn, "")
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del sflreflexolpremioColumn
                            del idGetRegistro
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)

                        # 05 SFLREFLEXOGSEMESTRAL - PARAMETRIZAR PARCELA REFLEXA
                        if sflreflexoapip == "S":
                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa(sflreflexoapipColumn, "")
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del sflreflexoapipColumn
                            del idGetRegistro
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)

                        # 06 SFLREFLEXOLPREMIO - PARAMETRIZAR PARCELA REFLEXA
                        if sflreflexohextras50 == "S":
                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa(sflreflexohextras50Column, snmverbaexpressopjecalc)
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del sflreflexohextras50Column
                            del idGetRegistro
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)

                        # 07 SFLREFLEXOAPIP - PARAMETRIZAR PARCELA REFLEXA
                        if sflreflexohextras100 == "S":
                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa(sflreflexohextras100Column, snmverbaexpressopjecalc)
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del sflreflexohextras100Column
                            del idGetRegistro
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)

                        # 08 SFLREFLEXOPLR- PARAMETRIZAR PARCELA REFLEXA
                        if sflreflexoplr == "S":
                            objVerbaReflexa = VerbaReflexa(self.sourceFileExcel)
                            idGetRegistro = objVerbaReflexa.filtro_parcelaReflexa(sflreflexoplrColumn, snmverbaexpressopjecalc)
                            if len(idGetRegistro) > 0:
                                objVerbaReflexa.functon_main_parcelaReflexa(driver, idGetRegistro, iidverbajrs, snmdescricaoverba)
                                self.selecionar_btnSalvarOperacao(driver)
                                self.objTime.aguardar_carregamento(driver)
                                sleep(delayG)
                                self.verificacao_2_new_parcelaReflexa(driver)

                            del sflreflexoplrColumn
                            del idGetRegistro
                            self.objTime.limparFilesTemp()
                            gc.collect(generation=0)
                            gc.collect(generation=1)
                            gc.collect(generation=2)

                sleep(1.5)

                # LIMPAR VARIÁVEL
                del dado
                del snmdescricaoverba
                del stpvalor
                del sflreflexo13terceiro
                del sflreflexoaviso
                del sflreflexoferias
                del sflreflexoart477clt
                del sflreflexoart467clt
                del sflreflexorepouso
                del sflsepararferias
                # del result_repouso
                del sflreflexoganual
                del sflreflexogsemestral
                del sflreflexolpremio
                del sflreflexoapip
                del sflreflexohextras50
                del sflreflexohextras100
                del sflreflexoplr
                # del result_477clt
                # del result_467clt
                self.objTime.limparFilesTemp()
                gc.collect(generation=0)
                gc.collect(generation=1)
                gc.collect(generation=2)
                # break

            # - Regerar sobrescrevendo
            self.selecionar_todos_checks(driver)
            self.objTime.aguardar_carregamento(driver)
            sleep(self.delayP)
            self.selecionar_sobrescrever(driver)
            self.click_regerar(driver)
            self.click_confirmar_regeracao(driver)
            self.objTime.aguardar_carregamento(driver)
            sleep(self.delayP)

        else:
            print('- [Sem registros para Verbas]')
            return 1

if __name__ == '__main__':
    pass
    # # OBJETO VERBA
    # objVerba = Verbas("")
    # # FUNÇÃO PRINCIPAL
    # objVerba.functionMain()
