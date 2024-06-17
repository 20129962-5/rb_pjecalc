from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import xlrd
import time
import os
import gc
# from Calculo.pjecalc_dados_calculo import DadosCalculo
from Tools.pjecalc_control import Control


class Custas:


    def __init__(self, source):
        self.source = source
        self.planilha_base = pd.read_excel(self.source, sheet_name='PJE-BD', header=1)
        self.tamanho_planilha_base = len(self.planilha_base)
        self.delay = 10
        self.delayG = 1.5
        self.objTools = Control()
        self.qtd_custas_recolhidas = 0


    # Variáveis de controle
    var_controle_float = 0.0
    var_controle_string = ""
    var_controle_int = 1

    # Atributos — Custas Devidas
    base_para_custas_conhecimento = ''
    custas_do_reclamante = ''
    custas_reclamado_conhecimento = ''
    custas_recda_conhecimento_info_vcto = ""
    custas_recda_conhecimento_info_valor = ""

    juros_correcao_monetaria = ""
    vencimento_custas_devidas = ""
    valor_custas_devidas = ""

    # Instanciando um objeto da classe Controle
    # objeto_controle = Control()

    # Atributos - Custas Recolhidas
    custas_recolhidas_rcdo1_vencimento = ''
    custas_recolhidas_rcdo1_valor = ''


    def acessar_custas_judiciais(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageCustasJudiciais"))).click()

    def preencher_custas_devidas(self, driver):

        for i in range(self.tamanho_planilha_base):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                continue
            # Base para Custas de Conhecimento e Liquidação
            elif coluna_identificador == 'base_para_custas_conhecimento':
                self.base_para_custas_conhecimento = coluna_informacao
                print('- Base para Custas de Conhecimento e Liquidação: ', self.base_para_custas_conhecimento)
            elif coluna_identificador == 'custas_do_reclamante':
                self.custas_do_reclamante = coluna_informacao
                print('- Custas do Reclamante - Conhecimento: ', self.custas_do_reclamante)
            elif coluna_identificador == 'custas_reclamado_conhecimento':
                self.custas_reclamado_conhecimento = coluna_informacao
                print('- Custas do Reclamado: ', self.custas_reclamado_conhecimento)

                # Preencher os dados do PJeCalc
                # * Base para Custas de Conhecimento e Liquidação
                campo_base_custas = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:baseParaCustasCalculadas')))
                selecionar_base_custas = Select(campo_base_custas)
                selecionar_base_custas.select_by_visible_text(self.base_para_custas_conhecimento)

                # Tempo de controle
                time.sleep(1)

                # * Custas do Reclamante — Conhecimento
                if self.custas_do_reclamante == "Não se Aplica":
                    opcao_nao_aplica = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeCustasDeConhecimentoDoReclamante:0')))
                    opcao_nao_aplica.click()
                elif self.custas_do_reclamante == "Calculada 2%":
                    opcao_calculada = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeCustasDeConhecimentoDoReclamante:1')))
                    opcao_calculada.click()
                elif self.custas_do_reclamante == "Informada":
                    opcao_informada = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeCustasDeConhecimentoDoReclamante:2')))
                    opcao_informada.click()

                # Tempo de controle
                time.sleep(1)

                # * Custas do Reclamado
                if self.custas_reclamado_conhecimento == "Não se Aplica":
                    opcao_nao_aplica = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeCustasDeConhecimentoDoReclamado:0')))
                    opcao_nao_aplica.click()
                elif self.custas_reclamado_conhecimento == "Calculada 2%":
                    opcao_calculada = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeCustasDeConhecimentoDoReclamado:1')))
                    opcao_calculada.click()
                elif self.custas_reclamado_conhecimento == "Informada":
                    opcao_informada = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeCustasDeConhecimentoDoReclamado:2')))
                    opcao_informada.click()

        # Tempo de controle
        time.sleep(1)

        self.salvar_operacao(driver)
        # Aguardar Processamento
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(1)

        # Verificar status da operação
        # self.verificacao_custas_devidas(driver)

        # Teste de Nova Função de Verificação
        self.verificacao_cnae(driver)

    def definir_base_custas_conhecimento_liquidacao(self, driver):

        for i in range(self.tamanho_planilha_base):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                continue
            # Base para Custas de Conhecimento e Liquidação
            elif coluna_identificador == "base_para_custas_conhecimento":
                self.base_para_custas_conhecimento = coluna_informacao
                print('- Base para Custas de Conhecimento e Liquidação: ', self.base_para_custas_conhecimento)

                # Preencher os dados do PJeCalc
                # * Base para Custas de Conhecimento e Liquidação
                campo_base_custas = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:baseParaCustasCalculadas')))
                selecionar_base_custas = Select(campo_base_custas)
                selecionar_base_custas.select_by_visible_text(self.base_para_custas_conhecimento)

    def definir_correcao_monetaria_custas_reclamado_v3_34(self, driver):

        for i in range(self.tamanho_planilha_base):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                continue
            # Base para Custas de Conhecimento e Liquidação
            elif coluna_identificador == "custas_reclamado_conhecimento":
                self.custas_reclamado_conhecimento = coluna_informacao
                print("- Custas do Reclamado - Conhecimento: ", self.juros_correcao_monetaria)
            elif coluna_identificador == "custas_recda_conhecimento_info_vcto":
                self.custas_recda_conhecimento_info_vcto = coluna_informacao
                print("- Vencimento - Custas Devidas: ", self.custas_recda_conhecimento_info_vcto)
            elif coluna_identificador == "custas_recda_conhecimento_info_valor":
                self.custas_recda_conhecimento_info_valor = coluna_informacao
                print("- Valor - Custas Devidas: ", self.custas_recda_conhecimento_info_valor)

                # PJeCalc
                if self.custas_reclamado_conhecimento == "Não se Aplica":
                    WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:tipoDeCustasDeConhecimentoDoReclamado:0"))).click()
                elif self.custas_reclamado_conhecimento == "Calculada 2%":
                    WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:tipoDeCustasDeConhecimentoDoReclamado:1"))).click()
                elif self.custas_reclamado_conhecimento == "Informada":
                    WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:tipoDeCustasDeConhecimentoDoReclamado:2"))).click()
                    # Aguardar campo condicional
                    WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.ID, "formulario:dataVencimentoConhecimentoDoReclamadoInputDate")))
                    # Preencher valor do vencimento
                    campo_vencimento = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:dataVencimentoConhecimentoDoReclamadoInputDate")))
                    campo_vencimento.send_keys(self.custas_recda_conhecimento_info_vcto)
                    # Tempo de controle
                    time.sleep(1)
                    # Preencher valor
                    # Tratamento do valor
                    self.custas_recda_conhecimento_info_valor = f"{self.custas_recda_conhecimento_info_valor:.2f}"
                    campo_valor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:valorConhecimentoDoReclamado")))
                    campo_valor.send_keys(self.custas_recda_conhecimento_info_valor)

    def salvar_operacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:salvar'))).click()

    def verificacao_cnae(self, driver):

        delay = 5
        mensagem_sucesso = ""
        mensagem_erro = ""
        conteudo = ""
        # elemento_cancelar_status = ""
        # elemento_cancelar_status_2 = ""

        elemento_titulo = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'barraTitulo')))
        titulo_pagina = elemento_titulo.text
        # print('-- Título da Página Atual: ', titulo_pagina)
        titulo_pagina = titulo_pagina.replace(">", ":")

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            # file_txt_log.write('- ' + campo + ' | ' + status + '\n')
            file_txt_log.write(f'- {campo} : {titulo_pagina} | {status}\n')
            return file_txt_log.close()

        def gerar_relatorio_erro(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            # file_txt_log.write('- ' + campo + ' | ' + status + '\n')
            file_txt_log.write(f'- {campo} | {status}\n')
            return file_txt_log.close()

        try:
            mensagem_sucesso = WebDriverWait(driver, delay).until(
                EC.visibility_of_element_located((By.CLASS_NAME, "sucesso")))
        except TimeoutException:
            pass

        try:
            mensagem_erro = WebDriverWait(driver, delay).until(
                EC.visibility_of_element_located((By.CLASS_NAME, "erro")))
        except TimeoutException:
            pass

        if mensagem_sucesso:
            msg = mensagem_sucesso.text
            msg = msg.replace("\n", "")
            print("- ", msg)
            gerar_relatorio('Contribuição Social', 'Ok')
        elif mensagem_erro:
            msg = mensagem_erro.text
            print("!! - ", msg, " - !!")
            # Script
            elementos_erro = WebDriverWait(driver, delay).until(
                EC.visibility_of_all_elements_located((By.CLASS_NAME, "linkErro")))
            if elementos_erro:
                # print("- Tamanho da Lista: ", len(elementos_erro))
                for elemento in elementos_erro:
                    # Início Tratamento
                    erro = elemento.get_attribute("textContent")
                    erros = erro.split("//<!")
                    print(f"!! - {str(erros[0])} - !!")
                    conteudo = str(erros[0])
                    driver.execute_script(f"alert('{conteudo}')")
                    WebDriverWait(driver, delay).until(EC.alert_is_present())
                    alerta = Alert(driver)
                    time.sleep(4)
                    try:
                        alerta.accept()
                    except NoAlertPresentException:
                        continue
            gerar_relatorio_erro(f'Contribuição Social: {conteudo}', '---------- Erro! ----------')
            #
            try:
                elemento_cancelar_status = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.NAME, "formulario:cancelar")))
                elemento_cancelar_status.click()
                # print("1 - ", elemento_cancelar_status)
            except TimeoutException:
                pass

            try:
                elemento_cancelar_status_2 = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.NAME, "formulario:cancelarGeracao")))
                elemento_cancelar_status_2.click()
                # print("2 - ", elemento_cancelar_status_2)
            except TimeoutException:
                pass

        time.sleep(2)


    def verificacao_custas_devidas(self, driver):

        local = "Custas Devidas"

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            # file_txt_log.write('- ' + campo + ' | ' + status + '\n')
            file_txt_log.write(f'- {campo} : {local} | {status}\n')
            return file_txt_log.close()

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Custas Judiciais', 'Ok')
            else:
                # print('* ERRO!', msg)
                gerar_relatorio('Custas Judiciais', '---------- Erro! ----------')
        except TimeoutException:
            print('- [Except][Custas] - Página demorou a responder ou o elemento não encontrado. Encerrando...')

        # Tempo de controle
        time.sleep(2)


    def verificacao_custas_recolhidas(self, driver):

        local = "Custas Recolhidas"

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            # file_txt_log.write('- ' + campo + ' | ' + status + '\n')
            file_txt_log.write(f'- {campo} : {local} | {status}\n')
            return file_txt_log.close()

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Custas Judiciais', 'Ok')
            else:
                # print('* ERRO!', msg)
                gerar_relatorio('Custas Judiciais', '---------- Erro! ----------')

        except TimeoutException:
            print('- [Except][Custas] - Página demorou a responder ou o elemento não encontrado. Encerrando...')

        # Tempo de controle
        time.sleep(2)

    def acessar_aba_custas_recolhidas(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tabCustasPagas_lbl'))).click()

    def verificar_quantidadeCustasRecolhidas(self):

        for i in range(self.tamanho_planilha_base):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # - Pular Linhas Em Branco
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == "qtd_custas_recolhidas":
                self.qtd_custas_recolhidas = coluna_informacao
                # print(f"- QTD. CUSTAS JUDICIAIS RECOLHIDAS (Main): {self.qtd_custas_recolhidas} {type(self.qtd_custas_recolhidas)}")

                return self.qtd_custas_recolhidas


    def preencher_custas_recolhidas(self, driver):

        indice = 1

        for i in range(self.tamanho_planilha_base):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                continue
            elif coluna_identificador == "qtd_custas_recolhidas":
                qtd_custas_recolhidas = coluna_informacao
                self.qtd_custas_recolhidas = coluna_informacao
                print("- Quantidade - Custas Recolhidas: ", qtd_custas_recolhidas)
                if qtd_custas_recolhidas == 0:
                    break
            # Reclamado - Vencimento
            elif f'custas_recolhidas_rcdo{indice}_vencimento' in coluna_identificador:
                self.custas_recolhidas_rcdo1_vencimento = coluna_informacao

                if type(self.custas_recolhidas_rcdo1_vencimento) == type(self.var_controle_int):
                    self.custas_recolhidas_rcdo1_vencimento = xlrd.xldate_as_datetime(self.custas_recolhidas_rcdo1_vencimento, 0)
                    self.custas_recolhidas_rcdo1_vencimento = self.custas_recolhidas_rcdo1_vencimento.strftime('%d/%m/%Y')
                print(f'- Custas Recolhidas - Vencimento - {indice}: ', self.custas_recolhidas_rcdo1_vencimento)

            elif f'custas_recolhidas_rcdo{indice}_valor' in coluna_identificador:
                self.custas_recolhidas_rcdo1_valor = coluna_informacao

                if type(self.custas_recolhidas_rcdo1_valor) != type(self.var_controle_string):
                    self.custas_recolhidas_rcdo1_valor = float(self.custas_recolhidas_rcdo1_valor)
                    self.custas_recolhidas_rcdo1_valor = '{:.2f}'.format(self.custas_recolhidas_rcdo1_valor)
                print(f'- Custas Recolhidas - Valor - {indice}: ', self.custas_recolhidas_rcdo1_valor)

                # Preencher os dados coletados no PJeCalc
                for k in range(self.tamanho_planilha_base):

                    if '-' in self.custas_recolhidas_rcdo1_vencimento or '-' in self.custas_recolhidas_rcdo1_valor:
                        break
                    else:
                        # Preencher data de vencimento
                        campo_vencimento = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:dataVencimentoRDInputDate')))
                        campo_vencimento.send_keys(self.custas_recolhidas_rcdo1_vencimento)

                        # Tempo de controle
                        time.sleep(1)

                        campo_valor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:valorRD')))
                        campo_valor.send_keys(self.custas_recolhidas_rcdo1_valor)

                        # Tempo de controle
                        time.sleep(1)

                        # Adicionar
                        btn_add = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:cmdIncluirRD')))
                        btn_add.click()

                        # Aguardar Processamento
                        self.objTools.aguardar_carregamento(driver)

                        # Tempo de controle
                        time.sleep(1)

                        # Saída para coletar os próximos valores
                        break

                # Próximo índice - vencimento e valor
                indice += 1

        self.salvar_operacao(driver)

        self.objTools.aguardar_carregamento(driver)

        # Tempo de controle
        time.sleep(1)

        # Verificação
        # self.verificacao_custas_recolhidas(driver)

        # Nova Função de verificação
        self.verificacao_cnae(driver)

    def main_custas(self, driver):

        self.acessar_custas_judiciais(driver)
        # Aguardar Processamento
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.definir_correcao_monetaria_custas_reclamado_v3_34(driver)
        # Tempo de controle
        time.sleep(self.delayG)

        result_qtd = self.verificar_quantidadeCustasRecolhidas()
        if result_qtd > 0:
            self.acessar_aba_custas_recolhidas(driver)
            # Tempo de controle
            time.sleep(self.delayG)
            self.preencher_custas_recolhidas(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        # - Limpar Temp
        self.objTools.limparFilesTemp()

        print('-- Fim - (Custas Judiciais) --')