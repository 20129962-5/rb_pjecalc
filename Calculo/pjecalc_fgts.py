from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import gc
from Calculo.pjecalc_dados_calculo import DadosCalculo
from Tools.pjecalc_control import Control


class FGTS(DadosCalculo):


    def __init__(self, source):
        super().__init__(source)
        self.delay = 10
        self.delayG = 1.5
        self.objTools = Control()


    destino_fgts = ''
    compor_principal = ''
    multa_fgts = ''

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
                gerar_relatorio('FGTS', 'Ok')
            else:
                # print('* ERRO!', msg)
                gerar_relatorio('FGTS', '---------- Erro! ----------')
        except TimeoutException:
            print('- [Except][FGTS] - Elemento não encontrado/A Página demorou para responder. Encerrando...')

        # Tempo de controle
        time.sleep(2)

    def acessar_fgts(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageFgts"))).click()
        # WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:j_id46:0:j_id49:14:j_id1141')))

    def selecionar_destino(self, driver):


        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                # print('* Pulando linhas em branco ...')
                continue
            # Dados de FGTS
            elif coluna_identificador == 'destino_fgts':
                self.destino_fgts = coluna_informacao
                print('- Destino FGTS: ', self.destino_fgts)
                # Após encontrar o conteúdo, interromper o loop.
                break

        if self.destino_fgts == 'Pagar':
            opcao_pagar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeVerba:0')))
            checkbox_pagar = opcao_pagar.is_selected()
            print('- Status - Opção Pagar: ', checkbox_pagar)
            if checkbox_pagar:
                print('- OptionButton - Destino "Pagar" - Já Habilitado.')
            else:
                print('- OptionButton - Destino "Pagar" - Foi Habilitado.')
                opcao_pagar.click()
        else:
            opcao_recolher = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeVerba:1')))
            checkbox_recolher = opcao_recolher.is_selected()
            print('- Status - Opção Recolher: ', checkbox_recolher)
            if checkbox_recolher:
                print('- OptionButton - Destino "Recolher" - Já Habilitado.')
            else:
                print('- OptionButton - Destino "Recolher" - Foi Habilitado.')
                opcao_recolher.click()

    def selecionar_compor_principal(self, driver):


        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float) and type(coluna_informacao) == type(self.var_controle_float):
                # print('* Pulando linhas em branco ...')
                continue
            elif coluna_identificador == 'compor_principal':
                self.compor_principal = coluna_informacao
                print('- Compor Principal: ', self.compor_principal)
                # Após encontrar o elemento e atribuir o valor, encerrar o loop
                break


        if self.compor_principal == 'Sim':
            opcao_sim = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:comporPrincipal:0')))
            checkbox_sim = opcao_sim.is_selected()
            print('- Status - Compor Principal: ', checkbox_sim)
            if checkbox_sim:
                print('- OptionButton - Compor Princial "Sim" - Já Habilitado.')
            else:
                print('- OptionButton - Compor Princial "Sim" - Foi Habilitado.')
                opcao_sim.click()

        elif self.compor_principal == 'Não':
            opcao_nao = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:comporPrincipal:1')))
            checkbox_nao = opcao_nao.is_selected()
            print('- Status - Compor Principal: ', checkbox_nao)
            if checkbox_nao:
                print('- OptionButton - Compor Princial "Não" - Já Habilitado.')
            else:
                opcao_nao.click()

    def checkbox_multa(self, driver):

        # elementoCheckboxMulta = ""

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']
            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                # print('* Pulando linhas em branco ...')
                continue
            # Dados de FGTS
            elif coluna_identificador == 'multa_fgts':
                self.multa_fgts = coluna_informacao
                print('- Multa: ', self.multa_fgts)
                # Valor coletado, atribuído e loop encerrado.
                break

        if self.multa_fgts == 'Sim':
            try:
                elementoCheckboxMulta = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, 'formulario:multa')))
                checkbox_multas = elementoCheckboxMulta.is_selected()
                print('- Status - Multa: ', checkbox_multas)
                if checkbox_multas:
                    print('- Checkbox - Multa - Já Habilitado.')
                    # --- Nova Funcionalidade ---#
                    # Nova Funcionalidade — Desmarcar a opção 'Excluir da Base da Multa o valor de FGTS sobre Aviso prévio'
                    checkbox_excluir_base = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "formulario:excluirAvisoDaMulta")))
                    if checkbox_excluir_base.is_selected():
                        checkbox_excluir_base.click()
                        print("- Checkbox Desabilitado - 'Excluir da Base da Multa o valor de FGTS sobre Aviso prévio'")
                else:
                    elementoCheckboxMulta.click()
                    print('- Checkbox - Multa - Foi Habilitado.')
                    # --- Nova Funcionalidade ---#
                    # Nova Funcionalidade — Desmarcar a opção 'Excluir da Base da Multa o valor de FGTS sobre Aviso prévio'
                    checkbox_excluir_base = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "formulario:excluirAvisoDaMulta")))
                    if checkbox_excluir_base.is_selected():
                        checkbox_excluir_base.click()
                        print("- Checkbox Desabilitado - 'Excluir da Base da Multa o valor de FGTS sobre Aviso prévio'")
            except TimeoutException:
                pass
        else:
            try:
                campo_checkbox = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, 'formulario:multa')))
                checkbox_multas = campo_checkbox.is_selected()
                print('- Status - Multa: ', checkbox_multas)
                if checkbox_multas:
                    print('- Checkbox - Multa - Foi Desabilitado.')
                    campo_checkbox.click()
            except TimeoutException:
                pass

    def salvar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:salvar'))).click()

    def main_fgts(self, driver):

        self.acessar_fgts(driver)
        # Aguardar Processamento
        self.objTools.aguardar_carregamento(driver)
        # Tempo de Controle
        time.sleep(5)
        self.selecionar_destino(driver)
        # Tempo de Controle
        time.sleep(self.delayG)
        self.selecionar_compor_principal(driver)
        # Tempo de Controle
        time.sleep(self.delayG)
        self.checkbox_multa(driver)
        # Tempo de Controle
        time.sleep(self.delayG)
        self.salvar(driver)
        # Aguardar Processamento
        self.objTools.aguardar_carregamento(driver)
        # Tempo de Controle
        time.sleep(5)
        self.verificacao(driver)
        # Tempo de Controle
        time.sleep(self.delayG)

        # - Limpar Temp
        self.objTools.limparFilesTemp()
        gc.collect(generation=0)
        gc.collect(generation=1)
        gc.collect(generation=2)

        print('-- Fim - (FGTS) --')