import os
import gc
import xlrd
import pandas as pd
from time import sleep
from Tools.pjecalc_control import Control
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait, TimeoutException


class ContribuicaoSocial:


    def __init__(self, source):
        self.source = source
        self.planilha_base_inss = pd.read_excel(self.source, sheet_name='PJE HIST-VAL', header=3)
        # self.planilha_base = pd.read_excel(self.source, sheet_name='PJE-BD', header=1)
        self.planilha_base = pd.read_excel(self.source, sheet_name='PJE-BD', index_col=7)
        self.delay = 10
        self.delayG = 1.5
        self.objTools = Control()
        self.delayDefault = 0.7
        self.qtdTentativas = 3
        self.cont_social_apurar_segurado = ""
        self.cont_social_cobrar_reclamante = ""
        self.cont_social_com_correcao_trabalhista = ""
        self.cont_social_salario_pagos = ""
        self.cont_social_aliquota_segurado = ""
        self.cont_social_aliquota_empregador = ""
        self.cont_social_apurar_empresa = ""
        self.cont_social_apurar_sat = ""
        self.cont_social_apurar_terceiros = ""
        self.processo_de_agroindustria = ""
        self.empresa_enquadrada_simples = ""
        self.empresa_com_desoneracao = ""
        self.empresa_agro = ""
        self.empresa_desonerada = ""
        self.entidade_filantropica = ""


    def gerar_relatorio(self, campo, subatividade, status):
        file_txt_log = open(os.getcwd() + '\log.txt', "a")
        file_txt_log.write(f'- {campo} : {subatividade} | {status}\n')
        return file_txt_log.close()

    def mensagem_alert_frontend(self, driver, conteudo):
        driver.execute_script(f"alert('{conteudo}')")
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alerta = Alert(driver)
        sleep(4)
        try:
            alerta.accept()
        except NoAlertPresentException:
            pass


    def registrar_msg_log(self, area, subatividade, status):
        with open(f"{os.getcwd()}\log.txt", "a") as f:
            if subatividade != "":
                content = f"- {area} : {subatividade} | {status}\n"
            else:
                content = f"- {area} | {status}\n"
            f.write(content)
            f.close()

    def verificacao_cnae(self, driver):

        delay = 5
        mensagem_sucesso = ""
        mensagem_erro = ""
        conteudo = ""


        elemento_titulo = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'barraTitulo')))
        titulo_pagina = elemento_titulo.text
        titulo_pagina = titulo_pagina.replace(">", ":")

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo} : {titulo_pagina} | {status}\n')
            return file_txt_log.close()

        def gerar_relatorio_erro(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo} | {status}\n')
            return file_txt_log.close()

        try:
            mensagem_sucesso = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.CLASS_NAME, "sucesso")))
        except TimeoutException:
            pass

        try:
            mensagem_erro = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.CLASS_NAME, "erro")))
        except TimeoutException:
            pass

        if mensagem_sucesso:
            msg = mensagem_sucesso.text
            msg = msg.replace("\n", "")
            print("- ", msg)
            gerar_relatorio('Contribuição Social', 'Ok')
        elif mensagem_erro:
            # msg = mensagem_erro.text
            # Script
            elementos_erro = WebDriverWait(driver, delay).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, "linkErro")))
            if elementos_erro:
                # print("- Tamanho da Lista: ", len(elementos_erro))
                for elemento in elementos_erro:
                    # Início Tratamento
                    erro = elemento.get_attribute("textContent")
                    erros = erro.split("//<!")
                    conteudo = str(erros[0])
                    driver.execute_script(f"alert('{conteudo}')")
                    WebDriverWait(driver, delay).until(EC.alert_is_present())
                    alerta = Alert(driver)
                    sleep(3)
                    try:
                        alerta.accept()
                    except NoAlertPresentException:
                        continue
            gerar_relatorio_erro(f'Contribuição Social: {conteudo}', '---------- Erro! ----------')
            #
            try:
                elemento_cancelar_status = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.NAME, "formulario:cancelarGeracao")))
                elemento_cancelar_status.click()
            except TimeoutException:
                pass

            try:
                elemento_cancelar_status_2 = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.NAME, "formulario:cancelar")))
                elemento_cancelar_status_2.click()
            except TimeoutException:
                pass

        sleep(2)


    def clicar_btnCancelar(self, driver):
        try:
            field = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:cancelar"]')))
            field.click()
        except Exception as e:
            print(f"- [except][clicar_btnCancelar]: {e}")


    def verificar_statusOperacao(self, driver, subarea):

        try:
            elemento_msg = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//div[@id="divMensagem"]'))).get_attribute('textContent')
            print(f"- [STATUS_OPERACAO] {elemento_msg}")
            if 'sucesso' in elemento_msg:
                self.registrar_msg_log(f'Constribuição Social', f'{subarea}', 'Ok')
            elif 'erro' in elemento_msg or 'não pôde':
                self.registrar_msg_log(f'Constribuição Social', f'{subarea}', '---------- Erro! ----------')
                self.clicar_btnCancelar(driver)
                self.objTools.aguardar_carregamento(driver)
                sleep(self.delayDefault)
            else:
                self.registrar_msg_log(f'Constribuição Social', f'{subarea}', '---------- Erro! ----------')
        except Exception as e:
            print(f"- [except][verificar_statusOperacao]: {e}")
            self.registrar_msg_log('Constribuição Social', f'{subarea}', f'---------- Erro! ---------- : {e}')


    def verificacao(self, driver):

        delay = 10

        elemento_titulo = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'barraTitulo')))
        titulo_pagina = elemento_titulo.text

        titulo_pagina = titulo_pagina.replace(">", ":")

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo} : {titulo_pagina} | {status}\n')
            return file_txt_log.close()

        try:
            mensagem = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Contribuição Social', 'Ok')
            else:
                print('* ERRO!', msg)
                gerar_relatorio('Contribuição Social', '---------- Erro! ----------')
                try:
                    self.cancelar_operacao(driver)
                except TimeoutException:
                    pass
        except TimeoutException:
            print('- [Except][INSS] - Elemento não encontrado/A Página demorou para responder. Encerrando...')

        # Tempo de controle
        sleep(2)


    # [ACESSAR][NEW]
    def acessar_menuContribuicaoSocial(self, driver):

        for _ in range(self.qtdTentativas):
            try:
                btn_menuINSS = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//li[@id="li_calculo_inss"]//a[text()="Contribuição Social"]')))
                btn_menuINSS.click()
                return [True, '']
            except Exception as e:
                print(f"- [except][acessar_menuContribuicaoSocial]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][acessar_menuContribuicaoSocial]"
            print(f"- {msg}")
            return [False, msg]

    # [ACESSAR][OLD]
    def acessar_contribuicao_social(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageInss"))).click()


    # [SALARIOS_DEVIDO][NEW]
    # [1.APURAR_SEGURADO]
    def marcar_checkbox_apurarSeguro(self, driver, value):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:apurarInssSeguradoDevido"]')))
                if 'true' in value.lower() and field.is_enabled() and not field.is_selected():
                    field.click()
                print("- [APURAR_SEGURADO]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][marcar_checkbox_apurarSeguro]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][marcar_checkbox_apurarSeguro]"
            print(f"- {msg}")
            return [False, msg]

    def get_value_apurarSegurado_e_marcar(self, driver):
        cont_social_apurar_segurado = self.planilha_base.loc['cont_social_apurar_segurado', self.planilha_base.columns[7]]
        print(f"- [APURAR_SEGURADO]: {cont_social_apurar_segurado} | [TIPO]: {type(cont_social_apurar_segurado)}")
        # [TRATAR_VALORES_EM_BRANCO]
        if not pd.isna(cont_social_apurar_segurado):
            self.marcar_checkbox_apurarSeguro(driver, cont_social_apurar_segurado)

    # [2.COBRAR_DO_RECLAMANTE]
    def marcar_checkbox_cobrarDoReclamante(self, driver, value):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:cobrarDoReclamanteDevido"]')))
                if 'true' in value.lower() and field.is_enabled() and not field.is_selected():
                    field.click()
                print("- [COBRAR_DO_RECLAMANTE]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][marcar_checkbox_cobrarDoReclamante]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][marcar_checkbox_cobrarDoReclamante]"
            print(f"- {msg}")
            return [False, msg]

    def get_value_cobrarDoReclamante_e_marcar(self, driver):
        cont_social_cobrar_reclamante = self.planilha_base.loc['cont_social_cobrar_reclamante', self.planilha_base.columns[7]]
        print(f"- [COBRAR_DO_RECLAMANTE]: {cont_social_cobrar_reclamante} | [TIPO]: {type(cont_social_cobrar_reclamante)}")
        # [TRATAR_VALORES_EM_BRANCO]
        if not pd.isna(cont_social_cobrar_reclamante):
            self.marcar_checkbox_cobrarDoReclamante(driver, cont_social_cobrar_reclamante)

    # [3.COM_CORRECAO_TRABALHISTA]
    def marcar_checkbox_comCorrecaoTrabalhista(self, driver, value):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:corrigirDescontoReclamante"]')))
                if 'true' in value.lower() and field.is_enabled() and not field.is_selected():
                    field.click()
                print("- [COM_CORRECAO_TRABALHISTA]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][marcar_checkbox_comCorrecaoTrabalhista]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][marcar_checkbox_comCorrecaoTrabalhista]"
            print(f"- {msg}")
            return [False, msg]

    def get_value_comCorrecaoTrabalhista_e_marcar(self, driver):
        cont_social_com_correcao_trabalhista = self.planilha_base.loc['cont_social_com_correcao_trabalhista', self.planilha_base.columns[7]]
        print(f"- [COM_CORRECAO_TRABALHISTA]: {cont_social_com_correcao_trabalhista} | [TIPO]: {type(cont_social_com_correcao_trabalhista)}")
        # [TRATAR_VALORES_EM_BRANCO]
        if not pd.isna(cont_social_com_correcao_trabalhista):
            self.marcar_checkbox_comCorrecaoTrabalhista(driver, cont_social_com_correcao_trabalhista)

    # [SALARIO_DEVIDO][OLD]
    def selecionar_salarios_devidos(self, driver):

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == 'cont_social_apurar_segurado':
                self.cont_social_apurar_segurado = coluna_informacao
            elif coluna_identificador == 'cont_social_cobrar_reclamante':
                self.cont_social_cobrar_reclamante = coluna_informacao
            elif coluna_identificador == 'cont_social_com_correcao_trabalhista':
                self.cont_social_com_correcao_trabalhista = coluna_informacao


        if self.cont_social_apurar_segurado == 'True':
            elemento_apurar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:apurarInssSeguradoDevido')))
            checkbox_apurar = elemento_apurar.is_selected()
            if not checkbox_apurar:
                elemento_apurar.click()
        else:
            elemento_apurar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:apurarInssSeguradoDevido')))
            checkbox_apurar = elemento_apurar.is_selected()
            if checkbox_apurar:
                elemento_apurar.click()

        if self.cont_social_cobrar_reclamante == 'True':
            elemento_cobrar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:cobrarDoReclamanteDevido')))
            checkbox_cobrar = elemento_cobrar.is_selected()
            if not checkbox_cobrar:
                elemento_cobrar.click()
        else:
            elemento_cobrar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:cobrarDoReclamanteDevido')))
            checkbox_cobrar = elemento_cobrar.is_selected()
            if checkbox_cobrar:
                elemento_cobrar.click()

        if self.cont_social_com_correcao_trabalhista == 'True':
            elemento_com_correcao = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:corrigirDescontoReclamante')))
            checkbox_com_correcao = elemento_com_correcao.is_selected()
            if not checkbox_com_correcao:
                elemento_com_correcao.click()
        else:
            elemento_com_correcao = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:corrigirDescontoReclamante')))
            checkbox_com_correcao = elemento_com_correcao.is_selected()
            if checkbox_com_correcao:
                elemento_com_correcao.click()


    # [SALARIOS_PAGOS_APURAR][NEW]
    def marcar_checkbox_salariosPagos_apurar(self, driver, value):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:apurarSalariosPagos"]')))
                if 'true' in value.lower() and field.is_enabled() and not field.is_selected():
                    field.click()
                print("- [SALARIOS_PAGOS_APURAR]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][marcar_checkbox_salariosPagos_apurar]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][marcar_checkbox_salariosPagos_apurar]"
            print(f"- {msg}")
            return [False, msg]

    def get_value_salariosPagos_apurar_e_marcar(self, driver):
        cont_social_salario_pagos = self.planilha_base.loc['cont_social_salario_pagos', self.planilha_base.columns[7]]
        print(f"- [SALARIOS_PAGOS_APURAR]: {cont_social_salario_pagos} | [TIPO]: {type(cont_social_salario_pagos)}")
        # [TRATAR_VALORES_EM_BRANCO]
        if not pd.isna(cont_social_salario_pagos):
            self.marcar_checkbox_salariosPagos_apurar(driver, cont_social_salario_pagos)

    # [SALARIOS_PAGOS][OLD]
    def selecionar_salarios_pagos(self, driver):

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == 'cont_social_salario_pagos':
                self.cont_social_salario_pagos = coluna_informacao
                break

        if self.cont_social_salario_pagos == 'True':
            elemento_apurar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:apurarSalariosPagos')))
            checkbox_apurar = elemento_apurar.is_selected()
            if not checkbox_apurar:
                elemento_apurar.click()
        else:
            elemento_apurar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:apurarSalariosPagos')))
            checkbox_apurar = elemento_apurar.is_selected()
            if checkbox_apurar:
                elemento_apurar.click()


    # [SALVAR][NEW]
    def clicar_btnSalvar(self, driver):

        for _ in range(self.qtdTentativas):
            try:
                btn_salvar = WebDriverWait(driver, self.delay).until(
                    EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:salvar"]')))
                btn_salvar.click()
                return [True, '']
            except Exception as e:
                print(f"- [except][clicar_btnSalvar]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][clicar_btnSalvar]"
            print(f"- {msg}")
            return [False, msg]


    # [SALVAR][OLD]
    def salvar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.botao#formulario\:salvar'))).click()



    # [OCORRENCIAS][NEW]
    def clicar_btnOcorrencias(self, driver):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:ocorrencias"]')))
                field.click()
                return [True, '']
            except Exception:
                sleep(1)
                continue
        else:
            msg = "[tentativas_esgotadas][clicar_btnOcorrencias]"
            print(f"- {msg}")
            return [False, msg]

    # [OCORRENCIAS][OLD]
    def acessar_ocorrencias(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:ocorrencias'))).click()

    # [REGERAR][NEW]
    def clicar_btnRegerar(self, driver):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:regerar"]')))
                field.click()
                return [True, '']
            except Exception:
                sleep(1)
                continue
        else:
            msg = "[tempo_esgotado][clicar_btnRegerar]"
            print(f"- {msg}")
            return [False, msg]

    # [REGERAR][OLD]
    def regerar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:regerar'))).click()


    def clicar_opcaoSobrescrever(self, driver):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:tipoRegeracao:1"]')))
                field.click()
                return [True, '']
            except Exception:
                sleep(1)
                continue
        else:
            msg = "[tempo_esgotado][clicar_opcaoSobrescrever]"
            print(f"- {msg}")
            return [False, msg]


    def clicar_btnConfirmar(self, driver):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:confirmarGeracao"]')))
                field.click()
                return [True, '']
            except Exception:
                sleep(1)
                continue
        else:
            msg = "[tempo_esgotado][clicar_btnConfirmar]"
            print(f"- {msg}")
            return [False, msg]





    # [ALIQUOTA_SEGURADO][NEW]
    def marcar_checkbox_aliquotaSegurado(self, driver, value):


        # //input[@name="formulario:aliquotaEmpregado"][@value="SEGURADO_EMPREGADO"]
        # //input[@name="formulario:aliquotaEmpregado"][@value="EMPREGADO_DOMESTICO"]
        # //input[@name="formulario:aliquotaEmpregado"][@value="FIXA"]

        if 'empregado' in value.lower():
            value = "SEGURADO_EMPREGADO"
        elif 'doméstico' in value.lower():
            value = "EMPREGADO_DOMESTICO"
        else:
            value = "FIXA"


        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, f'//input[@name="formulario:aliquotaEmpregado"][@value="{value}"]')))
                if 'true' in value.lower() and field.is_enabled() and not field.is_selected():
                    field.click()
                print("- [ALIQUOTA_SEGURADO]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][marcar_checkbox_aliquotaSegurado]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][marcar_checkbox_aliquotaSegurado]"
            print(f"- {msg}")
            return [False, msg]

    def get_value_aliquotaSegurado_e_marcar(self, driver):
        cont_social_aliquota_segurado = self.planilha_base.loc['cont_social_aliquota_segurado', self.planilha_base.columns[7]]
        print(f"- [ALIQUOTA_SEGURADO]: {cont_social_aliquota_segurado} | [TIPO]: {type(cont_social_aliquota_segurado)}")
        # [TRATAR_VALORES_EM_BRANCO]
        if not pd.isna(cont_social_aliquota_segurado):
            self.marcar_checkbox_aliquotaSegurado(driver, cont_social_aliquota_segurado)


    # [ALIQUOTA_SEGURADO][OLD]
    def selecionar_aliquota_segurado(self, driver):

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == 'cont_social_aliquota_segurado':
                self.cont_social_aliquota_segurado = coluna_informacao
                print('- Alíquota Segurado: ', self.cont_social_aliquota_segurado)
                # Após encontrar e atribuir o valor, encerrar o loop
                break

        if self.cont_social_aliquota_segurado == 'Segurado Empregado':
            elemento_segurado = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:aliquotaEmpregado:0')))
            checkbox_segurado = elemento_segurado.is_selected()
            if checkbox_segurado:
                print('- Checkbox - "Segurado Empregado" - Já Habilitado.')
            else:
                print('- Checkbox - "Segurado Empregado" - Foi Habilitado.')
                elemento_segurado.click()

        elif self.cont_social_aliquota_segurado == 'Empregado Doméstico':
            elemento_empregado = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:aliquotaEmpregado:1')))
            checkbox_empregado = elemento_empregado.is_selected()
            if checkbox_empregado:
                print('- Checkbox - "Empregado Doméstico" - Já Habilitado.')
            else:
                print('- Checkbox - "Empregado Doméstico" - Foi Habilitado.')
                elemento_empregado.click()
        elif self.cont_social_aliquota_segurado == 'Fixa':
            elemento_fixa = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:aliquotaEmpregado:2')))
            checkbox_fixa = elemento_fixa.is_selected()
            if checkbox_fixa:
                print('- Checkbox - "Fixa" - Já Habilitado.')
            else:
                print('- Checkbox - "Fixa" - Foi Habilitado.')
                elemento_fixa.click()
        else:
            print('!! Parâmetros das Ocorrências -> Alíquota Segurado: Erro!!')


    # [ALIQUOTA_EMPREGADOR_FIXA][NEW]
    def marcar_optionButton_aliquotaEmpregador_fixa(self, driver, value):

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@name="formulario:aliquotaEmpregador"][@value="FIXA"]')))
                if 'true' in value.lower() and field.is_enabled() and not field.is_selected():
                    field.click()
                print("- [ALIQUOTA_EMPREGADOR_FIXA]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][marcar_optionButton_aliquotaEmpregador_fixa]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][marcar_optionButton_aliquotaEmpregador_fixa]"
            print(f"- {msg}")
            return [False, msg]

    def digitar_percentual_aliquotaEmpregador_fixa_empresa(self, driver, value):

        # "20,0000"
        valor_fixo_empresa = value

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:aliquotaEmpresaFixa"]')))
                field.clear()
                field.send_keys(valor_fixo_empresa)
                print("- [ALIQUOTA_EMPREGADOR_FIXA][PERCENTUAL_EMPRESA]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][digitar_percentual_aliquotaEmpregador_fixa_empresa]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][digitar_percentual_aliquotaEmpregador_fixa_empresa]"
            print(f"- {msg}")
            return [False, msg]

    def digitar_percentual_aliquotaEmpregador_fixa_SAT(self, driver, value):

        # "3,0000"

        valor_fixo_SAT = value

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:aliquotaRatFixa"]')))
                field.clear()
                field.send_keys(valor_fixo_SAT)
                print("- [ALIQUOTA_EMPREGADOR_FIXA][PERCENTUAL_SAT (%)]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][digitar_percentual_aliquotaEmpregador_fixa_SAT]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][digitar_percentual_aliquotaEmpregador_fixa_SAT]"
            print(f"- {msg}")
            return [False, msg]


    # [ALIQUOTA_EMPREGADOR_FIXA][OLD]
    def definar_aliquota_empregador_fixa(self, driver):

        val_fixo_empresa = "20,0000"
        val_fixo_sat = "3,0000"
        # Option Button
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:aliquotaEmpregador:2"))).click()
        sleep(0.5)
        # Empresa (%)
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:aliquotaEmpresaFixa"))).send_keys(val_fixo_empresa)
        sleep(0.5)
        # SAT (%)
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:aliquotaRatFixa"))).send_keys(val_fixo_sat)
        sleep(0.5)


    def definir_parametros_aliquotaEmpregador_fixa(self, driver, value):

        if 'true' in value.lower():
            # self.marcar_optionButton_aliquotaEmpregador_fixa(driver)
            self.marcar_optionButton_aliquotaEmpregador_fixa(driver, value)
            sleep(self.delayDefault)
            self.digitar_percentual_aliquotaEmpregador_fixa_empresa(driver, "0")
            sleep(self.delayDefault)
            self.digitar_percentual_aliquotaEmpregador_fixa_SAT(driver, "0")
            sleep(self.delayDefault)



    # [ATIVIDADE_ECONOMICA (CNAE)][NEW]

    # [ATIVIDADE_ECONOMICA][CHECKBOX][APURAR_EMPRESA][NEW]
    def marcar_checkbox_atividadeEconomica_apurarEmpresa(self, driver, value):

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:apurarEmpresaPorAtividade"]')))
                if 'true' in value.lower() and field.is_enabled() and not field.is_selected():
                    field.click()
                elif 'false' in value.lower() and field.is_enabled() and field.is_selected():
                    field.click()
                print("- [APURAR_EMPRESA]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][marcar_checkbox_atividadeEconomica_apurarEmpresa]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][marcar_checkbox_atividadeEconomica_apurarEmpresa]"
            print(f"- {msg}")
            return [False, msg]

    def get_value_atividadeEconomica_apurarEmpresa_e_marcar(self, driver):

        cont_social_apurar_empresa = self.planilha_base.loc['cont_social_apurar_empresa', self.planilha_base.columns[7]]
        print(f"- [APURAR_EMPRESA]: {cont_social_apurar_empresa}")
        # [TRATAR_VALORES_EM_BRANCO]
        if not pd.isna(cont_social_apurar_empresa):
            self.marcar_checkbox_atividadeEconomica_apurarEmpresa(driver, cont_social_apurar_empresa)


    # [ATIVIDADE_ECONOMICA][CHECKBOX][APURAR_SAT][NEW]
    def marcar_checkbox_atividadeEconomica_apurarSAT(self, driver, value):

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:apurarRATPorAtividade"]')))
                if 'true' in value.lower() and field.is_enabled() and not field.is_selected():
                    field.click()
                elif 'false' in value.lower() and field.is_enabled() and field.is_selected():
                    field.click()
                print("- [APURAR_SAT]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][marcar_checkbox_atividadeEconomica_apurarSAT]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][marcar_checkbox_atividadeEconomica_apurarSAT]"
            print(f"- {msg}")
            return [False, msg]

    def get_value_atividadeEconomica_apurarSAT_e_marcar(self, driver):

        cont_social_apurar_sat = self.planilha_base.loc['cont_social_apurar_sat', self.planilha_base.columns[7]]
        print(f"- [APURAR_SAT]: {cont_social_apurar_sat}")
        # [TRATAR_VALORES_EM_BRANCO]
        if not pd.isna(cont_social_apurar_sat):
            self.marcar_checkbox_atividadeEconomica_apurarSAT(driver, cont_social_apurar_sat)

    # [ATIVIDADE_ECONOMICA][CHECKBOX][APURAR_TERCEIROS][NEW]

    def marcar_checkbox_atividadeEconomica_apurarTerceiros(self, driver, value):

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:apurarTerceirosPorAtividade"]')))
                if 'true' in value.lower() and field.is_enabled() and not field.is_selected():
                    field.click()
                elif 'false' in value.lower() and field.is_enabled() and field.is_selected():
                    field.click()
                print("- [APURAR_TERCEIROS]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][marcar_checkbox_atividadeEconomica_apurarTerceiros]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][marcar_checkbox_atividadeEconomica_apurarTerceiros]"
            print(f"- {msg}")
            return [False, msg]

    def get_value_atividadeEconomica_apurarTerceiros_e_marcar(self, driver):

        cont_social_apurar_terceiros = self.planilha_base.loc['cont_social_apurar_terceiros', self.planilha_base.columns[7]]
        print(f"- [APURAR_TERCEIROS]: {cont_social_apurar_terceiros}")
        # [TRATAR_VALORES_EM_BRANCO]
        if not pd.isna(cont_social_apurar_terceiros):
            self.marcar_checkbox_atividadeEconomica_apurarTerceiros(driver, cont_social_apurar_terceiros)


    # [SELECIONAR_CNAE][OLD]
    def selecionar_atividade_economica_checkbox(self, driver):

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == 'cont_social_apurar_empresa':
                self.cont_social_apurar_empresa = coluna_informacao
                print('- Apurar Empresa: ', self.cont_social_apurar_empresa)

                # PJeCalc
                if self.cont_social_apurar_empresa == 'True':
                    elemento_apurar_empresa = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:apurarEmpresaPorAtividade')))
                    status_checkbox = elemento_apurar_empresa.is_selected()
                    if status_checkbox:
                        print('- Checkbox - "Apurar Empresa" - Já Habilitado.')
                    else:
                        print('- Checkbox - "Apurar Empresa" - Foi Habilitado.')
                        elemento_apurar_empresa.click()
                else:
                    elemento_apurar_empresa = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:apurarEmpresaPorAtividade')))
                    status_checkbox = elemento_apurar_empresa.is_selected()
                    if status_checkbox:
                        print('- Checkbox - "Apurar Empresa" - Foi Desabilitado.')
                        elemento_apurar_empresa.click()

            elif coluna_identificador == 'cont_social_apurar_sat':
                self.cont_social_apurar_sat = coluna_informacao
                print('- Apurar SAT: ', self.cont_social_apurar_sat)

                # PJeCalc
                if self.cont_social_apurar_sat == 'True':
                    elemento_apurar_sat = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:apurarRATPorAtividade')))
                    status_checkbox = elemento_apurar_sat.is_selected()
                    if status_checkbox:
                        print('- Checkbox - "Apurar SAT" - Já Habilitado.')
                    else:
                        print('- Checkbox - "Apurar SAT" - Foi Habilitado.')
                        elemento_apurar_sat.click()
                else:
                    elemento_apurar_sat = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:apurarRATPorAtividade')))
                    status_checkbox = elemento_apurar_sat.is_selected()
                    if status_checkbox:
                        print('- Checkbox - "Apurar Empresa" - Foi Desabilitado.')
                        elemento_apurar_sat.click()

            elif coluna_identificador == 'cont_social_apurar_terceiros':
                self.cont_social_apurar_terceiros = coluna_informacao
                print('- Apurar Terceiros: ', self.cont_social_apurar_terceiros)

                # PJeCalc
                if self.cont_social_apurar_terceiros == 'True':
                    elemento_apurar_terceiro = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:apurarTerceirosPorAtividade')))
                    status_checkbox = elemento_apurar_terceiro.is_selected()
                    if status_checkbox:
                        print('- Checkbox - "Apurar Terceiros" - Já Habilitado.')
                    else:
                        print('- Checkbox - "Apurar Terceiros" - Foi Habilitado.')
                        elemento_apurar_terceiro.click()
                else:
                    elemento_apurar_terceiro = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:apurarTerceirosPorAtividade')))
                    status_checkbox = elemento_apurar_terceiro.is_selected()
                    if status_checkbox:
                        print('- Checkbox - "Apurar Empresa" - Foi Desabilitado.')
                        elemento_apurar_terceiro.click()


    # [ATIVIDADE_ECONOMICA (CNAE)][NEW]
    def get_value_aliquotaEmpregador_porAtividadeEconomoca_cnae(self):
        cont_social_aliquota_empregador = self.planilha_base.loc['cont_social_aliquota_empregador', self.planilha_base.columns[7]]
        print(f"- [CNAE]: {cont_social_aliquota_empregador}")
        return cont_social_aliquota_empregador


    def marcar_optionButton_aliquotaEmpregador_porAtividadeEconomica_cnae(self, driver):

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@name="formulario:aliquotaEmpregador"][@value="POR_ATIVIDADE_ECONOMICA"]')))
                if field.is_enabled() and not field.is_selected():
                    field.click()
                print("- [ATIVIDADE_ECONOMICA (CNAE)]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][marcar_optionButton_aliquotaEmpregador_porAtividadeEconomica_cnae]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][marcar_optionButton_aliquotaEmpregador_porAtividadeEconomica_cnae]"
            print(f"- {msg}")
            return [False, msg]

    # [DIGITAR_E_SELECIONAR_CNAE]
    # TABELA: //table[@id="formulario:suggestionautoCompleteAtividades:suggest"]//tr//td

    def digitar_atividadeEconomica_cnae(self, driver, value):

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//textarea[@id="formulario:atividadesEconomicas"]')))
                field.clear()
                field.click()
                field.send_keys(value)
                print("- [CNAE]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][digitar_atividadeEconomica_cnae]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][digitar_atividadeEconomica_cnae]"
            print(f"- {msg}")
            return [False, msg]


    def clicar_atividadeEconomica_cnae_listado(self, driver, value):

        for _ in range(self.qtdTentativas):
            try:
                fields = WebDriverWait(driver, self.delay).until(EC.presence_of_all_elements_located((By.XPATH, '//table[@id="formulario:suggestionautoCompleteAtividades:suggest"]//tr//td')))
                print(f"- [QTD_ELEMENTOS_ENCONTRADO]: {len(fields)}")
                for item in fields:
                    conteudo_cnae = item.text
                    print(f"- [CNAE_LISTA_PJECALC]: {conteudo_cnae.lower()} | [CNAE_PLANILHA_BASE]: {value.lower()}")
                    if conteudo_cnae.lower() == value.lower():
                        item.click()
                        return [True, '']
                else:

                    for item in fields:
                        conteudo_cnae = item.text
                        print(f"- [CNAE_LISTA_PJECALC]: {conteudo_cnae.lower()} | [CNAE_PLANILHA_BASE]: {value.lower()}")
                        if value.lower() in conteudo_cnae.lower():
                            item.click()
                            return [True, '']

                    else:
                        msg = "[falha][cnae_nao_localizado][clicar_atividadeEconomica_cnae_listado]"
                        print(f"- {msg}")
                        return [False, msg]

            except Exception as e:
                print(f"- [except][clicar_atividadeEconomica_cnae_listado]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][clicar_atividadeEconomica_cnae_listado]"
            print(f"- {msg}")
            return [False, msg]


    def definir_aliquotaEmpregador_porAtividadeEconomica_cnae(self, driver):

        # [PEGAR_VALOR_CNAE]
        valor_retorno_cnae = self.get_value_aliquotaEmpregador_porAtividadeEconomoca_cnae()
        if not pd.isna(valor_retorno_cnae):
            self.marcar_optionButton_aliquotaEmpregador_porAtividadeEconomica_cnae(driver)
            sleep(self.delayDefault)
            self.digitar_atividadeEconomica_cnae(driver, valor_retorno_cnae)
            sleep(self.delayDefault)
            status_operacao = self.clicar_atividadeEconomica_cnae_listado(driver, valor_retorno_cnae)
            if status_operacao[0] is False:
                sleep(self.delayDefault)
                self.mensagem_alert_frontend(driver, 'Registro não encontrado. Será adicionado a Alíquota do Empregador Fixa.')
                self.registrar_msg_log('Constribuição Social', 'Ocorrências > Regerar', f'---------- Erro! ---------- : {status_operacao[1]}')
                # [ALIQOTA_EMPREGADOR_FIXA]
                self.definir_parametros_aliquotaEmpregador_fixa(driver, 'True')
                sleep(self.delayDefault)


    # [ATIVIDADE_ECONOMICA (CNAE)][ADICIONAR][OLD]
    def preencher_atividade_economica(self, driver):

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == 'cont_social_aliquota_empregador':
                self.cont_social_aliquota_empregador = coluna_informacao
                # Remover espaços em branco
                self.cont_social_aliquota_empregador = self.cont_social_aliquota_empregador.strip()
                print('- Atividade Econômica *: ', self.cont_social_aliquota_empregador)
                # Após identificar e atribuir o valor, encerrar o loop
                break

        # Selecionar Alíquota Empregador → Até o momento será sempre a opção 'Por Atividade Econômica'
        elemento_por_ativ_economica = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:aliquotaEmpregador:0')))
        status_checkbox = elemento_por_ativ_economica.is_selected()
        if status_checkbox:
            print('- Checkbox - "Por Atividade Econômica" - Já Habilitado.')
        else:
            print('- Checkbox - "Por Atividade Econômica" - Foi Habilitado.')
            elemento_por_ativ_economica.click()

        # Tempo de controle
        sleep(1)
        campo_atividade_economica = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:atividadesEconomicas')))
        campo_atividade_economica.send_keys(self.cont_social_aliquota_empregador)
        # Tempo de controle
        sleep(2)
        # Trecho de código novo
        indice = 1
        try:
            while True:
                listagemCnae = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, f"/html/body/form/div/div[3]/table/tbody/tr/td[3]/span/div/table[1]/tbody/tr[1]/td/span[3]/fieldset/fieldset/div[1]/div/fieldset/div[1]/span/table/tbody/tr[1]/td/span/div[1]/div[1]/div/table/tbody/tr/td/div/table/tbody/tr[{indice}]/td")))

                atividadeEconomicaListaCnae = listagemCnae.text
                if self.cont_social_aliquota_empregador.lower() == atividadeEconomicaListaCnae.lower():
                    # print("- Nomenclaruta Idênticas.")
                    listagemCnae.click()
                    # Tempo de controle
                    sleep(1)
                    # Selecionar Checkboxs da 'Alíquota Empregador'
                    self.selecionar_atividade_economica_checkbox(driver)
                    # Sair
                    break
                indice += 1
                continue
        except TimeoutException:
            indice = 1
            try:
                while True:
                    listagemCnae = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH,
                                                                                                  f"/html/body/form/div/div[3]/table/tbody/tr/td[3]/span/div/table[1]/tbody/tr[1]/td/span[3]/fieldset/fieldset/div[1]/div/fieldset/div[1]/span/table/tbody/tr[1]/td/span/div[1]/div[1]/div/table/tbody/tr/td/div/table/tbody/tr[{indice}]/td")))
                    atividadeEconomicaListaCnae = listagemCnae.text
                    if self.cont_social_aliquota_empregador.lower() in atividadeEconomicaListaCnae.lower():
                        # print("- Nomenclaruta Compatível.")
                        listagemCnae.click()
                        # Tempo de controle
                        sleep(1)
                        # Selecionar Checkboxs da 'Alíquota Empregador'
                        self.selecionar_atividade_economica_checkbox(driver)
                        # Sair
                        break
                    indice += 1
                    continue
            except TimeoutException:
                # Trecho de código para o caso do cnae não ser encontrado
                elemento = WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.ID, "formulario:suggestionautoCompleteAtividadesNothingLabel")))
                retorno_registro = elemento.get_attribute("textContent")
                print("- Retorno: ", retorno_registro)
                if "Registro não encontrado" in retorno_registro:
                    self.mensagem_alert_frontend(driver, "Registro não encontrado. Será adicionado a Alíquota do Empregador Fixa.")
                    self.definar_aliquota_empregador_fixa(driver)
                    file_txt_log = open(os.getcwd() + '\log.txt', "a")
                    conteudo_log = f"- Contribuição Social : CNAE : {retorno_registro} - Será adicionado a Alíquota do Empregador Fixa "
                    file_txt_log.write(conteudo_log)
                    file_txt_log.close()
                    # Sair
                else:
                    # ESSA FUNÇÃO SERÁ EXECUTADA APENAS SE O CONTEÚDO DO CNAE DA PLANILHA NÃO BATER COM O DA BASE DO PJECALC
                    self.definar_aliquota_empregador_fixa(driver)


    # [ALIQUTA_EMPREGADOR_AGRO][NEW]
    def get_value_processoAgroindustria(self):
        processo_de_agroindustria = self.planilha_base.loc['processo_de_agroindustria', self.planilha_base.columns[7]]
        print(f"- [processo_de_agroindustria]: {processo_de_agroindustria}")
        return processo_de_agroindustria

    def definir_parametros_processoAgroindustria(self, driver, value):

        if not pd.isna(value):
            if 'true' in value.lower():
                self.marcar_optionButton_aliquotaEmpregador_fixa(driver)
                sleep(self.delayDefault)
                self.digitar_percentual_aliquotaEmpregador_fixa_empresa(driver, "0")
                sleep(self.delayDefault)
                self.digitar_percentual_aliquotaEmpregador_fixa_SAT(driver, "0")
                sleep(self.delayDefault)

    # [ALIQUTA_EMPREGADOR_AGRO][OLD]
    def definir_aliquota_empregador_agro(self, driver):

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == "processo_de_agroindustria":
                self.processo_de_agroindustria = coluna_informacao
                print("- Processo de Agroindústria: ", self.processo_de_agroindustria)
                break

        if self.processo_de_agroindustria == "True":
            # Selecionar a opção Fixa
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:aliquotaEmpregador:2"))).click()
            sleep(1)
            # Empresa(%)
            campo_empresa = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:aliquotaEmpresaFixa")))
            campo_empresa.click()
            campo_empresa.send_keys("0")
            # SAT (%)
            campo_sat = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:aliquotaRatFixa")))
            campo_sat.click()
            campo_sat.send_keys("0")
            # Terceiros (%)


    # [ALIQUTA_EMPREGADOR_ENTIDADE_FILANTROPICA][NEW]
    def get_value_processoEntidadeFilantropica(self):
        entidade_filantropica = self.planilha_base.loc['entidade_filantropica', self.planilha_base.columns[7]]
        print(f"- [entidade_filantropica]: {entidade_filantropica}")
        return entidade_filantropica

    def definir_parametros_processoEntidadeFilantropica(self, driver, value):

        if not pd.isna(value):
            if 'true' in value.lower():
                self.marcar_optionButton_aliquotaEmpregador_fixa(driver)
                sleep(self.delayDefault)
                self.digitar_percentual_aliquotaEmpregador_fixa_empresa(driver, "0")
                sleep(self.delayDefault)
                self.digitar_percentual_aliquotaEmpregador_fixa_SAT(driver, "0")
                sleep(self.delayDefault)


    # [ALIQUTA_EMPREGADOR_ENTIDADE_FILANTROPICA][OLD]
    def definir_aliquota_empregador_entidade_filantropica(self, driver):

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == "entidade_filantropica":
                self.entidade_filantropica = coluna_informacao
                print("-Entidade Filantrópica: ", self.entidade_filantropica)
                break

        if self.entidade_filantropica == "True":
            # Selecionar a opção Fixa
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:aliquotaEmpregador:2"))).click()
            sleep(1)
            # Empresa(%)
            campo_empresa = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:aliquotaEmpresaFixa")))
            campo_empresa.click()
            campo_empresa.send_keys("0")
            # SAT (%)
            campo_sat = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:aliquotaRatFixa")))
            campo_sat.click()
            campo_sat.send_keys("0")
            # Terceiros (%)


    # [EMPRESA_ENQUADRA_SIMPLES][NEW]
    def get_value_empresaEnquadrada_simplesNacional_e_digitarPeriodo(self, driver):
        empresa_enquadrada_simples = self.planilha_base.loc['empresa_enquadrada_simples', self.planilha_base.columns[7]]
        print(f"- [empresa_enquadrada_simples]: {empresa_enquadrada_simples}")

        if not pd.isna(empresa_enquadrada_simples):
            if 'true' in empresa_enquadrada_simples.lower():

                planilha = pd.read_excel(self.source, sheet_name='PJE-DIV', header=20)

                for indice, value in enumerate(planilha['SN_INICIO']):
                    if not pd.isna(value):
                        dataInicial = xlrd.xldate.xldate_as_datetime(int(value), 0).strftime('%m/%Y')
                        dataFinal = int(planilha.loc[indice, 'SN_TERMINO'])
                        dataFinal = xlrd.xldate.xldate_as_datetime(dataFinal, 0).strftime('%m/%Y')
                        print(f"- [INICIO]: {dataInicial} | - [FIM]: {dataFinal}")
                        self.definir_periodoInicial_simplesNacional(driver, dataInicial)
                        sleep(self.delayDefault)
                        self.definir_periodoFinal_simplesNacional(driver, dataFinal)
                        sleep(self.delayDefault)
                        self.clicar_btnIncluirperiodo_simplesNacional(driver)
                        self.objTools.aguardar_carregamento(driver)
                        sleep(self.delayDefault)

                self.objTools.limparFilesTemp()

    def definir_periodoInicial_simplesNacional(self, driver, value):

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:dataInicioSimplesInputDate"]')))
                field.clear()
                field.click()
                field.send_keys(value)
                print("- [SIMPLES_NACIONAL_INICIO]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][definir_periodoInicial_simplesNacional]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][definir_periodoInicial_simplesNacional]"
            print(f"- {msg}")
            return [False, msg]

    def definir_periodoFinal_simplesNacional(self, driver, value):

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:dataTerminoSimplesInputDate"]')))
                field.clear()
                field.click()
                field.send_keys(value)
                print("- [SIMPLES_NACIONAL_FIM]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][definir_periodoFinal_simplesNacional]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][definir_periodoFinal_simplesNacional]"
            print(f"- {msg}")
            return [False, msg]

    def clicar_btnIncluirperiodo_simplesNacional(self, driver):

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//a[@id="formulario:cmdIncluirPeriodoSimples"]')))
                field.click()
                print("- [INCLUIR_PERIODO]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][clicar_btnIncluirperiodo_simplesNacional]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][clicar_btnIncluirperiodo_simplesNacional]"
            print(f"- {msg}")
            return [False, msg]

    # [EMPRESA_ENQUADRA_SIMPLES][OLD]
    def definir_periodos_simples_nacional(self, driver):

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == "empresa_enquadrada_simples":
                self.empresa_enquadrada_simples = coluna_informacao
                print("- Empresa enquadra no Simples Nacional: ", self.empresa_enquadrada_simples)
                break

        if self.empresa_enquadrada_simples == "True":

            planilha = pd.read_excel(self.source, sheet_name='PJE-DIV', header=20)

            for i in range(len(planilha)):
                inicio = planilha.loc[i, "SN_INICIO"]
                fim = planilha.loc[i, "SN_TERMINO"]

                if inicio > 0 and fim > 0:
                    comp_inicial = int(inicio)
                    comp_final = int(fim)
                    # Conversão para Datetime
                    comp_inicial = xlrd.xldate_as_datetime(comp_inicial, 0)
                    comp_final = xlrd.xldate_as_datetime(comp_final, 0)
                    # Conversão para Data str
                    comp_inicial_dt = comp_inicial.strftime("%m/%Y")
                    comp_fim_dt = comp_final.strftime("%m/%Y")
                    print(comp_inicial_dt, " - ", comp_fim_dt)
                    # Início
                    comp_inicial = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:dataInicioSimplesInputDate")))
                    comp_inicial.send_keys(comp_inicial_dt)
                    # Fim
                    comp_fim = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:dataTerminoSimplesInputDate")))
                    comp_fim.send_keys(comp_fim_dt)
                    # Adicionar
                    WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:cmdIncluirPeriodoSimples"))).click()
                    self.objTools.aguardar_carregamento(driver)
                    sleep(1)



    # [EMPRESA_DESONARADA][POR_PERIODO][NOVO]
    def clicar_aliquotaEmpregador_porPeriodo(self, driver):

        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@name="formulario:aliquotaEmpregador"][@value="POR_PERIODO"]')))
                field.click()
                print("- [EMPRESA_DESONARADA][POR_PERIODO]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][clicar_aliquotaEmpregador_porPeriodo]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][clicar_aliquotaEmpregador_porPeriodo]"
            print(f"- {msg}")
            return [False, msg]

    def digitar_aliquotaEmpregador_porPeriodo_inicio(self, driver, value):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@name="formulario:dataInicioPeriodoInputDate"]')))
                field.clear()
                field.click()
                field.send_keys(value)
                print("- [EMPRESA_DESONARADA][POR_PERIODO][INICIO]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][digitar_aliquotaEmpregador_porPeriodo_inicio]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][digitar_aliquotaEmpregador_porPeriodo_inicio]"
            print(f"- {msg}")
            return [False, msg]

    def digitar_aliquotaEmpregador_porPeriodo_fim(self, driver, value):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@name="formulario:dataTerminoPeriodoInputDate"]')))
                field.clear()
                field.click()
                field.send_keys(value)
                print("- [EMPRESA_DESONARADA][POR_PERIODO][FIM]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][digitar_aliquotaEmpregador_porPeriodo_fim]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][digitar_aliquotaEmpregador_porPeriodo_fim]"
            print(f"- {msg}")
            return [False, msg]

    def digitar_aliquotaEmpregador_porPeriodo_empresa(self, driver, value):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@name="formulario:aliquotaEmpresaPorPeriodo"]')))
                field.clear()
                field.click()
                field.send_keys(value)
                print("- [EMPRESA_DESONARADA][POR_PERIODO][EMPRESA]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][digitar_aliquotaEmpregador_porPeriodo_empresa]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][digitar_aliquotaEmpregador_porPeriodo_empresa]"
            print(f"- {msg}")
            return [False, msg]

    def digitar_aliquotaEmpregador_porPeriodo_SAT(self, driver, value):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@name="formulario:aliquotaRatPorPeriodo"]')))
                field.clear()
                field.click()
                field.send_keys(value)
                print("- [EMPRESA_DESONARADA][POR_PERIODO][SAT]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][digitar_aliquotaEmpregador_porPeriodo_SAT]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][digitar_aliquotaEmpregador_porPeriodo_SAT]"
            print(f"- {msg}")
            return [False, msg]

    def clicar_btnIncluir_aliquotaEmpregador_porPeriodo(self, driver):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//a[@id="formulario:cmdIncluirPorPeriodo"]')))
                field.click()
                print("- [EMPRESA_DESONARADA][BTN_INCLUIR]: [OK]")
                return [True, '']
            except Exception as e:
                print(f"- [except][clicar_btnIncluir_aliquotaEmpregador_porPeriodo]: {e}")
                sleep(1)
        else:
            msg = "[tentativas_esgotadas][clicar_btnIncluir_aliquotaEmpregador_porPeriodo]"
            print(f"- {msg}")
            return [False, msg]

    def get_value_empresaComDesoneracao(self):

        empresa_com_desoneracao = self.planilha_base.loc['empresa_com_desoneracao', self.planilha_base.columns[7]]
        print(f"- [empresa_com_desoneracao]: {empresa_com_desoneracao}")
        return empresa_com_desoneracao

    def definir_aliquotaEmpregador_porPeriodo_empresaComDesoneracao(self, driver, value):

        if not pd.isna(value):
            if 'true' in value.lower():

                self.clicar_aliquotaEmpregador_porPeriodo(driver)
                sleep(self.delayDefault)

                planilha_desoneracao = pd.read_excel(self.source, sheet_name='PJE-DIV', header=7)

                for indice, value in enumerate(planilha_desoneracao['DES_INICIO']):
                    if not pd.isna(value):
                        if isinstance(value, int):
                            inicio = value
                            fim = planilha_desoneracao.loc[indice, "DES_TERMINO"]
                            status_desoneracao = planilha_desoneracao.loc[indice, "DES"]
                            empresa_percentual = planilha_desoneracao.loc[indice, "EMP"]
                            sat_percentual = planilha_desoneracao.loc[indice, "SAT (%)"]
                            print(f"[1] - [INICIO]: {inicio} | [FIM]: {fim} | [STATUS]: {status_desoneracao} | [EMPRESA]: {empresa_percentual} | [SAT]: {sat_percentual}")
                            inicio = xlrd.xldate_as_datetime(inicio, 0).strftime('%m/%Y')
                            fim = xlrd.xldate_as_datetime(fim, 0).strftime('%m/%Y')
                            empresa_percentual = f"{empresa_percentual:.4%}".replace("%", "").replace(".", ",")
                            sat_percentual = f"{sat_percentual:.4%}".replace("%", "").replace(".", ",")
                            print(f"[2] - [INICIO]: {inicio} | [FIM]: {fim} | [STATUS]: {status_desoneracao} | [EMPRESA]: {empresa_percentual} | [SAT]: {sat_percentual}")
                            self.digitar_aliquotaEmpregador_porPeriodo_inicio(driver, inicio)
                            sleep(self.delayDefault)
                            self.digitar_aliquotaEmpregador_porPeriodo_fim(driver, fim)
                            sleep(self.delayDefault)
                            self.digitar_aliquotaEmpregador_porPeriodo_empresa(driver, empresa_percentual)
                            sleep(self.delayDefault)
                            self.digitar_aliquotaEmpregador_porPeriodo_SAT(driver, sat_percentual)
                            sleep(self.delayDefault)
                            self.clicar_btnIncluir_aliquotaEmpregador_porPeriodo(driver)
                            self.objTools.aguardar_carregamento(driver)
                            sleep(self.delayDefault)

                self.objTools.limparFilesTemp()

    # [EMPRESA_DESONARADA][OLD] [PAREI_AQUI]
    def definir_empresa_desonerada(self, driver):

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == "empresa_com_desoneracao":
                self.empresa_com_desoneracao = coluna_informacao
                print("- Empresa com Desoneração: ", self.empresa_com_desoneracao)
                break

        if self.empresa_com_desoneracao == "True":

            # Selecionar a opção Por Período
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:aliquotaEmpregador:1"))).click()

            planilha = pd.read_excel(self.source, sheet_name='PJE-DIV', header=7)
            var_str = ""
            for i in range(len(planilha)):
                inicio = planilha.loc[i, "DES_INICIO"]
                fim = planilha.loc[i, "DES_TERMINO"]
                des = planilha.loc[i, "DES"]
                emp = planilha.loc[i, "EMP"]
                sat = planilha.loc[i, "SAT (%)"]

                if type(inicio) != type(var_str):

                    if inicio > 0 and fim > 0:
                        if des == "SIM":
                            emp = f"{emp:.4%}"
                            sat = f"{sat:.4%}"
                            print("- Empresa (%): ", emp, " -- SAT (%): ", sat)
                        else:
                            emp = f"{emp:.4%}"
                            sat = f"{sat:.4%}"
                            print("- Empresa (%): ", emp, " -- SAT (%): ", sat)
                        # Continua
                        comp_inicial = int(inicio)
                        comp_final = int(fim)
                        # Conversão para Datetime
                        comp_inicial = xlrd.xldate_as_datetime(comp_inicial, 0)
                        comp_final = xlrd.xldate_as_datetime(comp_final, 0)
                        # Conversão para Data str
                        comp_inicial_dt = comp_inicial.strftime("%m/%Y")
                        comp_fim_dt = comp_final.strftime("%m/%Y")
                        print(comp_inicial_dt, " - ", comp_fim_dt)

                        # INÍCIO
                        campo_inicio = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:dataInicioPeriodoInputDate")))
                        campo_inicio.send_keys(comp_inicial_dt)
                        # FIM
                        campo_fim = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:dataTerminoPeriodoInputDate")))
                        campo_fim.send_keys(comp_fim_dt)
                        # EMPRESA
                        campo_empresa = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:aliquotaEmpresaPorPeriodo")))
                        campo_empresa.send_keys(emp)
                        # SAT
                        campo_sat = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:aliquotaRatPorPeriodo")))
                        campo_sat.send_keys(sat)
                        # Adicionar
                        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:cmdIncluirPorPeriodo"))).click()
                        self.objTools.aguardar_carregamento(driver)
                        sleep(1)

    def definir_aliquota_empregador_v3_34(self, driver):

        for i in range(len(self.planilha_base)):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            elif coluna_identificador == "processo_de_agroindustria":
                self.empresa_agro = coluna_informacao
                print("- Processo de Agroidústria: ", self.empresa_agro)
            elif coluna_identificador == "empresa_com_desoneracao":
                self.empresa_desonerada = coluna_informacao
                print("- Empresa com Desoneração: ", self.empresa_desonerada)
            elif coluna_identificador == "entidade_filantropica":
                self.entidade_filantropica = coluna_informacao
                print("- Empresa com Desoneração: ", self.entidade_filantropica)

                # Se a empresa é agro
                if self.empresa_agro == "True":
                    self.definir_aliquota_empregador_agro(driver)
                    sleep(1)
                # Se a empresa é desonerada
                elif self.empresa_desonerada == "True":
                    self.definir_empresa_desonerada(driver)
                    sleep(1)
                elif self.entidade_filantropica == "True":
                    self.definir_aliquota_empregador_entidade_filantropica(driver)
                    sleep(1)
                else:
                    # Senão, CNAE
                    self.preencher_atividade_economica(driver)
                    sleep(1)
                    # self.selecionar_atividade_economica_checkbox(driver)
                    # sleep(1)

                # Simples Nacional
                self.definir_periodos_simples_nacional(driver)
                sleep(1)

    def confirmar_operacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:confirmarGeracao'))).click()

    def cancelar_operacao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:cancelarGeracao'))).click()

    def clicar_btnCancelar(self, driver):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.XPATH, '//input[@id="formulario:cancelarGeracao"]')))
                field.click()
                return [True, '']
            except Exception:
                sleep(1)
                continue
        else:
            msg = "[tempo_esgotado][clicar_btnCancelar]"
            print(f"- {msg}")
            return [False, msg]


    def preencher_salarios_pagos_historico(self, driver, admissao, rescisao, inicio_calculo, termino_calculo):

        # Condição para checar intervalo do cálculo. Não pode passar de 30 anos.
        print(f"- Data de Admissão: {admissao}\n- Data de Demissão: {rescisao}\n- Data Inicial do Cálculo: {inicio_calculo}\n- Data Final do Cálculo: {termino_calculo}\n")
        script = "alert('O escopo do cálculo é superior ao limite da planilha base (30 anos). Irei ignorar esta operação.')"
        resultado = 0

        if inicio_calculo != "" and termino_calculo != "":
            resultado = termino_calculo.year - inicio_calculo.year
            print("- Início e Término do cálculo preenchidos!")
            print(f"- Escopo do Cálculo: {resultado} anos.")

        elif inicio_calculo == "" and termino_calculo == "":
            resultado = rescisao.year - admissao.year
            print("- Início e Término do cálculo vazio!")
            print(f"- Escopo do Cálculo: {resultado} anos.")

        elif inicio_calculo == "":
            resultado = termino_calculo.year - admissao.year
            print("- Data de Cálculo Inicial vazio!")
            print(f"- Escopo do Cálculo: {resultado} anos.")

        elif termino_calculo == "":
            resultado = rescisao.year - inicio_calculo.year
            print("- Data de Término do Cálculo vazio!")
            print(f"- Escopo do Cálculo: {resultado} anos.")

        print("\n - Valor Final do Escopo do Cálculo: ", resultado)

        if resultado > 30:
            driver.execute_script(script)
            WebDriverWait(driver, self.delay).until(EC.alert_is_present())
            alerta = Alert(driver)
            self.gerar_relatorio("Contribuição Social", "O intervalo do Cálculo é superior ao limite da planilha (30 anos)", "---------- Erro! ----------")
            sleep(5)
            alerta.accept()
        else:

            indice = 0
            # Coletar a quantidade de competência na coluna do Excel AA5
            quantidade_compentencias = self.planilha_base_inss.iloc[0, 26]
            print("- Quantidade de Competências: ", quantidade_compentencias)
            # Loop para percorrer a coluna AB do Excel
            for k in range(1, quantidade_compentencias + 1):
                col_base_inss = self.planilha_base_inss.iloc[k, 27]
                # Preencher PJeCalc
                try:
                    col_base_inss = f"{col_base_inss:.2f}"
                    campo_salarios_pagos = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, f'formulario:listagemOcorrenciasDevidos:{indice}:baseHistoricoDevido')))
                    campo_salarios_pagos.click()
                    campo_salarios_pagos.send_keys(Keys.CONTROL, "a")
                    campo_salarios_pagos.send_keys(col_base_inss)
                    indice += 1
                except TimeoutException:
                    sleep(1.5)
                    break

            # Tempo de controle
            sleep(1)
            self.salvar(driver)
            self.objTools.aguardar_carregamento(driver)
            # Tempo de controle
            sleep(1)
            self.verificacao(driver)
            # Tempo de controle
            sleep(1.5)




    # [DIGITAR_VALORES_COLUNA_SALARIOS_PAGOS][NEW]
    def calcular_escopoDoCalculo_porAno(self, admissao, demissao, inicioCalc, fimCalc):
        qdtAnos = 0

        print(f"- [ADMISSAO]: {admissao} {type(admissao)}")
        print(f"- [DEMISSAO]: {demissao} {type(demissao)}")
        print(f"- [INICIO_CALCULO]: {inicioCalc} {type(inicioCalc)}")
        print(f"- [TERMINO_CALCULO]: {fimCalc} {type(fimCalc)}")

        try:
            if inicioCalc != "" and fimCalc != "":
                qdtAnos = fimCalc.year - inicioCalc.year
                print("- [1]")
            elif admissao != "" and demissao != "":
                qdtAnos = demissao.year - admissao.year
                print("- [2]")
            elif demissao == "":
                qdtAnos = fimCalc.year - admissao.year
                print("- [3]")
            elif admissao == "":
                qdtAnos = demissao.year - inicioCalc.year
                print("- [4]")
            else:
                print("- [FALHA_AO_CALCULAR_ESCOPO_DO_CALCULO]")
        except Exception as e:
            print(f"- [except][FALHA_AO_CALCULAR_ESCOPO_DO_CALCULO]: {e}")

        print(f"- [ESCOPO_DO_CALCULO]: {qdtAnos} anos.")
        return qdtAnos



    def digitar_valor_salariosPagos(self, driver, indice, valor):
        for _ in range(self.qtdTentativas):
            try:
                field = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, f'//input[@name="formulario:listagemOcorrenciasDevidos:{indice}:baseHistoricoDevido"]')))
                field.click()
                # field.clear()
                field.send_keys(valor)
                return [True, '']
            except Exception:
                sleep(1)
                continue
        else:
            print("- [tempo_esgotado][digitar_valor_salariosPagos]")
            return [False, '']


    def get_value_salariosPagos_historico_e_digitar(self, driver):

        planilha_base_inss = pd.read_excel(self.source, sheet_name='PJE HIST-VAL', header=3)
        # [PEGAR_TOTAL_ITENS][COLUNA_NA_COLUNA_DA_PLANILHA_AA5]
        qtd_registros = planilha_base_inss.iloc[0, 26]
        # [ALTERAR_O_INDICE_DA_LINHA]
        planilha_base_inss = planilha_base_inss.iloc[2:].reset_index(drop=True)
        planilha_base_inss.columns = planilha_base_inss.iloc[0]
        print(f"- [MES/ANO]{' ' * 2} | {' ' * 2} [VALOR]")
        print(f"-----------------------------------")
        for indice in range(qtd_registros):

            competencia_base_inss = planilha_base_inss.iloc[indice, 26]
            valor_base_inss = planilha_base_inss.iloc[indice, 27]

            if isinstance(competencia_base_inss, int):
                competencia_base_inss = xlrd.xldate_as_datetime(competencia_base_inss, 0).strftime('%m/%Y')
            valor_base_inss = f"{valor_base_inss:_.2f}".replace(".", ",").replace("_", ".")
            print(f"- {competencia_base_inss} {' ' * 3} | {' ' * 3} {valor_base_inss}")
            status_operacao = self.digitar_valor_salariosPagos(driver, indice, valor_base_inss)
            if status_operacao[0] is False:
                return [False, '']
        self.objTools.limparFilesTemp()
        sleep(self.delayDefault)
        return [True, '']



    # [DIGITAR_VALORES_COLUNA_SALARIOS_PAGOS][OLD]
    def preencher_salarios_pagos_historico_v34(self, driver, admissao, rescisao, inicio_calculo, termino_calculo):

        print(f"- Data de Admissão: {admissao}\n- Data de Demissão: {rescisao}\n- Data Inicial do Cálculo: {inicio_calculo}\n- Data Final do Cálculo: {termino_calculo}\n")
        script = "alert('O escopo do cálculo é superior ao limite da planilha base (30 anos). Irei ignorar esta operação.')"
        resultado = 0

        if inicio_calculo != "" and termino_calculo != "":
            resultado = termino_calculo.year - inicio_calculo.year
            print("- Início e Término do cálculo preenchidos!")
            print(f"- Escopo do Cálculo: {resultado} anos.")

        elif inicio_calculo == "" and termino_calculo == "":
            resultado = rescisao.year - admissao.year
            print("- Início e Término do cálculo vazio!")
            print(f"- Escopo do Cálculo: {resultado} anos.")

        elif inicio_calculo == "":
            resultado = termino_calculo.year - admissao.year
            print("- Data de Cálculo Inicial vazio!")
            print(f"- Escopo do Cálculo: {resultado} anos.")

        elif termino_calculo == "":
            resultado = rescisao.year - inicio_calculo.year
            print("- Data de Término do Cálculo vazio!")
            print(f"- Escopo do Cálculo: {resultado} anos.")

        # print("\n - Valor Final do Escopo do Cálculo: ", resultado)

        if resultado > 30:
            driver.execute_script(script)
            WebDriverWait(driver, self.delay).until(EC.alert_is_present())
            alerta = Alert(driver)
            self.gerar_relatorio("Contribuição Social",
                                 "O intervalo do Cálculo é superior ao limite da planilha (30 anos)",
                                 "---------- Erro! ----------")
            sleep(5)
            alerta.accept()
        else:
            indice = 0
            # Coletar a quantidade de competência na coluna do Excel AA5
            quantidade_compentencias = self.planilha_base_inss.iloc[0, 26]
            print("- Quantidade de Competências 1: ", quantidade_compentencias)
            quantidade_compentencias = int(quantidade_compentencias)
            print("- Quantidade de Competências 2: ", quantidade_compentencias)
            # Loop para percorrer a coluna AB do Excel
            for k in range(2, quantidade_compentencias + 2):
                col_base_inss = self.planilha_base_inss.iloc[k, 27]
                # Preencher PJeCalc
                try:
                    col_base_inss = f"{col_base_inss:.2f}"
                    # //input[@id="formulario:listagemOcorrenciasDevidos:0:baseHistoricoDevido"]
                    campo_salarios_pagos = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, f'//input[@id="formulario:listagemOcorrenciasDevidos:{indice}:baseHistoricoDevido"]')))
                    # campo_salarios_pagos = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, f'formulario:listagemOcorrenciasDevidos:{indice}:baseHistoricoDevido')))
                    campo_salarios_pagos.clear()
                    campo_salarios_pagos.send_keys(col_base_inss)
                    indice += 1
                except Exception as e:
                    print(f"- [except][INSS][preencher_salarios_pagos]: {e}")
                    sleep(1.5)
                    break

            # sleep(2.5)
            self.salvar(driver)
            sleep(1.5)


    def main_contribuicao_social_bkp(self, driver, admissao, rescisao, inicio_calculo, termino_calculo):

        self.acessar_menuContribuicaoSocial(driver)
        self.objTools.aguardar_carregamento(driver)
        sleep(self.delayDefault)

        # Aguardar Processamento
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        sleep(self.delayG)
        self.selecionar_salarios_devidos(driver)
        # Tempo de controle
        sleep(self.delayG)
        self.selecionar_salarios_pagos(driver)
        # Tempo de controle
        sleep(self.delayG)
        self.salvar(driver)
        # Aguardar Processamento
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        sleep(self.delayG)
        # Nova Função de Verificação
        self.verificacao_cnae(driver)
        # Tempo de controle
        sleep(self.delayG)
        self.acessar_ocorrencias(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        sleep(self.delayG)
        self.regerar(driver)
        # Tempo de controle
        sleep(self.delayG)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        sleep(self.delayG)
        self.selecionar_aliquota_segurado(driver)
        # Tempo de controle
        sleep(self.delayG)
        # Para planilha base v3.34 - Início
        self.definir_aliquota_empregador_v3_34(driver)
        # Tempo de controle
        sleep(self.delayG)
        self.confirmar_operacao(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        sleep(self.delayG)
        self.verificacao_cnae(driver)
        # Tempo de controle
        sleep(self.delayG)
        # Para planilha base v3.34 - Início
        try:
            self.preencher_salarios_pagos_historico_v34(driver, admissao, rescisao, inicio_calculo, termino_calculo)
        except Exception as e:
            print(f"- [except][inss]: {e}")
            driver.refresh()
            sleep(3)
            self.preencher_salarios_pagos_historico_v34(driver, admissao, rescisao, inicio_calculo, termino_calculo)
        # Para planilha base v3.34 - Fim
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        sleep(self.delayG)

        # - Limpar Temp
        self.objTools.limparFilesTemp()
        gc.collect(generation=0)
        gc.collect(generation=1)
        gc.collect(generation=2)

        print('-- Fim - (Contribuição Social) --')

    def main_contribuicao_social(self, driver, admissao, rescisao, inicio_calculo, termino_calculo):

        print('# ========== [CONTRIBUICAO_SOCIAL] ========== #')

        self.acessar_menuContribuicaoSocial(driver)
        self.objTools.aguardar_carregamento(driver)
        sleep(self.delayDefault)

        # [CONTRIBUICAO_SOCIAL_SOBRE_SALARIOS_DEVIDOS]
        self.get_value_apurarSegurado_e_marcar(driver)
        sleep(self.delayDefault)
        self.get_value_cobrarDoReclamante_e_marcar(driver)
        sleep(self.delayDefault)
        self.get_value_comCorrecaoTrabalhista_e_marcar(driver)
        sleep(self.delayDefault)

        # [CONTRIBUICAO_SOCIAL_SOBRE_SALARIOS_PAGOS]
        self.get_value_salariosPagos_apurar_e_marcar(driver)
        sleep(self.delayDefault)
        self.clicar_btnSalvar(driver)
        self.objTools.aguardar_carregamento(driver)
        sleep(self.delayDefault)
        self.verificar_statusOperacao(driver, '')
        sleep(self.delayDefault)

        # [CONTRIBUICAO_SOCIAL_OCORRENCIAS]
        self.clicar_btnOcorrencias(driver)
        self.objTools.aguardar_carregamento(driver)
        sleep(self.delayDefault)

        # [CONTRIBUICAO_SOCIAL_REGERAR]
        self.clicar_btnRegerar(driver)
        self.objTools.aguardar_carregamento(driver)
        sleep(self.delayDefault)

        # [CONTRIBUICAO_SOCIAL_ALIQUOTA_SEGURADO]
        self.get_value_aliquotaSegurado_e_marcar(driver)
        sleep(self.delayDefault)

        # [CONTRIBUICAO_SOCIAL_ALIQUOTA_EMPREGADOR]
        valor_retorno_processoAgroindustria = self.get_value_processoAgroindustria()
        print(f"- [PROCESSO_AGROINDUSTRIA]: {valor_retorno_processoAgroindustria}")
        if not pd.isna(valor_retorno_processoAgroindustria) and 'true' in valor_retorno_processoAgroindustria.lower():
            self.definir_parametros_processoAgroindustria(driver, valor_retorno_processoAgroindustria)
            sleep(self.delayDefault)
        else:
            valor_retorno_empresaComDesoneracao = self.get_value_empresaComDesoneracao()
            print(f"- [EMPRESA_COM_DESONERACAO]: {valor_retorno_empresaComDesoneracao}")
            if not pd.isna(valor_retorno_empresaComDesoneracao) and 'true' in valor_retorno_empresaComDesoneracao.lower():
                self.definir_aliquotaEmpregador_porPeriodo_empresaComDesoneracao(driver, valor_retorno_empresaComDesoneracao)
                sleep(self.delayDefault)
            else:
                valor_retorno_processo_entidade_filantropica = self.get_value_processoEntidadeFilantropica()
                print(f"- [PROCESSO_ENTIDADE_FILANTROPICA]: {valor_retorno_processo_entidade_filantropica}")
                if not pd.isna(valor_retorno_processo_entidade_filantropica) and 'true' in valor_retorno_processo_entidade_filantropica.lower():
                    self.definir_parametros_processoEntidadeFilantropica(driver, valor_retorno_processo_entidade_filantropica)
                    sleep(self.delayDefault)
                else:
                    print(f"- [PROCESSO_POR_ATIVIDADE_ECONOMIDA]")
                    self.definir_aliquotaEmpregador_porAtividadeEconomica_cnae(driver)
                    sleep(self.delayDefault)

        # [EMPRESA_ENQUADRADA_SIMPLES_NACIONAL]
        self.get_value_empresaEnquadrada_simplesNacional_e_digitarPeriodo(driver)
        sleep(self.delayDefault)

        # [CONTRIBUICAO_SOCIAL_CONFIRMAR_OPERACAO]
        self.clicar_btnConfirmar(driver)
        self.objTools.aguardar_carregamento(driver)
        sleep(self.delayDefault)

        # [CONTRIBUICAO_SOCIAL_VERIFICAR_STATUS_OPERACAO]
        self.verificar_statusOperacao(driver, 'Ocorrências > Regerar')
        sleep(self.delayDefault)

        retorno_escopo_calculo_por_ano = self.calcular_escopoDoCalculo_porAno(admissao, rescisao, inicio_calculo, termino_calculo)
        if retorno_escopo_calculo_por_ano <= 30:
            # [DIGITAR_VALORES_SALARIOS_PAGOS]
            status_operacao = self.get_value_salariosPagos_historico_e_digitar(driver)
            print(f"- [STATUS_OPERACAO_DIGITAR_SALARIOS_PAGOS]: {status_operacao}")
            if status_operacao[0] is False:

                self.clicar_btnSalvar(driver)
                self.objTools.aguardar_carregamento(driver)
                sleep(self.delayDefault)
                self.verificar_statusOperacao(driver, '')
                sleep(self.delayDefault)
                self.objTools.limparFilesTemp()
                return [True, '']

                # [REVER_ESSA_ROTINA]

                # print(f'- [REPETINDO_OPERACAO]')
                # driver.refresh()
                # sleep(6)
                # status_operacao = self.get_value_salariosPagos_historico_e_digitar(driver)
                # print(f"- [STATUS_OPERACAO_DIGITAR_SALARIOS_PAGOS]: {status_operacao}")
                # if status_operacao[0] is False:
                #     self.mensagem_alert_frontend(driver, 'Falha ao tentar digitar os valores dos SALÁRIOS PAGOS pela segunda tentativa.')
                #     self.registrar_msg_log('Contribuição Social', 'Ocorrências de Contribuição Social sobre Salários Devidos',
                #                            '---------- Erro! ---------- : (Falha ao tentar digitar os valores dos SALÁRIOS PAGOS pela segunda tentativa.)')
                #     self.clicar_btnCancelar(driver)
                #     return [False, '']
            else:

                self.clicar_btnSalvar(driver)
                self.objTools.aguardar_carregamento(driver)
                sleep(self.delayDefault)
                self.verificar_statusOperacao(driver, '')
                sleep(self.delayDefault)
                self.objTools.limparFilesTemp()
                return [True, '']
        else:
            self.mensagem_alert_frontend(driver, 'O escopo do cálculo é superior ao limite da planilha base (30 anos). Irei ignorar esta operação.')
            self.registrar_msg_log('Contribuição Social', 'Ocorrências de Contribuição Social sobre Salários Devidos', '---------- Erro! ---------- : (O escopo do cálculo é superior ao limite da planilha base (30 anos))')
            self.clicar_btnCancelar(driver)