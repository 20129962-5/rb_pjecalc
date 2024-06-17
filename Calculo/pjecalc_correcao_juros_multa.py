from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
import time
import xlrd
import os
import gc
from Calculo.pjecalc_dados_calculo import DadosCalculo
from Tools.pjecalc_control import Control


class Correcao(DadosCalculo):

    def __init__(self, source):
        super().__init__(source)
        self.objeto_calculo = DadosCalculo(source)
        self.planilha_base = self.objeto_calculo.planilha_base
        self.tamanho_plan = len(self.planilha_base)
        self.delay = 10
        self.delayG = 1.5
        self.objTools = Control()


    var_controle_float = 0.0
    var_controle_string = ""
    var_controle_int = 1

    indice_trabalhista = ""
    combinar_com_outro_indice = ""
    outro_indice_trabalhista = ""

    outro_indice_trabalhista_a_partir_de = ""

    ignorar_taxa_negativa = ""
    aplicar_juros_fase_pre_judicial = ""
    tabelas_de_juros = ""
    tabelas_de_juros2 = ""
    combinar_com_outra_tabela_juros = ""
    outra_tabela_a_partir_de_juros = ""
    verbas_base_de_juros = ""
    cs_sp_previdencia_correcao = ""
    aplicar_sumula_368_TST = ""
    cs_sd_lei_11941 = ""
    cs_sd_lei_11941_a_partir_de = ""
    cs_sd_limitar_multa = ""
    cs_sd_limitar_multa_a_partir_de = ""
    cs_sd_trabalhista_correcao = ""


    def verificacao_dados_gerais(self, driver):

        local = "Dados Gerais"

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            # file_txt_log.write('* ' + campo + ' | ' + status + '\n')
            file_txt_log.write(f'- {campo} : {local} | {status}\n')
            return file_txt_log.close()

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Correção, Juros e Multa', 'Ok')
            elif 'Existem erros no formulário.' in msg or 'erro' in msg or 'Erro' in msg:
                # print('* ERRO!', msg)
                gerar_relatorio('Correção, Juros e Multa', '---------- Erro! ----------')
            else:
                print('#- ', msg)
        except TimeoutException:
            print('- [Except][Correção/Juros/Multa][1] - Elemento não encontrado/A Página demorou para responder. Encerrando...')
            exit()

        # Tempo de controle
        time.sleep(2)

    def verificacao_dados_especificos(self, driver):

        local = "Dados Específicos"

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo} : {local} | {status}\n')
            return file_txt_log.close()

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Correção, Juros e Multa', 'Ok')
            elif 'Existem erros no formulário.' in msg or 'erro' in msg or 'Erro' in msg:
                print('* ERRO!', msg)
                gerar_relatorio('Correção, Juros e Multa', '---------- Erro! ----------')
            else:
                print('#- ', msg)
        except TimeoutException:
            print('- [Except][Correção/Juros/Multa][2] - Elemento não encontrado/A Página demorou para responder. Encerrando...')
            exit()

        # Tempo de controle
        time.sleep(2)

    def acessar_correcao(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageParamAtualizacao"))).click()

    def salvar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:salvar'))).click()

    def preencher_dados_gerais_correcao_monetaria(self, driver):

        indice = 1
        # Correção
        for i in range(self.tamanho_plan):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            # Índice Trabalhista
            elif coluna_identificador == 'indice_trabalhista':
                self.indice_trabalhista = coluna_informacao
                print('- Índice Trabalhista: ', self.indice_trabalhista)
                # -- Indice Trabalhista
                campo_outro_indice = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:indiceTrabalhista')))
                selecionar_outro_indice = Select(campo_outro_indice)
                selecionar_outro_indice.select_by_visible_text(self.indice_trabalhista)
            # Opção — Combinar com Outro Índice
            elif coluna_identificador == 'combinar_com_outro_indice':
                self.combinar_com_outro_indice = coluna_informacao
                print('- Combinar com Outro Índice: ', self.combinar_com_outro_indice)
            # Outros Índices Trabalhistas
            elif coluna_identificador == f'outro_indice_trabalhista{indice}':
                self.outro_indice_trabalhista = coluna_informacao
                print(f'- Outro Índice Trabalhista {indice}: ', self.outro_indice_trabalhista)
            # Campo — A partir de
            elif coluna_identificador == f'outro_indice_trabalhista{indice}_a_partir_de':
                self.outro_indice_trabalhista_a_partir_de = coluna_informacao
                if type(self.outro_indice_trabalhista_a_partir_de) == int:
                    self.outro_indice_trabalhista_a_partir_de = xlrd.xldate_as_datetime(self.outro_indice_trabalhista_a_partir_de, 0)
                    self.outro_indice_trabalhista_a_partir_de = self.outro_indice_trabalhista_a_partir_de.strftime('%d/%m/%Y')
                print(f'- A partir de {indice}: ', self.outro_indice_trabalhista_a_partir_de)
            # Opção — Ignorar Taxa Negativa para Índice(s) selecionado(s)
            elif coluna_identificador == 'ignorar_taxa_negativa':
                self.ignorar_taxa_negativa = coluna_informacao
                print('- Ignorar Taxa Negativa para Índice(s) selecionado(s): ', self.ignorar_taxa_negativa)

                # Preencher o PJeCalc
                # -- Combinar com Outro Índice
                if self.combinar_com_outro_indice == "False":
                    # — Verificar se o botão já está habilitado
                    checkbox_combinar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:combinarOutroIndice')))
                    status_checkbox_combinar = checkbox_combinar.is_selected()
                    print('-- Status do Checkbox: "Combinar com Outro Índice": ', status_checkbox_combinar)
                    if status_checkbox_combinar:
                        checkbox_combinar.click()
                        print('* Checkbox - "Combinar com Outro Índice" - Foi Desabilitado.')
                    else:
                        print('* Checkbox - "Combinar com Outro Índice" - Desabilitado.')
                elif self.combinar_com_outro_indice == "True":
                    # — Verificar se o botão já está habilitado
                    checkbox_combinar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:combinarOutroIndice')))
                    status_checkbox_combinar = checkbox_combinar.is_selected()
                    print('-- Status do Checkbox: "Combinar com Outro Índice": ', status_checkbox_combinar)
                    if status_checkbox_combinar:
                        print('* Checkbox - "Combinar com Outro Índice" - Já Habilitado.')
                    else:
                        checkbox_combinar.click()
                        print('* Checkbox - "Combinar com Outro Índice" - Foi Habilitado.')
                    # Aguardar campo condicional aparecer
                    WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.ID, 'formulario:outroIndiceTrabalhista')))
                    # Selecionar Outro Índice Trabalhista
                    if type(self.outro_indice_trabalhista) != float:
                        # Selecionar Outro Índice Trabalhista
                        campo_outro_indice = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:outroIndiceTrabalhista')))
                        selecionar_outro_indice = Select(campo_outro_indice)
                        selecionar_outro_indice.select_by_visible_text(self.outro_indice_trabalhista)
                        # Data — A partir de
                        campo_a_partir_de = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:apartirDeOutroIndiceInputDate')))
                        campo_a_partir_de.send_keys(self.outro_indice_trabalhista_a_partir_de)
                        # Botão Adicionar
                        btn_adicionar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:addOutroIndice')))
                        btn_adicionar.click()
                        # Aguardar
                        self.objTools.aguardar_carregamento(driver)
                indice += 1

                # -- Ignorar Taxa Negativa para Índice(s) selecionado(s)
                if self.ignorar_taxa_negativa == 'True':
                    # Verificar se a opção já se encontra habilitada
                    elemento_ignorar_taxa = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:ignorarTaxaNegativa')))
                    status_checkbox_ignorar_taxa = elemento_ignorar_taxa.is_selected()
                    print('-- Status do Checkbox: "Ignorar Taxa Negativa para Índice(s) selecionado(s)": ', status_checkbox_ignorar_taxa)
                    if status_checkbox_ignorar_taxa:
                        print('* Checkbox - "Combinar com Outro Índice" - Já Habilitada.')
                    else:
                        elemento_ignorar_taxa.click()
                        print('* Checkbox - "Combinar com Outro Índice" - Habilitei.')
                elif self.ignorar_taxa_negativa == 'False':
                    # Verificar se a opção já se encontra habilitada
                    elemento_ignorar_taxa = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:ignorarTaxaNegativa')))
                    status_checkbox_ignorar_taxa = elemento_ignorar_taxa.is_selected()
                    print('-- Status do Checkbox: "Ignorar Taxa Negativa para Índice(s) selecionado(s)": ', status_checkbox_ignorar_taxa)
                    if status_checkbox_ignorar_taxa:
                        elemento_ignorar_taxa.click()
                        print('* Checkbox - "Combinar com Outro Índice" - Desabilitei.')

    def preencher_dados_gerais_juros_de_mora(self, driver):

        indice = 2
        for i in range(self.tamanho_plan):
            # print("- Índice: ", indice)

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == float:
                continue
            # Aplicar Juros na Fase Pré-Judicial
            elif coluna_identificador == 'aplicar_juros_fase_pre_judicial':
                self.aplicar_juros_fase_pre_judicial = coluna_informacao
                print('- Aplicar Juros na Fase Pré-Judicial: ', self.aplicar_juros_fase_pre_judicial)
            # Tabela de Juros
            elif coluna_identificador == "tabelas_de_juros":
                self.tabelas_de_juros = coluna_informacao
                print('- Tabela de Juros: ', self.tabelas_de_juros)
            elif coluna_identificador == "combinar_com_outra_tabela_juros":
                self.combinar_com_outra_tabela_juros = coluna_informacao
                print("- Combinar com Outra Tabela de Juros: ", self.combinar_com_outra_tabela_juros)
            # Combinar com Outra Tabela de Juros
            elif coluna_identificador == f"tabelas_de_juros{indice}":
                self.tabelas_de_juros2 = coluna_informacao
                print(f'- Tabela de Juros ({indice}): ', self.tabelas_de_juros2)
            # Outra Tabela Juros
            elif coluna_identificador == f"outra_tabela_a_partir_de_juros_{indice}":
                self.outra_tabela_a_partir_de_juros = coluna_informacao
                if type(self.outra_tabela_a_partir_de_juros) == int:
                    self.outra_tabela_a_partir_de_juros = xlrd.xldate_as_datetime(self.outra_tabela_a_partir_de_juros, 0)
                    self.outra_tabela_a_partir_de_juros = self.outra_tabela_a_partir_de_juros.strftime("%d/%m/%Y")
                print(f'- Outra Tabela Juros ({indice}): ', self.outra_tabela_a_partir_de_juros)

                for k in range(3):
                    # Aplicar Juros na Fase Pré-Judicial
                    if self.aplicar_juros_fase_pre_judicial == "True":
                        # Verificar se a opção já se encontra habilitada
                        elemento_aplicar_juros = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:aplicarJurosFasePreJudicial')))
                        status_checkbox_aplicar_juros = elemento_aplicar_juros.is_selected()
                        print('- Status do Checkbox: "Aplicar Juros na Fase Pré-Judicial": ', status_checkbox_aplicar_juros)
                        if status_checkbox_aplicar_juros:
                            print('* Checkbox - "Aplicar Juros na Fase Pré-Judicial" - Já Habilitada.')
                        else:
                            elemento_aplicar_juros.click()
                            print('* Checkbox - "Aplicar Juros na Fase Pré-Judicial" - Habilitei.')
                    elif self.aplicar_juros_fase_pre_judicial == "False":
                        # Verificar se a opção já se encontra habilitada
                        elemento_aplicar_juros = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:aplicarJurosFasePreJudicial')))
                        status_checkbox_aplicar_juros = elemento_aplicar_juros.is_selected()
                        print('- Status do Checkbox: "Aplicar Juros na Fase Pré-Judicial": ', status_checkbox_aplicar_juros)
                        if status_checkbox_aplicar_juros:
                            elemento_aplicar_juros.click()
                            print('* Checkbox - "Aplicar Juros na Fase Pré-Judicial" - Desabilitei.')
                    # Tabela de Juros
                    campo_tabela_juros = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:juros')))
                    selecionar_tabela_juros = Select(campo_tabela_juros)
                    selecionar_tabela_juros.select_by_visible_text(self.tabelas_de_juros)
                    # Combinar com Outra Tabela de juros
                    if self.combinar_com_outra_tabela_juros == "False":
                        # Verificar se a opção já se encontra habilitada
                        elemento_combinar_tabela_juros = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:combinarOutroJuros')))
                        status_checkbox_combinar_tabela_juros = elemento_combinar_tabela_juros.is_selected()
                        print('-- Status do Checkbox: "Combinar com Outra Tabela de Juros": ', status_checkbox_combinar_tabela_juros)
                        if status_checkbox_combinar_tabela_juros:
                            elemento_combinar_tabela_juros.click()
                            print('* Checkbox - "Combinar com Outra Tabela de Juros" - Foi Desabilitado.')
                    elif self.combinar_com_outra_tabela_juros == "True":
                        # Verificar se a opção já se encontra habilitada
                        elemento_combinar_tabela_juros = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:combinarOutroJuros')))
                        status_checkbox_combinar_tabela_juros = elemento_combinar_tabela_juros.is_selected()
                        print('-- Status do Checkbox: "Combinar com Outra Tabela de Juros": ', status_checkbox_combinar_tabela_juros)
                        if status_checkbox_combinar_tabela_juros:

                            # — Verificar se já há um índice adicionado — #
                            indice = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "formulario:j_id113:tb")))
                            if indice:
                                # Desmarcar Checkbox - Combinar com Outra Tabela de Juros
                                WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:combinarOutroJuros'))).click()
                                status_checkbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:combinarOutroJuros'))).is_selected()
                                if not status_checkbox:
                                    # Habilitar - Causando a limpeza do índice
                                    WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:combinarOutroJuros'))).click()
                        else:
                            elemento_combinar_tabela_juros.click()
                            print('* Checkbox - "Combinar com Outra Tabela de Juros" - Habilitei.')
                        # Aguardar campo condicional
                        WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.ID, 'formulario:outroJuros')))
                        # Condição para preencher no PJeCalc se o conteúdo for diferente de vazio
                        if type(self.combinar_com_outra_tabela_juros) != float:
                            # Preencher Tabela de Juros
                            campo_combinar_tabela_juros = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:outroJuros')))
                            selecionar_combinar_tabela_juros = Select(campo_combinar_tabela_juros)
                            selecionar_combinar_tabela_juros.select_by_visible_text(self.tabelas_de_juros2)
                            # Preencher Data - A partir de
                            campo_data_a_partir_de = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:apartirDeOutroJurosInputDate')))
                            campo_data_a_partir_de.send_keys(self.outra_tabela_a_partir_de_juros)
                            # Adicionar
                            btn_adicionar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:addOutroJuros')))
                            btn_adicionar.click()
                            # Aguardar
                            self.objTools.aguardar_carregamento(driver)
                            time.sleep(1)

                    # Sair
                    break
                # indice = indice + 1
                # continue

    def acessar_aba_dados_especificos(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tabDadosEspecificos_lbl'))).click()

    def aplicar_sumula_368_tst_v3_34(self, driver):

        for i in range(self.tamanho_plan):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                continue
            # Contribuição Social
            elif coluna_identificador == "aplicar_sumula_368_TST":
                self.aplicar_sumula_368_TST = coluna_informacao
                print("- Aplicar Súmula 368 TST: ", self.aplicar_sumula_368_TST)
            #
            elif coluna_identificador == "cs_sd_lei_11941":
                self.cs_sd_lei_11941 = coluna_informacao
                print("- Lei: Lei nº 11.941/2009", self.cs_sd_lei_11941)
            elif coluna_identificador == "cs_sd_lei_11941_a_partir_de":
                self.cs_sd_lei_11941_a_partir_de = coluna_informacao
                if type(self.cs_sd_lei_11941_a_partir_de) == int:
                    self.cs_sd_lei_11941_a_partir_de = xlrd.xldate_as_datetime(self.cs_sd_lei_11941_a_partir_de, 0)
                    self.cs_sd_lei_11941_a_partir_de = self.cs_sd_lei_11941_a_partir_de.strftime("%d/%m/%Y")
                print("- A partir de: ", self.cs_sd_lei_11941_a_partir_de)
            elif coluna_identificador == "cs_sd_limitar_multa":
                self.cs_sd_limitar_multa = coluna_informacao
                print("- Limitar multa: ", self.cs_sd_limitar_multa)
            elif coluna_identificador == "cs_sd_limitar_multa_a_partir_de":
                self.cs_sd_limitar_multa_a_partir_de = coluna_informacao
                print("- A partir de: ", self.cs_sd_limitar_multa_a_partir_de)
            elif coluna_identificador == "cs_sd_trabalhista_correcao":
                self.cs_sd_trabalhista_correcao = coluna_informacao
                print("- Correção: ", self.cs_sd_trabalhista_correcao)

                # PJeCalc
                if self.cs_sd_lei_11941 == "True":
                    # Verificar status dos checkboxs (Lei nº 11.941/2009, Limitar multa e Correção)
                    # Lei nº 11.941/2009
                    campo_lei_11941_2009 = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:correcaoLei11941")))
                    status_checkbox = campo_lei_11941_2009.is_selected()
                    if status_checkbox:
                        print("- Checkbox - 'Lei nº 11.941/2009' - Já Habilitado.")
                    else:
                        campo_lei_11941_2009.click()
                        print("- Checkbox - 'Lei nº 11.941/2009' - Foi Habilitado.")

                    # Campo - 'A partir de'
                    # Aguardar o campo da data
                    if self.cs_sd_lei_11941_a_partir_de == "05/03/2009" or self.cs_sd_lei_11941_a_partir_de == "<oculto>":
                        pass
                        # print("!! Data da planilha é idêntico ao do PJeCalc !!")
                    else:
                        try:
                            WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID, "formulario:aplicarAteLei11941InputDate")))
                        except TimeoutException:
                            time.sleep(2)

                        # Preencher data
                        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "formulario:aplicarAteLei11941InputDate"))).send_keys(self.cs_sd_lei_11941_a_partir_de)
                        time.sleep(0.5)

                    # Checkbox - 'Limitar Multa'
                    if self.cs_sd_limitar_multa == "True":
                        # Limitar Multa
                        limitar_multa = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:correcaoLei11941Multa")))
                        status_checkbox = limitar_multa.is_selected()
                        if status_checkbox:
                            print("- Checkbox - 'Limitar Multa' - Já Habilitado.")
                        else:
                            limitar_multa.click()
                            print("- Checkbox - 'Limitar Multa' - Foi Habilitado.")

                        # Limitir multar - A partir de
                        # Aguardar o campo da data
                        if self.cs_sd_limitar_multa_a_partir_de == "<em branco>" or self.cs_sd_limitar_multa_a_partir_de == "<oculto>":
                            pass
                        else:
                            try:
                                WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID, "formulario:aplicarAteLei11941MultaInputDate")))
                            except TimeoutException:
                                time.sleep(2)
                            # Preencher data
                            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "formulario:aplicarAteLei11941MultaInputDate"))).send_keys(self.cs_sd_limitar_multa_a_partir_de)
                            time.sleep(0.5)
                    elif self.cs_sd_limitar_multa == "False":
                        # Limitar Multa
                        limitar_multa = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:correcaoLei11941Multa")))
                        status_checkbox = limitar_multa.is_selected()
                        if status_checkbox:
                            limitar_multa.click()
                            print("- Checkbox - 'Limitar Multa' - Foi Habilitado.")

                else:
                    campo_lei_11941_2009 = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:correcaoLei11941")))
                    status_checkbox = campo_lei_11941_2009.is_selected()
                    if status_checkbox:
                        campo_lei_11941_2009.click()
                        print("- Checkbox - 'Lei nº 11.941/2009' - Foi Desabilitado.")

                # Checkbox — Trabalhista — Correção
                if self.cs_sd_trabalhista_correcao == "True":
                    correcao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:correcaoTrabalhistaDosSalariosDevidosDoINSS")))
                    status_checkbox = correcao.is_selected()
                    if status_checkbox:
                        print("- Checkbox - 'Correção' - Já Habilitado.")
                    else:
                        correcao.click()
                        print("- Checkbox - 'Correção' - Foi Habilitado.")
                else:
                    correcao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:correcaoTrabalhistaDosSalariosDevidosDoINSS")))
                    status_checkbox = correcao.is_selected()
                    if status_checkbox:
                        correcao.click()
                        print("- Checkbox - 'Correção' - Foi Desabilitado.")

    def preencher_dados_especificos(self, driver):

        for i in range(self.tamanho_plan):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                continue
            # Índice Trabalhista
            elif coluna_identificador == 'verbas_base_de_juros':
                self.verbas_base_de_juros = coluna_informacao

                # Preencher PJeCalc
                campo_bases_juros = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:baseDeJurosDasVerbas')))
                selecionar_base_juros = Select(campo_bases_juros)
                selecionar_base_juros.select_by_visible_text(self.verbas_base_de_juros)

            elif coluna_identificador == 'cs_sp_previdencia_correcao':
                self.cs_sp_previdencia_correcao = coluna_informacao
                print('- Previdenciária - Correção: ', self.cs_sp_previdencia_correcao)
                # Sair
                break

        # Tempo de controle
        time.sleep(1)

        # Checkbox — Previdenciária — Correção
        if self.cs_sp_previdencia_correcao == 'True':
            elemento_correcao = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:correcaoPrevidenciariaDosSalariosPagosDoINSS')))
            status_checkbox = elemento_correcao.is_selected()
            if status_checkbox:
                print('- Checkbox - "Previdenciário - Correção" - Já Habilitado.')
            else:
                print('- Checkbox - "Previdenciário - Correção" - Foi Habilitado.')
                elemento_correcao.click()
        else:
            elemento_correcao = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:correcaoPrevidenciariaDosSalariosPagosDoINSS')))
            status_checkbox = elemento_correcao.is_selected()
            if status_checkbox:
                print('- Checkbox - "Previdenciário - Correção" - Foi Desabilitado.')
                elemento_correcao.click()

        # Salvar Operação
        self.salvar(driver)

        # Aguardar
        self.objTools.aguardar_carregamento(driver)

        # Tempo de controle
        time.sleep(1)

        # Verificar
        self.verificacao_dados_especificos(driver)

    def main_correcao(self, driver):

        self.acessar_correcao(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.preencher_dados_gerais_correcao_monetaria(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.preencher_dados_gerais_juros_de_mora(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        # Salvar
        self.salvar(driver)
        # Aguardar
        self.objTools.aguardar_carregamento(driver)
        # Verificar
        self.verificacao_dados_gerais(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.acessar_aba_dados_especificos(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        # Apenas para versão da planilha 3.34 - Início
        self.aplicar_sumula_368_tst_v3_34(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.preencher_dados_especificos(driver)
        # Tempo de controle
        time.sleep(self.delayG)

        # - Limpar Temp
        self.objTools.limparFilesTemp()
        print('-- Fim - (Correção) --')