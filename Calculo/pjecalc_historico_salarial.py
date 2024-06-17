from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
import pandas as pd
import time
import xlrd
import os
import gc
#
from Tools.pjecalc_control import Control


class HistoricoSalarial:

    def __init__(self, source):
        self.source = source
        self.planilha_base = pd.read_excel(self.source, sheet_name='PJE HIST-VAL', header=3)
        self.delay = 10
        self.delayG = 1.5
        self.objTools = Control()

    competencia_datetime = ""
    var_controle_float = 0.0
    var_controle_str = ""

    time_controle = 5
    nome_coluna = ""
    titulo_coluna = ""
    titulo_coluna_fgts = ""
    contador = 0
    data_demissao = ""
    data_final_calc = ""

    def gerar_relatorio(self, campo, subatividade, status):
        file_txt_log = open(os.getcwd() + '\log.txt', "a")
        file_txt_log.write(f'- {campo} : {subatividade} | {status}\n')
        return file_txt_log.close()

    def verificacao(self, driver):

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            # file_txt_log.write('- ' + campo + ' | ' + status + '\n')
            file_txt_log.write(f'- {campo} : {self.nome_coluna.title()} | {status}\n')
            return file_txt_log.close()

        def cancelar_operacao():
            btn_cancelar = WebDriverWait(driver, self.delay).until(
                EC.element_to_be_clickable((By.ID, 'formulario:cancelar')))
            btn_cancelar.click()
            self.objTools.aguardar_carregamento(driver)

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Histórico Salarial', 'Ok')
            else:
                # print('* ERRO!', msg)
                gerar_relatorio('Histórico Salarial', '---------- Erro! ----------')
                cancelar_operacao()
        except TimeoutException:
            print('- [Except][HS] - Elemento não encontrado/A Página demorou para responder. Encerrando...')

        # Tempo de controle
        time.sleep(2)

    def entrar_historico_salarial(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageBase"))).click()

    def click_grade_ocorrencias(self, driver):
        WebDriverWait(driver, self.delay).until(
            EC.element_to_be_clickable((By.ID, 'formulario:visualizarOcorrencias'))).click()

    def preencher_gratificacao_semestral(self, driver):

        inicio = 10
        campo_id = 0
        for i in range(49):
            valores = self.planilha_base.loc[inicio, 'col_2']
            val_format = '{:.2f}'.format(valores)
            # print(str(i + 1), '-', 'R$ ', val_format)
            salario_base = WebDriverWait(driver, self.delay).until(
                EC.presence_of_element_located((By.NAME, 'formulario:tabOcorrencias:' + str(campo_id) + ':linha1')))
            salario_base.send_keys(val_format)
            # Próximos valores
            campo_id += 1
            inicio += 1
        # time.sleep(3)

    def salvar(self, driver):
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'formulario:salvar'))).click()

    def criar_novo(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:incluir'))).click()

    def preencher_dados_historico_salarial(self, driver, admissao, rescisao, inicio_calculo, termino_calculo):

        print(
            f"- Data de Admissão: {admissao}\n- Data de Demissão: {rescisao}\n- Data Inicial do Cálculo: {inicio_calculo}\n- Data Final do Cálculo: {termino_calculo}\n")
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

        print("\n- Valor Final do Escopo do Cálculo: ", resultado)

        if resultado > 42:
            driver.execute_script(script)
            WebDriverWait(driver, self.delay).until(EC.alert_is_present())
            alerta = Alert(driver)
            self.gerar_relatorio("Histórico Salarial",
                                 "O intervalo do Cálculo é superior ao limite da planilha (30 anos)",
                                 "---------- Erro! ----------")
            time.sleep(5)
            try:
                alerta.accept()
            except NoAlertPresentException:
                pass
        else:

            ct_cp = []

            for m in range(1, len(self.planilha_base)):
                competencia = self.planilha_base.loc[m, 'MÊS/ANO']
                if type(competencia) != type(self.var_controle_float):
                    try:
                        competencia = int(competencia)
                        competencia_int = xlrd.xldate_as_datetime(competencia, 0)
                        competencia_int = competencia_int.strftime('%m/%Y')
                        ct_cp.append(competencia_int)
                    except:
                        pass
            print('- Competencia Qtd: ', len(ct_cp))
            competencia_inicial = ct_cp[0]
            competencia_final = ct_cp[-1]
            print('- Mês/Ano - Inicial: ', competencia_inicial)
            print('- Mês/Ano - Final: ', competencia_final)

            # Colunas
            for c in range(1, 20):
                try:
                    try:
                        # Pegar o nome da coluna
                        for l in range(len(ct_cp)):
                            coluna = self.planilha_base.loc[l, c]
                            # Pular valores 'nan' tipo float
                            if type(coluna) == type(self.var_controle_str):
                                self.nome_coluna = coluna
                                # print(self.nome_coluna)
                                print('- Contagem: ', self.contador)
                                self.contador = 1
                                break
                            else:
                                self.contador = 0
                    except:
                        print('- Exceção - Pegar nome da coluna.')
                        break

                    if self.contador == 0:
                        continue
                        # break
                    else:
                        # Definir um limite
                        for s in range(self.contador):
                            self.criar_novo(driver)
                            self.objTools.aguardar_carregamento(driver)
                            # Tempo de controle
                            time.sleep(1)
                            try:
                                # Colocar nome da coluna
                                campo_nome_coluna = WebDriverWait(driver, self.delay).until(
                                    EC.presence_of_element_located((By.NAME, 'formulario:nome')))
                                campo_nome_coluna.send_keys(self.nome_coluna)
                            except:
                                print('- Exceção - Nome.')
                                pass
                            # Tempo de controle
                            time.sleep(1)
                            # Competência Inicial
                            campo_competencia_inicial = WebDriverWait(driver, self.delay).until(
                                EC.presence_of_element_located((By.NAME, 'formulario:competenciaInicialInputDate')))
                            campo_competencia_inicial.send_keys(competencia_inicial)
                            # Tempo de controle
                            time.sleep(1)
                            # Competência Final
                            campo_competencia_final = WebDriverWait(driver, self.delay).until(
                                EC.presence_of_element_located((By.NAME, 'formulario:competenciaFinalInputDate')))
                            campo_competencia_final.send_keys(competencia_final)
                            # Tempo de controle
                            time.sleep(1)
                            try:
                                # Valor 0
                                campo_valor = WebDriverWait(driver, self.delay).until(
                                    EC.presence_of_element_located((By.NAME, 'formulario:valorParaBaseDeCalculo')))
                                campo_valor.send_keys('0')
                            except:
                                print('- Exceção - Preencher Valor.')
                                pass
                            # Tempo de controle
                            time.sleep(1)
                            try:
                                # Adicionar
                                btn_gerar_ocorrencia = WebDriverWait(driver, self.delay).until(
                                    EC.element_to_be_clickable((By.ID, 'formulario:cmdGerarOcorrencias')))
                                btn_gerar_ocorrencia.click()
                            except:
                                print('- Exceção - Adicionar.')
                                pass
                            # Aguardar
                            self.objTools.aguardar_carregamento(driver)
                            # Tempo de controle
                            time.sleep(1)
                            indice = 0
                            for k in range(1, len(ct_cp) + 1):
                                dados = self.planilha_base.loc[k, c]
                                # print("***** Dados da Planilha: ", dados, type(dados))
                                # if type(dados) != type(self.var_controle_float):
                                # Formatação
                                dados = '{:.2f}'.format(dados)
                                # print(dados)
                                campo_valor = WebDriverWait(driver, self.delay).until(
                                    EC.presence_of_element_located((By.NAME, f'formulario:listagemMC:{indice}:valor')))
                                campo_valor.send_keys(dados)
                                indice += 1

                            # Tempo de controle
                            time.sleep(1)
                            # Novo Trecho
                            self.salvar(driver)
                            self.objTools.aguardar_carregamento(driver)
                            # Tempo de controle
                            time.sleep(1)
                            self.verificacao(driver)
                            # Tempo de controle
                            time.sleep(1)
                            break
                except:
                    print('- Exceção - Geral.')
                    break

    def preencher_dados_hist_salarial_and_fgts_v3_34(self, driver, admissao, rescisao, inicio_calculo, termino_calculo):

        print(
            f"- Data de Admissão: {admissao}\n- Data de Demissão: {rescisao}\n- Data Inicial do Cálculo: {inicio_calculo}\n- Data Final do Cálculo: {termino_calculo}\n")
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

        print("\n- Valor Final do Escopo do Cálculo: ", resultado)

        if resultado > 42:
            driver.execute_script(script)
            WebDriverWait(driver, self.delay).until(EC.alert_is_present())
            alerta = Alert(driver)
            self.gerar_relatorio("Histórico Salarial",
                                 "O intervalo do Cálculo é superior ao limite da planilha (30 anos)",
                                 "---------- Erro! ----------")
            time.sleep(5)
            try:
                alerta.accept()
            except NoAlertPresentException:
                pass
        else:
            qtd_competencias = self.planilha_base.loc[0, "MÊS/ANO"]
            competencia_inicial_f1 = self.planilha_base.loc[2, "MÊS/ANO"]
            competencia_final_f2 = self.planilha_base.loc[qtd_competencias + 1, "MÊS/ANO"]
            # Formatação Competência Inicial
            competencia_inicial = xlrd.xldate_as_datetime(competencia_inicial_f1, 0)
            competencia_inicial = competencia_inicial.strftime("%m/%Y")
            # Formatação Competência Final

            competencia_final = xlrd.xldate_as_datetime(competencia_final_f2, 0)
            competencia_final = competencia_final.strftime("%m/%Y")
            # Output
            print(f"- [COMPETENCIA_INICIAL]: {competencia_inicial}")
            print(f"- [COMPETENCIA_FINAL]: {competencia_final}\n")
            coletar = True
            coletar_fgts = True
            indice = 0
            for c in range(1, 21):  # Colunas
                for l in range(2, qtd_competencias + 2):  # Quantidade de linhas iniciando com índice 2 da coluna 'MÊS/ANO'
                    # Condição para verificar se a coluna tem um título e a próxima linha for igual a 0
                    if type(self.planilha_base.loc[0, c]) == type(self.var_controle_str) and self.planilha_base.loc[1, c] == 0:
                        # PLANILHA — Coletar apenas o título da coluna
                        while coletar:
                            coletar = False
                            self.titulo_coluna = self.planilha_base.loc[0, c]
                            print("- Título da Coluna: ", self.titulo_coluna)

                            # PJECALC
                            # Novo
                            self.criar_novo(driver)
                            self.objTools.aguardar_carregamento(driver)
                            time.sleep(2)
                            # Digitar Nome da coluna
                            campoNome = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:nome")))
                            campoNome.send_keys(self.titulo_coluna)
                            time.sleep(0.5)
                            # webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                            # Digitar a competência inicial
                            field_competencia_inicial = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:competenciaInicialInputDate")))
                            field_competencia_inicial.clear()
                            field_competencia_inicial.send_keys(competencia_inicial)
                            time.sleep(0.5)
                            # Digitar a competência Final
                            field_competencia_final = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:competenciaFinalInputDate")))
                            field_competencia_final.clear()
                            field_competencia_final.send_keys(competencia_final)
                            time.sleep(0.5)
                            # Digitar o valor 0 inicialmente
                            WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:valorParaBaseDeCalculo"))).send_keys("0")
                            # Clique no botão add
                            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:cmdGerarOcorrencias"))).click()
                            # Aguardar carregamento
                            self.objTools.aguardar_carregamento(driver)
                            time.sleep(2)

                        # PLANILHA — Percorrer todos os valores da coluna que atende a condição
                        conteudo = self.planilha_base.loc[l, c]
                        conteudo_formato = f"{conteudo:.2f}"
                        # print(conteudo_formato)

                        # PJECALC
                        # Digitar todos os valores da coluna no PJeCalc
                        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located(
                            (By.NAME, f"formulario:listagemMC:{indice}:valor"))).send_keys(conteudo_formato)
                        indice += 1

                    # Condição para verificar se a coluna tem um título e a próxima linha for igual a SIM
                    elif type(self.planilha_base.loc[0, c]) == type(self.var_controle_str) and self.planilha_base.loc[1, c] == "SIM":
                        # Coletar apenas o título da coluna
                        while coletar_fgts:
                            coletar_fgts = False
                            self.titulo_coluna_fgts = self.planilha_base.loc[0, c]
                            print("- Título da Coluna FGTS: ", self.titulo_coluna_fgts)

                            # PJECALC
                            # Novo
                            self.criar_novo(driver)
                            self.objTools.aguardar_carregamento(driver)
                            time.sleep(2)
                            # Digitar Nome da coluna
                            WebDriverWait(driver, self.delay).until(
                                EC.presence_of_element_located((By.NAME, "formulario:nome"))).send_keys(
                                self.titulo_coluna_fgts)
                            # Checkbox — Incidência — Marcar FGTS
                            WebDriverWait(driver, self.delay).until(
                                EC.element_to_be_clickable((By.ID, "formulario:fgts"))).click()
                            # Checkbox — Incidência — Desmarcar FGTS
                            WebDriverWait(driver, self.delay).until(
                                EC.visibility_of_element_located((By.ID, "formulario:proporcionalizarFGTS"))).click()
                            # Digitar a competência inicial
                            WebDriverWait(driver, self.delay).until(EC.presence_of_element_located(
                                (By.NAME, "formulario:competenciaInicialInputDate"))).send_keys(competencia_inicial)
                            # Digitar a competência Final
                            WebDriverWait(driver, self.delay).until(EC.presence_of_element_located(
                                (By.NAME, "formulario:competenciaFinalInputDate"))).send_keys(competencia_final)
                            # Digitar o valor 0 inicialmente
                            WebDriverWait(driver, self.delay).until(EC.presence_of_element_located(
                                (By.NAME, "formulario:valorParaBaseDeCalculo"))).send_keys("0")
                            # Clique no botão add
                            WebDriverWait(driver, self.delay).until(
                                EC.element_to_be_clickable((By.ID, "formulario:cmdGerarOcorrencias"))).click()
                            # Aguardar carregamento
                            self.objTools.aguardar_carregamento(driver)
                            time.sleep(2)

                        # Percorrer todos os valores da coluna que atende a condição
                        conteudo_fgts = self.planilha_base.loc[l, c]
                        conteudo_fgts_formatado = f"{conteudo_fgts:.2f}"
                        # print(conteudo_fgts_formatado)

                        # PJECALC
                        # Digitar todos os valores da coluna no PJeCalc
                        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located(
                            (By.NAME, f"formulario:listagemMC:{indice}:valor"))).send_keys(conteudo_fgts_formatado)
                        indice += 1
                    else:
                        continue

                # Salvar após digitar todos os valores
                try:
                    self.salvar(driver)
                    self.objTools.aguardar_carregamento(driver)
                    time.sleep(2)
                    # Verificação
                    self.verificacao(driver)
                except TimeoutException:
                    break

                coletar = True
                coletar_fgts = True
                indice = 0
                print("\n")


    def preencher_hist_fgts(self, driver):

        planilha = pd.read_excel(self.source, sheet_name="PJE HIST-VAL", header=2)

        def montar_estrutura_dados():

            # PJeCalc
            # Novo
            self.criar_novo(driver)
            # Aguardar
            self.objTools.aguardar_carregamento(driver)
            time.sleep(1)
            # Digitar Nome
            WebDriverWait(driver, self.delay).until(
                EC.presence_of_element_located((By.NAME, "formulario:nome"))).send_keys(nome)
            time.sleep(0.5)
            # Habilitar Checkbox do FGTS
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:fgts"))).click()
            time.sleep(0.5)
            # Aguardar checkbox condicional 'Proporcionalizar FGTS' e desmarcar
            WebDriverWait(driver, self.delay).until(
                EC.visibility_of_element_located((By.ID, "formulario:proporcionalizarFGTS"))).click()
            time.sleep(0.5)
            # Colocar a Competência Inicial
            campo_competencia_inicial = WebDriverWait(driver, self.delay).until(
                EC.presence_of_element_located((By.NAME, "formulario:competenciaInicialInputDate")))
            campo_competencia_inicial.send_keys(Keys.BACKSPACE)
            time.sleep(0.5)
            campo_competencia_inicial.send_keys(data_inicial)
            time.sleep(0.5)
            # Colocar a Competência Final
            campo_competencia_final = WebDriverWait(driver, self.delay).until(
                EC.presence_of_element_located((By.NAME, "formulario:competenciaFinalInputDate")))
            campo_competencia_final.send_keys(Keys.BACKSPACE)
            time.sleep(0.5)
            campo_competencia_final.send_keys(data_final)
            time.sleep(0.5)
            # Colocar inicialmente o valor 0
            WebDriverWait(driver, self.delay).until(
                EC.presence_of_element_located((By.NAME, "formulario:valorParaBaseDeCalculo"))).send_keys(0)
            time.sleep(0.5)
            # Clique em adicionar
            WebDriverWait(driver, self.delay).until(
                EC.element_to_be_clickable((By.ID, "formulario:cmdGerarOcorrencias"))).click()
            # Aguardar
            self.objTools.aguardar_carregamento(driver)
            time.sleep(1)

        # Datas coletadas da planilha
        data_inicial = planilha.iloc[1, 31]
        data_final = planilha.iloc[2, 31]
        qtd_competencias = planilha.iloc[2, 32]

        # Coluna 'AH' da Planilha Base
        coluna_1 = planilha.iloc[0, 33]
        # Coluna 'AK' da Planilha Base
        coluna_2 = planilha.iloc[0, 36]
        # Coluna 'AN' da Planilha Base
        coluna_3 = planilha.iloc[0, 39]

        # Tratamento de exceção — Se o valor da competência estiver em branco, ignorar
        try:
            # Conversão - 'Competência Inicial' - float -> datetime
            data_inicial = xlrd.xldate_as_datetime(data_inicial, 0)
            data_inicial = data_inicial.strftime("%m/%Y")
        except ValueError:
            pass

        # Tratamento de exceção — Se o valor da competência estiver em branco, ignorar
        try:
            # Conversão - 'Competência Final' - float -> datetime
            data_final = xlrd.xldate_as_datetime(data_final, 0)
            data_final = data_final.strftime("%m/%Y")
        except ValueError:
            pass

        print("- Competência Inicial: ", data_inicial)
        print("- Competência Final: ", data_final)
        print("- Qtd de Competências: ", qtd_competencias)
        print("- Indicador da coluna FGTS: ", coluna_1, ' - ', coluna_2, ' - ', coluna_3)
        print()

        if coluna_1 == 1:
            nome = planilha.iloc[1, 33]
            # print("- ", nome)
            # PJeCalc
            montar_estrutura_dados()
            # Iterar e preecher os valores
            contador = 1
            for i in range(3, qtd_competencias + 3):
                base = planilha.iloc[i, 33]
                # Conversão
                base = f"{base:.2f}"
                # PJeCalc - Preencher os valores
                try:
                    WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.NAME, f"formulario:listagemMC:{i - 3}:valor"))).send_keys(base)
                except TimeoutException:
                    time.sleep(1.5)
                    break
                # Próxima iteração
                contador += 1

            # Tempo de controle
            time.sleep(1.5)

            # PJeCalc - Desmarcar o Checkbox Principal do Incide FGTS
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3"))).click()
            time.sleep(0.5)

            # print("Incide  Recolhido")
            # Checkbox - 'Incide FGTS' - Marcar ou desmarcar
            for i in range(3, qtd_competencias + 3):
                # Coluna 'AI' da Planilha Base
                incide = planilha.iloc[i, 34]
                # Coluna 'AJ' da Planilha Base
                recolhido = planilha.iloc[i, 35]
                # Output
                # print(incide, " ", recolhido)
                if incide == "Sim":
                    # PJeCalc
                    try:
                        checkbox = WebDriverWait(driver, 4).until(
                            EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:incideFGTS")))
                        if checkbox.is_selected():
                            pass
                            # print("- Checkbox - 'Incide FGTS' - Já selecionado.")
                        else:
                            # print("- Checkbox - 'Incide FGTS' - Selecionado.")
                            checkbox.click()
                    except TimeoutException:
                        break
                if recolhido == "Sim":
                    # PJeCalc
                    try:
                        checkbox = WebDriverWait(driver, 4).until(
                            EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:recolhidoFGTS")))
                        if checkbox.is_selected():
                            pass
                            # print("- Checkbox - 'FGTS Recolhido' - Já selecionado.")
                        else:
                            # print("- Checkbox - 'FGTS Recolhido - Selecionado.")
                            checkbox.click()
                    except TimeoutException:
                        break
                else:
                    continue

            # Tempo de controle
            time.sleep(1)
            # Salvar Operação
            self.salvar(driver)
            self.objTools.aguardar_carregamento(driver)
            time.sleep(1)
            self.verificacao(driver)
            time.sleep(2)

        if coluna_2 == 1:
            nome = planilha.iloc[1, 36]
            # print("- ", nome)

            # PJeCalc
            montar_estrutura_dados()

            # Iterar e preecher os valores
            contador = 1
            for i in range(3, qtd_competencias + 3):
                base = planilha.iloc[i, 36]
                # Conversão
                # contagem = f"{contador:0>3}"
                # Conversão para duas casas decimais
                base = f"{base:.2f}"
                # Output
                # print(contagem, " - ", base)
                # PJeCalc - Preencher os valores
                try:
                    WebDriverWait(driver, 4).until(
                        EC.presence_of_element_located((By.NAME, f"formulario:listagemMC:{i - 3}:valor"))).send_keys(base)
                except TimeoutException:
                    break
                # Próxima iteração
                contador += 1
            # Tempo de controle
            time.sleep(1)
            # PJeCalc - Desmarcar o Checkbox Principal do Incide FGTS
            WebDriverWait(driver, self.delay).until(
                EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3"))).click()
            time.sleep(0.5)

            # print("Incide  Recolhido")
            # Checkbox - 'Incide FGTS' - Marcar ou desmarcar
            for i in range(3, qtd_competencias + 3):
                # Coluna 'AL' da Planilha Base
                incide = planilha.iloc[i, 37]
                # Coluna 'AM' da Planilha Base
                recolhido = planilha.iloc[i, 38]
                # Output
                # print(incide, " ", recolhido)

                if incide == "Sim":
                    # PJeCalc
                    try:
                        checkbox = WebDriverWait(driver, 4).until(
                            EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:incideFGTS")))
                        if checkbox.is_selected():
                            pass
                            # print("- Checkbox - 'Incide FGTS' - Já selecionado.")
                        else:
                            # print("- Checkbox - 'Incide FGTS' - Selecionado.")
                            checkbox.click()
                    except TimeoutException:
                        break
                if recolhido == "Sim":
                    # PJeCalc
                    try:
                        checkbox = WebDriverWait(driver, 4).until(
                            EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:recolhidoFGTS")))
                        if checkbox.is_selected():
                            pass
                            # print("- Checkbox - 'FGTS Recolhido' - Já selecionado.")
                        else:
                            # print("- Checkbox - 'FGTS Recolhido - Selecionado.")
                            checkbox.click()
                    except TimeoutException:
                        break
                else:
                    continue

            # Tempo de controle
            time.sleep(1)
            # Salvar Operação
            self.salvar(driver)
            self.objTools.aguardar_carregamento(driver)
            time.sleep(1)
            self.verificacao(driver)
            time.sleep(2)

        if coluna_3 == 1:
            nome = planilha.iloc[1, 39]
            print("- ", nome)

            # PJeCalc
            montar_estrutura_dados()

            # Iterar e preecher os valores
            contador = 1
            for i in range(3, qtd_competencias + 3):
                base = planilha.iloc[i, 39]
                # Conversão
                # contagem = f"{contador:0>3}"
                # Conversão para duas casas decimais
                base = f"{base:.2f}"
                # Output
                # print(contagem, " - ", base)
                # PJeCalc - Preencher os valores
                try:
                    WebDriverWait(driver, 4).until(
                        EC.presence_of_element_located((By.NAME, f"formulario:listagemMC:{i - 3}:valor"))).send_keys(base)
                except TimeoutException:
                    break
                # Próxima iteração
                contador += 1
            # Tempo de controle
            time.sleep(1)
            # PJeCalc - Desmarcar o Checkbox Principal do Incide FGTS
            WebDriverWait(driver, self.delay).until(
                EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3"))).click()
            time.sleep(0.5)

            print("Incide  Recolhido")
            # Checkbox - 'Incide FGTS' - Marcar ou desmarcar
            for i in range(3, qtd_competencias + 3):
                # Coluna 'AO' da Planilha Base
                incide = planilha.iloc[i, 40]
                # Coluna 'AP' da Planilha Base
                recolhido = planilha.iloc[i, 41]
                # Output
                print(incide, " ", recolhido)

                if incide == "Sim":
                    # PJeCalc
                    try:
                        checkbox = WebDriverWait(driver, 4).until(
                            EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:incideFGTS")))
                        if checkbox.is_selected():
                            pass
                            # print("- Checkbox - 'Incide FGTS' - Já selecionado.")
                        else:
                            # print("- Checkbox - 'Incide FGTS' - Selecionado.")
                            checkbox.click()
                    except TimeoutException:
                        break
                if recolhido == "Sim":
                    # PJeCalc
                    try:
                        checkbox = WebDriverWait(driver, 4).until(
                            EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:recolhidoFGTS")))
                        if checkbox.is_selected():
                            pass
                            # print("- Checkbox - 'FGTS Recolhido' - Já selecionado.")
                        else:
                            # print("- Checkbox - 'FGTS Recolhido - Selecionado.")
                            checkbox.click()
                    except TimeoutException:
                        break
                else:
                    continue
            # Tempo de controle
            time.sleep(1)
            # Salvar Operação
            self.salvar(driver)
            self.objTools.aguardar_carregamento(driver)
            time.sleep(1)
            self.verificacao(driver)
            time.sleep(2)

    def preencher_hist_fgts_novo(self, driver):

        planilha = pd.read_excel(self.source, sheet_name="PJE HIST-VAL", header=2)

        def montar_estrutura_dados():

            # PJeCalc
            # Novo
            self.criar_novo(driver)
            # Aguardar
            self.objTools.aguardar_carregamento(driver)
            time.sleep(1)
            # Digitar Nome
            WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, "formulario:nome"))).send_keys(nome)
            time.sleep(0.5)
            # Habilitar Checkbox do FGTS
            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:fgts"))).click()
            time.sleep(0.5)
            # Aguardar checkbox condicional 'Proporcionalizar FGTS' e desmarcar
            WebDriverWait(driver, self.delay).until(
                EC.visibility_of_element_located((By.ID, "formulario:proporcionalizarFGTS"))).click()
            time.sleep(0.5)
            # Colocar a Competência Inicial
            campo_competencia_inicial = WebDriverWait(driver, self.delay).until(
                EC.presence_of_element_located((By.NAME, "formulario:competenciaInicialInputDate")))
            campo_competencia_inicial.send_keys(Keys.BACKSPACE)
            time.sleep(0.5)
            campo_competencia_inicial.send_keys(data_inicial)
            time.sleep(0.5)
            # Colocar a Competência Final
            campo_competencia_final = WebDriverWait(driver, self.delay).until(
                EC.presence_of_element_located((By.NAME, "formulario:competenciaFinalInputDate")))
            campo_competencia_final.send_keys(Keys.BACKSPACE)
            time.sleep(0.5)
            campo_competencia_final.send_keys(data_final)
            time.sleep(0.5)
            # Colocar inicialmente o valor 0
            WebDriverWait(driver, self.delay).until(
                EC.presence_of_element_located((By.NAME, "formulario:valorParaBaseDeCalculo"))).send_keys(0)
            time.sleep(0.5)
            # Clique em adicionar
            WebDriverWait(driver, self.delay).until(
                EC.element_to_be_clickable((By.ID, "formulario:cmdGerarOcorrencias"))).click()
            # Aguardar
            self.objTools.aguardar_carregamento(driver)
            time.sleep(1)

        # Datas coletadas da planilha
        data_inicial = planilha.iloc[1, 31]
        data_final = planilha.iloc[2, 31]
        qtd_competencias = planilha.iloc[2, 32]

        # Coluna 'AH' da Planilha Base
        coluna_1 = planilha.iloc[0, 33]
        # Coluna 'AK' da Planilha Base
        coluna_2 = planilha.iloc[0, 36]
        # Coluna 'AN' da Planilha Base
        coluna_3 = planilha.iloc[0, 39]

        # Tratamento de exceção — Se o valor da competência estiver em branco, ignorar
        try:
            # Conversão - 'Competência Inicial' - float -> datetime
            data_inicial = xlrd.xldate_as_datetime(data_inicial, 0)
            data_inicial = data_inicial.strftime("%m/%Y")
        except ValueError:
            pass

        # Tratamento de exceção — Se o valor da competência estiver em branco, ignorar
        try:
            # Conversão - 'Competência Final' - float -> datetime
            data_final = xlrd.xldate_as_datetime(data_final, 0)
            data_final = data_final.strftime("%m/%Y")
        except ValueError:
            pass

        print("- Competência Inicial: ", data_inicial)
        print("- Competência Final: ", data_final)
        print("- Qtd de Competências: ", qtd_competencias)
        print("- Indicador da coluna FGTS: ", coluna_1, ' - ', coluna_2, ' - ', coluna_3)
        print()

        # Verificar a utilização do Ckeckbox Geral
        plan = pd.read_excel(self.source, sheet_name="PJE HIST-VAL", header=1)
        # Dados referente a primeira parte do FGTS (Incide/Recolhido)
        quantidade_competencias = plan.iloc[0, 32]
        quantidade_indice = plan.iloc[0, 34]
        quantidade_recolhido = plan.iloc[0, 35]
        # Dados referente a segunda parte do FGTS (Incide/Recolhido)
        quantidade_indice_2 = plan.iloc[0, 37]
        quantidade_recolhido_2 = plan.iloc[0, 38]
        # Dados referente a terceira parte do FGTS (Incide/Recolhido)
        quantidade_indice_3 = plan.iloc[0, 40]
        quantidade_recolhido_3 = plan.iloc[0, 41]
        print("---------------------------------------------------------------------------------------")
        print(f"- Qtd. Competências: {quantidade_competencias}")
        print("---------------------------------------------------------------------------------------")
        print(f"1 - Qtd. Incide: {quantidade_indice} | - Qtd. Recolhido: {quantidade_recolhido}")
        print("---------------------------------------------------------------------------------------")
        print(f"2 - Qtd. Incide: {quantidade_indice_2} | - Qtd. Recolhido: {quantidade_recolhido_2}")
        print("---------------------------------------------------------------------------------------")
        print(f"3 - Qtd. Incide: {quantidade_indice_3} | - Qtd. Recolhido: {quantidade_recolhido_3}")
        print("---------------------------------------------------------------------------------------")

        if coluna_1 == 1:
            # Coletar nome da coluna
            nome = planilha.iloc[1, 33]
            # PJeCalc
            montar_estrutura_dados()
            # Após habilitar o checkbox FGTS, por padrão o checkbox INCIDE FGTS vem ativo. Então, para a lógica do código fazer sentido, desabitarei o checkbox logo em seguida.
            WebDriverWait(driver, self.delay).until(
                EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3"))).click()
            # Iterar preenchendo os valores
            for i in range(3, qtd_competencias + 3):
                base = planilha.iloc[i, 33]
                # Conversão para duas casas decimais
                base = f"{base:.2f}"
                # PJeCalc - Preencher os valores
                try:
                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.NAME, f"formulario:listagemMC:{i - 3}:valor"))).send_keys(
                        base)
                except TimeoutException:
                    # Sair
                    break

            # Tempo de controle
            time.sleep(1)

            # - INCIDE FGTS - #
            # Comparação da quantidade de competencias com a quantidade da coluna Indice selecionados
            if quantidade_competencias == quantidade_indice:
                # Incide FGTS
                checkbox = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3")))
                status = checkbox.is_selected()
                print("- Status do Ckeckbox - 'Incide FGTS' - ", status)
                if not status:
                    checkbox.click()
                    # checkbox.click()
                    print("- Checkbos Geral - (Incide FGTS) - Foi Habilitado.")
            else:
                # print("- 1 - Incide FGTS - É diferente da quantidade de competencias!")
                # PJeCalc - Desmarcar o Checkbox Principal do Incide FGTS
                checkboxIncide1 = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3")))
                print("1 - Incide FGTS - Status do Checkbox: ", checkboxIncide1.is_selected())
                if checkboxIncide1.is_selected():
                    checkboxIncide1.click()
                # Tempo de controle
                time.sleep(0.5)

                # Checkbox - 'Incide FGTS' - Marcar ou desmarcar
                for i in range(3, qtd_competencias + 3):
                    # Coluna 'AI' da Planilha Base
                    incide = planilha.iloc[i, 34]
                    # Verificar se o campo foi habilitado na planilha
                    if incide == "Sim":
                        # PJeCalc
                        try:
                            checkbox = WebDriverWait(driver, 3).until(
                                EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:incideFGTS")))
                            if checkbox.is_selected():
                                pass
                            else:
                                checkbox.click()
                        except TimeoutException:
                            break
                    else:
                        continue

            # - FGTS RECOLHIDO - #
            if quantidade_competencias == quantidade_recolhido:
                # FGTS Recolhido
                checkbox_2 = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel1")))
                status = checkbox_2.is_selected()
                print("1 - FGTS Recolhido - Status do Ckeckbox: ", status)
                if not status:
                    checkbox_2.click()
                    print("- Checkbos Geral - (FGTS Recolhido) - Foi Habilitado.")
            else:
                # print("1 - FGTS Recolhido - É diferente da quantidade de competencias!")
                # PJeCalc - Desmarcar o Checkbox Principal do Incide FGTS
                checkboxRecolhido = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel1")))
                print("1 - FGTS Recolhido - Status do Checkbox: ", checkboxRecolhido.is_selected())
                if checkboxRecolhido.is_selected():
                    checkboxRecolhido.click()
                # Tempo de controle
                time.sleep(0.5)
                # Checkbox - 'Incide FGTS' - Marcar ou desmarcar
                for i in range(3, qtd_competencias + 3):
                    # Coluna 'AJ' da Planilha Base
                    recolhido = planilha.iloc[i, 35]
                    # Verificação se o campo foi selecionado
                    if recolhido == "Sim":
                        # PJeCalc
                        try:
                            checkbox = WebDriverWait(driver, 3).until(
                                EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:recolhidoFGTS")))
                            if checkbox.is_selected():
                                pass
                            else:
                                checkbox.click()
                        except TimeoutException:
                            break
                    else:
                        continue

            # Tempo de controle
            time.sleep(1)
            # Salvar Operação
            self.salvar(driver)
            self.objTools.aguardar_carregamento(driver)
            time.sleep(1)
            self.verificacao(driver)
            time.sleep(2)

        if coluna_2 == 1:
            # Coletar nome da coluna
            nome = planilha.iloc[1, 36]
            # PJeCalc
            montar_estrutura_dados()
            # Após habilitar o checkbox FGTS, por padrão o checkbox INCIDE FGTS vem ativo. Então, para a lógica do código fazer sentido, desabitarei o checkbox logo em seguida.
            WebDriverWait(driver, self.delay).until(
                EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3"))).click()
            # Iterar preenchendo os valores
            for i in range(3, qtd_competencias + 3):
                base = planilha.iloc[i, 36]
                # Conversão para duas casas decimais
                base = f"{base:.2f}"
                # PJeCalc - Preencher os valores
                try:
                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.NAME, f"formulario:listagemMC:{i - 3}:valor"))).send_keys(
                        base)
                except TimeoutException:
                    # Sair
                    break
            # Tempo de controle
            time.sleep(1)

            # - INCIDE FGTS - #
            if quantidade_competencias == quantidade_indice_2:
                # Incide FGTS
                checkbox = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3")))
                print("2 - Incide FGTS - Status do Ckeckbox: ", checkbox.is_selected())
                if not checkbox.is_selected():
                    checkbox.click()
                    # checkbox.click()
                    print("2 - Checkbos Geral - (Incide FGTS) - Foi Habilitado.")
            else:
                # print("2 - Incide FGTS - É diferente da quantidade de competencias!")
                # PJeCalc - Desmarcar o Checkbox Principal do Incide FGTS
                checkboxIncideFGTS = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3")))
                print("- Status Checkbox 'Incide FGTS': ", checkboxIncideFGTS.is_selected())
                if checkboxIncideFGTS.is_selected():
                    checkboxIncideFGTS.click()

                time.sleep(0.5)
                # Checkbox - 'Incide FGTS' - Marcar ou desmarcar
                for i in range(3, qtd_competencias + 3):
                    # Coluna 'AL' da Planilha Base
                    incide = planilha.iloc[i, 37]
                    if incide == "Sim":
                        # PJeCalc
                        try:
                            checkbox = WebDriverWait(driver, 3).until(
                                EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:incideFGTS")))
                            if checkbox.is_selected():
                                pass
                            else:
                                checkbox.click()
                        except TimeoutException:
                            break
                    else:
                        continue

            # - FGTS RECOLHIDO - #
            if quantidade_competencias == quantidade_recolhido_2:
                checkbox_2 = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel1")))
                print("2 - FGTS Recolhido - Status do Ckeckbox: ", checkbox_2.is_selected())
                if not checkbox_2.is_selected():
                    checkbox_2.click()
                    print("2 - Checkbos Geral - (FGTS Recolhido) - Foi Habilitado.")
            else:
                # print("- 2 - FGTS Recolhido - É diferente da quantidade de competencias!")
                # PJeCalc - Desmarcar o Checkbox Principal do Incide FGTS
                checkboxRecolhido2 = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel1")))
                print("2 - FGTS Recolhido - Status do Checkbox: ", checkboxRecolhido2.is_selected())
                if checkboxRecolhido2.is_selected():
                    checkboxRecolhido2.click()
                # Tempo de controle
                time.sleep(0.5)
                # Checkbox - 'Incide FGTS' - Marcar ou desmarcar
                for i in range(3, qtd_competencias + 3):
                    # Coluna 'AM' da Planilha Base
                    recolhido = planilha.iloc[i, 38]
                    # Verificação se o campo foi selecionado
                    if recolhido == "Sim":
                        # PJeCalc
                        try:
                            checkbox = WebDriverWait(driver, 3).until(
                                EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:recolhidoFGTS")))
                            if checkbox.is_selected():
                                pass
                            else:
                                checkbox.click()
                        except TimeoutException:
                            # Sair
                            break
                    else:
                        continue

            # Tempo de controle
            time.sleep(1)
            # Salvar Operação
            self.salvar(driver)
            self.objTools.aguardar_carregamento(driver)
            time.sleep(1)
            self.verificacao(driver)
            time.sleep(2)

        if coluna_3 == 1:
            # Coletar nome da coluna
            nome = planilha.iloc[1, 39]
            # PJeCalc
            montar_estrutura_dados()
            # Após habilitar o checkbox FGTS, por padrão o checkbox INCIDE FGTS vem ativo. Então, para a lógica do código fazer sentido, desabitarei o checkbox logo em seguida.
            WebDriverWait(driver, self.delay).until(
                EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3"))).click()
            # Iterar e preecher os valores
            for i in range(3, qtd_competencias + 3):
                base = planilha.iloc[i, 39]
                # Conversão para duas casas decimais
                base = f"{base:.2f}"
                # PJeCalc - Preencher os valores
                try:
                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.NAME, f"formulario:listagemMC:{i - 3}:valor"))).send_keys(
                        base)
                except TimeoutException:
                    # Sair
                    break
            # Tempo de controle
            time.sleep(1)

            # - INCIDE FGTS - #
            if quantidade_competencias == quantidade_indice_3:
                # Incide FGTS
                checkbox = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3")))
                print("3 - Incide FGTS - Status do Ckeckbox: ", checkbox.is_selected())
                if not checkbox.is_selected():
                    checkbox.click()
                    # checkbox.click()
                    print("3 - Checkbos Geral - (Incide FGTS) - Foi Habilitado.")
            else:
                # print("3 - Incide FGTS - É diferente da quantidade de competencias!")
                # PJeCalc - Desmarcar o Checkbox Principal do Incide FGTS
                checkboxIncideFGTS3 = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel3")))
                print("3 - Incide FGTS- Status do Checkbox: ", checkboxIncideFGTS3.is_selected())
                if checkboxIncideFGTS3.is_selected():
                    checkboxIncideFGTS3.click()
                # Tempo de controle
                time.sleep(0.5)
                # Checkbox - 'Incide FGTS' - Marcar ou desmarcar
                for i in range(3, qtd_competencias + 3):
                    # Coluna 'AO' da Planilha Base
                    incide = planilha.iloc[i, 40]
                    # Verificação se o campo foi habilitado
                    if incide == "Sim":
                        # PJeCalc
                        try:
                            checkbox = WebDriverWait(driver, 3).until(
                                EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:incideFGTS")))
                            if checkbox.is_selected():
                                pass
                            else:
                                checkbox.click()
                        except TimeoutException:
                            # Sair
                            break
                    else:
                        continue

            # - FGTS RECOLHIDO - #
            if quantidade_competencias == quantidade_recolhido_3:
                checkbox_2 = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel1")))
                print("- Status do Ckeckbox - 'FGTS Recolhido' - ", checkbox_2.is_selected())
                if not checkbox_2.is_selected():
                    checkbox_2.click()
                    print("- Checkbos Geral - (FGTS Recolhido) - Foi Habilitado.")
            else:
                # PJeCalc - Desmarcar o Checkbox Principal do Incide FGTS
                checkboxRecolhido3 = WebDriverWait(driver, self.delay).until(
                    EC.element_to_be_clickable((By.ID, "selecionarTodosLabel1")))
                print("3 - Status do Checkbox 'FGTS Recolhido': ", checkboxRecolhido3.is_selected())
                if checkboxRecolhido3.is_selected():
                    checkboxRecolhido3.click()
                # Tempo de controle
                time.sleep(0.5)
                # Checkbox - 'FGTS Recolhido' - Marcar ou desmarcar
                for i in range(3, qtd_competencias + 3):
                    # Coluna 'AP' da Planilha Base
                    recolhido = planilha.iloc[i, 41]
                    # Verificação se o campo foi habilitado
                    if recolhido == "Sim":
                        # PJeCalc
                        try:
                            checkbox = WebDriverWait(driver, 3).until(
                                EC.element_to_be_clickable((By.ID, f"formulario:listagemMC:{i - 3}:recolhidoFGTS")))
                            if checkbox.is_selected():
                                pass
                            else:
                                checkbox.click()
                        except TimeoutException:
                            # Sair
                            break
                    else:
                        continue
            # Tempo de controle
            time.sleep(1)
            # Salvar Operação
            self.salvar(driver)
            self.objTools.aguardar_carregamento(driver)
            time.sleep(1)
            self.verificacao(driver)
            time.sleep(2)

    def main_historico_salarial(self, driver, admissao, rescisao, inicio_calculo, termino_calculo):

        print('\n# ========== [HISTORICO_SALARIAL] ========== #\n')
        self.entrar_historico_salarial(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(1)
        # Para Planilha base v3.34 - Início
        self.preencher_dados_hist_salarial_and_fgts_v3_34(driver, admissao, rescisao, inicio_calculo, termino_calculo)
        time.sleep(self.delayG)
        self.preencher_hist_fgts_novo(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        # - Limpar Temp
        self.objTools.limparFilesTemp()
        print('\n# ===================================== #\n')