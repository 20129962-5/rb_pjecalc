from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import os
import time
import xlrd
import gc
#
from Calculo.pjecalc_dados_calculo import DadosCalculo
from Tools.pjecalc_control import Control


class Honorarios(DadosCalculo):

    def __init__(self, source):
        super().__init__(source)
        self.planilha_base = self.planilha_base
        self.tamanho_plan = len(self.planilha_base)
        self.objTools = Control()
        self.delay = 10
        self.delayG = 1.5


    var_controle_float = 1.1
    var_controle_data = datetime.now()
    var_controle_string = ""
    honorarios = ""
    qtd_honorarios_periciais = ''
    honorarios_periciais_tipo_pericia_pje_calc_1 = ''
    honorarios_periciais_descricao_1 = ''
    honorarios_periciais_devedor_1 = ''
    honorarios_periciais_pgto_rcte_1 = ''
    honorarios_periciais_data_1 = ''
    honorarios_periciais_valor_1 = ''
    honorarios_periciais_nome_credor_1 = ''
    honorarios_periciais_cm = ""

    # HONORÁRIOS CONTRATUAIS
    qtd_honorarios_contratuais = ''
    honorarios_contratuais_tipob_1 = ''
    honorarios_contratuais_descricao_1 = ''
    honorarios_contratuais_devedor_1 = ''
    honorarios_contratuais_partic_1 = ''
    honorarios_contratuais_tipo_1 = ''
    honorarios_contratuais_aliquota_1 = ''
    honorarios_contratuais_base_1 = ''
    honorarios_contratuais_nome_1 = ''

    # HONORÁRIOS SUCUMBENCIAIS
    qtd_honorarios_sucumbenciais = ''
    honorarios_sucumbenciais_tipo_e_descricao_1 = ''
    honorarios_sucumbenciais_descricao_1 = ''
    honorarios_sucumbenciais_devedor_1 = ''
    honorarios_sucumbenciais_form_pagto_1 = ''
    honorarios_sucumbenciais_tipo_1 = ''
    # honorarios_sucumbenciais_dt_base_correc_1 = ''
    honorarios_sucumbenciais_vcto_1 = ''
    honorarios_sucumbenciais_valor_1 = ''
    honorarios_sucumbenciais_aliquota_1 = ''
    honorarios_sucumbenciais_base_1 = ''
    honorarios_sucumbenciais_nome_1 = ''


    def acessar_horonarios(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageHonorarios"))).click()
        # WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:j_id46:0:j_id49:20:j_id54'))).click()

    def verificacao(self, driver):

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            # file_txt_log.write('* ' + campo + ' | ' + status + '\n')
            if self.honorarios_periciais_tipo_pericia_pje_calc_1 != "-":
                self.honorarios = self.honorarios_periciais_tipo_pericia_pje_calc_1
            elif self.honorarios_contratuais_tipob_1 != "-":
                self.honorarios = self.honorarios_contratuais_tipob_1
            elif self.honorarios_sucumbenciais_tipo_e_descricao_1 != "-":
                self.honorarios = self.honorarios_sucumbenciais_tipo_e_descricao_1
            file_txt_log.write(f'- {campo} : {self.honorarios} | {status}\n')
            return file_txt_log.close()

        def cancelar_operacao():
            btn_cancelar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:cancelar')))
            btn_cancelar.click()
            self.objTools.aguardar_carregamento(driver)

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                # print('* Operação realizada com sucesso.')
                gerar_relatorio('Honorários', 'Ok')
            elif 'Existem erros no formulário.' in msg or 'erro' in msg or 'Erro' in msg:
                # print('* ERRO!', msg)
                gerar_relatorio('Honorários', '---------- Erro! ----------')
                cancelar_operacao()
            else:
                print('#- ', msg)
        except TimeoutException:
            print('- [Except][Honorários] - Elemento não encontrado/A Página demorou para responder. Encerrando...')

        # Tempo de controle
        time.sleep(2)

    def selecionar_novo(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:incluir'))).click()

    def verificar_qtdHonorariosPericiais(self):
        pass
    def verificar_qtdHonorariosContratuais(self):
        pass
    def verificar_qtdHonorariosSucumbenciais(self):
        pass

    def preencher_dados_honorarios_periciais(self, driver):

        indice = 1
        # Honorários Periciais
        for i in range(self.tamanho_plan):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                continue
            elif coluna_informacao == "qtd_honorarios_periciais":
                qtd_honorarios_periciais = coluna_informacao
                print("\n- Quantidade - Honorários Periciais - ", qtd_honorarios_periciais)
                # Condição para verificar se há honorários
                if qtd_honorarios_periciais == 0:
                    break
            # Quantidade a preencher
            elif coluna_identificador == "qtd_honorarios_periciais":
                self.qtd_honorarios_periciais = coluna_informacao
                print('- Quantidade - Honorários Periciais: ', self.qtd_honorarios_periciais)
            # Tipo de Honorário
            elif coluna_identificador == f"honorarios_periciais_tipo_pericia_pje_calc_{indice}":
                self.honorarios_periciais_tipo_pericia_pje_calc_1 = coluna_informacao
                print('- Tipo de Honorário *: ', self.honorarios_periciais_tipo_pericia_pje_calc_1)
            # Descrição
            elif coluna_identificador == f"honorarios_periciais_descricao_{indice}":
                self.honorarios_periciais_descricao_1 = coluna_informacao
                print('- Descrição *: ', self.honorarios_periciais_descricao_1)
            # Devedor
            elif coluna_identificador == f"honorarios_periciais_devedor_{indice}":
                self.honorarios_periciais_devedor_1 = coluna_informacao
                print('- Devedor *: ', self.honorarios_periciais_devedor_1)
            # Pagamento Reclamante
            elif coluna_identificador == f"honorarios_periciais_pgto_rcte_{indice}":
                self.honorarios_periciais_pgto_rcte_1 = coluna_informacao
                print('- Pagamento Reclamante: ', self.honorarios_periciais_pgto_rcte_1)
            # Vencimento
            elif coluna_identificador == f"honorarios_periciais_data_{indice}":
                self.honorarios_periciais_data_1 = coluna_informacao

                if type(self.honorarios_periciais_data_1) != type(self.var_controle_string):
                    self.honorarios_periciais_data_1 = xlrd.xldate_as_datetime(self.honorarios_periciais_data_1, 0)
                    self.honorarios_periciais_data_1 = self.honorarios_periciais_data_1.strftime('%d/%m/%Y')
                print('- Vencimento *: ', self.honorarios_periciais_data_1)
            # Valor
            elif coluna_identificador == f"honorarios_periciais_valor_{indice}":
                self.honorarios_periciais_valor_1 = coluna_informacao

                if type(self.honorarios_periciais_valor_1) != type(self.var_controle_string):
                    self.honorarios_periciais_valor_1 = float(self.honorarios_periciais_valor_1)
                    self.honorarios_periciais_valor_1 = '{:.2f}'.format(self.honorarios_periciais_valor_1)
                print('- Valor *: ', self.honorarios_periciais_valor_1)
            # Credor
            elif coluna_identificador == f"honorarios_periciais_nome_credor_{indice}":
                self.honorarios_periciais_nome_credor_1 = coluna_informacao
                print('- Credor - Nome Completo *: ', self.honorarios_periciais_nome_credor_1)
            # Correção Monetária — Utilizar outro índice
            elif coluna_identificador == f"honorarios_periciais_cm_{indice}":
                self.honorarios_periciais_cm = coluna_informacao
                print('- Índice: ', self.honorarios_periciais_cm)

                # ** PREENCHER CONTEÚDO NO PJECALC
                for j in range(self.tamanho_plan):

                    # Verificar se há conteúdo para preencher
                    if '-' in self.honorarios_periciais_descricao_1:
                        break
                    else:
                        # Novo
                        self.selecionar_novo(driver)
                        # Aguardar
                        self.objTools.aguardar_carregamento(driver)
                        # Tempo de controle
                        time.sleep(2)

                        # Tipo de Honorário - Inicialmente, conteúdo fixo
                        campo_tipo_honorario = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:tpHonorario')))
                        selecionar_tipo_honorario = Select(campo_tipo_honorario)
                        selecionar_tipo_honorario.select_by_visible_text(self.honorarios_periciais_tipo_pericia_pje_calc_1)

                        # Tempo de controle
                        time.sleep(1)

                        # Descrição
                        campo_descricao = WebDriverWait(driver, self.delay).until(
                            EC.presence_of_element_located((By.NAME, 'formulario:descricao')))
                        campo_descricao.send_keys(Keys.CONTROL, 'a')
                        campo_descricao.send_keys(self.honorarios_periciais_descricao_1)

                        # Devedor
                        if 'RECLAMADA' in self.honorarios_periciais_devedor_1:
                            opcao_reclamado = WebDriverWait(driver, self.delay).until(
                                EC.element_to_be_clickable((By.ID, 'formulario:tipoDeDevedor:1')))
                            opcao_reclamado.click()
                        elif 'RECLAMANTE' in self.honorarios_periciais_devedor_1:
                            opcao_reclamante = WebDriverWait(driver, self.delay).until(
                                EC.element_to_be_clickable((By.ID, 'formulario:tipoDeDevedor:0')))
                            opcao_reclamante.click()

                            # Tempo de controle
                            time.sleep(2)
                            # Aguardar campos condicionais
                            # self.aguardar_campos_condicionais_honorarios_devedor()
                            # Tempo de controle
                            # time.sleep(1)

                            if 'DESCONTAR' in self.honorarios_periciais_pgto_rcte_1:
                                opcao_descontar = WebDriverWait(driver, self.delay).until(
                                    EC.element_to_be_clickable((By.ID, 'formulario:tipoCobrancaReclamante:0')))
                                opcao_descontar.click()
                            else:
                                opcao_cobrar = WebDriverWait(driver, self.delay).until(
                                    EC.element_to_be_clickable((By.ID, 'formulario:tipoCobrancaReclamante:1')))
                                opcao_cobrar.click()
                        else:
                            print('- Devedor - * Nenhuma condição atendida! (Verificar)')

                        # Tempo de controle
                        time.sleep(1)

                        # Tipo Valor (Calculado) - Futuro

                        # Tipo Valor (Informado)
                        selecionar_tipo_informado = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoValor:0')))
                        selecionar_tipo_informado.click()

                        # Tempo de controle
                        time.sleep(2)

                        # Aguardar campos condicionais
                        try:
                            WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.ID, "formulario:dataVencimentoInputDate")))
                        except TimeoutException:
                            time.sleep(2)

                        # Vencimento
                        campo_vencimento = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:dataVencimentoInputDate')))
                        campo_vencimento.send_keys(self.honorarios_periciais_data_1)

                        # Valor
                        campo_valor = WebDriverWait(driver, self.delay).until(
                            EC.presence_of_element_located((By.NAME, 'formulario:valor')))
                        campo_valor.send_keys(self.honorarios_periciais_valor_1)

                        # Tempo de controle
                        time.sleep(1)

                        # --- Nova Funcionalidade ---#
                        # Correção Monetária — Utilizar outro índice
                        if self.honorarios_periciais_cm == "IPCA-E":
                            # Checkbox - 'Utilizar outro índice'
                            WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, "formulario:tipoDeIndiceDeCorrecao:1"))).click()
                            # Listbox - 'IPCA-E'
                            campo_de_selecao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, "formulario:outroIndiceDeCorrecao")))
                            elemento_option = Select(campo_de_selecao)
                            elemento_option.select_by_visible_text(self.honorarios_periciais_cm)

                        # Credor
                        campo_credor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:nomeCredor')))
                        campo_credor.send_keys(self.honorarios_periciais_nome_credor_1)

                        # Salvar
                        self.salvar(driver)
                        # Aguardar Processamento
                        self.objTools.aguardar_carregamento(driver)
                        # Tempo de controle
                        time.sleep(1)
                        # Verificação
                        self.verificacao(driver)

                        break
                        
                # O índice só passa para a próxima iteração, somente após encontrar a última condição
                indice += 1

                # Quebra de Linha
                print('\n') 
   
    def preencher_dados_honorarios_contratuais(self, driver):

        # Honorários Contratuais
        indice = 1
        for i in range(self.tamanho_plan):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                continue
            # ** HONORÁRIOS CONTRATUAIS
            elif coluna_identificador == "qtd_honorarios_contratuais":
                self.qtd_honorarios_contratuais = coluna_informacao
                print('\n- Qtd Honorários Contratuais: ', self.qtd_honorarios_contratuais)
                if self.qtd_honorarios_contratuais == 0:
                    break
            elif coluna_identificador == f"honorarios_contratuais_tipob_{indice}":
                self.honorarios_contratuais_tipob_1 = coluna_informacao
                self.honorarios_contratuais_tipob_1 = self.honorarios_contratuais_tipob_1.title()
                print('- Honorários - Tipo: ', self.honorarios_contratuais_tipob_1)

            elif coluna_identificador == f"honorarios_contratuais_descricao_{indice}":
                self.honorarios_contratuais_descricao_1 = coluna_informacao
                print('- Honorários - Descrição: ', self.honorarios_contratuais_descricao_1)

            elif coluna_identificador == f"honorarios_contratuais_devedor_{indice}":
                self.honorarios_contratuais_devedor_1 = coluna_informacao
                print('- Honorários - Devedor: ', self.honorarios_contratuais_devedor_1)
            #
            elif coluna_identificador == f"honorarios_contratuais_partic_{indice}":
                self.honorarios_contratuais_partic_1 = coluna_informacao
                print('- Honorários - (Descontar/Cobrar) Reclamente: ', self.honorarios_contratuais_partic_1)
            elif coluna_identificador == f"honorarios_contratuais_tipo_{indice}":
                self.honorarios_contratuais_tipo_1 = coluna_informacao
                print('- Honorários - Tipo (Informado/Calculado): ', self.honorarios_contratuais_tipo_1)

            elif coluna_identificador == f"honorarios_contratuais_alíquota_{indice}":
                self.honorarios_contratuais_aliquota_1 = coluna_informacao
                if type(self.honorarios_contratuais_aliquota_1) == type(self.var_controle_float):
                    self.honorarios_contratuais_aliquota_1 = '{:.2%}'.format(self.honorarios_contratuais_aliquota_1)
                print('- Honorários - Alíquota: ', self.honorarios_contratuais_aliquota_1)

            elif coluna_identificador == f"honorarios_contratuais_base_{indice}":
                self.honorarios_contratuais_base_1 = coluna_informacao
                self.honorarios_contratuais_base_1 = self.honorarios_contratuais_base_1.title()
                print('- Honorários - Base Apuração: ', self.honorarios_contratuais_base_1)

            elif coluna_identificador == f"honorarios_contratuais_nome_{indice}":
                self.honorarios_contratuais_nome_1 = coluna_informacao
                print('- Honorários - (Credor/Nome): ', self.honorarios_contratuais_nome_1)

                # ** PREENCHER CONTEÚDO NO PJECALC
                for j in range(len(self.planilha_base)):

                    # Verificar se há conteúdo para preencher
                    if '-' in self.honorarios_contratuais_tipob_1 or self.honorarios_contratuais_partic_1 == 'INFORMAR':
                        print("- O honorário será ignorado (Particularidade=Informar/tipo=null)")
                        break
                    else:
                        # Novo
                        self.selecionar_novo(driver)
                        # Aguardar
                        self.objTools.aguardar_carregamento(driver)
                        # Tempo de controle
                        time.sleep(2)

                        # Tipo de Honorário - Inicialmente, conteúdo fixo
                        campo_tipo_honorario = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:tpHonorario')))
                        selecionar_tipo_honorario = Select(campo_tipo_honorario)
                        selecionar_tipo_honorario.select_by_visible_text(self.honorarios_contratuais_tipob_1)

                        # Tempo de controle
                        time.sleep(1)

                        # Descrição
                        campo_descricao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:descricao')))
                        campo_descricao.send_keys(Keys.CONTROL, 'a')
                        campo_descricao.send_keys(self.honorarios_contratuais_descricao_1)

                        # Devedor
                        if 'RECLAMADA' in self.honorarios_contratuais_devedor_1:
                            opcao_reclamado = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeDevedor:1')))
                            opcao_reclamado.click()
                        elif 'RECLAMANTE' in self.honorarios_contratuais_devedor_1:
                            opcao_reclamante = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeDevedor:0')))
                            opcao_reclamante.click()

                            # Tempo de controle
                            time.sleep(2)
                            # Aguardar campos condicionais

                            if 'DEDUZIR' in self.honorarios_contratuais_partic_1:
                                opcao_descontar = WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.ID, 'formulario:tipoCobrancaReclamante:0')))
                                opcao_descontar.click()
                            else:
                                opcao_cobrar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoCobrancaReclamante:1')))
                                opcao_cobrar.click()
                        else:
                            print('- Devedor - * Nenhuma condição atendida! (Verificar)')

                        # Tempo de controle
                        time.sleep(1)

                        # Tipo Valor (Calculado)
                        if 'CALCULADO' in self.honorarios_contratuais_tipo_1:
                            selecionar_tipo_calculado = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoValor:1')))
                            selecionar_tipo_calculado.click()

                            # Tempo de controle
                            time.sleep(1)

                            # Alíquota
                            campo_aliquota = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:aliquota')))
                            campo_aliquota.send_keys(self.honorarios_contratuais_aliquota_1)

                            # Tempo de controle
                            time.sleep(1)

                            # Base para Apuração
                            campo_seletor_base = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:baseParaApuracao')))
                            selecionar_base = Select(campo_seletor_base)
                            selecionar_base.select_by_visible_text(self.honorarios_contratuais_base_1)

                            # Tempo de controle
                            time.sleep(1)

                            # Credor/Nome Completo
                            campo_nome = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:nomeCredor')))
                            campo_nome.send_keys(self.honorarios_contratuais_nome_1)

                            # Tempo de controle
                            time.sleep(1)

                        # Tipo Valor (Informado)
                        elif 'INFORMADO' in self.honorarios_contratuais_tipo_1:

                            print('- * Honorários Contratuais - "Informado" ainda não desenvolvido.')
                            break

                            # selecionar_tipo_informado = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoValor:0')))
                            # selecionar_tipo_informado.click()
                            #
                            # # Aguardar campos condicionais
                            # self.aguardar_campos_condicionais_honorarios_tipo_valor()
                            # # Tempo de controle
                            # time.sleep(1)
                            #
                            # # Vencimento
                            # campo_vencimento = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:dataVencimentoInputDate')))
                            # campo_vencimento.send_keys('vencimento')
                            #
                            # # Valor
                            # campo_valor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:valor')))
                            # campo_valor.send_keys('valor')
                            #
                            # # Tempo de controle
                            # time.sleep(1)
                            #
                            # # Credor
                            # campo_credor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:nomeCredor')))
                            # campo_credor.send_keys(self.honorarios_periciais_nome_credor_1)

                        # Salvar
                        self.salvar(driver)
                        # Aguardar Processamento
                        self.objTools.aguardar_carregamento(driver)
                        # Tempo de controle
                        time.sleep(1)
                        # Verificação
                        self.verificacao(driver)

                        break

                # O índice só passa para a próxima iteração, somente após encontrar a última condição
                indice += 1

                # Quebra de Linha
                print('\n')

    def preencher_dados_honorarios_sucumbenciais(self, driver):
        # Honorários Sucumbenciais
        indice = 1
        for i in range(self.tamanho_plan):

            coluna_identificador = self.planilha_base.loc[i, 'IDENTIFICADOR']
            coluna_informacao = self.planilha_base.loc[i, 'INFORMACAO']

            # Condição para pular as linhas em branco da coluna Identificador na planilha base
            if type(coluna_identificador) == type(self.var_controle_float):
                continue
            elif coluna_identificador == "qtd_honorarios_sucumbenciais":
                self.qtd_honorarios_sucumbenciais = coluna_informacao
                print("\n- Qtd Honorários Sucumbenciais - ", self.qtd_honorarios_sucumbenciais)
                # Condição para verificar se há honorários
                if self.qtd_honorarios_sucumbenciais == 0:
                    break
            # Tipo de Honorário
            # ** HONORÁRIOS SUCUMBENCIAIS
            elif f'honorarios_sucumbenciais_tipo_e_descricao_{indice}' in coluna_identificador:
                self.honorarios_sucumbenciais_tipo_e_descricao_1 = coluna_informacao
                # Condição para ignorar se o conteúdo é igual a '-'
                if self.honorarios_sucumbenciais_tipo_e_descricao_1 == '-':
                    pass
                else:
                    # print('0.', self.honorarios_sucumbenciais_tipo_e_descricao_1)
                    sucumbenciais_tipo = self.honorarios_sucumbenciais_tipo_e_descricao_1.title()
                    # print('1.', sucumbenciais_tipo)
                    self.honorarios_sucumbenciais_tipo_e_descricao_1 = sucumbenciais_tipo.split()
                    # print('2.', self.honorarios_sucumbenciais_tipo_e_descricao_1)
                    self.honorarios_sucumbenciais_tipo_e_descricao_1 = self.honorarios_sucumbenciais_tipo_e_descricao_1[0].title() + ' ' + self.honorarios_sucumbenciais_tipo_e_descricao_1[1].lower() + ' ' + self.honorarios_sucumbenciais_tipo_e_descricao_1[2].title()
                    # print('3.', self.honorarios_sucumbenciais_tipo_e_descricao_1)
                    print('- Sucumbenciais Tipo: ', self.honorarios_sucumbenciais_tipo_e_descricao_1)

            elif f'honorarios_sucumbenciais_descricao_{indice}' in coluna_identificador:
                self.honorarios_sucumbenciais_descricao_1 = coluna_informacao
                print('- Sucumbenciais - Descrição: ', self.honorarios_sucumbenciais_descricao_1)

            elif f'honorarios_sucumbenciais_devedor_{indice}' in coluna_identificador:
                self.honorarios_sucumbenciais_devedor_1 = coluna_informacao
                print('- Sucumbenciais - Devedor: ', self.honorarios_sucumbenciais_devedor_1)

            elif f'honorarios_sucumbenciais_form_pagto_{indice}' in coluna_identificador:
                self.honorarios_sucumbenciais_form_pagto_1 = coluna_informacao
                print('- Sucumbenciais - Forma de Pagamento: ', self.honorarios_sucumbenciais_form_pagto_1)

            elif f'honorarios_sucumbenciais_tipo_{indice}' in coluna_identificador:
                self.honorarios_sucumbenciais_tipo_1 = coluna_informacao
                print('- Sucumbenciais - Tipo: ', self.honorarios_sucumbenciais_tipo_1)

            elif f'honorarios_sucumbenciais_vcto_{indice}' in coluna_identificador:
                self.honorarios_sucumbenciais_vcto_1 = coluna_informacao
                if type(self.honorarios_sucumbenciais_vcto_1) != type(self.var_controle_string):
                    self.honorarios_sucumbenciais_vcto_1 = xlrd.xldate_as_datetime(self.honorarios_sucumbenciais_vcto_1, 0)
                    self.honorarios_sucumbenciais_vcto_1 = self.honorarios_sucumbenciais_vcto_1.strftime('%d/%m/%Y')
                print('- Sucumbenciais - Vencimento: ', self.honorarios_sucumbenciais_vcto_1)

            elif f'honorarios_sucumbenciais_valor_{indice}' in coluna_identificador:
                self.honorarios_sucumbenciais_valor_1 = coluna_informacao
                if type(self.honorarios_sucumbenciais_valor_1) != type(self.var_controle_string):
                    self.honorarios_sucumbenciais_valor_1 = float(self.honorarios_sucumbenciais_valor_1)
                    self.honorarios_sucumbenciais_valor_1 = '{:.2f}'.format(self.honorarios_sucumbenciais_valor_1)
                print('- Sucumbenciais - Valor: ', self.honorarios_sucumbenciais_valor_1)

            elif f'honorarios_sucumbenciais_aliquota_{indice}' in coluna_identificador:
                self.honorarios_sucumbenciais_aliquota_1 = coluna_informacao
                if type(self.honorarios_sucumbenciais_aliquota_1) == type(self.var_controle_float):
                    self.honorarios_sucumbenciais_aliquota_1 = '{:.2%}'.format(self.honorarios_sucumbenciais_aliquota_1)
                print('- Sucumbenciais - Alíquota: ', self.honorarios_sucumbenciais_aliquota_1)

            elif f'honorarios_sucumbenciais_base_{indice}' in coluna_identificador:
                self.honorarios_sucumbenciais_base_1 = coluna_informacao
                # print('- Sucumbenciais - Base 1: ', self.honorarios_sucumbenciais_base_1)
                self.honorarios_sucumbenciais_base_1 = self.honorarios_sucumbenciais_base_1.title()
                print('- Sucumbenciais - Base: ', self.honorarios_sucumbenciais_base_1)

            elif f'honorarios_sucumbenciais_nome_{indice}' in coluna_identificador:
                self.honorarios_sucumbenciais_nome_1 = coluna_informacao
                print('- Sucumbenciais - Nome: ', self.honorarios_sucumbenciais_nome_1)


            # ** PREENCHER CONTEÚDO NO PJECALC
                for j in range(self.tamanho_plan):

                    # Verificar se há conteúdo para preencher
                    if '-' in self.honorarios_sucumbenciais_tipo_e_descricao_1:
                        break
                    else:
                        # Novo
                        self.selecionar_novo(driver)
                        # Aguardar
                        self.objTools.aguardar_carregamento(driver)
                        # Tempo de controle
                        time.sleep(2)

                        # Tipo de Honorário - Inicialmente, conteúdo fixo
                        campo_tipo_honorario = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:tpHonorario')))
                        selecionar_tipo_honorario = Select(campo_tipo_honorario)
                        selecionar_tipo_honorario.select_by_visible_text(self.honorarios_sucumbenciais_tipo_e_descricao_1)

                        # Tempo de controle
                        time.sleep(1)

                        # Descrição
                        campo_descricao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:descricao')))
                        campo_descricao.send_keys(Keys.CONTROL, 'a')
                        campo_descricao.send_keys(self.honorarios_sucumbenciais_descricao_1)

                        # Devedor
                        if 'RECLAMADO' in self.honorarios_sucumbenciais_devedor_1:
                            opcao_reclamado = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeDevedor:1')))
                            opcao_reclamado.click()
                        elif 'RECLAMANTE' in self.honorarios_sucumbenciais_devedor_1:
                            opcao_reclamante = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeDevedor:0')))
                            opcao_reclamante.click()

                            # Tempo de controle
                            time.sleep(2)

                            # Aguardar campos condicionais
                            if 'DESCONTAR' in self.honorarios_sucumbenciais_form_pagto_1:
                                opcao_descontar = WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.ID, 'formulario:tipoCobrancaReclamante:0')))
                                opcao_descontar.click()
                            else:
                                opcao_cobrar = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoCobrancaReclamante:1')))
                                opcao_cobrar.click()
                        else:
                            print('- Devedor - * Nenhuma condição atendida! (Verificar)')
                            exit()

                        # Tempo de controle
                        time.sleep(1)

                        # Tipo Valor (Calculado)
                        if 'CALCULADO' in self.honorarios_sucumbenciais_tipo_1:
                            selecionar_tipo_calculado = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoValor:1')))
                            selecionar_tipo_calculado.click()

                            # Tempo de controle
                            time.sleep(1)

                            # Alíquota
                            if '-' in self.honorarios_sucumbenciais_aliquota_1:
                                pass
                            else:
                                campo_aliquota = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:aliquota')))
                                campo_aliquota.send_keys(self.honorarios_sucumbenciais_aliquota_1)
                            # Base para Apuração

                            # Tempo de controle
                            time.sleep(1)

                            # Base para Apuração
                            if '-' in self.honorarios_sucumbenciais_base_1:
                                pass
                            else:
                                campo_base_apuracao = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:baseParaApuracao')))
                                selecionar_base = Select(campo_base_apuracao)
                                selecionar_base.select_by_visible_text(self.honorarios_sucumbenciais_base_1)

                            # Tempo de controle
                            time.sleep(1)
                            
                            
                        elif 'INFORMADO' in self.honorarios_sucumbenciais_tipo_1:
                            # Tipo Valor (Informado)
                            selecionar_tipo_informado = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoValor:0')))
                            selecionar_tipo_informado.click()

                            # Tempo de controle
                            time.sleep(2)
                            # Aguardar campos condicionais
                            # Vencimento
                            campo_vencimento = WebDriverWait(driver, self.delay).until(EC.visibility_of_element_located((By.NAME, 'formulario:dataVencimentoInputDate')))
                            campo_vencimento.send_keys(self.honorarios_sucumbenciais_vcto_1)

                            # Tempo de controle
                            time.sleep(1)

                            # Valor
                            campo_valor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:valor')))
                            campo_valor.send_keys(self.honorarios_sucumbenciais_valor_1)
                            
                        else:
                            print('- * Tipo do Valor inválido! Verificar o código.')


                        # Tempo de controle
                        time.sleep(1)

                        # Nome Completo
                        campo_credor = WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.NAME, 'formulario:nomeCredor')))
                        campo_credor.send_keys(self.honorarios_sucumbenciais_nome_1)

                        # Salvar
                        self.salvar(driver)
                        # Aguardar Processamento
                        self.objTools.aguardar_carregamento(driver)
                        # Tempo de controle
                        time.sleep(1)
                        # Verificação
                        self.verificacao(driver)
                        break
                # O índice só passa para a próxima iteração, somente após encontrar a última condição
                indice += 1
                # Quebra de Linha
                print('\n')

    def salvar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:salvar'))).click()

    def main_honorarios(self, driver):

        self.acessar_horonarios(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.preencher_dados_honorarios_periciais(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.preencher_dados_honorarios_contratuais(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.preencher_dados_honorarios_sucumbenciais(driver)
        # Tempo de controle
        time.sleep(self.delayG)

        # - Limpar Temp
        self.objTools.limparFilesTemp()
        gc.collect(generation=0)
        gc.collect(generation=1)
        gc.collect(generation=2)

        print('-- Fim - (Honorários) --')