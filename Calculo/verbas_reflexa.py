from time import sleep
import pandas as pd
from . verbas import Verbas
from Tools.pjecalc_control import Control
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait, TimeoutException


class VerbaReflexa(Verbas):


    def __init__(self, source):
        super().__init__(source)
        self.delay = 10
        self.delayG = 2.5
        self.delayP = 1
        self.delayN = 0.5
        self.objTime = Control()
        # self.planilha_tbverbabd = self.planilha_base
        self.planilha_tbreflexa = pd.read_excel(source, sheet_name="TBREFLEXA")


    def click_parametrizarReflexo(self, driver, indice):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, f'formulario:listagem:0:listaReflexo:{indice}:j_id573'))).click()


    def click_btnManual(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:incluir'))).click()


    def filtro_parcelaReflexa(self, snmverbareflexa, snmverbaexpressopjecalc):

        if snmverbaexpressopjecalc != "":
            filtro = (self.planilha_tbreflexa['SNMVERBAREFLEXA'] == snmverbareflexa) & (self.planilha_tbreflexa['SNMVERBAEXPRESSOPJECALC'] == snmverbaexpressopjecalc)
        else:
            filtro = (self.planilha_tbreflexa['SNMVERBAREFLEXA'] == snmverbareflexa)

        indice_registro = self.planilha_tbreflexa.index[filtro].tolist()

        return indice_registro


    def selecionar_comportamento(self, driver, valor):

        # GABARITO:
        # VM: Valor Mensal
        # MA: Média Pelo Valor Absoluto
        # MC: Média Pelo Valor  Corrigido
        # MQ: Média Pela Quantidade

        dicionario = {"VM": 0, "MA": 1, "MC": 2, "MQ": 3}

        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, f'formulario:comportamentoDoReflexo:{dicionario[valor]}'))).click()
        except TimeoutException as e:
            print(f"- [except][Comportamento]: {e}")

        del valor


    def selecionar_PeriodoDaMedia(self, driver, valor):

        # GABARITO:
        # PA: Período Aquisitivo
        # AC: Ano Civil
        # DM: Últimos Doze Meses do Contrato
        # DV: Doze Meses Anteriores ao Vencimento da Verba

        dicionario = {"PA": "PERIODO_AQUISITIVO", "AC": "ANO_CIVIL", "DM": "ULTIMOS_DOZE_MESES_DO_CONTRATO", "DV": "DOZE_MESES_ANTERIORES_AO_VENCIMENTO_DA_PARCELA"}

        try:
            if WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.NAME, 'formulario:j_id339'))):

                element = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.NAME, 'formulario:j_id339')))
                fieldSelector = Select(element)
                fieldSelector.select_by_value(dicionario[valor])

        except TimeoutException as e:
            print(f"- [except][PeríodoDaMédia]: {e}")

        del valor


    def selecionar_TratamentoDaFracaoDeMes(self, driver, valor):

        # GABARITO:
        # M: Manter
        # I: Integralizar
        # D: Desprezar
        # DMQ: Desprezar Menor que 15 Dias


        dicionario = {"M": "MANTER", "I": "INTEGRALIZAR", "D": "DESPREZAR", "DMQ": "DESPREZAR_MENOR_QUE_15_DIAS"}

        try:
            if WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.NAME, 'formulario:j_id344'))):

                element = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.NAME, 'formulario:j_id344')))
                fieldSelector = Select(element)
                fieldSelector.select_by_value(dicionario[valor])

        except TimeoutException as e:
            print(f"- [except][TratamentoDaFracaoDeMes]: {e}")

        del valor

    def selecionar_verbaPrincipal(self, driver, pk, valor):

        try:
            filtro = (self.planilha_base['IIDVERBAJRS'] == pk) & (self.planilha_base['SNMDESCRICAOVERBA'] == valor)
            indice_parcela_principal = self.planilha_base.index[filtro].tolist()
            parcela_princial = self.planilha_base.loc[indice_parcela_principal[0], 'SNMDESCRICAOVERBA']
            parcela_princial = parcela_princial.upper()
            print(f"- Parcela Princial: {parcela_princial} |- Pk: {pk}")
        except IndexError as e:
            print(f"- [except][parcela_principal_nao_localizada]: {e}")
            parcela_princial = valor.upper()

        try:
            if WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, 'formulario:baseVerbaDeCalculo'))):
                campoSelecao = WebDriverWait(driver, self.delay).until(
                    EC.presence_of_element_located((By.ID, 'formulario:baseVerbaDeCalculo')))
                elementoSelecao = Select(campoSelecao)
                elementoSelecao.select_by_visible_text(parcela_princial)
        except TimeoutException:
            print("- [except][Campo de seleçãoo 'VERDA' não localizado]")
        except:
            pass

        del parcela_princial
        del valor
        del pk


    def functon_main_parcelaReflexa(self, driver, indice, iidverbajrs, snmverbaexpressopjecalc):

        stpvalor = ""
        sflincidenciainss = ""
        sflincidenciairpf = ""
        sflincidenciafgts = ""
        sflincidenciaprevprivada = ""
        sflincidenciapensao = ""

        self.click_btnManual(driver)
        self.objTime.aguardar_carregamento(driver)
        sleep(self.delayG)

        vl = indice[0]
        for vc in self.planilha_tbreflexa.columns:

            dado = self.planilha_tbreflexa.loc[vl, vc]

            result = self.verificar_valorTb(dado)
            if result is False:
                continue

            if vc == 'SNMDESCRICAOVERBAREFLEXA':
                snmdescricaoverba = dado
                print("- SNMDESCRICAOVERBA: ", snmdescricaoverba)
                self.modificar_snmdescricaoverba(driver, snmdescricaoverba)
                continue

            if vc == 'ICDASSUNTO':
                icdassunto = dado
                print("- ICDASSUNTO: ", icdassunto)
                self.selecionarAssuntoCNJ(driver)
                sleep(self.delayN)
                self.definirAssuntoCNJ(driver, icdassunto)
                sleep(self.delayN)
                self.btnSelecionarAssuntoCNJ(driver)
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
                    sleep(self.delayP)
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

                if vc == 'STPCOMPORTAMENTOREFLEXO':
                    stpcomportamentoreflexo = dado
                    print(f"- STPCOMPORTAMENTOREFLEXO: {stpcomportamentoreflexo} - Tipo: {type(stpcomportamentoreflexo)}")
                    if stpcomportamentoreflexo != "" or stpcomportamentoreflexo is not None:
                        self.selecionar_comportamento(driver, stpcomportamentoreflexo)
                    continue

                if vc == 'STPPERIODOMEDIAREFLEXO':
                    stpperiodomediareflexo = dado
                    print(f"- STPPERIODOMEDIAREFLEXO: {stpperiodomediareflexo} - Tipo: {type(stpperiodomediareflexo)}")
                    if stpperiodomediareflexo != "" or stpperiodomediareflexo is not None:
                        self.selecionar_PeriodoDaMedia(driver, stpperiodomediareflexo)
                    continue

                if vc == 'STPTRATAMENTOFRACAOMESREFLEXO':
                    stptratamentofracaomesreflexo = dado
                    print(f"- STPTRATAMENTOFRACAOMESREFLEXO: {stptratamentofracaomesreflexo} - Tipo: {type(stptratamentofracaomesreflexo)}")
                    if stptratamentofracaomesreflexo != "" or stptratamentofracaomesreflexo is not None:
                        self.selecionar_TratamentoDaFracaoDeMes(driver, stptratamentofracaomesreflexo)
                    continue


                # if vc == 'SFLPROPORCIONALIDADE':
                #     sflproporcionalidade = dado
                #
                # if vc == 'STPBASEDEVIDO':
                #     print("- STPBASEDEVIDO: ", dado)
                #     self.modificar_stpbasedevido(driver, dado)
                #     if dado == 'HS':
                #         self.incluir_bases_HistoricoSalarial(driver, iidverbajrs)
                #     else:
                #         print("- SFLPROPORCIONALIDADE: ", sflproporcionalidade)
                #         self.marcar_sflproporcionalidade_basecalculo(driver, sflproporcionalidade)
                #
                #     # self.selecionar_verba(driver, parcela)
                #     # self.click_btnIncluirVerba(driver)
                #     # sleep(1)
                #     # ADICIONAR HISTÓRICO
                #     continue

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


        # self.selecionar_verba(driver, snmverbaexpressopjecalc)
        self.selecionar_verbaPrincipal(driver, iidverbajrs, snmverbaexpressopjecalc)
        self.click_btnIncluirVerba(driver)
        self.objTime.aguardar_carregamento(driver)
        sleep(1)


if __name__ == '__main__':
    pass