import os
import pyautogui as pa
from time import sleep

from browser_pjecalc import WebDriver
from Tools.script_get_source_file import GetDadosGUI
from selenium.common.exceptions import WebDriverException

from Tools.pjecalc_control import Control

from Calculo.pjecalc_fgts import FGTS
from Calculo.pjecalc_faltas import Faltas
from Calculo.pjecalc_ferias import Ferias
from Calculo.pjecalc_verbas import VerbasModel
from Calculo.pjecalc_honorarios import Honorarios
from Calculo.pjecalc_custas_judiciais import Custas
from Calculo.pjecalc_dados_calculo import DadosCalculo
from Calculo.pjecalc_correcao_juros_multa import Correcao
from Calculo.pjecalc_pagina_inicial import PjecalcPaginaInicial
from Calculo.pjecalc_historico_salarial import HistoricoSalarial
from Calculo.pjecalc_contribuicao_social import ContribuicaoSocial
from Calculo.pjecalc_multas_e_indenizacoes import MultasIndenizacoes

from Operacoes.pjecalc_liquidar import Liquidar
from Operacoes.pjecalc_imprimir import Imprimir
from Operacoes.pjecalc_exportar import Exportar


def limparFiles():
    # [Zerar arquivos (log.txt e source_plan.txt)]
    l = open(fr"{os.getcwd()}\log.txt", "w")
    l.close()
    # s = open(fr"{os.getcwd()}\source_plan.txt", "w")
    # s.close()


def main():

    objTools = Control()

    # [Limpar Log]
    limparFiles()

    # - Interface Gráfica
    objGUI = GetDadosGUI()
    objGUI.main()
    locale_plabase = objGUI.source_planilha

    try:

        objGeckdriver = WebDriver()
        driver = objGeckdriver.driver_web()
        objHome = PjecalcPaginaInicial()

        try:
            driver.get('http://localhost:9257/pjecalc/pages/principal.jsf')
            sleep(5)
            objHome.clicar_paginaInicial_pjecalc(driver)
            sleep(2)
        except WebDriverException:
            pa.alert(title='Robô-PJeCalc', text='O PJeCalc não encontra-se em execução. Execute-o e inicialize a automação novamente.')
            driver.close()
            exit()

        # ---------- PJeCalc - Página Inicial ---------- #
        objHome.criar_novo_calculo(driver)
        # ---------- PJeCalc - Dados do Cálculo ---------- #
        objDados = DadosCalculo(locale_plabase)
        objDados.registrar_horario_inicial()
        dados = objDados.main_dados_calculo(driver)
        nome_reclamente = dados[0]
        id_processo = dados[1]
        admissao = dados[2]
        rescisao = dados[3]
        inicio_calculo = dados[4]
        termino_calculo = dados[5]
        # numero_processo = dados[6]
        objTools.limparFilesTemp()

        # ---------- PJeCalc - Faltas ---------- #
        objFaltas = Faltas(locale_plabase)
        objFaltas.main_faltas(driver)
        objTools.limparFilesTemp()

        # # ---------- PJeCalc - Férias ---------- #
        obFerias = Ferias(locale_plabase)
        obFerias.main_ferias(driver)
        objTools.limparFilesTemp()

        # ---------- PJeCalc - Histórico Salarial ---------- #
        objHistorico = HistoricoSalarial(locale_plabase)
        objHistorico.main_historico_salarial(driver, admissao, rescisao, inicio_calculo, termino_calculo)
        objTools.limparFilesTemp()

        # ---------- PJeCalc - Verbas ---------- #
        objVerbas = VerbasModel()
        objVerbas.main_verbas(driver)
        objTools.limparFilesTemp()

        # ---------- PJeCalc - FGTS ---------- #
        objFgts = FGTS(locale_plabase)
        objFgts.main_fgts(driver)
        objTools.limparFilesTemp()

        # # ---------- PJeCalc - Contribuição Social ---------- #
        objContribuicaoSocial = ContribuicaoSocial(locale_plabase)
        objContribuicaoSocial.main_contribuicao_social(driver, admissao, rescisao, inicio_calculo, termino_calculo)
        objTools.limparFilesTemp()

        # ---------- PJeCalc - Multas e Indenizações ---------- #
        objMultas = MultasIndenizacoes(locale_plabase)
        objMultas.main_multas_indenizacoes(driver)
        objTools.limparFilesTemp()

        # ---------- PJeCalc - Honorários ---------- #
        objHonorarios = Honorarios(locale_plabase)
        objHonorarios.main_honorarios(driver)
        objTools.limparFilesTemp()

        # ---------- PJeCalc - Custas ---------- #
        objCustas = Custas(locale_plabase)
        objCustas.main_custas(driver)
        objTools.limparFilesTemp()

        # ---------- PJeCalc - Correção ---------- #
        objCorrecao = Correcao(locale_plabase)
        objCorrecao.main_correcao(driver)
        objTools.limparFilesTemp()

        # ---------- PJeCalc - Liquidar ---------- #
        objLiquidar = Liquidar()
        objLiquidar.main_liquidar(driver)
        objTools.limparFilesTemp()

        # ---------- PJeCalc - Imprimir ---------- #
        objImprimir = Imprimir(locale_plabase)
        objImprimir.main_imprimir(driver, nome_reclamente)
        objTools.limparFilesTemp()

        # ---------- PJeCalc - Exportar ---------- #
        objExportar = Exportar(locale_plabase)
        objExportar.main_exportar(driver)

        # ---------- PJeCalc - Outras Funções ---------- #
        objDados.registrar_horario_final()
        objDados.encaminhar_log(id_processo)
        # Fechar arquivo
        # objDados.f.close()
        # Tempo de controle
        # sleep(2)
        pa.alert(title='Robô-PJeCalc', text='Automação Finalizada!')

        # Fechar o PJeCalc
        driver.quit()

        # [Limpar arquivos temporários]
        objTools.limparFilesTemp()

    except IndexError as e:
        print(f"# - Erro [Base]: {e}")
        pa.alert(title='Aviso', text=f'{e}')
        

if __name__ == '__main__':
    main()