from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from datetime import date
import shutil
import glob
import time
import os
from Tools.pjecalc_control import Control


class Imprimir:


    def __init__(self, source):
        self.source = source
        self.delay = 10
        self.delayG = 1.5
        self.objTools = Control()


    def verificacao(self, driver):

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            file_txt_log.write(f'- {campo} | {status}\n')
            return file_txt_log.close()

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                print('* Operação realizada com sucesso.')
                gerar_relatorio('Imprimir', 'Ok')
            else:
                print('* ERRO!', msg)
                gerar_relatorio('Imprimir', '---------- Erro! ----------')
        except TimeoutException:
            print('- {}* Exceção - Verificação -  A Página demorou para responder ou o elemento não foi encontrado. Encerrando...')
            exit()

        # Tempo de controle
        time.sleep(2)

    def acessar_guia_imprimir(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImagePrint"))).click()

    def selecionar_pdf(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:formatoSaida:0'))).click()


    def selecionar_todos_checkbox(self, driver):
        # Marcar todos
        checkbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.CLASS_NAME, 'css-label')))
        if not checkbox.is_selected():
            checkbox.click()
            time.sleep(1)
        # Marcar Todos
        # checkbox = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.CLASS_NAME, 'css-label')))
        # checkbox.click()


    def desmarcar_relatorios(self, driver):

        # Desmarcar Dados do Cálculo
        opcao_dados_calc = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeRelatorio:2')))
        opcao_dados_calc.click()

        # Desmarcar Faltas e Férias
        opcao_faltas_ferias = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeRelatorio:3')))
        opcao_faltas_ferias.click()

        # Desmarcar Histórico Salarial
        opcao_historico_salarial = WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:tipoDeRelatorio:6')))
        opcao_historico_salarial.click()

    def click_imprimir(self, driver):
        WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:imprimirConsolidado'))).click()


    def aguardar_carregamento(self, driver):
        try:
            while WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID, 'formulario:msgAguardeContentTable'))):
                print("...", end="")
                time.sleep(1)
        except TimeoutException as e:
            print(f"\n- [Except] - Elemento não encontrado na página. Erro: {e}")
            time.sleep(1)

    def get_numeroProcesso(self, driver):

        listaContent = []

        fields = WebDriverWait(driver, self.delay).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'labelVerbaCalculo')))

        # print(f"- Qtd. Elementos encontrados: {len(fields)}")

        calculo = fields[0].text.split(" ")[1]
        processo = fields[1].text.split(" ")[1]

        listaContent.append(calculo)
        listaContent.append(processo)

        # print(listaContent)
        return listaContent

    def mover_relarorio_renomeando_old(self, driver, nome_reclamente, destino, numero_processo):

        diretorio_destino = os.path.dirname(destino)

        # - Pegar número do processo e Cálculo
        dados = self.get_numeroProcesso(driver)
        # print(f"- Dados: {dados}")

        # Pegar a data atual para identificar o arquivo mais recente
        data_atual = date.today()
        data_atual = data_atual.strftime('%d%m%Y')

        numero_processo = numero_processo.replace(".", "").replace("-", "")

        """
        PROCESSO_*_CALCULO_*_DATA_{data_atual}_*.PJC'
        """

        # Pasta local - Onde o relatório baixado será direcionado
        arquivo_pdf = glob.glob(os.getcwd() + f"\downloads\RELATORIO_PROCESSO_{numero_processo}_CALCULO_{dados[0]}_DATA_{data_atual}_*.pdf")
        # Aguardar até o download do arquivo
        ct = 0
        while not len(arquivo_pdf) >= 1:
            print('...', end='')
            time.sleep(1)
            # Verificação
            arquivo_pdf = glob.glob(os.getcwd() + f"\downloads\RELATORIO_PROCESSO_{numero_processo}_CALCULO_{dados[0]}_DATA_{data_atual}_*.pdf")
            ct += 1
            if ct == 90:
                break
        print()
        # Exceção para os casos em que o Download do arquivo não seja concluído e pasta continue vazia ou igual a zero
        try:
            if os.path.exists(arquivo_pdf[0]):
                target = diretorio_destino + r'\03 Automação\Relatório de Cálculo - ' + nome_reclamente + '.pdf'
                print('- Pasta de origem: ', arquivo_pdf[0])
                print('- Pasta de destino: ', target)
                shutil.move(arquivo_pdf[0], target)
        except:
            try:
                if os.path.exists(arquivo_pdf[0]):
                    target = diretorio_destino + '\Relatório de Cálculo - ' + nome_reclamente + '.pdf'
                    print('- Pasta de origem: ', arquivo_pdf[0])
                    print('- Pasta de destino: ', target)
                    shutil.move(arquivo_pdf[0], target)
            except:
                print("- [except][imprimir]")


    def mover_relatorio_renomeando(self, driver, nome_reclamante):

        diretorio_destino = os.path.dirname(self.source)
        # - Pegar número do processo e Cálculo
        dados = self.get_numeroProcesso(driver)
        print(f"\n- Dados: {dados}")
        # Pegar a data atual para identificar o arquivo mais recente
        data_atual = date.today()
        data_atual = data_atual.strftime('%d%m%Y')
        # [Nº_PROCESSO]
        numero_processo = dados[1]
        numero_processo = numero_processo.replace(".", "").replace("-", "")
        contagem = 90  # 01:30

        while True:
            arquivo_pdf = glob.glob(os.getcwd() + f"\downloads\RELATORIO_PROCESSO_{numero_processo}_CALCULO_{dados[0]}_DATA_{data_atual}_*.pdf")
            try:
                if os.path.exists(arquivo_pdf[0]):
                    tamanho_inicial = os.path.getsize(arquivo_pdf[0])
                    time.sleep(1)  # Esperar um pouco antes de verificar novamente
                    tamanho_final = os.path.getsize(arquivo_pdf[0])
                    if tamanho_inicial == tamanho_final:  # Verifica se o tamanho do arquivo não está mais mudando
                        print("- [NOTIFICAÇÃO]: DOWNLOAD CONCLUÍDO.")
                        target = fr"{diretorio_destino}\03 Automação\Relatório de Cálculo - {nome_reclamante}.pdf"
                        print(f"- [ORIGEM_ARQUIVO]: {arquivo_pdf[0]}")
                        print(f"- [DESTINO_ARQUIVO]: {target}\n{'=' * 10}\n")
                        shutil.move(arquivo_pdf[0], target)
                        print(f"- [STATUS_OPERACAO][1]: PJC MOVIDO COM SUCESSO")
                        time.sleep(5)
                        break
                    else:
                        print("- [NOTIFICAÇÃO]: DOWNLOAD EM ANDAMENTO.")
                        continue
            except IndexError as e:
                print(f"- [except][1]: {e}")
                time.sleep(1)
                contagem -= 1
                if contagem == 0:
                    print("- [NOTIFICAÇÃO]: FALHA NO DOWNLOAD DO PDF")
                    print(f"\n{'=' * 10}\n")
                    break
                continue

            except Exception as e:
                print(f"- [except][2]: {e}")

                try:
                    if os.path.exists(arquivo_pdf[0]):
                        target = fr"{diretorio_destino}\Relatório de Cálculo - {nome_reclamante}.pdf"
                        print(f"- [ORIGEM_ARQUIVO][2]: {arquivo_pdf[0]}")
                        print(f"- [DESTINO_ARQUIVO][2]: {target}\n{'=' * 10}\n")
                        shutil.move(arquivo_pdf[0], target)
                        print(f"- [STATUS_OPERACAO][2]: PDF MOVIDO COM SUCESSO")
                        time.sleep(5)
                        break
                except Exception as e:
                    print(f"- [except][3]: {e}")
                    print(f"\n- [ALERTA]: [!!] NAO_FOI_POSSIVEL_MOVER_O_PDF [!!]\n")
                    time.sleep(5)
                    break

        del dados


    def main_imprimir(self, driver, nome_reclamente):

        self.acessar_guia_imprimir(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(self.delayG)
        self.selecionar_pdf(driver)
        self.click_imprimir(driver)
        self.objTools.aguardar_carregamento(driver)
        time.sleep(self.delayG)
        self.mover_relatorio_renomeando(driver, nome_reclamente)
        time.sleep(self.delayG)
        print('- Fim - (Imprimir) --')