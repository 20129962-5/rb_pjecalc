from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from datetime import date
import shutil
import glob
import time
import os
from Tools.pjecalc_control import Control


class Exportar:


    def __init__(self, source):
        self.source = source
        self.delay = 10
        self.delayG = 1.5
        self.objTools = Control()


    def verificacao(self, driver):

        def gerar_relatorio(campo, status):
            file_txt_log = open(os.getcwd() + '\log.txt', "a")
            # file_txt_log.write('* ' + campo + ' | ' + status + '\n')
            file_txt_log.write(f'- {campo} | {status}\n')
            return file_txt_log.close()

        delay = 10
        try:
            mensagem = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.ID, 'formulario:painelMensagens:j_id69')))
            msg = mensagem.text
            if 'Operação realizada com sucesso.' in msg:
                print('* Operação realizada com sucesso.')
                gerar_relatorio('Exportar', 'Ok')
            else:
                print('* ERRO!', msg)
                gerar_relatorio('Exportar', '---------- Erro! ----------')
        except TimeoutException:
            print(
                '* Exceção - Verificação -  A Página demorou para responder ou o elemento não foi encontrado. Encerrando...')
            exit()

        # Tempo de controle
        time.sleep(2)

    def entrar_exportar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.CLASS_NAME, "menuImageExport"))).click()
        # WebDriverWait(driver, self.delay).until(EC.element_to_be_clickable((By.ID, 'formulario:j_id46:2:j_id49:4:j_id54'))).click()

    def click_exportar(self, driver):
        WebDriverWait(driver, self.delay).until(EC.presence_of_element_located((By.ID, 'formulario:exportar'))).click()


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


    def mover_pjc_renomeando_bkp(self, driver):

        diretorio_destino = os.path.dirname(self.source)
        # - Pegar número do processo e Cálculo
        dados = self.get_numeroProcesso(driver)
        print(f"- Dados: {dados}")
        # Pegar a data atual para identificar o arquivo mais recente
        data_atual = date.today()
        data_atual = data_atual.strftime('%d%m%Y')
        # [Nº_PROCESSO]
        numero_processo = dados[1]
        numero_processo = numero_processo.replace(".", "").replace("-", "")
        arquivo_pjc = glob.glob(os.getcwd() + f'\downloads\PROCESSO_{numero_processo}_CALCULO_{dados[0]}_DATA_{data_atual}_*.PJC')

        # Aguardar até o download do arquivo
        ct = 0
        while not len(arquivo_pjc) >= 1:
            print('...', end='')
            time.sleep(1)
            # Verificação
            arquivo_pjc = glob.glob(os.getcwd() + f'\downloads\PROCESSO_{numero_processo}_CALCULO_{dados[0]}_DATA_{data_atual}_*.PJC')
            ct += 1
            if ct == 90:
                break
        # Exceção para os casos em que o Download do arquivo não seja concluído e pasta continue vazia ou igual a zero
        try:
            if os.path.exists(arquivo_pjc[0]):
                # Tratamento para pegar somente o arquivo PJC
                nome_pjc = arquivo_pjc[0].replace("\\", " ")
                nome_pjc = nome_pjc.split()
                target = diretorio_destino + rf'\03 Automação\{nome_pjc[-1]}'
                print('- Pasta de origem: ', arquivo_pjc[0])
                print('- Pasta de destino: ', target)
                shutil.move(arquivo_pjc[0], target)
        except:

            try:
                if os.path.exists(arquivo_pjc[0]):
                    # Tratamento para pegar somente o arquivo PJC
                    nome_pjc = arquivo_pjc[0].replace("\\", " ")
                    nome_pjc = nome_pjc.split()
                    target = diretorio_destino + rf'\{nome_pjc[-1]}'
                    print('- Pasta de origem: ', arquivo_pjc[0])
                    print('- Pasta de destino: ', target)
                    shutil.move(arquivo_pjc[0], target)
            except:
                print("- [except][exportar]")

    def mover_pjc_renomeando(self, driver):

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
        # arquivo_pjc = glob.glob(os.getcwd() + f'\downloads\PROCESSO_{numero_processo}_CALCULO_{dados[0]}_DATA_{data_atual}_*.PJC')
        contagem = 90  # 01:30

        while True:
            arquivo_pjc = glob.glob(os.getcwd() + f'\downloads\PROCESSO_{numero_processo}_CALCULO_{dados[0]}_DATA_{data_atual}_*.PJC')
            # relatorio_excel = glob.glob(fr"{diretorioSource}\{nome_file_source}")
            try:
                if os.path.exists(arquivo_pjc[0]):
                    tamanho_inicial = os.path.getsize(arquivo_pjc[0])
                    time.sleep(1)  # Esperar um pouco antes de verificar novamente
                    tamanho_final = os.path.getsize(arquivo_pjc[0])
                    if tamanho_inicial == tamanho_final:  # Verifica se o tamanho do arquivo não está mais mudando
                        print("- [NOTIFICAÇÃO]: DOWNLOAD CONCLUÍDO.")
                        # ... código para mover o arquivo ...
                        print("- [NOTIFICAÇÃO]: ARQUIVO LOCALIZADO.")
                        # Tratamento para pegar somente o arquivo PJC
                        nome_pjc = arquivo_pjc[0].replace("\\", " ")
                        nome_pjc = nome_pjc.split()
                        target = diretorio_destino + rf'\03 Automação\{nome_pjc[-1]}'
                        print(f"- [ORIGEM_ARQUIVO]: {arquivo_pjc[0]}")
                        print(f"- [DESTINO_ARQUIVO]: {target}\n{'=' * 5}\n")
                        shutil.move(arquivo_pjc[0], target)
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
                    print("- [NOTIFICAÇÃO]: FALHA NO DOWNLOAD DO PJC")
                    print(f"\n{'=' * 5}\n")
                    break
                continue

            except Exception as e:
                print(f"- [except][2]: {e}")

                try:
                    if os.path.exists(arquivo_pjc[0]):
                        # Tratamento para pegar somente o arquivo PJC
                        nome_pjc = arquivo_pjc[0].replace("\\", " ")
                        nome_pjc = nome_pjc.split()
                        # target = diretorio_destino + rf'\{nome_pjc[-1]}'
                        target = fr"{diretorio_destino}\{nome_pjc[-1]}"
                        print(f"- [ORIGEM_ARQUIVO][2]: {arquivo_pjc[0]}")
                        print(f"- [DESTINO_ARQUIVO][2]: {target}\n{'=' * 5}\n")
                        shutil.move(arquivo_pjc[0], target)
                        print(f"- [STATUS_OPERACAO][2]: PJC MOVIDO COM SUCESSO")
                        time.sleep(5)
                        break
                except Exception as e:
                    print(f"- [except][3]: {e}")
                    print(f"\n- [ALERTA]: [!!] NAO_FOI_POSSIVEL_MOVER_O_PJC [!!]\n")
                    time.sleep(5)
                    break

        del dados


    def main_exportar(self, driver):

        print('\n# ========== [EXPORTAR_PJC] ========== #\n')
        self.entrar_exportar(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(1)
        self.click_exportar(driver)
        self.objTools.aguardar_carregamento(driver)
        # Tempo de controle
        time.sleep(1)
        self.mover_pjc_renomeando(driver)
        print('\n# ===================================== #\n')