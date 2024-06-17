import os
import gc
import time
import shutil
from selenium.webdriver.common.by import By
from selenium.common import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



class Control:


    def recarregarPagina(self, driver):
        driver.refresh()



    def verificar_erro_paginaPJeCalc(self, driver):

        try:
            field_text = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, '//div[@class="boxErro"]'))).get_attribute('textContent')
            print(f"- [STATUS_PAGINA]: {field_text}")
            if 'erro' in field_text.lower():
                pass
            else:
                return [True, '']
        except Exception as e:
            print(f"- [except][verificar_erro_paginaPJeCalc]: {e}")


    def limparFilesTemp_v2(self):

        dirs_temp = [
            "C:\\Windows\\Temp",
            os.path.expanduser("~\\AppData\\Local\\Temp"),
            os.path.expanduser("~\\Recent"),
        ]

        for dir_temp in dirs_temp:
            try:
                for item in os.listdir(dir_temp):
                    path = os.path.join(dir_temp, item)

                    if 'base_' in item:
                        print(f"- [!!] [{item}] [!!]")
                    try:
                        if os.path.isfile(path) or os.path.islink(path):
                            os.remove(path)
                            print(f"- [!!] [{item}] - [DELETADO] [!!]")
                        elif os.path.isdir(path):
                            shutil.rmtree(path)
                    except Exception:
                        continue
            except Exception:
                continue

        gc.collect(generation=0)
        gc.collect(generation=1)
        gc.collect(generation=2)


    def limparFilesTemp(self):

        dirSystemTemp = "C:\Windows\Temp"
        dirUserTemp = os.path.expanduser("~\AppData\Local\Temp")
        dirUserRecent = os.path.expanduser("~\Recent")

        try:
            for f in os.listdir(dirSystemTemp):
                try:
                    os.remove(os.path.join(dirSystemTemp, f))
                except PermissionError:
                    continue
        except:
            pass
            # print("- [Except dirSystemTemp]")

        for g in os.listdir(dirUserTemp):

            try:
                os.remove(os.path.join(dirUserTemp, g))
            except PermissionError:
                continue

        try:
            for h in os.listdir(dirUserRecent):
                try:
                    os.remove(os.path.join(dirUserTemp, h))
                except PermissionError:
                    continue
        except:
            pass
            # print("- [Except Folder Recent]")

        gc.collect(generation=0)
        gc.collect(generation=1)
        gc.collect(generation=2)


    def aguardar_carregamento(self, driver):
        try:
            while WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID, 'formulario:msgAguardeContentTable'))):
                # print("...", end="")
                time.sleep(1)
        except TimeoutException:
            time.sleep(1.5)

    def remover_filesDiretorioProcesso(self, diretorio):

        if diretorio:
            for file in diretorio:
                # print(f"- [Delete]: {file}")
                if '03 Automação' in file:
                    try:
                        os.remove(file)
                    except FileNotFoundError as e:
                        print(f"- [except][delete]: {e}")
                else:
                    continue

    # - Implement
    def enviar_relatorio_email(self, id_processo):

        texto = f"Não foi possível encontrar os arquivos (Planinhas Base/Verbas) no diretório 03 Automação, referente ao processo: {id_processo}\n\nO responsável pode ter esquecido de copiar os arquivos para a pasta.\n\nATT.: Rô-berto"

        contas = ["marcos.santos@jrspericia.com.br", "wylber.andrade@jrspericia.com.br", "lucas.fonseca@jrspericia.com.br"]

        # Assunto
        assunto = "Automação - PJeCalc"

        # Percorrer a lista de emails de destino e enviar o relatório do Push
        for email in contas:
            msg = MIMEMultipart()
            mensagem = texto
            password = 'Jrs-2018'
            msg['Subject'] = assunto
            msg['From'] = "ro-berto@jrspericia.com.br"
            msg['To'] = email
            msg.attach(MIMEText(mensagem, 'plain'))
            # create server
            server = smtplib.SMTP('mail.jrspericia.com.br', 587)
            server.starttls()
            # Login Credentials for sending the mail
            server.login(msg['From'], password)
            # send the message via the server.
            server.sendmail(msg['From'], msg['To'], msg.as_string())
            server.quit()
            print("E-mail enviado com sucesso para %s:" % (msg['To']))
            sleep(0.5)


if __name__ == '__main__':

    objTools = Control()
    diretorio = [r"J:\01 Perícias\01 Análise Pericial\9652_0000518-68.2022.5.06.0311 PJe\01 Laudo Pericial\03 Automação\9652 - Planilha Base v3.37.4.xlsb", r"J:\01 Perícias\01 Análise Pericial\9652_0000518-68.2022.5.06.0311 PJe\01 Laudo Pericial\03 Automação\9652 - Verbas v1.08.xlsm"]
    objTools.remover_filesDiretorioProcesso(diretorio)