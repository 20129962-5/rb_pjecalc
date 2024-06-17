import os
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.common.exceptions import UnexpectedAlertPresentException



class WebDriver:

    def driver_web(self):
        try:
            profile = Options()
            # profile.add_argument('--headless')  # Ocultar instância webdriver
            profile.set_preference("browser.download.folderList", 2)  # 0 = Área de Trabalho | 1 = Local de Download padrão | 2 = Pasta personazada
            profile.set_preference("browser.download.manager.showWhenStarting", False)
            # profile.set_preference("browser.download.dir", "C:\\Users\lucas.fonseca\downloads\\")
            profile.set_preference("browser.download.dir", os.getcwd()+'\downloads')
            # profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf;" + "application/zip;charset=iso-8859-1" + "application/x-www-form-urlencoded")
            profile.set_preference("browser.helperApps.neverAsk.saveToDisk",
                                   "application/pdf" + "application/x-www-form-urlencoded" + "application/zip;charset=iso-8859-1")
            profile.set_preference("pdfjs.disabled", True)
            # service = Service('.\geckodriver.exe')
            # dir_path = os.path.dirname(os.path.realpath(__file__)) + '\geckodriver.exe'
            # dir_source = Service(executable_path=dir_path)
            driver = webdriver.Firefox(options=profile)
            driver.maximize_window()
            return driver
        except UnexpectedAlertPresentException:
            pass