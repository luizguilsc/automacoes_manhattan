from selenium import webdriver
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from time import sleep

class AutoClick:
    def __init__(self, driver):
        self.driver = driver

    def _encontrar_elemento(self, xpath: str, timeout: int, condition):
        return WebDriverWait(self.driver, timeout).until(condition((By.XPATH, xpath)))

    def click_elemento(self, xpath: str, timeout: int = 30):
        elemento = self._encontrar_elemento(xpath, timeout, EC.element_to_be_clickable)
        elemento.click()

    def enviar_keys(self, xpath: str, texto: str, timeout: int = 30):
        elemento = self._encontrar_elemento(xpath, timeout, EC.element_to_be_clickable)
        elemento.send_keys(texto)
    
    def pressionar_enter(self, xpath: str, timeout: int = 30):
        elemento = self._encontrar_elemento(xpath, timeout, EC.element_to_be_clickable)
        elemento.send_keys(Keys.RETURN)
    
    def selecionar(self, xpath: str, timeout: int = 30):
        elemento = WebDriverWait(self.driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))
        print("Elemento encontrado:", elemento.text)
        return elemento.text

    

class Automation():
    def __init__(self, wait: int = 30):
        self.wait = wait
        self.driver = None

    def setup_driver(self, link: str = "https://viavp-auth.sce.manh.com/discover_user"):
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.get(link)
        AutoClick(self.driver).click_elemento("//input[@id='discover-user-submit']", timeout=self.wait)

    def selecionar_tela(self):
        try:
            auto_click = AutoClick(self.driver)
            aguardando = Automation(self.driver)
            menu = "//ion-menu-toggle"
            barra = "/html/body/app-root/ion-app/div/ion-split-pane/ion-menu/ui-dm-hamburger/ion-header/ion-input/input"
            inventory_details_menu = "//button[@id='inventoryDetails']"
            aguardando.popup_please_wait()
            auto_click.click_elemento(menu, self.wait)
            auto_click.click_elemento(barra, self.wait)
            auto_click.enviar_keys(barra,"inventory details", self.wait)
            sleep(1)
            auto_click.click_elemento(inventory_details_menu, self.wait)
            aguardando.popup_please_wait()

        except Exception as e:
            print(f"Verificar o erro {e}")

    def popup_please_wait(self):
        if self.driver:
            try:
                please_wait = "//*[@id='loading-2-msg']"
                WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH,please_wait)))
                print("Elemento foi encontrado, passou pelo código e esperou até ficar invisivel")
            except Exception as e:
                print(f"Elemento não foi encontrado ou não ficou visivel{e}")    

    def reason_code_auto(self):
        #Label da ILPN
        inventory_container_path = "/html/body/app-root/ion-app/div/ion-split-pane/ion-router-outlet/inventory-grid/dm-list-layout/div/div/div[2]/dm-filter/div[2]/div/div[2]/div[2]/text-field-filter/div/ion-row/div/div/ion-input/input"
        checkbox = "//*[@id='main']/inventory-grid/dm-list-layout/div/div/div[3]/div[2]/div[1]/ion-content/grid-view/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[1]/datatable-body-cell/div/label/input"
        more_options = "//*[@id='main']/inventory-grid/dm-list-layout/div/div/div[3]/div[2]/footer-actions/ion-grid/ion-row/ion-col[3]/div/div/more-actions/ion-button"
        reidentify_item = "//*[@id='mat-menu-panel-2']/div/div[2]/button"
        source_da_ilpn = "//*[@id='ion-overlay-12']/transfer-popup/ion-content/ion-row[2]/div/div/ion-row[2]/ion-col[1]/div/ion-row/ion-col/div/ion-row"
        copiar_sku = copiar_sku = "//span[@data-component-id]"

        campo_sku_rdyi = "//*[@id='ion-overlay-6']/transfer-popup/ion-content/ion-row[2]/div/div/ion-row[2]/ion-col[2]/div/ion-row/ion-col/div/form/ion-row/ion-col/div/ion-row[1]/ion-input"
        popup_reidentify = "//*[@id='ion-overlay-12']//ion-backdrop"
        auto_click = AutoClick(self.driver)
        aguardando = Automation(self.driver)
        #Criar Looping para ler as ILPN's
        auto_click.click_elemento(inventory_container_path, 30)
        auto_click.enviar_keys(inventory_container_path, "Q1144789",30)
        auto_click.pressionar_enter(inventory_container_path, 10)
        # campo_ilpn = self.driver.find_element(By.XPATH,inventory_container_path)
        auto_click.click_elemento(checkbox, 10)
        sleep(1)
        auto_click.click_elemento(more_options, 10)
        auto_click.click_elemento(reidentify_item, 10)
        auto_click.click_elemento(copiar_sku, 30)
        input("Verificando")
        # campo_ilpn.clear()     

    def close_driver(self):
        if self.driver:
            self.driver.quit()

class Login():
    def __init__(self, driver, login: str, senha: str):
        self.driver = driver
        self.login = login
        self.senha = senha

    # Fazer com que reconheça o login e senha e aplique no código
    def logando(self):
        
        email= "//input[@id='i0116']"
        senha= "//input[@id='i0118']"
        avancar = "//input[@id='idSIButton9']"
        home = "//button[@value='Home']"
        auto_click = AutoClick(self.driver)

        def tentar_acao(acao,descricao,tentativas=3):
            for tentativa in range(tentativas):
                try:
                    acao()
                    return True
                except Exception as e:
                    print(f"[Erro] {descricao} na tentativa {tentativa + 1}: {e}")
            return False
        try:
            if not tentar_acao(lambda: auto_click.click_elemento(email, 10), "Tentando clicar no Label do e-mail"):
                return
            if not tentar_acao(lambda:auto_click.enviar_keys(email, self.login, 10), "Tentando enviar o e-mail"):
                return
            if not tentar_acao(lambda: auto_click.click_elemento(avancar,10), "Tentando avançar"):
                return
        except Exception as e:
            print("Encerrando")
            
        
        try:
            if not tentar_acao(lambda: auto_click.click_elemento(senha, 10),"Tentando clicar no Label da senha"):
                return
            if not tentar_acao(lambda:auto_click.enviar_keys(senha, self.senha, 10),"Tentando enviar a senha"):
                return
            if not tentar_acao(lambda: auto_click.click_elemento(avancar,10), "Tentando clicar em avançar"):
                return
        except Exception as e:
            print("Encerrando")
            
        
        
        sleep(1)
        if not tentar_acao(lambda: auto_click.click_elemento(avancar, 10), "Tentando clicar em avançar"):
            return
        if not tentar_acao(lambda: auto_click.click_elemento(home, 10), "Tentando clicar no home"):
            return


if __name__ == "__main__":
    automation = Automation()
    automation.setup_driver()
    login = Login(automation.driver, "luiz.leite@viavarejo.com.br", "Python20")
    login.logando()
    automation.selecionar_tela()
    automation.reason_code_auto()
    # automation.close_driver()
    input("Pressione qualquer tecla para finalizar")