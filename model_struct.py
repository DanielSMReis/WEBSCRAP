from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import time, os, shutil, glob 

class GETSU:
    def __init__(self, driver):
        self.driver = driver
        self.url = 'https://gets.ceb.unicamp.br/nec/view/inicio/home.jsf'
        self.bar_username = 'j_username'  # Name
        self.bar_password = 'j_password'  # Name
        self.btn_entrar = 'Entrar'  # Valor
        self.links = {
            'home': 'https://gets.ceb.unicamp.br/nec/view/inicio/home.jsf',
            'login': 'https://gets.ceb.unicamp.br/nec/view/inicio/login.jsf',
            'relatorios': 'https://gets.ceb.unicamp.br/nec/view/menus/relatorios.jsf',
            'os': 'https://gets.ceb.unicamp.br/nec/view/formrelatorios/ordens_pendentes.jsf',
            'os_encerradas':'https://gets.ceb.unicamp.br/nec/view/formrelatorios/ordens_encerradas.jsf'
        }

    def navigate(self):
        self.driver.get(self.url)  # instancia do selenium
    

    # loga com perfil de usuario tecnico
    def inicialize_tecnico(self):
        self.navigate()
        self.logar_tecnico()

    def inicialize_engenheiro(self):
        self.navigate()
        self.logar_engenheiro()

    # navega pelos principais links da pagina, salvos no atributo da classe
    def ir_para(self, acao):
        if acao in self.links:
            self.driver.get(self.links[acao])
        else:
            print("Ação inválida. Tente novamente.")


    def logar_tecnico(self, username='daniel.reis@ebserh.gov.br', password='Desktop1'):
        self.driver.find_element(By.NAME, self.bar_username).send_keys(username)
        self.driver.find_element(By.NAME, self.bar_password).send_keys(password)
        self.driver.find_element(By.XPATH, "//input[@value='Entrar']").click()


    def logar_engenheiro(self, username='victor.bomfim@ebserh.gov.br', password='251712victor'):
        self.driver.find_element(By.NAME, self.bar_username).send_keys(username)
        self.driver.find_element(By.NAME, self.bar_password).send_keys(password)
        self.driver.find_element(By.XPATH, "//input[@value='Entrar']").click()

    def gerar_planilha_abertas(self,driver, choice:str):
        ff = driver

        # Localize o botão usando XPath absoluto
        botao = ff.find_element(By.XPATH, "/html/body/div[3]/div[2]/form/div[2]/div[1]/div/table/tbody/tr[3]/td[2]/span[1]/button/span[1]")
        botao.click()

        btndia1 = ff.find_element(By.XPATH,"/html/body/div[13]/table/tbody/tr[1]/td[5]/a")
        btndia1.click()


        time.sleep(1)


        btnclasse_da_os = ff.find_element(By.XPATH, "/html/body/div[3]/div[2]/form/div[2]/div[1]/div/table/tbody/tr[6]/td[2]/div/div[3]/span")
        
        if(choice == '1'):
            btnclasse_da_os.click()
            btnclasse_corretiva = ff.find_element(By.XPATH,"/html/body/div[17]/div/ul/li[2]")
            btnclasse_corretiva.click()

        elif(choice == '2'):
            btnclasse_da_os.click()
            btnclasse_programada = ff.find_element(By.XPATH, "/html/body/div[17]/div/ul/li[3]")
            btnclasse_programada.click()

        btnplanilha = ff.find_element(By.XPATH,"/html/body/div[3]/div[2]/form/div[3]/button[2]/span[2]")
        btnplanilha.click()


    def gerar_planilha_fechadas(self,driver, choice:str):
        ff = driver
        btn_data_encerra = ff.find_element(By.XPATH,"/html/body/div[3]/div[2]/form/div[2]/div[1]/div/table/tbody/tr[2]/td[2]/span/span[1]/button" )
        btn_data_encerra.click()
        btn_1_encerra = ff.find_element(By.XPATH,"/html/body/div[11]/table/tbody/tr[1]/td[5]/a")
        btn_1_encerra.click()

        btn_data_aberta = ff.find_element(By.XPATH, "/html/body/div[3]/div[2]/form/div[2]/div[1]/div/table/tbody/tr[2]/td[1]/span/span[1]/button")
        btn_data_aberta.click()
        btn_1_aberta = ff.find_element(By.XPATH, "/html/body/div[11]/table/tbody/tr[1]/td[5]/a")
        btn_1_aberta.click()

        btn_classe_da_os = ff.find_element(By.XPATH, "/html/body/div[3]/div[2]/form/div[2]/div[1]/div/table/tbody/tr[7]/td[1]/div/div[3]/span")

        if (choice == '1'):
            btn_classe_da_os.click()
            btn_corretiva = ff.find_element(By.XPATH,"/html/body/div[16]/div/ul/li[2]")
            btn_corretiva.click()

            btn_gerar_planilha = ff.find_element(By.XPATH,"/html/body/div[3]/div[2]/form/div[3]/button[2]/span[2]")
            btn_gerar_planilha.click()
            btn_baixar = ff.find_element(By.XPATH, "/html/body/div[3]/div[2]/form/div[4]/div/div/p[1]/button[1]/span[2]")
            btn_baixar.click()

        elif(choice == '2'):
            btn_classe_da_os.click()
            btn_programadas = ff.find_element(By.XPATH, "/html/body/div[16]/div/ul/li[3]")
            btn_programadas.click()

            btn_gerar_planilha = ff.find_element(By.XPATH,"/html/body/div[3]/div[2]/form/div[3]/button[2]/span[2]")
            btn_gerar_planilha.click()
            btn_baixar = ff.find_element(By.XPATH, "/html/body/div[3]/div[2]/form/div[4]/div/div/p[1]/button[1]/span[2]")
            btn_baixar.click()
