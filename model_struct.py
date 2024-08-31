from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import time, os, shutil, glob, xlrd

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


    def transfer_download(self,choice1, choice2):
        # Obtém o caminho da pasta de downloads do usuário
        download_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
        destination_folder = os.path.join(download_folder, 'scraping_indicadores')

        # Cria a pasta de destino se ela não existir
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)

        # Encontra todos os arquivos .xls na pasta de downloads
        list_of_files = glob.glob(os.path.join(download_folder, '*.xls'))
        
        # Verifica se há arquivos na lista
        if not list_of_files:
            print("Nenhum arquivo .xls encontrado na pasta de downloads.")
            return
        
        # Encontra o arquivo mais recente
        latest_file = max(list_of_files, key=os.path.getctime)
        
        # Captura o nome do arquivo
        file_name = os.path.basename(latest_file)
        
        # Renomeia a variável de acordo com os parâmetros de entrada
        if choice1 == '1':
            first_part = 'corretivas'
        elif choice1 == '2':
            first_part = 'programadas'
        else:
            print("Escolha inválida para o primeiro parâmetro. Use 1 para corretivas ou 2 para programadas.")
            return
        
        if choice2 == '1':
            second_part = 'abertas'
        elif choice2 == '2':
            second_part = 'fechadas'
        else:
            print("Escolha inválida para o segundo parâmetro. Use 1 para abertas ou 2 para fechadas.")
            return
        
        new_file_name = f'planilha_{first_part}_{second_part}.xls'
        
        # Define o caminho completo do novo arquivo
        new_file_path = os.path.join(destination_folder, new_file_name)
        
        # Move e renomeia o arquivo para a pasta de destino
        shutil.move(latest_file, new_file_path)
        print(f"Arquivo {file_name} movido e renomeado para {new_file_name} na pasta {destination_folder}")

        return new_file_name




class automate_planilha:
    def __init__(self, file_name):
        # Constrói o caminho completo para o arquivo na pasta "scraping_indicadores" dentro de "Downloads"
        downloads_folder = os.path.join(os.environ['USERPROFILE'], 'Downloads', 'scraping_indicadores')
        self.file_path = os.path.join(downloads_folder, file_name)
        
        self.workbook = xlrd.open_workbook(self.file_path)
        self.sheet = self.workbook.sheet_by_index(0)

    def count_items_in_column(self, start_cell):
        # Extrai a coluna e a linha inicial da célula de início
        start_col = ord(start_cell[0].upper()) - ord('A')
        start_row = int(start_cell[1:]) - 1
        
        # Inicializa o contador
        count = 0
        
        # Itera pelas células na coluna, começando da linha inicial
        for row in range(start_row, self.sheet.nrows):
            cell_value = self.sheet.cell_value(row, start_col)
            if cell_value:
                count += 1
        
        return count