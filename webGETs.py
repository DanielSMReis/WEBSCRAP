from model_struct import GETSU
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import time, os 
\



choice = input('selecione uma das opções:\n\n1 - download da planilha de ordens para relatorio\n2 - download da planilha de indicadores\n')
if choice == '1':
    choice2 = input('\ndigite qual planilha do mes corrente deseja criar:\n\n1 - para OSs corretiva \n2 - para OSs programadas\n3 - para todas\n')
    #criando instancia da classe GETSU
    driver = webdriver.Firefox()
    perfil_tecnico = GETSU(driver)
    perfil_tecnico.inicialize_tecnico()
    perfil_tecnico.ir_para('os')
    perfil_tecnico.gerar_planilha_abertas(driver,choice2)

elif choice == '2':
    choice3 = input('Digite:\n1 - para indicador corretiva\n2 - para indicador programadas\n')
    driver = webdriver.Firefox()
    perfil_engenheiro = GETSU(driver)
    perfil_engenheiro.inicialize_engenheiro()
   # perfil_engenheiro.ir_para('os')
   # perfil_engenheiro.gerar_planilha_abertas(driver,choice3)
    
    #salvar na pasta para manipulação
    
    time.sleep(2)
    perfil_engenheiro.ir_para('os_encerradas')
    perfil_engenheiro.gerar_planilha_fechadas(driver,choice3)