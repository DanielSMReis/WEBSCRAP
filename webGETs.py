from model_struct import GETSU,automate_planilha
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
    driver.quit()

elif choice == '2':
    choice3 = input('Digite:\n1 - para indicador corretiva\n2 - para indicador programadas\n')
    driver = webdriver.Firefox()
    perfil_engenheiro = GETSU(driver)
    perfil_engenheiro.inicialize_engenheiro()
    perfil_engenheiro.ir_para('os')
    perfil_engenheiro.gerar_planilha_abertas(driver,choice3)

    planilha_abertas = perfil_engenheiro.transfer_download(choice3,'1')
    #salvar na pasta para manipulação
    
    perfil_engenheiro.ir_para('os_encerradas')
    perfil_engenheiro.gerar_planilha_fechadas(driver,choice3)
    planilha_fechadas = perfil_engenheiro.transfer_download(choice3,'2')
    driver.quit()

if choice3 == '1':
    corretivas_fechadas = 'planilha_corretivas_fechadas.xls'
    planilha_contador = automate_planilha(corretivas_fechadas)
    celula_inicial = 'H7'
    qnt_os_CORRETIVAS_fechadas = planilha_contador.count_items_in_column(celula_inicial)


    corretivas_abertas = 'planilha_corretivas_abertas.xls'
    planilha_contador = automate_planilha(corretivas_abertas)
    qnt_os_CORRETIVAS_abertas = planilha_contador.count_items_in_column(celula_inicial)



    indicador = 100*qnt_os_CORRETIVAS_fechadas / (qnt_os_CORRETIVAS_fechadas + qnt_os_CORRETIVAS_abertas)

    print (f'O indicador atual de Corretivas é: {indicador:.2f}%')
    
elif choice3 == '2':
    programadas_fechadas = 'planilha_programadas_fechadas.xls'
    planilha_contador = automate_planilha(programadas_fechadas)
    celula_inicial = 'H7'
    qnt_os_PROGRAMADAS_fechadas = planilha_contador.count_items_in_column(celula_inicial)


    programadas_abertas = 'planilha_programadas_abertas.xls'
    planilha_contador = automate_planilha(programadas_abertas)
    qnt_os_PROGRAMADAS_abertas = planilha_contador.count_items_in_column(celula_inicial)



    indicador = 100*qnt_os_PROGRAMADAS_fechadas / (qnt_os_PROGRAMADAS_fechadas + qnt_os_PROGRAMADAS_abertas)

    print (f'O indicador atual de Programadas é: {indicador:.2f}%')


