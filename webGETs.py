from model_struct import GETSU,automate_planilha,iniciar_interface
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import time, os ,sys
import tkinter as tk




def main(choice):
    choice3 = choice
    driver = webdriver.Firefox()
    perfil_engenheiro = GETSU(driver)
    perfil_engenheiro.inicialize_engenheiro()
    perfil_engenheiro.ir_para('os')
    perfil_engenheiro.gerar_planilha_abertas(driver,choice3)
    planilha_abertas = perfil_engenheiro.transfer_download(choice3,'1')
    perfil_engenheiro.gerar_planilha_abertas(driver,choice='2')
    planilha_abertas = perfil_engenheiro.transfer_download('2','1')
    #salvar na pasta para manipulação

    perfil_engenheiro.ir_para('os_encerradas')
    perfil_engenheiro.gerar_planilha_fechadas(driver,choice3)
    planilha_fechadas = perfil_engenheiro.transfer_download(choice3,'2')
    perfil_engenheiro.gerar_planilha_fechadas(driver,'2')
    planilha_fechadas = perfil_engenheiro.transfer_download('2','2')
    driver.quit()

    corretivas_fechadas = 'planilha_corretivas_fechadas.xls'
    planilha_contador = automate_planilha(corretivas_fechadas)
    celula_inicial = 'H7'
    qnt_os_CORRETIVAS_fechadas = planilha_contador.count_items_in_column(celula_inicial)


    corretivas_abertas = 'planilha_corretivas_abertas.xls'
    planilha_contador = automate_planilha(corretivas_abertas)
    qnt_os_CORRETIVAS_abertas = planilha_contador.count_items_in_column(celula_inicial)

    indicador = 100*qnt_os_CORRETIVAS_fechadas / (qnt_os_CORRETIVAS_fechadas + qnt_os_CORRETIVAS_abertas)

    print (f'\n\n\n\nO indicador atual de Corretivas é: {indicador:.2f}%')

    programadas_fechadas = 'planilha_programadas_fechadas.xls'
    planilha_contador = automate_planilha(programadas_fechadas)
    celula_inicial = 'H7'
    qnt_os_PROGRAMADAS_fechadas = planilha_contador.count_items_in_column(celula_inicial)

    programadas_abertas = 'planilha_programadas_abertas.xls'
    planilha_contador = automate_planilha(programadas_abertas)
    qnt_os_PROGRAMADAS_abertas = planilha_contador.count_items_in_column(celula_inicial)

    indicador = 100*qnt_os_PROGRAMADAS_fechadas / (qnt_os_PROGRAMADAS_fechadas + qnt_os_PROGRAMADAS_abertas)

    print (f'\nO indicador atual de Programadas é: {indicador:.2f}%')


choice3 = iniciar_interface(main)


