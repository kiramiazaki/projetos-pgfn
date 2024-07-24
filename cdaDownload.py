# pip install -r requirements.txt
from sys import stdout
from pandas import read_excel, DataFrame
from os import path, mkdir
from shutil import move
from datetime import date
from time import sleep
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import chromedriver_autoinstaller

stdout.reconfigure(encoding='utf-8')
data = date.today().strftime('%d%m%Y')
data_mod = date.today().strftime('%#d-%#m-%Y')
chromedriver_autoinstaller.install()
driver = webdriver.Chrome()

def lerExcel():
    excel_data = read_excel(fr'{Path.cwd()}\arquivos_gerados\AGRUPAMENTO - {data_mod}.xlsx', sheet_name=0, header=0, usecols=['Processo', 'Agrupamento'], dtype=str).dropna()
    excel_data_nome = excel_data['Processo'].tolist()
    excel_data_val = excel_data['Agrupamento'].tolist()

    excel_data = []
    for i in range(len(excel_data_nome)):
        excel_data.append([excel_data_nome[i], excel_data_val[i]])

    print('{} números de agrupamento lidos.\n'.format(len(excel_data)))
    return excel_data


def checarExistencia(xpath):
    try:
        driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True


def acessarSida():
    driver.get('https://sida.pgfn.fazenda/sida/#/sida/login')
    # Esperar fazer o login
    WebDriverWait(driver,300).until(EC.element_to_be_clickable((By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/fieldset/div/form/fieldset[1]/div[8]/div[1]')))
    return


def baixarPdf(apa):
    try:
        err = False
        driver.get('https://sida.pgfn.fazenda/sida/#/sida/consulta/busca')
        # Botão: Número de Agrupamento para Ajuizamento
        WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/fieldset/div/form/fieldset[1]/div[8]/div[1]')))
        driver.find_element(By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/fieldset/div/form/fieldset[1]/div[8]/div[1]').click()
        # Caixa de Texto: Número de Agrupamento para Ajuizamento
        WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/fieldset/div/form/fieldset[1]/div[8]/div[2]/input')))
        driver.find_element(By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/fieldset/div/form/fieldset[1]/div[8]/div[2]/input').send_keys(apa)
        # Botão: Buscar
        driver.find_element(By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/fieldset/div/form/div/button').click()

        WebDriverWait(driver,20).until(EC.presence_of_element_located((By.CLASS_NAME, 'sidaContainerContent')))
        multi = checarExistencia('/html/body/ng-include/article/div/div/ui-view/ui-view/section/div/div/div/div/div[3]/div/div/div/div/div[3]/div[1]/button[2]')
        if multi == True:
            # Botão: Selecionar tudo
            WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/div/div/div/div/div[3]/div/div/div/div/div[3]/div[1]/button[2]')))
            driver.find_element(By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/div/div/div/div/div[3]/div/div/div/div/div[3]/div[1]/button[2]').click()
            # Botão: Imprimir inscrições
            driver.find_element(By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/div/div/div/div/div[3]/div/div/div/div/div[3]/div[1]/button[3]').click()
            # Radio: Resumido
            WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/div/div/div/div[2]/relatorio-modal/sida-modal/div/div/div[2]/ng-transclude/div[2]/div[2]/input')))
            driver.find_element(By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/div/div/div/div[2]/relatorio-modal/sida-modal/div/div/div[2]/ng-transclude/div[2]/div[2]/input').click()

            # Div de Checkbox: Parâmetros
            WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.ID, 'accordion')))
            # Botão: OK
            driver.find_element(By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/section/div/div/div/div[2]/relatorio-modal/sida-modal/div/div/div[2]/ng-transclude/div[3]/button[2]').click()
        else:
            # Botão: Imprimir
            WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/div/div/div/div/div[2]/div/button[1]')))
            driver.find_element(By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/div/div/div/div/div[2]/div/button[1]').click()
            # Radio: Resumido
            WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/div/div/div/div/div[3]/relatorio-modal/sida-modal/div/div/div[2]/ng-transclude/div[2]/div[2]/input')))
            driver.find_element(By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/div/div/div/div/div[3]/relatorio-modal/sida-modal/div/div/div[2]/ng-transclude/div[2]/div[2]/input').click()

            # Div de Checkbox: Parâmetros
            WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.ID, 'accordion')))
            # Botão: OK
            driver.find_element(By.XPATH, '/html/body/ng-include/article/div/div/ui-view/ui-view/div/div/div/div/div[3]/relatorio-modal/sida-modal/div/div/div[2]/ng-transclude/div[3]/button[2]').click()
    except:
        err = True
    return err


def esperarDownload():
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        sleep(1)
        dl_wait = False
        if not path.exists(fr'{Path.home()}\Downloads\RelResumido-{data}.pdf'):
            dl_wait = True
        seconds += 1
    return seconds


def criarPasta():
    if not path.isdir(fr'{Path.cwd()}\arquivos_pdf\{data_mod}'):
        mkdir(fr'{Path.cwd()}\arquivos_pdf\{data_mod}')
    return
       

def moverPdf(nome):
    cont = 1
    while True:
        if not path.isfile(fr'{Path.cwd()}\arquivos_pdf\{data_mod}\RelResumido-{nome}.pdf'):
            move(fr'{Path.home()}\Downloads\RelResumido-{data}.pdf', fr'{Path.cwd()}\arquivos_pdf\{data_mod}\RelResumido-{nome}.pdf')
            break
        elif not path.isfile(fr'{Path.cwd()}\arquivos_pdf\{data_mod}\RelResumido-{nome} ({cont}).pdf'):
            move(fr'{Path.home()}\Downloads\RelResumido-{data}.pdf', fr'{Path.cwd()}\arquivos_pdf\{data_mod}\RelResumido-{nome} ({cont}).pdf')
            break
        cont += 1
    return


def main():
    erros = []
    excel_data = lerExcel()
    criarPasta()
    acessarSida()
    for i in range(len(excel_data)):
        print(f'{i+1}- Processo: {excel_data[i][0]} - Agrupamento: {excel_data[i][1]}', end='')
        test = baixarPdf(excel_data[i][1])
        if test == False:
            esperarDownload()
            moverPdf(excel_data[i][0])
            print()
        else:
            erros.append([excel_data[i][0], excel_data[i][1]])
            print(' - Erro')
    if len(erros) > 0:
        print(f'{len(erros)} erro(s) encontrado(s). Verifique o(s) número(s) na planilha de erros.')
        DataFrame(erros, columns=['Processo', 'Agrupamento']).to_excel(fr'{Path.cwd()}\Erros CDA - {data_mod}.xlsx', index=False, header=True)
    driver.quit()
    return


main()
