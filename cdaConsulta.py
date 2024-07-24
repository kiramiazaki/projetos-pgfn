# pip install -r requirements.txt
from sys import stdout
from pandas import read_excel, DataFrame
from datetime import date
from pathlib import Path
from string import punctuation
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
options = webdriver.ChromeOptions()
options.add_argument('--ignore-ssl-errors=yes')
options.add_argument('--ignore-certificate-errors')
driver = webdriver.Chrome(options=options)

def lerExcel():
    excel_data = read_excel(fr'{Path.cwd()}\arquivos_gerados\OUTPUT - Execução - {data_mod}.xlsx', sheet_name=0, header=0, usecols=['Processo']).dropna()
    list = excel_data['Processo'].str.translate(str.maketrans('', '', '[^\w\s]'+punctuation)).tolist()
    excel_data = []
    for l in list:
        excel_data.append([l])

    print('{} números de processos lidos.\n'.format(len(excel_data)))
    return excel_data


def checarExistencia(xpath):
    try:
        driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True


def escolherProcesso():
    for cont, mCab in enumerate(driver.find_element(By.XPATH, '//*[@id="consultarProcessoForm:dtProcessos"]/table/thead/tr').find_elements(By.TAG_NAME, 'th')):
        if mCab.text == 'CLASSE':
            for mValues in driver.find_element(By.XPATH, '//*[@id="consultarProcessoForm:dtProcessos_data"]').find_elements(By.TAG_NAME, 'tr'):
                if mValues.find_elements(By.TAG_NAME, 'td')[cont].text == 'Execução Fiscal (SIDA)':
                    mValues.find_element(By.TAG_NAME, 'a').click()
                    return True
    return False


def acessarSaj():
    driver.get('https://saj.pgfn.fazenda.gov.br/saj/login.jsf')
    # Esperar fazer o login
    WebDriverWait(driver,300).until(EC.presence_of_element_located((By.ID, 'j_idt15:formMenus:menuPerfilCadastroProcessos')))
    return


def consultar(processo):
    try:
        driver.get('https://saj.pgfn.fazenda.gov.br/saj/pages/consultarProcessos/consultarProcesso.jsf')
        # Form número processo
        WebDriverWait(driver,20).until(EC.presence_of_element_located((By.ID, 'consultarProcessoForm:numeroProcesso')))
        driver.find_element(By.ID, 'consultarProcessoForm:numeroProcesso').clear()
        driver.find_element(By.ID, 'consultarProcessoForm:numeroProcesso').click()
        driver.find_element(By.ID, 'consultarProcessoForm:numeroProcesso').send_keys(processo)
        # Botão Pesquisar
        driver.find_element(By.ID, 'consultarProcessoForm:consultarProcessos').click()

        # Checar se existe mais de um processo
        WebDriverWait(driver,20).until(EC.presence_of_element_located((By.ID, 'panelGroupConteudo')))
        multi = checarExistencia('//*[@id="consultarProcessoForm:resultadoPanel"]')
        if multi == True:
            result = escolherProcesso()
            if result == False:
                raise Exception

        # Campo Classe: Execução Fiscal (SIDA)
        WebDriverWait(driver,20).until(EC.presence_of_element_located((By.ID, 'frmDetalhar:j_idt104:0:pgDadosBasicos')))
        classe = False
        for cGeral in driver.find_elements(By.XPATH, '//*[@id="frmDetalhar:j_idt104:0:pgDadosBasicos"]/tbody/tr'):
            if cGeral.text == 'Classe: Execução Fiscal (SIDA)':
                classe = True
                break

        if classe:
            WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmDetalhar:j_idt104:inscricaoSidaTable"]/table/thead/tr')))
            for c, aNome in enumerate(driver.find_element(By.XPATH, '//*[@id="frmDetalhar:j_idt104:inscricaoSidaTable"]/table/thead/tr').find_elements(By.TAG_NAME, 'th')):
                if aNome.text == 'APA':
                    # Campo num APA
                    aNum = driver.find_element(By.XPATH, '//*[@id="frmDetalhar:j_idt104:inscricaoSidaTable_data"]/tr').find_elements(By.TAG_NAME, 'td')[c].text
                    break
        else:
            raise Exception
        
    except:
        aNum = ''
    return aNum


def main():
    excel_data = lerExcel()
    acessarSaj()
    for i in range(len(excel_data)):
        print(f'{i+1}- Processo: {excel_data[i][0]}', end=' ')
        dados = consultar(excel_data[i][0])
        if dados != '':
            excel_data[i].append(dados)
            print(f'- Agrupamento: {excel_data[i][1]}')
        else:
            print('- Não é Execução Fiscal (SIDA)')
    DataFrame(excel_data, columns=['Processo', 'Agrupamento']).to_excel(fr'{Path.cwd()}\arquivos_gerados\AGRUPAMENTO - {data_mod}.xlsx', index=False, header=True)
    print('Planilha AGRUPAMENTO criada.\n')
    driver.quit()
    return


main()
