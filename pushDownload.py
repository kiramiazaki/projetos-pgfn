from sys import stdout
from pandas import read_excel, DataFrame
from os import path, mkdir
from datetime import date
from time import sleep
from pathlib import Path
from glob import glob
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import chromedriver_autoinstaller

stdout.reconfigure(encoding='utf-8')
data = date.today().strftime('%#d-%#m-%Y')
chromedriver_autoinstaller.install()
options = webdriver.ChromeOptions()
options.add_argument('start-maximized')
directory = fr'{Path.cwd()}\arquivos_pdf\{data}'
prefs = {'download.default_directory' : directory,
         'download.prompt_for_download': False,
         'download.directory_upgrade': True,
         'plugins.always_open_pdf_externally': True,
         'plugins.plugins_list': [{'enabled': False,
                                         'name': 'Chrome PDF Viewer'}],
        'download.extensions_to_open': '',}
options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(options=options)

def lerExcel():
    excel_data = read_excel(fr'{Path.cwd()}\arquivos_gerados\Processos PUSH - {data}.xlsx', sheet_name=0, header=0, usecols=['Processo'], dtype=str).dropna()
    excel_data_nome = excel_data['Processo'].tolist()

    excel_data = []
    for i in range(len(excel_data_nome)):
        excel_data.append([excel_data_nome[i]])
    print('{} processos lidos.\n'.format(len(excel_data)))
    return excel_data

def acessarPje():
    driver.get('https://pje1g.trf3.jus.br/')
    # Esperar fazer o login
    WebDriverWait(driver,300).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="j_id179"]/input[1]')))
    return


def esperarDownload(bef, processo):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        sleep(1)
        dl_wait = False
        if bef == glob(fr'{directory}\{processo}*'):
            dl_wait = True
        seconds += 1
    if dl_wait:
        return True
    else:
        return False


def baixarSisbajud(processo):
    try:
        initial_window_handle = driver.current_window_handle
        driver.get('https://pje1g.trf3.jus.br/pje/Processo/ConsultaProcesso/listView.seam')
        WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, 'fPP:numeroProcesso:numeroSequencial')))
        driver.find_element(By.ID, 'fPP:numeroProcesso:numeroSequencial').send_keys(processo.replace('.4.03.', '.'))
        WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.ID, 'fPP:searchProcessos')))
        driver.find_element(By.ID, 'fPP:searchProcessos').click()
        WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH, "//a[@class='btn-link btn-condensed']")))
        driver.find_element(By.XPATH, "//a[@class='btn-link btn-condensed']").click()
        WebDriverWait(driver,20).until(EC.new_window_is_opened(set(driver.window_handles)))
        driver.switch_to.window(driver.window_handles[-1])
        WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, 'divTimeLine:txtPesquisa')))
        driver.find_element(By.ID, 'divTimeLine:txtPesquisa').send_keys('sisbajud')

        bef = driver.find_element(By.ID, 'divTimeLine:eventosTimeLineElement').find_elements(By.CSS_SELECTOR, 'div.media.interno.tipo-D')
        ant = glob(fr'{directory}\{processo}*')
        WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.ID, 'divTimeLine:btnPesquisar')))
        driver.find_element(By.ID, 'divTimeLine:btnPesquisar').click()
        sleep(1)

        WebDriverWait(driver,20).until(EC.presence_of_element_located((By.ID, 'divTimeLine:eventosTimeLineElement')))
        now = driver.find_element(By.ID, 'divTimeLine:eventosTimeLineElement').find_elements(By.CSS_SELECTOR, 'div.media.interno.tipo-D')
        if len(bef) != len(now) and len(now) != 0:
            sMov1 = driver.find_element(By.ID, 'divTimeLine:eventosTimeLineElement').find_elements(By.CSS_SELECTOR, 'div.media.interno.tipo-D')[0]
            try:
                sMov1.find_elements(By.CSS_SELECTOR, 'div.media-body.box > div.anexos > ul > li')[-1]
            except IndexError:
                sMov1.find_element(By.CSS_SELECTOR, 'div.media-body.box > div.anexos > a').click()
            else:
                sMov2 = sMov1.find_elements(By.CSS_SELECTOR, 'div.media-body.box > div.anexos > ul > li')[-1]
                sMov2.find_element(By.CSS_SELECTOR, 'a').click()
            sleep(1)
            WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.ID, 'detalheDocumento:download')))
            driver.find_element(By.ID, 'detalheDocumento:download').click()
            driver.switch_to.alert.accept()

            err = esperarDownload(ant, processo)
            if err:
                WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="panelAlertContentDiv"]//a[@class="hidelink btn-fechar"]')))
                driver.find_element(By.XPATH, '//*[@id="panelAlertContentDiv"]//a[@class="hidelink btn-fechar"]').click()
                WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameBinario')))
                driver.find_element(By.ID, 'open-button').click()
                driver.switch_to.default_content()
                esperarDownload(ant, processo.split('.4.03')[0])
            rel = 'Download Sisbajud'
            print(' - Download Sisbajud')
        else:
            WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, 'divTimeLine:txtPesquisa')))
            driver.find_element(By.ID, 'divTimeLine:txtPesquisa').clear()
            driver.find_element(By.ID, 'divTimeLine:txtPesquisa').send_keys('bacenjud')
            WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.ID, 'divTimeLine:btnPesquisar')))
            driver.find_element(By.ID, 'divTimeLine:btnPesquisar').click()
            sleep(1)
            
            WebDriverWait(driver,20).until(EC.presence_of_element_located((By.ID, 'divTimeLine:eventosTimeLineElement')))
            now = driver.find_element(By.ID, 'divTimeLine:eventosTimeLineElement').find_elements(By.CSS_SELECTOR, 'div.media.interno.tipo-D')
            if len(bef) != len(now) and len(now) != 0:
                bMov1 = driver.find_element(By.ID, 'divTimeLine:eventosTimeLineElement').find_elements(By.CSS_SELECTOR, 'div.media.interno.tipo-D')[0]
                try:
                    bMov1.find_elements(By.CSS_SELECTOR, 'div.media-body.box > div.anexos > ul > li')[-1]
                except IndexError:
                    bMov1.find_element(By.CSS_SELECTOR, 'div.media-body.box > div.anexos > a').click()
                else:
                    bMov2 = bMov1.find_elements(By.CSS_SELECTOR, 'div.media-body.box > div.anexos > ul > li')[-1]
                    bMov2.find_element(By.CSS_SELECTOR, 'a').click()
                sleep(1)
                WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.ID, 'detalheDocumento:download')))
                driver.find_element(By.ID, 'detalheDocumento:download').click()
                driver.switch_to.alert.accept()

                err = esperarDownload(ant, processo)
                if err:
                    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="panelAlertContentDiv"]//a[@class="hidelink btn-fechar"]')))
                    driver.find_element(By.XPATH, '//*[@id="panelAlertContentDiv"]//a[@class="hidelink btn-fechar"]').click()
                    WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameBinario')))
                    driver.find_element(By.ID, 'open-button').click()
                    driver.switch_to.default_content()
                    esperarDownload(ant, processo.split('.4.03')[0])
                rel = 'Download Bacenjud'
                print(' - Download Bacenjud')
            else:
                WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, 'divTimeLine:txtPesquisa')))
                driver.find_element(By.ID, 'divTimeLine:txtPesquisa').clear()
                driver.find_element(By.ID, 'divTimeLine:txtPesquisa').send_keys('certidão')
                WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.ID, 'divTimeLine:btnPesquisar')))
                driver.find_element(By.ID, 'divTimeLine:btnPesquisar').click()
                sleep(1)

                WebDriverWait(driver,20).until(EC.presence_of_element_located((By.ID, 'divTimeLine:eventosTimeLineElement')))
                now = driver.find_element(By.ID, 'divTimeLine:eventosTimeLineElement').find_elements(By.CSS_SELECTOR, 'div.media.interno.tipo-D')
                if len(bef) != len(now) and len(now) != 0:
                    i = 0
                    tam = driver.find_element(By.ID, 'divTimeLine:eventosTimeLineElement').find_elements(By.CSS_SELECTOR, 'div.media.interno.tipo-D')
                    while True:
                        try:
                            bMov1 = driver.find_element(By.ID, 'divTimeLine:eventosTimeLineElement').find_elements(By.CSS_SELECTOR, 'div.media.interno.tipo-D')[i]
                            bMov1.find_elements(By.CSS_SELECTOR, 'div.media-body.box > div.anexos > ul > li')[-1]
                        except IndexError:
                            if i == len(tam):
                                rel = 'Não Encontrado'
                                print(' - Não Encontrado')
                                break
                            i += 1
                        else:
                            bMov2 = bMov1.find_elements(By.CSS_SELECTOR, 'div.media-body.box > div.anexos > ul > li')[-1]
                            texto = bMov2.find_element(By.CSS_SELECTOR, 'a').text
                            if 'OUTROS DOCUMENTOS' in texto.upper() or 'DETALHAMENTO DA ORDEM JUDICIAL DE BLOQUEIO DE VALORES' in texto.upper():
                                bMov2.find_element(By.CSS_SELECTOR, 'a').click()
                                sleep(1)
                                WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.ID, 'detalheDocumento:download')))
                                driver.find_element(By.ID, 'detalheDocumento:download').click()
                                driver.switch_to.alert.accept()

                                err = esperarDownload(ant, processo)
                                if err:
                                    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="panelAlertContentDiv"]//a[@class="hidelink btn-fechar"]')))
                                    driver.find_element(By.XPATH, '//*[@id="panelAlertContentDiv"]//a[@class="hidelink btn-fechar"]').click()
                                    WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameBinario')))
                                    driver.find_element(By.ID, 'open-button').click()
                                    driver.switch_to.default_content()
                                    esperarDownload(ant, processo.split('.4.03')[0])
                                rel = 'Download Certidão'
                                print(' - Download Certidão')
                                break
                            else:
                                i += 1
                else:
                    rel = 'Não Encontrado'
                    print(' - Não Encontrado')
        driver.close()
        driver.switch_to.window(initial_window_handle)
    except:
        rel = 'Erro'
        print(' - Erro')
    return rel


def main():
    excel_data = lerExcel()
    if not path.isdir(fr'{Path.cwd()}\arquivos_pdf\{data}'):
        mkdir(fr'{Path.cwd()}\arquivos_pdf\{data}')
    acessarPje()
    for i in range(len(excel_data)):
        print(f'{i+1}- Processo: {excel_data[i][0]}', end='')
        test = baixarSisbajud(excel_data[i][0])
        excel_data[i].append(test)
    DataFrame(excel_data, columns=['Processo', 'Situação']).to_excel(fr'{Path.cwd()}\arquivos_gerados\Relatório PUSH - {data}.xlsx', index=False, header=True)
    print(fr'Planilha Relatório PUSH criada.')
    driver.quit()
    return


main()
