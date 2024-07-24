from sys import stdout
from pandas import read_excel, DataFrame
from datetime import date
from time import sleep
from pathlib import Path
from glob import glob
import win32com.client as win32
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
    excel_data = read_excel(fr'{Path.cwd()}\arquivos_gerados\teste - {data}.xlsx', sheet_name=0, header=0, usecols=['Processo'], dtype=str).dropna()
    excel_data_nome = excel_data['Processo'].tolist()

    excel_data = []
    for i in range(len(excel_data_nome)):
        excel_data.append([excel_data_nome[i]])
    print('{} processos lidos.\n'.format(len(excel_data)))
    return excel_data

def acessarPje():
    driver.get('https://pje2g.trf3.jus.br/')
    # Esperar fazer o login
    WebDriverWait(driver,600).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="j_id179"]/input[1]')))
    return


def esperarElemento():
    seconds = 0
    el_wait = True
    while el_wait and seconds < 20:
        sleep(1)
        el_wait = False
        try:
            driver.find_element(By.XPATH, "//a[@class='btn-link btn-condensed']").click()
        except:
            try:
                driver.find_element(By.XPATH, "//span[@class='rich-messages-label']")
            except:
                el_wait = True
            else:
                return False
        else:
            return True
        finally:
            seconds += 1
    return False


def validar(i):
    try:
        driver.switch_to.default_content()
        WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="navbar"]/ul/li/a[1]')))
        driver.find_element(By.XPATH, '//*[@id="navbar"]/ul/li/a[1]').click()
        WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="poloAtivo"]/table/tbody')))
        polo = {'ativo': driver.find_element(By.XPATH, '//*[@id="poloAtivo"]/table/tbody').find_elements(By.TAG_NAME, 'tr'),
                'passivo': driver.find_element(By.XPATH, '//*[@id="poloPassivo"]/table/tbody').find_elements(By.TAG_NAME, 'tr')}
        for p in polo[i]:
            if 'UNIAO FEDERAL - FAZENDA NACIONAL' in p.find_element(By.CSS_SELECTOR, 'td > span').text:
                rel = True
                break
            else:
                rel = False
        driver.find_element(By.XPATH, '//*[@id="navbar"]/ul/li/a[1]').click()
        WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameHtml')))
    except:
        rel = False
    return rel


def decisao(src):
    if 'ACOLHER' in src.upper() or 'ACOLHEU' in src.upper():
        if 'DESACOLHER' in src.upper() or 'DESACOLHEU' in src.upper():
            rel = 'Desacolhido'
        elif 'ACOLHER PARCIALMENTE' in src.upper() or 'ACOLHEU PARCIALMENTE' in src.upper():
            rel = 'Parcial Acolhido'
        else:
            rel = 'Acolhido'
    elif 'PROVIMENTO' in src.upper():
        if ('NEGAR PROVIMENTO' in src.upper() or 'NEGOU PROVIMENTO' in src.upper() or 'NEGANDO PROVIMENTO' in src.upper()
            or 'NEGAR-LHE PROVIMENTO' in src.upper() or 'NEGOU-LHE PROVIMENTO' in src.upper() or 'NEGANDO-LHE PROVIMENTO' in src.upper()):
            rel = 'Improvido'
        elif 'PARCIAL PROVIMENTO' in src.upper():
            rel = 'Parcial Provido'
        else:
            rel = 'Provido'
    elif 'PROCEDENTE' in src.upper():
        if 'IMPROCEDENTE' in src.upper():
            rel = 'Improcedente'
        else:
            rel = 'Procedente'
    elif 'MANTER O ACÓRDÃO' in src.upper() or 'MANTEVE O ACÓRDÃO' in src.upper():
        rel = 'Manteve Acórdão'
    elif 'MANTER A DECISÃO' in src.upper() or 'MANTEVE A DECISÃO' in src.upper():
        rel = 'Manteve Decisão'
    elif 'REJEITAR' in src.upper() or 'REJEITOU' in src.upper():
        rel = 'Rejeitado'
    elif 'SUSPENDER O JULGAMENTO' in src.upper() or 'SUSPENSO O JULGAMENTO' in src.upper():
        rel = 'Julgamento Suspenso'
    # Possível: Não exerceu -> Decisão | Juízo de retratação -> Recurso
    elif 'NÃO EXERCEU JUÍZO DE RETRATAÇÃO' in src.upper():
        rel = 'Não exerceu juízo de retratação'
    elif 'SEM DECISÃO' in src.upper():
        rel = 'Sem Decisão'
    elif 'ADIADO' in src.upper():
        rel = 'Adiado'
    else:
        rel = 'Não encontrado'
    return rel


def recurso(src):
    fil = src.upper().split('CERTIDÃO DE JULGAMENTO')[1]
    keyword = ['AGRAVO DE INSTRUMENTO', 'AGRAVO INOMINADO', 'AGRAVO INTERNO', 'CONFLITO NEGATIVO DE COMPETÊNCIA', 'EMBARGOS DE DECLARAÇÃO', 'JUÍZO DE RETRATAÇÃO NEGATIVO']
    
    for kw in keyword:
        if kw in fil:
            rel = kw.capitalize()
            break
        else:
            rel = 'Não encontrado'

    if rel == 'Não encontrado':
        if 'APELAÇÃO' in fil:
            if 'REMESSA OFICIAL' in fil:
                rel = 'Remessa oficial e recurso de apelação'
            else:
                rel = 'Apelação'
        elif 'RECURSO' in fil:
            rel = 'Recurso'
        elif 'DESCONSTITUIR' in fil:
            rel = 'Desconstituir'
    return rel


def busca(src, i):
    kw = [['SESSÃO REALIZADA EM'], ['APELANTE:', 'AGRAVANTE:', 'SUSCITANTE:', 'AUTOR:', 'PARTE AUTORA:', 'EMBARGANTE:'],
          ['APELADO:', 'AGRAVADO:', 'SUSCITADO:', 'REU:', 'PARTE RE:', 'EMBARGADO:']]
    c = 0
    while True:
        if c == len(kw[i]):
            rel = 'Não encontrado'
            break
        if kw[i][c] in src.upper():
            res = src.upper().split(kw[i][c])[1]
            if i == 0:
                rel = res.split(',')[0].strip()
            elif i == 1:
                rel = res.split('<BR>')[0].strip()
                if 'UNIAO FEDERAL - FAZENDA NACIONAL' not in rel:
                    if validar('ativo'):
                        rel += ' (UNIAO FEDERAL - FAZENDA NACIONAL)'
            else:
                rel = res.split('</P>')[0].strip()
                if 'UNIAO FEDERAL - FAZENDA NACIONAL' not in rel:
                    if validar('passivo'):
                        rel += ' (UNIAO FEDERAL - FAZENDA NACIONAL)'
            break
        else:
            c += 1
    return rel


def quorum(src):
    if 'À UNANIMIDADE' in src.upper() or 'POR UNANIMIDADE' in src.upper():
        rel = 'Unanimidade'
    elif 'POR MAIORIA' in src.upper():
        rel = 'Maioria'
    else:
        rel = 'Não encontrado'
    return rel


def resultado(dados):
    fav = ['Acolhido', 'Parcial Acolhido', 'Provido', 'Parcial Provido', 'Procedente']
    dfav = ['Desacolhido', 'Improvido', 'Improcedente', 'Rejeitado']
    res = 'Indefinido'
    if 'UNIAO FEDERAL - FAZENDA NACIONAL' in dados[3] and 'UNIAO FEDERAL - FAZENDA NACIONAL' in dados[4]:
        res = 'União em ambos os polos'
    elif 'UNIAO FEDERAL - FAZENDA NACIONAL' in dados[3]:
        if dados[0] in fav:
            res = 'Favorável à União'
        elif dados[0] in dfav:
            res = 'Desfavorável à União'
    elif 'UNIAO FEDERAL - FAZENDA NACIONAL' in dados[4]:
        if dados[0] in fav:
            res = 'Desfavorável à União'
        elif dados[0] in dfav:
            res = 'Favorável à União'
    return res


def coletarDados(processo):
    try:
        initial_window_handle = driver.current_window_handle
        driver.get('https://pje2g.trf3.jus.br/pje/Processo/ConsultaProcesso/listView.seam')
        WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, 'fPP:numeroProcesso:numeroSequencial')))
        driver.find_element(By.ID, 'fPP:numeroProcesso:numeroSequencial').send_keys(processo.replace('.4.03.', '.'))
        WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.ID, 'fPP:searchProcessos')))
        driver.find_element(By.ID, 'fPP:searchProcessos').click()
        
        exists = esperarElemento()
        if not exists:
            raise NoSuchElementException
        WebDriverWait(driver,20).until(EC.new_window_is_opened(set(driver.window_handles)))
        driver.switch_to.window(driver.window_handles[-1])
        WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, 'divTimeLine:txtPesquisa')))
        driver.find_element(By.ID, 'divTimeLine:txtPesquisa').send_keys('certidão de julgamento')

        bef = driver.find_element(By.ID, 'divTimeLine:eventosTimeLineElement').find_elements(By.CSS_SELECTOR, 'div.media.interno.tipo-D')
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
                    mov = driver.find_element(By.ID, 'divTimeLine:eventosTimeLineElement').find_elements(By.CSS_SELECTOR, 'div.media.interno.tipo-D')[i]
                    nome = mov.find_element(By.CSS_SELECTOR, 'div.media-body.box > div.anexos > a').text
                except (IndexError, NoSuchElementException):
                    if i == len(tam):
                        rel = ['Sem Decisão']
                        break
                    i += 1
                else:
                    if 'CERTIDÃO DE JULGAMENTO' in nome.upper():
                        mov.find_element(By.CSS_SELECTOR, 'div.media-body.box > div.anexos > a').click()
                        sleep(1)
                        WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameHtml')))
                        src = driver.page_source
                        # decisão, tipo de recurso, data da sessão, recorrente e recorrido, quorum
                        rel = [decisao(src), recurso(src), busca(src, 0), busca(src, 1), busca(src, 2), quorum(src)]
                        rel.append(resultado(rel))
                        break
                    else:
                        i += 1
        else:
            rel = ['Sem Decisão']
        for r in rel:
            print(f' | {r}', end='')
        print()
        driver.close()
        driver.switch_to.window(initial_window_handle)
    except:
        rel = ['Erro']
        print(' | Erro')
    return rel


def main():
    excel_data = lerExcel()
    acessarPje()
    for i in range(len(excel_data)):
        print(f'{i+1}- Processo: {excel_data[i][0]}', end='')
        test = coletarDados(excel_data[i][0])
        for t in test:
            excel_data[i].append(t)
    DataFrame(excel_data, columns=['Processo', 'Decisão', 'Recurso', 'Data', 'Recorrente', 'Recorrido', 'Quorum', 'Resultado']).to_excel(fr'{Path.cwd()}\arquivos_gerados\Decisão - {data}.xlsx', index=False, header=True)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fr'{Path.cwd()}\arquivos_gerados\Decisão - {data}.xlsx')
    wb.Worksheets("Sheet1").Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()
    print(fr'Planilha DECISÃO criada.')
    driver.quit()
    return


main()
