from imap_tools import MailBox, AND
import openpyxl
import os 
from datetime import datetime

login = "push.demandas.prfn3regiao@pgfn.gov.br"
senha = "smybwsrarkugrcls"

try:
    meu_email = MailBox("imap.gmail.com").login(login, senha)

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    lista_emails = meu_email.fetch(AND(from_="pje@trf3.jus.br"))

    cabecalho = ["Processo"]
    sheet.append(cabecalho)

    for email in lista_emails:
        subject = email.subject
        split_subject = subject.split()
        
        if len(split_subject) >= 5:
            quinto_split = split_subject[4]
        else:
            quinto_split = "Não disponível"
        
        sheet.append([quinto_split])

    data_atual = datetime.now().strftime("Processos PUSH - %#d-%#m-%Y")

    nome_arquivo = f"{data_atual}.xlsx"

    pasta_saida = "arquivos_gerados"

    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)


    caminho_arquivo = os.path.join(pasta_saida, nome_arquivo)
    workbook.save(caminho_arquivo)
    print(f"Arquivo salvo em: {caminho_arquivo}")


except Exception as e:
    print(f'Ocorreu um erro: {e}')

finally:
    if meu_email:
        meu_email.logout()
