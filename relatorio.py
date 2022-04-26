from email.mime.base import MIMEBase
import threading
import click
from selenium import webdriver
import time
import os
from zipfile import ZipFile
import pandas as pd

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from email.mime.base import MIMEBase
from email import encoders

import smtplib
import email.message




# criamos a Classe relatorio


class relatorio:
    def __init__(self):
        # iniciamos os pacotes chromedrive
        options = webdriver.ChromeOptions()
        options.add_argument('lang=pt-br')
        # identificamos o login e a senha e colocamos em variaveis (self.log, self.senha)
        self.log = "INSIRA SEU LOGIN"
        self.senha = 'INSIRA SUA SENHA'
        self.planilha = 'relatorio_produtos.xlsx'
        self.driver = webdriver.Chrome(
            executable_path=r'./chromedriver.exe', chrome_options=options)
        # iniciamos o processo de login

    def LogarOnedriver(self):
        # solicitamos a abertura do site atrvés do metodo .GET (pegamos o http da planilha que desejamos)
        self.driver.get(
            'https://onedrive.live.com/?id=root&cid=CE93BB21C4BCE301')
        time.sleep(3)
        # identificamos a barra de login
        login = self.driver.find_element_by_id('i0116')
        # clicamos na barra de login
        login.click()
        time.sleep(2)
        # inserimos o email ja instanciado acima (self.log)
        login.send_keys(self.log)
        # selecionamos o botão avançar
        emailavancar = self.driver.find_element_by_id('idSIButton9')
        # clicamos no botão avancar
        emailavancar.click()

        time.sleep(3)

        # selecionamos a caixa de texto da senha
        senha = self.driver.find_element_by_id('i0118')
        # clicamos na barra de senha
        senha.click()
        # inserimos a senha na barra de senha
        senha.send_keys(self.senha)
        # selecionar o botão de logar
        entrar = self.driver.find_element_by_id('idSIButton9')
        # clicar para entrar
        entrar.click()
        time.sleep(1)

        # continuar conectado??
        conectado = self.driver.find_element_by_xpath('//*[@id="idSIButton9"]')
        # clicar para continuar conectado
        conectado.click()
        time.sleep(2)

        # clicar na pasta python
        self.driver.get(
            'https://onedrive.live.com/?id=CE93BB21C4BCE301%21232624&cid=CE93BB21C4BCE301')
        time.sleep(2)
        # cliar para fazer dowload da planilha atualizada
        dowload = self.driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div/div/div[1]/div[4]/button/span/span/span')
        dowload.click()

bot = relatorio()
bot.LogarOnedriver()

time.sleep(260)
# mover o arquivo/PASTA de download para essa pasta (relatorio)
DePasta = "/home/jota/Downloads"
ParaPasta = "/home/jota/Documentos/Automação_Python/relatorio"
# lupando a pasta
os.chdir(DePasta)
# percorre arquivos e manda para a pasta desejada
for f in os.listdir():
    os.rename(f, ParaPasta + f)

time.sleep(5)

# vamos extrair o arquivo
file_ext = "/home/jota/Documentos/Automação_Python/relatoriopython.zip"
with ZipFile(file_ext, "r") as extra:
    extra.printdir()
    extra.extractall()
time.sleep(5)

#TRATANDO OS DADOS

df = pd.read_csv('/home/jota/Documentos/Automação_Python/relatoriocaso_full.xlsx')
#agrupar
df = pd.DataFrame(df.groupby(['city','state']).sum()['new_confirmed'])

df.to_csv("Dados.xlsx")

time.sleep(5)

#ENVIAR EMAIL / LOGIN DO EMAIL(GMAIL)

host = "smtp.gmail.com"

port = "587"

login = "INSIRA SEU LOGIN"

senha = "INSIRA SUA SENHA"

server = smtplib.SMTP(host,port)
server.ehlo()
server.starttls()
server.login(login,senha)

corpo = '<b>Segue relatorio do dia!</b>'

email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = login
email_msg['Subject'] = "Boa noite João, aqui está o relatorio dos casos de covid do Brasil."
email_msg.attach(MIMEText(corpo, 'html'))

cam_arquivo = "/home/jota/Documentos/Automação_Python/relatorioDados.xlsx"
attchment = open(cam_arquivo, 'rb')

att = MIMEBase('application', 'octet-stream')
att.set_payload(attchment.read())
encoders.encode_base64(att)

att.add_header('Content-Disposition', f'attachment; filename=relatorioDados.xlsx')
attchment.close()
email_msg.attach(att)

server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
server.quit()

print("email enviado")

#excluindo arquivos que não tem ultilidade usando Os
if os.path.exists("/home/jota/Documentos/Automação_Python/"):
        os.remove("/home/jota/Documentos/Automação_Python/relatoriocaso_full.xlsx")
        os.remove("/home/jota/Documentos/Automação_Python/relatorioDados.xlsx")
        os.remove("/home/jota/Documentos/Automação_Python/relatoriopython.zip")
else:
        print("os arquivos foram removidos")


