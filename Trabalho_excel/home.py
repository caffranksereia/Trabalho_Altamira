# Utilizar pip install selenium
# Utilizar pip install webdriver-manager
# Utilizar pip install openpyxl
# Utilizar pip install pandas

import os            #Criar pasta destino
import time          #Coloca tempo de espera
import shutil        #Movimentação de arquivo
import zipfile       #Descompactação de arquivo
import pandas as pd  #Trabalhar com arquivos excel

#Envio de email com anexo
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

#Automatização do download de arquivo utilizando Chrome
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

#Solicita crendencial de acesso
user = input("Informe o login de acesso:")
password_user = input("Informe a senha de acesso:")

driver = webdriver.Chrome(ChromeDriverManager().install()) #Inicia o Chrome
driver.maximize_window() #Maximiza a janela do Chrome

# O time sleep no código, aguarda o carregamento antes de executar o próximo, assim dando tempo de carregar a pagina por completo.

#Acessa o site da faculdade e clica em "Entrar"
driver.get("https://estudante.estacio.br/")
time.sleep(5)

#Clica no botão Entrar para começar o login
button0 = driver.find_element_by_xpath("/html/body/div[1]/section/div/div/div/section/div[1]/button").click()
time.sleep(15)

#Pega o campo do login pelo ID e envia o email de acesso.
username_textbox=driver.find_element_by_id("i0116")
username_textbox.send_keys(user)

#Clica no botão para avançar pelo ID
button1 = driver.find_element_by_id("idSIButton9").click()
time.sleep(10)

#Pega o campo da senha pelo ID e envia a senha de acesso.
password=driver.find_element_by_id("i0118")
password.send_keys(password_user)

#Clica para efetuar o login no site
login_attempt = driver.find_element_by_id("idSIButton9")
login_attempt.submit()
time.sleep(10)

#Clica em "Não" na caixa de mensagem que pergunta se quer se manter conectado.
button2 = driver.find_element_by_id("idBtn_Back").click()
time.sleep(10)

#Acessa a diciplina de Python
button3 = driver.find_element_by_xpath("/html/body/div[1]/section/article[2]/article[2]/div/section[2]/div/div/div/article/div/div[1]/section/header/button").click()
time.sleep(10)

#Acessa o tema do trabalho
button4 = driver.find_element_by_xpath("/html/body/div[1]/section/article[2]/article[2]/section/div/div/div[1]/section/div/div/article/section[3]/div[2]/div[3]/div/section/div[5]/a").click()
time.sleep(10)

#Clica no conteudo complementar para efetuar o download do arquivo.
button5 = driver.find_element_by_xpath("/html/body/div[1]/section/article[2]/article[2]/section/div/div/div[2]/section/div/div/article/section/section/section/div[3]/div[2]/div/section/div/div[1]/button").click()
time.sleep(10)

#Fecha o Browser Google Chrome
driver.quit()

#Cria o diretório onde iremos salvar o arquivo baixado e verifica se a pasta já existe.
os.makedirs("./TrabalhoAltamira", exist_ok=False)

#Pega o arquivo baixado e move para o diretorio que criado.
source = r"C:/Users/201801196982/Downloads/combate_pirataria.zip"
destination = r"C:/TrabalhoAltamira/combate_pirataria.zip"
shutil.move(source, destination)
time.sleep(5)

#Extrai o arquivo baixado e coloca na mesma pasta que criamos
extrair = zipfile.ZipFile("C:/201801196982/TrabalhoAltamira/combate_pirataria.zip")
extrair.extractall("C:/TrabalhoAltamira/")
extrair.close()

#Leitura do arquivo original
df = pd.read_csv("C:/TrabalhoAltamira/Tabela_PACP.csv", sep=";")

#Total de cada coluna com quantidade
qt_un_apreendidas = df["QT_UN_APREENDIDAS"].sum()
qt_un_lacradas = df["QT_UN_LACRADAS"].sum()
qt_un_retidas = df["QT_UN_RETIDAS"].sum()
qt_un_retiradas = df["QT_UN_RETIRADAS"].sum()

#Quantidade de itens NÃO-NULOS
qtd_equip_apreendidas = df.QT_UN_APREENDIDAS.count()
qtd_equip_lacradas = df.QT_UN_LACRADAS.count()
qtd_equip_retidas = df.QT_UN_RETIDAS.count()
qtd_equip_retiradas = df.QT_UN_RETIRADAS.count()

#Média aritimética de cada coluna
media_apreendidas = qt_un_apreendidas/qtd_equip_apreendidas
media_lacradas = qt_un_lacradas/qtd_equip_lacradas
media_retidas = qt_un_retidas/qtd_equip_retidas
media_retiradas = qt_un_retiradas/qtd_equip_retiradas

#Menor valor de cada coluna
min_apreendidas = df["QT_UN_APREENDIDAS"].min(skipna=True)
min_lacradas = df["QT_UN_LACRADAS"].min(skipna=True)
min_retidas = df["QT_UN_RETIDAS"].min(skipna=True)
min_retiradas = df["QT_UN_RETIRADAS"].min(skipna=True)

#Maior valor de cada coluna
max_apreendidas = df["QT_UN_APREENDIDAS"].max(skipna=True)
max_lacradas = df["QT_UN_LACRADAS"].max(skipna=True)
max_retidas = df["QT_UN_RETIDAS"].max(skipna=True)
max_retiradas = df["QT_UN_RETIRADAS"].max(skipna=True)

#Cria um novo excel com os dados organizados e os totais
totais_list = [("TOTAIS:","","",qt_un_apreendidas,qt_un_lacradas,qt_un_retidas,"",qt_un_retiradas)]
qtd_itens = [("TOTAL ITENS:","","",qtd_equip_apreendidas,qtd_equip_lacradas,qtd_equip_retidas,"",qtd_equip_retiradas)]
media_col = [("MEDIA:","","",media_apreendidas,media_lacradas,media_retidas,"",media_retiradas)]
min_val = [("MENOR VALOR:","","",min_apreendidas,min_lacradas,min_retidas,"",min_retiradas)]
max_val = [("MAIOR VALOR:","","",max_apreendidas,max_lacradas,max_retidas,"",max_retiradas)]

dfcalc = pd.DataFrame(totais_list, columns=["ID", "Área","Equipamento","QT_UN_APREENDIDAS","QT_UN_LACRADAS","QT_UN_RETIDAS","Valor Estimado","QT_UN_RETIRADAS"])
dfqtd = pd.DataFrame(qtd_itens, columns=["ID", "Área","Equipamento","QT_UN_APREENDIDAS","QT_UN_LACRADAS","QT_UN_RETIDAS","Valor Estimado","QT_UN_RETIRADAS"])
dfmedia = pd.DataFrame(media_col, columns=["ID", "Área","Equipamento","QT_UN_APREENDIDAS","QT_UN_LACRADAS","QT_UN_RETIDAS","Valor Estimado","QT_UN_RETIRADAS"])
dfmin = pd.DataFrame(min_val, columns=["ID", "Área","Equipamento","QT_UN_APREENDIDAS","QT_UN_LACRADAS","QT_UN_RETIDAS","Valor Estimado","QT_UN_RETIRADAS"])
dfmax = pd.DataFrame(max_val, columns=["ID", "Área","Equipamento","QT_UN_APREENDIDAS","QT_UN_LACRADAS","QT_UN_RETIDAS","Valor Estimado","QT_UN_RETIRADAS"])

df = df.append(dfcalc,ignore_index=True)
df = df.append(dfqtd,ignore_index=True)
df = df.append(dfmedia,ignore_index=True)
df = df.append(dfmin,ignore_index=True)
df = df.append(dfmax,ignore_index=True)

df.to_excel(r"C:/Fabio/TrabalhoAltamira/Resultado_Tabela.xlsx", index=False, header=True)

#Envio de e-mail automatico utilizando outlook

#Login e senha, e autenticação com o servidor
host = "smtp.office365.com"
port = 587
user_email = input("Informe o e-mail de acesso:")
password_email = input("Informe a senha de acesso:")

server = smtplib.SMTP(host, port)

server.ehlo()
server.starttls()
server.login(user_email, password_email)

#Informa o destinatario, titulo, monta a mensagem do corpo(HTML) e anexo o arquivo.
email_msg = MIMEMultipart()
email_msg["From"] = user_email
email_msg["to"] = "fabioeduardocircuncisao@gmail.com"
email_msg["Subject"] = "Trabalho - Paradigmas de Linguagens de Programação em Python"

menssage = """<p>Olá, boa noite! Tudo bem?</p>
<p>
<p>Segue o arquivo em anexo conforme solicitado.</p>
</p>
<p>Atenciosamente,</p>
<p>Fabio Eduardo</p> &
<p>João Vitor </p>
<p>Matricula: 201801196982</p>
<p>Matricula: </p>
"""
email_msg.attach(MIMEText(menssage, "html"))

anexo_arq = "C:/Fabio/TrabalhoAltamira/Resultado_Tabela.xlsx"
attachment = open(anexo_arq, "rb")

att = MIMEBase("application","octet-stream")
att.set_payload(attachment.read())
encoders.encode_base64(att)

att.add_header("Content-Disposition", "attachment; filename=Resultado_Tabela.xlsx")
attachment.close()
email_msg.attach(att)

#Efetua o envio do email
server.sendmail(email_msg["From"], email_msg["To"], email_msg.as_string())
print("Email enviado com sucesso!")

#Finaliza a autenticação com o servidor.
server.quit()
