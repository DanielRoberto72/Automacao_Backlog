#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import selenium, os, time, pandas as pd, csv, warnings, shutil, sys, lxml, re, itertools, openpyxl, glob, mysql.connector, smtplib
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup as BS
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import win32com.client
import os
import glob
warnings.filterwarnings("ignore")
log = ''

tempo = datetime.now() - timedelta()
timestamp = tempo.strftime('%Y-%m-%d')
timestamp_envio = tempo.strftime('%d-%m')

#-----------------------------------------------------------------------------------------------
#SETANDO INFORMACOES FIXAS
login = "seu.login"
senha = "suaSenha123"
dirRaiz = 'C:/Prod/Python/BacklogOperacao/'
diretorio = dirRaiz + 'arquivos/'

#-----------------------------------------------------------------------------------------------
#INICIANDO O CHROMEDRIVER
chrome_options = webdriver.ChromeOptions()
chromedriver = dirRaiz+"Driver/chromedriver.exe"
prefs = {"download.default_directory": r"C:\Prod\Python\BacklogOperacao\arquivos"}
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument('ignore-certificate-errors')
driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=chromedriver)
#-----------------------------------------------------------------------------------------------
def wait_xpath_click(y):
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, y))).click()
#-----------------------------------------------------------------------------------------------
#BAIXANDO RELATÓRIO DO OTRS

try:
    driver.get("link.html")
    driver.find_element_by_name('Login').send_keys(login)
    driver.find_element_by_name ('Senha').send_keys(senha)
    print('lOGADO!')
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[1]/div[2]/form/button'))).click()
    driver.get('link2.html')
    print('Aguardando para fazer o download')
    time.sleep(120)
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div[2]/div/div/div[1]/div/button[2]'))).click()
    WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div[2]/div/div/div[1]/div/ul/li/a'))).click()
    print('Aguardando o download do relatório! (2 Minutos)')
    time.sleep(120)
    print('Download realizado')
    driver.close()
except:
    print('Falha ao executar o script, script finalizado!')
    sys.exit()

#-----------------------------------------------------------------------------------------------
#FAZENDO MANIPULAÇÃO DA PLANILHA
o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
dirRaiz = 'C:/Prod/Python/BacklogOperacao/'
diretorio = dirRaiz + 'arquivos/'
files = glob.glob(diretorio + "/*.xls")
filename= diretorio +'relatorio_servico_realtime'
file = os.path.basename(filename)
output = diretorio + '/' + file.replace('.xls','.xlsx')
name='relatorio_servico_realtime'
wb = o.Workbooks.Open(filename)
wb.ActiveSheet.SaveAs(name,51)
wb.Close(True)
#Deletando arquivo xls
os.remove(diretorio+'relatorio_servico_realtime.xls')

#-----------------------------------------------------------------------------------------------
#FAZENDO BATIMENTOS DOS CHAMADOS POR FILA
dirRaiz = 'C:/Prod/Python/BacklogOperacao/'
diretorio = dirRaiz + 'arquivos/'

df = pd.read_excel(diretorio+'relatorio_servico_realtime.xlsx').astype(str)

try:
    lista1 = ['NIVEL 2::','NIVEL 2::','NIVEL 2::']
    Mvno = ['MVNO']
    lista2 = ['NIVEL 2::','NIVEL 2::']
    lista3 = ['NIVEL 3::', 'NIVEL 2::']
    lista4 = ['NIVEL 2::']
    lista5 = ['NIVEL 3::']

    df1 = df[df['FILA'].isin(lista1)].reset_index(drop=True)
    dfMVNO = df1[df1['MVNO'].isin(Mvno)].reset_index(drop=True)
    df2 = df1[~df1['MVNO'].isin(Mvno)].reset_index(drop=True)
    df3 = df[df['FILA'].isin(lista2)].reset_index(drop=True)
    df4 = df[df['FILA'].isin(lista3)].reset_index(drop=True)
    df5 = df[df['FILA'].isin(lista4)].reset_index(drop=True)
    df6 = df[df['FILA'].isin(lista5)].reset_index(drop=True)

except:
    print('FALHA!!!!')
    the_type, the_value, the_traceback = sys.exc_info()
    erro = 'Falha nos batimentos das filas'
    print(the_type, ',' ,the_value,',', the_traceback)

#-----------------------------------------------------------------------------------------------
#CONTAGEM E TRANSFORMAÇÃO PARA STRING DOS CHAMADOS POR FILA
n1 = int(len(dfMVNO))
n2 = int(len(df3))
n3 = int(len(df2))
n4 = int(len(df4))
n5 = int(len(df5))
n6 = int(len(df6))

n1Str = str(n1)
n2Str = str(n2)
n3Str = str(n3)
n4Str = str(n4)
n5Str = str(n5)
n6Str = str(n6)

print(n1Str)
print(n2Str)
print(n3Str)
print(n4Str)
print(n5Str)
print(n6Str)


#-----------------------------------------------------------------------------------------------
#ENVIO DO EMAIL PARA OS DESTINATÁRIOS FIXOS
tempo = datetime.now() - timedelta()
timestamp_envio = tempo.strftime('%d-%m')
try:
    email = 'emaildeenvio@email.com.br'
    password = 'suasenha'
    send_to_email = ['lista@email','lista@email']
    subject = 'BACKLOG OPERAÇÃO '+timestamp_envio
    message ='''
Bom dia!

Segue análise das filas 

Fila 1: '''+n1Str+'''
Fila 2: '''+n2Str+'''
Fila 3: '''+n3Str+'''
Fila 4:  '''+n4Str+'''
Fila 5: '''+n5Str+'''
Fila 6: '''+n6Str+'''

Atenciosamente. '''
    
    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = ", ".join(send_to_email)
    msg['Subject'] = subject

    msg.attach(MIMEText(message, 'plain'))
    

    server = smtplib.SMTP('SMTP.office365.com',587)
    server.starttls()
    server.login(email, password)
    text = msg.as_string()
    server.sendmail(email, send_to_email, text)
    server.quit()
    
    print('Email enviado COM SUCESSO PARA OS DESTINATÁRIOS!')
except:
    print('Falha ao enviar o Email!')
    the_type, the_value, the_traceback = sys.exc_info()
    print(the_type, ',' ,the_value,',', the_traceback)
    pass

try:
    os.remove(diretorio+'relatorio_servico_realtime.xlsx')
    print('Processo de remover planilha finalizado!!!')
except:
    print('falha ao remover arquivo ou arquivo já foi removido')

