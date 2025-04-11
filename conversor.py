
import pandas as pd
import os



arquivo = []

tabelas = []

#leitura de arquivos    
#exceções obrigatorias: não pode clicar em converter sem fazer upload, limite até 10mb, arquivo quebrado, arquivo em formato diferente


def verifica_tamanho(arquivo):
    tamanho = os.path.getsize(arquivo)
    if tamanho > 10_485_760:
        return "Arquivo maior que 10mb"


def ler_csv(arquivo):
    if not arquivo: #primeiro verifica se foi enviado
        return "arquivo não enviado"
    if (arquivo.filename.endswith('csv')): #depois verifica se é csv
        df = pd.read_csv(arquivo) #depois lê
        return df
    

def ler_xls(arquivo, engine='xlrd'): 
    if not arquivo: #primeiro verifica se enviou
        return "arquivo não enviado"
    if (arquivo.filename.endswith('xls')): #depois se é xls
        df = pd.read_excel(arquivo, engine=engine) #depois lê
        return df


def ler_xlsx(arquivo, engine='openpyxl'):
    if not arquivo:
        return "arquivo não enviado"
    if (arquivo.filename.endswith('xlsx')):
        df = pd.read_excel(arquivo, engine=engine)
        return df


def ler_ods (arquivo: str, engine='odf'):
    if not arquivo:
        return "arquivo não enviado"
    if (arquivo.filename.endswith('ods')):
        df = pd.read_excel(arquivo, engine=engine)
        return df 

#conversão de arquivos
def csv_xlsx(arquivo, tabela):
    nome_excel = arquivo.replace('.csv', '.xlsx')
    tabela.to_excel(nome_excel, index=False, engine='openpyxl')
    return nome_excel

def csv_xls(arquivo, tabela):
    nome_excel = arquivo.replace('.csv', '.xls')
    tabela.to_excel(nome_excel, index=False, engine='xlwt')
    return nome_excel

def xlsx_csv(arquivo, tabela):
    nome_csv = arquivo.replace('.xlsx', '.csv')
    tabela.to_csv(nome_csv, index=False)
    return nome_csv

def xlsx_xls(arquivo, tabela):
    nome_excel = arquivo.replace('.xlsx', '.xls')
    tabela.to_excel(nome_excel, index=False, engine='xlwt')
    return nome_excel

def xls_csv(arquivo, tabela):
    nome_csv = arquivo.replace('xls', '.csv')
    tabela.to_csv(nome_csv, index=False)
    return nome_csv

def xls_xlsx(arquivo, tabela):
    nome_excel = arquivo.replace('xls', '.xlsx')
    tabela.to_excel(nome_excel, index=False, engine='openpyxl')
    return nome_excel

def ods_csv(arquivo, tabela):
    nome_csv = arquivo.replace('.ods', '.csv')
    tabela.to_csv(nome_csv, index=False)
    return nome_csv

def ods_xlsx(arquivo, tabela):
    nome_xlsx = arquivo.replace('.ods', '.xlsx')
    tabela.to_excel(nome_xlsx, index=False, engine='openpyxl')
    return nome_xlsx

def ods_xls(arquivo, tabela):
    nome_xls = arquivo.replace('.ods', '.xls')
    tabela.to_excel(nome_xls, index=False, engine='xlwt')
    return nome_xls
