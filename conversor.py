
import pandas as pd



arquivos = []

tabelas = []

#leitura de arquivos    
def ler_csv(arquivo):
    df = pd.read_csv(arquivo)
    return df

def ler_xls(arquivo, engine='xlrd'):
    df = pd.read_excel(arquivo, engine=engine)
    return df


def ler_xlsx(arquivo, engine='openpyxl'):
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