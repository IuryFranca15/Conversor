
import pandas as pd
import os


arquivo = []

tabelas = []

#leitura de arquivos    
#exceções obrigatorias: não pode clicar em converter sem fazer upload, limite até 10mb, arquivo quebrado, arquivo em formato diferente

#se o tamanho é até 10mb
def verifica_tamanho(arquivo):
    if not arquivo:
        return "arquivo não enviado"
    tamanho = os.path.getsize(arquivo)
    if tamanho > 10_485_760:
        return "Arquivo maior que 10mb"
    else:
        return verifica_extensao(arquivo)
        
def verifica_extensao(arquivo):
    if arquivo.filename.endswith('csv'):
        return ler_csv(arquivo)

    elif arquivo.filename.endswith('xls'):
        return ler_xls(arquivo)

    elif arquivo.filename.endswith('xlsx'):
        return ler_xlsx(arquivo)

    elif arquivo.filename.endswith('ods'):
        return ler_ods(arquivo)
    else:
        print("formato invalido")
        return None    

def ler_csv(arquivo):
    df = pd.read_csv(arquivo, engine = 'c') #depois lê
    return df

def ler_xls(arquivo, engine='xlrd'): 
    df = pd.read_excel(arquivo, engine=engine) #depois lê
    return df

def ler_xlsx(arquivo, engine='openpyxl'):
    df = pd.read_excel(arquivo, engine=engine)
    return df

def ler_ods (arquivo: str, engine='odf'):
    df = pd.read_excel(arquivo, engine=engine)
    return df

def to_ods(df, nome_ods, save_data):
    # Converte o DataFrame para um formato que o pyexcel-ods3 entende
    data = df.to_dict(orient='split')  # Converte para dicionário
    sheet_data = { "Sheet1": data['data'] }  # O pyexcel-ods3 precisa de uma lista de listas
    save_data(nome_ods, sheet_data)
    
#conversão de arquivo
def conversor(df, arquivo, opcao, save_data):
    match opcao:
        case 1:
            nome_excel = arquivo.replace('.csv', '.xlsx')
            df.to_excel(nome_excel, index=False, engine='openpyxl')
            return nome_excel
        case 2:
            nome_excel = arquivo.replace('.csv', '.xls')
            arquivo.to_excel(nome_excel, index=False, engine='xlwt')
            return nome_excel
        case 3: 
            nome_ods = arquivo.replace('.csv', '.ods')
            to_ods(df, nome_ods)  # Usa o método to_ods para converter
            return nome_ods
        case 4:
            nome_csv = arquivo.replace('.xlsx', '.csv')
            arquivo.to_csv(nome_csv, index=False)
            return nome_csv
        case 5:
            nome_excel = arquivo.replace('.xlsx', '.xls')
            arquivo.to_excel(nome_excel, index=False, engine='xlwt')
            return nome_excel
        case 6:
            nome_excel = arquivo.replace('.xlsx', '.ods')
            to_ods(df, nome_ods)
            return nome_ods
        case 7:
            nome_csv = arquivo.replace('.xls', '.csv')
            arquivo.to_csv(nome_csv, index=False)
            return nome_csv
        case 8:
            nome_excel = arquivo.replace('.xls', '.xlsx')
            arquivo.to_excel(nome_excel, index=False, engine='openpyxl')
            return nome_excel
        case 9:
            nome_excel = arquivo.replace('.xls', '.ods')
            to_ods(df, nome_ods)
            return nome_ods
        case 10:
            nome_csv = arquivo.replace('.ods', '.csv')
            arquivo.to_csv(nome_csv, index=False)
            return nome_csv
        case 10:
            nome_xlsx = arquivo.replace('.ods', '.xlsx')
            arquivo.to_excel(nome_xlsx, index=False, engine='openpyxl')
            return nome_xlsx
        case 9:
            nome_xls = arquivo.replace('.ods', '.xls')
            arquivo.to_excel(nome_xls, index=False, engine='xlwt')
            return nome_xls


    