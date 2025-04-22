import pandas as pd
import os
from pyexcel_ods3 import save_data

def verifica_tamanho(arquivo_path):
    if not arquivo_path:
        return "Arquivo não enviado"
    try:
        tamanho = os.path.getsize(arquivo_path)
        if tamanho > 10_485_760:
            return "Arquivo maior que 10MB"
        return None
    except Exception as e:
        return f"Erro: {str(e)}"

def verifica_extensao(arquivo_path):
    extensao = os.path.splitext(arquivo_path)[1].lower()
    match extensao:
        case '.csv': 
            return ler_csv(arquivo_path)
        case '.xls': 
            return ler_xls(arquivo_path)
        case '.xlsx': 
            return ler_xlsx(arquivo_path)
        case '.ods': 
            return ler_ods(arquivo_path)
        case _: 
            return None

def ler_csv(arquivo_path):
    try:
        return pd.read_csv(arquivo_path, engine='c')
    except Exception as e:
        return f"Erro ao ler CSV: {str(e)}"

def ler_xls(arquivo_path):
    try:
        return pd.read_excel(arquivo_path, engine='xlrd')
    except Exception as e:
        return f"Erro ao ler XLS: {str(e)}"

def ler_xlsx(arquivo_path):
    try:
        return pd.read_excel(arquivo_path, engine='openpyxl')
    except Exception as e:
        return f"Erro ao ler XLSX: {str(e)}"

def ler_ods(arquivo_path):
    try:
        return pd.read_excel(arquivo_path, engine='odf')
    except Exception as e:
        return f"Erro ao ler ODS: {str(e)}"

def to_ods(df, nome_ods):
    try:    
        sheet_data = {"Sheet1": [df.columns.tolist()] + df.values.tolist()}
        save_data(nome_ods, sheet_data)
        return nome_ods
    except Exception as e:
        return f"Erro ao salvar ODS: {str(e)}"

def conversor(arquivo_path, opcao):
    # Validação inicial
    tamanho_erro = verifica_tamanho(arquivo_path)
    if tamanho_erro:
        return tamanho_erro
    
    df = verifica_extensao(arquivo_path)
    if not isinstance(df, pd.DataFrame):
        return df if isinstance(df, str) else "Erro na leitura do arquivo"

    try:
        match opcao:
            case 1:
                nome = arquivo_path.replace('.csv', '.xlsx')
                df.to_excel(nome, index=False, engine='openpyxl')
                return nome
            case 2:
                nome = arquivo_path.replace('.csv', '.xls')
                df.to_excel(nome, index=False, engine='xlwt')
                return nome
            case 3:
                nome = arquivo_path.replace('.csv', '.ods')
                return to_ods(df, nome)
            case 4:
                nome = arquivo_path.replace('.xls', '.csv')
                df.to_csv(nome, index=False)
                return nome
            case 5:
                nome = arquivo_path.replace('.xls', '.xlsx')
                df.to_excel(nome, index=False, engine='openpyxl')
                return nome
            case 6:
                nome = arquivo_path.replace('.xls', '.ods')
                return to_ods(df, nome)
            case 7:
                nome = arquivo_path.replace('.xlsx', '.csv')
                df.to_csv(nome, index=False)
                return nome
            case 8:
                nome = arquivo_path.replace('.xlsx', '.xls')
                df.to_excel(nome, index=False, engine='xlwt')
                return nome
            case 9:
                nome = arquivo_path.replace('.xlsx', '.ods')
                return to_ods(df, nome)
            case 10:
                nome = arquivo_path.replace('.ods', '.csv')
                df.to_csv(nome, index=False)
                return nome
            case 11:
                nome = arquivo_path.replace('.ods', '.xlsx')
                df.to_excel(nome, index=False, engine='openpyxl')
                return nome
            case 12:
                nome = arquivo_path.replace('.ods', '.xls')
                df.to_excel(nome, index=False, engine='xlwt')
                return nome
            case _:
                return "Opção inválida"
    except Exception as e:
        return f"Erro na conversão: {str(e)}"
