
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
    if os.path.basename(arquivo).endswith('.csv'):
        return ler_csv(arquivo)

    elif os.path.basename(arquivo).endswith('.xls'):
        return ler_xls(arquivo)

    elif os.path.basename(arquivo).endswith('.xlsx'):
        return ler_xlsx(arquivo)

    elif os.path.basename(arquivo).endswith('.ods'):
        return ler_ods(arquivo)
    else:
        print("formato invalido") #formato diferente
        return None    

def ler_csv(arquivo):
    try:
        df = pd.read_csv(arquivo, engine = 'c') #depois lê
        return df
    except Exception as e:
        print(f"Erro ao ler CSV: {e}")
        return None

def ler_xls(arquivo, engine='xlrd'): 
    try:
        df = pd.read_excel(arquivo, engine=engine) #depois lê
        return df
    except Exception as e:
        print(f"Erro ao ler xls: {e}")
        return None
    
def ler_xlsx(arquivo, engine='openpyxl'):
    try:   
        df = pd.read_excel(arquivo, engine=engine)
        return df
    except Exception as e:
        print(f"Erro ao ler xlsx: {e}")
        return None

def ler_ods (arquivo: str, engine='odf'):
    try:
        df = pd.read_excel(arquivo, engine=engine)
        return df
    except Exception as e:
        print(f"Erro ao ler ods: {e}")
        return None

def to_ods(df, nome_ods, save_data):
        # Converte o DataFrame para um formato que o pyexcel-ods3 entende
    try:    
        data = df.to_dict(orient='split')  # Converte para dicionário
        sheet_data = { "Sheet1": data['data'] }  # O pyexcel-ods3 precisa de uma lista de listas
        save_data(nome_ods, sheet_data)
    except Exception as e:
        print(f"Erro ao converter para ODS: {e}")
        return None    
    
#conversão de arquivo
def conversor(df, arquivo, opcao, save_data):
    match opcao:
        # CSV para outros formatos
        case 1:
            try:
                nome_excel = arquivo.replace('.csv', '.xlsx')
                df.to_excel(nome_excel, index=False, engine='openpyxl')
                print(f"Arquivo {nome_excel} convertido com sucesso!")
                return nome_excel
            except Exception as e:
                print(f"Erro ao converter para .xlsx: {e}")
                return None
        case 2:
            try:    
                nome_excel = arquivo.replace('.csv', '.xls')
                df.to_excel(nome_excel, index=False, engine='xlwt')
                print(f"Arquivo {nome_excel} convertido com sucesso!")
                return nome_excel
            except Exception as e:
                print(f"Erro ao converter para .xls: {e}")
                return None
        case 3: 
            try:
                nome_ods = arquivo.replace('.csv', '.ods')
                to_ods(df, nome_ods)
                print(f"Arquivo {nome_ods} convertido com sucesso!")
                return nome_ods
            except Exception as e:
                print(f"Erro ao converter para .ods: {e}")
                return None

        # XLS para outros formatos
        case 4:
            try:
                nome_csv = arquivo.replace('.xls', '.csv')
                df.to_csv(nome_csv, index=False)
                print(f"Arquivo {nome_csv} convertido com sucesso!")
                return nome_csv
            except Exception as e:
                print(f"Erro ao converter para .csv: {e}")
                return None
        case 5:
            try:
                nome_excel = arquivo.replace('.xls', '.xlsx')
                df.to_excel(nome_excel, index=False, engine='openpyxl')
                print(f"Arquivo {nome_excel} convertido com sucesso!")
                return nome_excel
            except Exception as e:
                print(f"Erro ao converter para .xlsx: {e}")
                return None
        case 6:
            try:
                nome_ods = arquivo.replace('.xls', '.ods')
                to_ods(df, nome_ods)
                print(f"Arquivo {nome_ods} convertido com sucesso!")
                return nome_ods
            except Exception as e:
                print(f"Erro ao converter para .ods: {e}")
                return None

        # XLSX para outros formatos
        case 7:
            try:
                nome_csv = arquivo.replace('.xlsx', '.csv')
                df.to_csv(nome_csv, index=False)
                print(f"Arquivo {nome_csv} convertido com sucesso!")
                return nome_csv
            except Exception as e:
                print(f"Erro ao converter para .csv: {e}")
                return None
        case 8:
            try:
                nome_excel = arquivo.replace('.xlsx', '.xls')
                df.to_excel(nome_excel, index=False, engine='xlwt')
                print(f"Arquivo {nome_excel} convertido com sucesso!")
                return nome_excel
            except Exception as e:
                print(f"Erro ao converter para .xls: {e}")
                return None
        case 9:
            try:
                nome_ods = arquivo.replace('.xlsx', '.ods')
                to_ods(df, nome_ods)
                print(f"Arquivo {nome_ods} convertido com sucesso!")
                return nome_ods
            except Exception as e:
                print(f"Erro ao converter para .ods: {e}")
                return None

        # ODS para outros formatos
        case 10:
            try:
                nome_csv = arquivo.replace('.ods', '.csv')
                df.to_csv(nome_csv, index=False)
                print(f"Arquivo {nome_csv} convertido com sucesso!")
                return nome_csv
            except Exception as e:
                print(f"Erro ao converter para .csv: {e}")
                return None
        case 11:
            try:    
                nome_xlsx = arquivo.replace('.ods', '.xlsx')
                df.to_excel(nome_xlsx, index=False, engine='openpyxl')
                print(f"Arquivo {nome_xlsx} convertido com sucesso!")
                return nome_xlsx
            except Exception as e:
                print(f"Erro ao converter para .xlsx: {e}")
                return None
        case 12:
            try:
                nome_xls = arquivo.replace('.ods', '.xls')
                df.to_excel(nome_xls, index=False, engine='xlwt')
                print(f"Arquivo {nome_xls} convertido com sucesso!")
                return nome_xls
            except Exception as e:
                print(f"Erro ao converter para .xls: {e}")
                return None

        # opção inválida
        case _:
            print("opcao invalida")
            return None



    #função para receber o caminho convertido e abrir o arquivo
    