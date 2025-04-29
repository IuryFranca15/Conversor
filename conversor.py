
import pandas as pd
import os
import ezodf

arquivo = []



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

def ler_xls(arquivo):
    try:
        #ler como XLS normal (binário)
        df = pd.read_excel(arquivo, engine='xlrd', dtype=str)
        return df

    except Exception as e1:
        print(f"Falha ao ler como XLS: {e1}. Tentando como XLSX...")

        try:
            #ler como XLSX mesmo que o arquivo tenha extensão XLS
            df = pd.read_excel(arquivo, engine='openpyxl', dtype=str)
            return df

        except Exception as e2:
            print(f"Falha ao ler como XLSX: {e2}")
            raise ValueError("Arquivo inválido para leitura e conversão.")



    
def ler_xlsx(arquivo, engine='openpyxl'):
    try:   
        df = pd.read_excel(arquivo, engine=engine)
        return df
    except Exception as e:
        print(f"Erro ao ler xlsx: {e}")
        return None

def ler_ods (arquivo: str):
    try:
        ezodf.config.set_table_expand_strategy('all')
        ods = ezodf.opendoc(arquivo)
        sheet = ods.sheets[0]

        data = []
        for row in sheet.rows():
            row_data = [cell.value for cell in row]
            data.append(row_data)

        import pandas as pd
        df = pd.DataFrame(data)
        return df
    except Exception as e:
        print(f"Erro ao ler ODS com ezodf: {e}")
        return None

def to_ods(df, nome_ods, save_data):
    
    try:
        if df is None:
            raise ValueError("Arquivo vazio, não é possível converter para ODS.")    
        df = df.astype(str)
        data = df.to_dict(orient='split')  # Converte para dicionário
        sheet_data = { "Sheet1": data['data'] }  # O pyexcel-ods3 precisa de uma lista de listas
        save_data(nome_ods, sheet_data)
    except Exception as e:
        print(f"Erro ao converter para ODS: {e}")
        raise e    
    
#conversão de arquivo
def conversor(df, arquivo, opcao, save_data):
    match opcao:
            case 1:
                try:
                    nome_excel = arquivo.replace('.csv', '.xlsx')
                    df.to_excel(nome_excel, index=False, engine='openpyxl')
                    return nome_excel
                except Exception as e:
                    print(f"Erro no case 1 (csv → xlsx): {e}")
                    raise e
                
            case 2:
                try: 
                    nome_ods = arquivo.replace('.csv', '.ods')
                    to_ods(df, nome_ods, save_data)
                    return nome_ods
                except Exception as e:
                    print(f"Erro no case 2 (csv → ods): {e}")
                    raise e
            case 3:
                try:
                    nome_csv = arquivo.replace('.xls', '.csv')
                    df.to_csv(nome_csv, index=False)
                    return nome_csv
                except Exception as e:
                    print(f"Erro no case 2 (xls → csv): {e}")
                    raise e
            case 4:
                try:
                    nome_excel = arquivo.replace('.xls', '.xlsx')
                    df.to_excel(nome_excel, index=False, engine='openpyxl')
                    return nome_excel
                except Exception as e:
                    print(f"Erro no case 2 (xls → xlsx): {e}")
                    raise e
            case 5:
                try:
                    nome_ods = arquivo.replace('.xls', '.ods')
                    to_ods(df, nome_ods, save_data)
                    return nome_ods
                except Exception as e:
                    print(f"Erro no case 2 (xls → ods): {e}")
                    raise e
            case 6:
                try:
                    nome_csv = arquivo.replace('.xlsx', '.csv')
                    df.to_csv(nome_csv, index=False)
                    return nome_csv
                except Exception as e:
                    print(f"Erro no case 2 (xlsx → csv): {e}")
                    raise e 
            case 7:
                try:
                    nome_ods = arquivo.replace('.xlsx', '.ods')
                    to_ods(df, nome_ods, save_data)
                    return nome_ods
                except Exception as e:
                    print(f"Erro no case 2 (xlsx → ods): {e}")
                    raise e
            case 8:
                try:
                    nome_csv = arquivo.replace('.ods', '.csv')
                    df.to_csv(nome_csv, index=False)
                    return nome_csv
                except Exception as e:
                    print(f"Erro no case 2 (ods → csv): {e}")
                    raise e 
            case 9:
                try:
                    nome_xlsx = arquivo.replace('.ods', '.xlsx')
                    df.to_excel(nome_xlsx, index=False, engine='openpyxl')
                    return nome_xlsx
                except Exception as e:
                    print(f"Erro no case 2 (ods → xlsx): {e}")
                    raise e
            case _:
                print("opcao invalida")
                return None

