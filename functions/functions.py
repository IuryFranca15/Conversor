import pandas as pd
import os
from tkinter import filedialog, messagebox

def selecionar_arquivo(arquivo_path, status):
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo",
        filetypes=[("Arquivos Suportados", "*.csv *.xlsx *.xls *.ods")]
    )
    if caminho:
        arquivo_path.set(caminho)
        status.config(text=f"Arquivo carregado: {os.path.basename(caminho)}")


def converter_arquivo(caminho, formato, status):
    if not caminho:
        messagebox.showerror("Erro", "Nenhum arquivo selecionado!")
        return

    formato = formato.lower()

    try:
        if caminho.endswith('.csv'):
            df = pd.read_csv(caminho)
        else:
            df = pd.read_excel(caminho)

        output_path = caminho.rsplit('.', 1)[0] + f".{formato}"

        if formato == "csv":
            df.to_csv(output_path, index=False)
        else:
            df.to_excel(output_path, index=False, engine={
                "xlsx": "openpyxl",
                "xls": "xlwt",
                "ods": "odf"
            }[formato])

        messagebox.showinfo("Sucesso", f"Arquivo convertido: {output_path}")
        status.config(text=f"Convertido para {formato.upper()}!", foreground="green")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha na conversão:\n{str(e)}")
        status.config(text="Erro na conversão.", foreground="red")
