import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pyexcel_ods3 import save_data
import traceback
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'Backend')))
import conversor

import sys

if getattr(sys, 'frozen', False):
    caminho_icon = os.path.join(sys._MEIPASS, "conversor.ico")
else:
    caminho_icon = os.path.join(os.path.dirname(__file__), "conversor.ico")

def registrar_erro_log(exception_obj):
    with open("log_erros.txt", "a", encoding="utf-8") as f:
        f.write(traceback.format_exc())
        f.write("\n\n")

class ConversorApp:
    def __init__(self, root):
        self.root = root
        self.root.iconbitmap(caminho_icon)
        self.style = ttk.Style()
        self.style.theme_use('xpnative')
        self.root.title("Conversor de Planilhas")
        self.root.geometry("500x300")

        # Variáveis
        self.arquivo_path = tk.StringVar()
        self.formato_saida = tk.StringVar(value="CSV")

        self.criar_interface()

    def criar_interface(self):
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # Seletor de arquivo
        ttk.Label(frame, text="Arquivo de Entrada:").grid(row=0, column=0, sticky=tk.W, pady=30)
        ttk.Entry(frame, textvariable=self.arquivo_path, width=40).grid(row=0, column=1)
        ttk.Button(frame, text="Procurar", command=self.selecionar_arquivo).grid(row=0, column=2)

        # Combobox de formatos
        ttk.Label(frame, text="Converter para:").grid(row=1, column=0, sticky=tk.W, pady=30)
        formatos = ["CSV", "XLSX", "ODS"]
        ttk.Combobox(frame, textvariable=self.formato_saida, values=formatos, state="readonly").grid(row=1, column=1, sticky=tk.W)

        # Botão de conversão
        ttk.Button(frame, text="Converter", command=self.converter_arquivo).grid(row=2, column=1, pady=10)

        # Barra de status
        self.status = ttk.Label(frame, text="Aguardando arquivo.", foreground="black", font=10, padding=10)
        self.status.grid(row=3, column=0, columnspan=3)

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
        title="Selecione o arquivo",
        filetypes=[("Arquivos Suportados", "*.csv *.xlsx *.xls *.ods")]
    )
        if caminho:
            try:
                resultado = conversor.verifica_tamanho(caminho)
                if isinstance(resultado, str) or resultado is None:
                    mensagem = (
                        "Erro: Arquivo não reconhecido como planilha válida.\n"
                        "Certifique-se de enviar arquivos .xls, .xlsx ou .ods reais."
                    )
                    messagebox.showerror("Erro", mensagem)
                    self.status.config(text="Erro ao carregar arquivo.", foreground="red")
                    return
                else:
                    self.arquivo_path.set(caminho)
                    self.status.config(text=f"Arquivo carregado: {os.path.basename(caminho)}", foreground="black")

            except Exception as e:
                messagebox.showerror(
                    "Erro",
                    "Erro ao ler o arquivo. Certifique-se de que o arquivo é um .xls, .xlsx, .csv ou .ods válido."
                )
                registrar_erro_log(e)
                self.status.config(text="Erro ao carregar arquivo.", foreground="red")


    def converter_arquivo(self):
        caminho = self.arquivo_path.get()

        if not caminho:
            messagebox.showerror("Erro", "Nenhum arquivo selecionado!")
            return

        formato = self.formato_saida.get().lower()

        try:
            df = conversor.verifica_tamanho(caminho)
            if isinstance(df, str):
        # trata o caso e retorna uma string de erro
                mensagem_amigavel = (
                    "Erro: Arquivo não reconhecido como planilha válida.\n"
                    "Certifique-se de enviar arquivos .xls, .xlsx ou .ods reais."
                )
                messagebox.showerror("Erro", mensagem_amigavel)
                self.status.config(text="Erro ao carregar arquivo.", foreground="red")
                return
            # opções de conversão
            opcoes = {
                ("csv", "xlsx"): 1,
                ("csv", "ods"): 2,
                ("xls", "csv"): 3,
                ("xls", "xlsx"): 4,
                ("xls", "ods"): 5,
                ("xlsx", "csv"): 6,
                ("xlsx", "ods"): 7,
                ("ods", "csv"): 8,
                ("ods", "xlsx"): 9,
            }

            extensao_atual = os.path.splitext(caminho)[-1][1:].lower()
            chave = (extensao_atual, formato)

            if chave not in opcoes:
                messagebox.showerror("Erro", f"Conversão de {extensao_atual.upper()} para {formato.upper()} não suportada.")
                return

            opcao = opcoes[chave]
            
            nome_arquivo_convertido = conversor.conversor(df, caminho, opcao, save_data)
            if nome_arquivo_convertido:
                messagebox.showinfo("Sucesso", f"Arquivo convertido: {nome_arquivo_convertido}")
                self.status.config(text=f"Convertido para {formato.upper()}!", foreground="green")
            else:
                messagebox.showerror("Erro", "Falha na conversão.")
                self.status.config(text="Erro na conversão.", foreground="red")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha na conversão:\n{str(e)}")
            registrar_erro_log(e)
            self.status.config(text="Erro inesperado.", foreground="red")


if __name__ == "__main__":
    root = tk.Tk()
    app = ConversorApp(root)
    root.mainloop()
