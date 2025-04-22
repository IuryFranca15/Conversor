import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pyexcel_ods3 import save_data
import os
import conversor  # importa suas funções do conversor.py

class ConversorApp:
    def __init__(self, root):
        self.root = root
        self.root.iconbitmap("icone.ico")
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
        formatos = ["CSV", "XLSX", "XLS", "ODS"]
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
            # Chama a função do backend para verificar o arquivo
            resultado = conversor.verifica_tamanho(caminho)
            if isinstance(resultado, str):
                messagebox.showerror("Erro", resultado)
                self.status.config(text="Falha ao carregar arquivo.", foreground="red")
                self.arquivo_path.set("")
            else:
                self.arquivo_path.set(caminho)
                self.status.config(text=f"Arquivo carregado: {os.path.basename(caminho)}", foreground="black")

    def converter_arquivo(self):
        caminho = self.arquivo_path.get()

        if not caminho:
            messagebox.showerror("Erro", "Nenhum arquivo selecionado!")
            return

        formato = self.formato_saida.get().lower()

        try:
            df = conversor.verifica_tamanho(caminho)
            if isinstance(df, str):
                messagebox.showerror("Erro", df)
                self.status.config(text="Erro ao carregar arquivo.", foreground="red")
                return

            # Seleciona a opção de conversão pelo índice
            opcoes = {
                ("csv", "xlsx"): 1,
                ("csv", "xls"): 2,
                ("csv", "ods"): 3,
                ("xls", "csv"): 4,
                ("xls", "xlsx"): 5,
                ("xls", "ods"): 6,
                ("xlsx", "csv"): 7,
                ("xlsx", "xls"): 8,
                ("xlsx", "ods"): 9,
                ("ods", "csv"): 10,
                ("ods", "xlsx"): 11,
                ("ods", "xls"): 12
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
            self.status.config(text="Erro inesperado.", foreground="red")


if __name__ == "__main__":
    root = tk.Tk()
    app = ConversorApp(root)
    root.mainloop()
