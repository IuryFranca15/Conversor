import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class ConversorApp:
    def __init__(self, root):
        self.root = root
        self.root.iconbitmap("icone.ico")
        self.style = ttk.Style()
        self.style.theme_use('xpnative')
        self.root.title("Conversor de planilhas")
        self.root.geometry("500x300")
        
        # Variáveis
        self.arquivo_path = tk.StringVar()
        self.formato_saida = tk.StringVar(value="csv")
        
        self.criar_interface()

    def criar_interface(self):
        # Frame principal
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. Seletor de Arquivo
        ttk.Label(frame, text="Arquivo de Entrada:").grid(row=0, column=0, sticky=tk.W, pady=30)
        ttk.Entry(frame, textvariable=self.arquivo_path, width=40).grid(row=0, column=1)
        ttk.Button(frame, text="Procurar", command=self.selecionar_arquivo).grid(row=0, column=2)
        
        # 2. Formato de Saída
        ttk.Label(frame, text="Converter para:").grid(row=1, column=0, sticky=tk.W, pady=30)
        formatos = ["CSV", "XLSX", "XLS", "ODS"]
        ttk.Combobox(frame, textvariable=self.formato_saida, values=formatos).grid(row=1, column=1, sticky=tk.W)
        
        # 3. Botão de Conversão
        ttk.Button(frame, text="Converter", command=self.converter_arquivo).grid(row=2, column=1, pady=10)
        
        # 4. Barra de Status
        self.status = ttk.Label(frame, text="Aguardando arquivo.", foreground="black", font=10, padding=10)
        self.status.grid(row=3, column=0, columnspan=3)

    def selecionar_arquivo(self):
        # Abre uma janela para o usuário escolher um arquivo suportado
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo",
            filetypes=[("Arquivos Suportados", "*.csv *.xlsx *.xls *.ods")]
        )
        # Se o usuário selecionar um arquivo...
        if caminho:
            # Atualiza a variável com o caminho selecionado
            self.arquivo_path.set(caminho)
            # Atualiza o texto da barra de status com o nome do arquivo
            self.status.config(text=f"Arquivo carregado: {os.path.basename(caminho)}")

    def converter_arquivo(self):
        # Recupera o caminho do arquivo selecionado
        caminho = self.arquivo_path.get()
        
        # Se nenhum arquivo foi selecionado, mostra uma mensagem de erro
        if not caminho:
            messagebox.showerror("Erro", "Nenhum arquivo selecionado!")
            return
        
        # Pega o formato de saída selecionado no combobox e converte para minúsculo
        formato = self.formato_saida.get().lower()
        
        try:
            # Lê o arquivo dependendo da sua extensão
            if caminho.endswith('.csv'):
                df = pd.read_csv(caminho)  # Leitura de CSV
            else:
                df = pd.read_excel(caminho)  # Leitura de Excel (.xlsx, .xls, .ods)
            
            # Cria o nome do arquivo de saída com a nova extensão
            output_path = caminho.rsplit('.', 1)[0] + f".{formato}"
            
            # Salva o novo arquivo no formato desejado
            if formato == "csv":
                df.to_csv(output_path, index=False)  # Salva como CSV
            else:
                # Salva como Excel, especificando o engine correto para cada tipo
                df.to_excel(output_path, index=False, engine={
                    "xlsx": "openpyxl",
                    "xls": "xlwt",
                    "ods": "odf"
                }[formato])
            
            # Exibe mensagem de sucesso e atualiza a barra de status
            messagebox.showinfo("Sucesso", f"Arquivo convertido: {output_path}")
            self.status.config(text=f"Convertido para {formato.upper()}!", foreground="green")
        
        except Exception as e:
            # Em caso de erro, mostra mensagem e atualiza a barra de status
            messagebox.showerror("Erro", f"Falha na conversão:\n{str(e)}")
            self.status.config(text="Erro na conversão.", foreground="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = ConversorApp(root)
    root.mainloop()
