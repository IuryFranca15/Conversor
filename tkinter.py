import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import conversor

class ConversorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de Planilhas")
        self.root.geometry("500x300")
        
        self.mapeamento = {
            ('csv', 'xlsx'): 1,
            ('csv', 'xls'): 2,
            ('csv', 'ods'): 3,
            ('xls', 'csv'): 4,
            ('xls', 'xlsx'): 5,
            ('xls', 'ods'): 6,
            ('xlsx', 'csv'): 7,
            ('xlsx', 'xls'): 8,
            ('xlsx', 'ods'): 9,
            ('ods', 'csv'): 10,
            ('ods', 'xlsx'): 11,
            ('ods', 'xls'): 12
        }

        self.arquivo_path = tk.StringVar()
        self.formato_saida = tk.StringVar(value="CSV")

        self.criar_interface()

    def criar_interface(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Seletor de arquivo
        ttk.Label(frame, text="Arquivo de Entrada:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.arquivo_path, width=40, state='readonly').grid(row=0, column=1)
        ttk.Button(frame, text="Procurar", command=self.selecionar_arquivo).grid(row=0, column=2, padx=10)

        # Seletor de formato
        ttk.Label(frame, text="Converter para:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.combo = ttk.Combobox(
            frame, 
            textvariable=self.formato_saida, 
            values=["CSV", "XLSX", "XLS", "ODS"],
            state="readonly"
        )
        self.combo.grid(row=1, column=1, sticky=tk.W)

        # Botão de conversão
        ttk.Button(frame, text="Converter", command=self.converter).grid(row=2, column=1, pady=20)

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(filetypes=[("Planilhas", "*.csv *.xls *.xlsx *.ods")])
        if caminho:
            self.arquivo_path.set(caminho)

    def obter_opcao(self):
        origem = self.arquivo_path.get().split('.')[-1].lower()
        destino = self.formato_saida.get().lower()
        return self.mapeamento.get((origem, destino), 0)

    def converter(self):
        arquivo = self.arquivo_path.get()
        if not arquivo:
            messagebox.showerror("Erro", "Selecione um arquivo!")
            return

        opcao = self.obter_opcao()
        if opcao == 0:
            messagebox.showerror("Erro", "Combinação inválida de formatos!")
            return

        resultado = conversor.conversor(arquivo, opcao)
        if "Erro" in resultado:
            messagebox.showerror("Erro", resultado)
        else:
            messagebox.showinfo("Sucesso", f"Arquivo convertido:\n{resultado}")
            self.arquivo_path.set('')

if __name__ == "__main__":
    root = tk.Tk()
    app = ConversorApp(root)
    root.mainloop()
