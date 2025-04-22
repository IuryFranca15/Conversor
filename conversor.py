import tkinter as tk
from tkinter import ttk
from functions.functions import selecionar_arquivo, converter_arquivo

class ConversorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de planilhas")
        self.root.geometry("500x300")

        self.style = ttk.Style()
        self.style.theme_use('xpnative')

        self.arquivo_path = tk.StringVar()
        self.formato_saida = tk.StringVar(value="csv")

        self.criar_interface()

    def criar_interface(self):
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Arquivo de Entrada:").grid(row=0, column=0, sticky=tk.W, pady=30)
        ttk.Entry(frame, textvariable=self.arquivo_path, width=40).grid(row=0, column=1)
        ttk.Button(frame, text="Procurar", command=self.selecionar_arquivo_wrapper).grid(row=0, column=2)

        ttk.Label(frame, text="Converter para:").grid(row=1, column=0, sticky=tk.W, pady=30)
        formatos = ["CSV", "XLSX", "XLS", "ODS"]
        ttk.Combobox(frame, textvariable=self.formato_saida, values=formatos).grid(row=1, column=1, sticky=tk.W)

        ttk.Button(frame, text="Converter", command=self.converter_arquivo_wrapper).grid(row=2, column=1, pady=10)

        self.status = ttk.Label(frame, text="Aguardando arquivo.", foreground="black", font=10, padding=10)
        self.status.grid(row=3, column=0, columnspan=3)

    def selecionar_arquivo_wrapper(self):
        selecionar_arquivo(self.arquivo_path, self.status)

    def converter_arquivo_wrapper(self):
        converter_arquivo(self.arquivo_path.get(), self.formato_saida.get(), self.status)


if __name__ == "__main__":
    root = tk.Tk()
    app = ConversorApp(root)
    root.mainloop()
