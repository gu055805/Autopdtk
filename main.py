import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def process_inventory():
    try:
        # Selecionar arquivo Excel
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        # Ler arquivo
        df = pd.read_excel(file_path)

        # Corrigir valores negativos no estoque
        if "Estoque Atual" in df.columns:
            df["Estoque Atual"] = df["Estoque Atual"].apply(lambda x: max(x, 0))

        # Identificar produtos com estoque crítico
        df["Status"] = "OK"
        if "Estoque Mínimo" in df.columns:
            df.loc[df["Estoque Atual"] < df["Estoque Mínimo"], "Status"] = "⚠️ ESTOQUE BAIXO"
        if "Estoque Máximo" in df.columns:
            df.loc[df["Estoque Atual"] > df["Estoque Máximo"], "Status"] = "⚠️ ESTOQUE ALTO"

        # Salvar novo arquivo na pasta "Downloads"
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        output_path = os.path.join(downloads_path, "relatorio_estoque.xlsx")
        df.to_excel(output_path, index=False)

        messagebox.showinfo("Sucesso", f"Relatório salvo em:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

# Criar interface gráfica
root = tk.Tk()
root.title("Automação de Estoques")
root.geometry("400x200")

btn_process = tk.Button(root, text="Selecionar Arquivo e Processar", command=process_inventory, padx=10, pady=5)
btn_process.pack(pady=20)

btn_exit = tk.Button(root, text="Sair", command=root.quit, padx=10, pady=5)
btn_exit.pack()

root.mainloop()
