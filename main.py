import tkinter as tk
from tkinter import filedialog
import excelSimplific
import threading
from PIL import Image, ImageTk
"""
Projeto: ExcelSimplifc
Autor: Marcos Vinicius Tanganelli Lara
Data de Início: 24/09/2023
Última Atualização: 28/09/2023

Descrição:
O programa tem como objetivo, converter dois arquivos exel em um organizado, de acordo com as orientações 
do meu supervisor. O programa atingi dois arquivos excel com tipos de formatações especificas, caso  
haja mudança na formatação, deve ser mudado o codigo também
"""

def buscar_arquivo(entry):
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo:
        entry.delete(0, tk.END)  # Limpar o campo de endereço
        entry.insert(0, arquivo)   # Inserir o endereço do arquivo no campo

def criar_excel(entry1, entry2, loading_label):
    endereco1 = entry1.get()
    endereco2 = entry2.get()
    nome_arquivo = filedialog.asksaveasfilename(defaultextension=".xls", initialfile="planilha.xls", filetypes=[("Arquivos Excel", "*.xls")])
    
    def excel_thread():
        if endereco1 and endereco2 and nome_arquivo:
            # Chame a função do seu módulo para criar o arquivo Excel
            # Certifique-se de ter a função excel() no seu módulo excelSimplific
            excelSimplific.excel(endereco1, endereco2, nome_arquivo)
            loading_label.config(text="Concluído!")
        else:
            loading_label.config(text="Erro: Certifique-se de que todos os campos estejam preenchidos e o local de salvamento selecionado.")

    # Inicie a thread para executar a função excel() em segundo plano
    excel_thread = threading.Thread(target=excel_thread)
    excel_thread.start()

    # Atualize a janela de carregamento enquanto a thread está rodando
    loading_label.config(text="Carregando...")

# Criar a janela principal
janela = tk.Tk()
janela.title("ExcelSimplific")
janela.geometry("600x400")

# Campo de entrada para o primeiro endereço (Composições)
endereco_entry1 = tk.Entry(janela)
endereco_entry1.pack()

# Botão para buscar endereço no primeiro campo
buscar_button1 = tk.Button(janela, text="Buscar Endereço Composições", command=lambda: buscar_arquivo(endereco_entry1))
buscar_button1.pack()

# Campo de entrada para o segundo endereço (INSUMOS)
endereco_entry2 = tk.Entry(janela)
endereco_entry2.pack()

# Botão para buscar endereço no segundo campo
buscar_button2 = tk.Button(janela, text="Buscar Endereço INSUMOS", command=lambda: buscar_arquivo(endereco_entry2))
buscar_button2.pack()

# Janela de carregamento
loading_label = tk.Label(janela, text="")
loading_label.pack()

# Botão para criar o Excel com os endereços
criar_excel_button = tk.Button(janela, text="Criar Excel", command=lambda: criar_excel(endereco_entry1, endereco_entry2, loading_label))
criar_excel_button.pack()

janela.mainloop()
