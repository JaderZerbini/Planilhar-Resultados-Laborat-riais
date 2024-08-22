import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def selecionar_diretorio_entrada():
    diretorio = filedialog.askdirectory()
    entry_diretorio_entrada.delete(0, tk.END)
    entry_diretorio_entrada.insert(0, diretorio)

def selecionar_diretorio_saida():
    diretorio = filedialog.askdirectory()
    entry_diretorio_saida.delete(0, tk.END)
    entry_diretorio_saida.insert(0, diretorio)

def extrair_dados():
    diretorio_entrada = entry_diretorio_entrada.get()
    diretorio_saida = entry_diretorio_saida.get()

    # Lista para armazenar os dados extraídos de todas as colunas E
    dados_coluna_E = []
    nomes_arquivos = []
    extraiu_colunas_A_C_J = False

    ordem_analitos = [
        ['Benzeno', 'Tolueno', 'Etilbenzeno', 'Xilenos'],
        ["Óleos e Graxas", "Acenafteno", "Acenaftileno", "Fluoreno", "Fluoranteno", "Pireno", "Antraceno",
        "Benzo(a)antraceno", "Benzo(b)fluoranteno", "Benzo(k)fluoranteno", "Benzo(g,h,i)perileno",
        "Benzo(a)pireno", "Criseno", "Dibenzo(a,h)antraceno", "Fenantreno", "Indeno(1,2,3-cd)pireno", "Naftaleno"],
        ["Porcentagem de Sólidos", "Porcentagem de Umidade", "TPH Fracionado Fração Alifática (>C10-C12)",
        "TPH Fracionado Fração Alifática (>C12-C16)", "TPH Fracionado Fração Alifática (>C16-C21)",
        "TPH Fracionado Fração Alifática (>C21-C32)", "Fração Alifática Total", 
        "TPH Fracionado Fração Aromática (>C10-C12)", "TPH Fracionado Fração Aromática (>C12-C16)", 
        "TPH Fracionado Fração Aromática (>C16-C21)", "TPH Fracionado Fração Aromática (>C21-C32)", 
        "Fração Aromática Total", "TPH Alifático (>C06-C08)", "TPH Alifático (>C08-C10)", 
        "TPH Alifático (C06-C08)", "TPH Aromático (>C08-C10)", "Ftano", "Pristano", "TPH Total (C8-C40)", 
        "TPH - HPR - (Hidrocarbonetos Resolvidos de Petróleo)", "TPH - MCNR - (Mistura Complexa não Resolvida)"]
    ]

    for filename in os.listdir(diretorio_entrada):
        if filename.endswith('.xlsx'):
            filepath = os.path.join(diretorio_entrada, filename)
            df = pd.read_excel(filepath, engine='openpyxl')

            nome_arquivo = os.path.splitext(filename)[0]
            nomes_arquivos.append(nome_arquivo)

            coluna_A = df.iloc[[*range(19, 24), *range(26, 44), *range(45, 67)], 0]  # Coluna A
            coluna_C = df.iloc[[*range(19, 24), *range(26, 44), *range(45, 67)], 2]  # Coluna C
            coluna_E = df.iloc[[*range(19, 24), *range(26, 44), *range(45, 67)], 4]  # Coluna E
            coluna_J = df.iloc[[*range(19, 24), *range(26, 44), *range(45, 67)], 8]  # Coluna J

            dados_coluna_E.append(coluna_E)

            if not extraiu_colunas_A_C_J:
                df.rename(columns={'A': 'Analises'}, inplace=True)
                df.rename(columns={'B': 'Un'}, inplace=True)
                extraiu_colunas_A_C_J = True

    dados_extraidos = pd.DataFrame({
        'Analises': coluna_A,
        'Un': coluna_C,
        **{f'{nome_arquivo}': dados for nome_arquivo, dados in zip(nomes_arquivos, dados_coluna_E)},
        'J': coluna_J
    })

    dados_extraidos.rename(columns={'J': 'VMP - VOR CETESB - Nº 125/2021/E, de 09 de Dezembro de 2021 - Água Subterrânea'}, inplace=True)

    dados_extraidos = pd.concat([dados_extraidos[dados_extraidos['Analises'].isin(grupo)] for grupo in ordem_analitos])

    arquivo_saida = os.path.join(diretorio_saida, 'resultados.xlsx')
    dados_extraidos.to_excel(arquivo_saida, index=False)
    print("\nDados extraídos salvos em:", arquivo_saida)

    if messagebox.askyesno("Extrair mais dados?", "Deseja extrair mais dados?"):
        return
    else:
        root.destroy()

# Configuração da interface gráfica
root = tk.Tk()
root.title("Extrair Dados de Arquivos Excel")

frame_diretorio_entrada = tk.Frame(root)
frame_diretorio_entrada.pack(pady=10)

label_diretorio_entrada = tk.Label(frame_diretorio_entrada, text="Diretório de Entrada:")
label_diretorio_entrada.grid(row=0, column=0)

entry_diretorio_entrada = tk.Entry(frame_diretorio_entrada, width=50)
entry_diretorio_entrada.grid(row=0, column=1)

button_selecionar_entrada = tk.Button(frame_diretorio_entrada, text="Selecionar", command=selecionar_diretorio_entrada)
button_selecionar_entrada.grid(row=0, column=2)

frame_diretorio_saida = tk.Frame(root)
frame_diretorio_saida.pack(pady=10)

label_diretorio_saida = tk.Label(frame_diretorio_saida, text="Diretório de Saída:")
label_diretorio_saida.grid(row=0, column=0)

entry_diretorio_saida = tk.Entry(frame_diretorio_saida, width=50)
entry_diretorio_saida.grid(row=0, column=1)

button_selecionar_saida = tk.Button(frame_diretorio_saida, text="Selecionar", command=selecionar_diretorio_saida)
button_selecionar_saida.grid(row=0, column=2)

button_extrair = tk.Button(root, text="Extrair Dados", command=extrair_dados)
button_extrair.pack(pady=10)

root.mainloop()
