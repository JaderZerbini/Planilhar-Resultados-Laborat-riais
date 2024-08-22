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

    dados_coluna_E = []
    nomes_arquivos = []
    datas_coleta = []
    extraiu_colunas_A_C_J = False

    ordem_analitos = [
        ["Benzeno", "Etilbenzeno", "Tolueno", "Xilenos"],
        ["Naftaleno", "Acenaftileno", "Acenafteno", "Fluoreno", "Fenantreno", "Antraceno", "Fluoranteno", "Pireno", "Benzo(a)antraceno",
        "Criseno", "Benzo(b)fluoranteno", "Benzo(k)fluoranteno", "Benzo(a)pireno", "Indeno(1,2,3-cd)pireno", "Dibenzo(a,h)antraceno", "Benzo(g,h,i)perileno", 
        "Estireno", "Ftano", "Pristano",
        "TPH C5 - C8 (alifática)", "TPH Fracionado Fração Alifática (C9- C18)", "TPH Fracionado Fração Alifática (C19- C32)",
        "TPH C6 - C8 (aromática)", "TPH Fracionado Fração Aromática (C9- C16)", "TPH Fracionado Fração Aromática (C17-C32)",
        "TPH HRP", "TPH MCNR", "TPH Total (C8-C40)"]
    ]

    ordem_analitos_completa = ordem_analitos[0] + ordem_analitos[1]

    for filename in os.listdir(diretorio_entrada):
        if filename.endswith('.xlsx'):
            filepath = os.path.join(diretorio_entrada, filename)
            df = pd.read_excel(filepath, engine='openpyxl')  # Usar engine openpyxl

            nome_arquivo = os.path.splitext(filename)[0]
            nomes_arquivos.append(nome_arquivo)

            data_coleta = df.iloc[7, 0]  # A9 é a 9ª linha, 0ª coluna
            datas_coleta.append(data_coleta)

            
            coluna_A = df.iloc[[*range(19, 53), *range(106)], 0]  # Coluna A
            coluna_C = df.iloc[[*range(19, 53), *range(106)], 2]  # Coluna C
            coluna_E = df.iloc[[*range(19, 53), *range(106)], 4]  # Coluna E
            coluna_J = df.iloc[[*range(19, 53), *range(106)], 10]  # Coluna J
            
            dados_coluna_E.append(coluna_E)

            if not extraiu_colunas_A_C_J:
                df.rename(columns={'A': 'Analises'}, inplace=True)
                df.rename(columns={'B': 'Un'}, inplace=True)
                extraiu_colunas_A_C_J = True

    dados_extraidos = pd.DataFrame({
        'Analises': coluna_A,
        'Un': coluna_C,
        **{f'{nome_arquivo} - {data_coleta}': dados for nome_arquivo, data_coleta, dados in zip(nomes_arquivos, datas_coleta, dados_coluna_E)},
        'J': coluna_J
    })

    dados_extraidos.rename(columns={'J': 'VMP - VOR CETESB - Nº 125/2021/E, de 09 de Dezembro de 2021 - Água Subterrânea'}, inplace=True)

    dados_extraidos = pd.concat([dados_extraidos[dados_extraidos['Analises'].isin(grupo)] for grupo in ordem_analitos])

    dados_extraidos = dados_extraidos.set_index('Analises').reindex(ordem_analitos_completa).reset_index()

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
