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
    datas_coleta = []
    extraiu_colunas_A_C_J = False

    ordem_analitos = [
        ['Benzeno', 'Tolueno', 'Etilbenzeno', 'Xilenos'],
        ["Acenafteno", "Acenaftileno", "Antraceno", "Benzo(a)antraceno", "Benzo(a)pireno", "Benzo(b)fluoranteno", "Benzo(g,h,i)perileno",    
        "Benzo(k)fluoranteno", "Criseno", "Dibenzo(a,h)antraceno", "Estireno", "Fenantreno", "Fluoranteno", "Fluoreno", "Ftano",
        "Indeno(1,2,3-cd)pireno", "Naftaleno", "Óleos e Graxas", "Pireno", "Pristano"],    
        ["TPH >C8 - C10 (aromática)", "TPH C6 - C8 (aromática)", "TPH HRP", "TPH MCNR", "TPH Total (C8-C40)", "TPH C5 - C8 (alifática)",    
        "TPH Fracionado Fração Alifática (C19-C32)", "TPH Fracionado Fração Alifática (C9-C18)", "TPH Fracionado Fração Aromática (C10-C32)",    
        "TPH Fracionado Fração Aromática (C17-C32)", "TPH Fracionado Fração Aromática (C9-C16)"]
    ]

    for filename in os.listdir(diretorio_entrada):
        if filename.endswith('.xlsx'):
            filepath = os.path.join(diretorio_entrada, filename)
            df = pd.read_excel(filepath)

            nome_arquivo = os.path.splitext(filename)[0]

            # Extraindo a "Data de Coleta" da célula A9
            data_coleta = df.iloc[7, 0]  # Linha 9 (índice 8), coluna A (índice 0)
            nome_arquivo_com_data = f'{nome_arquivo} - {data_coleta}'
            nomes_arquivos.append(nome_arquivo_com_data)  # Adicionando o nome do arquivo com a data de coleta

            coluna_A = df.iloc[[*range(8, 0), *range(19, 49), *range(101, 108)], 0]  # Coluna A
            coluna_C = df.iloc[[*range(19, 49), *range(101, 108)], 2]  # Coluna C
            coluna_E = df.iloc[[*range(19, 49), *range(101, 108)], 4]  # Coluna E
            coluna_J = df.iloc[[*range(19, 49), *range(101, 108)], 9]  # Coluna J

            dados_coluna_E.append(coluna_E)

            if not extraiu_colunas_A_C_J:
                df.rename(columns={'A': 'Analises'}, inplace=True)
                df.rename(columns={'B': 'Un'}, inplace=True)
                extraiu_colunas_A_C_J = True

    # Preparando os dados para o DataFrame final
    dados_extraidos = {
        'Analises': coluna_A,
        'Un': coluna_C
    }

    for nome_arquivo, dados in zip(nomes_arquivos, dados_coluna_E):
        dados_extraidos[nome_arquivo] = dados

    dados_extraidos['J'] = coluna_J

    df_resultado = pd.DataFrame(dados_extraidos)

    df_resultado.rename(columns={'J': 'VMP - VOR CETESB - Nº 125/2021/E, de 09 de Dezembro de 2021 - Água Subterrânea'}, inplace=True)

    df_resultado = pd.concat([df_resultado[df_resultado['Analises'].isin(grupo)] for grupo in ordem_analitos])

    arquivo_saida = os.path.join(diretorio_saida, 'resultados.xlsx')
    df_resultado.to_excel(arquivo_saida, index=False)
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
