import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import csv

# Função para selecionar a pasta raiz
def selecionar_pasta_raiz():
    root = tk.Tk()
    root.withdraw()
    pasta_selecionada = filedialog.askdirectory(title="Selecione a pasta raiz contendo as subpastas")
    return pasta_selecionada

# Selecionar a pasta raiz
pasta_raiz = selecionar_pasta_raiz()

# Criar pasta "Arquivos_Combinados" na pasta raiz para armazenar os arquivos de saída
pasta_combinados = os.path.join(pasta_raiz, "Arquivos_Combinados")
os.makedirs(pasta_combinados, exist_ok=True)

# Lista de arquivos para combinar
arquivos_para_combinar = [
    'empresas.csv', 'ambientes.csv', 'cargos.csv', 'epcs.csv', 'epis.csv',
    'exames.csv', 'exames-cargos-regras.csv', 'funcionarios.csv', 'pessoas.csv',
    'relacao-exames-cargos.csv', 'riscos.csv'
]

# Combinar cada tipo de arquivo
for nome_arquivo in arquivos_para_combinar:
    dataframes = []

    # Percorrer todas as subpastas na pasta raiz
    for root, dirs, files in os.walk(pasta_raiz):
        for file in files:
            if file == nome_arquivo:
                # Carregar a planilha e adicioná-la à lista, lendo todos os campos como texto
                df = pd.read_csv(
                    os.path.join(root, file), teste
                    encoding='ISO-8859-1',
                    delimiter=';',
                    quotechar='"',
                    quoting=csv.QUOTE_ALL,
                    doublequote=True,
                    dtype=str  # Forçar todos os dados a serem tratados como texto
                )
                # Garantir que os dados sejam tratados como strings e manter zeros à esquerda
                df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
                dataframes.append(df)

    # Se encontrar pelo menos uma planilha com o nome especificado
    if dataframes:
        # Concatenar todas as planilhas em um único dataframe
        df_final = pd.concat(dataframes, ignore_index=True)

        # Caminho para o arquivo final de saída na nova pasta "Arquivos_Combinados"
        arquivo_saida = os.path.join(pasta_combinados, f"{nome_arquivo}")

        # Salvar o dataframe combinado no arquivo
        df_final.to_csv(
            arquivo_saida,
            index=False,
            sep=';',
            encoding='ISO-8859-1',
            quoting=csv.QUOTE_ALL,
            quotechar='"',
            doublequote=True,
            escapechar='\\'  # Garantir caracteres escapados
        )
        print(f"Arquivo gerado: {arquivo_saida}")
    else:
        print(f"Nenhuma planilha encontrada para {nome_arquivo}")
