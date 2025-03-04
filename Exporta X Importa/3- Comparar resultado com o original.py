import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import csv


def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo base_cliente.csv",
        filetypes=[("CSV files", "*.csv")]
    )
    return file_path


def select_folder():
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Selecione a pasta com os arquivos CSV")
    return folder_path


# Selecionar o arquivo base
base_path = select_file()
# Carregar o arquivo base usando vírgula como delimitador, garantindo que todos os dados sejam strings
base_df = pd.read_csv(base_path, encoding='ISO-8859-1', delimiter=',', dtype=str)

# Renomear a coluna caso haja erro de codificação ou inconsistência
base_df.rename(columns={'Documento NÃºmero': 'DocumentNumber'}, inplace=True)

# Criar um conjunto de documentos válidos para referência
documentos_validos = set(base_df['DocumentNumber'])

# Criar um dataframe de contagem com a estrutura do arquivo base
resultado_contagem = base_df.copy()
resultado_contagem['Ambientes'] = 0
resultado_contagem['Cargos'] = 0
resultado_contagem['Funcionarios'] = 0

# Selecionar a pasta com arquivos CSV
pasta_csv = select_folder()

# Definir o mapeamento de arquivos e colunas para buscar
arquivos_colunas = {
    'ambientes.csv': 'Num_Doc_Empresa**',
    'cargos.csv': 'Num_Doc_Empresa**',
    'funcionarios.csv': 'Num_Doc_Empresa**'
}

# Iterar sobre cada arquivo e coluna mapeados para contar ocorrências
for arquivo, coluna in arquivos_colunas.items():
    caminho_arquivo = os.path.join(pasta_csv, arquivo)
    if os.path.exists(caminho_arquivo):
        # Carregar o arquivo CSV atual como string para manter zeros à esquerda
        df = pd.read_csv(caminho_arquivo, encoding='ISO-8859-1', delimiter=';', dtype=str)

        # Contar quantas vezes cada 'DocumentNumber' aparece
        contagem = df[coluna].value_counts()

        # Atualizar a coluna de contagem no dataframe de resultado
        for doc in documentos_validos:
            if doc in contagem:
                if 'ambientes' in arquivo:
                    resultado_contagem.loc[resultado_contagem['DocumentNumber'] == doc, 'Ambientes'] += int(
                        contagem[doc])
                elif 'cargos' in arquivo:
                    resultado_contagem.loc[resultado_contagem['DocumentNumber'] == doc, 'Cargos'] += int(contagem[doc])
                elif 'funcionarios' in arquivo:
                    resultado_contagem.loc[resultado_contagem['DocumentNumber'] == doc, 'Funcionarios'] += int(
                        contagem[doc])
    else:
        print(f'Arquivo {arquivo} não encontrado na pasta selecionada.')

# Garantir que o arquivo de saída tenha a mesma estrutura do arquivo de origem
resultado_path = os.path.join(os.path.dirname(base_path), 'resultado_comparacao.csv')
resultado_contagem[base_df.columns].to_csv(
    resultado_path,
    index=False,
    sep=',',
    encoding='ISO-8859-1',
    quoting=csv.QUOTE_ALL
)

print(f'Resultado da comparação salvo em: {resultado_path}')
