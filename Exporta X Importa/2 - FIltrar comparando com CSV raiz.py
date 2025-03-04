import os
import pandas as pd
import csv
import tkinter as tk
from tkinter import filedialog

# Função para selecionar o arquivo T552
def selecionar_arquivo_base():
    root = tk.Tk()
    root.withdraw()
    arquivo_selecionado = filedialog.askopenfilename(title="Selecione o arquivo base (T552)", filetypes=[("CSV Files", "*.csv")])
    return arquivo_selecionado

# Função para selecionar a pasta com as planilhas
def selecionar_pasta_planilhas():
    root = tk.Tk()
    root.withdraw()
    pasta_selecionada = filedialog.askdirectory(title="Selecione a pasta contendo as planilhas para comparação")
    return pasta_selecionada

# Selecionar o arquivo base e a pasta de planilhas
caminho_base = selecionar_arquivo_base()
pasta_entrada = selecionar_pasta_planilhas()

# Carregar a T552, tratando todos os campos como texto
planilha_base = pd.read_csv(caminho_base, encoding='ISO-8859-1', dtype=str, keep_default_na=False)
planilha_base = planilha_base.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)

# Ajustar mapeamento de colunas da T552 para as planilhas
mapa_colunas = {
    "Name": "Nome_Fantasia**",
    "DocumentNumber": "Num_Doc_Empresa**"
}
planilha_base = planilha_base.rename(columns=mapa_colunas)

# Pasta de saída
pasta_resultados = os.path.join(pasta_entrada, "Arquivos_Filtrados")
os.makedirs(pasta_resultados, exist_ok=True)

# Índice de filtros
filtros = {
    "empresas.csv": ["Nome_Fantasia**", "Num_Doc_Empresa**"],
    "ambientes.csv": ["Nome_Empresa**", "Num_Doc_Empresa**"],
    "cargos.csv": ["Num_Doc_Empresa**"],
    "epcs.csv": None,
    "epis.csv": None,
    "exames.csv": None,
    "exames-cargos-regras.csv": None,
    "funcionarios.csv": ["Nome_Empresa**", "Num_Doc_Empresa**"],
    "pessoas.csv": None,  # Será gerada posteriormente
    "relacao-exames-cargos.csv": ["Nome_Empresa**", "Num_Doc_Empresa**"],
    "riscos.csv": None
}

# Dicionário para armazenar resultados intermediários
resultados_intermediarios = {}

# Filtrar e processar os arquivos
for arquivo, campos in filtros.items():
    caminho_arquivo = os.path.join(pasta_entrada, arquivo)

    # Verificar se o arquivo existe
    if not os.path.exists(caminho_arquivo):
        print(f"Arquivo {arquivo} não encontrado.")
        continue

    # Carregar a planilha atual sem converter valores
    df = pd.read_csv(caminho_arquivo, encoding='ISO-8859-1', sep=';', dtype=str, keep_default_na=False)
    df = df.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)

    # Aplicar filtro se necessário
    if campos:
        campos_comuns = {campo: campo for campo in campos if campo in df.columns and campo in planilha_base.columns}
        if campos_comuns:
            base_filtrada = planilha_base[list(campos_comuns.values())]
            df_filtrado = df[df[list(campos_comuns.keys())].isin(base_filtrada.to_dict(orient='list')).all(axis=1)]
        else:
            df_filtrado = pd.DataFrame(columns=df.columns)  # Nenhum campo em comum
    else:
        df_filtrado = df  # Manter igual a original

    # Salvar resultado intermediário
    resultados_intermediarios[arquivo] = df_filtrado

    # Salvar o resultado filtrado
    caminho_saida = os.path.join(pasta_resultados, arquivo)
    df_filtrado.to_csv(caminho_saida, index=False, sep=';', encoding='ISO-8859-1', quoting=csv.QUOTE_ALL)
    print(f"Arquivo processado e salvo: {caminho_saida}")

# Gerar a planilha `pessoas.csv` com base na filtragem de `funcionarios.csv`
if "funcionarios.csv" in resultados_intermediarios:
    funcionarios = resultados_intermediarios["funcionarios.csv"]
    caminho_pessoas = os.path.join(pasta_entrada, "pessoas.csv")

    if os.path.exists(caminho_pessoas):
        pessoas = pd.read_csv(caminho_pessoas, encoding='ISO-8859-1', sep=';', dtype=str, keep_default_na=False)
        pessoas = pessoas.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)

        if "CPF_Funcionario**" in funcionarios.columns and "CPF**" in pessoas.columns:
            pessoas_filtradas = pessoas[pessoas["CPF**"].isin(funcionarios["CPF_Funcionario**"])]

            # Salvar a planilha gerada
            caminho_saida_pessoas = os.path.join(pasta_resultados, "pessoas.csv")
            pessoas_filtradas.to_csv(caminho_saida_pessoas, index=False, sep=';', encoding='ISO-8859-1', quoting=csv.QUOTE_ALL)
            print(f"Planilha `pessoas.csv` gerada e salva: {caminho_saida_pessoas}")
        else:
            print("Colunas necessárias para filtrar `pessoas.csv` não encontradas.")
    else:
        print("Arquivo `pessoas.csv` não encontrado.")
