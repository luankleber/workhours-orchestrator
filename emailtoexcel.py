# Instalar dependências (terminal):
# pip install camelot-py[cv] pandas openpyxl

import camelot
import pandas as pd
import re

PDF = "C:\\Relatorios\\report.pdf"
SAIDA = "tabela_final_organizada.xlsx"

def limpar_texto(valor):
    if pd.isna(valor):
        return None
    v = re.sub(r"\s+", " ", str(valor)).strip()
    return v if v else None


def renomear_colunas_duplicadas(df):
    cols = df.columns
    new_cols = []
    contador = {}

    for col in cols:
        if col not in contador:
            contador[col] = 0
        else:
            contador[col] += 1

        if contador[col] == 0:
            new_cols.append(col)
        else:
            new_cols.append(f"{col}_{contador[col]}")

    df.columns = new_cols
    return df


def organizar_colunas(df):
    # Limpa texto usando map linha a linha
    df = df.map(limpar_texto)

    # Remove linhas completamente vazias
    df = df.dropna(how="all")

    # Reseta índice
    df = df.reset_index(drop=True)

    # Detecta cabeçalho
    primeira = df.iloc[0].astype(str).tolist()
    if any(re.search(r"[A-Za-z]", x) for x in primeira):
        df.columns = primeira
        df = df.drop(0).reset_index(drop=True)

    # Evita colunas duplicadas
    df = renomear_colunas_duplicadas(df)

    return df


def extrair_e_organizar(pdf_path, saida):
    print("Extraindo tabelas...")
    tabelas = camelot.read_pdf(pdf_path, pages="all")

    if len(tabelas) == 0:
        print("Nenhuma tabela encontrada.")
        return

    dfs_limpos = []

    for i, t in enumerate(tabelas):
        df = t.df
        df_limpo = organizar_colunas(df)
        dfs_limpos.append(df_limpo)

    print("Unificando...")

    # Aqui garantimos que todas as tabelas têm as mesmas colunas
    df_final = pd.concat(dfs_limpos, ignore_index=True, sort=False)

    # Limpa novamente
    df_final = df_final.map(limpar_texto)
    df_final = df_final.dropna(how="all")

    print(f"Salvando em: {saida}")
    df_final.to_excel(saida, index=False)

    print("Processo concluído, arquivo organizado criado com sucesso!")


extrair_e_organizar(PDF, SAIDA)