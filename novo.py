import pandas as pd
import numpy as np
import openpyxl
import os
import time
import streamlit as st

# Incorporate data
caminho = r'C:\Users\thiagoferreira\Desktop\incluir lojas giro_v3.xlsm'
df_franqueados = pd.read_excel(caminho, sheet_name="Planilha1")
df_base = pd.read_excel(caminho, sheet_name="BASE VENDA E ESTOQUE")
df_cores = pd.read_excel(caminho, sheet_name="Base Cores")
df_modelo = pd.read_excel(caminho, sheet_name="Base Produtos")

df_base_com_modelo = pd.merge(df_base, df_modelo[["Referencia", "Generico"]], on="Referencia", how="inner")
df_base_com_modelo.rename(columns={"Generico": "Generico"}, inplace=True)

# Referência a ser analisada
modelo_generico = "SANDALIA BAIXA"
referencia = 40004
cor = "BISCUIT"
filiais = df_franqueados['Lojas'].unique()

# Inicializar a coluna de vendas, se não existir
if "Venda" not in df_franqueados.columns:
    df_franqueados["Venda"] = 0

# Filtrar os dados da referência para evitar somas incorretas
df_referencia = df_base[(df_base["Referencia"] == referencia) & (df_base["Cor"] == cor) & (df_base["Modelo"] == modelo_generico)]

start_time = time.time()
# Verificar e calcular vendas específicas para cada filial
for filial in filiais:
    vendas_filial = df_referencia[df_referencia["FIL_RAZAO_SOCIAL"] == filial ]["Qtd_Venda_Liquida"].sum()
    estoques_filial = df_referencia[df_referencia["FIL_RAZAO_SOCIAL"] == filial ]["Qtd_Estoque"].sum()
    recebimentos_filial = vendas_filial + estoques_filial
    giros_filial = vendas_filial / recebimentos_filial
    df_franqueados.loc[df_franqueados['Lojas'] == filial, "Venda"] = vendas_filial
    df_franqueados.loc[df_franqueados['Lojas'] == filial, "Estoque"] = estoques_filial
    df_franqueados.loc[df_franqueados['Lojas'] == filial, "Receb."] = recebimentos_filial
    df_franqueados.loc[df_franqueados['Lojas'] == filial, "Giro"] = giros_filial

df_franqueados["Giro"] = df_franqueados["Giro"].fillna(0)

df_franqueados.to_excel(r'C:\Users\thiagoferreira\Desktop\df_franqueados.xlsx',index=False)

end_time = time.time()
tempo_execucao = end_time - start_time
print(f"Tempo de execução: {tempo_execucao:.4f} segundos")

# montando streamlit

st.write("Giro - Verão 25:")
st.dataframe(df_franqueados,hide_index=True)





