import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

st.set_page_config(layout="wide", page_title="Atendimentos - Cards")
st.title("📋 Visualização de Atendimentos")

def carregar_planilha_google():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    try:
        cred_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    except Exception as e:
        st.error("Credencial Google não encontrada nos segredos. Verifique o painel de secrets no Streamlit Cloud.")
        st.stop()

    creds = ServiceAccountCredentials.from_json_keyfile_dict(cred_dict, scope)
    client = gspread.authorize(creds)

    planilha = client.open_by_key("1pdw-vQlg8G69aOXx0ObqO49BGWzK85tpXyJ9gvD6m0Q")
    aba = planilha.worksheet("AGENDAMENTO DOMICILIAR")

    dados = aba.get_all_values()
    headers = dados[0]
    conteudo = dados[1:]

    seen = {}
    headers_corrigidos = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            headers_corrigidos.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            headers_corrigidos.append(h)

    df = pd.DataFrame(conteudo, columns=headers_corrigidos)
    return df

try:
    df_raw = carregar_planilha_google()
except Exception as e:
    st.error(f"Erro ao carregar dados do Google Sheets: {e}")
    st.stop()

coluna_data_entrada = df_raw.columns[17]

colunas_esperadas = [
    coluna_data_entrada, 'ORDEM DE SERVIÇO', 'Fabricante', 'Produto', 'Defeito Relatado',
    'Nome Completo', 'Whatsapp/Celular', 'Endereço', 'Número', 'Bairro/Cidade',
    'CEP', 'Complemento'
]

colunas_faltando = [col for col in colunas_esperadas if col not in df_raw.columns]
if colunas_faltando:
    st.error(f"As seguintes colunas esperadas não foram encontradas: {colunas_faltando}")
    st.stop()

df = df_raw[colunas_esperadas].copy()
df.columns = [
    'Data de Entrada', 'OS', 'Fabricante', 'Produto', 'Defeito',
    'Nome do Cliente', 'Contato', 'Endereço', 'Número', 'Bairro',
    'CEP', 'Complemento'
]

df['Data de Entrada'] = pd.to_datetime(df['Data de Entrada'], errors='coerce')
datas_unicas = sorted(df['Data de Entrada'].dropna().dt.date.unique())

data_filtro = st.selectbox("Selecione uma data:", datas_unicas)
df_filtrado = df[df['Data de Entrada'].dt.date == data_filtro]

cols = st.columns(2)

for idx, (_, row) in enumerate(df_filtrado.iterrows()):
    with cols[idx % 2]:
        card = {
            "nome": row['Nome do Cliente'],
            "data": row['Data de Entrada'].strftime('%d/%m/%Y') if pd.notnull(row['Data de Entrada']) else '',
            "os": row['OS'],
            "produto": row['Produto'],
            "fabricante": row['Fabricante'],
            "defeito": row['Defeito'],
            "endereco": row['Endereço'],
            "numero": row['Número'],
            "bairro": row['Bairro'],
            "cep": row['CEP'],
            "complemento": row['Complemento'],
            "contato": row['Contato']
        }

        fig, ax = plt.subplots(figsize=(4.5, 6))
        ax.axis('off')
        ax.add_patch(patches.Rectangle((0, 0), 1, 1, fill=True, color="#ffffff", transform=ax.transAxes))

        y = 0.92
        ax.text(0.02, y, f"{card['nome']}", fontsize=12, fontweight='bold', color="#333333", transform=ax.transAxes); y -= 0.06
        ax.text(0.02, y, f"Data: {card['data']}", fontsize=10, color="#006699", transform=ax.transAxes); y -= 0.05
        ax.text(0.02, y, f"OS: {card['os']}", fontsize=10, fontweight='bold', transform=ax.transAxes); y -= 0.05
        ax.text(0.02, y, f"Produto: {card['produto']}", fontsize=9, transform=ax.transAxes); y -= 0.045
        ax.text(0.02, y, f"Fabricante: {card['fabricante']}", fontsize=9, transform=ax.transAxes); y -= 0.045
        ax.text(0.02, y, f"Defeito: {card['defeito']}", fontsize=9, color="#cc0000", transform=ax.transAxes); y -= 0.055
        ax.text(0.02, y, f"Endereço: {card['endereco']}, Nº {card['numero']}", fontsize=9, transform=ax.transAxes); y -= 0.045
        ax.text(0.02, y, f"Bairro: {card['bairro']}", fontsize=9, transform=ax.transAxes); y -= 0.045
        ax.text(0.02, y, f"CEP: {card['cep']}  Compl: {card['complemento']}", fontsize=9, transform=ax.transAxes); y -= 0.045
        ax.text(0.02, y, f"Contato: {card['contato']}", fontsize=9, color="#336699", transform=ax.transAxes)

        st.pyplot(fig)
