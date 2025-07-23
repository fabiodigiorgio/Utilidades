
import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from io import BytesIO
from datetime import datetime
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt

# =============================================================
# CONFIGURAÇÃO DA PÁGINA
# =============================================================
st.set_page_config(layout="wide", page_title="Atendimentos - Cards")
st.title("📋 Visualização de Atendimentos")

# =============================================================
# CARREGAMENTO DA PLANILHA
# =============================================================
@st.cache_data(ttl=300, show_spinner="Carregando dados do Google Sheets...")
def carregar_planilha_google():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    cred_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(cred_dict, scope)
    client = gspread.authorize(creds)
    planilha = client.open_by_key("1pdw-vQlg8G69aOXx0ObqO49BGWzK85tpXyJ9gvD6m0Q")
    aba = planilha.worksheet("AGENDAMENTO DOMICILIAR")
    dados = aba.get_all_values()
    if not dados:
        return pd.DataFrame()
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

# =============================================================
# PREPARAÇÃO DO DATAFRAME
# =============================================================
def preparar_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    if "DATA" not in df_raw.columns:
        st.error("A coluna 'DATA' não foi encontrada na planilha.")
        st.write("Colunas encontradas:", df_raw.columns.tolist())
        st.stop()

    colunas_necessarias = [
        "DATA",
        "ORDEM DE SERVIÇO",
        "Fabricante",
        "Produto",
        "Defeito Relatado",
        "Nome Completo",
        "Whatsapp/Celular",
        "Endereço",
        "Número",
        "Bairro/Cidade",
        "CEP",
        "Complemento"
    ]

    faltando = [c for c in colunas_necessarias if c not in df_raw.columns]
    if faltando:
        st.warning(f"As seguintes colunas não foram encontradas: {faltando}")
        colunas_necessarias = [c for c in colunas_necessarias if c in df_raw.columns]

    df = df_raw[colunas_necessarias].copy()
    df.columns = [
        "DATA",
        "OS",
        "Fabricante",
        "Produto",
        "Defeito",
        "Nome do Cliente",
        "Contato",
        "Endereço",
        "Número",
        "Bairro",
        "CEP",
        "Complemento"
    ]
    df["DATA"] = pd.to_datetime(df["DATA"], dayfirst=True, errors="coerce")
    return df

# =============================================================
# EXPORTAÇÃO XLSX (usando openpyxl)
# =============================================================
def exportar_xlsx_openpyxl(df_export, titulo):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Atendimentos")
        worksheet = writer.sheets["Atendimentos"]
        worksheet["A1"] = titulo
    buffer.seek(0)
    return buffer

# =============================================================
# EXPORTAÇÃO PDF
# =============================================================
def exportar_pdf_tabular(df_export, titulo):
    buffer = BytesIO()
    with PdfPages(buffer) as pdf:
        fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4
        ax.axis("off")

        ax.set_title(titulo, fontsize=14, weight="bold", pad=20)
        df_display = df_export.copy()
        if "DATA" in df_display.columns:
            df_display["DATA"] = df_display["DATA"].dt.strftime("%d/%m/%Y")

        table = ax.table(
            cellText=df_display.values,
            colLabels=df_display.columns,
            cellLoc="left",
            loc="center"
        )
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        table.scale(1, 1.2)
        pdf.savefig(fig)
        plt.close(fig)
    buffer.seek(0)
    return buffer

# =============================================================
# MAIN
# =============================================================
try:
    df_raw = carregar_planilha_google()
except Exception as e:
    st.error(f"Erro ao carregar dados: {e}")
    st.stop()

if df_raw.empty:
    st.warning("Planilha vazia ou não lida.")
    st.stop()

df = preparar_dataframe(df_raw)

datas_unicas = sorted(df["DATA"].dropna().dt.normalize().unique())
datas_formatadas = [pd.to_datetime(d).strftime("%d/%m/%Y") for d in datas_unicas]
datas_selecionadas = st.multiselect("Selecione as datas:", datas_formatadas)
df_filtrado = df[df["DATA"].dt.strftime("%d/%m/%Y").isin(datas_selecionadas)] if datas_selecionadas else df.copy()

termo_busca = st.text_input("Buscar por Nome ou OS:")
if termo_busca:
    df_filtrado = df_filtrado[df_filtrado.apply(lambda row: termo_busca.lower() in str(row).lower(), axis=1)]

st.markdown(f"<h2 style='color: #2e86c1;'>Total de Atendimentos: {len(df_filtrado)}</h2>", unsafe_allow_html=True)

st.markdown("---")
st.subheader("Visualização Tabular")

df_tabular = df_filtrado[["DATA", "OS", "Nome do Cliente", "Produto", "Fabricante", "Defeito"]].copy()
st.dataframe(df_tabular)

titulo_export = "Relatório de Atendimentos"
if datas_selecionadas:
    titulo_export = f"Atendimentos do(s) dia(s): {', '.join(datas_selecionadas)}"

col1, col2 = st.columns(2)
with col1:
    xlsx_file = exportar_xlsx_openpyxl(df_tabular, titulo_export)
    st.download_button(
        label="📥 Baixar XLSX",
        data=xlsx_file,
        file_name="atendimentos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col2:
    pdf_file = exportar_pdf_tabular(df_tabular, titulo_export)
    st.download_button(
        label="📄 Baixar PDF Tabulado",
        data=pdf_file,
        file_name="atendimentos.pdf",
        mime="application/pdf"
    )
