
import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from io import BytesIO
from datetime import datetime
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
import math

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
# FUNÇÕES DE VISUALIZAÇÃO DE CARDS
# =============================================================
CARD_STYLE_BASE = '''
    border:1px solid #ddd;
    border-radius:8px;
    padding:15px;
    margin-bottom:15px;
    background-color:#f9f9f9;
    font-size:14px;
    min-height: 300px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
'''

def html_card(card: dict) -> str:
    return f'''
    <div style="{CARD_STYLE_BASE}">
        <h4 style="color:#2e86c1; margin-bottom:5px;">{card["nome"]}</h4>
        <p><strong>Data:</strong> {card["data"]} &nbsp;|&nbsp; <strong>OS:</strong> {card["os"]}</p>
        <p><strong>Produto:</strong> {card["produto"]} &nbsp;|&nbsp; <strong>Fabricante:</strong> {card["fabricante"]}</p>
        <p style="color:#c0392b;"><strong>Defeito:</strong> {card["defeito"]}</p>
        <p><strong>Endereço:</strong> {card["endereco"]}, Nº {card["numero"]} - {card["bairro"]}</p>
        <p><strong>CEP:</strong> {card["cep"]} &nbsp;|&nbsp; <strong>Compl.:</strong> {card["complemento"]}</p>
        <p style="color:#2980b9;"><strong>Contato:</strong> {card["contato"]}</p>
    </div>
    '''

def exibir_card(card: dict, container):
    container.markdown(html_card(card), unsafe_allow_html=True)

# =============================================================
# EXPORTAÇÃO XLSX (openpyxl)
# =============================================================
def exportar_xlsx(df_export, titulo):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Atendimentos")
        sheet = writer.sheets["Atendimentos"]
        sheet["A1"] = titulo
    buffer.seek(0)
    return buffer

# =============================================================
# EXPORTAÇÃO PDF TABULAR
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

# =====================
# EXIBIÇÃO DOS CARDS
# =====================
st.subheader("Cards de Atendimento")
page_size = st.selectbox("Cards por página:", [6, 9, 12, 24], index=2)
total_pages = max(1, math.ceil(len(df_filtrado) / page_size))
if "page_current" not in st.session_state:
    st.session_state.page_current = 1

col1, col2, col3 = st.columns([1, 2, 1])
with col1:
    if st.button("⬅️ Anterior", disabled=st.session_state.page_current <= 1):
        st.session_state.page_current -= 1
with col2:
    st.write(f"Página {st.session_state.page_current} de {total_pages}")
with col3:
    if st.button("Próxima ➡️", disabled=st.session_state.page_current >= total_pages):
        st.session_state.page_current += 1

start_idx = (st.session_state.page_current - 1) * page_size
end_idx = start_idx + page_size
df_page = df_filtrado.iloc[start_idx:end_idx]

cards = []
cols = st.columns(3)
col_idx = 0
for _, row in df_page.iterrows():
    card = {
        "nome": row["Nome do Cliente"],
        "data": row["DATA"].strftime("%d/%m/%Y") if pd.notnull(row["DATA"]) else "",
        "os": row["OS"],
        "produto": row["Produto"],
        "fabricante": row["Fabricante"],
        "defeito": row["Defeito"],
        "endereco": row["Endereço"],
        "numero": row["Número"],
        "bairro": row["Bairro"],
        "cep": row["CEP"],
        "complemento": row["Complemento"],
        "contato": row["Contato"]
    }
    cards.append(card)
    exibir_card(card, cols[col_idx])
    col_idx = (col_idx + 1) % 3

# =====================
# TABELA E EXPORTAÇÕES
# =====================
st.markdown("---")
st.subheader("Visualização Tabular")
df_tabular = df_filtrado[["DATA", "OS", "Nome do Cliente", "Produto", "Fabricante", "Defeito"]].copy()
st.dataframe(df_tabular)

titulo_export = "Relatório de Atendimentos"
if datas_selecionadas:
    titulo_export = f"Atendimentos do(s) dia(s): {', '.join(datas_selecionadas)}"

col1, col2 = st.columns(2)
with col1:
    xlsx_file = exportar_xlsx(df_tabular, titulo_export)
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
