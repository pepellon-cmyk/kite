import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# Configuração da página
st.set_page_config(page_title="Gestão de Avaliação 2025.2", layout="wide")

ficheiro = "kite f lifeavaliacao_de_desempenho_-_2025.2.py.xlsm"

def gravar_no_excel(folha, novos_dados):
    # Carrega o workbook mantendo as macros (.xlsm)
    wb = load_workbook(ficheiro, keep_vba=True)
    ws = wb[folha]
    
    # Adiciona os dados na próxima linha vazia
    ws.append(novos_dados)
    wb.save(ficheiro)
    st.success(f"Dados gravados com sucesso na {folha}!")

st.title("📝 Sistema de Avaliação Responsivo")

# Menu de navegação para as diferentes folhas detetadas no ficheiro
aba_selecionada = st.sidebar.selectbox("Selecionar Área", ["sheet1", "sheet2", "sheet3", "sheet4"])

# Interface de Input
st.subheader(f"Nova Entrada: {aba_selecionada}")
col1, col2 = st.columns(2)

with col1:
    nome = st.text_input("Nome do Colaborador")
    data = st.date_input("Data da Avaliação")

with col2:
    nota = st.number_input("Nota Final (0-10)", min_value=0.0, max_value=10.0, step=0.5)
    comentarios = st.text_area("Observações Adicionais")

if st.button("Gravar na Planilha"):
    # Organiza os dados conforme a estrutura das suas sheets
    linha_para_gravar = [nome, data, nota, comentarios]
    gravar_no_excel(aba_selecionada, linha_para_gravar)

st.divider()

# Visualização dos dados existentes (Gráfico Responsivo)
st.subheader("📊 Visualização de Desempenho Atual")
try:
    df = pd.read_excel(ficheiro, sheet_name=aba_selecionada)
    st.dataframe(df, use_container_width=True)
except Exception as e:
    st.info("Aguardando dados para exibição.")