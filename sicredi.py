import streamlit as st
import pandas as pd
import pyodbc
import os
import io
from datetime import datetime, timedelta
from utils import fetch_system_data, normalize_name, col_idx_to_excel_col

st.set_page_config(page_title='Comparação de Vendas: Sicredi vs Sistema', layout='wide')

# Recuperar as credenciais dos segredos (configuradas via secrets do Streamlit Cloud ou no arquivo .streamlit/secrets.toml)
server   = st.secrets["mssql"]["server"]
database = st.secrets["mssql"]["database"]
username = st.secrets["mssql"]["username"]
password = st.secrets["mssql"]["password"]

if not all([server, database, username, password]):
    st.error("Verifique se todas as variáveis estão definidas corretamente no secrets.")
    st.stop()
else:
    st.write(f"Servidor: {server}")
    st.write(f"Banco de Dados: {database}")
    st.write(f"Usuário: {username}")

st.title('Comparação de Vendas: Sicredi vs Sistema')

#####################################
# Seção 1: Carregar planilha da Sicredi
#####################################
st.write('### Carregar planilha da Sicredi')
uploaded_file_sicredi = st.file_uploader('Faça o upload do arquivo Sicredi aqui', type=['xlsx'], key='sicredi')

#####################################
# Seção 2: Buscar dados do Sistema via SQL
#####################################
st.write('### Buscar dados do Sistema')
st.write("Informe os filtros para consulta:")

# Inputs de data
start_date = st.date_input("Data Inicial", value=datetime.now())
end_date = st.date_input("Data Final", value=datetime.now())

# Mapeamento de empresas para o selectbox
id_empresa_mapping = {
    58: 'Araguaína II',
    66: 'Balsas II',
    55: 'Araguaína I',
    53: 'Imperatriz II', 
    51: 'Imperatriz I',
    65: 'Araguaína IV', 
    52: 'Imperatriz III',
    57: 'Araguaína III',
    50: 'Balsas I', 
    56: 'Gurupi I',  
    61: 'Colinas',
    60: 'Estreito',
    46: 'Formosa I',
    59: 'Guaraí'
}

company_options = sorted(id_empresa_mapping.values())
selected_company = st.selectbox("Selecione a empresa", options=company_options)

# Botão para disparar a consulta do Sistema
if st.button("Buscar dados do Sistema"):
    try:
        df_sistema = fetch_system_data(start_date, end_date, selected_company, id_empresa_mapping, st.secrets)
        st.success("Dados do Sistema carregados com sucesso!")
        st.session_state.df_sistema = df_sistema
    except Exception as e:
        st.error(str(e))

#####################################
# Seção 3: Processamento e Consolidação
#####################################
if uploaded_file_sicredi is not None and "df_sistema" in st.session_state:
    # Leitura e tratamento da planilha Sicredi
    colunas_sicredi = [
        'Data da venda', 'Cód. de autorização', 'Produto', 'Parcelas', 'Bandeira', 
        'Canal', 'Valor bruto', 'Valor da taxa', 'Valor líquido', 'Valor cancelado', 
        'Status', 'Número do terminal', 'Comprovante da venda', 'Cód. do pedido', 
        'Número do estabelecimento', 'Nome do estabelecimento', 'Descrição do link', 
        'Número do cartão', 'Cód. Ref. Cartão'
    ]
    
    try:
        df_sicredi = pd.read_excel(uploaded_file_sicredi, skiprows=16, names=colunas_sicredi, header=None)
        st.success("Planilha Sicredi carregada com sucesso!")
        #st.write("Colunas lidas da planilha Sicredi:", df_sicredi.columns.tolist())
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Sicredi: {e}")
        st.stop()

    # Dados do Sistema obtidos na consulta SQL
    df_sistema = st.session_state.df_sistema.copy()

    # Remove espaços extras dos nomes das colunas
    df_sicredi.columns = df_sicredi.columns.str.strip()
    df_sistema.columns = df_sistema.columns.str.strip()

    # Seleção das colunas desejadas
    colunas_desejadas_sicredi = ['Data da venda', 'Produto', 'Canal', 'Bandeira', 'Valor bruto', 'Número do estabelecimento', 'Cód. do pedido']
    colunas_desejadas_sistema = ['ID EMPRESA', 'EMPRESA', 'ID VENDA', 'FORMA DE PAGAMENTO', 'NOME', 
                                 'ID CAIXA', 'NSU', 'VALOR BRUTO', 'DATA DE FATURAMENTO', 'EMISSAO']

    for col in colunas_desejadas_sicredi:
        if col not in df_sicredi.columns:
            st.error(f"A coluna '{col}' não foi encontrada na planilha Sicredi.")
            st.stop()

    for col in colunas_desejadas_sistema:
        if col not in df_sistema.columns:
            st.error(f"A coluna '{col}' não foi encontrada nos dados do Sistema.")
            st.stop()

    df_sicredi = df_sicredi[colunas_desejadas_sicredi]
    df_sistema = df_sistema[colunas_desejadas_sistema]

    #####################################
    # Mapeamento dos códigos dos estabelecimentos na Sicredi
    #####################################
    establishment_mapping = {
        "92185778": "Araguaína I",
        "92185790": "Araguaína II",
        "92185788": "Araguaína III",
        "92139112": "Araguaína IV",
        "92187397": "Imperatriz I",
        "92187444": "Imperatriz II",
        "92187446": "Imperatriz III",
        "92187441": "Balsas I",
        "92187436": "Balsas II",
        "92197344": "Estreito",
        "92185785": "Gurupi I",
        "92197340": "Formosa I",
        "92187423": "Guaraí",
        "92187439": "Colinas"
    }

    df_sicredi['Número do estabelecimento'] = df_sicredi['Número do estabelecimento'].astype(str).str.strip()
    df_sicredi['Número do estabelecimento'] = df_sicredi['Número do estabelecimento'].map(establishment_mapping)
    
    if df_sicredi['Número do estabelecimento'].isnull().any():
        unmapped_codes = df_sicredi[df_sicredi['Número do estabelecimento'].isnull()]['Número do estabelecimento'].unique()
        st.error(f"Existem códigos de estabelecimento sem mapeamento: {', '.join(unmapped_codes)}. Verifique o mapeamento fornecido.")
        st.stop()

    #####################################
    # Normalização dos nomes dos estabelecimentos
    #####################################
    df_sicredi['Número do estabelecimento'] = df_sicredi['Número do estabelecimento'].apply(normalize_name)
    df_sistema['EMPRESA'] = df_sistema['EMPRESA'].apply(normalize_name)

    #####################################
    # Conversão das datas
    #####################################
    try:
        df_sicredi['Data da venda'] = pd.to_datetime(df_sicredi['Data da venda'], dayfirst=True)
    except Exception as e:
        st.error(f"Erro ao converter 'Data da venda' para datetime: {e}")
        st.stop()

    try:
        df_sistema['DATA DE FATURAMENTO'] = pd.to_datetime(df_sistema['DATA DE FATURAMENTO'], dayfirst=True)
    except Exception as e:
        st.error(f"Erro ao converter 'DATA DE FATURAMENTO' para datetime: {e}")
        st.stop()

    df_sicredi['Data da venda'] = df_sicredi['Data da venda'].
