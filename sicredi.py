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
        st.error(f"Existem códigos de estabelecimento sem mapeamento: {', '.join(map(str, unmapped_codes))}. Verifique o mapeamento fornecido.")
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

    df_sicredi['Data da venda'] = df_sicredi['Data da venda'].dt.strftime('%d/%m/%Y')
    df_sistema['DATA DE FATURAMENTO'] = df_sistema['DATA DE FATURAMENTO'].dt.strftime('%d/%m/%Y')

    #####################################
    # Conversão dos valores para float
    #####################################
    try:
        df_sicredi['Valor bruto sicredi'] = df_sicredi['Valor bruto'].astype(float)
    except Exception as e:
        st.error(f"Erro ao converter 'Valor bruto' da Sicredi para float: {e}")
        st.stop()

    # Criação da nova coluna com os valores convertidos do Sistema
    try:
        df_sistema['Valor bruto sistema'] = df_sistema['VALOR BRUTO'].str.replace(',', '.').astype(float)
    except Exception as e:
        st.error(f"Erro ao converter 'VALOR BRUTO' do Sistema para float: {e}")
        st.stop()

    # Remove as colunas originais que não serão mais usadas
    df_sicredi.drop(columns=['Valor bruto'], inplace=True)
    df_sistema.drop(columns=['VALOR BRUTO'], inplace=True)

    df_sicredi.reset_index(drop=True, inplace=True)
    df_sistema.reset_index(drop=True, inplace=True)

    #####################################
    # Lógica de comparação entre as planilhas
    #####################################
    indices_utilizados_sicredi = []
    indices_utilizados_sistema = []
    resultados = []

    for i, row_sicredi in df_sicredi.iterrows():
        # Primeira tentativa: data exata
        correspondencia = df_sistema[
            (df_sistema['EMPRESA'] == row_sicredi['Número do estabelecimento']) &
            (df_sistema['DATA DE FATURAMENTO'] == row_sicredi['Data da venda']) &
            (df_sistema['Valor bruto sistema'] == row_sicredi['Valor bruto sicredi']) &
            (~df_sistema.index.isin(indices_utilizados_sistema))
        ].head(1)

        if not correspondencia.empty:
            resultados.append((row_sicredi, correspondencia.iloc[0], 'Correspondido (Data Exata)'))
            indices_utilizados_sistema.append(correspondencia.index[0])
            indices_utilizados_sicredi.append(i)
        else:
            # Segunda tentativa: D+1
            data_d1 = (pd.to_datetime(row_sicredi['Data da venda'], format='%d/%m/%Y') + timedelta(days=1)).strftime('%d/%m/%Y')
            correspondencia_d1 = df_sistema[
                (df_sistema['EMPRESA'] == row_sicredi['Número do estabelecimento']) &
                (df_sistema['DATA DE FATURAMENTO'] == data_d1) &
                (df_sistema['Valor bruto sistema'] == row_sicredi['Valor bruto sicredi']) &
                (~df_sistema.index.isin(indices_utilizados_sistema))
            ].head(1)

            if not correspondencia_d1.empty:
                resultados.append((row_sicredi, correspondencia_d1.iloc[0], 'Correspondido (D+1)'))
                indices_utilizados_sistema.append(correspondencia_d1.index[0])
                indices_utilizados_sicredi.append(i)
            else:
                resultados.append((row_sicredi, pd.Series(), 'Não Correspondido'))
                indices_utilizados_sicredi.append(i)

    for j, row_sistema in df_sistema.iterrows():
        if j not in indices_utilizados_sistema:
            resultados.append((pd.Series(), row_sistema, 'Não Correspondido'))
            indices_utilizados_sistema.append(j)

    # Para garantir que mesmo as linhas sem correspondência do Sistema contenham todas as chaves,
    # criamos um dicionário padrão com todas as colunas do df_sistema.
    default_sistema = {col: "" for col in df_sistema.columns}
    final_result = pd.DataFrame([{
        **row_sicredi.to_dict(),
        **(row_sistema.to_dict() if not row_sistema.empty else default_sistema),
        'Status': status
    } for row_sicredi, row_sistema, status in resultados])

    final_result.fillna('', inplace=True)
    final_result['Diferença'] = ''
    cols = [col for col in final_result.columns if col not in ['Diferença', 'Status']] + ['Diferença', 'Status']
    final_result = final_result[cols]

    #####################################
    # Geração do arquivo Excel para download
    #####################################
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_result.to_excel(writer, index=False, sheet_name='Resultado')
            workbook = writer.book
            worksheet = writer.sheets['Resultado']

            max_row = len(final_result) + 1
            col_names = final_result.columns.tolist()
            col_valor_bruto_sicredi = col_names.index('Valor bruto sicredi')
            col_valor_bruto_sistema = col_names.index('Valor bruto sistema')
            col_diferenca = col_names.index('Diferença')

            col_letter_sicredi = col_idx_to_excel_col(col_valor_bruto_sicredi)
            col_letter_sistema = col_idx_to_excel_col(col_valor_bruto_sistema)
            col_letter_diferenca = col_idx_to_excel_col(col_diferenca)

            for row_num in range(2, max_row + 1):
                formula = f"={col_letter_sicredi}{row_num}-{col_letter_sistema}{row_num}"
                worksheet.write_formula(f"{col_letter_diferenca}{row_num}", formula)

            number_format = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column(col_diferenca, col_diferenca, 15, number_format)
    except Exception as e:
        st.error(f"Erro ao gerar o arquivo Excel consolidado: {e}")
        st.stop()

    processed_data = output.getvalue()
    st.download_button(
        label='Baixar planilha consolidada',
        data=processed_data,
        file_name='Resultado_Comparacao_Sicredi_Sistema_Final.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    st.success("Comparação concluída e arquivo pronto para download.")
