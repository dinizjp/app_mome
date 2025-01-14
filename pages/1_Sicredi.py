import streamlit as st
import pandas as pd
from io import BytesIO

st.title('Comparação de Vendas: Sicredi vs Sistema')

st.write('### Carregar planilha da Sicredi')
# Upload da planilha da Sicredi
uploaded_file_sicredi = st.file_uploader('Faça o upload do arquivo Sicredi aqui', type=['xlsx'], key='sicredi')

st.write('### Carregar planilha do Sistema')
# Upload da planilha do Sistema
uploaded_file_sistema = st.file_uploader('Faça o upload do arquivo do Sistema aqui', type=['xlsx'], key='sistema_sicredi')

# Verifica se ambos os arquivos foram carregados
if uploaded_file_sicredi is not None and uploaded_file_sistema is not None:
    # Leitura das planilhas
    df_sicredi = pd.read_excel(uploaded_file_sicredi, skiprows=9)
    df_sistema = pd.read_excel(uploaded_file_sistema)

    # Remover espaços extras dos nomes das colunas
    df_sicredi.columns = df_sicredi.columns.str.strip()
    df_sistema.columns = df_sistema.columns.str.strip()

    # Selecionar colunas desejadas
    colunas_desejadas_sicredi = ['Data de venda', 'Valor bruto', 'Número do estabelecimento']
    df_sicredi = df_sicredi[colunas_desejadas_sicredi]

    colunas_desejadas_sistema = ['DATA DE FATURAMENTO', 'VALOR BRUTO', 'EMPRESA']
    df_sistema = df_sistema[colunas_desejadas_sistema]

    # Definir o mapeamento dos códigos para os nomes (apenas na coluna 'Número do estabelecimento' da Sicredi)
    establishment_mapping = {
        '92139112': 'ARAGUAÍNA IV'
        # Adicione outros códigos se necessário
    }

    # Mapear os códigos na coluna 'Número do estabelecimento' da planilha da Sicredi
    df_sicredi['Número do estabelecimento'] = df_sicredi['Número do estabelecimento'].astype(str).str.strip()
    df_sicredi['Número do estabelecimento'] = df_sicredi['Número do estabelecimento'].map(establishment_mapping)

    # Verificar se há códigos não mapeados
    if df_sicredi['Número do estabelecimento'].isnull().any():
        unmapped_codes = df_sicredi[df_sicredi['Número do estabelecimento'].isnull()]['Número do estabelecimento'].unique()
        st.error(f"Existem códigos de estabelecimento sem mapeamento: {', '.join(unmapped_codes)}. Verifique o mapeamento fornecido.")
        st.stop()

    # Função para normalizar os nomes dos estabelecimentos
    def normalize_name(name):
        # Converter para maiúsculas
        name = name.upper()
        # Remover espaços extras
        name = ' '.join(name.split())
        # Substituir espaços por underscores
        name = name.replace(' ', '_')
        return name

    # Aplicar a normalização nas colunas 'Número do estabelecimento' e 'EMPRESA'
    df_sicredi['Número do estabelecimento'] = df_sicredi['Número do estabelecimento'].apply(normalize_name)
    df_sistema['EMPRESA'] = df_sistema['EMPRESA'].apply(normalize_name)

    # Converter as colunas de data para o tipo datetime
    df_sicredi['Data de venda'] = pd.to_datetime(df_sicredi['Data de venda'], dayfirst=True)
    df_sistema['DATA DE FATURAMENTO'] = pd.to_datetime(df_sistema['DATA DE FATURAMENTO'], dayfirst=True)

    # Formatar a data para string (dd/mm/yyyy) para garantir a consistência
    df_sicredi['Data de venda'] = df_sicredi['Data de venda'].dt.strftime('%d/%m/%Y')
    df_sistema['DATA DE FATURAMENTO'] = df_sistema['DATA DE FATURAMENTO'].dt.strftime('%d/%m/%Y')

    # Converter as colunas de valor para float
    df_sicredi['Valor bruto sicredi'] = df_sicredi['Valor bruto'].astype(float)
    df_sistema['Valor bruto sistema'] = df_sistema['VALOR BRUTO'].astype(float)

    # Remover colunas originais de valor bruto para evitar confusão
    df_sicredi.drop(columns=['Valor bruto'], inplace=True)
    df_sistema.drop(columns=['VALOR BRUTO'], inplace=True)

    # Resetar os índices para garantir uma comparação controlada
    df_sistema.reset_index(drop=True, inplace=True)
    df_sicredi.reset_index(drop=True, inplace=True)

    # Criar listas para armazenar os índices já utilizados nas duas planilhas
    indices_utilizados_sicredi = []
    indices_utilizados_sistema = []

    # Criar uma lista para armazenar os resultados
    resultados = []

    # Loop sobre cada linha da planilha da Sicredi
    for i, row_sicredi in df_sicredi.iterrows():
        # Encontrar a primeira correspondência no Sistema que ainda não foi utilizada
        correspondencia = df_sistema[
            (df_sistema['EMPRESA'] == row_sicredi['Número do estabelecimento']) &
            (df_sistema['DATA DE FATURAMENTO'] == row_sicredi['Data de venda']) &
            (df_sistema['Valor bruto sistema'] == row_sicredi['Valor bruto sicredi']) &
            (~df_sistema.index.isin(indices_utilizados_sistema))
        ].head(1)

        # Se uma correspondência for encontrada, armazená-la
        if not correspondencia.empty:
            resultados.append((row_sicredi, correspondencia.iloc[0], 'Correspondido'))
            indices_utilizados_sistema.append(correspondencia.index[0])
            indices_utilizados_sicredi.append(i)
        else:
            # Caso não haja correspondência, apenas adicione os dados da Sicredi
            resultados.append((row_sicredi, pd.Series(), 'Não Correspondido'))
            indices_utilizados_sicredi.append(i)

    # Adicionar as linhas da planilha do Sistema que não foram correspondidas
    for j, row_sistema in df_sistema.iterrows():
        if j not in indices_utilizados_sistema:
            # Adicionar o row_sistema com um row_sicredi vazio
            resultados.append((pd.Series(), row_sistema, 'Não Correspondido'))
            indices_utilizados_sistema.append(j)

    # Criar DataFrame final, mantendo todas as colunas das duas planilhas e adicionando o 'Status'
    final_result = pd.DataFrame([{
        **row_sicredi.to_dict(),
        **row_sistema.to_dict(),
        'Status': status
    } for row_sicredi, row_sistema, status in resultados])

    # Substituir NaN por vazio para melhor visualização
    final_result.fillna('', inplace=True)

    # **Adicionar a coluna 'Diferença' vazia ao DataFrame**
    final_result['Diferença'] = ''

    # Garantir que as colunas 'Diferença' e 'Status' sejam as últimas
    cols = [col for col in final_result.columns if col not in ['Diferença', 'Status']] + ['Diferença', 'Status']
    final_result = final_result[cols]

    # Converter o DataFrame final em um objeto BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_result.to_excel(writer, index=False, sheet_name='Resultado')
        workbook = writer.book
        worksheet = writer.sheets['Resultado']

        # Obter o número de linhas e colunas
        max_row = len(final_result) + 1  # +1 porque o Excel é 1-indexado (linha de cabeçalho)
        max_col = len(final_result.columns)

        # Encontrar as posições das colunas
        col_names = final_result.columns.tolist()
        col_valor_bruto_sicredi = col_names.index('Valor bruto sicredi')
        col_valor_bruto_sistema = col_names.index('Valor bruto sistema')
        col_diferenca = col_names.index('Diferença')

        # Converter índices de coluna para letras de Excel
        def col_idx_to_excel_col(idx):
            """Converte índice de coluna (zero-based) para letra da coluna no Excel"""
            idx += 1  # Ajuste para 1-based indexing do Excel
            col_str = ''
            while idx > 0:
                idx, remainder = divmod(idx - 1, 26)
                col_str = chr(65 + remainder) + col_str
            return col_str

        col_letter_sicredi = col_idx_to_excel_col(col_valor_bruto_sicredi)
        col_letter_sistema = col_idx_to_excel_col(col_valor_bruto_sistema)
        col_letter_diferenca = col_idx_to_excel_col(col_diferenca)

        # Escrever fórmulas na coluna 'Diferença'
        for row_num in range(2, max_row + 1):  # Começando da linha 2 (após o cabeçalho)
            formula = f"={col_letter_sicredi}{row_num}-{col_letter_sistema}{row_num}"
            worksheet.write_formula(f"{col_letter_diferenca}{row_num}", formula)

        # Formatar a coluna 'Diferença' como número com duas casas decimais
        number_format = workbook.add_format({'num_format': '#,##0.00'})
        worksheet.set_column(col_diferenca, col_diferenca, 15, number_format)

    processed_data = output.getvalue()

    # Botão para download da planilha consolidada
    st.download_button(
        label='Baixar planilha consolidada',
        data=processed_data,
        file_name='Resultado_Comparacao_Sicredi_Sistema_Final.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.success("Comparação concluída e arquivo pronto para download.")