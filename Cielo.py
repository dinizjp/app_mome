import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import timedelta  # Importado para calcular D+1

st.write('Carregar planilha da Cielo')
# Upload da planilha da Cielo
uploaded_file_cielo = st.file_uploader('Faça o upload do arquivo aqui', type=['xlsx'], key='cielo')

st.write('Carregar planilha do Sistema')
# Upload da planilha do Sistema
uploaded_file_sistema = st.file_uploader('Faça o upload do arquivo aqui', type=['xlsx'], key='sistema')

# Verifica se ambos os arquivos foram carregados
if uploaded_file_cielo is not None and uploaded_file_sistema is not None:
    # Leitura das planilhas
    df_cielo = pd.read_excel(uploaded_file_cielo, skiprows=9)
    df_sistema = pd.read_excel(uploaded_file_sistema)

    # Remover espaços extras dos nomes das colunas
    df_cielo.columns = df_cielo.columns.str.strip()
    df_sistema.columns = df_sistema.columns.str.strip()

    # Selecionar colunas desejadas
    colunas_desejadas_cielo = ['Data da venda', 'Hora da venda', 'Estabelecimento', 'Forma de pagamento',
                               'Valor bruto', 'NSU/DOC', 'Bandeira']
    df_cielo = df_cielo[colunas_desejadas_cielo]

    colunas_desejadas_sistema = ['ID EMPRESA', 'EMPRESA', 'ID VENDA', 'FORMA DE PAGAMENTO', 'NOME',
                                 'ID CAIXA', 'NSU', 'VALOR BRUTO', 'DATA DE FATURAMENTO', 'EMISSAO']
    df_sistema = df_sistema[colunas_desejadas_sistema]

    # Definir o mapeamento dos códigos para os nomes (apenas na coluna 'Estabelecimento' da Cielo)
    establishment_mapping = {
        '2853441991': 'ARAGUAÍNA II',
        '2857146692': 'ARAGUAÍNA I',
        '2859045451': 'FORMOSA I',
        '2859443554': 'IMPERATRIZ III',
        '2859558840': 'COLINAS',
        '2856942487': 'IMPERATRIZ I',
        '2886221362': 'ESTREITO',
        '2808425621': 'ESTREITO',
        '2893038330': 'ARAGUAÍNA IV',
        '2893716436': 'BALSAS II',
        '2892453156': 'IMPERATRIZ II',
        '2845701319': 'ARAGUAÍNA III',
        '2857798851': 'GURUPI I',
        '2892125809': 'BALSAS I',
        '2893143711': 'IMPERATRIZ III',
        '2894588490': 'FORMOSA I',
        '2857845108': 'IMPERATRIZ II',
        '2859875209': 'BALSAS I',
        '2892453024': 'GURUPI I',
        '2893927950': 'IMPERATRIZ I',
        '2893972475': 'ARAGUAÍNA III',
        '2894031291': 'ESTREITO',
        '2895310933': 'ARAGUAÍNA I',
        '2896720299': 'ARAGUAÍNA II',
        '2871227700': 'GUARAÍ',
        '2888776450': 'COLINAS',
        '2891499632': 'CANAÃ DOS CARAJÁS',
        '2808601667': 'COLINAS',
        '2893481730': 'COLINAS'
    }

    # Mapear os códigos na coluna 'Estabelecimento' da planilha da Cielo
    df_cielo['Estabelecimento'] = df_cielo['Estabelecimento'].astype(str).str.strip()
    df_cielo['Estabelecimento'] = df_cielo['Estabelecimento'].map(establishment_mapping)

    # Verificar se há códigos não mapeados
    if df_cielo['Estabelecimento'].isnull().any():
        unmapped_codes = df_cielo[df_cielo['Estabelecimento'].isnull()]['Estabelecimento'].unique()
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

    # Aplicar a normalização nas colunas 'Estabelecimento' e 'EMPRESA'
    df_cielo['Estabelecimento'] = df_cielo['Estabelecimento'].apply(normalize_name)
    df_sistema['EMPRESA'] = df_sistema['EMPRESA'].apply(normalize_name)

    # Converter as colunas de data para o tipo datetime
    df_cielo['Data da venda'] = pd.to_datetime(df_cielo['Data da venda'], dayfirst=True)
    df_sistema['DATA DE FATURAMENTO'] = pd.to_datetime(df_sistema['DATA DE FATURAMENTO'], dayfirst=True)

    # Formatar a data para string (dd/mm/yyyy) para garantir a consistência
    df_cielo['Data da venda'] = df_cielo['Data da venda'].dt.strftime('%d/%m/%Y')
    df_sistema['DATA DE FATURAMENTO'] = df_sistema['DATA DE FATURAMENTO'].dt.strftime('%d/%m/%Y')

    # Converter as colunas de valor para float
    df_cielo['Valor bruto cielo'] = df_cielo['Valor bruto'].astype(float)
    df_sistema['Valor bruto sistema'] = df_sistema['VALOR BRUTO'].astype(float)

    # Remover colunas originais de valor bruto para evitar confusão
    df_cielo.drop(columns=['Valor bruto'], inplace=True)
    df_sistema.drop(columns=['VALOR BRUTO'], inplace=True)

    # Resetar os índices para garantir uma comparação controlada
    df_sistema.reset_index(drop=True, inplace=True)
    df_cielo.reset_index(drop=True, inplace=True)

    # Criar listas para armazenar os índices já utilizados nas duas planilhas
    indices_utilizados_cielo = []
    indices_utilizados_sistema = []

    # Criar uma lista para armazenar os resultados
    resultados = []

    # Loop sobre cada linha da planilha da Cielo
    for i, row_cielo in df_cielo.iterrows():
        # Primeira tentativa: correspondência na data exata
        correspondencia = df_sistema[
            (df_sistema['EMPRESA'] == row_cielo['Estabelecimento']) &
            (df_sistema['DATA DE FATURAMENTO'] == row_cielo['Data da venda']) &
            (df_sistema['Valor bruto sistema'] == row_cielo['Valor bruto cielo']) &
            (~df_sistema.index.isin(indices_utilizados_sistema))
        ].head(1)

        # Se uma correspondência for encontrada na data exata
        if not correspondencia.empty:
            resultados.append((row_cielo, correspondencia.iloc[0], 'Correspondido (Data Exata)'))
            indices_utilizados_sistema.append(correspondencia.index[0])
            indices_utilizados_cielo.append(i)
        else:
            # Segunda tentativa: correspondência em D+1
            data_d1 = (pd.to_datetime(row_cielo['Data da venda'], format='%d/%m/%Y') + timedelta(days=1)).strftime('%d/%m/%Y')
            correspondencia_d1 = df_sistema[
                (df_sistema['EMPRESA'] == row_cielo['Estabelecimento']) &
                (df_sistema['DATA DE FATURAMENTO'] == data_d1) &
                (df_sistema['Valor bruto sistema'] == row_cielo['Valor bruto cielo']) &
                (~df_sistema.index.isin(indices_utilizados_sistema))
            ].head(1)

            # Se uma correspondência for encontrada em D+1
            if not correspondencia_d1.empty:
                resultados.append((row_cielo, correspondencia_d1.iloc[0], 'Correspondido (D+1)'))
                indices_utilizados_sistema.append(correspondencia_d1.index[0])
                indices_utilizados_cielo.append(i)
            else:
                # Caso não haja correspondência
                resultados.append((row_cielo, pd.Series(), 'Não Correspondido'))
                indices_utilizados_cielo.append(i)

    # Adicionar as linhas da planilha do Sistema que não foram correspondidas
    for j, row_sistema in df_sistema.iterrows():
        if j not in indices_utilizados_sistema:
            # Adicionar o row_sistema com um row_cielo vazio
            resultados.append((pd.Series(), row_sistema, 'Não Correspondido'))
            indices_utilizados_sistema.append(j)

    # Criar DataFrame final, mantendo todas as colunas das duas planilhas e adicionando o 'Status'
    final_result = pd.DataFrame([{
        **row_cielo.to_dict(),
        **row_sistema.to_dict(),
        'Status': status
    } for row_cielo, row_sistema, status in resultados])

    # Substituir NaN por vazio para melhor visualização
    final_result.fillna('', inplace=True)

    # Adicionar a coluna 'Diferença' vazia ao DataFrame
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
        col_valor_bruto_cielo = col_names.index('Valor bruto cielo')
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

        col_letter_cielo = col_idx_to_excel_col(col_valor_bruto_cielo)
        col_letter_sistema = col_idx_to_excel_col(col_valor_bruto_sistema)
        col_letter_diferenca = col_idx_to_excel_col(col_diferenca)

        # Escrever fórmulas na coluna 'Diferença'
        for row_num in range(2, max_row + 1):  # Começando da linha 2 (após o cabeçalho)
            formula = f"={col_letter_cielo}{row_num}-{col_letter_sistema}{row_num}"
            worksheet.write_formula(f"{col_letter_diferenca}{row_num}", formula)

        # Formatar a coluna 'Diferença' como número com duas casas decimais
        number_format = workbook.add_format({'num_format': '#,##0.00'})
        worksheet.set_column(col_diferenca, col_diferenca, 15, number_format)

    processed_data = output.getvalue()

    # Botão para download da planilha consolidada
    st.download_button(
        label='Baixar planilha consolidada',
        data=processed_data,
        file_name='Resultado_Comparacao_Cielo_Sistema_Final.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.success("Comparação concluída e arquivo pronto para download.")