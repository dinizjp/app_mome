import streamlit as st
import pandas as pd
from io import BytesIO

st.title('Comparação de Planilhas Cielo e Sistema')

# Upload da planilha da Cielo
uploaded_file_cielo = st.file_uploader('Carregar planilha da Cielo', type=['xlsx'], key='cielo')

# Upload da planilha do Sistema
uploaded_file_sistema = st.file_uploader('Carregar planilha do Sistema', type=['xlsx'], key='sistema')

# Verifica se ambos os arquivos foram carregados
if uploaded_file_cielo is not None and uploaded_file_sistema is not None:
    # Use skiprows para pular as 9 primeiras linhas na planilha da Cielo
    df_cielo = pd.read_excel(uploaded_file_cielo, skiprows=9)
    df_sistema = pd.read_excel(uploaded_file_sistema)
    
    # Selecione apenas as colunas desejadas
    colunas_desejadas = ['Data da venda', 'Hora da venda', 'Estabelecimento', 'Forma de pagamento',
                         'Valor bruto', 'NSU/DOC', 'Bandeira']
    df_cielo = df_cielo[colunas_desejadas]
    
    # Renomear colunas no df_sistema para corresponder ao df_cielo
    df_sistema.columns = df_sistema.columns.str.strip()
    
    # Converter as colunas de data para o tipo datetime
    df_sistema['DATA DE FATURAMENTO'] = pd.to_datetime(df_sistema['DATA DE FATURAMENTO'], dayfirst=True)
    df_cielo['Data da venda'] = pd.to_datetime(df_cielo['Data da venda'], dayfirst=True)
    
    # Formatar a data para string (dd/mm/yyyy) para garantir a consistência
    df_sistema['DATA DE FATURAMENTO'] = df_sistema['DATA DE FATURAMENTO'].dt.strftime('%d/%m/%Y')
    df_cielo['Data da venda'] = df_cielo['Data da venda'].dt.strftime('%d/%m/%Y')
    
    # Renomear a coluna 'VALOR BRUTO' no sistema para preservar as duas versões no resultado
    df_sistema.rename(columns={'VALOR BRUTO': 'Valor bruto sistema'}, inplace=True)
    
    # Ordenar os DataFrames por 'Valor bruto' para manter a consistência na comparação
    df_sistema.sort_values('Valor bruto sistema', ascending=True, inplace=True)
    df_cielo.sort_values('Valor bruto', ascending=True, inplace=True)
    
    # Resetar os índices para garantir uma comparação controlada
    df_sistema = df_sistema.reset_index(drop=True)
    df_cielo = df_cielo.reset_index(drop=True)
    
    # Criar listas para armazenar os índices já utilizados nas duas planilhas
    indices_utilizados_cielo = []
    indices_utilizados_sistema = []
    
    # Criar uma lista para armazenar os resultados
    resultados = []
    
    # Loop sobre cada linha da planilha da Cielo
    for i, row_cielo in df_cielo.iterrows():
        # Encontrar a primeira correspondência no Sistema que ainda não foi utilizada
        correspondencia = df_sistema[
            (df_sistema['DATA DE FATURAMENTO'] == row_cielo['Data da venda']) &
            (df_sistema['Valor bruto sistema'] == row_cielo['Valor bruto']) &
            (~df_sistema.index.isin(indices_utilizados_sistema))
        ].head(1)
        
        # Se uma correspondência for encontrada, armazená-la
        if not correspondencia.empty:
            resultados.append((row_cielo, correspondencia.iloc[0]))
            indices_utilizados_sistema.append(correspondencia.index[0])
            indices_utilizados_cielo.append(i)
        else:
            # Caso não haja correspondência, apenas adicione os dados da Cielo
            resultados.append((row_cielo, pd.Series()))
            indices_utilizados_cielo.append(i)
    
    # Adicionar as linhas da planilha do Sistema que não foram correspondidas
    for j, row_sistema in df_sistema.iterrows():
        if j not in indices_utilizados_sistema:
            # Adicionar o row_sistema com um row_cielo vazio
            resultados.append((pd.Series(), row_sistema))
            indices_utilizados_sistema.append(j)
    
    # Criar DataFrame final, mantendo todas as colunas das duas planilhas
    final_result = pd.DataFrame([{
        **row_cielo.to_dict(),
        **row_sistema.to_dict()
    } for row_cielo, row_sistema in resultados])
    
    # Substituir NaN por vazio para melhor visualização
    final_result.fillna('', inplace=True)
    
    # Reordenar as colunas para melhor visualização (opcional)
    # Você pode ajustar a ordem das colunas conforme necessário
    cols = list(final_result.columns)
    final_result = final_result[cols]
    
    # Converter o DataFrame final em um objeto BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_result.to_excel(writer, index=False, sheet_name='Resultado')
        workbook = writer.book
        worksheet = writer.sheets['Resultado']
        
        # Encontrar as posições das colunas 'Valor bruto' e 'Valor bruto sistema'
        valor_bruto_col = final_result.columns.get_loc('Valor bruto')
        valor_bruto_sistema_col = final_result.columns.get_loc('Valor bruto sistema')
        diferenca_col = len(final_result.columns)  # A coluna 'Diferença' será adicionada após a última coluna
        
        # Escrever a coluna 'Diferença' com a fórmula
        for row_num in range(1, len(final_result) + 1):
            cell_formula = f"=IF(ISBLANK(${chr(65 + valor_bruto_sistema_col)}{row_num + 1})," \
                           f"${chr(65 + valor_bruto_col)}{row_num + 1}," \
                           f"${chr(65 + valor_bruto_col)}{row_num + 1}-${chr(65 + valor_bruto_sistema_col)}{row_num + 1})"
            worksheet.write_formula(row_num, diferenca_col, cell_formula)
        
        # Adicionar o cabeçalho da coluna 'Diferença'
        worksheet.write(0, diferenca_col, 'Diferença')
    
    # Obter o conteúdo do BytesIO
    processed_data = output.getvalue()
    
    # Botão para download da planilha consolidada
    st.download_button(
        label='Baixar planilha consolidada',
        data=processed_data,
        file_name='Resultado_Comparacao_Cielo_Sistema_Final.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    
    st.success("Comparação concluída e arquivo pronto para download.")
