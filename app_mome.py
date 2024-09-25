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
    colunas_desejadas = ['Data da venda', 'Hora da venda', 'Estabelecimento', 'Forma de pagamento', 'Valor bruto', 'NSU/DOC', 'Bandeira']
    df_cielo = df_cielo[colunas_desejadas]
    
    # Renomear colunas no df_sistema para corresponder ao df_cielo
    df_sistema.columns = df_sistema.columns.str.strip()
    
    # Converter as colunas de data para o tipo datetime
    df_sistema['DATA DE FATURAMENTO'] = pd.to_datetime(df_sistema['DATA DE FATURAMENTO'], dayfirst=True)
    df_cielo['Data da venda'] = pd.to_datetime(df_cielo['Data da venda'], dayfirst=True)
    
    # Formatar a data para string (dd/mm/yyyy) para garantir a consistência
    df_sistema['DATA DE FATURAMENTO'] = df_sistema['DATA DE FATURAMENTO'].dt.strftime('%d/%m/%Y')
    df_cielo['Data da venda'] = df_cielo['Data da venda'].dt.strftime('%d/%m/%Y')
    
    # Renomear a coluna 'Valor bruto' no sistema para preservar as duas versões no resultado
    df_sistema.rename(columns={'VALOR BRUTO': 'Valor bruto sistema'}, inplace=True)
    
    # Ordenar os DataFrames por 'Valor bruto' para manter a consistência na comparação
    df_sistema.sort_values('Valor bruto sistema', ascending=True, inplace=True)
    df_cielo.sort_values('Valor bruto', ascending=True, inplace=True)
    
    # Resetar os índices para garantir uma comparação controlada
    df_sistema = df_sistema.reset_index(drop=True)
    df_cielo = df_cielo.reset_index(drop=True)
    
    # Criar uma lista para armazenar os índices já utilizados no df_sistema para evitar duplicatas
    indices_utilizados = []
    
    # Criar uma lista para armazenar os resultados
    resultados = []
    
    # Loop sobre cada linha da planilha da Cielo
    for i, row_cielo in df_cielo.iterrows():
        # Encontrar a primeira correspondência no Sistema que ainda não foi utilizada
        correspondencia = df_sistema[
            (df_sistema['DATA DE FATURAMENTO'] == row_cielo['Data da venda']) &
            (df_sistema['Valor bruto sistema'] == row_cielo['Valor bruto']) &
            (~df_sistema.index.isin(indices_utilizados))
        ].head(1)
        
        # Se uma correspondência for encontrada, armazená-la
        if not correspondencia.empty:
            resultados.append((row_cielo, correspondencia.iloc[0]))
            indices_utilizados.append(correspondencia.index[0])
        else:
            # Caso não haja correspondência, apenas adicione os dados da Cielo
            resultados.append((row_cielo, pd.Series()))
    
    # Criar DataFrame final, mantendo todas as colunas das duas planilhas
    final_result = pd.DataFrame([{
        **row_cielo.to_dict(),
        **row_sistema.to_dict(),
        'Diferença': f"=E{idx+2}-M{idx+2}" if not row_sistema.empty else f"=E{idx+2}"
    } for idx, (row_cielo, row_sistema) in enumerate(resultados)])
    
    # Converter o DataFrame final em um objeto BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_result.to_excel(writer, index=False)
        # writer.save()  # Removido porque o gerenciador de contexto já salva e fecha o arquivo
    
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
