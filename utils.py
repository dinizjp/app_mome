import pandas as pd
import pyodbc

def fetch_system_data(start_date, end_date, selected_company, id_empresa_mapping, secrets):
    """
    Busca os dados do Sistema diretamente do banco de dados usando os filtros de data e empresa.
    """
    # Recupera o ID da empresa a partir do mapeamento
    reverse_mapping = {v: k for k, v in id_empresa_mapping.items()}
    selected_id = reverse_mapping[selected_company]
    
    # Formata as datas para compor a query
    data_inicio = f"{start_date.strftime('%Y-%m-%d')} 00:00:00"
    data_fim = f"{end_date.strftime('%Y-%m-%d')} 23:59:59"
    
    query = f"""
    SELECT 
      e.ID_Empresa as [ID EMPRESA],
      E.NomeFantasia as [EMPRESA],
      CA.ID_Venda as [ID VENDA],
      FP.Descricao as [FORMA DE PAGAMENTO],
      U.Nome as [NOME],
      CA.ID_Caixa as [ID CAIXA],
      CA.Documento_Cartao as [NSU],
      Replace(Convert(Varchar,Convert(Decimal(18,2),CA.Valor)),'.',',') as [VALOR BRUTO],
      CASE 
         WHEN format(VS.Data_Faturamento, 'dd/MM/yyyy') IS NULL 
         THEN format(CA.datacadastro, 'dd/MM/yyyy') 
         ELSE format(VS.Data_Faturamento, 'dd/MM/yyyy') 
      END as [DATA DE FATURAMENTO],
      CASE 
         WHEN VS.Data_Faturamento IS NULL 
         THEN CA.datacadastro 
         ELSE VS.Data_Faturamento 
      END as [EMISSAO]
    FROM ContasAReceber CA
      LEFT JOIN FormasPagamento FP ON FP.ID_Forma = CA.ID_Forma
      INNER JOIN Fechamento_Caixas FC ON FC.ID_Empresa = CA.ID_Empresa 
         AND FC.ID_Caixa = CA.ID_Caixa 
         AND FC.ID_Origem_Caixa = CA.id_origem_caixa
      INNER JOIN Usuarios U ON U.ID_Usuario = FC.ID_Usuario
      LEFT JOIN Vendas_Sorveteria VS ON VS.ID_Empresa = CA.ID_Empresa 
         AND VS.ID_Venda = CA.ID_Venda
      INNER JOIN Empresas E ON CA.ID_Empresa = E.ID_Empresa
    WHERE 
      CA.ID_Forma IN (5,6)
      AND E.TipoEmpresa = 'Sorveteria'
      AND IsNull(CA.Emissao, VS.Data_Faturamento) >= '{data_inicio}'
      AND IsNull(CA.Emissao, VS.Data_Faturamento) <= '{data_fim}'
      AND e.ID_Empresa IN ({selected_id})
      AND ca.ID_Origem_Caixa = 1
    ORDER BY 
      CA.Valor,
      E.NomeFantasia,
      IsNull(VS.Data_Faturamento, CA.Emissao)
    """
    
    try:
        conn = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={secrets['mssql']['server']};"
            f"DATABASE={secrets['mssql']['database']};"
            f"UID={secrets['mssql']['username']};"
            f"PWD={secrets['mssql']['password']};"
            "Encrypt=yes;"
            "TrustServerCertificate=yes;"
        )
        df = pd.read_sql(query, conn)
        conn.close()
        return df
    except Exception as e:
        raise Exception(f"Erro ao buscar dados do Sistema: {e}")

def normalize_name(name: str) -> str:
    """
    Normaliza o nome convertendo para maiúsculas, removendo espaços extras e substituindo por underline.
    """
    name = name.upper()
    name = ' '.join(name.split())
    name = name.replace(' ', '_')
    return name

def col_idx_to_excel_col(idx: int) -> str:
    """
    Converte o índice da coluna (0-indexado) para a notação de coluna do Excel.
    """
    idx += 1
    col_str = ''
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return col_str
