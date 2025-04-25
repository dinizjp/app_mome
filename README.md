# app_mome

## Descrição

O projeto **app_mome** é uma aplicação que realiza a comparação de vendas entre uma planilha fornecida pelo Sicredi e os dados extraídos de um banco de dados SQL Server. Ela permite o upload de um arquivo Excel da Sicredi, consulta de dados do sistema via SQL, processamento, normalização de dados e cruzamento das informações para verificar correspondências e discrepâncias. A interface é construída com Streamlit, facilitando a interação do usuário com filtros de data e seleção de empresa.

## Sumário

- [Dependências](#dependências)
- [Instalação](#_instalação)
- [Uso](#uso)
- [Estrutura de Pastas](#estrutura-de-pastas)

## Dependências

As dependências necessárias estão listadas no arquivo `requirements.txt`:

- streamlit==1.32.0
- pandas==2.2.2
- XlsxWriter==3.2.0
- openpyxl==3.1.5
- pyodbc==5.2.0

## Instalação

Para configurar o ambiente, execute os comandos abaixo na pasta do projeto:

```bash
pip install -r requirements.txt
```

Certifique-se de que o arquivo `requirements.txt` esteja na mesma pasta que o terminal de comandos.

## Uso

Para iniciar a aplicação, execute:

```bash
streamlit run sicredi.py
```

Após o carregamento da aplicação, siga os passos:

1. Faça o upload do arquivo Excel da Sicredi na seção "Carregar planilha da Sicredi".
2. Informe as datas inicial e final para a consulta.
3. Selecione a empresa desejada na lista.
4. Clique no botão "Buscar dados do Sistema" para consultar os dados do banco.
5. Após o carregamento e processamento, a comparação será realizada automaticamente, exibindo os resultados na interface.

## Estrutura de Pastas

```
app_mome/
│
├── sicredi.py
├── requirements.txt
├── utils.py
└── maping.txt
```