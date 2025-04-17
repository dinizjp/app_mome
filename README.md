# README para o Projeto 'app_mome'

## Objetivo do Projeto

O projeto "app_mome" tem como finalidade realizar a comparação de dados de vendas entre um sistema interno e um arquivo de relatórios fornecido pelo Sicredi. 

Ele permite que os usuários analisem e reconciliem informações financeiras de maneira eficiente, destacando correspondências e divergências entre as duas fontes de dados.

Através de uma interface interativa, os usuários podem carregar o arquivo do Sicredi, selecionar filtros para busca de dados no sistema e gerar um relatório detalhado em Excel.

## Dependências e Como Instalá-las

O projeto possui as seguintes dependências listadas no arquivo `requirements.txt`:

- `streamlit==1.32.0`
- `pandas==2.2.2`
- `XlsxWriter==3.2.0`
- `openpyxl==3.1.5`
- `pyodbc==5.2.0`

Para instalar todas as dependências, execute:

```bash
pip install -r requirements.txt
```

Além disso, é necessário ter o driver ODBC Driver 17 for SQL Server instalado na sua máquina.

## Como Executar e Testar

1. **Configurar Credenciais:**  
   Configure as credenciais de acesso ao banco de dados no seu arquivo `secrets.toml` ou na interface do Streamlit Cloud, incluindo servidor, banco, usuário e senha.

2. **Iniciar o Aplicativo:**  
   Execute o comando abaixo para rodar a aplicação:

   ```bash
   streamlit run sicredi.py
   ```

3. **Acessar a Interface:**  
   A aplicação será acessível pelo navegador. 

4. **Carregar Arquivo Sicredi:**  
   Faça o upload do arquivo Excel do Sicredi na seção "Carregar planilha da Sicredi".

5. **Buscar Dados no Sistema:**  
   Selecione a loja desejada, período e clique em "Buscar dados do Sistema". Os dados serão carregados automaticamente.

6. **Realizar a Comparação:**  
   Depois de carregar a planilha e obter os dados do sistema, o aplicativo realizará a comparação e exibirá os resultados, além de gerar um relatório para download.

## Estrutura de Pastas

```
app_mome/
│
├── sicredi.py          # Script principal da aplicação com lógica da interface
├── utils.py            # Funções auxiliares para busca de dados, normalizações e conversões
├── requirements.txt    # Lista de dependências do projeto
├── packages.txt         # Pacotes adicionais necessários (ex: msodbcsql17)
├── maping.txt           # Mapeamento de IDs de empresas para nomes
```

---

Este README fornece uma visão clara e objetiva para instalação, execução e entendimento geral do projeto "app_mome".
