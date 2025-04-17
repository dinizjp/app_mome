```markdown
# README para o Projeto 'app_mome'

## Objetivo do Projeto

O projeto "app_mome" tem como finalidade realizar a comparação de dados de vendas entre um sistema interno e um arquivo de relatórios fornecido pelo Sicredi. Ele permite que os usuários analisem e reconciliem informações financeiras de maneira eficiente, destacando correspondências e divergências entre as duas fontes de dados. Através de uma interface interativa, os usuários podem carregar o arquivo do Sicredi, selecionar filtros para busca de dados no sistema e gerar um relatório detalhado em Excel.

## Dependências e Como Instalá-las

O projeto possui algumas dependências que precisam ser instaladas para seu correto funcionamento. As dependências são listadas no arquivo `requirements.txt`. Para instalá-las, execute o seguinte comando:

```bash
pip install -r requirements.txt
```

As bibliotecas necessárias são:

- `streamlit==1.32.0`: Para a criação da interface web interativa.
- `pandas==2.2.2`: Para manipulação e análise de dados.
- `XlsxWriter==3.2.0`: Para geração de arquivos Excel.
- `openpyxl==3.1.5`: Para ler/escrever arquivos Excel.
- `pyodbc==5.2.0`: Para conectar-se a bancos de dados SQL Server.

Além disso, é necessário ter o driver ODBC Driver 17 for SQL Server instalado na sua máquina.

## Como Executar e Testar

1. **Configuração das Credenciais**: Configure as credenciais de acesso ao banco de dados no seu arquivo `secrets.toml` ou na interface do Streamlit Cloud, incluindo informações como servidor, banco de dados, usuário e senha.

2. **Rodar o Aplicativo**: Utilize o seguinte comando para iniciar a aplicação:

   ```bash
   streamlit run sicredi.py
   ```

3. **Acesso à Interface**: O aplicativo será executado em um servidor local, e você poderá acessá-lo através do navegador.

4. **Carregar Dados**: Faça o upload do arquivo Excel fornecido pelo Sicredi e insira os filtros desejados (data inicial, data final e seleção da empresa).

5. **Executar Comparação**: Após carregar os dados, o aplicativo realizará a comparação e apresentará os resultados, além de permitir o download de um relatório em Excel.

## Estrutura de Pastas e Breve Descrição dos Arquivos Principais

- `README.md`: Este documento, que fornece informações e instruções sobre o projeto.
- `requirements.txt`: Lista de dependências do Python que devem ser instaladas.
- `packages.txt`: Informações de pacotes necessários para o projeto.
- `maping.txt`: Mapeamento utilizado para associar IDs de empresas a seus respectivos nomes; fundamental para a consulta ao sistema.
- `sicredi.py`: Script principal da aplicação que contém a lógica para carregar os dados, comparar as informações e interagir com o usuário por meio da interface Streamlit.
- `utils.py`: Arquivo utilitário que contém funções auxiliares para buscar dados do sistema, normalizar nomes e converter índices de coluna do Excel.

Com esses componentes, o projeto "app_mome" oferece uma solução robusta para a reconciliação de dados de vendas, permitindo uma análise mais ágil e precisa.
```