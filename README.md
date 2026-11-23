# Sistema Financeiro em Python e Excel

Aplicação desktop para controlar produtos e movimentações financeiras de
estoque. Os dados ficam em **MySQL**, a interface é feita em **Tkinter** e os
relatórios são enviados, já formatados, para uma planilha do **Google Sheets**,
que pode ser aberta ou baixada como arquivo Excel (`.xlsx`).

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![Interface](https://img.shields.io/badge/Interface-Tkinter-2D3A8C)
![Banco](https://img.shields.io/badge/Banco-MySQL-4479A1?logo=mysql&logoColor=white)
![Planilha](https://img.shields.io/badge/Planilha-Google%20Sheets-34A853?logo=googlesheets&logoColor=white)

## Funcionalidades

| Módulo | Descrição |
|---|---|
| Produtos | Cadastro, consulta, edição e exclusão |
| Movimentações | Entradas e saídas vinculadas a cada produto |
| Gráficos | Movimentações por categoria e saldo mensal |
| Relatórios | Dados consolidados e resumo gerencial na planilha |

Regras de negócio:

- o saldo de um produto nunca fica negativo;
- uma saída não pode superar o estoque disponível;
- toda gravação ou exclusão pede confirmação;
- os dados são validados antes de chegar ao banco.

## Relatórios

A planilha exportada tem duas abas:

1. **Dados Consolidados**: cada movimentação com o produto, a categoria e o
   valor total (consulta com `JOIN`).
2. **Resumo Gerencial**: estoque e valor por categoria, e totais por tipo de
   movimentação.

As abas recebem cabeçalho destacado, linhas alternadas, colunas ajustadas,
primeira linha congelada e formatos de data e moeda (R$). A exportação
substitui o conteúdo anterior das duas abas, por isso o sistema pede
confirmação antes de enviar.

## Como executar

### 1. Banco de dados

Instale o MySQL e execute o arquivo [`schema.sql`](schema.sql). Ele cria o banco
`sistema_financeiro`, as tabelas, **5 produtos** e **15 movimentações** de
exemplo.

### 2. Dependências

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 3. Variáveis de ambiente do banco

| Variável | Padrão |
|---|---|
| `DB_HOST` | `localhost` |
| `DB_USER` | `root` |
| `DB_PASSWORD` | vazio |
| `DB_NAME` | `sistema_financeiro` |

### 4. Google Sheets

1. Crie uma conta de serviço no Google Cloud e ative as APIs **Google Sheets**
   e **Google Drive**.
2. Compartilhe a planilha com o e-mail da conta de serviço.
3. Defina `GOOGLE_CREDENTIALS_FILE` com o caminho do JSON da chave.
4. Opcional: defina `GOOGLE_SHEETS_ID` para usar uma planilha já existente.

Sem `GOOGLE_SHEETS_ID`, o sistema procura a planilha `relatorio_financeiro` e a
cria se ela não existir.

> Não publique o JSON de credenciais no repositório. Guarde-o em local seguro e
> informe apenas o caminho por variável de ambiente.

### 5. Iniciar

```powershell
python "Sistema Financeiro em Python e Excel.py"
```

## Arquivos

| Arquivo | Conteúdo |
|---|---|
| `Sistema Financeiro em Python e Excel.py` | Interface, acesso ao banco, gráficos e exportação |
| `schema.sql` | Estrutura do banco e dados iniciais |
| `requirements.txt` | Dependências Python |

## Estrutura do código

- **Configuração e SQL**: constantes, paleta de cores e consultas no topo do arquivo.
- **Especificação das telas** (`ENTITIES`, `METRICS`): colunas, títulos e botões
  das abas são descritos como dados, e a interface é montada a partir deles.
- **Banco**: `fetch_frames`, `stock_balance` e `Transaction` (confirma ou desfaz).
- **Relatórios**: `build_reports`, `export_reports` e `style_worksheet`.
- **Interface**: classe `FinanceApp`.
