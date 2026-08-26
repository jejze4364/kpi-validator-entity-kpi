# KPI Validator

Aplicativo Streamlit para validação automática de KPIs, comparando um arquivo
enviado pelo usuário (abas **LOGS**, **BOPS** e **SL**) contra um arquivo de
referência (`consolidador.xlsm`, aba **SHAREPOINT**), usando exclusivamente a
chave composta **ENTITY + KPI_CODE**. A aba **DEFINITION BOOK** do arquivo de
referência é opcional e, quando presente, enriquece o resultado com nome do
KPI, fórmula, unidade de medida e responsável.

Não há banco de dados, autenticação nem integração direta com APIs do
Anaplan ou do SharePoint — tudo funciona a partir de arquivos Excel lidos em
memória durante a execução.

---

## 1. Objetivo

Automatizar a conferência periódica de KPIs (Logística, BOPS e Service
Level) contra uma base de referência consolidada, eliminando o trabalho
manual de cruzar planilhas, identificando rapidamente:

- KPIs **OK** (dentro da tolerância definida);
- KPIs **DIVERGENTE** (fora da tolerância);
- KPIs **NÃO ESTÁ NA REFERÊNCIA** (existe no arquivo enviado, mas não na
  referência);
- KPIs **SOMENTE NA REFERÊNCIA** (existe na referência, mas não foi
  enviado).

## 2. Arquitetura

```
Usuário → upload (.xlsx/.xlsm) ─┐
                                 ├─► engine.run() ─► engine.report()
consolidador.xlsm (no repo) ────┘        │                 │
                                          ▼                 ▼
                                   DataFrames de       Relatório Excel
                                   comparação          (download)
                                          │
                                          ▼
                                     app.py (Streamlit)
                                     cards, tabelas, filtros,
                                     diagnóstico, exportação
```

- **`app.py`** — camada de interface (Streamlit). Não contém lógica de
  negócio: apenas carrega o arquivo de referência, recebe o upload, chama
  `engine.run`/`engine.report` e renderiza os resultados.
- **`engine.py`** — toda a lógica de leitura, normalização, mapeamento de
  colunas por *alias* e comparação. Não depende do Streamlit; pode ser
  testado isoladamente (`import engine`).
- **`consolidador.xlsm`** — arquivo de referência, lido automaticamente pelo
  `app.py` a partir da mesma pasta (não é feito upload dele).

## 3. Estrutura do repositório

```
kpi-validator/
├── app.py               # Interface Streamlit
├── engine.py             # Lógica de leitura, normalização e comparação
├── consolidador.xlsm      # Arquivo de referência (aba SHAREPOINT + DEFINITION BOOK)
├── requirements.txt        # Dependências Python
├── runtime.txt             # Versão do Python (Streamlit Cloud)
├── packages.txt            # Dependências de sistema (vazio — não há nenhuma)
├── .gitignore
└── README.md
```

## 4. Regras de comparação

### 4.1 Chave

A comparação usa **exclusivamente** `ENTITY + KPI_CODE`. Não existe
fallback para outras chaves.

### 4.2 Localização de abas

- No arquivo de **referência**: aba obrigatória equivalente a
  `SHAREPOINT` (aceita variações de nome/caixa) e aba opcional equivalente
  a `DEFINITION BOOK`.
- No arquivo **enviado**: abas obrigatórias `LOGS`, `BOPS` e `SL`. Se
  qualquer uma estiver ausente, o app exibe um erro claro listando o que
  falta e interrompe o processamento.

### 4.3 Detecção de cabeçalho

Cada aba tem seu cabeçalho localizado automaticamente nas primeiras 80
linhas, pontuando linhas que contenham colunas de KPI_CODE, ENTITY e maior
quantidade de células preenchidas.

### 4.4 Reconhecimento de colunas (aliases PT/EN)

| Papel      | Exemplos de aliases reconhecidos                                   |
|------------|----------------------------------------------------------------------|
| Entidade   | `entity`, `entidade`, `unit name`, `unidade`, `bu`, `location`, `dc` |
| KPI        | `kpi code`, `kpi_code`, `codigo kpi`, `código kpi`, `cod kpi`         |
| Valor      | `current_value ac`, `valor anaplan`, `valor origem`, `actual`, `ac`, `value`, `valor` |
| Nome KPI   | `kpi name`, `nome kpi`, `description`, `descrição`                    |
| Fórmula    | `formula`, `fórmula`, `regra`                                         |
| Unidade    | `unit_of_measure`, `unit of measure`, `uom`                           |
| Responsável| `owner`, `responsavel`, `responsável`                                 |

A busca ignora acentos, maiúsculas/minúsculas e espaços excedentes, e aceita
correspondência parcial quando não há correspondência exata.

### 4.5 Normalização de texto e números

- Texto: remoção de acentos, colapso de espaços e conversão para
  maiúsculas.
- Código de KPI: extração do padrão `XX-Y0000` (ex.: `SL-K0031`) quando
  presente em um texto maior (ex.: `"SL-K0031 - Total Volume..."`).
- Números: reconhece tanto `1234.56` quanto o padrão brasileiro
  `1.234.567,89` (ponto como separador de milhar, vírgula como separador
  decimal).

### 4.6 Agregação

Linhas com a mesma chave `ENTITY + KPI_CODE` são agregadas por **soma** dos
valores, tanto no arquivo enviado quanto na referência.

### 4.7 Tolerâncias

- **Tolerância absoluta** — padrão `0.01`.
- **Tolerância relativa** — padrão `0.0001` (aplicada sobre o valor
  absoluto da referência).
- Ambas configuráveis na barra lateral da interface.
- Um KPI é considerado **OK** quando
  `|diferença| <= tolerância_absoluta + tolerância_relativa * |valor_referência|`.

### 4.8 Classificação de status

| Status                     | Significado                                              |
|-----------------------------|-----------------------------------------------------------|
| `OK`                         | Dentro da tolerância                                      |
| `DIVERGENTE`                 | Fora da tolerância                                        |
| `NÃO ESTÁ NA REFERÊNCIA`     | Chave existe no arquivo enviado, mas não na referência     |
| `SOMENTE NA REFERÊNCIA`      | Chave existe na referência, mas não foi enviada             |

## 5. Interface

- **Cards**: total comparado, quantidade OK, quantidade de achados e
  percentual de conformidade.
- **Aba Resumo**: contagem por `SOURCE` (LOGS/BOPS/SL) x `STATUS`.
- **Aba Comparação**: tabela detalhada com filtro por status e busca por
  `ENTITY` ou `KPI_CODE`.
- **Aba Diagnóstico**: abas localizadas, linha de cabeçalho detectada e
  colunas mapeadas — tanto do arquivo de referência quanto do arquivo
  enviado. Útil para identificar rapidamente por que uma coluna não foi
  reconhecida.
- **Aba Exportar**: botão para baixar o relatório Excel completo.
- **Mensagens de erro**: exibidas de forma clara quando o arquivo de
  referência está ausente/corrompido, quando faltam abas obrigatórias no
  arquivo enviado, ou quando colunas essenciais (entidade, KPI, valor) não
  puderam ser identificadas — neste último caso, a mensagem lista as
  colunas efetivamente detectadas na aba.

## 6. Relatório exportado

O botão **Baixar relatório Excel** gera um arquivo `.xlsx` com as abas:

| Aba                  | Conteúdo                                                            |
|-----------------------|------------------------------------------------------------------------|
| `Comparacao_Completa`  | Todas as linhas comparadas (LOGS + BOPS + SL)                        |
| `OK`                   | Somente as linhas com status `OK`                                    |
| `Achados`              | Todas as linhas com status diferente de `OK`                          |
| `Resumo`               | Contagem por `SOURCE` x `STATUS`                                       |
| `Definition_Book`      | Enriquecimento vindo da aba `DEFINITION BOOK` (quando disponível)      |
| `Parametros`           | Chave utilizada, nome do arquivo de referência e tolerâncias aplicadas |

## 7. Execução local

```bash
# 1. Clonar o repositório
git clone <url-do-repositorio>
cd kpi-validator

# 2. Criar e ativar um ambiente virtual (recomendado)
python3.12 -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate

# 3. Instalar as dependências
pip install -r requirements.txt

# 4. Garantir que consolidador.xlsm está na mesma pasta de app.py

# 5. Rodar o aplicativo
streamlit run app.py
```

O app abrirá em `http://localhost:8501`.

## 8. Publicação no Streamlit Community Cloud

1. Suba este repositório para o GitHub, **incluindo o arquivo
   `consolidador.xlsm`** na raiz do projeto (ao lado de `app.py`).
2. Em [share.streamlit.io](https://share.streamlit.io), clique em **New app**.
3. Selecione o repositório e a branch.
4. Em **Main file path**, informe `app.py`.
5. Em **Advanced settings**, selecione **Python 3.12**.
6. Clique em **Deploy**.

O Streamlit Cloud instalará automaticamente as dependências de
`requirements.txt` e (se houver) os pacotes de sistema listados em
`packages.txt`.

## 9. Posicionamento do consolidador.xlsm

O arquivo `consolidador.xlsm` deve ficar **na raiz do repositório, ao lado
de `app.py`**. O aplicativo o carrega automaticamente pelo caminho relativo
`Path(__file__).parent / "consolidador.xlsm"` — não é necessário fazer
upload dele pela interface. Para atualizar a base de referência, basta
substituir esse arquivo no repositório (mantendo o mesmo nome) e publicar
novamente.

Caso o arquivo não seja encontrado, o app exibe uma mensagem de erro clara
orientando a reposicioná-lo antes de continuar.

## 10. Limitações

- A comparação depende inteiramente da qualidade dos cabeçalhos e da
  presença de pelo menos uma coluna reconhecível para entidade, KPI e
  valor em cada aba; layouts muito fora do padrão podem exigir ajuste dos
  aliases em `engine.py`.
- O reconhecimento de código de KPI assume o padrão `XX-Y0000` (duas ou
  três letras, hífen, `K` ou `R`, 3 a 5 dígitos, com sufixo opcional de
  ano). Códigos em outros formatos são normalizados apenas como texto.
- Arquivos muito grandes (dezenas de milhares de linhas) podem levar
  alguns segundos para processar, especialmente na leitura do
  `consolidador.xlsm`.
- O app não persiste dados entre sessões: cada execução processa os
  arquivos em memória e nada é salvo em disco além do relatório baixado
  pelo próprio usuário.

## 11. Segurança

- Nenhuma credencial, token ou segredo é utilizado ou armazenado pelo
  aplicativo.
- Não há chamadas a APIs externas (Anaplan, SharePoint ou outras) — toda a
  leitura é feita localmente sobre os bytes dos arquivos Excel.
- O arquivo de referência (`consolidador.xlsm`) contém dados internos da
  empresa; avalie a visibilidade do repositório (público/privado) antes de
  publicá-lo, já que ele fica versionado no GitHub.
- Segredos do Streamlit (`.streamlit/secrets.toml`), caso venham a ser
  usados no futuro, já estão listados no `.gitignore` e nunca devem ser
  commitados.
- O upload do usuário é processado inteiramente em memória
  (`uploaded_file.getvalue()`), sem gravação em disco no servidor.

## 12. Principais funções (`engine.py`)

- `load_reference(data, filename)` — carrega e padroniza o arquivo de
  referência (`SHAREPOINT` + `DEFINITION BOOK` opcional).
- `run(upload_bytes, reference_bundle, abs_tol, rel_tol)` — carrega o
  arquivo enviado, compara cada aba (`LOGS`, `BOPS`, `SL`) contra a
  referência e retorna os resultados, diagnósticos e parâmetros usados.
- `report(payload)` — gera os bytes do relatório Excel final a partir do
  resultado de `run`.

Essas três funções cobrem todo o ciclo de uso do `app.py` e podem ser
reaproveitadas em scripts ou testes automatizados fora da interface
Streamlit.
