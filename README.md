# KPI Validator

Aplicativo Streamlit para validação automática de KPIs, comparando um arquivo enviado pelo usuário, com as abas **LOGS**, **BOPS** e **SL**, contra um arquivo de referência (`consolidador.xlsm`, aba **SHAREPOINT**), usando exclusivamente a chave composta **ENTITY + KPI_CODE**. A aba **DEFINITION BOOK** do arquivo de referência é opcional e, quando presente, enriquece o resultado com nome do KPI, fórmula, unidade de medida e responsável.

Não há banco de dados, autenticação nem integração direta com APIs do Anaplan ou do SharePoint. Todo o processamento é realizado a partir de arquivos Excel lidos em memória durante a execução.

## 1. Objetivo

Automatizar a conferência periódica de KPIs de Logística, BOPS e Service Level contra uma base de referência consolidada, eliminando o trabalho manual de cruzar planilhas e identificando rapidamente:

- KPIs **OK**, dentro da tolerância definida;
- KPIs **DIVERGENTE**, fora da tolerância;
- KPIs **NÃO ESTÁ NA REFERÊNCIA**, existentes no arquivo enviado, mas ausentes na referência;
- KPIs **SOMENTE NA REFERÊNCIA**, existentes na referência, mas não enviados.

## 2. Arquitetura

```text
Usuário -> upload (.xlsx/.xlsm) --+
                                  +--> engine.run() --> engine.report()
consolidador.xlsm (no repo) ------+        |                 |
                                           v                 v
                                    DataFrames de       Relatório Excel
                                    comparação          para download
                                           |
                                           v
                                      app.py (Streamlit)
                                      cards, tabelas, filtros,
                                      diagnóstico e exportação
```

- **app.py**: camada de interface Streamlit. Não contém lógica de negócio. Carrega o arquivo de referência, recebe o upload, chama `engine.run()` e `engine.report()` e renderiza os resultados.
- **engine.py**: concentra a lógica de leitura, normalização, mapeamento de colunas por alias e comparação. Não depende do Streamlit e pode ser testado isoladamente com `import engine`.
- **consolidador.xlsm**: arquivo de referência lido automaticamente pelo `app.py` a partir da mesma pasta. Não é feito upload desse arquivo pela interface.

## 3. Estrutura do repositório

```text
kpi-validator/
├── app.py                 # Interface Streamlit
├── engine.py              # Leitura, normalização e comparação
├── consolidador.xlsm      # Referência: SHAREPOINT + DEFINITION BOOK
├── requirements.txt       # Dependências Python
├── runtime.txt            # Versão do Python no Streamlit Cloud
├── packages.txt           # Dependências de sistema
├── .gitignore
└── README.md
```

## 4. Regras de comparação

### 4.1 Chave

A comparação usa **exclusivamente** `ENTITY + KPI_CODE`. Não existe fallback para outras chaves.

### 4.2 Localização de abas

- No arquivo de **referência**, é obrigatória uma aba equivalente a `SHAREPOINT`, aceitando variações de nome e caixa. Uma aba equivalente a `DEFINITION BOOK` é opcional.
- No arquivo **enviado**, são obrigatórias as abas `LOGS`, `BOPS` e `SL`. Se qualquer uma estiver ausente, o aplicativo exibe um erro claro, lista as abas faltantes e interrompe o processamento.

### 4.3 Detecção de cabeçalho

Cada aba tem seu cabeçalho localizado automaticamente nas primeiras 80 linhas. A detecção pontua as linhas que contenham colunas de `KPI_CODE`, `ENTITY` e a maior quantidade de células preenchidas.

### 4.4 Reconhecimento de colunas por aliases PT/EN

| Papel | Exemplos de aliases reconhecidos |
|---|---|
| Entidade | `entity`, `entidade`, `unit name`, `unidade`, `bu`, `location`, `dc` |
| KPI | `kpi code`, `kpi_code`, `codigo kpi`, `código kpi`, `cod kpi` |
| Valor | `current_value ac`, `valor anaplan`, `valor origem`, `actual`, `ac`, `value`, `valor` |
| Nome KPI | `kpi name`, `nome kpi`, `description`, `descrição` |
| Fórmula | `formula`, `fórmula`, `regra` |
| Unidade | `unit_of_measure`, `unit of measure`, `uom` |
| Responsável | `owner`, `responsavel`, `responsável` |

A busca ignora acentos, diferenças entre maiúsculas e minúsculas e espaços excedentes. Também aceita correspondência parcial quando não há correspondência exata.

### 4.5 Normalização de texto e números

- **Texto**: remoção de acentos, colapso de espaços e conversão para maiúsculas.
- **Código de KPI**: extração do padrão `XX-Y0000`, por exemplo `SL-K0031`, quando presente em um texto maior, como `SL-K0031 - Total Volume...`.
- **Números**: reconhecimento tanto de `1234.56` quanto do padrão brasileiro `1.234.567,89`.

### 4.6 Agregação

Linhas com a mesma chave `ENTITY + KPI_CODE` são agregadas por **soma** dos valores, tanto no arquivo enviado quanto na referência.

### 4.7 Tolerâncias

- **Tolerância absoluta**: padrão `0.01`.
- **Tolerância relativa**: padrão `0.0001`, aplicada sobre o valor absoluto da referência.
- Ambas são configuráveis na barra lateral da interface.
- Um KPI é considerado **OK** quando:

```text
|diferença| <= tolerância_absoluta + tolerância_relativa * |valor_referência|
```

### 4.8 Classificação de status

| Status | Significado |
|---|---|
| `OK` | Dentro da tolerância |
| `DIVERGENTE` | Fora da tolerância |
| `NÃO ESTÁ NA REFERÊNCIA` | Chave existe no arquivo enviado, mas não na referência |
| `SOMENTE NA REFERÊNCIA` | Chave existe na referência, mas não foi enviada |

## 5. Interface

- **Cards**: total comparado, quantidade OK, quantidade de achados e percentual de conformidade.
- **Aba Resumo**: contagem por `SOURCE`, considerando `LOGS`, `BOPS` e `SL`, versus `STATUS`.
- **Aba Comparação**: tabela detalhada com filtro por status e busca por `ENTITY` ou `KPI_CODE`.
- **Aba Diagnóstico**: abas localizadas, linha de cabeçalho detectada e colunas mapeadas, tanto no arquivo de referência quanto no arquivo enviado.
- **Aba Exportar**: botão para baixar o relatório Excel completo.
- **Mensagens de erro**: exibidas quando o arquivo de referência está ausente ou corrompido, quando faltam abas obrigatórias no arquivo enviado ou quando colunas essenciais não puderam ser identificadas.

## 6. Relatório exportado

O botão **Baixar relatório Excel** gera um arquivo `.xlsx` com as seguintes abas:

| Aba | Conteúdo |
|---|---|
| `Comparacao_Completa` | Todas as linhas comparadas de LOGS, BOPS e SL |
| `OK` | Somente as linhas com status `OK` |
| `Achados` | Todas as linhas com status diferente de `OK` |
| `Resumo` | Contagem por `SOURCE` versus `STATUS` |
| `Definition_Book` | Enriquecimento vindo da aba `DEFINITION BOOK`, quando disponível |
| `Parametros` | Chave utilizada, nome do arquivo de referência e tolerâncias aplicadas |

## 7. Execução local

```bash
# 1. Clonar o repositório
git clone <url-do-repositorio>
cd kpi-validator

# 2. Criar e ativar um ambiente virtual
python3.12 -m venv .venv
source .venv/bin/activate

# No Windows:
# .venv\Scripts\activate

# 3. Instalar as dependências
pip install -r requirements.txt

# 4. Garantir que consolidador.xlsm esteja na mesma pasta de app.py

# 5. Rodar o aplicativo
streamlit run app.py
```

O aplicativo abrirá em `http://localhost:8501`.

## 8. Publicação no Streamlit Community Cloud

1. Suba o repositório para o GitHub, incluindo o arquivo `consolidador.xlsm` na raiz do projeto, ao lado de `app.py`.
2. Acesse [Streamlit Community Cloud](https://share.streamlit.io/).
3. Clique em **New app**.
4. Selecione o repositório e a branch.
5. Em **Main file path**, informe `app.py`.
6. Em **Advanced settings**, selecione **Python 3.12**.
7. Clique em **Deploy**.

O Streamlit Cloud instalará automaticamente as dependências de `requirements.txt` e, se houver, os pacotes de sistema listados em `packages.txt`.

## 9. Posicionamento do consolidador.xlsm

O arquivo `consolidador.xlsm` deve ficar na raiz do repositório, ao lado de `app.py`. O aplicativo o carrega automaticamente pelo caminho relativo:

```python
Path(__file__).parent / "consolidador.xlsm"
```

Não é necessário fazer upload do arquivo de referência pela interface. Para atualizar a base, basta substituir o arquivo no repositório, mantendo o mesmo nome, e publicar novamente.

Caso o arquivo não seja encontrado, o aplicativo exibe uma mensagem de erro orientando seu reposicionamento antes de continuar.

## 10. Limitações

- A comparação depende da qualidade dos cabeçalhos e da presença de pelo menos uma coluna reconhecível para entidade, KPI e valor em cada aba.
- Layouts muito fora do padrão podem exigir ajuste dos aliases em `engine.py`.
- O reconhecimento de código de KPI assume o padrão `XX-Y0000`, com duas ou três letras, hífen, `K` ou `R`, de três a cinco dígitos e sufixo opcional de ano.
- Códigos em outros formatos são normalizados apenas como texto.
- Arquivos muito grandes podem exigir mais tempo de processamento, principalmente durante a leitura do `consolidador.xlsm`.
- O aplicativo não persiste dados entre sessões. Cada execução processa os arquivos em memória e nada é salvo em disco além dos arquivos baixados pelo próprio usuário.

## 11. Segurança

- Nenhuma credencial, token ou segredo é utilizado ou armazenado pelo aplicativo.
- Não há chamadas a APIs externas do Anaplan, SharePoint ou de outros serviços.
- Toda a leitura é realizada localmente sobre os bytes dos arquivos Excel.
- O arquivo `consolidador.xlsm` contém dados internos da empresa. A visibilidade do repositório deve ser avaliada antes da publicação.
- Segredos do Streamlit, como `.streamlit/secrets.toml`, devem permanecer no `.gitignore` e nunca ser versionados.
- O upload do usuário é processado em memória por `uploaded_file.getvalue()`, sem gravação em disco no servidor.

## 12. Principais funções de engine.py

- `load_reference(data, filename)`: carrega e padroniza o arquivo de referência, incluindo `SHAREPOINT` e `DEFINITION BOOK` opcional.
- `run(upload_bytes, reference_bundle, abs_tol, rel_tol)`: carrega o arquivo enviado, compara `LOGS`, `BOPS` e `SL` contra a referência e retorna resultados, diagnósticos e parâmetros.
- `report(payload)`: gera os bytes do relatório Excel final a partir do resultado de `run()`.

Essas três funções cobrem todo o ciclo de uso do `app.py` e podem ser reaproveitadas em scripts ou testes automatizados fora da interface Streamlit.

## 13. Possíveis alterações futuras

O aplicativo poderá ser ampliado para usar o `consolidador.xlsm` também como fonte de consulta e distribuição de arquivos. Assim, classificações, responsáveis, KPIs e downloads ficarão disponíveis imediatamente após a abertura do aplicativo, sem que o usuário precise carregar um arquivo.

O upload continuará sendo necessário apenas para executar a validação dos valores enviados contra a referência.

### 13.1 Carregamento automático das informações

Ao iniciar, o aplicativo deverá carregar automaticamente o `consolidador.xlsm` localizado na raiz do repositório e utilizar suas informações para preencher a página inicial.

Sem qualquer upload, o aplicativo poderá apresentar:

- classificações existentes;
- responsáveis identificados;
- quantidade de KPIs por responsável;
- quantidade de KPIs por classificação;
- quantidade de entidades;
- distribuição dos KPIs entre `LOGS`, `BOPS` e `SL`;
- código e nome dos KPIs;
- fórmula;
- unidade de medida;
- responsável;
- origem ou grupo do KPI;
- demais informações disponíveis na aba `DEFINITION BOOK`.

Caso algum campo não exista no consolidador, a interface deverá ocultar somente a informação correspondente, sem impedir o funcionamento das demais áreas.

### 13.2 Exibição das classificações

Deverá ser criada uma área chamada **Classificações**, preenchida automaticamente com os dados do `consolidador.xlsm`.

A consulta poderá ser organizada por:

- classificação;
- responsável;
- origem;
- código do KPI;
- nome do KPI;
- entidade;
- unidade de medida;
- categoria disponível no `DEFINITION BOOK`.

Filtros sugeridos:

- **Classificação**;
- **Responsável**;
- **SOURCE**;
- **ENTITY**;
- **KPI_CODE**;
- **Nome do KPI**;
- **Unidade de medida**.

A tabela deverá permitir busca textual e atualização automática de acordo com os filtros selecionados.

Se não houver uma coluna chamada `CLASSIFICAÇÃO`, o sistema poderá procurar uma coluna equivalente definida nos aliases do `engine.py`. Se nenhuma coluna compatível for encontrada, deverá exibir:

> A informação de classificação não foi encontrada no consolidador.

### 13.3 Templates padrões por responsável

O aplicativo deverá permitir o download do template padrão correspondente a cada responsável ou grupo de KPIs.

Os templates deverão ser gerados automaticamente a partir do `consolidador.xlsm`, sem necessidade de manter arquivos individuais no repositório.

O usuário deverá selecionar um responsável e clicar em **Baixar template do responsável**. O arquivo gerado deverá conter somente os registros associados ao responsável selecionado.

Quando disponíveis no consolidador, poderão ser incluídas as seguintes colunas:

- `SOURCE`;
- `ENTITY`;
- `KPI_CODE`;
- `KPI_NAME`;
- `CLASSIFICAÇÃO`;
- `FORMULA`;
- `UNIT_OF_MEASURE`;
- `OWNER`;
- campo destinado ao preenchimento do valor;
- demais campos obrigatórios para o processo de reporte.

O template deverá manter o layout necessário para posterior validação no aplicativo e poderá conter as seguintes abas:

- `LOGS`;
- `BOPS`;
- `SL`;
- `INSTRUÇÕES`.

Cada aba de origem deverá conter apenas os KPIs correspondentes ao responsável selecionado. A aba `INSTRUÇÕES` poderá apresentar:

- responsável selecionado;
- data e hora de geração;
- nome do arquivo de referência utilizado;
- quantidade de KPIs incluídos;
- orientações para preenchimento;
- identificação das colunas que não devem ser alteradas;
- versão do template.

Nome sugerido:

```text
template_<responsavel>.xlsx
```

Os nomes deverão ser normalizados para evitar espaços, acentos e caracteres incompatíveis.

### 13.4 Download de todos os templates

Além do download individual, o aplicativo poderá disponibilizar a opção **Baixar todos os templates**.

Essa opção deverá gerar um arquivo `.zip` com um template para cada responsável identificado no consolidador.

```text
templates_kpi.zip
├── template_responsavel_01.xlsx
├── template_responsavel_02.xlsx
├── template_responsavel_03.xlsx
└── resumo_templates.xlsx
```

O arquivo `resumo_templates.xlsx` poderá apresentar:

- responsável;
- nome do arquivo gerado;
- quantidade de KPIs;
- quantidade de entidades;
- quantidade de KPIs de `LOGS`;
- quantidade de KPIs de `BOPS`;
- quantidade de KPIs de `SL`;
- classificações encontradas.

Registros sem responsável deverão ser agrupados em:

```text
template_sem_responsavel.xlsx
```

### 13.5 Download do consolidador completo

Deverá existir um botão **Baixar consolidador completo**, disponível sem upload.

O botão deverá entregar o próprio `consolidador.xlsm` carregado pelo aplicativo, preservando:

- formato `.xlsm`;
- abas existentes;
- fórmulas;
- macros;
- formatações;
- estrutura original.

O arquivo não deverá ser convertido para `.xlsx`, pois essa conversão pode remover macros e outros elementos do arquivo original.

A interface também poderá exibir:

- nome do arquivo;
- tamanho;
- data de modificação disponível no ambiente;
- status de carregamento;
- quantidade de registros da aba `SHAREPOINT`;
- existência ou ausência da aba `DEFINITION BOOK`.

Exemplo:

```text
Arquivo de referência: consolidador.xlsm
Status: carregado com sucesso
Aba SHAREPOINT: localizada
Aba DEFINITION BOOK: localizada
```

### 13.6 Central de Downloads

A página inicial deverá possuir uma seção chamada **Central de Downloads**, disponível assim que o arquivo de referência for carregado.

Essa área poderá conter:

1. **Baixar consolidador completo**;
2. **Baixar template por responsável**;
3. **Baixar todos os templates**;
4. **Baixar catálogo de classificações**;
5. **Baixar Definition Book**, quando disponível;
6. **Baixar lista de responsáveis**;
7. **Baixar relação completa de KPIs**.

Nenhuma dessas opções deverá depender do upload do arquivo de validação.

O upload será exigido somente para:

- comparar os valores enviados com a referência;
- calcular diferenças;
- aplicar tolerâncias;
- classificar os resultados;
- gerar o relatório de comparação.

### 13.7 Catálogo de classificações

O botão **Baixar catálogo de classificações** deverá gerar um arquivo Excel com informações extraídas do consolidador.

| Aba | Conteúdo |
|---|---|
| `Classificacoes` | Relação completa das classificações disponíveis |
| `KPIs` | Códigos e nomes dos KPIs |
| `Responsaveis` | Responsáveis identificados e quantidade de KPIs |
| `Entidades` | Entidades existentes no consolidador |
| `Resumo` | Contagens consolidadas |
| `Definition_Book` | Informações complementares disponíveis |

Nome sugerido:

```text
catalogo_classificacoes_kpi.xlsx
```

O catálogo deverá ser gerado integralmente em memória.

### 13.8 Organização futura da interface

A interface poderá ser reorganizada nas seguintes páginas ou abas:

#### Início

Apresentação do aplicativo, status do consolidador e indicadores gerais da referência.

#### Classificações

Consulta de classificações, KPIs, responsáveis e entidades disponíveis.

#### Templates

Seleção de responsável e geração dos templates padrões.

#### Central de Downloads

Download do consolidador, catálogos, `DEFINITION BOOK`, relação de KPIs e pacote com todos os templates.

#### Validação

Upload do arquivo preenchido e execução da comparação.

#### Resultados

Cards, resumo, comparação detalhada e filtros dos resultados.

#### Diagnóstico

Detalhes das abas, cabeçalhos e colunas identificadas.

#### Exportar

Download do relatório final da validação.

### 13.9 Cards da página inicial

Antes de qualquer upload, o aplicativo poderá apresentar cards calculados exclusivamente a partir do `consolidador.xlsm`:

- **Total de KPIs**;
- **Total de entidades**;
- **Total de responsáveis**;
- **Total de classificações**;
- **KPIs de LOGS**;
- **KPIs de BOPS**;
- **KPIs de SL**;
- **KPIs sem responsável**;
- **KPIs sem classificação**.

Após o upload e a validação, poderão ser apresentados os cards de resultado:

- total comparado;
- quantidade OK;
- quantidade de achados;
- percentual de conformidade.

### 13.10 Regras para geração dos templates

A geração dos templates deverá:

1. Ler as informações diretamente do `consolidador.xlsm`.
2. Não exigir upload prévio.
3. Filtrar os registros pelo responsável selecionado.
4. Manter somente as colunas necessárias ao preenchimento e à validação.
5. Preservar `ENTITY` e `KPI_CODE` como identificadores obrigatórios.
6. Não criar registros duplicados silenciosamente.
7. Ordenar os registros por `SOURCE`, `ENTITY` e `KPI_CODE`.
8. Criar as abas `LOGS`, `BOPS` e `SL` conforme a origem dos KPIs.
9. Manter as abas obrigatórias mesmo quando não houver registros, caso isso seja necessário para manter a compatibilidade com a validação.
10. Registrar a data de geração e o arquivo de referência utilizado.
11. Normalizar o nome do responsável usado no nome do arquivo.
12. Gerar o arquivo integralmente em memória.
13. Não alterar o `consolidador.xlsm` original.

### 13.11 Tratamento de responsáveis

O sistema deverá normalizar os responsáveis para agrupamento e comparação, preservando o nome original para exibição.

A normalização poderá considerar:

- remoção de espaços excedentes;
- comparação sem diferenciação entre maiúsculas e minúsculas;
- remoção de acentos para agrupamento;
- tratamento de células vazias;
- separação de múltiplos responsáveis, quando houver um padrão definido.

Se um mesmo responsável estiver registrado com pequenas variações, o sistema deverá usar o valor normalizado para evitar templates duplicados e exibir a grafia original mais representativa.

### 13.12 Funcionamento sem arquivo enviado

Enquanto nenhum arquivo tiver sido carregado, o aplicativo deverá permanecer funcional para consulta da referência.

Nesse estado, o usuário poderá:

- visualizar classificações;
- consultar KPIs;
- filtrar responsáveis;
- consultar entidades;
- visualizar informações do `DEFINITION BOOK`;
- baixar templates;
- baixar o consolidador completo;
- baixar catálogos e relações auxiliares.

As áreas de comparação deverão exibir apenas a orientação:

> Envie o arquivo preenchido para iniciar a validação dos KPIs. As consultas e os downloads da base de referência já estão disponíveis.

O aplicativo não deverá apresentar cards zerados de validação antes do upload.

### 13.13 Alterações sugeridas na arquitetura

```text
consolidador.xlsm
        |
        +--> load_reference()
                 |
                 +--> informações da página inicial
                 +--> classificações
                 +--> responsáveis
                 +--> catálogo de KPIs
                 +--> geração de templates
                 +--> central de downloads

Usuário --> upload opcional
                 |
                 +--> run()
                        |
                        +--> comparação
                        +--> diagnóstico
                        +--> report()
```

O `app.py` deverá continuar responsável apenas pela interface. O `engine.py` poderá receber funções adicionais:

```python
get_reference_summary(reference_bundle)
get_classifications(reference_bundle)
get_responsibles(reference_bundle)
get_kpi_catalog(reference_bundle)
generate_responsible_template(reference_bundle, responsible)
generate_all_templates(reference_bundle)
generate_classification_catalog(reference_bundle)
get_reference_file_download(reference_bytes)
```

### 13.14 Estrutura futura do repositório

```text
kpi-validator/
├── app.py
├── engine.py
├── downloads.py
├── templates.py
├── consolidador.xlsm
├── requirements.txt
├── runtime.txt
├── packages.txt
├── .gitignore
└── README.md
```

Responsabilidades sugeridas:

- `app.py`: interface Streamlit;
- `engine.py`: leitura, normalização, comparação e diagnóstico;
- `templates.py`: geração dos templates por responsável;
- `downloads.py`: geração dos catálogos, arquivos auxiliares e pacote `.zip`;
- `consolidador.xlsm`: fonte única das informações de referência.

### 13.15 Atualização automática

Sempre que o `consolidador.xlsm` for substituído no repositório, classificações, responsáveis, KPIs, templates e arquivos de download deverão ser atualizados automaticamente na próxima execução.

Não deverá ser necessário alterar manualmente:

- listas de responsáveis;
- listas de classificações;
- arquivos de template individuais;
- catálogo de KPIs;
- opções dos filtros;
- arquivos disponíveis para download.

O `consolidador.xlsm` deverá permanecer como a fonte única da verdade.

### 13.16 Controle e rastreabilidade

Os arquivos gerados deverão incluir uma aba ou seção de controle com:

- nome do arquivo de referência;
- data e hora da geração;
- responsável selecionado;
- quantidade de registros exportados;
- quantidade de KPIs;
- quantidade de entidades;
- classificações incluídas;
- versão da estrutura do template;
- filtros aplicados.

Esse controle permitirá identificar qual base foi utilizada na geração de cada template ou catálogo.

### 13.17 Segurança dos downloads

Como o consolidador e os templates podem conter dados internos, a disponibilização dos arquivos deverá considerar a visibilidade do ambiente em que o aplicativo está publicado.

Antes da publicação, deverá ser avaliado se:

- o repositório é privado;
- o aplicativo exige controle de acesso;
- todas as informações podem ser disponibilizadas aos usuários;
- determinadas colunas devem ser removidas dos templates;
- o download do consolidador completo deve ser restrito.

A aplicação não deverá disponibilizar automaticamente informações restritas sem autorização prévia para uso no portal.
