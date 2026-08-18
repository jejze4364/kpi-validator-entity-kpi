# KPI Validator Entity + KPI

- Upload único do Consolidador.
- Lê somente SHAREPOINT, LOGS, BOPS, SL e Definition Book.
- Chave exclusiva: `ENTITY + KPI_CODE`.
- Nenhuma outra coluna participa do match.
- SharePoint usa preferencialmente `CURRENT_VALUE AC`.
- LOGS/BOPS/SL usam preferencialmente `Valor Anaplan` ou `AC`.
- Definition Book apenas enriquece nome, fórmula, UOM e owner.
- Duplicidades da mesma chave são somadas e sinalizadas.

Execute `run_app.bat`.
