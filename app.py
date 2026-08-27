from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from engine import load_reference, report, run


st.set_page_config(
    page_title="KPI Validator",
    page_icon="📊",
    layout="wide",
)

st.title("KPI Validator | ENTITY + KPI_CODE")
st.caption(
    "Validação automática das abas LOGS, BOPS e SL contra a aba "
    "SHAREPOINT do consolidador.xlsm."
)

reference_path = Path(__file__).resolve().parent / "consolidador.xlsm"


@st.cache_data(show_spinner=False)
def load_cached_reference(path_string, modified_time):
    path = Path(path_string)
    return load_reference(path.read_bytes(), path.name)


with st.sidebar:
    st.header("Parâmetros")

    absolute_tolerance = st.number_input(
        "Tolerância absoluta",
        min_value=0.0,
        value=0.01,
        format="%.8f",
    )

    relative_tolerance = st.number_input(
        "Tolerância relativa",
        min_value=0.0,
        value=0.0001,
        format="%.8f",
    )

    st.divider()
    st.caption("Chave exclusiva de comparação")
    st.code("ENTITY + KPI_CODE")

if not reference_path.exists():
    st.error(
        "O arquivo de referência consolidador.xlsm não foi encontrado. "
        "Posicione o arquivo na mesma pasta do app.py."
    )
    st.code(str(reference_path))
    st.stop()

try:
    reference_bundle = load_cached_reference(
        str(reference_path),
        reference_path.stat().st_mtime_ns,
    )
except Exception as error:
    st.error(
        f"Não foi possível carregar o arquivo de referência: "
        f"{type(error).__name__}: {error}"
    )
    st.stop()

if reference_bundle.get("error"):
    st.error(reference_bundle["error"])

    details = reference_bundle.get("details")
    if details:
        st.json(details)

    st.stop()

uploaded_file = st.file_uploader(
    "Arquivo com as abas LOGS, BOPS e SL",
    type=["xlsx", "xlsm"],
)

if uploaded_file is None:
    st.info("Envie o arquivo Excel que será comparado com o consolidador.")
    st.stop()

with st.spinner("Processando e comparando os KPIs..."):
    payload = run(
        uploaded_file.getvalue(),
        reference_bundle,
        absolute_tolerance,
        relative_tolerance,
    )

if payload.get("error"):
    st.error(payload["error"])

    details = payload.get("details")
    if details:
        st.json(details)

    st.stop()

results = payload.get("results", [])

if not results:
    st.warning("O processamento foi concluído, mas nenhuma comparação foi gerada.")
    st.stop()

comparison = pd.concat(results, ignore_index=True)

total_compared = len(comparison)
total_ok = int(comparison["STATUS"].eq("OK").sum())
total_findings = total_compared - total_ok
compliance = total_ok / total_compared * 100 if total_compared else 0.0

card_1, card_2, card_3, card_4 = st.columns(4)

card_1.metric("Comparados", f"{total_compared:,}".replace(",", "."))
card_2.metric("OK", f"{total_ok:,}".replace(",", "."))
card_3.metric("Achados", f"{total_findings:,}".replace(",", "."))
card_4.metric("Conformidade", f"{compliance:.2f}%")

summary_tab, comparison_tab, diagnostic_tab, export_tab = st.tabs(
    ["Resumo", "Comparação", "Diagnóstico", "Exportar"]
)

with summary_tab:
    summary = (
        comparison.groupby(["SOURCE", "STATUS"], dropna=False)
        .size()
        .reset_index(name="QUANTIDADE")
        .sort_values(["SOURCE", "STATUS"])
        .reset_index(drop=True)
    )

    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True,
    )

with comparison_tab:
    filter_column_1, filter_column_2, filter_column_3 = st.columns([2, 2, 3])

    available_sources = sorted(
        comparison["SOURCE"].dropna().astype(str).unique().tolist()
    )

    available_statuses = sorted(
        comparison["STATUS"].dropna().astype(str).unique().tolist()
    )

    with filter_column_1:
        selected_sources = st.multiselect(
            "Fonte",
            options=available_sources,
            default=available_sources,
        )

    with filter_column_2:
        selected_statuses = st.multiselect(
            "Status",
            options=available_statuses,
            default=available_statuses,
        )

    with filter_column_3:
        search_text = st.text_input(
            "Buscar ENTITY ou KPI_CODE",
            placeholder="Digite uma entidade ou código de KPI",
        )

    filtered = comparison[
        comparison["SOURCE"].isin(selected_sources)
        & comparison["STATUS"].isin(selected_statuses)
    ].copy()

    if search_text.strip():
        search_pattern = search_text.strip()

        filtered = filtered[
            filtered["ENTITY"]
            .astype(str)
            .str.contains(search_pattern, case=False, na=False, regex=False)
            |
            filtered["KPI_CODE"]
            .astype(str)
            .str.contains(search_pattern, case=False, na=False, regex=False)
        ]

    st.caption(
        f"{len(filtered):,} linha(s) exibida(s)"
        .replace(",", ".")
    )

    st.dataframe(
        filtered,
        use_container_width=True,
        hide_index=True,
        height=620,
    )

with diagnostic_tab:
    st.subheader("Arquivo enviado")

    uploaded_diagnostic = pd.DataFrame(
        [
            {
                "ABA_ESPERADA": report_name,
                "ABA_LOCALIZADA": payload["uploaded_sheets"].get(report_name),
                "LINHA_CABECALHO_EXCEL": payload["headers"].get(report_name),
                "COLUNA_ENTITY": payload["maps"].get(report_name, {}).get("entity"),
                "COLUNA_KPI_CODE": payload["maps"].get(report_name, {}).get("kpi"),
                "COLUNA_VALUE": payload["maps"].get(report_name, {}).get("value"),
                "COLUNA_KPI_NAME": payload["maps"].get(report_name, {}).get("kpi_name"),
            }
            for report_name in ["LOGS", "BOPS", "SL"]
        ]
    )

    st.dataframe(
        uploaded_diagnostic,
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Arquivo de referência")

    reference_diagnostic = pd.DataFrame(
        [
            {
                "ABA_ESPERADA": sheet_name,
                "ABA_LOCALIZADA": payload["reference_sheets"].get(sheet_name),
                "LINHA_CABECALHO_EXCEL": payload["headers"].get(sheet_name),
                "COLUNA_ENTITY": payload["maps"].get(sheet_name, {}).get("entity"),
                "COLUNA_KPI_CODE": payload["maps"].get(sheet_name, {}).get("kpi"),
                "COLUNA_VALUE": payload["maps"].get(sheet_name, {}).get("value"),
                "COLUNA_KPI_NAME": payload["maps"].get(sheet_name, {}).get("kpi_name"),
                "COLUNA_FORMULA": payload["maps"].get(sheet_name, {}).get("formula"),
                "COLUNA_UOM": payload["maps"].get(sheet_name, {}).get("uom"),
                "COLUNA_OWNER": payload["maps"].get(sheet_name, {}).get("owner"),
            }
            for sheet_name in ["SHAREPOINT", "DEFINITION BOOK"]
        ]
    )

    st.dataframe(
        reference_diagnostic,
        use_container_width=True,
        hide_index=True,
    )

    with st.expander("Abas disponíveis"):
        st.write("Arquivo enviado")
        st.write(payload.get("uploaded_sheet_names", []))

        st.write("Arquivo de referência")
        st.write(payload.get("reference_sheet_names", []))

with export_tab:
    st.write(
        "O relatório contém as abas Comparacao_Completa, OK, Achados, "
        "Resumo, Definition_Book e Parametros."
    )

    report_bytes = report(payload)

    st.download_button(
        label="Baixar relatório Excel",
        data=report_bytes,
        file_name=(
            f"KPI_Entity_KPI_"
            f"{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
    )
