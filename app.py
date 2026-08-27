from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from engine import load_reference, report, run

st.set_page_config(page_title="KPI Validator", page_icon="📊", layout="wide")
st.title("KPI Validator | ENTITY + KPI_CODE")
st.caption("Validação das abas LOGS, BOPS e SL contra SHAREPOINT do consolidador.xlsm.")

REFERENCE_PATH = Path(__file__).resolve().parent / "consolidador.xlsm"


@st.cache_data(show_spinner=False)
def load_cached_reference(path_string, modified_time):
    path = Path(path_string)
    return load_reference(path.read_bytes(), path.name)


with st.sidebar:
    st.header("Parâmetros")
    absolute_tolerance = st.number_input(
        "Tolerância absoluta", min_value=0.0, value=0.01, format="%.8f"
    )
    relative_tolerance = st.number_input(
        "Tolerância relativa", min_value=0.0, value=0.0001, format="%.8f"
    )
    st.divider()
    st.caption("Chave exclusiva de comparação")
    st.code("ENTITY + KPI_CODE")

if not REFERENCE_PATH.is_file():
    st.error(
        "O arquivo consolidador.xlsm não foi encontrado. "
        "Coloque-o na mesma pasta do app.py."
    )
    st.code(str(REFERENCE_PATH))
    st.stop()

try:
    reference_bundle = load_cached_reference(
        str(REFERENCE_PATH), REFERENCE_PATH.stat().st_mtime_ns
    )
except Exception as error:
    st.error(f"Não foi possível abrir a referência: {type(error).__name__}: {error}")
    st.stop()

if reference_bundle.get("error"):
    st.error(reference_bundle["error"])
    if reference_bundle.get("details"):
        st.json(reference_bundle["details"])
    st.stop()

uploaded_file = st.file_uploader(
    "Arquivo com as abas LOGS, BOPS e SL", type=["xlsx", "xlsm"]
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
    if payload.get("details"):
        st.json(payload["details"])
    st.stop()

results = payload.get("results", [])
if not results:
    st.warning("Nenhuma comparação foi gerada.")
    st.stop()

comparison = pd.concat(results, ignore_index=True)
total = len(comparison)
total_ok = int(comparison["STATUS"].eq("OK").sum())
findings = total - total_ok
compliance = total_ok / total * 100 if total else 0.0

card1, card2, card3, card4 = st.columns(4)
card1.metric("Comparados", f"{total:,}".replace(",", "."))
card2.metric("OK", f"{total_ok:,}".replace(",", "."))
card3.metric("Achados", f"{findings:,}".replace(",", "."))
card4.metric("Conformidade", f"{compliance:.2f}%")

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
    st.dataframe(summary, use_container_width=True, hide_index=True)

with comparison_tab:
    col1, col2, col3 = st.columns([2, 2, 3])
    sources = sorted(comparison["SOURCE"].dropna().astype(str).unique().tolist())
    statuses = sorted(comparison["STATUS"].dropna().astype(str).unique().tolist())
    with col1:
        selected_sources = st.multiselect("Fonte", sources, default=sources)
    with col2:
        selected_statuses = st.multiselect("Status", statuses, default=statuses)
    with col3:
        search_text = st.text_input("Buscar ENTITY ou KPI_CODE")

    filtered = comparison[
        comparison["SOURCE"].isin(selected_sources)
        & comparison["STATUS"].isin(selected_statuses)
    ].copy()

    if search_text.strip():
        query = search_text.strip()
        filtered = filtered[
            filtered["ENTITY"].astype(str).str.contains(
                query, case=False, na=False, regex=False
            )
            | filtered["KPI_CODE"].astype(str).str.contains(
                query, case=False, na=False, regex=False
            )
        ]

    st.caption(f"{len(filtered):,} linha(s) exibida(s)".replace(",", "."))
    st.dataframe(filtered, use_container_width=True, hide_index=True, height=620)

with diagnostic_tab:
    st.subheader("Arquivo enviado")
    uploaded_diagnostic = pd.DataFrame(
        [
            {
                "ABA_ESPERADA": name,
                "ABA_LOCALIZADA": payload["uploaded_sheets"].get(name),
                "LINHA_CABECALHO_EXCEL": payload["headers"].get(name),
                "COLUNA_ENTITY": payload["maps"].get(name, {}).get("entity"),
                "COLUNA_KPI_CODE": payload["maps"].get(name, {}).get("kpi"),
                "COLUNA_VALUE": payload["maps"].get(name, {}).get("value"),
                "COLUNA_KPI_NAME": payload["maps"].get(name, {}).get("kpi_name"),
            }
            for name in ["LOGS", "BOPS", "SL"]
        ]
    )
    st.dataframe(uploaded_diagnostic, use_container_width=True, hide_index=True)

    st.subheader("Arquivo de referência")
    reference_diagnostic = pd.DataFrame(
        [
            {
                "ABA_ESPERADA": name,
                "ABA_LOCALIZADA": payload["reference_sheets"].get(name),
                "LINHA_CABECALHO_EXCEL": payload["headers"].get(name),
                "COLUNA_ENTITY": payload["maps"].get(name, {}).get("entity"),
                "COLUNA_KPI_CODE": payload["maps"].get(name, {}).get("kpi"),
                "COLUNA_VALUE": payload["maps"].get(name, {}).get("value"),
                "COLUNA_KPI_NAME": payload["maps"].get(name, {}).get("kpi_name"),
                "COLUNA_FORMULA": payload["maps"].get(name, {}).get("formula"),
                "COLUNA_UOM": payload["maps"].get(name, {}).get("uom"),
                "COLUNA_OWNER": payload["maps"].get(name, {}).get("owner"),
            }
            for name in ["SHAREPOINT", "DEFINITION BOOK"]
        ]
    )
    st.dataframe(reference_diagnostic, use_container_width=True, hide_index=True)

    with st.expander("Abas disponíveis"):
        st.write("Arquivo enviado", payload.get("uploaded_sheet_names", []))
        st.write("Arquivo de referência", payload.get("reference_sheet_names", []))

with export_tab:
    st.write(
        "O relatório contém Comparacao_Completa, OK, Achados, Resumo, "
        "Definition_Book e Parametros."
    )
    report_bytes = report(payload)
    st.download_button(
        "Baixar relatório Excel",
        data=report_bytes,
        file_name=f"KPI_Entity_KPI_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
