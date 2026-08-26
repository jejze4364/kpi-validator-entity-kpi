"""
KPI Validator - Interface Streamlit
=====================================

Compara automaticamente os KPIs de um arquivo enviado pelo usuario (abas
LOGS, BOPS e SL) contra o arquivo de referencia `consolidador.xlsm`
(aba SHAREPOINT, com enriquecimento opcional da aba DEFINITION BOOK),
usando exclusivamente a chave ENTITY + KPI_CODE.

Toda a logica de leitura/normalizacao/comparacao vive em `engine.py`.
"""

from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from engine import run, report, load_reference

REFERENCE_FILENAME = "consolidador.xlsm"
REFERENCE_PATH = Path(__file__).parent / REFERENCE_FILENAME

st.set_page_config(page_title="KPI Validator", page_icon="📊", layout="wide")
st.title("KPI Validator | Chave: ENTITY + KPI_CODE")
st.caption(
    "Compara automaticamente os KPIs enviados (LOGS, BOPS, SL) contra a base "
    "de referência consolidador.xlsm (aba SHAREPOINT), com enriquecimento "
    "opcional pela aba DEFINITION BOOK."
)


# ---------------------------------------------------------------------------
# Carregamento do arquivo de referencia (cacheado)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def get_reference_bundle(reference_bytes: bytes, filename: str):
    return load_reference(reference_bytes, filename)


if not REFERENCE_PATH.exists():
    st.error(
        f"Arquivo de referência **{REFERENCE_FILENAME}** não encontrado na pasta do "
        "aplicativo. Posicione o arquivo ao lado de `app.py` no repositório antes "
        "de publicar (veja o README para instruções)."
    )
    st.stop()

reference_bundle = get_reference_bundle(REFERENCE_PATH.read_bytes(), REFERENCE_FILENAME)

if reference_bundle.get("error"):
    st.error(f"Falha ao carregar o arquivo de referência: {reference_bundle['error']}")
    with st.expander("Diagnóstico da referência"):
        st.write("Abas localizadas:", reference_bundle.get("resolved"))
    st.stop()


# ---------------------------------------------------------------------------
# Barra lateral - parametros
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("Parâmetros")
    abs_tol = st.number_input("Tolerância absoluta", min_value=0.0, value=0.01, step=0.01, format="%.4f")
    rel_tol = st.number_input("Tolerância relativa", min_value=0.0, value=0.0001, step=0.0001, format="%.6f")
    st.divider()
    st.caption(f"Referência carregada: **{REFERENCE_FILENAME}**")
    n_kpis_ref = reference_bundle["reference_df"]["KPI_CODE"].nunique()
    n_entities_ref = reference_bundle["reference_df"]["ENTITY"].nunique()
    st.caption(f"{n_kpis_ref} KPIs · {n_entities_ref} entidades na referência")


# ---------------------------------------------------------------------------
# Upload do arquivo a validar
# ---------------------------------------------------------------------------

uploaded_file = st.file_uploader(
    "Envie o arquivo a validar (deve conter as abas LOGS, BOPS e SL)",
    type=["xlsx", "xlsm"],
)

if not uploaded_file:
    st.info("Aguardando o upload do arquivo com as abas **LOGS**, **BOPS** e **SL**.")
    st.stop()

with st.spinner("Processando…"):
    payload = run(uploaded_file.getvalue(), reference_bundle, abs_tol, rel_tol)

if payload.get("error"):
    st.error(payload["error"])
    with st.expander("Diagnóstico"):
        st.write("Abas localizadas no arquivo enviado:", payload.get("resolved"))
        if payload.get("mapping"):
            st.write("Colunas mapeadas:", payload.get("mapping"))
    st.stop()


# ---------------------------------------------------------------------------
# Resultado consolidado
# ---------------------------------------------------------------------------

combined = pd.concat(payload["results"], ignore_index=True)
ok_count = int(combined["STATUS"].eq("OK").sum())
total = len(combined)
findings = total - ok_count
compliance = f"{ok_count / total * 100:.2f}%" if total else "0%"

c1, c2, c3, c4 = st.columns(4)
c1.metric("Comparados", total)
c2.metric("OK", ok_count)
c3.metric("Achados", findings)
c4.metric("Conformidade", compliance)

tab_summary, tab_detail, tab_diagnostics, tab_export = st.tabs(
    ["Resumo", "Comparação", "Diagnóstico", "Exportar"]
)

with tab_summary:
    summary = (
        combined.groupby(["SOURCE", "STATUS"]).size().reset_index(name="QUANTIDADE")
    )
    st.dataframe(summary, use_container_width=True, hide_index=True)

with tab_detail:
    status_options = sorted(combined["STATUS"].unique())
    selected_status = st.multiselect("Status", status_options, default=status_options)
    search_term = st.text_input("Buscar por Entity ou KPI_CODE")

    filtered = combined[combined["STATUS"].isin(selected_status)]
    if search_term:
        filtered = filtered[
            filtered["ENTITY"].str.contains(search_term, case=False, na=False)
            | filtered["KPI_CODE"].str.contains(search_term, case=False, na=False)
        ]
    st.dataframe(filtered, use_container_width=True, hide_index=True, height=620)

with tab_diagnostics:
    st.subheader("Referência (consolidador.xlsm)")
    st.write("Abas localizadas:", reference_bundle.get("resolved"))
    st.write("Colunas mapeadas:", reference_bundle.get("mapping"))

    st.subheader("Arquivo enviado")
    st.write("Abas localizadas:", payload.get("resolved"))
    st.write("Linha de cabeçalho detectada (0-indexed):", payload.get("headers"))
    st.write("Colunas mapeadas:", payload.get("mapping"))

with tab_export:
    st.write(
        "O relatório contém as abas: Comparacao_Completa, OK, Achados, Resumo, "
        "Definition_Book e Parametros."
    )
    st.download_button(
        "Baixar relatório Excel",
        report(payload),
        f"KPI_Validator_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        type="primary",
    )
