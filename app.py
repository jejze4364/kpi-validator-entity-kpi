from datetime import datetime
import pandas as pd
import streamlit as st
from engine import STATUS_ORDER, report, run, blank_template


st.set_page_config(page_title="KPI Validator", page_icon="📊", layout="wide")
st.title("KPI Validator | SHAREPOINT x ANAPLAN")
st.caption("O aplicativo le exclusivamente as abas SHAREPOINT e ANAPLAN.")
st.download_button(
    "📥 Baixar Template Padrão",
    data=blank_template(),
    file_name="Template_KPI_Padrao.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary"
)
with st.sidebar:
    st.subheader("Configuracoes")
    absolute_tolerance = st.number_input("Tolerancia absoluta", min_value=0.0, value=0.01, format="%.6f")
    relative_tolerance = st.number_input("Tolerancia relativa", min_value=0.0, value=0.0001, format="%.6f")
    st.divider()
    include_bops = st.checkbox("Incluir BOPS", value=True)
    include_sl = st.checkbox("Incluir SL", value=True)

uploaded = st.file_uploader("Arquivo Excel", type=["xlsx", "xlsm"])
if not uploaded:
    st.info("Selecione o arquivo com as abas SHAREPOINT e ANAPLAN.")
    st.stop()

with st.spinner("Processando SHAREPOINT e ANAPLAN..."):
    payload = run(uploaded.getvalue(), absolute_tolerance, relative_tolerance)
if payload.get("error"):
    st.error(payload["error"])
    st.stop()

data = pd.concat(payload["results"], ignore_index=True)
if not include_bops:
    data = data[data["TIPO"] != "BOPS"]
if not include_sl:
    data = data[data["TIPO"] != "SL"]

countries = sorted(value for value in data["COUNTRY"].dropna().unique() if str(value).strip())
operations = sorted(value for value in data["BUSINESS_OPERATION"].dropna().unique() if str(value).strip())
filter_country, filter_operation = st.columns(2)
selected_countries = filter_country.multiselect("Pais", countries, default=countries, placeholder="Todos os paises")
selected_operations = filter_operation.multiselect("Business Operation", operations, default=operations, placeholder="Todas as operacoes")
if countries:
    data = data[data["COUNTRY"].isin(selected_countries)]
if operations:
    data = data[data["BUSINESS_OPERATION"].isin(selected_operations)]

st.markdown("""
**Legenda dos status**
- 🟢 **OK:** valores dentro da tolerancia.
- 🔴 **DIVERGENTE:** os valores de SHAREPOINT e ANAPLAN sao diferentes.
- 🟠 **NAO ESTA NO SHAREPOINT:** a chave existe somente no ANAPLAN.
- 🟡 **SOMENTE NO SHAREPOINT:** a chave existe somente no SHAREPOINT.
- ⚪ **VALOR INVALIDO:** nao foi possivel comparar os valores.
""")

comparable = data["STATUS"].isin(["OK", "DIVERGENTE"])
ok_count = int(data["STATUS"].eq("OK").sum())
divergent_count = int(data["STATUS"].eq("DIVERGENTE").sum())
comparable_count = int(comparable.sum())
metric1, metric2, metric3, metric4 = st.columns(4)
metric1.metric("Comparaveis", comparable_count)
metric2.metric("OK", ok_count)
metric3.metric("Divergencias", divergent_count)
metric4.metric("Conformidade", f"{ok_count / comparable_count * 100:.2f}%" if comparable_count else "0.00%")

summary_tab, comparison_tab, export_tab = st.tabs(["Resumo", "Comparacao", "Exportar"])
with summary_tab:
    summary = data.groupby(["COUNTRY", "BUSINESS_OPERATION", "STATUS"], dropna=False).size().reset_index(name="QUANTIDADE")
    st.dataframe(summary, use_container_width=True, hide_index=True)
with comparison_tab:
    available_statuses = [status for status in STATUS_ORDER if status in data["STATUS"].unique()]
    selected_statuses = st.multiselect("Status", available_statuses, default=available_statuses)
    search = st.text_input("Buscar Entity ou KPI Code")
    view = data[data["STATUS"].isin(selected_statuses)]
    if search:
        view = view[view["ENTITY"].str.contains(search, case=False, na=False, regex=False) | view["KPI_CODE"].str.contains(search, case=False, na=False, regex=False)]
    st.dataframe(view, use_container_width=True, hide_index=True, height=620)
with export_tab:
    st.download_button(
        "Baixar Excel",
        report(payload, data),
        f"KPI_Validator_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
