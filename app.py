from datetime import datetime
import re

import pandas as pd
import streamlit as st

from engine import (
    STATUS_ORDER,
    blank_template,
    carregar_kpis_entidades,
    carregar_opcoes_template,
    gerar_template_do_consolidador,
    report,
    run,
)

st.set_page_config(page_title="KPI Validator", page_icon="📊", layout="wide")
st.title("KPI Validator | SHAREPOINT x ANAPLAN")
st.caption("Validação das abas SHAREPOINT e ANAPLAN e geração de template direto do Consolidador 2.0.xlsm.")

with st.sidebar:
    st.subheader("Configurações")
    absolute_tolerance = st.number_input("Tolerância absoluta", min_value=0.0, value=0.01, format="%.6f")
    relative_tolerance = st.number_input("Tolerância relativa", min_value=0.0, value=0.0001, format="%.6f")
    st.divider()
    include_bops = st.checkbox("Incluir BOPS", value=True)
    include_sl = st.checkbox("Incluir SL", value=True)

st.subheader("Gerador de Template KPI")
fonte_consolidador = st.file_uploader(
    "Consolidador de origem",
    type=["xlsm", "xlsx"],
    key="consolidador_origem",
    help="Opcional. Se não enviar, o app usa o arquivo Consolidador 2.0.xlsm incluído no repositório.",
)
fonte = fonte_consolidador.getvalue() if fonte_consolidador else None

col_vazio, col_preenchido = st.columns([1, 3])
with col_vazio:
    st.download_button(
        "Baixar template vazio",
        data=blank_template(),
        file_name="Template_KPI_Vazio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

try:
    opcoes_filtro = carregar_opcoes_template(fonte)
except Exception as erro:
    st.error(f"Não foi possível ler o consolidador: {erro}")
    opcoes_filtro = []

with col_preenchido:
    filtro_template = st.selectbox(
        "Opção da coluna R da aba SHAREPOINT",
        options=[""] + opcoes_filtro,
        index=0,
        key="filtro_template",
    )

if filtro_template:
    try:
        kpis_disponiveis, entidades_disponiveis = carregar_kpis_entidades(filtro_template, fonte)
        chave_contexto = f"{filtro_template}|{len(kpis_disponiveis)}|{len(entidades_disponiveis)}"
        if st.session_state.get("contexto_template") != chave_contexto:
            st.session_state["contexto_template"] = chave_contexto
            st.session_state["kpis_template"] = kpis_disponiveis.copy()
            st.session_state["entidades_template"] = entidades_disponiveis.copy()

        col_kpi, col_entidade = st.columns(2)
        with col_kpi:
            botoes_kpi = st.columns(2)
            if botoes_kpi[0].button("Marcar todos os KPIs", use_container_width=True):
                st.session_state["kpis_template"] = kpis_disponiveis.copy()
                st.rerun()
            if botoes_kpi[1].button("Limpar KPIs", use_container_width=True):
                st.session_state["kpis_template"] = []
                st.rerun()
            kpis_selecionados = st.multiselect("KPIs", options=kpis_disponiveis, key="kpis_template", placeholder="Selecione os KPIs")

        with col_entidade:
            botoes_entidade = st.columns(2)
            if botoes_entidade[0].button("Marcar todas as entidades", use_container_width=True):
                st.session_state["entidades_template"] = entidades_disponiveis.copy()
                st.rerun()
            if botoes_entidade[1].button("Limpar entidades", use_container_width=True):
                st.session_state["entidades_template"] = []
                st.rerun()
            entidades_selecionadas = st.multiselect("Entidades", options=entidades_disponiveis, key="entidades_template", placeholder="Selecione as entidades")

        st.caption(f"{len(kpis_selecionados)} KPI(s) e {len(entidades_selecionadas)} entidade(s) selecionados.")

        if kpis_selecionados and entidades_selecionadas:
            template_gerado = gerar_template_do_consolidador(
                filtro_selecionado=filtro_template,
                kpis_selecionados=kpis_selecionados,
                entidades_selecionadas=entidades_selecionadas,
                fonte=fonte,
            )
            nome_filtro = re.sub(r"[^A-Za-z0-9_-]+", "_", filtro_template).strip("_") or "FILTRO"
            st.download_button(
                "Baixar template pré-preenchido",
                data=template_gerado["arquivo"],
                file_name=f"Template_KPI_{nome_filtro}_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
            m1, m2, m3 = st.columns(3)
            m1.metric("SHAREPOINT", template_gerado["quantidade_sharepoint"])
            m2.metric("ANAPLAN", template_gerado["quantidade_anaplan"])
            m3.metric("DEFINITION BOOK", template_gerado["quantidade_definition"])
    except Exception as erro:
        st.error(f"Não foi possível gerar o template: {erro}")

st.divider()
st.subheader("Validar arquivo")
uploaded = st.file_uploader("Arquivo Excel para validação", type=["xlsx", "xlsm"], key="arquivo_validacao")
if not uploaded:
    st.info("Gere um template acima ou selecione um arquivo com as abas SHAREPOINT e ANAPLAN.")
    st.stop()

with st.spinner("Processando SHAREPOINT e ANAPLAN..."):
    payload = run(uploaded.getvalue(), absolute_tolerance, relative_tolerance)
if payload.get("error"):
    st.error(payload["error"])
    st.stop()

data = pd.concat(payload["results"], ignore_index=True)
if not include_bops and "TIPO" in data.columns:
    data = data[data["TIPO"].astype(str).str.upper() != "BOPS"]
if not include_sl and "TIPO" in data.columns:
    data = data[data["TIPO"].astype(str).str.upper() != "SL"]

countries = sorted(value for value in data["COUNTRY"].dropna().unique() if str(value).strip())
operations = sorted(value for value in data["BUSINESS_OPERATION"].dropna().unique() if str(value).strip())
filter_country, filter_operation = st.columns(2)
selected_countries = filter_country.multiselect("País", countries, default=countries, placeholder="Todos os países")
selected_operations = filter_operation.multiselect("Business Operation", operations, default=operations, placeholder="Todas as operações")
if countries:
    data = data[data["COUNTRY"].isin(selected_countries)]
if operations:
    data = data[data["BUSINESS_OPERATION"].isin(selected_operations)]

st.markdown("""
**Legenda dos status**
- 🟢 **OK:** valores dentro da tolerância.
- 🔴 **DIVERGENTE:** os valores de SHAREPOINT e ANAPLAN são diferentes.
- 🟠 **NAO ESTA NO SHAREPOINT:** a chave existe somente no ANAPLAN.
- 🟡 **SOMENTE NO SHAREPOINT:** a chave existe somente no SHAREPOINT.
- ⚪ **VALOR INVALIDO:** não foi possível comparar os valores.
""")

comparable = data["STATUS"].isin(["OK", "DIVERGENTE"])
ok_count = int(data["STATUS"].eq("OK").sum())
divergent_count = int(data["STATUS"].eq("DIVERGENTE").sum())
comparable_count = int(comparable.sum())
metric1, metric2, metric3, metric4 = st.columns(4)
metric1.metric("Comparáveis", comparable_count)
metric2.metric("OK", ok_count)
metric3.metric("Divergências", divergent_count)
metric4.metric("Conformidade", f"{ok_count / comparable_count * 100:.2f}%" if comparable_count else "0.00%")

summary_tab, comparison_tab, export_tab = st.tabs(["Resumo", "Comparação", "Exportar"])
with summary_tab:
    summary = data.groupby(["COUNTRY", "BUSINESS_OPERATION", "STATUS"], dropna=False).size().reset_index(name="QUANTIDADE")
    st.dataframe(summary, use_container_width=True, hide_index=True)
with comparison_tab:
    available_statuses = [status for status in STATUS_ORDER if status in data["STATUS"].unique()]
    selected_statuses = st.multiselect("Status", available_statuses, default=available_statuses)
    search = st.text_input("Buscar Entity ou KPI Code")
    view = data[data["STATUS"].isin(selected_statuses)]
    if search:
        view = view[
            view["ENTITY"].astype("string").str.contains(search, case=False, na=False, regex=False)
            | view["KPI_CODE"].astype("string").str.contains(search, case=False, na=False, regex=False)
        ]
    st.dataframe(view, use_container_width=True, hide_index=True, height=620)
with export_tab:
    st.download_button(
        "Baixar Excel",
        data=report(payload, data),
        file_name=f"KPI_Validator_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
