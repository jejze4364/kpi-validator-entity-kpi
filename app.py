from datetime import datetime
from pathlib import Path
import streamlit as st
import engine

st.set_page_config(page_title="KPI Validator", page_icon="📊", layout="wide")

REFERENCE_PATH = Path(__file__).parent / "consolidador.xlsm"

@st.cache_data(show_spinner=False)
def load_reference_cached(data, filename, modified_at):
    return engine.load_reference(data, filename, modified_at)

@st.cache_data(show_spinner=False)
def generate_template_cached(reference_bundle, responsible):
    return engine.generate_responsible_template(reference_bundle, responsible)

@st.cache_data(show_spinner=False)
def generate_all_templates_cached(reference_bundle):
    return engine.generate_all_templates(reference_bundle)

@st.cache_data(show_spinner=False)
def generate_catalog_cached(reference_bundle):
    return engine.generate_classification_catalog(reference_bundle)

st.title("KPI Validator | ENTITY + KPI_CODE")
st.caption("Consulta da referência, geração de templates e validação automática de LOGS, BOPS e SL.")

if not REFERENCE_PATH.exists():
    st.error("Arquivo de referência não encontrado. Posicione consolidador.xlsm na mesma pasta de app.py.")
    st.stop()

try:
    reference_bytes = REFERENCE_PATH.read_bytes()
    reference_modified = datetime.fromtimestamp(REFERENCE_PATH.stat().st_mtime)
    reference = load_reference_cached(reference_bytes, REFERENCE_PATH.name, reference_modified.isoformat())
except Exception as exc:
    st.error(f"Não foi possível carregar o arquivo de referência: {exc}")
    st.stop()

summary = engine.get_reference_summary(reference)
catalog = engine.get_kpi_catalog(reference)
responsibles = engine.get_responsibles(reference)
classification_data = engine.get_classifications(reference)

pages = st.tabs(["Início", "Classificações", "Templates", "Central de Downloads", "Validação", "Resultados", "Diagnóstico", "Exportar"])

with pages[0]:
    st.success("Arquivo de referência carregado com sucesso.")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de KPIs", summary["total_kpis"])
    c2.metric("Total de entidades", summary["total_entities"])
    c3.metric("Total de responsáveis", summary["total_responsibles"])
    c4.metric("Total de classificações", summary["total_classifications"])
    c5, c6, c7, c8, c9 = st.columns(5)
    c5.metric("KPIs de LOGS", summary["logs_kpis"])
    c6.metric("KPIs de BOPS", summary["bops_kpis"])
    c7.metric("KPIs de SL", summary["sl_kpis"])
    c8.metric("Sem responsável", summary["without_owner"])
    c9.metric("Sem classificação", summary["without_classification"])
    st.subheader("Arquivo de referência")
    st.write(f"Nome: {reference['filename']}")
    st.write(f"Tamanho: {len(reference_bytes) / 1024 / 1024:.2f} MB")
    st.write(f"Modificação: {reference_modified:%d/%m/%Y %H:%M:%S}")
    st.write(f"Aba SHAREPOINT: {reference['resolved'].get('SHAREPOINT') or 'não localizada'}")
    st.write(f"Aba DEFINITION BOOK: {reference['resolved'].get('DEFINITION BOOK') or 'não localizada'}")
    st.write(f"Registros padronizados da referência: {len(reference['reference_data'])}")

with pages[1]:
    st.subheader("Classificações e catálogo de KPIs")
    if catalog.empty:
        st.info("Nenhum registro disponível para consulta.")
    else:
        filtered = catalog.copy()
        filter_columns = [
            ("CLASSIFICATION", "Classificação"),
            ("OWNER", "Responsável"),
            ("SOURCE", "SOURCE"),
            ("ENTITY", "ENTITY"),
            ("KPI_CODE", "KPI_CODE"),
            ("KPI_NAME", "Nome do KPI"),
            ("UNIT_OF_MEASURE", "Unidade de medida")
        ]
        selections = {}
        available_filters = [(column, label) for column, label in filter_columns if column in filtered.columns and filtered[column].fillna("").astype(str).str.strip().ne("").any()]
        filter_boxes = st.columns(min(3, max(1, len(available_filters))))
        for index, (column, label) in enumerate(available_filters):
            options = sorted(filtered[column].fillna("").astype(str).loc[lambda s: s.str.strip().ne("")].unique().tolist())
            selections[column] = filter_boxes[index % len(filter_boxes)].multiselect(label, options)
        query = st.text_input("Buscar na tabela", placeholder="ENTITY, KPI_CODE, nome, responsável ou classificação")
        for column, selected in selections.items():
            if selected:
                filtered = filtered[filtered[column].isin(selected)]
        if query:
            searchable = filtered.astype(str).apply(lambda column: column.str.contains(query, case=False, na=False, regex=False))
            filtered = filtered[searchable.any(axis=1)]
        st.caption(f"{len(filtered)} registro(s) exibido(s).")
        st.dataframe(filtered, use_container_width=True, hide_index=True, height=620)
    if not classification_data["available"]:
        st.warning("A informação de classificação não foi encontrada no consolidador.")

with pages[2]:
    st.subheader("Templates padrões por responsável")
    owner_options = responsibles["display_name"].tolist() if not responsibles.empty else []
    selected_owner = st.selectbox("Responsável", owner_options, index=None, placeholder="Selecione um responsável")
    if selected_owner:
        template_bytes, template_name = generate_template_cached(reference, selected_owner)
        st.download_button("Baixar template do responsável", template_bytes, template_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    if owner_options:
        all_templates = generate_all_templates_cached(reference)
        st.download_button("Baixar todos os templates", all_templates, "templates_kpi.zip", "application/zip")
    else:
        st.info("Nenhum responsável foi identificado na referência.")

with pages[3]:
    st.subheader("Central de Downloads")
    d1, d2, d3 = st.columns(3)
    d1.download_button("Baixar consolidador completo", engine.get_reference_file_download(reference_bytes), reference["filename"], "application/vnd.ms-excel.sheet.macroEnabled.12")
    classification_catalog = generate_catalog_cached(reference)
    d2.download_button("Baixar catálogo de classificações", classification_catalog, "catalogo_classificacoes_kpi.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d3.download_button("Baixar relação completa de KPIs", engine.dataframe_to_excel(catalog, "KPIs"), "relacao_completa_kpis.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d4, d5, d6 = st.columns(3)
    if reference["definitions"].empty:
        d4.info("Definition Book não disponível.")
    else:
        d4.download_button("Baixar Definition Book", engine.dataframe_to_excel(reference["definitions"], "Definition_Book"), "definition_book.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d5.download_button("Baixar lista de responsáveis", engine.dataframe_to_excel(responsibles, "Responsaveis"), "responsaveis_kpi.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if owner_options:
        d6.download_button("Baixar todos os templates", generate_all_templates_cached(reference), "templates_kpi.zip", "application/zip")

with st.sidebar:
    st.header("Parâmetros de validação")
    abs_tol = st.number_input("Tolerância absoluta", min_value=0.0, value=0.01, format="%.8f")
    rel_tol = st.number_input("Tolerância relativa", min_value=0.0, value=0.0001, format="%.8f")

with pages[4]:
    st.subheader("Validação")
    uploaded = st.file_uploader("Envie o arquivo preenchido", type=["xlsx", "xlsm"])
    if uploaded is None:
        st.info("Envie o arquivo preenchido para iniciar a validação dos KPIs. As consultas e os downloads da base de referência já estão disponíveis.")
    else:
        signature = (uploaded.name, uploaded.size, abs_tol, rel_tol, reference["fingerprint"])
        if st.session_state.get("validation_signature") != signature:
            with st.spinner("Processando a validação..."):
                st.session_state["validation_payload"] = engine.run(uploaded.getvalue(), reference, abs_tol, rel_tol, uploaded.name)
                st.session_state["validation_signature"] = signature
        payload = st.session_state["validation_payload"]
        if payload.get("error"):
            st.error(payload["error"])
            if payload.get("missing_sheets"):
                st.write("Abas faltantes:", ", ".join(payload["missing_sheets"]))
            st.json({"abas_localizadas": payload.get("resolved", {}), "mapeamentos": payload.get("maps", {})})
        else:
            st.success("Validação concluída. Consulte as abas Resultados, Diagnóstico e Exportar.")

payload = st.session_state.get("validation_payload")
valid_payload = payload if payload and not payload.get("error") else None

with pages[5]:
    if not valid_payload:
        st.info("Envie o arquivo preenchido na aba Validação para visualizar os resultados.")
    else:
        result = valid_payload["combined"]
        ok_count = int(result["STATUS"].eq("OK").sum())
        compared = len(result)
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Total comparado", compared)
        r2.metric("OK", ok_count)
        r3.metric("Achados", compared - ok_count)
        r4.metric("Conformidade", f"{ok_count / compared * 100:.2f}%" if compared else "0.00%")
        st.dataframe(valid_payload["summary"], use_container_width=True, hide_index=True)
        statuses = sorted(result["STATUS"].dropna().unique().tolist())
        selected_statuses = st.multiselect("Status", statuses, default=statuses)
        text_search = st.text_input("Buscar ENTITY ou KPI_CODE", key="result_search")
        view = result[result["STATUS"].isin(selected_statuses)]
        if text_search:
            view = view[view["ENTITY"].str.contains(text_search, case=False, na=False, regex=False) | view["KPI_CODE"].str.contains(text_search, case=False, na=False, regex=False)]
        st.dataframe(view, use_container_width=True, hide_index=True, height=620)

with pages[6]:
    st.subheader("Diagnóstico")
    st.markdown("#### Referência")
    st.json({"abas_localizadas": reference["resolved"], "cabecalhos": reference["heads"], "mapeamentos": reference["maps"]})
    if valid_payload:
        st.markdown("#### Arquivo enviado")
        st.json({"abas_localizadas": valid_payload["resolved"], "cabecalhos": valid_payload["heads"], "mapeamentos": valid_payload["maps"]})
    else:
        st.info("O diagnóstico do arquivo enviado ficará disponível após a validação.")

with pages[7]:
    if not valid_payload:
        st.info("Execute uma validação para gerar o relatório final.")
    else:
        filename = f"KPI_Entity_KPI_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.download_button("Baixar relatório Excel", engine.report(valid_payload), filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
