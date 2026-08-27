import hashlib
import io
import re
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

REFERENCE_SHEETS = ["SHAREPOINT", "DEFINITION BOOK"]
REPORT_SHEETS = ["LOGS", "BOPS", "SL"]
TEMPLATE_VERSION = "2.0"

ALIASES = {
    "entity": ["entity", "entidade", "unit name", "unidade", "bu", "location", "plant", "dc"],
    "kpi": ["kpi code", "kpi_code", "codigo kpi", "código kpi", "cod kpi"],
    "value": ["current_value ac", "current value ac", "current value", "valor anaplan", "valor origem", "actual", "ac", "value", "valor"],
    "kpi_name": ["kpi name", "nome kpi", "description", "descrição", "descricao"],
    "formula": ["formula", "fórmula", "regra"],
    "uom": ["unit_of_measure", "unit of measure", "uom", "unidade de medida"],
    "owner": ["owner", "responsavel", "responsável"],
    "classification": ["classificacao", "classificação", "classification", "class", "categoria", "category"],
    "source": ["source", "origem", "grupo", "group", "processo", "process"]
}

STANDARD_COLUMNS = ["SOURCE", "ENTITY", "KPI_CODE", "KPI_NAME", "CLASSIFICATION", "FORMULA", "UNIT_OF_MEASURE", "OWNER", "VALUE"]


def noacc(value):
    return "".join(character for character in unicodedata.normalize("NFKD", str(value)) if not unicodedata.combining(character))


def norm(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", noacc(value).strip().upper())


def nc(value):
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", noacc(value).strip().lower())).strip()


def nkpi(value):
    text = norm(value)
    match = re.search(r"\b([A-Z]{2,3}-[KR]\d{3,5}(?:_\d{4})?)\b", text)
    return match.group(1) if match else text


def safe_filename(value):
    text = noacc(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text[:100] or "sem_responsavel"


def locate_sheet(names, target):
    exact = [name for name in names if norm(name) == norm(target)]
    if exact:
        return exact[0]
    aliases = {
        "SHAREPOINT": ["SHAREPOINT", "SHARE POINT", "BASE SHAREPOINT", "CONSOLIDADO SHAREPOINT"],
        "DEFINITION BOOK": ["DEFINITION BOOK", "DEFINITIONBOOK", "DEFINITIONS", "DEFINITION"],
        "LOGS": ["LOGS", "LOGISTICS", "LOGISTICA", "LOGÍSTICA"],
        "BOPS": ["BOPS"],
        "SL": ["SL", "SERVICE LEVEL", "SERVICELEVEL"]
    }
    for alias in aliases.get(target, [target]):
        partial = [name for name in names if norm(alias) in norm(name)]
        if partial:
            return partial[0]
    return None


def detect_header(raw):
    best_row = 0
    best_score = -1
    for index in range(min(80, len(raw))):
        values = [norm(value) for value in raw.iloc[index]]
        has_kpi = any("KPI" in value and ("CODE" in value or "CODIGO" in value or "COD" in value) for value in values)
        has_entity = any(value in {"ENTITY", "ENTIDADE", "UNIT NAME", "UNIDADE", "BU", "LOCATION", "PLANT", "DC"} for value in values)
        has_value = any(nc(value) in {nc(alias) for alias in ALIASES["value"]} for value in values)
        score = 1000 * has_kpi + 800 * has_entity + 400 * has_value + sum(bool(value) for value in values)
        if score > best_score:
            best_row = index
            best_score = score
    return best_row


def read_sheet(data, sheet_name):
    raw = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name, header=None, engine="openpyxl")
    header_index = detect_header(raw)
    columns = []
    seen = {}
    for position, value in enumerate(raw.iloc[header_index]):
        column = str(value).strip() if not pd.isna(value) else f"COL_{position + 1}"
        seen[column] = seen.get(column, 0) + 1
        if seen[column] > 1:
            column = f"{column}_{seen[column]}"
        columns.append(column)
    frame = raw.iloc[header_index + 1:].copy()
    frame.columns = columns
    frame = frame.dropna(how="all").reset_index(drop=True)
    return frame, header_index + 1


def find_column(frame, role):
    normalized_columns = {nc(column): column for column in frame.columns}
    for alias in ALIASES[role]:
        if nc(alias) in normalized_columns:
            return normalized_columns[nc(alias)]
    candidates = []
    for normalized, original in normalized_columns.items():
        for alias in ALIASES[role]:
            alias_normalized = nc(alias)
            if alias_normalized and (alias_normalized in normalized or normalized in alias_normalized):
                candidates.append((abs(len(normalized) - len(alias_normalized)), original))
    return sorted(candidates)[0][1] if candidates else None


def create_mapping(frame, source):
    mapped = {role: find_column(frame, role) for role in ALIASES}
    normalized_columns = {nc(column): column for column in frame.columns}
    preferred_values = ["current value ac", "current_value ac", "current value", "ac", "value", "valor"] if source == "SHAREPOINT" else ["valor anaplan", "valor origem", "actual", "ac", "value", "valor"]
    for preferred in preferred_values:
        if nc(preferred) in normalized_columns:
            mapped["value"] = normalized_columns[nc(preferred)]
            break
    return mapped


def to_number(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    text = series.astype(str).str.strip().replace({"": np.nan, "-": np.nan, "nan": np.nan, "None": np.nan})
    text = text.str.replace(r"\s+", "", regex=True)
    negative_parentheses = text.str.match(r"^\(.*\)$", na=False)
    text.loc[negative_parentheses] = "-" + text.loc[negative_parentheses].str[1:-1]
    brazilian = text.str.match(r"^-?\d{1,3}(?:\.\d{3})+(?:,\d+)?$", na=False) | text.str.match(r"^-?\d+,\d+$", na=False)
    text.loc[brazilian] = text.loc[brazilian].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    text.loc[~brazilian] = text.loc[~brazilian].str.replace(",", "", regex=False)
    return pd.to_numeric(text, errors="coerce")


def column_or_blank(frame, column):
    if column:
        return frame[column].fillna("").astype(str)
    return pd.Series([""] * len(frame), index=frame.index, dtype="object")


def standardize(frame, mapping, source, require_value=True):
    required = ["entity", "kpi"] + (["value"] if require_value else [])
    missing = [role for role in required if not mapping.get(role)]
    if missing:
        raise ValueError(f"{source}: colunas não localizadas: {', '.join(missing)}. Detectadas: {list(frame.columns)}")
    output = pd.DataFrame(index=frame.index)
    output["SOURCE"] = column_or_blank(frame, mapping.get("source")).map(norm)
    if source in REPORT_SHEETS:
        output["SOURCE"] = source
    output["ENTITY"] = frame[mapping["entity"]].map(norm)
    output["KPI_CODE"] = frame[mapping["kpi"]].map(nkpi)
    output["KPI_NAME"] = column_or_blank(frame, mapping.get("kpi_name")).str.strip()
    output["CLASSIFICATION"] = column_or_blank(frame, mapping.get("classification")).str.strip()
    output["FORMULA"] = column_or_blank(frame, mapping.get("formula")).str.strip()
    output["UNIT_OF_MEASURE"] = column_or_blank(frame, mapping.get("uom")).str.strip()
    output["OWNER"] = column_or_blank(frame, mapping.get("owner")).str.strip()
    output["VALUE"] = to_number(frame[mapping["value"]]) if mapping.get("value") else np.nan
    output = output[(output["ENTITY"] != "") & (output["KPI_CODE"] != "")].copy()
    output.loc[~output["SOURCE"].isin(REPORT_SHEETS), "SOURCE"] = output.loc[~output["SOURCE"].isin(REPORT_SHEETS), "KPI_CODE"].str.extract(r"^(LOGS|BOPS|SL)", expand=False).fillna("")
    return output.reset_index(drop=True)


def standardize_definitions(frame, mapping):
    if not mapping.get("kpi"):
        return pd.DataFrame(columns=["KPI_CODE", "DEF_KPI_NAME", "DEF_CLASSIFICATION", "DEF_FORMULA", "DEF_UOM", "DEF_OWNER", "DEF_SOURCE"])
    output = pd.DataFrame(index=frame.index)
    output["KPI_CODE"] = frame[mapping["kpi"]].map(nkpi)
    output["DEF_KPI_NAME"] = column_or_blank(frame, mapping.get("kpi_name")).str.strip()
    output["DEF_CLASSIFICATION"] = column_or_blank(frame, mapping.get("classification")).str.strip()
    output["DEF_FORMULA"] = column_or_blank(frame, mapping.get("formula")).str.strip()
    output["DEF_UOM"] = column_or_blank(frame, mapping.get("uom")).str.strip()
    output["DEF_OWNER"] = column_or_blank(frame, mapping.get("owner")).str.strip()
    output["DEF_SOURCE"] = column_or_blank(frame, mapping.get("source")).map(norm)
    output = output[output["KPI_CODE"] != ""]
    return output.drop_duplicates("KPI_CODE", keep="first").reset_index(drop=True)


def fill_reference_metadata(reference_data, definitions):
    if definitions.empty:
        return reference_data
    merged = reference_data.merge(definitions, on="KPI_CODE", how="left")
    replacements = {
        "KPI_NAME": "DEF_KPI_NAME",
        "CLASSIFICATION": "DEF_CLASSIFICATION",
        "FORMULA": "DEF_FORMULA",
        "UNIT_OF_MEASURE": "DEF_UOM",
        "OWNER": "DEF_OWNER",
        "SOURCE": "DEF_SOURCE"
    }
    for target, fallback in replacements.items():
        current = merged[target].fillna("").astype(str).str.strip()
        alternative = merged[fallback].fillna("").astype(str).str.strip()
        merged[target] = current.where(current.ne(""), alternative)
    return merged[STANDARD_COLUMNS]


def load_reference(data, filename="consolidador.xlsm", modified_at=None):
    names = pd.ExcelFile(io.BytesIO(data), engine="openpyxl").sheet_names
    resolved = {target: locate_sheet(names, target) for target in REFERENCE_SHEETS}
    if not resolved["SHAREPOINT"]:
        raise ValueError(f"Aba SHAREPOINT não encontrada no arquivo de referência. Abas disponíveis: {names}")
    raw = {}
    heads = {}
    maps = {}
    raw["SHAREPOINT"], heads["SHAREPOINT"] = read_sheet(data, resolved["SHAREPOINT"])
    maps["SHAREPOINT"] = create_mapping(raw["SHAREPOINT"], "SHAREPOINT")
    reference_data = standardize(raw["SHAREPOINT"], maps["SHAREPOINT"], "SHAREPOINT", True)
    if resolved["DEFINITION BOOK"]:
        raw["DEFINITION BOOK"], heads["DEFINITION BOOK"] = read_sheet(data, resolved["DEFINITION BOOK"])
        maps["DEFINITION BOOK"] = create_mapping(raw["DEFINITION BOOK"], "DEFINITION BOOK")
        definitions = standardize_definitions(raw["DEFINITION BOOK"], maps["DEFINITION BOOK"])
    else:
        definitions = pd.DataFrame(columns=["KPI_CODE", "DEF_KPI_NAME", "DEF_CLASSIFICATION", "DEF_FORMULA", "DEF_UOM", "DEF_OWNER", "DEF_SOURCE"])
    reference_data = fill_reference_metadata(reference_data, definitions)
    reference_data["OWNER_KEY"] = reference_data["OWNER"].map(norm)
    reference_data["CLASSIFICATION_KEY"] = reference_data["CLASSIFICATION"].map(norm)
    return {
        "filename": filename,
        "modified_at": modified_at,
        "fingerprint": hashlib.sha256(data).hexdigest(),
        "reference_bytes": data,
        "sheet_names": names,
        "resolved": resolved,
        "heads": heads,
        "maps": maps,
        "reference_data": reference_data,
        "definitions": definitions
    }


def get_kpi_catalog(reference_bundle):
    data = reference_bundle["reference_data"].copy()
    columns = [column for column in STANDARD_COLUMNS if column != "VALUE" and column in data.columns]
    return data[columns].drop_duplicates().sort_values(["SOURCE", "ENTITY", "KPI_CODE"], na_position="last").reset_index(drop=True)


def representative_names(data, key_column, display_column):
    valid = data[data[key_column] != ""].copy()
    if valid.empty:
        return pd.DataFrame(columns=["key", "display_name", "kpi_count", "entity_count"])
    frequencies = valid.groupby([key_column, display_column], dropna=False).size().reset_index(name="frequency")
    representative = frequencies.sort_values([key_column, "frequency", display_column], ascending=[True, False, True]).drop_duplicates(key_column)
    counts = valid.groupby(key_column).agg(kpi_count=("KPI_CODE", "nunique"), entity_count=("ENTITY", "nunique")).reset_index()
    output = representative.merge(counts, on=key_column, how="left")
    return output.rename(columns={key_column: "key", display_column: "display_name"})[["key", "display_name", "kpi_count", "entity_count"]].sort_values("display_name").reset_index(drop=True)


def get_responsibles(reference_bundle):
    data = reference_bundle["reference_data"]
    output = representative_names(data, "OWNER_KEY", "OWNER")
    if (data["OWNER_KEY"] == "").any():
        missing = pd.DataFrame([{"key": "", "display_name": "Sem responsável", "kpi_count": data.loc[data["OWNER_KEY"] == "", "KPI_CODE"].nunique(), "entity_count": data.loc[data["OWNER_KEY"] == "", "ENTITY"].nunique()}])
        output = pd.concat([output, missing], ignore_index=True)
    return output


def get_classifications(reference_bundle):
    data = reference_bundle["reference_data"]
    output = representative_names(data, "CLASSIFICATION_KEY", "CLASSIFICATION")
    return {"available": not output.empty, "data": output}


def get_reference_summary(reference_bundle):
    data = reference_bundle["reference_data"]
    return {
        "total_kpis": int(data["KPI_CODE"].nunique()),
        "total_entities": int(data["ENTITY"].nunique()),
        "total_responsibles": int(data.loc[data["OWNER_KEY"] != "", "OWNER_KEY"].nunique()),
        "total_classifications": int(data.loc[data["CLASSIFICATION_KEY"] != "", "CLASSIFICATION_KEY"].nunique()),
        "logs_kpis": int(data.loc[data["SOURCE"] == "LOGS", "KPI_CODE"].nunique()),
        "bops_kpis": int(data.loc[data["SOURCE"] == "BOPS", "KPI_CODE"].nunique()),
        "sl_kpis": int(data.loc[data["SOURCE"] == "SL", "KPI_CODE"].nunique()),
        "without_owner": int(data.loc[data["OWNER_KEY"] == "", "KPI_CODE"].nunique()),
        "without_classification": int(data.loc[data["CLASSIFICATION_KEY"] == "", "KPI_CODE"].nunique())
    }


def compare(report_data, reference_data, definitions, abs_tol, rel_tol, source):
    report_grouped = report_data.groupby(["ENTITY", "KPI_CODE"], as_index=False).agg(REPORT_VALUE=("VALUE", lambda values: values.sum(min_count=1)), REPORT_ROWS=("VALUE", "size"), KPI_NAME=("KPI_NAME", "first"))
    reference_source = reference_data[(reference_data["SOURCE"] == source) | (reference_data["SOURCE"] == "")].copy()
    reference_grouped = reference_source.groupby(["ENTITY", "KPI_CODE"], as_index=False).agg(SHAREPOINT_VALUE=("VALUE", lambda values: values.sum(min_count=1)), SHAREPOINT_ROWS=("VALUE", "size"), REFERENCE_KPI_NAME=("KPI_NAME", "first"), CLASSIFICATION=("CLASSIFICATION", "first"), FORMULA=("FORMULA", "first"), UNIT_OF_MEASURE=("UNIT_OF_MEASURE", "first"), OWNER=("OWNER", "first"))
    result = report_grouped.merge(reference_grouped, on=["ENTITY", "KPI_CODE"], how="outer", indicator=True)
    result["SOURCE"] = source
    result["DIFFERENCE"] = result["REPORT_VALUE"] - result["SHAREPOINT_VALUE"]
    result["DIFFERENCE_PCT"] = np.where(result["SHAREPOINT_VALUE"].abs() > abs_tol, result["DIFFERENCE"] / result["SHAREPOINT_VALUE"], np.nan)
    within_tolerance = (result["DIFFERENCE"].abs() <= abs_tol + rel_tol * result["SHAREPOINT_VALUE"].abs()).fillna(False)
    result["STATUS"] = np.select(
        [result["_merge"].eq("left_only"), result["_merge"].eq("right_only"), within_tolerance],
        ["NÃO ESTÁ NA REFERÊNCIA", "SOMENTE NA REFERÊNCIA", "OK"],
        default="DIVERGENTE"
    )
    result["KPI_NAME"] = result["KPI_NAME"].fillna("").where(result["KPI_NAME"].fillna("").astype(str).str.strip().ne(""), result["REFERENCE_KPI_NAME"].fillna(""))
    result["KEY_USED"] = "ENTITY + KPI_CODE"
    return result.drop(columns=["_merge", "REFERENCE_KPI_NAME"])


def run(upload_bytes, reference_bundle, abs_tol=0.01, rel_tol=0.0001, upload_filename="arquivo_enviado.xlsx"):
    try:
        names = pd.ExcelFile(io.BytesIO(upload_bytes), engine="openpyxl").sheet_names
    except Exception as exc:
        return {"error": f"Não foi possível abrir o arquivo enviado: {exc}"}
    resolved = {target: locate_sheet(names, target) for target in REPORT_SHEETS}
    missing = [target for target in REPORT_SHEETS if not resolved[target]]
    if missing:
        return {"error": "Abas obrigatórias não encontradas no arquivo enviado: " + ", ".join(missing), "missing_sheets": missing, "resolved": resolved, "maps": {}}
    raw = {}
    maps = {}
    heads = {}
    standardized = {}
    try:
        for source in REPORT_SHEETS:
            raw[source], heads[source] = read_sheet(upload_bytes, resolved[source])
            maps[source] = create_mapping(raw[source], source)
            standardized[source] = standardize(raw[source], maps[source], source, True)
        results = [compare(standardized[source], reference_bundle["reference_data"], reference_bundle["definitions"], abs_tol, rel_tol, source) for source in REPORT_SHEETS]
    except Exception as exc:
        return {"error": str(exc), "resolved": resolved, "maps": maps, "heads": heads}
    combined = pd.concat(results, ignore_index=True)
    summary = combined.groupby(["SOURCE", "STATUS"], dropna=False).size().reset_index(name="QUANTIDADE")
    return {
        "error": None,
        "upload_filename": upload_filename,
        "reference_filename": reference_bundle["filename"],
        "resolved": resolved,
        "maps": maps,
        "heads": heads,
        "results": results,
        "combined": combined,
        "summary": summary,
        "definitions": reference_bundle["definitions"],
        "parameters": {"KEY_USED": "ENTITY + KPI_CODE", "REFERENCE_FILE": reference_bundle["filename"], "UPLOAD_FILE": upload_filename, "ABS_TOL": abs_tol, "REL_TOL": rel_tol, "GENERATED_AT": datetime.now().isoformat(timespec="seconds")}
    }


def format_workbook(writer):
    workbook = writer.book
    header_format = workbook.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF", "border": 1})
    for worksheet in writer.sheets.values():
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, worksheet.dim_rowmax, worksheet.dim_colmax)
        worksheet.set_row(0, None, header_format)
        worksheet.set_column(0, max(0, worksheet.dim_colmax), 18)


def dataframe_to_excel(frame, sheet_name="Dados"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        frame.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        format_workbook(writer)
    return output.getvalue()


def report(payload):
    output = io.BytesIO()
    data = payload["combined"]
    parameters = pd.DataFrame([payload["parameters"]])
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        data.to_excel(writer, index=False, sheet_name="Comparacao_Completa")
        data[data["STATUS"] == "OK"].to_excel(writer, index=False, sheet_name="OK")
        data[data["STATUS"] != "OK"].to_excel(writer, index=False, sheet_name="Achados")
        payload["summary"].to_excel(writer, index=False, sheet_name="Resumo")
        payload["definitions"].to_excel(writer, index=False, sheet_name="Definition_Book")
        parameters.to_excel(writer, index=False, sheet_name="Parametros")
        format_workbook(writer)
    return output.getvalue()


def owner_filter(reference_bundle, responsible):
    data = reference_bundle["reference_data"].copy()
    key = "" if responsible == "Sem responsável" else norm(responsible)
    return data[data["OWNER_KEY"] == key].copy()


def template_control(data, reference_bundle, responsible, filters):
    classifications = sorted(value for value in data["CLASSIFICATION"].dropna().astype(str).str.strip().unique() if value)
    return pd.DataFrame({
        "CAMPO": ["ARQUIVO_REFERENCIA", "DATA_HORA_GERACAO", "RESPONSAVEL", "REGISTROS", "KPIS", "ENTIDADES", "CLASSIFICACOES", "VERSAO_TEMPLATE", "FILTROS_APLICADOS"],
        "VALOR": [reference_bundle["filename"], datetime.now().isoformat(timespec="seconds"), responsible, len(data), data["KPI_CODE"].nunique(), data["ENTITY"].nunique(), "; ".join(classifications), TEMPLATE_VERSION, filters]
    })


def generate_responsible_template(reference_bundle, responsible):
    data = owner_filter(reference_bundle, responsible)
    template_columns = ["SOURCE", "ENTITY", "KPI_CODE", "KPI_NAME", "CLASSIFICATION", "FORMULA", "UNIT_OF_MEASURE", "OWNER", "VALUE"]
    data = data[template_columns].drop_duplicates(["SOURCE", "ENTITY", "KPI_CODE"], keep="first").sort_values(["SOURCE", "ENTITY", "KPI_CODE"])
    data["VALUE"] = np.nan
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for source in REPORT_SHEETS:
            source_data = data[data["SOURCE"] == source].copy()
            source_data.to_excel(writer, index=False, sheet_name=source)
        instructions = pd.DataFrame({
            "INSTRUÇÕES": [
                f"Responsável: {responsible}",
                f"Gerado em: {datetime.now():%d/%m/%Y %H:%M:%S}",
                f"Referência: {reference_bundle['filename']}",
                f"Quantidade de registros: {len(data)}",
                "Preencha somente a coluna VALUE.",
                "Não altere ENTITY nem KPI_CODE.",
                f"Versão do template: {TEMPLATE_VERSION}"
            ]
        })
        instructions.to_excel(writer, index=False, sheet_name="INSTRUÇÕES")
        template_control(data, reference_bundle, responsible, f"OWNER={responsible}").to_excel(writer, index=False, sheet_name="CONTROLE")
        format_workbook(writer)
    return output.getvalue(), f"template_{safe_filename(responsible)}.xlsx"


def generate_all_templates(reference_bundle):
    responsibles = get_responsibles(reference_bundle)
    output = io.BytesIO()
    summary_rows = []
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for responsible in responsibles["display_name"].tolist():
            template_bytes, filename = generate_responsible_template(reference_bundle, responsible)
            archive.writestr(filename, template_bytes)
            data = owner_filter(reference_bundle, responsible)
            summary_rows.append({
                "RESPONSAVEL": responsible,
                "ARQUIVO": filename,
                "QUANTIDADE_REGISTROS": len(data),
                "QUANTIDADE_KPIS": data["KPI_CODE"].nunique(),
                "QUANTIDADE_ENTIDADES": data["ENTITY"].nunique(),
                "KPIS_LOGS": data.loc[data["SOURCE"] == "LOGS", "KPI_CODE"].nunique(),
                "KPIS_BOPS": data.loc[data["SOURCE"] == "BOPS", "KPI_CODE"].nunique(),
                "KPIS_SL": data.loc[data["SOURCE"] == "SL", "KPI_CODE"].nunique(),
                "CLASSIFICACOES": "; ".join(sorted(value for value in data["CLASSIFICATION"].dropna().astype(str).str.strip().unique() if value))
            })
        summary = pd.DataFrame(summary_rows)
        archive.writestr("resumo_templates.xlsx", dataframe_to_excel(summary, "Resumo_Templates"))
    return output.getvalue()


def generate_classification_catalog(reference_bundle):
    catalog = get_kpi_catalog(reference_bundle)
    classifications = get_classifications(reference_bundle)["data"]
    responsibles = get_responsibles(reference_bundle)
    kpis = catalog[[column for column in ["SOURCE", "KPI_CODE", "KPI_NAME", "CLASSIFICATION", "UNIT_OF_MEASURE", "OWNER"] if column in catalog.columns]].drop_duplicates()
    entities = pd.DataFrame({"ENTITY": sorted(catalog["ENTITY"].dropna().unique().tolist())})
    summary = pd.DataFrame([get_reference_summary(reference_bundle)])
    control = template_control(reference_bundle["reference_data"], reference_bundle, "Todos", "Nenhum")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        classifications.to_excel(writer, index=False, sheet_name="Classificacoes")
        kpis.to_excel(writer, index=False, sheet_name="KPIs")
        responsibles.to_excel(writer, index=False, sheet_name="Responsaveis")
        entities.to_excel(writer, index=False, sheet_name="Entidades")
        summary.to_excel(writer, index=False, sheet_name="Resumo")
        reference_bundle["definitions"].to_excel(writer, index=False, sheet_name="Definition_Book")
        control.to_excel(writer, index=False, sheet_name="Controle")
        format_workbook(writer)
    return output.getvalue()


def get_reference_file_download(reference_bytes):
    return reference_bytes
