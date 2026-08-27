import io
import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd

REPORTS = ["LOGS", "BOPS", "SL"]
REFERENCE_SHEET = "SHAREPOINT"
DEFINITION_SHEET = "DEFINITION BOOK"

ALIASES = {
    "entity": ["entity", "entidade", "unit name", "unidade", "bu", "location", "plant", "dc"],
    "kpi": ["kpi code", "kpi_code", "codigo kpi", "código kpi", "cod kpi"],
    "value": ["current_value ac", "current value ac", "valor anaplan", "valor origem", "actual", "ac", "value", "valor"],
    "kpi_name": ["kpi name", "nome kpi", "description", "descrição"],
    "formula": ["formula", "fórmula", "regra"],
    "uom": ["unit_of_measure", "unit of measure", "uom"],
    "owner": ["owner", "responsavel", "responsável"],
}


def noacc(value):
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", str(value))
        if not unicodedata.combining(char)
    )


def norm(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", noacc(value).strip().upper())


def nc(value):
    normalized = noacc(value).strip().lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", normalized)).strip()


def nkpi(value):
    normalized = norm(value)
    match = re.search(r"\b([A-Z]{2,3}-[KR]\d{3,5}(?:_\d{4})?)\b", normalized)
    return match.group(1) if match else normalized


def resolve_sheet(names, target):
    exact = [name for name in names if norm(name) == norm(target)]
    if exact:
        return exact[0]
    aliases = ["DEFINITION BOOK", "DEFINITIONBOOK", "DEFINITIONS"] if target == DEFINITION_SHEET else [target]
    for alias in aliases:
        partial = [name for name in names if norm(alias) in norm(name)]
        if partial:
            return partial[0]
    return None


def detect_header(raw):
    best = (0, -1)
    for index in range(min(80, len(raw))):
        values = [norm(value) for value in raw.iloc[index]]
        score = (
            1000 * any("KPI" in value and ("CODE" in value or "CODIGO" in value) for value in values)
            + 800 * any(value in ["ENTITY", "ENTIDADE", "UNIT NAME", "UNIDADE", "BU", "LOCATION"] for value in values)
            + sum(bool(value) for value in values)
        )
        if score > best[1]:
            best = (index, score)
    return best[0]


def read_sheet(source, sheet_name):
    raw = pd.read_excel(source, sheet_name=sheet_name, header=None, engine="openpyxl")
    header_index = detect_header(raw)
    columns = []
    seen = {}
    for position, value in enumerate(raw.iloc[header_index]):
        column = str(value).strip() if not pd.isna(value) else f"COL_{position + 1}"
        seen[column] = seen.get(column, 0) + 1
        if seen[column] > 1:
            column = f"{column}_{seen[column]}"
        columns.append(column)
    data = raw.iloc[header_index + 1:].copy()
    data.columns = columns
    return data.dropna(how="all").reset_index(drop=True), header_index + 1


def find_column(dataframe, role):
    columns = {nc(column): column for column in dataframe.columns}
    for alias in ALIASES[role]:
        if nc(alias) in columns:
            return columns[nc(alias)]
    for normalized_column, original_column in columns.items():
        if any(nc(alias) in normalized_column for alias in ALIASES[role]):
            return original_column
    return None


def build_mapping(dataframe, source_name):
    mapped = {role: find_column(dataframe, role) for role in ALIASES}
    columns = {nc(column): column for column in dataframe.columns}
    preferences = (
        ["current value ac", "current value", "ac", "value", "valor"]
        if source_name == REFERENCE_SHEET
        else ["valor anaplan", "ac", "actual", "valor origem", "value", "valor"]
    )
    for preference in preferences:
        if preference in columns:
            mapped["value"] = columns[preference]
            break
    return mapped


def parse_number(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    values = series.astype(str).str.strip().replace({"": np.nan, "-": np.nan, "nan": np.nan})
    brazilian = values.str.contains(r"^-?\d{1,3}(?:\.\d{3})+,\d+$", regex=True, na=False)
    values.loc[brazilian] = (
        values.loc[brazilian]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    values.loc[~brazilian] = values.loc[~brazilian].str.replace(",", ".", regex=False)
    return pd.to_numeric(values, errors="coerce")


def standardize(dataframe, mapped, source_name):
    missing = [role for role in ["entity", "kpi", "value"] if not mapped.get(role)]
    if missing:
        raise ValueError(
            f"{source_name}: colunas não localizadas: {missing}. "
            f"Colunas detectadas: {list(dataframe.columns)}"
        )
    output = pd.DataFrame(
        {
            "SOURCE": source_name,
            "ENTITY": dataframe[mapped["entity"]].map(norm),
            "KPI_CODE": dataframe[mapped["kpi"]].map(nkpi),
            "VALUE": parse_number(dataframe[mapped["value"]]),
        }
    )
    output["KPI_NAME"] = dataframe[mapped["kpi_name"]].astype(str) if mapped.get("kpi_name") else ""
    return output[(output["ENTITY"] != "") & (output["KPI_CODE"] != "")]


def build_definitions(dataframe, mapped):
    if not mapped.get("kpi"):
        return pd.DataFrame(columns=["KPI_CODE"])
    output = pd.DataFrame({"KPI_CODE": dataframe[mapped["kpi"]].map(nkpi)})
    output["DEF_KPI_NAME"] = dataframe[mapped["kpi_name"]].astype(str) if mapped.get("kpi_name") else ""
    output["DEF_FORMULA"] = dataframe[mapped["formula"]].astype(str) if mapped.get("formula") else ""
    output["DEF_UOM"] = dataframe[mapped["uom"]].astype(str) if mapped.get("uom") else ""
    output["DEF_OWNER"] = dataframe[mapped["owner"]].astype(str) if mapped.get("owner") else ""
    return output[output["KPI_CODE"] != ""].drop_duplicates("KPI_CODE")


def compare(report_data, reference_data, definitions, absolute_tolerance, relative_tolerance):
    report_grouped = (
        report_data.groupby(["ENTITY", "KPI_CODE"])
        .agg(
            REPORT_VALUE=("VALUE", "sum"),
            REPORT_ROWS=("VALUE", "size"),
            KPI_NAME=("KPI_NAME", "first"),
            SOURCE=("SOURCE", "first"),
        )
        .reset_index()
    )
    reference_grouped = (
        reference_data.groupby(["ENTITY", "KPI_CODE"])
        .agg(REFERENCE_VALUE=("VALUE", "sum"), REFERENCE_ROWS=("VALUE", "size"))
        .reset_index()
    )
    result = report_grouped.merge(
        reference_grouped,
        on=["ENTITY", "KPI_CODE"],
        how="outer",
        indicator=True,
    ).merge(definitions, on="KPI_CODE", how="left")
    result["DIFFERENCE"] = result["REPORT_VALUE"] - result["REFERENCE_VALUE"]
    result["DIFFERENCE_PCT"] = np.where(
        result["REFERENCE_VALUE"].abs() > absolute_tolerance,
        result["DIFFERENCE"] / result["REFERENCE_VALUE"],
        np.nan,
    )
    within_tolerance = (
        result["DIFFERENCE"].abs()
        <= absolute_tolerance + relative_tolerance * result["REFERENCE_VALUE"].abs()
    ).fillna(False)
    result["STATUS"] = np.select(
        [result["_merge"].eq("left_only"), result["_merge"].eq("right_only"), within_tolerance],
        ["NÃO ESTÁ NA REFERÊNCIA", "SOMENTE NA REFERÊNCIA", "OK"],
        default="DIVERGENTE",
    )
    result["KEY_USED"] = "ENTITY + KPI_CODE"
    return result.drop(columns="_merge")


def run(uploaded_data, reference_path, absolute_tolerance=0.01, relative_tolerance=0.0001):
    try:
        uploaded_source = io.BytesIO(uploaded_data)
        reference_source = Path(reference_path)
        uploaded_names = pd.ExcelFile(uploaded_source, engine="openpyxl").sheet_names
        reference_names = pd.ExcelFile(reference_source, engine="openpyxl").sheet_names

        uploaded_sheets = {name: resolve_sheet(uploaded_names, name) for name in REPORTS}
        reference_sheets = {
            REFERENCE_SHEET: resolve_sheet(reference_names, REFERENCE_SHEET),
            DEFINITION_SHEET: resolve_sheet(reference_names, DEFINITION_SHEET),
        }

        missing_reports = [name for name in REPORTS if not uploaded_sheets[name]]
        if missing_reports:
            return {
                "error": "Abas não encontradas no arquivo enviado: " + ", ".join(missing_reports),
                "details": {"uploaded_sheets": uploaded_names},
            }
        if not reference_sheets[REFERENCE_SHEET]:
            return {
                "error": "A aba SHAREPOINT não foi encontrada no consolidador.xlsm do repositório.",
                "details": {"reference_sheets": reference_names},
            }

        raw = {}
        maps = {}
        headers = {}

        for report_name in REPORTS:
            uploaded_source.seek(0)
            raw[report_name], headers[report_name] = read_sheet(uploaded_source, uploaded_sheets[report_name])
            maps[report_name] = build_mapping(raw[report_name], report_name)

        raw[REFERENCE_SHEET], headers[REFERENCE_SHEET] = read_sheet(reference_source, reference_sheets[REFERENCE_SHEET])
        maps[REFERENCE_SHEET] = build_mapping(raw[REFERENCE_SHEET], REFERENCE_SHEET)

        reference_data = standardize(raw[REFERENCE_SHEET], maps[REFERENCE_SHEET], REFERENCE_SHEET)

        if reference_sheets[DEFINITION_SHEET]:
            raw[DEFINITION_SHEET], headers[DEFINITION_SHEET] = read_sheet(reference_source, reference_sheets[DEFINITION_SHEET])
            maps[DEFINITION_SHEET] = build_mapping(raw[DEFINITION_SHEET], DEFINITION_SHEET)
            definitions = build_definitions(raw[DEFINITION_SHEET], maps[DEFINITION_SHEET])
        else:
            definitions = pd.DataFrame(columns=["KPI_CODE"])

        results = [
            compare(
                standardize(raw[report_name], maps[report_name], report_name),
                reference_data,
                definitions,
                absolute_tolerance,
                relative_tolerance,
            )
            for report_name in REPORTS
        ]

        return {
            "error": None,
            "uploaded_sheets": uploaded_sheets,
            "reference_sheets": reference_sheets,
            "maps": maps,
            "headers": headers,
            "definitions": definitions,
            "results": results,
            "absolute_tolerance": absolute_tolerance,
            "relative_tolerance": relative_tolerance,
        }
    except Exception as error:
        return {"error": f"Falha durante o processamento: {type(error).__name__}: {error}"}


def report(payload):
    output = io.BytesIO()
    result = pd.concat(payload["results"], ignore_index=True)
    parameters = pd.DataFrame(
        {
            "PARAMETRO": ["CHAVE", "BASE", "TOLERANCIA_ABSOLUTA", "TOLERANCIA_RELATIVA"],
            "VALOR": [
                "ENTITY + KPI_CODE",
                "consolidador.xlsm",
                payload["absolute_tolerance"],
                payload["relative_tolerance"],
            ],
        }
    )
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result.to_excel(writer, index=False, sheet_name="Comparacao_Completa")
        result[result["STATUS"] == "OK"].to_excel(writer, index=False, sheet_name="OK")
        result[result["STATUS"] != "OK"].to_excel(writer, index=False, sheet_name="Achados")
        result.groupby(["SOURCE", "STATUS"]).size().reset_index(name="QUANTIDADE").to_excel(
            writer, index=False, sheet_name="Resumo"
        )
        payload["definitions"].to_excel(writer, index=False, sheet_name="Definition_Book")
        parameters.to_excel(writer, index=False, sheet_name="Parametros")
    return output.getvalue()
