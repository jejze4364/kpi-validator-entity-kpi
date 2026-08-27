import io
import re
import unicodedata

import numpy as np
import pandas as pd

REPORTS = ["LOGS", "BOPS", "SL"]
REFERENCE_SHEET = "SHAREPOINT"
DEFINITION_SHEET = "DEFINITION BOOK"

ALIASES = {
    "entity": ["entity", "entidade", "unit name", "unidade", "bu", "location", "plant", "dc"],
    "kpi": ["kpi code", "kpi_code", "codigo kpi", "código kpi", "cod kpi"],
    "value": ["current_value ac", "current value ac", "current value", "valor anaplan", "valor origem", "actual", "ac", "value", "valor"],
    "kpi_name": ["kpi name", "nome kpi", "description", "descrição"],
    "formula": ["formula", "fórmula", "regra"],
    "uom": ["unit_of_measure", "unit of measure", "uom"],
    "owner": ["owner", "responsavel", "responsável"],
}


def noacc(value):
    return "".join(
        character
        for character in unicodedata.normalize("NFKD", str(value))
        if not unicodedata.combining(character)
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
    if not normalized:
        return ""
    match = re.search(r"\b([A-Z]{2,3}-[KR]\d{3,5}(?:_\d{4})?)\b", normalized)
    return match.group(1) if match else normalized


def resolve_sheet(sheet_names, target):
    exact = [name for name in sheet_names if norm(name) == norm(target)]
    if exact:
        return exact[0]
    if target == DEFINITION_SHEET:
        aliases = ["DEFINITION BOOK", "DEFINITIONBOOK", "DEFINITIONS", "DEFINITION"]
    elif target == REFERENCE_SHEET:
        aliases = ["SHAREPOINT", "SHARE POINT"]
    else:
        aliases = [target]
    for alias in aliases:
        partial = [name for name in sheet_names if norm(alias) in norm(name)]
        if partial:
            return partial[0]
    return None


def detect_header(raw):
    if raw.empty:
        raise ValueError("A aba está vazia.")
    value_aliases = {nc(alias) for alias in ALIASES["value"]}
    best_index = None
    best_score = -1
    for index in range(min(80, len(raw))):
        values = [norm(value) for value in raw.iloc[index].tolist()]
        has_kpi = any(
            "KPI" in value and ("CODE" in value or "CODIGO" in value or "COD" in value)
            for value in values
        )
        has_entity = any(
            value in {"ENTITY", "ENTIDADE", "UNIT NAME", "UNIDADE", "BU", "LOCATION", "PLANT", "DC"}
            for value in values
        )
        has_value = any(nc(value) in value_aliases for value in values if value)
        score = 1000 * has_kpi + 800 * has_entity + 500 * has_value + sum(bool(value) for value in values)
        if score > best_score:
            best_index = index
            best_score = score
    if best_index is None:
        raise ValueError("Não foi possível localizar o cabeçalho da aba.")
    return best_index


def make_unique_columns(header_values):
    columns = []
    occurrences = {}
    for position, value in enumerate(header_values):
        base = f"COL_{position + 1}" if pd.isna(value) or not str(value).strip() else str(value).strip()
        key = nc(base) or f"col {position + 1}"
        occurrences[key] = occurrences.get(key, 0) + 1
        count = occurrences[key]
        columns.append(f"{base}_{count}" if count > 1 else base)
    return columns


def read_sheet(source, sheet_name):
    raw = pd.read_excel(source, sheet_name=sheet_name, header=None, engine="openpyxl")
    header_index = detect_header(raw)
    data = raw.iloc[header_index + 1:].copy()
    data.columns = make_unique_columns(raw.iloc[header_index].tolist())
    return data.dropna(how="all").reset_index(drop=True), header_index + 1


def find_column(dataframe, role):
    normalized_columns = {}
    for column in dataframe.columns:
        normalized = nc(column)
        if normalized and normalized not in normalized_columns:
            normalized_columns[normalized] = column
    for alias in ALIASES[role]:
        normalized_alias = nc(alias)
        if normalized_alias in normalized_columns:
            return normalized_columns[normalized_alias]
    candidates = []
    for normalized_column, original_column in normalized_columns.items():
        for alias in ALIASES[role]:
            normalized_alias = nc(alias)
            if normalized_alias and (
                normalized_alias in normalized_column
                or normalized_column in normalized_alias
            ):
                candidates.append((abs(len(normalized_column) - len(normalized_alias)), len(normalized_column), original_column))
    if candidates:
        candidates.sort(key=lambda item: (item[0], item[1]))
        return candidates[0][2]
    return None


def build_mapping(dataframe, source_name):
    mapped = {role: find_column(dataframe, role) for role in ALIASES}
    normalized_columns = {nc(column): column for column in dataframe.columns}
    preferences = (
        ["current value ac", "current value", "ac", "value", "valor", "actual"]
        if source_name == REFERENCE_SHEET
        else ["valor anaplan", "valor origem", "actual", "ac", "value", "valor", "current value ac", "current value"]
    )
    for preference in preferences:
        if nc(preference) in normalized_columns:
            mapped["value"] = normalized_columns[nc(preference)]
            break
    return mapped


def parse_number(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    values = series.astype("string").str.strip().replace(
        {"": pd.NA, "-": pd.NA, "--": pd.NA, "nan": pd.NA, "NaN": pd.NA, "None": pd.NA, "<NA>": pd.NA}
    )
    negative = values.str.match(r"^\(.*\)$", na=False)
    values.loc[negative] = "-" + values.loc[negative].str.replace("(", "", regex=False).str.replace(")", "", regex=False)
    values = values.str.replace("\u00a0", "", regex=False).str.replace(" ", "", regex=False)
    br_decimal = values.str.match(r"^-?\d{1,3}(?:\.\d{3})+,\d+$", na=False)
    br_simple = values.str.match(r"^-?\d+,\d+$", na=False)
    intl_decimal = values.str.match(r"^-?\d{1,3}(?:,\d{3})+\.\d+$", na=False)
    dot_integer = values.str.match(r"^-?\d{1,3}(?:\.\d{3})+$", na=False)
    comma_integer = values.str.match(r"^-?\d{1,3}(?:,\d{3})+$", na=False)
    values.loc[br_decimal] = values.loc[br_decimal].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    values.loc[br_simple] = values.loc[br_simple].str.replace(",", ".", regex=False)
    values.loc[intl_decimal] = values.loc[intl_decimal].str.replace(",", "", regex=False)
    values.loc[dot_integer] = values.loc[dot_integer].str.replace(".", "", regex=False)
    values.loc[comma_integer] = values.loc[comma_integer].str.replace(",", "", regex=False)
    unresolved = ~(br_decimal | br_simple | intl_decimal | dot_integer | comma_integer)
    values.loc[unresolved] = values.loc[unresolved].str.replace(",", ".", regex=False)
    return pd.to_numeric(values, errors="coerce")


def standardize(dataframe, mapped, source_name):
    missing = [role for role in ["entity", "kpi", "value"] if not mapped.get(role)]
    if missing:
        raise ValueError(
            f"{source_name}: colunas obrigatórias não localizadas: {', '.join(missing)}. "
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
    output["KPI_NAME"] = (
        dataframe[mapped["kpi_name"]].fillna("").astype(str)
        if mapped.get("kpi_name")
        else ""
    )
    return output[output["ENTITY"].ne("") & output["KPI_CODE"].ne("")].reset_index(drop=True)


def build_definitions(dataframe, mapped):
    columns = ["KPI_CODE", "DEF_KPI_NAME", "DEF_FORMULA", "DEF_UOM", "DEF_OWNER"]
    if not mapped.get("kpi"):
        return pd.DataFrame(columns=columns)
    output = pd.DataFrame({"KPI_CODE": dataframe[mapped["kpi"]].map(nkpi)})
    fields = {"DEF_KPI_NAME": "kpi_name", "DEF_FORMULA": "formula", "DEF_UOM": "uom", "DEF_OWNER": "owner"}
    for output_column, role in fields.items():
        source_column = mapped.get(role)
        output[output_column] = dataframe[source_column].fillna("").astype(str) if source_column else ""
    return output[output["KPI_CODE"].ne("")].drop_duplicates("KPI_CODE").reset_index(drop=True)[columns]


def aggregate_report(dataframe, source_name):
    grouped = dataframe.groupby(["ENTITY", "KPI_CODE"], as_index=False, dropna=False).agg(
        REPORT_VALUE=("VALUE", lambda values: values.sum(min_count=1)),
        REPORT_ROWS=("VALUE", "size"),
        REPORT_VALID_VALUES=("VALUE", "count"),
        KPI_NAME=("KPI_NAME", "first"),
    )
    grouped["SOURCE"] = source_name
    return grouped


def aggregate_reference(dataframe):
    return dataframe.groupby(["ENTITY", "KPI_CODE"], as_index=False, dropna=False).agg(
        REFERENCE_VALUE=("VALUE", lambda values: values.sum(min_count=1)),
        REFERENCE_ROWS=("VALUE", "size"),
        REFERENCE_VALID_VALUES=("VALUE", "count"),
    )


def compare(report_data, reference_data, definitions, source_name, absolute_tolerance, relative_tolerance):
    result = aggregate_report(report_data, source_name).merge(
        aggregate_reference(reference_data), on=["ENTITY", "KPI_CODE"], how="outer", indicator=True
    )
    result["SOURCE"] = source_name
    result = result.merge(definitions, on="KPI_CODE", how="left")
    result["DIFFERENCE"] = result["REPORT_VALUE"] - result["REFERENCE_VALUE"]
    result["DIFFERENCE_PCT"] = np.where(
        result["REFERENCE_VALUE"].notna() & result["REFERENCE_VALUE"].abs().gt(absolute_tolerance),
        result["DIFFERENCE"] / result["REFERENCE_VALUE"],
        np.nan,
    )
    limit = absolute_tolerance + relative_tolerance * result["REFERENCE_VALUE"].abs()
    within = result["REPORT_VALUE"].notna() & result["REFERENCE_VALUE"].notna() & result["DIFFERENCE"].abs().le(limit)
    invalid_report = result["_merge"].ne("right_only") & result["REPORT_VALUE"].isna()
    invalid_reference = result["_merge"].ne("left_only") & result["REFERENCE_VALUE"].isna()
    result["STATUS"] = np.select(
        [result["_merge"].eq("left_only"), result["_merge"].eq("right_only"), invalid_report, invalid_reference, within],
        ["NÃO ESTÁ NA REFERÊNCIA", "SOMENTE NA REFERÊNCIA", "VALOR INVÁLIDO NO ARQUIVO ENVIADO", "VALOR INVÁLIDO NA REFERÊNCIA", "OK"],
        default="DIVERGENTE",
    )
    result["KEY_USED"] = "ENTITY + KPI_CODE"
    result = result.drop(columns="_merge")
    order = ["SOURCE", "ENTITY", "KPI_CODE", "KPI_NAME", "DEF_KPI_NAME", "DEF_FORMULA", "DEF_UOM", "DEF_OWNER", "REPORT_VALUE", "REFERENCE_VALUE", "DIFFERENCE", "DIFFERENCE_PCT", "REPORT_ROWS", "REFERENCE_ROWS", "REPORT_VALID_VALUES", "REFERENCE_VALID_VALUES", "STATUS", "KEY_USED"]
    for column in order:
        if column not in result:
            result[column] = pd.NA
    return result[order].sort_values(["STATUS", "ENTITY", "KPI_CODE"], na_position="last").reset_index(drop=True)


def load_reference(data, filename="consolidador.xlsm"):
    try:
        source = io.BytesIO(bytes(data))
        sheet_names = pd.ExcelFile(source, engine="openpyxl").sheet_names
        resolved = {
            REFERENCE_SHEET: resolve_sheet(sheet_names, REFERENCE_SHEET),
            DEFINITION_SHEET: resolve_sheet(sheet_names, DEFINITION_SHEET),
        }
        if not resolved[REFERENCE_SHEET]:
            return {
                "error": f"A aba SHAREPOINT não foi encontrada no arquivo de referência {filename}.",
                "details": {"reference_sheet_names": sheet_names, "resolved_sheets": resolved},
            }
        raw = {}
        maps = {}
        headers = {}
        source.seek(0)
        raw[REFERENCE_SHEET], headers[REFERENCE_SHEET] = read_sheet(source, resolved[REFERENCE_SHEET])
        maps[REFERENCE_SHEET] = build_mapping(raw[REFERENCE_SHEET], REFERENCE_SHEET)
        reference_data = standardize(raw[REFERENCE_SHEET], maps[REFERENCE_SHEET], REFERENCE_SHEET)
        if resolved[DEFINITION_SHEET]:
            source.seek(0)
            raw[DEFINITION_SHEET], headers[DEFINITION_SHEET] = read_sheet(source, resolved[DEFINITION_SHEET])
            maps[DEFINITION_SHEET] = build_mapping(raw[DEFINITION_SHEET], DEFINITION_SHEET)
            definitions = build_definitions(raw[DEFINITION_SHEET], maps[DEFINITION_SHEET])
        else:
            headers[DEFINITION_SHEET] = None
            maps[DEFINITION_SHEET] = {role: None for role in ALIASES}
            definitions = pd.DataFrame(columns=["KPI_CODE", "DEF_KPI_NAME", "DEF_FORMULA", "DEF_UOM", "DEF_OWNER"])
        return {
            "error": None,
            "filename": filename,
            "reference_sheet_names": sheet_names,
            "reference_sheets": resolved,
            "maps": maps,
            "headers": headers,
            "reference_data": reference_data,
            "definitions": definitions,
        }
    except Exception as error:
        return {
            "error": f"Falha ao carregar o arquivo de referência: {type(error).__name__}: {error}",
            "details": {"filename": filename},
        }


def run(upload_bytes, reference_bundle, absolute_tolerance=0.01, relative_tolerance=0.0001):
    try:
        if reference_bundle.get("error"):
            return reference_bundle
        source = io.BytesIO(bytes(upload_bytes))
        sheet_names = pd.ExcelFile(source, engine="openpyxl").sheet_names
        uploaded_sheets = {name: resolve_sheet(sheet_names, name) for name in REPORTS}
        missing = [name for name in REPORTS if not uploaded_sheets[name]]
        if missing:
            return {
                "error": "Abas obrigatórias não encontradas no arquivo enviado: " + ", ".join(missing),
                "details": {"missing_sheets": missing, "uploaded_sheet_names": sheet_names, "resolved_sheets": uploaded_sheets},
            }
        raw = {}
        maps = {}
        headers = {}
        standardized = {}
        for name in REPORTS:
            source.seek(0)
            raw[name], headers[name] = read_sheet(source, uploaded_sheets[name])
            maps[name] = build_mapping(raw[name], name)
            standardized[name] = standardize(raw[name], maps[name], name)
        for name in [REFERENCE_SHEET, DEFINITION_SHEET]:
            maps[name] = reference_bundle.get("maps", {}).get(name, {})
            headers[name] = reference_bundle.get("headers", {}).get(name)
        results = [
            compare(
                standardized[name],
                reference_bundle["reference_data"],
                reference_bundle["definitions"],
                name,
                float(absolute_tolerance),
                float(relative_tolerance),
            )
            for name in REPORTS
        ]
        return {
            "error": None,
            "uploaded_sheet_names": sheet_names,
            "reference_sheet_names": reference_bundle.get("reference_sheet_names", []),
            "uploaded_sheets": uploaded_sheets,
            "reference_sheets": reference_bundle.get("reference_sheets", {}),
            "maps": maps,
            "headers": headers,
            "definitions": reference_bundle["definitions"],
            "results": results,
            "reference_filename": reference_bundle.get("filename", "consolidador.xlsm"),
            "absolute_tolerance": float(absolute_tolerance),
            "relative_tolerance": float(relative_tolerance),
        }
    except Exception as error:
        return {
            "error": f"Falha durante o processamento: {type(error).__name__}: {error}",
            "details": {
                "uploaded_sheet_names": locals().get("sheet_names", []),
                "uploaded_sheets": locals().get("uploaded_sheets", {}),
                "maps": locals().get("maps", {}),
                "headers": locals().get("headers", {}),
            },
        }


def report(payload):
    output = io.BytesIO()
    result = pd.concat(payload["results"], ignore_index=True)
    summary = result.groupby(["SOURCE", "STATUS"], dropna=False).size().reset_index(name="QUANTIDADE").sort_values(["SOURCE", "STATUS"])
    parameters = pd.DataFrame(
        {
            "PARAMETRO": ["CHAVE", "ARQUIVO_REFERENCIA", "ABA_REFERENCIA", "TOLERANCIA_ABSOLUTA", "TOLERANCIA_RELATIVA"],
            "VALOR": [
                "ENTITY + KPI_CODE",
                payload.get("reference_filename", "consolidador.xlsm"),
                payload.get("reference_sheets", {}).get(REFERENCE_SHEET, REFERENCE_SHEET),
                payload["absolute_tolerance"],
                payload["relative_tolerance"],
            ],
        }
    )
    frames = {
        "Comparacao_Completa": result,
        "OK": result[result["STATUS"].eq("OK")],
        "Achados": result[result["STATUS"].ne("OK")],
        "Resumo": summary,
        "Definition_Book": payload["definitions"],
        "Parametros": parameters,
    }
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, dataframe in frames.items():
            dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            if len(dataframe.columns):
                worksheet.autofilter(0, 0, max(len(dataframe), 1), len(dataframe.columns) - 1)
                for index, column in enumerate(dataframe.columns):
                    maximum = len(str(column)) if dataframe.empty else max(len(str(column)), int(dataframe[column].fillna("").astype(str).map(len).max()))
                    worksheet.set_column(index, index, min(maximum + 2, 50))
    output.seek(0)
    return output.getvalue()
