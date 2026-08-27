import io
import re
import unicodedata

import numpy as np
import pandas as pd


REPORTS = ["LOGS", "BOPS", "SL"]
REFERENCE_SHEET = "SHAREPOINT"
DEFINITION_SHEET = "DEFINITION BOOK"

ALIASES = {
    "entity": [
        "entity",
        "entidade",
        "unit name",
        "unidade",
        "bu",
        "location",
        "plant",
        "dc",
    ],
    "kpi": [
        "kpi code",
        "kpi_code",
        "codigo kpi",
        "código kpi",
        "cod kpi",
    ],
    "value": [
        "current_value ac",
        "current value ac",
        "current value",
        "valor anaplan",
        "valor origem",
        "actual",
        "ac",
        "value",
        "valor",
    ],
    "kpi_name": [
        "kpi name",
        "nome kpi",
        "description",
        "descrição",
    ],
    "formula": [
        "formula",
        "fórmula",
        "regra",
    ],
    "uom": [
        "unit_of_measure",
        "unit of measure",
        "uom",
    ],
    "owner": [
        "owner",
        "responsavel",
        "responsável",
    ],
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

    return re.sub(
        r"\s+",
        " ",
        noacc(value).strip().upper(),
    )


def nc(value):
    normalized = noacc(value).strip().lower()

    return re.sub(
        r"\s+",
        " ",
        re.sub(r"[^a-z0-9]+", " ", normalized),
    ).strip()


def nkpi(value):
    normalized = norm(value)

    if not normalized:
        return ""

    match = re.search(
        r"\b([A-Z]{2,3}-[KR]\d{3,5}(?:_\d{4})?)\b",
        normalized,
    )

    return match.group(1) if match else normalized


def resolve_sheet(sheet_names, target):
    exact_matches = [
        sheet_name
        for sheet_name in sheet_names
        if norm(sheet_name) == norm(target)
    ]

    if exact_matches:
        return exact_matches[0]

    if target == DEFINITION_SHEET:
        aliases = [
            "DEFINITION BOOK",
            "DEFINITIONBOOK",
            "DEFINITIONS",
            "DEFINITION",
        ]
    elif target == REFERENCE_SHEET:
        aliases = [
            "SHAREPOINT",
            "SHARE POINT",
        ]
    else:
        aliases = [target]

    for alias in aliases:
        partial_matches = [
            sheet_name
            for sheet_name in sheet_names
            if norm(alias) in norm(sheet_name)
        ]

        if partial_matches:
            return partial_matches[0]

    return None


def detect_header(raw):
    if raw.empty:
        raise ValueError("A aba está vazia.")

    best_index = None
    best_score = -1

    row_limit = min(80, len(raw))

    for index in range(row_limit):
        values = [norm(value) for value in raw.iloc[index].tolist()]

        has_kpi = any(
            "KPI" in value
            and (
                "CODE" in value
                or "CODIGO" in value
                or "COD" in value
            )
            for value in values
        )

        has_entity = any(
            value in {
                "ENTITY",
                "ENTIDADE",
                "UNIT NAME",
                "UNIDADE",
                "BU",
                "LOCATION",
                "PLANT",
                "DC",
            }
            for value in values
        )

        has_value = any(
            nc(value) in {nc(alias) for alias in ALIASES["value"]}
            for value in values
            if value
        )

        filled_cells = sum(bool(value) for value in values)

        score = (
            1000 * has_kpi
            + 800 * has_entity
            + 500 * has_value
            + filled_cells
        )

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
        if pd.isna(value) or not str(value).strip():
            base_name = f"COL_{position + 1}"
        else:
            base_name = str(value).strip()

        normalized_name = nc(base_name)
        occurrence_key = normalized_name or f"col {position + 1}"

        occurrences[occurrence_key] = occurrences.get(occurrence_key, 0) + 1
        occurrence = occurrences[occurrence_key]

        if occurrence > 1:
            column_name = f"{base_name}_{occurrence}"
        else:
            column_name = base_name

        columns.append(column_name)

    return columns


def read_sheet(source, sheet_name):
    raw = pd.read_excel(
        source,
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl",
    )

    header_index = detect_header(raw)
    columns = make_unique_columns(raw.iloc[header_index].tolist())

    data = raw.iloc[header_index + 1:].copy()
    data.columns = columns

    data = (
        data.dropna(how="all")
        .reset_index(drop=True)
    )

    return data, header_index + 1


def find_column(dataframe, role):
    normalized_columns = {}

    for column in dataframe.columns:
        normalized = nc(column)

        if normalized and normalized not in normalized_columns:
            normalized_columns[normalized] = column

    for alias in ALIASESnormalized_alias = nc(alias)

        if normalized_alias in normalized_columns:
            return normalized_columns[normalized_alias]

    partial_candidates = []

    for normalized_column, original_column in normalized_columns.items():
        for alias in ALIASESnormalized_alias = nc(alias)

            if (
                normalized_alias
                and (
                    normalized_alias in normalized_column
                    or normalized_column in normalized_alias
                )
            ):
                partial_candidates.append(
                    (
                        abs(len(normalized_column) - len(normalized_alias)),
                        len(normalized_column),
                        original_column,
                    )
                )

    if partial_candidates:
        partial_candidates.sort(key=lambda item: (item[0], item[1]))
        return partial_candidates[0][2]

    return None


def build_mapping(dataframe, source_name):
    mapped = {
        role: find_column(dataframe, role)
        for role in ALIASES
    }

    normalized_columns = {
        nc(column): column
        for column in dataframe.columns
    }

    if source_name == REFERENCE_SHEET:
        value_preferences = [
            "current value ac",
            "current value",
            "ac",
            "value",
            "valor",
            "actual",
        ]
    else:
        value_preferences = [
            "valor anaplan",
            "valor origem",
            "actual",
            "ac",
            "value",
            "valor",
            "current value ac",
            "current value",
        ]

    for preference in value_preferences:
        normalized_preference = nc(preference)

        if normalized_preference in normalized_columns:
            mapped["value"] = normalized_columns[normalized_preference]
            break

    return mapped


def parse_number(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    values = series.astype("string").str.strip()

    values = values.replace(
        {
            "": pd.NA,
            "-": pd.NA,
            "--": pd.NA,
            "nan": pd.NA,
            "NaN": pd.NA,
            "None": pd.NA,
            "<NA>": pd.NA,
        }
    )

    negative_parentheses = values.str.match(
        r"^\(.*\)$",
        na=False,
    )

    values.loc[negative_parentheses] = (
        "-"
        + values.loc[negative_parentheses]
        .str.replace("(", "", regex=False)
        .str.replace(")", "", regex=False)
    )

    values = (
        values
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    brazilian_decimal = values.str.match(
        r"^-?\d{1,3}(?:\.\d{3})+,\d+$",
        na=False,
    )

    brazilian_simple = values.str.match(
        r"^-?\d+,\d+$",
        na=False,
    )

    international_thousands = values.str.match(
        r"^-?\d{1,3}(?:,\d{3})+\.\d+$",
        na=False,
    )

    integer_with_dots = values.str.match(
        r"^-?\d{1,3}(?:\.\d{3})+$",
        na=False,
    )

    integer_with_commas = values.str.match(
        r"^-?\d{1,3}(?:,\d{3})+$",
        na=False,
    )

    values.loc[brazilian_decimal] = (
        values.loc[brazilian_decimal]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    values.loc[brazilian_simple] = (
        values.loc[brazilian_simple]
        .str.replace(",", ".", regex=False)
    )

    values.loc[international_thousands] = (
        values.loc[international_thousands]
        .str.replace(",", "", regex=False)
    )

    values.loc[integer_with_dots] = (
        values.loc[integer_with_dots]
        .str.replace(".", "", regex=False)
    )

    values.loc[integer_with_commas] = (
        values.loc[integer_with_commas]
        .str.replace(",", "", regex=False)
    )

    unresolved = ~(
        brazilian_decimal
        | brazilian_simple
        | international_thousands
        | integer_with_dots
        | integer_with_commas
    )

    values.loc[unresolved] = (
        values.loc[unresolved]
        .str.replace(",", ".", regex=False)
    )

    return pd.to_numeric(values, errors="coerce")


def standardize(dataframe, mapped, source_name):
    required_roles = ["entity", "kpi", "value"]

    missing_roles = [
        role
        for role in required_roles
        if not mapped.get(role)
    ]

    if missing_roles:
        raise ValueError(
            f"{source_name}: colunas obrigatórias não localizadas: "
            f"{', '.join(missing_roles)}. "
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

    if mapped.get("kpi_name"):
        output["KPI_NAME"] = (
            dataframe[mapped["kpi_name"]]
            .fillna("")
            .astype(str)
            .replace("nan", "")
        )
    else:
        output["KPI_NAME"] = ""

    output = output[
        output["ENTITY"].ne("")
        & output["KPI_CODE"].ne("")
    ].copy()

    return output.reset_index(drop=True)


def build_definitions(dataframe, mapped):
    columns = [
        "KPI_CODE",
        "DEF_KPI_NAME",
        "DEF_FORMULA",
        "DEF_UOM",
        "DEF_OWNER",
    ]

    if not mapped.get("kpi"):
        return pd.DataFrame(columns=columns)

    output = pd.DataFrame(
        {
            "KPI_CODE": dataframe[mapped["kpi"]].map(nkpi),
        }
    )

    definition_fields = {
        "DEF_KPI_NAME": "kpi_name",
        "DEF_FORMULA": "formula",
        "DEF_UOM": "uom",
        "DEF_OWNER": "owner",
    }

    for output_column, mapped_role in definition_fields.items():
        source_column = mapped.get(mapped_role)

        if source_column:
            output[output_column] = (
                dataframe[source_column]
                .fillna("")
                .astype(str)
                .replace("nan", "")
            )
        else:
            output[output_column] = ""

    output = output[
        output["KPI_CODE"].ne("")
    ].copy()

    output = (
        output.drop_duplicates(
            subset=["KPI_CODE"],
            keep="first",
        )
        .reset_index(drop=True)
    )

    return output[columns]


def aggregate_report(report_data, source_name):
    grouped = (
        report_data.groupby(
            ["ENTITY", "KPI_CODE"],
            as_index=False,
            dropna=False,
        )
        .agg(
            REPORT_VALUE=(
                "VALUE",
                lambda values: values.sum(min_count=1),
            ),
            REPORT_ROWS=("VALUE", "size"),
            REPORT_VALID_VALUES=("VALUE", "count"),
            KPI_NAME=("KPI_NAME", "first"),
        )
    )

    grouped["SOURCE"] = source_name

    return grouped


def aggregate_reference(reference_data):
    return (
        reference_data.groupby(
            ["ENTITY", "KPI_CODE"],
            as_index=False,
            dropna=False,
        )
        .agg(
            REFERENCE_VALUE=(
                "VALUE",
                lambda values: values.sum(min_count=1),
            ),
            REFERENCE_ROWS=("VALUE", "size"),
            REFERENCE_VALID_VALUES=("VALUE", "count"),
        )
    )


def compare(
    report_data,
    reference_data,
    definitions,
    source_name,
    absolute_tolerance,
    relative_tolerance,
):
    report_grouped = aggregate_report(
        report_data,
        source_name,
    )

    reference_grouped = aggregate_reference(
        reference_data,
    )

    result = report_grouped.merge(
        reference_grouped,
        on=["ENTITY", "KPI_CODE"],
        how="outer",
        indicator=True,
    )

    result["SOURCE"] = source_name

    result = result.merge(
        definitions,
        on="KPI_CODE",
        how="left",
    )

    result["DIFFERENCE"] = (
        result["REPORT_VALUE"]
        - result["REFERENCE_VALUE"]
    )

    valid_difference_pct = (
        result["REFERENCE_VALUE"].notna()
        & result["REFERENCE_VALUE"].abs().gt(absolute_tolerance)
    )

    result["DIFFERENCE_PCT"] = np.nan

    result.loc[valid_difference_pct, "DIFFERENCE_PCT"] = (
        result.loc[valid_difference_pct, "DIFFERENCE"]
        / result.loc[valid_difference_pct, "REFERENCE_VALUE"]
    )

    tolerance_limit = (
        absolute_tolerance
        + relative_tolerance
        * result["REFERENCE_VALUE"].abs()
    )

    within_tolerance = (
        result["REPORT_VALUE"].notna()
        & result["REFERENCE_VALUE"].notna()
        & result["DIFFERENCE"].abs().le(tolerance_limit)
    )

    report_without_numeric_value = (
        result["_merge"].ne("right_only")
        & result["REPORT_VALUE"].isna()
    )

    reference_without_numeric_value = (
        result["_merge"].ne("left_only")
        & result["REFERENCE_VALUE"].isna()
    )

    result["STATUS"] = np.select(
        [
            result["_merge"].eq("left_only"),
            result["_merge"].eq("right_only"),
            report_without_numeric_value,
            reference_without_numeric_value,
            within_tolerance,
        ],
        [
            "NÃO ESTÁ NA REFERÊNCIA",
            "SOMENTE NA REFERÊNCIA",
            "VALOR INVÁLIDO NO ARQUIVO ENVIADO",
            "VALOR INVÁLIDO NA REFERÊNCIA",
            "OK",
        ],
        default="DIVERGENTE",
    )

    result["KEY_USED"] = "ENTITY + KPI_CODE"

    column_order = [
        "SOURCE",
        "ENTITY",
        "KPI_CODE",
        "KPI_NAME",
        "DEF_KPI_NAME",
        "DEF_FORMULA",
        "DEF_UOM",
        "DEF_OWNER",
        "REPORT_VALUE",
        "REFERENCE_VALUE",
        "DIFFERENCE",
        "DIFFERENCE_PCT",
        "REPORT_ROWS",
        "REFERENCE_ROWS",
        "REPORT_VALID_VALUES",
        "REFERENCE_VALID_VALUES",
        "STATUS",
        "KEY_USED",
    ]

    for column in column_order:
        if column not in result.columns:
            result[column] = pd.NA

    return (
        result[column_order]
        .sort_values(
            ["STATUS", "ENTITY", "KPI_CODE"],
            na_position="last",
        )
        .reset_index(drop=True)
    )


def load_reference(data, filename="consolidador.xlsm"):
    try:
        reference_bytes = bytes(data)
        excel_source = io.BytesIO(reference_bytes)

        reference_sheet_names = pd.ExcelFile(
            excel_source,
            engine="openpyxl",
        ).sheet_names

        resolved_sheets = {
            REFERENCE_SHEET: resolve_sheet(
                reference_sheet_names,
                REFERENCE_SHEET,
            ),
            DEFINITION_SHEET: resolve_sheet(
                reference_sheet_names,
                DEFINITION_SHEET,
            ),
        }

        if not resolved_sheetsreturn {
                "error": (
                    "A aba SHAREPOINT não foi encontrada no "
                    f"arquivo de referência {filename}."
                ),
                "details": {
                    "reference_sheet_names": reference_sheet_names,
                    "resolved_sheets": resolved_sheets,
                },
            }

        raw = {}
        maps = {}
        headers = {}

        excel_source.seek(0)

        raw[REFERENCE_SHEET], headers[REFERENCE_SHEET] = read_sheet(
            excel_source,
            resolved_sheets[REFERENCE_SHEET],
        )

        maps[REFERENCE_SHEET] = build_mapping(
            raw[REFERENCE_SHEET],
            REFERENCE_SHEET,
        )

        reference_data = standardize(
            raw[REFERENCE_SHEET],
            maps[REFERENCE_SHEET],
            REFERENCE_SHEET,
        )

        if resolved_sheetsexcel_source.seek(0)

            raw[DEFINITION_SHEET], headers[DEFINITION_SHEET] = read_sheet(
                excel_source,
                resolved_sheets[DEFINITION_SHEET],
            )

            maps[DEFINITION_SHEET] = build_mapping(
                raw[DEFINITION_SHEET],
                DEFINITION_SHEET,
            )

            definitions = build_definitions(
                raw[DEFINITION_SHEET],
                maps[DEFINITION_SHEET],
            )
        else:
            headers[DEFINITION_SHEET] = None
            maps[DEFINITION_SHEET] = {
                role: None
                for role in ALIASES
            }

            definitions = pd.DataFrame(
                columns=[
                    "KPI_CODE",
                    "DEF_KPI_NAME",
                    "DEF_FORMULA",
                    "DEF_UOM",
                    "DEF_OWNER",
                ]
            )

        return {
            "error": None,
            "filename": filename,
            "reference_sheet_names": reference_sheet_names,
            "reference_sheets": resolved_sheets,
            "maps": maps,
            "headers": headers,
            "reference_data": reference_data,
            "definitions": definitions,
        }

    except Exception as error:
        return {
            "error": (
                "Falha ao carregar o arquivo de referência: "
                f"{type(error).__name__}: {error}"
            ),
            "details": {
                "filename": filename,
            },
        }


def run(
    upload_bytes,
    reference_bundle,
    absolute_tolerance=0.01,
    relative_tolerance=0.0001,
):
    try:
        if reference_bundle.get("error"):
            return reference_bundle

        uploaded_source = io.BytesIO(bytes(upload_bytes))

        uploaded_sheet_names = pd.ExcelFile(
            uploaded_source,
            engine="openpyxl",
        ).sheet_names

        uploaded_sheets = {
            report_name: resolve_sheet(
                uploaded_sheet_names,
                report_name,
            )
            for report_name in REPORTS
        }

        missing_reports = [
            report_name
            for report_name in REPORTS
            if not uploaded_sheets[report_name]
        ]

        if missing_reports:
            return {
                "error": (
                    "Abas obrigatórias não encontradas no arquivo enviado: "
                    + ", ".join(missing_reports)
                ),
                "details": {
                    "missing_sheets": missing_reports,
                    "uploaded_sheet_names": uploaded_sheet_names,
                    "resolved_sheets": uploaded_sheets,
                },
            }

        raw = {}
        maps = {}
        headers = {}

        for report_name in REPORTS:
            uploaded_source.seek(0)

            raw[report_name], headers[report_name] = read_sheet(
                uploaded_source,
                uploaded_sheets[report_name],
            )

            maps[report_name] = build_mapping(
                raw[report_name],
                report_name,
            )

        reference_maps = reference_bundle.get("maps", {})
        reference_headers = reference_bundle.get("headers", {})

        maps[REFERENCE_SHEET] = reference_maps.get(
            REFERENCE_SHEET,
            {},
        )

        maps[DEFINITION_SHEET] = reference_maps.get(
            DEFINITION_SHEET,
            {},
        )

        headers[REFERENCE_SHEET] = reference_headers.get(
            REFERENCE_SHEET,
        )

        headers[DEFINITION_SHEET] = reference_headers.get(
            DEFINITION_SHEET,
        )

        standardized_reports = {}

        for report_name in REPORTS:
            standardized_reports[report_name] = standardize(
                raw[report_name],
                maps[report_name],
                report_name,
            )

        reference_data = reference_bundle["reference_data"]
        definitions = reference_bundle["definitions"]

        results = []

        for report_name in REPORTS:
            results.append(
                compare(
                    standardized_reports[report_name],
                    reference_data,
                    definitions,
                    report_name,
                    absolute_tolerance,
                    relative_tolerance,
                )
            )

        return {
            "error": None,
            "uploaded_sheet_names": uploaded_sheet_names,
            "reference_sheet_names": reference_bundle.get(
                "reference_sheet_names",
                [],
            ),
            "uploaded_sheets": uploaded_sheets,
            "reference_sheets": reference_bundle.get(
                "reference_sheets",
                {},
            ),
            "maps": maps,
            "headers": headers,
            "definitions": definitions,
            "results": results,
            "reference_filename": reference_bundle.get(
                "filename",
                "consolidador.xlsm",
            ),
            "absolute_tolerance": float(absolute_tolerance),
            "relative_tolerance": float(relative_tolerance),
        }

    except Exception as error:
        return {
            "error": (
                "Falha durante o processamento: "
                f"{type(error).__name__}: {error}"
            ),
            "details": {
                "uploaded_sheet_names": locals().get(
                    "uploaded_sheet_names",
                    [],
                ),
                "uploaded_sheets": locals().get(
                    "uploaded_sheets",
                    {},
                ),
                "maps": locals().get(
                    "maps",
                    {},
                ),
                "headers": locals().get(
                    "headers",
                    {},
                ),
            },
        }


def report(payload):
    output = io.BytesIO()

    result = pd.concat(
        payload["results"],
        ignore_index=True,
    )

    summary = (
        result.groupby(
            ["SOURCE", "STATUS"],
            dropna=False,
        )
        .size()
        .reset_index(name="QUANTIDADE")
        .sort_values(["SOURCE", "STATUS"])
        .reset_index(drop=True)
    )

    parameters = pd.DataFrame(
        {
            "PARAMETRO": [
                "CHAVE",
                "ARQUIVO_REFERENCIA",
                "ABA_REFERENCIA",
                "TOLERANCIA_ABSOLUTA",
                "TOLERANCIA_RELATIVA",
            ],
            "VALOR": [
                "ENTITY + KPI_CODE",
                payload.get(
                    "reference_filename",
                    "consolidador.xlsm",
                ),
                payload.get(
                    "reference_sheets",
                    {},
                ).get(
                    REFERENCE_SHEET,
                    REFERENCE_SHEET,
                ),
                payload["absolute_tolerance"],
                payload["relative_tolerance"],
            ],
        }
    )

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
    ) as writer:
        result.to_excel(
            writer,
            index=False,
            sheet_name="Comparacao_Completa",
        )

        result[
            result["STATUS"].eq("OK")
        ].to_excel(
            writer,
            index=False,
            sheet_name="OK",
        )

        result[
            result["STATUS"].ne("OK")
        ].to_excel(
            writer,
            index=False,
            sheet_name="Achados",
        )

        summary.to_excel(
            writer,
            index=False,
            sheet_name="Resumo",
        )

        payload["definitions"].to_excel(
            writer,
            index=False,
            sheet_name="Definition_Book",
        )

        parameters.to_excel(
            writer,
            index=False,
            sheet_name="Parametros",
        )

        for sheet_name, dataframe in {
            "Comparacao_Completa": result,
            "OK": result[result["STATUS"].eq("OK")],
            "Achados": result[result["STATUS"].ne("OK")],
            "Resumo": summary,
            "Definition_Book": payload["definitions"],
            "Parametros": parameters,
        }.items():
            worksheet = writer.sheets[sheet_name]

            if len(dataframe.columns):
                worksheet.freeze_panes(1, 0)
                worksheet.autofilter(
                    0,
                    0,
                    max(len(dataframe), 1),
                    len(dataframe.columns) - 1,
                )

                for column_index, column_name in enumerate(dataframe.columns):
                    if dataframe.empty:
                        maximum_length = len(str(column_name))
                    else:
                        maximum_length = max(
                            len(str(column_name)),
                            dataframe[column_name]
                            .fillna("")
                            .astype(str)
                            .map(len)
                            .max(),
                        )

                    worksheet.set_column(
                        column_index,
                        column_index,
                        min(maximum_length + 2, 50),
                    )

    output.seek(0)
    return output.getvalue()
