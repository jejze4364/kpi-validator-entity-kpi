from io import BytesIO
import re
import unicodedata

import numpy as np
import pandas as pd


REQUIRED_SHEETS = ["SHAREPOINT", "ANAPLAN"]

STATUS_ORDER = [
    "OK",
    "DIVERGENTE",
    "NAO ESTA NO SHAREPOINT",
    "SOMENTE NO SHAREPOINT",
    "VALOR INVALIDO",
]

ALIASES = {
    "entity": [
        "entity",
        "entidade",
        "unit name",
        "unidade",
        "location",
        "plant",
        "dc",
    ],
    "kpi": [
        "kpi code",
        "kpi_code",
        "codigo kpi",
        "cod kpi",
    ],
    "value": [
        "current_value ac",
        "current value ac",
        "valor anaplan",
        "valor sharepoint",
        "actual value",
        "actual",
        "ac",
        "value",
        "valor",
    ],
    "country": [
        "filtro pais sharepoint",
        "filtro / pais sharepoint",
        "country",
        "pais",
    ],
    "business_operation": [
        "business operation",
        "business operations",
        "business_operation",
        "business area",
        "business_area",
    ],
    "source_base": [
        "source base",
        "source_base",
        "origem",
        "base origem",
    ],
}


def _noacc(value):
    """Remove acentos de um texto."""
    return "".join(
        character
        for character in unicodedata.normalize("NFKD", str(value))
        if not unicodedata.combining(character)
    )


def _text(value):
    """Converte um valor em texto limpo."""
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _norm(value):
    """Normaliza textos usados nas chaves de comparação."""
    return _noacc(_text(value)).upper()


def _col(value):
    """Normaliza nomes de colunas para localização por aliases."""
    return re.sub(r"[^a-z0-9]+", " ", _noacc(_text(value)).lower()).strip()


def _kpi(value):
    """Normaliza e extrai o código do KPI."""
    text = _norm(value)
    match = re.search(r"\b([A-Z]{2,4}-[A-Z]?\d{3,6}(?:_\d{4})?)\b", text)
    return match.group(1) if match else text


def _number(series):
    """Converte uma série em número, aceitando formatos brasileiro e internacional."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    text = series.astype(str).str.strip().replace(
        {
            "": np.nan,
            "-": np.nan,
            "nan": np.nan,
            "None": np.nan,
            "NONE": np.nan,
        }
    )

    brazilian = text.str.contains(
        r"^-?\d{1,3}(?:\.\d{3})+,\d+$",
        regex=True,
        na=False,
    )

    text.loc[brazilian] = (
        text.loc[brazilian]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    text.loc[~brazilian] = text.loc[~brazilian].str.replace(",", ".", regex=False)
    return pd.to_numeric(text, errors="coerce")


def _header_index(raw):
    """Localiza automaticamente a linha mais provável do cabeçalho."""
    best_row = 0
    best_score = -1

    for index in range(min(80, len(raw))):
        cells = [_col(value) for value in raw.iloc[index].tolist()]
        score = 0
        score += 1000 * any(
            "kpi" in value and ("code" in value or "codigo" in value)
            for value in cells
        )
        score += 800 * any(value in ALIASES["entity"] for value in cells)
        score += 500 * any(value in ALIASES["value"] for value in cells)
        score += 300 * any(value in ALIASES["source_base"] for value in cells)
        score += sum(bool(value) for value in cells)

        if score > best_score:
            best_row = index
            best_score = score

    return best_row


def _read(data, sheet):
    """Lê uma aba e aplica o cabeçalho detectado automaticamente."""
    raw = pd.read_excel(
        BytesIO(data),
        sheet_name=sheet,
        header=None,
        dtype=object,
        engine="openpyxl",
    )

    raw = raw.dropna(how="all").dropna(axis=1, how="all")

    if raw.empty:
        return pd.DataFrame()

    header = _header_index(raw)
    columns = []
    used = {}

    for index, value in enumerate(raw.iloc[header].tolist()):
        name = _text(value) or f"COL_{index + 1}"
        used[name] = used.get(name, 0) + 1
        columns.append(name if used[name] == 1 else f"{name}_{used[name]}")

    table = raw.iloc[header + 1 :].copy()
    table.columns = columns

    return table.dropna(how="all").reset_index(drop=True)


def _find(table, field):
    """Localiza uma coluna usando seus aliases conhecidos."""
    normalized = {_col(column): column for column in table.columns}

    for alias in ALIASES[field]:
        normalized_alias = _col(alias)
        if normalized_alias in normalized:
            return normalized[normalized_alias]

    for alias in ALIASES[field]:
        normalized_alias = _col(alias)
        for normalized_name, original in normalized.items():
            if normalized_alias and normalized_alias in normalized_name:
                return original

    return None


def _standardize(table, source):
    """Padroniza as colunas necessárias para a comparação."""
    if table.empty:
        raise ValueError(f"{source}: a aba está vazia.")

    entity = _find(table, "entity")
    kpi = _find(table, "kpi")
    value = _find(table, "value")
    country = _find(table, "country")
    operation = _find(table, "business_operation")
    source_base = _find(table, "source_base") if source == "ANAPLAN" else None

    missing = [
        name
        for name, column in (
            ("ENTITY", entity),
            ("KPI_CODE", kpi),
            ("VALUE", value),
        )
        if column is None
    ]

    if source == "ANAPLAN" and source_base is None:
        missing.append("SOURCE BASE")

    if missing:
        raise ValueError(
            f"{source}: colunas não localizadas: {', '.join(missing)}. "
            f"Colunas detectadas: {list(table.columns)}"
        )

    output = pd.DataFrame(
        {
            "ENTITY": table[entity].map(_norm),
            "KPI_CODE": table[kpi].map(_kpi),
            "VALUE": _number(table[value]),
            "COUNTRY": table[country].map(_text) if country else "",
            "BUSINESS_OPERATION": table[operation].map(_text) if operation else "",
            "SOURCE_BASE": (
                table[source_base].map(_norm)
                if source_base
                else pd.Series("", index=table.index, dtype=object)
            ),
        }
    )

    output = output[
        (output["ENTITY"] != "")
        & (output["KPI_CODE"] != "")
    ].reset_index(drop=True)

    return output


def _first_nonempty(series):
    """Retorna o primeiro valor preenchido de uma série."""
    values = series.dropna().map(_text)
    values = values[values != ""]
    return values.iloc[0] if not values.empty else ""


def run(data, abs_tolerance=0.01, rel_tolerance=0.0001):
    """Processa o template e compara SHAREPOINT e ANAPLAN."""
    try:
        names = pd.ExcelFile(BytesIO(data), engine="openpyxl").sheet_names
        missing_sheets = [sheet for sheet in REQUIRED_SHEETS if sheet not in names]

        if missing_sheets:
            return {
                "error": "Abas obrigatórias não encontradas: "
                + ", ".join(missing_sheets)
            }

        sharepoint = _standardize(_read(data, "SHAREPOINT"), "SHAREPOINT")
        anaplan = _standardize(_read(data, "ANAPLAN"), "ANAPLAN")

        sp = (
            sharepoint.groupby(
                ["ENTITY", "KPI_CODE"],
                as_index=False,
                dropna=False,
            )
            .agg(
                SHAREPOINT_VALUE=("VALUE", "sum"),
                COUNTRY=("COUNTRY", _first_nonempty),
                BUSINESS_OPERATION=("BUSINESS_OPERATION", _first_nonempty),
            )
        )

        an = (
            anaplan.groupby(
                ["ENTITY", "KPI_CODE", "SOURCE_BASE"],
                as_index=False,
                dropna=False,
            )
            .agg(
                ANAPLAN_VALUE=("VALUE", "sum"),
                AN_COUNTRY=("COUNTRY", _first_nonempty),
                AN_BUSINESS_OPERATION=("BUSINESS_OPERATION", _first_nonempty),
            )
        )

        result = an.merge(
            sp,
            on=["ENTITY", "KPI_CODE"],
            how="outer",
            indicator=True,
        )

        result["SOURCE_BASE"] = result["SOURCE_BASE"].fillna("")
        result["TIPO"] = result["SOURCE_BASE"]

        result["COUNTRY"] = (
            result["COUNTRY"]
            .replace("", np.nan)
            .fillna(result["AN_COUNTRY"])
            .fillna("")
        )

        result["BUSINESS_OPERATION"] = (
            result["BUSINESS_OPERATION"]
            .replace("", np.nan)
            .fillna(result["AN_BUSINESS_OPERATION"])
            .fillna("")
        )

        result["DIFFERENCE"] = (
            result["ANAPLAN_VALUE"] - result["SHAREPOINT_VALUE"]
        )

        valid = (
            result["ANAPLAN_VALUE"].notna()
            & result["SHAREPOINT_VALUE"].notna()
        )

        tolerance = (
            float(abs_tolerance)
            + float(rel_tolerance) * result["SHAREPOINT_VALUE"].abs()
        )

        ok = valid & (result["DIFFERENCE"].abs() <= tolerance)

        result["STATUS"] = np.select(
            [
                result["_merge"].eq("left_only"),
                result["_merge"].eq("right_only"),
                valid & ~ok,
                ok,
            ],
            [
                "NAO ESTA NO SHAREPOINT",
                "SOMENTE NO SHAREPOINT",
                "DIVERGENTE",
                "OK",
            ],
            default="VALOR INVALIDO",
        )

        result = result[
            [
                "COUNTRY",
                "BUSINESS_OPERATION",
                "SOURCE_BASE",
                "TIPO",
                "ENTITY",
                "KPI_CODE",
                "SHAREPOINT_VALUE",
                "ANAPLAN_VALUE",
                "DIFFERENCE",
                "STATUS",
            ]
        ]

        result = result.sort_values(
            ["COUNTRY", "BUSINESS_OPERATION", "SOURCE_BASE", "ENTITY", "KPI_CODE"],
            kind="stable",
        ).reset_index(drop=True)

        return {
            "error": None,
            "results": [result],
        }

    except Exception as exc:
        return {
            "error": f"Não foi possível processar o arquivo: {exc}"
        }


def report(payload, filtered=None):
    """Gera o relatório Excel para download."""
    data = (
        filtered.copy()
        if filtered is not None
        else pd.concat(payload["results"], ignore_index=True)
    )

    summary = (
        data.groupby(
            ["COUNTRY", "BUSINESS_OPERATION", "SOURCE_BASE", "STATUS"],
            dropna=False,
        )
        .size()
        .reset_index(name="QUANTIDADE")
    )

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        data.to_excel(writer, index=False, sheet_name="Comparacao")
        summary.to_excel(writer, index=False, sheet_name="Resumo")

        comparison_sheet = writer.sheets["Comparacao"]
        summary_sheet = writer.sheets["Resumo"]

        comparison_sheet.freeze_panes(1, 0)
        summary_sheet.freeze_panes(1, 0)
        comparison_sheet.autofilter(0, 0, len(data), len(data.columns) - 1)
        summary_sheet.autofilter(0, 0, len(summary), len(summary.columns) - 1)

        for index, column in enumerate(data.columns):
            width = max(
                len(str(column)),
                data[column].astype(str).map(len).max() if not data.empty else 0,
            )
            comparison_sheet.set_column(index, index, min(width + 2, 45))

        for index, column in enumerate(summary.columns):
            width = max(
                len(str(column)),
                summary[column].astype(str).map(len).max() if not summary.empty else 0,
            )
            summary_sheet.set_column(index, index, min(width + 2, 45))

    return output.getvalue()
