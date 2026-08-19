from io import BytesIO
import re
import unicodedata
import numpy as np
import pandas as pd

REQUIRED_SHEETS = ["SHAREPOINT", "ANAPLAN"]
STATUS_ORDER = ["OK", "DIVERGENTE", "NAO ESTA NO SHAREPOINT", "SOMENTE NO SHAREPOINT", "VALOR INVALIDO"]

ALIASES = {
    "entity": ["entity", "entidade", "unit name", "unidade", "location", "plant", "dc"],
    "kpi": ["kpi code", "kpi_code", "codigo kpi", "cod kpi"],
    "value": ["current_value ac", "current value ac", "valor anaplan", "actual value", "actual", "ac", "value", "valor"],
    "country": ["country", "pais"],
    "business_operation": ["business operation", "business operations", "business_operation", "business area", "business_area"],
}


def _noacc(value):
    return "".join(c for c in unicodedata.normalize("NFKD", str(value)) if not unicodedata.combining(c))


def _text(value):
    return "" if pd.isna(value) else re.sub(r"\s+", " ", str(value).strip())


def _norm(value):
    return _noacc(_text(value)).upper()


def _col(value):
    return re.sub(r"[^a-z0-9]+", " ", _noacc(_text(value)).lower()).strip()


def _kpi(value):
    text = _norm(value)
    match = re.search(r"\b([A-Z]{2,4}-[A-Z]?\d{3,6}(?:_\d{4})?)\b", text)
    return match.group(1) if match else text


def _number(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    text = series.astype(str).str.strip().replace({"": np.nan, "-": np.nan, "nan": np.nan, "None": np.nan})
    brazilian = text.str.contains(r"^-?\d{1,3}(?:\.\d{3})+,\d+$", regex=True, na=False)
    text.loc[brazilian] = text.loc[brazilian].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    text.loc[~brazilian] = text.loc[~brazilian].str.replace(",", ".", regex=False)
    return pd.to_numeric(text, errors="coerce")


def _header_index(raw):
    best_row, best_score = 0, -1
    for i in range(min(80, len(raw))):
        cells = [_col(v) for v in raw.iloc[i].tolist()]
        score = 0
        score += 1000 * any("kpi" in v and ("code" in v or "codigo" in v) for v in cells)
        score += 800 * any(v in ALIASES["entity"] for v in cells)
        score += 500 * any(v in ALIASES["value"] for v in cells)
        score += sum(bool(v) for v in cells)
        if score > best_score:
            best_row, best_score = i, score
    return best_row


def _read(data, sheet):
    raw = pd.read_excel(BytesIO(data), sheet_name=sheet, header=None, dtype=object, engine="openpyxl")
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        return pd.DataFrame()
    header = _header_index(raw)
    columns, used = [], {}
    for index, value in enumerate(raw.iloc[header].tolist()):
        name = _text(value) or f"COL_{index + 1}"
        used[name] = used.get(name, 0) + 1
        columns.append(name if used[name] == 1 else f"{name}_{used[name]}")
    table = raw.iloc[header + 1:].copy()
    table.columns = columns
    return table.dropna(how="all").reset_index(drop=True)


def _find(table, field):
    normalized = {_col(column): column for column in table.columns}
    for alias in ALIASES[field]:
        if alias in normalized:
            return normalized[alias]
    for alias in ALIASES[field]:
        for normalized_name, original in normalized.items():
            if alias in normalized_name:
                return original
    return None


def _standardize(table, source):
    entity = _find(table, "entity")
    kpi = _find(table, "kpi")
    value = _find(table, "value")
    if source == "ANAPLAN":
        if entity is None and len(table.columns) >= 2:
            entity = table.columns[1]
        if kpi is None and len(table.columns) >= 5:
            kpi = table.columns[4]
        if value is None:
            excluded = {entity, kpi}
            candidates = [(table[c].map(lambda v: pd.to_numeric(str(v).replace(",", "."), errors="coerce")).notna().sum(), c) for c in table.columns if c not in excluded]
            value = max(candidates, default=(0, None))[1]
    missing = [name for name, column in (("ENTITY", entity), ("KPI_CODE", kpi), ("VALUE", value)) if column is None]
    if missing:
        raise ValueError(f"{source}: colunas nao localizadas: {', '.join(missing)}. Detectadas: {list(table.columns)}")
    country = _find(table, "country")
    operation = _find(table, "business_operation")
    output = pd.DataFrame({
        "ENTITY": table[entity].map(_norm),
        "KPI_CODE": table[kpi].map(_kpi),
        "VALUE": _number(table[value]),
        "COUNTRY": table[country].map(_text) if country else "",
        "BUSINESS_OPERATION": table[operation].map(_text) if operation else "",
    })
    return output[(output["ENTITY"] != "") & (output["KPI_CODE"] != "")].reset_index(drop=True)


def _type(row):
    text = " ".join((_norm(row["KPI_CODE"]), _norm(row["BUSINESS_OPERATION"])))
    if re.search(r"(^|\W)SL([\W_]|$)|SERVICE LEVEL", text):
        return "SL"
    if "BOPS" in text or "BREWERY OPERATIONS" in text:
        return "BOPS"
    return "PADRAO"


def run(data, abs_tolerance=0.01, rel_tolerance=0.0001):
    try:
        names = pd.ExcelFile(BytesIO(data), engine="openpyxl").sheet_names
        missing = [sheet for sheet in REQUIRED_SHEETS if sheet not in names]
        if missing:
            return {"error": "Abas obrigatorias nao encontradas: " + ", ".join(missing)}
        sharepoint = _standardize(_read(data, "SHAREPOINT"), "SHAREPOINT")
        anaplan = _standardize(_read(data, "ANAPLAN"), "ANAPLAN")
        sp = sharepoint.groupby(["ENTITY", "KPI_CODE"], as_index=False).agg(
            SHAREPOINT_VALUE=("VALUE", "sum"), COUNTRY=("COUNTRY", "first"), BUSINESS_OPERATION=("BUSINESS_OPERATION", "first")
        )
        an = anaplan.groupby(["ENTITY", "KPI_CODE"], as_index=False).agg(
            ANAPLAN_VALUE=("VALUE", "sum"), AN_COUNTRY=("COUNTRY", "first"), AN_BUSINESS_OPERATION=("BUSINESS_OPERATION", "first")
        )
        result = an.merge(sp, on=["ENTITY", "KPI_CODE"], how="outer", indicator=True)
        result["COUNTRY"] = result["COUNTRY"].replace("", np.nan).fillna(result["AN_COUNTRY"]).fillna("")
        result["BUSINESS_OPERATION"] = result["BUSINESS_OPERATION"].replace("", np.nan).fillna(result["AN_BUSINESS_OPERATION"]).fillna("")
        result["TIPO"] = result.apply(_type, axis=1)
        result["DIFFERENCE"] = result["ANAPLAN_VALUE"] - result["SHAREPOINT_VALUE"]
        valid = result["ANAPLAN_VALUE"].notna() & result["SHAREPOINT_VALUE"].notna()
        tolerance = float(abs_tolerance) + float(rel_tolerance) * result["SHAREPOINT_VALUE"].abs()
        ok = valid & (result["DIFFERENCE"].abs() <= tolerance)
        result["STATUS"] = np.select(
            [result["_merge"].eq("left_only"), result["_merge"].eq("right_only"), valid & ~ok, ok],
            ["NAO ESTA NO SHAREPOINT", "SOMENTE NO SHAREPOINT", "DIVERGENTE", "OK"],
            default="VALOR INVALIDO",
        )
        result = result[["COUNTRY", "BUSINESS_OPERATION", "TIPO", "ENTITY", "KPI_CODE", "SHAREPOINT_VALUE", "ANAPLAN_VALUE", "DIFFERENCE", "STATUS"]]
        return {"error": None, "results": [result]}
    except Exception as exc:
        return {"error": f"Nao foi possivel processar o arquivo: {exc}"}


def report(payload, filtered=None):
    data = filtered.copy() if filtered is not None else pd.concat(payload["results"], ignore_index=True)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        data.to_excel(writer, index=False, sheet_name="Comparacao")
        data.groupby(["COUNTRY", "BUSINESS_OPERATION", "STATUS"], dropna=False).size().reset_index(name="QUANTIDADE").to_excel(writer, index=False, sheet_name="Resumo")
    return output.getvalue()
