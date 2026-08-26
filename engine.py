"""
KPI Validator - Engine
=======================

Motor de leitura, normalizacao e comparacao de KPIs entre um arquivo de
referencia (consolidador.xlsm - aba SHAREPOINT) e um arquivo enviado pelo
usuario contendo as abas LOGS, BOPS e SL.

Chave de comparacao: exclusivamente ENTITY + KPI_CODE.

Este modulo nao depende do Streamlit e pode ser testado isoladamente.
Funcoes publicas principais: `run` e `report`.
"""

from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime
from typing import Optional

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Configuracao
# ---------------------------------------------------------------------------

# Abas obrigatorias no arquivo de REFERENCIA (consolidador.xlsm)
REFERENCE_REQUIRED = ["SHAREPOINT"]
# Aba opcional no arquivo de REFERENCIA
REFERENCE_OPTIONAL = ["DEFINITION BOOK"]
# Abas obrigatorias no arquivo enviado pelo usuario
UPLOAD_REQUIRED = ["LOGS", "BOPS", "SL"]

KEY_LABEL = "ENTITY + KPI_CODE"

STATUS_OK = "OK"
STATUS_DIVERGENT = "DIVERGENTE"
STATUS_NOT_IN_REFERENCE = "NÃO ESTÁ NA REFERÊNCIA"
STATUS_ONLY_IN_REFERENCE = "SOMENTE NA REFERÊNCIA"

# Aliases (PT-BR / EN) reconhecidos para cada papel de coluna.
# A busca e feita por normalizacao (sem acento, minusculo, sem pontuacao).
ALIASES = {
    "entity": [
        "entity", "entidade", "unit name", "unidade", "bu",
        "location", "plant", "dc", "unit",
    ],
    "kpi": [
        "kpi code", "kpi_code", "codigo kpi", "código kpi", "cod kpi",
        "kpicode",
    ],
    "value": [
        "current_value ac", "current value ac", "valor anaplan",
        "valor origem", "actual", "ac", "value", "valor",
        "valor sharepoint", "current value",
    ],
    "kpi_name": [
        "kpi name", "nome kpi", "description", "descrição", "descricao",
    ],
    "formula": [
        "formula", "fórmula", "regra",
    ],
    "uom": [
        "unit_of_measure", "unit of measure", "uom",
    ],
    "owner": [
        "owner", "responsavel", "responsável",
    ],
}

# Nomes alternativos aceitos para localizar as abas por titulo (alem do
# nome exato do papel). Usados por `find_sheet`.
SHEET_ALIASES = {
    "SHAREPOINT": ["SHAREPOINT", "SHARE POINT", "REFERENCIA", "REFERENCE", "BASE"],
    "DEFINITION BOOK": ["DEFINITION BOOK", "DEFINITIONBOOK", "DEFINITIONS", "DEFINITION"],
    "LOGS": ["LOGS", "LOG"],
    "BOPS": ["BOPS", "BOP"],
    "SL": ["SL", "SERVICE LEVEL"],
}


# ---------------------------------------------------------------------------
# Normalizacao de texto e numeros
# ---------------------------------------------------------------------------

def strip_accents(text) -> str:
    """Remove acentos de uma string."""
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", str(text))
        if not unicodedata.combining(ch)
    )


def norm_text(value) -> str:
    """Normaliza texto: remove acentos, colapsa espacos e converte para maiusculo."""
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", strip_accents(value).strip()).upper()


def norm_col(value) -> str:
    """Normaliza nome de coluna para comparacao de aliases (minusculo,
    somente letras/digitos separados por espaco)."""
    cleaned = re.sub(r"[^a-z0-9]+", " ", strip_accents(value).strip().lower())
    return re.sub(r"\s+", " ", cleaned).strip()


def norm_kpi_code(value) -> str:
    """Extrai/normaliza um codigo de KPI no formato XX-Y0000 (ex: SL-K0031)."""
    text = norm_text(value)
    match = re.search(r"\b([A-Z]{2,3}-[KR]\d{3,5}(?:_\d{4})?)\b", text)
    return match.group(1) if match else text


def parse_number(series: pd.Series) -> pd.Series:
    """Converte uma coluna para numerico, tratando separador decimal com
    virgula e separador de milhar com ponto (padrao BR)."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    text = series.astype(str).str.strip().replace(
        {"": np.nan, "-": np.nan, "nan": np.nan, "None": np.nan}
    )
    # Padrao brasileiro: 1.234.567,89
    br_pattern = text.str.contains(r"^-?\d{1,3}(?:\.\d{3})+,\d+$", regex=True, na=False)
    text.loc[br_pattern] = (
        text.loc[br_pattern]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    # Demais casos: apenas troca virgula decimal por ponto
    text.loc[~br_pattern] = text.loc[~br_pattern].str.replace(",", ".", regex=False)
    return pd.to_numeric(text, errors="coerce")


# ---------------------------------------------------------------------------
# Localizacao de abas e cabecalhos
# ---------------------------------------------------------------------------

def find_sheet(sheet_names: list[str], target: str) -> Optional[str]:
    """Localiza o nome real de uma aba na planilha, tolerando pequenas
    variacoes (case, acentos, espacos, nomes alternativos)."""
    for name in sheet_names:
        if norm_text(name) == norm_text(target):
            return name

    aliases = SHEET_ALIASES.get(target.upper(), [target])
    for alias in aliases:
        for name in sheet_names:
            if norm_text(alias) in norm_text(name):
                return name
    return None


def detect_header_row(raw: pd.DataFrame, max_scan: int = 80) -> int:
    """Detecta a linha de cabecalho dentro das primeiras `max_scan` linhas,
    pontuando linhas que parecem conter titulos de KPI_CODE e ENTITY."""
    best_row, best_score = 0, -1
    limit = min(max_scan, len(raw))

    for i in range(limit):
        values = [norm_text(v) for v in raw.iloc[i]]
        has_kpi_col = any("KPI" in v and ("CODE" in v or "CODIGO" in v) for v in values)
        has_entity_col = any(
            v in ["ENTITY", "ENTIDADE", "UNIT NAME", "UNIDADE", "BU", "LOCATION"]
            for v in values
        )
        filled_cells = sum(bool(v) for v in values)
        score = 1000 * has_kpi_col + 800 * has_entity_col + filled_cells
        if score > best_score:
            best_row, best_score = i, score

    return best_row


def read_sheet(data: bytes, sheet_name: str) -> tuple[pd.DataFrame, int]:
    """Le uma aba do arquivo Excel, detecta o cabecalho automaticamente e
    retorna (dataframe_padronizado, indice_da_linha_de_cabecalho)."""
    raw = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name, header=None, engine="openpyxl")
    header_row = detect_header_row(raw)

    columns = []
    seen: dict[str, int] = {}
    for j, value in enumerate(raw.iloc[header_row]):
        col_name = str(value).strip() if not pd.isna(value) else f"COL_{j + 1}"
        seen[col_name] = seen.get(col_name, 0) + 1
        if seen[col_name] > 1:
            col_name = f"{col_name}_{seen[col_name]}"
        columns.append(col_name)

    body = raw.iloc[header_row + 1:].copy()
    body.columns = columns
    body = body.dropna(how="all").reset_index(drop=True)
    return body, header_row + 1


# ---------------------------------------------------------------------------
# Mapeamento de colunas por alias
# ---------------------------------------------------------------------------

def find_column(df: pd.DataFrame, role: str) -> Optional[str]:
    """Encontra a coluna do dataframe que corresponde a um papel (entity,
    kpi, value, ...) usando a lista de aliases."""
    normalized_cols = {norm_col(c): c for c in df.columns}

    # Match exato primeiro
    for alias in ALIASES[role]:
        key = norm_col(alias)
        if key in normalized_cols:
            return normalized_cols[key]

    # Match parcial (alias contido no nome da coluna)
    for col_key, original in normalized_cols.items():
        if any(norm_col(alias) in col_key for alias in ALIASES[role]):
            return original

    return None


def build_mapping(df: pd.DataFrame, source: str) -> dict:
    """Monta o dicionario de mapeamento {papel: coluna} para um dataframe,
    aplicando uma preferencia especifica para a coluna de valor conforme a
    origem (referencia ou relatorio)."""
    mapping = {role: find_column(df, role) for role in ALIASES}

    normalized_cols = {norm_col(c): c for c in df.columns}
    value_preference = (
        ["current value ac", "current value", "valor sharepoint", "ac", "value", "valor"]
        if source == "SHAREPOINT"
        else ["valor anaplan", "ac", "actual", "valor origem", "value", "valor"]
    )
    for candidate in value_preference:
        key = norm_col(candidate)
        if key in normalized_cols:
            mapping["value"] = normalized_cols[key]
            break

    return mapping


# ---------------------------------------------------------------------------
# Padronizacao de dados
# ---------------------------------------------------------------------------

def standardize(df: pd.DataFrame, mapping: dict, source: str) -> pd.DataFrame:
    """Converte um dataframe bruto no formato padrao SOURCE/ENTITY/KPI_CODE/VALUE.
    Lanca ValueError com mensagem clara se colunas essenciais nao forem localizadas."""
    missing = [role for role in ("entity", "kpi", "value") if not mapping.get(role)]
    if missing:
        raise ValueError(
            f"{source}: colunas obrigatorias não localizadas: {missing}. "
            f"Colunas detectadas na aba: {list(df.columns)}"
        )

    out = pd.DataFrame({
        "SOURCE": source,
        "ENTITY": df[mapping["entity"]].map(norm_text),
        "KPI_CODE": df[mapping["kpi"]].map(norm_kpi_code),
        "VALUE": parse_number(df[mapping["value"]]),
    })
    out["KPI_NAME"] = df[mapping["kpi_name"]].astype(str) if mapping.get("kpi_name") else ""
    return out[(out["ENTITY"] != "") & (out["KPI_CODE"] != "")]


def standardize_definitions(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """Extrai as colunas de enriquecimento da aba DEFINITION BOOK (opcional)."""
    if not mapping.get("kpi"):
        return pd.DataFrame(columns=["KPI_CODE", "DEF_KPI_NAME", "DEF_FORMULA", "DEF_UOM", "DEF_OWNER"])

    out = pd.DataFrame({"KPI_CODE": df[mapping["kpi"]].map(norm_kpi_code)})
    out["DEF_KPI_NAME"] = df[mapping["kpi_name"]].astype(str) if mapping.get("kpi_name") else ""
    out["DEF_FORMULA"] = df[mapping["formula"]].astype(str) if mapping.get("formula") else ""
    out["DEF_UOM"] = df[mapping["uom"]].astype(str) if mapping.get("uom") else ""
    out["DEF_OWNER"] = df[mapping["owner"]].astype(str) if mapping.get("owner") else ""
    return out[out["KPI_CODE"] != ""].drop_duplicates("KPI_CODE")


# ---------------------------------------------------------------------------
# Comparacao
# ---------------------------------------------------------------------------

def compare(report_df: pd.DataFrame, reference_df: pd.DataFrame, definitions_df: pd.DataFrame,
            abs_tol: float, rel_tol: float) -> pd.DataFrame:
    """Compara um relatorio (LOGS/BOPS/SL) contra a referencia (SHAREPOINT)
    usando a chave ENTITY + KPI_CODE, agregando duplicidades pela soma."""
    report_agg = report_df.groupby(["ENTITY", "KPI_CODE"]).agg(
        REPORT_VALUE=("VALUE", "sum"),
        REPORT_ROWS=("VALUE", "size"),
        KPI_NAME=("KPI_NAME", "first"),
        SOURCE=("SOURCE", "first"),
    ).reset_index()

    reference_agg = reference_df.groupby(["ENTITY", "KPI_CODE"]).agg(
        REFERENCE_VALUE=("VALUE", "sum"),
        REFERENCE_ROWS=("VALUE", "size"),
    ).reset_index()

    merged = report_agg.merge(
        reference_agg, on=["ENTITY", "KPI_CODE"], how="outer", indicator=True
    ).merge(definitions_df, on="KPI_CODE", how="left")

    merged["DIFFERENCE"] = merged["REPORT_VALUE"] - merged["REFERENCE_VALUE"]
    merged["DIFFERENCE_PCT"] = np.where(
        merged["REFERENCE_VALUE"].abs() > abs_tol,
        merged["DIFFERENCE"] / merged["REFERENCE_VALUE"],
        np.nan,
    )

    within_tolerance = (
        merged["DIFFERENCE"].abs() <= abs_tol + rel_tol * merged["REFERENCE_VALUE"].abs()
    ).fillna(False)

    merged["STATUS"] = np.select(
        [
            merged["_merge"].eq("left_only"),
            merged["_merge"].eq("right_only"),
            within_tolerance,
        ],
        [STATUS_NOT_IN_REFERENCE, STATUS_ONLY_IN_REFERENCE, STATUS_OK],
        default=STATUS_DIVERGENT,
    )
    merged["KEY_USED"] = KEY_LABEL
    return merged.drop(columns="_merge")


# ---------------------------------------------------------------------------
# Carregamento de arquivos (referencia e upload)
# ---------------------------------------------------------------------------

def _load_workbook_sheets(data: bytes, required: list[str], optional: list[str]):
    """Localiza e le as abas necessarias/opcionais de um arquivo Excel.
    Retorna (raw_by_role, mapping_by_role, header_by_role, resolved_names, missing)."""
    sheet_names = pd.ExcelFile(io.BytesIO(data), engine="openpyxl").sheet_names

    resolved = {}
    for role in required + optional:
        resolved[role] = find_sheet(sheet_names, role)

    missing = [role for role in required if not resolved[role]]

    raw, mapping, headers = {}, {}, {}
    for role, sheet_name in resolved.items():
        if sheet_name:
            raw[role], headers[role] = read_sheet(data, sheet_name)
            mapping[role] = build_mapping(raw[role], role)

    return raw, mapping, headers, resolved, missing


def load_reference(data: bytes, filename: str = "consolidador.xlsm") -> dict:
    """Carrega e padroniza o arquivo de referencia (SHAREPOINT + DEFINITION BOOK opcional)."""
    raw, mapping, headers, resolved, missing = _load_workbook_sheets(
        data, REFERENCE_REQUIRED, REFERENCE_OPTIONAL
    )
    if missing:
        return {
            "error": f"Arquivo de referência inválido. Abas não encontradas: {', '.join(missing)}.",
            "resolved": resolved,
        }

    try:
        reference_df = standardize(raw["SHAREPOINT"], mapping["SHAREPOINT"], "SHAREPOINT")
    except ValueError as exc:
        return {"error": str(exc), "resolved": resolved, "mapping": mapping}

    if "DEFINITION BOOK" in raw:
        definitions_df = standardize_definitions(raw["DEFINITION BOOK"], mapping["DEFINITION BOOK"])
    else:
        definitions_df = pd.DataFrame(columns=["KPI_CODE", "DEF_KPI_NAME", "DEF_FORMULA", "DEF_UOM", "DEF_OWNER"])

    return {
        "error": None,
        "filename": filename,
        "reference_df": reference_df,
        "definitions_df": definitions_df,
        "resolved": resolved,
        "mapping": mapping,
        "headers": headers,
    }


def run(upload_bytes: bytes, reference_bundle: dict, abs_tol: float = 0.01, rel_tol: float = 0.0001) -> dict:
    """Orquestra a leitura do arquivo enviado e a comparacao contra a
    referencia ja carregada (ver `load_reference`).

    Retorna um dicionario com:
      - error: mensagem de erro (str) ou None
      - resolved / mapping / headers: diagnostico de abas e colunas
      - results: lista de dataframes de comparacao (um por aba LOGS/BOPS/SL)
      - definitions_df: dataframe de enriquecimento (Definition Book)
      - params: parametros usados na comparacao (para o relatorio)
    """
    if reference_bundle.get("error"):
        return {"error": f"Referência inválida: {reference_bundle['error']}"}

    raw, mapping, headers, resolved, missing = _load_workbook_sheets(
        upload_bytes, UPLOAD_REQUIRED, []
    )
    if missing:
        return {
            "error": f"Arquivo enviado não contém as abas obrigatórias: {', '.join(missing)}.",
            "resolved": resolved,
        }

    reference_df = reference_bundle["reference_df"]
    definitions_df = reference_bundle["definitions_df"]

    results = []
    try:
        for role in UPLOAD_REQUIRED:
            report_df = standardize(raw[role], mapping[role], role)
            results.append(compare(report_df, reference_df, definitions_df, abs_tol, rel_tol))
    except ValueError as exc:
        return {"error": str(exc), "resolved": resolved, "mapping": mapping}

    return {
        "error": None,
        "resolved": resolved,
        "mapping": mapping,
        "headers": headers,
        "results": results,
        "definitions_df": definitions_df,
        "params": {
            "key": KEY_LABEL,
            "reference_file": reference_bundle.get("filename", "consolidador.xlsm"),
            "abs_tol": abs_tol,
            "rel_tol": rel_tol,
            "generated_at": datetime.now(),
        },
    }


# ---------------------------------------------------------------------------
# Exportacao do relatorio
# ---------------------------------------------------------------------------

def report(payload: dict) -> bytes:
    """Gera o relatorio Excel final com as abas:
    Comparacao_Completa, OK, Achados, Resumo, Definition_Book, Parametros."""
    combined = pd.concat(payload["results"], ignore_index=True)
    params = payload["params"]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        combined.to_excel(writer, index=False, sheet_name="Comparacao_Completa")
        combined[combined["STATUS"] == STATUS_OK].to_excel(writer, index=False, sheet_name="OK")
        combined[combined["STATUS"] != STATUS_OK].to_excel(writer, index=False, sheet_name="Achados")

        summary = combined.groupby(["SOURCE", "STATUS"]).size().reset_index(name="QUANTIDADE")
        summary.to_excel(writer, index=False, sheet_name="Resumo")

        payload["definitions_df"].to_excel(writer, index=False, sheet_name="Definition_Book")

        params_df = pd.DataFrame([
            {"PARAMETRO": "Chave de comparação", "VALOR": params["key"]},
            {"PARAMETRO": "Arquivo de referência", "VALOR": params["reference_file"]},
            {"PARAMETRO": "Tolerância absoluta", "VALOR": params["abs_tol"]},
            {"PARAMETRO": "Tolerância relativa", "VALOR": params["rel_tol"]},
            {"PARAMETRO": "Gerado em", "VALOR": params["generated_at"].strftime("%Y-%m-%d %H:%M:%S")},
        ])
        params_df.to_excel(writer, index=False, sheet_name="Parametros")

    return output.getvalue()
