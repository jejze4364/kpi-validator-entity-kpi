from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Any, Iterable
import re
import unicodedata

import numpy as np
import pandas as pd

STATUS_ORDER = ["OK", "DIVERGENTE", "NAO ESTA NO SHAREPOINT", "SOMENTE NO SHAREPOINT", "VALOR INVALIDO"]

ABA_SHAREPOINT = "SHAREPOINT"
ABA_LOGS = "LOGS"
ABA_BOPS = "BOPS"
ABA_SL = "SL"
ABA_ANAPLAN = "ANAPLAN"
ABA_DEFINITION_BOOK = "DEFINITION BOOK"
ABAS_ANAPLAN = (ABA_LOGS, ABA_BOPS, ABA_SL)

# Indices Python, equivalentes a F/L/R, B/H e E no VBA.
COLUNA_KPI_SHAREPOINT = 5
COLUNA_ENTITY_SHAREPOINT = 11
COLUNA_FILTRO_SHAREPOINT = 17
COLUNA_ENTITY_ANAPLAN = 1
COLUNA_KPI_ANAPLAN = 7
COLUNA_KPI_DEFINITION_BOOK = 4

ROTULO_AMBOS = "SHAREPOINT + ANAPLAN"
ROTULO_SOMENTE_SHAREPOINT = "SOMENTE SHAREPOINT"
ROTULO_SOMENTE_ANAPLAN = "SOMENTE ANAPLAN"

CAMINHO_CONSOLIDADOR = Path(__file__).resolve().parent / "Consolidador 2.0.xlsm"


def texto_limpo(valor: Any) -> str:
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass
    texto = str(valor).replace("\xa0", " ").replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", texto).strip()


def normalizar_texto(valor: Any) -> str:
    texto = unicodedata.normalize("NFKD", texto_limpo(valor).upper())
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip()


def chave_comparacao(valor: Any) -> str:
    return normalizar_texto(valor)


def montar_chave(entidade: Any, kpi: Any) -> str:
    return f"{chave_comparacao(entidade)}|{chave_comparacao(kpi)}"


def _bytes_arquivo(fonte: Any) -> bytes:
    if isinstance(fonte, bytes):
        return fonte
    if isinstance(fonte, bytearray):
        return bytes(fonte)
    if hasattr(fonte, "getvalue"):
        return fonte.getvalue()
    if isinstance(fonte, (str, Path)):
        return Path(fonte).read_bytes()
    if hasattr(fonte, "read"):
        pos = fonte.tell() if hasattr(fonte, "tell") else None
        dados = fonte.read()
        if pos is not None and hasattr(fonte, "seek"):
            fonte.seek(pos)
        return dados
    raise TypeError("Fonte de arquivo Excel inválida.")


def _excel_file(fonte: Any) -> pd.ExcelFile:
    return pd.ExcelFile(BytesIO(_bytes_arquivo(fonte)), engine="openpyxl")


def nome_aba_real(nomes_abas: Iterable[str], nome_procurado: str) -> str | None:
    alvo = normalizar_texto(nome_procurado)
    return next((nome for nome in nomes_abas if normalizar_texto(nome) == alvo), None)


def cabecalhos_unicos(cabecalhos: Iterable[Any]) -> list[str]:
    resultado: list[str] = []
    contagem: dict[str, int] = {}
    for indice, cabecalho in enumerate(cabecalhos):
        nome = texto_limpo(cabecalho) or f"COLUNA_{indice + 1:03d}"
        chave = normalizar_texto(nome) or f"COLUNA_{indice + 1:03d}"
        contagem[chave] = contagem.get(chave, 0) + 1
        if contagem[chave] > 1:
            nome = f"{nome}_{contagem[chave]}"
        resultado.append(nome)
    return resultado


def remover_vazios(df: pd.DataFrame) -> pd.DataFrame:
    return df.dropna(axis=0, how="all").dropna(axis=1, how="all").reset_index(drop=True)


def linha_cabecalho_real(fonte: Any, nome_aba: str) -> int:
    previa = pd.read_excel(BytesIO(_bytes_arquivo(fonte)), sheet_name=nome_aba, header=None, nrows=50, engine="openpyxl")
    for indice, linha in previa.iterrows():
        if any("KPI_CODE" in normalizar_texto(v) or "KPI CODE" in normalizar_texto(v) for v in linha.tolist()):
            return int(indice)
    return 0


def _ler_aba(fonte: Any, nome_aba: str, header: int = 0) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(_bytes_arquivo(fonte)), sheet_name=nome_aba, header=header, engine="openpyxl")
    df.columns = cabecalhos_unicos(df.columns)
    return remover_vazios(df)


def validar_consolidador(fonte: Any | None = None) -> tuple[Any, pd.ExcelFile]:
    origem = fonte if fonte is not None else CAMINHO_CONSOLIDADOR
    if isinstance(origem, (str, Path)) and not Path(origem).exists():
        raise FileNotFoundError(f"O arquivo '{Path(origem).name}' não foi encontrado no projeto.")
    excel = _excel_file(origem)
    obrigatorias = [ABA_SHAREPOINT, ABA_LOGS, ABA_BOPS, ABA_SL, ABA_DEFINITION_BOOK]
    ausentes = [aba for aba in obrigatorias if not nome_aba_real(excel.sheet_names, aba)]
    if ausentes:
        raise ValueError("Abas obrigatórias não encontradas: " + ", ".join(ausentes))
    return origem, excel


def ler_sharepoint(fonte: Any, excel: pd.ExcelFile) -> pd.DataFrame:
    return _ler_aba(fonte, nome_aba_real(excel.sheet_names, ABA_SHAREPOINT), 0)


def ler_definition_book(fonte: Any, excel: pd.ExcelFile) -> pd.DataFrame:
    return _ler_aba(fonte, nome_aba_real(excel.sheet_names, ABA_DEFINITION_BOOK), 0)


def ler_base_anaplan(fonte: Any, excel: pd.ExcelFile, nome_base: str) -> pd.DataFrame:
    aba = nome_aba_real(excel.sheet_names, nome_base)
    return _ler_aba(fonte, aba, linha_cabecalho_real(fonte, aba))


def _validar_indice(df: pd.DataFrame, indice: int, descricao: str) -> None:
    if len(df.columns) <= indice:
        raise ValueError(f"A aba não possui a coluna esperada para {descricao} (posição {indice + 1}).")


def valores_unicos_ordenados(serie: pd.Series) -> list[str]:
    valores: dict[str, str] = {}
    for valor in serie.tolist():
        texto = texto_limpo(valor)
        chave = chave_comparacao(texto)
        if chave and chave not in valores:
            valores[chave] = texto
    return sorted(valores.values(), key=normalizar_texto)


def carregar_opcoes_template(fonte: Any | None = None) -> list[str]:
    origem, excel = validar_consolidador(fonte)
    sp = ler_sharepoint(origem, excel)
    _validar_indice(sp, COLUNA_FILTRO_SHAREPOINT, "filtro/país do SHAREPOINT")
    return valores_unicos_ordenados(sp.iloc[:, COLUNA_FILTRO_SHAREPOINT])


def carregar_kpis_entidades(filtro_selecionado: str, fonte: Any | None = None) -> tuple[list[str], list[str]]:
    origem, excel = validar_consolidador(fonte)
    sp = ler_sharepoint(origem, excel)
    for indice, desc in ((COLUNA_KPI_SHAREPOINT, "KPI"), (COLUNA_ENTITY_SHAREPOINT, "entidade"), (COLUNA_FILTRO_SHAREPOINT, "filtro/país")):
        _validar_indice(sp, indice, desc)
    mascara = sp.iloc[:, COLUNA_FILTRO_SHAREPOINT].map(chave_comparacao).eq(chave_comparacao(filtro_selecionado))
    filtrado = sp.loc[mascara]
    return valores_unicos_ordenados(filtrado.iloc[:, COLUNA_KPI_SHAREPOINT]), valores_unicos_ordenados(filtrado.iloc[:, COLUNA_ENTITY_SHAREPOINT])


def _selecionados(valores: Iterable[Any]) -> set[str]:
    return {chave_comparacao(v) for v in valores if chave_comparacao(v)}


def _filtrar_sharepoint(sp: pd.DataFrame, filtro: str, kpis: Iterable[str], entidades: Iterable[str]) -> pd.DataFrame:
    ks, es = _selecionados(kpis), _selecionados(entidades)
    mascara = (
        sp.iloc[:, COLUNA_FILTRO_SHAREPOINT].map(chave_comparacao).eq(chave_comparacao(filtro))
        & sp.iloc[:, COLUNA_KPI_SHAREPOINT].map(chave_comparacao).isin(ks)
        & sp.iloc[:, COLUNA_ENTITY_SHAREPOINT].map(chave_comparacao).isin(es)
    )
    return sp.loc[mascara].copy().reset_index(drop=True)


def _dicionario_filtro(sp: pd.DataFrame) -> dict[str, str]:
    resultado: dict[str, str] = {}
    for _, linha in sp.iterrows():
        kpi, entidade, filtro = linha.iloc[COLUNA_KPI_SHAREPOINT], linha.iloc[COLUNA_ENTITY_SHAREPOINT], texto_limpo(linha.iloc[COLUNA_FILTRO_SHAREPOINT])
        chave = montar_chave(entidade, kpi)
        if chave != "|" and filtro and chave not in resultado:
            resultado[chave] = filtro
    return resultado


def _chaves_sp(sp: pd.DataFrame) -> set[str]:
    return {montar_chave(l.iloc[COLUNA_ENTITY_SHAREPOINT], l.iloc[COLUNA_KPI_SHAREPOINT]) for _, l in sp.iterrows()}


def _carregar_bases(fonte: Any, excel: pd.ExcelFile) -> dict[str, pd.DataFrame]:
    return {nome: ler_base_anaplan(fonte, excel, nome) for nome in ABAS_ANAPLAN}


def _filtrar_bases(bases: dict[str, pd.DataFrame], kpis: Iterable[str], entidades: Iterable[str]) -> dict[str, pd.DataFrame]:
    ks, es = _selecionados(kpis), _selecionados(entidades)
    resultado: dict[str, pd.DataFrame] = {}
    for nome, df in bases.items():
        _validar_indice(df, COLUNA_KPI_ANAPLAN, f"KPI da aba {nome}")
        _validar_indice(df, COLUNA_ENTITY_ANAPLAN, f"entidade da aba {nome}")
        mascara = df.iloc[:, COLUNA_KPI_ANAPLAN].map(chave_comparacao).isin(ks) & df.iloc[:, COLUNA_ENTITY_ANAPLAN].map(chave_comparacao).isin(es)
        resultado[nome] = df.loc[mascara].copy().reset_index(drop=True)
    return resultado


def _chaves_anaplan(bases: dict[str, pd.DataFrame]) -> set[str]:
    resultado: set[str] = set()
    for df in bases.values():
        for _, linha in df.iterrows():
            resultado.add(montar_chave(linha.iloc[COLUNA_ENTITY_ANAPLAN], linha.iloc[COLUNA_KPI_ANAPLAN]))
    return resultado


def _identidade_cabecalho(cabecalho: Any, indice: int) -> tuple[str, str]:
    nome = texto_limpo(cabecalho) or f"COLUNA_{indice + 1:03d}"
    chave = chave_comparacao(nome)
    if chave == "SOURCE BASE":
        return "SOURCE BASE ORIGEM", "SOURCE BASE ORIGEM"
    return chave, nome


def _consolidar_anaplan(bases_filtradas: dict[str, pd.DataFrame], bases_completas: dict[str, pd.DataFrame], chaves_sp: set[str], filtros: dict[str, str]) -> pd.DataFrame:
    ordem: list[str] = []
    nomes: dict[str, str] = {}
    for df in bases_completas.values():
        for i, coluna in enumerate(df.columns):
            chave, nome = _identidade_cabecalho(coluna, i)
            if chave not in nomes:
                nomes[chave] = nome
                ordem.append(chave)
    colunas = ["SOURCE BASE"] + [nomes[c] for c in ordem] + ["FILTRO / PAÍS (SHAREPOINT)", "CORRESPONDÊNCIA"]
    registros: list[dict[str, Any]] = []
    for base, df in bases_filtradas.items():
        for _, linha in df.iterrows():
            registro = {c: None for c in colunas}
            registro["SOURCE BASE"] = base
            for i, coluna in enumerate(df.columns):
                chave, _ = _identidade_cabecalho(coluna, i)
                registro[nomes[chave]] = linha.iloc[i]
            chave = montar_chave(linha.iloc[COLUNA_ENTITY_ANAPLAN], linha.iloc[COLUNA_KPI_ANAPLAN])
            registro["FILTRO / PAÍS (SHAREPOINT)"] = filtros.get(chave, "NÃO LOCALIZADO NO SHAREPOINT")
            registro["CORRESPONDÊNCIA"] = ROTULO_AMBOS if chave in chaves_sp else ROTULO_SOMENTE_ANAPLAN
            registros.append(registro)
    return pd.DataFrame(registros, columns=colunas)


def _filtrar_definition(df: pd.DataFrame, kpis: Iterable[str]) -> pd.DataFrame:
    _validar_indice(df, COLUNA_KPI_DEFINITION_BOOK, "KPI do DEFINITION BOOK")
    ks = _selecionados(kpis)
    resultado = df.loc[df.iloc[:, COLUNA_KPI_DEFINITION_BOOK].map(chave_comparacao).isin(ks)].copy()
    chave = resultado.iloc[:, COLUNA_KPI_DEFINITION_BOOK].map(chave_comparacao)
    return resultado.loc[~chave.duplicated()].reset_index(drop=True)


def _largura(serie: pd.Series, cabecalho: Any) -> int:
    try:
        maior = serie.fillna("").astype("string").str.len().max() if not serie.empty else 0
        maior = 0 if pd.isna(maior) else int(maior)
        return min(max(len(str(cabecalho)), maior) + 2, 45)
    except Exception:
        return 20


def _formatar(writer: pd.ExcelWriter, aba: str, df: pd.DataFrame) -> None:
    wb, ws = writer.book, writer.sheets[aba]
    header = wb.add_format({"bold": True, "bg_color": "#FFCD00", "font_color": "#000000", "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1})
    ws.freeze_panes(1, 0)
    ws.set_row(0, 30)
    for i, coluna in enumerate(df.columns):
        ws.write(0, i, str(coluna), header)
        ws.set_column(i, i, _largura(df[coluna], coluna))
    if len(df.columns):
        ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)


def gerar_template_do_consolidador(filtro_selecionado: str, kpis_selecionados: Iterable[str], entidades_selecionadas: Iterable[str], fonte: Any | None = None) -> dict[str, Any]:
    kpis_selecionados, entidades_selecionadas = list(kpis_selecionados), list(entidades_selecionadas)
    if not texto_limpo(filtro_selecionado):
        raise ValueError("Selecione uma opção da coluna R.")
    if not kpis_selecionados:
        raise ValueError("Selecione pelo menos um KPI.")
    if not entidades_selecionadas:
        raise ValueError("Selecione pelo menos uma entidade.")
    origem, excel = validar_consolidador(fonte)
    sp_completo = ler_sharepoint(origem, excel)
    definition = ler_definition_book(origem, excel)
    bases_completas = _carregar_bases(origem, excel)
    sp = _filtrar_sharepoint(sp_completo, filtro_selecionado, kpis_selecionados, entidades_selecionadas)
    bases = _filtrar_bases(bases_completas, kpis_selecionados, entidades_selecionadas)
    chaves_sp, chaves_anaplan = _chaves_sp(sp), _chaves_anaplan(bases)
    sp["CORRESPONDÊNCIA"] = [ROTULO_AMBOS if montar_chave(l.iloc[COLUNA_ENTITY_SHAREPOINT], l.iloc[COLUNA_KPI_SHAREPOINT]) in chaves_anaplan else ROTULO_SOMENTE_SHAREPOINT for _, l in sp.iterrows()]
    anaplan = _consolidar_anaplan(bases, bases_completas, chaves_sp, _dicionario_filtro(sp_completo))
    definition = _filtrar_definition(definition, kpis_selecionados)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for aba, df in ((ABA_SHAREPOINT, sp), (ABA_ANAPLAN, anaplan), (ABA_DEFINITION_BOOK, definition)):
            df.to_excel(writer, sheet_name=aba, index=False)
            _formatar(writer, aba, df)
    return {"arquivo": output.getvalue(), "quantidade_sharepoint": len(sp), "quantidade_anaplan": len(anaplan), "quantidade_definition": len(definition)}


def blank_template() -> bytes:
    sp = pd.DataFrame(columns=["MACRO_PASTA", "YEAR_MONTH", "BUSINESS AREA", "THEME", "KPI_CODE", "KPI_NAME", "CURRENCY", "UNIT_OF_MEASURE", "CURRENT_VALUE AC", "ENTITY", "ENTITY_CODE", "BUS_AREA_CODE", "MONTH_CODE", "SOURCE_FILE", "COUNTRY"])
    anaplan = pd.DataFrame(columns=["SOURCE BASE", "ENTITY", "KPI_CODE"])
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for aba, df in ((ABA_SHAREPOINT, sp), (ABA_ANAPLAN, anaplan)):
            df.to_excel(writer, sheet_name=aba, index=False)
            _formatar(writer, aba, df)
    return output.getvalue()


def _coluna_por_nomes(df: pd.DataFrame, nomes: Iterable[str], obrigatoria: bool = True) -> str | None:
    mapa = {normalizar_texto(c).replace("_", " "): c for c in df.columns}
    for nome in nomes:
        chave = normalizar_texto(nome).replace("_", " ")
        if chave in mapa:
            return mapa[chave]
    for nome in nomes:
        chave = normalizar_texto(nome).replace("_", " ")
        for normalizado, original in mapa.items():
            if chave and chave in normalizado:
                return original
    if obrigatoria:
        raise ValueError("Coluna não encontrada. Esperado: " + ", ".join(nomes))
    return None


def _numero(serie: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")
    texto = serie.astype("string").str.strip().str.replace(" ", "", regex=False)
    ambos = texto.str.contains(",", na=False) & texto.str.contains(".", regex=False, na=False)
    texto = texto.where(~ambos, texto.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    so_virgula = texto.str.contains(",", na=False) & ~texto.str.contains(".", regex=False, na=False)
    texto = texto.where(~so_virgula, texto.str.replace(",", ".", regex=False))
    return pd.to_numeric(texto, errors="coerce")


def run(file_bytes: bytes, absolute_tolerance: float = 0.01, relative_tolerance: float = 0.0001) -> dict[str, Any]:
    try:
        excel = _excel_file(file_bytes)
        aba_sp = nome_aba_real(excel.sheet_names, ABA_SHAREPOINT)
        aba_an = nome_aba_real(excel.sheet_names, ABA_ANAPLAN)
        if not aba_sp or not aba_an:
            return {"error": "O arquivo deve conter as abas SHAREPOINT e ANAPLAN."}
        sp = _ler_aba(file_bytes, aba_sp, 0)
        an = _ler_aba(file_bytes, aba_an, 0)
        sp_kpi = _coluna_por_nomes(sp, ["KPI_CODE", "KPI CODE"])
        sp_ent = _coluna_por_nomes(sp, ["ENTITY", "ENTIDADE"])
        sp_val = _coluna_por_nomes(sp, ["CURRENT_VALUE AC", "CURRENT VALUE AC", "VALOR SHAREPOINT", "VALUE", "VALOR"])
        sp_country = _coluna_por_nomes(sp, ["COUNTRY", "FILTRO / PAÍS", "FILTRO / PAIS", "PAIS"], False)
        sp_bo = _coluna_por_nomes(sp, ["BUSINESS_OPERATION", "BUSINESS OPERATION", "BUSINESS AREA"], False)
        an_kpi = _coluna_por_nomes(an, ["KPI_CODE", "KPI CODE"])
        an_ent = _coluna_por_nomes(an, ["ENTITY", "ENTIDADE"])
        an_val = _coluna_por_nomes(an, ["AC MTH", "CURRENT_VALUE AC", "CURRENT VALUE AC", "VALUE", "VALOR", "ACTUAL"])
        an_country = _coluna_por_nomes(an, ["COUNTRY", "FILTRO / PAÍS (SHAREPOINT)", "FILTRO / PAIS (SHAREPOINT)", "PAIS"], False)
        an_bo = _coluna_por_nomes(an, ["BUSINESS_OPERATION", "BUSINESS OPERATION", "BUSINESS AREA"], False)
        an_tipo = _coluna_por_nomes(an, ["SOURCE BASE", "TIPO"], False)

        spx = pd.DataFrame({"ENTITY": sp[sp_ent].map(texto_limpo), "KPI_CODE": sp[sp_kpi].map(texto_limpo), "SHAREPOINT_VALUE": _numero(sp[sp_val]), "COUNTRY": sp[sp_country].map(texto_limpo) if sp_country else "", "BUSINESS_OPERATION": sp[sp_bo].map(texto_limpo) if sp_bo else ""})
        anx = pd.DataFrame({"ENTITY": an[an_ent].map(texto_limpo), "KPI_CODE": an[an_kpi].map(texto_limpo), "ANAPLAN_VALUE": _numero(an[an_val]), "COUNTRY_AN": an[an_country].map(texto_limpo) if an_country else "", "BUSINESS_OPERATION_AN": an[an_bo].map(texto_limpo) if an_bo else "", "TIPO": an[an_tipo].map(texto_limpo) if an_tipo else "ANAPLAN"})
        spx["_KEY"] = spx["ENTITY"].map(normalizar_texto) + "|" + spx["KPI_CODE"].map(normalizar_texto)
        anx["_KEY"] = anx["ENTITY"].map(normalizar_texto) + "|" + anx["KPI_CODE"].map(normalizar_texto)
        merged = spx.merge(anx, on="_KEY", how="outer", suffixes=("_SP", "_AN"))
        merged["ENTITY"] = merged["ENTITY_SP"].fillna(merged["ENTITY_AN"])
        merged["KPI_CODE"] = merged["KPI_CODE_SP"].fillna(merged["KPI_CODE_AN"])
        merged["COUNTRY"] = merged["COUNTRY"].fillna(merged["COUNTRY_AN"]).fillna("")
        merged["BUSINESS_OPERATION"] = merged["BUSINESS_OPERATION"].fillna(merged["BUSINESS_OPERATION_AN"]).fillna("")
        merged["TIPO"] = merged["TIPO"].fillna("SHAREPOINT")
        presente_sp = merged["ENTITY_SP"].notna()
        presente_an = merged["ENTITY_AN"].notna()
        validos = merged["SHAREPOINT_VALUE"].notna() & merged["ANAPLAN_VALUE"].notna()
        diferenca = (merged["SHAREPOINT_VALUE"] - merged["ANAPLAN_VALUE"]).abs()
        limite = float(absolute_tolerance) + float(relative_tolerance) * np.maximum(merged["SHAREPOINT_VALUE"].abs(), merged["ANAPLAN_VALUE"].abs())
        merged["STATUS"] = np.select([presente_sp & ~presente_an, ~presente_sp & presente_an, presente_sp & presente_an & ~validos, presente_sp & presente_an & validos & (diferenca <= limite)], ["SOMENTE NO SHAREPOINT", "NAO ESTA NO SHAREPOINT", "VALOR INVALIDO", "OK"], default="DIVERGENTE")
        merged["DIFERENCA_ABSOLUTA"] = diferenca
        colunas = ["COUNTRY", "BUSINESS_OPERATION", "TIPO", "ENTITY", "KPI_CODE", "SHAREPOINT_VALUE", "ANAPLAN_VALUE", "DIFERENCA_ABSOLUTA", "STATUS"]
        data = merged[colunas].copy()
        return {"results": [data], "sharepoint_rows": len(sp), "anaplan_rows": len(an)}
    except Exception as exc:
        return {"error": str(exc)}


def report(payload: dict[str, Any], filtered: pd.DataFrame | None = None) -> bytes:
    data = filtered.copy() if filtered is not None else pd.concat(payload.get("results", []), ignore_index=True)
    grupos = [c for c in ["COUNTRY", "BUSINESS_OPERATION", "TIPO", "STATUS"] if c in data.columns]
    summary = data.groupby(grupos, dropna=False).size().reset_index(name="QUANTIDADE") if grupos else pd.DataFrame({"QUANTIDADE": [len(data)]})
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        data.to_excel(writer, index=False, sheet_name="Comparacao")
        summary.to_excel(writer, index=False, sheet_name="Resumo")
        _formatar(writer, "Comparacao", data)
        _formatar(writer, "Resumo", summary)
    return output.getvalue()
