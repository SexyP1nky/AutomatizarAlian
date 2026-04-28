"""
Responsabilidade: extrair todos os registros (cliente, vendedor, apolice) do XLSX de comissões.

Percorre todas as abas, localiza dinamicamente as colunas CLIENTE, VENDEDOR e APÓLICE,
e retorna os registros encontrados junto com avisos de abas problemáticas.
"""

import re
from dataclasses import dataclass
from typing import List, Tuple

import pandas as pd


# Palavras-chave para detectar a coluna de apólice no cabeçalho
_APOLICE_KEYWORDS = [
    "APÓLICE", "APOLICE",
    "Nº APÓLICE", "N° APÓLICE", "Nº APOLICE", "N° APOLICE",
    "NÚMERO DA APÓLICE", "NUMERO DA APOLICE",
    "Nº DA APÓLICE", "N° DA APÓLICE",
    "Nº DA APOLICE", "N° DA APOLICE",
]


@dataclass(frozen=True)
class XLSXRecord:
    cliente: str
    vendedor: str
    sheet_name: str
    apolice: str


def is_excluded_vendor(name: str) -> bool:
    """
    Retorna True se o nome do vendedor deve ser excluído da tabela.
    Exclui APENAS o nome exato "SUELANE" (case-insensitive, strip).
    Variações como "SUELANE J." ou "SUELANEJ" NÃO são excluídas.
    """
    return name.strip().upper() == "SUELANE"


def _normalize_apolice(raw: str) -> str:
    """
    Normaliza número de apólice: remove espaços, pontos, hífens.
    """
    if not raw:
        return ""
    return re.sub(r"[\s.\-/]", "", str(raw).strip())


def _find_header_row(
    df: pd.DataFrame,
) -> Tuple[int | None, int | None, int | None, int | None]:
    """
    Procura a linha do cabeçalho que contém CLIENTE e VENDEDOR.
    Também tenta encontrar a coluna APÓLICE.
    Retorna (row_idx, cliente_col_idx, vendedor_col_idx, apolice_col_idx).
    apolice_col_idx pode ser None se não encontrado.
    """
    for row_idx, row in df.iterrows():
        row_upper = [
            str(v).upper().strip() if pd.notna(v) else "" for v in row
        ]
        if "CLIENTE" in row_upper and "VENDEDOR" in row_upper:
            cliente_col = row_upper.index("CLIENTE")
            vendedor_col = row_upper.index("VENDEDOR")

            # Tentar encontrar coluna de apólice
            apolice_col = None
            for col_idx, cell_text in enumerate(row_upper):
                if not cell_text:
                    continue
                for keyword in _APOLICE_KEYWORDS:
                    if keyword in cell_text:
                        apolice_col = col_idx
                        break
                if apolice_col is not None:
                    break

            return row_idx, cliente_col, vendedor_col, apolice_col
    return None, None, None, None


def _extract_records_from_sheet(
    df: pd.DataFrame, sheet_name: str
) -> Tuple[List[XLSXRecord], str | None]:
    """
    Extrai registros de uma aba.
    Retorna (registros, mensagem_de_aviso_ou_None).
    """
    header_row_idx, cliente_col, vendedor_col, apolice_col = _find_header_row(df)

    if header_row_idx is None:
        return [], f"Aba '{sheet_name}': cabeçalho CLIENTE/VENDEDOR não encontrado — aba ignorada."

    records: List[XLSXRecord] = []

    for _, row in df.iloc[header_row_idx + 1 :].iterrows():
        cliente_val = row.iloc[cliente_col]
        vendedor_val = row.iloc[vendedor_col]

        if pd.isna(cliente_val) or pd.isna(vendedor_val):
            continue

        cliente_str = str(cliente_val).strip()
        vendedor_str = str(vendedor_val).strip()

        if not cliente_str or not vendedor_str:
            continue

        # Filtrar vendedor exatamente "SUELANE" (não variações)
        if is_excluded_vendor(vendedor_str):
            continue

        # Extrair apólice se a coluna foi encontrada
        apolice_str = ""
        if apolice_col is not None:
            apolice_val = row.iloc[apolice_col]
            if pd.notna(apolice_val):
                # Proteção contra o Pandas convertendo colunas numéricas com NaN para float (ex: 12345.0)
                if isinstance(apolice_val, float) and apolice_val.is_integer():
                    raw_str = str(int(apolice_val))
                else:
                    raw_str = str(apolice_val).strip()
                apolice_str = _normalize_apolice(raw_str)

        records.append(
            XLSXRecord(
                cliente=cliente_str,
                vendedor=vendedor_str,
                sheet_name=sheet_name,
                apolice=apolice_str,
            )
        )

    return records, None


def extract_client_vendor_pairs(
    xlsx_path: str,
) -> Tuple[List[XLSXRecord], List[str]]:
    """
    Lê todas as abas do XLSX e retorna (registros, lista_de_avisos).
    """
    all_records: List[XLSXRecord] = []
    warnings: List[str] = []

    xl = pd.ExcelFile(xlsx_path)

    for sheet_name in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None, dtype=str)
        records, warning = _extract_records_from_sheet(df, sheet_name)
        all_records.extend(records)
        if warning:
            warnings.append(warning)

    return all_records, warnings
