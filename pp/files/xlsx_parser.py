"""
Responsabilidade: extrair todos os pares (cliente, vendedor) do XLSX de comissões.

Percorre todas as abas, localiza dinamicamente as colunas CLIENTE e VENDEDOR,
e retorna os registros encontrados junto com avisos de abas problemáticas.
"""

from dataclasses import dataclass
from typing import List, Tuple

import pandas as pd


@dataclass(frozen=True)
class XLSXRecord:
    cliente: str
    vendedor: str
    sheet_name: str


def _find_header_row(df: pd.DataFrame) -> Tuple[int | None, int | None, int | None]:
    """
    Procura a linha do cabeçalho que contém CLIENTE e VENDEDOR.
    Retorna (row_idx, cliente_col_idx, vendedor_col_idx) ou (None, None, None).
    """
    for row_idx, row in df.iterrows():
        row_upper = [
            str(v).upper().strip() if pd.notna(v) else "" for v in row
        ]
        if "CLIENTE" in row_upper and "VENDEDOR" in row_upper:
            return row_idx, row_upper.index("CLIENTE"), row_upper.index("VENDEDOR")
    return None, None, None


def _extract_records_from_sheet(
    df: pd.DataFrame, sheet_name: str
) -> Tuple[List[XLSXRecord], str | None]:
    """
    Extrai registros de uma aba.
    Retorna (registros, mensagem_de_aviso_ou_None).
    """
    header_row_idx, cliente_col, vendedor_col = _find_header_row(df)

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

        records.append(
            XLSXRecord(cliente=cliente_str, vendedor=vendedor_str, sheet_name=sheet_name)
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
