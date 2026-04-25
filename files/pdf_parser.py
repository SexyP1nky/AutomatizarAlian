"""
Responsabilidade: extrair do PDF os registros onde COMISSÃO é negativa.

Retorna lista de PDFRecord com (segurado, inicio_vig).
O valor da comissão é descartado — interessa apenas identificar o segurado.

Estratégia de extração:
- pdfplumber extrai cada linha da tabela como uma mini-tabela separada
- Linhas de dados têm P/E = 'P' ou 'E' na coluna 0
- Posições fixas confirmadas na estrutura do PDF:
    col 3 → INICIO VIG
    col 4 → SEGURADO
    col 12 → COMISSÃO
"""

from dataclasses import dataclass
from typing import List

import pdfplumber


# Posições de coluna confirmadas via inspeção do PDF real
_COL_INICIO_VIG = 3
_COL_SEGURADO = 4
_COL_COMISSAO = 12
_MIN_COLS = 13
_VALID_PE_VALUES = {"P", "E"}


@dataclass(frozen=True)
class PDFRecord:
    segurado: str
    inicio_vig: str  # formato original do PDF: DD/MM/YYYY


def _parse_br_float(value: str) -> float:
    """
    Converte string no formato brasileiro (ex: '-1.017,60') para float.
    Separador de milhar: ponto. Separador decimal: vírgula.
    """
    cleaned = value.strip().replace(".", "").replace(",", ".")
    return float(cleaned)


def _is_negative_commission(comissao_cell) -> bool:
    if comissao_cell is None:
        return False
    raw = str(comissao_cell).strip()
    if not raw or raw == "None":
        return False
    try:
        return _parse_br_float(raw) < 0
    except ValueError:
        return False


def _extract_record_from_row(row: list) -> PDFRecord | None:
    """
    Valida e extrai PDFRecord de uma linha da tabela.
    Retorna None se a linha não for uma linha de dados válida.
    """
    if len(row) < _MIN_COLS:
        return None
    if str(row[0]).strip() not in _VALID_PE_VALUES:
        return None

    segurado = row[_COL_SEGURADO]
    inicio_vig = row[_COL_INICIO_VIG]

    if not segurado or not inicio_vig:
        return None

    segurado_str = str(segurado).strip()
    inicio_vig_str = str(inicio_vig).strip()

    if not segurado_str or not inicio_vig_str:
        return None

    return PDFRecord(segurado=segurado_str, inicio_vig=inicio_vig_str)


def extract_negative_commission_records(pdf_path: str) -> List[PDFRecord]:
    """
    Abre o PDF e retorna todos os PDFRecord onde COMISSÃO < 0.
    """
    records: List[PDFRecord] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                # Cada tabela pode ter 1 ou mais linhas; a linha de dados é sempre a primeira
                data_row = table[0]
                record = _extract_record_from_row(data_row)
                if record is None:
                    continue
                if _is_negative_commission(data_row[_COL_COMISSAO]):
                    records.append(record)

    return records
