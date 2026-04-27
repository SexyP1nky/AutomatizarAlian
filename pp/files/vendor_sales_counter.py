"""
Responsabilidade: contar vendas positivas por vendedor por mês.

Fluxo:
  1. Recebe registros positivos do PDF e registros do XLSX
  2. Para cada registro positivo do PDF, encontra o vendedor no XLSX
  3. Extrai o mês/ano do campo inicio_vig (DD/MM/YYYY → MM/YYYY)
  4. Agrupa e conta por (vendedor_upper, mês_ano)
  5. Retorna dicionário com as contagens

A contagem é usada para decidir o valor de estorno:
  - 4+ vendas positivas em um mês → R$ 50
  - 3 ou menos → R$ 30
"""

from collections import Counter
from typing import Dict, List, Tuple

from matcher import _build_xlsx_index, _find_matches_for_segurado
from name_normalizer import normalize_name
from pdf_parser import PDFRecord
from xlsx_parser import XLSXRecord


def extract_month_year(date_str: str) -> str:
    """
    Extrai mês/ano de uma data no formato DD/MM/YYYY.

    Retorna string "MM/YYYY" (ex: "04/2026").
    Retorna string vazia se o formato for inválido.
    """
    try:
        parts = date_str.strip().split("/")
        if len(parts) == 3:
            month = parts[1]
            year = parts[2]
            return f"{month}/{year}"
    except (ValueError, IndexError, AttributeError):
        pass
    return ""


def _match_positive_record_to_vendor(
    pdf_rec: PDFRecord,
    index: Dict[str, List[XLSXRecord]],
    all_xlsx_records: List[XLSXRecord],
) -> str:
    """
    Encontra o vendedor correspondente a um registro positivo do PDF.

    Retorna o nome do vendedor em maiúsculas, ou string vazia se não encontrado.
    Usa a mesma lógica de correspondência do matcher.
    """
    normalized = normalize_name(pdf_rec.segurado)
    matches = _find_matches_for_segurado(pdf_rec, normalized, index, all_xlsx_records)

    if not matches:
        return ""

    # Usa o primeiro match (prioridade: exato > fuzzy)
    xlsx_rec, _ = matches[0]
    return xlsx_rec.vendedor.upper().strip()


def count_positive_sales_per_vendor_month(
    positive_records: List[PDFRecord],
    xlsx_records: List[XLSXRecord],
) -> Dict[Tuple[str, str], int]:
    """
    Conta vendas positivas por (vendedor, mês).

    Para cada registro positivo do PDF:
      - Cruza com o XLSX para descobrir o vendedor
      - Extrai o mês do inicio_vig
      - Incrementa o contador de (vendedor_upper, mês_ano)

    Retorna:
        dicionário {(vendedor_upper, "MM/YYYY"): contagem}
    """
    index = _build_xlsx_index(xlsx_records)

    counts: Counter = Counter()

    for pdf_rec in positive_records:
        vendor = _match_positive_record_to_vendor(pdf_rec, index, xlsx_records)
        if not vendor:
            continue

        month_year = extract_month_year(pdf_rec.inicio_vig)
        if not month_year:
            continue

        counts[(vendor, month_year)] += 1

    return dict(counts)


def get_estorno_value(
    vendor_sales_counts: Dict[Tuple[str, str], int],
    vendedor: str,
    inicio_vig: str,
) -> int:
    """
    Determina o valor de estorno (30 ou 50) para um vendedor em um mês.

    Parâmetros:
        vendor_sales_counts — contagens de vendas positivas por (vendedor, mês)
        vendedor            — nome do vendedor (será normalizado para upper/strip)
        inicio_vig          — data de início de vigência (DD/MM/YYYY)

    Retorna:
        50 se o vendedor tem 4+ vendas positivas naquele mês, 30 caso contrário
    """
    vendor_key = vendedor.upper().strip()
    month_year = extract_month_year(inicio_vig)

    if not month_year:
        return 30

    count = vendor_sales_counts.get((vendor_key, month_year), 0)
    return 50 if count >= 4 else 30
