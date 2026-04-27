"""
Responsabilidade: contar vendas por vendedor por mês usando o XLSX.

Fluxo:
  1. Recebe todos os registros do XLSX (já com sheet_name)
  2. Extrai o mês de cada aba (ex: "JAN 26", "FEV 26 ELIDA" → "JAN", "FEV")
  3. Agrupa e conta por (vendedor_upper, mês_normalizado)
  4. Retorna dicionário com as contagens

NOTA: Múltiplas abas podem representar o mesmo mês
  (ex: "FEV 26 ELIDA" e "FEV 26 SUELANE" → ambas "FEV/26").
  As contagens são somadas.

A contagem é usada para decidir o valor de estorno:
  - 4+ vendas em um mês → R$ 50
  - 3 ou menos → R$ 30
"""

import re
from collections import Counter
from typing import Dict, List, Tuple

from xlsx_parser import XLSXRecord


# Mapeamento de nomes de meses para número (para normalização)
_MONTH_NAMES = {
    "JAN": "01", "FEV": "02", "MAR": "03", "ABR": "04",
    "MAI": "05", "JUN": "06", "JUL": "07", "AGO": "08",
    "SET": "09", "OUT": "10", "NOV": "11", "DEZ": "12",
}


def extract_month_from_sheet_name(sheet_name: str) -> str:
    """
    Extrai o mês/ano normalizado do nome de uma aba do XLSX.

    Exemplos:
        "JAN 26"          → "01/26"
        "FEV 26 ELIDA"    → "02/26"
        "SETEMBRO 25"     → "09/25"
        "JUNHO_25"        → "06/25"
        "MARÇO 2025"      → "03/25"

    Retorna string "MM/YY" ou string vazia se não conseguir extrair.
    """
    upper = sheet_name.strip().upper()

    for name, num in _MONTH_NAMES.items():
        if upper.startswith(name):
            # Extrair o ano (primeira sequência de 2 a 4 dígitos que encontrar na string)
            year_match = re.search(r"(\d{2,4})", upper)
            if year_match:
                year = year_match.group(1)
                # Se ano com 4 dígitos, pegar últimos 2
                if len(year) == 4:
                    year = year[2:]
                return f"{num}/{year}"
            break

    return ""


def extract_month_from_date(date_str: str) -> str:
    """
    Extrai mês/ano de uma data DD/MM/YYYY e converte para formato "MM/YY".

    Exemplos:
        "12/01/2026" → "01/26"
        "30/03/2026" → "03/26"

    Retorna string vazia se o formato for inválido.
    """
    try:
        parts = date_str.strip().split("/")
        if len(parts) == 3:
            month = parts[1]
            year = parts[2][-2:]  # Últimos 2 dígitos
            return f"{month}/{year}"
    except (ValueError, IndexError, AttributeError):
        pass
    return ""


def count_sales_per_vendor_month(
    xlsx_records: List[XLSXRecord],
) -> Dict[Tuple[str, str], int]:
    """
    Conta vendas por (vendedor, mês) usando apenas os dados do XLSX.

    O mês é extraído do nome da aba (sheet_name).
    Múltiplas abas do mesmo mês são somadas.

    Retorna:
        dicionário {(vendedor_upper, "MM/YY"): contagem}
    """
    counts: Counter = Counter()

    for rec in xlsx_records:
        month = extract_month_from_sheet_name(rec.sheet_name)
        if not month:
            continue

        vendor_key = rec.vendedor.upper().strip()
        counts[(vendor_key, month)] += 1

    return dict(counts)


def get_estorno_value(
    vendor_sales_counts: Dict[Tuple[str, str], int],
    vendedor: str,
    sheet_name: str,
) -> int:
    """
    Determina o valor de estorno (30 ou 50) para um vendedor em um mês.

    O mês é extraído do nome da aba onde o vendedor foi encontrado.

    Parâmetros:
        vendor_sales_counts — contagens de vendas por (vendedor, mês) do XLSX
        vendedor            — nome do vendedor (será normalizado para upper/strip)
        sheet_name          — nome da aba do XLSX onde o vendedor foi encontrado

    Retorna:
        50 se o vendedor tem 4+ vendas naquele mês, 30 caso contrário
    """
    vendor_key = vendedor.upper().strip()
    month = extract_month_from_sheet_name(sheet_name)

    if not month:
        return 30

    count = vendor_sales_counts.get((vendor_key, month), 0)
    return 50 if count >= 4 else 30
