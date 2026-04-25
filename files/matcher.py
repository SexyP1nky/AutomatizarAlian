"""
Responsabilidade: correlacionar segurados do PDF com clientes do XLSX.

Para cada segurado com comissão negativa:
- Busca correspondência exata (após normalização) no XLSX
- Fallback: correspondência token a token com Ç ≈ C
- Se encontrar múltiplos vendedores (client em abas diferentes), gera uma linha por vendedor
- Se não encontrar, adiciona ao relatório de não encontrados
"""

from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, List, Tuple

from name_normalizer import names_match, normalize_name
from pdf_parser import PDFRecord
from xlsx_parser import XLSXRecord


@dataclass(frozen=True)
class MatchedRecord:
    segurado: str       # nome original do PDF
    inicio_vig: str     # DD/MM/YYYY
    vendedor: str       # nome original do XLSX


def _build_xlsx_index(
    xlsx_records: List[XLSXRecord],
) -> Dict[str, List[XLSXRecord]]:
    """
    Constrói índice: nome_normalizado → lista de XLSXRecord.
    """
    index: Dict[str, List[XLSXRecord]] = defaultdict(list)
    for rec in xlsx_records:
        key = normalize_name(rec.cliente)
        index[key].append(rec)
    return index


def _find_matches_for_segurado(
    normalized_segurado: str,
    index: Dict[str, List[XLSXRecord]],
) -> List[XLSXRecord]:
    """
    Tenta correspondência exata primeiro; depois, aproximação Ç ≈ C por token.
    """
    # 1. Correspondência exata
    exact = index.get(normalized_segurado)
    if exact:
        return exact

    # 2. Fallback: percorre o índice e testa token a token com regra do Ç
    for normalized_client, records in index.items():
        if names_match(normalized_segurado, normalized_client):
            return records

    return []


def match_records(
    pdf_records: List[PDFRecord],
    xlsx_records: List[XLSXRecord],
) -> Tuple[List[MatchedRecord], List[str]]:
    """
    Cruza os registros do PDF com os do XLSX.

    Retorna:
        matched   — lista de MatchedRecord (uma entrada por combinação segurado + vendedor)
        not_found — nomes de segurados sem nenhuma correspondência no XLSX
    """
    index = _build_xlsx_index(xlsx_records)

    matched: List[MatchedRecord] = []
    not_found: List[str] = []

    for pdf_rec in pdf_records:
        normalized = normalize_name(pdf_rec.segurado)
        xlsx_matches = _find_matches_for_segurado(normalized, index)

        if not xlsx_matches:
            not_found.append(pdf_rec.segurado)
            continue

        for xlsx_rec in xlsx_matches:
            matched.append(
                MatchedRecord(
                    segurado=pdf_rec.segurado,
                    inicio_vig=pdf_rec.inicio_vig,
                    vendedor=xlsx_rec.vendedor,
                )
            )

    return matched, not_found
