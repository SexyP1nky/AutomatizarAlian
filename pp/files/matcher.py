"""
Responsabilidade: correlacionar segurados do PDF com clientes do XLSX.

Para cada segurado com comissão negativa:
- Busca correspondência combinando apólice (100% igual se existir) e nome
- 3 níveis de correspondência:
    - EXATO: nome bate (exato ou ç≈c) E apólice bate (ou apólice não existe)
    - APOLICE_DIFERENTE: nome bate (exato ou ç≈c) MAS apólice é diferente
    - FUZZY: nome não bate exatamente, mas é > 85% similar.
- Se encontrar múltiplos vendedores (client em abas diferentes), gera uma linha por vendedor
- Se não encontrar, adiciona ao relatório de não encontrados
"""

from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, List, Tuple

from thefuzz import fuzz

from name_normalizer import names_match, normalize_name
from pdf_parser import PDFRecord
from xlsx_parser import XLSXRecord


@dataclass(frozen=True)
class MatchedRecord:
    segurado: str       # nome original do PDF
    inicio_vig: str     # DD/MM/YYYY
    vendedor: str       # nome original do XLSX
    match_type: str     # "EXATO", "APOLICE_DIFERENTE", "FUZZY", etc.
    apolice_pdf: str    # apólice que estava no PDF
    apolice_xlsx: str   # apólice que estava no XLSX
    sheet_name: str     # aba da planilha XLSX onde foi encontrado


def _build_xlsx_index(
    xlsx_records: List[XLSXRecord],
) -> Dict[str, List[XLSXRecord]]:
    """
    Constrói índice: nome_normalizado → lista de XLSXRecord.
    Para buscas exatas.
    """
    index: Dict[str, List[XLSXRecord]] = defaultdict(list)
    for rec in xlsx_records:
        key = normalize_name(rec.cliente)
        index[key].append(rec)
    return index


def _compare_apolice(apol_pdf: str, apol_xlsx: str) -> bool:
    if not apol_pdf or not apol_xlsx:
        return False
    a1 = apol_pdf[-5:] if len(apol_pdf) >= 5 else apol_pdf
    a2 = apol_xlsx[-5:] if len(apol_xlsx) >= 5 else apol_xlsx
    return a1 == a2


def _eval_match_type(
    pdf_name: str,
    pdf_apolice: str,
    xlsx_name: str,
    xlsx_apolice: str,
    is_exact_name: bool
) -> str:
    """
    Avalia o tipo de match baseado no nome e na apólice (últimos 5 dígitos).
    """
    if pdf_apolice and not xlsx_apolice:
        if is_exact_name:
            return "APOLICE_AUSENTE_XLSX"
        else:
            return "FUZZY_APOLICE_AUSENTE_XLSX"

    if pdf_apolice and xlsx_apolice and not _compare_apolice(pdf_apolice, xlsx_apolice):
        # Se as duas apólices existem e são diferentes nos últimos 5 dígitos
        if is_exact_name:
            return "APOLICE_DIFERENTE"
        else:
            return "FUZZY_APOLICE_DIFERENTE"
    else:
        # Apólices batem, ou o PDF não tem apólice
        if is_exact_name:
            return "EXATO"
        else:
            return "FUZZY"


def _find_matches_for_segurado(
    pdf_rec: PDFRecord,
    normalized_segurado: str,
    index: Dict[str, List[XLSXRecord]],
    all_xlsx_records: List[XLSXRecord]
) -> List[Tuple[XLSXRecord, str]]:
    """
    Tenta correspondência exata primeiro; depois, aproximação Ç ≈ C por token.
    Se falhar, tenta Fuzzy matching.
    Retorna lista de (XLSXRecord, match_type).
    """
    matches = []

    # 1. Correspondência exata no índice
    exact = index.get(normalized_segurado)
    if exact:
        if len(exact) > 1 and pdf_rec.apolice:
            with_apolice = [r for r in exact if _compare_apolice(pdf_rec.apolice, r.apolice)]
            if with_apolice:
                exact = with_apolice
        for rec in exact:
            m_type = _eval_match_type(normalized_segurado, pdf_rec.apolice, normalize_name(rec.cliente), rec.apolice, True)
            matches.append((rec, m_type))
        return matches

    # 2. Fallback 1: percorre o índice e testa token a token com regra do Ç
    for normalized_client, records in index.items():
        if names_match(normalized_segurado, normalized_client):
            if len(records) > 1 and pdf_rec.apolice:
                with_apolice = [r for r in records if _compare_apolice(pdf_rec.apolice, r.apolice)]
                if with_apolice:
                    records = with_apolice
            for rec in records:
                m_type = _eval_match_type(normalized_segurado, pdf_rec.apolice, normalized_client, rec.apolice, True)
                matches.append((rec, m_type))
            return matches

    # 3. Fallback 2: Fuzzy matching
    # Apenas se o nome for bem parecido (> 85)
    best_score = 0
    best_records = []
    
    for rec in all_xlsx_records:
        norm_client = normalize_name(rec.cliente)
        score = fuzz.token_sort_ratio(normalized_segurado, norm_client)
        if score > 85:
            if score > best_score:
                best_score = score
                best_records = [rec]
            elif score == best_score:
                best_records.append(rec)
                
    if best_records:
        if len(best_records) > 1 and pdf_rec.apolice:
            with_apolice = [r for r in best_records if _compare_apolice(pdf_rec.apolice, r.apolice)]
            if with_apolice:
                best_records = with_apolice

        for rec in best_records:
            m_type = _eval_match_type(normalized_segurado, pdf_rec.apolice, normalize_name(rec.cliente), rec.apolice, False)
            matches.append((rec, m_type))

    return matches


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
        xlsx_matches = _find_matches_for_segurado(pdf_rec, normalized, index, xlsx_records)

        if not xlsx_matches:
            not_found.append(pdf_rec.segurado)
            continue

        for xlsx_rec, match_type in xlsx_matches:
            matched.append(
                MatchedRecord(
                    segurado=pdf_rec.segurado,
                    inicio_vig=pdf_rec.inicio_vig,
                    vendedor=xlsx_rec.vendedor,
                    match_type=match_type,
                    apolice_pdf=pdf_rec.apolice,
                    apolice_xlsx=xlsx_rec.apolice,
                    sheet_name=xlsx_rec.sheet_name
                )
            )

    return matched, not_found
