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


def _filter_best_matches(
    matches: List[Tuple[XLSXRecord, str]],
) -> List[Tuple[XLSXRecord, str]]:
    """
    Se houver múltiplos matches para o mesmo cliente, tenta desempatar pela apólice.
    Se algum deles for um match EXATO de apólice, retorna apenas ele(s).
    Caso contrário, retorna todos (ambiguidade).
    """
    if len(matches) <= 1:
        return matches

    # Verifica se existe algum match com apólice exata E nome exato
    exact_matches = [m for m in matches if m[1] == "EXATO"]
    if exact_matches:
        # Se a apólice bate perfeitamente E o nome bate, e existem múltiplos na planilha, 
        # isso indica linhas duplicadas pela digitação. Retorna só o primeiro 
        # para evitar cobrar o estorno do vendedor duas vezes.
        return [exact_matches[0]]
        
    # Verifica se existe algum match FUZZY (nome aproximado, mas apólice exata)
    fuzzy_exact_apolice_matches = [m for m in matches if m[1] == "FUZZY"]
    if fuzzy_exact_apolice_matches:
        # Pega só o primeiro para evitar dupla cobrança por linhas duplicadas na planilha
        first_fuzzy = fuzzy_exact_apolice_matches[0]
        out = [first_fuzzy]
        # Repassa junto todas as outras ambiguidades de apólice (se houverem)
        out.extend([m for m in matches if m[1] not in ("EXATO", "FUZZY")])
        return out

    return matches


def _find_matches_for_segurado(
    pdf_rec: PDFRecord,
    normalized_segurado: str,
    index: Dict[str, List[XLSXRecord]],
    all_xlsx_records: List[XLSXRecord]
) -> List[Tuple[XLSXRecord, str]]:
    """
    Tenta correspondência exata primeiro; depois, aproximação Ç ≈ C por token.
    Se falhar, tenta Fuzzy matching.
    Filtra para retornar apenas o melhor match (desempate por apólice).
    """
    matches = []

    # 1. Correspondência exata no índice
    exact = index.get(normalized_segurado)
    if exact:
        for rec in exact:
            m_type = _eval_match_type(normalized_segurado, pdf_rec.apolice, normalize_name(rec.cliente), rec.apolice, True)
            matches.append((rec, m_type))
        return _filter_best_matches(matches)

    # 2. Fallback 1: percorre o índice e testa token a token com regra do Ç
    for normalized_client, records in index.items():
        if names_match(normalized_segurado, normalized_client):
            for rec in records:
                m_type = _eval_match_type(normalized_segurado, pdf_rec.apolice, normalized_client, rec.apolice, True)
                matches.append((rec, m_type))
            return _filter_best_matches(matches)

    # 3. Fallback 2: Fuzzy matching
    best_score_apolice_match = 0
    best_records_apolice_match = []
    
    best_score_no_apolice = 0
    best_records_no_apolice = []
    
    for rec in all_xlsx_records:
        norm_client = normalize_name(rec.cliente)
        score = fuzz.token_sort_ratio(normalized_segurado, norm_client)
        
        if score > 85:
            apolice_matches = _compare_apolice(pdf_rec.apolice, rec.apolice)
            if apolice_matches:
                # O nome é fuzzy, mas a apólice bate perfeitamente
                if score > best_score_apolice_match:
                    best_score_apolice_match = score
                    best_records_apolice_match = [rec]
                elif score == best_score_apolice_match:
                    best_records_apolice_match.append(rec)
            else:
                # A apólice NÃO bate. Só consideramos nomes incrivelmente próximos (>92)
                if score > 92:
                    if score > best_score_no_apolice:
                        best_score_no_apolice = score
                        best_records_no_apolice = [rec]
                    elif score == best_score_no_apolice:
                        best_records_no_apolice.append(rec)
                
    if best_records_apolice_match:
        for rec in best_records_apolice_match:
            m_type = _eval_match_type(normalized_segurado, pdf_rec.apolice, normalize_name(rec.cliente), rec.apolice, False)
            matches.append((rec, m_type))
            
    if best_records_no_apolice:
        for rec in best_records_no_apolice:
            m_type = _eval_match_type(normalized_segurado, pdf_rec.apolice, normalize_name(rec.cliente), rec.apolice, False)
            matches.append((rec, m_type))
            
    if matches:
        return _filter_best_matches(matches)

    return matches


def match_records(
    pdf_records: List[PDFRecord],
    xlsx_records: List[XLSXRecord],
) -> Tuple[List[MatchedRecord], List[PDFRecord]]:
    """
    Cruza os registros do PDF com os do XLSX.

    Retorna:
        matched   — lista de MatchedRecord (uma entrada por combinação segurado + vendedor)
        not_found — registros do PDF sem nenhuma correspondência no XLSX
    """
    index = _build_xlsx_index(xlsx_records)

    matched: List[MatchedRecord] = []
    not_found: List[PDFRecord] = []

    for pdf_rec in pdf_records:
        normalized = normalize_name(pdf_rec.segurado)
        xlsx_matches = _find_matches_for_segurado(pdf_rec, normalized, index, xlsx_records)

        if not xlsx_matches:
            not_found.append(pdf_rec)
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
