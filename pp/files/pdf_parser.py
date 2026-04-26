"""
Responsabilidade: extrair do PDF os registros onde COMISSÃO é negativa.

Retorna lista de PDFRecord com (segurado, inicio_vig).
O valor da comissão é descartado — interessa apenas identificar o segurado.

Estratégia de extração (dupla, para máxima cobertura):
  1. Tenta via extract_tables() do pdfplumber (funciona bem para PDFs de teste
     com tabela única e para a maioria dos registros do PDF real)
  2. Fallback via extract_text() com regex para capturar linhas de dados que
     o extrator de tabelas não detectou (cobre registros "perdidos" no PDF real)

Posições fixas na tabela:
    col 0 → P/E ('P' ou 'E')
    col 3 → INICIO VIG
    col 4 → SEGURADO
    col 12 → COMISSÃO
"""

import re
from dataclasses import dataclass
from typing import List, Set

import pdfplumber


# Posições de coluna confirmadas via inspeção do PDF real
_COL_INICIO_VIG = 3
_COL_SEGURADO = 4
_COL_COMISSAO = 12
_MIN_COLS = 13
_VALID_PE_VALUES = {"P", "E"}

# Regex para extrair linhas de dados do texto bruto do PDF
# Formato: P/E TIPO_SEGURO PROPOSTA DD/MM/YYYY NOME APOLICE ... COMISSAO
_LINE_RE = re.compile(
    r"^[PE]\s+"                     # P ou E
    r"(?:SEGURO NOV|RENOV CORR)\s+"  # tipo de seguro
    r"\d+\s+"                        # proposta
    r"(\d{2}/\d{2}/\d{4})\s+"       # inicio_vig (grupo 1)
    r"([A-ZÀ-ÿ\s]+?)\s+"           # segurado (grupo 2) — captura gulosa mínima de letras+espaços
    r"[\d]+\S*\s+"                   # apólice (número + possíveis letras)
    r"\S+\s+"                        # chassi
    r"\d+/\d+\s+"                    # PC
    r"[\d.,]+\s+"                    # prêmio líquido
    r"[\d.,]+\s+"                    # % rep
    r"(-?[\d.,]+)\s*$"              # comissão (grupo 3)
)


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


def _extract_from_tables(page) -> List[PDFRecord]:
    """
    Extrai registros com comissão negativa usando extract_tables().
    """
    records: List[PDFRecord] = []
    tables = page.extract_tables()
    for table in tables:
        if not table:
            continue
        # Percorre TODAS as linhas de cada tabela para capturar
        # tanto PDFs com uma tabela grande quanto PDFs com mini-tabelas
        for row in table:
            record = _extract_record_from_row(row)
            if record is None:
                continue
            if _is_negative_commission(row[_COL_COMISSAO]):
                records.append(record)
    return records


def _extract_from_text(page) -> List[PDFRecord]:
    """
    Fallback: extrai registros com comissão negativa via texto bruto + regex.
    Captura registros que extract_tables() perde.
    """
    records: List[PDFRecord] = []
    text = page.extract_text()
    if not text:
        return records

    for line in text.split("\n"):
        line = line.strip()
        match = _LINE_RE.match(line)
        if not match:
            continue
        inicio_vig = match.group(1)
        segurado = match.group(2).strip()
        comissao_str = match.group(3)

        if _is_negative_commission(comissao_str):
            records.append(PDFRecord(segurado=segurado, inicio_vig=inicio_vig))

    return records


def extract_negative_commission_records(pdf_path: str) -> List[PDFRecord]:
    """
    Abre o PDF e retorna todos os PDFRecord onde COMISSÃO < 0.

    Usa extração dupla (tabelas + texto) e mescla resultados sem duplicatas.
    A chave de deduplicação é (segurado_upper, inicio_vig).
    """
    seen: Set[tuple] = set()
    records: List[PDFRecord] = []

    def _add_if_new(rec: PDFRecord) -> None:
        key = (rec.segurado.upper(), rec.inicio_vig)
        if key not in seen:
            seen.add(key)
            records.append(rec)

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Estratégia 1: tabelas
            for rec in _extract_from_tables(page):
                _add_if_new(rec)

            # Estratégia 2: texto (fallback para registros perdidos)
            for rec in _extract_from_text(page):
                _add_if_new(rec)

    return records
