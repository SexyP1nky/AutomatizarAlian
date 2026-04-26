"""
Responsabilidade: extrair do PDF os registros onde COMISSÃO é negativa.

Retorna lista de PDFRecord com (segurado, inicio_vig).
O valor da comissão é descartado — interessa apenas identificar o segurado.

Estratégia de extração (dupla, para máxima cobertura):
  1. Tenta via extract_tables() do pdfplumber (funciona bem para PDFs de teste
     com tabela única e para a maioria dos registros do PDF real)
     - Detecção dinâmica de colunas via cabeçalho da tabela
     - Fallback para posições fixas se o cabeçalho não for encontrado
  2. Fallback via extract_text() com regex para capturar linhas de dados que
     o extrator de tabelas não detectou (cobre registros "perdidos" no PDF real)

Posições fixas de fallback na tabela:
    col 0 → P/E ('P' ou 'E')
    col 3 → INICIO VIG
    col 4 → SEGURADO
    col 12 → COMISSÃO
"""

import re
from dataclasses import dataclass
from typing import List, Set, Tuple

import pdfplumber


# Posições de coluna de fallback confirmadas via inspeção do PDF real
_DEFAULT_COL_INICIO_VIG = 3
_DEFAULT_COL_SEGURADO = 4
_DEFAULT_COL_COMISSAO = 12
_MIN_COLS = 13
_VALID_PE_VALUES = {"P", "E"}

# Palavras-chave para detecção dinâmica de colunas no cabeçalho da tabela
_HEADER_KEYWORDS = {
    "segurado": ["SEGURADO", "NOME", "CLIENTE"],
    "inicio_vig": ["INICIO VIG", "INÍCIO VIG", "INICIO", "INÍCIO", "VIGENCIA", "VIGÊNCIA"],
    "comissao": ["COMISSÃO", "COMISSAO", "COMISS"],
}

# Regex para extrair linhas de dados do texto bruto do PDF
# Formato: P/E TIPO_SEGURO PROPOSTA DD/MM/YYYY NOME APOLICE ... COMISSAO
# Expandido para capturar mais tipos de operação além de "SEGURO NOV" e "RENOV CORR"
_LINE_RE = re.compile(
    r"^[PE]\s+"                      # P ou E
    r"(?:"
    r"SEGURO NOV|RENOV CORR|"        # tipos originais
    r"ENDOSSO|CANCELAMENTO|"         # endosso e cancelamento
    r"SEGURO\s+\w+|"                 # qualquer tipo de seguro (SEGURO XXX)
    r"RENOV\s+\w+|"                  # qualquer tipo de renovação (RENOV XXX)
    r"[A-Z]{2,}\s+[A-Z]{2,}"        # padrão genérico (duas palavras maiúsculas)
    r")\s+"
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


def _detect_columns_from_header(
    table: list,
) -> Tuple[int | None, int | None, int | None]:
    """
    Tenta detectar as colunas SEGURADO, INICIO VIG e COMISSÃO
    a partir da primeira linha (cabeçalho) de uma tabela.

    Retorna (col_segurado, col_inicio_vig, col_comissao) ou (None, None, None).
    """
    if not table or not table[0]:
        return None, None, None

    header = [str(cell).upper().strip() if cell else "" for cell in table[0]]

    col_segurado = None
    col_inicio_vig = None
    col_comissao = None

    for col_idx, cell_text in enumerate(header):
        if not cell_text:
            continue

        # Detectar coluna SEGURADO
        if col_segurado is None:
            for keyword in _HEADER_KEYWORDS["segurado"]:
                if keyword in cell_text:
                    col_segurado = col_idx
                    break

        # Detectar coluna INICIO VIG
        if col_inicio_vig is None:
            for keyword in _HEADER_KEYWORDS["inicio_vig"]:
                if keyword in cell_text:
                    col_inicio_vig = col_idx
                    break

        # Detectar coluna COMISSÃO
        if col_comissao is None:
            for keyword in _HEADER_KEYWORDS["comissao"]:
                if keyword in cell_text:
                    col_comissao = col_idx
                    break

    # Só retorna se encontrou TODAS as colunas necessárias
    if col_segurado is not None and col_inicio_vig is not None and col_comissao is not None:
        return col_segurado, col_inicio_vig, col_comissao

    return None, None, None


def _extract_record_from_row(
    row: list,
    col_segurado: int,
    col_inicio_vig: int,
    col_comissao: int,
    min_cols: int,
) -> PDFRecord | None:
    """
    Valida e extrai PDFRecord de uma linha da tabela.
    Retorna None se a linha não for uma linha de dados válida.
    """
    if len(row) < min_cols:
        return None
    if str(row[0]).strip() not in _VALID_PE_VALUES:
        return None

    segurado = row[col_segurado]
    inicio_vig = row[col_inicio_vig]

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
    Tenta detecção dinâmica de colunas; usa fallback para posições fixas.
    """
    records: List[PDFRecord] = []
    tables = page.extract_tables()
    for table in tables:
        if not table:
            continue

        # Tenta detectar colunas dinamicamente pelo cabeçalho
        detected = _detect_columns_from_header(table)
        if detected != (None, None, None):
            col_segurado, col_inicio_vig, col_comissao = detected
            min_cols = max(col_segurado, col_inicio_vig, col_comissao) + 1
        else:
            # Fallback: posições fixas
            col_segurado = _DEFAULT_COL_SEGURADO
            col_inicio_vig = _DEFAULT_COL_INICIO_VIG
            col_comissao = _DEFAULT_COL_COMISSAO
            min_cols = _MIN_COLS

        # Percorre TODAS as linhas de cada tabela para capturar
        # tanto PDFs com uma tabela grande quanto PDFs com mini-tabelas
        for row in table:
            record = _extract_record_from_row(
                row, col_segurado, col_inicio_vig, col_comissao, min_cols
            )
            if record is None:
                continue
            if _is_negative_commission(row[col_comissao]):
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
