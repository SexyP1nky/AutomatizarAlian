"""
Responsabilidade: extrair do PDF os registros onde COMISSÃO é negativa.

Retorna lista de PDFRecord com (segurado, inicio_vig, apolice).
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
    col 5 → APÓLICE (pode estar na mesma row ou na row seguinte)
    col 12 → COMISSÃO
"""

import re
from dataclasses import dataclass
from typing import List, Set, Tuple

import pdfplumber


# Posições de coluna de fallback confirmadas via inspeção do PDF real
_DEFAULT_COL_INICIO_VIG = 3
_DEFAULT_COL_SEGURADO = 4
_DEFAULT_COL_APOLICE = 5
_DEFAULT_COL_COMISSAO = 12
_MIN_COLS = 13
_VALID_PE_VALUES = {"P", "E"}

# Palavras-chave para detecção dinâmica de colunas no cabeçalho da tabela
_HEADER_KEYWORDS = {
    "segurado": ["SEGURADO", "NOME", "CLIENTE"],
    "inicio_vig": ["INICIO VIG", "INÍCIO VIG", "INICIO", "INÍCIO", "VIGENCIA", "VIGÊNCIA"],
    "comissao": ["COMISSÃO", "COMISSAO", "COMISS"],
    "apolice": ["APÓLICE", "APOLICE", "Nº APÓLICE", "N° APÓLICE", "NÚMERO DA APÓLICE",
                 "NUMERO DA APOLICE", "N° DA APÓLICE", "Nº DA APÓLICE"],
}

# Regex para extrair linhas de dados do texto bruto do PDF
# Formato: P/E TIPO_SEGURO PROPOSTA DD/MM/YYYY NOME APOLICE CHASSI PC PREMIO %REP COMISSAO
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
    r"([A-ZÀ-ÿ\s]+?)\s+"           # segurado (grupo 2)
    r"(\d[\d\S]*)\s+"               # apólice (grupo 3) — número que inicia com dígito
    r"\S+\s+"                        # chassi
    r"\d+/\d+\s+"                    # PC
    r"[\d.,]+\s+"                    # prêmio líquido
    r"[\d.,]+\s+"                    # % rep
    r"(-?[\d.,]+)\s*$"              # comissão (grupo 4)
)


@dataclass(frozen=True)
class PDFRecord:
    segurado: str
    inicio_vig: str  # formato original do PDF: DD/MM/YYYY
    apolice: str     # número da apólice


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


def _is_positive_commission(comissao_cell) -> bool:
    if comissao_cell is None:
        return False
    raw = str(comissao_cell).strip()
    if not raw or raw == "None":
        return False
    try:
        return _parse_br_float(raw) > 0
    except ValueError:
        return False


def _normalize_apolice(raw: str) -> str:
    """
    Normaliza número de apólice: remove espaços, pontos, hífens.
    Retorna string limpa só com dígitos e letras.
    """
    if not raw:
        return ""
    return re.sub(r"[\s.\-/]", "", str(raw).strip())


def _detect_columns_from_header(
    table: list,
) -> Tuple[int | None, int | None, int | None, int | None]:
    """
    Tenta detectar as colunas SEGURADO, INICIO VIG, APÓLICE e COMISSÃO
    a partir da primeira linha (cabeçalho) de uma tabela.

    Retorna (col_segurado, col_inicio_vig, col_apolice, col_comissao) ou Nones.
    """
    if not table or not table[0]:
        return None, None, None, None

    header = [str(cell).upper().strip() if cell else "" for cell in table[0]]

    col_segurado = None
    col_inicio_vig = None
    col_comissao = None
    col_apolice = None

    for col_idx, cell_text in enumerate(header):
        if not cell_text:
            continue

        if col_segurado is None:
            for keyword in _HEADER_KEYWORDS["segurado"]:
                if keyword in cell_text:
                    col_segurado = col_idx
                    break

        if col_inicio_vig is None:
            for keyword in _HEADER_KEYWORDS["inicio_vig"]:
                if keyword in cell_text:
                    col_inicio_vig = col_idx
                    break

        if col_apolice is None:
            for keyword in _HEADER_KEYWORDS["apolice"]:
                if keyword in cell_text:
                    col_apolice = col_idx
                    break

        if col_comissao is None:
            for keyword in _HEADER_KEYWORDS["comissao"]:
                if keyword in cell_text:
                    col_comissao = col_idx
                    break

    if (col_segurado is not None and col_inicio_vig is not None
            and col_comissao is not None):
        # apolice is optional for detection success
        return col_segurado, col_inicio_vig, col_apolice, col_comissao

    return None, None, None, None


def _extract_apolice_from_table(
    table: list, data_row_idx: int, col_apolice: int
) -> str:
    """
    Extrai o nº da apólice de uma tabela. No PDF real, a apólice pode estar:
    - Na mesma row que os dados (row 0, col 5)
    - Na row seguinte (row 1, col 5) quando a row 0 tem a célula vazia

    Tenta ambas as posições.
    """
    # Tenta na row de dados
    row = table[data_row_idx]
    if col_apolice < len(row):
        val = _normalize_apolice(str(row[col_apolice]) if row[col_apolice] else "")
        if val:
            return val

    # Tenta na row seguinte (padrão do PDF real: apólice em sub-row)
    next_idx = data_row_idx + 1
    if next_idx < len(table):
        next_row = table[next_idx]
        if col_apolice < len(next_row):
            val = _normalize_apolice(
                str(next_row[col_apolice]) if next_row[col_apolice] else ""
            )
            if val:
                return val

    return ""


def _extract_record_from_row(
    row: list,
    col_segurado: int,
    col_inicio_vig: int,
    col_comissao: int,
    min_cols: int,
) -> Tuple[str | None, str | None]:
    """
    Valida e extrai (segurado, inicio_vig) de uma linha da tabela.
    Retorna (None, None) se a linha não for uma linha de dados válida.
    """
    if len(row) < min_cols:
        return None, None
    if str(row[0]).strip() not in _VALID_PE_VALUES:
        return None, None

    segurado = row[col_segurado]
    inicio_vig = row[col_inicio_vig]

    if not segurado or not inicio_vig:
        return None, None

    segurado_str = str(segurado).strip()
    inicio_vig_str = str(inicio_vig).strip()

    if not segurado_str or not inicio_vig_str:
        return None, None

    return segurado_str, inicio_vig_str


def _extract_from_tables(page, commission_filter) -> List[PDFRecord]:
    """
    Extrai registros de uma página usando extract_tables().
    Tenta detecção dinâmica de colunas; usa fallback para posições fixas.

    Parâmetros:
        page              — página do pdfplumber
        commission_filter  — função que recebe o valor da comissão e retorna bool
    """
    records: List[PDFRecord] = []
    tables = page.extract_tables()
    for table in tables:
        if not table:
            continue

        # Tenta detectar colunas dinamicamente pelo cabeçalho
        detected = _detect_columns_from_header(table)
        if detected[0] is not None:
            col_segurado, col_inicio_vig, col_apolice, col_comissao = detected
            if col_apolice is None:
                col_apolice = _DEFAULT_COL_APOLICE
            min_cols = max(col_segurado, col_inicio_vig, col_comissao) + 1
        else:
            col_segurado = _DEFAULT_COL_SEGURADO
            col_inicio_vig = _DEFAULT_COL_INICIO_VIG
            col_apolice = _DEFAULT_COL_APOLICE
            col_comissao = _DEFAULT_COL_COMISSAO
            min_cols = _MIN_COLS

        for row_idx, row in enumerate(table):
            segurado, inicio_vig = _extract_record_from_row(
                row, col_segurado, col_inicio_vig, col_comissao, min_cols
            )
            if segurado is None:
                continue
            if commission_filter(row[col_comissao]):
                apolice = _extract_apolice_from_table(table, row_idx, col_apolice)
                records.append(PDFRecord(
                    segurado=segurado,
                    inicio_vig=inicio_vig,
                    apolice=apolice,
                ))
    return records


def _extract_from_text(page, commission_filter) -> List[PDFRecord]:
    """
    Fallback: extrai registros via texto bruto + regex.
    Captura registros que extract_tables() perde.

    Parâmetros:
        page              — página do pdfplumber
        commission_filter  — função que recebe o valor da comissão e retorna bool
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
        apolice = _normalize_apolice(match.group(3))
        comissao_str = match.group(4)

        if commission_filter(comissao_str):
            records.append(PDFRecord(
                segurado=segurado,
                inicio_vig=inicio_vig,
                apolice=apolice,
            ))

    return records


def _deduplicate_and_collect(
    pdf_path: str, commission_filter
) -> List[PDFRecord]:
    """
    Abre o PDF e retorna PDFRecords filtrados por commission_filter.

    Usa extração dupla (tabelas + texto) e mescla resultados sem duplicatas.
    A chave de deduplicação é (segurado_upper, inicio_vig).
    Prioriza registros que têm apólice sobre os que não têm.
    """
    seen: dict[tuple, PDFRecord] = {}
    records: List[PDFRecord] = []

    def _add_if_new(rec: PDFRecord) -> None:
        key = (rec.segurado.upper(), rec.inicio_vig, rec.apolice)
        existing = seen.get(key)
        
        if existing is None:
            # Se não tem apólice, vamos checar se já existe um registro sem apólice para mesma pessoa/data
            if not rec.apolice:
                # Procura se já existe a mesma pessoa/data (com ou sem apolice)
                for k, v in list(seen.items()):
                    if k[0] == rec.segurado.upper() and k[1] == rec.inicio_vig:
                        # Se já existe, não adicionamos este novo (ele não tem apolice, não traz info nova)
                        return
                        
            seen[key] = rec
            records.append(rec)
            
        elif not existing.apolice and rec.apolice:
            # Substitui se o novo registro tem apólice e o existente não
            # Na verdade, com a nova key, isso só acontece se ambos tivessem apolice='' na key
            idx = records.index(existing)
            records[idx] = rec
            seen[key] = rec

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for rec in _extract_from_tables(page, commission_filter):
                _add_if_new(rec)

            for rec in _extract_from_text(page, commission_filter):
                _add_if_new(rec)

    return records


def extract_negative_commission_records(pdf_path: str) -> List[PDFRecord]:
    """Abre o PDF e retorna todos os PDFRecord onde COMISSÃO < 0."""
    return _deduplicate_and_collect(pdf_path, _is_negative_commission)


def extract_positive_commission_records(pdf_path: str) -> List[PDFRecord]:
    """Abre o PDF e retorna todos os PDFRecord onde COMISSÃO > 0."""
    return _deduplicate_and_collect(pdf_path, _is_positive_commission)
