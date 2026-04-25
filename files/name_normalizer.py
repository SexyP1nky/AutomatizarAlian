"""
Responsabilidade: normalização e comparação de nomes.

Regras:
- Remove acentos de todas as letras EXCETO ç/Ç
- Converte para maiúsculas
- Normaliza espaços múltiplos
- Comparação exata após normalização
- Exceção: tokens que contêm Ç permitem aproximação Ç ↔ C
"""

import unicodedata


def normalize_name(name: str) -> str:
    """
    Remove diacríticos exceto do ç/Ç, converte para maiúsculas
    e normaliza espaços.
    """
    if not name:
        return ""

    result = []
    for char in name:
        if char in ("ç", "Ç"):
            result.append("Ç")
            continue
        decomposed = unicodedata.normalize("NFD", char)
        without_diacritics = "".join(
            c for c in decomposed if unicodedata.category(c) != "Mn"
        )
        result.append(without_diacritics)

    normalized = "".join(result).upper()
    return " ".join(normalized.split())


def _tokens_match(tok_a: str, tok_b: str) -> bool:
    """
    Compara dois tokens individualmente.
    Se algum contém Ç, permite substituição Ç ↔ C nesse token específico.
    """
    if tok_a == tok_b:
        return True
    if "Ç" in tok_a or "Ç" in tok_b:
        return tok_a.replace("Ç", "C") == tok_b.replace("Ç", "C")
    return False


def names_match(name_a: str, name_b: str) -> bool:
    """
    Retorna True se dois nomes normalizados são equivalentes.
    Correspondência exata token a token, com exceção Ç ≈ C.
    """
    if name_a == name_b:
        return True

    tokens_a = name_a.split()
    tokens_b = name_b.split()

    if len(tokens_a) != len(tokens_b):
        return False

    return all(_tokens_match(a, b) for a, b in zip(tokens_a, tokens_b))
