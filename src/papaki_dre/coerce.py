from __future__ import annotations

import datetime as dt
import re

# Únicos textos aceitos nas colunas money_or_text (CP: Deduções, Alterado; CR: Desconto,
# Outras ded., Multa, Juros, IDA, Atraso). Levantado no bruto real de set/2026: só "-".
MONEY_TEXT_OK = frozenset({"-"})

_NUMBER = re.compile(r"^(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?$")


def parse_money(text: str) -> float:
    """'R$ 21.132,72' -> 21132.72. Aceita um único sinal ('-' antes ou depois de 'R$', ou
    parênteses) e número puro. Sinal duplo ('-R$ -5,00', '(-5)') é rejeitado."""
    s = text.replace(" ", " ").strip()
    signs = 0
    if s.startswith("(") and s.endswith(")"):
        signs, s = signs + 1, s[1:-1].strip()
    if s.startswith("-"):
        signs, s = signs + 1, s[1:].strip()
    if s.startswith("R$"):
        s = s[2:].strip()
    if s.startswith("-"):
        signs, s = signs + 1, s[1:].strip()
    if signs > 1 or not _NUMBER.match(s):
        raise ValueError("valor monetário inválido (texto omitido)")
    value = float(s.replace(".", "").replace(",", "."))
    return -value if signs else value


def parse_date(text: str) -> dt.datetime:
    s = text.strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError("data inválida (texto omitido)")


def parse_datetime(text: str) -> dt.datetime:
    s = text.strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            continue
    return parse_date(s)


def coerce_cell(kind: str, raw):
    """Converte um valor bruto de texto. Vazio vira None (nunca zero)."""
    if raw is None:
        return None
    s = raw if isinstance(raw, str) else str(raw)
    if s.strip() == "":
        return None
    if kind == "text":
        return s
    if kind == "date":
        return parse_date(s)
    if kind == "datetime":
        return parse_datetime(s)
    if kind == "money":
        return parse_money(s)
    if kind == "money_or_text":
        if s.strip() in MONEY_TEXT_OK:
            return s
        return parse_money(s)          # qualquer outro texto: ValueError (vira ContractError)
    raise ValueError(f"tipo de coluna desconhecido: {kind!r}")
