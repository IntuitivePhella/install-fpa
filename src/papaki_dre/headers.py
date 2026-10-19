from __future__ import annotations

import json
import unicodedata
from importlib import resources

from .types import ContractError


def normalize(name) -> str:
    return unicodedata.normalize("NFC", str(name).strip())


def load_headers() -> dict[str, list[str]]:
    text = resources.files("papaki_dre").joinpath("headers.json").read_text(encoding="utf-8")
    return json.loads(text)


def check_header(report: str, header: list[str], expected: list[str]) -> None:
    """Para a rodada se o cabeçalho divergir em quantidade ou ordem."""
    if len(header) != len(expected):
        raise ContractError(
            f"{report}: o arquivo tem {len(header)} colunas; esperadas {len(expected)}")
    for i, (got, want) in enumerate(zip(header, expected), start=1):
        if got != want:
            raise ContractError(f"{report}: coluna {i} é {got!r}; esperado {want!r}")
