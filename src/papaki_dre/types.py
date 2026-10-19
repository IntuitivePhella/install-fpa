from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


class ContractError(Exception):
    """Violação estrutural (cabeçalho, nº de colunas, erro inesperado). Para a rodada."""


class Formula(str):
    """Texto de fórmula do Excel, com o '=' inicial."""


class XlError(str):
    """Valor de erro literal do Excel (ex.: '#DIV/0!')."""


@dataclass(frozen=True)
class Removal:
    """Uma linha do bruto que não chegou ao resultado."""
    source_row: int  # nº do registro no bruto (o cabeçalho conta como 1 em CSV; nº da linha em planilhas)
    reason: str


@dataclass
class Sanitized:
    report: str
    header: list[str]
    rows: list[list[Any]]          # valores tipados; fórmulas como Formula (para o Excel)
    db_rows: list[list[Any]]       # somente valores (para o banco)
    source_rows: list[int]         # linha de origem de cada linha mantida
    raw_count: int                 # registros do bruto (sem o cabeçalho)
    technical: list[Removal] = field(default_factory=list)   # linhas vazias e rodapés (TOTAL)
    exclusions: list[Removal] = field(default_factory=list)  # exclusões de negócio
    warnings: list[str] = field(default_factory=list)
    title: Any = None
    encoding: str | None = None

    @property
    def kept(self) -> int:
        return len(self.rows)

    def balanced(self) -> bool:
        return self.raw_count == self.kept + len(self.technical) + len(self.exclusions)
