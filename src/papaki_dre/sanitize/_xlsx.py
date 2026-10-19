"""Partes comuns dos sanitizadores de planilha (Comissões e RRA)."""
from __future__ import annotations

from openpyxl.utils import get_column_letter

from ..formulas import readdress
from ..types import ContractError, XlError

EXPECTED_ERRORS = {"#DIV/0!"}


def blank(values) -> bool:
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in values)


def check_header_extent(report: str, header_cells, ncols: int) -> None:
    """Célula de cabeçalho preenchida além das `ncols` colunas do contrato para a rodada.
    Células vazias além (o bruto do RRA vai até a coluna 72) não contam."""
    extra = [c for c in header_cells[ncols:] if not blank([c.value])]
    if extra:
        refs = ", ".join(c.coordinate for c in extra)
        raise ContractError(f"{report}: cabeçalho com coluna extra além da {ncols}ª: {refs}")


class RowCopier:
    """Copia uma linha mantida célula a célula, com a mesma política nos dois relatórios:
    - erro em cache diferente de #DIV/0! -> ContractError;
    - fórmula -> reendereçada para a linha de destino; no banco vai o valor em cache
      (erro esperado vira None; sem cache vira None e gera um aviso agregado);
    - erro literal -> XlError na planilha, None no banco."""

    def __init__(self, report: str):
        self.report = report
        self.missing_cache: list[str] = []

    def copy(self, cells, vcells, r: int, dest: int, label) -> tuple[list, list]:
        row, db_row = [], []
        for c, vc in zip(cells, vcells):
            letter = get_column_letter(c.column)
            if vc.data_type == "e" and str(vc.value) not in EXPECTED_ERRORS:
                raise ContractError(
                    f"{self.report}: erro inesperado {vc.value} em {letter}{r} (job {label!r})")
            v = c.value
            if isinstance(v, str) and v.startswith("="):
                row.append(readdress(v, f"{letter}{r}", f"{letter}{dest}"))
                if vc.data_type == "e":
                    db_row.append(None)
                else:
                    if vc.value is None:
                        self.missing_cache.append(f"{letter}{r}")
                    db_row.append(vc.value)
            elif c.data_type == "e":
                row.append(XlError(v))
                db_row.append(None)
            else:
                row.append(v)
                db_row.append(v)
        return row, db_row

    def warnings(self) -> list[str]:
        if not self.missing_cache:
            return []
        n, first = len(self.missing_cache), self.missing_cache[0]
        return [f"{self.report}: {n} fórmula(s) de linhas mantidas sem valor em cache "
                f"(ex.: {first}); gravadas como nulo no banco"]
