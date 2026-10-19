from __future__ import annotations

from ..coerce import coerce_cell
from ..config import CR_FOOTER, CR_INDEX_KINDS, CR_KINDS, R_BLANK, R_FOOTER
from ..headers import check_header, load_headers, normalize
from ..readers import read_delimited
from ..types import ContractError, Removal, Sanitized

REPORT = "cr"


def _kind(i: int, name: str) -> str:
    return CR_INDEX_KINDS.get(i) or CR_KINDS.get(name, "text")


def _is_footer(raw: list[str]) -> bool:
    """Rodapés TOTAL PÁGINA / TOTAL GERAL: rótulo exato na posição em que o bruto o traz.
    Só são reconhecidos para serem descartados como linha técnica; os valores não são lidos."""
    i, labels = CR_FOOTER
    return len(raw) > i and normalize(raw[i]) in labels


def sanitize_cr(path, *, expected_header: list[str] | None = None) -> Sanitized:
    enc, rows = read_delimited(path)
    if not rows:
        raise ContractError("cr: arquivo vazio")
    header = [normalize(h) for h in rows[0]]
    check_header(REPORT, header, expected_header or load_headers()[REPORT])
    ncols = len(header)

    out = Sanitized(report=REPORT, header=header, rows=[], db_rows=[], source_rows=[],
                    raw_count=len(rows) - 1, encoding=enc)
    for n, raw in enumerate(rows[1:], start=2):
        if not any(c.strip() for c in raw):
            out.technical.append(Removal(n, R_BLANK))
            continue
        if _is_footer(raw):
            out.technical.append(Removal(n, R_FOOTER))
            continue
        if len(raw) != ncols:   # aspas não fechadas, registro truncado, rodapé desconhecido
            raise ContractError(
                f"cr: registro {n} tem {len(raw)} colunas; esperadas {ncols}")
        typed = []
        for i, cell in enumerate(raw):
            try:
                typed.append(coerce_cell(_kind(i, header[i]), cell))
            except ValueError as e:
                raise ContractError(f"cr: registro {n}, coluna {header[i]!r}: {e}") from e
        out.rows.append(typed)
        out.db_rows.append(list(typed))   # listas distintas: rows (Excel) e db_rows (banco)
        out.source_rows.append(n)
    return out
