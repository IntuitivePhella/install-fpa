from __future__ import annotations

from ..coerce import coerce_cell
from ..config import CP_FOOTER, CP_KINDS, PESQUISA, R_BLANK, R_FOOTER, R_JOB, R_LOC
from ..headers import check_header, load_headers, normalize
from ..produtoras import Regras, regras_padrao
from ..readers import read_delimited
from ..types import ContractError, Removal, Sanitized

REPORT = "cp"


def _is_footer(raw: list[str]) -> bool:
    """Rodapé VALOR TOTAL: rótulo exato na posição em que o bruto o traz. Só é reconhecido
    para ser descartado como linha técnica; seus valores nunca são lidos."""
    i, label = CP_FOOTER
    return len(raw) > i and raw[i].strip() == label


def sanitize_cp(path, *, expected_header: list[str] | None = None,
                regras: Regras | None = None) -> Sanitized:
    regras = regras or regras_padrao()      # o `importar` sempre informa as da produtora
    enc, rows = read_delimited(path)
    if not rows:
        raise ContractError("cp: arquivo vazio")
    header = [normalize(h) for h in rows[0]]
    check_header(REPORT, header, expected_header or load_headers()[REPORT])
    ncols = len(header)
    ix = {name: i for i, name in enumerate(header)}
    i_job, i_nome, i_loc = ix["Nº job"], ix["Job"], ix["Localidade"]

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
                f"cp: registro {n} tem {len(raw)} colunas; esperadas {ncols}")
        typed = []
        for i, cell in enumerate(raw):
            kind = CP_KINDS.get(header[i], "text")
            try:
                typed.append(coerce_cell(kind, cell))
            except ValueError as e:
                raise ContractError(f"cp: registro {n}, coluna {header[i]!r}: {e}") from e
        if raw[i_job].strip() and PESQUISA not in raw[i_nome].upper():
            out.exclusions.append(Removal(n, R_JOB))
            continue
        if normalize(raw[i_loc]) != regras.localidade_cp:
            out.exclusions.append(Removal(n, R_LOC))
            continue
        out.rows.append(typed)
        out.db_rows.append(list(typed))   # listas distintas: rows (Excel) e db_rows (banco)
        out.source_rows.append(n)
    return out
