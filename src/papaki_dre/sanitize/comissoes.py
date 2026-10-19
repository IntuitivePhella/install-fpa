from __future__ import annotations

from ..config import COMISSOES_FOOTER, DEST, PESQUISA, R_BLANK, R_FOOTER, R_H, R_MODELO, R_PESQ
from ..headers import check_header, load_headers, normalize
from ..produtoras import Regras, regras_padrao
from ..readers import open_xlsx
from ..types import Removal, Sanitized
from ._xlsx import RowCopier, blank, check_header_extent

REPORT = "comissoes"
NCOLS = 14
HEADER_ROW = 3
FIRST_RAW_ROW = 4


def _is_footer(values) -> bool:
    """Linha final TOTAL GERAL do bruto: coluna A vazia, rótulo exato na coluna H e só
    somas em J:N. Reconhecida para ser descartada como técnica; os valores não são lidos."""
    label = values[7]
    return (isinstance(label, str) and normalize(label) == COMISSOES_FOOTER
            and blank(values[:7]) and blank(values[8:9]))


def sanitize_comissoes(path, *, expected_header: list[str] | None = None,
                       regras: Regras | None = None) -> Sanitized:
    regras = regras or regras_padrao()
    excluir_h = {normalize(x).casefold() for x in regras.comissoes_excluir_h}
    ws = open_xlsx(path).worksheets[0]
    wv = open_xlsx(path, data_only=True).worksheets[0]
    header_cells = ws[HEADER_ROW]
    header = [normalize(c.value) if c.value is not None else "" for c in header_cells[:NCOLS]]
    check_header(REPORT, header, expected_header or load_headers()[REPORT])
    check_header_extent(REPORT, header_cells, NCOLS)

    first_dest = DEST[REPORT]["first_row"]
    out = Sanitized(report=REPORT, header=header, rows=[], db_rows=[], source_rows=[],
                    raw_count=max(ws.max_row - HEADER_ROW, 0), title=ws["A1"].value)
    copier = RowCopier(REPORT)
    seen_rows = 0
    # iter_rows com max_col fixo: `ws[r]` recalcula max_column a cada linha (quadrático).
    last = ws.max_row
    pairs = zip(ws.iter_rows(min_row=FIRST_RAW_ROW, max_row=last, max_col=NCOLS),
                wv.iter_rows(min_row=FIRST_RAW_ROW, max_row=last, max_col=NCOLS))
    for r, (cells, vcells) in enumerate(pairs, start=FIRST_RAW_ROW):
        values = [c.value for c in cells]
        seen_rows += 1
        if blank(values):
            out.technical.append(Removal(r, R_BLANK))
            continue
        if _is_footer(values):
            out.technical.append(Removal(r, R_FOOTER))
            continue
        modelo = normalize(values[0]) if values[0] is not None else ""
        titulo = str(values[6] or "").upper()
        opcao = normalize(values[7]).casefold() if values[7] is not None else ""
        if modelo != regras.modelo_comissoes:
            out.exclusions.append(Removal(r, R_MODELO))
            continue
        if PESQUISA in titulo:
            out.exclusions.append(Removal(r, R_PESQ))
            continue
        if opcao in excluir_h:
            out.exclusions.append(Removal(r, R_H))
            continue
        dest = first_dest + len(out.rows)
        row, db_row = copier.copy(cells, vcells, r, dest, values[1])
        out.rows.append(row)
        out.db_rows.append(db_row)
        out.source_rows.append(r)
    out.raw_count = seen_rows
    out.warnings.extend(copier.warnings())
    return out
