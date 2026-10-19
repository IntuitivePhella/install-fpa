import pytest

from papaki_dre.config import R_BLANK, R_FOOTER
from papaki_dre.headers import load_headers
from papaki_dre.sanitize.cr import sanitize_cr
from papaki_dre.types import ContractError

H = load_headers()["cr"]
SHIFT = 3   # no bruto real o rodapé tem 31 colunas, com os totais deslocados


def _r(d):
    row = [""] * len(H)
    for name, value in d.items():
        row[H.index(name)] = value
    return row


def _write(tmp_path, rows, header=None, enc="cp1252"):
    p = tmp_path / "cr.csv"
    p.write_bytes(("\n".join(";".join(r) for r in [header or H] + rows) + "\n").encode(enc))
    return p


def _footer(label, doc, cob, rec, dif):
    f = [""] * (len(H) + SHIFT)
    f[6] = label
    f[H.index("Vl Documento") + SHIFT] = doc
    f[H.index("Vl cobrado") + SHIFT] = cob
    f[H.index("Vl recebido") + SHIFT] = rec
    f[H.index("Dif. receber") + SHIFT] = dif
    return f


ROW = {"Cliente": "ITAU", "Emissão": "05/01/2026", "Vcto": "05/01/2026",
       "Vl Documento": "R$ 100,00", "Vl cobrado": "R$ 100,00", "Vl recebido": "R$ 60,00",
       "Dif. receber": "R$ 40,00", "Desconto": "-", "Competência": "2026-01-01",
       "Descrição": "RENDIMENTOS ação"}


def _file(tmp_path, extra=()):
    return _write(tmp_path, [_r(ROW), *extra,
                             _footer("TOTAL PÁGINA", "R$ 100,00", "R$ 100,00", "R$ 60,00", "R$ 40,00"),
                             _footer("TOTAL GERAL", "R$ 100,00", "R$ 100,00", "R$ 60,00", "R$ 40,00")])


def test_reads_cp1252_and_types_the_columns(tmp_path):
    s = sanitize_cr(_file(tmp_path))
    assert s.encoding == "cp1252"
    row = s.rows[0]
    assert row[H.index("Descrição")] == "RENDIMENTOS ação"
    assert row[H.index("Vl Documento")] == 100.0
    assert row[H.index("Emissão")].month == 1
    assert row[H.index("Competência")].year == 2026
    assert row[H.index("Desconto")] == "-"          # texto preservado


def test_second_recebido_column_stays_text(tmp_path):
    second = max(i for i, n in enumerate(H) if n == "Recebido")
    first = H.index("Recebido")
    row = _r({**ROW})
    row[first] = "06/01/2026"
    row[second] = "Sim"
    s = sanitize_cr(_write(tmp_path, [row]))
    assert s.rows[0][first].day == 6
    assert s.rows[0][second] == "Sim"


def test_footers_are_technical(tmp_path):
    s = sanitize_cr(_file(tmp_path))
    assert s.kept == 1
    assert [t.reason for t in s.technical] == [R_FOOTER, R_FOOTER]
    assert not hasattr(s, "controls")
    assert not s.warnings
    assert s.balanced()


def test_footer_values_are_never_read(tmp_path):
    p = _write(tmp_path, [_r(ROW), _footer("TOTAL GERAL", "lixo", "", "x", "")])
    s = sanitize_cr(p)
    assert [t.reason for t in s.technical] == [R_FOOTER] and not s.warnings


def test_missing_footer_is_silent(tmp_path):
    s = sanitize_cr(_write(tmp_path, [_r(ROW)]))
    assert s.kept == 1 and not s.warnings


def test_totalpass_client_is_data_not_footer(tmp_path):
    row = _r({**ROW, "Cliente": "TOTALPASS"})
    s = sanitize_cr(_file(tmp_path, [row]))
    assert s.kept == 2
    assert s.rows[1][H.index("Cliente")] == "TOTALPASS"
    assert [t.reason for t in s.technical] == [R_FOOTER, R_FOOTER]


def test_footer_label_only_counts_at_index_6(tmp_path):
    f = _footer("", "R$ 1,00", "R$ 1,00", "R$ 1,00", "R$ 1,00")
    f[2] = "TOTAL GERAL"
    with pytest.raises(ContractError, match="colunas"):
        sanitize_cr(_write(tmp_path, [_r(ROW), f]))


def test_startswith_total_is_not_a_footer(tmp_path):
    f = _footer("TOTALPASS", "R$ 1,00", "R$ 1,00", "R$ 1,00", "R$ 1,00")
    with pytest.raises(ContractError, match="colunas"):
        sanitize_cr(_write(tmp_path, [_r(ROW), f]))


def test_short_record_stops_the_run(tmp_path):
    with pytest.raises(ContractError, match="registro 3"):
        sanitize_cr(_write(tmp_path, [_r(ROW), _r(ROW)[:-2]]))


def test_unclosed_quote_stops_the_run(tmp_path):
    bad = _r({**ROW, "Cliente": '"ITAU'})
    with pytest.raises(ContractError):
        sanitize_cr(_write(tmp_path, [_r(ROW), bad, _r(ROW)]))


def test_quoted_multiline_field_is_valid(tmp_path):
    ok = _r({**ROW, "Descrição": '"linha1\nlinha2"'})
    s = sanitize_cr(_file(tmp_path, [ok]))
    assert s.kept == 2 and s.rows[1][H.index("Descrição")] == "linha1\nlinha2"


def test_never_excludes_business_rows(tmp_path):
    desconsiderar = _r({**ROW, "Plano de contas": "Desconsiderar"})
    s = sanitize_cr(_file(tmp_path, [desconsiderar]))
    assert s.kept == 2 and not s.exclusions


def test_blank_rows_are_technical(tmp_path):
    s = sanitize_cr(_file(tmp_path, [[""] * len(H)]))
    assert R_BLANK in [t.reason for t in s.technical]


def test_identifier_keeps_leading_zeros(tmp_path):
    s = sanitize_cr(_write(tmp_path, [_r({**ROW, "NFe": "000123", "Lote": "0042"})]))
    assert s.rows[0][H.index("NFe")] == "000123"
    assert s.rows[0][H.index("Lote")] == "0042"


def test_rows_and_db_rows_are_not_aliased(tmp_path):
    s = sanitize_cr(_write(tmp_path, [_r(ROW)]))
    assert s.rows[0] is not s.db_rows[0]
    s.rows[0][H.index("Cliente")] = "OUTRO"
    assert s.db_rows[0][H.index("Cliente")] == "ITAU"


def test_header_mismatch_stops_the_run(tmp_path):
    with pytest.raises(ContractError):
        sanitize_cr(_write(tmp_path, [_r(ROW)], header=H[:-1]))


@pytest.mark.real
def test_real_september_file(cr_real):
    s = sanitize_cr(cr_real)
    assert s.encoding == "cp1252"
    assert s.balanced()
    assert s.kept == 484
    assert s.raw_count == 486
    assert [t.reason for t in s.technical] == [R_FOOTER, R_FOOTER]
    assert not s.exclusions
