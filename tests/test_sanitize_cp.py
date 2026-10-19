import pytest

from papaki_dre.config import R_BLANK, R_FOOTER, R_JOB, R_LOC
from papaki_dre.headers import load_headers
from papaki_dre.produtoras import regras_padrao
from papaki_dre.sanitize.cp import sanitize_cp
from papaki_dre.types import ContractError

UTLD = regras_padrao().localidade_cp
H = load_headers()["cp"]


def _r(d):
    row = [""] * len(H)
    for name, value in d.items():
        row[H.index(name)] = value
    return row


def _write(tmp_path, rows, header=None, enc="utf-8-sig"):
    p = tmp_path / "cp.csv"
    lines = [";".join(r) for r in [header or H] + rows]
    p.write_bytes(("\n".join(lines) + "\n").encode(enc))
    return p


def _footer(total):
    f = [""] * 23
    f[4], f[5] = "VALOR TOTAL", total
    return f


BASE = {"Cód.": "untitled-1", "Valor": "R$ 100,00", "Localidade": UTLD,
        "Vcto": "05/01/2026", "Competência": "01/01/2026"}


def test_keeps_untitled_rows_without_job(tmp_path):
    p = _write(tmp_path, [_r(BASE), _footer("R$ 100,00")])
    s = sanitize_cp(p)
    assert s.kept == 1
    assert s.rows[0][H.index("Valor")] == 100.0
    assert s.rows[0][H.index("Vcto")].year == 2026
    assert s.balanced()


def test_excludes_filled_job_without_pesquisa(tmp_path):
    ex = _r({**BASE, "Cód.": "untitled-2", "Nº job": "#75111", "Job": "Meli | Gut"})
    s = sanitize_cp(_write(tmp_path, [_r(BASE), ex, _footer("R$ 200,00")]))
    assert s.kept == 1
    assert [e.reason for e in s.exclusions] == [R_JOB]
    assert s.exclusions[0].source_row == 3   # cabeçalho = 1; primeira linha de dados = 2


def test_keeps_filled_job_with_pesquisa(tmp_path):
    ok = _r({**BASE, "Nº job": "#75111", "Job": "PESQUISA E TRATAMENTO | X"})
    s = sanitize_cp(_write(tmp_path, [ok, _footer("R$ 100,00")]))
    assert s.kept == 1 and not s.exclusions


def test_excludes_other_localidade(tmp_path):
    other = _r({**BASE, "Localidade": "SOMA"})
    s = sanitize_cp(_write(tmp_path, [_r(BASE), other, _footer("R$ 200,00")]))
    assert s.kept == 1
    assert [e.reason for e in s.exclusions] == [R_LOC]


def test_job_rule_runs_before_localidade_filter(tmp_path):
    both = _r({**BASE, "Localidade": "SOMA", "Nº job": "#1", "Job": "x"})
    s = sanitize_cp(_write(tmp_path, [both, _footer("R$ 100,00")]))
    assert [e.reason for e in s.exclusions] == [R_JOB]


def test_footer_and_blank_are_technical(tmp_path):
    p = _write(tmp_path, [_r(BASE), [""] * len(H), _footer("R$ 100,00")])
    s = sanitize_cp(p)
    assert sorted(t.reason for t in s.technical) == sorted([R_BLANK, R_FOOTER])
    assert s.kept == 1 and s.balanced()
    assert not hasattr(s, "controls")


def test_footer_value_is_never_read(tmp_path):
    s = sanitize_cp(_write(tmp_path, [_r(BASE), _footer("ilegível")]))
    assert [t.reason for t in s.technical] == [R_FOOTER]
    assert not s.warnings


@pytest.mark.parametrize("pos,label", [(3, "VALOR TOTAL"), (4, "VALOR TOTAL GERAL"), (4, "TOTAL")])
def test_footer_label_must_be_exact_and_at_index_4(tmp_path, pos, label):
    loose = [""] * 23
    loose[pos], loose[5] = label, "R$ 1,00"
    with pytest.raises(ContractError, match="colunas"):
        sanitize_cp(_write(tmp_path, [_r(BASE), loose]))


def test_short_record_stops_the_run(tmp_path):
    with pytest.raises(ContractError, match="registro 3"):
        sanitize_cp(_write(tmp_path, [_r(BASE), _r(BASE)[:-1], _footer("R$ 1,00")]))


def test_unclosed_quote_stops_the_run(tmp_path):
    bad = _r({**BASE, "Cód.": '"sem fechar'})
    with pytest.raises(ContractError):
        sanitize_cp(_write(tmp_path, [_r(BASE), bad, _r(BASE), _footer("R$ 1,00")]))


def test_quoted_multiline_field_is_valid(tmp_path):
    ok = _r({**BASE, "Cód.": '"linha1\nlinha2"'})
    s = sanitize_cp(_write(tmp_path, [ok, _footer("R$ 1,00")]))
    assert s.kept == 1 and s.rows[0][H.index("Cód.")] == "linha1\nlinha2"


def test_duplicate_footer_is_silent(tmp_path):
    s = sanitize_cp(_write(tmp_path, [_r(BASE), _footer("R$ 1,00"), _footer("R$ 1,00")]))
    assert [t.reason for t in s.technical] == [R_FOOTER, R_FOOTER]
    assert not s.warnings


@pytest.mark.parametrize("text", ["-R$ 5,00", "R$ -5,00", "(R$ 5,00)"])
def test_negative_valor(tmp_path, text):
    s = sanitize_cp(_write(tmp_path, [_r({**BASE, "Valor": text})]))
    assert s.rows[0][H.index("Valor")] == -5.0


def test_rows_and_db_rows_are_not_aliased(tmp_path):
    s = sanitize_cp(_write(tmp_path, [_r(BASE)]))
    assert s.rows[0] is not s.db_rows[0]
    s.rows[0][H.index("Valor")] = "=X1"
    assert s.db_rows[0][H.index("Valor")] == 100.0


def test_header_mismatch_stops_the_run(tmp_path):
    bad = H[:-1] + ["Outra coisa"]
    with pytest.raises(ContractError, match="coluna 41"):
        sanitize_cp(_write(tmp_path, [_r(BASE)], header=bad))


def test_invalid_value_stops_the_run(tmp_path):
    bad = _r({**BASE, "Valor": "dez reais"})
    with pytest.raises(ContractError, match="Valor"):
        sanitize_cp(_write(tmp_path, [bad, _footer("R$ 0,00")]))


def test_missing_footer_is_silent(tmp_path):
    s = sanitize_cp(_write(tmp_path, [_r(BASE)]))
    assert s.kept == 1 and not s.warnings


@pytest.mark.real
def test_real_september_file(cp_real):
    s = sanitize_cp(cp_real)
    assert s.encoding == "utf-8-sig"
    assert s.balanced()
    assert s.raw_count == 36999
    assert s.kept == 1110
    assert len(s.exclusions) == 21505 + 14382
    assert sorted(t.reason for t in s.technical) == [R_BLANK, R_FOOTER]


def test_regras_de_outra_produtora_mudam_a_localidade_que_entra(tmp_path):
    from papaki_dre.produtoras import Regras
    outra = Regras(localidade_cp="OUTRA LTDA", nucleo_rra="Projetos - Outra",
                   modelo_comissoes="Outra - Padrão", comissoes_excluir_h=("Impostos",))
    da_untitled = _r({**BASE, "Cód.": "untitled-9"})
    da_outra = _r({**BASE, "Cód.": "outra-1", "Localidade": "OUTRA LTDA"})
    p = _write(tmp_path, [da_untitled, da_outra, _footer("R$ 200,00")])
    s = sanitize_cp(p, regras=outra)
    assert s.kept == 1
    assert s.rows[0][H.index("Cód.")] == "outra-1"
    assert [e.reason for e in s.exclusions] == [R_LOC]
    assert s.exclusions[0].source_row == 2             # a linha da Untitled saiu por não ser da Localidade
