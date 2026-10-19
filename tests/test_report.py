import pytest

from papaki_dre.build import BuildReport, SheetReport
from papaki_dre.config import R_BLANK, R_FOOTER, R_JOB, R_LOC
from papaki_dre.report import contract_blocked, conference, render_markdown
from papaki_dre.types import ContractError, Removal, Sanitized


def _s(report, kept=2, technical=1, exclusions=(), warnings=()):
    return Sanitized(
        report=report, header=["a"], rows=[["x"]] * kept, db_rows=[["x"]] * kept,
        source_rows=list(range(kept)), raw_count=kept + technical + len(exclusions),
        technical=[Removal(i, R_FOOTER) for i in range(technical)],
        exclusions=list(exclusions), warnings=list(warnings))


def _results(**over):
    base = {k: _s(k) for k in ("rra", "cp", "cr", "comissoes")}
    base.update(over)
    return base


def _build(**over):
    sheets = {
        "BASE": SheetReport("BASE", 5, 6, 2),
        "DESPESAS": SheetReport("DESPESAS", 3, 4, 2),
        "RECEITA": SheetReport("RECEITA", 2, 3, 2),
        "REAL X ORÇADO COMISSÕES": SheetReport("REAL X ORÇADO COMISSÕES", 2, 3, 2),
    }
    sheets.update(over)
    return BuildReport(sheets=sheets, parts={"changed": ["a"], "removed": ["xl/calcChain.xml"], "added": []})


def test_clean_run_is_ok():
    conf = conference(_results(), _build())
    assert conf["status"] == "ok" and not conf["alertas"] and not conf["bloqueios"]


def test_footer_rows_are_only_counted_never_checked():
    # 3 rodapés de TOTAL removidos como técnicos: entram na contagem e em mais nada
    cp = _s("cp", technical=3)
    conf = conference(_results(cp=cp), _build())
    assert conf["status"] == "ok"
    assert conf["relatorios"]["cp"]["tecnicas_por_motivo"] == {R_FOOTER: 3}
    assert "controles" not in conf["relatorios"]["cp"]


def test_count_flow_is_reported():
    ex = [Removal(10, R_JOB), Removal(11, R_LOC), Removal(12, R_LOC)]
    cp = _s("cp", kept=2, technical=1, exclusions=ex)
    cp.technical.append(Removal(99, R_BLANK))
    cp.raw_count += 1
    r = conference(_results(cp=cp), _build())["relatorios"]["cp"]
    assert r["registros_bruto"] == 7
    assert r["removidas_tecnicas"] == 2 and r["tecnicas_por_motivo"] == {R_FOOTER: 1, R_BLANK: 1}
    assert r["registros_de_dados"] == 5
    assert r["excluidas"] == {R_JOB: 1, R_LOC: 2}
    assert r["mantidas"] == 2 and r["carregadas"] == 2


def test_screen_count_difference_is_an_alert_not_an_error():
    cp = _s("cp", kept=3, technical=2)              # 3 registros de dados
    conf = conference(_results(cp=cp), _build(DESPESAS=SheetReport("DESPESAS", 3, 5, 3)),
                      screen_counts={"cp": 5})
    assert conf["status"] == "com_alertas" and not conf["bloqueios"]
    assert any("tela mostrava 5" in a and "tem 3" in a for a in conf["alertas"])


def test_matching_screen_count_is_silent():
    cp = _s("cp", kept=3, technical=2)
    conf = conference(_results(cp=cp), _build(DESPESAS=SheetReport("DESPESAS", 3, 5, 3)),
                      screen_counts={"cp": 3})
    assert conf["status"] == "ok"


def test_unbalanced_counts_block():
    s = _s("rra")
    s.raw_count += 5
    conf = conference(_results(rra=s), _build())
    assert conf["status"] == "bloqueada" and "não fecham" in conf["bloqueios"][0]


def test_kept_rows_must_all_be_loaded():
    conf = conference(_results(), _build(RECEITA=SheetReport("RECEITA", 2, 2, 1)))
    assert conf["status"] == "bloqueada"
    assert "RECEITA" in conf["bloqueios"][0] and "carregadas" in conf["bloqueios"][0]


def test_missing_model_formulas_block():
    sr = SheetReport("DESPESAS", 3, 4, 2, missing_formulas=["AP4", "AQ4"])
    conf = conference(_results(), _build(DESPESAS=sr))
    assert conf["status"] == "bloqueada" and "DESPESAS" in conf["bloqueios"][0]


def test_coverage_gaps_in_commissions_and_base_total_block():
    com = SheetReport("REAL X ORÇADO COMISSÕES", 2, 3, 2, missing_formulas=["L3", "M3", "N3"])
    base = SheetReport("BASE", 5, 6, 2, missing_formulas=["A8", "J8"], total_row=8)
    conf = conference(_results(), _build(BASE=base, **{"REAL X ORÇADO COMISSÕES": com}))
    assert conf["status"] == "bloqueada"
    text = "\n".join(conf["bloqueios"])
    assert "BASE" in text and "J8" in text and "COMISSÕES" in text and "L3" in text


def test_orphans_dates_and_warnings_are_alerts():
    sr = SheetReport("RECEITA", 2, 3, 2, orphan_formula_rows=179, date_style_warnings={"B": 2})
    warn = "comissoes: 3 fórmula(s) de linhas mantidas sem valor em cache (ex.: L9)"
    conf = conference(_results(comissoes=_s("comissoes", warnings=[warn])), _build(RECEITA=sr))
    assert conf["status"] == "com_alertas" and not conf["bloqueios"]
    text = "\n".join(conf["alertas"])
    assert "179" in text and "sem valor em cache" in text and "formato de data" in text


def test_informative_fields_are_shown_and_do_not_change_status():
    sr = SheetReport("DESPESAS", 3, 4, 2, overwritten_constants=439, unified_formula_rows=846,
                     autofilter_extended=True)
    conf = conference(_results(), _build(DESPESAS=sr))
    assert conf["status"] == "ok" and not conf["alertas"] and not conf["bloqueios"]
    aba = conf["abas"]["DESPESAS"]
    assert aba["constantes_substituidas"] == 439
    assert aba["formulas_unificadas"] == 846
    assert aba["autofiltro_estendido"] is True
    assert conf["abas"]["RECEITA"]["constantes_substituidas"] == 0
    assert conf["abas"]["RECEITA"]["autofiltro_estendido"] is False
    info = "\n".join(conf["informacoes"])
    assert "439" in info and "846" in info and "autofiltro" in info
    md = render_markdown(conf)
    assert "## Informações" in md and "439" in md and "846" in md
    assert "Constantes substituídas" in md and "Fórmulas unificadas" in md and "Autofiltro" in md
    assert "**Status: ok**" in md


def test_no_informative_fields_means_no_information_section():
    conf = conference(_results(), _build(), na_count=0)
    assert conf["informacoes"] == []
    assert "## Informações" not in render_markdown(conf)


def test_na_count_not_executed_is_informative_only():
    conf = conference(_results(), _build())
    assert conf["status"] == "ok" and not conf["alertas"] and not conf["bloqueios"]
    assert conf["informacoes"] == ["contagem de #N/D não executada"]
    assert "contagem de #N/D não executada" in render_markdown(conf)


def test_na_count_executed_or_failed_has_no_not_executed_note():
    assert "contagem de #N/D não executada" not in conference(_results(), _build(), na_count=0)["informacoes"]
    assert "contagem de #N/D não executada" not in conference(_results(), _build(), na_error="x")["informacoes"]


def test_missing_sheet_in_build_blocks_with_own_message():
    build = _build()
    del build.sheets["RECEITA"]
    conf = conference(_results(), build)
    assert conf["status"] == "bloqueada"
    assert any("aba RECEITA ausente no build" in b for b in conf["bloqueios"])
    assert not any("None" in b for b in conf["bloqueios"])


def test_missing_result_key_is_a_contract_error():
    res = _results()
    del res["cr"]
    with pytest.raises(ContractError, match="cr"):
        conference(res, _build())


def test_na_count_is_an_alert_with_the_count_never_a_block():
    conf = conference(_results(), _build(), na_count=231)
    assert conf["status"] == "com_alertas" and not conf["bloqueios"]
    text = "\n".join(conf["alertas"])
    assert "231" in text and "#N/D" in text and "AP:AQ" in text and "De-Para" in text
    assert "231" in render_markdown(conf)


def test_na_count_zero_is_silent():
    conf = conference(_results(), _build(), na_count=0)
    assert conf["status"] == "ok" and not conf["alertas"]


def test_na_error_is_a_warning_alert_never_a_block():
    conf = conference(_results(), _build(), na_error="Excel indisponível")
    assert conf["status"] == "com_alertas" and not conf["bloqueios"]
    assert any("não foi possível recalcular no Excel para contar #N/D: Excel indisponível" in a
               for a in conf["alertas"])


def test_na_alerts_never_land_in_blocks_even_together():
    conf = conference(_results(), _build(), na_count=3, na_error="falhou")
    assert conf["status"] == "com_alertas" and conf["bloqueios"] == []


def test_contract_error_is_a_blocked_conference():
    conf = contract_blocked("DESPESAS!AP: classificação ausente", meta={"competencia": "2026-09"})
    assert conf["status"] == "bloqueada"
    assert conf["bloqueios"] == ["erro de contrato: DESPESAS!AP: classificação ausente"]
    md = render_markdown(conf)
    assert "bloqueada" in md and "classificação ausente" in md and "2026-09" in md


def test_markdown_mentions_status_and_every_report():
    md = render_markdown(conference(_results(), _build(), meta={"produtora": "untitled", "competencia": "2026-09"}))
    assert "# Conferência da importação" in md and "untitled" in md and "2026-09" in md
    for label in ("RRA", "CP", "CR", "Comissões"):
        assert label in md
    assert "**Status: ok**" in md and "Carregadas" in md
    assert "Controles" not in md


# --- RESULTADO.md ---------------------------------------------------------------------------------

def test_render_resultado_resume_em_portugues_simples():
    from papaki_dre.report import render_resultado
    conf = {"status": "com_alertas", "alertas": ["algo para olhar"], "bloqueios": [],
            "meta": {"produtora": "untitled", "competencia": "2026-09", "periodo": "2026-01-01 a 2026-09-30",
                     "importacao_id": "abc"}}
    texto = render_resultado(conf)
    assert "Status: **com_alertas**" in texto and "foto vigente: sim" in texto
    assert "## Alertas (1)" in texto and "- algo para olhar" in texto and "## Bloqueios (0)" in texto


def test_render_resultado_bloqueada_ou_sem_banco_nao_e_vigente():
    from papaki_dre.report import render_resultado
    base = {"alertas": [], "bloqueios": ["x"], "meta": {"produtora": "untitled", "competencia": "2026-09"}}
    assert "foto vigente: não" in render_resultado({"status": "bloqueada", **base})
    sem_banco = render_resultado({"status": "ok", **base})
    assert "nenhuma (rodada sem banco)" in sem_banco and "foto vigente: não" in sem_banco
