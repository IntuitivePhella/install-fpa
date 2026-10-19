import datetime as dt
import json

import pytest

from papaki_dre import calendario as c

D = dt.date


def test_pascoa_de_anos_conhecidos():
    assert c.pascoa(2025) == D(2025, 4, 20)
    assert c.pascoa(2026) == D(2026, 4, 5)
    assert c.pascoa(2027) == D(2027, 3, 28)
    assert c.pascoa(2024) == D(2024, 3, 31)
    assert c.pascoa(2028) == D(2028, 4, 16)
    assert c.pascoa(2030) == D(2030, 4, 21)


def test_sexta_feira_santa_e_feriado():
    assert D(2026, 4, 3) in c.feriados(2026)
    assert not c.eh_dia_util(D(2026, 4, 3))


def test_consciencia_negra_so_a_partir_de_2024():
    assert D(2024, 11, 20) in c.feriados(2024)
    assert D(2023, 11, 20) not in c.feriados(2023)


def test_carnaval_e_corpus_christi_contam_como_dia_util():
    assert c.eh_dia_util(D(2026, 2, 17))
    assert c.eh_dia_util(D(2026, 6, 4))


def test_feriado_extra():
    assert c.eh_dia_util(D(2026, 7, 9))
    assert not c.eh_dia_util(D(2026, 7, 9), extras=("2026-07-09",))


def test_novembro_de_2026_o_segundo_dia_util_e_dia_4():
    # dia 1 é domingo, dia 2 é Finados
    assert c.dias_uteis(2026, 11)[:3] == [D(2026, 11, 3), D(2026, 11, 4), D(2026, 11, 5)]
    assert c.enesimo_dia_util(2026, 11, 2) == D(2026, 11, 4)


def test_janeiro_de_2027_e_a_competencia_de_dezembro():
    assert c.enesimo_dia_util(2027, 1, 1) == D(2027, 1, 4)
    assert c.competencia_alvo(D(2027, 1, 5)) == "2026-12"


def test_competencia_alvo_e_o_mes_anterior():
    assert c.competencia_alvo(D(2026, 11, 4)) == "2026-10"


def test_deve_rodar_no_segundo_dia_util():
    assert c.deve_rodar(D(2026, 11, 4), ja_importada=False) == (True, "")


def test_nao_roda_no_primeiro_dia_util():
    rodar, motivo = c.deve_rodar(D(2026, 11, 3), ja_importada=False)
    assert not rodar and "2º dia útil" in motivo


def test_roda_ate_o_quinto_dia_util_e_depois_nao():
    # novembro/2026: 3, 4, 5, 6, 9 são os cinco primeiros dias úteis
    assert c.deve_rodar(D(2026, 11, 9), ja_importada=False)[0] is True
    rodar, motivo = c.deve_rodar(D(2026, 11, 10), ja_importada=False)
    assert not rodar and "5º dia útil" in motivo


def test_fim_de_semana_e_feriado_nao_rodam():
    assert c.deve_rodar(D(2026, 11, 7), ja_importada=False)[0] is False
    assert c.deve_rodar(D(2026, 11, 2), ja_importada=False)[0] is False


def test_competencia_ja_importada_nao_roda():
    rodar, motivo = c.deve_rodar(D(2026, 11, 4), ja_importada=True)
    assert not rodar and "já importada" in motivo


def test_periodo_da_competencia():
    assert c.periodo_da_competencia("2026-09") == ("2026-09", D(2026, 1, 1), D(2026, 9, 30))
    assert c.periodo_da_competencia("2024-02")[2] == D(2024, 2, 29)
    for ruim in ("2026-13", "2026/09", "setembro", "", None):
        with pytest.raises(ValueError):
            c.periodo_da_competencia(ruim)


def _conf(tmp_path, **kw):
    conf = {"status": "ok", "meta": {"competencia": "2026-09", "periodo": "2026-01-01 a 2026-09-30",
                                       "importacao_id": "abc"}}
    conf.update(kw.pop("topo", {}))
    conf["meta"].update(kw)
    p = tmp_path / "conferencia.json"
    p.write_text(json.dumps(conf), encoding="utf-8")
    return p


def test_competencia_concluida(tmp_path):
    assert c.competencia_concluida(_conf(tmp_path), "2026-09") is True
    assert c.competencia_concluida(_conf(tmp_path, topo={"status": "com_alertas"}), "2026-09") is True


@pytest.mark.parametrize("mudanca", [
    {"topo": {"status": "bloqueada"}},
    {"topo": {"status": "falhou"}},
    {"importacao_id": None},                       # rodada sem banco
    {"competencia": "2026-08"},
    {"periodo": "2026-01-01 a 2026-09-15"},        # demonstrativo parcial
])
def test_competencia_nao_concluida(tmp_path, mudanca):
    assert c.competencia_concluida(_conf(tmp_path, **mudanca), "2026-09") is False


def test_competencia_concluida_sem_arquivo_ou_ilegivel(tmp_path):
    assert c.competencia_concluida(tmp_path / "nao-existe.json", "2026-09") is False
    ruim = tmp_path / "x.json"
    ruim.write_text("{não é json", encoding="utf-8")
    assert c.competencia_concluida(ruim, "2026-09") is False


def test_competencia_concluida_com_meta_malformada(tmp_path):
    for meta in ("x", [1], 5):
        p = tmp_path / "c.json"
        p.write_text(json.dumps({"status": "ok", "meta": meta}), encoding="utf-8")
        assert c.competencia_concluida(p, "2026-09") is False
