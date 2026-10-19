from pathlib import Path

import pytest

from papaki_dre import caminhos
from papaki_dre.caminhos import CaminhosErro, nomes_pastas


def test_padrao_sem_variaveis():
    assert nomes_pastas({}) == ("extracoes", "apuracoes")
    assert nomes_pastas(None) == ("extracoes", "apuracoes")


def test_nomes_do_ambiente_com_acento_e_espaco():
    env = {"PAPAKI_PASTA_EXTRACOES": "Extrações brutas", "PAPAKI_PASTA_APURACOES": '"Apurações DRE"'}
    assert nomes_pastas(env) == ("Extrações brutas", "Apurações DRE")


@pytest.mark.parametrize("ruim", ["a/b", "a\\b", "..", ".", "a.", "", "x" * 80, "a:b", "con<", "a|b"])
def test_nome_invalido_e_recusado(ruim):
    if ruim == "":
        assert nomes_pastas({"PAPAKI_PASTA_EXTRACOES": ruim})[0] == "extracoes"   # vazio = padrão
        return
    with pytest.raises(CaminhosErro, match="PAPAKI_PASTA_EXTRACOES"):
        nomes_pastas({"PAPAKI_PASTA_EXTRACOES": ruim})


def test_nomes_iguais_sao_recusados():
    with pytest.raises(CaminhosErro, match="diferentes"):
        nomes_pastas({"PAPAKI_PASTA_EXTRACOES": "Dados", "PAPAKI_PASTA_APURACOES": "dados"})


def test_pastas_usam_o_nome_pedido(tmp_path):
    raiz = tmp_path
    assert caminhos.pasta_extracoes(raiz, "2026-09", "untitled", nome="Brutos") == raiz / "Brutos" / "2026-09" / "untitled"
    assert caminhos.pasta_apuracao(raiz, "untitled", "2026-09", nome="DRE") == raiz / "DRE" / "untitled" / "2026-09"
    assert caminhos.rel_extracao("2026-09", "untitled", "20261016-1051", nome="Brutos") == \
        "Brutos/2026-09/untitled/20261016-1051"


def test_padroes_continuam_valendo(tmp_path):
    assert caminhos.pasta_extracoes(tmp_path, "2026-09", "untitled") == tmp_path / "extracoes" / "2026-09" / "untitled"
    assert caminhos.pasta_apuracao(tmp_path, "untitled", "2026-09") == tmp_path / "apuracoes" / "untitled" / "2026-09"
    assert caminhos.rel_extracao("2026-09", "untitled", "x") == "extracoes/2026-09/untitled/x"
