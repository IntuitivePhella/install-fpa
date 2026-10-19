import json

import pytest

from papaki_dre import produtoras
from papaki_dre.produtoras import ProdutoraErro, carregar, existe, ids, regras_padrao


def _valido(**mudar):
    d = {"id": "soma", "nome": "Soma", "localidade": "SOMA LTDA",
         "esferas": {"consolidado": "papaki.esferas.com.br", "produtora": "soma.esferas.com.br"},
         "regras": {"nucleo_rra": "Projetos - Soma", "modelo_comissoes": "Soma - Padrão",
                    "comissoes_excluir_h": ["Impostos"]},
         "linha_mestre": {"cp": {"AP": 10}}}
    d.update(mudar)
    return d


def _gravar(pasta, nome, dados):
    pasta.mkdir(exist_ok=True)
    (pasta / f"{nome}.json").write_text(json.dumps(dados, ensure_ascii=False), encoding="utf-8")


def test_untitled_tem_os_valores_que_estavam_no_config():
    p = carregar("untitled")
    assert p.id == "untitled" and p.nome == "Untitled"
    assert p.localidade == "UTLD PRODUÇÃO DE FILMES LTDA"
    assert p.esferas_consolidado == "papaki.esferas.com.br"
    assert p.esferas_produtora == "untitled.esferas.com.br"
    assert p.regras.localidade_cp == "UTLD PRODUÇÃO DE FILMES LTDA"
    assert p.regras.nucleo_rra == "Projetos - Untitled"
    assert p.regras.modelo_comissoes == "Untitled - Publicidade Padrão"
    assert p.regras.comissoes_excluir_h == ("Seguro Produção", "Taxa da produtora", "Comissão total 1",
                                            "Comissão total 2", "Impostos", "Taxa ADM")
    assert p.linha_mestre == {"cp": {"AP": 881}}


def test_regras_padrao_sao_as_da_untitled():
    assert regras_padrao() == carregar("untitled").regras


def test_ids_e_existe():
    assert "untitled" in ids()
    assert existe("untitled")
    for ruim in ("UNTITLED", "../untitled", "untitled.json", "", None, 3, "nao_existe"):
        assert not existe(ruim)


def test_produtora_desconhecida_nao_ecoa_texto_longo():
    with pytest.raises(ProdutoraErro, match="desconhecida"):
        carregar("x" * 200 + "!")


def test_carrega_de_outra_pasta(tmp_path):
    _gravar(tmp_path, "soma", _valido())
    p = carregar("soma", tmp_path)
    assert p.nome == "Soma" and p.linha_mestre == {"cp": {"AP": 10}}
    assert ids(tmp_path) == ["soma"]


def test_linha_mestre_e_opcional(tmp_path):
    d = _valido()
    del d["linha_mestre"]
    _gravar(tmp_path, "soma", d)
    assert carregar("soma", tmp_path).linha_mestre == {}


@pytest.mark.parametrize("mudar, trecho", [
    ({"id": "outra"}, "'id'"),
    ({"nome": ""}, "'nome'"),
    ({"localidade": 3}, "'localidade'"),
    ({"extra": 1}, "desconhecido"),
    ({"esferas": {"consolidado": "a"}}, "'produtora'"),
    ({"esferas": {"consolidado": "a", "produtora": "b", "mais": "c"}}, "desconhecido"),
    ({"regras": {"nucleo_rra": "x", "modelo_comissoes": "y"}}, "comissoes_excluir_h"),
    ({"regras": {"nucleo_rra": "x", "modelo_comissoes": "y", "comissoes_excluir_h": "Impostos"}},
     "comissoes_excluir_h"),
    ({"linha_mestre": {"xx": {"AP": 1}}}, "linha_mestre"),
    ({"linha_mestre": {"cp": {"ap": 1}}}, "linha_mestre"),
    ({"linha_mestre": {"cp": {"AP": 0}}}, "linha_mestre"),
    ({"linha_mestre": {"cp": {"AP": "881"}}}, "linha_mestre"),
])
def test_arquivo_invalido_e_recusado(tmp_path, mudar, trecho):
    _gravar(tmp_path, "soma", _valido(**mudar))
    with pytest.raises(ProdutoraErro, match=trecho):
        carregar("soma", tmp_path)


def test_json_quebrado_e_recusado(tmp_path):
    (tmp_path / "soma.json").write_text("{nao é json", encoding="utf-8")
    with pytest.raises(ProdutoraErro, match="ilegível"):
        carregar("soma", tmp_path)


def test_o_modulo_aponta_para_a_pasta_produtoras_do_projeto():
    assert produtoras.PASTA.name == "produtoras" and (produtoras.PASTA / "untitled.json").is_file()
