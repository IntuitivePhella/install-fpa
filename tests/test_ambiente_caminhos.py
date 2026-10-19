from pathlib import Path

import pytest

from papaki_dre import ambiente, caminhos


def test_resolver_devolve_o_dicionario_dado():
    env = {"A": "1"}
    assert ambiente.resolver(env) is env


def test_ler_env_o_dotenv_prevalece_e_vazio_nao_conta(monkeypatch):
    import dotenv
    monkeypatch.setenv("A", "do-windows")
    monkeypatch.setenv("B", "do-windows")
    monkeypatch.setattr(dotenv, "dotenv_values", lambda *a, **k: {"A": "do-env", "B": "", "C": None})
    env = ambiente.ler_env()
    assert env["A"] == "do-env" and env["B"] == "do-windows" and "C" not in env


def test_raiz_dados_exige_a_variavel():
    with pytest.raises(caminhos.CaminhosErro, match="PAPAKI_DADOS_DIR"):
        caminhos.raiz_dados({})


def test_raiz_dados_nao_pode_ficar_dentro_do_projeto():
    dentro = caminhos.RAIZ_PROJETO / "dados"
    with pytest.raises(caminhos.CaminhosErro, match="dentro da pasta do projeto"):
        caminhos.raiz_dados({"PAPAKI_DADOS_DIR": str(dentro)})
    with pytest.raises(caminhos.CaminhosErro, match="dentro da pasta do projeto"):
        caminhos.raiz_dados({"PAPAKI_DADOS_DIR": str(caminhos.RAIZ_PROJETO)})


def test_raiz_dados_fora_do_projeto(tmp_path):
    assert caminhos.raiz_dados({"PAPAKI_DADOS_DIR": str(tmp_path)}) == tmp_path.resolve()


def test_pastas_padrao(tmp_path):
    raiz = tmp_path
    assert caminhos.pasta_extracoes(raiz, "2026-09", "untitled") == raiz / "extracoes" / "2026-09" / "untitled"
    assert caminhos.pasta_apuracao(raiz, "untitled", "2026-09") == raiz / "apuracoes" / "untitled" / "2026-09"
    assert caminhos.rel_extracao("2026-09", "untitled", "20261104-0930") == \
        "extracoes/2026-09/untitled/20261104-0930"


def test_modelo_da_produtora(tmp_path):
    modelo = tmp_path / "Modelo DRE - Untitled - 2026.xlsx"
    modelo.write_bytes(b"x")
    assert caminhos.modelo_da_produtora("untitled", {"PAPAKI_MODELO_UNTITLED": str(modelo)}) == modelo
    with pytest.raises(caminhos.CaminhosErro, match="PAPAKI_MODELO_UNTITLED"):
        caminhos.modelo_da_produtora("untitled", {})
    with pytest.raises(caminhos.CaminhosErro, match="não existe") as exc:
        caminhos.modelo_da_produtora("untitled", {"PAPAKI_MODELO_UNTITLED": str(tmp_path / "x.xlsx")})
    assert "x.xlsx" not in str(exc.value) and "PAPAKI_MODELO_UNTITLED" in str(exc.value)


def test_raiz_dados_nao_pode_ser_pai_do_projeto():
    with pytest.raises(caminhos.CaminhosErro, match="dentro da pasta do projeto"):
        caminhos.raiz_dados({"PAPAKI_DADOS_DIR": str(caminhos.RAIZ_PROJETO.parent)})
    with pytest.raises(caminhos.CaminhosErro, match="dentro da pasta do projeto"):
        caminhos.raiz_dados({"PAPAKI_DADOS_DIR": Path(caminhos.RAIZ_PROJETO.anchor)})


def test_raiz_dados_recusa_caminho_relativo():
    with pytest.raises(caminhos.CaminhosErro, match="caminho completo"):
        caminhos.raiz_dados({"PAPAKI_DADOS_DIR": "dados"})


def test_modelo_recusa_caminho_relativo(tmp_path):
    with pytest.raises(caminhos.CaminhosErro, match="caminho completo"):
        caminhos.modelo_da_produtora("untitled", {"PAPAKI_MODELO_UNTITLED": "modelo.xlsx"})


def test_raiz_dados_aceita_aspas_e_espacos_nas_pontas(tmp_path):
    assert caminhos.raiz_dados({"PAPAKI_DADOS_DIR": f'  "{tmp_path}"  '}) == tmp_path.resolve()
    assert caminhos.raiz_dados({"PAPAKI_DADOS_DIR": f"'{tmp_path}'"}) == tmp_path.resolve()


def test_modelo_aceita_aspas_nas_pontas(tmp_path):
    modelo = tmp_path / "modelo.xlsx"
    modelo.write_bytes(b"x")
    assert caminhos.modelo_da_produtora("untitled", {"PAPAKI_MODELO_UNTITLED": f'"{modelo}"'}) == modelo


def test_valor_so_com_espacos_conta_como_ausente(monkeypatch):
    import dotenv
    monkeypatch.setenv("PAPAKI_DADOS_DIR", "   ")
    monkeypatch.setattr(dotenv, "dotenv_values", lambda *a, **k: {})
    env = ambiente.ler_env()
    assert "PAPAKI_DADOS_DIR" not in env
    with pytest.raises(caminhos.CaminhosErro, match="PAPAKI_DADOS_DIR"):
        caminhos.raiz_dados({"PAPAKI_DADOS_DIR": "   "})


def test_ler_env_le_o_env_do_projeto_de_qualquer_pasta(tmp_path, monkeypatch):
    projeto = tmp_path / "projeto"
    projeto.mkdir()
    (projeto / ".env").write_text("PAPAKI_TESTE_X=valor-falso\n", encoding="utf-8")
    outra = tmp_path / "outra"
    outra.mkdir()
    monkeypatch.setattr(ambiente, "RAIZ_PROJETO", projeto)
    monkeypatch.chdir(outra)
    monkeypatch.delenv("PAPAKI_TESTE_X", raising=False)
    assert ambiente.ler_env()["PAPAKI_TESTE_X"] == "valor-falso"
