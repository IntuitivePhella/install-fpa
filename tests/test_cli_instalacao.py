import json
from pathlib import Path

import pytest

from papaki_dre import cli


def _rodar(capsys, argv, env=None):
    codigo = cli.main(argv, env=env)
    saida = capsys.readouterr()
    return codigo, saida.out, saida.err


@pytest.fixture
def modelo(tmp_path):
    arq = tmp_path / "Modelo DRE - Untitled - base 2026.xlsx"
    arq.write_bytes(b"x")
    return arq


def test_configurar_grava_so_nomes_e_nao_imprime_valores(capsys, tmp_path, modelo):
    arq_env = tmp_path / ".env"
    dados = tmp_path / "dados da papaki"
    codigo, out, err = _rodar(capsys, ["configurar", "--dados-dir", str(dados), "--criar-pastas",
                                       "--pasta-extracoes", "Brutos", "--pasta-apuracoes", "DRE",
                                       "--modelo", f"untitled={modelo}", "--usuario-esferas", "ana@exemplo.com",
                                       "--arquivo-env", str(arq_env)])
    assert codigo == 0, err
    d = json.loads(out)
    assert d["arquivo_env"] == "criado"
    assert d["gravadas"] == ["PAPAKI_DADOS_DIR", "PAPAKI_PASTA_EXTRACOES", "PAPAKI_PASTA_APURACOES",
                             "PAPAKI_MODELO_UNTITLED", "PAPAKI_USUARIO_ESFERAS"]
    assert d["pastas_criadas"] == ["Brutos", "DRE"]
    assert str(dados) not in out and str(modelo) not in out and "ana@exemplo.com" not in out
    assert (dados / "Brutos").is_dir() and (dados / "DRE").is_dir()
    texto = arq_env.read_text(encoding="utf-8")
    assert f"PAPAKI_MODELO_UNTITLED={modelo}" in texto and "PAPAKI_PASTA_EXTRACOES=Brutos" in texto


def test_configurar_segunda_vez_atualiza(capsys, tmp_path):
    arq_env = tmp_path / ".env"
    _rodar(capsys, ["configurar", "--dados-dir", str(tmp_path / "a"), "--arquivo-env", str(arq_env)])
    codigo, out, _ = _rodar(capsys, ["configurar", "--dados-dir", str(tmp_path / "b"),
                                     "--arquivo-env", str(arq_env)])
    assert codigo == 0 and json.loads(out)["arquivo_env"] == "atualizado"
    assert f"PAPAKI_DADOS_DIR={(tmp_path / 'b').resolve()}" in arq_env.read_text(encoding="utf-8")


@pytest.mark.parametrize("extra, trecho", [
    (["--dados-dir", "relativa\\pasta"], "completo"),
    (["--modelo", "naoexiste=C:\\x.xlsx"], "produtora desconhecida"),
    (["--modelo", "untitled=relativo.xlsx"], "caminho completo"),
    (["--pasta-extracoes", "a/b"], "PAPAKI_PASTA_EXTRACOES"),
    (["--usuario-esferas", "sem-arroba"], "e-mail"),
    ([], "nada a configurar"),
])
def test_configurar_erros_de_uso(capsys, tmp_path, extra, trecho):
    codigo, out, err = _rodar(capsys, ["configurar", *extra, "--arquivo-env", str(tmp_path / ".env")])
    assert codigo == 3 and trecho in err and out == ""
    assert not (tmp_path / ".env").exists()


def test_configurar_recusa_pasta_de_dados_dentro_do_projeto(capsys, tmp_path):
    from papaki_dre import caminhos
    codigo, _, err = _rodar(capsys, ["configurar", "--dados-dir", str(caminhos.RAIZ_PROJETO / "dados"),
                                     "--arquivo-env", str(tmp_path / ".env")])
    assert codigo == 3 and "dentro da pasta do projeto" in err


def test_configurar_recusa_modelo_que_nao_e_xlsx(capsys, tmp_path):
    ruim = tmp_path / "m.txt"
    ruim.write_text("x")
    codigo, _, err = _rodar(capsys, ["configurar", "--modelo", f"untitled={ruim}",
                                     "--arquivo-env", str(tmp_path / ".env")])
    assert codigo == 3 and ".xlsx" in err


def test_produtora_listar_e_mostrar(capsys):
    codigo, out, _ = _rodar(capsys, ["produtora", "listar"])
    assert codigo == 0 and "untitled" in json.loads(out)["produtoras"]
    codigo, out, _ = _rodar(capsys, ["produtora", "mostrar", "--produtora", "untitled"])
    d = json.loads(out)
    assert codigo == 0 and d["localidade"] == "UTLD PRODUÇÃO DE FILMES LTDA"
    assert d["esferas"] == {"consolidado": "papaki.esferas.com.br", "produtora": "untitled.esferas.com.br"}
    assert d["linha_mestre"] == {"cp": {"AP": 881}} and "Impostos" in d["regras"]["comissoes_excluir_h"]


def test_produtora_mostrar_desconhecida(capsys):
    codigo, _, err = _rodar(capsys, ["produtora", "mostrar", "--produtora", "nao_existe"])
    assert codigo == 3 and "desconhecida" in err


def test_instalacao_mostrar_nao_revela_caminhos_nem_chaves(capsys, tmp_path, modelo):
    env = {"PAPAKI_DADOS_DIR": str(tmp_path / "dados"), "PAPAKI_MODELO_UNTITLED": str(modelo),
           "PAPAKI_USUARIO_ESFERAS": "ana@exemplo.com", "SUPABASE_SERVICE_KEY": "segredo-inventado"}
    codigo, out, _ = _rodar(capsys, ["instalacao", "mostrar"], env=env)
    d = json.loads(out)
    assert codigo == 0
    assert d == {"dados_configurado": True, "pasta_extracoes": "extracoes", "pasta_apuracoes": "apuracoes",
                 "usuario_esferas": "ana@exemplo.com", "modelos": {"untitled": True}}
    assert str(tmp_path) not in out and "segredo-inventado" not in out


def test_instalacao_mostrar_sem_nada_configurado(capsys):
    codigo, out, _ = _rodar(capsys, ["instalacao", "mostrar"], env={})
    assert codigo == 0
    assert json.loads(out) == {"dados_configurado": False, "pasta_extracoes": "extracoes",
                               "pasta_apuracoes": "apuracoes", "usuario_esferas": None,
                               "modelos": {"untitled": False}}


def test_configurar_recusa_caminho_com_separador_de_linha_unicode(capsys, tmp_path):
    codigo, out, err = _rodar(capsys, ["configurar", "--dados-dir", str(tmp_path / "a SUPABASE_URL=x"),
                                       "--arquivo-env", str(tmp_path / ".env")])
    assert codigo == 3 and "valor inválido" in err and out == ""
    assert not (tmp_path / ".env").exists()
