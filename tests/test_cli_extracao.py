import json

import pytest

from extracao_factory import CONTEUDO, mensagem_rra, png, tela
from papaki_dre import cli
from papaki_dre.extracao import Extracao


@pytest.fixture
def env(tmp_path):
    return {"PAPAKI_DADOS_DIR": str(tmp_path / "dados")}


def _rodar(capsys, argv, env):
    codigo = cli.main(argv, env=env)
    saida = capsys.readouterr()
    return codigo, saida.out, saida.err


def _iniciar(capsys, env, *extra):
    codigo, out, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                       "--competencia", "2026-09", *extra], env)
    assert codigo == 0, err
    return json.loads(out)["pasta"]


def _tela_args(rel, pasta, print_, **mudar):
    t = tela(rel, **mudar)
    args = ["extracao", "tela", "--pasta", pasta, "--relatorio", rel, "--localidade", t["localidade"],
            "--criterio", t["criterio"], "--data-inicio", t["data_inicio"], "--data-fim", t["data_fim"],
            "--status", t["status"], "--total", str(t["total"]), "--print", str(print_)]
    if t.get("pesquisar_em"):
        args += ["--pesquisar-em", t["pesquisar_em"]]
    return args


# --- agenda --------------------------------------------------------------------------------------

def test_agenda_no_segundo_dia_util(capsys, env):
    codigo, out, _ = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "2026-11-04"], env)
    assert codigo == 0
    d = json.loads(out)
    assert (d["rodar"], d["competencia"], d["motivo"]) == (True, "2026-10", "")


def test_agenda_nao_roda_no_feriado(capsys, env):
    _, out, _ = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "2026-11-02"], env)
    d = json.loads(out)
    assert (d["rodar"], d["competencia"], d["motivo"]) == (False, "2026-10", "hoje não é dia útil")


def test_agenda_nao_roda_se_a_competencia_ja_foi_importada(capsys, env, tmp_path):
    apuracao = tmp_path / "dados" / "apuracoes" / "untitled" / "2026-10"
    apuracao.mkdir(parents=True)
    (apuracao / "conferencia.json").write_text(json.dumps(
        {"status": "ok", "meta": {"competencia": "2026-10", "periodo": "2026-01-01 a 2026-10-31",
                                  "importacao_id": "abc"}}), encoding="utf-8")
    _, out, _ = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "2026-11-04"], env)
    assert json.loads(out)["rodar"] is False and "já importada" in json.loads(out)["motivo"]


def test_agenda_exige_a_pasta_de_dados(capsys):
    codigo, _, err = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "2026-11-04"], {})
    assert codigo == 3 and "PAPAKI_DADOS_DIR" in err


def test_agenda_data_invalida(capsys, env):
    codigo, _, err = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "ontem"], env)
    assert codigo == 3 and "--hoje" in err and "AAAA-MM-DD" in err


@pytest.mark.parametrize("hoje", ["2026-13-45", "20261104", "2026-1-4", "04/11/2026"])
def test_agenda_hoje_so_no_formato_extenso(capsys, env, hoje):
    codigo, _, err = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", hoje], env)
    assert codigo == 3 and "--hoje" in err and "AAAA-MM-DD" in err


def test_agenda_usa_o_dia_de_brasilia(capsys, env, monkeypatch):
    import datetime as dt
    from papaki_dre import cli_extracao
    from papaki_dre.config import FUSO_BRASILIA
    # 01:00 UTC de 04/11 = 22:00 de 03/11 em Brasilia (1o dia util: nao roda; no dia UTC rodaria)
    monkeypatch.setattr(cli_extracao, "_agora", lambda: dt.datetime(2026, 11, 3, 22, 0, tzinfo=FUSO_BRASILIA))
    _, out, _ = _rodar(capsys, ["agenda", "--produtora", "untitled"], env)
    assert json.loads(out)["rodar"] is False and json.loads(out)["motivo"].startswith("ainda não é o")


def test_agenda_padrao_pede_o_relogio_com_fuso_de_brasilia():
    import datetime as dt
    from papaki_dre import cli_extracao
    assert cli_extracao._agora().utcoffset() == dt.timedelta(hours=-3)


# --- abreviações do argparse e falhas de arquivo ---------------------------------------------------

@pytest.mark.parametrize("flag", ["--nov", "--ret"])
def test_abreviacao_no_iniciar_e_recusada_sem_criar_pasta(capsys, env, tmp_path, flag):
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                     "--competencia", "2026-09", flag], env)
    assert codigo == 3 and "ERRO DE USO" in err
    assert not (tmp_path / "dados").exists()


def test_abreviacao_nao_abandona_a_extracao_aberta(capsys, env):
    pasta = _iniciar(capsys, env)
    codigo, _, _ = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                   "--competencia", "2026-09", "--nov"], env)
    assert codigo == 3
    _, out, _ = _rodar(capsys, ["extracao", "estado", "--pasta", pasta], env)
    assert json.loads(out)["abandonada"] is False


def test_abreviacao_no_agenda_e_recusada(capsys, env):
    codigo, out, err = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoj", "2026-11-04"], env)
    assert codigo == 3 and out == "" and "ERRO DE USO" in err


def test_abreviacao_no_importar_tambem_e_recusada(capsys):
    codigo, _, err = _rodar(capsys, ["importar", "--produtora", "untitled", "--competencia", "2026-09",
                                     "--sem-ban"], {})
    assert codigo == 3 and "ERRO DE USO" in err


def test_total_nao_numerico(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    args = _tela_args("cp", pasta, png(tmp_path))
    args[args.index("--total") + 1] = "abc"
    codigo, _, err = _rodar(capsys, args, env)
    assert codigo == 3 and "ERRO DE USO" in err


def test_pasta_inexistente(capsys, env, tmp_path):
    codigo, _, err = _rodar(capsys, ["extracao", "estado", "--pasta", str(tmp_path / "nao_existe")], env)
    assert codigo == 3 and "ERRO" in err


def test_estado_nao_grava_nada(capsys, env):
    from pathlib import Path
    pasta = _iniciar(capsys, env)
    antes = {p.name: p.read_bytes() for p in Path(pasta).iterdir()}
    assert _rodar(capsys, ["extracao", "estado", "--pasta", pasta], env)[0] == 0
    depois = {p.name: p.read_bytes() for p in Path(pasta).iterdir()}
    assert antes == depois and not any(n.startswith(".trava") for n in depois)
    assert not (Path(pasta).parent / ".trava").exists()


def test_candidatas_com_elementos_nao_dict_nao_escolhem_nada(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    assert _rodar(capsys, _tela_args("rra", pasta, png(tmp_path)), env)[0] == 0
    _rodar(capsys, ["extracao", "rra-ids-antes", "--pasta", pasta, "--ids", ""], env)
    assert _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", "rra"], env)[0] == 0
    cand = tmp_path / "estranho.json"
    cand.write_text(json.dumps([1, "x", None, ["a"]]), encoding="utf-8")
    codigo, _, err = _rodar(capsys, ["extracao", "rra-escolher", "--pasta", pasta, "--candidatas", str(cand)], env)
    assert codigo == 4 and "ainda" in err


def test_erro_de_arquivo_nao_vaza_o_caminho(capsys, env, tmp_path, monkeypatch):
    from papaki_dre import cli_extracao

    def quebra(*a, **k):
        raise PermissionError(13, "negado", str(tmp_path / "segredo" / "estado.json"))
    monkeypatch.setattr(cli_extracao.Extracao, "abrir", staticmethod(quebra))
    codigo, _, err = _rodar(capsys, ["extracao", "estado", "--pasta", str(tmp_path)], env)
    assert codigo == 3 and "PermissionError" in err and "estado.json" in err and "segredo" not in err


# --- extracao: fluxo completo --------------------------------------------------------------------

def test_fluxo_completo_pelo_cli(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    print_ = png(tmp_path)
    downloads = tmp_path / "Downloads"
    downloads.mkdir()
    assert _rodar(capsys, ["extracao", "rra-ids-antes", "--pasta", pasta, "--ids", "velha1,velha2"], env)[0] == 0
    for rel in ("cp", "cr", "rra", "comissoes"):
        assert _rodar(capsys, _tela_args(rel, pasta, print_), env)[0] == 0
        codigo, out, err = _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", rel], env)
        assert codigo == 0, err
        if rel == "rra":
            ex = Extracao.abrir(pasta)
            cand = tmp_path / "cand.json"
            cand.write_text(json.dumps([mensagem_rra(ex, "novo1")]), encoding="utf-8")
            codigo, out, err = _rodar(capsys, ["extracao", "rra-escolher", "--pasta", pasta,
                                                "--candidatas", str(cand)], env)
            assert codigo == 0 and json.loads(out)["mensagem"] == "novo1"
        nome, dados = CONTEUDO[rel]
        (downloads / nome).write_bytes(dados)
        codigo, out, err = _rodar(capsys, ["extracao", "arquivo", "--pasta", pasta, "--relatorio", rel,
                                           "--caminho", str(downloads / nome)], env)
        assert codigo == 0, err
        assert json.loads(out)["arquivo"] == f"{rel}{'.csv' if rel in ('cp', 'cr') else '.xls' if rel == 'rra' else '.xlsx'}"
    codigo, out, err = _rodar(capsys, ["extracao", "fechar", "--pasta", pasta], env)
    assert codigo == 0, err
    assert json.loads(out)["relatorios"] == ["cp", "cr", "rra", "comissoes"]
    codigo, out, _ = _rodar(capsys, ["extracao", "estado", "--pasta", pasta], env)
    assert json.loads(out)["fechada"] is True


def test_saida_e_sempre_ascii_mesmo_com_acento_no_caminho(capsys, tmp_path):
    env = {"PAPAKI_DADOS_DIR": str(tmp_path / "dados é ç")}
    codigo, out, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                       "--competencia", "2026-09"], env)
    assert codigo == 0, err
    assert out.isascii() and "é" in json.loads(out)["pasta"]


def test_tela_invalida_devolve_codigo_3_e_a_mensagem(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    codigo, _, err = _rodar(capsys, _tela_args("cr", pasta, png(tmp_path), localidade="UTLD PRODUCAO"), env)
    assert codigo == 3 and "Localidade" in err and "NÃO exporte" in err


def test_disparar_sem_tela_e_recusado(capsys, env):
    pasta = _iniciar(capsys, env)
    codigo, _, err = _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", "cp"], env)
    assert codigo == 3 and "registre a tela" in err


def test_segundo_disparo_e_recusado(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    assert _rodar(capsys, _tela_args("cp", pasta, png(tmp_path)), env)[0] == 0
    assert _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", "cp"], env)[0] == 0
    codigo, _, err = _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", "cp"], env)
    assert codigo == 3 and "NÃO dispare de novo" in err


def test_rra_escolher_sem_mensagem_devolve_4(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    assert _rodar(capsys, _tela_args("rra", pasta, png(tmp_path)), env)[0] == 0
    _rodar(capsys, ["extracao", "rra-ids-antes", "--pasta", pasta, "--ids", ""], env)
    assert _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", "rra"], env)[0] == 0
    cand = tmp_path / "vazio.json"
    cand.write_text("[]", encoding="utf-8")
    codigo, _, err = _rodar(capsys, ["extracao", "rra-escolher", "--pasta", pasta, "--candidatas", str(cand)], env)
    assert codigo == 4 and "ainda" in err


def test_rra_escolher_candidatas_ilegiveis(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    ruim = tmp_path / "ruim.json"
    ruim.write_text("{não", encoding="utf-8")
    codigo, _, err = _rodar(capsys, ["extracao", "rra-escolher", "--pasta", pasta, "--candidatas", str(ruim)], env)
    assert codigo == 3 and "candidatas ilegível" in err


def test_iniciar_de_novo_pede_retomar_e_estado_mostra_o_que_falta(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                     "--competencia", "2026-09"], env)
    assert codigo == 3 and "--retomar" in err
    assert _iniciar(capsys, env, "--retomar") == pasta
    codigo, out, _ = _rodar(capsys, ["extracao", "estado", "--pasta", pasta], env)
    estado = json.loads(out)
    assert codigo == 0 and estado["relatorios"]["cp"] == {"tela": False, "disparo": False, "arquivo": False}
    assert estado["abandonada"] is False and estado["fechada"] is False


def test_nova_abandona_a_aberta_anterior(capsys, env):
    antiga = _iniciar(capsys, env)
    nova = _iniciar(capsys, env, "--nova")
    assert nova != antiga
    _, out, _ = _rodar(capsys, ["extracao", "estado", "--pasta", antiga], env)
    assert json.loads(out)["abandonada"] is True
    # a abandonada não aceita mais nada
    _, _, err = _rodar(capsys, ["extracao", "rra-ids-antes", "--pasta", antiga, "--ids", "x"], env)
    assert "abandonada" in err


def test_nova_e_recusada_com_disparo_sem_arquivo(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    assert _rodar(capsys, _tela_args("cp", pasta, png(tmp_path)), env)[0] == 0
    assert _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", "cp"], env)[0] == 0
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                     "--competencia", "2026-09", "--nova"], env)
    assert codigo == 3 and "disparo sem arquivo" in err


def test_ajuda_do_nova_diz_que_abandona_em_definitivo(capsys):
    with pytest.raises(SystemExit):
        cli.main(["extracao", "iniciar", "--help"])
    assert "abandona em definitivo" in capsys.readouterr().out.replace("\n", " ")


def test_nao_ha_comando_para_abandonar_nem_redisparar(capsys):
    for acao in ("abandonar", "redisparar"):
        assert cli.main(["extracao", acao]) == 3
        assert "ERRO DE USO" in capsys.readouterr().err


def test_iniciar_periodo_fim_fora_da_competencia(capsys, env):
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled", "--competencia", "2026-09",
                                     "--periodo-fim", "2026-10-05"], env)
    assert codigo == 3 and "dentro da competência" in err


def test_iniciar_sem_pasta_de_dados(capsys):
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                     "--competencia", "2026-09"], {})
    assert codigo == 3 and "PAPAKI_DADOS_DIR" in err


def test_iniciar_competencia_invalida(capsys, env):
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                     "--competencia", "2026-13"], env)
    assert codigo == 3 and "AAAA-MM" in err


# --- caminhos que o Claude precisa (ele não abre o .env) ---------------------------------------------

def test_agenda_devolve_as_pastas_de_apuracao_e_de_extracoes(capsys, env, tmp_path):
    _, out, _ = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "2026-11-04"], env)
    d = json.loads(out)
    raiz = (tmp_path / "dados").resolve()
    assert d["pasta_apuracao"] == str(raiz / "apuracoes" / "untitled" / "2026-10")
    assert d["pasta_extracoes"] == str(raiz / "extracoes" / "2026-10" / "untitled")


def test_agenda_devolve_as_pastas_mesmo_quando_nao_roda(capsys, env):
    _, out, _ = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "2026-11-02"], env)
    d = json.loads(out)
    assert d["rodar"] is False and d["pasta_apuracao"] and d["pasta_extracoes"]


def test_iniciar_devolve_tambem_a_pasta_de_apuracao(capsys, env, tmp_path):
    codigo, out, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                       "--competencia", "2026-09"], env)
    assert codigo == 0, err
    d = json.loads(out)
    assert d["pasta_apuracao"] == str((tmp_path / "dados").resolve() / "apuracoes" / "untitled" / "2026-09")
    assert d["pasta"].startswith(str((tmp_path / "dados").resolve() / "extracoes" / "2026-09" / "untitled"))


def test_retomar_sem_aberta_mas_com_fechada_diz_para_reimportar(capsys, env, tmp_path):
    from extracao_factory import extracao_fechada
    ex = extracao_fechada(tmp_path)
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                     "--competencia", "2026-09", "--retomar"], env)
    assert codigo == 3
    assert "não há extração aberta" in err and f"({ex.pasta.name}) está fechada" in err
    assert "importar --extracao" in err and "não extraia de novo sem o usuário pedir" in err
    assert str(tmp_path) not in err


def test_retomar_sem_nenhuma_extracao_mantem_a_mensagem_simples(capsys, env):
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                     "--competencia", "2026-09", "--retomar"], env)
    assert codigo == 3 and "não há extração aberta" in err and "fechada" not in err
