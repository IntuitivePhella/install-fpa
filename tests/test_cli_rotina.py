"""importar com --extracao, modelo/saída opcionais, fuso e RESULTADO.md. Dublês do test_cli (sanitizadores,
build e contar_nd trocados); nenhum teste toca o banco real nem o Excel, e todos passam `env` explícito ou
--modelo/--saida (nunca dependem do ambiente real da máquina). Dados: só inventados."""
import json

import pytest

import extracao_factory
from extracao_factory import extracao_fechada
from fake_supabase import FakeClient
from papaki_dre import cli
from papaki_dre.db import sha256_file
from papaki_dre.headers import load_headers
from test_cli import OUT_OK, env  # noqa: F401  (fixture `env` e nome do arquivo de saída)

HEADERS = load_headers()
# CSVs com o cabeçalho certo (sem linhas de dados): passam pela conferência prévia de cabeçalho
CP_CSV = b"\xef\xbb\xbf" + ";".join(HEADERS["cp"]).encode("utf-8") + b"\r\n"
CR_CSV = ";".join(HEADERS["cr"]).encode("cp1252") + b"\r\n"


def _argv_ext(env, pasta, *extra, modelo=True, saida=True):
    a = ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--extracao", str(pasta)]
    if modelo:
        a += ["--modelo", str(env["model"])]
    if saida:
        a += ["--saida", str(env["out"])]
    return a + list(extra)


@pytest.fixture
def extracao(env, monkeypatch):
    monkeypatch.setitem(extracao_factory.CONTEUDO, "cp", ("Contas a pagar.csv", CP_CSV))
    monkeypatch.setitem(extracao_factory.CONTEUDO, "cr", ("Contas a receber.csv", CR_CSV))
    return extracao_fechada(env["out"].parent)


def _manifesto(ex):
    return json.loads((ex.pasta / "extracao.json").read_text(encoding="utf-8"))


def _trocar_conteudo(ex, rel, dados: bytes):
    """Troca o conteúdo de um arquivo da extração e acerta hash/tamanho no manifesto (o manifesto passa
    na conferência de integridade; só o conteúdo está errado)."""
    m = _manifesto(ex)
    arq = ex.pasta / m["relatorios"][rel]["arquivo"]
    arq.write_bytes(dados)
    m["relatorios"][rel]["sha256"] = sha256_file(arq)
    m["relatorios"][rel]["tamanho"] = arq.stat().st_size
    (ex.pasta / "extracao.json").write_text(json.dumps(m, ensure_ascii=False), encoding="utf-8")


def test_importar_com_extracao_grava_proveniencia_e_resultado(env, extracao):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta, "--gatilho", "agendado"), client=client, env={}) == 0
    brutos = {r["relatorio"]: r for r in client.tables["arquivos_brutos"]}
    assert set(brutos) == {"cp", "cr", "rra", "comissoes"}
    cp = brutos["cp"]
    assert cp["total_tela"] == 37030
    assert cp["filtros_tela"]["localidade"] == "Todas" and cp["filtros_tela"]["status"] == "Todos"
    assert cp["prova_caminho"] == f"extracoes/2026-09/untitled/{extracao.pasta.name}/prints/cp.png"
    assert brutos["cr"]["filtros_tela"]["localidade"] == "UTLD PRODUÇÃO DE FILMES LTDA"
    # no banco vai o nome ORIGINAL baixado, não o nome fixo da cópia na pasta da extração
    assert {k: r["nome_arquivo"] for k, r in brutos.items()} == {
        "cp": "Contas a pagar.csv", "cr": "Contas a receber.csv", "rra": "Relatorio Rentabilidade.xls",
        "comissoes": "Comissoes.xlsx"}
    m = _manifesto(extracao)
    assert all(r["sha256"] == m["relatorios"][k]["sha256"] for k, r in brutos.items())
    imp = client.tables["importacoes"][0]
    assert imp["gatilho"] == "agendado" and imp["origem"] == "Esferas (Chrome)" and imp["vigente"] is True
    assert imp["data_extracao"] == m["data_extracao"] and imp["data_extracao"].endswith("-03:00")
    resultado = (env["out"] / "RESULTADO.md").read_text(encoding="utf-8")
    assert "Competência: 2026-09" in resultado and "foto vigente: sim" in resultado
    assert imp["id"] in resultado


def test_extracao_passa_pelos_mesmos_sanitizadores_com_os_arquivos_da_pasta(env, extracao, monkeypatch):
    vistos = {}

    def espiao(rel, original):
        def inner(path, **kw):
            vistos[rel] = path
            return original(path, **kw)
        return inner

    for rel in ("cp", "cr", "comissoes", "rra"):
        nome = f"sanitize_{rel}"
        monkeypatch.setattr(cli, nome, espiao(rel, getattr(cli, nome)))
    assert cli.main(_argv_ext(env, extracao.pasta), client=FakeClient(), env={}) == 0
    assert {k: p.name for k, p in vistos.items()} == {
        "cp": "cp.csv", "cr": "cr.csv", "rra": "rra.xls", "comissoes": "comissoes.xlsx"}
    assert all(p.parent == extracao.pasta for p in vistos.values())


def test_a_tela_do_registro_alimenta_a_conferencia(env, extracao):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 0
    conf = json.loads((env["out"] / "conferencia.json").read_text(encoding="utf-8"))
    assert any("tela mostrava 37030" in a for a in conf["alertas"])


@pytest.mark.parametrize("opcao", [("--cp", "x"), ("--cr", "x"), ("--rra", "x"), ("--comissoes", "x"),
                                   ("--tela", "cp=1"), ("--tela", ""),
                                   ("--data-extracao", "2026-10-08T09:00:00"), ("--origem", "outra")])
def test_extracao_nao_combina_com_arquivos_soltos(env, extracao, opcao, capsys):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta, *opcao), client=client, env={}) == 3
    assert "não combina" in capsys.readouterr().err and client.calls == []
    assert not env["out"].exists()


def test_extracao_de_outra_competencia(env, extracao, capsys):
    client = FakeClient()
    argv = _argv_ext(env, extracao.pasta)
    argv[argv.index("--competencia") + 1] = "2026-08"
    assert cli.main(argv, client=client, env={}) == 3
    assert "outra produtora/competência" in capsys.readouterr().err and client.calls == []


def test_extracao_adulterada_nao_grava_nada(env, extracao, capsys):
    nome = _manifesto(extracao)["relatorios"]["cp"]["arquivo"]
    (extracao.pasta / nome).write_bytes(b"\xef\xbb\xbfoutro;conteudo\n")
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 3
    assert "não confere" in capsys.readouterr().err and client.calls == []
    assert not env["out"].exists()


def test_cp_e_cr_trocados_no_manifesto_sao_barrados_antes_do_banco(env, extracao, capsys):
    """Os dois são .csv: o manifesto íntegro (hash acertado) não basta; o cabeçalho é conferido antes
    de qualquer gravação."""
    _trocar_conteudo(extracao, "cp", CR_CSV)
    _trocar_conteudo(extracao, "cr", CP_CSV)
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 3
    err = capsys.readouterr().err
    assert "ERRO DE USO" in err and "cp:" in err and "colunas" in err
    assert client.calls == [] and not env["out"].exists()


def test_cabecalho_errado_so_no_cr_tambem_barra(env, extracao, capsys):
    _trocar_conteudo(extracao, "cr", CR_CSV.replace(b"Cliente", b"Fornecedor", 1))
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 3
    assert "cr: coluna 1" in capsys.readouterr().err and client.calls == []


def test_arquivo_com_assinatura_errada_e_barrado(env, extracao, capsys):
    _trocar_conteudo(extracao, "rra", extracao_factory.ZIP)          # um .xlsx dentro do rra.xls
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 3
    assert "rra" in capsys.readouterr().err and client.calls == []


def test_extracao_nao_fechada(env, tmp_path, capsys):
    pasta = tmp_path / "aberta"
    pasta.mkdir()
    client = FakeClient()
    assert cli.main(_argv_ext(env, pasta), client=client, env={}) == 3
    assert "não está fechada" in capsys.readouterr().err and client.calls == []


def test_periodo_fim_diferente_do_registro_e_erro(env, extracao, capsys):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta, "--periodo-fim", "2026-09-15"), client=client, env={}) == 3
    assert "período" in capsys.readouterr().err and client.calls == []


def test_periodo_fim_igual_ao_registro_e_aceito(env, extracao):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta, "--periodo-fim", "2026-09-30"), client=client, env={}) == 0
    assert client.tables["importacoes"][0]["periodo_fim"] == "2026-09-30"


def test_periodo_fim_invalido_com_extracao(env, extracao, capsys):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta, "--periodo-fim", "30/09/2026"), client=client, env={}) == 3
    assert "AAAA-MM-DD" in capsys.readouterr().err and client.calls == []


def test_sem_extracao_faltam_arquivos(env, capsys):
    argv = ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--modelo", str(env["model"]),
            "--saida", str(env["out"]), "--sem-banco"]
    assert cli.main(argv, env={}) == 3
    assert "faltam: --rra, --cp, --cr, --comissoes" in capsys.readouterr().err


def test_modelo_e_saida_vem_do_ambiente(env, extracao, tmp_path):
    dados = tmp_path / "dados"
    envs = {"PAPAKI_DADOS_DIR": str(dados), "PAPAKI_MODELO_UNTITLED": str(env["model"])}
    argv = _argv_ext(env, extracao.pasta, "--sem-banco", modelo=False, saida=False)
    assert cli.main(argv, env=envs) == 0
    pasta = dados / "apuracoes" / "untitled" / "2026-09"
    assert (pasta / OUT_OK).is_file()
    assert "nenhuma (rodada sem banco)" in (pasta / "RESULTADO.md").read_text(encoding="utf-8")


def test_sem_modelo_nem_ambiente_e_erro_de_uso(env, extracao, capsys):
    argv = _argv_ext(env, extracao.pasta, "--sem-banco", modelo=False)
    assert cli.main(argv, env={}) == 3
    assert "PAPAKI_MODELO_UNTITLED" in capsys.readouterr().err


def test_sem_saida_nem_ambiente_e_erro_de_uso(env, extracao, capsys):
    argv = _argv_ext(env, extracao.pasta, "--sem-banco", saida=False)
    assert cli.main(argv, env={}) == 3
    assert "PAPAKI_DADOS_DIR" in capsys.readouterr().err


def test_raiz_de_dados_dentro_do_projeto_e_recusada(env, extracao, capsys):
    from papaki_dre.caminhos import RAIZ_PROJETO
    argv = _argv_ext(env, extracao.pasta, "--sem-banco", saida=False)
    assert cli.main(argv, env={"PAPAKI_DADOS_DIR": str(RAIZ_PROJETO / "dados"),
                               "PAPAKI_MODELO_UNTITLED": str(env["model"])}) == 3
    assert "dentro da pasta do projeto" in capsys.readouterr().err
    assert not (RAIZ_PROJETO / "dados").exists()


def test_arquivos_soltos_sem_resultado_md(env):
    f = env["files"]
    argv = ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--cp", str(f["cp"]),
            "--cr", str(f["cr"]), "--rra", str(f["rra"]), "--comissoes", str(f["comissoes"]),
            "--modelo", str(env["model"]), "--saida", str(env["out"])]
    client = FakeClient()
    assert cli.main(argv, client=client, env={}) == 0
    assert not (env["out"] / "RESULTADO.md").exists()
    imp = client.tables["importacoes"][0]
    assert imp["origem"] == "arquivos locais"
    assert imp["data_extracao"].endswith("+00:00")                      # sem data: agora, com fuso
    assert {r["nome_arquivo"] for r in client.tables["arquivos_brutos"]} == {"cp.dat", "cr.dat", "rra.dat",
                                                                               "comissoes.dat"}
    assert all(r["filtros_tela"] is None and r["prova_caminho"] is None for r in client.tables["arquivos_brutos"])


def test_data_extracao_sem_fuso_vale_brasilia(env):
    f = env["files"]
    argv = ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--cp", str(f["cp"]),
            "--cr", str(f["cr"]), "--rra", str(f["rra"]), "--comissoes", str(f["comissoes"]),
            "--modelo", str(env["model"]), "--saida", str(env["out"]), "--data-extracao", "2026-10-08T09:00:00"]
    client = FakeClient()
    assert cli.main(argv, client=client, env={}) == 0
    assert client.tables["importacoes"][0]["data_extracao"] == "2026-10-08T09:00:00-03:00"


def test_resultado_md_na_rodada_bloqueada(env, extracao):
    env["state"]["missing"] = ["AP3"]
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 2
    texto = (env["out"] / "RESULTADO.md").read_text(encoding="utf-8")
    assert "Status: **bloqueada**" in texto and "foto vigente: não" in texto


def test_resultado_md_no_bloqueio_por_contrato(env, extracao, monkeypatch):
    from papaki_dre.types import ContractError

    def falha(*a, **kw):
        raise ContractError("cp: coluna 3 é 'X'; esperado 'Y'")
    monkeypatch.setattr(cli, "sanitize_cp", falha)
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 2
    assert "Status: **bloqueada**" in (env["out"] / "RESULTADO.md").read_text(encoding="utf-8")


def test_falha_ao_gravar_resultado_nao_vira_falha(env, extracao, monkeypatch, capsys):
    def quebra(conf):
        raise RuntimeError("disco cheio")
    monkeypatch.setattr(cli, "render_resultado", quebra)
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 0
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "concluida" and imp["vigente"] is True
    assert "RESULTADO.md" in capsys.readouterr().err
    assert json.loads((env["out"] / "conferencia.json").read_text(encoding="utf-8"))["status"] != "falhou"


# --- revisão: RESULTADO.md coerente com a pasta, reimportação, integridade e menores ---------------

def test_promover_falha_e_o_resultado_antigo_nao_diz_vigente(env, extracao, monkeypatch):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 0
    assert "foto vigente: sim" in (env["out"] / "RESULTADO.md").read_text(encoding="utf-8")
    from papaki_dre.db import Repository

    def cai(self, imp_id):
        raise RuntimeError("rede caiu")
    monkeypatch.setattr(Repository, "promover", cai)
    assert cli.main(_argv_ext(env, extracao.pasta, "--forcar"), client=client, env={}) == 3
    conf = json.loads((env["out"] / "conferencia.json").read_text(encoding="utf-8"))
    assert conf["status"] == "falhou"
    resultado = env["out"] / "RESULTADO.md"
    assert not resultado.exists() or "foto vigente: sim" not in resultado.read_text(encoding="utf-8")


def test_rodada_solta_bloqueada_sobre_pasta_com_resultado_antigo(env, extracao):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 0
    assert (env["out"] / "RESULTADO.md").exists()
    f = env["files"]
    argv = ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--cp", str(f["cp"]),
            "--cr", str(f["cr"]), "--rra", str(f["rra"]), "--comissoes", str(f["comissoes"]),
            "--modelo", str(env["model"]), "--saida", str(env["out"])]
    env["state"]["missing"] = ["AP3"]
    assert cli.main(argv, client=client, env={}) == 2
    resultado = env["out"] / "RESULTADO.md"
    assert not resultado.exists() or "foto vigente: sim" not in resultado.read_text(encoding="utf-8")


def test_resultado_so_e_escrito_depois_de_promover(env, extracao, monkeypatch):
    client = FakeClient()
    original = cli.render_resultado
    momento = []

    def espiao(conf):
        momento.append(("rpc", "promover_importacao", None) in client.calls)
        return original(conf)
    monkeypatch.setattr(cli, "render_resultado", espiao)
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 0
    assert momento == [True]


def _outra_extracao(env, monkeypatch):
    monkeypatch.setitem(extracao_factory.CONTEUDO, "cp", ("Contas a pagar.csv", CP_CSV + b"\r\n"))
    return extracao_fechada(env["out"].parent / "outra")


def test_reimportar_extracao_que_nao_e_mais_a_vigente_avisa(env, extracao, monkeypatch, capsys):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 0
    id_a = client.tables["importacoes"][0]["id"]
    capsys.readouterr()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 0
    out = capsys.readouterr().out
    assert "idênticos já importados" in out and "é a foto vigente" in out and "NÃO" not in out
    b = _outra_extracao(env, monkeypatch)
    assert cli.main(_argv_ext(env, b.pasta), client=client, env={}) == 0
    assert f"substituiu a importação {id_a} como foto vigente" in capsys.readouterr().out
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 0
    out = capsys.readouterr().out
    assert "idênticos já importados" in out and "esta NÃO é a foto vigente" in out
    assert "use --forcar para voltar a ela" in out
    assert len(client.tables["importacoes"]) == 2


def test_primeira_importacao_nao_diz_que_substituiu(env, extracao, capsys):
    assert cli.main(_argv_ext(env, extracao.pasta), client=FakeClient(), env={}) == 0
    assert "substituiu" not in capsys.readouterr().out


def test_arquivo_muda_entre_o_manifesto_e_o_hash(env, extracao, monkeypatch, capsys):
    original = cli._conferir_brutos_da_extracao

    def e_muda(files):
        original(files)
        files["cp"].write_bytes(CP_CSV + b"\r\n")
    monkeypatch.setattr(cli, "_conferir_brutos_da_extracao", e_muda)
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 3
    assert "mudou" in capsys.readouterr().err and client.calls == []
    assert not env["out"].exists()


def test_extracao_fora_do_lugar_padrao_grava_prova_nula_e_avisa(env, extracao, tmp_path, capsys):
    import shutil
    outra = tmp_path / "pendrive" / "qualquer"
    shutil.copytree(extracao.pasta, outra)
    client = FakeClient()
    assert cli.main(_argv_ext(env, outra), client=client, env={}) == 0
    assert all(r["prova_caminho"] is None for r in client.tables["arquivos_brutos"])
    assert all(r["filtros_tela"] for r in client.tables["arquivos_brutos"])
    assert "prova_caminho" in capsys.readouterr().err


@pytest.mark.parametrize("ruim", [".", "..", "", "   ", None])
def test_nome_original_invalido_cai_no_nome_fixo(env, extracao, ruim):
    m = _manifesto(extracao)
    m["relatorios"]["cp"]["arquivo_original"] = ruim
    (extracao.pasta / "extracao.json").write_text(json.dumps(m, ensure_ascii=False), encoding="utf-8")
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client, env={}) == 0
    brutos = {r["relatorio"]: r for r in client.tables["arquivos_brutos"]}
    assert brutos["cp"]["nome_arquivo"] == "cp.csv"


def test_arquivo_nao_encontrado_sem_caminho_completo(env, tmp_path, capsys):
    f = env["files"]
    argv = ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--cp", str(tmp_path / "sumiu.csv"),
            "--cr", str(f["cr"]), "--rra", str(f["rra"]), "--comissoes", str(f["comissoes"]),
            "--modelo", str(env["model"]), "--saida", str(env["out"])]
    client = FakeClient()
    assert cli.main(argv, client=client, env={}) == 3
    err = capsys.readouterr().err
    assert "sumiu.csv" in err and str(tmp_path) not in err and client.calls == []


# --- Task 8: importar --restaurar-banco -----------------------------------------------------------
# AdminFalso no lugar da API de gerenciamento; nenhum teste fala com o Supabase nem lê o .env real.

from papaki_dre.supabase_admin import AdminErro  # noqa: E402

TOKEN_FALSO = "sbp_falso_0123456789"
URL_FALSA = "https://refinventado.supabase.co"


REF_DUBLE = "refdoadminfalso"
TOKEN_DUBLE = "sbp_tokendoadminfalso"


class AdminFalso:
    """Acordar 'liga' o cliente falso (que estava pausado), como a restauração real faria. Tem `ref` e
    `token` reconhecíveis: nenhum deles pode aparecer na saída."""

    def __init__(self, client, resultado="restaurado", erro=None, liga=True):
        self.client, self.resultado, self.erro, self.liga, self.chamadas = client, resultado, erro, liga, 0
        self.ref, self.token, self.kwargs = REF_DUBLE, TOKEN_DUBLE, None

    def acordar(self, *a, **k):
        self.chamadas += 1
        self.kwargs = k
        if self.erro:
            raise self.erro
        if self.liga:
            self.client.fail_with = None
        return self.resultado


@pytest.fixture
def sem_espera(monkeypatch):
    esperas = []
    monkeypatch.setattr(cli, "_dormir", esperas.append)
    return esperas


def _com_extracao(env, extracao, *extra):
    return _argv_ext(env, extracao.pasta, *extra)


def _sem_segredos(saida, *extra):
    texto = saida.out + saida.err
    assert all(s not in texto for s in (REF_DUBLE, TOKEN_DUBLE, TOKEN_FALSO, "refinventado", *extra))


def test_banco_pausado_com_restaurar_acorda_e_importa(env, extracao, sem_espera):
    client = FakeClient(fail_with=ConnectionError("project paused"))
    admin = AdminFalso(client)
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin, env={}) == 0
    assert admin.chamadas == 1 and client.tables["importacoes"][0]["vigente"] is True


def test_sem_a_opcao_o_banco_pausado_continua_so_avisando(env, extracao, capsys):
    client = FakeClient(fail_with=ConnectionError("project paused"))
    admin = AdminFalso(client)
    assert cli.main(_com_extracao(env, extracao), client=client, admin=admin, env={}) == 3
    assert admin.chamadas == 0 and "banco pausado: restaure o projeto e rode de novo" in capsys.readouterr().err
    assert "importacoes" not in client.tables


def test_restaurar_sem_chave_e_erro_de_uso_antes_de_tudo(env, extracao, capsys):
    client = FakeClient()
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, env={}) == 3
    err = capsys.readouterr().err
    assert "ERRO DE USO" in err and "SUPABASE_ACCESS_TOKEN" in err and client.calls == []
    assert not env["out"].exists()


def test_restaurar_sem_url_e_erro_de_uso_sem_mostrar_o_token(env, extracao, capsys):
    client = FakeClient()
    argv = _com_extracao(env, extracao, "--restaurar-banco")
    assert cli.main(argv, client=client, env={"SUPABASE_ACCESS_TOKEN": TOKEN_FALSO}) == 3
    saida = capsys.readouterr()
    assert "SUPABASE_URL" in saida.err and TOKEN_FALSO not in saida.err + saida.out and client.calls == []


def test_restaurar_com_url_invalida_nao_mostra_a_url(env, extracao, capsys):
    client = FakeClient()
    envs = {"SUPABASE_ACCESS_TOKEN": TOKEN_FALSO, "SUPABASE_URL": "https://outro-host.exemplo.com"}
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, env=envs) == 3
    err = capsys.readouterr().err
    assert "SUPABASE_URL inválida" in err and "outro-host" not in err and client.calls == []


def test_sem_banco_nao_combina_com_restaurar(env, extracao, capsys):
    argv = _com_extracao(env, extracao, "--restaurar-banco", "--sem-banco")
    assert cli.main(argv, env={}) == 3
    err = capsys.readouterr().err
    assert "--sem-banco" in err and "--restaurar-banco" in err and not env["out"].exists()


def test_falha_ao_acordar_vira_codigo_3_com_aviso(env, extracao, sem_espera, capsys):
    client = FakeClient(fail_with=ConnectionError("project paused"))
    admin = AdminFalso(client, erro=AdminErro("o projeto não ficou ativo em 10 min (último estado: RESTORING)"))
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin, env={}) == 3
    saida = capsys.readouterr()
    err = saida.err
    assert "não ficou ativo" in err and "banco pausado" in err and "importacoes" not in client.tables
    assert admin.chamadas == 1
    _sem_segredos(saida)


def test_ativo_mas_banco_ainda_nao_responde_desiste_sem_esperar(env, extracao, sem_espera, capsys):
    """A API diz que o projeto já estava ativo: a falha não é pausa. Nada é restaurado, não há espera
    longa e o aviso atual continua (código 3)."""
    client = FakeClient(fail_with=ConnectionError("fora do ar"))
    admin = AdminFalso(client, resultado="ja_ativo", liga=False)
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin, env={}) == 3
    saida = capsys.readouterr()
    err = saida.err
    assert admin.chamadas == 1 and "banco pausado" in err and "não é pausa" in err
    _sem_segredos(saida)
    assert sem_espera == [] and "importacoes" not in client.tables


def test_restaurado_mas_banco_nao_volta_desiste_depois_das_tentativas(env, extracao, sem_espera, capsys):
    client = FakeClient(fail_with=ConnectionError("ainda subindo"))
    admin = AdminFalso(client, resultado="restaurado", liga=False)
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin, env={}) == 3
    saida = capsys.readouterr()
    err = saida.err
    _sem_segredos(saida)
    assert admin.chamadas == 1                                     # restaura no máximo uma vez
    assert admin.kwargs["espera_max"] == cli.ESPERA_ACORDAR < cli.PRAZO_RESTAURACAO
    assert sem_espera == [cli.INTERVALO_APOS_ACORDAR] * (cli.TENTATIVAS_APOS_ACORDAR - 1)
    assert "banco inacessível" in err and "importacoes" not in client.tables


def test_banco_volta_depois_de_algumas_tentativas(env, extracao, monkeypatch):
    client = FakeClient(fail_with=ConnectionError("project paused"))
    admin = AdminFalso(client, resultado="aguardado", liga=False)
    esperas = []

    def liga_na_terceira(s):
        esperas.append(s)
        if len(esperas) == 2:
            client.fail_with = None
    monkeypatch.setattr(cli, "_dormir", liga_na_terceira)
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin, env={}) == 0
    assert admin.chamadas == 1 and len(esperas) == 2
    assert client.tables["importacoes"][0]["vigente"] is True


def test_banco_acessivel_com_restaurar_nem_cria_o_admin(env, extracao, monkeypatch):
    criados = []
    monkeypatch.setattr(cli, "AdminSupabase", lambda *a, **k: criados.append(a))
    envs = {"SUPABASE_ACCESS_TOKEN": TOKEN_FALSO, "SUPABASE_URL": URL_FALSA}
    client = FakeClient()
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, env=envs) == 0
    assert criados == []


def test_restaurar_com_token_no_ambiente_cria_o_admin_real(env, extracao, monkeypatch, capsys):
    criados = {}
    client = FakeClient(fail_with=ConnectionError("project paused"))
    client.supabase_url = URL_FALSA + "/"          # o cliente real (supabase-py) guarda a URL assim

    class AdminCriado(AdminFalso):
        def __init__(self, token, ref):
            criados["token"], criados["ref"] = token, ref
            super().__init__(client)

        def fechar(self):
            criados["fechado"] = True

    monkeypatch.setattr(cli, "AdminSupabase", AdminCriado)
    monkeypatch.setattr(cli, "_dormir", lambda s: None)
    envs = {"SUPABASE_ACCESS_TOKEN": TOKEN_FALSO, "SUPABASE_URL": URL_FALSA}
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, env=envs) == 0
    assert criados == {"token": TOKEN_FALSO, "ref": "refinventado", "fechado": True}
    saida = capsys.readouterr()
    assert all(s not in saida.err + saida.out for s in (TOKEN_FALSO, "refinventado", URL_FALSA))


# --- revisão da Task 8: mesma fonte do .env, mesmo projeto, prazo global e __context__ --------------

class _RelogioFalso:
    def __init__(self):
        self.agora = 1000.0

    def __call__(self):
        return self.agora


def test_sem_a_opcao_com_token_no_ambiente_nenhum_admin_e_criado(env, extracao, monkeypatch, capsys):
    criados = []
    monkeypatch.setattr(cli, "AdminSupabase", lambda *a, **k: criados.append(a))
    envs = {"SUPABASE_ACCESS_TOKEN": TOKEN_FALSO, "SUPABASE_URL": URL_FALSA}
    client = FakeClient(fail_with=ConnectionError("project paused"))
    client.supabase_url = URL_FALSA + "/"          # mesmo projeto: só a falta da opção impede restaurar
    assert cli.main(_com_extracao(env, extracao), client=client, env=envs) == 3
    assert criados == [] and "banco pausado" in capsys.readouterr().err


def test_restaurar_com_url_mas_sem_token_e_erro_de_uso(env, extracao, capsys):
    client = FakeClient()
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client,
                    env={"SUPABASE_URL": URL_FALSA}) == 3
    saida = capsys.readouterr()
    assert "SUPABASE_ACCESS_TOKEN" in saida.err and "SUPABASE_URL" not in saida.err.replace(
        "SUPABASE_ACCESS_TOKEN", "") and client.calls == []
    _sem_segredos(saida, URL_FALSA)


@pytest.mark.parametrize("url_cliente", ["https://outroprojeto.supabase.co/", None])
def test_cliente_de_outro_projeto_nao_restaura(env, extracao, monkeypatch, capsys, url_cliente):
    criados = []
    monkeypatch.setattr(cli, "AdminSupabase", lambda *a, **k: criados.append(a))
    client = FakeClient(fail_with=ConnectionError("project paused"))
    if url_cliente:
        client.supabase_url = url_cliente
    envs = {"SUPABASE_ACCESS_TOKEN": TOKEN_FALSO, "SUPABASE_URL": URL_FALSA}
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, env=envs) == 3
    saida = capsys.readouterr()
    assert criados == [] and "mesmo projeto" in saida.err and "importacoes" not in client.tables
    _sem_segredos(saida, "outroprojeto")


def test_sem_cliente_injetado_o_banco_usa_o_mesmo_ambiente_da_restauracao(env, extracao, monkeypatch):
    vistos = []
    client = FakeClient()
    monkeypatch.setattr(cli, "make_client", lambda e=None: vistos.append(e) or client)
    envs = {"SUPABASE_ACCESS_TOKEN": TOKEN_FALSO, "SUPABASE_URL": URL_FALSA}
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), env=envs) == 0
    assert vistos == [envs]


def test_prazo_global_da_restauracao(env, extracao, monkeypatch, capsys):
    """Relógio falso: o acordar consome quase todo o prazo; as tentativas seguintes param no prazo global
    (sem espera real) e a mensagem é a fixa."""
    relogio = _RelogioFalso()
    esperas = []

    def dorme(s):
        esperas.append(s)
        relogio.agora += s

    monkeypatch.setattr(cli, "_relogio", relogio)
    monkeypatch.setattr(cli, "_dormir", dorme)
    tentativas = []

    class Contando(FakeClient):
        def table(self, name):
            tentativas.append(relogio.agora)
            return super().table(name)

    client = Contando(fail_with=ConnectionError("ainda subindo"))

    class Demorado(AdminFalso):
        def acordar(self, *a, **k):
            relogio.agora += cli.PRAZO_RESTAURACAO - 10          # 10 s antes do prazo
            return super().acordar(*a, **k)

    admin = Demorado(client, resultado="restaurado", liga=False)
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin, env={}) == 3
    saida = capsys.readouterr()
    assert f"o banco não respondeu em {cli.PRAZO_RESTAURACAO // 60} min; rode de novo" in saida.err
    assert "nada foi gravado" in saida.err and "importacoes" not in client.tables
    assert esperas == [6, 4]
    # 1 antes de acordar + 2 depois; nenhuma tentativa começa com o prazo já esgotado
    inicio = 1000.0
    assert len(tentativas) == 3 and all(t - inicio < cli.PRAZO_RESTAURACAO for t in tentativas)
    assert admin.chamadas == 1
    assert cli.PRAZO_RESTAURACAO <= 9 * 60 and cli.ESPERA_ACORDAR < cli.PRAZO_RESTAURACAO
    _sem_segredos(saida)


def test_prazo_estourado_dentro_do_acordar_testa_uma_vez_e_desiste(env, extracao, monkeypatch, capsys):
    relogio = _RelogioFalso()
    monkeypatch.setattr(cli, "_relogio", relogio)
    monkeypatch.setattr(cli, "_dormir", lambda s: pytest.fail("não espera depois do prazo"))
    client = FakeClient(fail_with=ConnectionError("ainda subindo"))

    class Estourou(AdminFalso):
        def acordar(self, *a, **k):
            relogio.agora += cli.PRAZO_RESTAURACAO + 1
            return super().acordar(*a, **k)

    admin = Estourou(client, resultado="restaurado", liga=False)
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin, env={}) == 3
    assert "o banco não respondeu em" in capsys.readouterr().err


def test_banco_indisponivel_do_garantir_banco_sem_context():
    from papaki_dre.db import BancoIndisponivel, Repository
    repo = Repository(FakeClient(fail_with=ConnectionError("https://refinventado.supabase.co apikey=x")))
    with pytest.raises(BancoIndisponivel) as exc:
        cli._garantir_banco(repo, AdminFalso(repo.client, erro=AdminErro("falhou")))
    assert exc.value.__context__ is None and exc.value.__cause__ is None
    with pytest.raises(BancoIndisponivel) as exc:
        cli._garantir_banco(repo)
    assert exc.value.__context__ is None


def test_url_invalida_na_restauracao_sem_context():
    with pytest.raises(cli.UsageError) as exc:
        cli._fabrica_admin({"SUPABASE_ACCESS_TOKEN": TOKEN_FALSO, "SUPABASE_URL": "https://x.exemplo.com"})
    assert exc.value.__context__ is None and "exemplo" not in str(exc.value)


# --- Task 9: rodadas órfãs -----------------------------------------------------------------------

import datetime as dt  # noqa: E402


def _orfa(client, id_, **kw):
    linha = {"id": id_, "produtora_id": "untitled", "competencia": "2026-09", "status": "em_andamento",
             "vigente": False, "gatilho": "manual", "criado_em": "2026-01-01T00:00:00+00:00"}
    linha.update(kw)
    client.tables.setdefault("importacoes", []).append(linha)


@pytest.fixture
def relogio_orfas(monkeypatch):
    monkeypatch.setattr(cli, "_agora", lambda: dt.datetime(2026, 11, 4, 12, 0, tzinfo=dt.timezone.utc))


def test_importar_encerra_orfas_antigas_da_mesma_competencia(env, extracao, relogio_orfas):
    client = FakeClient()
    _orfa(client, "orfa")
    _orfa(client, "recente", criado_em="2026-11-04T11:00:00+00:00")
    _orfa(client, "outra-comp", competencia="2026-08")
    _orfa(client, "outra-prod", produtora_id="outra")
    assert cli.main(_argv_ext(env, extracao.pasta), client=client) == 0
    por_id = {r["id"]: r for r in client.tables["importacoes"]}
    assert por_id["orfa"]["status"] == "falhou" and "interrompida" in por_id["orfa"]["motivo_falha"]
    assert [por_id[k]["status"] for k in ("recente", "outra-comp", "outra-prod")] == ["em_andamento"] * 3
    assert len([r for r in client.tables["importacoes"] if r["vigente"]]) == 1


def test_arquivos_identicos_ja_importados_nao_mexem_em_nada(env, extracao, relogio_orfas):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client) == 0
    _orfa(client, "orfa")                                   # duplicado: a rodada não acontece
    antes = [dict(r) for r in client.tables["importacoes"]]
    assert cli.main(_argv_ext(env, extracao.pasta), client=client) == 0
    assert client.tables["importacoes"] == antes


def test_falha_na_limpeza_de_orfas_nao_derruba_a_rodada(env, extracao, monkeypatch, capsys):
    monkeypatch.setenv("SUPABASE_URL", "https://refinventado.supabase.co")
    monkeypatch.setenv("SUPABASE_SERVICE_KEY", "segredo")

    def quebra(self, *a, **k):
        raise ConnectionError("https://refinventado.supabase.co apikey=segredo")

    monkeypatch.setattr(cli.Repository, "encerrar_orfas", quebra)
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client) == 0
    err = capsys.readouterr().err
    assert "AVISO" in err and "rodadas interrompidas" in err
    assert "refinventado" not in err and "segredo" not in err
    assert len([r for r in client.tables["importacoes"] if r["vigente"]]) == 1


def test_sem_banco_nao_toca_em_orfas(env, extracao, monkeypatch):
    monkeypatch.setattr(cli.Repository, "encerrar_orfas",
                        lambda *a, **k: pytest.fail("não deve limpar sem banco"))
    assert cli.main(_argv_ext(env, extracao.pasta, "--sem-banco"), client=FakeClient()) == 0


# --- pasta de apuração na saída, caminhos fora das mensagens, prazo da restauração ---------------------

def test_importar_imprime_a_pasta_de_apuracao_uma_vez_em_linha_propria(env, extracao, capsys):
    assert cli.main(_com_extracao(env, extracao), client=FakeClient(), env={}) == 0
    linhas = [l for l in capsys.readouterr().out.splitlines() if l.startswith("Pasta de apuracao: ")]
    assert linhas == [f"Pasta de apuracao: {env['out']}"]


def test_importar_com_arquivos_identicos_tambem_imprime_a_pasta(env, extracao, capsys):
    client = FakeClient()
    assert cli.main(_com_extracao(env, extracao), client=client, env={}) == 0
    capsys.readouterr()
    assert cli.main(_com_extracao(env, extracao), client=client, env={}) == 0
    saida = capsys.readouterr().out
    assert "Arquivos idênticos já importados" in saida and saida.count("Pasta de apuracao: ") == 1


def test_status_final_cita_so_o_nome_do_arquivo(env, extracao, capsys):
    assert cli.main(_com_extracao(env, extracao), client=FakeClient(), env={}) == 0
    status = [l for l in capsys.readouterr().out.splitlines() if l.startswith("Status:")][0]
    assert status.endswith(OUT_OK) and str(env["out"]) not in status


def test_erro_de_saida_de_rodada_real_cita_so_o_nome_da_pasta(env, extracao, capsys):
    assert cli.main(_com_extracao(env, extracao), client=FakeClient(), env={}) == 0
    capsys.readouterr()
    assert cli.main(_com_extracao(env, extracao, "--sem-banco"), env={}) == 3
    err = capsys.readouterr().err
    assert "use outra pasta em --saida" in err and str(env["out"]) not in err


def test_espera_de_acordar_e_um_minuto_menor_que_o_prazo_global():
    assert cli.PRAZO_RESTAURACAO == 540 and cli.ESPERA_ACORDAR == cli.PRAZO_RESTAURACAO - 60 == 480


def test_restauracao_de_oito_minutos_ainda_funciona(env, extracao, monkeypatch):
    relogio = _RelogioFalso()
    monkeypatch.setattr(cli, "_relogio", relogio)
    monkeypatch.setattr(cli, "_dormir", lambda s: None)
    client = FakeClient(fail_with=ConnectionError("project paused"))

    class OitoMinutos(AdminFalso):
        def acordar(self, *a, **k):
            assert k["espera_max"] == 480
            relogio.agora += 480
            return super().acordar(*a, **k)

    admin = OitoMinutos(client)
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin, env={}) == 0


def test_restauracao_de_mais_de_nove_minutos_falha_com_a_mensagem_fixa(env, extracao, monkeypatch, capsys):
    relogio = _RelogioFalso()
    monkeypatch.setattr(cli, "_relogio", relogio)
    monkeypatch.setattr(cli, "_dormir", lambda s: None)
    client = FakeClient(fail_with=ConnectionError("ainda subindo"))

    class NoveMinutosEUmSegundo(AdminFalso):
        def acordar(self, *a, **k):
            relogio.agora += 541
            return super().acordar(*a, **k)

    admin = NoveMinutosEUmSegundo(client, liga=False)
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin, env={}) == 3
    assert "o banco não respondeu em 9 min; rode de novo" in capsys.readouterr().err
