"""Orquestração do `importar`. Tudo com dublês: FakeClient no lugar do Supabase, sanitizadores,
`build_workbook` e `contar_nd` trocados por monkeypatch. Nenhum teste toca o banco real nem o Excel.
Dados sensíveis: só sintéticos."""
import hashlib
import io
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from postgrest import APIError

from papaki_dre import cli
from papaki_dre.build import BuildReport, SheetReport
from papaki_dre.config import R_FOOTER
from papaki_dre.types import ContractError, Removal, Sanitized
from fake_supabase import FakeClient

ROOT = Path(__file__).resolve().parents[1]
OUT_OK = "Modelo DRE - Untitled - 09_2026.xlsx"
OUT_BLOCKED = "Modelo DRE - Untitled - 09_2026_BLOQUEADO.xlsx"
CPF_FALSO = "12345678900"           # CPF inventado, só para provar que não vaza


def _san(report, kept=2):
    header = ["Núm. JOB", "Valor"] if report in ("rra", "comissoes") else ["Cód.", "Valor"]
    return Sanitized(report=report, header=header, rows=[["a", 1.0]] * kept,
                     db_rows=[["a", 1.0]] * kept, source_rows=list(range(10, 10 + kept)),
                     raw_count=kept + 1, technical=[Removal(2, R_FOOTER)])


def _build_report(missing=()):
    names = {"BASE": 5, "DESPESAS": 3, "RECEITA": 2, "REAL X ORÇADO COMISSÕES": 2}
    sheets = {n: SheetReport(n, f, f + 1, 2, missing_formulas=list(missing) if n == "DESPESAS" else [])
              for n, f in names.items()}
    return BuildReport(sheets=sheets, parts={"changed": ["x"], "removed": [], "added": []})


@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.delenv("PAPAKI_DEBUG", raising=False)
    files = {}
    for k in ("cp", "cr", "rra", "comissoes"):
        p = tmp_path / f"{k}.dat"
        p.write_bytes(k.encode())
        files[k] = p
    model = tmp_path / "modelo.xlsx"
    model.write_bytes(b"modelo")
    state = {"missing": (), "na": 0, "nd_calls": [], "rra_kwargs": None, "build_args": None, "runs": 0,
             "san_kwargs": {}}

    def fake_san(key):
        def inner(*a, **kw):
            state["san_kwargs"][key] = kw
            if key == "rra":
                state["rra_kwargs"] = kw
            return _san(key)
        return inner

    monkeypatch.setattr(cli, "sanitize_cp", fake_san("cp"))
    monkeypatch.setattr(cli, "sanitize_cr", fake_san("cr"))
    monkeypatch.setattr(cli, "sanitize_comissoes", fake_san("comissoes"))
    monkeypatch.setattr(cli, "sanitize_rra", fake_san("rra"))

    def fake_build(template, out, results, **kw):
        state["build_args"] = (template, out, kw)
        state["runs"] += 1
        out.write_bytes(f"xlsx-{len(results)}-rodada{state['runs']}".encode())
        return _build_report(missing=state["missing"])

    monkeypatch.setattr(cli, "build_workbook", fake_build)

    def fake_contar_nd(xlsx, aba, colunas, primeira, ultima):
        state["nd_calls"].append((Path(xlsx).name, aba, tuple(colunas), primeira, ultima))
        if isinstance(state["na"], BaseException):
            raise state["na"]
        return state["na"]

    monkeypatch.setattr(cli, "contar_nd", fake_contar_nd)
    out = tmp_path / "saida"
    yield {"files": files, "model": model, "out": out, "state": state}
    # em nenhum caminho sobra arquivo de preparo
    if out.exists():
        assert [n for n in os.listdir(out) if n.endswith(".novo.xlsx")] == []


def _argv(env, *extra):
    f = env["files"]
    return ["importar", "--produtora", "untitled", "--competencia", "2026-09",
            "--cp", str(f["cp"]), "--cr", str(f["cr"]), "--rra", str(f["rra"]),
            "--comissoes", str(f["comissoes"]), "--modelo", str(env["model"]),
            "--saida", str(env["out"]), "--data-extracao", "2026-10-08T09:00:00-03:00", *extra]


def _conf_json(env):
    return json.loads((env["out"] / "conferencia.json").read_text(encoding="utf-8"))


class FlakyClient(FakeClient):
    """FakeClient com falhas programáveis (simula erros do PostgREST), ligáveis entre rodadas:
    `fail_insert` = tabela cujo insert falha; `fail_update_conf` = o update só de `conferencia`
    (salvar_conferencia) falha; `fail_rpc` = 'antes' (nada muda) ou 'depois' (a promoção é
    confirmada no banco, mas a resposta se perde)."""

    def __init__(self, exc, fail_insert=None):
        super().__init__()
        self.exc, self.fail_insert = exc, fail_insert
        self.fail_update_conf = False
        self.fail_rpc = None

    def table(self, name):
        q = super().table(name)
        if name == self.fail_insert:
            def insert(*a, **kw):
                raise self.exc
            q.insert = insert
        if name == "importacoes" and self.fail_update_conf:
            orig_update = q.update

            def update(values):
                if set(values) == {"conferencia"}:
                    raise self.exc
                return orig_update(values)
            q.update = update
        return q

    def rpc(self, name, params):
        r = super().rpc(name, params)
        if self.fail_rpc == "antes":
            def execute():
                raise self.exc
            r.execute = execute
        elif self.fail_rpc == "depois":
            orig = r.execute

            def execute():
                orig()
                raise self.exc
            r.execute = execute
        return r


def FailingInsertClient(table, exc):
    return FlakyClient(exc, fail_insert=table)


PREP = f".{OUT_OK[:-5]}.novo.xlsx"


# --- caminho feliz -------------------------------------------------------------------------------

def test_happy_path_without_database(env):
    code = cli.main(_argv(env, "--sem-banco"))
    assert code == 0
    assert (env["out"] / OUT_OK).exists()
    assert (env["out"] / "conferencia.md").exists()
    assert (env["out"] / "conferencia.json").exists()


def test_happy_path_with_database_promotes_the_snapshot(env):
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "concluida" and imp["vigente"] is True
    assert imp["periodo_inicio"] == "2026-01-01" and imp["periodo_fim"] == "2026-09-30"
    assert imp["data_extracao"] == "2026-10-08T09:00:00-03:00"
    assert len(client.tables["arquivos_brutos"]) == 4
    assert len(client.tables["cp_despesas"]) == 2
    assert len(client.tables["exclusoes"]) == 4                   # 1 técnica por relatório
    assert len(client.tables["cargas_excel"]) == 4
    assert imp["conferencia"]["status"] in ("ok", "com_alertas")
    assert imp["conferencia"]["meta"]["importacao_id"] == imp["id"]


def test_build_and_rra_receive_the_expected_arguments(env):
    from papaki_dre.produtoras import carregar
    prod = carregar("untitled")
    assert cli.main(_argv(env, "--sem-banco")) == 0
    template, out, kw = env["state"]["build_args"]
    assert Path(template) == env["model"] and out == env["out"] / PREP   # grava no preparo
    assert kw == {"master_formula_rows": prod.linha_mestre}               # linha-mestre vem da produtora
    assert env["state"]["rra_kwargs"] == {"work_dir": env["out"] / "work", "regras": prod.regras}


def test_todos_os_sanitizadores_recebem_as_regras_da_produtora(env):
    from papaki_dre.produtoras import carregar
    assert cli.main(_argv(env, "--sem-banco")) == 0
    regras = carregar("untitled").regras
    for chave in ("cp", "comissoes", "rra"):             # o CR não tem regra de negócio por produtora
        assert env["state"]["san_kwargs"][chave].get("regras") == regras, chave
    assert "regras" not in env["state"]["san_kwargs"]["cr"]


def test_same_files_again_are_ignored(env):
    client = FakeClient()
    cli.main(_argv(env), client=client)
    assert cli.main(_argv(env), client=client) == 0
    assert len(client.tables["importacoes"]) == 1            # nada novo foi criado


def test_forcar_creates_a_new_snapshot_and_supersedes_the_old_one(env):
    client = FakeClient()
    cli.main(_argv(env), client=client)
    assert cli.main(_argv(env, "--forcar"), client=client) == 0
    rows = client.tables["importacoes"]
    assert len(rows) == 2 and [r["vigente"] for r in rows] == [False, True]


# --- contagem de #N/D ----------------------------------------------------------------------------

def test_nd_count_uses_the_despesas_range(env):
    assert cli.main(_argv(env, "--sem-banco")) == 0
    assert env["state"]["nd_calls"] == [(PREP, "DESPESAS", ("AP", "AQ"), 3, 4)]


def test_sem_recalculo_skips_contar_nd_and_says_so(env):
    assert cli.main(_argv(env, "--sem-banco", "--sem-recalculo")) == 0
    assert env["state"]["nd_calls"] == []
    conf = _conf_json(env)
    assert "contagem de #N/D não executada" in conf["informacoes"]
    assert conf["status"] == "ok"


def test_nd_cells_become_an_alert_and_go_to_the_conference(env):
    env["state"]["na"] = 231
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "concluida" and imp["vigente"] is True
    assert imp["conferencia"]["status"] == "com_alertas"
    assert any("231" in a and "#N/D" in a for a in imp["conferencia"]["alertas"])
    conf = _conf_json(env)
    assert conf["status"] == "com_alertas"
    assert any("231" in a for a in conf["alertas"])
    assert "231" in (env["out"] / "conferencia.md").read_text(encoding="utf-8")


def test_nd_count_failure_is_a_warning_never_a_block(env):
    env["state"]["na"] = OSError("Excel não respondeu")
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "concluida" and imp["vigente"] is True
    conf = imp["conferencia"]
    assert conf["status"] == "com_alertas" and conf["bloqueios"] == []
    assert any("não foi possível recalcular" in a and "Excel não respondeu" in a for a in conf["alertas"])
    assert (env["out"] / OUT_OK).exists()


# --- bloqueios -----------------------------------------------------------------------------------

def test_rows_without_model_formula_block_and_do_not_promote(env):
    env["state"]["missing"] = ["AP3", "AQ3"]
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 2
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "bloqueada" and imp["vigente"] is False
    assert imp["conferencia"]["status"] == "bloqueada"
    assert len(client.tables["cp_despesas"]) == 2                # decisão 2: linhas ficam gravadas
    assert client.tables["cargas_excel"][0]["arquivo_saida"] == OUT_BLOCKED
    assert (env["out"] / OUT_BLOCKED).exists()
    assert not (env["out"] / OUT_OK).exists()


def test_blocked_output_replaces_an_old_blocked_file(env):
    env["state"]["missing"] = ["AP3"]
    env["out"].mkdir()
    (env["out"] / OUT_BLOCKED).write_bytes(b"velho")
    assert cli.main(_argv(env, "--sem-banco")) == 2
    assert (env["out"] / OUT_BLOCKED).read_bytes() == b"xlsx-4-rodada1"


def test_contract_error_in_the_build_blocks_and_is_recorded(env, monkeypatch):
    def boom(*a, **kw):
        raise ContractError("DESPESAS!AP: classificação ausente")
    monkeypatch.setattr(cli, "build_workbook", boom)
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 2
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "bloqueada" and imp["vigente"] is False
    assert "classificação ausente" in imp["motivo_falha"]
    assert imp["conferencia"]["bloqueios"] == ["erro de contrato: DESPESAS!AP: classificação ausente"]
    assert client.tables.get("cp_despesas", []) == []          # nada carregado
    assert "bloqueada" in (env["out"] / "conferencia.md").read_text(encoding="utf-8")
    assert env["state"]["nd_calls"] == []


def test_sanitizer_contract_error_blocks_and_is_recorded(env, monkeypatch):
    def boom(*a, **kw):
        raise ContractError("cr: coluna 3 é 'X'; esperado 'Y'")
    monkeypatch.setattr(cli, "sanitize_cr", boom)
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 2
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "bloqueada" and "coluna 3" in imp["motivo_falha"]
    assert client.tables.get("arquivos_brutos", []) == [] and client.tables.get("cr_receita", []) == []


def test_contract_error_without_database_still_writes_the_conference(env, monkeypatch):
    def boom(*a, **kw):
        raise ContractError("rra: registro com nº de colunas inesperado")
    monkeypatch.setattr(cli, "sanitize_rra", boom)
    assert cli.main(_argv(env, "--sem-banco")) == 2
    conf = (env["out"] / "conferencia.json").read_text(encoding="utf-8")
    assert "bloqueada" in conf and "colunas inesperado" in conf


# --- falhas --------------------------------------------------------------------------------------

def test_unexpected_error_marks_failure(env, monkeypatch, capsys):
    def boom(*a, **kw):
        raise OSError("disco cheio")
    monkeypatch.setattr(cli, "build_workbook", boom)
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 3
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "falhou" and imp["vigente"] is False
    assert "disco cheio" in imp["motivo_falha"]
    err = capsys.readouterr().err
    assert "FALHA: OSError: disco cheio" in err
    assert "Traceback" not in err                                # rastro só com PAPAKI_DEBUG=1


def test_debug_env_prints_the_traceback(env, monkeypatch, capsys):
    def boom(*a, **kw):
        raise OSError("disco cheio")
    monkeypatch.setattr(cli, "build_workbook", boom)
    monkeypatch.setenv("PAPAKI_DEBUG", "1")
    assert cli.main(_argv(env, "--sem-banco")) == 3
    assert "Traceback" in capsys.readouterr().err


def test_database_error_text_never_leaks_row_details(env, capsys):
    exc = APIError({"message": "new row violates check constraint", "code": "23514",
                    "details": f"Failing row contains ({CPF_FALSO}, Fulano Sintético, 1.0)", "hint": None})
    client = FailingInsertClient("cp_despesas", exc)
    assert cli.main(_argv(env), client=client) == 3
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "falhou" and imp["vigente"] is False
    assert "violates check constraint" in imp["motivo_falha"]
    assert CPF_FALSO not in imp["motivo_falha"] and "Failing row" not in imp["motivo_falha"]
    captured = capsys.readouterr()
    assert CPF_FALSO not in captured.err and CPF_FALSO not in captured.out
    assert "FALHA:" in captured.err


def test_database_error_before_promotion_marks_failure(env):
    client = FailingInsertClient("cargas_excel", APIError({"message": "timeout", "code": "57014",
                                                           "details": None, "hint": None}))
    assert cli.main(_argv(env), client=client) == 3
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "falhou" and imp["vigente"] is False


def test_error_after_promotion_never_undoes_the_snapshot(env, monkeypatch):
    # a última coisa depois de promover é o resumo no terminal; um console fechado/inválido
    # (OSError [Errno 22] no Windows) é uma falha realista nesse ponto
    real_print = print

    def fragile_print(*a, **kw):
        if a and str(a[0]).startswith("Status:"):
            raise OSError(22, "Invalid argument")
        real_print(*a, **kw)

    monkeypatch.setattr(cli, "print", fragile_print, raising=False)
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 3
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "concluida" and imp["vigente"] is True
    assert "motivo_falha" not in imp
    rpc_at = client.calls.index(("rpc", "promover_importacao", None))
    assert ("importacoes", "update", None) not in client.calls[rpc_at + 1:]   # nem tenta marcar


def test_keyboard_interrupt_marks_failure_and_propagates(env, monkeypatch):
    def interrupt(*a, **kw):
        raise KeyboardInterrupt
    monkeypatch.setattr(cli, "build_workbook", interrupt)
    client = FakeClient()
    with pytest.raises(KeyboardInterrupt):
        cli.main(_argv(env), client=client)
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "falhou" and imp["vigente"] is False
    assert imp["motivo_falha"] == "interrompida"


def test_paused_database_is_detected_and_nothing_is_written(env, capsys):
    client = FakeClient(fail_with=ConnectionError("project paused"))
    assert cli.main(_argv(env), client=client) == 3
    err = capsys.readouterr().err
    assert "banco pausado: restaure o projeto e rode de novo" in err
    assert "importacoes" not in client.tables and client.calls == []
    assert not (env["out"] / "conferencia.json").exists()
    assert not (env["out"] / OUT_OK).exists()
    assert env["state"]["nd_calls"] == []


# --- arquivos de saída: preparo, finalização e consistência com o banco --------------------------

def _snapshot(env):
    out = env["out"]
    return {n: (out / n).read_bytes() for n in (OUT_OK, "conferencia.json", "conferencia.md")}


def test_cargas_record_the_final_name_with_the_hash_of_the_delivered_file(env):
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    final_hash = hashlib.sha256((env["out"] / OUT_OK).read_bytes()).hexdigest()
    assert {(c["arquivo_saida"], c["sha256_saida"]) for c in client.tables["cargas_excel"]} == \
        {(OUT_OK, final_hash)}


def test_duplicate_run_says_no_file_was_generated(env, capsys):
    client = FakeClient()
    cli.main(_argv(env), client=client)
    capsys.readouterr()
    assert cli.main(_argv(env), client=client) == 0
    assert "nenhum arquivo foi gerado nesta pasta" in capsys.readouterr().out


def test_blocked_rerun_keeps_the_previous_ok_file_intact(env):
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    before = (env["out"] / OUT_OK).read_bytes()
    env["state"]["missing"] = ["AP3"]
    assert cli.main(_argv(env, "--forcar"), client=client) == 2
    assert (env["out"] / OUT_OK).read_bytes() == before
    assert (env["out"] / OUT_BLOCKED).read_bytes() == b"xlsx-4-rodada2"
    assert [r["vigente"] for r in client.tables["importacoes"]] == [True, False]


def test_database_failure_on_rerun_keeps_previous_files_and_conference(env):
    client = FlakyClient(APIError({"message": "timeout", "code": "57014", "details": None, "hint": None}))
    assert cli.main(_argv(env), client=client) == 0
    before = _snapshot(env)
    client.fail_insert = "cp_despesas"
    assert cli.main(_argv(env, "--forcar"), client=client) == 3
    assert _snapshot(env) == before                                   # nada da rodada 2 em disco
    assert not (env["out"] / OUT_BLOCKED).exists()
    imps = client.tables["importacoes"]
    assert [(r["status"], r["vigente"]) for r in imps] == [("concluida", True), ("falhou", False)]


@pytest.mark.parametrize("ponto", ["cargas", "salvar_conferencia", "promover"])
def test_failure_late_in_the_run_never_leaves_an_ok_conference(env, ponto):
    client = FlakyClient(APIError({"message": "conexão perdida", "code": "08006", "details": None,
                                   "hint": None}))
    if ponto == "cargas":
        client.fail_insert = "cargas_excel"
    elif ponto == "salvar_conferencia":
        client.fail_update_conf = True
    else:
        client.fail_rpc = "antes"
    assert cli.main(_argv(env), client=client) == 3
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "falhou" and imp["vigente"] is False
    conf_path = env["out"] / "conferencia.json"
    if ponto == "cargas":                       # antes da finalização: nada é publicado na pasta
        assert not conf_path.exists() and not (env["out"] / OUT_OK).exists()
    else:                                       # depois: a conferência em disco diz que falhou
        conf = _conf_json(env)
        assert conf["status"] == "falhou"
        assert f"arquivo {OUT_OK} gerado, mas a importação falhou" in conf["motivo_falha"]
        assert "conexão perdida" in conf["motivo_falha"]
        assert "falhou" in (env["out"] / "conferencia.md").read_text(encoding="utf-8")
        assert f"arquivo {OUT_OK} gerado" in imp["motivo_falha"]


def test_promotion_confirmed_but_response_lost_keeps_the_snapshot(env):
    client = FlakyClient(ConnectionError("resposta perdida"))
    client.fail_rpc = "depois"
    assert cli.main(_argv(env), client=client) == 3
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "concluida" and imp["vigente"] is True
    assert _conf_json(env)["status"] == "ok"            # coerente com o banco: a foto é vigente


def test_output_open_in_excel_is_a_clear_error_and_touches_nothing(env, monkeypatch, capsys):
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    before = _snapshot(env)
    real_replace = os.replace

    def locked_replace(src, dst):
        if str(src).endswith(".novo.xlsx"):
            raise PermissionError(13, "O arquivo já está sendo usado por outro processo", str(dst))
        return real_replace(src, dst)

    monkeypatch.setattr(cli.os, "replace", locked_replace)
    capsys.readouterr()
    assert cli.main(_argv(env, "--forcar"), client=client) == 3
    err = capsys.readouterr().err
    assert f"feche o arquivo {OUT_OK} no Excel e rode de novo" in err
    assert _snapshot(env) == before
    imps = client.tables["importacoes"]
    assert imps[0]["status"] == "concluida" and imps[0]["vigente"] is True
    assert imps[1]["status"] == "falhou" and "feche o arquivo" in imps[1]["motivo_falha"]


def test_interrupt_right_after_creating_the_import_marks_failure(env, monkeypatch):
    def interrupt(*a, **kw):
        raise KeyboardInterrupt
    monkeypatch.setattr(cli, "sanitize_cp", interrupt)
    client = FakeClient()
    with pytest.raises(KeyboardInterrupt):
        cli.main(_argv(env), client=client)
    assert client.tables["importacoes"][0]["status"] == "falhou"


# --- argumentos ----------------------------------------------------------------------------------

def test_partial_period_end_and_screen_counts(env):
    client = FakeClient()
    cli.main(_argv(env, "--periodo-fim", "2026-09-15", "--tela", "cp=9,cr=8"), client=client)
    assert client.tables["importacoes"][0]["periodo_fim"] == "2026-09-15"
    by_rel = {r["relatorio"]: r for r in client.tables["arquivos_brutos"]}
    assert by_rel["cp"]["total_tela"] == 9 and by_rel["rra"]["total_tela"] is None
    conf = client.tables["importacoes"][0]["conferencia"]
    assert any("tela mostrava 9" in a for a in conf["alertas"])


@pytest.mark.parametrize("bad", ["2026-13", "2026/09", "setembro", "26-09"])
def test_invalid_competencia_is_rejected(env, bad):
    argv = _argv(env, "--sem-banco")
    argv[argv.index("--competencia") + 1] = bad
    assert cli.main(argv) == 3


def test_unknown_produtora_is_rejected(env):
    argv = _argv(env, "--sem-banco")
    argv[argv.index("--produtora") + 1] = "outra"
    assert cli.main(argv) == 3


@pytest.mark.parametrize("extra", [
    ("--periodo-fim", "2025-12-31"),          # antes do início do período (2026-01-01)
    ("--periodo-fim", "15/09/2026"),
    ("--data-extracao", "ontem"),
    ("--tela", "xx=1"),
])
def test_invalid_options_are_usage_errors(env, extra, capsys):
    client = FakeClient()
    assert cli.main(_argv(env, *extra), client=client) == 3
    assert "ERRO DE USO" in capsys.readouterr().err
    assert client.calls == [] and not env["out"].exists()


def test_missing_required_argument_is_a_usage_error(env, capsys):
    argv = _argv(env, "--sem-banco")
    i = argv.index("--modelo")
    del argv[i:i + 2]
    assert cli.main(argv, env={}) == 3                           # não o código 2 do argparse
    assert "ERRO DE USO" in capsys.readouterr().err


def test_parse_tela():
    assert cli.parse_tela("cp=37030,cr=483") == {"cp": 37030, "cr": 483}
    with pytest.raises(cli.UsageError):
        cli.parse_tela("xx=1")
    with pytest.raises(cli.UsageError):
        cli.parse_tela("cp=abc")


def test_usage_errors_touch_nothing(env):
    client = FakeClient()
    argv = _argv(env)
    argv[argv.index("--competencia") + 1] = "2026-13"
    assert cli.main(argv, client=client) == 3
    assert client.calls == [] and "importacoes" not in client.tables
    assert not (env["out"] / "conferencia.json").exists()


# --- revisão final do Plano 3 --------------------------------------------------------------------

SEM_BANCO_INFO = "rodada sem banco: não registrada no Supabase"


def test_sem_banco_is_stated_in_the_conference(env):
    assert cli.main(_argv(env, "--sem-banco")) == 0
    conf = _conf_json(env)
    assert conf["meta"]["banco"] is False and "importacao_id" not in conf["meta"]
    assert SEM_BANCO_INFO in conf["informacoes"]
    assert SEM_BANCO_INFO in (env["out"] / "conferencia.md").read_text(encoding="utf-8")


def test_sem_banco_contract_block_is_stated_too(env, monkeypatch):
    def boom(*a, **kw):
        raise ContractError("cp: coluna 2 é 'X'; esperado 'Y'")
    monkeypatch.setattr(cli, "sanitize_cp", boom)
    assert cli.main(_argv(env, "--sem-banco")) == 2
    conf = _conf_json(env)
    assert conf["meta"]["banco"] is False and SEM_BANCO_INFO in conf["informacoes"]


def test_sem_banco_refuses_to_overwrite_the_output_of_a_real_run(env, capsys):
    assert cli.main(_argv(env), client=FakeClient()) == 0
    before = _snapshot(env)
    capsys.readouterr()
    assert cli.main(_argv(env, "--sem-banco")) == 3
    assert "ERRO DE USO" in capsys.readouterr().err
    assert _snapshot(env) == before


def test_sem_banco_may_rerun_over_its_own_output(env):
    assert cli.main(_argv(env, "--sem-banco")) == 0
    assert cli.main(_argv(env, "--sem-banco")) == 0


@pytest.mark.parametrize("nome", [OUT_OK, OUT_BLOCKED])
def test_output_that_would_overwrite_the_model_is_a_usage_error(env, nome, capsys):
    env["out"].mkdir()
    modelo = env["out"] / nome
    modelo.write_bytes(b"modelo-original")
    argv = _argv(env, "--sem-banco")
    # outra grafia do mesmo arquivo: o destino é comparado pelo caminho resolvido
    argv[argv.index("--modelo") + 1] = str(env["out"] / "sub" / ".." / nome)
    (env["out"] / "sub").mkdir()
    assert cli.main(argv) == 3
    assert "ERRO DE USO" in capsys.readouterr().err
    assert modelo.read_bytes() == b"modelo-original"
    assert not (env["out"] / "conferencia.json").exists()


def test_unencodable_final_message_never_turns_a_good_run_into_a_failure(env, monkeypatch):
    real_print = print

    def cp1252_print(*a, **kw):
        if a and str(a[0]).startswith("Status:"):
            raise UnicodeEncodeError("cp1252", "Ł", 0, 1, "character maps to <undefined>")
        real_print(*a, **kw)

    monkeypatch.setattr(cli, "print", cp1252_print, raising=False)
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "concluida" and imp["vigente"] is True


def test_cp1252_console_is_tolerated(env, monkeypatch):
    stream = io.TextIOWrapper(io.BytesIO(), encoding="cp1252", errors="strict")
    monkeypatch.setattr(sys, "stdout", stream)
    argv = _argv(env, "--sem-banco")
    out = env["out"].with_name("saida_Łódź")         # caminho com letras fora do cp1252
    argv[argv.index("--saida") + 1] = str(out)
    assert cli.main(argv) == 0
    assert stream.errors == "replace"
    stream.flush()
    assert b"Status: ok" in stream.buffer.getvalue()
    assert (out / OUT_OK).exists()


# --- __main__ ------------------------------------------------------------------------------------

def test_python_dash_m_help_exits_zero(tmp_path):
    env = {**os.environ, "PYTHONPATH": str(ROOT / "src"), "PYTHONIOENCODING": "utf-8"}
    for args in (["--help"], ["importar", "--help"]):
        r = subprocess.run([sys.executable, "-m", "papaki_dre", *args], cwd=tmp_path, env=env,
                           capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60)
        assert r.returncode == 0, r.stderr
        assert "importar" in r.stdout


def test_blocked_rerun_warns_that_the_old_ok_file_is_not_from_this_run(env):
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    env["state"]["missing"] = ["AP3"]
    assert cli.main(_argv(env, "--forcar"), client=client) == 2
    conf = json.loads((env["out"] / "conferencia.json").read_text(encoding="utf-8"))
    assert conf["status"] == "bloqueada"
    assert any(OUT_OK in i and "NÃO é desta rodada" in i for i in conf["informacoes"])


def test_importar_grava_a_regra_da_produtora_em_exclusoes(env, monkeypatch):
    from papaki_dre.produtoras import carregar
    recebidas = []

    def espia(regras):
        recebidas.append(regras)
        return {R_FOOTER: "REGRA-DA-PRODUTORA"}               # texto inconfundível: prova de onde veio

    monkeypatch.setattr(cli, "regras_texto", espia)
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    assert recebidas and all(r == carregar("untitled").regras for r in recebidas)
    assert {r["regra"] for r in client.tables["exclusoes"]} == {"REGRA-DA-PRODUTORA"}


def test_importar_garante_a_linha_da_produtora_no_banco(env):
    client = FakeClient()
    client.tables["produtoras"] = []                       # banco novo, sem a linha da produtora
    assert cli.main(_argv(env), client=client) == 0
    assert [r["id"] for r in client.tables["produtoras"]] == ["untitled"]


# --- várias máquinas: máquina gravada e reserva liberada ------------------------------------------

def _reservar(client, maquina):
    from papaki_dre.db import Repository
    return Repository(client).reservar_rodada("untitled", "2026-09", maquina, 6)


def _status_reservas(client):
    return [r["status"] for r in client.tables.get("reservas_rodada", [])]


def test_importar_grava_a_maquina_e_libera_a_reserva(env, monkeypatch):
    monkeypatch.setenv("PAPAKI_MAQUINA", "principal")
    client = FakeClient()
    _reservar(client, "principal")
    assert cli.main(_argv(env), client=client) == 0
    assert client.tables["importacoes"][0]["maquina"] == "principal"
    assert _status_reservas(client) == ["liberada"]


def test_importar_sem_nome_da_maquina_nao_manda_a_coluna_e_nao_mexe_na_reserva(env):
    client = FakeClient()
    _reservar(client, "grupo-a")
    assert cli.main(_argv(env), client=client) == 0
    assert "maquina" not in client.tables["importacoes"][0]
    assert _status_reservas(client) == ["ativa"]


def test_importar_nao_libera_a_reserva_de_outra_maquina(env, monkeypatch):
    monkeypatch.setenv("PAPAKI_MAQUINA", "principal")
    client = FakeClient()
    _reservar(client, "grupo-a")
    assert cli.main(_argv(env), client=client) == 0
    assert _status_reservas(client) == ["ativa"]


def test_importar_bloqueada_tambem_libera_a_reserva(env, monkeypatch):
    monkeypatch.setenv("PAPAKI_MAQUINA", "principal")

    def bloqueia(*a, **kw):
        raise ContractError("cabeçalho diferente")

    monkeypatch.setattr(cli, "sanitize_cp", bloqueia)
    client = FakeClient()
    _reservar(client, "principal")
    assert cli.main(_argv(env), client=client) == 2
    assert _status_reservas(client) == ["liberada"]


def test_importar_identico_tambem_libera_a_reserva(env, monkeypatch):
    monkeypatch.setenv("PAPAKI_MAQUINA", "principal")
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    _reservar(client, "principal")
    assert cli.main(_argv(env), client=client) == 0           # mesmos arquivos: "Arquivos idênticos já importados"
    assert len(client.tables["importacoes"]) == 1
    assert _status_reservas(client) == ["liberada"]


def test_falha_ao_liberar_nao_derruba_a_importacao(env, monkeypatch):
    monkeypatch.setenv("PAPAKI_MAQUINA", "principal")
    client = FakeClient()
    monkeypatch.setattr("papaki_dre.db.Repository.liberar_rodada",
                        lambda *a, **kw: (_ for _ in ()).throw(RuntimeError("banco caiu")))
    assert cli.main(_argv(env), client=client) == 0
    assert client.tables["importacoes"][0]["status"] == "concluida"
