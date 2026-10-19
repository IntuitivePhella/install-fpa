import datetime as dt
import hashlib
import os

import pytest

from papaki_dre.config import R_BLANK, R_FOOTER, R_H, R_JOB, R_LOC, R_MODELO, R_NUCLEO, R_PESQ
from papaki_dre.db import (BATCH, BancoIndisponivel, REGRAS, Repository, json_value, make_client,
                           mensagem_segura, row_to_json, sha256_file)
from papaki_dre.types import Formula, Removal, Sanitized, XlError
from fake_supabase import FakeAPIError, FakeClient


def test_sha256_file(tmp_path):
    p = tmp_path / "a.bin"
    p.write_bytes(b"abc")
    assert sha256_file(p) == hashlib.sha256(b"abc").hexdigest()


def test_json_value_types():
    assert json_value(dt.datetime(2026, 1, 5)) == "2026-01-05T00:00:00"
    assert json_value(dt.date(2026, 1, 5)) == "2026-01-05"
    assert json_value(1.5) == 1.5 and json_value(3) == 3 and json_value("x") == "x"
    assert json_value(True) is True and json_value(None) is None
    assert json_value(float("nan")) is None
    assert json_value(XlError("#DIV/0!")) is None       # mesma política do Plano 1 (RowCopier): erro -> nulo no banco


def test_json_value_rejects_formulas():
    # db_rows é só valores; uma fórmula aqui significa que `rows` foi passado no lugar de `db_rows`
    with pytest.raises(TypeError, match="db_rows"):
        json_value(Formula("=SUM(A1:A2)"))


def test_every_removal_reason_has_a_rule():
    for r in (R_BLANK, R_FOOTER, R_JOB, R_LOC, R_NUCLEO, R_PESQ, R_MODELO, R_H):
        assert REGRAS[r]


def test_row_to_json_disambiguates_duplicate_headers():
    d = row_to_json(["Recebido", "Valor", "Recebido"], [dt.datetime(2026, 1, 6), 10.0, "Sim"])
    assert d == {"Recebido": "2026-01-06T00:00:00", "Valor": 10.0, "Recebido (2)": "Sim"}


def _san(report, rows, header=("Cód.", "Valor"), exclusions=(), technical=()):
    return Sanitized(report=report, header=list(header), rows=rows, db_rows=rows,
                     source_rows=[10 + i for i in range(len(rows))], raw_count=len(rows),
                     exclusions=list(exclusions), technical=list(technical))


def _repo():
    return Repository(FakeClient())


def _new(repo, comp="2026-09"):
    return repo.criar_importacao(produtora="untitled", competencia=comp, periodo_inicio="2026-01-01",
                                 periodo_fim="2026-09-30", data_extracao="2026-10-08T16:34:00+00:00",
                                 origem="teste", gatilho="manual")


def _imps(repo):
    return {r["id"]: r for r in repo.client.tables["importacoes"]}


def test_criar_importacao_starts_in_progress_and_not_current():
    repo = _repo()
    imp = _new(repo)
    row = repo.client.tables["importacoes"][0]
    assert row["id"] == imp and row["status"] == "em_andamento" and row["vigente"] is False


def test_inserir_linhas_stores_key_destination_row_and_json():
    repo = _repo()
    imp = _new(repo)
    s = _san("cp", [["untitled-1", 10.0], ["untitled-2", 20.0]])
    repo.inserir_linhas(imp, "untitled", "2026-09", "cp", s)
    rows = repo.client.tables["cp_despesas"]
    assert [r["cod"] for r in rows] == ["untitled-1", "untitled-2"]
    assert [r["linha_origem"] for r in rows] == [10, 11]
    assert [r["linha_destino"] for r in rows] == [3, 4]          # DEST["cp"]["first_row"] = 3
    assert rows[0]["dados"] == {"Cód.": "untitled-1", "Valor": 10.0}
    assert {(r["importacao_id"], r["produtora_id"], r["competencia"]) for r in rows} == {(imp, "untitled", "2026-09")}


def test_inserir_linhas_uses_db_rows_not_excel_rows():
    repo = _repo()
    imp = _new(repo)
    s = Sanitized(report="comissoes", header=["Núm. JOB", "Real"], rows=[["75131", Formula("=B2*2")]],
                  db_rows=[["75131", 40.0]], source_rows=[4], raw_count=1)
    repo.inserir_linhas(imp, "untitled", "2026-09", "comissoes", s)
    row = repo.client.tables["comissoes"][0]
    assert row["dados"] == {"Núm. JOB": "75131", "Real": 40.0} and row["num_job"] == "75131"
    assert row["linha_origem"] == 4 and row["linha_destino"] == 2      # DEST["comissoes"]["first_row"] = 2


def test_inserir_linhas_numeric_key_becomes_text():
    repo = _repo()
    imp = _new(repo)
    s = _san("rra", [[75131, 1.0], [None, 2.0]], header=("Núm. JOB", "Valor"))
    repo.inserir_linhas(imp, "untitled", "2026-09", "rra", s)
    assert [r["num_job"] for r in repo.client.tables["rra_base"]] == ["75131", None]
    assert repo.client.tables["rra_base"][0]["linha_destino"] == 5     # DEST["rra"]["first_row"] = 5


def test_inserir_linhas_batches_large_inputs():
    repo = _repo()
    imp = _new(repo)
    s = _san("cp", [[f"c{i}", 1.0] for i in range(1201)])
    repo.inserir_linhas(imp, "untitled", "2026-09", "cp", s)
    assert len(repo.client.tables["cp_despesas"]) == 1201
    inserts = [n for t, op, n in repo.client.calls if t == "cp_despesas" and op == "insert"]
    assert inserts == [500, 500, 201] and BATCH == 500


def test_inserir_linhas_with_no_rows_makes_no_call():
    repo = _repo()
    imp = _new(repo)
    repo.inserir_linhas(imp, "untitled", "2026-09", "cp", _san("cp", []))
    assert not [c for c in repo.client.calls if c[0] == "cp_despesas"]


def test_cr_has_no_promoted_key():
    repo = _repo()
    imp = _new(repo)
    repo.inserir_linhas(imp, "untitled", "2026-09", "cr", _san("cr", [["x", 1.0]], header=("Cliente", "Vl")))
    assert "cod" not in repo.client.tables["cr_receita"][0]
    assert "num_job" not in repo.client.tables["cr_receita"][0]


def test_inserir_exclusoes_records_technical_and_business():
    repo = _repo()
    imp = _new(repo)
    s = _san("cp", [], exclusions=[Removal(5, R_LOC)], technical=[Removal(2, R_FOOTER)])
    repo.inserir_exclusoes(imp, "cp", s)
    got = sorted((r["tipo"], r["motivo"], r["linha_origem"]) for r in repo.client.tables["exclusoes"])
    assert got == [("negocio", R_LOC, 5), ("tecnica", R_FOOTER, 2)]
    assert {r["regra"] for r in repo.client.tables["exclusoes"]} == {REGRAS[R_LOC], REGRAS[R_FOOTER]}
    assert {r["relatorio"] for r in repo.client.tables["exclusoes"]} == {"cp"}


def test_inserir_exclusoes_large_round_is_batched():
    # ~40 mil exclusões por rodada: lotes de 500, nunca um round-trip por linha
    repo = _repo()
    imp = _new(repo)
    s = _san("cp", [], exclusions=[Removal(i, R_LOC) for i in range(40_000)],
             technical=[Removal(40_001, R_FOOTER)])
    repo.inserir_exclusoes(imp, "cp", s)
    assert len(repo.client.tables["exclusoes"]) == 40_001
    inserts = [n for t, op, n in repo.client.calls if t == "exclusoes" and op == "insert"]
    assert len(inserts) == 81 and max(inserts) == 500


def test_registrar_arquivo_and_find_identical():
    repo = _repo()
    imp = _new(repo)
    hashes = {"rra": "h1", "cp": "h2", "cr": "h3", "comissoes": "h4"}
    for rel, h in hashes.items():
        repo.registrar_arquivo(imp, rel, nome=f"{rel}.x", sha256=h, tamanho=10, total_tela=None)
    row = next(r for r in repo.client.tables["arquivos_brutos"] if r["relatorio"] == "cp")
    assert row["nome_arquivo"] == "cp.x" and row["sha256"] == "h2" and row["tamanho_bytes"] == 10
    assert repo.find_identical("untitled", "2026-09", hashes) is None     # ainda em andamento
    repo.promover(imp)
    assert repo.find_identical("untitled", "2026-09", hashes) == imp
    assert repo.find_identical("untitled", "2026-09", {**hashes, "cp": "OUTRO"}) is None
    assert repo.find_identical("untitled", "2026-10", hashes) is None
    assert repo.find_identical("outra", "2026-09", hashes) is None


def test_find_identical_ignores_blocked_or_failed_imports():
    repo = _repo()
    hashes = {"rra": "h1", "cp": "h2", "cr": "h3", "comissoes": "h4"}
    for status in ("bloqueada", "falhou"):
        imp = _new(repo)
        for rel, h in hashes.items():
            repo.registrar_arquivo(imp, rel, nome=f"{rel}.x", sha256=h, tamanho=10)
        repo.marcar(imp, status, motivo="x")
    assert repo.find_identical("untitled", "2026-09", hashes) is None


def test_find_identical_accepts_a_concluded_non_current_import():
    # spec §4: mesmo hash = ignorado, mesmo que a foto já tenha sido substituída por outra
    repo = _repo()
    hashes = {"rra": "h1", "cp": "h2", "cr": "h3", "comissoes": "h4"}
    a = _new(repo)
    for rel, h in hashes.items():
        repo.registrar_arquivo(a, rel, nome=f"{rel}.x", sha256=h, tamanho=10)
    repo.promover(a)
    b = _new(repo)
    repo.promover(b)
    assert _imps(repo)[a]["vigente"] is False
    assert repo.find_identical("untitled", "2026-09", hashes) == a


def test_find_identical_requires_the_same_set_of_files():
    repo = _repo()
    imp = _new(repo)
    hashes = {"rra": "h1", "cp": "h2", "cr": "h3", "comissoes": "h4"}
    for rel in ("rra", "cp", "cr"):
        repo.registrar_arquivo(imp, rel, nome=f"{rel}.x", sha256=hashes[rel], tamanho=10)
    repo.promover(imp)
    assert repo.find_identical("untitled", "2026-09", hashes) is None
    # mesmos hashes em relatórios trocados também não é a mesma foto
    assert repo.find_identical("untitled", "2026-09", {"rra": "h2", "cp": "h1", "cr": "h3"}) is None
    assert repo.find_identical("untitled", "2026-09", {"rra": "h1", "cp": "h2", "cr": "h3"}) == imp


def test_find_identical_filters_by_hash_on_the_server_despite_the_row_limit():
    # o PostgREST devolve no máximo 1000 linhas por consulta: a busca não pode ler a tabela inteira
    repo = _repo()
    for i in range(260):                                     # 1040 arquivos de outras fotos
        other = _new(repo)
        for rel in ("rra", "cp", "cr", "comissoes"):
            repo.registrar_arquivo(other, rel, nome="x", sha256=f"o{i}-{rel}", tamanho=1)
        repo.marcar(other, "falhou", motivo="x")
    imp = _new(repo)
    hashes = {"rra": "h1", "cp": "h2", "cr": "h3", "comissoes": "h4"}
    for rel, h in hashes.items():
        repo.registrar_arquivo(imp, rel, nome=f"{rel}.x", sha256=h, tamanho=10)
    repo.promover(imp)
    assert len(repo.client.tables["arquivos_brutos"]) > 1000
    assert repo.find_identical("untitled", "2026-09", hashes) == imp


def test_find_identical_with_no_hashes_is_none():
    assert _repo().find_identical("untitled", "2026-09", {}) is None


def test_promover_replaces_the_previous_current_one_without_deleting_it():
    repo = _repo()
    a, b = _new(repo), _new(repo)
    repo.promover(a)
    repo.promover(b)
    by_id = _imps(repo)
    assert by_id[a]["vigente"] is False and by_id[a]["status"] == "concluida"
    assert by_id[b]["vigente"] is True and by_id[b]["status"] == "concluida"


def test_promover_only_in_progress():
    repo = _repo()
    a = _new(repo)
    repo.promover(a)
    with pytest.raises(FakeAPIError, match="não está em andamento"):
        repo.promover(a)
    b = _new(repo)
    repo.marcar(b, "bloqueada", motivo="x")
    with pytest.raises(FakeAPIError):
        repo.promover(b)
    assert _imps(repo)[a]["vigente"] is True and _imps(repo)[b]["vigente"] is False


def test_other_competencia_keeps_its_own_current():
    repo = _repo()
    a, b = _new(repo, "2026-08"), _new(repo, "2026-09")
    repo.promover(a)
    repo.promover(b)
    assert {r["competencia"]: r["vigente"] for r in repo.client.tables["importacoes"]} == {
        "2026-08": True, "2026-09": True}


def test_marcar_falhou_and_bloqueada():
    repo = _repo()
    a, b = _new(repo), _new(repo)
    assert repo.marcar(a, "falhou", motivo="erro X") is True
    repo.marcar(b, "bloqueada", motivo="ContractError: classificação ausente",
                conferencia={"status": "bloqueada"})
    rows = _imps(repo)
    assert rows[a]["status"] == "falhou" and rows[a]["motivo_falha"] == "erro X"
    assert rows[b]["status"] == "bloqueada" and rows[b]["vigente"] is False
    assert "classificação ausente" in rows[b]["motivo_falha"]
    assert rows[b]["conferencia"] == {"status": "bloqueada"}


def test_marcar_never_undoes_a_concluded_current_import():
    # uma exceção depois de promover() não pode desfazer a foto vigente
    repo = _repo()
    a = _new(repo)
    repo.promover(a)
    assert repo.marcar(a, "falhou", motivo="erro depois de promover") is False
    row = _imps(repo)[a]
    assert row["status"] == "concluida" and row["vigente"] is True and row.get("motivo_falha") is None


def test_marcar_does_not_touch_an_already_failed_import():
    repo = _repo()
    a = _new(repo)
    repo.marcar(a, "bloqueada", motivo="primeiro")
    assert repo.marcar(a, "falhou", motivo="segundo") is False
    assert _imps(repo)[a]["status"] == "bloqueada" and _imps(repo)[a]["motivo_falha"] == "primeiro"


def test_marcar_rejects_invalid_status():
    repo = _repo()
    a = _new(repo)
    with pytest.raises(ValueError):
        repo.marcar(a, "concluida")
    with pytest.raises(ValueError):
        repo.marcar(a, "xpto")
    assert _imps(repo)[a]["status"] == "em_andamento"


def test_fake_enforces_vigente_only_if_concluded():
    # garante que o dublê imita o banco: vigente com status != concluida é rejeitado
    repo = _repo()
    a = _new(repo)
    with pytest.raises(FakeAPIError, match="vigente_so_se_concluida"):
        repo.client.table("importacoes").update({"vigente": True}).eq("id", a).execute()
    assert _imps(repo)[a]["vigente"] is False


def test_fake_enforces_one_current_per_competencia():
    repo = _repo()
    a, b = _new(repo), _new(repo)
    repo.promover(a)
    repo.client.table("importacoes").update({"status": "concluida"}).eq("id", b).execute()
    with pytest.raises(FakeAPIError, match="importacoes_uma_vigente"):
        repo.client.table("importacoes").update({"vigente": True}).eq("id", b).execute()


def test_salvar_conferencia():
    repo = _repo()
    a = _new(repo)
    repo.salvar_conferencia(a, {"status": "ok"})
    assert _imps(repo)[a]["conferencia"] == {"status": "ok"}


def test_registrar_carga():
    repo = _repo()
    imp = _new(repo)
    repo.registrar_carga(imp, aba="BASE", primeira=5, ultima=76, arquivo="m.xlsx", sha256="zz")
    row = repo.client.tables["cargas_excel"][0]
    assert row["linhas"] == 72 and row["aba"] == "BASE"
    assert (row["linha_inicial"], row["linha_final"], row["arquivo_saida"], row["sha256_saida"]) == (
        5, 76, "m.xlsx", "zz")


def test_verificar_acesso_detects_paused_or_unreachable_database():
    Repository(FakeClient()).verificar_acesso()                     # banco ativo: não levanta
    repo = Repository(FakeClient(fail_with=ConnectionError("name resolution failed")))
    with pytest.raises(BancoIndisponivel, match="name resolution failed"):
        repo.verificar_acesso()


def test_verificar_acesso_never_leaks_url_or_key(monkeypatch):
    url, key = "https://refinventado.supabase.co", "chave-inventada-123"
    monkeypatch.setenv("SUPABASE_URL", url)
    monkeypatch.setenv("SUPABASE_SERVICE_KEY", key)
    err = ConnectionError(f"falhou {url}/rest/v1 host refinventado.supabase.co apikey={key}")
    with pytest.raises(BancoIndisponivel) as exc:
        Repository(FakeClient(fail_with=err)).verificar_acesso()
    msg = str(exc.value)
    assert "ConnectionError" in msg and key not in msg and "refinventado" not in msg


def _no_dotenv(monkeypatch):
    import dotenv
    monkeypatch.setattr(dotenv, "dotenv_values", lambda *a, **k: {})    # nunca lê o .env real


def test_make_client_names_only_the_missing_variable(monkeypatch):
    _no_dotenv(monkeypatch)
    monkeypatch.setenv("SUPABASE_URL", "https://refinventado.supabase.co")
    monkeypatch.delenv("SUPABASE_SERVICE_KEY", raising=False)
    with pytest.raises(RuntimeError) as exc:
        make_client()
    msg = str(exc.value)
    assert "SUPABASE_SERVICE_KEY" in msg and "refinventado" not in msg
    assert "SUPABASE_URL" not in msg.replace("SUPABASE_SERVICE_KEY", "")


def test_make_client_names_both_missing_variables(monkeypatch):
    _no_dotenv(monkeypatch)
    monkeypatch.delenv("SUPABASE_URL", raising=False)
    monkeypatch.delenv("SUPABASE_SERVICE_KEY", raising=False)
    with pytest.raises(RuntimeError, match="SUPABASE_URL.*SUPABASE_SERVICE_KEY"):
        make_client()


def test_make_client_uses_the_environment(monkeypatch):
    import supabase
    _no_dotenv(monkeypatch)
    monkeypatch.setenv("SUPABASE_URL", "https://refinventado.supabase.co")
    monkeypatch.setenv("SUPABASE_SERVICE_KEY", "chave-inventada")
    seen = {}
    monkeypatch.setattr(supabase, "create_client", lambda u, k: seen.update(u=u, k=k) or "cliente")
    assert make_client() == "cliente"
    assert seen == {"u": "https://refinventado.supabase.co", "k": "chave-inventada"}


def test_make_client_dotenv_overrides_windows_variables(monkeypatch):
    """O .env do projeto vale sobre variáveis já definidas no Windows (nunca mistura URL/chave de
    outro projeto). O dotenv é um dublê com valores inventados: o .env real nunca é lido."""
    import dotenv
    import supabase
    monkeypatch.setenv("SUPABASE_URL", "https://outroprojeto.supabase.co")
    monkeypatch.setenv("SUPABASE_SERVICE_KEY", "chave-de-outro-projeto")
    arquivo = {"SUPABASE_URL": "https://refinventado.supabase.co", "SUPABASE_SERVICE_KEY": "chave-inventada"}

    lidos = []

    def fake_dotenv_values(caminho=None, *a, **k):
        lidos.append(caminho)
        return dict(arquivo)

    monkeypatch.setattr(dotenv, "dotenv_values", fake_dotenv_values)
    seen = {}
    monkeypatch.setattr(supabase, "create_client", lambda u, k: seen.update(u=u, k=k) or "cliente")
    assert make_client() == "cliente"
    assert seen == {"u": "https://refinventado.supabase.co", "k": "chave-inventada"}
    from papaki_dre import ambiente
    assert lidos == [ambiente.RAIZ_PROJETO / ".env"]       # o mesmo arquivo que `ambiente` (e --restaurar-banco) lê
    # mensagem_segura continua escondendo os valores usados pelo cliente
    assert os.environ["SUPABASE_URL"] == "https://refinventado.supabase.co"


def test_make_client_com_env_explicito_nao_le_o_dotenv(monkeypatch):
    import dotenv
    import supabase
    monkeypatch.setattr(dotenv, "dotenv_values", lambda *a, **k: pytest.fail("leu o .env"))
    for nome in ("SUPABASE_URL", "SUPABASE_SERVICE_KEY", "SUPABASE_ACCESS_TOKEN"):
        monkeypatch.setenv(nome, "x")             # registra o estado para o monkeypatch restaurar
        monkeypatch.delenv(nome)
    seen = {}
    monkeypatch.setattr(supabase, "create_client", lambda u, k: seen.update(u=u, k=k) or "cliente")
    env = {"SUPABASE_URL": "https://refinventado.supabase.co", "SUPABASE_SERVICE_KEY": "chave-inventada",
           "SUPABASE_ACCESS_TOKEN": "sbp_inventado"}
    assert make_client(env) == "cliente"
    assert seen == {"u": "https://refinventado.supabase.co", "k": "chave-inventada"}
    err = ConnectionError("falhou https://refinventado.supabase.co chave-inventada sbp_inventado")
    msg = mensagem_segura(err)
    assert "refinventado" not in msg and "chave-inventada" not in msg and "sbp_inventado" not in msg


def test_make_client_valor_vazio_conta_como_ausente(monkeypatch):
    with pytest.raises(RuntimeError, match="SUPABASE_SERVICE_KEY"):
        make_client({"SUPABASE_URL": "https://refinventado.supabase.co", "SUPABASE_SERVICE_KEY": "  "})


def test_banco_indisponivel_sem_context(monkeypatch):
    with pytest.raises(BancoIndisponivel) as exc:
        Repository(FakeClient(fail_with=ConnectionError("https://refinventado.supabase.co"))).verificar_acesso()
    assert exc.value.__context__ is None and exc.value.__cause__ is None


# --- revisão de qualidade da Task 3 ---------------------------------------------------------------

def _inserts_returning(repo, table):
    return [c for c in repo.client.calls if c[0] == table and c[1] == "insert"]


def test_bulk_inserts_ask_for_minimal_return():
    from postgrest.types import ReturnMethod
    repo = _repo()
    imp = _new(repo)
    repo.registrar_arquivo(imp, "cp", nome="cp.x", sha256="h", tamanho=1)
    repo.inserir_linhas(imp, "untitled", "2026-09", "cp", _san("cp", [["c", 1.0]]))
    repo.inserir_exclusoes(imp, "cp", _san("cp", [], exclusions=[Removal(5, R_LOC)]))
    repo.registrar_carga(imp, aba="DESPESAS", primeira=3, ultima=3, arquivo="m.xlsx", sha256="z")
    assert repo.client.returning == {
        "importacoes": [None],                               # criar_importacao precisa do id de volta
        "arquivos_brutos": [ReturnMethod.minimal], "cp_despesas": [ReturnMethod.minimal],
        "exclusoes": [ReturnMethod.minimal], "cargas_excel": [ReturnMethod.minimal]}


def test_fake_minimal_insert_returns_no_data():
    from postgrest.types import ReturnMethod
    c = FakeClient()
    res = c.table("produtoras").insert({"id": "x", "nome": "X", "localidade": "L"},
                                       returning=ReturnMethod.minimal).execute()
    assert res.data == []


def test_find_identical_not_hidden_by_many_failed_imports_with_the_same_hashes():
    repo = _repo()
    hashes = {"rra": "h1", "cp": "h2", "cr": "h3", "comissoes": "h4"}
    for i in range(260):                                     # 1040 linhas com os MESMOS hashes
        other = _new(repo)
        for rel, h in hashes.items():
            repo.registrar_arquivo(other, rel, nome="x", sha256=h, tamanho=1)
        repo.marcar(other, "falhou" if i % 2 else "bloqueada", motivo="x")
    imp = _new(repo)
    for rel, h in hashes.items():
        repo.registrar_arquivo(imp, rel, nome=f"{rel}.x", sha256=h, tamanho=10)
    repo.promover(imp)
    assert len(repo.client.tables["arquivos_brutos"]) > 1000
    assert repo.find_identical("untitled", "2026-09", hashes) == imp


def test_find_identical_subset_of_hashes_does_not_match():
    repo = _repo()
    imp = _new(repo)
    hashes = {"rra": "h1", "cp": "h2", "cr": "h3", "comissoes": "h4"}
    for rel, h in hashes.items():
        repo.registrar_arquivo(imp, rel, nome=f"{rel}.x", sha256=h, tamanho=10)
    repo.promover(imp)
    assert repo.find_identical("untitled", "2026-09", {"rra": "h1", "cp": "h2", "cr": "h3"}) is None


class _ErroSintetico(Exception):
    """Imita postgrest.APIError: message/code/hint/details."""
    def __init__(self, message, code, details, hint=None):
        self.message, self.code, self.details, self.hint = message, code, details, hint
        super().__init__(f"{message} {details} {hint}")


URL_INVENTADA = "https://refinventado.supabase.co"
CPF_INVENTADO = "12345678900"


def _erro_com_linha():
    return _ErroSintetico(
        message=f'new row for relation "cp_despesas" violates check constraint em {URL_INVENTADA}/rest/v1',
        code="23514",
        details=f"Failing row contains ({CPF_INVENTADO}, Fulano Inventado, 99.999.999/0001-99).",
        hint=f"veja {URL_INVENTADA}")


def test_mensagem_segura_drops_details_and_secrets(monkeypatch):
    monkeypatch.setenv("SUPABASE_URL", URL_INVENTADA)
    monkeypatch.setenv("SUPABASE_SERVICE_KEY", "chave-inventada-xyz")
    msg = mensagem_segura(_erro_com_linha())
    assert msg.startswith("_ErroSintetico") and "23514" in msg and "violates check constraint" in msg
    for proibido in (CPF_INVENTADO, "Failing row", "Fulano", "refinventado", "chave-inventada"):
        assert proibido not in msg


def test_mensagem_segura_real_postgrest_error(monkeypatch):
    from postgrest.exceptions import APIError
    monkeypatch.setenv("SUPABASE_URL", URL_INVENTADA)
    e = APIError({"message": "duplicate key value violates unique constraint \"x\"", "code": "23505",
                  "details": f"Key (cod)=({CPF_INVENTADO}) already exists.", "hint": None})
    msg = mensagem_segura(e)
    assert "APIError" in msg and "23505" in msg and CPF_INVENTADO not in msg


def test_mensagem_segura_generic_exception_strips_failing_row(monkeypatch):
    monkeypatch.setenv("SUPABASE_URL", URL_INVENTADA)
    e = RuntimeError(f"falhou em {URL_INVENTADA}. Failing row contains ({CPF_INVENTADO}, x)")
    msg = mensagem_segura(e)
    assert msg.startswith("RuntimeError: falhou em") and CPF_INVENTADO not in msg and "refinventado" not in msg


def test_mensagem_segura_is_truncated():
    assert len(mensagem_segura(RuntimeError("x" * 5000))) <= 500


def test_marcar_scrubs_and_truncates_motivo(monkeypatch):
    monkeypatch.setenv("SUPABASE_URL", URL_INVENTADA)
    monkeypatch.setenv("SUPABASE_SERVICE_KEY", "chave-inventada-xyz")
    repo = _repo()
    a, b = _new(repo), _new(repo)
    repo.marcar(a, "falhou", motivo=mensagem_segura(_erro_com_linha()))
    repo.marcar(b, "falhou", motivo=f"bruto {URL_INVENTADA} chave-inventada-xyz Failing row contains "
                                    f"({CPF_INVENTADO}) " + "y" * 2000)
    for imp in (a, b):
        motivo = _imps(repo)[imp]["motivo_falha"]
        assert len(motivo) <= 500
        for proibido in (CPF_INVENTADO, "Failing row", "refinventado", "chave-inventada"):
            assert proibido not in motivo


def test_verificar_acesso_uses_safe_message(monkeypatch):
    monkeypatch.setenv("SUPABASE_URL", URL_INVENTADA)
    with pytest.raises(BancoIndisponivel) as exc:
        Repository(FakeClient(fail_with=_erro_com_linha())).verificar_acesso()
    assert CPF_INVENTADO not in str(exc.value) and "refinventado" not in str(exc.value)


def test_json_value_strips_nul_characters():
    assert json_value("a\x00b") == "ab"
    assert row_to_json(["Co\x00l"], ["v\x00"]) == {"Col": "v"}


def test_fake_rejects_nul_and_nan_and_non_json_payloads():
    c = FakeClient()
    with pytest.raises(FakeAPIError):
        c.table("produtoras").insert({"id": "x\x00", "nome": "X", "localidade": "L"}).execute()
    with pytest.raises((FakeAPIError, ValueError)):
        c.table("produtoras").insert({"id": "y", "nome": float("nan"), "localidade": "L"}).execute()
    with pytest.raises((FakeAPIError, TypeError)):
        c.table("produtoras").insert({"id": "z", "nome": dt.date(2026, 1, 1), "localidade": "L"}).execute()
    with pytest.raises(FakeAPIError):
        c.table("produtoras").update({"nome": "a\x00"}).eq("id", "untitled").execute()
    assert [r["id"] for r in c.tables["produtoras"]] == ["untitled"]


def test_fake_select_projects_only_requested_columns():
    repo = _repo()
    imp = _new(repo)
    got = repo.client.table("importacoes").select("id, status").eq("id", imp).execute().data
    assert got == [{"id": imp, "status": "em_andamento"}]
    full = repo.client.table("importacoes").select("*").eq("id", imp).execute().data[0]
    assert {"competencia", "vigente", "gatilho"} <= set(full)
    assert set(repo.client.table("importacoes").select().eq("id", imp).execute().data[0]) == set(full)


def test_identica_diz_se_a_importacao_identica_e_a_vigente():
    repo = _repo()
    hashes = {"rra": "h1", "cp": "h2", "cr": "h3", "comissoes": "h4"}
    assert repo.identica("untitled", "2026-09", hashes) is None and repo.vigente("untitled", "2026-09") is None
    a = _new(repo)
    for rel, h in hashes.items():
        repo.registrar_arquivo(a, rel, nome=f"{rel}.x", sha256=h, tamanho=10)
    repo.promover(a)
    assert repo.identica("untitled", "2026-09", hashes) == {"id": a, "vigente": True}
    assert repo.vigente("untitled", "2026-09") == a
    b = _new(repo)
    repo.promover(b)
    assert repo.identica("untitled", "2026-09", hashes) == {"id": a, "vigente": False}
    assert repo.vigente("untitled", "2026-09") == b and repo.vigente("untitled", "2026-10") is None
    assert repo.find_identical("untitled", "2026-09", hashes) == a       # assinatura antiga intacta


# --- rodadas órfãs (Plano 4, Task 9) ------------------------------------------------------------

AGORA = dt.datetime(2026, 11, 4, 12, 0, tzinfo=dt.timezone.utc)


def _imp(client, status="em_andamento", criado_em=None, produtora="untitled", competencia="2026-09",
         vigente=False):
    imps = client.tables.setdefault("importacoes", [])
    imps.append({"id": f"id-{len(imps)}", "produtora_id": produtora, "competencia": competencia,
                 "status": status, "vigente": vigente, "gatilho": "manual", "criado_em": criado_em})
    return imps[-1]["id"]


def test_encerrar_orfas_marca_so_as_antigas_em_andamento():
    c = FakeClient()
    velha = _imp(c, criado_em="2026-11-04T05:00:00+00:00")           # 7 h
    recente = _imp(c, criado_em="2026-11-04T09:00:00+00:00")         # 3 h
    concluida = _imp(c, status="concluida", criado_em="2026-10-01T00:00:00+00:00")
    vigente = _imp(c, status="concluida", criado_em="2026-09-01T00:00:00+00:00", vigente=True)
    bloqueada = _imp(c, status="bloqueada", criado_em="2026-10-01T00:00:00+00:00")
    outra_comp = _imp(c, criado_em="2026-10-01T00:00:00+00:00", competencia="2026-08")
    outra_prod = _imp(c, criado_em="2026-10-01T00:00:00+00:00", produtora="outra")
    sem_data = _imp(c, criado_em=None)
    assert Repository(c).encerrar_orfas("untitled", "2026-09", horas=6, agora=AGORA) == 1
    por_id = {r["id"]: r for r in c.tables["importacoes"]}
    assert por_id[velha]["status"] == "falhou" and "interrompida" in por_id[velha]["motivo_falha"]
    assert por_id[velha]["vigente"] is False
    assert por_id[recente]["status"] == "em_andamento"
    assert por_id[concluida]["status"] == "concluida"
    assert por_id[vigente]["status"] == "concluida" and por_id[vigente]["vigente"] is True
    assert por_id[bloqueada]["status"] == "bloqueada"
    for k in (outra_comp, outra_prod, sem_data):
        assert por_id[k]["status"] == "em_andamento"
    assert len(c.tables["importacoes"]) == 8                       # nunca apaga linhas


@pytest.mark.parametrize("criado_em", ["2026-11-04T01:00:00Z", "2026-11-04T01:00:00.123456+00:00",
                                        "2026-11-04T01:00:00", "2026-11-03T22:00:00-03:00"])
def test_encerrar_orfas_aceita_formatos_de_data(criado_em):
    c = FakeClient()
    _imp(c, criado_em=criado_em)
    assert Repository(c).encerrar_orfas("untitled", "2026-09", agora=AGORA) == 1


def test_encerrar_orfas_fuso_conta_o_instante_real():
    c = FakeClient()
    _imp(c, criado_em="2026-11-04T08:00:00-03:00")                  # 11:00 UTC: 1 h atrás
    assert Repository(c).encerrar_orfas("untitled", "2026-09", agora=AGORA) == 0


@pytest.mark.parametrize("ruim", ["ontem", "", "2026-13-45"])
def test_encerrar_orfas_ignora_data_ilegivel(ruim):
    c = FakeClient()
    _imp(c, criado_em=ruim)
    assert Repository(c).encerrar_orfas("untitled", "2026-09", agora=AGORA) == 0
    assert c.tables["importacoes"][0]["status"] == "em_andamento"


def test_encerrar_orfas_limite_exato_nao_encerra():
    c = FakeClient()
    _imp(c, criado_em="2026-11-04T06:00:00+00:00")                  # exatamente 6 h
    assert Repository(c).encerrar_orfas("untitled", "2026-09", horas=6, agora=AGORA) == 0


def test_encerrar_orfas_sem_nada_so_le():
    c = FakeClient()
    assert Repository(c).encerrar_orfas("untitled", "2026-09", agora=AGORA) == 0
    assert c.calls == [("importacoes", "select", None)]


def _importacao_vigente(repo, periodo_fim, produtora="untitled", competencia="2026-09"):
    imp = repo.criar_importacao(produtora=produtora, competencia=competencia, periodo_inicio="2026-01-01",
                                periodo_fim=periodo_fim, data_extracao="2026-10-16T10:00:00-03:00",
                                origem="teste", gatilho="manual")
    repo.promover(imp)
    return imp


def test_garantir_produtora_cria_uma_vez():
    repo = Repository(FakeClient())
    assert repo.garantir_produtora("untitled", "Untitled", "UTLD PRODUÇÃO DE FILMES LTDA") is False
    assert repo.garantir_produtora("soma", "Soma", "SOMA LTDA") is True
    assert repo.garantir_produtora("soma", "Soma", "SOMA LTDA") is False
    assert [r["id"] for r in repo.client.tables["produtoras"]] == ["untitled", "soma"]


def test_vigente_completa_so_vale_com_o_periodo_ate_o_fim_do_mes():
    repo = Repository(FakeClient())
    assert repo.vigente_completa("untitled", "2026-09", "2026-09-30") is None
    parcial = _importacao_vigente(repo, "2026-09-15")
    assert repo.vigente_completa("untitled", "2026-09", "2026-09-30") is None   # período parcial não conta
    completa = _importacao_vigente(repo, "2026-09-30")
    assert repo.vigente_completa("untitled", "2026-09", "2026-09-30") == completa
    assert completa != parcial
    assert repo.vigente_completa("untitled", "2026-08", "2026-08-31") is None


def _regras_outra():
    from papaki_dre.produtoras import Regras
    return Regras(localidade_cp="OUTRA LTDA", nucleo_rra="Projetos - Outra",
                  modelo_comissoes="Outra - Padrão", comissoes_excluir_h=("Impostos",))


def test_regras_texto_com_regras_nao_padrao_traz_os_valores_da_produtora():
    from papaki_dre.db import regras_texto
    t = regras_texto(_regras_outra())
    assert "OUTRA LTDA" in t[R_LOC]
    assert "Projetos - Outra" in t[R_NUCLEO]
    assert "Outra - Padrão" in t[R_MODELO]
    assert t[R_H].endswith("Impostos")
    assert "UTLD" not in t[R_LOC]                      # nada da Untitled vaza para outra produtora


def test_inserir_exclusoes_grava_o_texto_de_regra_recebido():
    from papaki_dre.db import regras_texto
    repo = _repo()
    imp = _new(repo)
    regras_txt = regras_texto(_regras_outra())
    s = _san("cp", [], exclusions=[Removal(5, R_LOC)], technical=[Removal(2, R_FOOTER)])
    repo.inserir_exclusoes(imp, "cp", s, regras_txt)
    por_motivo = {r["motivo"]: r["regra"] for r in repo.client.tables["exclusoes"]}
    assert por_motivo == {R_LOC: regras_txt[R_LOC], R_FOOTER: regras_txt[R_FOOTER]}
    assert "OUTRA LTDA" in por_motivo[R_LOC]
