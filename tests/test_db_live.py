"""Teste ao vivo, opcional, contra o Supabase real. Só roda com PAPAKI_LIVE=1 e credenciais válidas.

Usa apenas dados SINTÉTICOS na competência '2099-01' da produtora 'untitled' e apaga tudo no fim
(o `on delete cascade` das tabelas filhas leva o resto). Nunca imprime URL nem chave.
    PAPAKI_LIVE=1 py -m pytest tests/test_db_live.py -v -s
"""
import os

import pytest

from papaki_dre.config import R_FOOTER, R_LOC
from papaki_dre.db import Repository, make_client
from papaki_dre.types import Removal, Sanitized

pytestmark = pytest.mark.live

COMP = "2099-01"
PROD = "untitled"


@pytest.fixture
def client():
    if os.environ.get("PAPAKI_LIVE") != "1":
        pytest.skip("teste ao vivo: defina PAPAKI_LIVE=1 para rodar")
    try:
        c = make_client()
    except Exception as e:  # noqa: BLE001
        pytest.skip(f"sem cliente do Supabase ({type(e).__name__})")
    return c


def _limpar(client):
    client.table("importacoes").delete().eq("competencia", COMP).execute()


def _nova(repo):
    return repo.criar_importacao(produtora=PROD, competencia=COMP, periodo_inicio="2099-01-01",
                                 periodo_fim="2099-01-31", data_extracao="2099-02-01T12:00:00+00:00",
                                 origem="teste ao vivo (sintético)", gatilho="manual")


def test_live_verificar_acesso(client):
    Repository(client).verificar_acesso()
    print("conectou")


def test_live_full_round_trip(client):
    repo = Repository(client)
    repo.verificar_acesso()
    _limpar(client)                      # sobra de uma execução interrompida
    try:
        a, b = _nova(repo), _nova(repo)
        hashes = {"cp": "sintetico-hash-cp-a"}
        repo.registrar_arquivo(a, "cp", nome="cp_sintetico.csv", sha256=hashes["cp"], tamanho=123)
        s = Sanitized(report="cp", header=["Cód.", "Descrição", "Valor"],
                      rows=[], db_rows=[["teste-1", "item inventado", 1.5],
                                        ["teste-2", "outro item inventado", 2.0],
                                        ["teste-3", "mais um", -3.25]],
                      source_rows=[2, 3, 5], raw_count=5,
                      technical=[Removal(6, R_FOOTER)], exclusions=[Removal(4, R_LOC)])
        repo.inserir_linhas(a, PROD, COMP, "cp", s)
        repo.inserir_exclusoes(a, "cp", s)
        repo.registrar_carga(a, aba="DESPESAS", primeira=3, ultima=5, arquivo="sintetico.xlsx", sha256="zz")

        linhas = client.table("cp_despesas").select("cod,linha_destino,dados").eq("importacao_id", a).execute().data
        assert sorted(r["cod"] for r in linhas) == ["teste-1", "teste-2", "teste-3"]
        assert sorted(r["linha_destino"] for r in linhas) == [3, 4, 5]
        exc = client.table("exclusoes").select("tipo,motivo").eq("importacao_id", a).execute().data
        assert sorted((r["tipo"], r["motivo"]) for r in exc) == [("negocio", R_LOC), ("tecnica", R_FOOTER)]

        assert repo.find_identical(PROD, COMP, hashes) is None            # ainda em andamento
        repo.promover(a)
        repo.promover(b)
        imps = {r["id"]: r for r in client.table("importacoes").select("id,status,vigente")
                .eq("competencia", COMP).execute().data}
        assert (imps[a]["status"], imps[a]["vigente"]) == ("concluida", False)
        assert (imps[b]["status"], imps[b]["vigente"]) == ("concluida", True)
        assert repo.find_identical(PROD, COMP, hashes) == a              # concluída, mesmo não vigente

        assert repo.marcar(b, "falhou", motivo="teste: não pode desfazer a vigente") is False
        row = client.table("importacoes").select("status,vigente,motivo_falha").eq("id", b).execute().data[0]
        assert (row["status"], row["vigente"], row["motivo_falha"]) == ("concluida", True, None)
    except BaseException:
        # limpeza de melhor esforço que nunca mascara a falha original
        try:
            _limpar(client)
        except Exception:  # noqa: BLE001
            pass
        raise
    _limpar(client)
    assert client.table("importacoes").select("id").eq("competencia", COMP).execute().data == []
