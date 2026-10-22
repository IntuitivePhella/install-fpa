import datetime as dt

from fake_supabase import FakeClient
from papaki_dre.db import Repository

T0 = dt.datetime(2026, 11, 3, 8, 0, tzinfo=dt.timezone.utc)


def _banco():
    """(repo, client, relogio): o relogio é o 'now()' do banco falso; os testes o adiantam."""
    relogio = {"agora": T0}
    client = FakeClient()
    client.agora = lambda: relogio["agora"]
    return Repository(client), client, relogio


def test_primeira_reserva_e_obtida():
    repo, client, _ = _banco()
    r = repo.reservar_rodada("untitled", "2026-10", "principal", 6)
    assert r["reservada"] is True and r["maquina"] == "principal"
    assert dt.datetime.fromisoformat(r["expira_em"]) == T0 + dt.timedelta(hours=6)
    assert [x["status"] for x in client.tables["reservas_rodada"]] == ["ativa"]


def test_outra_maquina_nao_consegue_e_ve_quem_tem():
    repo, _, _ = _banco()
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    r = repo.reservar_rodada("untitled", "2026-10", "principal", 6)
    assert r["reservada"] is False and r["maquina"] == "grupo-a"


def test_a_mesma_maquina_renova_o_prazo():
    repo, _, relogio = _banco()
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    relogio["agora"] = T0 + dt.timedelta(hours=2)
    r = repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    assert r["reservada"] is True
    assert dt.datetime.fromisoformat(r["expira_em"]) == T0 + dt.timedelta(hours=8)


def test_reserva_vencida_deixa_outra_maquina_assumir():
    repo, client, relogio = _banco()
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    relogio["agora"] = T0 + dt.timedelta(hours=6, minutes=1)
    r = repo.reservar_rodada("untitled", "2026-10", "principal", 6)
    assert r["reservada"] is True and r["maquina"] == "principal"
    assert sorted(x["status"] for x in client.tables["reservas_rodada"]) == ["ativa", "expirada"]


def test_liberar_deixa_outra_maquina_reservar():
    repo, _, _ = _banco()
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    assert repo.liberar_rodada("untitled", "2026-10", "grupo-a") is True
    assert repo.reservar_rodada("untitled", "2026-10", "principal", 6)["reservada"] is True


def test_liberar_so_vale_para_a_propria_maquina():
    repo, _, _ = _banco()
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    assert repo.liberar_rodada("untitled", "2026-10", "principal") is False
    assert repo.reservar_rodada("untitled", "2026-10", "principal", 6)["reservada"] is False


def test_liberar_sem_reserva_devolve_falso():
    repo, _, _ = _banco()
    assert repo.liberar_rodada("untitled", "2026-10", "principal") is False


def test_produtoras_e_competencias_diferentes_sao_independentes():
    repo, _, _ = _banco()
    assert repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)["reservada"] is True
    assert repo.reservar_rodada("outra", "2026-10", "principal", 6)["reservada"] is True
    assert repo.reservar_rodada("untitled", "2026-11", "principal", 6)["reservada"] is True


def _criar(repo, **extra):
    return repo.criar_importacao(produtora="untitled", competencia="2026-10", periodo_inicio="2026-01-01",
                                 periodo_fim="2026-10-31", data_extracao="2026-11-03T10:00:00-03:00",
                                 origem="teste", gatilho="agendado", **extra)


def test_criar_importacao_grava_a_maquina():
    repo, client, _ = _banco()
    _criar(repo, maquina="grupo-a")
    assert client.tables["importacoes"][0]["maquina"] == "grupo-a"


def test_criar_importacao_sem_maquina_nao_manda_a_coluna():
    """Banco ainda sem a migração 0002: a coluna nem pode ser enviada."""
    repo, client, _ = _banco()
    _criar(repo)
    assert "maquina" not in client.tables["importacoes"][0]
