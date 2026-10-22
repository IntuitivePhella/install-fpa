import json

import pytest

from fake_supabase import FakeClient
from papaki_dre import cli
from papaki_dre.db import Repository


@pytest.fixture
def env(tmp_path):
    return {"PAPAKI_DADOS_DIR": str(tmp_path / "dados"), "PAPAKI_MAQUINA": "principal"}


def _rodar(capsys, argv, env, client):
    codigo = cli.main(argv, env=env, client=client)
    saida = capsys.readouterr()
    return codigo, (json.loads(saida.out) if saida.out.strip() else None), saida.err


def _agenda(capsys, env, client, *extra, hoje="2026-11-04"):
    return _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", hoje, *extra], env, client)


def test_sem_reservar_o_agenda_nao_mexe_nas_reservas(capsys, env):
    client = FakeClient()
    _, d, _ = _agenda(capsys, env, client)
    assert d["reserva"] == "nao_pedida" and d["rodar"] is True
    assert client.tables.get("reservas_rodada", []) == []


def test_reservar_com_a_rodada_livre_obtem_a_reserva(capsys, env):
    client = FakeClient()
    codigo, d, _ = _agenda(capsys, env, client, "--reservar")
    assert codigo == 0 and d["rodar"] is True and d["reserva"] == "obtida"
    reservas = client.tables["reservas_rodada"]
    assert [(r["produtora_id"], r["competencia"], r["maquina"], r["status"]) for r in reservas] == [
        ("untitled", "2026-10", "principal", "ativa")]


def test_reserva_de_outra_maquina_manda_nao_rodar(capsys, env):
    client = FakeClient()
    Repository(client).reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    codigo, d, _ = _agenda(capsys, env, client, "--reservar")
    assert codigo == 0 and d["rodar"] is False and d["reserva"] == "de_outra_maquina"
    assert "outra máquina" in d["motivo"]


def test_a_propria_reserva_deixa_rodar_de_novo(capsys, env):
    client = FakeClient()
    _agenda(capsys, env, client, "--reservar")
    _, d, _ = _agenda(capsys, env, client, "--reservar")
    assert d["rodar"] is True and d["reserva"] == "obtida"


def test_sem_nome_da_maquina_segue_e_avisa(capsys, tmp_path):
    env = {"PAPAKI_DADOS_DIR": str(tmp_path / "dados")}
    client = FakeClient()
    _, d, _ = _agenda(capsys, env, client, "--reservar")
    assert d["rodar"] is True and d["reserva"] == "sem_nome_da_maquina"
    assert client.tables.get("reservas_rodada", []) == []


def test_banco_indisponivel_segue_e_avisa(capsys, env):
    _, d, err = _agenda(capsys, env, FakeClient(fail_with=ConnectionError("sem rede")), "--reservar")
    assert d["rodar"] is True and d["banco"] == "indisponivel" and d["reserva"] == "nao_consultada"
    assert "sem rede" not in err


def test_sem_banco_nao_reserva(capsys, env):
    client = FakeClient()
    _, d, _ = _agenda(capsys, env, client, "--reservar", "--sem-banco")
    assert d["banco"] == "pulado" and d["reserva"] == "nao_consultada" and d["rodar"] is True


def test_competencia_ja_importada_nao_reserva(capsys, env):
    client = FakeClient()
    repo = Repository(client)
    imp = repo.criar_importacao(produtora="untitled", competencia="2026-10", periodo_inicio="2026-01-01",
                                periodo_fim="2026-10-31", data_extracao="2026-11-03T10:00:00-03:00",
                                origem="teste", gatilho="agendado")
    repo.promover(imp)
    _, d, _ = _agenda(capsys, env, client, "--reservar")
    assert d["rodar"] is False and d["reserva"] == "nao_pedida"
    assert client.tables.get("reservas_rodada", []) == []


def test_fora_do_dia_de_rodar_nao_reserva(capsys, env):
    client = FakeClient()
    _, d, _ = _agenda(capsys, env, client, "--reservar", hoje="2026-11-02")      # feriado
    assert d["rodar"] is False and d["reserva"] == "nao_pedida"
    assert client.calls == []


def test_reserva_obter_para_a_competencia(capsys, env):
    client = FakeClient()
    codigo, d, _ = _rodar(capsys, ["reserva", "obter", "--produtora", "untitled", "--competencia", "2026-10"],
                          env, client)
    assert codigo == 0 and d["estado"] == "obtida" and d["maquina"] == "principal"
    assert d["produtora"] == "untitled" and d["competencia"] == "2026-10" and d["expira_em"]


def test_reserva_obter_mostra_quem_tem_a_reserva(capsys, env):
    client = FakeClient()
    Repository(client).reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    _, d, _ = _rodar(capsys, ["reserva", "obter", "--produtora", "untitled", "--competencia", "2026-10"],
                     env, client)
    assert d["estado"] == "de_outra_maquina" and d["maquina"] == "grupo-a"


@pytest.mark.parametrize("argv, trecho", [
    (["reserva", "obter", "--produtora", "naoexiste", "--competencia", "2026-10"], "produtora desconhecida"),
    (["reserva", "obter", "--produtora", "untitled", "--competencia", "2026-13"], "2026-13"),
])
def test_reserva_obter_erros_de_uso(capsys, env, argv, trecho):
    codigo, d, err = _rodar(capsys, argv, env, FakeClient())
    assert codigo == 3 and d is None and trecho in err
