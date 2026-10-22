import datetime as dt
import json

from fake_supabase import FakeClient
from papaki_dre import cli, produtoras
from papaki_dre.db import Repository


def _situacao(capsys, client, *extra):
    codigo = cli.main(["produtora", "situacao", "--competencia", "2026-10", *extra], client=client)
    saida = capsys.readouterr()
    return codigo, (json.loads(saida.out) if saida.out.strip() else None), saida.err


def _importar(repo, maquina="grupo-a", periodo_fim="2026-10-31"):
    imp = repo.criar_importacao(produtora="untitled", competencia="2026-10", periodo_inicio="2026-01-01",
                                periodo_fim=periodo_fim, data_extracao="2026-11-03T10:00:00-03:00",
                                origem="teste", gatilho="agendado", maquina=maquina)
    repo.promover(imp)


def _item(d, produtora="untitled"):
    return next(p for p in d["produtoras"] if p["produtora"] == produtora)


def test_lista_todas_as_produtoras_cadastradas(capsys):
    codigo, d, _ = _situacao(capsys, FakeClient())
    assert codigo == 0 and d["competencia"] == "2026-10"
    assert [p["produtora"] for p in d["produtoras"]] == produtoras.ids()
    assert _item(d) == {"produtora": "untitled", "importada": False, "periodo_fim": None, "maquina": None,
                        "reserva": None}


def test_mostra_a_foto_vigente_e_quem_importou(capsys):
    client = FakeClient()
    _importar(Repository(client))
    _, d, _ = _situacao(capsys, client)
    assert _item(d) == {"produtora": "untitled", "importada": True, "periodo_fim": "2026-10-31",
                        "maquina": "grupo-a", "reserva": None}


def test_importacao_antiga_sem_maquina_mostra_nulo(capsys):
    client = FakeClient()
    _importar(Repository(client), maquina=None)
    _, d, _ = _situacao(capsys, client)
    assert _item(d)["importada"] is True and _item(d)["maquina"] is None


def test_mostra_a_reserva_ativa(capsys):
    client = FakeClient()
    Repository(client).reservar_rodada("untitled", "2026-10", "principal", 6)
    _, d, _ = _situacao(capsys, client)
    reserva = _item(d)["reserva"]
    assert reserva["maquina"] == "principal" and reserva["expira_em"]


def test_ignora_reserva_vencida_e_reserva_liberada(capsys):
    client = FakeClient()
    client.agora = lambda: dt.datetime(2020, 1, 1, tzinfo=dt.timezone.utc)     # reserva criada "no passado"
    repo = Repository(client)
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    _, d, _ = _situacao(capsys, client)
    assert _item(d)["reserva"] is None
    client.agora = lambda: dt.datetime.now(dt.timezone.utc)
    repo.reservar_rodada("untitled", "2026-10", "principal", 6)
    repo.liberar_rodada("untitled", "2026-10", "principal")
    _, d, _ = _situacao(capsys, client)
    assert _item(d)["reserva"] is None


def test_competencia_invalida(capsys):
    codigo = cli.main(["produtora", "situacao", "--competencia", "2026-13"], client=FakeClient())
    assert codigo == 3 and "2026-13" in capsys.readouterr().err
