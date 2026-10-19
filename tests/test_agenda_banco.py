import json

import pytest

from fake_supabase import FakeClient
from papaki_dre import cli
from papaki_dre.db import Repository


@pytest.fixture
def env(tmp_path):
    return {"PAPAKI_DADOS_DIR": str(tmp_path / "dados")}


def _agenda(capsys, env, client=None, *extra, hoje="2026-11-04", admin=None):
    codigo = cli.main(["agenda", "--produtora", "untitled", "--hoje", hoje, *extra],
                      env=env, client=client, admin=admin)
    saida = capsys.readouterr()
    return codigo, (json.loads(saida.out) if saida.out.strip() else None), saida.err


def _importar_vigente(client, periodo_fim="2026-10-31", competencia="2026-10"):
    repo = Repository(client)
    imp = repo.criar_importacao(produtora="untitled", competencia=competencia, periodo_inicio="2026-01-01",
                                periodo_fim=periodo_fim, data_extracao="2026-11-03T10:00:00-03:00",
                                origem="teste", gatilho="agendado")
    repo.promover(imp)


def test_banco_com_a_competencia_importada_por_outra_maquina_nao_roda(capsys, env):
    client = FakeClient()
    _importar_vigente(client)
    codigo, d, _ = _agenda(capsys, env, client)
    assert codigo == 0
    assert d["rodar"] is False and d["banco"] == "consultado" and "banco" in d["motivo"]


def test_banco_sem_a_competencia_deixa_rodar(capsys, env):
    codigo, d, _ = _agenda(capsys, env, FakeClient())
    assert codigo == 0 and d["rodar"] is True and d["banco"] == "consultado" and d["motivo"] == ""


def test_importacao_parcial_nao_conta(capsys, env):
    client = FakeClient()
    _importar_vigente(client, periodo_fim="2026-10-15")
    assert _agenda(capsys, env, client)[1]["rodar"] is True


def test_outra_competencia_nao_conta(capsys, env):
    client = FakeClient()
    _importar_vigente(client, periodo_fim="2026-09-30", competencia="2026-09")
    assert _agenda(capsys, env, client)[1]["rodar"] is True


def test_sem_chaves_do_supabase_vale_so_a_pasta_local(capsys, env):
    codigo, d, _ = _agenda(capsys, env)           # sem client e sem SUPABASE_* no env
    assert codigo == 0 and d["rodar"] is True and d["banco"] == "nao_configurado"


def test_banco_inacessivel_cai_para_a_pasta_local(capsys, env):
    codigo, d, err = _agenda(capsys, env, FakeClient(fail_with=ConnectionError("sem rede")))
    assert codigo == 0 and d["rodar"] is True and d["banco"] == "indisponivel"
    assert "sem rede" not in err


def test_sem_banco_pula_a_consulta(capsys, env):
    client = FakeClient()
    _importar_vigente(client)
    chamadas_antes = len(client.calls)
    codigo, d, _ = _agenda(capsys, env, client, "--sem-banco")
    assert d["rodar"] is True and d["banco"] == "pulado" and len(client.calls) == chamadas_antes


def test_o_banco_so_e_consultado_quando_o_calendario_manda_rodar(capsys, env):
    client = FakeClient()
    codigo, d, _ = _agenda(capsys, env, client, hoje="2026-11-02")    # feriado: não é dia útil
    assert d["rodar"] is False and d["banco"] == "pulado" and client.calls == []


def test_restaurar_banco_acorda_o_projeto_pausado_e_consulta(capsys, env):
    client = FakeClient(fail_with=ConnectionError("projeto pausado"))

    class Admin:
        chamadas = 0

        def acordar(self, espera_max, intervalo):
            Admin.chamadas += 1
            client.fail_with = None
            return "restaurado"

    codigo, d, _ = _agenda(capsys, env, client, "--restaurar-banco", admin=Admin())
    assert codigo == 0 and Admin.chamadas == 1 and d["banco"] == "consultado" and d["rodar"] is True


def test_restaurar_banco_sem_o_token_so_consulta(capsys, env):
    codigo, d, _ = _agenda(capsys, env, FakeClient(fail_with=ConnectionError("pausado")), "--restaurar-banco")
    assert codigo == 0 and d["banco"] == "indisponivel" and d["rodar"] is True
