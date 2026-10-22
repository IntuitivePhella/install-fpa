import sys
from pathlib import Path

import pytest

from fake_supabase import FakeClient
from papaki_dre import cli, verificacao
from papaki_dre.verificacao import AVISO, FALHA, OK, Item, render, verificar


@pytest.fixture
def projeto(tmp_path):
    p = tmp_path / "projeto"
    (p / ".venv" / "Scripts").mkdir(parents=True)
    return p


@pytest.fixture
def modelo(tmp_path):
    arq = tmp_path / "Modelo.xlsx"
    arq.write_bytes(b"x")
    return arq


def _env(tmp_path, modelo, **mais):
    dados = tmp_path / "dados"
    (dados / "extracoes").mkdir(parents=True, exist_ok=True)
    (dados / "apuracoes").mkdir(parents=True, exist_ok=True)
    e = {"PAPAKI_DADOS_DIR": str(dados), "PAPAKI_MODELO_UNTITLED": str(modelo),
         "PAPAKI_USUARIO_ESFERAS": "ana@exemplo.com", "PAPAKI_MAQUINA": "principal", "SUPABASE_URL": "https://inventado.supabase.co",
         "SUPABASE_SERVICE_KEY": "chave-inventada", "SUPABASE_ACCESS_TOKEN": "token-inventado"}
    e.update(mais)
    return e


def _rodar(tmp_path, modelo, projeto, env=None, **kw):
    kw.setdefault("excel_registrado", lambda: True)
    kw.setdefault("importar", lambda nome: object())
    kw.setdefault("plataforma", "win32")
    kw.setdefault("versao", (3, 12, 1))
    kw.setdefault("executavel", projeto / ".venv" / "Scripts" / "python.exe")
    kw.setdefault("git", lambda: True)
    kw.setdefault("client", FakeClient())
    return verificar(env if env is not None else _env(tmp_path, modelo), projeto=projeto, **kw)


def _por_nome(itens):
    return {i.nome: i for i in itens}


def test_tudo_certo(tmp_path, modelo, projeto):
    itens = _rodar(tmp_path, modelo, projeto)
    assert [i for i in itens if i.nivel != OK] == [], itens


def test_sem_banco_nao_consulta_nem_exige_as_chaves(tmp_path, modelo, projeto):
    env = _env(tmp_path, modelo)
    del env["SUPABASE_URL"], env["SUPABASE_SERVICE_KEY"]
    itens = _rodar(tmp_path, modelo, projeto, env=env, com_banco=False, client=None)
    assert "banco" not in _por_nome(itens)
    assert not [i for i in itens if i.nivel == FALHA]


def test_python_velho_e_fora_da_venv(tmp_path, modelo, projeto):
    itens = _por_nome(_rodar(tmp_path, modelo, projeto, versao=(3, 10, 0), executavel=Path(sys.executable)))
    assert itens["python"].nivel == FALHA and "3.11" in itens["python"].detalhe
    assert itens["venv"].nivel == AVISO and ".venv/Scripts/python.exe" in itens["venv"].detalhe


def test_sistema_que_nao_e_windows(tmp_path, modelo, projeto):
    assert _por_nome(_rodar(tmp_path, modelo, projeto, plataforma="linux"))["sistema"].nivel == FALHA


def test_biblioteca_faltando_e_excel_ausente(tmp_path, modelo, projeto):
    def importar(nome):
        if nome == "lxml":
            raise ImportError(nome)
        return object()
    itens = _por_nome(_rodar(tmp_path, modelo, projeto, importar=importar, excel_registrado=lambda: False))
    assert itens["bibliotecas"].nivel == FALHA and "lxml" in itens["bibliotecas"].detalhe
    assert itens["excel"].nivel == FALHA


def test_git_ausente_so_avisa(tmp_path, modelo, projeto):
    assert _por_nome(_rodar(tmp_path, modelo, projeto, git=lambda: False))["git"].nivel == AVISO


def test_variaveis_do_env_faltando(tmp_path, modelo, projeto):
    env = _env(tmp_path, modelo)
    for k in ("SUPABASE_URL", "SUPABASE_ACCESS_TOKEN", "PAPAKI_USUARIO_ESFERAS"):
        del env[k]
    itens = _por_nome(_rodar(tmp_path, modelo, projeto, env=env))
    assert itens["SUPABASE_URL"].nivel == FALHA
    assert itens["SUPABASE_ACCESS_TOKEN"].nivel == AVISO
    assert itens["PAPAKI_USUARIO_ESFERAS"].nivel == AVISO


def test_pasta_de_dados_ausente_dentro_do_projeto_ou_inexistente(tmp_path, modelo, projeto):
    env = _env(tmp_path, modelo)
    del env["PAPAKI_DADOS_DIR"]
    assert _por_nome(_rodar(tmp_path, modelo, projeto, env=env))["dados"].nivel == FALHA
    env["PAPAKI_DADOS_DIR"] = str(projeto / "dados")
    item = _por_nome(_rodar(tmp_path, modelo, projeto, env=env))["dados"]
    assert item.nivel == FALHA and "dentro da pasta do projeto" in item.detalhe
    env["PAPAKI_DADOS_DIR"] = str(tmp_path / "nao_existe")
    item = _por_nome(_rodar(tmp_path, modelo, projeto, env=env))["dados"]
    assert item.nivel == FALHA and "configurar --criar-pastas" in item.detalhe


def test_subpastas_que_ainda_nao_existem_so_avisam(tmp_path, modelo, projeto):
    env = _env(tmp_path, modelo, PAPAKI_PASTA_EXTRACOES="Brutos")
    itens = _por_nome(_rodar(tmp_path, modelo, projeto, env=env))
    assert itens["dados"].nivel == OK and itens["subpastas"].nivel == AVISO


def test_modelo_configurado_inexistente_falha_e_sem_modelo_nenhum_falha(tmp_path, modelo, projeto):
    env = _env(tmp_path, modelo, PAPAKI_MODELO_UNTITLED=str(tmp_path / "sumiu.xlsx"))
    assert _por_nome(_rodar(tmp_path, modelo, projeto, env=env))["modelo untitled"].nivel == FALHA
    del env["PAPAKI_MODELO_UNTITLED"]
    itens = _por_nome(_rodar(tmp_path, modelo, projeto, env=env))
    assert itens["modelo untitled"].nivel == AVISO and itens["modelos"].nivel == FALHA


def test_banco_inacessivel(tmp_path, modelo, projeto):
    itens = _por_nome(_rodar(tmp_path, modelo, projeto,
                             client=FakeClient(fail_with=ConnectionError("sem rede"))))
    assert itens["banco"].nivel == FALHA and "pausado" in itens["banco"].detalhe


def test_banco_sem_tabelas(tmp_path, modelo, projeto):
    class SemTabela(FakeClient):
        def table(self, name):
            if name == "cargas_excel":
                raise RuntimeError("relation does not exist")
            return super().table(name)
    item = _por_nome(_rodar(tmp_path, modelo, projeto, client=SemTabela()))["banco"]
    assert item.nivel == FALHA and "cargas_excel" in item.detalhe and "0001_schema.sql" in item.detalhe


def test_banco_erro_que_nao_e_tabela_ausente_nao_sugere_reaplicar_o_schema(tmp_path, modelo, projeto):
    class Cai(FakeClient):
        def table(self, name):
            if name == "cargas_excel":
                raise ConnectionError("queda")
            return super().table(name)
    item = _por_nome(_rodar(tmp_path, modelo, projeto, client=Cai()))["banco"]
    assert item.nivel == FALHA and "ConnectionError" in item.detalhe
    assert "0001_schema.sql" not in item.detalhe and "queda" not in item.detalhe


def test_banco_tabela_ausente_por_codigo_do_postgrest(tmp_path, modelo, projeto):
    class Erro(Exception):
        code = "PGRST205"

    class SemTabela(FakeClient):
        def table(self, name):
            if name == "comissoes":
                raise Erro("x")
            return super().table(name)
    item = _por_nome(_rodar(tmp_path, modelo, projeto, client=SemTabela()))["banco"]
    assert item.nivel == FALHA and "comissoes" in item.detalhe and "0001_schema.sql" in item.detalhe


def test_nenhum_detalhe_revela_valores(tmp_path, modelo, projeto):
    env = _env(tmp_path, modelo)
    texto = render(_rodar(tmp_path, modelo, projeto, env=env))
    for segredo in ("chave-inventada", "token-inventado", "inventado.supabase.co", "ana@exemplo.com",
                    str(tmp_path)):
        assert segredo not in texto


def test_render_resume(capsys):
    itens = [Item(OK, "a", "tudo bem"), Item(AVISO, "b", "atenção"), Item(FALHA, "c", "quebrou")]
    texto = render(itens)
    assert "OK     a: tudo bem" in texto and "FALHA  c: quebrou" in texto
    assert texto.rstrip().endswith("RESULTADO: 1 ok, 1 aviso(s), 1 falha(s)")


def test_comando_verificar_codigo_de_saida(capsys, tmp_path, modelo):
    codigo = cli.main(["verificar", "--sem-banco"], env={"PAPAKI_DADOS_DIR": str(tmp_path / "x")})
    saida = capsys.readouterr().out
    assert codigo == 3 and "FALHA" in saida and "RESULTADO:" in saida


def test_verificar_por_padrao_consulta_o_banco(monkeypatch, capsys, tmp_path, modelo):
    chamadas = []
    monkeypatch.setattr(verificacao, "verificar",
                        lambda env, **kw: chamadas.append(kw) or [Item(OK, "x", "ok")])
    assert cli.main(["verificar"], env={}) == 0
    assert chamadas[0]["com_banco"] is True
    assert cli.main(["verificar", "--sem-banco"], env={}) == 0
    assert chamadas[1]["com_banco"] is False


def test_sem_nome_da_maquina_avisa(tmp_path, modelo, projeto):
    env = _env(tmp_path, modelo)
    del env["PAPAKI_MAQUINA"]
    item = _por_nome(_rodar(tmp_path, modelo, projeto, env=env))["PAPAKI_MAQUINA"]
    assert item.nivel == AVISO and "--maquina" in item.detalhe and "principal" not in item.detalhe


def test_tabela_de_reservas_presente_e_ok(tmp_path, modelo, projeto):
    assert _por_nome(_rodar(tmp_path, modelo, projeto))["reservas"].nivel == OK


def test_tabela_de_reservas_ausente_e_so_aviso(tmp_path, modelo, projeto):
    class SemReservas(FakeClient):
        def table(self, name):
            if name == "reservas_rodada":
                e = Exception("relation does not exist")
                e.code = "42P01"
                raise e
            return super().table(name)

    itens = _por_nome(_rodar(tmp_path, modelo, projeto, client=SemReservas()))
    assert itens["reservas"].nivel == AVISO and "0002_reservas.sql" in itens["reservas"].detalhe
    assert itens["banco"].nivel == OK                      # as 9 tabelas da 0001 continuam o que decide o banco
