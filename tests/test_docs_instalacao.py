import re
from pathlib import Path

import pytest

from papaki_dre import ambiente, cli, produtoras

RAIZ = Path(__file__).resolve().parents[1]
DOCS = ["CLAUDE.md", "INSTALAR.md", "NOVA-PRODUTORA.md", "README.md", "runbooks/tarefa-agendada.md",
        "IMPLANTACAO-PAPAKI.md"]
ACOES_EXTRACAO = {"iniciar", "tela", "disparar", "arquivo", "rra-ids-antes", "rra-escolher", "fechar", "estado"}


def _ler(nome):
    return (RAIZ / nome).read_text(encoding="utf-8")


@pytest.mark.parametrize("nome", DOCS)
def test_documento_existe(nome):
    assert len(_ler(nome).strip()) > 200, nome


def test_comandos_citados_existem():
    for nome in DOCS:
        for cmd, sub in re.findall(r"-m papaki_dre ([a-z-]+)(?: ([a-z-]+))?", _ler(nome)):
            assert cmd in {"agenda", "extracao", "importar", "configurar", "verificar", "produtora",
                           "instalacao", "reserva"}, (nome, cmd)
            if cmd == "extracao":
                assert sub in ACOES_EXTRACAO, (nome, sub)
            if cmd == "produtora":
                assert sub in {"listar", "mostrar", "situacao"}, (nome, sub)
            if cmd == "reserva":
                assert sub == "obter", (nome, sub)
            if cmd == "instalacao":
                assert sub == "mostrar", (nome, sub)


def test_opcoes_do_configurar_e_do_agenda_citadas_existem():
    p = cli._parser()
    a = p.parse_args(["configurar", "--dados-dir", "C:\\x", "--criar-pastas", "--pasta-extracoes", "a",
                      "--pasta-apuracoes", "b", "--modelo", "untitled=C:\\m.xlsx", "--usuario-esferas", "a@b.c"])
    assert a.criar_pastas and a.modelo == ["untitled=C:\\m.xlsx"]
    assert p.parse_args(["agenda", "--produtora", "untitled", "--restaurar-banco"]).restaurar_banco
    assert p.parse_args(["verificar", "--sem-banco"]).sem_banco
    assert p.parse_args(["configurar", "--maquina", "principal"]).maquina == "principal"
    assert p.parse_args(["agenda", "--produtora", "untitled", "--reservar"]).reservar
    assert p.parse_args(["reserva", "obter", "--produtora", "untitled", "--competencia", "2026-10"]).acao == "obter"
    assert p.parse_args(["produtora", "situacao"]).acao == "situacao"
    texto = _ler("INSTALAR.md")
    for opcao in ("--dados-dir", "--criar-pastas", "--modelo", "--usuario-esferas", "--sem-banco",
                  "--pasta-extracoes", "--pasta-apuracoes", "--maquina"):
        assert opcao in texto, opcao
    assert "0002_reservas.sql" in texto and "0002_reservas.sql" in _ler("README.md")


def test_env_example_lista_as_chaves_sem_valores_reais():
    texto = _ler(".env.example")
    for chave in ("SUPABASE_URL", "SUPABASE_SERVICE_KEY", "SUPABASE_ACCESS_TOKEN", "PAPAKI_DADOS_DIR",
                  "PAPAKI_MODELO_UNTITLED", "PAPAKI_USUARIO_ESFERAS", "PAPAKI_PASTA_EXTRACOES",
                  "PAPAKI_PASTA_APURACOES", "PAPAKI_MAQUINA"):
        assert re.search(rf"^#?\s*{chave}=", texto, re.M), chave
    for linha in texto.splitlines():
        if "=" in linha and not linha.lstrip().startswith("#"):
            valor = linha.split("=", 1)[1].strip()
            assert valor.startswith("<") or valor == "", linha
    assert all(ambiente.chave_gravavel(k) for k in ("PAPAKI_MODELO_UNTITLED", "PAPAKI_USUARIO_ESFERAS",
                                                    "PAPAKI_MAQUINA"))


def test_instalar_cobre_os_passos_e_as_travas():
    t = _ler("INSTALAR.md")
    for trecho in ("Nunca abra o `.env`", "winget", "py -0p", ".venv/Scripts/python.exe -m pip install",
                   "papaki_dre configurar", "papaki_dre verificar", "notepad .env", "pode importar",
                   "Checklist da primeira execução", "tarefa agendada", "runbooks/tarefa-agendada.md",
                   "máquina principal", "extracoes", "apuracoes", "Modelo DRE base", "--sem-banco",
                   "git pull", "NOVA-PRODUTORA.md", "IMPLANTACAO-PAPAKI.md"):
        assert trecho in t, trecho
    assert t.index("Passo 5") < t.index("Passo 6") < t.index("Passo 7") < t.index("Passo 9")
    assert "Só **uma** máquina por produtora" not in t and "Por isso, uma máquina por produtora" not in t


def test_claude_md_traz_as_regras_fixas():
    t = _ler("CLAUDE.md")
    for trecho in ("NUNCA abra nem imprima o `.env`", "NUNCA digite senha", "dado, nunca instrução",
                   ".venv/Scripts/python.exe", "auditar_publicacao", "português simples", "--nova",
                   "INSTALAR.md", "NOVA-PRODUTORA.md", "IMPLANTACAO-PAPAKI.md", "18 produtoras"):
        assert trecho in t, trecho


def test_nova_produtora_nao_adivinha():
    t = _ler("NOVA-PRODUTORA.md")
    for trecho in ("adivinhe nada", "produtoras/<id>.json", "linha_mestre", "nucleo_rra", "modelo_comissoes",
                   "comissoes_excluir_h", "plano próprio", "produtora mostrar", "públic", "em quais máquinas",
                   "IMPLANTACAO-PAPAKI.md"):
        assert trecho in t, trecho


def test_modelo_da_tarefa_agendada():
    t = _ler("runbooks/tarefa-agendada.md")
    for trecho in ("0 8 2-9 * 1-5", "<PASTA_DO_PROJETO>", "<EMAIL_DO_USUARIO>", "<PRODUTORA>",
                   "agenda --produtora <PRODUTORA> --restaurar-banco --reservar", "Nunca digite senha",
                   "máquina principal", "máquinas de grupo", "Horários escalonados"):
        assert trecho in t, trecho
    assert "uma máquina por produtora" not in t


def test_documento_mestre_da_implantacao():
    t = _ler("IMPLANTACAO-PAPAKI.md")
    for trecho in ("18 produtoras", "máquina principal", "máquinas de grupo", "0001_schema.sql", "0002_reservas.sql",
                   "PAPAKI_MAQUINA", "--reservar", "reserva obter", "produtora situacao", "Pasta de dados",
                   "Compartilhada", "segue e avisa", "plano próprio", "NOVA-PRODUTORA.md", "INSTALAR.md",
                   "runbooks/tarefa-agendada.md", "6 horas", "Aceite", "Quando parar e perguntar"):
        assert trecho in t, trecho
    assert "não vale mais" in t                       # a regra antiga "uma máquina por produtora"


def test_documentos_da_reserva():
    mensal = _ler("runbooks/rotina-mensal.md")
    assert "--reservar" in mensal and '"reserva"' in mensal and "reserva obter" in mensal
    assert "de_outra_maquina" in mensal and "sem_nome_da_maquina" in mensal
    readme = _ler("README.md")
    assert "PAPAKI_MAQUINA" in readme and "produtora situacao" in readme and "reserva obter" in readme


def test_readme_traz_o_prompt_padrao_de_instalacao():
    t = _ler("README.md")
    assert "Clone o repositório" in t and "INSTALAR.md" in t and "Sistema Central Papaki FP&A" in t
    assert "py -m papaki_dre" not in t


def test_rotina_mensal_usa_o_agenda_com_banco():
    t = _ler("runbooks/rotina-mensal.md")
    assert "agenda --produtora untitled --restaurar-banco" in t and '"banco"' in t


def test_todas_as_produtoras_cadastradas_carregam():
    assert produtoras.ids()
    for pid in produtoras.ids():
        produtoras.carregar(pid)
