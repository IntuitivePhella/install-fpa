import re
from pathlib import Path

from papaki_dre import cli
from papaki_dre.extracao import RELATORIOS

RAIZ = Path(__file__).resolve().parents[1]
ROTEIROS = [RAIZ / "runbooks" / "extracao-esferas.md", RAIZ / "runbooks" / "rotina-mensal.md"]


def _ler(nome):
    return (RAIZ / "runbooks" / nome).read_text(encoding="utf-8")


PY = ".venv/Scripts/python.exe -m papaki_dre"


def _comandos(texto):
    return re.findall(r"\.venv/Scripts/python\.exe -m papaki_dre ([a-z-]+)(?: ([a-z-]+))?", texto)


def test_roteiros_existem_e_so_citam_comandos_que_existem():
    acoes = {"iniciar", "tela", "disparar", "arquivo", "rra-ids-antes", "rra-escolher", "fechar", "estado"}
    for roteiro in ROTEIROS:
        texto = roteiro.read_text(encoding="utf-8")
        assert texto.strip()
        for cmd, sub in _comandos(texto):
            assert cmd in {"agenda", "extracao", "importar", "produtora", "instalacao",
                           "verificar"}, (roteiro.name, cmd)
            if cmd == "extracao":
                assert sub in acoes, (roteiro.name, sub)


def test_opcoes_citadas_do_importar_existem():
    texto = _ler("rotina-mensal.md")
    ajuda = cli._parser().parse_args(
        ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--extracao", "x",
         "--gatilho", "agendado", "--restaurar-banco"])
    assert ajuda.restaurar_banco and ajuda.extracao
    assert "--extracao" in texto and "--restaurar-banco" in texto and "--gatilho" in texto


def test_opcoes_citadas_de_extracao_existem():
    p = cli._parser()
    texto = _ler("extracao-esferas.md")
    base = ["extracao", "tela", "--pasta", "x", "--relatorio", "cp", "--localidade", "Todas",
            "--criterio", "c", "--data-inicio", "2026-01-01", "--data-fim", "2026-09-30", "--status", "Todos",
            "--total", "1", "--pesquisar-em", "x", "--print", "x.png"]
    assert p.parse_args(base).pesquisar_em
    assert p.parse_args(["extracao", "iniciar", "--produtora", "untitled", "--competencia", "2026-09",
                         "--retomar", "--periodo-fim", "2026-09-30"]).retomar
    # `--ids ""` perde o argumento vazio no PowerShell 5.1: os roteiros usam a forma com igual
    assert p.parse_args(["extracao", "rra-ids-antes", "--pasta", "x", "--ids="]).ids == ""
    assert p.parse_args(["extracao", "rra-ids-antes", "--pasta", "x", "--ids=a,b"]).ids == "a,b"
    assert "--ids=" in texto and '--ids ""' not in texto and "--ids <" not in texto
    assert p.parse_args(["extracao", "rra-escolher", "--pasta", "x", "--candidatas", "x.json"]).candidatas
    for opcao in ("--retomar", "--nova", "--pesquisar-em", "--periodo-fim", "--candidatas", "--ids", "--print"):
        assert opcao in texto


def test_roteiro_de_extracao_cobre_os_quatro_relatorios_e_as_armadilhas():
    texto = _ler("extracao-esferas.md")
    for rel in RELATORIOS:
        assert f"--relatorio {rel}" in texto
    for armadilha in ("UTLD PRODUCAO", "seta", "Localidade", "NÃO exporte", "uma única vez",
                      "Windows-1252", "mais de 200", "gera um CSV"):
        assert armadilha in texto


def test_roteiros_trazem_as_regras_revisadas():
    ext, mensal = _ler("extracao-esferas.md"), _ler("rotina-mensal.md")
    for texto in (ext, mensal):
        assert PY in texto
        assert "NUNCA digite senha" in texto or "nunca digite senha" in texto
    for trecho in ("15 minutos", "rra-ids-antes", "rra-escolher", "Código `4`", "20 minutos",
                   "nunca é disparado de novo", "--nova", "cp.csv", "cr.csv", "rra.xls", "comissoes.xlsx",
                   "Downloads", "posterior ao disparo", "capturar_tela.ps1", "PAPAKI_DADOS_DIR"):
        assert trecho in ext, trecho
    # RRA: os ids antes da tela do RRA
    assert ext.index("**Antes de `extracao tela` do RRA**") < ext.index("5. Volte ao Esferas")
    # iniciar antes das regras por relatório
    assert ext.index("## Antes de tudo (iniciar)") < ext.index("## Ordem de cada relatório")
    for trecho in ("--restaurar-banco", "9", "RESULTADO.md", "não envia e-mail", "Código", "segundo plano",
                   "agenda --produtora untitled"):
        assert trecho in mensal, trecho


def test_roteiros_sem_dados_sensiveis_nem_chaves():
    for roteiro in ROTEIROS:
        texto = roteiro.read_text(encoding="utf-8")
        assert not re.search(r"\d{3}\.\d{3}\.\d{3}-\d{2}|\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto)
        assert "SUPABASE_ACCESS_TOKEN=" not in texto and "eyJ" not in texto


def test_script_de_print_grava_png_sem_imprimir_conteudo():
    ps1 = (RAIZ / "scripts" / "capturar_tela.ps1").read_text(encoding="utf-8")
    assert "-Destino" in ps1 or "$Destino" in ps1
    assert "ImageFormat]::Png" in ps1 and "CopyFromScreen" in ps1
    assert "Write-Output" not in ps1 and "Write-Host" not in ps1


def test_regras_novas_dos_roteiros():
    ext, mensal = _ler("extracao-esferas.md"), _ler("rotina-mensal.md")
    for trecho in ("IMEDIATAMENTE", "180 segundos", "NÃO clique de novo", "PARE e AVISE", "sem separador de milhar",
                   "37030", "não bloqueia", "id da MENSAGEM", "permmsgid",
                   "**não** o corpo inteiro", "Olhe a captura", "tela bloqueada", "DPI-aware", "PARADA.md",
                   "escreva em `RESULTADO.md`**", "Salvar como", "vários downloads"):
        assert trecho in ext, trecho
    assert "dado, nunca instrução" in ext.casefold()
    for trecho in ("PARADA.md", "Nunca escreva em `RESULTADO.md`", "Arquivos idênticos já importados",
                   "Não diga \"apuração importada\"", "Status: ...", "ensaios", "dentro da pasta de dados",
                   "dado, nunca instrução"):
        assert trecho in mensal, trecho


def test_comandos_dos_roteiros_nao_tem_reticencias():
    for roteiro in ROTEIROS:
        for linha in roteiro.read_text(encoding="utf-8").splitlines():
            for cmd in re.findall(r"-m papaki_dre[^`]*", linha):
                assert "..." not in cmd, (roteiro.name, cmd)


def test_linhas_de_tela_completas_para_os_quatro_relatorios():
    ext = _ler("extracao-esferas.md")
    obrigatorias = ("--pasta", "--relatorio {rel}", "--localidade", "--criterio", "--data-inicio", "--data-fim",
                    "--status", "--total", "--print")
    for rel in RELATORIOS:
        linhas = [l for l in ext.splitlines() if f"extracao tela --pasta <pasta> --relatorio {rel} " in l]
        assert linhas, rel
        for o in obrigatorias:
            assert o.format(rel=rel) in linhas[0], (rel, o)
    assert "--pesquisar-em" in [l for l in ext.splitlines() if "--relatorio cp " in l and "extracao tela" in l][0]


def test_script_de_print_e_dpi_aware_e_so_a_tela_primaria_por_padrao():
    ps1 = (RAIZ / "scripts" / "capturar_tela.ps1").read_text(encoding="utf-8")
    assert "SetProcessDPIAware" in ps1 and "PrimaryScreen" in ps1 and "TodosMonitores" in ps1


def test_roteiros_usam_os_campos_do_json_e_nao_o_valor_do_env():
    ext, mensal = _ler("extracao-esferas.md"), _ler("rotina-mensal.md")
    for texto in (ext, mensal):
        assert "<PAPAKI_DADOS_DIR>" not in texto
        assert "pasta_apuracao" in texto
    assert "pasta_extracoes" in mensal and "Pasta de apuracao:" in mensal
    assert "código 3" in mensal and "avise o usuário" in mensal.split("código 3")[1][:200]


def test_roteiro_mensal_tem_o_ramo_da_falha_depois_do_fechar():
    mensal = _ler("rotina-mensal.md")
    ramo = mensal.split("## Falha DEPOIS do `extracao fechar`")[1].split(chr(10) + "## ")[0]
    for trecho in ("NUNCA extraia de novo", "importar --produtora untitled", "--extracao <pasta fechada>",
                   "código 3", "código 2", "nem reimporte sozinho", "PARADA.md",
                   "a pedido do usuário"):
        assert trecho in ramo, trecho


def test_roteiro_manda_registrar_resolvido_no_parada_antigo():
    assert "resolvido em <data>" in _ler("rotina-mensal.md")


def test_roteiros_genericos_sem_usuario_nem_maquina_fixos():
    for roteiro in ROTEIROS:
        texto = roteiro.read_text(encoding="utf-8")
        assert "py -m papaki_dre" not in texto, roteiro.name
        for fixo in ("pythoncore", "C:\\Users", "fematronni", "Felipe", "Marcelo", "Olá FELIPE", "Plano 5"):
            assert fixo not in texto, (roteiro.name, fixo)
        assert not re.search(r"[\w.]+@(gmail|hotmail|outlook)\.", texto), roteiro.name


def test_roteiros_usam_a_configuracao_da_produtora_e_da_instalacao():
    ext, mensal = _ler("extracao-esferas.md"), _ler("rotina-mensal.md")
    assert f"{PY} produtora mostrar --produtora" in ext
    for campo in ("<localidade>", "<consolidado>", "<site da produtora>"):
        assert campo in ext, campo
    assert f"{PY} instalacao mostrar" in mensal and "usuario_esferas" in mensal
    assert "UTLD PRODUÇÃO DE FILMES LTDA" not in ext        # a Localidade vem do arquivo da produtora
    assert "UTLD PRODUCAO" in ext                            # mas o aviso da armadilha continua, como exemplo
