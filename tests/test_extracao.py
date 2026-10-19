import datetime as dt
import json
import os
import time
import unicodedata
from pathlib import Path

import pytest

from extracao_factory import CONTEUDO, OLE, extracao_fechada, mensagem_rra, png, tela
from papaki_dre import extracao as mod
from papaki_dre.config import FUSO_BRASILIA
from papaki_dre.extracao import (ARQ_ESTADO, ARQ_MANIFESTO, ARQ_TRAVA, RELATORIOS, Extracao, ExtracaoErro,
                                 carregar_manifesto, iniciar)


@pytest.fixture
def base(tmp_path):
    return tmp_path / "dados" / "extracoes" / "2026-09" / "untitled"


@pytest.fixture
def ex(base):
    return iniciar(base, "untitled", "2026-09", "2026-09-30")


@pytest.fixture
def print_(tmp_path):
    return png(tmp_path)


def _arquivo_novo(tmp_path, rel, nome=None):
    nome_padrao, dados = CONTEUDO[rel]
    p = tmp_path / (nome or nome_padrao)
    p.write_bytes(dados)
    return p


class Relogio:
    """Relógio controlado: devolve `atual` e avança `passo` segundos a cada leitura."""

    def __init__(self, inicio, passo=1):
        self.atual, self.passo = inicio, dt.timedelta(seconds=passo)

    def __call__(self):
        agora = self.atual
        self.atual += self.passo
        return agora


def _estado_em_disco(ex):
    return json.loads((ex.pasta / ARQ_ESTADO).read_text(encoding="utf-8"))


def _completar(ex, rel, print_, tmp_path):
    ex.registrar_tela(rel, tela(rel), print_)
    ex.disparar(rel)
    if rel == "rra":
        ex.escolher_rra([mensagem_rra(ex)])
    return ex.registrar_arquivo(rel, _arquivo_novo(tmp_path, rel))


# --- iniciar / retomar ---------------------------------------------------------------------------

def test_iniciar_cria_a_pasta_com_estado(ex):
    assert (ex.pasta / ARQ_ESTADO).is_file()
    assert ex.estado["periodo_inicio"] == "2026-01-01" and ex.estado["periodo_fim"] == "2026-09-30"


def test_produtora_desconhecida(tmp_path):
    with pytest.raises(ExtracaoErro, match="produtora"):
        iniciar(tmp_path / "x", "outra", "2026-09", "2026-09-30")


@pytest.mark.parametrize("competencia,fim,trecho", [
    ("2026-13", "2026-09-30", "AAAA-MM"),
    ("26-09", "2026-09-30", "AAAA-MM"),
    ("2026-09", "30/09/2026", "AAAA-MM-DD"),
    ("2026-09", "2026-10-31", "dentro da competência"),
    ("2026-09", "2025-09-30", "dentro da competência"),
])
def test_competencia_e_periodo_invalidos(tmp_path, competencia, fim, trecho):
    with pytest.raises(ExtracaoErro, match=trecho):
        iniciar(tmp_path / "x", "untitled", competencia, fim)
    assert not (tmp_path / "x").exists()


def test_extracao_aberta_exige_retomar(ex):
    with pytest.raises(ExtracaoErro, match="--retomar"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30")
    de_novo = iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", retomar=True)
    assert de_novo.pasta == ex.pasta


def test_retomar_e_nova_juntos(ex):
    with pytest.raises(ExtracaoErro, match="não os dois"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", retomar=True, nova=True)


def test_retomar_sem_nada_aberto(tmp_path):
    with pytest.raises(ExtracaoErro, match="não há extração aberta"):
        iniciar(tmp_path / "x", "untitled", "2026-09", "2026-09-30", retomar=True)


def test_retomar_de_outro_periodo(ex):
    with pytest.raises(ExtracaoErro, match="outra competência/período"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-15", retomar=True)


def test_nova_nao_e_permitida_com_disparo_sem_arquivo(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    with pytest.raises(ExtracaoErro, match="disparo sem arquivo") as e:
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", nova=True)
    assert "Downloads" in str(e.value) and "--retomar" in str(e.value)
    assert len([p for p in ex.pasta.parent.iterdir() if p.is_dir()]) == 1     # nenhuma pasta nova


def test_nova_com_extracao_aberta_sem_pendencia(ex):
    nova = iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", nova=True)
    assert nova.pasta != ex.pasta


def test_retomar_continua_de_onde_parou(ex, print_, tmp_path):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    ex.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))
    de_novo = iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", retomar=True)
    assert de_novo.estado["relatorios"]["cp"]["arquivo"]["nome_original"] == "Contas a pagar.csv"
    with pytest.raises(ExtracaoErro, match="já disparado"):
        de_novo.disparar("cp")


def test_retomar_disparado_sem_arquivo_nao_dispara_e_orienta(ex, print_, tmp_path):
    """O processo morreu entre gravar o disparo e o clique (ou antes do download): nunca dispara de novo."""
    ex.registrar_tela("cr", tela("cr"), print_)
    ex.disparar("cr")
    de_novo = iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", retomar=True)
    with pytest.raises(ExtracaoErro, match="NÃO dispare de novo") as e:
        de_novo.disparar("cr")
    assert "Downloads" in str(e.value) and "extracao arquivo" in str(e.value) and "avise o usuário" in str(e.value)
    with pytest.raises(ExtracaoErro, match="já foi disparado"):
        de_novo.registrar_tela("cr", tela("cr"), print_)     # nem "revalidar" a tela para disparar de novo
    de_novo.registrar_arquivo("cr", _arquivo_novo(tmp_path, "cr"))   # o caminho é achar o arquivo
    assert de_novo.status()["relatorios"]["cr"]["arquivo"] is True


def test_extracao_fechada_nao_conta_como_aberta(tmp_path):
    fechada = extracao_fechada(tmp_path)
    nova = iniciar(fechada.pasta.parent, "untitled", "2026-09", "2026-09-30")
    assert (fechada.pasta / ARQ_MANIFESTO).is_file()
    assert nova.pasta != fechada.pasta and (nova.pasta / ARQ_ESTADO).is_file()


def test_dois_inicios_no_mesmo_minuto_nao_colidem(base):
    fixo = lambda: dt.datetime(2026, 10, 14, 9, 30, 15, tzinfo=FUSO_BRASILIA)
    pastas = [iniciar(base, "untitled", "2026-09", "2026-09-30", agora=fixo).pasta]
    for _ in range(11):
        pastas.append(iniciar(base, "untitled", "2026-09", "2026-09-30", nova=True, agora=fixo).pasta)
    assert len(set(pastas)) == 12
    assert pastas[0].name == "20261014-0930" and pastas[1].name == "20261014-0930-2"
    # retomar escolhe a mais recente mesmo com sufixo de dois dígitos (ordem numérica, não alfabética)
    assert iniciar(base, "untitled", "2026-09", "2026-09-30", retomar=True).pasta == pastas[-1]


def _desabandonar(pasta):
    """Edição à mão: tira a marca de abandonada (só para montar cenários com duas extrações abertas)."""
    e = json.loads((pasta / ARQ_ESTADO).read_text(encoding="utf-8"))
    e.pop("abandonada_em", None)
    e.pop("abandonada_por", None)
    (pasta / ARQ_ESTADO).write_text(json.dumps(e, ensure_ascii=False), encoding="utf-8")


def test_retomar_prefere_a_extracao_com_disparo_pendente(base, print_):
    a = iniciar(base, "untitled", "2026-09", "2026-09-30")
    b = iniciar(base, "untitled", "2026-09", "2026-09-30", nova=True)
    _desabandonar(a.pasta)
    Extracao.abrir(a.pasta).registrar_tela("cp", tela("cp"), print_)
    Extracao.abrir(a.pasta).disparar("cp")
    assert iniciar(base, "untitled", "2026-09", "2026-09-30", retomar=True).pasta == a.pasta
    e = _estado_em_disco(b)                                   # B com disparo pendente (editado à mão)
    e["relatorios"]["cr"] = copy_rel = _estado_em_disco(a)["relatorios"]["cp"]
    copy_rel["prova"] = "prints/cr.png"
    copy_rel["tela"] = {**tela("cr"), "vista_em": copy_rel["tela"]["vista_em"]}
    (b.pasta / ARQ_ESTADO).write_text(json.dumps(e, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="mais de uma"):
        iniciar(base, "untitled", "2026-09", "2026-09-30", retomar=True)


def test_estado_corrompido_bloqueia_nova_e_retomar(ex):
    (ex.pasta / ARQ_ESTADO).write_text("{corrompido", encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="ilegível"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", nova=True)
    with pytest.raises(ExtracaoErro, match="ilegível"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", retomar=True)
    with pytest.raises(ExtracaoErro, match="ilegível"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30")


def test_estado_corrompido_de_extracao_fechada_valida_nao_bloqueia(tmp_path):
    fechada = extracao_fechada(tmp_path)
    (fechada.pasta / ARQ_ESTADO).write_text("{corrompido", encoding="utf-8")
    nova = iniciar(fechada.pasta.parent, "untitled", "2026-09", "2026-09-30")
    assert nova.pasta != fechada.pasta
    assert iniciar(fechada.pasta.parent, "untitled", "2026-09", "2026-09-30", retomar=True).pasta == nova.pasta


def test_estado_corrompido_com_manifesto_invalido_continua_bloqueando(tmp_path):
    fechada = extracao_fechada(tmp_path)
    (fechada.pasta / "cp.csv").write_bytes(b"adulterado;\n")                 # manifesto deixa de valer
    (fechada.pasta / ARQ_ESTADO).write_text("{corrompido", encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="ilegível"):
        iniciar(fechada.pasta.parent, "untitled", "2026-09", "2026-09-30")
    (fechada.pasta / ARQ_MANIFESTO).write_text("{ruim", encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="ilegível"):
        iniciar(fechada.pasta.parent, "untitled", "2026-09", "2026-09-30", nova=True)


def test_agora_sem_fuso_e_recusado(tmp_path):
    with pytest.raises(ExtracaoErro, match="fuso"):
        iniciar(tmp_path / "x", "untitled", "2026-09", "2026-09-30", agora=lambda: dt.datetime(2026, 10, 14, 9))


def test_horarios_sao_gravados_em_brasilia(base, print_):
    utc = lambda: dt.datetime(2026, 10, 14, 12, 0, tzinfo=dt.timezone.utc)
    ex = iniciar(base, "untitled", "2026-09", "2026-09-30", agora=utc)
    ex.registrar_tela("cp", tela("cp"), print_)
    assert ex.disparar("cp") == "2026-10-14T09:00:00-03:00"


# --- abrir com estado inválido -------------------------------------------------------------------

def test_abrir_pasta_invalida(tmp_path):
    with pytest.raises(ExtracaoErro, match="inválida"):
        Extracao.abrir(tmp_path)


@pytest.mark.parametrize("estrago", [
    lambda e: e.pop("produtora"),
    lambda e: e.update(produtora="outra"),
    lambda e: e.update(versao=99),
    lambda e: e.update(competencia="2026-9"),
    lambda e: e.update(periodo_fim="ontem"),
    lambda e: e.update(periodo_inicio="2026-02-01"),
    lambda e: e.update(relatorios=[]),
    lambda e: e["relatorios"].update(xx={}),
    lambda e: e["relatorios"].update(cp={"disparo": "2026-10-14T09:00:00"}),            # sem fuso
    lambda e: e["relatorios"].update(cp={"disparo": "não é data"}),
    lambda e: e["relatorios"].update(cp={"arquivo": {"nome": "cp.csv"}}),               # arquivo sem disparo
    lambda e: e.update(rra_ids_antes="abc"),
    lambda e: e.update(rra_mensagem=5),
    lambda e: e.clear(),
])
def test_abrir_estado_com_campos_faltando_ou_errados(ex, estrago):
    estado = _estado_em_disco(ex)
    estrago(estado)
    (ex.pasta / ARQ_ESTADO).write_text(json.dumps(estado), encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="inválida"):
        Extracao.abrir(ex.pasta)


def test_abrir_estado_que_nao_e_objeto(ex):
    (ex.pasta / ARQ_ESTADO).write_text("[1, 2]", encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="inválida"):
        Extracao.abrir(ex.pasta)


# --- travas da tela ------------------------------------------------------------------------------

@pytest.mark.parametrize("rel,mudanca,trecho", [
    ("cr", dict(localidade="UTLD PRODUCAO"), "Localidade"),                 # sem cedilha/til: proibida
    ("cr", dict(localidade="UTLD PRODUCAO DE FILMES LTDA"), "Localidade"),  # o nome inteiro sem acento também
    ("cr", dict(localidade="UTLD PRODUÇAO DE FILMES LTDA"), "Localidade"),  # só sem o til
    ("cr", dict(localidade="UTLD PRODUCÃO DE FILMES LTDA"), "Localidade"),  # só sem a cedilha
    ("comissoes", dict(localidade="UTLD PRODUCAO DE FILMES LTDA"), "Localidade"),
    ("cr", dict(localidade="utld produção de filmes ltda"), "Localidade"),  # exatamente, não "parecido"
    ("cr", dict(localidade="UTLD PRODUÇÃO DE FILMES LTDA - FILIAL"), "Localidade"),
    ("cr", dict(localidade="Todas"), "Localidade"),
    ("cr", dict(localidade=None), "Localidade"),
    ("comissoes", dict(localidade="Todas as localidades"), "Localidade"),
    ("cp", dict(localidade="UTLD PRODUÇÃO DE FILMES LTDA"), "Localidade"),
    ("cp", dict(localidade="Todas exceto UTLD"), "Localidade"),
    ("rra", dict(localidade="Todas, menos uma"), "Localidade"),
    ("rra", dict(localidade="UTLD PRODUÇÃO DE FILMES LTDA"), "Localidade"),
    ("cp", dict(criterio="Venc. alterado"), "Critério"),
    ("rra", dict(criterio="Data competência"), "Critério"),
    ("comissoes", dict(criterio="Data competência"), "Critério"),
    ("cp", dict(data_inicio="2026-02-01"), "Data inicial"),
    ("cp", dict(data_inicio="2025-01-01"), "Data inicial"),
    ("cp", dict(data_fim="2026-09-15"), "Data final"),
    ("cp", dict(data_fim="amanhã"), "Data final"),
    ("cr", dict(status="Pendentes"), "Status"),
    ("rra", dict(status=None), "Status"),
    ("cp", dict(pesquisar_em="Individuais"), "Pesquisar em"),
    ("cp", dict(pesquisar_em=None), "Pesquisar em"),
    ("cp", dict(total=-1), "Total"),
    ("cp", dict(total="muitos"), "Total"),
    ("cp", dict(total=True), "Total"),
    ("cp", dict(total=3.5), "Total"),
])
def test_tela_invalida_e_recusada(ex, print_, rel, mudanca, trecho):
    with pytest.raises(ExtracaoErro, match=trecho) as e:
        ex.registrar_tela(rel, tela(rel, **mudanca), print_)
    assert "NÃO exporte" in str(e.value)
    assert "tela" not in ex.estado["relatorios"].get(rel, {})
    assert "tela" not in _estado_em_disco(ex)["relatorios"].get(rel, {})
    assert not (ex.pasta / "prints").exists() or not any((ex.pasta / "prints").iterdir())


def test_localidade_da_produtora_vem_da_configuracao_e_tem_acento():
    from papaki_dre.produtoras import carregar
    assert carregar("untitled").localidade == "UTLD PRODUÇÃO DE FILMES LTDA"
    assert mod.localidade_esperada("cr", "untitled") == "UTLD PRODUÇÃO DE FILMES LTDA"
    assert mod.localidade_esperada("cp", "untitled") is None


def test_localidade_com_acento_decomposto_e_a_mesma(ex, print_):
    """O mesmo texto com os acentos em forma decomposta (NFD) é o mesmo nome; não é a variante sem acento."""
    nfd = unicodedata.normalize("NFD", "UTLD PRODUÇÃO DE FILMES LTDA")
    assert nfd != "UTLD PRODUÇÃO DE FILMES LTDA"
    ex.registrar_tela("cr", tela("cr", localidade=nfd), print_)
    assert ex.estado["relatorios"]["cr"]["tela"]["localidade"] == "UTLD PRODUÇÃO DE FILMES LTDA"


def test_tela_aceita_variantes_legitimas(ex, print_):
    ex.registrar_tela("cp", tela("cp", localidade="Todas as localidades", data_inicio="01/01/2026",
                                data_fim="30/09/2026", status=" todos "), print_)
    t = ex.estado["relatorios"]["cp"]["tela"]
    assert t["data_inicio"] == "2026-01-01" and t["data_fim"] == "2026-09-30" and t["total"] == 37030
    ex.registrar_tela("cr", tela("cr", localidade="  UTLD  PRODUÇÃO DE FILMES LTDA "), print_)


def test_tela_sem_print_e_recusada(ex, tmp_path):
    with pytest.raises(ExtracaoErro, match="print"):
        ex.registrar_tela("cp", tela("cp"), tmp_path / "nao-existe.png")


def test_print_que_nao_e_imagem_e_recusado(ex, tmp_path):
    falso = tmp_path / "tela.png"
    falso.write_bytes(b"isto nao e uma imagem")
    with pytest.raises(ExtracaoErro, match="print"):
        ex.registrar_tela("cp", tela("cp"), falso)
    txt = tmp_path / "tela.txt"
    txt.write_bytes(b"\x89PNG\r\n\x1a\n")
    with pytest.raises(ExtracaoErro, match="print"):
        ex.registrar_tela("cp", tela("cp"), txt)


def test_print_e_copiado_para_a_pasta(ex, print_):
    ex.registrar_tela("cr", tela("cr"), print_)
    assert (ex.pasta / "prints" / "cr.png").is_file()
    assert ex.estado["relatorios"]["cr"]["prova"] == "prints/cr.png"


def test_prints_tem_nome_por_relatorio_e_troca_nao_deixa_sobra(ex, tmp_path, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.registrar_tela("cr", tela("cr"), print_)               # o mesmo arquivo de print serve aos dois
    jpg = tmp_path / "outra.JPEG"
    jpg.write_bytes(b"\xff\xd8\xff\xe0" + b"\x00" * 16)
    ex.registrar_tela("cp", tela("cp"), jpg)                   # nova leitura da tela antes do disparo
    nomes = sorted(p.name for p in (ex.pasta / "prints").iterdir())
    assert nomes == ["cp.jpg", "cr.png"]
    assert ex.estado["relatorios"]["cp"]["prova"] == "prints/cp.jpg"
    ex.registrar_tela("cr", tela("cr"), ex.pasta / "prints" / "cr.png")   # o próprio print copiado
    assert (ex.pasta / "prints" / "cr.png").is_file()


def test_tela_que_falha_descarta_a_tela_anterior(ex, print_):
    """A tela mudou e agora está errada: a validação antiga não pode mais liberar o disparo."""
    ex.registrar_tela("cr", tela("cr"), print_)
    with pytest.raises(ExtracaoErro, match="Localidade"):
        ex.registrar_tela("cr", tela("cr", localidade="UTLD PRODUCAO"), print_)
    assert "tela" not in _estado_em_disco(ex)["relatorios"].get("cr", {})
    assert "tela" not in ex.estado["relatorios"].get("cr", {})
    with pytest.raises(ExtracaoErro, match="registre a tela"):
        ex.disparar("cr")


def test_relatorio_desconhecido(ex, print_):
    with pytest.raises(ExtracaoErro, match="relatório"):
        ex.registrar_tela("xx", tela("cp"), print_)


def test_mensagem_de_erro_nao_ecoa_texto_longo(ex, print_):
    """O que o Claude digitou como 'Localidade' pode ser lixo copiado da tela: a mensagem só ecoa o começo."""
    lixo = "UTLD PRODUCAO " + "x" * 40 + " CLIENTE SIGILOSO 123.456.789-00"
    with pytest.raises(ExtracaoErro) as e:
        ex.registrar_tela("cr", tela("cr", localidade=lixo), print_)
    assert "SIGILOSO" not in str(e.value) and "123.456" not in str(e.value)


# --- disparo único -------------------------------------------------------------------------------

def test_nao_dispara_sem_tela_validada(ex):
    with pytest.raises(ExtracaoErro, match="registre a tela"):
        ex.disparar("cp")


def test_um_disparo_por_relatorio_e_o_registro_vai_para_o_disco(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    assert _estado_em_disco(ex)["relatorios"]["cp"]["disparo"]
    with pytest.raises(ExtracaoErro, match="NÃO dispare de novo"):
        ex.disparar("cp")


def test_disparo_visto_por_outro_processo(ex, print_):
    """Dois objetos (dois processos) sobre a mesma pasta: o segundo relê o disco e recusa."""
    ex.registrar_tela("cp", tela("cp"), print_)
    outro = Extracao.abrir(ex.pasta)
    ex.disparar("cp")
    with pytest.raises(ExtracaoErro, match="NÃO dispare de novo"):
        outro.disparar("cp")


def _trava(ex):
    """A trava fica na pasta-base (competência/produtora): vale para iniciar e para todas as extrações dela."""
    return ex.pasta.parent / ARQ_TRAVA


def _pid_morto():
    import subprocess
    import sys
    p = subprocess.Popen([sys.executable, "-c", "pass"])
    p.wait()
    return p.pid


def _trava_de_outro(ex, idade_s, pid=None):
    _trava(ex).write_text(json.dumps({"pid": os.getpid() if pid is None else pid, "id": "outro",
                                      "criada": time.time() - idade_s}), encoding="utf-8")


def test_operacao_com_trava_presa_falha_sem_gravar(ex, print_, monkeypatch):
    monkeypatch.setattr(mod, "TRAVA_ESPERA_S", 0.1)
    ex.registrar_tela("cp", tela("cp"), print_)
    _trava(ex).write_text("", encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="em andamento"):
        ex.disparar("cp")
    assert "disparo" not in _estado_em_disco(ex)["relatorios"]["cp"]


def test_trava_de_processo_vivo_bloqueia_mesmo_velha(ex, print_, monkeypatch):
    """Dono vivo e lento (antivírus, notebook suspenso): a trava NÃO é tomada antes do limite de abandono."""
    monkeypatch.setattr(mod, "TRAVA_ESPERA_S", 0.2)
    ex.registrar_tela("cp", tela("cp"), print_)
    _trava_de_outro(ex, idade_s=mod.TRAVA_ABANDONO_S - 60)              # pid = este processo, vivo
    with pytest.raises(ExtracaoErro, match="em andamento"):
        ex.disparar("cp")
    assert _trava(ex).is_file()                                          # a trava alheia continua lá
    assert "disparo" not in _estado_em_disco(ex)["relatorios"]["cp"]


def test_trava_recente_liberada_durante_a_espera(ex, print_):
    import threading
    ex.registrar_tela("cp", tela("cp"), print_)
    _trava_de_outro(ex, idade_s=0)
    threading.Timer(0.2, lambda: mod._apagar(_trava(ex))).start()        # o "outro" libera
    assert ex.disparar("cp")


@pytest.mark.parametrize("caso", ["pid_morto", "abandonada", "vazia_abandonada"])
def test_trava_orfa_e_tomada_sozinha(ex, print_, monkeypatch, caso):
    monkeypatch.setattr(mod, "TRAVA_ESPERA_S", 0.2)
    ex.registrar_tela("cp", tela("cp"), print_)
    if caso == "pid_morto":                                              # processo morto: tomada já
        _trava_de_outro(ex, idade_s=1, pid=_pid_morto())
    elif caso == "abandonada":                                           # pid "vivo" (reuso), mas velha demais
        _trava_de_outro(ex, idade_s=mod.TRAVA_ABANDONO_S + 30)
    else:                                                                # morreu antes de escrever o conteúdo
        _trava(ex).write_text("", encoding="utf-8")
        velho = time.time() - mod.TRAVA_ABANDONO_S - 30
        os.utime(_trava(ex), (velho, velho))
    ex.disparar("cp")
    assert _estado_em_disco(ex)["relatorios"]["cp"]["disparo"]
    assert not _trava(ex).exists()


def test_processo_vivo():
    assert mod._processo_vivo(os.getpid()) is True
    assert mod._processo_vivo(_pid_morto()) is False
    for desconhecido in (None, "12", -1, 0, True):                       # sem pid confiável: considera vivo
        assert mod._processo_vivo(desconhecido) is True


def test_trava_guarda_pid_e_horario(ex):
    with mod._trava(ex.pasta.parent):
        dono = json.loads(_trava(ex).read_text(encoding="utf-8"))
    assert dono["pid"] == os.getpid() and abs(dono["criada"] - time.time()) < 5 and dono["id"]
    assert not _trava(ex).exists()


def test_liberar_nao_apaga_trava_de_outro_dono(ex):
    with mod._trava(ex.pasta.parent):
        _trava_de_outro(ex, idade_s=0)                                   # foi tomada (ex.: julgada órfã)
    assert json.loads(_trava(ex).read_text(encoding="utf-8"))["id"] == "outro"


def test_tomada_simultanea_de_trava_orfa_nao_da_dois_donos(ex):
    import threading
    donos, maximo, erros = [0], [0], []
    lock = threading.Lock()
    morto = _pid_morto()

    def trabalhar(barreira):
        try:
            barreira.wait()
            with mod._trava(ex.pasta.parent):
                with lock:
                    donos[0] += 1
                    maximo[0] = max(maximo[0], donos[0])
                time.sleep(0.01)
                with lock:
                    donos[0] -= 1
        except Exception as e:                                           # noqa: BLE001
            erros.append(e)

    for _ in range(15):
        _trava_de_outro(ex, idade_s=1, pid=morto)
        barreira = threading.Barrier(6)
        threads = [threading.Thread(target=trabalhar, args=(barreira,)) for _ in range(6)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
    assert not erros and maximo[0] == 1
    assert not _trava(ex).exists()


def test_tomada_reconfere_a_trava_antes_de_apagar(ex, monkeypatch):
    """Intercalação determinística: entre a primeira leitura (trava órfã) e a reconferência, outro
    processo tomou a órfã e criou a sua trava nova. A tomada não pode apagar a trava nova."""
    morto = _pid_morto()
    _trava_de_outro(ex, idade_s=1, pid=morto)
    ler = mod._ler_trava
    chamadas = [0]

    def ler_e_trocar(caminho):
        chamadas[0] += 1
        conteudo = ler(caminho)
        if chamadas[0] == 1:                                             # depois da 1ª leitura, outro tomou
            _trava_de_outro(ex, idade_s=0)                               # trava nova, de dono vivo
        return conteudo
    monkeypatch.setattr(mod, "_ler_trava", ler_e_trocar)
    mod._tomar_trava_orfa(_trava(ex))
    monkeypatch.undo()
    assert json.loads(_trava(ex).read_text(encoding="utf-8"))["id"] == "outro"
    assert not (ex.pasta.parent / (ARQ_TRAVA + ".tomada")).exists()


def test_dono_lento_que_perdeu_a_trava_nao_grava_por_cima(ex, print_, monkeypatch):
    """Cenário P2: o dono vivo demora mais que o limite, outro toma a trava e dispara; quando o lento
    termina, ele NÃO pode gravar o estado que calculou antes (apagaria o disparo)."""
    import threading
    monkeypatch.setattr(mod, "TRAVA_ABANDONO_S", 0.5)
    ex.registrar_tela("cp", tela("cp"), print_)
    erros = []

    def lento(e):
        time.sleep(1.5)
        e["rra_ids_antes"] = ["x"]

    def rodar():
        try:
            ex._mudar(lento)
        except ExtracaoErro as err:
            erros.append(err)
    t = threading.Thread(target=rodar)
    t.start()
    time.sleep(0.8)
    disparo = Extracao.abrir(ex.pasta).disparar("cp")
    t.join()
    assert erros and "nada foi gravado" in str(erros[0])
    disco = _estado_em_disco(ex)
    assert disco["relatorios"]["cp"]["disparo"] == disparo and disco["rra_ids_antes"] is None
    with pytest.raises(ExtracaoErro, match="NÃO dispare de novo"):
        Extracao.abrir(ex.pasta).disparar("cp")


def test_estado_mudado_no_disco_durante_a_operacao_aborta(ex, print_):
    """Mesmo com a trava, se o estado.json mudou desde a leitura, a gravação é abortada."""
    ex.registrar_tela("cp", tela("cp"), print_)

    def mexe_no_disco(e):
        outro = json.loads((ex.pasta / ARQ_ESTADO).read_text(encoding="utf-8"))
        outro["rra_ids_antes"] = ["de outro processo"]
        (ex.pasta / ARQ_ESTADO).write_text(json.dumps(outro), encoding="utf-8")
        e["rra_ids_antes"] = ["meu"]
    with pytest.raises(ExtracaoErro, match="nada foi gravado"):
        ex._mudar(mexe_no_disco)
    assert _estado_em_disco(ex)["rra_ids_antes"] == ["de outro processo"]


def test_criar_trava_sem_permissao_nao_vira_em_andamento(ex, print_, monkeypatch):
    """Pasta somente leitura: erro claro de permissão, não 'em andamento' nem traceback."""
    ex.registrar_tela("cp", tela("cp"), print_)
    abrir_os = os.open

    def sem_permissao(caminho, *a, **k):
        if str(caminho).endswith(ARQ_TRAVA):
            raise PermissionError("somente leitura")
        return abrir_os(caminho, *a, **k)
    monkeypatch.setattr(mod.os, "open", sem_permissao)
    with pytest.raises(ExtracaoErro, match="sem permissão"):
        ex.disparar("cp")


def test_criar_tomada_sem_permissao_nao_da_traceback(ex, print_, monkeypatch):
    monkeypatch.setattr(mod, "TRAVA_ESPERA_S", 0.2)
    ex.registrar_tela("cp", tela("cp"), print_)
    _trava_de_outro(ex, idade_s=1, pid=_pid_morto())
    abrir_os = os.open

    def sem_permissao(caminho, *a, **k):
        if str(caminho).endswith(".tomada"):
            raise PermissionError("somente leitura")
        return abrir_os(caminho, *a, **k)
    monkeypatch.setattr(mod.os, "open", sem_permissao)
    with pytest.raises(ExtracaoErro, match="sem permissão"):
        ex.disparar("cp")
    monkeypatch.undo()
    assert "disparo" not in _estado_em_disco(ex)["relatorios"]["cp"]


def test_replace_do_estado_insiste_quando_o_windows_segura_o_arquivo(ex, print_, monkeypatch):
    ex.registrar_tela("cp", tela("cp"), print_)
    trocar = os.replace
    falhas = [3]

    def ocupado(*a, **k):
        if falhas[0]:
            falhas[0] -= 1
            raise PermissionError("arquivo em uso por outro processo")
        return trocar(*a, **k)
    monkeypatch.setattr(mod.os, "replace", ocupado)
    assert ex.disparar("cp")
    monkeypatch.undo()
    assert _estado_em_disco(ex)["relatorios"]["cp"]["disparo"]


def test_disparar_com_leitor_do_estado_aberto(ex, print_):
    """Outro processo lendo o estado.json (ex.: 'extracao estado') não derruba o disparo."""
    import threading
    ex.registrar_tela("cp", tela("cp"), print_)
    f = open(ex.pasta / ARQ_ESTADO, encoding="utf-8")
    threading.Timer(0.3, f.close).start()
    assert ex.disparar("cp")
    assert _estado_em_disco(ex)["relatorios"]["cp"]["disparo"]


def test_tela_vista_no_futuro_expira(base, print_):
    """O relógio voltou: uma tela registrada 'no futuro' não libera disparo."""
    hora = [dt.datetime(2026, 10, 14, 12, 0, tzinfo=FUSO_BRASILIA)]
    ex = iniciar(base, "untitled", "2026-09", "2026-09-30", agora=lambda: hora[0])
    ex.registrar_tela("cp", tela("cp"), print_)
    hora[0] = dt.datetime(2026, 10, 14, 10, 0, tzinfo=FUSO_BRASILIA)
    with pytest.raises(ExtracaoErro, match="registre a tela de novo"):
        ex.disparar("cp")


# --- valores gravados são revalidados ------------------------------------------------------------

def _editar_estado(ex, fn):
    e = _estado_em_disco(ex)
    fn(e)
    (ex.pasta / ARQ_ESTADO).write_text(json.dumps(e, ensure_ascii=False), encoding="utf-8")


@pytest.mark.parametrize("rel,estrago", [
    ("cr", lambda t: t.update(localidade="UTLD PRODUCAO")),
    ("cr", lambda t: t.update(status="Pendentes")),
    ("cp", lambda t: t.update(data_inicio="2026-09-30", data_fim="2026-01-01")),
    ("cp", lambda t: t.update(total=-1)),
    ("cp", lambda t: t.pop("localidade")),
    ("cp", lambda t: t.pop("pesquisar_em")),
])
def test_estado_com_tela_editada_a_mao_e_recusado(ex, print_, rel, estrago):
    ex.registrar_tela(rel, tela(rel), print_)
    _editar_estado(ex, lambda e: estrago(e["relatorios"][rel]["tela"]))
    with pytest.raises(ExtracaoErro, match="inválida"):
        Extracao.abrir(ex.pasta)
    with pytest.raises(ExtracaoErro):
        ex.disparar(rel)                                  # o objeto já aberto relê o disco e recusa
    assert "disparo" not in _estado_em_disco(ex)["relatorios"][rel]


def test_disparar_revalida_a_tela_gravada(ex, print_, monkeypatch):
    """Mesmo que a leitura do estado deixasse passar, o disparo roda as travas da tela de novo."""
    ex.registrar_tela("cr", tela("cr"), print_)
    _editar_estado(ex, lambda e: e["relatorios"]["cr"]["tela"].update(localidade="UTLD PRODUCAO"))
    monkeypatch.setattr(mod, "_problema_no_estado", lambda e: None)
    with pytest.raises(ExtracaoErro, match="Localidade") as err:
        ex.disparar("cr")
    assert "NÃO exporte" in str(err.value)
    assert "disparo" not in _estado_em_disco(ex)["relatorios"]["cr"]


@pytest.mark.parametrize("prova", ["../../x.png", "prints/cr.png", "prints/cp.gif", "C:/x/cp.png", 5])
def test_estado_com_prova_fora_do_padrao_e_recusado(ex, print_, prova):
    ex.registrar_tela("cp", tela("cp"), print_)
    _editar_estado(ex, lambda e: e["relatorios"]["cp"].update(prova=prova))
    with pytest.raises(ExtracaoErro, match="inválida"):
        Extracao.abrir(ex.pasta)


def test_estado_com_arquivo_completo_sem_disparo_e_recusado(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    _editar_estado(ex, lambda e: e["relatorios"]["cp"].update(
        arquivo={"nome": "cp.csv", "nome_original": "x.csv", "sha256": "a" * 64, "tamanho": 10,
                 "registrado_em": "2026-10-14T09:00:00-03:00"}))
    with pytest.raises(ExtracaoErro, match="inválida"):
        Extracao.abrir(ex.pasta)


@pytest.mark.parametrize("estrago", [
    lambda m: m["relatorios"]["cr"]["filtros_tela"].update(localidade="UTLD PRODUCAO"),
    lambda m: m["relatorios"]["cr"]["filtros_tela"].update(status="Pendentes"),
    lambda m: m["relatorios"]["cp"]["filtros_tela"].update(data_inicio="2026-09-30", data_fim="2026-01-01"),
    lambda m: m["relatorios"]["cp"]["filtros_tela"].pop("localidade"),
    lambda m: m["relatorios"]["cp"].update(total_tela=-5),
    lambda m: m.update(produtora="outra"),
    lambda m: m.update(competencia="xx"),
    lambda m: m.update(periodo_inicio="2025-01-01"),
    lambda m: m.update(periodo_fim="2026-08-31"),
    lambda m: m.update(periodo_fim="2026-10-31", competencia="2026-10"),
])
def test_manifesto_com_filtros_ou_periodo_editados_e_recusado(tmp_path, estrago):
    ex = extracao_fechada(tmp_path)
    caminho = ex.pasta / ARQ_MANIFESTO
    m = json.loads(caminho.read_text(encoding="utf-8"))
    estrago(m)
    caminho.write_text(json.dumps(m, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="inválido"):
        carregar_manifesto(ex.pasta)


def test_copia_que_nao_confere_e_recusada(disparado, tmp_path, monkeypatch):
    def copia_errada(origem, destino, *a, **k):
        Path(destino).write_bytes(Path(origem).read_bytes() + b"x")
    monkeypatch.setattr(mod.shutil, "copyfile", copia_errada)
    with pytest.raises(ExtracaoErro, match="cópia não confere"):
        disparado.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))
    assert "arquivo" not in _estado_em_disco(disparado)["relatorios"]["cp"]
    assert not (disparado.pasta / "cp.csv").exists()


def test_copia_com_tamanho_igual_e_conteudo_diferente_e_recusada(disparado, tmp_path, monkeypatch):
    def copia_trocada(origem, destino, *a, **k):
        dados = Path(origem).read_bytes()
        Path(destino).write_bytes(dados[:-1] + b"9")
    monkeypatch.setattr(mod.shutil, "copyfile", copia_trocada)
    with pytest.raises(ExtracaoErro, match="cópia não confere"):
        disparado.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))


# --- iniciar sob trava e pastas fora do padrão ---------------------------------------------------

def test_dois_iniciar_simultaneos_nao_criam_duas_abertas(base, monkeypatch):
    import threading
    base.mkdir(parents=True)
    varrer = mod._pastas_com_estado

    def varrer_devagar(b):
        achadas = varrer(b)
        time.sleep(0.3)                                  # os dois varrem antes de qualquer um criar
        return achadas
    monkeypatch.setattr(mod, "_pastas_com_estado", varrer_devagar)
    criadas, erros = [], []
    barreira = threading.Barrier(2)

    def rodar():
        barreira.wait()
        try:
            criadas.append(iniciar(base, "untitled", "2026-09", "2026-09-30"))
        except ExtracaoErro as e:
            erros.append(e)
    ts = [threading.Thread(target=rodar) for _ in range(2)]
    for t in ts:
        t.start()
    for t in ts:
        t.join()
    assert len(criadas) == 1 and len(erros) == 1 and "--retomar" in str(erros[0])
    assert len([p for p in base.iterdir() if p.is_dir()]) == 1


def test_disparo_e_iniciar_usam_a_mesma_trava(ex, print_, monkeypatch):
    """Um disparo em andamento numa extração segura o iniciar de outra (a varredura não perde o disparo)."""
    monkeypatch.setattr(mod, "TRAVA_ESPERA_S", 0.2)
    _trava_de_outro(ex, idade_s=0)
    with pytest.raises(ExtracaoErro, match="em andamento"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", nova=True)


def test_pasta_fora_do_padrao_nao_e_a_mais_recente(base):
    a = iniciar(base, "untitled", "2026-09", "2026-09-30",
                agora=lambda: dt.datetime(2026, 10, 1, 9, tzinfo=FUSO_BRASILIA))
    os.rename(a.pasta, base / "zz-antiga-revisar")
    b = iniciar(base, "untitled", "2026-09", "2026-09-30", nova=True,
                agora=lambda: dt.datetime(2026, 10, 14, 9, tzinfo=FUSO_BRASILIA))
    _desabandonar(base / "zz-antiga-revisar")               # duas abertas: só a ordem decide
    assert iniciar(base, "untitled", "2026-09", "2026-09-30", retomar=True).pasta == b.pasta


def test_pasta_fora_do_padrao_com_disparo_pendente_continua_bloqueando(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    os.rename(ex.pasta, ex.pasta.parent / "renomeada")
    with pytest.raises(ExtracaoErro, match="disparo sem arquivo"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", nova=True)


def test_pasta_no_padrao_sem_estado_bloqueia(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    os.replace(ex.pasta / ARQ_ESTADO, ex.pasta / "estado.json.bak")
    with pytest.raises(ExtracaoErro, match="avise o usuário"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", nova=True)
    with pytest.raises(ExtracaoErro, match="avise o usuário"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", retomar=True)


def test_pasta_no_padrao_vazia_nao_bloqueia(base):
    """Pasta criada e vazia (processo morto antes de gravar o estado): não pode ter disparo."""
    (base / "20261014-0900").mkdir(parents=True)
    assert iniciar(base, "untitled", "2026-09", "2026-09-30").pasta.name != "20261014-0900"


@pytest.mark.parametrize("residuos", [["estado.json.tmp"], [".trava"], [".trava.tomada", "estado.json.tmp"]])
def test_pasta_no_padrao_so_com_residuos_nao_bloqueia(base, residuos):
    """Só sobras de gravação interrompida (estado.json.tmp, .trava*): o estado nunca chegou a existir."""
    pasta = base / "20261014-0900"
    pasta.mkdir(parents=True)
    for nome in residuos:
        (pasta / nome).write_text("{", encoding="utf-8")
    assert iniciar(base, "untitled", "2026-09", "2026-09-30").pasta != pasta


def test_pasta_no_padrao_com_residuo_e_outro_arquivo_bloqueia(base):
    pasta = base / "20261014-0900"
    pasta.mkdir(parents=True)
    (pasta / "estado.json.tmp").write_text("{", encoding="utf-8")
    (pasta / "cp.csv").write_bytes(CONTEUDO["cp"][1])
    with pytest.raises(ExtracaoErro, match="avise o usuário"):
        iniciar(base, "untitled", "2026-09", "2026-09-30")


# --- --nova abandona as anteriores; disparo olha as outras extrações da base ---------------------

def test_nova_abandona_a_anterior_e_ela_nao_dispara_de_novo(base, print_):
    """Cenário da revisão: A aberta, --nova cria B, B dispara cp; A NÃO pode disparar cp de novo."""
    a = iniciar(base, "untitled", "2026-09", "2026-09-30")
    a.registrar_tela("cp", tela("cp"), print_)                 # A chegou a conferir a tela
    b = iniciar(base, "untitled", "2026-09", "2026-09-30", nova=True)
    disco_a = _estado_em_disco(a)
    assert disco_a["abandonada_em"] and disco_a["abandonada_por"] == b.pasta.name
    b.registrar_tela("cp", tela("cp"), print_)
    assert b.disparar("cp")
    with pytest.raises(ExtracaoErro, match="abandonada") as e:
        a.disparar("cp")
    assert b.pasta.name in str(e.value)
    with pytest.raises(ExtracaoErro, match="abandonada"):
        Extracao.abrir(a.pasta).disparar("cp")
    assert "disparo" not in _estado_em_disco(a)["relatorios"].get("cp", {})
    assert Extracao.abrir(a.pasta).status()["abandonada"] is True
    assert iniciar(base, "untitled", "2026-09", "2026-09-30", retomar=True).pasta == b.pasta


def test_extracao_abandonada_recusa_qualquer_mudanca(base, print_, tmp_path):
    a = iniciar(base, "untitled", "2026-09", "2026-09-30")
    iniciar(base, "untitled", "2026-09", "2026-09-30", nova=True)
    acoes = [
        lambda: a.registrar_tela("cp", tela("cp"), print_),
        lambda: a.rra_ids_antes([]),
        lambda: a.disparar("cp"),
        lambda: a.escolher_rra([]),
        lambda: a.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp")),
        lambda: a.fechar(),
    ]
    antes = (a.pasta / ARQ_ESTADO).read_bytes()
    for acao in acoes:
        with pytest.raises(ExtracaoErro, match="abandonada"):
            acao()
    assert (a.pasta / ARQ_ESTADO).read_bytes() == antes


def test_abandonada_nao_conta_como_aberta(base):
    import shutil
    a = iniciar(base, "untitled", "2026-09", "2026-09-30")
    b = iniciar(base, "untitled", "2026-09", "2026-09-30", nova=True)
    shutil.rmtree(b.pasta)                                      # sobra só a abandonada
    with pytest.raises(ExtracaoErro, match="não há extração aberta"):
        iniciar(base, "untitled", "2026-09", "2026-09-30", retomar=True)
    c = iniciar(base, "untitled", "2026-09", "2026-09-30")     # sem --nova: a abandonada não atrapalha
    assert c.pasta != a.pasta and c.status()["abandonada"] is False


def test_nova_nao_abandona_extracao_com_disparo_pendente(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    with pytest.raises(ExtracaoErro, match="disparo sem arquivo"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", nova=True)
    assert "abandonada_em" not in _estado_em_disco(ex)


def test_disparar_recusa_se_outra_extracao_tem_disparo_pendente(base, print_):
    """Mesmo com duas abertas (estado editado à mão), o disparo olha as outras pastas da base."""
    a = iniciar(base, "untitled", "2026-09", "2026-09-30")
    b = iniciar(base, "untitled", "2026-09", "2026-09-30", nova=True)
    _desabandonar(a.pasta)
    a.registrar_tela("cp", tela("cp"), print_)
    a.disparar("cp")                                             # A fica com cp disparado sem arquivo
    b.registrar_tela("cr", tela("cr"), print_)
    with pytest.raises(ExtracaoErro, match="disparo sem arquivo") as e:
        b.disparar("cr")
    assert a.pasta.name in str(e.value) and "NÃO dispare" in str(e.value)
    assert "disparo" not in _estado_em_disco(b)["relatorios"]["cr"]


@pytest.mark.parametrize("estrago", ["ilegivel", "sem_estado"])
def test_disparar_recusa_se_outra_pasta_da_base_e_duvidosa(base, print_, estrago):
    a = iniciar(base, "untitled", "2026-09", "2026-09-30")
    b = iniciar(base, "untitled", "2026-09", "2026-09-30", nova=True)
    if estrago == "ilegivel":
        (a.pasta / ARQ_ESTADO).write_text("{corrompido", encoding="utf-8")
    else:
        os.replace(a.pasta / ARQ_ESTADO, a.pasta / "estado.json.bak")
    b.registrar_tela("cp", tela("cp"), print_)
    with pytest.raises(ExtracaoErro, match="avise o usuário"):
        b.disparar("cp")
    assert "disparo" not in _estado_em_disco(b)["relatorios"]["cp"]


def test_disparar_ignora_extracao_fechada_valida_da_base(tmp_path, print_):
    fechada = extracao_fechada(tmp_path)
    nova = iniciar(fechada.pasta.parent, "untitled", "2026-09", "2026-09-30")
    nova.registrar_tela("cp", tela("cp"), print_)
    assert nova.disparar("cp")


@pytest.mark.parametrize("estrago", [
    lambda e: e.update(abandonada_em="ontem", abandonada_por="x"),
    lambda e: e.update(abandonada_em="2026-10-14T09:00:00-03:00"),                        # sem abandonada_por
    lambda e: e.update(abandonada_em="2026-10-14T09:00:00-03:00", abandonada_por=5),
])
def test_abrir_recusa_marca_de_abandono_malformada(ex, estrago):
    _editar_estado(ex, estrago)
    with pytest.raises(ExtracaoErro, match="inválida"):
        Extracao.abrir(ex.pasta)


def test_abrir_recusa_abandonada_com_disparo_pendente(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    _editar_estado(ex, lambda e: e.update(abandonada_em="2026-10-14T09:00:00-03:00", abandonada_por="x"))
    with pytest.raises(ExtracaoErro, match="inválida"):
        Extracao.abrir(ex.pasta)


# --- menores: posse da trava no iniciar, mensagem, cópias só depois da conferência ---------------

def test_iniciar_com_trava_roubada_nao_grava(base, monkeypatch):
    """A trava muda de dono entre a varredura e a gravação (o estado não muda): nada é criado."""
    base.mkdir(parents=True)
    momento = mod._momento

    def roubar(relogio):
        (base / ARQ_TRAVA).write_text(json.dumps({"pid": os.getpid(), "id": "ladrao", "criada": time.time()}),
                                      encoding="utf-8")
        return momento(relogio)
    monkeypatch.setattr(mod, "_momento", roubar)
    with pytest.raises(ExtracaoErro, match="nada foi gravado"):
        iniciar(base, "untitled", "2026-09", "2026-09-30")
    assert not [p for p in base.iterdir() if p.is_dir()]


def test_nova_com_trava_roubada_nao_abandona_nem_cria(ex, monkeypatch):
    base = ex.pasta.parent
    momento = mod._momento

    def roubar(relogio):
        (base / ARQ_TRAVA).write_text(json.dumps({"pid": os.getpid(), "id": "ladrao", "criada": time.time()}),
                                      encoding="utf-8")
        return momento(relogio)
    monkeypatch.setattr(mod, "_momento", roubar)
    with pytest.raises(ExtracaoErro, match="nada foi gravado"):
        iniciar(base, "untitled", "2026-09", "2026-09-30", nova=True)
    assert "abandonada_em" not in _estado_em_disco(ex)
    assert [p for p in base.iterdir() if p.is_dir()] == [ex.pasta]


def test_mensagem_de_trava_ocupada_cita_o_limite_de_10_min(ex, print_, monkeypatch):
    monkeypatch.setattr(mod, "TRAVA_ESPERA_S", 0.1)
    ex.registrar_tela("cp", tela("cp"), print_)
    _trava_de_outro(ex, idade_s=0)
    with pytest.raises(ExtracaoErro, match="10 min") as e:
        ex.disparar("cp")
    assert "pid" in str(e.value)


def _conferencia_que_falha(monkeypatch):
    def falha(*_a, **_k):
        raise ExtracaoErro("a trava foi tomada por outro processo durante a operação; nada foi gravado")
    monkeypatch.setattr(mod, "_conferir_posse", falha)


def test_print_so_e_trocado_depois_da_conferencia(base, print_, tmp_path, monkeypatch):
    ex = iniciar(base, "untitled", "2026-09", "2026-09-30",
                 agora=Relogio(dt.datetime(2026, 10, 14, 9, 0, tzinfo=FUSO_BRASILIA)))   # cada tela, outro horário
    ex.registrar_tela("cp", tela("cp"), print_)
    original = (ex.pasta / "prints" / "cp.png").read_bytes()
    outro = tmp_path / "outro.png"
    outro.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x07" * 16)
    jpg = tmp_path / "outro.jpg"
    jpg.write_bytes(b"\xff\xd8\xff\xe0" + b"\x00" * 16)
    _conferencia_que_falha(monkeypatch)
    for novo in (outro, jpg):
        with pytest.raises(ExtracaoErro, match="nada foi gravado"):
            ex.registrar_tela("cp", tela("cp"), novo)
    assert (ex.pasta / "prints" / "cp.png").read_bytes() == original       # a prova da tela gravada fica
    assert sorted(p.name for p in (ex.pasta / "prints").iterdir()) == ["cp.png"]


def test_arquivo_so_e_copiado_depois_da_conferencia(disparado, tmp_path, monkeypatch):
    _conferencia_que_falha(monkeypatch)
    with pytest.raises(ExtracaoErro, match="nada foi gravado"):
        disparado.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))
    assert sorted(p.name for p in disparado.pasta.iterdir() if p.is_file()) == [ARQ_ESTADO]


def test_falha_ao_gravar_o_disparo_nao_deixa_disparo(ex, print_, monkeypatch):
    ex.registrar_tela("cp", tela("cp"), print_)

    def falha(*_):
        raise OSError("disco cheio")
    monkeypatch.setattr(mod.os, "replace", falha)
    with pytest.raises(OSError):
        ex.disparar("cp")
    monkeypatch.undo()
    assert "disparo" not in ex.estado["relatorios"]["cp"]               # memória igual ao disco
    assert "disparo" not in _estado_em_disco(ex)["relatorios"]["cp"]
    assert not list(ex.pasta.glob("*.tmp"))
    ex.disparar("cp")                                                   # o clique não aconteceu: pode disparar


def test_estado_e_gravado_de_forma_atomica(ex, print_, monkeypatch):
    """Se a escrita do temporário quebra no meio, o estado.json antigo continua inteiro."""
    antes = (ex.pasta / ARQ_ESTADO).read_bytes()
    ex.registrar_tela("cp", tela("cp"), print_)
    depois_tela = (ex.pasta / ARQ_ESTADO).read_bytes()
    assert depois_tela != antes

    def quebra(*_a, **_k):
        raise OSError("falha no meio da escrita")
    monkeypatch.setattr(mod.os, "fsync", quebra)
    with pytest.raises(OSError):
        ex.disparar("cp")
    assert (ex.pasta / ARQ_ESTADO).read_bytes() == depois_tela
    assert not list(ex.pasta.glob("*.tmp"))


def test_tela_velha_nao_libera_disparo(base, print_):
    relogio = Relogio(dt.datetime(2026, 10, 14, 9, 0, tzinfo=FUSO_BRASILIA))
    ex = iniciar(base, "untitled", "2026-09", "2026-09-30", agora=relogio)
    ex.registrar_tela("cp", tela("cp"), print_)
    relogio.atual += dt.timedelta(seconds=mod.TELA_VALIDADE_S + 1)
    with pytest.raises(ExtracaoErro, match="registre a tela de novo"):
        ex.disparar("cp")
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")


def test_tela_nao_muda_depois_do_disparo(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    with pytest.raises(ExtracaoErro, match="já foi disparado"):
        ex.registrar_tela("cp", tela("cp", total=1), print_)


def test_rra_exige_ids_antes_registrados(ex, print_):
    ex.registrar_tela("rra", tela("rra"), print_)
    with pytest.raises(ExtracaoErro, match="rra-ids-antes"):
        ex.disparar("rra")
    ex.rra_ids_antes([])
    ex.disparar("rra")
    with pytest.raises(ExtracaoErro, match="depois do disparo"):
        ex.rra_ids_antes(["x"])


def test_rra_ids_antes_exige_lista_de_textos(ex):
    with pytest.raises(ExtracaoErro, match="lista"):
        ex.rra_ids_antes("abc")               # uma string viraria ['a', 'b', 'c']
    with pytest.raises(ExtracaoErro, match="lista"):
        ex.rra_ids_antes([1, 2])
    assert _estado_em_disco(ex)["rra_ids_antes"] is None


# --- arquivos ------------------------------------------------------------------------------------

def test_arquivo_exige_disparo(ex, tmp_path):
    with pytest.raises(ExtracaoErro, match="disparo"):
        ex.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))


@pytest.fixture
def disparado(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    return ex


def test_arquivo_registrado_com_hash(disparado, tmp_path):
    disparado.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))
    a = disparado.estado["relatorios"]["cp"]["arquivo"]
    assert a["nome"] == "cp.csv" and a["nome_original"] == "Contas a pagar.csv"
    copia = disparado.pasta / a["nome"]
    assert copia.is_file() and copia.read_bytes() == CONTEUDO["cp"][1]
    assert len(a["sha256"]) == 64 and a["tamanho"] == len(CONTEUDO["cp"][1])
    with pytest.raises(ExtracaoErro, match="já tem arquivo"):
        disparado.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))


def test_relatorios_com_o_mesmo_nome_de_arquivo(ex, print_, tmp_path):
    """O Esferas pode dar o mesmo nome a CP e CR: cada um é guardado com o nome do relatório."""
    (tmp_path / "a").mkdir()
    (tmp_path / "b").mkdir()
    for rel, sub in (("cp", "a"), ("cr", "b")):
        ex.registrar_tela(rel, tela(rel), print_)
        ex.disparar(rel)
        p = tmp_path / sub / "Relatorio.csv"
        p.write_bytes(CONTEUDO[rel][1])
        ex.registrar_arquivo(rel, p)
    assert (ex.pasta / "cp.csv").read_bytes() == CONTEUDO["cp"][1]
    assert (ex.pasta / "cr.csv").read_bytes() == CONTEUDO["cr"][1]


def test_o_mesmo_arquivo_nao_serve_a_dois_relatorios(ex, print_, tmp_path):
    for rel in ("cp", "cr"):
        ex.registrar_tela(rel, tela(rel), print_)
        ex.disparar(rel)
    p = _arquivo_novo(tmp_path, "cp")
    ex.registrar_arquivo("cp", p)
    with pytest.raises(ExtracaoErro, match="mesmo arquivo"):
        ex.registrar_arquivo("cr", p)
    with pytest.raises(ExtracaoErro, match="dentro da pasta"):
        ex.registrar_arquivo("cr", ex.pasta / "cp.csv")
    assert "arquivo" not in ex.estado["relatorios"]["cr"]


def test_sobra_de_copia_interrompida_nao_trava(disparado, tmp_path):
    """Se o processo morreu depois de copiar e antes de gravar o estado, a sobra é substituída."""
    (disparado.pasta / "cp.csv").write_bytes(b"sobra de uma copia interrompida;\n")
    disparado.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))
    assert (disparado.pasta / "cp.csv").read_bytes() == CONTEUDO["cp"][1]


def test_arquivo_inexistente(disparado, tmp_path):
    with pytest.raises(ExtracaoErro, match="não encontrado"):
        disparado.registrar_arquivo("cp", tmp_path / "sumiu.csv")


def test_arquivo_vazio(disparado, tmp_path):
    p = tmp_path / "vazio.csv"
    p.write_bytes(b"")
    with pytest.raises(ExtracaoErro, match="vazio"):
        disparado.registrar_arquivo("cp", p)


def test_extensao_errada(disparado, tmp_path):
    p = tmp_path / "x.xlsx"
    p.write_bytes(b"a;b\n")
    with pytest.raises(ExtracaoErro, match="extensão"):
        disparado.registrar_arquivo("cp", p)


@pytest.mark.parametrize("dados", [
    b"\x00\x01\x02 sem ponto e virgula",
    b"sem ponto e virgula nenhum\n",
    b"<!DOCTYPE html><html>&nbsp; sessao expirada</html>",       # página de erro salva como .csv
    b"\xef\xbb\xbf  <html><body>a;b</body></html>",
])
def test_assinatura_errada_csv(disparado, tmp_path, dados):
    p = tmp_path / "x.csv"
    p.write_bytes(dados)
    with pytest.raises(ExtracaoErro, match="não parece"):
        disparado.registrar_arquivo("cp", p)
    assert "arquivo" not in _estado_em_disco(disparado)["relatorios"]["cp"]


def test_comissoes_zip_que_nao_e_xlsx(ex, print_, tmp_path):
    import io
    import zipfile
    ex.registrar_tela("comissoes", tela("comissoes"), print_)
    ex.disparar("comissoes")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("leia-me.txt", "qualquer zip")
    p = tmp_path / "x.xlsx"
    p.write_bytes(buf.getvalue())
    with pytest.raises(ExtracaoErro, match="não parece"):
        ex.registrar_arquivo("comissoes", p)
    p.write_bytes(b"PK\x03\x04" + b"\x00" * 64)                       # assinatura certa, zip quebrado
    with pytest.raises(ExtracaoErro, match="não parece"):
        ex.registrar_arquivo("comissoes", p)


def test_arquivo_anterior_ao_disparo_e_recusado(disparado, tmp_path):
    p = _arquivo_novo(tmp_path, "cp")
    velho = time.time() - 3600
    os.utime(p, (velho, velho))
    with pytest.raises(ExtracaoErro, match="anterior ao disparo"):
        disparado.registrar_arquivo("cp", p)


def test_rra_exige_mensagem_escolhida_e_assinatura_ole2(ex, print_, tmp_path):
    ex.registrar_tela("rra", tela("rra"), print_)
    ex.rra_ids_antes([])
    ex.disparar("rra")
    p = tmp_path / "Relatorio.xls"
    p.write_bytes(OLE)
    with pytest.raises(ExtracaoErro, match="mensagem"):
        ex.registrar_arquivo("rra", p)
    ex.escolher_rra([mensagem_rra(ex)])
    ex.registrar_arquivo("rra", p)
    ruim = tmp_path / "Outro.xls"
    ruim.write_bytes(b"isto nao e um xls de verdade")
    ex2 = iniciar(tmp_path / "d2", "untitled", "2026-09", "2026-09-30")
    ex2.registrar_tela("rra", tela("rra"), print_)
    ex2.rra_ids_antes([])
    ex2.disparar("rra")
    ex2.escolher_rra([mensagem_rra(ex2)])
    with pytest.raises(ExtracaoErro, match="não parece"):
        ex2.registrar_arquivo("rra", ruim)


def test_escolher_rra_antes_do_disparo(ex):
    with pytest.raises(ExtracaoErro, match="ainda não foi disparado"):
        ex.escolher_rra([])


def test_escolher_rra_e_idempotente(ex, print_):
    ex.registrar_tela("rra", tela("rra"), print_)
    ex.rra_ids_antes(["velha"])
    ex.disparar("rra")
    assert ex.escolher_rra([mensagem_rra(ex, "m1")]) == "m1"
    assert ex.escolher_rra([mensagem_rra(ex, "m1"), mensagem_rra(ex, "m2")]) == "m1"
    assert _estado_em_disco(ex)["rra_mensagem"] == "m1"


def test_escolher_rra_ambiguo_nao_grava(ex, print_):
    from papaki_dre.rra_email import RraEmailErro
    ex.registrar_tela("rra", tela("rra"), print_)
    ex.rra_ids_antes([])
    ex.disparar("rra")
    with pytest.raises(RraEmailErro, match="mais de uma"):
        ex.escolher_rra([mensagem_rra(ex, "m1"), mensagem_rra(ex, "m2")])
    assert _estado_em_disco(ex)["rra_mensagem"] is None


# --- fechar e manifesto --------------------------------------------------------------------------

def test_fechar_incompleto_lista_o_que_falta(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    with pytest.raises(ExtracaoErro) as e:
        ex.fechar()
    for rel in RELATORIOS:
        assert rel in str(e.value)
    assert not (ex.pasta / ARQ_MANIFESTO).exists()


def test_fechar_com_tres_de_quatro_nao_gera_manifesto(ex, print_, tmp_path):
    ex.rra_ids_antes([])
    for rel in ("cp", "cr", "rra"):
        _completar(ex, rel, print_, tmp_path)
    with pytest.raises(ExtracaoErro, match="comissoes"):
        ex.fechar()
    assert not (ex.pasta / ARQ_MANIFESTO).exists()


def test_fechar_confere_os_arquivos_de_novo(ex, print_, tmp_path):
    ex.rra_ids_antes([])
    for rel in RELATORIOS:
        _completar(ex, rel, print_, tmp_path)
    (ex.pasta / "cr.csv").write_bytes(b"alterado;depois\n")
    with pytest.raises(ExtracaoErro, match="não confere"):
        ex.fechar()
    assert not (ex.pasta / ARQ_MANIFESTO).exists()


def test_fechar_gera_manifesto_valido(tmp_path):
    ex = extracao_fechada(tmp_path)
    m = carregar_manifesto(ex.pasta)
    assert m["produtora"] == "untitled" and m["competencia"] == "2026-09" and m["periodo_fim"] == "2026-09-30"
    assert m["periodo_inicio"] == "2026-01-01"
    assert set(m["relatorios"]) == set(RELATORIOS)
    cp = m["relatorios"]["cp"]
    assert cp["total_tela"] == 37030 and cp["prova"] == "prints/cp.png" and cp["arquivo"] == "cp.csv"
    assert cp["arquivo_original"] == "Contas a pagar.csv"
    assert cp["filtros_tela"]["localidade"] == "Todas" and "total" not in cp["filtros_tela"]
    assert m["data_extracao"] == min(r["disparo"] for r in m["relatorios"].values())
    assert m["origem"] == "Esferas (Chrome)"
    assert m["rra_mensagem"] == "msg1"
    assert _estado_em_disco(ex)["fechada_em"]


def test_data_extracao_e_o_primeiro_disparo_mesmo_com_fusos_diferentes(tmp_path):
    """Comparar os textos ISO daria errado com fusos diferentes; compara-se o instante."""
    # horários no passado: os arquivos de teste (criados agora) ficam depois dos disparos
    horarios = iter([dt.datetime(2026, 1, 5, 12, 0, tzinfo=dt.timezone.utc)]
                    + [dt.datetime(2026, 1, 5, 9, 0, s, tzinfo=FUSO_BRASILIA) if s % 2
                       else dt.datetime(2026, 1, 5, 12, 0, s, tzinfo=dt.timezone.utc) for s in range(1, 60)])
    ex = extracao_fechada(tmp_path, agora=lambda: next(horarios))
    m = carregar_manifesto(ex.pasta)
    instantes = [dt.datetime.fromisoformat(r["disparo"]) for r in m["relatorios"].values()]
    assert dt.datetime.fromisoformat(m["data_extracao"]) == min(instantes)
    assert m["data_extracao"].endswith("-03:00")


def test_extracao_fechada_nao_aceita_mudancas(tmp_path, print_):
    ex = extracao_fechada(tmp_path)
    with pytest.raises(ExtracaoErro, match="fechada"):
        ex.fechar()
    with pytest.raises(ExtracaoErro, match="fechada"):
        Extracao.abrir(ex.pasta).rra_ids_antes([])
    with pytest.raises(ExtracaoErro, match="fechada"):
        Extracao.abrir(ex.pasta).registrar_tela("cp", tela("cp"), print_)
    assert carregar_manifesto(ex.pasta)


def test_manifesto_com_arquivo_adulterado_e_recusado(tmp_path):
    ex = extracao_fechada(tmp_path)
    (ex.pasta / "cp.csv").write_bytes(b"\xef\xbb\xbfoutro;conteudo\n")
    with pytest.raises(ExtracaoErro, match="não confere"):
        carregar_manifesto(ex.pasta)


def test_manifesto_com_arquivo_do_mesmo_tamanho_adulterado(tmp_path):
    ex = extracao_fechada(tmp_path)
    original = (ex.pasta / "cr.csv").read_bytes()
    (ex.pasta / "cr.csv").write_bytes(original.replace(b"10", b"99"))
    with pytest.raises(ExtracaoErro, match="não confere"):
        carregar_manifesto(ex.pasta)


def test_manifesto_com_arquivo_sumido(tmp_path):
    ex = extracao_fechada(tmp_path)
    (ex.pasta / "comissoes.xlsx").unlink()
    with pytest.raises(ExtracaoErro, match="não encontrado"):
        carregar_manifesto(ex.pasta)


def test_manifesto_com_print_sumido_ou_trocado(tmp_path):
    ex = extracao_fechada(tmp_path)
    (ex.pasta / "prints" / "cr.png").write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x01" * 16)
    with pytest.raises(ExtracaoErro, match="não confere"):
        carregar_manifesto(ex.pasta)
    (ex.pasta / "prints" / "cr.png").unlink()
    with pytest.raises(ExtracaoErro, match="não encontrado"):
        carregar_manifesto(ex.pasta)


@pytest.mark.parametrize("estrago", [
    lambda m: m["relatorios"]["cp"].update(arquivo="../fora.csv"),
    lambda m: m["relatorios"]["cp"].update(arquivo="C:/fora.csv"),
    lambda m: m["relatorios"]["cp"].pop("sha256"),
    lambda m: m["relatorios"]["cp"].update(tamanho="12"),
    lambda m: m["relatorios"].pop("rra"),
    lambda m: m.update(versao=2),
    lambda m: m.update(relatorios=[]),
    lambda m: m.pop("data_extracao"),
])
def test_manifesto_malformado(tmp_path, estrago):
    ex = extracao_fechada(tmp_path)
    caminho = ex.pasta / ARQ_MANIFESTO
    m = json.loads(caminho.read_text(encoding="utf-8"))
    estrago(m)
    caminho.write_text(json.dumps(m), encoding="utf-8")
    with pytest.raises(ExtracaoErro):
        carregar_manifesto(ex.pasta)


def test_pasta_sem_manifesto_ou_ilegivel(tmp_path):
    with pytest.raises(ExtracaoErro, match="não está fechada"):
        carregar_manifesto(tmp_path)
    (tmp_path / ARQ_MANIFESTO).write_text("{ruim", encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="ilegível"):
        carregar_manifesto(tmp_path)
    (tmp_path / ARQ_MANIFESTO).write_text("[]", encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="ilegível|inválido"):
        carregar_manifesto(tmp_path)


def test_status_resume_o_que_ja_foi_feito(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    s = ex.status()
    assert s["relatorios"]["cp"] == {"tela": True, "disparo": True, "arquivo": False}
    assert s["relatorios"]["cr"] == {"tela": False, "disparo": False, "arquivo": False}
    assert s["rra_ids_antes"] is False and s["rra_mensagem"] is False and s["fechada"] is False
