import datetime as dt

import pytest

from papaki_dre.config import FUSO_BRASILIA
from papaki_dre.rra_email import RraEmailErro, SemMensagemNova, escolher_mensagem, horario_solicitado

A = "@"  # arroba separada: o texto-fonte não casa com o padrão de e-mail da auditoria de publicação
DISPARO = dt.datetime(2026, 11, 4, 9, 30, 0, tzinfo=FUSO_BRASILIA)


def msg(id_, quando="04/11/2026 09:30:20", remetente="Papaki <notas@papaki.com.br>",
        assunto="Relatório de rentabilidade gerado com sucesso"):
    return {"id": id_, "remetente": remetente, "assunto": assunto,
            "corpo": f"O relatório solicitado em {quando} está anexo."}


def test_horario_solicitado():
    assert horario_solicitado("... solicitado em 04/11/2026 09:31:07 ...") == \
        dt.datetime(2026, 11, 4, 9, 31, 7, tzinfo=FUSO_BRASILIA)
    assert horario_solicitado("sem horário") is None
    assert horario_solicitado("solicitado em 31/02/2026 09:31:07") is None


def test_escolhe_a_unica_mensagem_nova_e_no_horario():
    assert escolher_mensagem([msg("a1")], ids_antes=[], disparo=DISPARO) == "a1"


def test_ignora_mensagem_que_ja_existia_antes_do_disparo():
    with pytest.raises(SemMensagemNova):
        escolher_mensagem([msg("velha")], ids_antes=["velha"], disparo=DISPARO)


def test_ignora_remetente_e_assunto_errados():
    outras = [msg("b1", remetente=f"alguem{A}outro.com"), msg("b2", assunto="Outro assunto")]
    with pytest.raises(SemMensagemNova):
        escolher_mensagem(outras, ids_antes=[], disparo=DISPARO)


@pytest.mark.parametrize("remetente", [
    "Papaki <notas@papaki.com.br>",
    "notas@papaki.com.br",
    "<notas@papaki.com.br>",
    "NOTAS@Papaki.com.br",
    '"Papaki, Notas" <notas@papaki.com.br>',
])
def test_remetente_exato_aceito_em_varios_formatos(remetente):
    assert escolher_mensagem([msg("g1", remetente=remetente)], ids_antes=[], disparo=DISPARO) == "g1"


@pytest.mark.parametrize("remetente", [
    f"notas{A}papaki.com.br.falso.com",
    f"Papaki <notas{A}papaki.com.br.falso.com>",
    f"xnotas{A}papaki.com.br",
    f"Papaki <xnotas{A}papaki.com.br>",
    f"notas{A}papaki.com.brx",
    f"notas{A}sub.papaki.com.br",
    "notas@papaki.com.br@falso.com",
    "notas@papaki.com.br <golpe@falso.com>",
    "Fulano notas@papaki.com.br",
    "notas@papaki.com.br, golpe@falso.com",
    "golpe@falso.com, notas@papaki.com.br",
    "notas@papaki.com.br, notas@papaki.com.br",
    "",
    None,
])
def test_remetente_parecido_nao_passa(remetente):
    with pytest.raises(SemMensagemNova):
        escolher_mensagem([msg("h1", remetente=remetente)], ids_antes=[], disparo=DISPARO)


@pytest.mark.parametrize("quando", ["04/11/2026 09:29:50", "04/11/2026 09:30:00", "04/11/2026 09:33:00"])
def test_horarios_dentro_da_janela(quando):
    assert escolher_mensagem([msg("c1", quando)], ids_antes=[], disparo=DISPARO) == "c1"


@pytest.mark.parametrize("quando", ["04/11/2026 09:29:49", "04/11/2026 09:33:01", "03/11/2026 09:30:00"])
def test_horarios_fora_da_janela(quando):
    with pytest.raises(RraEmailErro, match="fora da janela do disparo; pare e avise o usuário") as e:
        escolher_mensagem([msg("d1", quando)], ids_antes=[], disparo=DISPARO)
    assert not isinstance(e.value, SemMensagemNova)


def test_fora_da_janela_nunca_e_escolhida_mas_nao_atrapalha_a_de_dentro():
    lista = [msg("fora", "04/11/2026 09:40:00"), msg("dentro", "04/11/2026 09:30:30")]
    assert escolher_mensagem(lista, ids_antes=[], disparo=DISPARO) == "dentro"


def test_fora_da_janela_so_conta_se_remetente_e_assunto_conferem():
    lista = [msg("x", "04/11/2026 09:40:00", remetente="golpe@falso.com"),
             msg("y", "04/11/2026 09:40:00", assunto="outro assunto"),
             msg("velha", "04/11/2026 09:40:00")]
    with pytest.raises(SemMensagemNova):
        escolher_mensagem(lista, ids_antes=["velha"], disparo=DISPARO)


def test_horario_do_corpo_e_de_brasilia_mesmo_com_disparo_em_outro_fuso():
    # 09:30 em Brasília == 12:30 UTC
    disparo_utc = dt.datetime(2026, 11, 4, 12, 30, 0, tzinfo=dt.timezone.utc)
    assert escolher_mensagem([msg("u1")], ids_antes=[], disparo=disparo_utc) == "u1"


def test_disparo_sem_fuso_e_recusado():
    with pytest.raises(RraEmailErro):
        escolher_mensagem([msg("n1")], ids_antes=[], disparo=dt.datetime(2026, 11, 4, 9, 30, 0))


def test_duas_candidatas_validas_param_sem_escolher():
    with pytest.raises(RraEmailErro, match="mais de uma") as e:
        escolher_mensagem([msg("e1"), msg("e2", "04/11/2026 09:30:40")], ids_antes=[], disparo=DISPARO)
    assert not isinstance(e.value, SemMensagemNova)


def test_duas_candidatas_com_o_mesmo_id_contam_como_uma():
    assert escolher_mensagem([msg("m1"), msg("m1")], ids_antes=[], disparo=DISPARO) == "m1"


def test_candidata_sem_id_e_ignorada():
    sem_id = msg("x")
    sem_id["id"] = None
    assert escolher_mensagem([sem_id, msg("ok")], ids_antes=[], disparo=DISPARO) == "ok"


def test_candidata_invalida_nao_atrapalha_a_valida():
    lista = [msg("z1", remetente=f"notas{A}papaki.com.br.falso.com"), msg("z2")]
    assert escolher_mensagem(lista, ids_antes=[], disparo=DISPARO) == "z2"


def test_mensagem_sem_horario_no_corpo_nao_serve():
    sem = {"id": "f1", "remetente": "notas@papaki.com.br",
           "assunto": "Relatório de rentabilidade gerado com sucesso", "corpo": "sem horário"}
    with pytest.raises(SemMensagemNova):
        escolher_mensagem([sem], ids_antes=[], disparo=DISPARO)


@pytest.mark.parametrize("lixo", [None, "texto", 42, ["a"], ("a",), {"id": ["x"], "remetente": "notas@papaki.com.br"}])
def test_candidata_que_nao_e_dict_ou_malformada_e_ignorada(lixo):
    assert escolher_mensagem([lixo, msg("ok")], ids_antes=[], disparo=DISPARO) == "ok"


@pytest.mark.parametrize("id_", [123, 1.5, b"x", ("a",), ["a"], {"a": 1}])
def test_id_que_nao_e_texto_e_ignorado(id_):
    ruim = msg("x")
    ruim["id"] = id_
    with pytest.raises(SemMensagemNova):
        escolher_mensagem([ruim], ids_antes=[], disparo=DISPARO)


@pytest.mark.parametrize("campo", ["corpo", "assunto", "remetente"])
@pytest.mark.parametrize("valor", [None, 5, b"bytes", ["a"], {"a": 1}])
def test_campo_que_nao_e_texto_nao_estoura_e_nao_serve(campo, valor):
    ruim = msg("r1")
    ruim[campo] = valor
    with pytest.raises(SemMensagemNova):
        escolher_mensagem([ruim], ids_antes=[], disparo=DISPARO)


def test_horario_solicitado_com_corpo_que_nao_e_texto():
    assert horario_solicitado(None) is None
    assert horario_solicitado(5) is None
    assert horario_solicitado(["solicitado em 04/11/2026 09:31:07"]) is None


def test_ids_antes_none_vale_como_vazio():
    assert escolher_mensagem([msg("a1")], ids_antes=None, disparo=DISPARO) == "a1"


def test_ids_antes_com_itens_estranhos_nao_estoura():
    assert escolher_mensagem([msg("a1")], ids_antes=[["lista"], None, "outro"], disparo=DISPARO) == "a1"
    with pytest.raises(SemMensagemNova):
        escolher_mensagem([msg("a1")], ids_antes=[["lista"], "a1"], disparo=DISPARO)
