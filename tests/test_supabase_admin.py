"""`supabase_admin` só com API falsa: nenhum teste chama a API real nem usa o token real
(token e ref abaixo são valores inventados)."""
import json
import logging

import httpx
import pytest

from papaki_dre.db import mensagem_segura
from papaki_dre.supabase_admin import API, AdminErro, AdminSupabase, cliente_http, ref_do_projeto

TOKEN = "sbp_valor_falso_para_teste_123"
REF = "refinventadoabcdefgh"


class Resp:
    def __init__(self, status=200, corpo=None, json_invalido=False):
        self.status_code = status
        self._corpo = {} if corpo is None else corpo
        self._json_invalido = json_invalido

    def json(self):
        if self._json_invalido:
            # como o httpx: a mensagem do erro de decodificação traz trechos do corpo
            raise json.JSONDecodeError(f"corpo com {TOKEN} e {REF}", "doc", 0)
        return self._corpo


class FakeHttp:
    """GET devolve os estados em sequência (o último se repete); POST devolve `post_status`."""

    def __init__(self, estados, post_status=200, get_status=200, erro_rede=None, corpo_get=None,
                 json_invalido=False):
        self.estados, self.post_status, self.get_status, self.erro_rede = list(estados), post_status, get_status, erro_rede
        self.corpo_get, self.json_invalido = corpo_get, json_invalido
        self.chamadas = []

    def request(self, metodo, caminho, headers=None):
        self.chamadas.append((metodo, caminho, headers))
        if self.erro_rede:
            raise self.erro_rede
        if metodo == "GET":
            estado = self.estados.pop(0) if len(self.estados) > 1 else self.estados[0]
            corpo = self.corpo_get if self.corpo_get is not None else {"status": estado}
            return Resp(self.get_status, corpo, self.json_invalido)
        return Resp(self.post_status, {})


class Relogio:
    def __init__(self, anda=True):
        self.t = 0.0
        self.anda = anda
        self.sonos = 0

    def __call__(self):
        return self.t

    def dormir(self, s):
        self.sonos += 1
        if self.anda:
            self.t += s


def _admin(http, relogio=None):
    relogio = relogio or Relogio()
    return AdminSupabase(TOKEN, REF, http=http, dormir=relogio.dormir, relogio=relogio), relogio


def _posts(http):
    return [c for c in http.chamadas if c[0] == "POST"]


def _limpa(texto):
    assert TOKEN not in texto and REF not in texto and "://" not in texto and "supabase.co" not in texto
    assert "<html>" not in texto and "/v1/" not in texto


def test_ref_do_projeto():
    assert ref_do_projeto(f"https://{REF}.supabase.co") == REF
    assert ref_do_projeto(f"https://{REF}.supabase.co/") == REF


@pytest.mark.parametrize("url", ["", None, "nao-e-url", "https://exemplo.com.br", f"https://{REF}.supabase.co.exemplo.com",
                                 "https://ref_com_simbolo!.supabase.co"])
def test_ref_do_projeto_invalido_nao_cita_url(url):
    with pytest.raises(AdminErro) as e:
        ref_do_projeto(url)
    if url:
        assert url not in str(e.value)
    _limpa(str(e.value))


def test_sem_token_ou_ref_recusa_antes_de_chamar():
    http = FakeHttp(["INACTIVE"])
    with pytest.raises(AdminErro, match="SUPABASE_ACCESS_TOKEN"):
        AdminSupabase("", REF, http=http)
    with pytest.raises(AdminErro):
        AdminSupabase(TOKEN, "../outro", http=http)
    assert http.chamadas == []


def test_ja_ativo_nao_restaura():
    http = FakeHttp(["ACTIVE_HEALTHY"])
    admin, _ = _admin(http)
    assert admin.acordar() == "ja_ativo"
    assert _posts(http) == []


def test_pausado_restaura_uma_vez_e_espera_ficar_ativo():
    http = FakeHttp(["INACTIVE", "INACTIVE", "RESTORING", "COMING_UP", "ACTIVE_HEALTHY"])
    admin, _ = _admin(http)
    assert admin.acordar() == "restaurado"
    assert len(_posts(http)) == 1
    assert _posts(http)[0][1] == f"/v1/projects/{REF}/restore"
    assert http.chamadas[0][1] == f"/v1/projects/{REF}"
    assert http.chamadas[0][2] == {"Authorization": f"Bearer {TOKEN}"}


def test_restaurando_so_espera_sem_pedir_de_novo():
    http = FakeHttp(["RESTORING", "ACTIVE_HEALTHY"])
    admin, _ = _admin(http)
    assert admin.acordar() == "aguardado"
    assert _posts(http) == []


@pytest.mark.parametrize("estado", ["PAUSING", "GOING_DOWN", "UNKNOWN", "ESTADO_NOVO_QUALQUER"])
def test_estado_que_nao_e_pausa_nao_restaura_e_para_na_hora(estado):
    """Restaurar só com pausa confirmada; estado desconhecido não vira espera até o tempo esgotar."""
    http = FakeHttp([estado])
    admin, relogio = _admin(http)
    with pytest.raises(AdminErro) as e:
        admin.acordar()
    assert _posts(http) == [] and relogio.sonos == 0 and len(http.chamadas) == 1
    _limpa(str(e.value))


def test_estado_desconhecido_durante_a_espera_para_na_hora():
    http = FakeHttp(["INACTIVE", "RESTORING", "ESTADO_NOVO_QUALQUER"])
    admin, relogio = _admin(http)
    with pytest.raises(AdminErro, match="inesperado"):
        admin.acordar()
    assert len(_posts(http)) == 1 and relogio.sonos == 1


def test_tempo_esgotado():
    http = FakeHttp(["INACTIVE", "RESTORING"])
    admin, relogio = _admin(http)
    with pytest.raises(AdminErro, match="não ficou ativo em 10 min"):
        admin.acordar(espera_max=600, intervalo=15)
    assert relogio.t >= 600
    assert len(_posts(http)) == 1          # nunca pede a restauração de novo


def test_continua_inativo_depois_do_pedido_nao_pede_de_novo():
    http = FakeHttp(["INACTIVE"])
    admin, _ = _admin(http)
    with pytest.raises(AdminErro, match="não ficou ativo"):
        admin.acordar(espera_max=60, intervalo=15)
    assert len(_posts(http)) == 1


def test_espera_limitada_mesmo_se_o_relogio_nao_andar():
    """O limite não depende só do relógio: com relógio parado, para pelo número de consultas."""
    http = FakeHttp(["INACTIVE", "RESTORING"])
    admin, relogio = _admin(http, Relogio(anda=False))
    with pytest.raises(AdminErro, match="não ficou ativo"):
        admin.acordar(espera_max=600, intervalo=15)
    assert relogio.sonos == 40
    assert len([c for c in http.chamadas if c[0] == "GET"]) == 42   # inicial + 41 na espera


@pytest.mark.parametrize("espera_max,intervalo", [(0, 15), (600, 0), (-1, 15), (600, -5)])
def test_parametros_de_espera_invalidos(espera_max, intervalo):
    http = FakeHttp(["INACTIVE"])
    admin, _ = _admin(http)
    with pytest.raises(ValueError):
        admin.acordar(espera_max=espera_max, intervalo=intervalo)
    assert http.chamadas == []


@pytest.mark.parametrize("estado", ["RESTORE_FAILED", "INIT_FAILED", "REMOVED", "PAUSE_FAILED"])
def test_estados_de_falha_param_na_hora(estado):
    http = FakeHttp(["INACTIVE", estado])
    admin, _ = _admin(http)
    with pytest.raises(AdminErro, match=estado):
        admin.acordar()


@pytest.mark.parametrize("estado", ["RESTORE_FAILED", "REMOVED"])
def test_estado_de_falha_logo_de_inicio_nao_restaura(estado):
    http = FakeHttp([estado])
    admin, _ = _admin(http)
    with pytest.raises(AdminErro, match=estado):
        admin.acordar()
    assert _posts(http) == []


@pytest.mark.parametrize("status", [401, 403])
def test_chave_recusada_nao_vaza_nada(status):
    admin, _ = _admin(FakeHttp(["INACTIVE"], get_status=status))
    with pytest.raises(AdminErro) as e:
        admin.acordar()
    assert str(status) in str(e.value)
    _limpa(str(e.value))


def test_erro_de_rede_so_diz_o_tipo():
    admin, _ = _admin(FakeHttp(["x"], erro_rede=ConnectionError(f"falhou em https://{REF}.supabase.co {TOKEN}")))
    with pytest.raises(AdminErro) as e:
        admin.estado()
    assert "ConnectionError" in str(e.value)
    _limpa(str(e.value))
    assert e.value.__context__ is None and e.value.__cause__ is None


def test_erro_http_do_restore():
    admin, _ = _admin(FakeHttp(["INACTIVE"], post_status=500))
    with pytest.raises(AdminErro, match="HTTP 500") as e:
        admin.acordar()
    _limpa(str(e.value))


def test_resposta_sem_json_valido_vira_admin_erro_limpo():
    admin, _ = _admin(FakeHttp(["INACTIVE"], json_invalido=True))
    with pytest.raises(AdminErro) as e:
        admin.estado()
    _limpa(str(e.value))
    assert e.value.__context__ is None and e.value.__cause__ is None


@pytest.mark.parametrize("corpo", [[], "texto", {"sem_status": 1}, {"status": 3}, {"status": ""},
                                   {"status": f"https://{REF}.supabase.co {TOKEN}"}])
def test_corpo_inesperado_nao_vaza_e_nao_restaura(corpo):
    http = FakeHttp(["x"], corpo_get=corpo)
    admin, _ = _admin(http)
    with pytest.raises(AdminErro) as e:
        admin.acordar()
    _limpa(str(e.value))
    assert _posts(http) == []


def test_mensagens_passam_por_mensagem_segura(monkeypatch):
    monkeypatch.setenv("SUPABASE_ACCESS_TOKEN", TOKEN)
    monkeypatch.setenv("SUPABASE_URL", f"https://{REF}.supabase.co")
    admin, _ = _admin(FakeHttp(["INACTIVE"], get_status=401))
    with pytest.raises(AdminErro) as e:
        admin.acordar()
    assert mensagem_segura(e.value).startswith("AdminErro: ")
    _limpa(mensagem_segura(e.value))


def test_token_nunca_aparece_em_mensagem_segura(monkeypatch):
    monkeypatch.setenv("SUPABASE_ACCESS_TOKEN", TOKEN)
    assert TOKEN not in mensagem_segura(RuntimeError(f"erro com {TOKEN} no meio"))


def test_cliente_httpx_de_verdade_com_transporte_falso():
    """Exercita o httpx real (base_url, cabeçalho, .json()) sem rede: MockTransport responde."""
    vistos = []
    estados = iter(["INACTIVE", "ACTIVE_HEALTHY"])

    def responder(req: httpx.Request):
        vistos.append((req.method, req.url.host, req.url.path, req.headers.get("authorization")))
        if req.method == "GET":
            return httpx.Response(200, json={"status": next(estados)})
        return httpx.Response(200, json={})

    with cliente_http(transport=httpx.MockTransport(responder)) as http:
        admin, _ = _admin(http)
        assert admin.acordar() == "restaurado"
    assert API == "https://api.supabase.com"
    assert vistos == [("GET", "api.supabase.com", f"/v1/projects/{REF}", f"Bearer {TOKEN}"),
                      ("POST", "api.supabase.com", f"/v1/projects/{REF}/restore", f"Bearer {TOKEN}"),
                      ("GET", "api.supabase.com", f"/v1/projects/{REF}", f"Bearer {TOKEN}")]


def test_httpx_sem_json_e_erro_de_conexao_reais_nao_vazam():
    def sem_json(req):
        return httpx.Response(200, text=f"<html>{TOKEN} {REF}</html>")

    with cliente_http(transport=httpx.MockTransport(sem_json)) as http:
        admin, _ = _admin(http)
        with pytest.raises(AdminErro) as e:
            admin.estado()
        _limpa(str(e.value))

    def sem_rede(req):
        raise httpx.ConnectError(f"falhou {req.url}", request=req)

    with cliente_http(transport=httpx.MockTransport(sem_rede)) as http:
        admin, _ = _admin(http)
        with pytest.raises(AdminErro, match="ConnectError") as e:
            admin.estado()
        _limpa(str(e.value))


def test_cliente_padrao_nao_segue_redirecionamento():
    """Um redirecionamento não deve levar o token para outro host."""
    def redireciona(req):
        return httpx.Response(302, headers={"location": "https://outro.exemplo/x"})

    with cliente_http(transport=httpx.MockTransport(redireciona)) as http:
        admin, _ = _admin(http)
        with pytest.raises(AdminErro, match="HTTP 302") as e:
            admin.estado()
        _limpa(str(e.value))


# --- revisão: traceback sem segredos, falha do POST, ACTIVE_UNHEALTHY, logger do httpx ---------------

def _sem_cadeia(exc):
    assert exc.__context__ is None and exc.__cause__ is None


def test_erro_de_rede_httpx_nao_prende_token_no_traceback():
    """A exceção original do httpx traz request.headers (token) e url (ref): não pode ficar encadeada."""
    def sem_rede(req):
        raise httpx.ConnectError("falhou", request=req)

    with cliente_http(transport=httpx.MockTransport(sem_rede)) as http:
        admin, _ = _admin(http)
        with pytest.raises(AdminErro) as e:
            admin.estado()
    _sem_cadeia(e.value)


def test_json_invalido_httpx_nao_prende_corpo_no_traceback():
    def sem_json(req):
        return httpx.Response(200, text=f"<html>{TOKEN}</html>")

    with cliente_http(transport=httpx.MockTransport(sem_json)) as http:
        admin, _ = _admin(http)
        with pytest.raises(AdminErro) as e:
            admin.estado()
    _sem_cadeia(e.value)


def test_falhas_do_fake_tambem_sem_cadeia():
    admin, _ = _admin(FakeHttp(["x"], erro_rede=ConnectionError(TOKEN)))
    with pytest.raises(AdminErro) as e:
        admin.estado()
    _sem_cadeia(e.value)
    admin, _ = _admin(FakeHttp(["INACTIVE"], json_invalido=True))
    with pytest.raises(AdminErro) as e:
        admin.estado()
    _sem_cadeia(e.value)


def test_post_429_e_continua_pausado_falha_com_texto_proprio():
    http = FakeHttp(["INACTIVE"], post_status=429)
    admin, relogio = _admin(http)
    with pytest.raises(AdminErro, match="limite de pedidos da API; tente de novo em alguns minutos") as e:
        admin.acordar()
    _limpa(str(e.value))
    assert len(_posts(http)) == 1 and relogio.sonos == 0
    assert [c[0] for c in http.chamadas] == ["GET", "POST", "GET"]   # consultou de novo antes de falhar


@pytest.mark.parametrize("status", [400, 409, 500])
def test_post_falha_e_continua_pausado_orienta_o_painel(status):
    http = FakeHttp(["INACTIVE"], post_status=status)
    admin, _ = _admin(http)
    with pytest.raises(AdminErro, match="painel") as e:
        admin.acordar()
    assert f"HTTP {status}" in str(e.value) and "90 dias" in str(e.value)
    _limpa(str(e.value))
    assert len(_posts(http)) == 1


@pytest.mark.parametrize("depois", ["RESTORING", "COMING_UP"])
def test_post_falha_mas_outro_ja_pediu_segue_esperando_sem_novo_post(depois):
    """Outra execução (ou o painel) já pediu: o POST falha, mas o projeto já está subindo.
    O rótulo é 'aguardado': o nosso pedido não foi aceito, só esperamos o que já estava em curso."""
    http = FakeHttp(["INACTIVE", depois, "COMING_UP", "ACTIVE_HEALTHY"], post_status=400)
    admin, _ = _admin(http)
    assert admin.acordar() == "aguardado"
    assert len(_posts(http)) == 1


def test_post_falha_mas_ja_ficou_ativo():
    http = FakeHttp(["INACTIVE", "ACTIVE_HEALTHY"], post_status=429)
    admin, relogio = _admin(http)
    assert admin.acordar() == "aguardado"
    assert len(_posts(http)) == 1 and relogio.sonos == 0


def test_post_erro_de_rede_consulta_de_novo_sem_novo_post():
    class RedeCaiNoPost(FakeHttp):
        def request(self, metodo, caminho, headers=None):
            if metodo == "POST":
                self.chamadas.append((metodo, caminho, headers))
                raise ConnectionError(TOKEN)
            return super().request(metodo, caminho, headers)

    http = RedeCaiNoPost(["INACTIVE", "RESTORING", "ACTIVE_HEALTHY"])
    admin, _ = _admin(http)
    assert admin.acordar() == "aguardado"
    assert len(_posts(http)) == 1


@pytest.mark.parametrize("status", [401, 403])
def test_post_com_chave_recusada_para_na_hora(status):
    http = FakeHttp(["INACTIVE"], post_status=status)
    admin, _ = _admin(http)
    with pytest.raises(AdminErro, match=f"HTTP {status}") as e:
        admin.acordar()
    _sem_cadeia(e.value)
    assert len(_posts(http)) == 1 and len(http.chamadas) == 2


def test_active_unhealthy_no_inicio_espera_sem_restaurar():
    http = FakeHttp(["ACTIVE_UNHEALTHY", "ACTIVE_UNHEALTHY", "ACTIVE_HEALTHY"])
    admin, _ = _admin(http)
    assert admin.acordar() == "aguardado"
    assert _posts(http) == []


def test_active_unhealthy_persistente_falha_com_mensagem_propria():
    http = FakeHttp(["ACTIVE_UNHEALTHY"])
    admin, relogio = _admin(http)
    with pytest.raises(AdminErro, match="no ar mas não está saudável") as e:
        admin.acordar(espera_max=60, intervalo=15)
    _limpa(str(e.value))
    assert _posts(http) == [] and relogio.t >= 60


def test_cliente_http_silencia_o_log_do_httpx(caplog):
    logging.getLogger("httpx").setLevel(logging.INFO)
    estados = iter(["ACTIVE_HEALTHY"])

    def responder(req):
        return httpx.Response(200, json={"status": next(estados)})

    with caplog.at_level(logging.INFO):
        with cliente_http(transport=httpx.MockTransport(responder)) as http:
            assert logging.getLogger("httpx").level >= logging.WARNING
            admin, _ = _admin(http)
            assert admin.acordar() == "ja_ativo"
    assert REF not in caplog.text and TOKEN not in caplog.text
