"""Estado e restauração do projeto Supabase pela API de gerenciamento (plano gratuito: o projeto pausa
após ~7 dias sem uso). Usa SUPABASE_ACCESS_TOKEN (chave pessoal, só no .env).

Regras (spec §4.2 e §5):
- só pede restauração quando a API CONFIRMA que o projeto está pausado, e no máximo UMA vez por chamada;
  se o pedido falhar (429, 4xx porque outra execução/o painel já pediu, rede), consulta o estado de novo
  antes de decidir: já subindo/ativo → só espera; ainda pausado → erro com orientação;
- qualquer outro estado que não seja ativo, subindo ou falha conhecida para na hora, sem restaurar;
- a espera tem limite de tempo E de número de consultas (não depende só do relógio);
- as mensagens de erro são montadas só com textos fixos e estados conhecidos: nunca levam o token, o ref
  do projeto, a URL nem o corpo da resposta da API; e são lançadas FORA dos blocos `except`, para a
  exceção original (com `request.headers`, `url` e corpo) não ficar encadeada no traceback;
- o logger `httpx` fica em WARNING (no nível INFO ele registra a URL, que contém o ref)."""
from __future__ import annotations

import logging
import math
import re
import time
from urllib.parse import urlparse

import httpx

API = "https://api.supabase.com"
ATIVO = "ACTIVE_HEALTHY"
NAO_SAUDAVEL = "ACTIVE_UNHEALTHY"
PAUSADO = frozenset({"INACTIVE", "PAUSED"})
# Já está no ar ou subindo (ninguém precisa pedir restauração): só espera.
SUBINDO = frozenset({"RESTORING", "COMING_UP", "RESTARTING", "UPGRADING", "RESIZING", NAO_SAUDAVEL})
# Durante a espera também aceita pausado (o pedido de restauração ainda não refletiu).
ESPERAVEL = SUBINDO | PAUSADO
FALHA = frozenset({"RESTORE_FAILED", "INIT_FAILED", "PAUSE_FAILED", "REMOVED"})
CONHECIDOS = ESPERAVEL | FALHA | {ATIVO, "PAUSING", "GOING_DOWN", "UNKNOWN"}
_FORMATO_ESTADO = re.compile(r"[A-Z][A-Z_]{0,39}")
_FORMATO_REF = re.compile(r"[a-z0-9]{1,64}")
_LOGGERS_SILENCIADOS = ("httpx", "httpcore")


class AdminErro(RuntimeError):
    """Falha ao consultar ou restaurar o projeto (mensagem já limpa). `status`: código HTTP, se houve."""

    def __init__(self, mensagem: str, status: int | None = None, chave_recusada: bool = False):
        super().__init__(mensagem)
        self.status = status
        self.chave_recusada = chave_recusada


def ref_do_projeto(url: str | None) -> str:
    """Ref do projeto a partir de SUPABASE_URL (`https://<ref>.supabase.co`). O erro não cita a URL."""
    try:
        host = (urlparse(url or "").hostname or "").lower()
    except ValueError:
        host = ""
    partes = host.split(".")
    if len(partes) != 3 or partes[1:] != ["supabase", "co"] or not _FORMATO_REF.fullmatch(partes[0]):
        raise AdminErro("SUPABASE_URL inválida: não foi possível identificar o projeto")
    return partes[0]


def cliente_http(transport=None) -> httpx.Client:
    """Cliente da API de gerenciamento. Não segue redirecionamentos (o token não vai para outro host) e
    deixa o log do httpx em WARNING (o INFO registraria a URL com o ref)."""
    for nome in _LOGGERS_SILENCIADOS:
        logger = logging.getLogger(nome)
        if logger.getEffectiveLevel() < logging.WARNING:
            logger.setLevel(logging.WARNING)
    return httpx.Client(base_url=API, timeout=30, follow_redirects=False, transport=transport)


def _duracao(segundos: float) -> str:
    return f"{int(segundos) // 60} min" if segundos % 60 == 0 else f"{segundos:g} s"


class AdminSupabase:
    def __init__(self, token: str, ref: str, *, http=None, dormir=time.sleep, relogio=time.monotonic):
        if not token:
            raise AdminErro("defina SUPABASE_ACCESS_TOKEN para acordar o projeto pausado")
        if not isinstance(ref, str) or not _FORMATO_REF.fullmatch(ref):
            raise AdminErro("ref do projeto Supabase inválido")
        self.token, self.ref = token, ref
        self._proprio = http is None
        self.http = cliente_http() if http is None else http
        self.dormir, self.relogio = dormir, relogio

    def fechar(self) -> None:
        if self._proprio:
            self.http.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.fechar()

    def _erro(self, texto: str, **extra) -> AdminErro:
        """Defesa extra: mesmo que um texto fixo um dia cite algo, token e ref nunca saem."""
        for segredo in sorted((self.token, self.ref), key=len, reverse=True):
            texto = texto.replace(segredo, "***")
        return AdminErro(texto, **extra)

    def _chamar(self, metodo: str, caminho: str):
        tipo_falha = None
        try:
            r = self.http.request(metodo, caminho, headers={"Authorization": f"Bearer {self.token}"})
        except Exception as e:  # noqa: BLE001 - guarda só o tipo: a exceção do httpx traz URL e cabeçalhos
            tipo_falha = type(e).__name__
        if tipo_falha is not None:       # fora do except: nada de __context__ com o token
            raise self._erro(f"API do Supabase inacessível ({tipo_falha})")
        status = r.status_code
        if status in (401, 403):
            raise self._erro(f"a chave pessoal do Supabase foi recusada (HTTP {status})",
                             status=status, chave_recusada=True)
        if not 200 <= status < 300:          # inclui 3xx: redirecionamento não é seguido
            raise self._erro(f"a API do Supabase respondeu HTTP {status}", status=status)
        return r

    def estado(self) -> str:
        r = self._chamar("GET", f"/v1/projects/{self.ref}")
        json_ok = True
        try:
            corpo = r.json()
        except Exception:  # noqa: BLE001 - a mensagem do erro de JSON traz trechos do corpo
            json_ok = False
        if not json_ok:                  # fora do except: nada de __context__ com o corpo
            raise self._erro("a API do Supabase devolveu uma resposta que não é JSON válido")
        estado = corpo.get("status") if isinstance(corpo, dict) else None
        if not isinstance(estado, str) or not _FORMATO_ESTADO.fullmatch(estado):
            raise self._erro("a API do Supabase devolveu o estado do projeto em formato inesperado")
        return estado

    def restaurar(self) -> None:
        self._chamar("POST", f"/v1/projects/{self.ref}/restore")

    def _nao_reconhecido(self, estado: str, quando: str) -> AdminErro:
        nome = estado if estado in CONHECIDOS else "não reconhecido"
        return self._erro(f"estado inesperado do projeto {quando} ({nome}); nada foi restaurado "
                          "automaticamente — verifique o projeto pelo painel do Supabase")

    def _falha(self, estado: str) -> AdminErro:
        return self._erro(f"o Supabase informou {estado}; restaure o projeto pelo painel")

    def _pedir_restauracao(self) -> str:
        """Faz o ÚNICO pedido de restauração. Se ele falhar (exceto chave recusada), consulta o estado de
        novo: já subindo/ativo → 'aguardado' (o nosso pedido não foi aceito; outra execução ou o painel
        já tinha pedido, e só esperamos o que estava em curso); ainda pausado → erro orientado."""
        falha = None
        try:
            self.restaurar()
        except AdminErro as e:
            falha = e
        if falha is None:
            return "restaurado"
        if falha.chave_recusada:
            raise self._erro(str(falha), status=falha.status, chave_recusada=True)
        estado = self.estado()
        if estado == ATIVO or estado in SUBINDO:
            return "aguardado"
        if estado in FALHA:
            raise self._falha(estado)
        if estado in PAUSADO:
            if falha.status == 429:
                raise self._erro("o pedido de restauração foi recusado por limite de pedidos da API; "
                                 "tente de novo em alguns minutos", status=429)
            motivo = f"HTTP {falha.status}" if falha.status else "sem resposta da API"
            raise self._erro(
                f"o pedido de restauração falhou ({motivo}) e o projeto continua pausado; restaure pelo "
                "painel do Supabase (projeto pausado há mais de 90 dias ou limite de projetos ativos no "
                "plano gratuito impedem a restauração pela API)", status=falha.status)
        raise self._nao_reconhecido(estado, "depois do pedido de restauração")

    def acordar(self, espera_max: float = 600, intervalo: float = 15) -> str:
        """'ja_ativo' (nada a fazer), 'restaurado' (o nosso pedido de restauração foi aceito e esperamos)
        ou 'aguardado' (já estava subindo/no ar sem saúde, ou outro já tinha pedido a restauração).
        Só pede restauração se a API confirmar que o projeto está pausado, e no máximo uma vez.
        Espera até `espera_max` segundos, consultando a cada `intervalo`."""
        if espera_max <= 0 or intervalo <= 0:
            raise ValueError("espera_max e intervalo devem ser positivos")
        estado = self.estado()
        if estado == ATIVO:
            return "ja_ativo"
        if estado in FALHA:
            raise self._falha(estado)
        if estado in PAUSADO:
            acao = self._pedir_restauracao()
        elif estado in SUBINDO:
            acao = "aguardado"
        else:
            raise self._nao_reconhecido(estado, "antes de restaurar")

        tentativas = math.ceil(espera_max / intervalo)     # limite que não depende do relógio
        fim = self.relogio() + espera_max
        for i in range(tentativas + 1):
            estado = self.estado()
            if estado == ATIVO:
                return acao
            if estado in FALHA:
                raise self._falha(estado)
            if estado not in ESPERAVEL:
                raise self._nao_reconhecido(estado, "durante a espera")
            if i == tentativas or self.relogio() >= fim:
                break
            self.dormir(intervalo)
        if estado == NAO_SAUDAVEL:
            raise self._erro(f"o projeto está no ar mas não está saudável há {_duracao(espera_max)}; "
                             "verifique o projeto pelo painel do Supabase")
        raise self._erro(f"o projeto não ficou ativo em {_duracao(espera_max)} (último estado: {estado})")
