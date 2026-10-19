"""Escolha da mensagem do RRA entre as que o Claude leu no Gmail. Função pura: nada de rede."""
from __future__ import annotations

import datetime as dt
import re
from email.utils import getaddresses

from .config import FUSO_BRASILIA

REMETENTE = "notas@papaki.com.br"
ASSUNTO = "Relatório de rentabilidade gerado com sucesso"
ANTES_S = 10          # o "solicitado em" do Esferas pode ficar até 10 s antes do disparo registrado
DEPOIS_S = 180        # o disparo é gravado ANTES do clique; o pedido real vem até 3 min depois
_SOLICITADO = re.compile(r"solicitad[oa] em\s*(\d{2})/(\d{2})/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})", re.I)


class RraEmailErro(RuntimeError):
    """Não foi possível escolher a mensagem com segurança (ambígua, por exemplo)."""


class SemMensagemNova(RraEmailErro):
    """Ainda não há mensagem nova que combine: o chamador espera e tenta de novo."""


def horario_solicitado(corpo) -> dt.datetime | None:
    if not isinstance(corpo, str):
        return None
    m = _SOLICITADO.search(corpo)
    if not m:
        return None
    dia, mes, ano, hh, mm, ss = (int(x) for x in m.groups())
    try:
        return dt.datetime(ano, mes, dia, hh, mm, ss, tzinfo=FUSO_BRASILIA)
    except ValueError:
        return None


def _remetente_confere(remetente) -> bool:
    """O endereço (não o nome de exibição) é exatamente `notas@papaki.com.br`.
    Aceita "Nome <endereço>", mas só quando há um único endereço no cabeçalho."""
    if not isinstance(remetente, str) or not remetente.strip():
        return False
    enderecos = [a for _, a in getaddresses([remetente]) if a]
    return len(enderecos) == 1 and enderecos[0].strip().casefold() == REMETENTE


def escolher_mensagem(candidatas, ids_antes, disparo: dt.datetime, *,
                      antes_s: int = ANTES_S, depois_s: int = DEPOIS_S) -> str:
    """Id da única mensagem que: não existia antes do disparo, vem do endereço exato `notas@papaki.com.br`,
    tem o assunto esperado e cujo "solicitado em" (Brasília) cai entre `antes_s` antes e `depois_s` depois
    do disparo. Nenhuma -> SemMensagemNova; mais de uma (ids diferentes) -> RraEmailErro (nunca escolhe no palpite)."""
    if disparo.tzinfo is None:
        raise RraEmailErro("o horário do disparo precisa ter fuso (comparar sem fuso daria resultado errado)")
    ja_existiam = {i for i in (ids_antes or ()) if isinstance(i, str)}
    validas = set()
    fora_da_janela = False
    for c in candidatas:
        if not isinstance(c, dict):
            continue
        id_ = c.get("id")
        if not isinstance(id_, str) or not id_ or id_ in ja_existiam:
            continue
        if not _remetente_confere(c.get("remetente")):
            continue
        assunto = c.get("assunto")
        if not isinstance(assunto, str) or ASSUNTO.casefold() not in assunto.casefold():
            continue
        quando = horario_solicitado(c.get("corpo"))
        if quando is None:
            continue
        if -antes_s <= (quando - disparo).total_seconds() <= depois_s:
            validas.add(id_)
        else:
            fora_da_janela = True
    if not validas and fora_da_janela:
        # a mensagem existe, mas não combina com o disparo: esperar mais não resolve (e nunca se escolhe fora)
        raise RraEmailErro("chegou um e-mail do RRA fora da janela do disparo; pare e avise o usuário")
    if not validas:
        raise SemMensagemNova("ainda não há mensagem nova do RRA que combine com o disparo")
    if len(validas) > 1:
        raise RraEmailErro(f"mais de uma mensagem combina com o disparo do RRA (ids: {', '.join(sorted(validas))}); "
                           "pare e peça ajuda ao usuário")
    return next(iter(validas))
