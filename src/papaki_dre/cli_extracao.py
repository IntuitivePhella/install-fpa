"""Subcomandos de apoio à rotina mensal: `agenda` e `extracao ...`. A saída que o Claude lê é sempre
JSON em ASCII (caminhos com acento chegam escapados e voltam intactos ao decodificar).

Não existe subcomando para abandonar uma extração nem para redisparar um relatório já disparado:
isso é decisão do usuário, fora do alcance do Claude."""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
from pathlib import Path

from . import ambiente, calendario, caminhos, produtoras
from .config import FERIADOS_EXTRAS, FUSO_BRASILIA
from .extracao import RELATORIOS, Extracao, ExtracaoErro, iniciar
from .rra_email import RraEmailErro, SemMensagemNova

EXIT_OK, EXIT_FALHA, EXIT_AGUARDAR = 0, 3, 4


def registrar(sub) -> None:
    ag = sub.add_parser("agenda", help="diz se hoje é dia de rodar a rotina mensal")
    ag.add_argument("--produtora", required=True)
    ag.add_argument("--hoje", help="AAAA-MM-DD; padrão = hoje")
    ag.add_argument("--sem-banco", action="store_true", help="não consulta o banco central")
    ag.add_argument("--restaurar-banco", action="store_true",
                    help="se o banco estiver pausado, acorda o projeto antes de consultar (exige "
                         "SUPABASE_ACCESS_TOKEN no .env)")
    ag.add_argument("--reservar", action="store_true",
                    help="se for dia de rodar, reserva a rodada no banco central para esta máquina (PAPAKI_MAQUINA); "
                         "se outra máquina já a reservou, responde 'não rodar'")

    rv = sub.add_parser("reserva", help="reserva da rodada de uma produtora e competência (usado pelo Claude)")
    rvs = rv.add_subparsers(dest="acao", required=True)
    o = rvs.add_parser("obter", help="tenta reservar a rodada para esta máquina e diz quem a tem")
    o.add_argument("--produtora", required=True)
    o.add_argument("--competencia", required=True, help="AAAA-MM")

    ex = sub.add_parser("extracao", help="registro da extração do Esferas (usado pelo Claude)")
    acoes = ex.add_subparsers(dest="acao", required=True)

    p = acoes.add_parser("iniciar", help="cria (ou, com --retomar, reabre) a pasta de uma extração")
    p.add_argument("--produtora", required=True)
    p.add_argument("--competencia", required=True, help="AAAA-MM")
    p.add_argument("--periodo-fim", help="AAAA-MM-DD, dentro da competência; padrão = último dia da competência")
    p.add_argument("--retomar", action="store_true", help="reabre a extração aberta (a que tem disparo sem "
                                                         "arquivo tem prioridade)")
    p.add_argument("--nova", action="store_true",
                   help="começa outra extração e abandona em definitivo a extração aberta anterior (ela não "
                        "aceita mais nada); recusado se algum relatório dela já foi disparado sem arquivo")

    def com_pasta(nome, ajuda, relatorio=False):
        q = acoes.add_parser(nome, help=ajuda)
        q.add_argument("--pasta", required=True, type=Path)
        if relatorio:
            q.add_argument("--relatorio", required=True, choices=RELATORIOS)
        return q

    t = com_pasta("tela", "registra e valida o que a tela mostra ANTES de exportar", relatorio=True)
    t.add_argument("--localidade", required=True)
    t.add_argument("--criterio", required=True, help="Data competência (CP, CR) ou Aprovação (RRA, Comissões)")
    t.add_argument("--data-inicio", required=True)
    t.add_argument("--data-fim", required=True)
    t.add_argument("--status", required=True, help="Todos")
    t.add_argument("--total", required=True, type=int, help="registros que a tela mostra")
    t.add_argument("--pesquisar-em", help="só no CP: Individuais e agrupadas")
    t.add_argument("--print", dest="print_", required=True, type=Path, help="imagem .png/.jpg da tela")
    com_pasta("disparar", "registra o disparo (antes do clique em Exportar/Relatórios); a tela vale por 15 min",
              relatorio=True)
    a = com_pasta("arquivo", "registra o arquivo baixado", relatorio=True)
    a.add_argument("--caminho", required=True, type=Path)
    r = com_pasta("rra-ids-antes", "registra os ids de e-mails já existentes (antes de disparar o RRA)")
    r.add_argument("--ids", required=True, help="ids separados por vírgula (vazio se não há nenhum)")
    c = com_pasta("rra-escolher", "escolhe a mensagem do RRA entre as candidatas (JSON)")
    c.add_argument("--candidatas", required=True, type=Path)
    com_pasta("fechar", "fecha a extração e gera o registro para o importar")
    com_pasta("estado", "mostra o que já foi feito (para retomar)")


def _agora() -> dt.datetime:
    return dt.datetime.now(FUSO_BRASILIA)


def _hoje(texto) -> dt.date:
    """`--hoje` no formato extenso AAAA-MM-DD; sem ele, o dia de Brasília (não o do relógio local/UTC)."""
    if texto is None:
        return _agora().date()
    try:
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", texto):
            raise ValueError
        return dt.date.fromisoformat(texto)
    except ValueError:
        raise ExtracaoErro(f"--hoje deve ser uma data real no formato AAAA-MM-DD; recebido {texto[:20]!r}") from None


def _json(dados) -> None:
    print(json.dumps(dados, ensure_ascii=True))


def executar(a, env=None, consulta_banco=None, reservar_banco=None) -> int:
    try:
        if a.cmd == "agenda":
            return _agenda(a, env, consulta_banco, reservar_banco)
        if a.cmd == "reserva":
            return _reserva(a, reservar_banco)
        return _extracao(a, env)
    except SemMensagemNova as e:
        print(f"AGUARDE: {e}", file=sys.stderr)
        return EXIT_AGUARDAR
    except (ExtracaoErro, caminhos.CaminhosErro, RraEmailErro, ValueError) as e:
        print(f"ERRO: {e}", file=sys.stderr)
        return EXIT_FALHA
    except OSError as e:                       # só o tipo e o nome do arquivo, nunca o caminho completo
        nome = Path(str(e.filename)).name if e.filename else "?"
        print(f"ERRO: falha de arquivo ({type(e).__name__}) em {nome!r}", file=sys.stderr)
        return EXIT_FALHA


def _agenda(a, env, consulta_banco=None, reservar_banco=None) -> int:
    if not produtoras.existe(a.produtora):
        raise ExtracaoErro(f"produtora desconhecida: {a.produtora!r}")
    hoje = _hoje(a.hoje)
    competencia = calendario.competencia_alvo(hoje)
    amb = ambiente.resolver(env)
    raiz = caminhos.raiz_dados(amb)
    n_ext, n_apu = caminhos.nomes_pastas(amb)
    apuracao = caminhos.pasta_apuracao(raiz, a.produtora, competencia, nome=n_apu)
    ja = calendario.competencia_concluida(apuracao / "conferencia.json", competencia)
    rodar, motivo = calendario.deve_rodar(hoje, ja, extras=FERIADOS_EXTRAS)
    banco = "pulado"
    if rodar and not a.sem_banco and consulta_banco is not None:
        _, _, fim = calendario.periodo_da_competencia(competencia)
        banco, ja_no_banco = consulta_banco(a.produtora, competencia, fim.isoformat(), a.restaurar_banco)
        if ja_no_banco:
            rodar, motivo = False, "competência já importada (banco central)"
    reserva = "nao_pedida"
    if rodar and a.reservar:
        reserva = "nao_consultada"
        if banco == "consultado" and reservar_banco is not None:
            reserva = reservar_banco(a.produtora, competencia)["estado"]
            if reserva == "de_outra_maquina":
                rodar, motivo = False, "rodada em andamento em outra máquina"
    _json({"rodar": rodar, "competencia": competencia, "motivo": motivo, "banco": banco, "reserva": reserva,
           "pasta_apuracao": str(apuracao),
           "pasta_extracoes": str(caminhos.pasta_extracoes(raiz, competencia, a.produtora, nome=n_ext))})
    return EXIT_OK


def _reserva(a, reservar_banco) -> int:
    if not produtoras.existe(a.produtora):
        raise ExtracaoErro(f"produtora desconhecida: {a.produtora!r}")
    competencia, _, _ = calendario.periodo_da_competencia(a.competencia)   # ValueError -> ERRO (código 3)
    _json({"produtora": a.produtora, "competencia": competencia, **reservar_banco(a.produtora, competencia)})
    return EXIT_OK


def _extracao(a, env) -> int:
    if a.acao == "iniciar":
        competencia, _, fim = calendario.periodo_da_competencia(a.competencia)
        amb = ambiente.resolver(env)
        raiz = caminhos.raiz_dados(amb)
        n_ext, n_apu = caminhos.nomes_pastas(amb)
        ex = iniciar(caminhos.pasta_extracoes(raiz, competencia, a.produtora, nome=n_ext), a.produtora,
                     competencia, a.periodo_fim or fim.isoformat(), retomar=a.retomar, nova=a.nova)
        _json({"pasta": str(ex.pasta),
               "pasta_apuracao": str(caminhos.pasta_apuracao(raiz, a.produtora, competencia, nome=n_apu))})
        return EXIT_OK
    ex = Extracao.abrir(a.pasta)
    if a.acao == "tela":
        ex.registrar_tela(a.relatorio, {
            "localidade": a.localidade, "criterio": a.criterio, "data_inicio": a.data_inicio,
            "data_fim": a.data_fim, "status": a.status, "total": a.total, "pesquisar_em": a.pesquisar_em}, a.print_)
        _json({"relatorio": a.relatorio, "tela": "validada"})
    elif a.acao == "disparar":
        _json({"relatorio": a.relatorio, "disparo": ex.disparar(a.relatorio)})
    elif a.acao == "arquivo":
        arq = ex.registrar_arquivo(a.relatorio, a.caminho)
        _json({"relatorio": a.relatorio, "arquivo": arq["nome"], "sha256": arq["sha256"][:12]})
    elif a.acao == "rra-ids-antes":
        ex.rra_ids_antes([i.strip() for i in a.ids.split(",") if i.strip()])
        _json({"rra_ids_antes": "registrado"})
    elif a.acao == "rra-escolher":
        try:
            candidatas = json.loads(Path(a.candidatas).read_text(encoding="utf-8"))
        except (OSError, ValueError):
            raise ExtracaoErro("arquivo de candidatas ilegível (esperado: lista JSON de mensagens)") from None
        if not isinstance(candidatas, list):
            raise ExtracaoErro("arquivo de candidatas ilegível (esperado: lista JSON de mensagens)")
        _json({"mensagem": ex.escolher_rra(candidatas)})
    elif a.acao == "fechar":
        m = ex.fechar()
        _json({"fechada": True, "pasta": str(ex.pasta), "relatorios": list(m["relatorios"])})
    elif a.acao == "estado":
        _json(ex.status())
    return EXIT_OK
