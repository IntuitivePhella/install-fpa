"""Comando `importar`: sanitiza os 4 brutos, carrega a cópia do Modelo DRE, confere e grava a foto
(importação) no Supabase. Códigos de saída: 0 ok/com alertas/duplicado; 2 bloqueada; 3 falha, banco
pausado/inacessível ou erro de uso."""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
import time
import traceback
from pathlib import Path

from . import ambiente, calendario, caminhos, cli_extracao, cli_instalacao, produtoras
from . import extracao as extracao_mod
from .build import build_workbook
from .config import DEST, FUSO_BRASILIA
from .db import BancoIndisponivel, Repository, make_client, mensagem_segura, regras_texto, sha256_file
from .extracao import ExtracaoErro, carregar_manifesto
from .headers import check_header, load_headers, normalize
from .readers import read_delimited
from .recalc import contar_nd
from .report import contract_blocked, conference, render_markdown, render_resultado
from .sanitize import sanitize_comissoes, sanitize_cp, sanitize_cr, sanitize_rra
from .supabase_admin import AdminErro, AdminSupabase, ref_do_projeto
from .types import ContractError

REPORTS = ("rra", "cp", "cr", "comissoes")
EXIT_OK, EXIT_BLOCKED, EXIT_FAILED = 0, 2, 3
ND_COLUNAS = ("AP", "AQ")                 # classificação em DESPESAS (alerta de #N/D, nunca bloqueio)
MSG_PAUSADO = "banco pausado: restaure o projeto e rode de novo"
_dormir = time.sleep                      # trocáveis nos testes (sem espera real)
_relogio = time.monotonic
PRAZO_RESTAURACAO = 540                   # s: prazo GLOBAL do início da restauração até desistir (9 min)
ESPERA_ACORDAR = PRAZO_RESTAURACAO - 60   # s: espera máxima dentro de AdminSupabase.acordar (8 min)
INTERVALO_ACORDAR = 15                    # s entre as consultas de estado da API
TENTATIVAS_APOS_ACORDAR = 10
INTERVALO_APOS_ACORDAR = 6                # segundos entre as tentativas de falar com o banco


class UsageError(ValueError):
    """Argumento inválido do comando. Nada é gravado (nem banco nem arquivos)."""


class _Parser(argparse.ArgumentParser):
    """Erro de argumento vira UsageError (código 3), não o SystemExit(2) do argparse, que se
    confundiria com 'bloqueada'."""

    def __init__(self, *args, **kwargs):
        # sem abreviações: `--nov` nunca pode virar `--nova` (abandona uma extração em definitivo).
        # Os subparsers são criados com esta mesma classe, então herdam o padrão.
        kwargs["allow_abbrev"] = False
        super().__init__(*args, **kwargs)

    def error(self, message):
        raise UsageError(message)


def parse_tela(text: str | None) -> dict[str, int]:
    out: dict[str, int] = {}
    if not text:
        return out
    for part in text.split(","):
        key, _, val = part.partition("=")
        key = key.strip()
        if key not in REPORTS:
            raise UsageError(f"--tela: relatório desconhecido {key!r} (use {', '.join(REPORTS)})")
        try:
            out[key] = int(val)
        except ValueError:
            raise UsageError(f"--tela: valor inválido para {key}: {val!r}") from None
    return out


def parse_competencia(text: str) -> tuple[str, dt.date, dt.date]:
    try:
        return calendario.periodo_da_competencia(text)
    except ValueError as e:
        raise UsageError(str(e)) from None


def _parser() -> argparse.ArgumentParser:
    p = _Parser(prog="papaki_dre")
    sub = p.add_subparsers(dest="cmd", required=True)
    i = sub.add_parser("importar", help="sanitiza os 4 brutos, grava no Supabase e gera o Modelo DRE")
    i.add_argument("--produtora", required=True)
    i.add_argument("--competencia", required=True, help="AAAA-MM")
    for r in REPORTS:
        i.add_argument(f"--{r}", type=Path, help="bruto do relatório (ou use --extracao)")
    i.add_argument("--extracao", type=Path,
                   help="pasta de uma extração fechada (no lugar de --cp/--cr/--rra/--comissoes, --tela, "
                        "--data-extracao e --origem)")
    i.add_argument("--modelo", type=Path, help="modelo anual vazio (.xlsx); padrão: PAPAKI_MODELO_<PRODUTORA>")
    i.add_argument("--saida", type=Path,
                   help="pasta de saída; padrão: <PAPAKI_DADOS_DIR>/apuracoes/<produtora>/<competencia>")
    i.add_argument("--periodo-fim", help="AAAA-MM-DD; padrão = último dia da competência (ou o da extração)")
    i.add_argument("--gatilho", choices=["manual", "agendado"], default="manual")
    i.add_argument("--origem", help="padrão: 'arquivos locais'")
    i.add_argument("--data-extracao", help="ISO-8601; sem fuso vale Brasília; padrão = agora")
    i.add_argument("--tela", help="totais vistos na tela, ex.: cp=37030,cr=483,rra=1410")
    i.add_argument("--sem-banco", action="store_true")
    i.add_argument("--restaurar-banco", action="store_true",
                   help="se o banco estiver inacessível e a API do Supabase confirmar que o projeto está "
                        "pausado, acorda o projeto (exige SUPABASE_ACCESS_TOKEN e SUPABASE_URL no .env)")
    i.add_argument("--sem-recalculo", action="store_true",
                   help="não abre o Excel para contar #N/D em DESPESAS!AP:AQ")
    i.add_argument("--forcar", action="store_true", help="importa mesmo que os arquivos já tenham sido importados")
    cli_extracao.registrar(sub)
    cli_instalacao.registrar(sub)
    return p


def _out_name(nome: str, competencia: str) -> str:
    year, month = competencia.split("-")
    return f"Modelo DRE - {nome} - {month}_{year}.xlsx"


ARQ_RESULTADO = "RESULTADO.md"


def _write_conference(out_dir: Path, conf: dict) -> None:
    """Grava a conferência e tira o RESULTADO.md da pasta: o resumo de uma rodada anterior nunca fica
    contradizendo a conferência nova (as rodadas por --extracao reescrevem o resumo só depois de
    publicadas). Se não der para apagar o resumo, a conferência é gravada mesmo assim e o erro sobe."""
    erro = None
    try:
        (out_dir / ARQ_RESULTADO).unlink(missing_ok=True)
    except OSError as e:
        erro = e
    (out_dir / "conferencia.json").write_text(json.dumps(conf, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "conferencia.md").write_text(render_markdown(conf), encoding="utf-8")
    if erro is not None:
        raise erro


def _contar_nd(build, out_xlsx: Path) -> tuple[int | None, str | None]:
    """(na_count, na_error) para a conferência. Qualquer falha do Excel vira aviso, nunca bloqueio."""
    sr = build.sheets.get(DEST["cp"]["sheet"])
    if sr is None or sr.rows == 0:
        return 0, None
    try:
        return contar_nd(out_xlsx, sr.sheet, ND_COLUNAS, sr.first_row, sr.last_row), None
    except Exception as e:  # noqa: BLE001 - COM levanta tipos variados; só alerta
        return None, mensagem_segura(e)


def _marcar_falha(repo, imp_id, motivo: str) -> bool | None:
    """Marca 'falhou' sem mascarar o erro original se o próprio banco falhar. Devolve se a
    importação mudou (False = ela já não estava 'em_andamento', p.ex. promovida), ou None se o
    banco não respondeu."""
    try:
        return repo.marcar(imp_id, "falhou", motivo=motivo)
    except Exception as e:  # noqa: BLE001
        print(f"AVISO: não foi possível marcar a importação {imp_id} como 'falhou': {mensagem_segura(e)}",
              file=sys.stderr)
        return None


class ArquivoEmUso(RuntimeError):
    """O destino final está aberto (no Excel) e não pôde ser substituído."""


def _preparo(final: Path) -> Path:
    """Arquivo de preparo oculto na pasta de saída; nunca tem o nome final nem _BLOQUEADO."""
    return final.with_name(f".{final.stem}.novo.xlsx")


def _destinos(final_ok: Path) -> tuple[Path, Path, Path]:
    """Tudo o que a rodada pode gravar com o nome do Excel: final ok, _BLOQUEADO e o preparo."""
    return final_ok, final_ok.with_name(final_ok.stem + "_BLOQUEADO.xlsx"), _preparo(final_ok)


def _saida_de_rodada_real(out_dir: Path) -> bool:
    """A pasta já tem a conferência de uma importação registrada no banco (com importacao_id)."""
    try:
        conf = json.loads((out_dir / "conferencia.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False
    return isinstance(conf, dict) and bool((conf.get("meta") or {}).get("importacao_id"))


SEM_BANCO_INFO = "rodada sem banco: não registrada no Supabase"


def _sem_banco_info(conf: dict, repo) -> dict:
    if repo is None:
        conf.setdefault("informacoes", []).append(SEM_BANCO_INFO)
    return conf


def _aviso_arquivo_antigo(conf: dict, final_ok: Path) -> dict:
    """Rodada bloqueada numa pasta que ainda tem o .xlsx de uma rodada anterior: a conferência avisa
    que aquele arquivo NÃO é desta rodada."""
    if final_ok.exists():
        conf.setdefault("informacoes", []).append(
            f"a pasta ainda tem {final_ok.name} de uma rodada anterior; ele NÃO é desta rodada "
            "(que está bloqueada): não use esse arquivo")
    return conf


def _agora() -> dt.datetime:
    """Relógio da limpeza de rodadas órfãs (UTC); os testes o substituem."""
    return dt.datetime.now(dt.timezone.utc)


def _status(msg: str, *, file=None) -> None:
    """Resumo final no terminal. Um console que não aceita algum caractere (cp1252) nunca transforma
    uma rodada concluída em falha: o texto sai com '?' no lugar."""
    stream = file or sys.stdout
    try:
        print(msg, file=stream)
    except UnicodeEncodeError:
        enc = getattr(stream, "encoding", None) or "ascii"
        try:
            print(msg.encode(enc, errors="replace").decode(enc), file=stream)
        except Exception:  # noqa: BLE001 - só o resumo; a rodada já terminou
            pass


def _apagar(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass                                     # melhor esforço


def _finalizar(preparo: Path, final: Path) -> None:
    try:
        os.replace(preparo, final)
    except PermissionError:
        raise ArquivoEmUso(f"feche o arquivo {final.name} no Excel e rode de novo") from None


def _regravar_falha(out_dir: Path, conf: dict, motivo: str) -> None:
    """A pasta nunca fica com uma conferência 'ok' de uma rodada que não está no banco."""
    falha = {**conf, "status": "falhou", "motivo_falha": motivo,
             "bloqueios": [*conf.get("bloqueios", []), f"a importação falhou: {motivo}"]}
    try:
        _write_conference(out_dir, falha)
    except Exception as e:  # noqa: BLE001
        print(f"AVISO: não foi possível regravar a conferência: {mensagem_segura(e)}", file=sys.stderr)


def _iso_com_fuso(texto: str | None) -> str:
    """ISO-8601 com fuso. Sem fuso, vale o horário de Brasília. Sem data: agora (UTC, com fuso)."""
    if not texto:
        return dt.datetime.now(dt.timezone.utc).isoformat()
    try:
        valor = dt.datetime.fromisoformat(texto)
    except ValueError:
        raise UsageError(f"--data-extracao deve ser ISO-8601; recebido {texto!r}") from None
    if valor.tzinfo is None:
        valor = valor.replace(tzinfo=FUSO_BRASILIA)
    return valor.isoformat()


def _data_periodo_fim(texto: str) -> dt.date:
    try:
        return dt.date.fromisoformat(texto)
    except ValueError:
        raise UsageError(f"--periodo-fim deve ser AAAA-MM-DD; recebido {texto!r}") from None


def _nome_original(r: dict) -> str:
    """Nome do arquivo como foi baixado (vai para arquivos_brutos.nome_arquivo); sem ele, o nome fixo da
    cópia na pasta da extração. Só o nome, nunca um caminho."""
    original = r.get("arquivo_original")
    nome = Path(original.strip()).name.strip() if isinstance(original, str) else ""
    return r["arquivo"] if nome in ("", ".", "..") else nome


def _conferir_brutos_da_extracao(files: dict[str, Path]) -> None:
    """Antes de qualquer gravação: cada arquivo tem a assinatura do seu formato e CP/CR têm o cabeçalho
    do próprio relatório (os dois são .csv: trocados, o manifesto íntegro não basta). Os sanitizadores
    conferem tudo de novo depois; aqui só se barra cedo o que nem é o relatório certo."""
    try:
        for k in REPORTS:
            if not extracao_mod._assinatura_ok(k, files[k]):
                raise UsageError(f"--extracao: {k}: o arquivo {files[k].name} não parece um "
                                 f"{extracao_mod.EXTENSAO[k]} válido; nada foi gravado")
        esperados = load_headers()
        for k in ("cp", "cr"):
            _, rows = read_delimited(files[k])
            if not rows:
                raise ContractError(f"{k}: arquivo vazio")
            check_header(k, [normalize(h) for h in rows[0]], esperados[k])
    except ContractError as e:
        raise UsageError(f"--extracao: {e} (o arquivo não é o relatório certo?); nada foi gravado") from None
    except OSError as e:
        raise UsageError(f"--extracao: não foi possível ler os arquivos: {mensagem_segura(e)}") from None


def _entradas(a, competencia: str, env) -> dict:
    """Arquivos, modelo, saída e proveniência da rodada: do registro da extração (--extracao) ou das
    opções soltas. Qualquer inconsistência é UsageError (nada gravado)."""
    ent = {"tela": {}, "filtros": {}, "provas": {}, "nomes": {}, "hashes": None, "data_extracao": None,
           "periodo_fim": None, "origem": a.origem or "arquivos locais"}
    if a.extracao is not None:
        soltas = [n for n in (*REPORTS, "tela", "data_extracao", "origem") if getattr(a, n) is not None]
        if soltas:
            raise UsageError("--extracao não combina com: " + ", ".join("--" + n.replace("_", "-") for n in soltas))
        pasta = Path(a.extracao).resolve()
        try:
            m = carregar_manifesto(pasta)
        except ExtracaoErro as e:
            raise UsageError(f"--extracao: {e}; nada foi gravado") from None
        if m["produtora"] != a.produtora or m["competencia"] != competencia:
            raise UsageError("--extracao é de outra produtora/competência; nada foi gravado")
        if a.periodo_fim and _data_periodo_fim(a.periodo_fim) != dt.date.fromisoformat(m["periodo_fim"]):
            raise UsageError(f"--periodo-fim difere do período da extração ({m['periodo_fim']})")
        try:
            n_ext, _ = caminhos.nomes_pastas(ambiente.resolver(env))
        except caminhos.CaminhosErro as e:
            raise UsageError(str(e)) from None
        no_lugar = (pasta.parent.name == a.produtora and pasta.parent.parent.name == competencia
                    and pasta.parent.parent.parent.name == n_ext)
        if not no_lugar:
            _status("AVISO: a pasta da extração não está no lugar padrão "
                    f"({n_ext}/<competência>/<produtora>/<pasta>); prova_caminho fica vazio "
                    "no banco", file=sys.stderr)
        rel_pasta = caminhos.rel_extracao(competencia, a.produtora, pasta.name, nome=n_ext)
        rels = m["relatorios"]
        files = {k: pasta / rels[k]["arquivo"] for k in REPORTS}
        _conferir_brutos_da_extracao(files)
        ent.update(
            files=files,
            nomes={k: _nome_original(rels[k]) for k in REPORTS},
            hashes={k: rels[k]["sha256"] for k in REPORTS},
            tela={k: rels[k]["total_tela"] for k in REPORTS},
            filtros={k: rels[k]["filtros_tela"] for k in REPORTS},
            provas={k: f"{rel_pasta}/{rels[k]['prova']}" if no_lugar else None for k in REPORTS},
            data_extracao=m["data_extracao"], periodo_fim=m["periodo_fim"], origem=m["origem"])
    else:
        faltam = [f"--{k}" for k in REPORTS if not getattr(a, k)]
        if faltam:
            raise UsageError("faltam: " + ", ".join(faltam) + " (ou use --extracao)")
        files = {k: Path(getattr(a, k)) for k in REPORTS}
        ent.update(files=files, nomes={k: p.name for k, p in files.items()}, tela=parse_tela(a.tela),
                   data_extracao=a.data_extracao)
    # o ambiente só é lido quando falta --modelo ou --saida (testes e uso avulso passam os dois)
    amb = ambiente.resolver(env) if a.modelo is None or a.saida is None else None
    try:
        ent["modelo"] = Path(a.modelo) if a.modelo else caminhos.modelo_da_produtora(a.produtora, amb)
        ent["saida"] = Path(a.saida) if a.saida else caminhos.pasta_apuracao(
            caminhos.raiz_dados(amb), a.produtora, competencia, nome=caminhos.nomes_pastas(amb)[1])
    except caminhos.CaminhosErro as e:
        raise UsageError(str(e)) from None
    return ent


def _escrever_resultado(out_dir: Path, conf: dict) -> None:
    """RESULTADO.md só nas rodadas por --extracao. Nunca transforma uma rodada encerrada em falha."""
    try:
        (out_dir / ARQ_RESULTADO).write_text(render_resultado(conf), encoding="utf-8")
    except Exception as e:  # noqa: BLE001 - só o resumo; a rodada já terminou
        _status(f"AVISO: não foi possível gravar RESULTADO.md: {mensagem_segura(e)}", file=sys.stderr)


class _Restauracao:
    """O que --restaurar-banco precisa, conferido antes de tudo: o ref do projeto (derivado de
    SUPABASE_URL, nunca impresso) e quem cria o AdminSupabase (só se o banco estiver inacessível)."""

    def __init__(self, token: str, ref: str):
        self._token, self.ref = token, ref

    def criar(self):
        return AdminSupabase(self._token, self.ref)


def _fabrica_admin(env) -> _Restauracao:
    """Confere o que --restaurar-banco exige (token e SUPABASE_URL com um ref válido), lendo do MESMO
    lugar que `make_client` (`ambiente`). As mensagens citam só o nome das variáveis, nunca valores."""
    e = ambiente.resolver(env)
    faltam = [v for v in ("SUPABASE_ACCESS_TOKEN", "SUPABASE_URL") if not (e.get(v) or "").strip()]
    if faltam:
        raise UsageError(f"--restaurar-banco exige {' e '.join(faltam)} no .env; nada foi gravado")
    motivo = None
    try:
        ref = ref_do_projeto(e["SUPABASE_URL"].strip())
    except AdminErro as exc:                   # mensagem fixa, sem a URL
        motivo = str(exc)
    if motivo is not None:                     # fora do except: nada de __context__
        raise UsageError(f"--restaurar-banco: {motivo}")
    return _Restauracao(e["SUPABASE_ACCESS_TOKEN"].strip(), ref)


def _ref_do_cliente(client) -> str | None:
    """Ref do projeto para o qual o cliente do banco aponta (supabase-py guarda `supabase_url`)."""
    url = getattr(client, "supabase_url", None)
    if url is None:
        return None
    try:
        return ref_do_projeto(str(url))
    except AdminErro:
        return None


def _verificar(repo) -> BancoIndisponivel | None:
    try:
        repo.verificar_acesso()
        return None
    except BancoIndisponivel as e:
        return e


def _fechar(admin) -> None:
    fechar = getattr(admin, "fechar", None)
    if fechar is not None:
        try:
            fechar()
        except Exception:  # noqa: BLE001 - só libera o cliente http
            pass


def _msg_prazo() -> str:
    return (f"o banco não respondeu em {PRAZO_RESTAURACAO // 60} min; rode de novo (o projeto pode ainda "
            "estar subindo; nada foi gravado)")


def _garantir_banco(repo, admin=None, restauracao: _Restauracao | None = None) -> None:
    """Testa o banco. Se estiver inacessível e há `admin` (dublê) ou `restauracao` (só com
    --restaurar-banco), pergunta à API de gerenciamento: `acordar` só pede restauração se a API confirmar
    que o projeto está pausado, e no máximo uma vez (esta função o chama uma única vez por execução).
    Antes, confere que o cliente do banco e a restauração apontam para o MESMO projeto.
    Prazo GLOBAL de PRAZO_RESTAURACAO s desde o início da restauração: `acordar` espera no máximo
    ESPERA_ACORDAR s e as tentativas de falar com o banco depois dele param quando o prazo acaba.
    Se a API diz que o projeto já estava ativo, a falha não é pausa: testa só mais uma vez, sem esperar.
    Falha final -> BancoIndisponivel (mensagem sem URL, ref nem chave)."""
    erro = _verificar(repo)
    if erro is None:
        return
    if admin is None and restauracao is None:
        raise erro
    if restauracao is not None and _ref_do_cliente(repo.client) != restauracao.ref:
        raise BancoIndisponivel(f"{erro}; SUPABASE_URL do banco e da restauração não indicam o mesmo projeto "
                                "(ou não deu para identificá-lo); nada foi restaurado")
    _status("Banco inacessível; verificando o projeto no Supabase...", file=sys.stderr)
    inicio = _relogio()
    proprio = admin is None
    falha_admin = None
    resultado = None
    try:
        if proprio:
            admin = restauracao.criar()
        resultado = admin.acordar(espera_max=ESPERA_ACORDAR, intervalo=INTERVALO_ACORDAR)
    except AdminErro as e:
        falha_admin = mensagem_segura(e)
    finally:
        if proprio and admin is not None:
            _fechar(admin)
    if falha_admin is not None:                # fora do except: nada de __context__
        raise BancoIndisponivel(f"{erro}; {falha_admin}")
    if resultado == "ja_ativo":
        erro = _verificar(repo)
        if erro is None:
            return
        raise BancoIndisponivel(f"{erro}; o projeto já estava ativo no Supabase: a falha não é pausa e nada "
                                "foi restaurado (confira SUPABASE_URL, SUPABASE_SERVICE_KEY e a rede)")
    _status("Projeto ativo no Supabase; aguardando o banco responder...", file=sys.stderr)
    for tentativa in range(TENTATIVAS_APOS_ACORDAR):
        if tentativa and _relogio() - inicio >= PRAZO_RESTAURACAO:
            raise BancoIndisponivel(_msg_prazo())     # não começa outra tentativa depois do prazo
        if _verificar(repo) is None:
            return
        restante = PRAZO_RESTAURACAO - (_relogio() - inicio)
        if restante <= 0:
            raise BancoIndisponivel(_msg_prazo())
        if tentativa < TENTATIVAS_APOS_ACORDAR - 1:
            _dormir(min(INTERVALO_APOS_ACORDAR, restante))
    raise BancoIndisponivel(f"o banco não respondeu depois de {TENTATIVAS_APOS_ACORDAR} tentativas; rode de "
                            "novo (o projeto pode ainda estar subindo; nada foi gravado)")


def _consulta_agenda(produtora: str, competencia: str, periodo_fim: str, restaurar: bool, *, env=None,
                     client=None, admin=None) -> tuple[str, bool]:
    """Para o `agenda`: (estado, já importada). estado = consultado | nao_configurado | indisponivel.
    Com `restaurar`, acorda o projeto pausado (mesma regra do `importar`: só com o token no .env e com a API
    confirmando a pausa). Nunca levanta: sem banco, vale só a pasta local. Nenhum texto de exceção sai daqui."""
    try:
        repo = Repository(client if client is not None else make_client(env))
    except RuntimeError:                                  # SUPABASE_URL / SUPABASE_SERVICE_KEY ausentes
        return "nao_configurado", False
    except Exception:  # noqa: BLE001 - cliente com URL inválida etc.
        return "indisponivel", False
    restauracao = None
    if restaurar and admin is None:
        try:
            restauracao = _fabrica_admin(env)
        except UsageError:                                # sem o token: só consulta
            restauracao = None
    try:
        _garantir_banco(repo, admin if restaurar else None, restauracao)
        return "consultado", repo.vigente_completa(produtora, competencia, periodo_fim) is not None
    except Exception:  # noqa: BLE001 - BancoIndisponivel ou erro de consulta: a pasta local decide
        return "indisponivel", False


def importar(a, client=None, admin=None, env=None) -> int:
    # 1. argumentos (UsageError -> código 3, nada gravado)
    restauracao = None
    if not a.restaurar_banco:
        admin = None                           # só a opção explícita autoriza restaurar o projeto
    elif a.sem_banco:
        raise UsageError("--sem-banco não combina com --restaurar-banco")
    elif admin is None:
        restauracao = _fabrica_admin(env)      # sem chave: erro de uso ANTES de tudo
    if not produtoras.existe(a.produtora):
        raise UsageError(f"produtora desconhecida: {a.produtora!r} (cadastradas: {', '.join(produtoras.ids())})")
    prod = produtoras.carregar(a.produtora)
    competencia, inicio, fim = parse_competencia(a.competencia)
    ent = _entradas(a, competencia, env)
    periodo_fim = a.periodo_fim or ent["periodo_fim"]
    if periodo_fim:
        fim = _data_periodo_fim(periodo_fim)
        if fim < inicio:
            raise UsageError(f"--periodo-fim {fim.isoformat()} é anterior ao início do período "
                             f"({inicio.isoformat()})")
    extracao = _iso_com_fuso(ent["data_extracao"])
    tela, files, modelo_path, out_dir = ent["tela"], ent["files"], ent["modelo"], ent["saida"]
    # o Claude não abre o .env: é por aqui que ele fica sabendo onde ler RESULTADO.md e gravar PARADA.md
    _status(f"Pasta de apuracao: {out_dir}")
    nome = prod.nome
    faltando = [p.name for p in [*files.values(), modelo_path] if not p.is_file()]
    if faltando:
        raise UsageError(f"arquivo(s) não encontrado(s): {', '.join(faltando)}")
    destinos = _destinos(out_dir / _out_name(nome, competencia))
    modelo = modelo_path.resolve()
    if any(d.resolve() == modelo for d in destinos):
        raise UsageError(f"a saída sobrescreveria o modelo {modelo_path.name}; use outra pasta em --saida")
    if a.sem_banco and _saida_de_rodada_real(out_dir):
        raise UsageError(f"a pasta {out_dir.name} tem a saída de uma importação registrada no banco; --sem-banco não a "
                         "sobrescreve: use outra pasta em --saida")

    # 2. banco: acesso, dedupe por hash (nada é gravado até aqui)
    hashes = {k: sha256_file(p) for k, p in files.items()}
    if ent["hashes"] is not None and hashes != ent["hashes"]:
        raise UsageError("--extracao: um arquivo da extração mudou durante a conferência; nada foi gravado")
    repo = None
    anterior = None
    if not a.sem_banco:
        repo = Repository(client if client is not None else make_client(env))
        try:
            _garantir_banco(repo, admin, restauracao)  # plano gratuito: pausa após ~7 dias sem uso
        except BancoIndisponivel as e:         # a mensagem já vem sem URL/chave (mensagem_segura)
            print(f"ERRO: banco inacessível ({e}).\n{MSG_PAUSADO} "
                  "(painel do Supabase ou MCP restore_project). Nada foi gravado.", file=sys.stderr)
            return EXIT_FAILED
        repo.garantir_produtora(a.produtora, prod.nome, prod.localidade)
        same = None if a.forcar else repo.identica(a.produtora, competencia, hashes)
        if same:
            if same["vigente"]:
                _status(f"Arquivos idênticos já importados (importação {same['id']}, que é a foto vigente); "
                        "nada a fazer e nenhum arquivo foi gerado nesta pasta. Use --forcar para repetir.")
            else:
                _status(f"Arquivos idênticos já importados (importação {same['id']}); esta NÃO é a foto "
                        "vigente (a vigente é outra importação). Nada foi feito e nenhum arquivo foi gerado "
                        "nesta pasta; use --forcar para voltar a ela.")
            return EXIT_OK
        anterior = repo.vigente(a.produtora, competencia)    # só para avisar quem foi substituída
        try:                                   # limpeza de rodadas esquecidas: nunca derruba a rodada
            repo.encerrar_orfas(a.produtora, competencia, agora=_agora())
        except Exception as e:  # noqa: BLE001 - o cliente levanta tipos variados
            print(f"AVISO: não foi possível encerrar rodadas interrompidas: {mensagem_segura(e)}",
                  file=sys.stderr)

    out_dir.mkdir(parents=True, exist_ok=True)
    final_ok = out_dir / _out_name(nome, competencia)
    final = final_ok
    preparo = _preparo(final_ok)              # build/contagem/hash trabalham aqui, nunca no final
    imp_id = None
    conf = None
    substituido = False                        # o preparo já virou o arquivo final
    publicado = False                          # rodada encerrada com sucesso (promovida ou sem banco)
    promovida = False                          # depois de promovida, a foto nunca mais é marcada
    try:
        if repo:
            # criada antes de sanitizar: toda rodada (inclusive a bloqueada por contrato) fica registrada
            imp_id = repo.criar_importacao(
                produtora=a.produtora, competencia=competencia, periodo_inicio=inicio.isoformat(),
                periodo_fim=fim.isoformat(), data_extracao=extracao, origem=ent["origem"], gatilho=a.gatilho)
        meta = {"produtora": a.produtora, "competencia": competencia,
                "periodo": f"{inicio.isoformat()} a {fim.isoformat()}", "data_extracao": extracao,
                **({"importacao_id": imp_id} if imp_id else {"banco": False})}
        _apagar(preparo)                      # resto de uma rodada anterior interrompida

        # 3. sanitizar, carregar o Excel (no preparo) e conferir
        try:
            results = {
                "cp": sanitize_cp(files["cp"], regras=prod.regras),
                "cr": sanitize_cr(files["cr"]),
                "comissoes": sanitize_comissoes(files["comissoes"], regras=prod.regras),
                # o XLSX convertido fica em work/<stem>-<hash12>.xlsx (xlsx_cache_path) e é reaproveitado
                "rra": sanitize_rra(files["rra"], work_dir=out_dir / "work", regras=prod.regras),
            }
            build = build_workbook(modelo_path, preparo, results, master_formula_rows=prod.linha_mestre)
        except ContractError as e:             # mensagem escrita pelo próprio código, sem dados de linha
            _apagar(preparo)
            blocked_conf = _aviso_arquivo_antigo(
                _sem_banco_info(contract_blocked(str(e), meta=meta), repo), final_ok)
            if repo:
                repo.marcar(imp_id, "bloqueada", motivo=f"ContractError: {e}", conferencia=blocked_conf)
            _write_conference(out_dir, blocked_conf)
            publicado = True
            if a.extracao is not None:
                _escrever_resultado(out_dir, blocked_conf)
            _status(f"Status: bloqueada - erro de contrato: {e}", file=sys.stderr)
            return EXIT_BLOCKED

        na_count, na_error = (None, None) if a.sem_recalculo else _contar_nd(build, preparo)
        conf = _sem_banco_info(conference(results, build, screen_counts=tela, meta=meta, na_count=na_count,
                                          na_error=na_error), repo)
        blocked = conf["status"] == "bloqueada"
        if blocked:
            final = final_ok.with_name(final_ok.stem + "_BLOQUEADO.xlsx")
            _aviso_arquivo_antigo(conf, final_ok)

        # 4. banco: arquivos, linhas, exclusões e cargas (com o nome final e o hash do preparo)
        if repo:
            for k in REPORTS:
                repo.registrar_arquivo(imp_id, k, nome=ent["nomes"][k], sha256=hashes[k],
                                       tamanho=files[k].stat().st_size, total_tela=tela.get(k),
                                       filtros_tela=ent["filtros"].get(k), prova_caminho=ent["provas"].get(k))
                repo.inserir_linhas(imp_id, a.produtora, competencia, k, results[k])
                repo.inserir_exclusoes(imp_id, k, results[k], regras_texto(prod.regras))
            out_hash = sha256_file(preparo)
            for k in REPORTS:
                sr = build.sheets[DEST[k]["sheet"]]
                repo.registrar_carga(imp_id, aba=sr.sheet, primeira=sr.first_row, ultima=sr.last_row,
                                     arquivo=final.name, sha256=out_hash)

        # 5. finalizar os arquivos: só agora o .xlsx final e a conferência da pasta mudam
        _finalizar(preparo, final)
        substituido = True
        _write_conference(out_dir, conf)

        # 6. desfecho no banco (promover é sempre a última chamada)
        if repo:
            if blocked:
                repo.marcar(imp_id, "bloqueada", conferencia=conf)
            else:
                repo.salvar_conferencia(imp_id, conf)
                repo.promover(imp_id)
                promovida = True
        publicado = True
        if a.extracao is not None:             # depois de publicado: falhar aqui não vira 'falhou'
            _escrever_resultado(out_dir, conf)
        _status(f"Status: {conf['status']} - arquivo: {final.name}")
        if promovida and anterior and anterior != imp_id:
            _status(f"Esta importação ({imp_id}) substituiu a importação {anterior} como foto vigente.")
        return EXIT_BLOCKED if blocked else EXIT_OK
    except BaseException as e:                 # inclui KeyboardInterrupt: registra e relança
        motivo = mensagem_segura(e) if isinstance(e, Exception) else "interrompida"
        if not substituido:
            _apagar(preparo)                   # o final e a conferência anteriores ficam intactos
        elif not publicado:
            motivo = f"arquivo {final.name} gerado, mas a importação falhou: {motivo}"
        mudou = None
        if repo and imp_id and not promovida:
            mudou = _marcar_falha(repo, imp_id, motivo)
        # mudou is False: a importação já não estava em andamento (a promoção foi confirmada)
        if substituido and not publicado and mudou is not False and conf is not None:
            _regravar_falha(out_dir, conf, motivo)
        raise


def _tolerar_console() -> None:
    """Console do Windows sem UTF-8 (cp1252): caractere que não cabe vira '?' em vez de exceção."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:  # noqa: BLE001 - stream sem reconfigure (substituído por outro programa)
            pass


def main(argv=None, client=None, admin=None, env=None) -> int:
    _tolerar_console()
    try:
        args = _parser().parse_args(argv)
        if args.cmd == "importar":
            return importar(args, client=client, admin=admin, env=env)
        if args.cmd in cli_instalacao.COMANDOS:
            return cli_instalacao.executar(args, env, client=client)
        consulta = None
        if args.cmd == "agenda":
            consulta = lambda produtora, competencia, fim, restaurar: _consulta_agenda(  # noqa: E731
                produtora, competencia, fim, restaurar, env=env, client=client, admin=admin)
        return cli_extracao.executar(args, env, consulta_banco=consulta)
    except UsageError as e:
        print(f"ERRO DE USO: {e}", file=sys.stderr)
        return EXIT_FAILED
    except Exception as e:  # noqa: BLE001 - qualquer falha vira código 3 (a importação já foi marcada 'falhou')
        if os.environ.get("PAPAKI_DEBUG") == "1":
            traceback.print_exc()
        print(f"FALHA: {mensagem_segura(e)}", file=sys.stderr)
        return EXIT_FAILED
