"""Subcomandos de instalação: `configurar`, `verificar`, `produtora` e `instalacao`.
Nenhum imprime valor do `.env` nem caminho: só nomes de variáveis e respostas sim/não. A saída que o Claude
lê é JSON em ASCII."""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from . import ambiente, caminhos, produtoras

COMANDOS = ("configurar", "verificar", "produtora", "instalacao")
EXIT_OK, EXIT_FALHA = 0, 3
_EMAIL = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+")


def registrar(sub) -> None:
    c = sub.add_parser("configurar", help="grava no .env as respostas da instalação (nunca segredos)")
    c.add_argument("--dados-dir", type=Path, help="pasta de dados (completa, fora do projeto)")
    c.add_argument("--pasta-extracoes", help="nome da subpasta das extrações (padrão: extracoes)")
    c.add_argument("--pasta-apuracoes", help="nome da subpasta das apurações (padrão: apuracoes)")
    c.add_argument("--modelo", action="append", default=[], metavar="PRODUTORA=CAMINHO",
                   help="Modelo DRE base da produtora (caminho completo do .xlsx); pode repetir")
    c.add_argument("--usuario-esferas", help="e-mail do usuário do Esferas/Gmail desta máquina")
    c.add_argument("--criar-pastas", action="store_true", help="cria a pasta de dados e as duas subpastas")
    c.add_argument("--arquivo-env", type=Path, help=argparse.SUPPRESS)

    v = sub.add_parser("verificar", help="confere o ambiente (sem mostrar valores)")
    v.add_argument("--sem-banco", action="store_true", help="não consulta o banco central")

    p = sub.add_parser("produtora", help="configuração das produtoras (sem segredos)")
    ps = p.add_subparsers(dest="acao", required=True)
    ps.add_parser("listar", help="ids das produtoras cadastradas")
    m = ps.add_parser("mostrar", help="configuração de uma produtora")
    m.add_argument("--produtora", required=True)

    i = sub.add_parser("instalacao", help="o que o Claude precisa saber da instalação local")
    isub = i.add_subparsers(dest="acao", required=True)
    isub.add_parser("mostrar", help="usuário do Esferas, nomes das subpastas e se há modelo (sem caminhos)")


def _json(dados) -> None:
    print(json.dumps(dados, ensure_ascii=True))


def executar(a, env=None, client=None) -> int:
    try:
        if a.cmd == "configurar":
            return _configurar(a)
        if a.cmd == "verificar":
            return _verificar(a, env, client)
        if a.cmd == "produtora":
            return _produtora(a)
        return _instalacao(env)
    except ValueError as e:                   # inclui CaminhosErro e ProdutoraErro
        print(f"ERRO: {e}", file=sys.stderr)
        return EXIT_FALHA
    except OSError as e:                      # só o tipo e o nome do arquivo, nunca o caminho completo
        nome = Path(str(e.filename)).name if e.filename else "?"
        print(f"ERRO: falha de arquivo ({type(e).__name__}) em {nome!r}", file=sys.stderr)
        return EXIT_FALHA


def _configurar(a) -> int:
    pares: dict[str, str] = {}
    if a.dados_dir is not None:
        pares["PAPAKI_DADOS_DIR"] = str(caminhos.raiz_dados({"PAPAKI_DADOS_DIR": str(a.dados_dir)}))
    if a.pasta_extracoes:
        pares["PAPAKI_PASTA_EXTRACOES"] = a.pasta_extracoes
    if a.pasta_apuracoes:
        pares["PAPAKI_PASTA_APURACOES"] = a.pasta_apuracoes
    if pares.keys() & {"PAPAKI_PASTA_EXTRACOES", "PAPAKI_PASTA_APURACOES"}:
        caminhos.nomes_pastas(pares)          # valida os nomes antes de gravar qualquer coisa
    for item in a.modelo:
        produtora, _, caminho = item.partition("=")
        produtora = produtora.strip()
        if not produtoras.existe(produtora):
            raise ValueError(f"--modelo: produtora desconhecida (cadastradas: {', '.join(produtoras.ids())})")
        arq = Path(caminho.strip().strip('"'))
        if not arq.is_absolute():
            raise ValueError("--modelo: use o caminho completo do arquivo")
        if arq.suffix.lower() != ".xlsx" or not arq.is_file():
            raise ValueError(f"--modelo {produtora}: arquivo .xlsx não encontrado")
        pares[f"{caminhos.PREFIXO_MODELO}{produtora.upper()}"] = str(arq)
    if a.usuario_esferas:
        if not _EMAIL.fullmatch(a.usuario_esferas.strip()):
            raise ValueError("--usuario-esferas: informe um e-mail")
        pares["PAPAKI_USUARIO_ESFERAS"] = a.usuario_esferas.strip()
    if not pares and not a.criar_pastas:
        raise ValueError("nada a configurar: informe ao menos uma opção")
    arquivo = a.arquivo_env or ambiente.arquivo_env()
    existia = arquivo.exists()
    gravadas = ambiente.gravar_env(arquivo, pares) if pares else []
    criadas: list[str] = []
    if a.criar_pastas:
        atual = ambiente.ler_arquivo(arquivo)
        raiz = caminhos.raiz_dados(atual)
        for nome in caminhos.nomes_pastas(atual):
            (raiz / nome).mkdir(parents=True, exist_ok=True)
            criadas.append(nome)
    _json({"arquivo_env": "atualizado" if existia else "criado", "gravadas": gravadas, "pastas_criadas": criadas})
    return EXIT_OK


def _produtora(a) -> int:
    if a.acao == "listar":
        _json({"produtoras": produtoras.ids()})
        return EXIT_OK
    p = produtoras.carregar(a.produtora)
    _json({"id": p.id, "nome": p.nome, "localidade": p.localidade,
           "esferas": {"consolidado": p.esferas_consolidado, "produtora": p.esferas_produtora},
           "regras": {"nucleo_rra": p.regras.nucleo_rra, "modelo_comissoes": p.regras.modelo_comissoes,
                      "comissoes_excluir_h": list(p.regras.comissoes_excluir_h)},
           "linha_mestre": p.linha_mestre})
    return EXIT_OK


def _instalacao(env) -> int:
    amb = ambiente.resolver(env)
    n_ext, n_apu = caminhos.nomes_pastas(amb)
    modelos = {}
    for pid in produtoras.ids():
        try:
            caminhos.modelo_da_produtora(pid, amb)
            modelos[pid] = True
        except caminhos.CaminhosErro:
            modelos[pid] = False
    _json({"dados_configurado": bool((amb.get("PAPAKI_DADOS_DIR") or "").strip()),
           "pasta_extracoes": n_ext, "pasta_apuracoes": n_apu,
           "usuario_esferas": (amb.get("PAPAKI_USUARIO_ESFERAS") or "").strip() or None, "modelos": modelos})
    return EXIT_OK


def _verificar(a, env, client) -> int:
    from . import verificacao
    itens = verificacao.verificar(ambiente.resolver(env), com_banco=not a.sem_banco, client=client)
    print(verificacao.render(itens), end="")
    return EXIT_FALHA if any(i.nivel == verificacao.FALHA for i in itens) else EXIT_OK
