"""Variáveis de ambiente do sistema: as do Windows mais as do `.env` do projeto. O `.env` prevalece
(mesma regra de `db.make_client`: nunca mistura projetos) e valor vazio no `.env` conta como ausente."""
from __future__ import annotations

import os
import re
from pathlib import Path

from .caminhos import PREFIXO_MODELO, RAIZ_PROJETO


def ler_env() -> dict[str, str]:
    from dotenv import dotenv_values

    env = {chave: valor for chave, valor in os.environ.items() if valor.strip()}
    for chave, valor in dotenv_values(RAIZ_PROJETO / ".env").items():
        if valor and valor.strip():
            env[chave] = valor
    return env


def resolver(env):
    """`env` dado (testes) ou o ambiente real."""
    return env if env is not None else ler_env()


CHAVES_GRAVAVEIS = ("PAPAKI_DADOS_DIR", "PAPAKI_PASTA_EXTRACOES", "PAPAKI_PASTA_APURACOES",
                    "PAPAKI_USUARIO_ESFERAS", "PAPAKI_MAQUINA")
_MODELO = re.compile(r"[A-Z][A-Z0-9_]{1,30}")
CABECALHO_NOVO = (
    "# Cole abaixo, VOCÊ MESMO, os dados do projeto Supabase da Papaki (o Claude não deve abrir este arquivo).",
    "SUPABASE_URL=",
    "SUPABASE_SERVICE_KEY=",
    "# Opcional: chave pessoal do Supabase, só para acordar o projeto pausado.",
    "SUPABASE_ACCESS_TOKEN=",
)


def arquivo_env() -> Path:
    """O `.env` do projeto (resolvido na hora da chamada)."""
    return RAIZ_PROJETO / ".env"


def chave_gravavel(chave: str) -> bool:
    return chave in CHAVES_GRAVAVEIS or (
        chave.startswith(PREFIXO_MODELO) and bool(_MODELO.fullmatch(chave[len(PREFIXO_MODELO):])))


def ler_arquivo(arquivo) -> dict[str, str]:
    """Só as chaves com valor de um `.env` qualquer (não mistura com as variáveis do Windows)."""
    from dotenv import dotenv_values

    arquivo = Path(arquivo)
    if not arquivo.is_file():
        return {}
    return {k: v for k, v in dotenv_values(arquivo).items() if v and v.strip()}




def _valor_valido(valor) -> bool:
    """Valor que o `.env` guarda e o dotenv relê idêntico, em uma única linha física."""
    return (isinstance(valor, str) and bool(valor) and valor == valor.strip() and valor.isprintable()
            and len(valor.splitlines()) == 1 and "\t" not in valor and "${" not in valor and " #" not in valor
            and valor[0] not in "'\"" and valor[-1] not in "'\"")


def _nome_da_linha(core: str) -> tuple[str, str | None]:
    """(prefixo `export ` ou '', nome da chave) de uma linha de atribuição; nome None se não for atribuição."""
    s = core.lstrip()
    if s.startswith("#") or "=" not in s:
        return "", None
    esquerda = s.split("=", 1)[0]
    if esquerda.startswith("export ") or esquerda.startswith("export\t"):
        return "export ", esquerda[7:].strip()
    return "", esquerda.strip()


def gravar_env(arquivo, pares: dict[str, str]) -> list[str]:
    """Grava ou atualiza no `.env` SÓ as chaves permitidas (nunca segredos), preservando o resto do arquivo
    byte a byte (terminadores, BOM, comentários), sem aspas nem escape. Atualiza TODAS as ocorrências de cada
    chave (aceita `export `). Devolve os nomes gravados. Arquivo novo nasce com as linhas dos segredos vazias,
    para o usuário preencher. A gravação é atômica (arquivo temporário na mesma pasta + troca): o original
    nunca fica truncado. Nunca imprime valores nem caminhos."""
    arquivo = Path(arquivo)
    for chave, valor in pares.items():
        if not chave_gravavel(chave):
            raise ValueError(f"{chave}: esta chave não pode ser gravada por aqui (segredos são colados por "
                             "você direto no .env)")
        if not _valor_valido(valor):
            raise ValueError(f"{chave}: valor inválido (vazio, com espaço nas pontas, aspas nas pontas, tabulação, "
                             "caractere de controle ou quebra de linha, '${' ou ' #')")
    bom = ""
    if arquivo.exists():
        try:
            with open(arquivo, encoding="utf-8", newline="") as f:
                texto = f.read()
        except UnicodeDecodeError:
            raise ValueError("o .env não está em UTF-8: abra-o no Bloco de Notas e use Salvar como > UTF-8; "
                             "nada foi alterado") from None
        if texto.startswith("﻿"):
            bom, texto = "﻿", texto[1:]
        partes = texto.split("\n")                    # só \n separa linhas (nunca splitlines)
        if partes and partes[-1] == "":
            partes.pop()
        linhas = partes
        crlf = texto.count("\r\n")
        usa_crlf = crlf > texto.count("\n") - crlf
    else:
        linhas = [x for x in CABECALHO_NOVO]
        usa_crlf = False
    fim = "\r" if usa_crlf else ""
    vistas: set[str] = set()
    novas: list[str] = []
    for linha in linhas:
        cr = "\r" if linha.endswith("\r") else ""
        core = linha[:-1] if cr else linha
        prefixo, nome = _nome_da_linha(core)
        if nome is not None and nome in pares:
            novas.append(f"{prefixo}{nome}={pares[nome]}{cr}")
            vistas.add(nome)
        else:
            novas.append(linha)
    if novas and not novas[-1].endswith("\r"):
        novas[-1] += fim                              # última linha sem terminador ganha o do arquivo
    novas.extend(f"{k}={v}{fim}" for k, v in pares.items() if k not in vistas)
    try:
        dados = (bom + "\n".join(novas) + "\n").encode("utf-8")   # codifica ANTES de tocar no disco
    except UnicodeEncodeError:
        raise ValueError("valor inválido: caractere que não pode ser gravado em UTF-8") from None
    temporario = arquivo.with_name(arquivo.name + ".tmp")
    try:
        with open(temporario, "wb") as f:
            f.write(dados)
            f.flush()
            os.fsync(f.fileno())                          # o conteúdo está no disco antes da troca
        os.replace(temporario, arquivo)
    except BaseException as e:                            # inclui KeyboardInterrupt: nunca deixa .env.tmp
        try:
            temporario.unlink()
        except OSError:
            pass
        if isinstance(e, OSError):
            raise ValueError("não foi possível gravar o .env (arquivo em uso, pasta sem permissão ou inexistente); "
                             "nada foi alterado") from None
        raise
    return list(pares)
