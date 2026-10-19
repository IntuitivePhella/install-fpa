"""Onde ficam os arquivos. Uma só peça conhece a estrutura de pastas; o resto do código pede o caminho
a ela. A raiz vem de PAPAKI_DADOS_DIR (definida na instalação, fora do projeto) e o modelo base de cada
produtora de PAPAKI_MODELO_<PRODUTORA> (até o Plano 5 trazer um arquivo de configuração). Os nomes das
subpastas de extrações e apurações são os padrões do `config`, trocáveis na instalação
(PAPAKI_PASTA_EXTRACOES/PAPAKI_PASTA_APURACOES opcionais)."""
from __future__ import annotations

import re
from pathlib import Path

from .config import NOME_PASTA_APURACOES, NOME_PASTA_EXTRACOES

RAIZ_PROJETO = Path(__file__).resolve().parents[2]


class CaminhosErro(ValueError):
    """Configuração de pastas ausente ou inválida."""


def _normalizar(valor) -> str:
    """Tira espaços e aspas simples/duplas das pontas (o `.env` e o Windows costumam deixá-las)."""
    texto = str(valor or "").strip()
    if len(texto) >= 2 and texto[0] == texto[-1] and texto[0] in "'\"":
        texto = texto[1:-1].strip()
    return texto


def _caminho_completo(env, var: str) -> Path | None:
    valor = _normalizar((env or {}).get(var))
    if not valor:
        return None
    caminho = Path(valor).expanduser()
    if not caminho.is_absolute():
        raise CaminhosErro(f"{var} deve ser um caminho completo (ex.: C:\\Pasta\\Dados), não relativo")
    return caminho


def raiz_dados(env, projeto: Path | None = None) -> Path:
    caminho = _caminho_completo(env, "PAPAKI_DADOS_DIR")
    if caminho is None:
        raise CaminhosErro("defina PAPAKI_DADOS_DIR no .env (pasta de dados, fora do projeto; veja .env.example)")
    raiz = caminho.resolve()
    proj = (projeto or RAIZ_PROJETO).resolve()
    if raiz == proj or proj in raiz.parents or raiz in proj.parents:
        raise CaminhosErro("a pasta de dados não pode conter nem ficar dentro da pasta do projeto: "
                           "os dados reais nunca vão para o repositório")
    return raiz


_NOME_PASTA = re.compile(r"\w[\w .-]{0,60}")
_PROIBIDOS_WINDOWS = set('<>:"/\\|?*')


def _nome_de_pasta(env, var: str, padrao: str) -> str:
    valor = _normalizar((env or {}).get(var)) or padrao
    if (not _NOME_PASTA.fullmatch(valor) or valor.endswith((" ", ".")) or valor in (".", "..")
            or _PROIBIDOS_WINDOWS & set(valor)):
        raise CaminhosErro(f"{var}: nome de pasta inválido (use só o nome, sem barras, sem : * ? \" < > |, "
                           "sem ponto ou espaço no fim e com até 60 caracteres)")
    return valor


def nomes_pastas(env) -> tuple[str, str]:
    """(pasta das extrações, pasta das apurações) sob a raiz de dados: PAPAKI_PASTA_EXTRACOES e
    PAPAKI_PASTA_APURACOES do ambiente (a instalação escolhe), ou os padrões do `config`."""
    ext = _nome_de_pasta(env, "PAPAKI_PASTA_EXTRACOES", NOME_PASTA_EXTRACOES)
    apu = _nome_de_pasta(env, "PAPAKI_PASTA_APURACOES", NOME_PASTA_APURACOES)
    if ext.casefold() == apu.casefold():
        raise CaminhosErro("os nomes das pastas de extrações e de apurações devem ser diferentes")
    return ext, apu


def pasta_extracoes(raiz: Path, competencia: str, produtora: str, *, nome: str = NOME_PASTA_EXTRACOES) -> Path:
    return Path(raiz) / nome / competencia / produtora


def pasta_apuracao(raiz: Path, produtora: str, competencia: str, *, nome: str = NOME_PASTA_APURACOES) -> Path:
    return Path(raiz) / nome / produtora / competencia


def rel_extracao(competencia: str, produtora: str, nome_pasta: str, *, nome: str = NOME_PASTA_EXTRACOES) -> str:
    """Caminho da pasta de uma extração relativo à raiz de dados (o que vai para `prova_caminho`)."""
    return f"{nome}/{competencia}/{produtora}/{nome_pasta}"


PREFIXO_MODELO = "PAPAKI_MODELO_"


def modelo_da_produtora(produtora: str, env) -> Path:
    var = f"{PREFIXO_MODELO}{produtora.upper()}"
    caminho = _caminho_completo(env, var)
    if caminho is None:
        raise CaminhosErro(f"defina {var} no .env (caminho do Modelo DRE base anual da produtora)")
    if not caminho.is_file():
        raise CaminhosErro(f"o arquivo de {var} não existe")
    return caminho
