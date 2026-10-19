"""Configuração por produtora: um arquivo JSON por produtora em `produtoras/` (na raiz do projeto,
versionado e igual para todos os usuários). Guarda só o que é regra de negócio da produtora; o que é da
máquina (pastas, caminho do modelo, chaves) fica no `.env` (veja `ambiente` e `caminhos`).

Este módulo não importa nenhum outro do pacote (evita ciclos: `config`, `sanitize`, `db` e `cli` o usam)."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

PASTA = Path(__file__).resolve().parents[2] / "produtoras"
_ID = re.compile(r"[a-z][a-z0-9_]{1,30}")
_COLUNA = re.compile(r"[A-Z]{1,3}")
_RELATORIOS = ("rra", "cp", "cr", "comissoes")


class ProdutoraErro(ValueError):
    """Arquivo de produtora ausente ou inválido. A mensagem cita o arquivo e o campo, nunca valores."""


@dataclass(frozen=True)
class Regras:
    """O que os sanitizadores precisam saber da produtora (spec do MVP, §6)."""
    localidade_cp: str                       # CP: só entra a linha com esta Localidade
    nucleo_rra: str                          # RRA: só entra o Núcleo igual a este
    modelo_comissoes: str                    # Comissões: só entra este Modelo de Orçamento
    comissoes_excluir_h: tuple[str, ...]     # Comissões: opções da coluna H que não entram


@dataclass(frozen=True)
class Produtora:
    id: str
    nome: str
    localidade: str                          # nome da Localidade no Esferas (tela do CR e das Comissões)
    esferas_consolidado: str                 # endereço do consolidado (CP, CR, RRA)
    esferas_produtora: str                   # endereço da produtora (Comissões)
    regras: Regras
    linha_mestre: dict                       # {"cp": {"AP": 881}}: linha do modelo com a fórmula completa


def _pasta(pasta) -> Path:
    return Path(pasta) if pasta is not None else PASTA


def existe(produtora_id, pasta=None) -> bool:
    return (isinstance(produtora_id, str) and bool(_ID.fullmatch(produtora_id))
            and (_pasta(pasta) / f"{produtora_id}.json").is_file())


def ids(pasta=None) -> list[str]:
    base = _pasta(pasta)
    return sorted(p.stem for p in base.glob("*.json") if _ID.fullmatch(p.stem)) if base.is_dir() else []


def _texto(d: dict, chave: str, onde: str) -> str:
    v = d.get(chave)
    if not isinstance(v, str) or not v.strip():
        raise ProdutoraErro(f"{onde}: o campo '{chave}' é obrigatório e deve ser um texto")
    return v.strip()


def _objeto(d, chaves: set[str], onde: str) -> dict:
    if not isinstance(d, dict):
        raise ProdutoraErro(f"{onde}: esperado um objeto")
    sobra = sorted(set(d) - chaves)
    if sobra:
        raise ProdutoraErro(f"{onde}: campo(s) desconhecido(s): {', '.join(sobra)}")
    return d


def _linha_mestre(d, onde: str) -> dict:
    if not isinstance(d, dict):
        raise ProdutoraErro(f"{onde}: 'linha_mestre' deve ser um objeto")
    saida = {}
    for rel, colunas in d.items():
        if rel not in _RELATORIOS or not isinstance(colunas, dict):
            raise ProdutoraErro(f"{onde}: 'linha_mestre' só aceita {', '.join(_RELATORIOS)} com colunas")
        saida[rel] = {}
        for col, linha in colunas.items():
            if (not isinstance(col, str) or not _COLUNA.fullmatch(col) or isinstance(linha, bool)
                    or not isinstance(linha, int) or linha < 1):
                raise ProdutoraErro(f"{onde}: 'linha_mestre' pede coluna em maiúsculas (ex.: AP) e linha inteira ≥ 1")
            saida[rel][col] = linha
    return saida


def carregar(produtora_id, pasta=None) -> Produtora:
    if not existe(produtora_id, pasta):
        raise ProdutoraErro(f"produtora desconhecida: {str(produtora_id)[:40]!r}")
    arq = _pasta(pasta) / f"{produtora_id}.json"
    onde = arq.name
    try:
        d = json.loads(arq.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        raise ProdutoraErro(f"{onde}: arquivo ilegível (JSON inválido)") from None
    _objeto(d, {"id", "nome", "localidade", "esferas", "regras", "linha_mestre"}, onde)
    if _texto(d, "id", onde) != produtora_id:
        raise ProdutoraErro(f"{onde}: o campo 'id' deve ser '{produtora_id}'")
    esf = _objeto(d.get("esferas"), {"consolidado", "produtora"}, f"{onde} (esferas)")
    reg = _objeto(d.get("regras"), {"nucleo_rra", "modelo_comissoes", "comissoes_excluir_h"}, f"{onde} (regras)")
    excluir = reg.get("comissoes_excluir_h")
    if not isinstance(excluir, list) or not all(isinstance(x, str) and x.strip() for x in excluir):
        raise ProdutoraErro(f"{onde} (regras): 'comissoes_excluir_h' deve ser uma lista de textos")
    localidade = _texto(d, "localidade", onde)
    return Produtora(
        id=produtora_id, nome=_texto(d, "nome", onde), localidade=localidade,
        esferas_consolidado=_texto(esf, "consolidado", f"{onde} (esferas)"),
        esferas_produtora=_texto(esf, "produtora", f"{onde} (esferas)"),
        regras=Regras(localidade_cp=localidade, nucleo_rra=_texto(reg, "nucleo_rra", f"{onde} (regras)"),
                      modelo_comissoes=_texto(reg, "modelo_comissoes", f"{onde} (regras)"),
                      comissoes_excluir_h=tuple(x.strip() for x in excluir)),
        linha_mestre=_linha_mestre(d.get("linha_mestre", {}), onde))


def regras_padrao() -> Regras:
    """Regras da Untitled (a única produtora do MVP), usadas quando quem chama não informa as da produtora:
    testes e uso avulso. O comando `importar` SEMPRE informa as da produtora pedida."""
    return carregar("untitled").regras
