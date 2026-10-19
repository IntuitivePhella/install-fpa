"""Auditoria do que vai para o GitHub (repositório PÚBLICO). Lê os arquivos que o git rastreia e falha se achar:
e-mail fora da lista permitida, CPF/CNPJ formatado fora de `tests/`, chave ou código de projeto do Supabase,
caminho `C:\\Users\\<nome>` e planilha, CSV, PDF ou `.env` rastreados.

Uso: .venv/Scripts/python.exe scripts/auditar_publicacao.py [pasta do repositório]
Nunca imprime o trecho achado (pode ser dado sensível): só arquivo, linha e tipo. Código de saída 0 = nada achado."""
from __future__ import annotations

import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

EXTENSOES_PROIBIDAS = {".xls", ".xlsx", ".xlsm", ".csv", ".pdf", ".db", ".sqlite"}
EMAILS_PERMITIDOS = {"notas@papaki.com.br", "noreply@anthropic.com", "golpe@falso.com"}
DOMINIOS_PERMITIDOS = {"example.com", "exemplo.com", "exemplo.com.br", "falso.com"}

_EMAIL = re.compile(r"[A-Za-z0-9._%+-]+@((?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,})")
_DOC = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b|\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b")
_CAMINHO = re.compile(r"[A-Za-z]:[\\/]+Users[\\/]+(?!<)[^\\/\s`'\"<>)]+")
_SUPABASE = (
    (re.compile(r"\b[a-z0-9]{20}\.supabase\.co"), "endereço de projeto Supabase"),
    (re.compile(r"project_ref=[a-z0-9]{20}"), "código de projeto Supabase"),
    (re.compile(r"\bref\b[^A-Za-z0-9\n]{0,8}[a-z]{20}\b"), "código de projeto Supabase"),
    (re.compile(r"eyJ[A-Za-z0-9_-]{20,}"), "chave JWT"),
    (re.compile(r"\bsb_(?:secret|publishable)_[A-Za-z0-9_-]{10,}"), "chave Supabase"),
    (re.compile(r"\bsbp_[0-9a-f]{20,}"), "token pessoal Supabase"),
)


@dataclass(frozen=True)
class Achado:
    arquivo: str
    linha: int          # 0 = o problema é o arquivo (nome ou extensão), não uma linha
    tipo: str


def _nome_proibido(arquivo: str) -> str | None:
    p = Path(arquivo)
    if p.suffix.lower() in EXTENSOES_PROIBIDAS:
        return "planilha, CSV ou PDF rastreado"
    if p.name == ".env":
        return ".env rastreado"
    if p.name.upper().startswith("HANDOFF"):
        return "HANDOFF rastreado (conversa interna)"
    return None


def _email_ruim(linha: str) -> bool:
    for m in _EMAIL.finditer(linha):
        email, dominio = m.group(0).lower(), m.group(1).lower()
        if email not in EMAILS_PERMITIDOS and dominio not in DOMINIOS_PERMITIDOS:
            return True
    return False


def auditar(raiz, arquivos) -> list[Achado]:
    raiz = Path(raiz)
    achados: list[Achado] = []
    for arquivo in arquivos:
        motivo = _nome_proibido(arquivo)
        if motivo:
            achados.append(Achado(arquivo, 0, motivo))
            continue
        caminho = raiz / arquivo
        if not caminho.is_file():
            continue
        try:
            texto = caminho.read_text(encoding="utf-8")
        except (UnicodeDecodeError, OSError):
            continue                                     # binário: nada a ler
        em_tests = Path(arquivo).parts[:1] == ("tests",)
        for n, linha in enumerate(texto.splitlines(), start=1):
            if _email_ruim(linha):
                achados.append(Achado(arquivo, n, "e-mail fora da lista permitida"))
            if not em_tests and _DOC.search(linha):
                achados.append(Achado(arquivo, n, "CPF ou CNPJ formatado"))
            if _CAMINHO.search(linha):
                achados.append(Achado(arquivo, n, "caminho de usuário do Windows"))
            for padrao, tipo in _SUPABASE:
                if padrao.search(linha):
                    achados.append(Achado(arquivo, n, tipo))
                    break
    return achados


def git_arquivos(raiz) -> list[str]:
    saida = subprocess.run(["git", "-C", str(raiz), "ls-files", "-z"], capture_output=True, check=True).stdout
    return [a for a in saida.decode("utf-8").split("\0") if a]


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    raiz = Path(argv[0]) if argv else Path(__file__).resolve().parents[1]
    arquivos = git_arquivos(raiz)
    achados = auditar(raiz, arquivos)
    for a in achados:
        print(f"{a.arquivo}:{a.linha}: {a.tipo}")
    if achados:
        print(f"AUDITORIA FALHOU: {len(achados)} achado(s). Corrija antes de publicar.")
        return 1
    print(f"Auditoria de publicação: nada encontrado em {len(arquivos)} arquivos.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
