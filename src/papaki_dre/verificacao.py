"""Conferências do comando `verificar`. Cada item é OK, AVISO ou FALHA, com uma frase fixa que nunca traz
valor do `.env`, caminho, e-mail nem chave (só nomes de variáveis e de arquivos). As partes que dependem da
máquina (Excel, Git, bibliotecas, sistema) entram por parâmetro para os testes não dependerem dela."""
from __future__ import annotations

import importlib
import shutil
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

from . import caminhos, produtoras

OK, AVISO, FALHA = "OK", "AVISO", "FALHA"
BIBLIOTECAS = (("openpyxl", "openpyxl"), ("lxml", "lxml"), ("win32com.client", "pywin32"),
               ("supabase", "supabase"), ("dotenv", "python-dotenv"), ("httpx", "httpx"))
TABELAS = ("produtoras", "importacoes", "arquivos_brutos", "rra_base", "cp_despesas", "cr_receita",
           "comissoes", "exclusoes", "cargas_excel")


@dataclass(frozen=True)
class Item:
    nivel: str
    nome: str
    detalhe: str


def _excel_registrado() -> bool:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
            return True
    except (ImportError, OSError):
        return False


def _git() -> bool:
    return shutil.which("git") is not None


def verificar(env, *, projeto=None, com_banco: bool = True, client=None, excel_registrado=None, importar=None,
              plataforma=None, versao=None, executavel=None, git=None) -> list[Item]:
    projeto = Path(projeto) if projeto else caminhos.RAIZ_PROJETO
    excel_registrado = excel_registrado or _excel_registrado
    importar = importar or importlib.import_module
    git = git or _git
    itens: list[Item] = []

    itens.append(Item(OK, "sistema", "Windows") if (plataforma or sys.platform) == "win32"
                 else Item(FALHA, "sistema", "o sistema usa o Excel do Windows"))
    v = tuple(versao or sys.version_info[:3])
    itens.append(Item(OK, "python", f"{v[0]}.{v[1]}.{v[2]}") if v >= (3, 11)
                 else Item(FALHA, "python", f"{v[0]}.{v[1]}.{v[2]}: é preciso Python 3.11 ou mais novo"))
    exe = Path(executavel or sys.executable).resolve()
    try:
        exe.relative_to((projeto / ".venv").resolve())
        itens.append(Item(OK, "venv", "usando o Python da .venv do projeto"))
    except ValueError:
        itens.append(Item(AVISO, "venv", "não é o Python da .venv do projeto: use sempre "
                                         ".venv/Scripts/python.exe -m papaki_dre ..."))
    faltam = []
    for modulo, pacote in BIBLIOTECAS:
        try:
            importar(modulo)
        except ImportError:
            faltam.append(pacote)
    itens.append(Item(OK, "bibliotecas", "todas instaladas") if not faltam
                 else Item(FALHA, "bibliotecas", "faltam: " + ", ".join(faltam)
                           + ' (rode: .venv/Scripts/python.exe -m pip install -e ".[dev]")'))
    itens.append(Item(OK, "excel", "Microsoft Excel registrado") if excel_registrado()
                 else Item(FALHA, "excel", "Microsoft Excel não encontrado (o programa converte o RRA e confere "
                                           "pelo Excel)"))
    itens.append(Item(OK, "git", "Git encontrado") if git()
                 else Item(AVISO, "git", "Git não encontrado (só é preciso para atualizar o sistema)"))

    _variaveis(env, com_banco, itens)
    _dados(env, projeto, itens)
    _modelos(env, itens)
    if com_banco:
        _banco(env, client, itens)
    return itens


def _tem(env, nome: str) -> bool:
    return bool((env.get(nome) or "").strip())


def _variaveis(env, com_banco: bool, itens: list[Item]) -> None:
    if com_banco:
        for nome in ("SUPABASE_URL", "SUPABASE_SERVICE_KEY"):
            itens.append(Item(OK, nome, "definida") if _tem(env, nome)
                         else Item(FALHA, nome, "falta no .env: cole o valor do projeto Supabase da Papaki"))
    itens.append(Item(OK, "SUPABASE_ACCESS_TOKEN", "definida") if _tem(env, "SUPABASE_ACCESS_TOKEN")
                 else Item(AVISO, "SUPABASE_ACCESS_TOKEN", "não definida: sem ela o banco pausado não acorda "
                                                           "sozinho"))
    itens.append(Item(OK, "PAPAKI_USUARIO_ESFERAS", "definida") if _tem(env, "PAPAKI_USUARIO_ESFERAS")
                 else Item(AVISO, "PAPAKI_USUARIO_ESFERAS", "não definida: o Claude precisa dela para escolher o "
                                                            "Chrome certo (configurar --usuario-esferas)"))
    itens.append(Item(OK, "PAPAKI_MAQUINA", "definida") if _tem(env, "PAPAKI_MAQUINA")
                 else Item(AVISO, "PAPAKI_MAQUINA", "não definida: sem o nome da máquina a rotina agendada não "
                                                    "reserva a produtora (configurar --maquina)"))


def _dados(env, projeto: Path, itens: list[Item]) -> None:
    try:
        n_ext, n_apu = caminhos.nomes_pastas(env)
    except caminhos.CaminhosErro as e:
        itens.append(Item(FALHA, "subpastas", str(e)))
        return
    if not _tem(env, "PAPAKI_DADOS_DIR"):
        itens.append(Item(FALHA, "dados", "PAPAKI_DADOS_DIR não definida (configurar --dados-dir ...)"))
        return
    try:
        raiz = caminhos.raiz_dados(env, projeto)
    except caminhos.CaminhosErro as e:
        itens.append(Item(FALHA, "dados", str(e)))
        return
    if not raiz.is_dir():
        itens.append(Item(FALHA, "dados", "a pasta de dados não existe (rode configurar --criar-pastas)"))
        return
    try:
        with tempfile.NamedTemporaryFile(dir=raiz):
            pass
    except OSError:
        itens.append(Item(FALHA, "dados", "não consigo gravar na pasta de dados (permissão?)"))
        return
    itens.append(Item(OK, "dados", "pasta de dados fora do projeto e gravável"))
    ausentes = [n for n in (n_ext, n_apu) if not (raiz / n).is_dir()]
    itens.append(Item(OK, "subpastas", f"{n_ext} e {n_apu} existem") if not ausentes
                 else Item(AVISO, "subpastas", "ainda não existem: " + ", ".join(ausentes)
                           + " (configurar --criar-pastas)"))


def _modelos(env, itens: list[Item]) -> None:
    com_modelo = 0
    for pid in produtoras.ids():
        var = f"{caminhos.PREFIXO_MODELO}{pid.upper()}"
        if not _tem(env, var):
            itens.append(Item(AVISO, f"modelo {pid}", f"sem {var} (sem problema se você não usa esta produtora)"))
            continue
        try:
            caminhos.modelo_da_produtora(pid, env)
        except caminhos.CaminhosErro as e:
            itens.append(Item(FALHA, f"modelo {pid}", str(e)))
            continue
        com_modelo += 1
        itens.append(Item(OK, f"modelo {pid}", "arquivo encontrado"))
    if not com_modelo:
        itens.append(Item(FALHA, "modelos", "nenhuma produtora tem Modelo DRE base configurado "
                                            "(configurar --modelo produtora=caminho)"))


def _tabela_inexistente(e: Exception) -> bool:
    """True só para o erro do PostgREST de tabela inexistente (42P01 / PGRST205 / mensagem equivalente)."""
    if str(getattr(e, "code", "") or "") in ("42P01", "PGRST205"):
        return True
    texto = str(e).lower()
    return "does not exist" in texto or "could not find the table" in texto


def _banco(env, client, itens: list[Item]) -> None:
    from .db import BancoIndisponivel, Repository, make_client, mensagem_segura

    try:
        cliente = client if client is not None else make_client(env)
    except RuntimeError as e:                  # a mensagem diz só QUAIS variáveis faltam
        itens.append(Item(FALHA, "banco", str(e)))
        return
    except Exception as e:  # noqa: BLE001 - o cliente levanta tipos variados
        itens.append(Item(FALHA, "banco", f"não foi possível criar o cliente: {mensagem_segura(e)}"))
        return
    try:
        Repository(cliente).verificar_acesso()
    except BancoIndisponivel as e:
        itens.append(Item(FALHA, "banco", f"banco inacessível ({e}); se o projeto está pausado, restaure-o no "
                                          "painel do Supabase"))
        return
    ausentes = []
    for tabela in TABELAS:
        try:
            cliente.table(tabela).select("*").limit(1).execute()
        except Exception as e:  # noqa: BLE001 - o erro do PostgREST/rede tem tipos variados
            if _tabela_inexistente(e):
                ausentes.append(tabela)
                continue
            # outro erro (rede, permissão...): não é falta de schema; nada do texto da exceção vai à saída
            itens.append(Item(FALHA, "banco", f"erro ao consultar as tabelas ({type(e).__name__})"))
            return
    itens.append(Item(OK, "banco", "responde e as 9 tabelas existem") if not ausentes
                 else Item(FALHA, "banco", "faltam tabelas: " + ", ".join(ausentes)
                           + " (aplique supabase/migrations/0001_schema.sql no projeto)"))
    if not ausentes:
        _reservas(cliente, itens)


def _reservas(cliente, itens: list[Item]) -> None:
    """Migração 0002 (várias máquinas): sem ela o sistema funciona, mas sem a proteção da reserva."""
    try:
        cliente.table("reservas_rodada").select("*").limit(1).execute()
    except Exception as e:  # noqa: BLE001 - o erro do PostgREST/rede tem tipos variados
        if _tabela_inexistente(e):
            itens.append(Item(AVISO, "reservas", "falta a tabela reservas_rodada (aplique "
                                                 "supabase/migrations/0002_reservas.sql): sem ela duas máquinas "
                                                 "podem extrair a mesma produtora ao mesmo tempo"))
        else:
            itens.append(Item(AVISO, "reservas", f"erro ao consultar as reservas ({type(e).__name__})"))
        return
    itens.append(Item(OK, "reservas", "tabela de reservas existe"))


def render(itens: list[Item]) -> str:
    linhas = [f"{i.nivel:<6} {i.nome}: {i.detalhe}" for i in itens]
    n = {nivel: sum(1 for i in itens if i.nivel == nivel) for nivel in (OK, AVISO, FALHA)}
    linhas.append(f"RESULTADO: {n[OK]} ok, {n[AVISO]} aviso(s), {n[FALHA]} falha(s)")
    return "\n".join(linhas) + "\n"
