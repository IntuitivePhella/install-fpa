"""Registro de cada extração do Esferas: o que a tela mostrou, o disparo único de cada relatório e os
arquivos baixados. O Claude chama estes passos pelo CLI (`extracao ...`); as travas valem AQUI, não só
no roteiro. Nada neste módulo lê o conteúdo dos relatórios (só assinatura de formato, tamanho e hash).

Regras que este módulo impõe:
- um disparo por relatório; o horário vai para o disco ANTES do clique; disparado sem arquivo nunca é
  disparado de novo (nem numa retomada, nem numa extração "nova");
- a tela (Localidade, datas, Status, critério, "Pesquisar em") é validada antes de qualquer disparo;
- arquivo só entra se o relatório foi disparado, com extensão e assinatura certas, não vazio, não anterior
  ao disparo, e é copiado com hash e tamanho conferidos; o RRA só depois de escolhida a mensagem do e-mail;
- o manifesto só sai com os 4 relatórios completos e `carregar_manifesto` recusa arquivo alterado ou sumido;
- `estado.json` é gravado de forma atômica (temporário + fsync + replace), sob a trava da pasta-base,
  conferindo logo antes da troca que a trava ainda é do processo e que o arquivo não mudou."""
from __future__ import annotations

import contextlib
import copy
import datetime as dt
import json
import os
import re
import shutil
import time
import unicodedata
import uuid
import zipfile
from pathlib import Path

from . import produtoras, rra_email
from .config import FUSO_BRASILIA
from .db import sha256_file

RELATORIOS = ("cp", "cr", "rra", "comissoes")
ARQ_ESTADO = "estado.json"
ARQ_MANIFESTO = "extracao.json"
ARQ_TRAVA = ".trava"
VERSAO = 1
ORIGEM = "Esferas (Chrome)"
CRITERIO = {"cp": "Data competência", "cr": "Data competência", "rra": "Aprovação", "comissoes": "Aprovação"}
EXTENSAO = {"cp": ".csv", "cr": ".csv", "rra": ".xls", "comissoes": ".xlsx"}
TODAS = ("todas", "todas as localidades")
PESQUISAR_EM_CP = "Individuais e agrupadas"
OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
BOM = b"\xef\xbb\xbf"
PRINT_MAGICA = {".png": b"\x89PNG\r\n\x1a\n", ".jpg": b"\xff\xd8\xff", ".jpeg": b"\xff\xd8\xff"}
TOLERANCIA_MTIME_S = 5
TELA_VALIDADE_S = 15 * 60          # tela vista há mais tempo que isso não libera disparo: confira de novo
TRAVA_ESPERA_S = 10
TRAVA_ABANDONO_S = 10 * 60         # trava mais velha que isso é tomada mesmo com o pid aparentemente vivo
ECO_MAX = 40                       # quanto de um valor digitado a mensagem de erro repete

_RE_COMPETENCIA = re.compile(r"\d{4}-(0[1-9]|1[0-2])")
_RE_DATA_ISO = re.compile(r"\d{4}-\d{2}-\d{2}")
_RE_PASTA = re.compile(r"(\d{8}-\d{4})(?:-(\d+))?")
_RE_SHA = re.compile(r"[0-9a-f]{64}")
_RE_PROVA = re.compile(r"prints/(cp|cr|rra|comissoes)\.(png|jpg)")


class ExtracaoErro(ValueError):
    """Trava violada ou uso fora de ordem. A mensagem não leva dados dos relatórios."""


# --- utilitários ---------------------------------------------------------------------------------

def _relogio_padrao() -> dt.datetime:
    return dt.datetime.now(FUSO_BRASILIA)


def _momento(relogio) -> dt.datetime:
    """Lê o relógio (injetável nos testes) e devolve o instante em Brasília, sem frações de segundo.
    Relógio sem fuso é recusado: comparar horários sem fuso daria resultado errado."""
    m = relogio()
    if not isinstance(m, dt.datetime) or m.utcoffset() is None:
        raise ExtracaoErro("o relógio precisa devolver data e hora com fuso")
    return m.astimezone(FUSO_BRASILIA).replace(microsecond=0)


def _instante(texto) -> dt.datetime | None:
    """ISO-8601 com fuso -> datetime; qualquer outra coisa -> None."""
    if not isinstance(texto, str):
        return None
    try:
        m = dt.datetime.fromisoformat(texto)
    except ValueError:
        return None
    return m if m.utcoffset() is not None else None


def _texto(x) -> str | None:
    """Mesma forma Unicode (NFC) e espaços colapsados. NÃO tira acento: 'PRODUCAO' continua diferente
    de 'PRODUÇÃO'."""
    if not isinstance(x, str):
        return None
    return " ".join(unicodedata.normalize("NFC", x).split())


def _norm(x) -> str | None:
    t = _texto(x)
    return None if t is None else t.casefold()


def _eco(x) -> str:
    """Repete só o começo do que foi digitado (pode ser lixo copiado da tela com dados de cliente)."""
    s = repr(x)
    return s if len(s) <= ECO_MAX else s[:ECO_MAX - 4] + "...'"


def _data(texto) -> dt.date | None:
    if not isinstance(texto, str):
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(texto.strip(), fmt).date()
        except ValueError:
            pass
    return None


def _data_iso(texto) -> dt.date | None:
    if not isinstance(texto, str) or not _RE_DATA_ISO.fullmatch(texto):
        return None
    try:
        return dt.date.fromisoformat(texto)
    except ValueError:
        return None


def _insistir(fn, tentativas=200):
    """No Windows, apagar, ler ou substituir um arquivo que outro processo está abrindo dá PermissionError
    por instantes: tenta de novo (até ~2 s)."""
    for _ in range(tentativas - 1):
        try:
            return fn()
        except PermissionError:
            time.sleep(0.01)
    return fn()


def _apagar(caminho: Path) -> None:
    with contextlib.suppress(FileNotFoundError):
        _insistir(caminho.unlink)


def _gravar_json(caminho: Path, dados: dict, conferir=None) -> None:
    """Atômico: escreve um temporário, força para o disco e troca de uma vez. Se algo falha, o arquivo
    antigo fica intacto e o temporário é apagado. `conferir()` roda imediatamente antes da troca e pode
    abortar (ExtracaoErro) sem gravar nada."""
    tmp = caminho.with_name(caminho.name + ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        if conferir is not None:
            conferir()
        _insistir(lambda: os.replace(tmp, caminho))
    except BaseException:
        with contextlib.suppress(OSError):
            tmp.unlink()
        raise


# --- trava ---------------------------------------------------------------------------------------
# Uma trava por pasta-base (<extracoes>/<competencia>/<produtora>): iniciar e todas as mudanças de estado
# das extrações dessa base passam por ela, então a varredura do iniciar nunca perde um disparo em curso.
# A trava guarda {pid, id, criada}. Ela é ÓRFÃ (tomada sozinha, sem pessoa) quando:
#   - o pid gravado não existe mais (processo morto) — tomada imediata; ou
#   - tem mais de TRAVA_ABANDONO_S (10 min), qualquer que seja o pid (pid reutilizado, conteúdo vazio).
# Um dono vivo e lento nunca perde a trava antes de 10 min; e, se perder, `_mudar` confere imediatamente
# antes de gravar que a trava ainda é dele e que o estado.json não mudou — senão aborta sem gravar.

def _processo_vivo(pid) -> bool:
    """False só quando há certeza de que o processo não existe. Qualquer dúvida -> True (vivo)."""
    if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        return True
    if pid == os.getpid():
        return True
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes
            k32 = ctypes.WinDLL("kernel32", use_last_error=True)
            k32.OpenProcess.restype = wintypes.HANDLE
            k32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
            k32.GetExitCodeProcess.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD))
            k32.CloseHandle.argtypes = (wintypes.HANDLE,)
            h = k32.OpenProcess(0x1000, False, pid)          # PROCESS_QUERY_LIMITED_INFORMATION
            if not h:
                return ctypes.get_last_error() != 87          # 87 = parâmetro inválido: pid não existe
            try:
                codigo = wintypes.DWORD()
                if not k32.GetExitCodeProcess(h, ctypes.byref(codigo)):
                    return True
                return codigo.value == 259                    # STILL_ACTIVE
            finally:
                k32.CloseHandle(h)
        except Exception:                                     # noqa: BLE001 - na dúvida, vivo
            return True
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except OSError:
        return True
    return True


def _criar_exclusivo(caminho: Path, dados: bytes) -> bool:
    """Cria `caminho` só se não existe. True = criou; False = já existe (ou ainda está sendo apagado).
    PermissionError persistente sem o arquivo existir = pasta sem permissão -> ExtracaoErro."""
    for _ in range(50):
        try:
            fd = os.open(caminho, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            return False
        except PermissionError:
            if caminho.exists():
                return False
            time.sleep(0.01)                                  # Windows: arquivo sendo apagado
            continue
        try:
            os.write(fd, dados)
        except BaseException:
            os.close(fd)
            with contextlib.suppress(OSError):
                _apagar(caminho)
            raise
        os.close(fd)
        return True
    raise ExtracaoErro(f"sem permissão para criar {caminho.name} na pasta {caminho.parent.name} "
                       "(pasta somente leitura?); pare e avise o usuário")


def _ler_trava(caminho: Path) -> dict | None:
    """Conteúdo da trava ({pid, id, criada}); None se sumiu, está vazia ou ilegível."""
    try:
        dados = json.loads(_insistir(lambda: caminho.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        return None
    return dados if isinstance(dados, dict) else None


def _idade_trava(caminho: Path, conteudo: dict | None) -> float | None:
    """Segundos desde a criação (pelo horário gravado; sem ele, pela data do arquivo). None = sumiu."""
    criada = (conteudo or {}).get("criada")
    if isinstance(criada, (int, float)) and not isinstance(criada, bool):
        return time.time() - criada
    try:
        return time.time() - caminho.stat().st_mtime
    except OSError:
        return None


def _orfa(caminho: Path, conteudo: dict | None) -> bool | None:
    """True = órfã; False = de dono (possivelmente) vivo; None = sumiu."""
    idade = _idade_trava(caminho, conteudo)
    if idade is None:
        return None
    if idade > TRAVA_ABANDONO_S:
        return True
    return conteudo is not None and not _processo_vivo(conteudo.get("pid"))


def _tomar_trava_orfa(caminho: Path) -> bool:
    """Remove a trava se ela é órfã. True = vale tentar criar de novo já.
    Só quem cria `.trava.tomada` (criação exclusiva) pode remover, e só depois de reconferir que a trava
    ainda é a mesma e ainda é órfã: duas tomadas simultâneas não removem a trava nova uma da outra."""
    conteudo = _ler_trava(caminho)
    orfa = _orfa(caminho, conteudo)
    if orfa is None:
        return True                                  # foi liberada nesse meio-tempo
    if not orfa:
        return False
    tomada = caminho.with_name(caminho.name + ".tomada")
    if not _criar_exclusivo(tomada, json.dumps({"pid": os.getpid(), "criada": time.time()}).encode("utf-8")):
        if _orfa(tomada, _ler_trava(tomada)):        # uma tomada interrompida também sai
            _apagar(tomada)
        return False
    try:
        atual = _ler_trava(caminho)
        if atual == conteudo and _orfa(caminho, atual):
            _apagar(caminho)
    finally:
        with contextlib.suppress(OSError):
            _apagar(tomada)
    return True


@contextlib.contextmanager
def _trava(base: Path):
    """Uma operação por vez na pasta-base. Devolve o dono ({pid, id, criada}) para a conferência final."""
    caminho = base / ARQ_TRAVA
    dono = {"pid": os.getpid(), "id": uuid.uuid4().hex, "criada": time.time()}
    limite = time.monotonic() + TRAVA_ESPERA_S
    while not _criar_exclusivo(caminho, json.dumps(dono).encode("utf-8")):
        if _tomar_trava_orfa(caminho) and time.monotonic() < limite:
            continue
        if time.monotonic() >= limite:
            raise ExtracaoErro(f"outra operação nesta extração está em andamento ({ARQ_TRAVA} de processo vivo); "
                               "tente de novo em instantes. A trava de um processo que morreu é tomada sozinha; "
                               "a de um processo ainda vivo (ou de pid reutilizado por outro programa) só é "
                               f"liberada sozinha depois de {TRAVA_ABANDONO_S // 60} min")
        time.sleep(0.05)
    try:
        yield dono
    finally:
        if (_ler_trava(caminho) or {}).get("id") == dono["id"]:     # nunca apaga a trava de outro dono
            with contextlib.suppress(OSError):
                _apagar(caminho)


def _conferir_posse(base: Path, dono: dict, arquivo: Path | None = None, bruto: bytes | None = None) -> None:
    """Imediatamente antes de gravar: a trava ainda é minha e o arquivo não mudou desde a leitura."""
    if (_ler_trava(base / ARQ_TRAVA) or {}).get("id") != dono["id"]:
        raise ExtracaoErro("a trava foi tomada por outro processo durante a operação; nada foi gravado. "
                           "Confira o estado (extracao estado) antes de repetir")
    if arquivo is not None:
        try:
            atual = _insistir(arquivo.read_bytes)
        except FileNotFoundError:
            atual = None
        if atual != bruto:
            raise ExtracaoErro(f"{arquivo.name} mudou durante a operação; nada foi gravado. "
                               "Confira o estado (extracao estado) antes de repetir")


# --- travas da tela ------------------------------------------------------------------------------

def localidade_esperada(rel: str, produtora: str) -> str | None:
    """None = 'Todas'. CR e Comissões são exatamente a Localidade da produtora (nunca 'UTLD PRODUCAO')."""
    return None if rel in ("cp", "rra") else produtoras.carregar(produtora).localidade


def validar_tela(rel: str, estado: dict, tela: dict) -> list[str]:
    erros = []
    esperado = localidade_esperada(rel, estado["produtora"])
    visto = tela.get("localidade")
    if esperado is None:
        if _norm(visto) not in TODAS:
            erros.append(f"Localidade deve ser 'Todas'; a tela mostra {_eco(visto)}")
    elif _texto(visto) != _texto(esperado):       # exatamente: com acento/cedilha e maiúsculas
        erros.append(f"Localidade deve ser exatamente {esperado!r}; a tela mostra {_eco(visto)}")
    if _norm(tela.get("criterio")) != _norm(CRITERIO[rel]):
        erros.append(f"Critério de data deve ser {CRITERIO[rel]!r}; a tela mostra {_eco(tela.get('criterio'))}")
    if _data(tela.get("data_inicio")) != dt.date.fromisoformat(estado["periodo_inicio"]):
        erros.append(f"Data inicial deve ser {estado['periodo_inicio']}; a tela mostra {_eco(tela.get('data_inicio'))}")
    if _data(tela.get("data_fim")) != dt.date.fromisoformat(estado["periodo_fim"]):
        erros.append(f"Data final deve ser {estado['periodo_fim']}; a tela mostra {_eco(tela.get('data_fim'))}")
    if _norm(tela.get("status")) != "todos":
        erros.append(f"Status/aba deve ser 'Todos'; a tela mostra {_eco(tela.get('status'))}")
    if rel == "cp" and _norm(tela.get("pesquisar_em")) != _norm(PESQUISAR_EM_CP):
        erros.append(f"Pesquisar em deve ser {PESQUISAR_EM_CP!r}; a tela mostra {_eco(tela.get('pesquisar_em'))}")
    total = tela.get("total")
    if isinstance(total, bool) or not isinstance(total, int) or total < 0:
        erros.append(f"Total de registros deve ser um inteiro >= 0; recebido {_eco(total)}")
    return erros


def _problema_print(caminho: Path) -> str | None:
    if not caminho.is_file():
        return "print da tela não encontrado"
    magica = PRINT_MAGICA.get(caminho.suffix.casefold())
    if magica is None:
        return "o print deve ser .png ou .jpg"
    with open(caminho, "rb") as f:
        if not f.read(len(magica)) == magica:
            return "o print não parece uma imagem .png/.jpg"
    return None


def _assinatura_ok(rel: str, caminho: Path) -> bool:
    with open(caminho, "rb") as f:
        cab = f.read(4096)
    if rel == "rra":
        return cab.startswith(OLE2)
    if rel == "comissoes":
        if not cab.startswith(b"PK\x03\x04"):
            return False
        try:
            with zipfile.ZipFile(caminho) as z:
                nomes = set(z.namelist())
        except (zipfile.BadZipFile, OSError, ValueError):
            return False
        return "[Content_Types].xml" in nomes and "xl/workbook.xml" in nomes
    corpo = cab[len(BOM):] if cab.startswith(BOM) else cab     # CSV (UTF-8 com BOM ou Windows-1252)
    if b"\x00" in corpo or b";" not in corpo:
        return False
    return not corpo.lstrip().startswith(b"<")                  # página HTML de erro salva como .csv


# --- validação do estado.json --------------------------------------------------------------------

def _problema_no_periodo(d: dict) -> str | None:
    """Produtora conhecida, competência AAAA-MM, início = 01/01 do ano, fim dentro da competência.
    Vale para o estado.json e para o manifesto."""
    if not produtoras.existe(d.get("produtora")):
        return "produtora desconhecida"
    comp = d.get("competencia")
    if not isinstance(comp, str) or not _RE_COMPETENCIA.fullmatch(comp):
        return "competência inválida"
    if d.get("periodo_inicio") != f"{comp[:4]}-01-01":
        return "período inicial inválido"
    fim = _data_iso(d.get("periodo_fim"))
    if fim is None or fim.strftime("%Y-%m") != comp:
        return "período final inválido"
    return None


def _problema_na_tela_gravada(rel: str, d: dict, tela) -> str | None:
    """A tela gravada (estado ou manifesto, possivelmente editada à mão) passa de novo pelas travas."""
    if not isinstance(tela, dict):
        return f"{rel}: tela inválida"
    erros = validar_tela(rel, d, tela)
    return f"{rel}: tela gravada não passa nas travas - " + "; ".join(erros) if erros else None


def _prova_ok(rel: str, prova) -> bool:
    return isinstance(prova, str) and bool(_RE_PROVA.fullmatch(prova)) and prova.startswith(f"prints/{rel}.")


def _problema_no_estado(e) -> str | None:
    if not isinstance(e, dict):
        return "não é um objeto"
    if e.get("versao") != VERSAO:
        return "versão inesperada"
    problema = _problema_no_periodo(e)
    if problema:
        return problema
    ids = e.get("rra_ids_antes")
    if ids is not None and not (isinstance(ids, list) and all(isinstance(i, str) for i in ids)):
        return "rra_ids_antes inválido"
    msg = e.get("rra_mensagem")
    if msg is not None and not (isinstance(msg, str) and msg):
        return "rra_mensagem inválida"
    rels = e.get("relatorios")
    if not isinstance(rels, dict):
        return "relatorios inválido"
    for rel, r in rels.items():
        if rel not in RELATORIOS or not isinstance(r, dict):
            return f"relatório inesperado: {_eco(rel)}"
        if "tela" in r:
            t = r["tela"]
            if not isinstance(t, dict) or _instante(t.get("vista_em")) is None or not _prova_ok(rel, r.get("prova")):
                return f"{rel}: tela inválida"
            problema = _problema_na_tela_gravada(rel, e, t)
            if problema:
                return problema
        if "disparo" in r and (_instante(r["disparo"]) is None or "tela" not in r):
            return f"{rel}: disparo inválido"
        if "arquivo" in r:
            a = r["arquivo"]
            if ("disparo" not in r or not isinstance(a, dict) or a.get("nome") != f"{rel}{EXTENSAO[rel]}"
                    or not isinstance(a.get("sha256"), str) or not _RE_SHA.fullmatch(a["sha256"])
                    or isinstance(a.get("tamanho"), bool) or not isinstance(a.get("tamanho"), int)
                    or a["tamanho"] <= 0):
                return f"{rel}: arquivo inválido"
    if rels.get("rra", {}).get("arquivo") and not msg:
        return "rra: arquivo sem mensagem escolhida"
    if "abandonada_em" in e:
        if _instante(e["abandonada_em"]) is None or not isinstance(e.get("abandonada_por"), str) \
                or not e["abandonada_por"]:
            return "marca de abandono inválida"
        if _pendentes(e):
            return "abandonada com relatório disparado sem arquivo"
    return None


def _ler_estado_bruto(pasta: Path) -> tuple[dict, bytes]:
    """Estado validado e os bytes lidos (para conferir, antes de gravar, que ninguém mudou o arquivo)."""
    try:
        bruto = _insistir((pasta / ARQ_ESTADO).read_bytes)
        estado = json.loads(bruto.decode("utf-8"))
    except (OSError, ValueError):
        raise ExtracaoErro(f"pasta de extração inválida ({ARQ_ESTADO} ausente ou ilegível): {pasta.name}") from None
    problema = _problema_no_estado(estado)
    if problema:
        raise ExtracaoErro(f"pasta de extração inválida ({ARQ_ESTADO} ilegível: {problema}): {pasta.name}")
    return estado, bruto


def _ler_estado(pasta: Path) -> dict:
    return _ler_estado_bruto(pasta)[0]


def _pendentes(estado: dict) -> list[str]:
    """Relatórios disparados e ainda sem arquivo."""
    return [rel for rel, r in estado["relatorios"].items() if r.get("disparo") and not r.get("arquivo")]


def _orientacao_pendente() -> str:
    return ("procure o arquivo já exportado em Downloads (no RRA, a mensagem no e-mail) e registre com "
            "'extracao arquivo'; se não achar, pare a rodada e avise o usuário")


# --- a extração ----------------------------------------------------------------------------------

class Extracao:
    def __init__(self, pasta, estado: dict, agora=None):
        self.pasta, self.estado, self._relogio = Path(pasta), estado, agora or _relogio_padrao
        self._conferir = None
        self._trocas = []

    @classmethod
    def abrir(cls, pasta, agora=None) -> "Extracao":
        pasta = Path(pasta)
        return cls(pasta, _ler_estado(pasta), agora)

    def _agora(self) -> dt.datetime:
        return _momento(self._relogio)

    def fechada(self) -> bool:
        return (self.pasta / ARQ_MANIFESTO).is_file() or bool(self.estado.get("fechada_em"))

    def _mudar(self, fn):
        """Lê o estado do DISCO sob a trava (outro processo pode ter mudado), aplica `fn` numa cópia e só
        grava se mudou. Se `fn` falha, nada é gravado e a memória fica igual ao disco. Arquivos que `fn`
        prepara (print, relatório) ficam num nome temporário e só tomam o lugar definitivo depois da
        conferência final de posse, imediatamente antes da troca do estado.json."""
        base = self.pasta.parent
        with _trava(base) as dono:
            estado, bruto = _ler_estado_bruto(self.pasta)
            self.estado = estado
            if self.fechada():
                raise ExtracaoErro("esta extração já está fechada; não aceita mudanças")
            if estado.get("abandonada_em"):
                raise ExtracaoErro(f"esta extração foi abandonada (--nova criou {estado['abandonada_por']}); não "
                                   "aceita mais tela, disparo, arquivo nem fechamento. Use a extração nova")
            novo = copy.deepcopy(estado)
            self._trocas = []

            def conferir_e_trocar():
                _conferir_posse(base, dono, self.pasta / ARQ_ESTADO, bruto)
                while self._trocas:
                    self._trocas.pop(0)()

            self._conferir = conferir_e_trocar
            try:
                resultado = fn(novo)
                if novo != estado:
                    _gravar_json(self.pasta / ARQ_ESTADO, novo, conferir=self._conferir)
                    self.estado = novo
            finally:
                self._conferir = None
                for tmp in self.pasta.rglob("*.novo"):          # o que não chegou a ser trocado
                    with contextlib.suppress(OSError):
                        _apagar(tmp)
            return resultado

    def _preparar_copia(self, origem: Path, destino: Path) -> Path:
        """Copia `origem` para `<destino>.novo` agora; a troca para `destino` fica para depois da conferência."""
        tmp = destino.with_name(destino.name + ".novo")
        shutil.copyfile(origem, tmp)
        self._trocas.append(lambda: _insistir(lambda: os.replace(tmp, destino)))
        return tmp

    @staticmethod
    def _checar_rel(rel) -> None:
        if rel not in RELATORIOS:
            raise ExtracaoErro(f"relatório desconhecido: {_eco(rel)} (use {', '.join(RELATORIOS)})")

    # --- tela -----------------------------------------------------------------------------------
    def registrar_tela(self, rel: str, tela: dict, print_path) -> None:
        self._checar_rel(rel)
        if not isinstance(tela, dict):
            raise ExtracaoErro(f"{rel}: tela deve ser um dicionário de campos. NÃO exporte.")
        print_path = Path(print_path)

        def fn(e):
            r = e["relatorios"].setdefault(rel, {})
            if r.get("disparo"):
                raise ExtracaoErro(f"{rel}: já foi disparado; a tela registrada não pode mais mudar")
            erros = validar_tela(rel, e, tela)
            problema = _problema_print(print_path)
            if problema:
                erros.append(problema)
            if erros:
                # a tela mudou e está errada: uma validação anterior não pode mais liberar o disparo
                r.pop("tela", None)
                r.pop("prova", None)
                if not r:
                    del e["relatorios"][rel]
                return f"{rel}: trava da tela falhou - " + "; ".join(erros) + ". NÃO exporte."
            ext = ".jpg" if print_path.suffix.casefold() in (".jpg", ".jpeg") else ".png"
            prints = self.pasta / "prints"
            prints.mkdir(exist_ok=True)
            destino = prints / f"{rel}{ext}"
            self._preparar_copia(print_path, destino)

            def tirar_prints_velhos():
                for velho in prints.glob(f"{rel}.*"):
                    if velho.name != destino.name and velho.suffix != ".novo":
                        with contextlib.suppress(OSError):
                            _apagar(velho)
            self._trocas.append(tirar_prints_velhos)
            pesquisar_em = _texto(tela.get("pesquisar_em"))
            r["tela"] = {
                "localidade": _texto(tela["localidade"]), "criterio": _texto(tela["criterio"]),
                "data_inicio": _data(tela["data_inicio"]).isoformat(), "data_fim": _data(tela["data_fim"]).isoformat(),
                "status": _texto(tela["status"]), "total": tela["total"],
                **({"pesquisar_em": pesquisar_em} if pesquisar_em else {}),
                "vista_em": self._agora().isoformat()}
            r["prova"] = f"prints/{destino.name}"
            return None

        erro = self._mudar(fn)
        if erro:
            raise ExtracaoErro(erro)

    # --- disparo --------------------------------------------------------------------------------
    def rra_ids_antes(self, ids) -> None:
        if not isinstance(ids, (list, tuple)) or not all(isinstance(i, str) for i in ids):
            raise ExtracaoErro("rra: os ids anteriores devem ser uma lista de textos")
        limpos = sorted({i.strip() for i in ids if i.strip()})

        def fn(e):
            if e["relatorios"].get("rra", {}).get("disparo"):
                raise ExtracaoErro("rra: os ids anteriores só podem ser registrados antes do disparo; "
                                   "já existe disparo, não dá mais para registrar depois do disparo")
            e["rra_ids_antes"] = limpos

        self._mudar(fn)

    def disparar(self, rel: str) -> str:
        """Grava o horário do disparo ANTES do clique. Se isto falha, NÃO clique."""
        self._checar_rel(rel)

        def fn(e):
            r = e["relatorios"].get(rel, {})
            if r.get("disparo"):
                raise ExtracaoErro(f"{rel}: já disparado em {r['disparo']}: NÃO dispare de novo. "
                                   + _orientacao_pendente())
            if not r.get("tela"):
                raise ExtracaoErro(f"{rel}: registre a tela (extracao tela) antes de disparar")
            erros = validar_tela(rel, e, r["tela"])          # o que está gravado passa de novo pelas travas
            if erros:
                raise ExtracaoErro(f"{rel}: a tela gravada não passa nas travas - " + "; ".join(erros)
                                   + ". NÃO exporte; registre a tela de novo")
            agora = self._agora()
            vista = _instante(r["tela"]["vista_em"])
            if not 0 <= (agora - vista).total_seconds() <= TELA_VALIDADE_S:     # velha ou "no futuro"
                raise ExtracaoErro(f"{rel}: a tela não foi conferida nos últimos {TELA_VALIDADE_S // 60} min; "
                                   "confira e registre a tela de novo antes de disparar")
            if rel == "rra" and e.get("rra_ids_antes") is None:
                raise ExtracaoErro("rra: registre os ids dos e-mails existentes (extracao rra-ids-antes) "
                                   "antes de disparar")
            outras = _varrer(self.pasta.parent, exceto=self.pasta)      # já sob a trava da base
            duvida = _duvidas(outras)
            if duvida:
                raise ExtracaoErro(f"{rel}: NÃO dispare - {duvida}")
            if outras["pendentes"]:
                raise ExtracaoErro(f"{rel}: há outra extração desta competência com disparo sem arquivo "
                                   f"({', '.join(p.name for p in outras['pendentes'])}): NÃO dispare (o relatório "
                                   "sairia duas vezes). Retome aquela extração e " + _orientacao_pendente())
            r["disparo"] = agora.isoformat()
            return r["disparo"]

        return self._mudar(fn)

    def escolher_rra(self, candidatas) -> str:
        if not isinstance(candidatas, (list, tuple)):
            raise ExtracaoErro("rra: as candidatas devem ser uma lista de mensagens")

        def fn(e):
            if e.get("rra_mensagem"):
                return e["rra_mensagem"]                 # idempotente: a primeira escolha vale
            disparo = e["relatorios"].get("rra", {}).get("disparo")
            if not disparo:
                raise ExtracaoErro("rra: ainda não foi disparado")
            escolhida = rra_email.escolher_mensagem(candidatas, e.get("rra_ids_antes") or [], _instante(disparo))
            e["rra_mensagem"] = escolhida
            return escolhida

        return self._mudar(fn)

    # --- arquivo --------------------------------------------------------------------------------
    def registrar_arquivo(self, rel: str, caminho) -> dict:
        self._checar_rel(rel)
        src = Path(caminho)

        def fn(e):
            r = e["relatorios"].get(rel, {})
            if not r.get("disparo"):
                raise ExtracaoErro(f"{rel}: sem disparo registrado; não há o que baixar")
            if r.get("arquivo"):
                raise ExtracaoErro(f"{rel}: já tem arquivo registrado")
            if rel == "rra" and not e.get("rra_mensagem"):
                raise ExtracaoErro("rra: escolha antes a mensagem do e-mail (extracao rra-escolher)")
            if not src.is_file():
                raise ExtracaoErro(f"{rel}: arquivo não encontrado: {_eco(src.name)}")
            if src.resolve().is_relative_to(self.pasta.resolve()):
                raise ExtracaoErro(f"{rel}: o arquivo está dentro da pasta da extração; registre o arquivo baixado")
            st = src.stat()
            if st.st_size == 0:
                raise ExtracaoErro(f"{rel}: o arquivo está vazio")
            if src.suffix.casefold() != EXTENSAO[rel]:
                raise ExtracaoErro(f"{rel}: extensão {_eco(src.suffix)} inesperada (esperada {EXTENSAO[rel]})")
            if not _assinatura_ok(rel, src):
                raise ExtracaoErro(f"{rel}: o arquivo não parece um {EXTENSAO[rel]} válido")
            disparo = _instante(r["disparo"])
            modificado = dt.datetime.fromtimestamp(st.st_mtime, FUSO_BRASILIA)
            if modificado < disparo - dt.timedelta(seconds=TOLERANCIA_MTIME_S):
                raise ExtracaoErro(f"{rel}: o arquivo é anterior ao disparo (provavelmente um download antigo)")
            sha = sha256_file(src)
            for outro, ro in e["relatorios"].items():
                if outro != rel and (ro.get("arquivo") or {}).get("sha256") == sha:
                    raise ExtracaoErro(f"{rel}: é o mesmo arquivo já registrado para {outro}")
            destino = self.pasta / f"{rel}{EXTENSAO[rel]}"
            tmp = self._preparar_copia(src, destino)
            if tmp.stat().st_size != st.st_size or sha256_file(tmp) != sha:
                raise ExtracaoErro(f"{rel}: a cópia não confere com o original (o arquivo mudou durante a cópia?)")
            r["arquivo"] = {"nome": destino.name, "nome_original": src.name, "sha256": sha,
                            "tamanho": st.st_size, "registrado_em": self._agora().isoformat()}
            return dict(r["arquivo"])

        return self._mudar(fn)

    # --- fechamento -----------------------------------------------------------------------------
    def status(self) -> dict:
        rels = {rel: {"tela": bool(self.estado["relatorios"].get(rel, {}).get("tela")),
                      "disparo": bool(self.estado["relatorios"].get(rel, {}).get("disparo")),
                      "arquivo": bool(self.estado["relatorios"].get(rel, {}).get("arquivo"))}
                for rel in RELATORIOS}
        return {"pasta": self.pasta.name, "competencia": self.estado["competencia"], "relatorios": rels,
                "rra_ids_antes": self.estado.get("rra_ids_antes") is not None,
                "rra_mensagem": bool(self.estado.get("rra_mensagem")),
                "fechada": self.fechada(), "abandonada": bool(self.estado.get("abandonada_em"))}

    def fechar(self) -> dict:
        def fn(e):
            faltam = []
            for rel in RELATORIOS:
                r = e["relatorios"].get(rel, {})
                for chave, texto in (("tela", "tela validada"), ("disparo", "disparo"), ("arquivo", "arquivo")):
                    if not r.get(chave):
                        faltam.append(f"{rel}: falta {texto}")
            if not e.get("rra_mensagem"):
                faltam.append("rra: falta a mensagem do e-mail escolhida")
            if faltam:
                raise ExtracaoErro("extração incompleta - " + "; ".join(faltam))
            rels = {}
            for rel in RELATORIOS:
                r = e["relatorios"][rel]
                a = r["arquivo"]
                arq = self.pasta / a["nome"]
                if not arq.is_file():
                    raise ExtracaoErro(f"{rel}: arquivo não encontrado na extração: {a['nome']}")
                if arq.stat().st_size != a["tamanho"] or sha256_file(arq) != a["sha256"]:
                    raise ExtracaoErro(f"{rel}: o arquivo {a['nome']} não confere com o registrado (foi alterado)")
                prova = self.pasta / r["prova"]
                if not prova.is_file():
                    raise ExtracaoErro(f"{rel}: print não encontrado na extração: {r['prova']}")
                rels[rel] = {"arquivo": a["nome"], "arquivo_original": a.get("nome_original"),
                             "sha256": a["sha256"], "tamanho": a["tamanho"], "total_tela": r["tela"]["total"],
                             "filtros_tela": {k: v for k, v in r["tela"].items() if k != "total"},
                             "prova": r["prova"], "prova_sha256": sha256_file(prova), "disparo": r["disparo"]}
            primeiro = min(_instante(r["disparo"]) for r in rels.values())
            manifesto = {"versao": VERSAO, "produtora": e["produtora"],
                         "competencia": e["competencia"], "periodo_inicio": e["periodo_inicio"],
                         "periodo_fim": e["periodo_fim"],
                         "data_extracao": primeiro.astimezone(FUSO_BRASILIA).isoformat(), "origem": ORIGEM,
                         "rra_mensagem": e["rra_mensagem"], "relatorios": rels}
            _gravar_json(self.pasta / ARQ_MANIFESTO, manifesto, conferir=self._conferir)
            e["fechada_em"] = self._agora().isoformat()
            return manifesto

        return self._mudar(fn)


# --- iniciar / retomar ---------------------------------------------------------------------------

def _ordem(p: Path):
    """Pastas fora do padrão AAAAMMDD-HHMM[-n] (renomeadas à mão) vêm PRIMEIRO: nunca são 'a mais
    recente' para o --retomar. As do padrão seguem em ordem de data e sufixo numérico."""
    m = _RE_PASTA.fullmatch(p.name)
    return (1, m.group(1), int(m.group(2) or 1)) if m else (0, p.name, 0)


def _pastas_com_estado(base: Path) -> list[Path]:
    if not base.is_dir():
        return []
    return sorted((p for p in base.iterdir() if p.is_dir() and (p / ARQ_ESTADO).is_file()), key=_ordem)


def _pastas_no_padrao_sem_estado(base: Path) -> list[str]:
    """Pastas com nome de extração, sem estado.json e com algum conteúdo (estado renomeado, apagado...):
    não dá para saber se houve disparo. Uma pasta vazia (processo morto entre criar a pasta e gravar o
    estado) não pode ter disparo e não conta."""
    if not base.is_dir():
        return []
    return sorted(p.name for p in base.iterdir()
                  if p.is_dir() and _RE_PASTA.fullmatch(p.name) and not (p / ARQ_ESTADO).is_file()
                  and any(not _residuo(q) for q in p.iterdir()))


def _residuo(p: Path) -> bool:
    """Sobra de gravação interrompida, que não prova nada: estado.json.tmp e .trava*."""
    return p.is_file() and (p.name == ARQ_ESTADO + ".tmp" or p.name.startswith(ARQ_TRAVA))


def _varrer(base: Path, exceto: Path | None = None) -> dict:
    """Situação das extrações da base (fechadas com manifesto válido são puladas):
    sem_estado / ilegiveis = pastas em que não dá para saber se houve disparo; pendentes = com relatório
    disparado sem arquivo; abertas = nem fechadas nem abandonadas; fechadas = com manifesto válido."""
    v = {"sem_estado": [n for n in _pastas_no_padrao_sem_estado(base) if exceto is None or n != exceto.name],
         "ilegiveis": [], "pendentes": [], "abertas": [], "fechadas": []}
    for p in _pastas_com_estado(base):
        if exceto is not None and p.name == exceto.name:
            continue
        if (p / ARQ_MANIFESTO).is_file():
            try:
                carregar_manifesto(p)
                v["fechadas"].append(p)
                continue        # fechada e conferida: não tem disparo pendente, o estado.json não importa
            except ExtracaoErro:
                pass            # manifesto que não vale: segue as regras de uma pasta sem manifesto
        try:
            est = _ler_estado(p)
        except ExtracaoErro:
            v["ilegiveis"].append(p.name)
            continue
        if _pendentes(est):
            v["pendentes"].append(p)
        if not (p / ARQ_MANIFESTO).is_file() and not est.get("fechada_em") and not est.get("abandonada_em"):
            v["abertas"].append(p)
    return v


def _duvidas(v: dict) -> str | None:
    if v["sem_estado"]:
        return (f"pasta de extração sem {ARQ_ESTADO} ({', '.join(v['sem_estado'])}): não é seguro continuar "
                "(pode haver relatório disparado sem arquivo); pare e avise o usuário")
    if v["ilegiveis"]:
        return (f"{ARQ_ESTADO} ilegível ou inválido em {', '.join(v['ilegiveis'])}: não é seguro continuar "
                "(pode haver relatório disparado sem arquivo); pare e avise o usuário")
    return None


def iniciar(base, produtora: str, competencia: str, periodo_fim: str, *, retomar: bool = False,
            nova: bool = False, agora=None) -> Extracao:
    """Cria a pasta de uma extração em `base` (= .../extracoes/<competencia>/<produtora>) ou, com
    `retomar`, reabre a aberta. Uma extração com disparo sem arquivo nunca é abandonada, e um estado
    ilegível bloqueia (não dá para saber se há disparo pendente). Varredura, criação da pasta e gravação
    do estado correm sob a trava da base: dois `iniciar` simultâneos não criam duas extrações abertas."""
    if not produtoras.existe(produtora):
        raise ExtracaoErro(f"produtora desconhecida: {_eco(produtora)}")
    if retomar and nova:
        raise ExtracaoErro("use --retomar ou --nova, não os dois")
    if not isinstance(competencia, str) or not _RE_COMPETENCIA.fullmatch(competencia):
        raise ExtracaoErro(f"competência deve ser AAAA-MM; recebido {_eco(competencia)}")
    fim = _data_iso(periodo_fim)
    if fim is None:
        raise ExtracaoErro(f"o fim do período deve ser AAAA-MM-DD; recebido {_eco(periodo_fim)}")
    if fim.strftime("%Y-%m") != competencia:
        raise ExtracaoErro(f"o fim do período ({periodo_fim}) deve estar dentro da competência {competencia}")
    base = Path(base)
    base.mkdir(parents=True, exist_ok=True)
    with _trava(base) as dono:
        return _iniciar_travado(base, dono, produtora, competencia, periodo_fim, retomar, nova, agora)


def _iniciar_travado(base, dono, produtora, competencia, periodo_fim, retomar, nova, agora) -> Extracao:
    v = _varrer(base)
    duvida = _duvidas(v)
    if duvida:
        raise ExtracaoErro(duvida)
    pendentes, abertas = v["pendentes"], v["abertas"]

    if retomar:
        if not abertas:
            if v["fechadas"]:
                raise ExtracaoErro(f"não há extração aberta; a última ({v['fechadas'][-1].name}) está fechada: "
                                   "reimporte-a com importar --extracao; não extraia de novo sem o usuário pedir")
            raise ExtracaoErro("não há extração aberta para retomar")
        com_pendencia = [p for p in abertas if p in pendentes]
        if len(com_pendencia) > 1:
            raise ExtracaoErro("há mais de uma extração aberta com disparo sem arquivo "
                               f"({', '.join(p.name for p in com_pendencia)}); pare e avise o usuário")
        ex = Extracao.abrir(com_pendencia[0] if com_pendencia else abertas[-1], agora)
        if (ex.estado["periodo_fim"] != periodo_fim or ex.estado["competencia"] != competencia
                or ex.estado["produtora"] != produtora):
            raise ExtracaoErro("a extração aberta é de outra competência/período")
        return ex
    if pendentes:
        raise ExtracaoErro(f"há extração com disparo sem arquivo ({pendentes[-1].name}): NÃO comece outra "
                           "(dispararia o relatório de novo). Use --retomar e " + _orientacao_pendente())
    if abertas and not nova:
        raise ExtracaoErro(f"há extração aberta ({abertas[-1].name}): use --retomar "
                           "(ou --nova, que a abandona em definitivo: ela não aceitará mais nada)")

    momento = _momento(agora or _relogio_padrao)
    raiz = momento.strftime("%Y%m%d-%H%M")
    for n in range(1, 1000):                    # dois inícios no mesmo minuto: -2, -3, ...
        pasta = base / (raiz if n == 1 else f"{raiz}-{n}")
        try:
            pasta.mkdir()
            break
        except FileExistsError:
            continue
    else:
        raise ExtracaoErro(f"não foi possível criar uma pasta nova para {raiz}")
    estado = {"versao": VERSAO, "produtora": produtora, "competencia": competencia,
              "periodo_inicio": f"{competencia[:4]}-01-01", "periodo_fim": periodo_fim,
              "criada_em": momento.isoformat(), "rra_ids_antes": None, "rra_mensagem": None, "relatorios": {}}
    try:
        # --nova: as abertas (todas sem disparo pendente, conferido acima) ficam abandonadas e não aceitam
        # mais nada, para nunca haver duas extrações ativas disparando os mesmos relatórios
        for p in abertas:
            antigo, bruto = _ler_estado_bruto(p)
            antigo["abandonada_em"], antigo["abandonada_por"] = momento.isoformat(), pasta.name
            _gravar_json(p / ARQ_ESTADO, antigo,
                         conferir=lambda p=p, bruto=bruto: _conferir_posse(base, dono, p / ARQ_ESTADO, bruto))
        _gravar_json(pasta / ARQ_ESTADO, estado, conferir=lambda: _conferir_posse(base, dono))
    except BaseException:
        with contextlib.suppress(OSError):
            pasta.rmdir()
        raise
    return Extracao(pasta, estado, agora)


# --- manifesto -----------------------------------------------------------------------------------

def _problema_no_manifesto(m) -> str | None:
    if not isinstance(m, dict):
        return "não é um objeto"
    if m.get("versao") != VERSAO:
        return "versão inesperada"
    rels = m.get("relatorios")
    if not isinstance(rels, dict) or set(rels) != set(RELATORIOS):
        return "relatórios inesperados"
    problema = _problema_no_periodo(m)
    if problema:
        return problema
    for chave in ("origem", "rra_mensagem"):
        if not isinstance(m.get(chave), str) or not m[chave]:
            return f"campo {chave} ausente"
    if _instante(m.get("data_extracao")) is None:
        return "data_extracao ausente ou sem fuso"
    for rel in RELATORIOS:
        r = rels[rel]
        if not isinstance(r, dict) or r.get("arquivo") != f"{rel}{EXTENSAO[rel]}":
            return f"{rel}: nome de arquivo inesperado"
        for chave in ("sha256", "prova_sha256"):
            if not isinstance(r.get(chave), str) or not _RE_SHA.fullmatch(r[chave]):
                return f"{rel}: {chave} inválido"
        for chave in ("tamanho", "total_tela"):
            if isinstance(r.get(chave), bool) or not isinstance(r.get(chave), int):
                return f"{rel}: {chave} inválido"
        if not _prova_ok(rel, r.get("prova")):
            return f"{rel}: prova inválida"
        if not isinstance(r.get("filtros_tela"), dict) or _instante(r.get("disparo")) is None:
            return f"{rel}: filtros ou disparo inválidos"
        problema = _problema_na_tela_gravada(rel, m, {**r["filtros_tela"], "total": r["total_tela"]})
        if problema:
            return problema
    return None


def carregar_manifesto(pasta) -> dict:
    """Lê e confere `extracao.json`: estrutura, os 4 relatórios e cada arquivo e print (existência,
    tamanho, hash). Qualquer divergência -> ExtracaoErro."""
    pasta = Path(pasta)
    caminho = pasta / ARQ_MANIFESTO
    if not caminho.is_file():
        raise ExtracaoErro(f"a extração {pasta.name} não está fechada (sem {ARQ_MANIFESTO})")
    try:
        m = json.loads(caminho.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        raise ExtracaoErro(f"{ARQ_MANIFESTO} ilegível em {pasta.name}") from None
    problema = _problema_no_manifesto(m)
    if problema:
        raise ExtracaoErro(f"{ARQ_MANIFESTO} de {pasta.name} inválido: {problema}")
    for rel in RELATORIOS:
        r = m["relatorios"][rel]
        arq = pasta / r["arquivo"]
        if not arq.is_file():
            raise ExtracaoErro(f"{rel}: arquivo não encontrado na extração: {r['arquivo']}")
        if arq.stat().st_size != r["tamanho"] or sha256_file(arq) != r["sha256"]:
            raise ExtracaoErro(f"{rel}: o arquivo {r['arquivo']} não confere com o registrado (foi alterado)")
        prova = pasta / r["prova"]
        if not prova.is_file():
            raise ExtracaoErro(f"{rel}: print não encontrado na extração: {r['prova']}")
        if sha256_file(prova) != r["prova_sha256"]:
            raise ExtracaoErro(f"{rel}: o print {r['prova']} não confere com o registrado (foi alterado)")
    return m
