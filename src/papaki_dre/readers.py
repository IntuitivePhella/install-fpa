from __future__ import annotations

import csv
import hashlib
import io
import os
import re
import uuid
import zipfile
from pathlib import Path

import openpyxl

from .types import ContractError

_AUTOFILTER = re.compile(rb"<autoFilter\b[^>]*/>|<autoFilter\b[^>]*>.*?</autoFilter>", re.S)
_XL_OPEN_XML_WORKBOOK = 51


def read_delimited(path, delimiter: str = ";") -> tuple[str, list[list[str]]]:
    """Lê CSV do Esferas. A codificação varia por relatório: tenta UTF-8 (com ou
    sem BOM) e cai para Windows-1252 (CR) e, por último, Latin-1."""
    raw = Path(path).read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:  # pragma: no cover - latin-1 decodifica qualquer byte
        raise ContractError(f"não foi possível decodificar {path}")
    rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delimiter))
    return enc, rows


def _strip_autofilters(path) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                data = _AUTOFILTER.sub(b"", data)
            zout.writestr(item, data)
    buf.seek(0)
    return buf


def open_xlsx(path, *, data_only: bool = False) -> openpyxl.Workbook:
    """Abre um XLSX do Esferas. O arquivo de comissões vem com `autoFilter`
    de referência inválida, que o openpyxl não aceita; ele é removido em memória
    (o arquivo original não é alterado)."""
    return openpyxl.load_workbook(_strip_autofilters(path), data_only=data_only)


_RPC_E_CALL_REJECTED = -2147418111   # o Excel ainda está ocupado ao iniciar/abrir


def _rejeicao_escondida(e: AttributeError) -> bool:
    """O win32com dinâmico engole o 'chamada rejeitada' da busca do nome (GetIDsOfNames) e levanta
    AttributeError('<obj>.<nome>'). Só esse caso (levantado dentro de win32com/client/dynamic.py)
    é tratado como Excel ocupado; um AttributeError do nosso código continua falhando na hora."""
    import traceback

    tb = traceback.extract_tb(e.__traceback__)
    return bool(tb) and tb[-1].filename.replace("\\", "/").endswith("win32com/client/dynamic.py")


def com_retry(fn, tries: int = 15, wait: float = 2.0):
    """Repete uma chamada COM enquanto o Excel responde 'chamada rejeitada' (ocupado), inclusive
    quando o win32com a disfarça de AttributeError."""
    import time

    import pywintypes

    for i in range(tries):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult != _RPC_E_CALL_REJECTED or i == tries - 1:
                raise
            time.sleep(wait)
        except AttributeError as e:
            if not _rejeicao_escondida(e) or i == tries - 1:
                raise
            time.sleep(wait)


def com_tipado(obj):
    """Logo depois de Workbooks.Open o Excel rejeita chamadas por alguns segundos e o win32com
    embrulha o resultado SEM a type info (o GetTypeInfo rejeitado é engolido): sem ela, um método
    vira 'property get' (wb.Close executa sem argumentos e devolve bool). Espera o Excel aceitar o
    GetTypeInfo e reembrulha o mesmo objeto, agora com a type info."""
    import win32com.client

    ole = getattr(obj, "_oleobj_", None)
    if ole is None:                      # não é um objeto do win32com (ex.: dublê dos testes)
        return obj
    com_retry(lambda: ole.GetTypeInfo())
    return win32com.client.Dispatch(ole)


def xlsx_cache_path(src, dst) -> Path:
    """Destino da conversão no diretório de `dst`: `<stem de dst>-<sha256[:12] do conteúdo
    de src>.xlsx`. Dois brutos de mesmo nome e conteúdos diferentes nunca compartilham
    o destino; o mesmo conteúdo reaproveita a conversão."""
    h = hashlib.sha256()
    with open(src, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    dst = Path(dst)
    return dst.with_name(f"{dst.stem}-{h.hexdigest()[:12]}.xlsx")


def _materialize(target: Path, produce) -> Path:
    """Reaproveita `target` se existir; senão chama `produce(tmp)` com um temporário no
    mesmo diretório e o move para `target` com `os.replace`. Assim `target` só existe
    completo: uma conversão interrompida nunca é reaproveitada."""
    target = Path(target)
    if target.exists():
        return target
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_name(f".{target.stem}.{uuid.uuid4().hex[:8]}.tmp.xlsx")
    try:
        produce(tmp)
        if not tmp.exists():
            raise ContractError(f"a conversão não gerou {target}")
        os.replace(tmp, target)
    finally:
        if tmp.exists():
            tmp.unlink()
    return target


def _quietly(fn) -> None:
    """Limpeza do COM: uma falha aqui não pode mascarar a exceção original."""
    try:
        com_retry(fn)
    except Exception:  # noqa: BLE001 - limpeza de melhor esforço
        pass


def _excel_save_as_xlsx(src: Path, out: Path) -> None:
    import win32com.client  # import tardio: só existe no Windows com Excel

    excel = com_retry(lambda: win32com.client.DispatchEx("Excel.Application"))
    try:
        com_retry(lambda: setattr(excel, "Visible", False))
        com_retry(lambda: setattr(excel, "DisplayAlerts", False))
        wb = com_tipado(com_retry(lambda: excel.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=True)))
        try:
            com_retry(lambda: wb.SaveAs(str(out), FileFormat=_XL_OPEN_XML_WORKBOOK))
        finally:
            _quietly(lambda: wb.Close(False))
    finally:
        _quietly(lambda: excel.Quit())


def convert_xls_to_xlsx(src, dst) -> Path:
    """Converte XLS binário em XLSX real, preservando fórmulas, pelo próprio Excel (COM).
    O arquivo gerado fica no diretório de `dst`, com o hash do conteúdo de `src` no nome
    (ver `xlsx_cache_path`); a função devolve esse caminho."""
    src = Path(src).resolve()
    target = xlsx_cache_path(src, Path(dst).resolve())
    return _materialize(target, lambda tmp: _excel_save_as_xlsx(src, tmp))
