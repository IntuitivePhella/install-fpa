from __future__ import annotations

import shutil
import tempfile
import time
import warnings
from pathlib import Path

from .readers import _quietly, com_retry, com_tipado
from .types import ContractError

# Código COM de #N/D em Range.Value2: erros de célula chegam como inteiros negativos
# (xlErrNA = 2042, convertido por CVErr para -2146826246).
XL_ERR_NA = -2146826246


def _count_na(values) -> int:
    """Conta os #N/D numa região lida de Range.Value2 (tupla de tuplas; uma célula só vem como
    valor solto, região vazia como None). Só o inteiro de #N/D conta; os demais erros de célula
    (também inteiros, mas outros códigos) e os números normais não entram."""
    if values is None:
        return 0
    if not isinstance(values, (tuple, list)):
        values = ((values,),)
    n = 0
    for row in values:
        cells = row if isinstance(row, (tuple, list)) else (row,)
        n += sum(1 for v in cells if isinstance(v, int) and not isinstance(v, bool) and v == XL_ERR_NA)
    return n


def _remove_dir(path: Path) -> None:
    """Apaga a pasta temporária; se o Excel ainda segura o arquivo, tenta de novo (3 x 1 s) e, por
    fim, avisa o caminho em vez de falhar calado."""
    for i in range(3):
        shutil.rmtree(path, ignore_errors=True)
        if not path.exists():
            return
        if i < 2:
            time.sleep(1)
    warnings.warn(f"não foi possível apagar a pasta temporária do recálculo: {path}")


def contar_nd(xlsx, aba: str, colunas: tuple[str, str] | list[str], primeira: int, ultima: int) -> int:
    """Conta as células com #N/D em `aba`!<colunas[0]><primeira>:<colunas[-1]><ultima> depois de
    um recálculo completo feito pelo Excel. NUNCA recalcula o arquivo entregue: trabalha numa cópia
    temporária, aberta somente leitura numa instância PRÓPRIA do Excel (DispatchEx, invisível, sem
    alertas), fechada sem salvar; só essa instância é encerrada (as planilhas abertas do usuário não
    são tocadas) e a cópia é apagada no fim. Faixa vazia (`ultima < primeira`) devolve 0; aba
    inexistente levanta ContractError. Débito conhecido: sem timeout se o Excel travar."""
    if primeira < 1:
        raise ValueError(f"primeira linha inválida: {primeira}")
    if ultima < primeira:
        return 0
    import pywintypes
    import win32com.client  # import tardio: só existe no Windows com Excel

    xlsx = Path(xlsx)
    ref = f"{colunas[0]}{primeira}:{colunas[-1]}{ultima}"
    tmp_dir = Path(tempfile.mkdtemp(prefix="papaki_nd_"))
    excel = wb = values = None
    try:
        copia = tmp_dir / xlsx.name
        shutil.copyfile(xlsx, copia)
        excel = com_retry(lambda: win32com.client.DispatchEx("Excel.Application"))
        try:
            com_retry(lambda: setattr(excel, "Visible", False))
            com_retry(lambda: setattr(excel, "DisplayAlerts", False))
            wb = com_tipado(com_retry(lambda: excel.Workbooks.Open(str(copia), UpdateLinks=0, ReadOnly=True)))
            com_retry(lambda: wb.Application.CalculateFull())
            try:
                values = com_retry(lambda: wb.Worksheets(aba).Range(ref).Value2)
            except pywintypes.com_error as e:
                nomes = com_retry(lambda: [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)])
                if aba not in nomes:
                    raise ContractError(f"aba {aba!r} ausente no arquivo") from e
                raise
        finally:
            if wb is not None:
                _quietly(lambda: wb.Close(False))
            _quietly(lambda: excel.Quit())
        return _count_na(values)
    finally:
        wb = excel = values = None          # solta as referências COM antes de apagar a cópia
        _remove_dir(tmp_dir)
