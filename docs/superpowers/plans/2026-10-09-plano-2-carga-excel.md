# Plano 2 — Carga no Excel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Gravar os resultados dos sanitizadores (`Sanitized`) numa **cópia** do Modelo DRE, somente nas quatro abas de base, sem alterar nenhuma outra aba, tabela dinâmica ou gráfico do arquivo.

**Architecture:** Edição cirúrgica do XML das quatro abas dentro do `.xlsx` (lxml). Todas as demais partes do pacote são copiadas byte a byte. Textos entram como `inlineStr` (não toca `sharedStrings.xml`); estilos existentes das células são preservados; fórmulas do modelo em colunas de base (`AP:AQ`, `AC`) são **estendidas** às linhas novas (resolvendo fórmulas compartilhadas); `O:Q` das comissões vêm de constantes do modelo de setembro quando o modelo não as traz. O `calcChain.xml` é removido e `fullCalcOnLoad` ligado, para o Excel recalcular ao abrir. Depois de gravar, a **cobertura de fórmulas** é conferida no arquivo gerado (spec §8): `AP:AQ` (DESPESAS), `AC` (RECEITA), `L:N` (fórmulas de origem reendereçadas) e `O:Q` (Comissões) em **todas** as linhas carregadas, e a linha `TOTAL` da `BASE` (rótulo e as 35 somas). Lacunas vão para `SheetReport.missing_formulas`; no fluxo completo (Plano 3) qualquer lacuna deixa a importação `bloqueada` (código de saída 2).

**Decisões do usuário (09/10/2026):** as fórmulas de classificação/cálculo do modelo precisam valer para **todas** as linhas carregadas dos relatórios sanitizados; só se detecta classificação **ausente** (linha carregada sem a fórmula, ou modelo sem fórmula para estender). Não há totais de controle: as linhas `TOTAL` dos brutos são só técnicas (Plano 1); a única linha `TOTAL` deste plano é a da `BASE`, recriada pelo sistema.

**Tech Stack:** Python ≥ 3.11, lxml, openpyxl (só `Translator` e verificação de leitura), pytest, Excel COM (apenas no teste de abertura).

**Pré-requisito:** Plano 1 concluído (`Sanitized`, `Formula`, `XlError`, `ContractError`, `DEST`, os quatro sanitizadores).

**Spec:** `docs/superpowers/specs/2026-10-09-papaki-esferas-mvp-design.md` (§7 carga no Excel; §8 conferência).

**Nota de execução (13/10/2026):** o plano foi executado e revisado; o código vigente (`src/papaki_dre/xlsx_inject.py`, `src/papaki_dre/build.py`) diverge de alguns blocos abaixo, marcados **SUPERSEDIDO**. Siga o código, não esses blocos. Decisões e mudanças estão detalhadas na "Nota de execução" da seção Self-Review, no fim.

**Fatos do modelo-base verificados em 09/10/2026** (`Modelo DRE - Untitled - base.xlsx`): 170 partes; 24 tabelas dinâmicas, gráficos, `calcChain.xml`; as quatro abas de base estão **vazias**, só com estilos, inclusive os **cabeçalhos** (linhas 1–3 da `BASE`, linha 2 de `DESPESAS`, linha 1 de `RECEITA`, linha 1 das comissões); `DESPESAS!AP:AQ` têm fórmulas nas linhas 3 a 1159 (`AP` compartilhada, `AQ` explícita); `RECEITA!AC` tem fórmulas compartilhadas nas linhas 2 a 664; a aba de comissões (`REAL X ORÇADO COMISSÕES`) **não tem fórmulas** em `O:Q`; as linhas das abas são esparsas (só existem células estilizadas).

---

## File Structure

```
src/papaki_dre/xlsx_inject.py     # DateStyles, SheetEditor, WorkbookSession, changed_parts, to_serial
src/papaki_dre/build.py           # build_workbook(): leiautes por aba, extensão de fórmulas, total da BASE;
                                  # formula_coverage(): cobertura AP:AQ, AC, L:N, O:Q e TOTAL da BASE
tests/xlsx_factory.py             # fábrica de .xlsx mínimos para testes (XML escrito à mão)
tests/test_xlsx_inject.py
tests/test_build.py
tests/test_build_real.py          # modelo real + brutos de setembro (+ abertura no Excel)
```

---

### Task 1: Fábrica de modelos sintéticos e `DateStyles`

**Files:**
- Create: `tests/xlsx_factory.py`, `src/papaki_dre/xlsx_inject.py` (começa só com `DateStyles` e helpers)
- Test: `tests/test_xlsx_inject.py`

- [ ] **Step 1: Criar a fábrica `tests/xlsx_factory.py`**

```python
"""Gera um .xlsx mínimo, com XML escrito à mão, que imita o que importa do modelo real:
linhas esparsas, fórmulas compartilhadas, calcChain, uma aba que NÃO pode mudar."""
import zipfile
from pathlib import Path

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

# estilos: 0 geral | 1 data (14) | 2 dinheiro (4) | 3 data custom (164)
STYLES = HEAD + f"""<styleSheet xmlns="{MAIN}">
<numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts>
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

SHEET_NAMES = ["BASE", "DESPESAS", "RECEITA", "REAL X ORÇADO COMISSÕES", "OUTRA"]


def _sheet(dim, rows_xml):
    return HEAD + (f'<worksheet xmlns="{MAIN}" xmlns:r="{RNS}"><dimension ref="{dim}"/>'
                   f'<sheetData>{rows_xml}</sheetData></worksheet>')


def _styled_rows(rows, cells):
    """cells: lista de (coluna, estilo) para cada linha; linhas esparsas e com spans."""
    return "".join(
        f'<row r="{r}" spans="1:3">' + "".join(f'<c r="{c}{r}" s="{s}"/>' for c, s in cells) + "</row>"
        for r in rows)


def make_template(path, *, desp_formula_rows=5, rec_formula_rows=3, calc_chain=True):
    path = Path(path)
    base = _sheet("A1:C12", _styled_rows(range(1, 13), [("A", 0), ("B", 3), ("C", 2)]))

    desp_rows = ['<row r="2" spans="1:43"><c r="AP2" t="inlineStr"><is><t>De-Para</t></is></c></row>']
    last = 2 + desp_formula_rows
    for i, r in enumerate(range(3, last + 1)):
        master = (f'<f t="shared" ref="AP3:AP{last}" si="0">IF(A3="x","a","b")</f>' if i == 0
                  else '<f t="shared" si="0"/>')
        desp_rows.append(
            f'<row r="{r}" spans="3:43"><c r="C{r}" s="3"/><c r="AP{r}" t="e">{master}<v>#N/A</v></c>'
            f'<c r="AQ{r}" t="e"><f>LEN(AP{r})&amp;"!"</f><v>#N/A</v></c></row>')
    despesas = _sheet(f"A2:AQ{last}", "".join(desp_rows))

    rec_rows = ['<row r="1" spans="1:29"><c r="AC1" t="inlineStr"><is><t>DRE</t></is></c></row>']
    rlast = 1 + rec_formula_rows
    for i, r in enumerate(range(2, rlast + 1)):
        master = (f'<f t="shared" ref="AC2:AC{rlast}" si="0">IF(A2="","",A2)</f>' if i == 0
                  else '<f t="shared" si="0"/>')
        rec_rows.append(f'<row r="{r}" spans="1:29"><c r="AC{r}" t="str">{master}<v/></c></row>')
    receita = _sheet(f"A1:AC{rlast}", "".join(rec_rows))

    comissoes = _sheet("C2:O8", _styled_rows(range(2, 9), [("C", 3)]))
    outra = _sheet("A1:B1", '<row r="1"><c r="A1"><v>5</v></c><c r="B1"><f>A1*2</f><v>10</v></c></row>')
    sheets = [base, despesas, receita, comissoes, outra]

    wb_sheets = "".join(f'<sheet name="{n}" sheetId="{i}" r:id="rId{i}"/>'
                        for i, n in enumerate(SHEET_NAMES, start=1))
    workbook = HEAD + (f'<workbook xmlns="{MAIN}" xmlns:r="{RNS}"><sheets>{wb_sheets}</sheets>'
                       '<calcPr calcId="191029"/></workbook>')
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="{RNS}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, 6))
    rels += f'<Relationship Id="rId50" Type="{RNS}/styles" Target="styles.xml"/>'
    if calc_chain:
        rels += f'<Relationship Id="rId51" Type="{RNS}/calcChain" Target="calcChain.xml"/>'
    wb_rels = HEAD + f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>'

    ct = ''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, 6))
    ct += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    if calc_chain:
        ct += '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
    content_types = HEAD + ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                            '<Default Extension="xml" ContentType="application/xml"/>'
                            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                            f'{ct}</Types>')
    root_rels = HEAD + (f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        f'<Relationship Id="rId1" Type="{RNS}/officeDocument" Target="xl/workbook.xml"/></Relationships>')

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", STYLES)
        for i, xml in enumerate(sheets, start=1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", xml)
        if calc_chain:
            z.writestr("xl/calcChain.xml", HEAD + f'<calcChain xmlns="{MAIN}"><c r="AQ3" i="2"/></calcChain>')
    return path
```

- [ ] **Step 2: Escrever os testes de `DateStyles` e `to_serial` (falham)**

`tests/test_xlsx_inject.py` (começo; as próximas tasks acrescentam testes ao mesmo arquivo):

```python
import datetime as dt

import pytest

from papaki_dre.xlsx_inject import DateStyles, to_serial
from xlsx_factory import STYLES, make_template


def test_date_styles_recognise_builtin_and_custom_date_formats():
    ds = DateStyles(STYLES.encode("utf-8"))
    assert not ds.is_date(None)       # estilo 0: geral
    assert ds.is_date("1")            # numFmtId 14
    assert not ds.is_date("2")        # numFmtId 4 (dinheiro)
    assert ds.is_date("3")            # custom dd/mm/yyyy
    assert not ds.is_date("99")       # fora da tabela


def test_date_styles_without_styles_part():
    assert not DateStyles(None).is_date("1")


def test_to_serial_known_values():
    assert to_serial(dt.datetime(2026, 1, 1)) == 46023
    assert to_serial(dt.datetime(1900, 3, 1)) == 61
    assert to_serial(dt.datetime(2026, 1, 1, 12)) == pytest.approx(46023.5)
    assert to_serial(dt.date(2026, 1, 2)) == 46024
```

- [ ] **Step 3: Rodar e ver falhar**

Run: `py -m pytest tests/test_xlsx_inject.py -v`
Expected: FAIL com `ModuleNotFoundError: No module named 'papaki_dre.xlsx_inject'`

- [ ] **Step 4: Criar `xlsx_inject.py` (primeira parte)**

```python
"""Edição cirúrgica de um .xlsx: só o XML das abas pedidas muda; o resto do pacote
(tabelas dinâmicas, gráficos, estilos, strings compartilhadas) é copiado byte a byte."""
from __future__ import annotations

import bisect
import datetime as dt
import math
import re
import unicodedata
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import range_boundaries

from .types import ContractError, Formula, XlError

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
_EPOCH = dt.datetime(1899, 12, 30)
_ILLEGAL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")
_REF = re.compile(r"^([A-Z]+)(\d+)$")


def _q(tag: str) -> str:
    return f"{{{NS}}}{tag}"


def to_serial(value) -> float | int:
    """datetime/date -> número de série do Excel (sistema 1900)."""
    if not isinstance(value, dt.datetime):
        value = dt.datetime(value.year, value.month, value.day)
    delta = value - _EPOCH
    serial = delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6
    return int(serial) if serial == int(serial) else serial


def _num(x) -> str:
    if isinstance(x, int):
        return str(x)
    if not math.isfinite(x):
        raise ContractError(f"número inválido para o Excel: {x!r}")
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    return repr(x)


class DateStyles:
    """Diz se um índice de estilo (`s=`) é um formato de data, lendo styles.xml."""
    BUILTIN = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47} | set(range(50, 59))

    def __init__(self, styles_xml: bytes | None):
        self._flags: list[bool] = []
        if not styles_xml:
            return
        root = etree.fromstring(styles_xml)
        codes = {}
        numfmts = root.find(_q("numFmts"))
        if numfmts is not None:
            for e in numfmts:
                codes[int(e.get("numFmtId"))] = e.get("formatCode", "")
        xfs = root.find(_q("cellXfs"))
        if xfs is not None:
            for xf in xfs:
                fid = int(xf.get("numFmtId", "0"))
                self._flags.append(self._is_date_format(fid, codes.get(fid)))

    @classmethod
    def _is_date_format(cls, fid: int, code: str | None) -> bool:
        if fid in cls.BUILTIN:
            return True
        if code is None:
            return False
        stripped = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", code)
        return bool(re.search(r"[dmyhs]", stripped, re.I))

    def is_date(self, style) -> bool:
        i = int(style) if style not in (None, "") else 0
        return i < len(self._flags) and self._flags[i]
```

- [ ] **Step 5: Rodar e ver passar**

Run: `py -m pytest tests/test_xlsx_inject.py -v`
Expected: 3 passed

- [ ] **Step 6: Commit**

```bash
git add tests/xlsx_factory.py tests/test_xlsx_inject.py src/papaki_dre/xlsx_inject.py && git commit -m "feat: fábrica de modelos sintéticos, reconhecimento de estilos de data e serial do Excel"
```

---

### Task 2: `SheetEditor` — gravar células sem mexer no resto

**Files:**
- Modify: `src/papaki_dre/xlsx_inject.py`
- Test: `tests/test_xlsx_inject.py`

- [ ] **Step 1: Acrescentar os testes (falham)**

Acrescente ao final de `tests/test_xlsx_inject.py`:

```python
import zipfile

from papaki_dre.types import ContractError, Formula, XlError
from papaki_dre.xlsx_inject import SheetEditor, WorkbookSession


def _editor(tmp_path, sheet="BASE", **kw):
    tpl = make_template(tmp_path / "t.xlsx", **kw)
    s = WorkbookSession(tpl)
    return s, s.editor(sheet)


def test_write_all_value_types_and_read_back_with_openpyxl(tmp_path):
    import openpyxl
    s, ed = _editor(tmp_path)
    ed.write(5, 1, "texto com <&> e acentuação ç", date_styles=s.date_styles)
    ed.write(5, 2, dt.datetime(2026, 3, 17), date_styles=s.date_styles)
    ed.write(5, 3, 1234.5, date_styles=s.date_styles)
    ed.write(6, 1, 7, date_styles=s.date_styles)
    ed.write(6, 2, True)
    ed.write(6, 3, Formula("=A6*2"))
    ed.write(7, 1, XlError("#DIV/0!"))
    s.save(tmp_path / "o.xlsx")
    ws = openpyxl.load_workbook(tmp_path / "o.xlsx")["BASE"]
    assert ws["A5"].value == "texto com <&> e acentuação ç"
    assert ws["B5"].value == dt.datetime(2026, 3, 17)       # estilo de data preservado
    assert ws["C5"].value == 1234.5
    assert ws["A6"].value == 7 and ws["B6"].value is True
    assert ws["C6"].value == "=A6*2"
    assert ws["A7"].value == "#DIV/0!" and ws["A7"].data_type == "e"


def test_none_writes_nothing_and_creates_no_row(tmp_path):
    _, ed = _editor(tmp_path)
    ed.write(500, 1, None)
    assert not ed.row_exists(500)


def test_existing_style_is_kept_and_new_cells_copy_style_from_reference(tmp_path):
    s, ed = _editor(tmp_path, "DESPESAS")
    ed.write(3, 3, "x")                              # C3 já existe com s=3
    assert ed.style_of(3, 3) == "3"
    ed.write(20, 3, "y", style_from=(3, 3))           # linha nova copia o estilo de C3
    assert ed.style_of(20, 3) == "3"


def test_rows_and_cells_stay_in_order_when_inserting(tmp_path):
    s, ed = _editor(tmp_path, "DESPESAS")
    ed.write(15, 5, "e")
    ed.write(15, 1, "a")
    ed.write(10, 2, "b")
    out = tmp_path / "o.xlsx"
    s.save(out)
    import re
    xml = zipfile.ZipFile(out).read("xl/worksheets/sheet2.xml").decode()
    rows = [int(x) for x in re.findall(r'<row r="(\d+)"', xml)]
    assert rows == sorted(rows)
    assert re.findall(r'<c r="([A-Z]+)15"', xml) == ["A", "E"]


def test_refuses_to_overwrite_a_model_formula(tmp_path):
    _, ed = _editor(tmp_path, "DESPESAS")
    with pytest.raises(ContractError, match="AP3"):
        ed.write(3, 42, "x")


def test_illegal_xml_characters_are_removed(tmp_path):
    import openpyxl
    s, ed = _editor(tmp_path)
    ed.write(5, 1, "a\x00b\x0bc")
    s.save(tmp_path / "o.xlsx")
    assert openpyxl.load_workbook(tmp_path / "o.xlsx")["BASE"]["A5"].value == "abc"


def test_date_in_cell_without_date_format_is_counted(tmp_path):
    s, ed = _editor(tmp_path, "RECEITA")
    ed.write(2, 2, dt.datetime(2026, 1, 5), date_styles=s.date_styles)
    assert ed.date_warnings == {"B": 1}


def test_nan_is_rejected(tmp_path):
    _, ed = _editor(tmp_path)
    with pytest.raises(ContractError):
        ed.write(5, 1, float("nan"))


def test_unsupported_type_is_rejected(tmp_path):
    _, ed = _editor(tmp_path)
    with pytest.raises(ContractError):
        ed.write(5, 1, object())


def test_nonempty_cells_detects_dirty_region(tmp_path):
    _, ed = _editor(tmp_path, "DESPESAS")
    assert ed.nonempty_cells(range(3, 8), {1, 2, 3}) == []
    ed.write(4, 1, "sujeira")
    assert ed.nonempty_cells(range(3, 8), {1, 2, 3}) == ["A4"]
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_xlsx_inject.py -v`
Expected: FAIL com `ImportError: cannot import name 'SheetEditor'`

- [ ] **Step 3: Implementar `SheetEditor` e `WorkbookSession` (parte de gravação)**

Acrescente a `src/papaki_dre/xlsx_inject.py`:

```python
class SheetEditor:
    def __init__(self, xml: bytes):
        parser = etree.XMLParser(huge_tree=True, remove_blank_text=False)
        self.root = etree.fromstring(xml, parser)
        self.sheet_data = self.root.find(_q("sheetData"))
        if self.sheet_data is None:
            raise ContractError("a aba não tem sheetData")
        self._rows = {int(r.get("r")): r for r in self.sheet_data.findall(_q("row"))}
        self._rownums = sorted(self._rows)
        self._masters: dict[str, tuple[str, str]] | None = None
        self.date_warnings: dict[str, int] = {}
        self.touched_rows: set[int] = set()

    # ---- navegação -------------------------------------------------------
    def row_exists(self, r: int) -> bool:
        return r in self._rows

    def _row(self, r: int, create: bool = True):
        row = self._rows.get(r)
        if row is None and create:
            row = etree.Element(_q("row"))
            row.set("r", str(r))
            pos = bisect.bisect_left(self._rownums, r)
            self.sheet_data.insert(pos, row)
            self._rownums.insert(pos, r)
            self._rows[r] = row
        return row

    def _cell(self, r: int, col: int, create: bool = True):
        row = self._row(r, create)
        if row is None:
            return None
        pos = 0
        for i, c in enumerate(row.findall(_q("c"))):
            cc = column_index_from_string(_REF.match(c.get("r")).group(1))
            if cc == col:
                return c
            if cc > col:
                break
            pos = i + 1
        if not create:
            return None
        cell = etree.Element(_q("c"))
        cell.set("r", f"{get_column_letter(col)}{r}")
        row.insert(pos, cell)
        return cell

    def style_of(self, r: int, col: int) -> str | None:
        c = self._cell(r, col, create=False)
        return None if c is None else c.get("s")

    def nonempty_cells(self, rows, cols: set[int]) -> list[str]:
        found = []
        for r in rows:
            row = self._rows.get(r)
            if row is None:
                continue
            for c in row.findall(_q("c")):
                ci = column_index_from_string(_REF.match(c.get("r")).group(1))
                if ci in cols and any(c.find(_q(t)) is not None for t in ("v", "f", "is")):
                    found.append(c.get("r"))
        return found

    # ---- gravação ----------------------------------------------------------
    def write(self, r: int, col: int, value, *, style_from: tuple[int, int] | None = None,
              date_styles: DateStyles | None = None) -> None:
        if value is None:
            return
        c = self._cell(r, col)
        if c.find(_q("f")) is not None:
            raise ContractError(f"{c.get('r')} contém fórmula do modelo; não será sobrescrita")
        if c.get("s") is None and style_from is not None:
            ref = self._cell(style_from[0], style_from[1], create=False)
            if ref is not None and ref.get("s") is not None:
                c.set("s", ref.get("s"))
        if (isinstance(value, (dt.datetime, dt.date)) and date_styles is not None
                and not date_styles.is_date(c.get("s"))):
            letter = get_column_letter(col)
            self.date_warnings[letter] = self.date_warnings.get(letter, 0) + 1
        self._set(c, value)
        self._rows[r].attrib.pop("spans", None)
        self.touched_rows.add(r)

    def _set(self, c, value) -> None:
        for child in list(c):
            c.remove(child)
        c.attrib.pop("t", None)
        if isinstance(value, Formula):
            text = str(value)
            etree.SubElement(c, _q("f")).text = text[1:] if text.startswith("=") else text
        elif isinstance(value, XlError):
            c.set("t", "e")
            etree.SubElement(c, _q("v")).text = str(value)
        elif isinstance(value, bool):
            c.set("t", "b")
            etree.SubElement(c, _q("v")).text = "1" if value else "0"
        elif isinstance(value, (dt.datetime, dt.date)):
            etree.SubElement(c, _q("v")).text = _num(to_serial(value))
        elif isinstance(value, (int, float)):
            etree.SubElement(c, _q("v")).text = _num(value)
        elif isinstance(value, str):
            c.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(c, _q("is")), _q("t"))
            t.text = _ILLEGAL.sub("", value)
            t.set(f"{{{XML_NS}}}space", "preserve")
        else:
            raise ContractError(f"tipo de valor não suportado em {c.get('r')}: {type(value).__name__}")

    # ---- serialização ------------------------------------------------------
    def to_bytes(self) -> bytes:
        self._update_dimension()
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _update_dimension(self) -> None:
        dim = self.root.find(_q("dimension"))
        if dim is None:
            return
        lo_r = lo_c = 10 ** 9
        hi_r = hi_c = 0
        for row in self.sheet_data.findall(_q("row")):
            for c in row.findall(_q("c")):
                m = _REF.match(c.get("r"))
                ci, ri = column_index_from_string(m.group(1)), int(m.group(2))
                lo_r, hi_r = min(lo_r, ri), max(hi_r, ri)
                lo_c, hi_c = min(lo_c, ci), max(hi_c, ci)
        if not hi_r:
            return
        c1, r1, c2, r2 = range_boundaries(dim.get("ref"))
        dim.set("ref", f"{get_column_letter(min(c1, lo_c))}{min(r1, lo_r)}:"
                       f"{get_column_letter(max(c2, hi_c))}{max(r2, hi_r)}")


def _sheet_parts(zf: zipfile.ZipFile) -> dict[str, str]:
    wb = etree.fromstring(zf.read("xl/workbook.xml"))
    rels = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels}
    parts = {}
    for sheet in wb.find(_q("sheets")):
        target = targets[sheet.get(f"{{{R_NS}}}id")]
        part = target.lstrip("/") if target.startswith("/") else "xl/" + target
        parts[unicodedata.normalize("NFC", sheet.get("name"))] = part
    return parts


class WorkbookSession:
    """Abre um .xlsx modelo, edita só as abas pedidas e grava uma cópia."""

    def __init__(self, template):
        self.template = Path(template)
        self._zin = zipfile.ZipFile(self.template)
        self.parts = _sheet_parts(self._zin)
        self._editors: dict[str, SheetEditor] = {}
        names = self._zin.namelist()
        self.date_styles = DateStyles(
            self._zin.read("xl/styles.xml") if "xl/styles.xml" in names else None)

    def editor(self, sheet_name: str) -> SheetEditor:
        key = unicodedata.normalize("NFC", sheet_name)
        if key not in self.parts:
            raise ContractError(f"o modelo não tem a aba {sheet_name!r}")
        if key not in self._editors:
            self._editors[key] = SheetEditor(self._zin.read(self.parts[key]))
        return self._editors[key]

    def save(self, out_path) -> Path:
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        replaced = {self.parts[k]: ed.to_bytes() for k, ed in self._editors.items()}
        has_chain = "xl/calcChain.xml" in self._zin.namelist()
        with zipfile.ZipFile(out_path, "w") as zout:
            for info in self._zin.infolist():
                if info.filename == "xl/calcChain.xml":
                    continue
                data = replaced.get(info.filename)
                if data is None:
                    data = self._zin.read(info.filename)
                    if info.filename == "xl/workbook.xml":
                        data = _force_full_calc(data)
                    elif info.filename == "[Content_Types].xml" and has_chain:
                        data = re.sub(rb'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', b"", data)
                    elif info.filename == "xl/_rels/workbook.xml.rels" and has_chain:
                        data = re.sub(rb'<Relationship[^>]*Target="[^"]*calcChain\.xml"[^>]*/>', b"", data)
                zi = zipfile.ZipInfo(info.filename, info.date_time)
                zi.compress_type = info.compress_type
                zi.external_attr = info.external_attr
                zout.writestr(zi, data)
        return out_path

    def close(self) -> None:
        self._zin.close()


def _force_full_calc(data: bytes) -> bytes:
    if b"<calcPr" not in data:
        raise ContractError("workbook.xml sem <calcPr>; não é possível forçar o recálculo ao abrir")
    if b"fullCalcOnLoad" in data:
        return data
    return re.sub(rb"<calcPr\b", b'<calcPr fullCalcOnLoad="1"', data, count=1)
```

> **SUPERSEDIDO em 13/10/2026:** ver código vigente em `src/papaki_dre/xlsx_inject.py` (`WorkbookSession.save` grava num temporário da mesma pasta e só substitui a saída no sucesso, recusa gravar sobre o próprio modelo e monta as partes de ligação em `_plumbing`/`_sync_filter_names`; `_force_full_calc` aceita qualquer ordem de atributos, aspas simples ou duplas e prefixo de namespace, e força `fullCalcOnLoad="1"` mesmo quando o modelo traz `"0"`; `SheetEditor` recebe o nome da aba); a nota de execução no topo explica.

- [ ] **Step 4: Rodar e ver passar**

Run: `py -m pytest tests/test_xlsx_inject.py -v`
Expected: todos passam. Se `test_rows_and_cells_stay_in_order_when_inserting` falhar, o problema é a posição de inserção de linhas/células: depure `_row`/`_cell`, não o teste.

- [ ] **Step 5: Commit**

```bash
git add src/papaki_dre/xlsx_inject.py tests/test_xlsx_inject.py && git commit -m "feat: SheetEditor e WorkbookSession para gravação cirúrgica de células"
```

---

### Task 3: Fórmulas compartilhadas, extensão e cópia byte a byte

**Files:**
- Modify: `src/papaki_dre/xlsx_inject.py`
- Test: `tests/test_xlsx_inject.py`

- [ ] **Step 1: Acrescentar os testes (falham)**

```python
from papaki_dre.xlsx_inject import changed_parts


def test_formula_at_resolves_shared_formula_children(tmp_path):
    _, ed = _editor(tmp_path, "DESPESAS")
    assert ed.formula_at(3, 42) == '=IF(A3="x","a","b")'     # mestre
    assert ed.formula_at(5, 42) == '=IF(A5="x","a","b")'     # filha: traduzida para a linha 5
    assert ed.formula_at(5, 43) == '=LEN(AP5)&"!"'            # explícita
    assert ed.formula_at(4, 1) is None


def test_text_at_reads_inline_strings_only(tmp_path):
    _, ed = _editor(tmp_path)
    ed.write(8, 1, "TOTAL")
    ed.write(9, 1, 5)
    assert ed.text_at(8, 1) == "TOTAL"
    assert ed.text_at(9, 1) is None          # número
    assert ed.text_at(500, 1) is None        # célula inexistente


def test_last_formula_row(tmp_path):
    _, ed = _editor(tmp_path, "DESPESAS", desp_formula_rows=5)
    assert ed.last_formula_row(42) == 7
    assert ed.last_formula_row(43) == 7
    assert ed.last_formula_row(1) is None


def test_extension_writes_explicit_formulas_translated_from_last_row(tmp_path):
    s, ed = _editor(tmp_path, "DESPESAS", desp_formula_rows=5)
    base = ed.formula_at(7, 42)
    for r in (8, 9):
        ed.write(r, 42, Formula(Translator(base, origin="AP7").translate_formula(f"AP{r}")))
    assert ed.formula_at(9, 42) == '=IF(A9="x","a","b")'


def test_save_removes_calc_chain_and_forces_full_calc(tmp_path):
    s, ed = _editor(tmp_path)
    out = s.save(tmp_path / "o.xlsx")
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        assert "xl/calcChain.xml" not in names
        assert b'fullCalcOnLoad="1"' in z.read("xl/workbook.xml")
        assert b"calcChain" not in z.read("[Content_Types].xml")
        assert b"calcChain" not in z.read("xl/_rels/workbook.xml.rels")


def test_all_other_parts_are_byte_identical(tmp_path):
    tpl = make_template(tmp_path / "t.xlsx")
    s = WorkbookSession(tpl)
    s.editor("DESPESAS").write(3, 1, "x")
    out = s.save(tmp_path / "o.xlsx")
    diff = changed_parts(tpl, out)
    assert diff["changed"] == sorted(["xl/worksheets/sheet2.xml", "xl/workbook.xml",
                                      "[Content_Types].xml", "xl/_rels/workbook.xml.rels"])
    assert diff["removed"] == ["xl/calcChain.xml"]
    assert diff["added"] == []


def test_untouched_sheet_keeps_formula_and_cached_value(tmp_path):
    import openpyxl
    tpl = make_template(tmp_path / "t.xlsx")
    s = WorkbookSession(tpl)
    s.editor("BASE").write(5, 1, "x")
    out = s.save(tmp_path / "o.xlsx")
    ws = openpyxl.load_workbook(out, data_only=True)["OUTRA"]
    assert ws["B1"].value == 10        # valor em cache da aba que não pode mudar


def test_dimension_grows_with_new_rows(tmp_path):
    import re
    s, ed = _editor(tmp_path, "DESPESAS")
    ed.write(300, 1, "x")
    out = s.save(tmp_path / "o.xlsx")
    xml = zipfile.ZipFile(out).read("xl/worksheets/sheet2.xml").decode()
    assert re.search(r'<dimension ref="A2:AQ300"', xml)


def test_unknown_sheet_is_a_contract_error(tmp_path):
    s = WorkbookSession(make_template(tmp_path / "t.xlsx"))
    with pytest.raises(ContractError, match="NAO_EXISTE"):
        s.editor("NAO_EXISTE")
```

Esses testes usam o `Translator` do openpyxl: acrescente `from openpyxl.formula.translate import Translator` ao bloco de imports no topo de `tests/test_xlsx_inject.py`.

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_xlsx_inject.py -v`
Expected: FAIL (`formula_at`, `text_at`, `last_formula_row`, `changed_parts` ainda não existem)

- [ ] **Step 3: Implementar `formula_at`, `text_at`, `last_formula_row` e `changed_parts`**

Acrescente dentro da classe `SheetEditor` (antes de `to_bytes`):

```python
    # ---- fórmulas ------------------------------------------------------------
    def _build_masters(self) -> None:
        self._masters = {}
        for c in self.sheet_data.iter(_q("c")):
            f = c.find(_q("f"))
            if f is not None and f.get("t") == "shared" and f.text:
                self._masters[f.get("si")] = (c.get("r"), f.text)

    def formula_at(self, r: int, col: int) -> str | None:
        """Texto da fórmula da célula (com '='), resolvendo fórmulas compartilhadas."""
        c = self._cell(r, col, create=False)
        if c is None:
            return None
        f = c.find(_q("f"))
        if f is None:
            return None
        if f.get("t") == "shared" and not f.text:
            if self._masters is None:
                self._build_masters()
            origin, text = self._masters[f.get("si")]
            return Translator("=" + text, origin=origin).translate_formula(c.get("r"))
        return "=" + (f.text or "")

    def text_at(self, r: int, col: int) -> str | None:
        """Texto de uma célula `inlineStr` (o tipo que este módulo grava); None nos demais casos.
        Usado na cobertura para conferir o rótulo `TOTAL` da BASE."""
        c = self._cell(r, col, create=False)
        if c is None or c.get("t") != "inlineStr":
            return None
        is_ = c.find(_q("is"))
        return None if is_ is None else "".join(is_.itertext())

    def last_formula_row(self, col: int) -> int | None:
        for r in reversed(self._rownums):
            c = self._cell(r, col, create=False)
            if c is not None and c.find(_q("f")) is not None:
                return r
        return None
```

> **SUPERSEDIDO em 13/10/2026:** ver código vigente em `src/papaki_dre/xlsx_inject.py` (`last_formula_row(col, first=1)` devolve a última fórmula **da mesma família** das linhas de dados, ignora resumos/rodapés e levanta `ContractError` quando a coluna só tem fórmulas ilegíveis; `formula_at` resolve filhas compartilhadas com `FormulaShifter`); a nota de execução no topo explica.

E, no fim do módulo:

```python
def changed_parts(template, output) -> dict[str, list[str]]:
    """Compara dois .xlsx parte a parte (bytes)."""
    with zipfile.ZipFile(template) as a, zipfile.ZipFile(output) as b:
        an, bn = set(a.namelist()), set(b.namelist())
        changed = sorted(n for n in an & bn if a.read(n) != b.read(n))
    return {"changed": changed, "removed": sorted(an - bn), "added": sorted(bn - an)}
```

- [ ] **Step 4: Rodar e ver passar**

Run: `py -m pytest tests/test_xlsx_inject.py -v`
Expected: todos passam.

- [ ] **Step 5: Commit**

```bash
git add src/papaki_dre/xlsx_inject.py tests/test_xlsx_inject.py && git commit -m "feat: resolução de fórmulas compartilhadas, leitura de texto, remoção do calcChain e verificação byte a byte"
```

---

### Task 4: `build_workbook` — leiaute por aba, extensão de fórmulas, total da BASE, cobertura

**Files:**
- Create: `src/papaki_dre/build.py`
- Test: `tests/test_build.py`

Regras (spec §7): destino por relatório vem de `DEST`; escrever **cabeçalhos** (o modelo-base os traz em branco) e, no RRA, o título em `A1`; dados a partir de `first_row`; `DESPESAS!AP:AQ` e `RECEITA!AC` são estendidas a partir da última fórmula do modelo a **todas** as linhas carregadas (sem fórmula nenhuma no modelo = `ContractError`, "classificação ausente bloqueia a publicação"; no fluxo completo do Plano 3 isso vira importação `bloqueada` com o motivo, nenhuma linha gravada, código de saída 2); comissões: as fórmulas de origem `L:N` chegam do sanitizador já reendereçadas (`Formula` em `Sanitized.rows`) e são gravadas como estão; se o modelo não tiver fórmulas em `O:Q`, usar as constantes de setembro; `BASE`: linha de total = última linha de dados + 2, `SUM(X4:X<última+1>)` nas 35 colunas do modelo de setembro; região de destino precisa estar vazia; fórmulas órfãs (abaixo da última linha carregada) não são removidas, só contadas.

**Cobertura (spec §8), conferida no arquivo gravado por `formula_coverage`:** `AP:AQ` (DESPESAS), `AC` (RECEITA), `L:N` e `O:Q` (Comissões) em cada linha de `first_row` a `last_row`, e a linha `TOTAL` da `BASE` (rótulo `TOTAL` em `A` e `=SUM(X4:X<última+1>)` exata em cada uma das 35 colunas). Cada lacuna entra em `SheetReport.missing_formulas` (ex.: `L3`, `AP1112`, `A78`, `J78`). `build_workbook` **não** levanta erro por lacuna de cobertura: devolve o relatório, e a conferência do Plano 3 marca a importação `bloqueada` (código de saída 2; a saída vira `*_BLOQUEADO.xlsx`; linhas e exclusões ficam no banco sem se tornarem vigentes).

- [ ] **Step 1: Escrever os testes que falham**

```python
import datetime as dt

import openpyxl
import pytest

from openpyxl.utils import column_index_from_string

from papaki_dre.build import SUM_COLUMNS, SheetReport, build_workbook, formula_coverage
from papaki_dre.types import ContractError, Formula, Sanitized
from papaki_dre.xlsx_inject import WorkbookSession, changed_parts
from xlsx_factory import make_template

D = dt.datetime(2026, 3, 17)
COM_HEADER = ["Modelo", "Núm. JOB", "Data", *[f"col{j}" for j in range(4, 15)]]   # 14 colunas (A:N)


def _san(report, header, rows, title=None):
    return Sanitized(report=report, header=header, rows=rows, db_rows=rows,
                     source_rows=list(range(len(rows))), raw_count=len(rows), title=title)


def _com_row(i, with_formulas=True):
    """Linha de comissões como o sanitizador entrega: A:K valores, L:N fórmulas de origem
    já reendereçadas para a linha de destino (2 + i)."""
    r = 2 + i
    head = ["M", f"#{i}", D, *([None] * 6), 100.0, 10.0]                             # A:K
    tail = ([Formula(f"=J{r}-K{r}"), Formula(f"=L{r}/J{r}"), Formula(f"=K{r}*2")]   # L:N
            if with_formulas else [90.0, 0.9, 20.0])
    return head + tail


def _results(n_rra=2, n_cp=3, n_cr=2, n_com=2):
    return {
        "rra": _san("rra", ["Núm. JOB", "Data", "VALOR"],
                    [[f"75{i}", D, 100.0 + i] for i in range(n_rra)], title="RENDIMENTOS DOS JOBS"),
        "cp": _san("cp", ["Cód.", "Valor", "Vcto"], [[f"x-{i}", 10.5 + i, D] for i in range(n_cp)]),
        "cr": _san("cr", ["Cliente", "Emissão", "Vl"], [[f"c{i}", D, 5.0] for i in range(n_cr)]),
        "comissoes": _san("comissoes", COM_HEADER, [_com_row(i) for i in range(n_com)]),
    }


def _build(tmp_path, results=None, **tpl):
    tpl_path = make_template(tmp_path / "t.xlsx", **tpl)
    out = tmp_path / "o.xlsx"
    report = build_workbook(tpl_path, out, results or _results())
    return tpl_path, out, report


def test_writes_headers_title_and_data_in_the_right_places(tmp_path):
    _, out, _ = _build(tmp_path)
    wb = openpyxl.load_workbook(out)
    base = wb["BASE"]
    assert base["A1"].value == "RENDIMENTOS DOS JOBS"
    assert [c.value for c in base[3][:3]] == ["Núm. JOB", "Data", "VALOR"]
    assert base["A5"].value == "750" and base["B5"].value == D and base["C6"].value == 101.0
    desp = wb["DESPESAS"]
    assert [c.value for c in desp[2][:3]] == ["Cód.", "Valor", "Vcto"]
    assert desp["A3"].value == "x-0" and desp["C5"].value == D
    assert wb["RECEITA"]["A1"].value == "Cliente" and wb["RECEITA"]["A3"].value == "c1"
    com = wb["REAL X ORÇADO COMISSÕES"]
    assert com["A1"].value == "Modelo" and com["B2"].value == "#0" and com["C3"].value == D


def test_base_total_row_is_two_rows_below_the_last_job(tmp_path):
    _, out, report = _build(tmp_path)
    base = openpyxl.load_workbook(out)["BASE"]
    # 2 jobs: linhas 5 e 6 -> total na 8, somando da linha 4 até a 7
    assert base["A8"].value == "TOTAL"
    assert base["J8"].value == "=SUM(J4:J7)"
    assert base["AU8"].value == "=SUM(AU4:AU7)"
    assert len(SUM_COLUMNS) == 35 and SUM_COLUMNS[0] == "J" and SUM_COLUMNS[-1] == "AU"
    assert report.sheets["BASE"].total_row == 8
    assert report.sheets["BASE"].missing_formulas == []        # TOTAL completo na cobertura


def test_coverage_flags_a_missing_base_total(tmp_path):
    tpl = make_template(tmp_path / "t.xlsx")                    # BASE do modelo: sem TOTAL
    missing = formula_coverage(tpl, "rra", SheetReport("BASE", 5, 6, 2, total_row=8))
    assert missing[0] == "A8"                                   # rótulo ausente
    assert missing[1:] == [f"{c}8" for c in SUM_COLUMNS]        # e as 35 somas


def test_coverage_flags_a_base_total_with_the_wrong_range(tmp_path):
    s = WorkbookSession(make_template(tmp_path / "t.xlsx"))
    ed = s.editor("BASE")
    ed.write(8, 1, "TOTAL")
    for letter in SUM_COLUMNS:
        end = 6 if letter == "J" else 7                         # J soma até a linha errada
        ed.write(8, column_index_from_string(letter), Formula(f"=SUM({letter}4:{letter}{end})"))
    out = s.save(tmp_path / "o.xlsx")
    s.close()
    assert formula_coverage(out, "rra", SheetReport("BASE", 5, 6, 2, total_row=8)) == ["J8"]


def test_formulas_of_the_model_are_extended_beyond_the_template(tmp_path):
    # o modelo tem fórmulas até a linha 7; carregando 8 linhas (3..10) precisa estender 8..10
    _, out, report = _build(tmp_path, _results(n_cp=8), desp_formula_rows=5)
    sess = WorkbookSession(out)
    ed = sess.editor("DESPESAS")
    assert ed.formula_at(10, 42) == '=IF(A10="x","a","b")'
    assert ed.formula_at(10, 43) == '=LEN(AP10)&"!"'
    assert report.sheets["DESPESAS"].missing_formulas == []
    assert report.sheets["DESPESAS"].orphan_formula_rows == 0


def test_orphan_formulas_below_loaded_rows_are_kept_and_counted(tmp_path):
    _, out, report = _build(tmp_path, _results(n_cp=2), desp_formula_rows=5)   # carrega 3..4; modelo vai até 7
    assert report.sheets["DESPESAS"].orphan_formula_rows == 3
    assert WorkbookSession(out).editor("DESPESAS").formula_at(7, 42) is not None


def test_receita_formula_column_is_extended(tmp_path):
    _, out, report = _build(tmp_path, _results(n_cr=6), rec_formula_rows=3)
    assert WorkbookSession(out).editor("RECEITA").formula_at(7, 29) == '=IF(A7="","",A7)'
    assert report.sheets["RECEITA"].missing_formulas == []


def test_comissoes_use_september_formulas_when_the_model_has_none(tmp_path):
    _, out, report = _build(tmp_path)
    ws = openpyxl.load_workbook(out)["REAL X ORÇADO COMISSÕES"]
    assert ws["O3"].value == "=DATE(YEAR(C3),MONTH(C3),1)"
    assert ws["P3"].value.startswith("=IF(COUNTIFS($B$2:$B3,$B3,$H$2:$H3,$H3,$G$2:$G3,$G3,$C$2:$C3,$C3)=1,$J3,0)")
    assert ws["Q2"].value.startswith('=IF($H2="Diretor",$P2/')
    assert report.sheets["REAL X ORÇADO COMISSÕES"].missing_formulas == []


def test_comissoes_source_formulas_in_L_to_N_are_written_and_covered(tmp_path):
    _, out, report = _build(tmp_path)
    ws = openpyxl.load_workbook(out)["REAL X ORÇADO COMISSÕES"]
    assert ws["L3"].value == "=J3-K3" and ws["N2"].value == "=K2*2"
    assert report.sheets["REAL X ORÇADO COMISSÕES"].missing_formulas == []


def test_comissoes_row_without_formulas_in_L_to_N_is_a_coverage_gap(tmp_path):
    res = _results()
    res["comissoes"].rows[1] = _com_row(1, with_formulas=False)     # linha 3 chegou só com valores
    _, _, report = _build(tmp_path, res)
    # não levanta: a lacuna vai para o relatório; no Plano 3 a conferência fica 'bloqueada' (código 2)
    assert report.sheets["REAL X ORÇADO COMISSÕES"].missing_formulas == ["L3", "M3", "N3"]


def test_missing_model_formulas_block_the_publication(tmp_path):
    # classificação ausente: o modelo não tem AP:AQ para estender. No Plano 3 este ContractError
    # vira importação 'bloqueada' com o motivo, nenhuma linha gravada, código de saída 2.
    tpl = make_template(tmp_path / "t.xlsx", desp_formula_rows=0)
    with pytest.raises(ContractError, match="classificação ausente"):
        build_workbook(tpl, tmp_path / "o.xlsx", _results())


def test_dirty_template_is_rejected_and_nothing_is_written(tmp_path):
    tpl = make_template(tmp_path / "t.xlsx")
    s = WorkbookSession(tpl)
    s.editor("DESPESAS").write(3, 1, "resto de outra rodada")
    dirty = s.save(tmp_path / "dirty.xlsx")
    with pytest.raises(ContractError, match="A3"):
        build_workbook(dirty, tmp_path / "o.xlsx", _results())
    assert not (tmp_path / "o.xlsx").exists()


def test_only_the_four_base_sheets_change(tmp_path):
    tpl, out, report = _build(tmp_path)
    changed = set(report.parts["changed"])
    assert changed == {"xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml",
                       "xl/worksheets/sheet3.xml", "xl/worksheets/sheet4.xml",
                       "xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"}
    assert "xl/worksheets/sheet5.xml" not in changed          # aba "OUTRA" intacta
    assert report.parts["removed"] == ["xl/calcChain.xml"]


def test_date_in_unformatted_cells_is_reported(tmp_path):
    _, _, report = _build(tmp_path)
    assert report.sheets["RECEITA"].date_style_warnings == {"B": 2}   # RECEITA do modelo sintético não tem formato de data


def test_zero_rows_still_builds(tmp_path):
    res = _results(n_rra=0, n_cp=0, n_cr=0, n_com=0)
    _, out, report = _build(tmp_path, res)
    assert report.sheets["DESPESAS"].rows == 0
    assert openpyxl.load_workbook(out)["BASE"]["A6"].value == "TOTAL"   # sem jobs: 4 + 2
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_build.py -v`
Expected: FAIL com `ModuleNotFoundError: No module named 'papaki_dre.build'`

- [ ] **Step 3: Implementar `build.py`**

```python
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter

from .config import DEST
from .types import ContractError, Formula, Sanitized
from .xlsx_inject import WorkbookSession, changed_parts

# Colunas somadas na linha TOTAL da BASE (iguais às do modelo de setembro/2026).
SUM_COLUMNS = ("J K L N P Q R S T U V W X Y Z AA AB AC AD AE AF AG AH AI AJ AK AL AM AN AO "
               "AP AQ AR AS AU").split()

# Colunas do modelo que são fórmulas e apenas se estendem às linhas carregadas.
FORMULA_COLUMNS = {"cp": ["AP", "AQ"], "cr": ["AC"], "comissoes": ["O", "P", "Q"]}

# Colunas conferidas na cobertura (spec §8): as estendidas acima e, nas comissões, também as
# fórmulas de origem L:N (reendereçadas pelo sanitizador). A linha TOTAL da BASE é conferida à parte.
COVERAGE_COLUMNS = {"cp": ["AP", "AQ"], "cr": ["AC"], "comissoes": ["L", "M", "N", "O", "P", "Q"]}

# O:Q das comissões: o modelo-base não as traz; origem = linha 2 do modelo de setembro/2026.
COMISSOES_FORMULAS = {
    "O": "=DATE(YEAR(C2),MONTH(C2),1)",
    "P": "=IF(COUNTIFS($B$2:$B2,$B2,$H$2:$H2,$H2,$G$2:$G2,$G2,$C$2:$C2,$C2)=1,$J2,0)",
    "Q": '=IF($H2="Diretor",$P2/(LEN($D2)-LEN(SUBSTITUTE($D2,",",""))+1),0)',
}
ORDER = ("rra", "cp", "cr", "comissoes")


@dataclass
class SheetReport:
    sheet: str
    first_row: int
    last_row: int
    rows: int
    missing_formulas: list[str] = field(default_factory=list)   # lacunas de cobertura (vazia = ok)
    orphan_formula_rows: int = 0
    date_style_warnings: dict[str, int] = field(default_factory=dict)
    total_row: int | None = None                                 # só na BASE: linha do TOTAL recriado


@dataclass
class BuildReport:
    sheets: dict[str, SheetReport] = field(default_factory=dict)
    parts: dict[str, list[str]] = field(default_factory=dict)


def _translate(formula: str, origin: str, dest: str) -> Formula:
    return Formula(Translator(formula, origin=origin).translate_formula(dest))


def _check_empty(ed, sheet: str, rows, ncols: int) -> None:
    dirty = ed.nonempty_cells(rows, set(range(1, ncols + 1)))
    if dirty:
        raise ContractError(
            f"{sheet}: o modelo não está vazio na região de destino "
            f"({len(dirty)} células, ex.: {', '.join(dirty[:5])})")


def _load(session: WorkbookSession, key: str, data: Sanitized) -> SheetReport:
    lay = DEST[key]
    sheet, first = lay["sheet"], lay["first_row"]
    ed = session.editor(sheet)
    ncols, n = len(data.header), len(data.rows)
    last = first + n - 1
    top = lay.get("title_row", lay["header_row"])
    bottom = last + 2 if key == "rra" else max(last, first)
    _check_empty(ed, sheet, range(top, bottom + 1), ncols)

    ds = session.date_styles
    if key == "rra" and data.title is not None:
        ed.write(lay["title_row"], 1, data.title)
    for j, name in enumerate(data.header, start=1):
        ed.write(lay["header_row"], j, name, style_from=None)
    for i, row in enumerate(data.rows):
        r = first + i
        for j, v in enumerate(row, start=1):
            ed.write(r, j, v, style_from=(first, j), date_styles=ds)

    report = SheetReport(sheet, first, last, n)
    for letter in FORMULA_COLUMNS.get(key, []):
        _extend_formulas(ed, key, sheet, letter, first, last, report)
    if key == "rra":
        total_row = last + 2
        ed.write(total_row, 1, "TOTAL")
        for letter in SUM_COLUMNS:
            ed.write(total_row, column_index_from_string(letter),
                     Formula(f"=SUM({letter}4:{letter}{last + 1})"))
        report.total_row = total_row
    report.date_style_warnings = dict(ed.date_warnings)
    return report


def _extend_formulas(ed, key, sheet, letter, first, last, report) -> None:
    col = column_index_from_string(letter)
    last_f = ed.last_formula_row(col)
    if last_f is None:
        if key == "comissoes":                       # o modelo não traz O:Q: usar as de setembro
            for r in range(first, last + 1):
                ed.write(r, col, _translate(COMISSOES_FORMULAS[letter], f"{letter}2", f"{letter}{r}"))
            return
        raise ContractError(
            f"{sheet}!{letter}: classificação ausente — o modelo não tem fórmula para estender")
    base = ed.formula_at(last_f, col)
    for r in range(max(first, last_f + 1), last + 1):
        ed.write(r, col, _translate(base, f"{letter}{last_f}", f"{letter}{r}"),
                 style_from=(last_f, col))
    if last_f > last:
        report.orphan_formula_rows = max(report.orphan_formula_rows, last_f - last)


def _total_gaps(ed, sr: SheetReport) -> list[str]:
    """Linha TOTAL da BASE: rótulo em A e =SUM(X4:X<última+1>) exata em cada coluna de SUM_COLUMNS."""
    t = sr.total_row
    if t is None:
        return ["BASE: linha TOTAL não registrada"]
    gaps = [] if ed.text_at(t, 1) == "TOTAL" else [f"A{t}"]
    for letter in SUM_COLUMNS:
        expected = f"=SUM({letter}4:{letter}{sr.last_row + 1})"
        if ed.formula_at(t, column_index_from_string(letter)) != expected:
            gaps.append(f"{letter}{t}")
    return gaps


def formula_coverage(path, key: str, sr: SheetReport) -> list[str]:
    """Lê o arquivo gravado e devolve as células sem a fórmula exigida (spec §8): COVERAGE_COLUMNS
    em todas as linhas carregadas e, na BASE, a linha TOTAL. Lista vazia = cobertura completa."""
    session = WorkbookSession(path)
    try:
        ed = session.editor(sr.sheet)
        missing = []
        for letter in COVERAGE_COLUMNS.get(key, []):
            col = column_index_from_string(letter)
            for r in range(sr.first_row, sr.last_row + 1):
                if ed.formula_at(r, col) is None:
                    missing.append(f"{letter}{r}")
        if key == "rra":
            missing.extend(_total_gaps(ed, sr))
        return missing
    finally:
        session.close()


def build_workbook(template, out_path, results: dict[str, Sanitized]) -> BuildReport:
    """Copia o modelo e carrega as quatro bases. Nenhuma outra aba é tocada."""
    missing = [k for k in ORDER if k not in results]
    if missing:
        raise ContractError(f"faltam resultados para: {', '.join(missing)}")
    out_path = Path(out_path)
    report = BuildReport()
    session = WorkbookSession(template)
    try:
        for key in ORDER:
            report.sheets[DEST[key]["sheet"]] = _load(session, key, results[key])
        session.save(out_path)
    finally:
        session.close()
    for key in ORDER:                      # lacuna de cobertura não levanta: a conferência (Plano 3) bloqueia
        sr = report.sheets[DEST[key]["sheet"]]
        sr.missing_formulas = formula_coverage(out_path, key, sr)
    report.parts = changed_parts(template, out_path)
    return report
```

- [ ] **Step 4: Rodar e ver passar**

Run: `py -m pytest tests/test_build.py -v`
Expected: 15 passed

`test_comissoes_row_without_formulas_in_L_to_N_is_a_coverage_gap`, `test_coverage_flags_a_missing_base_total` e `test_coverage_flags_a_base_total_with_the_wrong_range` falham se `formula_coverage` deixar de conferir `L:N` ou a linha `TOTAL` (esse é o objetivo deles). Nunca relaxe a cobertura para um teste passar.

Se `test_date_in_unformatted_cells_is_reported` falhar, confira o estilo das células de `RECEITA` na fábrica (sem `s`) e a coluna do teste: o teste espera as 2 datas da coluna `B` (`Emissão`).

- [ ] **Step 5: Commit**

```bash
git add src/papaki_dre/build.py tests/test_build.py && git commit -m "feat: carga das quatro bases no modelo com extensão de fórmulas, total da BASE e cobertura"
```

---

### Task 5: Integração com o modelo real e os brutos de setembro

**Files:**
- Test: `tests/test_build_real.py`

Este teste usa o modelo real (`Modelo DRE - Untitled - base.xlsx`), roda os quatro sanitizadores nos brutos de 08/10 e grava a cópia. Não escreve nada no modelo original.

- [ ] **Step 1: Escrever o teste**

`tests/test_build_real.py`:

```python
import openpyxl
import pytest

from papaki_dre.build import build_workbook
from papaki_dre.sanitize import sanitize_comissoes, sanitize_cp, sanitize_cr, sanitize_rra
from papaki_dre.xlsx_inject import WorkbookSession
from conftest import BRUTO, REAL_FILES, TEMPLATE

pytestmark = [pytest.mark.real, pytest.mark.excel]


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    needed = [TEMPLATE, *REAL_FILES.values()]
    if not all(p.exists() for p in needed):
        pytest.skip("modelo ou brutos de referência ausentes")
    work = tmp_path_factory.mktemp("real")
    results = {
        "cp": sanitize_cp(REAL_FILES["cp"]),
        "cr": sanitize_cr(REAL_FILES["cr"]),
        "comissoes": sanitize_comissoes(REAL_FILES["comissoes"]),
        "rra": sanitize_rra(REAL_FILES["rra"], work_dir=work),
    }
    out = work / "Modelo DRE - Untitled - 09_2026.xlsx"
    report = build_workbook(TEMPLATE, out, results)
    return results, out, report


def test_counts_match_the_sanitizers(built):
    results, _, report = built
    assert report.sheets["BASE"].rows == results["rra"].kept == 72
    assert report.sheets["DESPESAS"].rows == results["cp"].kept == 1110
    assert report.sheets["RECEITA"].rows == results["cr"].kept == 484
    assert report.sheets["REAL X ORÇADO COMISSÕES"].rows == results["comissoes"].kept == 649


def test_only_base_sheets_and_workbook_plumbing_changed(built):
    _, out, report = built
    sess = WorkbookSession(out)
    names = {sess.parts[n] for n in ("BASE", "DESPESAS", "RECEITA", "REAL X ORÇADO COMISSÕES")}
    sess.close()
    assert set(report.parts["changed"]) == names | {
        "xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"}
    assert report.parts["removed"] == ["xl/calcChain.xml"]
    assert report.parts["added"] == []


def test_every_loaded_row_has_the_model_formulas(built):
    # cobertura completa: AP:AQ, AC, L:N + O:Q em todas as linhas carregadas e o TOTAL da BASE
    _, _, report = built
    for name, sr in report.sheets.items():
        assert sr.missing_formulas == [], (name, sr.missing_formulas[:5])
    assert report.sheets["BASE"].total_row == 78


def test_orphan_formulas_are_counted_not_removed(built):
    _, _, report = built
    assert report.sheets["DESPESAS"].orphan_formula_rows == 1159 - (2 + 1110)
    assert report.sheets["RECEITA"].orphan_formula_rows == 664 - (1 + 484)


def test_output_opens_with_openpyxl_and_has_expected_cells(built):
    results, out, _ = built
    wb = openpyxl.load_workbook(out)
    assert wb["DESPESAS"]["A3"].value == results["cp"].rows[0][0]
    assert wb["BASE"]["A3"].value == "Núm. JOB"
    assert wb["BASE"]["A78"].value == "TOTAL"                 # 5 + 72 - 1 + 2
    assert wb["BASE"]["J78"].value == "=SUM(J4:J77)"
    assert wb["REAL X ORÇADO COMISSÕES"]["A1"].value == "Modelo de Orçamento"


def test_pivot_tables_charts_and_drawings_survive(built):
    _, out, _ = built
    import zipfile
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
    assert sum(n.startswith("xl/pivotTables/pivotTable") for n in names) >= 24
    assert any(n.startswith("xl/charts/chart") for n in names)


def test_excel_opens_the_output(built):
    results, out, _ = built
    import win32com.client
    from papaki_dre.readers import com_retry
    excel = com_retry(lambda: win32com.client.DispatchEx("Excel.Application"))
    com_retry(lambda: setattr(excel, "Visible", False))
    com_retry(lambda: setattr(excel, "DisplayAlerts", False))
    try:
        wb = com_retry(lambda: excel.Workbooks.Open(str(out), ReadOnly=True))
        try:
            ws = com_retry(lambda: wb.Worksheets("DESPESAS"))
            assert com_retry(lambda: ws.Range("A3").Value) == results["cp"].rows[0][0]
            assert str(com_retry(lambda: ws.Range("AP3").Formula)).startswith("=")
            assert com_retry(lambda: wb.Worksheets.Count) == 17
        finally:
            com_retry(lambda: wb.Close(False))
    finally:
        com_retry(lambda: excel.Quit())
```

- [ ] **Step 2: Rodar**

Run: `py -m pytest tests/test_build_real.py -v`
Expected: 7 passed. A carga e a abertura no Excel levam de 30 s a 2 min.

Se `test_every_loaded_row_has_the_model_formulas` falhar para `REAL X ORÇADO COMISSÕES`: em `O:Q`, confira se o modelo real realmente não tem essas fórmulas (a aba tem `dim=C2:O1160` e 0 fórmulas, então deve usar as constantes de setembro); em `L:N`, as fórmulas vêm do bruto via sanitizador (em 09/10/2026 as 649 linhas mantidas tinham `Formula` em `L`, `M` e `N`) — **PARE e investigue**, não relaxe a cobertura. Uma lacuna real é exatamente o que o Plano 3 transforma em importação `bloqueada` (código 2).

- [ ] **Step 3: Conferência manual (obrigatória, não automatizável)**

Abra o arquivo gerado em `<tmp>/Modelo DRE - Untitled - 09_2026.xlsx` (o caminho aparece com `pytest --basetemp=out/pytest`; rode `py -m pytest tests/test_build_real.py --basetemp=out/pytest` e abra `out/pytest/real0/...`) **no Excel normal** e confirme:
1. O Excel **não** exibe aviso de reparo ao abrir.
2. As abas `DRE - Untitled`, `JOB's Detalhado` etc. abrem e recalculam sem erro de estrutura.
3. As tabelas dinâmicas continuam presentes (a atualização delas é do FP&A).

Anote o resultado no relatório da execução. Um aviso de reparo é **bloqueante**.

- [ ] **Step 4: Commit**

```bash
git add tests/test_build_real.py && git commit -m "test: integração da carga com o modelo real e os brutos de setembro"
```

---

## Self-Review

**Cobertura da spec §7/§8 (carga):** cópia do modelo e escrita só nas quatro abas (T4); método cirúrgico, demais partes idênticas, `calcChain` removido, `fullCalcOnLoad` (T2–T3, T5); intervalos `A5:AW`/`A3:AO`/`A2:AB`/`A1:N` e cabeçalhos (T4); extensão de `AP:AQ`/`AC` e `O:Q` a todas as linhas carregadas (T4); `L:N` das comissões gravadas como chegam do sanitizador, já reendereçadas (T4); linha `TOTAL` da `BASE` (T4); **cobertura da spec §8 completa** — `AP:AQ`, `AC`, `L:N`, `O:Q` em cada linha carregada e a linha `TOTAL` da `BASE` (rótulo + 35 somas exatas), lida do arquivo gravado por `formula_coverage` e devolvida em `SheetReport.missing_formulas`, com testes que falham quando falta fórmula em `L:N` e quando falta ou está errado o `TOTAL` (T4, T5); região vazia (T4); fórmulas órfãs contadas, não removidas (T4); classificação ausente bloqueia (T4): `ContractError` quando o modelo não tem fórmula para estender, lacuna em `missing_formulas` quando uma linha carregada fica sem fórmula — no fluxo completo, ambos viram importação `bloqueada` no Plano 3 (código de saída 2); fórmulas de abas fora da base **não** são conferidas (por desenho); avisos de data em célula sem formato (T2, T4); tipo da célula preservado (T2).

**Fora do MVP (decisão de 09/10/2026):** qualquer critério de classificação além de "ausente"; conferência por valores dos rodapés `TOTAL` dos brutos (são só técnicos).

**Interface entregue ao Plano 3:** `build_workbook(template, out_path, results, *, master_formula_rows=None) -> BuildReport` (`None` = `config.MASTER_FORMULA_ROWS`); `BuildReport(sheets, parts)`; `SheetReport(sheet, first_row, last_row, rows, missing_formulas, orphan_formula_rows, date_style_warnings, total_row=None, overwritten_constants=0, autofilter_extended=False, unified_formula_rows=0)` — os quatro últimos campos são novos e têm padrão, então os construtores posicionais do Plano 3 continuam válidos; `missing_formulas` inclui `L:N`, o `TOTAL` da `BASE` e, desde 13/10, `<f/>` vazio e filha compartilhada sem mestre; `formula_coverage(path, key, sr, session=None)` (o `session` opcional reaproveita uma sessão aberta). O Plano 3 exibe `overwritten_constants`, `unified_formula_rows` e `autofilter_extended` no relatório de conferência e **não** bloqueia por eles.

**Nota de execução (13/10/2026):** (1) **Decisão "opção a":** constantes digitadas em `DESPESAS!AP:AQ` (439 no modelo de setembro) e em `RECEITA!AC` dentro das linhas carregadas são substituídas pela fórmula do modelo e contadas em `SheetReport.overwritten_constants`. (2) A base de extensão é a última fórmula **da mesma família** das linhas de dados (rodapés/resumos como `=COUNTA(AP3:AP39)` são ignorados). (3) O `autoFilter` da aba (e o nome oculto `_xlnm._FilterDatabase` correspondente) é estendido até a última linha carregada quando ela passa do fim do filtro, nunca encolhido (`SheetReport.autofilter_extended`); no modelo de setembro nenhum filtro precisa crescer — isto substitui o risco (2) abaixo. (4) A linha `TOTAL` da `BASE` copia o estilo da última linha de dados; o início das somas é `BASE_SUM_START_ROW = 4`. (5) `#N/D` em `AP:AQ` (classificação sem correspondência no De-Para; 327 células no arquivo de setembro, incluindo linhas órfãs) **não** bloqueia no MVP (decisão 8: só a ausente), mas deve aparecer como alerta no relatório do Plano 3. (6) **Revisão final (13/10/2026):** dentro das linhas carregadas, fórmula de **outra família** (`<f/>` vazio, ilegível, resumo que sai da planilha ao ser levado para a 1ª linha, resumo só com referências absolutas) também é substituída pela fórmula de dados e contada em `overwritten_constants` (que passa a significar "constantes e fórmulas de outra família substituídas"); variante com referências relativas fica intacta (o modelo de setembro tem uma em `DESPESAS!AP881:AP1070`; **superado pelo item 7 para `AP`**); mestre compartilhado de outra família e células de fórmula de matriz → `ContractError`. `write` trata texto vazio após a limpeza de caracteres inválidos como `None`; `_sync_filter_names` ignora nomes com várias áreas e aceita prefixo de namespace. Saída do build real idêntica byte a byte (sha256) à anterior. (7) **Decisão do usuário de 13/10/2026 — fórmula completa em toda a `DESPESAS!AP`:** substitui a "variante de `AP881:AP1070` mantida / decisão do FP&A". A referência é a linha 881 do modelo-base de setembro da Untitled (a classificação com o ramo `"Transferência entre Empresas Coligadas"` → `"Desconsiderar"`), configurada em `config.MASTER_FORMULA_ROWS = {"cp": {"AP": 881}}` (por produtora no Plano 5). `_unify_formulas` (build) grava essa fórmula, traduzida, em todas as linhas de `first` até max(`last`, última fórmula de dados da coluna) — no real `AP3:AP1159`, inclusive as órfãs —, deixando intactas as já idênticas; o mestre de um grupo compartilhado só é trocado com o `ref` inteiro (e todas as filhas do `si`) no mesmo lote (`SheetEditor.write_batch`/`check_batch`, validação antes de gravar), senão `ContractError`; idem para linha configurada sem fórmula, ilegível ou que não se traduz para `first`. Nova contagem `SheetReport.unified_formula_rows` (fórmulas de dados existentes trocadas; 846 no build real), separada de `overwritten_constants` (439, inalterado); o Plano 3 a exibe sem bloquear. `AQ`, `AC` e `O:Q` não mudam. Verificado: 0 filhas compartilhadas órfãs em DESPESAS; via COM, 0 erros em `AP3:AP1112` e 231 `#N/D` em `AQ3:AQ1112` (igual a antes).

**Fora deste plano:** relatório de conferência consolidado e CLI (Plano 3), banco (Plano 3).

**Riscos assumidos explicitamente:** (1) tabelas dinâmicas ficam com cache antigo: atualizar é do FP&A; (2) ~~o `autoFilter` de `DESPESAS`/`RECEITA` mantém o intervalo do modelo~~ (resolvido em 13/10: é estendido, nunca encolhido); (3) a abertura no Excel sem aviso de reparo só se comprova manualmente (T5, Step 3), porque `DisplayAlerts=False` pode reparar em silêncio.

**Consistência de nomes:** `DateStyles`, `SheetEditor` (`write`, `formula_at`, `text_at`, `last_formula_row`, `formula_family`, `is_foreign_formula`, `array_formula_at`, `has_formula_element`, `nonempty_cells`, `style_of`, `row_exists`, `date_warnings`, `extend_autofilter`, `_inherit_style`, `_count_unformatted_date`, `_prepare`), `WorkbookSession` (`editor`, `save`, `close`, `parts`, `date_styles`), `FormulaShifter`, `translate_formula`, `MAX_EXACT_INT`, `changed_parts`, `to_serial` (T1–T3, `xlsx_inject.py`); `build_workbook`, `formula_coverage`, `BuildReport`, `SheetReport` (com `total_row`), `SUM_COLUMNS`, `FORMULA_COLUMNS`, `COVERAGE_COLUMNS`, `COMISSOES_FORMULAS`, `BASE_SUM_START_ROW`, `_has_formula` (T4, `build.py`). Do Plano 1, só nomes do código final: `Sanitized`, `Formula`, `XlError`, `ContractError` (`types.py`), `DEST` (`config.py`), `sanitize_cp/cr/comissoes`, `sanitize_rra(xls_path, *, work_dir)`.
