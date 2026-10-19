# Plano 1 — Fundação e Sanitizadores Implementation Plan

> **Status:** Executado em 09/10/2026; este documento descreve o código como construído (commits listados). Os blocos de código e de teste abaixo são cópias fiéis dos arquivos finais (após a rodada de correções A–G); a suíte final tem 138 testes.

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Criar o pacote Python `papaki_dre` com leitores de arquivo e os quatro sanitizadores (CP, CR, RRA, Comissões) que aplicam as regras do `instrucoes_relatorios_ base.md`, com testes sintéticos e testes contra os brutos reais de setembro/2026.

**Architecture:** Funções puras: entra o caminho do arquivo bruto, sai um objeto `Sanitized` (linhas tratadas para o Excel em `rows`, linhas só de valores para o banco em `db_rows` — listas distintas —, linha de origem de cada linha mantida, linhas técnicas e exclusões de negócio com motivo, e avisos). Linhas de TOTAL dos brutos são reconhecidas por rótulo e posição exatos e apenas descartadas como técnicas (`R_FOOTER`); seus valores nunca são lidos. Estrutura inválida (cabeçalho divergente ou com coluna extra preenchida, registro de CSV com nº de colunas inesperado, valor ilegível, erro inesperado do Excel em linha mantida) levanta `ContractError` e para a rodada. Diferença de contagem entre tela e arquivo nunca é erro. O RRA (XLS binário com 11 mil fórmulas) é convertido para XLSX pelo próprio Excel (COM) para preservar fórmulas; o resultado fica num arquivo com o hash do conteúdo no nome, gravado de forma atômica. Comissões e RRA compartilham a cópia de linha (`sanitize/_xlsx.py`).

**Tech Stack:** Python ≥ 3.11, openpyxl, lxml, pywin32 (Excel COM), pytest.

**Spec:** `docs/superpowers/specs/2026-10-09-papaki-esferas-mvp-design.md` (§6 sanitização; decisões de 09/10/2026 em §10).

**Planos seguintes (independentes, cada um depende deste):** Plano 2 — carga no Excel; Plano 3 — banco (Supabase) e comando `importar`; Plano 4 — automação do Chrome (runbook).

**Convenções:** comandos rodam na raiz `<pasta do projeto>` com bash (Git Bash). Mensagens de commit seguem o padrão `feat:`/`test:`/`chore:`/`fix:`/`perf:`. Ao commitar, a execução acrescenta a linha de co-autoria exigida pelo ambiente.

---

## File Structure

```
pyproject.toml
.gitignore
scripts/capture_headers.py            # gera src/papaki_dre/headers.json a partir dos brutos de referência
src/papaki_dre/__init__.py
src/papaki_dre/config.py              # constantes de negócio, tipos por coluna, motivos de remoção, rótulos de rodapé
src/papaki_dre/types.py               # Formula, XlError, ContractError, Removal, Sanitized
src/papaki_dre/coerce.py              # parse_money/parse_date/parse_datetime/coerce_cell, MONEY_TEXT_OK
src/papaki_dre/formulas.py            # readdress() — reendereça referências relativas
src/papaki_dre/readers.py             # read_delimited, open_xlsx (repara autoFilter), convert_xls_to_xlsx (Excel COM, cache por hash)
src/papaki_dre/headers.py             # normalize(), load_headers(), check_header()
src/papaki_dre/headers.json           # contrato de cabeçalhos (gerado, versionado)
src/papaki_dre/sanitize/__init__.py
src/papaki_dre/sanitize/_xlsx.py      # partes comuns de Comissões e RRA: RowCopier, blank, check_header_extent
src/papaki_dre/sanitize/cp.py
src/papaki_dre/sanitize/cr.py
src/papaki_dre/sanitize/comissoes.py
src/papaki_dre/sanitize/rra.py
tests/conftest.py                     # fixtures dos brutos reais (skip se ausentes)
tests/xlsxcache.py                    # helper: injeta valores em cache (<v>) em fórmulas de um XLSX do openpyxl
tests/test_types.py
tests/test_coerce.py
tests/test_formulas.py
tests/test_readers.py
tests/test_headers.py
tests/test_sanitize_cp.py
tests/test_sanitize_cr.py
tests/test_sanitize_comissoes.py
tests/test_sanitize_rra.py
```

Os brutos de referência (somente leitura, não vão para o git) estão em `amostras/2026-09/bruto/`:
`CP_papaki_consolidado_2026-01-01_a_2026-09-30.csv`, `CR_untitled_UTLD_2026-01-01_a_2026-09-30.csv`, `COMISSOES_untitled_UTLD_2026-01-01_a_2026-09-30.xlsx`, `Relatório Rentabilidade.xls`.

---

### Task 1: Scaffolding do projeto

Commit: 8e23478

**Files:**
- Create: `pyproject.toml`, `.gitignore`, `src/papaki_dre/__init__.py`, `src/papaki_dre/sanitize/__init__.py`, `tests/conftest.py`

- [x] **Step 1: Inicializar git e criar o `.gitignore`**

```bash
cd "<pasta do projeto>" && git init
```

Create `.gitignore`:

```gitignore
# dados reais da Papaki: nunca versionar (CPF/CNPJ, dados bancários)
amostras/
out/
*.xls
*.xlsx
*.csv
.env
__pycache__/
*.egg-info/
.pytest_cache/
.venv/
```

Observação: o modelo `Modelo DRE - Untitled - base.xlsx` fica fora do git por causa de `*.xlsx` (é arquivo da Papaki). Os testes que dependem dele são ignorados (skip) se ele não existir.

- [x] **Step 2: Criar `pyproject.toml`**

```toml
[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[project]
name = "papaki-dre"
version = "0.1.0"
description = "Importação das bases do Esferas para o Modelo DRE da Papaki"
requires-python = ">=3.11"
dependencies = [
  "openpyxl>=3.1",
  "lxml>=5",
  "pywin32>=306; sys_platform == 'win32'",
  "supabase>=2.0",
  "python-dotenv>=1.0",
]

[project.optional-dependencies]
dev = ["pytest>=8"]

[tool.setuptools.packages.find]
where = ["src"]

[tool.setuptools.package-data]
papaki_dre = ["headers.json"]

[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["src"]
markers = [
  "excel: exige Microsoft Excel instalado (COM)",
  "real: usa os brutos reais de amostras/2026-09 (ignorado se ausentes)",
]
```

- [x] **Step 3: Criar os pacotes vazios e o `conftest.py`**

`src/papaki_dre/__init__.py`:

```python
"""Importação das bases do Esferas para o Modelo DRE da Papaki."""
__version__ = "0.1.0"
```

`src/papaki_dre/sanitize/__init__.py` (vazio por enquanto; a Task 10 define os exports):

```python
"""Sanitizadores dos quatro relatórios do Esferas."""
```

`tests/conftest.py`:

```python
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
BRUTO = ROOT / "amostras" / "2026-09" / "bruto"
TEMPLATE = ROOT / "Modelo DRE - Untitled - base.xlsx"

REAL_FILES = {
    "cp": BRUTO / "CP_papaki_consolidado_2026-01-01_a_2026-09-30.csv",
    "cr": BRUTO / "CR_untitled_UTLD_2026-01-01_a_2026-09-30.csv",
    "comissoes": BRUTO / "COMISSOES_untitled_UTLD_2026-01-01_a_2026-09-30.xlsx",
    "rra": BRUTO / "Relatório Rentabilidade.xls",
}


def _real(name):
    path = REAL_FILES[name]
    if not path.exists():
        pytest.skip(f"bruto de referência ausente: {path}")
    return path


@pytest.fixture
def cp_real():
    return _real("cp")


@pytest.fixture
def cr_real():
    return _real("cr")


@pytest.fixture
def comissoes_real():
    return _real("comissoes")


@pytest.fixture
def rra_real():
    return _real("rra")


@pytest.fixture
def template_real():
    if not TEMPLATE.exists():
        pytest.skip(f"modelo ausente: {TEMPLATE}")
    return TEMPLATE
```

- [x] **Step 4: Instalar e verificar que o pytest roda**

Run: `python -m pip install -e ".[dev]" && python -m pytest -q`
Expected: instala sem erro; pytest termina com `no tests ran` (exit code 5 é aceitável aqui).

- [x] **Step 5: Commit**

```bash
git add pyproject.toml .gitignore src tests && git commit -m "chore: scaffolding do pacote papaki_dre"
```

---

### Task 2: Tipos compartilhados e configuração

Commit: 7a496e9 (ajustado depois em 9984a05 — remoção dos totais de controle e rótulos de rodapé — e fb97ad2 — remoção de `R_ODD`)

**Files:**
- Create: `src/papaki_dre/types.py`, `src/papaki_dre/config.py`
- Test: `tests/test_types.py`

- [x] **Step 1: Escrever o teste que falha**

`tests/test_types.py`:

```python
import dataclasses

from papaki_dre.types import Formula, Removal, Sanitized, XlError


def test_sanitized_has_no_control_totals():
    names = {f.name for f in dataclasses.fields(Sanitized)}
    assert "controls" not in names


def test_sanitized_invariant_counts():
    s = Sanitized(
        report="cp", header=["a"], rows=[["x"]], db_rows=[["x"]], source_rows=[2],
        raw_count=3, technical=[Removal(3, "linha_vazia")], exclusions=[Removal(4, "x")],
    )
    assert s.kept == 1
    assert s.balanced()


def test_formula_and_error_are_str_subclasses():
    assert isinstance(Formula("=A1"), str)
    assert isinstance(XlError("#DIV/0!"), str)
```

- [x] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_types.py -v`
Expected: FAIL com `ModuleNotFoundError: No module named 'papaki_dre.types'`

- [x] **Step 3: Implementar `types.py`**

`Sanitized` guarda linhas mantidas (`rows`/`db_rows`/`source_rows`), linhas técnicas, exclusões e avisos; `balanced()` confere `raw_count = mantidas + técnicas + exclusões`.

```python
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


class ContractError(Exception):
    """Violação estrutural (cabeçalho, nº de colunas, erro inesperado). Para a rodada."""


class Formula(str):
    """Texto de fórmula do Excel, com o '=' inicial."""


class XlError(str):
    """Valor de erro literal do Excel (ex.: '#DIV/0!')."""


@dataclass(frozen=True)
class Removal:
    """Uma linha do bruto que não chegou ao resultado."""
    source_row: int  # nº do registro no bruto (o cabeçalho conta como 1 em CSV; nº da linha em planilhas)
    reason: str


@dataclass
class Sanitized:
    report: str
    header: list[str]
    rows: list[list[Any]]          # valores tipados; fórmulas como Formula (para o Excel)
    db_rows: list[list[Any]]       # somente valores (para o banco)
    source_rows: list[int]         # linha de origem de cada linha mantida
    raw_count: int                 # registros do bruto (sem o cabeçalho)
    technical: list[Removal] = field(default_factory=list)   # linhas vazias e rodapés (TOTAL)
    exclusions: list[Removal] = field(default_factory=list)  # exclusões de negócio
    warnings: list[str] = field(default_factory=list)
    title: Any = None
    encoding: str | None = None

    @property
    def kept(self) -> int:
        return len(self.rows)

    def balanced(self) -> bool:
        return self.raw_count == self.kept + len(self.technical) + len(self.exclusions)
```

- [x] **Step 4: Implementar `config.py`**

Os rótulos de rodapé (`CP_FOOTER`, `CR_FOOTER`, `RRA_FOOTER`, `COMISSOES_FOOTER`) só servem para reconhecer a linha de TOTAL e descartá-la como técnica (`R_FOOTER`).

```python
"""Constantes de negócio. Fonte: instrucoes_relatorios_ base.md."""

UTLD = "UTLD PRODUÇÃO DE FILMES LTDA"          # nunca "UTLD PRODUCAO"
PESQUISA = "PESQUISA E TRATAMENTO"
NUCLEO_RRA = "Projetos - Untitled"
MODELO_COMISSOES = "Untitled - Publicidade Padrão"
COMISSOES_EXCLUIR_H = (
    "Seguro Produção", "Taxa da produtora", "Comissão total 1",
    "Comissão total 2", "Impostos", "Taxa ADM",
)

# Primeira linha de dados e linha de cabeçalho no DESTINO (modelo Excel).
DEST = {
    "rra": {"sheet": "BASE", "title_row": 1, "header_row": 3, "first_row": 5},
    "cp": {"sheet": "DESPESAS", "header_row": 2, "first_row": 3},
    "cr": {"sheet": "RECEITA", "header_row": 1, "first_row": 2},
    "comissoes": {"sheet": "REAL X ORÇADO COMISSÕES", "header_row": 1, "first_row": 2},
}

# Motivos de remoção (usados em `exclusoes` no banco e no relatório).
R_BLANK = "linha_vazia"
R_FOOTER = "rodape_total"
R_JOB = "job_preenchido_sem_pesquisa"
R_LOC = "fora_localidade_untitled"
R_NUCLEO = "nucleo_diferente_de_untitled"
R_PESQ = "pesquisa_e_tratamento"
R_MODELO = "modelo_diferente"
R_H = "comissao_excluida_coluna_h"

# Tipo de cada coluna, por nome de cabeçalho. O que não estiver aqui é texto.
# text | date | datetime | money (estrito) | money_or_text
CP_KINDS = {
    **{n: "date" for n in (
        "Vcto", "Vcto alterado", "Data prevista", "Data entrada", "Emissão",
        "Competência", "Compensado", "Data documento", "Comp.")},
    "Criado em": "datetime",
    "Valor": "money",
    "Deduções": "money_or_text",
    "Alterado": "money_or_text",
}
CR_KINDS = {
    **{n: "date" for n in ("Emissão", "Vcto", "Competência")},
    **{n: "money" for n in ("Vl Documento", "Vl cobrado", "Vl recebido", "Dif. receber")},
    **{n: "money_or_text" for n in ("Desconto", "Outras ded.", "Multa", "Juros", "IDA", "Atraso")},
}
# O CR tem duas colunas "Recebido": a de índice 5 é data; a de índice 24 é Sim/Não (texto).
CR_INDEX_KINDS = {5: "date"}

# Rótulos dos rodapés de TOTAL. Só servem para reconhecer a linha como técnica
# (R_FOOTER): os valores dos rodapés nunca são lidos.
CP_FOOTER = (4, "VALOR TOTAL")                       # (índice da coluna, rótulo exato)
CR_FOOTER = (6, ("TOTAL PÁGINA", "TOTAL GERAL"))
RRA_FOOTER = "TOTAL"                                 # coluna A
COMISSOES_FOOTER = "TOTAL GERAL"                     # coluna H, com a coluna A vazia
```

- [x] **Step 5: Rodar e ver passar**

Run: `python -m pytest tests/test_types.py -v`
Expected: 3 passed

- [x] **Step 6: Commit**

```bash
git add src/papaki_dre/types.py src/papaki_dre/config.py tests/test_types.py && git commit -m "feat: tipos compartilhados e configuração de negócio"
```

---

### Task 3: Conversão de valores (`coerce.py`)

Commit: a2a8c80 (ajustado depois em b84919c — lista fechada de texto em `money_or_text` e rejeição de sinal duplo)

**Files:**
- Create: `src/papaki_dre/coerce.py`
- Test: `tests/test_coerce.py`

Regras: `parse_money` aceita um único sinal (`-` antes ou depois de `R$`, ou parênteses), `R$` só no início, e número puro; sinal duplo é rejeitado. Nas colunas `money_or_text`, o único texto aceito é o da lista fechada `MONEY_TEXT_OK = {"-"}` (levantada no bruto real de set/2026); qualquer outro valor que não seja dinheiro válido gera `ValueError`, que os sanitizadores convertem em `ContractError`. Vazio vira `None`, nunca zero.

- [x] **Step 1: Escrever os testes que falham**

```python
import datetime as dt

import pytest

from papaki_dre.coerce import coerce_cell, parse_date, parse_datetime, parse_money


@pytest.mark.parametrize("text,expected", [
    ("R$ 21.132,72", 21132.72),
    ("R$ 3,31", 3.31),
    ("R$ 0,00", 0.0),
    ("R$ -1.050,16", -1050.16),
    ("-R$ 8,73", -8.73),
    ("(R$ 8,73)", -8.73),
    ("1234", 1234.0),
    ("1.234,5", 1234.5),
    ("0", 0.0),
])
def test_parse_money(text, expected):
    assert parse_money(text) == pytest.approx(expected)


@pytest.mark.parametrize("bad", ["abc", "R$ ", "1.5", "12,3,4", "-", "-R$ -5,00", "(-R$ 5,00)",
                                 "(R$ -5,00)", "--5", "R$ --5,00"])
def test_parse_money_rejects(bad):
    with pytest.raises(ValueError):
        parse_money(bad)


def test_parse_date_formats():
    assert parse_date("05/01/2026") == dt.datetime(2026, 1, 5)
    assert parse_date("2026-01-05") == dt.datetime(2026, 1, 5)


def test_parse_date_rejects_invalid():
    with pytest.raises(ValueError):
        parse_date("31/09/2026")


def test_parse_datetime():
    assert parse_datetime("14/11/2024 18:06:29") == dt.datetime(2024, 11, 14, 18, 6, 29)
    assert parse_datetime("14/11/2024") == dt.datetime(2024, 11, 14)


def test_coerce_empty_is_none_not_zero():
    assert coerce_cell("money", "") is None
    assert coerce_cell("text", "   ") is None
    assert coerce_cell("date", "") is None


def test_coerce_text_preserves_leading_zeros_and_spaces():
    assert coerce_cell("text", "00123") == "00123"
    assert coerce_cell("text", "UNDERDOGS ") == "UNDERDOGS "


def test_coerce_money_or_text_keeps_dash():
    assert coerce_cell("money_or_text", "-") == "-"
    assert coerce_cell("money_or_text", "R$ 10,00") == 10.0
    assert coerce_cell("money_or_text", "135") == 135.0


@pytest.mark.parametrize("bad", ["R$ 1,2,3", "12abc", "R$", "1.5", "n/d", "--", "-R$ -5,00"])
def test_coerce_money_or_text_rejects_anything_outside_the_closed_list(bad):
    with pytest.raises(ValueError):
        coerce_cell("money_or_text", bad)


def test_coerce_money_or_text_dash_tolerates_surrounding_spaces():
    assert coerce_cell("money_or_text", " - ") == " - "


def test_coerce_unknown_kind():
    with pytest.raises(ValueError):
        coerce_cell("banana", "x")
```

- [x] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_coerce.py -v`
Expected: FAIL com `ModuleNotFoundError: No module named 'papaki_dre.coerce'`

- [x] **Step 3: Implementar `coerce.py`**

```python
from __future__ import annotations

import datetime as dt
import re

# Únicos textos aceitos nas colunas money_or_text (CP: Deduções, Alterado; CR: Desconto,
# Outras ded., Multa, Juros, IDA, Atraso). Levantado no bruto real de set/2026: só "-".
MONEY_TEXT_OK = frozenset({"-"})

_NUMBER = re.compile(r"^(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?$")


def parse_money(text: str) -> float:
    """'R$ 21.132,72' -> 21132.72. Aceita um único sinal ('-' antes ou depois de 'R$', ou
    parênteses) e número puro. Sinal duplo ('-R$ -5,00', '(-5)') é rejeitado."""
    s = text.replace(" ", " ").strip()
    signs = 0
    if s.startswith("(") and s.endswith(")"):
        signs, s = signs + 1, s[1:-1].strip()
    if s.startswith("-"):
        signs, s = signs + 1, s[1:].strip()
    if s.startswith("R$"):
        s = s[2:].strip()
    if s.startswith("-"):
        signs, s = signs + 1, s[1:].strip()
    if signs > 1 or not _NUMBER.match(s):
        raise ValueError(f"valor monetário inválido: {text!r}")
    value = float(s.replace(".", "").replace(",", "."))
    return -value if signs else value


def parse_date(text: str) -> dt.datetime:
    s = text.strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f"data inválida: {text!r}")


def parse_datetime(text: str) -> dt.datetime:
    s = text.strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            continue
    return parse_date(s)


def coerce_cell(kind: str, raw):
    """Converte um valor bruto de texto. Vazio vira None (nunca zero)."""
    if raw is None:
        return None
    s = raw if isinstance(raw, str) else str(raw)
    if s.strip() == "":
        return None
    if kind == "text":
        return s
    if kind == "date":
        return parse_date(s)
    if kind == "datetime":
        return parse_datetime(s)
    if kind == "money":
        return parse_money(s)
    if kind == "money_or_text":
        if s.strip() in MONEY_TEXT_OK:
            return s
        return parse_money(s)          # qualquer outro texto: ValueError (vira ContractError)
    raise ValueError(f"tipo de coluna desconhecido: {kind!r}")
```

- [x] **Step 4: Rodar e ver passar**

Run: `python -m pytest tests/test_coerce.py -v`
Expected: 34 passed

- [x] **Step 5: Commit**

```bash
git add src/papaki_dre/coerce.py tests/test_coerce.py && git commit -m "feat: conversão de valores monetários, datas e vazios"
```

---

### Task 4: Reendereçamento de fórmulas

Commit: d3de098

**Files:**
- Create: `src/papaki_dre/formulas.py`
- Test: `tests/test_formulas.py`

- [x] **Step 1: Escrever os testes que falham**

```python
from papaki_dre.formulas import readdress
from papaki_dre.types import Formula


def test_relative_references_shift_with_the_row():
    assert readdress("=IF(K4=0,J4,0)", "L4", "L2") == "=IF(K2=0,J2,0)"


def test_absolute_references_stay():
    assert readdress("=SUM($A$1:A4)", "B4", "B2") == "=SUM($A$1:A2)"


def test_returns_formula_type():
    assert isinstance(readdress("=A1", "B1", "B2"), Formula)


def test_same_cell_is_identity():
    assert readdress("=J5+K5", "L5", "L5") == "=J5+K5"


def test_moving_down():
    assert readdress("=J5+K5", "L5", "L10") == "=J10+K10"
```

- [x] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_formulas.py -v`
Expected: FAIL com `ModuleNotFoundError`

- [x] **Step 3: Implementar**

```python
from openpyxl.formula.translate import Translator

from .types import Formula


def readdress(formula: str, origin: str, dest: str) -> Formula:
    """Move uma fórmula de `origin` para `dest` (ex.: 'L4' -> 'L2'), ajustando
    referências relativas e mantendo as absolutas."""
    return Formula(Translator(formula, origin=origin).translate_formula(dest))
```

- [x] **Step 4: Rodar e ver passar**

Run: `python -m pytest tests/test_formulas.py -v`
Expected: 5 passed

- [x] **Step 5: Commit**

```bash
git add src/papaki_dre/formulas.py tests/test_formulas.py && git commit -m "feat: reendereçamento de referências relativas de fórmulas"
```

---

### Task 5: Leitores de arquivo

Commit: 6cc7644 (ajustado depois em dd748ec — destino com hash do conteúdo, gravação atômica, limpeza segura do Excel, `UpdateLinks=0`)

**Files:**
- Create: `src/papaki_dre/readers.py`
- Test: `tests/test_readers.py`

Comportamento de `convert_xls_to_xlsx(src, dst)`: o XLSX é gravado no diretório de `dst` com o nome `<stem de dst>-<sha256[:12] do conteúdo de src>.xlsx` (`xlsx_cache_path`) e a função devolve esse caminho. Mesmo conteúdo reaproveita a conversão; brutos de mesmo nome e conteúdos diferentes nunca compartilham o destino. `_materialize` grava num temporário do mesmo diretório e só então faz `os.replace` para o destino, de modo que uma conversão interrompida nunca é reaproveitada. O Excel abre o arquivo com `UpdateLinks=0` e `ReadOnly=True`; `Close`/`Quit` sempre rodam e uma falha na limpeza não mascara a exceção original. Os testes do cache e da limpeza rodam sem Excel (COM falso via `monkeypatch`).

- [x] **Step 1: Escrever os testes que falham**

```python
import io
import re
import zipfile

import openpyxl
import pytest

from papaki_dre.readers import (_materialize, convert_xls_to_xlsx, open_xlsx, read_delimited,
                                xlsx_cache_path)
from papaki_dre.types import ContractError


def test_read_delimited_utf8_with_bom(tmp_path):
    p = tmp_path / "a.csv"
    p.write_bytes("Cód.;Valor\nx-1;R$ 1,00\n".encode("utf-8-sig"))
    enc, rows = read_delimited(p)
    assert enc == "utf-8-sig"
    assert rows == [["Cód.", "Valor"], ["x-1", "R$ 1,00"]]


def test_read_delimited_falls_back_to_cp1252(tmp_path):
    p = tmp_path / "b.csv"
    p.write_bytes("Emissão;Descrição\n05/01/2026;ação\n".encode("cp1252"))
    enc, rows = read_delimited(p)
    assert enc == "cp1252"
    assert rows[1] == ["05/01/2026", "ação"]


def test_read_delimited_keeps_quoted_semicolons_and_newlines(tmp_path):
    p = tmp_path / "c.csv"
    p.write_bytes('a;b\n"x;y";"linha1\nlinha2"\n'.encode("utf-8"))
    _, rows = read_delimited(p)
    assert rows[1] == ["x;y", "linha1\nlinha2"]


def _xlsx_with_bad_autofilter(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BASE COMISSÕES"
    ws.append(["a", "b"])
    ws.append([1, 2])
    ws.auto_filter.ref = "A1:B2"
    wb.save(path)
    buf = io.BytesIO()
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                data = re.sub(rb'<autoFilter ref="[^"]*"', b"<autoFilter ref=\"'X'!A3:N4\"", data)
            zout.writestr(item, data)
    path.write_bytes(buf.getvalue())


def test_open_xlsx_repairs_invalid_autofilter(tmp_path):
    p = tmp_path / "bad.xlsx"
    _xlsx_with_bad_autofilter(p)
    with pytest.raises(ValueError):
        openpyxl.load_workbook(p)          # confirma que o problema é real
    ws = open_xlsx(p).worksheets[0]
    assert [c.value for c in ws[2]] == [1, 2]


def test_open_xlsx_keeps_formulas_or_values_on_request(tmp_path):
    p = tmp_path / "f.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = 2
    wb.active["B1"] = "=A1*3"
    wb.save(p)
    assert open_xlsx(p).active["B1"].value == "=A1*3"


@pytest.mark.excel
@pytest.mark.real
def test_convert_xls_to_xlsx_keeps_formulas(rra_real, tmp_path):
    out = convert_xls_to_xlsx(rra_real, tmp_path / "rra.xlsx")
    ws = openpyxl.load_workbook(out)["Plan1"]
    assert ws.max_column >= 49
    formulas = sum(1 for r in ws.iter_rows(min_row=4) for c in r
                   if isinstance(c.value, str) and c.value.startswith("="))
    assert formulas > 1000


# --- cache da conversão XLS -> XLSX (sem Excel) ------------------------------------------

def test_cache_path_depends_on_content_not_on_name(tmp_path):
    a, b, c = tmp_path / "a", tmp_path / "b", tmp_path / "c"
    for d in (a, b, c):
        d.mkdir()
    (a / "rel.xls").write_bytes(b"conteudo 1")
    (b / "rel.xls").write_bytes(b"conteudo 2")
    (c / "rel.xls").write_bytes(b"conteudo 1")
    dst = tmp_path / "work" / "rel.xlsx"
    pa, pb, pc = (xlsx_cache_path(d / "rel.xls", dst) for d in (a, b, c))
    assert pa != pb                       # mesmo nome, conteúdos diferentes: destinos diferentes
    assert pa == pc                       # mesmo conteúdo: reaproveita
    assert pa.parent == dst.parent and pa.suffix == ".xlsx"
    assert pa.name.startswith("rel-")


def test_materialize_never_reuses_a_partial_file(tmp_path):
    target = tmp_path / "work" / "rel-abc.xlsx"

    def broken(tmp):
        tmp.write_bytes(b"parcial")
        raise RuntimeError("Excel caiu no meio do SaveAs")

    with pytest.raises(RuntimeError, match="no meio"):
        _materialize(target, broken)
    assert not target.exists()
    assert list(target.parent.iterdir()) == []          # temporário removido

    calls = []

    def good(tmp):
        calls.append(tmp)
        assert tmp.parent == target.parent and tmp != target
        tmp.write_bytes(b"completo")

    assert _materialize(target, good) == target
    assert target.read_bytes() == b"completo"
    assert _materialize(target, good) == target           # agora reaproveita
    assert len(calls) == 1


def test_materialize_requires_the_output(tmp_path):
    with pytest.raises(ContractError):
        _materialize(tmp_path / "x.xlsx", lambda tmp: None)


class _FakeWorkbook:
    def __init__(self, log):
        self.log = log

    def SaveAs(self, path, FileFormat):
        self.log.append(("SaveAs", FileFormat))
        raise RuntimeError("falha original no SaveAs")

    def Close(self, save):
        self.log.append("Close")
        raise RuntimeError("falha no Close")


class _FakeExcel:
    def __init__(self, log):
        self.log = log
        outer = self

        class _Books:
            def Open(self, path, **kw):
                outer.log.append(("Open", kw))
                return _FakeWorkbook(outer.log)
        self.Workbooks = _Books()

    def __setattr__(self, name, value):
        if name in ("Visible", "DisplayAlerts"):
            self.log.append((name, value))
        object.__setattr__(self, name, value)

    def Quit(self):
        self.log.append("Quit")
        raise RuntimeError("falha no Quit")


def test_excel_is_always_quit_without_masking_the_original_error(tmp_path, monkeypatch):
    win32 = pytest.importorskip("win32com.client")
    log = []
    monkeypatch.setattr(win32, "DispatchEx", lambda name: _FakeExcel(log))
    src = tmp_path / "rel.xls"
    src.write_bytes(b"x")
    with pytest.raises(RuntimeError, match="falha original no SaveAs"):
        convert_xls_to_xlsx(src, tmp_path / "work" / "rel.xlsx")
    opened = [e for e in log if isinstance(e, tuple) and e[0] == "Open"]
    assert opened and opened[0][1].get("UpdateLinks") == 0 and opened[0][1].get("ReadOnly")
    assert "Close" in log and log[-1] == "Quit"
    assert not any((tmp_path / "work").glob("*.xlsx"))


def test_excel_is_quit_even_if_setup_fails(tmp_path, monkeypatch):
    win32 = pytest.importorskip("win32com.client")
    log = []

    class _Bad(_FakeExcel):
        def __setattr__(self, name, value):
            if name == "Visible":
                raise RuntimeError("falha ao configurar")
            super().__setattr__(name, value)

    monkeypatch.setattr(win32, "DispatchEx", lambda name: _Bad(log))
    src = tmp_path / "rel.xls"
    src.write_bytes(b"y")
    with pytest.raises(RuntimeError, match="falha ao configurar"):
        convert_xls_to_xlsx(src, tmp_path / "work" / "rel.xlsx")
    assert log[-1] == "Quit"
```

- [x] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_readers.py -v -m "not excel"`
Expected: FAIL com `ModuleNotFoundError: No module named 'papaki_dre.readers'`

- [x] **Step 3: Implementar `readers.py`**

```python
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


def com_retry(fn, tries: int = 15, wait: float = 2.0):
    """Repete uma chamada COM enquanto o Excel responde 'chamada rejeitada' (ocupado)."""
    import time

    import pywintypes

    for i in range(tries):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult != _RPC_E_CALL_REJECTED or i == tries - 1:
                raise
            time.sleep(wait)


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
        wb = com_retry(lambda: excel.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=True))
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
```

- [x] **Step 4: Rodar e ver passar**

Run: `python -m pytest tests/test_readers.py -v`
Expected: 11 passed, incluindo o teste `excel` (abre o Excel por alguns segundos). Se o Excel não estiver disponível, rode com `-m "not excel"` e registre isso no relatório.

- [x] **Step 5: Commit**

```bash
git add src/papaki_dre/readers.py tests/test_readers.py && git commit -m "feat: leitores de CSV (codificação variável), XLSX com autoFilter inválido e XLS via Excel"
```

---

### Task 6: Contrato de cabeçalhos

Commit: 9e32451

**Files:**
- Create: `scripts/capture_headers.py`, `src/papaki_dre/headers.py`, `src/papaki_dre/headers.json` (gerado)
- Test: `tests/test_headers.py`

- [x] **Step 1: Escrever os testes que falham**

```python
import pytest

from papaki_dre.headers import check_header, load_headers
from papaki_dre.types import ContractError


def test_contract_has_the_expected_column_counts():
    h = load_headers()
    assert len(h["rra"]) == 49
    assert len(h["cp"]) == 41
    assert len(h["cr"]) == 28
    assert len(h["comissoes"]) == 14


def test_contract_has_no_blank_names():
    for name, cols in load_headers().items():
        assert all(c.strip() for c in cols), name


def test_check_header_accepts_equal():
    check_header("cp", ["a", "b"], ["a", "b"])


def test_check_header_rejects_wrong_count():
    with pytest.raises(ContractError, match="2 colunas.*esperadas 3"):
        check_header("cp", ["a", "b"], ["a", "b", "c"])


def test_check_header_rejects_wrong_order_and_names_the_column():
    with pytest.raises(ContractError, match=r"coluna 2.*'c'.*'b'"):
        check_header("cp", ["a", "c"], ["a", "b"])
```

- [x] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_headers.py -v`
Expected: FAIL com `ModuleNotFoundError: No module named 'papaki_dre.headers'`

- [x] **Step 3: Implementar `headers.py`**

```python
from __future__ import annotations

import json
import unicodedata
from importlib import resources

from .types import ContractError


def normalize(name) -> str:
    return unicodedata.normalize("NFC", str(name).strip())


def load_headers() -> dict[str, list[str]]:
    text = resources.files("papaki_dre").joinpath("headers.json").read_text(encoding="utf-8")
    return json.loads(text)


def check_header(report: str, header: list[str], expected: list[str]) -> None:
    """Para a rodada se o cabeçalho divergir em quantidade ou ordem."""
    if len(header) != len(expected):
        raise ContractError(
            f"{report}: o arquivo tem {len(header)} colunas; esperadas {len(expected)}")
    for i, (got, want) in enumerate(zip(header, expected), start=1):
        if got != want:
            raise ContractError(f"{report}: coluna {i} é {got!r}; esperado {want!r}")
```

- [x] **Step 4: Escrever o script que gera `headers.json` a partir dos brutos de referência**

`scripts/capture_headers.py`:

```python
"""Gera src/papaki_dre/headers.json a partir dos brutos de referência (setembro/2026).
Uso: python scripts/capture_headers.py
Revise o JSON gerado: ele é o contrato de cabeçalhos dos 4 relatórios."""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from papaki_dre.headers import normalize  # noqa: E402
from papaki_dre.readers import convert_xls_to_xlsx, open_xlsx, read_delimited  # noqa: E402
import openpyxl  # noqa: E402

BRUTO = ROOT / "amostras" / "2026-09" / "bruto"
OUT = ROOT / "src" / "papaki_dre" / "headers.json"
WORK = ROOT / "out" / "work"


def main():
    _, cp = read_delimited(BRUTO / "CP_papaki_consolidado_2026-01-01_a_2026-09-30.csv")
    _, cr = read_delimited(BRUTO / "CR_untitled_UTLD_2026-01-01_a_2026-09-30.csv")
    com = open_xlsx(BRUTO / "COMISSOES_untitled_UTLD_2026-01-01_a_2026-09-30.xlsx").worksheets[0]
    xlsx = convert_xls_to_xlsx(BRUTO / "Relatório Rentabilidade.xls", WORK / "rra.xlsx")
    rra = openpyxl.load_workbook(xlsx)["Plan1"]

    headers = {
        "cp": [normalize(c) for c in cp[0]],
        "cr": [normalize(c) for c in cr[0]],
        "comissoes": [normalize(c.value) for c in com[3][:14]],
        "rra": [normalize(c.value) for c in rra[3][:49]],
    }
    expected = {"cp": 41, "cr": 28, "comissoes": 14, "rra": 49}
    for k, n in expected.items():
        assert len(headers[k]) == n, f"{k}: {len(headers[k])} colunas, esperadas {n}"
        assert all(h and h != "None" for h in headers[k]), f"{k}: cabeçalho vazio"
    OUT.write_text(json.dumps(headers, ensure_ascii=False, indent=2), encoding="utf-8")
    print("headers.json gerado:", {k: len(v) for k, v in headers.items()})


if __name__ == "__main__":
    main()
```

`headers.json` é um objeto com as chaves `cp`, `cr`, `comissoes` e `rra`, cada uma com a lista ordenada dos nomes de coluna (41, 28, 14 e 49 nomes).

- [x] **Step 5: Gerar e revisar o contrato**

Run: `python scripts/capture_headers.py`
Expected: `headers.json gerado: {'cp': 41, 'cr': 28, 'comissoes': 14, 'rra': 49}`

Abra `src/papaki_dre/headers.json` e confira: `cp[0] == "Cód."`, `cp` contém `"Nº job"`, `"Job"`, `"Localidade"`, `"Valor"`; `cr` contém `"Vl Documento"` e `"Descrição"` e a coluna `"Recebido"` duas vezes (posições 6 e 25); `rra[0] == "Núm. JOB"` e contém `"Núcleo"` e `"Título do job"`; `comissoes[0] == "Modelo de Orçamento"` e `comissoes[7] == "Comissão"`.

- [x] **Step 6: Rodar e ver passar**

Run: `python -m pytest tests/test_headers.py -v`
Expected: 5 passed

- [x] **Step 7: Commit**

```bash
git add scripts/capture_headers.py src/papaki_dre/headers.py src/papaki_dre/headers.json tests/test_headers.py && git commit -m "feat: contrato de cabeçalhos dos 4 relatórios (49/41/28/14 colunas)"
```

---

### Task 7: Sanitizador do CP (Contas a Pagar)

Commit: 0bf76e4 (ajustado depois em 9984a05, fb97ad2 e 84d3786)

**Files:**
- Create: `src/papaki_dre/sanitize/cp.py`
- Test: `tests/test_sanitize_cp.py`

Regras (spec §6): validar cabeçalho (41 colunas); remover linhas técnicas — vazia (`R_BLANK`) e o rodapé `VALOR TOTAL` (`R_FOOTER`), reconhecido só pelo rótulo exato no índice 4; o valor do rodapé nunca é lido, e rodapé ausente ou repetido não gera aviso; registro não vazio, que não seja rodapé reconhecido e tenha nº de colunas diferente do cabeçalho (aspas não fechadas, registro truncado, rodapé desconhecido) → `ContractError`; tipar as colunas (valor ilegível → `ContractError`); excluir se `Nº job` preenchido **e** `Job` sem `PESQUISA E TRATAMENTO` (na base consolidada inteira); depois manter só `Localidade = UTLD PRODUÇÃO DE FILMES LTDA`. `rows` e `db_rows` são listas distintas.

- [x] **Step 1: Escrever os testes que falham**

```python
import pytest

from papaki_dre.config import R_BLANK, R_FOOTER, R_JOB, R_LOC, UTLD
from papaki_dre.headers import load_headers
from papaki_dre.sanitize.cp import sanitize_cp
from papaki_dre.types import ContractError

H = load_headers()["cp"]


def _r(d):
    row = [""] * len(H)
    for name, value in d.items():
        row[H.index(name)] = value
    return row


def _write(tmp_path, rows, header=None, enc="utf-8-sig"):
    p = tmp_path / "cp.csv"
    lines = [";".join(r) for r in [header or H] + rows]
    p.write_bytes(("\n".join(lines) + "\n").encode(enc))
    return p


def _footer(total):
    f = [""] * 23
    f[4], f[5] = "VALOR TOTAL", total
    return f


BASE = {"Cód.": "untitled-1", "Valor": "R$ 100,00", "Localidade": UTLD,
        "Vcto": "05/01/2026", "Competência": "01/01/2026"}


def test_keeps_untitled_rows_without_job(tmp_path):
    p = _write(tmp_path, [_r(BASE), _footer("R$ 100,00")])
    s = sanitize_cp(p)
    assert s.kept == 1
    assert s.rows[0][H.index("Valor")] == 100.0
    assert s.rows[0][H.index("Vcto")].year == 2026
    assert s.balanced()


def test_excludes_filled_job_without_pesquisa(tmp_path):
    ex = _r({**BASE, "Cód.": "untitled-2", "Nº job": "#75111", "Job": "Meli | Gut"})
    s = sanitize_cp(_write(tmp_path, [_r(BASE), ex, _footer("R$ 200,00")]))
    assert s.kept == 1
    assert [e.reason for e in s.exclusions] == [R_JOB]
    assert s.exclusions[0].source_row == 3   # cabeçalho = 1; primeira linha de dados = 2


def test_keeps_filled_job_with_pesquisa(tmp_path):
    ok = _r({**BASE, "Nº job": "#75111", "Job": "PESQUISA E TRATAMENTO | X"})
    s = sanitize_cp(_write(tmp_path, [ok, _footer("R$ 100,00")]))
    assert s.kept == 1 and not s.exclusions


def test_excludes_other_localidade(tmp_path):
    other = _r({**BASE, "Localidade": "SOMA"})
    s = sanitize_cp(_write(tmp_path, [_r(BASE), other, _footer("R$ 200,00")]))
    assert s.kept == 1
    assert [e.reason for e in s.exclusions] == [R_LOC]


def test_job_rule_runs_before_localidade_filter(tmp_path):
    both = _r({**BASE, "Localidade": "SOMA", "Nº job": "#1", "Job": "x"})
    s = sanitize_cp(_write(tmp_path, [both, _footer("R$ 100,00")]))
    assert [e.reason for e in s.exclusions] == [R_JOB]


def test_footer_and_blank_are_technical(tmp_path):
    p = _write(tmp_path, [_r(BASE), [""] * len(H), _footer("R$ 100,00")])
    s = sanitize_cp(p)
    assert sorted(t.reason for t in s.technical) == sorted([R_BLANK, R_FOOTER])
    assert s.kept == 1 and s.balanced()
    assert not hasattr(s, "controls")


def test_footer_value_is_never_read(tmp_path):
    s = sanitize_cp(_write(tmp_path, [_r(BASE), _footer("ilegível")]))
    assert [t.reason for t in s.technical] == [R_FOOTER]
    assert not s.warnings


@pytest.mark.parametrize("pos,label", [(3, "VALOR TOTAL"), (4, "VALOR TOTAL GERAL"), (4, "TOTAL")])
def test_footer_label_must_be_exact_and_at_index_4(tmp_path, pos, label):
    loose = [""] * 23
    loose[pos], loose[5] = label, "R$ 1,00"
    with pytest.raises(ContractError, match="colunas"):
        sanitize_cp(_write(tmp_path, [_r(BASE), loose]))


def test_short_record_stops_the_run(tmp_path):
    with pytest.raises(ContractError, match="registro 3"):
        sanitize_cp(_write(tmp_path, [_r(BASE), _r(BASE)[:-1], _footer("R$ 1,00")]))


def test_unclosed_quote_stops_the_run(tmp_path):
    bad = _r({**BASE, "Cód.": '"sem fechar'})
    with pytest.raises(ContractError):
        sanitize_cp(_write(tmp_path, [_r(BASE), bad, _r(BASE), _footer("R$ 1,00")]))


def test_quoted_multiline_field_is_valid(tmp_path):
    ok = _r({**BASE, "Cód.": '"linha1\nlinha2"'})
    s = sanitize_cp(_write(tmp_path, [ok, _footer("R$ 1,00")]))
    assert s.kept == 1 and s.rows[0][H.index("Cód.")] == "linha1\nlinha2"


def test_duplicate_footer_is_silent(tmp_path):
    s = sanitize_cp(_write(tmp_path, [_r(BASE), _footer("R$ 1,00"), _footer("R$ 1,00")]))
    assert [t.reason for t in s.technical] == [R_FOOTER, R_FOOTER]
    assert not s.warnings


@pytest.mark.parametrize("text", ["-R$ 5,00", "R$ -5,00", "(R$ 5,00)"])
def test_negative_valor(tmp_path, text):
    s = sanitize_cp(_write(tmp_path, [_r({**BASE, "Valor": text})]))
    assert s.rows[0][H.index("Valor")] == -5.0


def test_rows_and_db_rows_are_not_aliased(tmp_path):
    s = sanitize_cp(_write(tmp_path, [_r(BASE)]))
    assert s.rows[0] is not s.db_rows[0]
    s.rows[0][H.index("Valor")] = "=X1"
    assert s.db_rows[0][H.index("Valor")] == 100.0


def test_header_mismatch_stops_the_run(tmp_path):
    bad = H[:-1] + ["Outra coisa"]
    with pytest.raises(ContractError, match="coluna 41"):
        sanitize_cp(_write(tmp_path, [_r(BASE)], header=bad))


def test_invalid_value_stops_the_run(tmp_path):
    bad = _r({**BASE, "Valor": "dez reais"})
    with pytest.raises(ContractError, match="Valor"):
        sanitize_cp(_write(tmp_path, [bad, _footer("R$ 0,00")]))


def test_missing_footer_is_silent(tmp_path):
    s = sanitize_cp(_write(tmp_path, [_r(BASE)]))
    assert s.kept == 1 and not s.warnings


@pytest.mark.real
def test_real_september_file(cp_real):
    s = sanitize_cp(cp_real)
    assert s.encoding == "utf-8-sig"
    assert s.balanced()
    assert s.raw_count == 36999
    assert s.kept == 1110
    assert len(s.exclusions) == 21505 + 14382
    assert sorted(t.reason for t in s.technical) == [R_BLANK, R_FOOTER]
```

- [x] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_sanitize_cp.py -v`
Expected: FAIL com `ModuleNotFoundError: No module named 'papaki_dre.sanitize.cp'`

- [x] **Step 3: Implementar `sanitize/cp.py`**

```python
from __future__ import annotations

from ..coerce import coerce_cell
from ..config import CP_FOOTER, CP_KINDS, PESQUISA, R_BLANK, R_FOOTER, R_JOB, R_LOC, UTLD
from ..headers import check_header, load_headers, normalize
from ..readers import read_delimited
from ..types import ContractError, Removal, Sanitized

REPORT = "cp"


def _is_footer(raw: list[str]) -> bool:
    """Rodapé VALOR TOTAL: rótulo exato na posição em que o bruto o traz. Só é reconhecido
    para ser descartado como linha técnica; seus valores nunca são lidos."""
    i, label = CP_FOOTER
    return len(raw) > i and raw[i].strip() == label


def sanitize_cp(path, *, expected_header: list[str] | None = None) -> Sanitized:
    enc, rows = read_delimited(path)
    if not rows:
        raise ContractError("cp: arquivo vazio")
    header = [normalize(h) for h in rows[0]]
    check_header(REPORT, header, expected_header or load_headers()[REPORT])
    ncols = len(header)
    ix = {name: i for i, name in enumerate(header)}
    i_job, i_nome, i_loc = ix["Nº job"], ix["Job"], ix["Localidade"]

    out = Sanitized(report=REPORT, header=header, rows=[], db_rows=[], source_rows=[],
                    raw_count=len(rows) - 1, encoding=enc)

    for n, raw in enumerate(rows[1:], start=2):
        if not any(c.strip() for c in raw):
            out.technical.append(Removal(n, R_BLANK))
            continue
        if _is_footer(raw):
            out.technical.append(Removal(n, R_FOOTER))
            continue
        if len(raw) != ncols:   # aspas não fechadas, registro truncado, rodapé desconhecido
            raise ContractError(
                f"cp: registro {n} tem {len(raw)} colunas; esperadas {ncols}")
        typed = []
        for i, cell in enumerate(raw):
            kind = CP_KINDS.get(header[i], "text")
            try:
                typed.append(coerce_cell(kind, cell))
            except ValueError as e:
                raise ContractError(f"cp: registro {n}, coluna {header[i]!r}: {e}") from e
        if raw[i_job].strip() and PESQUISA not in raw[i_nome].upper():
            out.exclusions.append(Removal(n, R_JOB))
            continue
        if normalize(raw[i_loc]) != UTLD:
            out.exclusions.append(Removal(n, R_LOC))
            continue
        out.rows.append(typed)
        out.db_rows.append(list(typed))   # listas distintas: rows (Excel) e db_rows (banco)
        out.source_rows.append(n)
    return out
```

- [x] **Step 4: Rodar e ver passar**

Run: `python -m pytest tests/test_sanitize_cp.py -v`
Expected: 22 passed (o teste `real` roda se o bruto existir; ele leva alguns segundos).

Se `test_real_september_file` divergir em `len(s.exclusions)`: pare e investigue antes de ajustar o número. Os valores esperados vêm da análise de 08/10/2026 e foram confirmados no bruto real: 36.999 registros (sem o cabeçalho) = 1.110 mantidos + 2 técnicos (1 vazio + 1 rodapé) + 21.505 excluídos pela regra do job + 14.382 fora da Localidade.

- [x] **Step 5: Commit**

```bash
git add src/papaki_dre/sanitize/cp.py tests/test_sanitize_cp.py && git commit -m "feat: sanitizador do CP com regra de job, filtro de Localidade e total de controle"
```

(Mensagem histórica do commit 0bf76e4; o total de controle foi removido depois, em 9984a05.)

---

### Task 8: Sanitizador do CR (Contas a Receber)

Commit: 4c15da6 (ajustado depois em 9984a05, fb97ad2 e 84d3786)

**Files:**
- Create: `src/papaki_dre/sanitize/cr.py`
- Test: `tests/test_sanitize_cr.py`

Regras: validar 28 colunas; CSV em Windows-1252; **nenhuma exclusão de linha** ("Desconsiderar" é classificação); remover só linhas técnicas: vazias e os rodapés `TOTAL PÁGINA`/`TOTAL GERAL` (31 colunas no bruto real), reconhecidos pelo rótulo exato no índice 6 — um cliente como `TOTALPASS` é dado, não rodapé; os valores dos rodapés nunca são lidos e rodapé ausente não gera aviso. Registro não vazio, não rodapé e com nº de colunas diferente do cabeçalho → `ContractError`. Tipar datas e valores; a primeira coluna `Recebido` (índice 5) é data e a segunda (índice 24) é texto; identificadores ficam como texto com zeros à esquerda. `rows` e `db_rows` são listas distintas.

- [x] **Step 1: Escrever os testes que falham**

```python
import pytest

from papaki_dre.config import R_BLANK, R_FOOTER
from papaki_dre.headers import load_headers
from papaki_dre.sanitize.cr import sanitize_cr
from papaki_dre.types import ContractError

H = load_headers()["cr"]
SHIFT = 3   # no bruto real o rodapé tem 31 colunas, com os totais deslocados


def _r(d):
    row = [""] * len(H)
    for name, value in d.items():
        row[H.index(name)] = value
    return row


def _write(tmp_path, rows, header=None, enc="cp1252"):
    p = tmp_path / "cr.csv"
    p.write_bytes(("\n".join(";".join(r) for r in [header or H] + rows) + "\n").encode(enc))
    return p


def _footer(label, doc, cob, rec, dif):
    f = [""] * (len(H) + SHIFT)
    f[6] = label
    f[H.index("Vl Documento") + SHIFT] = doc
    f[H.index("Vl cobrado") + SHIFT] = cob
    f[H.index("Vl recebido") + SHIFT] = rec
    f[H.index("Dif. receber") + SHIFT] = dif
    return f


ROW = {"Cliente": "ITAU", "Emissão": "05/01/2026", "Vcto": "05/01/2026",
       "Vl Documento": "R$ 100,00", "Vl cobrado": "R$ 100,00", "Vl recebido": "R$ 60,00",
       "Dif. receber": "R$ 40,00", "Desconto": "-", "Competência": "2026-01-01",
       "Descrição": "RENDIMENTOS ação"}


def _file(tmp_path, extra=()):
    return _write(tmp_path, [_r(ROW), *extra,
                             _footer("TOTAL PÁGINA", "R$ 100,00", "R$ 100,00", "R$ 60,00", "R$ 40,00"),
                             _footer("TOTAL GERAL", "R$ 100,00", "R$ 100,00", "R$ 60,00", "R$ 40,00")])


def test_reads_cp1252_and_types_the_columns(tmp_path):
    s = sanitize_cr(_file(tmp_path))
    assert s.encoding == "cp1252"
    row = s.rows[0]
    assert row[H.index("Descrição")] == "RENDIMENTOS ação"
    assert row[H.index("Vl Documento")] == 100.0
    assert row[H.index("Emissão")].month == 1
    assert row[H.index("Competência")].year == 2026
    assert row[H.index("Desconto")] == "-"          # texto preservado


def test_second_recebido_column_stays_text(tmp_path):
    second = max(i for i, n in enumerate(H) if n == "Recebido")
    first = H.index("Recebido")
    row = _r({**ROW})
    row[first] = "06/01/2026"
    row[second] = "Sim"
    s = sanitize_cr(_write(tmp_path, [row]))
    assert s.rows[0][first].day == 6
    assert s.rows[0][second] == "Sim"


def test_footers_are_technical(tmp_path):
    s = sanitize_cr(_file(tmp_path))
    assert s.kept == 1
    assert [t.reason for t in s.technical] == [R_FOOTER, R_FOOTER]
    assert not hasattr(s, "controls")
    assert not s.warnings
    assert s.balanced()


def test_footer_values_are_never_read(tmp_path):
    p = _write(tmp_path, [_r(ROW), _footer("TOTAL GERAL", "lixo", "", "x", "")])
    s = sanitize_cr(p)
    assert [t.reason for t in s.technical] == [R_FOOTER] and not s.warnings


def test_missing_footer_is_silent(tmp_path):
    s = sanitize_cr(_write(tmp_path, [_r(ROW)]))
    assert s.kept == 1 and not s.warnings


def test_totalpass_client_is_data_not_footer(tmp_path):
    row = _r({**ROW, "Cliente": "TOTALPASS"})
    s = sanitize_cr(_file(tmp_path, [row]))
    assert s.kept == 2
    assert s.rows[1][H.index("Cliente")] == "TOTALPASS"
    assert [t.reason for t in s.technical] == [R_FOOTER, R_FOOTER]


def test_footer_label_only_counts_at_index_6(tmp_path):
    f = _footer("", "R$ 1,00", "R$ 1,00", "R$ 1,00", "R$ 1,00")
    f[2] = "TOTAL GERAL"
    with pytest.raises(ContractError, match="colunas"):
        sanitize_cr(_write(tmp_path, [_r(ROW), f]))


def test_startswith_total_is_not_a_footer(tmp_path):
    f = _footer("TOTALPASS", "R$ 1,00", "R$ 1,00", "R$ 1,00", "R$ 1,00")
    with pytest.raises(ContractError, match="colunas"):
        sanitize_cr(_write(tmp_path, [_r(ROW), f]))


def test_short_record_stops_the_run(tmp_path):
    with pytest.raises(ContractError, match="registro 3"):
        sanitize_cr(_write(tmp_path, [_r(ROW), _r(ROW)[:-2]]))


def test_unclosed_quote_stops_the_run(tmp_path):
    bad = _r({**ROW, "Cliente": '"ITAU'})
    with pytest.raises(ContractError):
        sanitize_cr(_write(tmp_path, [_r(ROW), bad, _r(ROW)]))


def test_quoted_multiline_field_is_valid(tmp_path):
    ok = _r({**ROW, "Descrição": '"linha1\nlinha2"'})
    s = sanitize_cr(_file(tmp_path, [ok]))
    assert s.kept == 2 and s.rows[1][H.index("Descrição")] == "linha1\nlinha2"


def test_never_excludes_business_rows(tmp_path):
    desconsiderar = _r({**ROW, "Plano de contas": "Desconsiderar"})
    s = sanitize_cr(_file(tmp_path, [desconsiderar]))
    assert s.kept == 2 and not s.exclusions


def test_blank_rows_are_technical(tmp_path):
    s = sanitize_cr(_file(tmp_path, [[""] * len(H)]))
    assert R_BLANK in [t.reason for t in s.technical]


def test_identifier_keeps_leading_zeros(tmp_path):
    s = sanitize_cr(_write(tmp_path, [_r({**ROW, "NFe": "000123", "Lote": "0042"})]))
    assert s.rows[0][H.index("NFe")] == "000123"
    assert s.rows[0][H.index("Lote")] == "0042"


def test_rows_and_db_rows_are_not_aliased(tmp_path):
    s = sanitize_cr(_write(tmp_path, [_r(ROW)]))
    assert s.rows[0] is not s.db_rows[0]
    s.rows[0][H.index("Cliente")] = "OUTRO"
    assert s.db_rows[0][H.index("Cliente")] == "ITAU"


def test_header_mismatch_stops_the_run(tmp_path):
    with pytest.raises(ContractError):
        sanitize_cr(_write(tmp_path, [_r(ROW)], header=H[:-1]))


@pytest.mark.real
def test_real_september_file(cr_real):
    s = sanitize_cr(cr_real)
    assert s.encoding == "cp1252"
    assert s.balanced()
    assert s.kept == 484
    assert s.raw_count == 486
    assert [t.reason for t in s.technical] == [R_FOOTER, R_FOOTER]
    assert not s.exclusions
```

- [x] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_sanitize_cr.py -v`
Expected: FAIL com `ModuleNotFoundError: No module named 'papaki_dre.sanitize.cr'`

- [x] **Step 3: Implementar `sanitize/cr.py`**

```python
from __future__ import annotations

from ..coerce import coerce_cell
from ..config import CR_FOOTER, CR_INDEX_KINDS, CR_KINDS, R_BLANK, R_FOOTER
from ..headers import check_header, load_headers, normalize
from ..readers import read_delimited
from ..types import ContractError, Removal, Sanitized

REPORT = "cr"


def _kind(i: int, name: str) -> str:
    return CR_INDEX_KINDS.get(i) or CR_KINDS.get(name, "text")


def _is_footer(raw: list[str]) -> bool:
    """Rodapés TOTAL PÁGINA / TOTAL GERAL: rótulo exato na posição em que o bruto o traz.
    Só são reconhecidos para serem descartados como linha técnica; os valores não são lidos."""
    i, labels = CR_FOOTER
    return len(raw) > i and normalize(raw[i]) in labels


def sanitize_cr(path, *, expected_header: list[str] | None = None) -> Sanitized:
    enc, rows = read_delimited(path)
    if not rows:
        raise ContractError("cr: arquivo vazio")
    header = [normalize(h) for h in rows[0]]
    check_header(REPORT, header, expected_header or load_headers()[REPORT])
    ncols = len(header)

    out = Sanitized(report=REPORT, header=header, rows=[], db_rows=[], source_rows=[],
                    raw_count=len(rows) - 1, encoding=enc)
    for n, raw in enumerate(rows[1:], start=2):
        if not any(c.strip() for c in raw):
            out.technical.append(Removal(n, R_BLANK))
            continue
        if _is_footer(raw):
            out.technical.append(Removal(n, R_FOOTER))
            continue
        if len(raw) != ncols:   # aspas não fechadas, registro truncado, rodapé desconhecido
            raise ContractError(
                f"cr: registro {n} tem {len(raw)} colunas; esperadas {ncols}")
        typed = []
        for i, cell in enumerate(raw):
            try:
                typed.append(coerce_cell(_kind(i, header[i]), cell))
            except ValueError as e:
                raise ContractError(f"cr: registro {n}, coluna {header[i]!r}: {e}") from e
        out.rows.append(typed)
        out.db_rows.append(list(typed))   # listas distintas: rows (Excel) e db_rows (banco)
        out.source_rows.append(n)
    return out
```

- [x] **Step 4: Rodar e ver passar**

Run: `python -m pytest tests/test_sanitize_cr.py -v`
Expected: 17 passed (contagem real: 486 registros = 484 mantidos + 2 rodapés).

- [x] **Step 5: Commit**

```bash
git add src/papaki_dre/sanitize/cr.py tests/test_sanitize_cr.py && git commit -m "feat: sanitizador do CR (Windows-1252, rodapé deslocado, sem exclusões de negócio)"
```

(Mensagem histórica do commit 4c15da6; a leitura do rodapé deslocado foi removida depois, em 9984a05.)

---

### Task 9: Sanitizador das Comissões (Real x Orçado)

Commit: fa0b6bb (ajustado depois em 9984a05, ddffb42, 84d3786 e 8d85af7)

**Files:**
- Create: `src/papaki_dre/sanitize/_xlsx.py` (extraído em ddffb42; usado também pelo RRA na Task 10), `src/papaki_dre/sanitize/comissoes.py`
- Test: `tests/test_sanitize_comissoes.py`, `tests/xlsxcache.py` (helper de teste)

Regras: XLSX com título na linha 1, cabeçalho na linha 3 (14 colunas) e dados a partir da linha 4; célula de cabeçalho preenchida além da 14ª coluna → `ContractError` (células vazias além são ignoradas). A linha final `TOTAL GERAL` (rótulo exato na coluna H, colunas A..G e I vazias, somas em J:N) é técnica (`R_FOOTER`) e seus valores não são lidos. Manter só `Modelo de Orçamento = Untitled - Publicidade Padrão`; excluir `Título do Job` com `PESQUISA E TRATAMENTO`; excluir coluna H nas seis opções (lista explícita; opções novas ficam). Ordem de precedência dos motivos: modelo, pesquisa, coluna H. Fórmulas de origem (colunas L:N) são preservadas e **reendereçadas** para a linha de destino (`DEST["comissoes"]["first_row"] + posição`). A leitura usa `iter_rows(max_col=NCOLS)`.

Cópia de linha (`RowCopier`, comum a Comissões e RRA): erro em cache diferente de `#DIV/0!` numa linha mantida → `ContractError` (em linha excluída, não bloqueia); fórmula → `Formula` reendereçada em `rows` e valor em cache em `db_rows` (`#DIV/0!` vira `None`; fórmula sem valor em cache vira `None` e gera **um** aviso agregado em `warnings`); erro literal → `XlError` em `rows` e `None` em `db_rows`.

- [x] **Step 1: Escrever o helper de teste e os testes que falham**

`tests/xlsxcache.py` (o openpyxl não calcula fórmulas; o helper grava `<v>` no XML como o Excel faria):

```python
"""Helper de teste: grava valores em cache (<v>) em células com fórmula de um XLSX
gerado pelo openpyxl (que não calcula fórmulas), simulando o que o Excel salvaria."""
import io
import re
import zipfile
from pathlib import Path


def inject_cache(path, values: dict[str, object]) -> None:
    """`values`: {"L4": 100, "N5": "#DIV/0!"}. Texto começando com '#' vira erro (t="e")."""
    path = Path(path)
    pending = dict(values)
    buf = io.BytesIO()
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                for ref in list(pending):
                    val = pending[ref]
                    pat = re.compile(rb'<c r="' + ref.encode() + rb'"[^>]*><f>(.*?)</f><v\s*/>|'
                                     rb'<c r="' + ref.encode() + rb'"[^>]*><f>(.*?)</f><v>[^<]*</v>')
                    m = pat.search(data)
                    if not m:
                        continue
                    f = m.group(1) if m.group(1) is not None else m.group(2)
                    if isinstance(val, str) and val.startswith("#"):
                        t, v = b' t="e"', val.encode()
                    elif isinstance(val, str):
                        t, v = b' t="str"', val.encode()
                    else:
                        t, v = b"", repr(val).encode()
                    new = b'<c r="' + ref.encode() + b'"' + t + b"><f>" + f + b"</f><v>" + v + b"</v>"
                    data = data[:m.start()] + new + data[m.end():]
                    del pending[ref]
            zout.writestr(item, data)
    if pending:
        raise AssertionError(f"células sem fórmula para receber cache: {sorted(pending)}")
    path.write_bytes(buf.getvalue())
```

`tests/test_sanitize_comissoes.py`:

```python
import datetime as dt

import openpyxl
import pytest

from papaki_dre.config import R_BLANK, R_FOOTER, R_H, R_MODELO, R_PESQ
from papaki_dre.headers import load_headers
from papaki_dre.sanitize.comissoes import sanitize_comissoes
from papaki_dre.types import ContractError, Formula
from xlsxcache import inject_cache

H = load_headers()["comissoes"]
MODELO = "Untitled - Publicidade Padrão"


def _build(tmp_path, data_rows, header=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "BASE DE COMISSÕES POR JOB — REAL x ORÇADO"
    for j, name in enumerate(header or H, start=1):
        ws.cell(3, j, name)
    for i, row in enumerate(data_rows, start=4):
        for j, v in enumerate(row, start=1):
            ws.cell(i, j, v)
    p = tmp_path / "com.xlsx"
    wb.save(p)
    return p


def _line(r, modelo=MODELO, titulo="Job X", comissao="Diretor", forn="ACME", orcado=100, real=0):
    return [modelo, "#75131", dt.datetime(2026, 9, 22), "Fulano", "Cliente", "Agência", titulo,
            comissao, forn, orcado, real, f"=IF(K{r}=0,J{r},0)", f"=J{r}-K{r}",
            f"=IF(J{r}=0,0,(J{r}-K{r})/J{r})"]


def test_keeps_matching_rows_and_preserves_types(tmp_path):
    s = sanitize_comissoes(_build(tmp_path, [_line(4)]))
    assert s.kept == 1 and s.balanced()
    assert s.rows[0][2] == dt.datetime(2026, 9, 22)
    assert s.rows[0][9] == 100


def test_formulas_are_readdressed_to_destination_rows(tmp_path):
    rows = [_line(4, modelo="Plena - Publicidade Padrão"), _line(5), _line(6)]
    s = sanitize_comissoes(_build(tmp_path, rows))
    assert s.kept == 2
    first, second = s.rows
    assert isinstance(first[11], Formula)
    assert first[11] == "=IF(K2=0,J2,0)"            # linha de destino 2
    assert second[11] == "=IF(K3=0,J3,0)"
    assert second[12] == "=J3-K3"
    assert s.source_rows == [5, 6]


def test_db_rows_have_values_not_formulas(tmp_path):
    p = _build(tmp_path, [_line(4, orcado=100, real=0)])
    inject_cache(p, {"L4": 100, "M4": 100, "N4": 1})      # o que o Excel teria calculado
    s = sanitize_comissoes(p)
    assert s.db_rows[0][11:14] == [100, 100, 1]
    assert all(isinstance(v, Formula) for v in s.rows[0][11:14])
    assert not s.warnings


def test_excludes_other_models(tmp_path):
    s = sanitize_comissoes(_build(tmp_path, [_line(4, modelo="Vision - Publicidade Padrão ")]))
    assert s.kept == 0 and [e.reason for e in s.exclusions] == [R_MODELO]


def test_model_comparison_ignores_surrounding_spaces(tmp_path):
    s = sanitize_comissoes(_build(tmp_path, [_line(4, modelo=MODELO + " ")]))
    assert s.kept == 1


def test_excludes_pesquisa_e_tratamento(tmp_path):
    s = sanitize_comissoes(_build(tmp_path, [_line(4, titulo="x | PESQUISA E TRATAMENTO | y")]))
    assert [e.reason for e in s.exclusions] == [R_PESQ]


@pytest.mark.parametrize("opcao", ["Seguro Produção", "Taxa da produtora", "Comissão total 1",
                                   "Comissão total 2", "Impostos", "Taxa ADM"])
def test_excludes_the_six_column_h_options(tmp_path, opcao):
    s = sanitize_comissoes(_build(tmp_path, [_line(4, comissao=opcao)]))
    assert [e.reason for e in s.exclusions] == [R_H]


def test_new_column_h_options_are_preserved(tmp_path):
    s = sanitize_comissoes(_build(tmp_path, [_line(4, comissao="Nany"), _line(5, comissao="Opção Nova")]))
    assert s.kept == 2


def test_reason_precedence_model_then_pesquisa_then_h(tmp_path):
    row = _line(4, modelo="Outro", titulo="PESQUISA E TRATAMENTO", comissao="Impostos")
    s = sanitize_comissoes(_build(tmp_path, [row]))
    assert [e.reason for e in s.exclusions] == [R_MODELO]


def test_reason_precedence_pesquisa_before_h(tmp_path):
    row = _line(4, titulo="x PESQUISA E TRATAMENTO", comissao="Impostos")
    s = sanitize_comissoes(_build(tmp_path, [row]))
    assert [e.reason for e in s.exclusions] == [R_PESQ]


def test_blank_rows_are_technical(tmp_path):
    s = sanitize_comissoes(_build(tmp_path, [_line(4), [None] * 14, _line(6)]))
    assert [t.reason for t in s.technical] == [R_BLANK]


def _footer(first, last):
    """Formato do bruto real: coluna A vazia, rótulo em H e somas em J:N."""
    row = [None] * 14
    row[7] = "TOTAL GERAL"
    for j, col in enumerate("JKLM", start=9):
        row[j] = f"=SUM({col}{first}:{col}{last})"
    row[13] = f"=IF(J{last + 2}=0,0,(J{last + 2}-K{last + 2})/J{last + 2})"
    return row


def test_total_geral_footer_is_technical_not_model_exclusion(tmp_path):
    s = sanitize_comissoes(_build(tmp_path, [_line(4), _line(5), [None] * 14, _footer(4, 5)]))
    assert s.kept == 2
    assert [t.reason for t in s.technical] == [R_BLANK, R_FOOTER]
    assert not s.exclusions
    assert s.balanced()


def test_total_geral_outside_column_h_is_not_footer(tmp_path):
    row = _footer(4, 4)
    row[7], row[6] = None, "TOTAL GERAL"
    s = sanitize_comissoes(_build(tmp_path, [_line(4), row]))
    assert R_FOOTER not in [t.reason for t in s.technical]


def test_unexpected_cached_error_in_kept_row_blocks(tmp_path):
    p = _build(tmp_path, [_line(4)])
    inject_cache(p, {"L4": "#REF!", "M4": 100, "N4": 1})
    with pytest.raises(ContractError, match="#REF!"):
        sanitize_comissoes(p)


def test_cached_error_in_excluded_row_does_not_block(tmp_path):
    p = _build(tmp_path, [_line(4, modelo="Outro")])
    inject_cache(p, {"L4": "#REF!"})
    assert sanitize_comissoes(p).kept == 0


def test_expected_div0_in_kept_row_becomes_null(tmp_path):
    p = _build(tmp_path, [_line(4)])
    inject_cache(p, {"L4": 100, "M4": 100, "N4": "#DIV/0!"})
    s = sanitize_comissoes(p)
    assert s.db_rows[0][13] is None
    assert s.rows[0][13] == "=IF(J2=0,0,(J2-K2)/J2)"


def test_formula_without_cache_warns_and_is_null(tmp_path):
    s = sanitize_comissoes(_build(tmp_path, [_line(4)]))      # openpyxl não grava cache
    assert s.db_rows[0][11:14] == [None, None, None]
    assert any("cache" in w for w in s.warnings)


def test_filled_header_beyond_ncols_stops(tmp_path):
    p = _build(tmp_path, [_line(4)])
    wb = openpyxl.load_workbook(p)
    wb.active.cell(3, 20, "Coluna extra")
    wb.save(p)
    with pytest.raises(ContractError, match="extra|além"):
        sanitize_comissoes(p)


def test_empty_cells_beyond_ncols_are_fine(tmp_path):
    p = _build(tmp_path, [_line(4)])
    wb = openpyxl.load_workbook(p)
    wb.active.cell(3, 20, "   ")
    wb.active.cell(4, 30, None)
    wb.save(p)
    assert sanitize_comissoes(p).kept == 1


def test_header_mismatch_stops(tmp_path):
    bad = H[:-1] + ["Var X"]
    with pytest.raises(ContractError):
        sanitize_comissoes(_build(tmp_path, [_line(4)], header=bad))


@pytest.mark.real
def test_real_september_file(comissoes_real):
    s = sanitize_comissoes(comissoes_real)
    assert s.balanced()
    assert s.raw_count == 1222
    assert s.kept == 649
    assert sorted(t.reason for t in s.technical) == [R_BLANK, R_FOOTER]
    assert len([e for e in s.exclusions if e.reason == R_MODELO]) == 81
    assert all(isinstance(r[11], Formula) for r in s.rows)
```

- [x] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_sanitize_comissoes.py -v`
Expected: FAIL com `ModuleNotFoundError`

- [x] **Step 3: Implementar `sanitize/_xlsx.py`**

```python
"""Partes comuns dos sanitizadores de planilha (Comissões e RRA)."""
from __future__ import annotations

from openpyxl.utils import get_column_letter

from ..formulas import readdress
from ..types import ContractError, XlError

EXPECTED_ERRORS = {"#DIV/0!"}


def blank(values) -> bool:
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in values)


def check_header_extent(report: str, header_cells, ncols: int) -> None:
    """Célula de cabeçalho preenchida além das `ncols` colunas do contrato para a rodada.
    Células vazias além (o bruto do RRA vai até a coluna 72) não contam."""
    extra = [c for c in header_cells[ncols:] if not blank([c.value])]
    if extra:
        refs = ", ".join(c.coordinate for c in extra)
        raise ContractError(f"{report}: cabeçalho com coluna extra além da {ncols}ª: {refs}")


class RowCopier:
    """Copia uma linha mantida célula a célula, com a mesma política nos dois relatórios:
    - erro em cache diferente de #DIV/0! -> ContractError;
    - fórmula -> reendereçada para a linha de destino; no banco vai o valor em cache
      (erro esperado vira None; sem cache vira None e gera um aviso agregado);
    - erro literal -> XlError na planilha, None no banco."""

    def __init__(self, report: str):
        self.report = report
        self.missing_cache: list[str] = []

    def copy(self, cells, vcells, r: int, dest: int, label) -> tuple[list, list]:
        row, db_row = [], []
        for c, vc in zip(cells, vcells):
            letter = get_column_letter(c.column)
            if vc.data_type == "e" and str(vc.value) not in EXPECTED_ERRORS:
                raise ContractError(
                    f"{self.report}: erro inesperado {vc.value} em {letter}{r} (job {label!r})")
            v = c.value
            if isinstance(v, str) and v.startswith("="):
                row.append(readdress(v, f"{letter}{r}", f"{letter}{dest}"))
                if vc.data_type == "e":
                    db_row.append(None)
                else:
                    if vc.value is None:
                        self.missing_cache.append(f"{letter}{r}")
                    db_row.append(vc.value)
            elif c.data_type == "e":
                row.append(XlError(v))
                db_row.append(None)
            else:
                row.append(v)
                db_row.append(v)
        return row, db_row

    def warnings(self) -> list[str]:
        if not self.missing_cache:
            return []
        n, first = len(self.missing_cache), self.missing_cache[0]
        return [f"{self.report}: {n} fórmula(s) de linhas mantidas sem valor em cache "
                f"(ex.: {first}); gravadas como nulo no banco"]
```

- [x] **Step 4: Implementar `sanitize/comissoes.py`**

```python
from __future__ import annotations

from ..config import (COMISSOES_EXCLUIR_H, COMISSOES_FOOTER, DEST, MODELO_COMISSOES, PESQUISA,
                      R_BLANK, R_FOOTER, R_H, R_MODELO, R_PESQ)
from ..headers import check_header, load_headers, normalize
from ..readers import open_xlsx
from ..types import Removal, Sanitized
from ._xlsx import RowCopier, blank, check_header_extent

REPORT = "comissoes"
NCOLS = 14
HEADER_ROW = 3
FIRST_RAW_ROW = 4
_EXCLUIR_H = {normalize(x).casefold() for x in COMISSOES_EXCLUIR_H}


def _is_footer(values) -> bool:
    """Linha final TOTAL GERAL do bruto: coluna A vazia, rótulo exato na coluna H e só
    somas em J:N. Reconhecida para ser descartada como técnica; os valores não são lidos."""
    label = values[7]
    return (isinstance(label, str) and normalize(label) == COMISSOES_FOOTER
            and blank(values[:7]) and blank(values[8:9]))


def sanitize_comissoes(path, *, expected_header: list[str] | None = None) -> Sanitized:
    ws = open_xlsx(path).worksheets[0]
    wv = open_xlsx(path, data_only=True).worksheets[0]
    header_cells = ws[HEADER_ROW]
    header = [normalize(c.value) if c.value is not None else "" for c in header_cells[:NCOLS]]
    check_header(REPORT, header, expected_header or load_headers()[REPORT])
    check_header_extent(REPORT, header_cells, NCOLS)

    first_dest = DEST[REPORT]["first_row"]
    out = Sanitized(report=REPORT, header=header, rows=[], db_rows=[], source_rows=[],
                    raw_count=max(ws.max_row - HEADER_ROW, 0), title=ws["A1"].value)
    copier = RowCopier(REPORT)
    seen_rows = 0
    # iter_rows com max_col fixo: `ws[r]` recalcula max_column a cada linha (quadrático).
    last = ws.max_row
    pairs = zip(ws.iter_rows(min_row=FIRST_RAW_ROW, max_row=last, max_col=NCOLS),
                wv.iter_rows(min_row=FIRST_RAW_ROW, max_row=last, max_col=NCOLS))
    for r, (cells, vcells) in enumerate(pairs, start=FIRST_RAW_ROW):
        values = [c.value for c in cells]
        seen_rows += 1
        if blank(values):
            out.technical.append(Removal(r, R_BLANK))
            continue
        if _is_footer(values):
            out.technical.append(Removal(r, R_FOOTER))
            continue
        modelo = normalize(values[0]) if values[0] is not None else ""
        titulo = str(values[6] or "").upper()
        opcao = normalize(values[7]).casefold() if values[7] is not None else ""
        if modelo != MODELO_COMISSOES:
            out.exclusions.append(Removal(r, R_MODELO))
            continue
        if PESQUISA in titulo:
            out.exclusions.append(Removal(r, R_PESQ))
            continue
        if opcao in _EXCLUIR_H:
            out.exclusions.append(Removal(r, R_H))
            continue
        dest = first_dest + len(out.rows)
        row, db_row = copier.copy(cells, vcells, r, dest, values[1])
        out.rows.append(row)
        out.db_rows.append(db_row)
        out.source_rows.append(r)
    out.raw_count = seen_rows
    out.warnings.extend(copier.warnings())
    return out
```

- [x] **Step 5: Rodar e ver passar**

Run: `python -m pytest tests/test_sanitize_comissoes.py -v`
Expected: 26 passed (o teste real confirma 649 linhas mantidas).

Se o teste real não fechar em 649: o número vem da comparação de 08/10/2026 e foi confirmado no bruto real: 1.222 linhas a partir da linha 4 = 649 mantidas + 2 técnicas (1 vazia + o rodapé `TOTAL GERAL`) + 81 de outro modelo + 126 de `PESQUISA E TRATAMENTO` + 364 da coluna H. Investigue antes de mudar o número.

- [x] **Step 6: Commit**

```bash
git add src/papaki_dre/sanitize/comissoes.py tests/test_sanitize_comissoes.py && git commit -m "feat: sanitizador de comissões com fórmulas reendereçadas e lista explícita da coluna H"
```

---

### Task 10: Sanitizador do RRA

Commit: 72b19ee (ajustado depois em 9984a05, ddffb42, 84d3786 e 8d85af7)

**Files:**
- Create: `src/papaki_dre/sanitize/rra.py`
- Modify: `src/papaki_dre/sanitize/__init__.py`
- Test: `tests/test_sanitize_rra.py`

Regras: XLS binário → XLSX real (Excel COM, via `convert_xls_to_xlsx`); aba `Plan1`; título na linha 1, cabeçalho na linha 3 (49 colunas; célula de cabeçalho preenchida além da 49ª → `ContractError`, vazias além são ignoradas — o bruto vai até a coluna 72), dados a partir da linha 4 (a 4 costuma ser vazia; a 5 é o primeiro job); linha com coluna A igual a `TOTAL` é técnica (`R_FOOTER`) e seus valores não são lidos; manter `Núcleo = Projetos - Untitled`; excluir `Título do job` com `PESQUISA E TRATAMENTO`; preservar fórmulas (reendereçadas para `BASE` linha 5 em diante) e o `#DIV/0!` esperado; **bloquear** qualquer outro erro nas linhas mantidas; nunca trocar erro por "Não calculável". A cópia de linha é a mesma das Comissões (`RowCopier`, Task 9). A leitura usa `iter_rows(max_col=NCOLS)`.

- [x] **Step 1: Escrever os testes que falham**

Os testes unitários criam um XLSX sintético (sem Excel) e chamam o núcleo `sanitize_rra_xlsx`; letras de coluna vêm de `get_column_letter`; valores em cache são injetados com `tests/xlsxcache.py`. O teste real chama `sanitize_rra` com o XLS (precisa do Excel).

```python
import datetime as dt

import openpyxl
import pytest
from openpyxl.utils import get_column_letter

from papaki_dre.config import NUCLEO_RRA, R_BLANK, R_FOOTER, R_NUCLEO, R_PESQ
from papaki_dre.headers import load_headers
from papaki_dre.sanitize.rra import sanitize_rra, sanitize_rra_xlsx
from papaki_dre.types import ContractError, Formula
from xlsxcache import inject_cache

H = load_headers()["rra"]
I_JOB, I_NUCLEO, I_TITULO = 0, H.index("Núcleo"), H.index("Título do job")
I_FAT = H.index("VALOR FATURADO")


def _job(r, job="75131", nucleo=NUCLEO_RRA, titulo="Job X", fat=100.0, formula=True):
    row = [None] * len(H)
    row[I_JOB], row[1] = job, dt.datetime(2026, 3, 17)
    row[I_NUCLEO], row[I_TITULO], row[I_FAT] = nucleo, titulo, fat
    if formula:
        row[I_FAT + 1] = f"={get_column_letter(I_FAT + 1)}{r}*2"      # coluna ao lado do VALOR FATURADO
    return row


def _build(tmp_path, jobs, total=True, header=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Plan1"
    ws["A1"] = "RENDIMENTOS DOS JOBS"
    for j, name in enumerate(header or H, start=1):
        ws.cell(3, j, name)
    r = 5
    for row in jobs:
        for j, v in enumerate(row, start=1):
            ws.cell(r, j, v)
        r += 1
    if total:
        ws.cell(r, 1, "TOTAL")
        col = get_column_letter(I_FAT + 1)
        ws[f"{col}{r}"] = f"=SUM({col}5:{col}{r - 1})"
    p = tmp_path / "rra.xlsx"
    wb.save(p)
    return p


def test_keeps_only_untitled_nucleo(tmp_path):
    jobs = [_job(5), _job(6, job="1", nucleo="Projetos - Kaus"), _job(7, job="2")]
    s = sanitize_rra_xlsx(_build(tmp_path, jobs))
    assert s.kept == 2
    assert [e.reason for e in s.exclusions] == [R_NUCLEO]
    assert s.balanced()


def test_excludes_pesquisa_e_tratamento(tmp_path):
    s = sanitize_rra_xlsx(_build(tmp_path, [_job(5, titulo="x PESQUISA E TRATAMENTO y")]))
    assert s.kept == 0 and [e.reason for e in s.exclusions] == [R_PESQ]


def test_total_row_is_technical(tmp_path):
    s = sanitize_rra_xlsx(_build(tmp_path, [_job(5, fat=100.0), _job(6, fat=50.0)]))
    assert [t.reason for t in s.technical] == [R_BLANK, R_FOOTER]
    assert s.kept == 2 and s.balanced()
    assert not hasattr(s, "controls")


def test_formulas_are_readdressed_after_exclusions(tmp_path):
    jobs = [_job(5, job="1", nucleo="Projetos - Kaus"), _job(6, job="2"), _job(7, job="3")]
    s = sanitize_rra_xlsx(_build(tmp_path, jobs))
    col = get_column_letter(I_FAT + 1)
    f = s.rows[0][I_FAT + 1]
    assert isinstance(f, Formula)
    assert f == f"={col}5*2"                 # primeira linha mantida vai para a linha 5
    assert s.rows[1][I_FAT + 1] == f"={col}6*2"


def test_expected_div0_is_preserved_in_db_as_null(tmp_path):
    row = _job(5, formula=False)
    row[I_FAT + 1] = "=1/0"
    p = _build(tmp_path, [row])
    inject_cache(p, {f"{get_column_letter(I_FAT + 2)}5": "#DIV/0!"})   # cache t="e", como o Excel
    s = sanitize_rra_xlsx(p)                                          # não bloqueia
    assert s.kept == 1
    assert s.rows[0][I_FAT + 1] == "=1/0"
    assert s.db_rows[0][I_FAT + 1] is None
    assert not s.warnings


def test_unexpected_error_in_kept_row_blocks(tmp_path):
    row = _job(5, formula=False)
    ws_path = _build(tmp_path, [row])
    wb = openpyxl.load_workbook(ws_path)
    cell = wb["Plan1"].cell(5, I_FAT + 3)
    cell.value = "#REF!"
    cell.data_type = "e"
    wb.save(ws_path)
    with pytest.raises(ContractError, match="#REF!"):
        sanitize_rra_xlsx(ws_path)


def test_error_in_excluded_row_does_not_block(tmp_path):
    row = _job(5, nucleo="Projetos - Kaus", formula=False)
    p = _build(tmp_path, [row])
    wb = openpyxl.load_workbook(p)
    cell = wb["Plan1"].cell(5, I_FAT + 3)
    cell.value = "#REF!"
    cell.data_type = "e"
    wb.save(p)
    assert sanitize_rra_xlsx(p).kept == 0


def test_blank_row_is_technical(tmp_path):
    s = sanitize_rra_xlsx(_build(tmp_path, [_job(5)]))
    assert R_BLANK in [t.reason for t in s.technical]      # a linha 4 é vazia


def test_title_is_kept(tmp_path):
    assert sanitize_rra_xlsx(_build(tmp_path, [_job(5)])).title == "RENDIMENTOS DOS JOBS"


def test_unexpected_cached_error_in_formula_blocks(tmp_path):
    p = _build(tmp_path, [_job(5)])
    inject_cache(p, {f"{get_column_letter(I_FAT + 2)}5": "#N/A"})
    with pytest.raises(ContractError, match="#N/A"):
        sanitize_rra_xlsx(p)


def test_formula_without_cache_warns_and_is_null(tmp_path):
    s = sanitize_rra_xlsx(_build(tmp_path, [_job(5)]))
    assert s.db_rows[0][I_FAT + 1] is None
    assert any("cache" in w for w in s.warnings)


def test_filled_header_beyond_ncols_stops(tmp_path):
    p = _build(tmp_path, [_job(5)])
    wb = openpyxl.load_workbook(p)
    wb["Plan1"].cell(3, len(H) + 5, "Coluna extra")
    wb.save(p)
    with pytest.raises(ContractError, match="extra|além"):
        sanitize_rra_xlsx(p)


def test_empty_cells_beyond_ncols_are_fine(tmp_path):
    p = _build(tmp_path, [_job(5)])
    wb = openpyxl.load_workbook(p)
    wb["Plan1"].cell(3, 72, "")
    wb["Plan1"].cell(6, 72, None)
    wb.save(p)
    assert sanitize_rra_xlsx(p).kept == 1


def test_header_mismatch_stops(tmp_path):
    bad = H[:-1] + ["Coluna Nova"]
    with pytest.raises(ContractError):
        sanitize_rra_xlsx(_build(tmp_path, [_job(5)], header=bad))


@pytest.mark.excel
@pytest.mark.real
def test_real_september_file(rra_real, tmp_path):
    s = sanitize_rra(rra_real, work_dir=tmp_path)
    assert s.balanced()
    assert s.kept == 72
    assert len([e for e in s.exclusions if e.reason == R_PESQ]) == 9
    assert len(s.exclusions) == (1410 - 81) + 9
    assert s.raw_count == 2408
    assert len(s.technical) == 998
    assert [t.reason for t in s.technical].count(R_FOOTER) == 1
    assert not s.warnings
    assert any(isinstance(v, Formula) for v in s.rows[0])
```

- [x] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_sanitize_rra.py -v -m "not excel"`
Expected: FAIL com `ModuleNotFoundError`

- [x] **Step 3: Implementar `sanitize/rra.py`**

```python
from __future__ import annotations

from pathlib import Path

import openpyxl

from ..config import DEST, NUCLEO_RRA, PESQUISA, R_BLANK, R_FOOTER, R_NUCLEO, R_PESQ, RRA_FOOTER
from ..headers import check_header, load_headers, normalize
from ..readers import convert_xls_to_xlsx
from ..types import Removal, Sanitized
from ._xlsx import RowCopier, blank, check_header_extent

REPORT = "rra"
NCOLS = 49
HEADER_ROW = 3
FIRST_RAW_ROW = 4


def sanitize_rra(xls_path, *, work_dir, expected_header: list[str] | None = None) -> Sanitized:
    """XLS do Esferas -> XLSX (Excel COM, preserva fórmulas) -> sanitização."""
    xlsx = convert_xls_to_xlsx(xls_path, Path(work_dir) / (Path(xls_path).stem + ".xlsx"))
    return sanitize_rra_xlsx(xlsx, expected_header=expected_header)


def sanitize_rra_xlsx(xlsx_path, *, expected_header: list[str] | None = None) -> Sanitized:
    wf = openpyxl.load_workbook(xlsx_path)
    wv = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wf["Plan1"] if "Plan1" in wf.sheetnames else wf.worksheets[0]
    wsv = wv[ws.title]
    header_cells = ws[HEADER_ROW]
    header = [normalize(c.value) if c.value is not None else "" for c in header_cells[:NCOLS]]
    check_header(REPORT, header, expected_header or load_headers()[REPORT])
    check_header_extent(REPORT, header_cells, NCOLS)
    i_nucleo = header.index("Núcleo")
    i_titulo = header.index("Título do job")
    first_dest = DEST[REPORT]["first_row"]

    out = Sanitized(report=REPORT, header=header, rows=[], db_rows=[], source_rows=[],
                    raw_count=0, title=ws["A1"].value)
    copier = RowCopier(REPORT)

    # iter_rows com max_col fixo: `ws[r]` recalcula max_column a cada linha (quadrático).
    last = ws.max_row
    pairs = zip(ws.iter_rows(min_row=FIRST_RAW_ROW, max_row=last, max_col=NCOLS),
                wsv.iter_rows(min_row=FIRST_RAW_ROW, max_row=last, max_col=NCOLS))
    for r, (cells, vcells) in enumerate(pairs, start=FIRST_RAW_ROW):
        values = [c.value for c in cells]
        out.raw_count += 1
        if blank(values):
            out.technical.append(Removal(r, R_BLANK))
            continue
        if str(values[0]).strip().upper() == RRA_FOOTER:     # rodapé: técnico, valores não lidos
            out.technical.append(Removal(r, R_FOOTER))
            continue
        nucleo = normalize(values[i_nucleo]) if values[i_nucleo] is not None else ""
        if nucleo != NUCLEO_RRA:
            out.exclusions.append(Removal(r, R_NUCLEO))
            continue
        if PESQUISA in str(values[i_titulo] or "").upper():
            out.exclusions.append(Removal(r, R_PESQ))
            continue

        dest = first_dest + len(out.rows)
        row, db_row = copier.copy(cells, vcells, r, dest, values[0])
        out.rows.append(row)
        out.db_rows.append(db_row)
        out.source_rows.append(r)
    out.warnings.extend(copier.warnings())
    return out
```

- [x] **Step 4: Atualizar `sanitize/__init__.py` para exportar os quatro**

```python
from .comissoes import sanitize_comissoes
from .cp import sanitize_cp
from .cr import sanitize_cr
from .rra import sanitize_rra, sanitize_rra_xlsx

__all__ = ["sanitize_cp", "sanitize_cr", "sanitize_comissoes", "sanitize_rra", "sanitize_rra_xlsx"]
```

- [x] **Step 5: Rodar e ver passar**

Run: `python -m pytest tests/test_sanitize_rra.py -v`
Expected: 15 passed (o teste `excel`/`real` converte o XLS real; leva alguns segundos). Sem Excel, rode `-m "not excel"` e registre a ressalva.

Se o teste real divergir: pare e investigue; não ajuste o número. Contagem confirmada no bruto real: 2.408 linhas a partir da linha 4 = 72 mantidas + 998 técnicas (997 vazias + 1 rodapé `TOTAL`) + 1.329 de outro núcleo + 9 de `PESQUISA E TRATAMENTO`, sem avisos.

- [x] **Step 6: Rodar a suíte completa**

Run: `python -m pytest -v`
Expected: 138 passed (testes `real`/`excel` pulam se os brutos ou o Excel estiverem ausentes). Contagens reais obtidas: CP 1.110, CR 484, Comissões 649, RRA 72.

- [x] **Step 7: Commit**

```bash
git add src/papaki_dre/sanitize tests/test_sanitize_rra.py && git commit -m "feat: sanitizador do RRA (XLS via Excel, fórmulas preservadas, erro inesperado bloqueia)"
```

---

## Self-Review

**Cobertura da spec §6:** leitura com codificação por arquivo — UTF-8 com BOM no CP, Windows-1252 no CR (T5); validação de cabeçalho e nº de colunas, que para a rodada, inclusive coluna extra preenchida no cabeçalho de Comissões/RRA e registro de CSV com largura inesperada no CP/CR (T6, T7–T10); remoção só de linhas técnicas por regra — vazias e linhas de TOTAL reconhecidas por rótulo e posição exatos (CP `VALOR TOTAL` no índice 4; CR `TOTAL PÁGINA`/`TOTAL GERAL` no índice 6; RRA coluna A `TOTAL`; Comissões `TOTAL GERAL` na coluna H com A..G e I vazias), sem leitura de valores do rodapé (T7–T10); tipagem, vazio ≠ zero, identificadores como texto com zeros à esquerda, lista fechada de texto em colunas de dinheiro (T3, T7, T8); exclusões de negócio com motivo e linha de origem rastreados, e balanço `raw_count = mantidas + técnicas + exclusões` (T2, T7–T10); CR sem exclusão de linhas (T8); `#DIV/0!` preservado, erro inesperado bloqueia em Comissões e RRA, nunca "Não calculável" (T9, T10); fórmulas L:N das Comissões e do RRA reendereçadas, com valores em cache no `db_rows` e aviso agregado quando falta cache (T4, T9, T10); XLS → XLSX real com cache por hash do conteúdo e gravação atômica (T5, T10); `autoFilter` inválido (T5). Spec §9 (aceite parcial): contagens de setembro 72/1.110/484/649 como testes reais, com a decomposição completa de cada bruto.

**Fora deste plano (planos 2 e 3):** carga no Excel, linha de total da `BASE`, extensão de fórmulas `AP:AQ`/`AC`/`O:Q`, relatório de conferência, banco, reimportação, CLI.

**Consistência de nomes:** `Sanitized`, `Removal`, `Formula`, `XlError`, `ContractError` (T2); `sanitize_cp/cr/comissoes/rra/rra_xlsx` (T7–T10); `RowCopier`, `blank`, `check_header_extent`, `EXPECTED_ERRORS` (`sanitize/_xlsx.py`, T9); `read_delimited`, `open_xlsx`, `convert_xls_to_xlsx`, `xlsx_cache_path`, `com_retry` (T5); `load_headers`, `check_header`, `normalize` (T6); `readdress` (T4); `coerce_cell`, `parse_money`, `MONEY_TEXT_OK` (T3); constantes `R_*`, `DEST`, `CP_KINDS`, `CR_KINDS`, `CR_INDEX_KINDS`, `CP_FOOTER`, `CR_FOOTER`, `RRA_FOOTER`, `COMISSOES_FOOTER` (T2); helper de teste `inject_cache` (`tests/xlsxcache.py`).

---

## Rodada de correções pós-revisão (A–G)

Depois dos commits das Tasks 1–10, a revisão gerou sete commits de correção, todos já refletidos no código e nos testes acima:

- **A — 9984a05** `fix:` rodapés de TOTAL viram linhas técnicas por rótulo e posição exatos; os totais de controle foram eliminados por decisão do usuário (09/10/2026): saíram `ControlTotal`, `Sanitized.controls` e as constantes de controle do CR (`CR_CONTROL_COLUMNS`, `CR_FOOTER_SHIFT`) e `TOLERANCIA_CONTROLE`; o rodapé das Comissões deixou de cair em `modelo_diferente` (`R_MODELO` 82 → 81 no real).
- **B — fb97ad2** `fix:` CP e CR param a rodada em registro não vazio, não rodapé e com nº de colunas inesperado; o motivo `R_ODD` (descarte com aviso) foi removido.
- **C — dd748ec** `fix:` conversão XLS com destino nomeado pelo hash do conteúdo, gravação atômica (`_materialize` + `os.replace`), limpeza segura do Excel e `UpdateLinks=0`, no lugar do reaproveitamento por data de modificação.
- **D — b84919c** `fix:` `money_or_text` aceita como texto só a lista fechada (`MONEY_TEXT_OK = {"-"}`) e `parse_money` rejeita sinal duplo.
- **E — ddffb42** `fix:` cópia de linha comum a Comissões e RRA (`sanitize/_xlsx.py`); erro inesperado em cache bloqueia nas duas; fórmula sem cache vira nulo com aviso agregado; cabeçalho com coluna extra preenchida para a rodada; helper `tests/xlsxcache.py`.
- **F — 84d3786** `test:` testes que passavam pelo motivo errado foram corrigidos (cache injetado, `#DIV/0!` real, `get_column_letter`, zeros à esquerda, valores negativos); CP e CR passam a guardar listas distintas em `rows` e `db_rows`.
- **G — 8d85af7** `perf:` RRA e Comissões percorrem as linhas com `iter_rows(max_col=NCOLS)` (RRA 91,6 s → 5,2 s; Comissões 4,3 s → 1,1 s, com resultados idênticos); o teste real do RRA passa a conferir raw 2.408, 998 técnicas (1 rodapé) e ausência de avisos.
