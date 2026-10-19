from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.utils import column_index_from_string, get_column_letter

from .config import DEST
from .types import ContractError, Formula, Sanitized
from .xlsx_inject import FormulaShifter, WorkbookSession, changed_parts

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

# Primeira linha das somas do TOTAL da BASE: a linha 4 fica vazia entre o cabeçalho (3) e o 1º job (5),
# herdada do modelo de setembro/2026, e as somas a incluem (=SUM(J4:J<última+1>)).
BASE_SUM_START_ROW = 4


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
    overwritten_constants: int = 0                               # constantes e fórmulas de outra família substituídas pela fórmula de dados do modelo
    autofilter_extended: bool = False                            # o <autoFilter> da aba foi estendido até a última linha carregada
    # Colunas da `linha_mestre` da produtora (decisão de 13/10/2026): nº de linhas cuja fórmula de DADOS
    # existente (célula com <f>, outra versão da classificação) foi trocada pela fórmula completa da
    # linha configurada, inclusive nas linhas órfãs abaixo das carregadas. NÃO inclui constantes nem
    # fórmulas de outra família (resumos, <f/> vazio, ilegíveis), que seguem em overwritten_constants.
    # Informativo: o relatório (Plano 3) exibe, sem bloquear.
    unified_formula_rows: int = 0


@dataclass
class BuildReport:
    sheets: dict[str, SheetReport] = field(default_factory=dict)
    parts: dict[str, list[str]] = field(default_factory=dict)


def _check_empty(ed, sheet: str, rows, ncols: int) -> None:
    dirty = ed.nonempty_cells(rows, set(range(1, ncols + 1)))
    if dirty:
        raise ContractError(
            f"{sheet}: o modelo não está vazio na região de destino "
            f"({len(dirty)} células, ex.: {', '.join(dirty[:5])})")


def _load(session: WorkbookSession, key: str, data: Sanitized,
          master_rows: dict[str, int] | None = None) -> SheetReport:
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
        ed.write(lay["header_row"], j, name)
    for i, row in enumerate(data.rows):
        r = first + i
        for j, v in enumerate(row, start=1):
            ed.write(r, j, v, style_from=(first, j), date_styles=ds)

    report = SheetReport(sheet, first, last, n)
    for letter in FORMULA_COLUMNS.get(key, []):
        _extend_formulas(ed, key, sheet, letter, first, last, report,
                         master_row=(master_rows or {}).get(letter))
    if key == "rra":
        total_row = last + 2                       # estilo de cada célula: o da última linha de dados
        ed.write(total_row, 1, "TOTAL", style_from=(last, 1))
        for letter in SUM_COLUMNS:
            col = column_index_from_string(letter)
            ed.write(total_row, col, Formula(f"=SUM({letter}{BASE_SUM_START_ROW}:{letter}{last + 1})"),
                     style_from=(last, col))
        report.total_row = total_row
    if n:
        report.autofilter_extended = ed.extend_autofilter(last) is not None
    report.date_style_warnings = dict(ed.date_warnings)
    return report


def _extend_formulas(ed, key, sheet, letter, first, last, report, master_row: int | None = None) -> None:
    """Garante a fórmula de dados do modelo em `letter`, linhas first..last.

    Com `master_row` (coluna da `linha_mestre` da produtora) a regra é outra: ver
    `_unify_formulas` (a fórmula da linha configurada vale em todas as linhas da coluna). Sem ele:

    A base é a última fórmula da família de dados (`SheetEditor.formula_family`). Em cada linha
    carregada: fórmula de dados (da família, inclusive filhas compartilhadas e o mestre, ou variante
    com referências relativas) fica intacta; célula vazia recebe a fórmula; constante digitada OU
    fórmula de outra família (`SheetEditor.is_foreign_formula`: um resumo =COUNTA(AP3:AP11) na linha
    12 que caiu dentro da faixa, um resumo só com referências absolutas, <f/> vazio, fórmula
    ilegível, matriz de uma célula só) é SOBRESCRITA pela fórmula de dados e contada em
    `report.overwritten_constants`. Recusa (ContractError) trocar o mestre de uma fórmula
    compartilhada de outra família e qualquer célula coberta por uma fórmula de matriz de várias
    células (as filhas da matriz só têm valor).

    Atômico por coluna: todas as gravações são planejadas e validadas (`SheetEditor.check_write`)
    antes da primeira; se houver ContractError, nenhuma célula da coluna muda e nada é contado."""
    col = column_index_from_string(letter)
    family = ed.formula_family(col, first)
    plan = []                                         # (linha, fórmula, troca <f>?, conta?)
    if family is None:
        if key != "comissoes":
            raise ContractError(
                f"{sheet}!{letter}: classificação ausente — o modelo não tem fórmula para estender")
        # o modelo não traz O:Q: usar as de setembro
        shifter = FormulaShifter(COMISSOES_FORMULAS[letter], f"{letter}2", where=sheet)
        plan = [(r, Formula(shifter.to(f"{letter}{r}")), False, False) for r in range(first, last + 1)]
        style_from = last_f = None
    elif master_row is not None:
        last_f, signature = family
        _unify_formulas(ed, sheet, letter, first, last, master_row, signature, report)
    else:
        last_f, signature = family
        shifter = FormulaShifter(ed.formula_at(last_f, col), f"{letter}{last_f}", where=sheet)
        style_from = (last_f, col)
        for r in range(first, last + 1):
            has_f = ed.has_formula_element(r, col)
            if has_f and not ed.is_foreign_formula(r, col, first, signature):
                continue                              # fórmula de dados (família ou variante): intacta
            # constante digitada ou fórmula de outra família no lugar da fórmula de dados
            replaced = has_f or bool(ed.nonempty_cells([r], {col}))
            plan.append((r, Formula(shifter.to(f"{letter}{r}")), has_f, replaced))
    for r, formula, has_f, _ in plan:                 # valida tudo antes de mudar qualquer célula
        ed.check_write(r, col, formula, replace_formula=has_f)
    for r, formula, has_f, replaced in plan:
        ed.write(r, col, formula, style_from=style_from, replace_formula=has_f)
        if replaced:
            report.overwritten_constants += 1
    if last_f is not None and last_f > last:
        report.orphan_formula_rows = max(report.orphan_formula_rows, last_f - last)


def _unify_formulas(ed, sheet, letter, first, last, master_row, signature, report) -> None:
    """Coluna da `linha_mestre` da produtora (decisão do usuário de 13/10/2026): a fórmula da
    linha `master_row` do modelo (em DESPESAS!AP, a classificação COMPLETA da linha 881) é traduzida e
    gravada explicitamente em TODAS as linhas de `first` até max(`last`, última linha com fórmula de
    dados da coluna), inclusive nas órfãs abaixo das carregadas (no modelo real 3..1159), para que
    nenhum grupo de fórmula compartilhada fique sem mestre.

    Linhas que já têm exatamente a fórmula-base traduzida ficam intactas, salvo filhas de um grupo
    compartilhado cujo mestre é trocado (trocadas junto). As demais são trocadas: fórmula de dados de
    outra versão (com <f>) conta em `report.unified_formula_rows`; constante digitada ou fórmula de
    outra família (`SheetEditor.is_foreign_formula` com a assinatura `signature` da família) conta em
    `report.overwritten_constants`; célula vazia não conta. Um mestre compartilhado só é trocado se
    todas as células do seu `ref` (e todas as suas filhas) também forem (`SheetEditor.write_batch`).

    ContractError (nada muda na coluna): linha configurada sem fórmula, fórmula dela ilegível ou que
    não se traduz para `first`, grupo compartilhado não coberto por inteiro, célula em matriz."""
    col = column_index_from_string(letter)
    where = f"{sheet}!{letter}{master_row}"
    base = None
    if ed.has_formula_element(master_row, col):
        try:
            base = ed.formula_at(master_row, col)
        except ContractError as e:
            raise ContractError(f"{where}: a fórmula da linha configurada em MASTER_FORMULA_ROWS é "
                                f"ilegível ({e})") from e
    if base is None or not base[1:].strip():
        raise ContractError(f"{where}: a linha configurada em MASTER_FORMULA_ROWS não tem fórmula; "
                            "não há fórmula completa para aplicar à coluna")
    shifter = FormulaShifter(base, f"{letter}{master_row}", where=sheet)
    try:
        shifter.to(f"{letter}{master_row}")
    except ContractError as e:
        raise ContractError(f"{where}: a fórmula da linha configurada em MASTER_FORMULA_ROWS é "
                            f"ilegível ({e})") from e
    try:
        base_sig = shifter.to(f"{letter}{first}")
    except ContractError as e:
        raise ContractError(f"{where}: a fórmula da linha configurada em MASTER_FORMULA_ROWS não se "
                            f"traduz para a linha {first} ({e})") from e

    data_rows = [r for r in ed.formula_rows(col, first) if not ed.is_foreign_formula(r, col, first, signature)]
    end = max([last, *data_rows])
    keep = {r for r in range(first, end + 1)
            if ed.signature_at(r, col, first) == base_sig and ed.formula_at(r, col) == shifter.to(f"{letter}{r}")}
    dropped = {g[0] for r in range(first, end + 1)
               if r not in keep and (g := ed.shared_group(r, col)) is not None and g[1]}
    keep = {r for r in keep if (g := ed.shared_group(r, col)) is None or g[0] not in dropped}

    plan, unified, overwritten = [], 0, 0
    for r in range(first, end + 1):
        if r in keep:
            continue
        if ed.has_formula_element(r, col):
            if ed.is_foreign_formula(r, col, first, signature):
                overwritten += 1
            else:
                unified += 1
        elif ed.nonempty_cells([r], {col}):
            overwritten += 1
        plan.append((r, col, Formula(shifter.to(f"{letter}{r}"))))
    ed.write_batch(plan, style_from=(master_row, col))   # valida tudo antes de gravar
    report.unified_formula_rows += unified
    report.overwritten_constants += overwritten


def _total_gaps(ed, sr: SheetReport) -> list[str]:
    """Linha TOTAL da BASE: rótulo em A e =SUM(X<BASE_SUM_START_ROW>:X<última+1>) exata em cada
    coluna de SUM_COLUMNS."""
    t = sr.total_row
    if t is None:
        return ["BASE: linha TOTAL não registrada"]
    gaps = [] if ed.text_at(t, 1) == "TOTAL" else [f"A{t}"]
    for letter in SUM_COLUMNS:
        expected = f"=SUM({letter}{BASE_SUM_START_ROW}:{letter}{sr.last_row + 1})"
        col = column_index_from_string(letter)
        if not _has_formula(ed, t, col) or ed.formula_at(t, col) != expected:
            gaps.append(f"{letter}{t}")
    return gaps


def _has_formula(ed, r: int, col: int) -> bool:
    """Há fórmula de verdade: <f> com texto, ou filha compartilhada que resolve para uma fórmula.
    <f/> vazio (ou só com espaços) e filha sem mestre/intraduzível contam como lacuna."""
    try:
        f = ed.formula_at(r, col)
    except ContractError:
        return False
    return f is not None and f[1:].strip() != ""


def formula_coverage(path, key: str, sr: SheetReport, session: WorkbookSession | None = None) -> list[str]:
    """Lê o arquivo gravado e devolve as células sem a fórmula exigida (spec §8): COVERAGE_COLUMNS
    em todas as linhas carregadas e, na BASE, a linha TOTAL. Lista vazia = cobertura completa.
    `session`: sessão já aberta sobre `path` (reaproveitada entre as abas; quem a abriu a fecha)."""
    if session is None:
        with WorkbookSession(path) as own:
            return formula_coverage(path, key, sr, own)
    ed = session.editor(sr.sheet)
    missing = []
    for letter in COVERAGE_COLUMNS.get(key, []):
        col = column_index_from_string(letter)
        for r in range(sr.first_row, sr.last_row + 1):
            if not _has_formula(ed, r, col):
                missing.append(f"{letter}{r}")
    if key == "rra":
        missing.extend(_total_gaps(ed, sr))
    return missing


def build_workbook(template, out_path, results: dict[str, Sanitized], *,
                   master_formula_rows: dict[str, dict[str, int]] | None = None) -> BuildReport:
    """Copia o modelo e carrega as quatro bases. Nenhuma outra aba é tocada.
    master_formula_rows: colunas cuja fórmula completa vem de uma linha fixa do modelo (a `linha_mestre`
    do arquivo da produtora, ver `_unify_formulas`); sem ele, nenhuma coluna é unificada."""
    missing = [k for k in ORDER if k not in results]
    if missing:
        raise ContractError(f"faltam resultados para: {', '.join(missing)}")
    if master_formula_rows is None:
        master_formula_rows = {}
    out_path = Path(out_path)
    report = BuildReport()
    with WorkbookSession(template) as session:
        for key in ORDER:
            report.sheets[DEST[key]["sheet"]] = _load(session, key, results[key],
                                                      master_formula_rows.get(key))
        session.save(out_path)
    # lacuna de cobertura não levanta: vai para missing_formulas e a conferência posterior bloqueia
    with WorkbookSession(out_path) as written:
        for key in ORDER:
            sr = report.sheets[DEST[key]["sheet"]]
            sr.missing_formulas = formula_coverage(out_path, key, sr, session=written)
    report.parts = changed_parts(template, out_path)
    return report
