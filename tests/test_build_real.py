import openpyxl
import pytest

from papaki_dre.build import build_workbook
from papaki_dre.sanitize import sanitize_comissoes, sanitize_cp, sanitize_cr, sanitize_rra
from papaki_dre.xlsx_inject import WorkbookSession
from conftest import BRUTO, REAL_FILES, TEMPLATE

pytestmark = [pytest.mark.real, pytest.mark.excel]

# As contagens fixas abaixo (72/1110/484/649 linhas, órfãs, 439 constantes) valem para a extração de
# 2026-09 feita em 08/10/2026 (amostras/2026-09/bruto) e o modelo-base de setembro. Outra extração
# muda os números: não ajuste para passar sem entender a diferença.


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
    from papaki_dre.produtoras import carregar
    report = build_workbook(TEMPLATE, out, results, master_formula_rows=carregar("untitled").linha_mestre)
    return results, out, report


def test_counts_match_the_sanitizers(built):
    # extração de 2026-09 de 08/10
    results, _, report = built
    assert report.sheets["BASE"].rows == results["rra"].kept == 72
    assert report.sheets["DESPESAS"].rows == results["cp"].kept == 1110
    assert report.sheets["RECEITA"].rows == results["cr"].kept == 484
    assert report.sheets["REAL X ORÇADO COMISSÕES"].rows == results["comissoes"].kept == 649


def test_only_base_sheets_and_workbook_plumbing_changed(built):
    _, out, report = built
    with WorkbookSession(out) as sess:
        names = {sess.parts[n] for n in ("BASE", "DESPESAS", "RECEITA", "REAL X ORÇADO COMISSÕES")}
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
    # extração de 2026-09 de 08/10: fórmulas do modelo até 1159 (DESPESAS) e 664 (RECEITA)
    _, _, report = built
    assert report.sheets["DESPESAS"].orphan_formula_rows == 1159 - (2 + 1110)
    assert report.sheets["RECEITA"].orphan_formula_rows == 664 - (1 + 484)


def test_typed_constants_replaced_and_filters_untouched(built):
    # medido em 13/10/2026 para a extração de 08/10: 439 constantes digitadas em DESPESAS!AP:AQ dentro
    # das linhas carregadas viram a fórmula do modelo (decisão opção a); nenhuma nas outras abas.
    # Os filtros do modelo (DESPESAS A2:AR1159, RECEITA A1:AC664) já cobrem as linhas carregadas.
    _, _, report = built
    assert {k: sr.overwritten_constants for k, sr in report.sheets.items()} == {
        "BASE": 0, "DESPESAS": 439, "RECEITA": 0, "REAL X ORÇADO COMISSÕES": 0}
    assert all(sr.missing_formulas == [] for sr in report.sheets.values())
    assert not any(sr.autofilter_extended for sr in report.sheets.values())


def test_complete_classification_formula_in_every_ap_row(built):
    # decisão do usuário de 13/10/2026 (linha_mestre da Untitled): DESPESAS!AP inteira (3..1159,
    # inclusive as órfãs 1113..1159) com a fórmula COMPLETA da linha 881 do modelo-base, traduzida
    from papaki_dre.xlsx_inject import translate_formula
    _, out, report = built
    with WorkbookSession(TEMPLATE) as tpl:
        full = tpl.editor("DESPESAS").formula_at(881, 42)
    assert "Transferência entre Empresas Coligadas" in full
    with WorkbookSession(out) as sess:
        ed = sess.editor("DESPESAS")
        for r in (3, 500, 880, 881, 1070, 1071, 1112, 1159):
            assert ed.formula_at(r, 42) == translate_formula(full, "AP881", f"AP{r}"), r
    # medido em 13/10/2026 (extração de 08/10): 1036 fórmulas em AP3:AP1159, menos as 190 já completas
    # (AP881:AP1070); constantes seguem só em overwritten_constants (439)
    assert report.sheets["DESPESAS"].unified_formula_rows == 846
    assert all(sr.unified_formula_rows == 0 for k, sr in report.sheets.items() if k != "DESPESAS")


def test_no_orphan_shared_formula_children_in_despesas(built):
    import zipfile
    from xlsx_factory import orphan_shared_children
    _, out, _ = built
    with WorkbookSession(out) as sess:
        part = sess.parts["DESPESAS"]
    with zipfile.ZipFile(out) as z:
        assert orphan_shared_children(z.read(part)) == []


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
    from papaki_dre.readers import com_retry, com_tipado
    excel = com_retry(lambda: win32com.client.DispatchEx("Excel.Application"))
    com_retry(lambda: setattr(excel, "Visible", False))
    com_retry(lambda: setattr(excel, "DisplayAlerts", False))
    try:
        wb = com_tipado(com_retry(lambda: excel.Workbooks.Open(str(out), ReadOnly=True)))
        try:
            ws = com_retry(lambda: wb.Worksheets("DESPESAS"))
            assert com_retry(lambda: ws.Range("A3").Value) == results["cp"].rows[0][0]
            assert str(com_retry(lambda: ws.Range("AP3").Formula)).startswith("=")
            assert com_retry(lambda: wb.Worksheets.Count) == 17
        finally:
            com_retry(lambda: wb.Close(False))
    finally:
        com_retry(lambda: excel.Quit())
