"""Gera um .xlsx mínimo, com XML escrito à mão, que imita o que importa do modelo real:
linhas esparsas, fórmulas compartilhadas, calcChain, uma aba que NÃO pode mudar."""
import re
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


# As duas versões da classificação de DESPESAS!AP (ver make_template, desp_ap_variant): a curta e a
# COMPLETA, que tem um ramo a mais na frente (como "Transferência entre Empresas Coligadas" no modelo).
AP_SHORT = 'IF(A{r}="x","a","b")'
AP_FULL = 'IF(B{r}="t","Desconsiderar",IF(A{r}="x","a","b"))'


def orphan_shared_children(sheet_xml: bytes) -> list[str]:
    """Células com fórmula compartilhada filha (<f t="shared" si=X/>, sem texto) cujo si não tem mestre
    (<f t="shared" si=X>texto</f>) na aba: o Excel repara/corrompe o arquivo. Vazia = ok."""
    from lxml import etree
    root = etree.fromstring(sheet_xml)
    fs = [(f.getparent().get("r"), f) for f in root.iter(f"{{{MAIN}}}f") if f.get("t") == "shared"]
    masters = {f.get("si") for _, f in fs if f.text}
    return [ref for ref, f in fs if not f.text and f.get("si") not in masters]


def _sheet(dim, rows_xml, tail=""):
    """tail: elementos depois de <sheetData> (ex.: <autoFilter>)."""
    return HEAD + (f'<worksheet xmlns="{MAIN}" xmlns:r="{RNS}"><dimension ref="{dim}"/>'
                   f'<sheetData>{rows_xml}</sheetData>{tail}</worksheet>')


def _styled_rows(rows, cells):
    """cells: lista de (coluna, estilo) para cada linha; linhas esparsas e com spans."""
    return "".join(
        f'<row r="{r}" spans="1:3">' + "".join(f'<c r="{c}{r}" s="{s}"/>' for c, s in cells) + "</row>"
        for r in rows)


def _const(ref):
    """Constante digitada à mão no lugar da fórmula (como o 'Desconsiderar' do modelo real)."""
    return f'<c r="{ref}" t="inlineStr"><is><t>Desconsiderar</t></is></c>'


def replace_part(src, dst, name, data, *, drop=False):
    """Copia o .xlsx `src` para `dst` trocando (ou removendo, com drop=True) a parte `name`."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename == name:
                if not drop:
                    zout.writestr(name, data)
            else:
                zout.writestr(info, zin.read(info.filename))
    return Path(dst)


def sheet_xml(rows_xml, dim="A1:A1"):
    """XML de uma aba qualquer (para trocar uma parte com replace_part)."""
    return _sheet(dim, rows_xml)


def make_template(path, *, desp_formula_rows=5, rec_formula_rows=3, calc_chain=True,
                  desp_constant_rows=(), rec_constant_rows=(), calc_pr='<calcPr calcId="191029"/>',
                  desp_footer_row=None, desp_autofilter=None, desp_ap_variant=None):
    """desp_ap_variant: (s, e) -> AP traz duas versões da classificação, como o modelo-base de setembro
    (AP881:AP1070): AP_FULL (com um ramo a mais) no grupo compartilhado si=1 em s..e e AP_SHORT nos
    grupos si=0 (3..s-1) e si=2 (e+1..última).
    desp_constant_rows / rec_constant_rows: linhas em que AP e AQ (DESPESAS) / AC (RECEITA) trazem
    uma constante de texto em vez da fórmula. O mestre da fórmula compartilhada fica na primeira linha
    de fórmula que NÃO é constante (normalmente 3 em DESPESAS e 2 em RECEITA); assim dá para cercar o
    mestre de constantes (ex.: desp_constant_rows=(3, 5) põe o mestre em AP4).
    desp_footer_row: linha de resumo abaixo dos dados, com fórmulas de OUTRA família em AP
    (=COUNTA(AP3:AP<linha-1>), intraduzível para a linha 3) e AQ (=$A$1&"x", traduzível mas diferente).
    desp_autofilter: ref do <autoFilter> de DESPESAS (ex.: "A2:AQ7"), com o nome oculto
    _xlnm._FilterDatabase correspondente no workbook.xml, como no modelo real; há sempre também um
    _FilterDatabase da aba OUTRA e um nome comum, que nunca devem mudar."""
    path = Path(path)
    base = _sheet("A1:C12", _styled_rows(range(1, 13), [("A", 0), ("B", 3), ("C", 2)]))

    desp_rows = ['<row r="2" spans="1:43"><c r="AP2" t="inlineStr"><is><t>De-Para</t></is></c></row>']
    last = 2 + desp_formula_rows
    # grupos compartilhados de AP: (1ª linha, última linha, fórmula); com desp_ap_variant=(s, e) há três
    # grupos, como no modelo real: curta em 3..s-1, COMPLETA (ramo a mais) em s..e, curta em e+1..last
    if desp_ap_variant:
        s, e = desp_ap_variant
        assert 3 < s <= e < last, "a variante fica no meio das fórmulas de AP"
        segments = [(3, s - 1, AP_SHORT), (s, e, AP_FULL), (e + 1, last, AP_SHORT)]
    else:
        segments = [(3, last, AP_SHORT)]
    ap_f = {}
    for si, (a, b, text) in enumerate(segments):
        m = next((r for r in range(a, b + 1) if r not in desp_constant_rows), None)   # linha do mestre
        for r in range(a, b + 1):
            ap_f[r] = (f'<f t="shared" ref="AP{m}:AP{b}" si="{si}">{text.format(r=m)}</f>' if r == m
                       else f'<f t="shared" si="{si}"/>')
    for r in range(3, last + 1):
        if r in desp_constant_rows:
            ap, aq = _const(f"AP{r}"), _const(f"AQ{r}")
        else:
            ap = f'<c r="AP{r}" t="e">{ap_f[r]}<v>#N/A</v></c>'
            aq = f'<c r="AQ{r}" t="e"><f>LEN(AP{r})&amp;"!"</f><v>#N/A</v></c>'
        desp_rows.append(f'<row r="{r}" spans="3:43"><c r="C{r}" s="3"/>{ap}{aq}</row>')
    if desp_footer_row is not None:
        f = desp_footer_row
        assert f > last + 1, "o resumo fica abaixo das fórmulas de dados"
        desp_rows.append(f'<row r="{f}"><c r="AP{f}"><f>COUNTA(AP3:AP{f - 1})</f><v>0</v></c>'
                         f'<c r="AQ{f}" t="str"><f>$A$1&amp;"x"</f><v>x</v></c></row>')
        last = f
    af = f'<autoFilter ref="{desp_autofilter}"/>' if desp_autofilter else ""
    despesas = _sheet(f"A2:AQ{last}", "".join(desp_rows), af)

    rec_rows = ['<row r="1" spans="1:29"><c r="AC1" t="inlineStr"><is><t>DRE</t></is></c></row>']
    rlast = 1 + rec_formula_rows
    rm = next((r for r in range(2, rlast + 1) if r not in rec_constant_rows), None)
    for r in range(2, rlast + 1):
        master = (f'<f t="shared" ref="AC{rm}:AC{rlast}" si="0">IF(A{rm}="","",A{rm})</f>' if r == rm
                  else '<f t="shared" si="0"/>')
        cell = _const(f"AC{r}") if r in rec_constant_rows else f'<c r="AC{r}" t="str">{master}<v/></c>'
        rec_rows.append(f'<row r="{r}" spans="1:29">{cell}</row>')
    # RECEITA não tem nenhuma célula estilizada nas colunas de dados (A:C): uma data gravada ali cai
    # sem formato de data e é contada em date_style_warnings (test_date_in_unformatted_cells_is_reported).
    receita = _sheet(f"A1:AC{rlast}", "".join(rec_rows))

    comissoes = _sheet("C2:O8", _styled_rows(range(2, 9), [("C", 3)]))
    outra = _sheet("A1:B1", '<row r="1"><c r="A1"><v>5</v></c><c r="B1"><f>A1*2</f><v>10</v></c></row>')
    sheets = [base, despesas, receita, comissoes, outra]

    wb_sheets = "".join(f'<sheet name="{n}" sheetId="{i}" r:id="rId{i}"/>'
                        for i, n in enumerate(SHEET_NAMES, start=1))
    names = ('<definedName name="_xlnm._FilterDatabase" localSheetId="4" hidden="1">OUTRA!$A$1:$B$1</definedName>'
             '<definedName name="Taxa">OUTRA!$A$1</definedName>')
    if desp_autofilter:
        c1, c2 = desp_autofilter.split(":")
        absolute = lambda ref: re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", ref)
        names = (f'<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">'
                 f'DESPESAS!{absolute(c1)}:{absolute(c2)}</definedName>') + names
    workbook = HEAD + (f'<workbook xmlns="{MAIN}" xmlns:r="{RNS}"><sheets>{wb_sheets}</sheets>'
                       f'<definedNames>{names}</definedNames>{calc_pr or ""}</workbook>')
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
