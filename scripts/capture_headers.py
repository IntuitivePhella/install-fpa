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
