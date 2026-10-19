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
