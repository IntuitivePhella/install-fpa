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
