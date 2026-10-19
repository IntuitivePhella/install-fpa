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
