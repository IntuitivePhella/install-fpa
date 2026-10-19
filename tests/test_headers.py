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
