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


def test_coercion_errors_never_echo_the_cell_text():
    segredo = "Fulano Inventado 111.222.333-44"
    for fn in (parse_money, parse_date):
        with pytest.raises(ValueError) as e:
            fn(segredo)
        assert "Fulano" not in str(e.value) and "111.222" not in str(e.value)
