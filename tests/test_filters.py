import pytest
import math
from kisorlib.filters import filter_number


def test_filter_number():
    # None and NaN cases
    assert filter_number(None) == "0"
    assert filter_number(float("nan")) == "0"

    # Integer and Float cases
    assert filter_number(1500000) == "1.500.000"
    assert filter_number(1500000.0) == "1.500.000"
    assert filter_number(1.5) == "2"  # rounding check

    # Vietnamese / US string formats
    assert filter_number("1.500") == "1.500"
    assert filter_number("1,500,000") == "1.500.000"
    assert filter_number("1.500.000") == "1.500.000"
    assert filter_number("1234567") == "1.234.567"
    assert filter_number("  1.234  ") == "1.234"
