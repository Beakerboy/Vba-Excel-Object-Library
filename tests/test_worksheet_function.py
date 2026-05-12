import pytest
import vba_types
from vba_excel_obj_lib.worksheet_function import WorksheetFunction


@pytest.mark.parametrize(
    "input, expected", [
        (vba_types.VBAArray(1, 2, 3), vba_types.VBADouble(3.0)),
    ])
def test_max(input, expected) -> None:
    result = WorksheetFunction.max(input)
    assert result == expected


def test_min() -> None:
    assert WorksheetFunction.min(vba_types.VBAArray(1, 2, 3)) == vba_types.VBADouble(1.0)


def test_average_array() -> None:
    assert WorksheetFunction.average(vba_types.VBAArray(1, 2, 3)) == vba_types.VBADouble(2.0)


def test_average_ints() -> None:
    assert WorksheetFunction.average(1, 2, 3) == vba_types.VBADouble(2.0)


def test_average_mix() -> None:
    assert WorksheetFunction.average(vba_types.VBAArray(1, 2), 3) == vba_types.VBADouble(2.0)


def stdev() -> None:
    assert WorksheetFunction.stdev(vba_types.VBAArray(1, 2, 3)) == vba_types.VBADouble(1.0)
