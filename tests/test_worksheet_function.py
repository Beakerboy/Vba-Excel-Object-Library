import pytest
import vba_types
from vba_excel_obj_lib.worksheet_function import WorksheetFunction


array_of_ints = vba_types.VBAArray(vba_types.VBAInteger(1),
                                   vba_types.VBAInteger(2),
                                   vba_types.VBAInteger(3))
@pytest.mark.parametrize(
    "input, expected", [
        (array_of_ints, vba_types.VBADouble(3.0)),
    ])
def test_max(input, expected) -> None:
    result = WorksheetFunction.max(input)
    assert result == expected


def test_min() -> None:
    assert WorksheetFunction.min(array_of_ints) == vba_types.VBADouble(1.0)


def test_average_array() -> None:
    assert WorksheetFunction.average(array_of_ints) == vba_types.VBADouble(2.0)


def test_average_ints() -> None:
    assert WorksheetFunction.average(vba_types.VBAInteger(1),
                                     vba_types.VBAInteger(2),
                                     vba_types.VBAInteger(3)) == vba_types.VBADouble(2.0)


def test_average_mix() -> None:
    result = WorksheetFunction.average(
        vba_types.VBAArray(vba_types.VBAInteger(1), vba_types.VBAInteger(2)),
        vba_types.VBAInteger(3)
    )
    assert result == vba_types.VBADouble(2.0)


def stdev() -> None:
    assert WorksheetFunction.stdev(array_of_ints) == vba_types.VBADouble(1.0)
