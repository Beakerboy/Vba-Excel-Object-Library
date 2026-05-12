import vba_types
from vba_excel_obj_lib.worksheet_function import WorksheetFunction


def test_max() -> None:
    assert WorksheetFunction.max(vba_types.VBAArray(1, 2, 3)) == 1


def test_min() -> None:
    assert WorksheetFunction.min(vba_types.VBAArray(1, 2, 3)) == 1


def test_average() -> None:
    assert WorksheetFunction.average(vba_types.VBAArray(1, 2, 3)) == 2.0
