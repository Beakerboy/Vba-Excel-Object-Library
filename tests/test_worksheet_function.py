from vba_excel_obj_lib.worksheet_function import WorksheetFunction


def test_min() -> None:
    assert WorksheetFunction.min(1, 2, 3) == 1
