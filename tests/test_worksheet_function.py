import vba_excel_obj_lib.worksheet_function as worksheet_function


def test_min() -> None:
    assert worksheet_function.min(1, 2, 3) == 1
