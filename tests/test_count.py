import vba_types
from vba_excel_obj_lib.worksheet_function import WorksheetFunction


def test_count_mixed_scalars_and_array() -> None:
    arr = vba_types.VBAArray(
        vba_types.VBADouble(10.0),
        vba_types.VBADouble(20.0),
    )
    result = WorksheetFunction.count(
        vba_types.VBADouble(1.0),
        arr,
        vba_types.VBADouble(5.0),
    )
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 4.0


# Need to verify behavior in VBA
def test_count_ignores_non_numeric_values() -> None:
    items = vba_types.VBAArray(
        vba_types.VBADouble(42.0),
        vba_types.VBAString("hello"),
        vba_types.VBABoolean(True),
    )
    result = WorksheetFunction.count(
        items,
        vba_types.VBAString("world"),
        vba_types.VBABoolean(False),
    )
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 1.0


def test_count_empty_args_and_empty_array() -> None:
    assert WorksheetFunction.count().value == 0.0
    empty_arr = vba_types.VBAArray()
    assert WorksheetFunction.count(empty_arr).value == 0.0
