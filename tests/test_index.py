import pytest
import vba_types
from vba_excel_obj_lib.worksheet_function import WorksheetFunction
from vba_types.exceptions import SubscriptOutOfRangeError


def _make_2d_vba_array(
    rows: list[list[vba_types.VBATypeBase]],
    r_low: int = 1,
    c_low: int = 1,
) -> vba_types.VBAArray:
    num_r = len(rows)
    num_c = len(rows[0])
    arr = vba_types.VBAArray.initialize(
        (r_low, r_low + num_r - 1),
        (c_low, c_low + num_c - 1),
        empty=vba_types.VBADouble(0),
    )
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            arr[r_low + r_idx, c_low + c_idx] = val
    return arr


def test_1d_vba_array_lookup() -> None:
    arr = vba_types.VBAArray(
        vba_types.VBADouble(10),
        vba_types.VBADouble(20),
        vba_types.VBADouble(30),
    )
    result = WorksheetFunction.index(arr, vba_types.VBADouble(2))
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 20.0


def test_1d_vba_array_row_zero_returns_full_array() -> None:
    arr = vba_types.VBAArray(
        vba_types.VBADouble(10),
        vba_types.VBADouble(20),
        vba_types.VBADouble(30),
    )
    result = WorksheetFunction.index(arr, 0)
    assert isinstance(result, vba_types.VBAArray)
    assert [x.value for x in result] == [10.0, 20.0, 30.0]


def test_2d_vba_array_lookup() -> None:
    matrix = _make_2d_vba_array(
        [
            [vba_types.VBADouble(1), vba_types.VBADouble(2)],
            [vba_types.VBADouble(3), vba_types.VBADouble(4)],
        ]
    )
    result = WorksheetFunction.index(
        matrix, vba_types.VBADouble(2), vba_types.VBADouble(1)
    )
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 3.0


def test_2d_default_column_is_one() -> None:
    matrix = _make_2d_vba_array(
        [
            [vba_types.VBADouble(5), vba_types.VBADouble(6)],
            [vba_types.VBADouble(7), vba_types.VBADouble(8)],
        ]
    )
    result = WorksheetFunction.index(matrix, vba_types.VBADouble(2))
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 7.0


def test_2d_vba_array_custom_and_negative_bounds() -> None:
    matrix = _make_2d_vba_array(
        [
            [vba_types.VBADouble(11), vba_types.VBADouble(12)],
            [vba_types.VBADouble(21), vba_types.VBADouble(22)],
        ],
        r_low=-1,
        c_low=0,
    )
    result = WorksheetFunction.index(
        matrix,
        vba_types.VBADouble(2),
        vba_types.VBADouble(2)
    )
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 22.0


def test_full_row_when_col_num_is_zero() -> None:
    matrix = _make_2d_vba_array(
        [
            [vba_types.VBADouble(10), vba_types.VBADouble(20)],
            [vba_types.VBADouble(30), vba_types.VBADouble(40)],
        ]
    )
    row_result = WorksheetFunction.index(
        matrix, vba_types.VBADouble(2), vba_types.VBADouble(0)
    )
    assert isinstance(row_result, vba_types.VBAArray)
    assert [x.value for x in row_result] == [30.0, 40.0]


def test_full_column_when_row_num_is_zero() -> None:
    matrix = _make_2d_vba_array(
        [
            [vba_types.VBADouble(10), vba_types.VBADouble(20)],
            [vba_types.VBADouble(30), vba_types.VBADouble(40)],
        ]
    )
    col_result = WorksheetFunction.index(
        matrix, vba_types.VBADouble(0), vba_types.VBADouble(2)
    )
    assert isinstance(col_result, vba_types.VBAArray)
    assert [x.value for x in col_result] == [20.0, 40.0]


def test_1d_vba_array_out_of_range_raises() -> None:
    arr = vba_types.VBAArray(vba_types.VBADouble(1), vba_types.VBADouble(2))
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(arr, 3)
