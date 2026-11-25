import sys
import types
import numpy as np
import pytest
import vba_types
from vba_excel_obj_lib.worksheet_function import WorksheetFunction

def test_1d_vba_array_lookup() -> None:
    arr = vba_types.VBAArray(
        [vba_types.VBADouble(10), vba_types.VBADouble(20), vba_types.VBADouble(30)]
    )
    result = WorksheetFunction.index(arr, vba_types.VBADouble(2))
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 20.0


def test_2d_nested_vba_array_lookup() -> None:
    matrix = vba_types.VBAArray(
        [
            vba_types.VBAArray([vba_types.VBADouble(1), vba_types.VBADouble(2)]),
            vba_types.VBAArray([vba_types.VBADouble(3), vba_types.VBADouble(4)]),
        ]
    )
    result = WorksheetFunction.index(
        matrix, vba_types.VBADouble(2), vba_types.VBADouble(1)
    )
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 3.0


def test_2d_default_column_is_one() -> None:
    matrix = vba_types.VBAArray(
        [
            vba_types.VBAArray([vba_types.VBADouble(5), vba_types.VBADouble(6)]),
            vba_types.VBAArray([vba_types.VBADouble(7), vba_types.VBADouble(8)]),
        ]
    )
    result = WorksheetFunction.index(matrix, vba_types.VBADouble(2))
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 7.0


def test_full_row_when_col_num_is_zero() -> None:
    matrix = vba_types.VBAArray(
        [
            vba_types.VBAArray([vba_types.VBADouble(10), vba_types.VBADouble(20)]),
            vba_types.VBAArray([vba_types.VBADouble(30), vba_types.VBADouble(40)]),
        ]
    )
    row_result = WorksheetFunction.index(
        matrix, vba_types.VBADouble(2), vba_types.VBADouble(0)
    )
    assert isinstance(row_result, vba_types.VBAArray)
    assert [x.value for x in row_result] == [30.0, 40.0]


def test_full_column_when_row_num_is_zero() -> None:
    matrix = vba_types.VBAArray(
        [
            vba_types.VBAArray([vba_types.VBADouble(10), vba_types.VBADouble(20)]),
            vba_types.VBAArray([vba_types.VBADouble(30), vba_types.VBADouble(40)]),
        ]
    )
    col_result = WorksheetFunction.index(
        matrix, vba_types.VBADouble(0), vba_types.VBADouble(2)
    )
    assert isinstance(col_result, vba_types.VBAArray)
    assert [x.value for x in col_result] == [20.0, 40.0]


def test_numpy_2d_array_wraps_numeric_result() -> None:
    arr = np.array([[100, 200], [300, 400]])
    result = WorksheetFunction.index(arr, 1, 2)
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 200.0


def test_multi_area_selection() -> None:
    area1 = [[1, 2], [3, 4]]
    area2 = [[50, 60], [70, 80]]
    result = WorksheetFunction.index([area1, area2], 2, 2, 2)
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 80.0


@pytest.mark.parametrize(
    "row_num, col_num",
    [
        (0, 1),   # 1D-style single row with row_num=0 and omitted col_num
        (5, 1),   # row_num exceeds row count
        (-1, 1),  # negative row_num
    ],
)
def test_1d_list_invalid_indices_raise(row_num, col_num) -> None:
    arr = [10, 20, 30]
    with pytest.raises(IndexError):
        WorksheetFunction.index(arr, row_num)


def test_2d_out_of_bounds_column_raises() -> None:
    matrix = [[1, 2], [3, 4]]
    with pytest.raises(IndexError):
        WorksheetFunction.index(matrix, 1, 3)


def test_both_row_and_col_zero_raises_value_error() -> None:
    matrix = [[1, 2], [3, 4]]
    with pytest.raises(ValueError):
        WorksheetFunction.index(matrix, 0, 0)


def test_empty_array_raises_value_error():
    with pytest.raises(ValueError):
        WorksheetFunction.index([], 1, 1)
