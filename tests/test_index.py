from typing import Any
import numpy as np
import pytest
import vba_types
from vba_excel_obj_lib.worksheet_function import WorksheetFunction
from vba_types.exceptions import SubscriptOutOfRangeError, TypeMismatchError


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
    result = WorksheetFunction.index(matrix, 2, 2)
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


def test_1d_list_and_tuple_lookups() -> None:
    tup = (100, 200, 300)
    assert WorksheetFunction.index(tup, 1).value == 100.0
    assert WorksheetFunction.index(tup, 3).value == 300.0


def test_1d_list_row_zero_returns_full_array() -> None:
    arr = [10, 20, 30]
    result = WorksheetFunction.index(arr, 0)
    assert isinstance(result, vba_types.VBAArray)
    assert [x.value for x in result] == [10.0, 20.0, 30.0]


def test_numpy_1d_and_2d_arrays() -> None:
    arr_1d = np.array([10, 20, 30])
    assert WorksheetFunction.index(arr_1d, 2).value == 20.0

    arr_2d = np.array([[100, 200], [300, 400]])
    result = WorksheetFunction.index(arr_2d, 1, 2)
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 200.0


def test_scalar_input_lookup() -> None:
    result = WorksheetFunction.index(42, 1, 1)
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 42.0


def test_multi_area_selection_default_and_explicit() -> None:
    area1 = [[1, 2], [3, 4]]
    area2 = [[50, 60], [70, 80]]
    areas = [area1, area2]

    first = WorksheetFunction.index(areas, 2, 1)
    assert first.value == 3.0

    second = WorksheetFunction.index(areas, 2, 2, 2)
    assert second.value == 80.0


def test_float_and_string_index_coercion() -> None:
    arr = [10, 20, 30]
    assert WorksheetFunction.index(arr, 1.6).value == 20.0
    assert WorksheetFunction.index(arr, 2.0).value == 20.0
    assert WorksheetFunction.index(arr, "3").value == 30.0


@pytest.mark.parametrize("row_num", [5, -1])
def test_1d_list_out_of_range_raises(row_num: int) -> None:
    arr = [10, 20, 30]
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(arr, row_num)


def test_1d_vba_array_out_of_range_raises() -> None:
    arr = vba_types.VBAArray(vba_types.VBADouble(1), vba_types.VBADouble(2))
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(arr, 3)


def test_2d_out_of_bounds_row_and_column_raise() -> None:
    matrix = [[1, 2], [3, 4]]
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(matrix, 3, 1)
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(matrix, 1, 3)
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(matrix, -1, 1)
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(matrix, 1, -1)


def test_both_row_and_col_zero_on_2d_raises() -> None:
    matrix = [[1, 2], [3, 4]]
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(matrix, 0, 0)


def test_invalid_area_num_raises() -> None:
    areas = [[[1, 2]], [[3, 4]]]
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(areas, 1, 1, 0)
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(areas, 1, 1, 3)


def test_3d_numpy_array_raises() -> None:
    arr_3d = np.zeros((2, 2, 2))
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index(arr_3d, 1, 1)


def test_empty_containers_raise() -> None:
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index([], 1, 1)
    with pytest.raises(SubscriptOutOfRangeError):
        WorksheetFunction.index((), 1, 1)


@pytest.mark.parametrize(
    "bad_idx",
    [True, False, float("nan"), float("inf"), "abc", object()],
)
def test_invalid_index_type_raises_type_mismatch(bad_idx: Any) -> None:
    arr = [10, 20, 30]
    with pytest.raises(TypeMismatchError):
        WorksheetFunction.index(arr, bad_idx)
