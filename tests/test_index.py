import sys
import types
import numpy as np
import pytest


# Provide minimal stubs if vba_types is not already in sys.modules
if "vba_types" not in sys.modules:
    vba_stub = types.ModuleType("vba_types")

    class VBATypeBase:
        def __init__(self, value):
            self.value = value

        def __eq__(self, other):
            if isinstance(other, VBATypeBase):
                return self.value == other.value
            return self.value == other

        def __repr__(self):
            return f"{self.__class__.__name__}({self.value!r})"

    class VBADouble(VBATypeBase):
        def __init__(self, value):
            super().__init__(float(value))

    class VBAArray(VBATypeBase):
        def __init__(self, items):
            super().__init__(list(items))

        def __iter__(self):
            return iter(self.value)

        def __len__(self):
            return len(self.value)

    vba_stub.VBATypeBase = VBATypeBase
    vba_stub.VBADouble = VBADouble
    vba_stub.VBAArray = VBAArray
    sys.modules["vba_types"] = vba_stub

import vba_types


class WorksheetFunction:
    @staticmethod
    def index(
        arg1,
        arg2,
        arg3=None,
        arg4=None,
    ):
        def _unwrap(val):
            return val.value if hasattr(val, "value") else val

        def _to_grid(container):
            if (
                isinstance(container, (list, tuple))
                and container
                and isinstance(container[0], (list, tuple, vba_types.VBAArray))
            ):
                first_inner = _unwrap(container[0])
                if (
                    isinstance(first_inner, (list, tuple))
                    and first_inner
                    and isinstance(_unwrap(first_inner[0]), (list, tuple))
                ):
                    area_idx = int(_unwrap(arg4)) - 1 if arg4 is not None else 0
                    if area_idx < 0 or area_idx >= len(container):
                        raise IndexError("area_num out of range")
                    container = container[area_idx]

            if isinstance(container, np.ndarray):
                if container.ndim == 1:
                    return [list(container)]
                if container.ndim == 2:
                    return [list(row) for row in container]
                raise ValueError("Only 1D and 2D arrays are supported")

            if isinstance(container, vba_types.VBAArray):
                items = list(container)
            elif isinstance(container, (list, tuple)):
                items = list(container)
            else:
                return [[container]]

            if not items:
                raise ValueError("Array cannot be empty")

            unwrapped_first = _unwrap(items[0])
            if isinstance(items[0], vba_types.VBAArray) or isinstance(
                unwrapped_first, (list, tuple, np.ndarray)
            ):
                grid = []
                for row in items:
                    row_items = (
                        list(row)
                        if isinstance(row, (vba_types.VBAArray, list, tuple, np.ndarray))
                        else [row]
                    )
                    grid.append(row_items)
                return grid

            return [items]

        def _wrap_result(val):
            if isinstance(val, vba_types.VBATypeBase):
                return val
            unwrapped = _unwrap(val)
            if isinstance(unwrapped, (int, np.integer)) and not isinstance(unwrapped, bool):
                return vba_types.VBADouble(float(unwrapped))
            if isinstance(unwrapped, (float, np.floating)):
                return vba_types.VBADouble(float(unwrapped))
            return val

        grid = _to_grid(arg1)
        num_rows = len(grid)
        num_cols = max(len(row) for row in grid)

        rect_grid = [list(row) + [None] * (num_cols - len(row)) for row in grid]

        row_num = int(_unwrap(arg2)) if arg2 is not None else 0
        col_num = int(_unwrap(arg3)) if arg3 is not None else 0

        if num_rows == 1 and arg3 is None and row_num > 0:
            if row_num > num_cols:
                raise IndexError("row_num out of range")
            return _wrap_result(rect_grid[0][row_num - 1])

        if row_num < 0 or row_num > num_rows:
            raise IndexError("row_num out of range")
        if col_num < 0 or col_num > num_cols:
            raise IndexError("col_num out of range")

        if row_num == 0 and col_num > 0:
            col_slice = [
                _wrap_result(rect_grid[r][col_num - 1]) for r in range(num_rows)
            ]
            return vba_types.VBAArray(col_slice)

        if col_num == 0 and row_num > 0:
            row_slice = [_wrap_result(val) for val in rect_grid[row_num - 1]]
            return vba_types.VBAArray(row_slice)

        effective_col = col_num if col_num > 0 else 1
        if row_num == 0:
            raise ValueError("row_num and col_num cannot both be 0")

        return _wrap_result(rect_grid[row_num - 1][effective_col - 1])


# --- Unit Tests ---


def test_1d_vba_array_lookup():
    arr = vba_types.VBAArray(
        [vba_types.VBADouble(10), vba_types.VBADouble(20), vba_types.VBADouble(30)]
    )
    result = WorksheetFunction.index(arr, vba_types.VBADouble(2))
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 20.0


def test_2d_nested_vba_array_lookup():
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


def test_2d_default_column_is_one():
    matrix = vba_types.VBAArray(
        [
            vba_types.VBAArray([vba_types.VBADouble(5), vba_types.VBADouble(6)]),
            vba_types.VBAArray([vba_types.VBADouble(7), vba_types.VBADouble(8)]),
        ]
    )
    result = WorksheetFunction.index(matrix, vba_types.VBADouble(2))
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 7.0


def test_full_row_when_col_num_is_zero():
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


def test_full_column_when_row_num_is_zero():
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


def test_numpy_2d_array_wraps_numeric_result():
    arr = np.array([[100, 200], [300, 400]])
    result = WorksheetFunction.index(arr, 1, 2)
    assert isinstance(result, vba_types.VBADouble)
    assert result.value == 200.0


def test_multi_area_selection():
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
def test_1d_list_invalid_indices_raise(row_num, col_num):
    arr = [10, 20, 30]
    with pytest.raises(IndexError):
        WorksheetFunction.index(arr, row_num)


def test_2d_out_of_bounds_column_raises():
    matrix = [[1, 2], [3, 4]]
    with pytest.raises(IndexError):
        WorksheetFunction.index(matrix, 1, 3)


def test_both_row_and_col_zero_raises_value_error():
    matrix = [[1, 2], [3, 4]]
    with pytest.raises(ValueError):
        WorksheetFunction.index(matrix, 0, 0)


def test_empty_array_raises_value_error():
    with pytest.raises(ValueError):
        WorksheetFunction.index([], 1, 1)
