import math
import numpy as np
import scipy
import vba_types
from typing import Any


class WorksheetFunction:
    @staticmethod
    def average(*args: vba_types.VBATypeBase) -> vba_types.VBADouble:
        vals = []
        for arg in args:
            if isinstance(arg, vba_types.VBAArray):
                new_list = list(map(lambda x: x.value, arg))
                vals.extend(new_list)
            else:
                vals.append(arg.value)
        arr = np.array(vals)
        return vba_types.VBADouble(np.average(arr))

    @staticmethod
    def acos(arg1: vba_types.VBATypeBase) -> vba_types.VBADouble:
        return vba_types.VBADouble(math.acos(arg1.value))

    @staticmethod
    def count(*args: vba_types.VBATypeBase) -> vba_types.VBADouble:
        vals = []
        for arg in args:
            if isinstance(arg, vba_types.VBAArray):
                new_list = list(map(lambda x: x.value, arg))
                vals.extend(new_list)
            else:
                vals.append(arg.value)
        return vba_types.VBADouble(len(vals))

    @staticmethod
    def index(
        arg1: Any,
        arg2: Any,
        arg3: Any = None,
        arg4: Any = None,
    ) -> Any:
        """
        Implements VBA's WorksheetFunction.Index(array_or_range, row_num,
        [col_num], [area_num]). Uses 1-based indexing; passing 0 for row_num or
        col_num returns the full column or row.
        """
        def _unwrap(val: Any) -> Any:
            return val.value if hasattr(val, "value") else val

        def _to_grid(container: Any) -> list[list[Any]]:
            # Handle multi-area containers if arg4 is provided
            if (
                isinstance(container, (list, tuple)) and
                container and
                isinstance(container[0], (list, tuple, vba_types.VBAArray))
            ):
                # Check if 2D (list of 1D rows) or 3D (list of 2D areas)
                first_inner = _unwrap(container[0])
                if (
                    isinstance(first_inner, (list, tuple)) and
                    first_inner and
                    isinstance(_unwrap(first_inner[0]), (list, tuple))
                ):
                    area_idx = (
                        int(_unwrap(arg4)) - 1
                        if arg4 is not None else 0
                    )
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

            # Determine whether elements are row containers (2D) or scalar
            # VBATypeBase items (1D)
            unwrapped_first = _unwrap(items[0])
            if (
                isinstance(items[0], vba_types.VBAArray) or
                isinstance(unwrapped_first, (list, tuple, np.ndarray))
            ):
                grid: list[list[Any]] = []
                for row in items:
                    is_instance = isinstance(
                        row,
                        (vba_types.VBAArray, list, tuple, np.ndarray)
                    )
                    row_items = list(row) if is_instance else [row]
                    grid.append(row_items)
                return grid

            # 1D array treated as a single row (or column if accessed with
            # row_num)
            return [items]

        def _wrap_result(val: Any) -> Any:
            if isinstance(val, vba_types.VBATypeBase):
                return val
            unwrapped = _unwrap(val)
            if (
                isinstance(unwrapped, (int, np.integer)) and
                not isinstance(unwrapped, bool)
            ):
                return vba_types.VBADouble(float(unwrapped))
            if isinstance(unwrapped, (float, np.floating)):
                return vba_types.VBADouble(float(unwrapped))
            return val

        grid = _to_grid(arg1)
        num_rows = len(grid)
        num_cols = max(len(row) for row in grid)

        # Normalize to rectangular grid
        rect_grid = [
            list(row) + [None] * (num_cols - len(row))
            for row in grid
        ]

        row_num = int(_unwrap(arg2)) if arg2 is not None else 0
        col_num = int(_unwrap(arg3)) if arg3 is not None else 0

        # If a 1D list was passed as a single row and row_num is used without
        # col_num, treat row_num as the index along that 1D sequence.
        if num_rows == 1 and arg3 is None and row_num > 0:
            if row_num > num_cols:
                raise IndexError("row_num out of range")
            return _wrap_result(rect_grid[0][row_num - 1])

        if row_num < 0 or row_num > num_rows:
            raise IndexError("row_num out of range")
        if col_num < 0 or col_num > num_cols:
            raise IndexError("col_num out of range")

        # Full column when row_num == 0 and col_num > 0
        if row_num == 0 and col_num > 0:
            col_slice = [
                _wrap_result(rect_grid[r][col_num - 1])
                for r in range(num_rows)
            ]
            return vba_types.VBAArray(col_slice)

        # Full row when col_num == 0 and row_num > 0
        if col_num == 0 and row_num > 0:
            row_slice = [_wrap_result(val) for val in rect_grid[row_num - 1]]
            return vba_types.VBAArray(row_slice)

        # Default col_num to 1 when omitted and row_num > 0
        effective_col = col_num if col_num > 0 else 1
        if row_num == 0:
            raise ValueError("row_num and col_num cannot both be 0")
        return _wrap_result(rect_grid[row_num - 1][effective_col - 1])

    @staticmethod
    def max(*args: vba_types.VBATypeBase) -> vba_types.VBADouble:
        """
        Should be able to accept numbers and lists of numbers
        """
        vals = []
        for arg in args:
            if isinstance(arg, vba_types.VBAArray):
                new_list = list(map(lambda x: x.value, arg))
                vals.extend(new_list)
            else:
                vals.append(arg.value)
        return vba_types.VBADouble(max(vals))

    @staticmethod
    def min(*args: vba_types.VBATypeBase) -> vba_types.VBADouble:
        vals = []
        for arg in args:
            if isinstance(arg, vba_types.VBAArray):
                new_list = list(map(lambda x: x.value, arg))
                vals.extend(new_list)
            else:
                vals.append(arg.value)
        return vba_types.VBADouble(min(vals))

    @staticmethod
    def mmult(array1: vba_types.VBAArray,
              array2: vba_types.VBAArray) -> vba_types.VBAArray:
        a1 = WorksheetFunction._to_numpy(array1)
        a2 = WorksheetFunction._to_numpy(array2)
        if a1.ndim == 1:
            a1 = a1.reshape(1, -1)
        if a2.ndim == 1:
            a2 = a2.reshape(-1, 1)
        result = a1 @ a2
        return WorksheetFunction._to_vba_array(result)

    @staticmethod
    def pi() -> vba_types.VBADouble:
        return vba_types.VBADouble(math.pi)

    @staticmethod
    def small(values: vba_types.VBAArray,
              k: vba_types.VBAInteger) -> vba_types.VBADouble:
        arr = np.array(values._data)
        return vba_types.VBADouble(np.partition(arr, int(k) - 1)[int(k) - 1])

    @staticmethod
    def stdev(*args: vba_types.VBATypeBase) -> vba_types.VBADouble:
        vals = []
        for arg in args:
            if isinstance(arg, vba_types.VBAArray):
                new_list = list(map(lambda x: x.value, arg))
                vals.extend(new_list)
            else:
                vals.append(arg.value)
        arr = np.array(vals)
        return vba_types.VBADouble(np.std(arr, ddof=1))

    @staticmethod
    def t_inv(p: vba_types.VBADouble,
              df: vba_types.VBADouble) -> vba_types.VBADouble:
        result = scipy.special.stdtrit(df.value, p.value)
        return vba_types.VBADouble(float(result))

    @staticmethod
    def transpose(arg1: vba_types.VBAArray) -> vba_types.VBAArray:
        arr = WorksheetFunction._to_numpy(arg1)
        if arr.ndim == 1:
            result: np.ndarray[Any, Any] = arr.reshape(-1, 1)
        elif arr.ndim == 2 and arr.shape[1] == 1:
            result = arr.reshape(-1)
        else:
            result = arr.T
        return WorksheetFunction._to_vba_array(result)

    @staticmethod
    def _extract_raw(item: object) -> Any:
        if isinstance(item, (list, tuple)):
            return [WorksheetFunction._extract_raw(x) for x in item]
        if isinstance(item, vba_types.VBATypeBase):
            return item.value
        return item

    @staticmethod
    def _to_numpy(arr: vba_types.VBAArray) -> np.ndarray:
        if not isinstance(arr, vba_types.VBAArray):
            return np.array([[float(arr.value)]], dtype=float)
        raw = WorksheetFunction._extract_raw(arr._data)
        return np.array(raw, dtype=float)

    @staticmethod
    def _to_vba_array(np_arr: np.ndarray[Any, Any]) -> vba_types.VBAArray:
        """
        Helper to wrap a 1D or 2D NumPy array back into a 1-based VBAArray.
        """
        if np_arr.ndim == 1:
            items = [vba_types.VBADouble(float(v)) for v in np_arr]
            return vba_types.VBAArray(*items, base=1)

        res = vba_types.VBAArray(base=1)
        res._data = [
            [vba_types.VBADouble(float(v)) for v in row]
            for row in np_arr
        ]
        res._bounds = [(1, np_arr.shape[0]), (1, np_arr.shape[1])]
        return res


api = (
    vba_types.LibraryBuilder("excel")
    .add_module("worksheetfunction", WorksheetFunction)
    .build()["modules"]["worksheetfunction"]
)
