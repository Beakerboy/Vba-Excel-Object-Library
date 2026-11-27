import math
import numpy as np
import scipy
import vba_types
from typing import Any
from vba_types.exceptions import (
    SubscriptOutOfRangeError,
    TypeMismatchError,
    VBAException
)


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
    def chisq_inv(
        probability: vba_types.VBATypeBase,
        deg_freedom: vba_types.VBATypeBase,
    ) -> vba_types.VBADouble:
        """VBA WorksheetFunction.ChiSq_Inv(probability, deg_freedom)."""

        def _to_float(val: vba_types.VBATypeBase) -> float:
            raw = val.value
            if isinstance(raw, bool):
                raise TypeMismatchError()
            if isinstance(raw, (int, float)):
                return float(raw)
            if isinstance(raw, str):
                try:
                    return float(raw.strip())
                except ValueError as exc:
                    raise TypeMismatchError() from exc
            raise TypeMismatchError()

        prob = _to_float(probability)
        df = _to_float(deg_freedom)

        if not math.isfinite(prob) or prob <= 0.0 or prob >= 1.0:
            raise VBAException()
        if not math.isfinite(df) or df < 1.0:
            raise VBAException()

        result = float(scipy.stats.chi2.ppf(prob, df=df))
        return vba_types.VBADouble(result)

    @staticmethod
    def count(
        *args: vba_types.VBATypeBase,
    ) -> vba_types.VBADouble:
        """VBA WorksheetFunction.Count over numeric values."""

        def _count_item(item: vba_types.VBATypeBase) -> int:
            if isinstance(item, vba_types.VBAArray):
                return sum(_count_item(elem) for elem in item)
            if isinstance(item, vba_types.VBABoolean):
                return 0
            if isinstance(item, vba_types.VBANumericType):
                return 1
            return 0


        total = sum(_count_item(arg, is_top_level=True) for arg in args)
        return vba_types.VBADouble(float(total))

    @staticmethod
    def index(
        arg1: vba_types.VBAArray,
        arg2: vba_types.VBATypeBase | None = None,
        arg3: vba_types.VBATypeBase | None = None,
        arg4: vba_types.VBATypeBase | None = None,
    ) -> vba_types.VBATypeBase:
        """VBA WorksheetFunction.Index using VBAArray bounds."""

        def _unwrap(val: Any) -> Any:
            if isinstance(val, vba_types.VBAArray):
                return list(val)
            if hasattr(val, "value"):
                return val.value
            return val

        def _to_index_int(
            val: vba_types.VBATypeBase | None,
            default: int = 0,
        ) -> int:
            if val is None:
                return default
            raw = _unwrap(val)
            if isinstance(raw, bool):
                raise TypeMismatchError()
            if isinstance(raw, (int, np.integer)):
                return int(raw)
            if isinstance(raw, (float, np.floating)):
                if not np.isfinite(raw):
                    raise TypeMismatchError()
                if float(raw).is_integer():
                    return int(raw)
                return int(round(float(raw)))
            if isinstance(raw, str):
                try:
                    return int(float(raw.strip()))
                except ValueError as exc:
                    raise TypeMismatchError() from exc
            raise TypeMismatchError()

        def _bounds(
            arr: vba_types.VBAArray,
            dim: int,
        ) -> tuple[int, int]:
            dim_arg = vba_types.VBALong(dim)
            low = _to_index_int(arr.lbound(dim_arg))
            high = _to_index_int(arr.ubound(dim_arg))
            return low, high

        def _vba_array_to_grid(
            arr: vba_types.VBAArray,
        ) -> list[list[Any]]:
            ndim = len(arr._bounds)
            if ndim == 0:
                raise SubscriptOutOfRangeError()

            r_low, r_high = _bounds(arr, 1)
            row_keys = list(range(r_low, r_high + 1))
            if not row_keys:
                raise SubscriptOutOfRangeError()

            if ndim == 1:
                return [[arr[r_key] for r_key in row_keys]]

            if ndim == 2:
                c_low, c_high = _bounds(arr, 2)
                col_keys = list(range(c_low, c_high + 1))
                if not col_keys:
                    raise SubscriptOutOfRangeError()
                return [
                    [arr[r_key, c_key] for c_key in col_keys]
                    for r_key in row_keys
                ]

            raise SubscriptOutOfRangeError()

        def _wrap_result(val: Any) -> vba_types.VBATypeBase:
            if isinstance(val, vba_types.VBATypeBase):
                return val
            unwrapped = _unwrap(val)
            if isinstance(unwrapped, bool):
                if hasattr(vba_types, "VBABoolean"):
                    return vba_types.VBABoolean(unwrapped)
                raise TypeMismatchError()
            if isinstance(
                unwrapped, (int, float, np.integer, np.floating)
            ):
                return vba_types.VBADouble(float(unwrapped))
            if isinstance(unwrapped, str) and hasattr(
                vba_types, "VBAString"
            ):
                return vba_types.VBAString(unwrapped)
            raise TypeMismatchError()

        if arg2 is None and arg3 is None:
            raise TypeMismatchError()

        grid = _vba_array_to_grid(arg1)
        num_rows = len(grid)
        num_cols = max(len(row) for row in grid)

        rect_grid = [
            list(row) + [None] * (num_cols - len(row)) for row in grid
        ]

        row_num = _to_index_int(arg2, default=0)
        if arg3 is None and row_num > 0 and num_rows > 1:
            col_num = 1
        else:
            col_num = _to_index_int(arg3, default=0)

        if num_rows == 1 and arg3 is None:
            if row_num == 0:
                return vba_types.VBAArray(
                    *[_wrap_result(x) for x in rect_grid[0]]
                )
            if row_num < 0 or row_num > num_cols:
                raise SubscriptOutOfRangeError()
            return _wrap_result(rect_grid[0][row_num - 1])

        if row_num < 0 or row_num > num_rows:
            raise SubscriptOutOfRangeError()
        if col_num < 0 or col_num > num_cols:
            raise SubscriptOutOfRangeError()

        if row_num == 0 and col_num > 0:
            col_slice = [
                _wrap_result(rect_grid[r][col_num - 1])
                for r in range(num_rows)
            ]
            return vba_types.VBAArray(*col_slice)

        if col_num == 0 and row_num > 0:
            row_slice = [
                _wrap_result(val) for val in rect_grid[row_num - 1]
            ]
            return vba_types.VBAArray(*row_slice)

        if row_num == 0:
            raise SubscriptOutOfRangeError()

        effective_col = col_num if col_num > 0 else 1
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
    def norm_s_dist(
        z: vba_types.VBATypeBase,
        cumulative: vba_types.VBATypeBase,
    ) -> vba_types.VBADouble:
        """VBA WorksheetFunction.Norm_S_Dist(z, cumulative)."""
        z_val = float(z.value)
        is_cdf = bool(cumulative.value)
        if is_cdf:
            result = float(scipy.stats.norm.cdf(z_val, loc=0.0, scale=1.0))
        else:
            result = float(scipy.stats.norm.pdf(z_val, loc=0.0, scale=1.0))
        return vba_types.VBADouble(result)

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
