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
