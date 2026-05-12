import numpy as np
import vba_types
from typing import Any


class WorksheetFunction:
    @staticmethod
    def max(*args: Any) -> Any:
        """
        Should be able to accept numbers and lists of numbers
        """
        return max(*args)

    @staticmethod
    def min(*args: Any) -> Any:
        """
        Should be able to accept numbers and lists of numbers
        """
        return min(*args)

    @staticmethod
    def small(values: vba_types.VBAArray, k: vba_types.VBAInteger) -> Any:
        arr = np.array(values._data)
        return np.partition(arr, int(k) - 1)[int(k) - 1]
