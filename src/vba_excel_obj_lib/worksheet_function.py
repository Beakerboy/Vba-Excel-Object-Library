from __future__ import annotations
import numpy as np
import scipy
import vba_types


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

    @ataticmethod
    def t_inv(p, df) -> VBADouble:
        result = scipy.special.stdtrit(df, p)
        return VBADouble(result)

arg1 = {
    "name": "arg1",
    "optional": False,
    "default": ""
}


args = [arg1]
i = 2
for n in range(29):
    args.append({
        "name": f"arg{i}",
        "optional": True,
        "default": ""
    })


api = {
    "name": "worksheetfunction",
    "type": "module",
    "functions": {
        "average": {
            "name": "average",
            "type": "function",
            "project": "excel",
            "module": "worksheetfunction",
            "handle": getattr(WorksheetFunction, "average"),
            "params": args
        },
        "count": {
            "name": "count",
            "type": "function",
            "project": "excel",
            "module": "worksheetfunction",
            "handle": getattr(WorksheetFunction, "count"),
            "params": args
        },
        "max": {
            "name": "max",
            "type": "function",
            "project": "excel",
            "module": "worksheetfunction",
            "handle": getattr(WorksheetFunction, "max"),
            "params": args
        },
        "min": {
            "name": "min",
            "type": "function",
            "project": "excel",
            "module": "worksheetfunction",
            "handle": getattr(WorksheetFunction, "min"),
            "params": args
        },
        "small": {
            "name": "small",
            "type": "function",
            "project": "excel",
            "module": "worksheetfunction",
            "handle": getattr(WorksheetFunction, "small"),
            "params": [
                {
                    "name": "arg1",
                    "optional": False,
                    "default": ""
                },
                {
                    "name": "arg2",
                    "optional": False,
                    "default": ""
                }
            ]
        },
        "stdev": {
            "name": "stdev",
            "type": "function",
            "project": "excel",
            "module": "worksheetfunction",
            "handle": getattr(WorksheetFunction, "stdev"),
            "params": args
        },
    }
}
