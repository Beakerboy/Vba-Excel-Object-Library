from vba_excel_obj_lib.worksheet_function import WorksheetFunction
arg1 = {
    "name": "arg1",
    "optional": False,
    "default": ""
}
args = [arg1]
i = 2
for n in range(29):
    args.append({
        "name": "arg2",
        "optional": True,
        "default": ""
    })

api = {
    "name": "excel",
    "type": "project",
    "modules": {
        "worksheetfunction": {
            "name": "worksheetfunction",
            "type": "module",
            "functions": {
                "min": {
                    "name": "min",
                    "type": "function",
                    "project": "excel",
                    "module": "worksheetfunction",
                    "handle": getattr(WorksheetFunction, "min"),
                     "params": args
                }
            }
        }
    }
}
