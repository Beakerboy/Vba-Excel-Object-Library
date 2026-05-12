from vba_excel_obj_lib.worksheet_function import WorksheetFunction

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
                     "params": [
                         {
                            "name": "arg1",
                            "optional": False,
                            "default": ""
                        },
                        {
                            "name": "arg2",
                            "optional": True,
                            "default": ""
                        }
                     ]
                }
            }
        }
    }
}
