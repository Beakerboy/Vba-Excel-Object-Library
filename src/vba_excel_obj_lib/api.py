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
                    "handle": getattr(WorksheetFunction, "min"),
                }
            }
        }
    }
}
