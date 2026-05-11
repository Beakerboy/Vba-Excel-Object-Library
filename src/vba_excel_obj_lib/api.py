from vba_excel_obj_lib.worksheet_function import WorksheetFunction

api = {
    "name": "excel",
    "type": "project",
    "modules": {
        "name": "worksheetfunction",
        "type": "module",
        "functions": {
            "msgbox": {
                "name": "min",
                "type": "function",
                "handle": getattr(WorksheetFunction, "min"),
            }
        }
    }
}
