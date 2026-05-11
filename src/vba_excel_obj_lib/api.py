from vba_excel_obj_lib.worksheet_function import WorksheetFunction

api = {
    "name": "excel",
    "modules": {
        "name": "worksheetfunction",
        "functions": {
            "msgbox": {
                "name": "min",
                "handle": getattr(WorksheetFunction, "min"),
            }
        }
    }
}
