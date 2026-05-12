from .worksheet_function import api as worksheet_function_api


api = {
    "name": "excel",
    "type": "project",
    "classes": {
        "worksheetfunction": worksheet_function_api
    }
}
