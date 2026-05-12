from .worksheet_function import worksheet_function
arg1 = {
    "name": "arg1",
    "optional": False,
    "default": ""
}
args = [arg1]
i = 2
for n in range(29):
    args.append({
        "name": "arg" + i,
        "optional": True,
        "default": ""
    })

api = {
    "name": "excel",
    "type": "project",
    "modules": {
        "worksheetfunction": worksheet_function.api
    }
}
