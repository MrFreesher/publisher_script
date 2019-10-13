# import openpyxl module
import openpyxl
import json
import sys


if len(sys.argv) < 2:
    print("Execute with parameters")
else:

    # Give the location of the file
    inputPath = sys.argv[1]
    outputPath = sys.argv[2]
    startRow = 3

    wb_obj = openpyxl.load_workbook(inputPath)
    sheet_obj = wb_obj.active
    max_col = sheet_obj.max_column
    max_row = sheet_obj.max_row
    categories = []
    withoutCategory = 0
    # Loop will print all columns name
    for i in range(1, max_col):
        cell_obj = sheet_obj.cell(row=3, column=i)
        if cell_obj.value is not None:
            categories.append(cell_obj.value)
        else:
            withoutCategory += 1

    print(withoutCategory)
    fields = []
    for i in range(2, withoutCategory + 1):
        field = sheet_obj.cell(row=4, column=i).value
        if field in fields:
            fields.append("{}_2".format(field))
        else:
            fields.append(field)

    print(fields)
    print(categories[0])

    dane = []
    for i in range(5, max_row):
        licznik = 0
        obj = {}
        categoryIndex = 0
        for j in range(2, withoutCategory + 1):
            obj[fields[licznik]] = sheet_obj.cell(row=i, column=j).value
            licznik += 1
        rowCategories = []
        for j in range(withoutCategory + 1, max_col):
            temp_cell = sheet_obj.cell(row=i, column=j).value
            if temp_cell is "x":
                rowCategories.append(categories[categoryIndex])
            categoryIndex += 1
        obj["dziedziny"] = rowCategories
        dane.append(obj)

    json_content = json.dumps(dane, indent=4, sort_keys=True)
    file = open(outputPath, "w", encoding="utf-8")
    file.write(json_content)

