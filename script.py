# import openpyxl module
import openpyxl
import json


class Category:
    def __init__(self):
        self.index = 0
        self.name = ""


class Magazine:
    def __init__(self):
        self.title1 = ""
        self.issn = ""
        self.eissn = ""
        self.title2 = ""
        self.issn2 = ""
        self.eisnn2 = ""
        self.points = ""
        self.categories = []


# Give the location of the file
path = "czasopisma.xlsx"

# to open the workbook
# workbook object is created

startRow = 3

wb_obj = openpyxl.load_workbook(path)
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
categoryFile = open("categoryfile.txt", "w", encoding="utf-8")

dane = []
for i in range(5, max_row):
    licznik = 0
    obj = {}
    for j in range(2, withoutCategory + 1):
        obj[fields[licznik]] = sheet_obj.cell(row=i, column=j).value
        licznik += 1
    dane.append(obj)


json_content = json.dumps(dane, indent=4, sort_keys=True)
file = open("abc.json", "w", encoding="utf-8")
file.write(json_content)
