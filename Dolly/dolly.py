import os
from sys import argv
from xlrd import open_workbook
from docxtpl import DocxTemplate

if len(argv) <= 3:
    print("Not arguments")
    exit()


# Checks if the folder already exists
# and if not then creates it
def create_folder(path):
    try:
        if not os.path.exists(path):
            os.makedirs(path)
    except PermissionError:
        print("Path not found")
        exit()


# Verifies the correctness of the indicated paths
def verification_of_paths(path_1, path_2):
    if not os.path.exists(path_1) or not os.path.exists(path_2):
        print("Path not found")
        exit()
    if not os.access(path_1, os.R_OK) or not os.access(path_2, os.R_OK):
        print("No permission to access the file")
        exit()


# Scans rows of the table and populates the dictionary
# where the key is the name of the colum and the value
# is a record
def parser_excel_file(number):
    data = {}
    j = 0
    for row in sheet.row(0):
        if type(sheet.row(number)[j].value) == float:
            data[row.value.replace(" ", "")] = int(sheet.row(number)[j].value)
        else:
            data[row.value.replace(" ", "")] = sheet.row(number)[j].value

        j = j + 1

    return data


# Substitute the value from the dictionary into
# the tamplate and create as many records as the rows

def substitution_into_a_template():
    for i in range(1, len(sheet.col(0))):
        doc = DocxTemplate(argv[2])
        doc.render(parser_excel_file(i))
        doc.save(argv[3] + '/Invoice{0}.docx'.format(i))


verification_of_paths(argv[1], argv[2])

book = open_workbook(argv[1], on_demand=True)
sheet = book.sheet_by_name(book.sheet_names()[0])

create_folder(argv[3])

substitution_into_a_template()
