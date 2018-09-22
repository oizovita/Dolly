import os
import json
import pdfrw
import argparse
from xlrd import open_workbook
from docxtpl import DocxTemplate

FILE_DOCX = 'docx'
FILE_PDF = 'pdf'
ANNOT_KEY = '/Annots'
ANNOT_FIELD_KEY = '/T'
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'


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
        if type(sheet.row(number)[j].value) == float and sheet.row(number)[j].value == int(sheet.row(number)[j].value):
            data[row.value.replace(" ", "")] = int(sheet.row(number)[j].value)
        else:
            data[row.value.replace(" ", "")] = sheet.row(number)[j].value

        j = j + 1

    return data


# Substitute the value from the dictionary into
# the tamplate and create as many records as the rows

def substitution_into_a_template(template, output_folder):
    if template[-3:] == FILE_PDF:
        print('Inavalif file format')
        exit()
    for i in range(1, len(sheet.col(0))):
        doc = DocxTemplate(template)
        doc.render(parser_excel_file(i))
        doc.save(output_folder + '/Invoice{0}.docx'.format(i))


# Reads a pdf-file, searches in it for the same keys from the file
# and dictionary and replaces the values in the file. Saves a new file
def record_in_pdf_template(template, output_folder, data_dict):
    if template[-4:] == FILE_DOCX:
        print('Inavalif file format')
        exit()
    template_pdf = pdfrw.PdfReader(template)
    annotations = template_pdf.pages[0][ANNOT_KEY]
    for annotation in annotations:
        if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
            if annotation[ANNOT_FIELD_KEY]:
                key = annotation[ANNOT_FIELD_KEY][1:-1]
                if key in data_dict.keys():
                    annotation.update(pdfrw.PdfDict(V='{}'.format(data_dict[key])))

    pdfrw.PdfWriter().write(output_folder, template_pdf)


# Reads the json file by replacing the values in the dictionary
# and passes it to record_in_pdf_template
def create_pdf_templates(template, output_folder, json_file):
    data_json = json.load(open(json_file))
    for i in range(1, len(sheet.col(0))):
        data_dict = {}
        for data in data_json:
            for exel_row in parser_excel_file(i):
                if data_json[data] == exel_row:
                    data_dict[data] = parser_excel_file(i)[exel_row]

        record_in_pdf_template(template, output_folder + '/Invoice{0}.pdf'.format(i), data_dict)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Template script')
    parser.add_argument('-t', action="store", default=None, dest="type_file", help="type file: docx - d, pdf - p")
    parser.add_argument('data', action="store", help="path to the data file")
    parser.add_argument('template', action="store", help="path to the template file")
    parser.add_argument('output_folder', action="store", help="folder path with output")
    parser.add_argument('-f', action="store", default="", dest="json_file", help="path to the json file")
    args = parser.parse_args()

    verification_of_paths(args.data, args.template)

    book = open_workbook(args.data, on_demand=True)
    sheet = book.sheet_by_name(book.sheet_names()[0])

    create_folder(args.output_folder)

    if args.type_file is None:
        print('Enter the file type')
        exit()
    elif args.type_file == 'd':
        substitution_into_a_template(args.template, args.output_folder)
    elif args.type_file == 'p':
        if not os.path.exists(args.json_file):
            print("Path not found")
            exit()

        create_pdf_templates(args.template, args.output_folder, args.json_file)
