import xlrd


def get_list_2d_all(sheet):
    return [sheet.row_values(row) for row in range(sheet.nrows)]


def array2plaintext(array):
    text_list = [x for stack in array for x in stack]
    plaintext = ""
    for x in text_list:
        if type(x) != str:
            plaintext += str(x)+","
        else:
            plaintext += x
    return plaintext


def excel2doc(file_path):
    wb = xlrd.open_workbook(file_path)
    # sheetを開く
    sheet_list = wb.sheet_names()
    plain_text =""

    for sheet_name in sheet_list:
        sheet = wb.sheet_by_name(sheet_name)
        plain_text += array2plaintext(get_list_2d_all(sheet))
    return plain_text







