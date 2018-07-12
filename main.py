import glob, argparse
from gooey import Gooey

from itertools import zip_longest
import xlrd
import xlsxwriter


def colnum_string(n):
    n += 1
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def create_errors_excel(file, sheets):
    try:
        workbook = xlsxwriter.Workbook(file)
        for sheet in sheets:
            worksheet = workbook.add_worksheet(sheet['sheet-name'])
            for error in sheet['errors']:
                if not error['miss-row']:
                    worksheet.write(error['row'], error['col'], error['value1'])
                    worksheet.write(error['row'] + error['max-row'], error['col'], error['value2'])
                    worksheet.write(error['row'] + 2 * error['max-row'] + 1, error['col'], str(
                        error['value1'] == error['value2']) +
                                    f'        EXACT({colnum_string(error["col"])}{error["row"]+1}'
                                    f';{colnum_string(error["col"])}{error["row"]+1 + error["max-row"]})')
                else:
                    worksheet.write(error['row'], 1, 'Потеряная строка')
        workbook.close()
    except PermissionError as e:
        print('нужно закрыть файл' + file)


def compare(file1, file2):
    rb1 = xlrd.open_workbook(file1)
    rb2 = xlrd.open_workbook(file2)

    if len(rb2.sheet_names()) != len(rb1.sheet_names()):
        return 'не равны по количеству шитов'
    sheet_names = list(rb1.sheet_names())

    for i in range(len(sheet_names)):
        sheet1 = rb1.sheet_by_index(i)
        sheet2 = rb2.sheet_by_index(i)
        max_row = max(sheet1.nrows, sheet2.nrows)
        errors = []
        for rownum in range(max_row):
            if rownum < sheet1.nrows and rownum < sheet2.nrows:
                row_rb1 = sheet1.row_values(rownum)
                row_rb2 = sheet2.row_values(rownum)
                row = []
                flag = True
                for colnum, (c1, c2) in enumerate(zip_longest(row_rb1, row_rb2)):
                    row.append(
                        {'miss-row': False, 'row': rownum, 'col': colnum, 'value1': c1, 'value2': c2,
                         'max-row': max_row + 1})
                    if c1 != c2:
                        flag = False
                if not flag:
                    errors += row
            else:
                errors += [{'miss-row': True, 'row': rownum, 'max-row': max_row + 1}]
        if len(errors) > 0:
            yield {'sheet-name': sheet_names[i], 'errors': errors}

    return 'Good'


def get_excels(dir_path1, dir_path2):
    for path in zip(glob.glob(f"{dir_path1.lower()}\\*.xlsx"), glob.glob(f"{dir_path2.lower()}\\*.xlsx")):
        yield path


@Gooey
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-p1", "--path-to-1", type=str,
                        help="Введите путь до 1 папки")
    parser.add_argument("-p2", "--path-to-2", type=str,
                        help="Введите путь до 2 папки")
    parser.add_argument("-to", "--path-to-errors", type=str,
                        help="Путь для папки ошибочных файлов")
    args = parser.parse_args()
    path1 = args.path_to_1
    path2 = args.path_to_2
    pth_to = args.path_to_errors
    for file1, file2 in get_excels(path1, path2):
        create_errors_excel(pth_to + file1.split('\\')[-1][:-5] + 'сравнение.xlsx',
                            compare(file1, file2))


if __name__ == '__main__':
    # path = r'C:\Work\compare-excel'
    # for file1, file2 in get_excels(path + r'\1', path + r'\2'):
    #     create_errors_excel(path + r'\Ошибки\\' + file1.split('\\')[-1][:-5] + 'сравнение.xlsx', compare(file1, file2))
    main()
