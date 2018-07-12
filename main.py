import glob, argparse
from work_excel import compare, create_errors_excel
from gooey import Gooey


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
