import argparse
import pandas as pd
from pathlib import Path
from src.concat import join_cells, join_cells_with_colnames
from src.xlsx import create_xlsx, create_xlsx_with_qrs, get_xlsx_path


def main():
    parser = argparse.ArgumentParser(
        prog='chagirka_labels_converter')

    parser.add_argument('-p', '--path', required=True, help="Путь до эксель-файла. Желательно взять его из адресной строки в проводнике.")
    parser.add_argument('-o', '--filename', help="Название конечного файла.")
    parser.add_argument('-q', '--no-qrcode', action='store_true', help="Если есть, то не будут генерироваться QR-коды.")
    parser.add_argument('-c', '--no-colnames', action='store_true', help="Если есть, то в сгенерированном лэйбл не будут добавляться названия колонок.")

    path = Path(parser.parse_args().path)
    filename = parser.parse_args().filename
    skip_qr_codes = parser.parse_args().no_qrcode
    skip_col_names = parser.parse_args().no_colnames

    if path.suffix == '':
        path = Path(path.parent, path.stem + '.xlsx')

    if not path.suffix.endswith('xlsx'):
        print(f'Неподходящее расширение {path.suffix}. Обрабатываются только `.xlsx`')
        return 0

    if path is None or not path.exists():
        print(f'Файл {path.absolute()} не найден.')
        return 0

    # Читаем эксель
    df = pd.read_excel(path)

    # Конкатенируем строку в зависимости от алгоритма: без названия колонок или с ними.
    if skip_col_names:
        labels = join_cells(df)
    else:
        labels = join_cells_with_colnames(df)

    # Генерируем название нового файла (либо на основе filename, либо добавлением префикса к существующему.)
    # Файлы не перезаписываются, при совпадении добавляется временная характеристика в название (см. src/xlsx.py).
    result_xlsx_path = get_xlsx_path(path, filename)

    # Сохраняем лэйблы в новый файл.
    if skip_qr_codes:
        create_xlsx(labels, result_xlsx_path, sheet_name='Labels')
    else:
        create_xlsx_with_qrs(labels, result_xlsx_path, sheet_name='QRcodes', qr_size=3)

    print(f'Done. \"{result_xlsx_path.absolute()}\"')
    return 0


if __name__ == "__main__":
    main()
