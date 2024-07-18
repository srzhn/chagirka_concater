import argparse
import pandas as pd
from pathlib import Path
from tempfile import TemporaryDirectory
import xlsxwriter
import qrcode

def vals2str(df, sep='_'):
    """
    Версия, которая добавляет конкатенирует все значения.
    """
    labels = []
    for row in df.values:
        labels.append(sep.join(
            list(map(str,
                     filter(lambda x: not pd.isna(x), row)))))
    return labels

def cols_vals2str(df, sep='_'):
    """
    Версия, которая добавляет в строку название колонки если есть.
    """
    labels = []
    empty_columns = list(map(lambda x: x.startswith('Unnamed'), df.columns))
    
    for row in df.values:
        s = []
        for i in range(len(row)):
            # пропускаем, если значения нет
            if pd.isna(row[i]):
                continue

            # если есть колонка - добавляем
            if not empty_columns[i]:
                s.append(df.columns[i])
                
            # добавляем значение
            s.append(str(row[i]))
            
        labels.append(sep.join(s))
    return labels

def make_qr(data, path):
    img = qrcode.make(data=data, border=1)
    img.save(path)
    return path

def save_with_qrs(labels, filepath):
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet('QRcodes')
    with TemporaryDirectory() as tempdir:
        for i, value in enumerate(labels.values, 1):
            qr_path = Path(tempdir, f"{i}.png")
            make_qr(data=value, path=qr_path)
            
            worksheet.write(f'A{i}', value)
            worksheet.insert_image(f'B{i}', qr_path)
        workbook.close()
    
def convert_excel(path: Path, save=True):
    path = Path(path)
    if not path.exists():
        print('File is not found.')
        return

    df = pd.read_excel(path)
    
#     labels = vals2str(df)
    labels = cols_vals2str(df)
    labels = pd.Series(labels)

    if save:
        filepath = path.parent / ('converted_' + path.name)
        # labels.to_excel(filepath, index=0)
        save_with_qrs(labels=labels, filepath=filepath)

        print(f'Done. Path = {filepath}')
    
    return labels

def main():
    path = "/Users/srzhn/smth/chagirka/xlsx_cols2lines/examples/Goat Cave_2021_final.xlsx"

    parser = argparse.ArgumentParser(
                    prog='chagirka_labels_converter')


    parser.add_argument('-p', '--path')

    path = parser.parse_args().path
    if path is None:
        print(f'Path is empty. Close.')
        return 0
    
    convert_excel(path)
    return 0

if __name__=="__main__":
    main()