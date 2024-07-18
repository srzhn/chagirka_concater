import xlsxwriter
from tempfile import TemporaryDirectory
from pathlib import Path
import pandas as pd
import qrcode
# from hashlib import sha256
from typing import Optional
from datetime import datetime


def make_qr(data, path):
    img = qrcode.make(data=data, border=1)
    img.save(path)
    return path


def create_xlsx(labels: pd.Series,
                filepath: Path,
                sheetname: str = 'Labels'):
    labels.to_excel(filepath, index=0, sheet_name=sheetname)


def create_xlsx_with_qrs(labels: pd.Series,
                         filepath: Path,
                         sheetname: str = 'QRcodes'):
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet(sheetname)

    # Сформированные qr-коды сохраняем во временную папку.
    with TemporaryDirectory() as tempdir:
        for index, value in enumerate(labels.values, 1):
            qr_path = Path(tempdir, f"qrcode_{index}.png")
            make_qr(data=value, path=qr_path)

            # Пишем label в первый столбец
            worksheet.write(f'A{index}', value)

            # Вставляем qr-код во второй столбец
            worksheet.insert_image(f'B{index}', qr_path)

        # Пока существует папка - закрываем xlsx.
        workbook.close()


def get_xlsx_path(path: Path,
                  filename: Optional[str]=None):
    
    if filename is None:
        filename = 'converted_' + path.name
        
    if not filename.endswith('xlsx'):
        filename += '.xlsx'

    filepath = Path(path.parent, filename)

    if filepath.exists():
        filepath = Path(path.parent, f'converted_{datetime.now().strftime(format="%y-%m-%d-%H%M%S")}_{path.name}')
        print(f'File {filename} already exists. Result file is renamed to {filepath.name}.')

    return filepath
