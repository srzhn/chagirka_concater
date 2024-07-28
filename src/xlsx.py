import xlsxwriter
from tempfile import TemporaryDirectory
from pathlib import Path
import pandas as pd
import qrcode
# from hashlib import sha256
from typing import Optional, Union
from datetime import datetime


def make_qr(data: str, path: Path, border: float=1, box_size: float=3):
    img = qrcode.make(data=data, 
                      border=border, 
                      box_size=box_size)
    img.save(str(path.absolute()))
    return path

def prepare_dataframe(data):
    columns = ['[site]', '[prefix]', '[point_id]', '[point_sub_id]', '[layer]', '[sublayer]', '[tag]', '[data_x]', '[data_y]', '[data_z]']
    
    df = data.loc[:, columns]
    
    df.columns = [col[1:-1] for col in df.columns]
    
    df['data_x'] = df['data_x'].round(2)
    df['data_y'] = df['data_y'].round(2)
    df['data_z'] = df['data_z'].round(2)
    
    df = df.groupby(['site', 'prefix', 'point_id']).first().reset_index().drop(columns=['point_sub_id'])
    
    column_mapper = {
                     'point_id': 'N',
                     'layer': 'L',
                     'sublayer': 'h',
                    #  'tag': 'Unnamed 0',
                    #  'site': 'Unnamed 1',
                    #  'tag': '',
                    #  'site': '',
                    #  'prefix': '',
                     'data_x': 'X',
                     'data_y': 'Y',
                     'data_z': 'Z',
                     }
    df = df.rename(column_mapper, axis=1)
    
    # df.to_excel('results/2024-07-28_test.xlsx', index=0)
    
    return df


def create_xlsx(data: Union[pd.Series, pd.DataFrame],
                filepath: Path,
                sheet_name: str = 'Labels'):
    data.to_excel(filepath, index=0, sheet_name=sheet_name)


def create_xlsx_with_qrs(labels: pd.Series,
                         filepath: Path,
                         sheet_name: str = 'QRcodes',
                         qr_size=3):
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet(sheet_name)

    # Сформированные qr-коды сохраняем во временную папку.
    with TemporaryDirectory() as tempdir:
        for index, value in enumerate(labels.values, 1):
            qr_path = Path(tempdir, f"qrcode_{index}.png")
            make_qr(data=value, path=qr_path, box_size=qr_size)

            # Пишем label в первый столбец
            worksheet.write(f'A{index}', value)

            # Вставляем qr-код во второй столбец
            worksheet.insert_image(f'B{index}', qr_path)
        
        worksheet.autofit()

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
