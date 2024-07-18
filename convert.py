import argparse
import pandas as pd
from pathlib import Path
from src.concat import join_cells, join_cells_with_colnames
from src.xlsx import create_xlsx, create_xlsx_with_qrs, get_xlsx_path
from typing import Optional


def main():
    parser = argparse.ArgumentParser(
                    prog='chagirka_labels_converter')


    parser.add_argument('-p', '--path', required=True)
    parser.add_argument('-o', '--filename')
    parser.add_argument('-q', '--qrcode', action='store_true')
    parser.add_argument('-c', '--colnames', action='store_true')
    
    path = Path(parser.parse_args().path)
    filename = parser.parse_args().filename
    
    
    if path is None or not path.exists():
        print(f'File {path.absolute()} doesn\'t exist.')
        return 0

    
    df = pd.read_excel(path)
    
#     labels = join_cells(df)
    labels = join_cells_with_colnames(df)
    
    result_xlsx_path = get_xlsx_path(path, filename)
    
    # create_xlsx(labels, result_xlsx_path, sheetname='Labels')
    create_xlsx_with_qrs(labels, result_xlsx_path, sheetname='QRcodes')
    
    print(f'Done. \"{result_xlsx_path.absolute()}\"')
    return 0

if __name__=="__main__":
    main()