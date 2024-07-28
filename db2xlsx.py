import argparse
import pandas as pd
from pathlib import Path
from src.xlsx import prepare_dataframe, get_xlsx_path, create_xlsx

def main():
    parser = argparse.ArgumentParser(
        prog='chagirka_db2xlsx')

    parser.add_argument('-p', '--path', required=True, help="Путь до эксель-файла. Желательно взять его из адресной строки в проводнике.")
    parser.add_argument('-o', '--filename', help="Название конечного файла.")

    path = Path(parser.parse_args().path)
    filename = parser.parse_args().filename

    if path.suffix.endswith('xlsx'):
        df = pd.read_excel(path)

    elif path.suffix.endswith('csv'):
        df = pd.read_csv(path)
        
    else:
        print('Неизвестное расширение, обрабатывается только csv или xlsx.')
        return 0
    
    df = prepare_dataframe(df)
    
    result_xlsx_path = get_xlsx_path(path, filename)
    create_xlsx(df, filepath=result_xlsx_path, sheet_name='Data')
    
    print(f'Done. \"{result_xlsx_path.absolute()}\"')
    return 0

if __name__=="__main__":
    main()