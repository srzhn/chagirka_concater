# Converter

## Installation

1. Install [python](https://github.com/PackeTsar/Install-Python/blob/master/README.md)
2. install libraries from `requirements.txt`: 
`python -m pip install -r requirements.txt`


## Colab

Ссылка на [colab-notebook](https://colab.research.google.com/drive/1A69cbJopPOcaNy94iUfKdwxoFV8S-2Br?usp=sharing) для запуска в интернете. Для использования необходимо склонировать в свой Google Drive.

## Usage

### Генерация лэйблов с qr-кодами.

`python3 generate_labels.py -p PATH [-o FILENAME] [-q] [-c]`

По умолчанию, обрабатывается первый лист в эксель-файле.

- `-h/--help` - Справка о функционале.
- `-p/--path` - Путь до файла. Обрабатывается только первый лист в файле.
- `-o/--filename` - свое название файла. Результат будет лежать в той же папке.
- `-q/--no_qrcode` - qr-code не будет генерироваться.
- `-c/--no-colnames` - название колонок не будет добавляться в label.



### Генерация файла нужного формата из csv/xlsx.

`python3 db2xlsx.py -p PATH [-o FILENAME]`

- `-p/--path` - Путь до файла. Обрабатывается только первый лист в файле.
- `-o/--filename` - свое название файла. Результат будет лежать в той же папке.



## КАК НАСТРОИТЬ ПЕЧАТЬ ДЛЯ ПРИНТЕРА.
- Все коллонтитулы на 0, левый на 0.4. Снизу установить на 1.5-2, чтоб не пучаталось сразу несколько на одном. Количество строк должно совпадать с количеством распечатанного.
- Печатать все столбцы, масштаб сам подстроится.
- Вид -> страничный режим.
- Сначала проверить на одной строке перед запуском всего.