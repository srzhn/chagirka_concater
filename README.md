# Converter

## Installation

install python

install libraries from requirements.txt: 
`python3 -m pip install -r requirements`

## Usage

`python3 convert.py -p "PATH_TO_FILE"`

- `-p/--path` - Путь до файла. Обрабатывается только первый лист в файле.
- `-o/--filename` - свое название файла. Результат будет лежать в той же папке.
- `-q/--no_qrcode` - qr-code не будет генерироваться.
- `-c/--no-colnames` - название колонок не будет добавляться в label.