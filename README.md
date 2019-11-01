# Скрипт для перевода данных о чеках из формата XLSX в JSON

## Инструкция по установке для 64-битной Windows

Шаг 1. Скачать установщик Python 3.7: [https://www.python.org/ftp/python/3.7.5/python-3.7.5-amd64.exe](https://www.python.org/ftp/python/3.7.5/python-3.7.5-amd64.exe)

Шаг 2. Установить Python 3.7, при установке включить опцию «Add Python 3.7 to PATH», при этом рекомендуется устанавливать в папку, не содержащую в пути символов кириллицы

Шаг 3. Установить Poetry, выполнив в командной строке команду:

```bat
curl -sSL https://raw.githubusercontent.com/sdispater/poetry/master/get-poetry.py | python
```

Шаг 4. Запустить командную строку в папке с этим файлом README и выполнить там команды:

```bat
poetry install --no-dev
```

## Инструкция по быстрому использованию

Положить файл в директорию `input`, назвать `input.xlsx`, запустить файл `run.bat`, выходной файл будет в папке `output` с именем `output.json`.
