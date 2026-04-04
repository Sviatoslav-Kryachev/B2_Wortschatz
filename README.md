# B2 Wortschatz

Скрипт для добавления слов из текстового файла в Google Sheets-словарь B2 по выбранному `Kapitel`.

## Что нужно

- Python с рабочим виртуальным окружением `.venv`
- локальный `config.json`
- локальный JSON-ключ Service Account для Google Sheets

Для репозитория используйте шаблоны:

- [config.example.json](/d:/нужно%20перенести%20на%20hard%20disk/my%20codding/Python/B2_Wortschatz/config.example.json)
- [service-account.example.json](/d:/нужно%20перенести%20на%20hard%20disk/my%20codding/Python/B2_Wortschatz/service-account.example.json)

Настоящие `config.json` и JSON-ключ с секретами не должны попадать в Git.

## Установка зависимостей

Для этого проекта рекомендуется всегда использовать явный интерпретатор из локального окружения:

```powershell
.\.venv\Scripts\python.exe -m pip install google-api-python-client google-auth
```

Если PowerShell блокирует активацию окружения, можно временно разрешить её в текущей сессии:

```powershell
(Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned) ; (& ".\.venv\Scripts\Activate.ps1")
```

Но даже после активации надежнее запускать команды через явный путь:

```powershell
.\.venv\Scripts\python.exe ...
```

Это важно, потому что в PowerShell команды `python`, `pip` и `py` могут указывать на другое окружение.

## Подготовка к GitHub

В репозиторий безопасно коммитить только шаблоны конфигурации, а не реальные секреты.

Локально оставьте:

- `config.json`
- ваш настоящий JSON-ключ Service Account

В Git коммитьте:

- `config.example.json`
- `service-account.example.json`
- `.gitignore`

## Формат входного файла

Файл `input.txt`:

```text
=================================
Существительное:
das Wort — слово
der Begriff — понятие
=================================
Глаголы:
lernen — учить
```

## Проверка без записи

```powershell
.\.venv\Scripts\python.exe add_words.py --kapitel 2 --dry-run
```

## Реальный запуск

```powershell
.\.venv\Scripts\python.exe add_words.py --kapitel 2
```

Для другого `Kapitel` меняется только номер:

```powershell
.\.venv\Scripts\python.exe add_words.py --kapitel 3
.\.venv\Scripts\python.exe add_words.py --kapitel 4 --dry-run
```

## Дополнительные аргументы

```powershell
.\.venv\Scripts\python.exe add_words.py --kapitel 2 --file input.txt --config config.json
```

## Удобный alias для PowerShell

Чтобы не писать длинную команду каждый раз, можно создать функцию:

```powershell
function bw { & ".\.venv\Scripts\python.exe" ".\add_words.py" @args }
```

После этого доступны короткие команды:

```powershell
bw --kapitel 2
bw --kapitel 3 --dry-run
```

## Как работает

- читает категории и слова из `input.txt`
- определяет нужный лист Google Sheets
- ищет нужный `Kapitel`
- заполняет свободные строки в существующих `часть / Teil`
- если места нет, создает новую часть
- дубликаты не добавляет, а повышает приоритет

## Файлы

- [add_words.py](/d:/нужно%20перенести%20на%20hard%20disk/my%20codding/Python/B2_Wortschatz/add_words.py) — точка входа
- [config.json](/d:/нужно%20перенести%20на%20hard%20disk/my%20codding/Python/B2_Wortschatz/config.json) — настройки проекта
- [input.txt](/d:/нужно%20перенести%20на%20hard%20disk/my%20codding/Python/B2_Wortschatz/input.txt) — входные слова
- [SPEC.md](/d:/нужно%20перенести%20на%20hard%20disk/my%20codding/Python/B2_Wortschatz/SPEC.md) — полные требования
