# B2 Wortschatz

Утилита для импорта слов из `input.txt` в Google Sheets-словарь уровня B2 с разбиением по `Kapitel`, поддержкой `dry-run` и защитой от дублей.

## Что умеет

- читает категории и слова из текстового файла
- раскладывает слова по нужным листам Google Sheets
- добавляет слова в выбранный `Kapitel`
- заполняет свободные слоты в существующих `Teil`
- автоматически создает новую часть, если места больше нет
- не дублирует существующие слова, а повышает их приоритет
- поддерживает безопасную проверку через `--dry-run`

## Стек

- Python
- Google Sheets API
- Service Account JSON credentials

## Быстрый старт

### 1. Клонирование

```bash
git clone https://github.com/Sviatoslav-Kryachev/B2_Wortschatz.git
cd B2_Wortschatz
```

### 2. Создание виртуального окружения

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Если PowerShell блокирует активацию:

```powershell
(Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned) ; (& ".\.venv\Scripts\Activate.ps1")
```

### 3. Настройка конфигурации

Создайте локальные файлы на основе шаблонов:

```powershell
Copy-Item .\config.example.json .\config.json
Copy-Item .\service-account.example.json .\service-account.credentials.json
```

После этого:

- вставьте в `config.json` ваш `sheet_id`
- проверьте путь в `google_sheets.credentials_file`
- вставьте реальные данные service account в `service-account.credentials.json`

Настоящие `config.json` и JSON-ключ не коммитьте в Git. Они уже добавлены в `.gitignore`.

## Запуск

Надежный способ для этого проекта:

```powershell
.\.venv\Scripts\python.exe add_words.py --kapitel 2 --dry-run
.\.venv\Scripts\python.exe add_words.py --kapitel 2
```

Для другого `Kapitel` меняется только номер:

```powershell
.\.venv\Scripts\python.exe add_words.py --kapitel 3
.\.venv\Scripts\python.exe add_words.py --kapitel 4 --dry-run
```

Дополнительные аргументы:

```powershell
.\.venv\Scripts\python.exe add_words.py --kapitel 2 --file input.txt --config config.json
```

## Удобный alias для PowerShell

```powershell
function bw { & ".\.venv\Scripts\python.exe" ".\add_words.py" @args }
```

Примеры:

```powershell
bw --kapitel 2
bw --kapitel 3 --dry-run
```

## Формат входного файла

```text
=================================
Существительное:
das Wort — слово
der Begriff — понятие
=================================
Глаголы:
lernen — учить
```

Обрабатываются только строки с разделителем `—`.

## Файлы конфигурации

- `config.example.json` — шаблон конфигурации проекта
- `service-account.example.json` — шаблон service-account JSON
- `config.json` — локальный рабочий конфиг, не для Git
- `service-account.credentials.json` — локальный ключ Google API, не для Git

## Как это работает

1. Скрипт читает категории из `input.txt`.
2. Определяет, в какой лист Google Sheets отправить каждую категорию.
3. Ищет нужный `Kapitel`.
4. Заполняет свободные строки в существующих частях.
5. Если места нет, создает новую `Teil`.
6. Для дублей не создает новую запись, а повышает приоритет.

## Безопасность

- секреты не хранятся в репозитории
- реальные `config.json` и service-account JSON исключены через `.gitignore`
- для публикации используются только шаблоны `.example.json`

## Полезные файлы

- [README.md](/d:/нужно%20перенести%20на%20hard%20disk/my%20codding/Python/B2_Wortschatz/README.md)
- [SPEC.md](/d:/нужно%20перенести%20на%20hard%20disk/my%20codding/Python/B2_Wortschatz/SPEC.md)
- [config.example.json](/d:/нужно%20перенести%20на%20hard%20disk/my%20codding/Python/B2_Wortschatz/config.example.json)
- [service-account.example.json](/d:/нужно%20перенести%20на%20hard%20disk/my%20codding/Python/B2_Wortschatz/service-account.example.json)
