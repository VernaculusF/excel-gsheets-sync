# Excel Google Sheets Sync

Python CLI для двусторонней синхронизации данных между файлами Excel и Google Sheets. При экспорте инструмент форматирует книгу, а перед импортом создаёт резервный лист.

## Возможности

- Экспорт Google Sheets в `.xlsx`.
- Импорт `.xlsx` в Google Sheets с проверкой структуры.
- Резервное копирование листа перед импортом.
- Форматирование заголовков, границ, выравнивания и ширины столбцов.
- Условное выделение числовых значений.

## Стек

- Python 3.10+
- gspread и Google Auth
- openpyxl
- python-dotenv

## Быстрый старт

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
cp creds.json.example creds.json
```

Укажите идентификатор таблицы и имя листа в `.env`, затем замените содержимое `creds.json` ключом Google Service Account. Таблица должна быть доступна адресу `client_email` из этого ключа.

```bash
python sync.py --export
python sync.py --import --file report.xlsx
```

## Структура проекта

```text
sync.py             CLI и запуск синхронизации
sheets_client.py    операции Google Sheets API
excel_handler.py    чтение и запись Excel
formatter.py        правила форматирования
config.py           загрузка конфигурации
requirements.txt    зависимости Python
```

## Лицензия

MIT
