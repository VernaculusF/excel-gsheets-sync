# Excel ↔ Google Sheets Sync

Bidirectional data synchronization project between Google Sheets and Excel files with automatic formatting.

**English** | [Русский](README.ru.md)

## 🎯 Features

- ✅ **Export from Google Sheets to Excel** with professional formatting
- ✅ **Import from Excel to Google Sheets** with structure validation
- ✅ **Automatic formatting**: headers, alignment, borders
- ✅ **Conditional formatting**: values > 100 highlighted in green
- ✅ **Auto-adjust column width** based on content
- ✅ **Automatic backup** before import
- ✅ **Data structure validation**
- ✅ **Error handling** with clear logging

## 📋 Requirements

- Python 3.10+
- Google Service Account with Google Sheets API access

## 🚀 Quick Start

### 1. Clone the repository

```bash
git clone https://github.com/VernaculusF/excel-gsheets-sync.git
cd excel-gsheets-sync
```

### 2. Create virtual environment

**Windows (PowerShell):**
```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**Windows (CMD):**
```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

**Linux/macOS:**
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Get Google Service Account credentials

#### Step 1: Create project in Google Cloud Console

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. In the navigation menu, select **APIs & Services** → **Library**
4. Find and enable **Google Sheets API**

#### Step 2: Create Service Account

1. Go to **APIs & Services** → **Credentials**
2. Click **Create Credentials** → **Service Account**
3. Fill in:
   - **Service account name**: e.g., `excel-sheets-sync`
   - **Service account ID**: will be auto-generated
4. Click **Create and Continue**
5. Role can be left empty (or select **Editor**)
6. Click **Done**

#### Step 3: Create key

1. In the Service Accounts list, find the created account
2. Click on it, go to the **Keys** tab
3. Click **Add Key** → **Create new key**
4. Select **JSON** format
5. Click **Create** — the file will automatically download

#### Step 4: Configure the key

1. Rename the downloaded file to `creds.json`
2. Move the file to the project root folder

### 5. Configure Google Sheets access

1. Open the `creds.json` file
2. Find the `"client_email"` line, copy the email address
   ```json
   "client_email": "excel-sheets-sync@your-project.iam.gserviceaccount.com"
   ```
3. Open your Google spreadsheet
4. Click the **Share** button
5. Paste the email from `client_email`
6. Select **Editor** role
7. Click **Send** (without notification)

### 6. Configure environment variables

1. Copy `.env.example` to `.env`:
   ```bash
   copy .env.example .env    # Windows
   cp .env.example .env      # Linux/macOS
   ```

2. Open `.env` and fill in:
   ```env
   SPREADSHEET_ID=your_spreadsheet_id_here
   SHEET_NAME=Sheet1
   ```

**How to get SPREADSHEET_ID:**
- Open your Google spreadsheet
- Copy the ID from URL:
  ```
  https://docs.google.com/spreadsheets/d/1abc_DEFGH-ijk123LMN456opq/edit
                                      ^^^^^^^^^^^^^^^^^^^^^^^^
                                      This is your SPREADSHEET_ID
  ```

**SHEET_NAME:**
- Sheet name in the spreadsheet (visible at the bottom of the screen)
- Default: `Sheet1`

## 📖 Usage

### Mode 1: Export (Google Sheets → Excel)

Export data from Google Sheets to Excel file with formatting:

```bash
python sync.py --export
```

**Result:**
- Creates file `report_YYYYMMDD_HHMMSS.xlsx`
- Applies formatting:
  - Headers: bold font, gray background
  - Auto column width
  - Alignment (text left, numbers right)
  - Cell borders
  - Conditional formatting: cells with value > 100 highlighted in green

### Mode 2: Import (Excel → Google Sheets)

Import data from Excel to Google Sheets:

```bash
python sync.py --import --file "report.xlsx"
```

**With full path:**
```bash
python sync.py --import --file "C:\Users\user\Documents\data.xlsx"
```

**What happens:**
1. File existence is checked
2. Column count is validated against the spreadsheet
3. Google Sheets backup is created (new sheet `backup_YYYYMMDD_HHMMSS`)
4. Data is imported (complete overwrite)

## 📁 Project Structure

```
excel-gsheets-sync/
├── sync.py                # main script with CLI argument parsing
├── sheets_client.py       # Google Sheets API operations
├── excel_handler.py       # Excel read/write/processing
├── formatter.py           # Excel styling configuration
├── config.py              # .env configuration loader
├── requirements.txt       # Python dependencies
├── .env                   # environment variables (NOT IN GIT!)
├── .env.example           # configuration example
├── creds.json             # Google credentials (NOT IN GIT!)
├── creds.json.example     # credentials structure example
├── .gitignore             # git exclusions
├── README.md              # documentation (English)
├── README.ru.md           # documentation (Russian)
├── plan.md                # development plan
└── .venv/                 # virtual environment (NOT IN GIT!)
```

## 🎨 Excel Formatting

### Headers (first row)
- **Font**: bold, 11pt
- **Fill**: gray color (#D3D3D3)
- **Alignment**: center
- **Borders**: thin black

### Data
- **Alignment**: 
  - Text: left
  - Numbers: right
- **Borders**: thin black

### Conditional Formatting
- Cells with value > 100 → green fill (#90EE90)
- Works for all numeric columns
- Applied to all cells except headers

### Auto Column Width
- Minimum width: 10
- Maximum width: 50
- Automatically adjusts to content

## 🔧 Customizing Styles

To change styles, edit the `formatter.py` file:

```python
# Headers
HEADER_FONT = Font(bold=True, size=11)
HEADER_FILL = PatternFill(start_color="D3D3D3", ...)

# Conditional formatting
HIGHLIGHT_FILL = PatternFill(start_color="90EE90", ...)
HIGHLIGHT_THRESHOLD = 100  # Highlighting threshold
```

## ⚠️ Error Handling

### Empty Cells
- Saved as empty string `''`
- Don't cause synchronization errors

### Structure Validation on Import
Excel column count must match Google Sheets:

```
✗ ERROR: Structure mismatch:
  Expected: 5 columns
  In Excel: 7 columns
```

### Backup
A backup is automatically created before import:

```
✓ Backup created: backup_20260220_143052
```

The backup is created as a new sheet in the same spreadsheet.

## 🐛 Troubleshooting

### Error: "File creds.json not found"
- Make sure the `creds.json` file is in the project root
- Check that the file is not renamed

### Error: "Spreadsheet not found or no access"
- Check the `SPREADSHEET_ID` in `.env`
- Make sure the spreadsheet is shared with the email from `creds.json`
- Check that Google Sheets API is enabled in Google Cloud Console

### Error: "Worksheet not found"
- Check the sheet name (`SHEET_NAME`) in `.env`
- The name must match exactly (including case and spaces)

### Error: "Environment variable not found"
- Make sure the `.env` file is created and filled
- Check for typos in variable names

### Error: "Structure mismatch"
- Make sure the column count in Excel matches Google Sheets
- Check the first row (headers) in both files

## 📊 Log Examples

### Successful Export
```
[2026-02-20 14:30:45] [INFO] ============================================================
[2026-02-20 14:30:45] [INFO] MODE: EXPORT (Google Sheets → Excel)
[2026-02-20 14:30:45] [INFO] ============================================================
[2026-02-20 14:30:45] [INFO] Authenticating to Google Sheets...
[2026-02-20 14:30:46] [INFO] ✓ Authentication successful
[2026-02-20 14:30:46] [INFO] Opening spreadsheet...
[2026-02-20 14:30:47] [INFO] ✓ Spreadsheet opened
[2026-02-20 14:30:47] [INFO] Reading data from Google Sheets...
[2026-02-20 14:30:48] [INFO] ✓ Read 150 rows
[2026-02-20 14:30:48] [INFO] Creating Excel file report_20260220_143048.xlsx...
[2026-02-20 14:30:48] [INFO] ✓ Data written
[2026-02-20 14:30:49] [INFO] Applying formatting to Excel sheet...
[2026-02-20 14:30:49] [INFO] ✓ Formatting successfully applied
[2026-02-20 14:30:49] [INFO] ✓ File saved: report_20260220_143048.xlsx
[2026-02-20 14:30:49] [INFO] ============================================================
[2026-02-20 14:30:49] [INFO] ✓ EXPORT COMPLETED
[2026-02-20 14:30:49] [INFO]   File: g:\Code\Google Sheets\excel-gsheets-sync\report_20260220_143048.xlsx
[2026-02-20 14:30:49] [INFO]   Rows: 150
[2026-02-20 14:30:49] [INFO] ============================================================
```

### Successful Import
```
[2026-02-20 14:35:12] [INFO] ============================================================
[2026-02-20 14:35:12] [INFO] MODE: IMPORT (Excel → Google Sheets)
[2026-02-20 14:35:12] [INFO] ============================================================
[2026-02-20 14:35:12] [INFO] Reading Excel file data.xlsx...
[2026-02-20 14:35:12] [INFO] ✓ Read 120 rows
[2026-02-20 14:35:12] [INFO] Authenticating to Google Sheets...
[2026-02-20 14:35:13] [INFO] ✓ Authentication successful
[2026-02-20 14:35:13] [INFO] Validating data structure...
[2026-02-20 14:35:14] [INFO] ✓ Excel file structure matches Google Sheets
[2026-02-20 14:35:14] [INFO] Creating backup 'backup_20260220_143514'...
[2026-02-20 14:35:15] [INFO] ✓ Backup created: backup_20260220_143514
[2026-02-20 14:35:15] [INFO] Writing data to Google Sheets (120 rows)...
[2026-02-20 14:35:17] [INFO] ✓ Data successfully written
[2026-02-20 14:35:17] [INFO] ============================================================
[2026-02-20 14:35:17] [INFO] ✓ IMPORT COMPLETED
[2026-02-20 14:35:17] [INFO]   File: g:\Code\Google Sheets\excel-gsheets-sync\data.xlsx
[2026-02-20 14:35:17] [INFO]   Rows: 120
[2026-02-20 14:35:17] [INFO]   Backup: backup_20260220_143514
[2026-02-20 14:35:17] [INFO] ============================================================
```

## 🔒 Security

### Files NOT for Git:
- `.env` — contains spreadsheet ID
- `creds.json` — contains Google private keys
- `*.xlsx` — Excel data files
- `.venv/` — virtual environment

All these files are already added to `.gitignore`.

### Recommendations:
- Never commit `creds.json` to the repository
- Don't share `creds.json` via messengers or email
- Regularly update access keys
- Use minimal access rights for Service Account

## 📝 License

Project created for educational purposes.

## 👨‍💻 Author

Developed using GitHub Copilot (Claude Sonnet 4.5)

---

**Happy syncing! 🚀**


```bash
pip install -r requirements.txt
```

### 4. Получение учетных данных Google Service Account

#### Шаг 1: Создание проекта в Google Cloud Console

1. Перейдите на [Google Cloud Console](https://console.cloud.google.com/)
2. Создайте новый проект или выберите существующий
3. В меню навигации выберите **APIs & Services** → **Library**
4. Найдите и включите **Google Sheets API**

#### Шаг 2: Создание Service Account

1. Перейдите в **APIs & Services** → **Credentials**
2. Нажмите **Create Credentials** → **Service Account**
3. Заполните:
   - **Service account name**: например, `excel-sheets-sync`
   - **Service account ID**: автоматически сгенерируется
4. Нажмите **Create and Continue**
5. Роль можно не указывать (или выбрать **Editor**)
6. Нажмите **Done**

#### Шаг 3: Создание ключа

1. В списке Service Accounts найдите созданный аккаунт
2. Нажмите на него, перейдите на вкладку **Keys**
3. Нажмите **Add Key** → **Create new key**
4. Выберите формат **JSON**
5. Нажмите **Create** — файл автоматически скачается

#### Шаг 4: Настройка ключа

1. Переименуйте скачанный файл в `creds.json`
2. Переместите файл в корневую папку проекта:
   ```
   g:\Code\Google Sheets\excel-gsheets-sync\creds.json
   ```

### 5. Настройка доступа к Google Sheets

1. Откройте файл `creds.json`
2. Найдите строку `"client_email"`, скопируйте email адрес
   ```json
   "client_email": "excel-sheets-sync@your-project.iam.gserviceaccount.com"
   ```
3. Откройте вашу Google таблицу
4. Нажмите кнопку **Share** (Поделиться)
5. Вставьте email из `client_email`
6. Выберите роль **Editor** (Редактор)
7. Нажмите **Send** (без отправки уведомления)

### 6. Настройка переменных окружения

1. Скопируйте `.env.example` в `.env`:
   ```bash
   copy .env.example .env
   ```

2. Откройте `.env` и заполните:
   ```env
   SPREADSHEET_ID=your_spreadsheet_id_here
   SHEET_NAME=Sheet1
   ```

**Как получить SPREADSHEET_ID:**
- Откройте вашу Google таблицу
- Скопируйте ID из URL:
  ```
  https://docs.google.com/spreadsheets/d/1abc_DEFGH-ijk123LMN456opq/edit
                                      ^^^^^^^^^^^^^^^^^^^^^^^^
                                      Это ваш SPREADSHEET_ID
  ```

**SHEET_NAME:**
- Название листа в таблице (видно внизу экрана)
- По умолчанию: `Sheet1` или `Лист1`

## 📖 Использование

### Режим 1: Экспорт (Google Sheets → Excel)

Экспортировать данные из Google Sheets в Excel файл с форматированием:

```bash
python sync.py --export
```

**Результат:**
- Создается файл `report_YYYYMMDD_HHMMSS.xlsx`
- Применяется форматирование:
  - Заголовки: жирный шрифт, серая заливка
  - Автоматическая ширина колонок
  - Выравнивание (текст влево, числа вправо)
  - Границы всех ячеек
  - Условное форматирование: ячейки со значением > 100 выделяются зеленым

### Режим 2: Импорт (Excel → Google Sheets)

Импортировать данные из Excel в Google Sheets:

```bash
python sync.py --import --file "report.xlsx"
```

**С полным путем:**
```bash
python sync.py --import --file "C:\Users\user\Documents\data.xlsx"
```

**Что происходит:**
1. Проверяется существование файла
2. Проверяется совпадение количества колонок с таблицей
3. Создается резервная копия Google Sheets (новый лист `backup_YYYYMMDD_HHMMSS`)
4. Данные импортируются (полная перезапись)

## 📁 Структура проекта

```
excel-gsheets-sync/
├── sync.py                # основной скрипт с парсингом аргументов
├── sheets_client.py       # работа с Google Sheets API
├── excel_handler.py       # чтение/запись/обработка Excel
├── formatter.py           # настройки стилей для Excel
├── config.py              # загрузка конфигурации из .env
├── requirements.txt       # зависимости Python
├── .env                   # переменные окружения (НЕ В GIT!)
├── .env.example           # пример конфигурации
├── creds.json             # учетные данные Google (НЕ В GIT!)
├── creds.json.example     # пример структуры ключей
├── .gitignore             # исключения для git
├── README.md              # документация
├── plan.md                # план разработки
└── .venv/                 # виртуальное окружение (НЕ В GIT!)
```

## 🎨 Форматирование Excel

### Заголовки (первая строка)
- **Шрифт**: жирный, 11pt
- **Заливка**: серый цвет (#D3D3D3)
- **Выравнивание**: по центру
- **Границы**: тонкие черные

### Данные
- **Выравнивание**: 
  - Текст: влево
  - Числа: вправо
- **Границы**: тонкие черные

### Условное форматирование
- Ячейки со значением > 100 → зеленая заливка (#90EE90)
- Работает для любых числовых колонок
- Применяется ко всем ячейкам, кроме заголовков

### Автоширина колонок
- Минимальная ширина: 10
- Максимальная ширина: 50
- Автоматически подстраивается под содержимое

## 🔧 Настройка стилей

Чтобы изменить стили, отредактируйте файл `formatter.py`:

```python
# Заголовки
HEADER_FONT = Font(bold=True, size=11)
HEADER_FILL = PatternFill(start_color="D3D3D3", ...)

# Условное форматирование
HIGHLIGHT_FILL = PatternFill(start_color="90EE90", ...)
HIGHLIGHT_THRESHOLD = 100  # Порог для выделения
```

## ⚠️ Обработка ошибок

### Пустые ячейки
- Сохраняются как пустая строка `''`
- Не вызывают ошибок при синхронизации

### Проверка структуры при импорте
Количество колонок в Excel должно совпадать с Google Sheets:

```
✗ ОШИБКА: Несоответствие структуры:
  Ожидается: 5 колонок
  В Excel: 7 колонок
```

### Резервное копирование
Перед импортом автоматически создается резервная копия:

```
✓ Создана резервная копия: backup_20260220_143052
```

Копия создается как новый лист в той же таблице.

## 🐛 Troubleshooting

### Ошибка: "Файл creds.json не найден"
- Убедитесь, что файл `creds.json` находится в корне проекта
- Проверьте, что файл не переименован

### Ошибка: "Таблица не найдена или нет доступа"
- Проверьте правильность `SPREADSHEET_ID` в `.env`
- Убедитесь, что таблица расшарена для email из `creds.json`
- Проверьте, что включен Google Sheets API в Google Cloud Console

### Ошибка: "Лист не найден"
- Проверьте название листа (`SHEET_NAME`) в `.env`
- Название должно точно совпадать (включая регистр и пробелы)

### Ошибка: "Переменная окружения не найдена"
- Убедитесь, что файл `.env` создан и заполнен
- Проверьте, что нет опечаток в названиях переменных

### Ошибка: "Несоответствие структуры"
- Убедитесь, что количество колонок в Excel совпадает с Google Sheets
- Проверьте первую строку (заголовки) в обоих файлах

## 📊 Примеры логов

### Успешный экспорт
```
[2026-02-20 14:30:45] [INFO] ============================================================
[2026-02-20 14:30:45] [INFO] РЕЖИМ: ЭКСПОРТ (Google Sheets → Excel)
[2026-02-20 14:30:45] [INFO] ============================================================
[2026-02-20 14:30:45] [INFO] Авторизация в Google Sheets...
[2026-02-20 14:30:46] [INFO] ✓ Авторизация успешна
[2026-02-20 14:30:46] [INFO] Открытие таблицы...
[2026-02-20 14:30:47] [INFO] ✓ Таблица открыта
[2026-02-20 14:30:47] [INFO] Чтение данных из Google Sheets...
[2026-02-20 14:30:48] [INFO] ✓ Прочитано 150 строк
[2026-02-20 14:30:48] [INFO] Создание Excel файла report_20260220_143048.xlsx...
[2026-02-20 14:30:48] [INFO] ✓ Данные записаны
[2026-02-20 14:30:49] [INFO] Применение форматирования к листу Excel...
[2026-02-20 14:30:49] [INFO] ✓ Форматирование успешно применено
[2026-02-20 14:30:49] [INFO] ✓ Файл сохранен: report_20260220_143048.xlsx
[2026-02-20 14:30:49] [INFO] ============================================================
[2026-02-20 14:30:49] [INFO] ✓ ЭКСПОРТ ЗАВЕРШЕН
[2026-02-20 14:30:49] [INFO]   Файл: g:\Code\Google Sheets\excel-gsheets-sync\report_20260220_143048.xlsx
[2026-02-20 14:30:49] [INFO]   Строк: 150
[2026-02-20 14:30:49] [INFO] ============================================================
```

### Успешный импорт
```
[2026-02-20 14:35:12] [INFO] ============================================================
[2026-02-20 14:35:12] [INFO] РЕЖИМ: ИМПОРТ (Excel → Google Sheets)
[2026-02-20 14:35:12] [INFO] ============================================================
[2026-02-20 14:35:12] [INFO] Чтение Excel файла data.xlsx...
[2026-02-20 14:35:12] [INFO] ✓ Прочитано 120 строк
[2026-02-20 14:35:12] [INFO] Авторизация в Google Sheets...
[2026-02-20 14:35:13] [INFO] ✓ Авторизация успешна
[2026-02-20 14:35:13] [INFO] Проверка структуры данных...
[2026-02-20 14:35:14] [INFO] ✓ Структура Excel файла совпадает с Google Sheets
[2026-02-20 14:35:14] [INFO] Создание резервной копии 'backup_20260220_143514'...
[2026-02-20 14:35:15] [INFO] ✓ Резервная копия создана: backup_20260220_143514
[2026-02-20 14:35:15] [INFO] Запись данных в Google Sheets (120 строк)...
[2026-02-20 14:35:17] [INFO] ✓ Данные успешно записаны
[2026-02-20 14:35:17] [INFO] ============================================================
[2026-02-20 14:35:17] [INFO] ✓ ИМПОРТ ЗАВЕРШЕН
[2026-02-20 14:35:17] [INFO]   Файл: g:\Code\Google Sheets\excel-gsheets-sync\data.xlsx
[2026-02-20 14:35:17] [INFO]   Строк: 120
[2026-02-20 14:35:17] [INFO]   Резервная копия: backup_20260220_143514
[2026-02-20 14:35:17] [INFO] ============================================================
```

## 🔒 Безопасность

### Файлы НЕ для Git:
- `.env` — содержит ID таблицы
- `creds.json` — содержит приватные ключи Google
- `*.xlsx` — Excel файлы с данными
- `.venv/` — виртуальное окружение

Все эти файлы уже добавлены в `.gitignore`.

### Рекомендации:
- Никогда не коммитьте `creds.json` в репозиторий
- Не делитесь `creds.json` через мессенджеры или email
- Регулярно обновляйте ключи доступа
- Используйте минимальные права доступа для Service Account

## 📝 Лицензия

Проект создан для образовательных целей.

## 👨‍💻 Автор

Разработано с использованием GitHub Copilot (Claude Sonnet 4.5)

---

**Успешной синхронизации! 🚀**
