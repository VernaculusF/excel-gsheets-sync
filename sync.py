#!/usr/bin/env python
"""
Основной скрипт для синхронизации между Google Sheets и Excel.

Режимы работы:
1. --export: Экспорт данных из Google Sheets в Excel
2. --import --file "путь": Импорт данных из Excel в Google Sheets
"""
import argparse
import logging
import sys
from pathlib import Path
from datetime import datetime
from config import get_config
from sheets_client import SheetsClient
from excel_handler import ExcelHandler


# Настройка логирования
def setup_logging():
    """Настройка логирования в консоль."""
    logging.basicConfig(
        level=logging.INFO,
        format='[%(asctime)s] [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


def export_to_excel(config) -> None:
    """
    Экспорт данных из Google Sheets в Excel.
    
    Args:
        config: Объект конфигурации.
    """
    logger = logging.getLogger(__name__)
    logger.info("=" * 60)
    logger.info("РЕЖИМ: ЭКСПОРТ (Google Sheets → Excel)")
    logger.info("=" * 60)
    
    try:
        # Инициализация клиента Google Sheets
        sheets_client = SheetsClient(
            creds_file=config.creds_file,
            spreadsheet_id=config.spreadsheet_id,
            sheet_name=config.sheet_name
        )
        
        # Получение данных из Google Sheets
        data = sheets_client.get_all_values()
        
        if not data:
            logger.warning("Таблица Google Sheets пустая. Нечего экспортировать.")
            sys.exit(0)
        
        # Создание имени выходного файла
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = Path(f"report_{timestamp}.xlsx")
        
        # Создание Excel файла с форматированием
        excel_handler = ExcelHandler()
        excel_handler.create_excel(data, output_file)
        
        logger.info("=" * 60)
        logger.info("✓ ЭКСПОРТ ЗАВЕРШЕН")
        logger.info(f"  Файл: {output_file.absolute()}")
        logger.info(f"  Строк: {len(data)}")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"✗ ОШИБКА ПРИ ЭКСПОРТЕ: {e}")
        sys.exit(1)


def import_from_excel(config, excel_file_path: str) -> None:
    """
    Импорт данных из Excel в Google Sheets.
    
    Args:
        config: Объект конфигурации.
        excel_file_path: Путь к Excel файлу.
    """
    logger = logging.getLogger(__name__)
    logger.info("=" * 60)
    logger.info("РЕЖИМ: ИМПОРТ (Excel → Google Sheets)")
    logger.info("=" * 60)
    
    try:
        # Проверка существования файла
        excel_path = Path(excel_file_path)
        if not excel_path.exists():
            logger.error(f"Файл {excel_path} не найден")
            sys.exit(1)
        
        # Инициализация обработчика Excel
        excel_handler = ExcelHandler()
        
        # Чтение данных из Excel
        excel_data = excel_handler.read_excel(excel_path)
        
        if not excel_data:
            logger.error("Excel файл пустой")
            sys.exit(1)
        
        # Инициализация клиента Google Sheets
        sheets_client = SheetsClient(
            creds_file=config.creds_file,
            spreadsheet_id=config.spreadsheet_id,
            sheet_name=config.sheet_name
        )
        
        # Получение заголовков из Google Sheets для проверки структуры
        logger.info("Проверка структуры данных...")
        expected_headers = sheets_client.get_headers()
        
        # Проверка совпадения структуры
        is_valid, error_msg = excel_handler.validate_structure(excel_data, expected_headers)
        if not is_valid:
            logger.error(error_msg)
            sys.exit(1)
        
        # Создание резервной копии перед импортом
        backup_name = sheets_client.create_backup()
        logger.info(f"✓ Создана резервная копия: {backup_name}")
        
        # Обновление данных в Google Sheets
        sheets_client.update_values(excel_data)
        
        logger.info("=" * 60)
        logger.info("✓ ИМПОРТ ЗАВЕРШЕН")
        logger.info(f"  Файл: {excel_path.absolute()}")
        logger.info(f"  Строк: {len(excel_data)}")
        logger.info(f"  Резервная копия: {backup_name}")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"✗ ОШИБКА ПРИ ИМПОРТЕ: {e}")
        sys.exit(1)


def main():
    """Точка входа в приложение."""
    setup_logging()
    logger = logging.getLogger(__name__)
    
    # Создание парсера аргументов
    parser = argparse.ArgumentParser(
        description='Синхронизация данных между Google Sheets и Excel',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python sync.py --export
  python sync.py --import --file "report.xlsx"
  python sync.py --import --file "C:\\Users\\user\\Documents\\data.xlsx"
        """
    )
    
    # Определение режима работы
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        '--export',
        action='store_true',
        help='Экспорт данных из Google Sheets в Excel'
    )
    group.add_argument(
        '--import',
        action='store_true',
        dest='import_mode',
        help='Импорт данных из Excel в Google Sheets'
    )
    
    # Аргумент для пути к файлу (обязателен при --import)
    parser.add_argument(
        '--file',
        type=str,
        help='Путь к Excel файлу (обязательно при --import)'
    )
    
    # Парсинг аргументов
    args = parser.parse_args()
    
    # Проверка аргументов
    if args.import_mode and not args.file:
        parser.error("--import требует указания --file")
    
    try:
        # Загрузка конфигурации
        config = get_config()
        
        # Выполнение операции в зависимости от режима
        if args.export:
            export_to_excel(config)
        elif args.import_mode:
            import_from_excel(config, args.file)
            
    except Exception as e:
        logger.error(f"✗ КРИТИЧЕСКАЯ ОШИБКА: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
