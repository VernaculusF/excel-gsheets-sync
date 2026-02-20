"""
Модуль для работы с Google Sheets API.
"""
import gspread
from google.auth.exceptions import GoogleAuthError
from gspread.exceptions import SpreadsheetNotFound, WorksheetNotFound, APIError
from gspread import Client, Spreadsheet, Worksheet
from typing import List, Any, Optional
from datetime import datetime
import logging
from pathlib import Path

logger = logging.getLogger(__name__)


class SheetsClient:
    """Класс для работы с Google Sheets."""
    
    def __init__(self, creds_file: Path, spreadsheet_id: str, sheet_name: str):
        """
        Инициализация клиента Google Sheets.
        
        Args:
            creds_file: Путь к файлу с учетными данными сервисного аккаунта.
            spreadsheet_id: ID таблицы Google Sheets.
            sheet_name: Название листа в таблице.
            
        Raises:
            GoogleAuthError: Ошибка авторизации.
            FileNotFoundError: Файл с учетными данными не найден.
        """
        self.creds_file = creds_file
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.client: Optional[Client] = None
        self.spreadsheet: Optional[Spreadsheet] = None
        self.worksheet: Optional[Worksheet] = None
        
        self._authenticate()
    
    def _authenticate(self) -> None:
        """
        Авторизация через сервисный аккаунт Google.
        
        Raises:
            GoogleAuthError: Ошибка авторизации.
            FileNotFoundError: Файл с учетными данными не найден.
        """
        try:
            logger.info("Авторизация в Google Sheets...")
            self.client = gspread.service_account(filename=str(self.creds_file))
            logger.info("✓ Авторизация успешна")
        except FileNotFoundError:
            logger.error(f"Файл {self.creds_file} не найден")
            raise
        except GoogleAuthError as e:
            logger.error(f"Ошибка авторизации Google: {e}")
            raise
        except Exception as e:
            logger.error(f"Неожиданная ошибка при авторизации: {e}")
            raise
    
    def _get_worksheet(self) -> None:
        """
        Получить доступ к листу таблицы.
        
        Raises:
            SpreadsheetNotFound: Таблица не найдена или нет доступа.
            WorksheetNotFound: Лист не найден.
        """
        assert self.client is not None, "Клиент не инициализирован"
        
        try:
            logger.info(f"Открытие таблицы {self.spreadsheet_id}...")
            self.spreadsheet = self.client.open_by_key(self.spreadsheet_id)
            logger.info("✓ Таблица открыта")
            
            logger.info(f"Открытие листа '{self.sheet_name}'...")
            self.worksheet = self.spreadsheet.worksheet(self.sheet_name)
            logger.info("✓ Лист открыт")
        except SpreadsheetNotFound:
            logger.error(
                f"Таблица {self.spreadsheet_id} не найдена или нет доступа. "
                f"Убедитесь, что таблица расшарена для сервисного аккаунта."
            )
            raise
        except WorksheetNotFound:
            logger.error(
                f"Лист '{self.sheet_name}' не найден в таблице. "
                f"Проверьте название листа в .env."
            )
            raise
        except APIError as e:
            logger.error(f"Ошибка API Google Sheets: {e}")
            raise
    
    def get_all_values(self) -> List[List[Any]]:
        """
        Получить все значения из листа.
        
        Returns:
            Список списков со значениями ячеек. Пустые ячейки представлены как ''.
            
        Raises:
            SpreadsheetNotFound: Таблица не найдена.
            WorksheetNotFound: Лист не найден.
        """
        self._get_worksheet()
        assert self.worksheet is not None, "Лист не инициализирован"
        
        logger.info("Чтение данных из Google Sheets...")
        try:
            data = self.worksheet.get_all_values()
            logger.info(f"✓ Прочитано {len(data)} строк")
            return data
        except APIError as e:
            logger.error(f"Ошибка при чтении данных: {e}")
            raise
    
    def update_values(self, data: List[List[Any]]) -> None:
        """
        Обновить все значения на листе.
        
        Args:
            data: Данные для записи (список списков).
            
        Raises:
            SpreadsheetNotFound: Таблица не найдена.
            WorksheetNotFound: Лист не найден.
        """
        self._get_worksheet()
        assert self.worksheet is not None, "Лист не инициализирован"
        
        logger.info(f"Запись данных в Google Sheets ({len(data)} строк)...")
        try:
            # Очищаем лист перед записью
            self.worksheet.clear()
            
            # Записываем новые данные
            if data:
                self.worksheet.update(range_name='A1', values=data)
            
            logger.info("✓ Данные успешно записаны")
        except APIError as e:
            logger.error(f"Ошибка при записи данных: {e}")
            raise
    
    def create_backup(self) -> str:
        """
        Создать резервную копию текущего листа.
        
        Returns:
            Название созданного листа с резервной копией.
            
        Raises:
            SpreadsheetNotFound: Таблица не найдена.
            WorksheetNotFound: Лист не найден.
        """
        self._get_worksheet()
        assert self.worksheet is not None, "Лист не инициализирован"
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"backup_{timestamp}"
        
        logger.info(f"Создание резервной копии '{backup_name}'...")
        try:
            # Дублируем лист
            self.worksheet.duplicate(new_sheet_name=backup_name)
            logger.info(f"✓ Резервная копия создана: {backup_name}")
            return backup_name
        except APIError as e:
            logger.error(f"Ошибка при создании резервной копии: {e}")
            raise
    
    def get_headers(self) -> List[str]:
        """
        Получить заголовки (первую строку) листа.
        
        Returns:
            Список заголовков.
            
        Raises:
            SpreadsheetNotFound: Таблица не найдена.
            WorksheetNotFound: Лист не найден.
        """
        self._get_worksheet()
        assert self.worksheet is not None, "Лист не инициализирован"
        
        try:
            headers = self.worksheet.row_values(1)
            logger.debug(f"Заголовки: {headers}")
            return headers
        except APIError as e:
            logger.error(f"Ошибка при чтении заголовков: {e}")
            raise
    
    def __repr__(self) -> str:
        """Строковое представление клиента."""
        return (
            f"SheetsClient(spreadsheet_id='{self.spreadsheet_id}', "
            f"sheet_name='{self.sheet_name}')"
        )
