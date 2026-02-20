"""
Модуль для загрузки конфигурации из переменных окружения.
"""
import os
import logging
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv

logger = logging.getLogger(__name__)


class Config:
    """Класс для управления конфигурацией приложения."""
    
    def __init__(self, env_file: Optional[str] = None):
        """
        Инициализация конфигурации.
        
        Args:
            env_file: Путь к .env файлу. Если не указан, используется .env в текущей директории.
        """
        # Загружаем переменные из .env файла
        if env_file:
            load_dotenv(env_file)
        else:
            load_dotenv()
        
        # Загружаем обязательные переменные
        self.spreadsheet_id: str = self._get_required_env('SPREADSHEET_ID')
        self.sheet_name: str = self._get_required_env('SHEET_NAME')
        
        # Путь к файлу с учетными данными Google
        self.creds_file: Path = Path('creds.json')
        
        # Проверяем наличие файла с учетными данными
        if not self.creds_file.exists():
            raise FileNotFoundError(
                f"Файл {self.creds_file} не найден. "
                "Убедитесь, что вы создали файл с учетными данными Google Service Account."
            )
        
        logger.info("Конфигурация успешно загружена")
        logger.info(f"- SPREADSHEET_ID: {self.spreadsheet_id}")
        logger.info(f"- SHEET_NAME: {self.sheet_name}")
        logger.info(f"- CREDS_FILE: {self.creds_file}")
    
    def _get_required_env(self, key: str) -> str:
        """
        Получить обязательную переменную окружения.
        
        Args:
            key: Название переменной.
            
        Returns:
            Значение переменной.
            
        Raises:
            ValueError: Если переменная не найдена или пустая.
        """
        value = os.getenv(key)
        if not value:
            raise ValueError(
                f"Переменная окружения {key} не найдена или пустая. "
                f"Убедитесь, что файл .env содержит все необходимые переменные."
            )
        return value
    
    def __repr__(self) -> str:
        """Строковое представление конфигурации."""
        return (
            f"Config(spreadsheet_id='{self.spreadsheet_id}', "
            f"sheet_name='{self.sheet_name}', "
            f"creds_file='{self.creds_file}')"
        )


def get_config(env_file: Optional[str] = None) -> Config:
    """
    Получить конфигурацию приложения.
    
    Args:
        env_file: Путь к .env файлу (опционально).
        
    Returns:
        Объект конфигурации.
    """
    return Config(env_file)
