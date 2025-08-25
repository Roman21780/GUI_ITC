import sys
import os
import pyodbc
from database import AccessDatabase
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def initialize_database():
    """Инициализирует базу данных и создает таблицы"""
    try:
        print("Создание базы данных...")
        db = AccessDatabase()

        # Проверяем соединение
        conn = db.get_connection()
        if conn:
            print("✓ База данных успешно создана и подключена")

        # Создаем таблицы
        print("Создание таблиц...")
        db.create_tables()

        db.close()
        print("Инициализация завершена успешно!")

    except Exception as e:
        print(f"Ошибка инициализации: {e}")


if __name__ == "__main__":
    initialize_database()