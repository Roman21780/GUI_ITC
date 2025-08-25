# database.py
import pyodbc
import pandas as pd
import logging
import datetime
import os


class AccessDatabase:
    def __init__(self, db_path=None):
        """
        Инициализирует соединение с базой данных.
        :param db_path: Путь к файлу базы данных.
        """
        if db_path is None:
            db_path = r'C:\Work\GUI_ITC\research_data.accdb'
        self.db_path = db_path

        # Создаем базу данных, если она не существует
        if not os.path.exists(self.db_path):
            self.create_database()

        self.conn = None
        self.get_connection()  # Устанавливаем соединение при создании экземпляра

    def get_connection(self):
        """
        Устанавливает соединение с базой данных.
        :return: Соединение с базой данных.
        """
        if self.conn is None:
            try:
                conn_str = (
                    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                    f'DBQ={self.db_path};'
                )
                self.conn = pyodbc.connect(conn_str)
                logging.info(f"Успешное подключение к базе данных: {self.db_path}")
            except Exception as e:
                logging.error(f"Ошибка подключения к базе данных: {str(e)}")
                raise
        return self.conn

    def create_database(self):
        """
        Создает новую базу данных Access, если она не существует.
        """
        try:
            import win32com.client

            # Создаем экземпляр Access
            access_app = win32com.client.Dispatch("Access.Application")
            access_app.Visible = False  # Работаем в фоновом режиме

            # Создаем новую базу данных
            access_app.NewCurrentDatabase(self.db_path)

            # Закрываем базу и приложение
            access_app.CloseCurrentDatabase()
            access_app.Quit()

            # Освобождаем COM объекты
            del access_app

            logging.info(f"База данных успешно создана через MS Access: {self.db_path}")

        except Exception as e:
            logging.error(f"Ошибка создания базы данных через Access: {str(e)}")
            raise

    def create_tables(self):
        """
        Создает таблицы в базе данных с упрощенным синтаксисом.
        """
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Создаем таблицы по одной с простым синтаксисом
            tables_sql = [
                # Таблица main_data - упрощенная версия
                '''
                CREATE TABLE main_data
                (
                    id            COUNTER PRIMARY KEY,
                    company       TEXT,
                    field         TEXT,
                    well          TEXT,
                    date_research DATE,
                    created_date  DATETIME
                )
                ''',

                # Таблица research_params
                '''
                CREATE TABLE research_params
                (
                    id           COUNTER PRIMARY KEY,
                    main_data_id INTEGER,
                    section      INTEGER,
                    param_name   TEXT,
                    param_value  TEXT,
                    created_date DATETIME
                )
                ''',

                # Таблица corrections
                '''
                CREATE TABLE corrections
                (
                    id              COUNTER PRIMARY KEY,
                    main_data_id    INTEGER,
                    correction_type TEXT,
                    [value] FLOAT,
                    created_date    DATETIME
                )
                '''
            ]

            table_names = ['main_data', 'research_params', 'corrections']

            for i, (table_name, create_sql) in enumerate(zip(table_names, tables_sql)):
                try:
                    # Проверяем, существует ли таблица
                    cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
                    print(f"✓ Таблица {table_name} уже существует")
                except pyodbc.Error:
                    # Таблица не существует - создаем
                    try:
                        cursor.execute(create_sql)
                        print(f"✓ Таблица {table_name} создана успешно")
                    except pyodbc.Error as e:
                        print(f"✗ Ошибка создания таблицы {table_name}: {e}")
                        # Покажем конкретную ошибку
                        if 'research_params' in str(e):
                            print("Таблица research_params уже существует - это нормально")

            conn.commit()
            cursor.close()

        except Exception as e:
            print(f"Ошибка создания таблиц: {e}")
            raise

    def close(self):
        """
        Закрывает соединение с базой данных.
        """
        if self.conn:
            self.conn.close()
            self.conn = None
            logging.info("Соединение с базой данных закрыто")

    # Дополнительные методы для проверки
    def table_exists(self, table_name):
        """Проверяет существование таблицы"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            cursor.close()
            return True
        except pyodbc.Error:
            return False

    def get_table_info(self, table_name):
        """Возвращает информацию о таблице"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Получаем информацию о колонках
            cursor.execute(f"SELECT TOP 1 * FROM {table_name}")
            columns = [column[0] for column in cursor.description]

            # Получаем количество записей
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            count = cursor.fetchone()[0]

            cursor.close()
            return {'columns': columns, 'row_count': count}

        except pyodbc.Error as e:
            return {'error': str(e)}