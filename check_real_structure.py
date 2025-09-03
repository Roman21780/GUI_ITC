import sys
import os

sys.path.append(r'C:\Work\GUI_ITC')

from db_access import AccessDatabase


def check_real_structures():
    """Проверяет реальную структуру всех таблиц"""
    try:
        db = AccessDatabase()

        tables = [
            'success',
            'researchClass',
            'pressureLastPoint',
            'estimatedTime',
            'density',
            'calculatedParameters',
            'calculatedPressure',
            'InputData',
            'ModelVNK',
            'PressureVNK',
            'amendments',
            'amendments2',
            'amendments3',
            'amendments4',
            'ModelKSD',
            'calculate',
            'prevData',
            'dampingTable'
        ]

        for table in tables:
            print(f"\n=== {table} ===")
            columns = db.check_table_structure(table)
            print(f"Поля: {columns}")

        db.close()

    except Exception as e:
        print(f"Ошибка: {e}")


def check_column_types():
    """Проверяет точные типы данных столбцов в таблице"""
    try:
        db = AccessDatabase()

        tables = [
            'success',
            'researchClass',
            'pressureLastPoint',
            'estimatedTime',
            'density',
            'calculatedParameters',
            'calculatedPressure',
            'InputData',
            'ModelVNK',
            'PressureVNK',
            'amendments',
            'ModelKSD',
            'calculate',
            'prevData'
        ]

        for table in tables:
            print(f"\n=== {table} ===")
            columns = db.check_column_types(table)
            print(f"Поля: {columns}")

        db.close()

    except Exception as e:
        print(f"Ошибка при проверке типов данных: {e}")

def get_odbc_type_name(self, type_code):
    """Преобразует код типа ODBC в читаемое имя"""
    odbc_types = {
        1: 'SQL_CHAR',
        2: 'SQL_NUMERIC',
        3: 'SQL_DECIMAL',
        4: 'SQL_INTEGER',
        5: 'SQL_SMALLINT',
        6: 'SQL_FLOAT',
        7: 'SQL_REAL',
        8: 'SQL_DOUBLE',
        9: 'SQL_DATE',
        10: 'SQL_TIME',
        11: 'SQL_TIMESTAMP',
        12: 'SQL_VARCHAR',
        # ... добавьте другие коды по необходимости
    }
    return odbc_types.get(type_code, f'UNKNOWN_TYPE_{type_code}')


if __name__ == "__main__":
    check_real_structures()
    # check_column_types()
