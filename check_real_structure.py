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
            'amendments'
        ]

        for table in tables:
            print(f"\n=== {table} ===")
            columns = db.check_table_structure(table)
            print(f"Поля: {columns}")

        db.close()

    except Exception as e:
        print(f"Ошибка: {e}")


if __name__ == "__main__":
    check_real_structures()