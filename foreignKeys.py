import sys
import os

sys.path.append(r'C:\Work\GUI_ITC')

from db_access import AccessDatabase


def add_foreign_keys():
    """Добавляет поля связи во все таблицы"""
    try:
        db = AccessDatabase()
        db.add_foreign_keys_to_tables()
        db.close()
        print("Готово! Поля связи добавлены.")

    except Exception as e:
        print(f"Ошибка: {e}")


if __name__ == "__main__":
    add_foreign_keys()