import pyodbc
import pandas as pd
import os

# Путь к базе данных Access
db_path = r'C:\Work\GUI_ITC\research_data.accdb'

# Подключение к базе данных
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    f'DBQ={db_path};'
)
conn = pyodbc.connect(conn_str)

# Получение списка таблиц
cursor = conn.cursor()
tables = [table.table_name for table in cursor.tables(tableType='TABLE')]

# Создание папки для CSV-файлов
output_dir = r'C:\Work\GUI_ITC\csv_exports'
os.makedirs(output_dir, exist_ok=True)

# Экспорт каждой таблицы в CSV
for table in tables:
    query = f"SELECT * FROM [{table}]"
    df = pd.read_sql(query, conn)
    output_path = os.path.join(output_dir, f"{table}.csv")
    df.to_csv(output_path, index=False, encoding='utf-8')
    print(f"Таблица {table} экспортирована в {output_path}")

# Закрытие соединения
conn.close()