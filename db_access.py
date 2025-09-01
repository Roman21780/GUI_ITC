import pyodbc
import os
from utils import logger, logging
import pandas as pd
from datetime import datetime
from tkinter import messagebox
import math

class AccessDatabase:
    def __init__(self, db_path='research_data.accdb'):
        self.db_path = os.path.abspath(db_path)
        self.connection_string = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                r'DBQ=' + self.db_path + ';'
        )
        self.conn = None
        self.current_main_data_id = None

    def get_connection(self):
        """Устанавливает и возвращает соединение с базой данных."""
        if self.conn is None:
            try:
                self.conn = pyodbc.connect(self.connection_string)
                logger.info(f"Успешное подключение к базе данных: {self.db_path}")
            except pyodbc.Error as e:
                logger.error(f"Ошибка подключения к базе данных: {e}")
                raise
        return self.conn

    def close(self):
        """Закрывает соединение с базой данных."""
        if self.conn:
            self.conn.close()
            self.conn = None
            logger.info("Соединение с базой данных закрыто.")

    def get_last_record(self):
        """Возвращает последнюю запись из InputData"""
        try:
            conn = self.get_connection()
            query = "SELECT TOP 1 * FROM InputData ORDER BY ID DESC"
            with conn.cursor() as cursor:
                cursor.execute(query)
                columns = [column[0] for column in cursor.description]
                data = cursor.fetchall()
                df = pd.DataFrame(data, columns=columns)
            return df
        except Exception as e:
            logging.error(f"Ошибка получения последней записи: {str(e)}")
            return pd.DataFrame()

    def get_model_vnk_ordered(self, input_data_id):
        """Получает данные ModelVNK в исходном порядке"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute("""
                           SELECT empty, DT, deltaP, DP, PressureVnkModel
                           FROM ModelVNK
                           WHERE InputData_ID = ?
                           ORDER BY sort_order
                           """, (input_data_id,))

            results = cursor.fetchall()
            return results

        except Exception as e:
            print(f"Ошибка получения данных: {e}")
            return []

    def clear_data(self):
        """Очищает все данные из таблиц базы данных."""
        conn = self.get_connection()
        cursor = conn.cursor()
        tables = [
            'dampingTable', 'calculatedPressure', 'calculatedParameters',
            'amendments', 'Parameters', 'InputData',
            'researchClass', 'success', 'density', 'estimatedTime',
            'ModelVNK', 'PressureVNK', 'pressureLastPoint', 'TextParameters', 'ModelKSD'
        ]
        try:
            for table in tables:
                try:
                    cursor.execute(f"DELETE FROM {table}")
                except pyodbc.Error:
                    logger.warning(f"Таблица {table} не найдена, пропускаем")
            conn.commit()
            logger.info("Все данные из таблиц успешно удалены.")
            return True
        except pyodbc.Error as e:
            conn.rollback()
            logger.error(f"Ошибка при очистке таблиц: {e}")
            return False
        finally:
            cursor.close()

    def insert_main_data(self, data_dict):
        """Вставляет основные данные в таблицу InputData."""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Отладочный вывод
            print("Полученные данные для вставки:")
            for key, value in data_dict.items():
                print(f"  {key}: {value}")

            sql = """
                  INSERT INTO InputData (Company, Localoredenie, Skvazhina, VNK, Data_issledovaniya, Plast, \
                                         Interval_perforacii, Tip_pribora, Glubina_ustanovki_pribora, Interpretator, \
                                         Data_interpretacii, Vremya_issledovaniya, Obvodnennost, Nalicie_paktera, \
                                         Data_GRP, Vid_issledovaniya, created_date) \
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) \
                  """

            # Безопасное получение значений с значениями по умолчанию
            params = (
                data_dict.get('Company', ''),
                data_dict.get('Localoredenie', ''),
                data_dict.get('Skvazhina', ''),
                data_dict.get('VNK', ''),
                data_dict.get('Data_issledovaniya'),
                data_dict.get('Plast', ''),
                data_dict.get('Interval_perforacii', ''),
                data_dict.get('Tip_pribora', ''),
                data_dict.get('Glubina_ustanovki_pribora', ''),
                data_dict.get('Interpretator', ''),
                data_dict.get('Data_interpretacii'),
                data_dict.get('Vremya_issledovaniya'),
                data_dict.get('Obvodnennost', ''),
                data_dict.get('Nalicie_paktera', ''),
                data_dict.get('Data_GRP'),
                data_dict.get('Vid_issledovaniya', ''),
                datetime.now()
            )

            print("Параметры для SQL:")
            for i, param in enumerate(params, 1):
                print(f"  Параметр {i}: {param}")

            cursor.execute(sql, params)
            conn.commit()

            # Получаем ID вставленной записи
            cursor.execute("SELECT @@IDENTITY")
            last_id = cursor.fetchone()[0]

            logging.info(f"Данные успешно добавлены в InputData с ID: {last_id}")
            return last_id

        except Exception as e:
            logging.error(f"Ошибка при вставке в InputData: {e}")
            raise

    def insert_all_calculated_parameters_from_clipboard(self, input_data_id, clipboard_data):
        """Вставляет все параметры (числовые и текстовые) в соответствующие таблицы"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Разбираем данные из буфера обмена
            rows = [r.split('\t') for r in clipboard_data.split('\n') if r.strip()]

            numeric_count = 0
            text_count = 0
            updated_count = 0

            for row in rows:
                if len(row) >= 2:
                    param_name = row[0].strip()
                    param_value = row[1].strip()
                    unit = row[2].strip() if len(row) >= 3 else ""

                    if param_name and param_value:
                        # Пробуем вставить как число
                        try:
                            numeric_value = float(param_value.replace(',', '.'))

                            # ПРОВЕРЯЕМ СУЩЕСТВОВАНИЕ ЗАПИСИ
                            check_sql = "SELECT COUNT(*) FROM calculatedParameters WHERE InputData_ID = ? AND calcParam = ?"
                            cursor.execute(check_sql, (input_data_id, param_name))
                            exists = cursor.fetchone()[0] > 0

                            if exists:
                                # ОБНОВЛЯЕМ существующую запись
                                update_sql = "UPDATE calculatedParameters SET Val = ?, unit = ? WHERE InputData_ID = ? AND calcParam = ?"
                                cursor.execute(update_sql, (numeric_value, unit, input_data_id, param_name))
                                updated_count += 1
                            else:
                                # ВСТАВЛЯЕМ новую запись
                                sql = "INSERT INTO calculatedParameters (InputData_ID, calcParam, Val, unit) VALUES (?, ?, ?, ?)"
                                cursor.execute(sql, (input_data_id, param_name, numeric_value, unit))
                                numeric_count += 1

                        except ValueError:
                            # Если не число, вставляем как текст в другую таблицу
                            # Сначала проверим, существует ли таблица TextParameters
                            try:
                                check_sql = "SELECT COUNT(*) FROM TextParameters WHERE InputData_ID = ? AND ParamName = ?"
                                cursor.execute(check_sql, (input_data_id, param_name))
                                exists = cursor.fetchone()[0] > 0

                                if exists:
                                    update_sql = "UPDATE TextParameters SET ParamValue = ?, Unit = ? WHERE InputData_ID = ? AND ParamName = ?"
                                    cursor.execute(update_sql, (param_value, unit, input_data_id, param_name))
                                else:
                                    sql = "INSERT INTO TextParameters (InputData_ID, ParamName, ParamValue, Unit) VALUES (?, ?, ?, ?)"
                                    cursor.execute(sql, (input_data_id, param_name, param_value, unit))
                                text_count += 1
                            except:
                                # Если таблицы TextParameters нет, просто пропускаем текстовые параметры
                                print(f"Пропущен текстовый параметр: {param_name} = {param_value}")
                                continue

            conn.commit()
            print(
                f"Вставлено новых: {numeric_count} числовых, {text_count} текстовых. Обновлено: {updated_count} записей")

        except Exception as e:
            logging.error(f"Ошибка вставки параметров: {str(e)}")
            raise

    def insert_model_vnk(self, input_data_id, data_list):
        """Вставляет данные в ModelVNK с сохранением порядка"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            inserted_count = 0

            for i, data in enumerate(data_list):
                sql = "INSERT INTO ModelVNK (empty, Dat, PressureVnkModel, InputData_ID, sort_order) VALUES (?, ?, ?, ?, ?)"
                cursor.execute(sql, (
                    data.get('empty', ''),
                    data.get('Dat'),
                    data.get('PressureVnkModel'),
                    input_data_id,
                    i  # Порядковый номер для сортировки
                ))
                inserted_count += 1

            conn.commit()
            print(f"Успешно вставлено записей: {inserted_count}")
            logging.info(f"Данные ModelVNK сохранены для ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки ModelVNK: {str(e)}")
            raise

    def convert_scientific_notation(self, value):
        """Корректно преобразует научную нотацию в float"""
        if not value:
            return 0.0

        try:
            # Убираем пробелы и заменяем запятые на точки
            clean_value = str(value).strip().replace(' ', '').replace(',', '.')

            # Обрабатываем научную нотацию
            if 'E' in clean_value.upper():
                # Для научной нотации: 4.81E-04 -> 0.000481
                return float(clean_value)
            else:
                # Для обычных чисел
                return float(clean_value)

        except (ValueError, TypeError) as e:
            print(f"Ошибка преобразования '{value}': {e}")
            return 0.0

    def calculate_pressure_vnk_model(self, input_data_id, research_type):
        """Рассчитывает PressureVnkModel на основе данных исследования"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Получаем все записи для этого исследования
            cursor.execute("""
                           SELECT ID, DT, deltaP, DP, PressureVnkModel
                           FROM ModelVNK
                           WHERE InputData_ID = ?
                           ORDER BY ID
                           """, (input_data_id,))

            records = cursor.fetchall()

            if not records:
                print("Нет данных для расчета")
                return

            # Получаем PzabVnk из calculatedParameters
            cursor.execute("""
                           SELECT Val
                           FROM calculatedParameters
                           WHERE InputData_ID = ?
                             AND calcParam = 'PzabVnk'
                           """, (input_data_id,))

            pzab_vnk_result = cursor.fetchone()
            pzab_vnk = pzab_vnk_result[0] if pzab_vnk_result else None

            if pzab_vnk is None:
                print("Не найдено PzabVnk для расчета")
                return

            print(f"PzabVnk для расчета: {pzab_vnk}")
            print(f"Тип исследования: {research_type}")
            print(f"Найдено записей: {len(records)}")

            # Расчет значений PressureVnkModel
            for i, record in enumerate(records):
                record_id, dt, delta_p, dp, current_value = record

                if "КПД" in research_type.upper():
                    # Логика для КПД исследований
                    if i == 0:
                        new_value = pzab_vnk
                    elif i == 1:
                        new_value = pzab_vnk - delta_p
                    else:
                        prev_record = records[i - 1]
                        new_value = prev_record[4] - delta_p + records[i - 1][
                            2]  # prev PressureVnkModel - deltaP + prev deltaP
                else:
                    # Логика для других исследований
                    if i == 0:
                        new_value = pzab_vnk
                    elif i == 1:
                        new_value = pzab_vnk + delta_p
                    else:
                        prev_record = records[i - 1]
                        new_value = prev_record[4] + delta_p - records[i - 1][
                            2]  # prev PressureVnkModel + deltaP - prev deltaP

                # Обновляем запись
                cursor.execute("""
                               UPDATE ModelVNK
                               SET PressureVnkModel = ?
                               WHERE ID = ?
                               """, (new_value, record_id))

                print(f"Запись {i}: DT={dt}, deltaP={delta_p}, PressureVnkModel={new_value}")

            conn.commit()
            print("Расчет PressureVnkModel завершен")

        except Exception as e:
            print(f"Ошибка расчета PressureVnkModel: {e}")
            raise

    def get_research_type(self, input_data_id):
        """Получает тип исследования по ID записи"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute("SELECT Vid_issledovaniya FROM InputData WHERE ID = ?", (input_data_id,))
            result = cursor.fetchone()

            return result[0] if result else ""

        except Exception as e:
            print(f"Ошибка получения типа исследования: {e}")
            return ""

    def has_model_vnk_data(self, input_data_id):
        """Проверяет есть ли данные в ModelVNK"""
        try:
            conn = self.get_connection()
            query = "SELECT COUNT(*) as count FROM ModelVNK WHERE InputData_ID = ?"
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df.iloc[0]['count'] > 0
        except Exception as e:
            logging.error(f"Ошибка проверки ModelVNK: {str(e)}")
            return False

    def get_last_model_vnk_pressure(self, input_data_id):
        """Получает последнее значение PressureVnkModel из ModelVNK"""
        try:
            conn = self.get_connection()
            query = """
                    SELECT TOP 1 PressureVnkModel \
                    FROM ModelVNK
                    WHERE InputData_ID = ? \
                      AND PressureVnkModel IS NOT NULL
                    ORDER BY ID DESC \
                    """
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df.iloc[0]['PressureVnkModel'] if not df.empty else None
        except Exception as e:
            logging.error(f"Ошибка получения давления ModelVNK: {str(e)}")
            return None

    def insert_pressure_vnk(self, input_data_id, data_list):
        """Вставляет данные в PressureVNK пакетно"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Пакетная вставка
            batch_size = 5000  # Вставляем по 5000 записей за раз
            for i in range(0, len(data_list), batch_size):
                batch = data_list[i:i + batch_size]

                # Создаем пакетный запрос
                values = []
                for data in batch:
                    values.append((data.get('Dat'), data.get('PressureVnk')))

                # Пакетная вставка
                cursor.executemany(
                    "INSERT INTO PressureVNK (Dat, PressureVnk) VALUES (?, ?)",
                    values
                )
                print(f"Вставлено {len(batch)} записей...")

                # Коммит после каждого пакета
                conn.commit()

            logging.info(f"Данные PressureVNK сохранены для записи ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки PressureVNK: {str(e)}")
            raise


    def get_last_pressure_vnk(self, input_data_id):
        """Получает последнее значение PressureVnk из PressureVNK"""
        try:
            conn = self.get_connection()
            query = """
                    SELECT TOP 1 PressureVnk \
                    FROM PressureVNK
                    WHERE InputData_ID = ? \
                      AND PressureVnk IS NOT NULL
                    ORDER BY ID DESC \
                    """
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df.iloc[0]['PressureVnk'] if not df.empty else None
        except Exception as e:
            logging.error(f"Ошибка получения давления PressureVNK: {str(e)}")
            return None

    def insert_model_ksd(self, input_data_id, data_list):
        """Вставляет данные в ModelKSD с сохранением порядка"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            inserted_count = 0

            for i, data in enumerate(data_list):
                try:
                    sql = """INSERT INTO ModelKSD (empty, Dat, PressureVnkModel, InputData_ID, sort_order)
                             VALUES (?, ?, ?, ?, ?)"""
                    cursor.execute(sql, (
                        data.get('empty', ''),
                        data.get('Dat'),
                        data.get('PressureVnkModel'),
                        input_data_id,
                        i  # Порядковый номер
                    ))
                    inserted_count += 1

                except Exception as insert_error:
                    print(f"Ошибка вставки записи {i}: {insert_error}")
                    continue

            conn.commit()
            print(f"Успешно вставлено записей: {inserted_count}")
            logging.info(f"Данные ModelKSD сохранены для ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки ModelKSD: {str(e)}")
            raise

    def has_model_ksd_data(self, input_data_id):
        """Проверяет есть ли данные в ModelKSD"""
        try:
            conn = self.get_connection()
            query = "SELECT COUNT(*) as count FROM ModelKSD WHERE InputData_ID = ?"
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df.iloc[0]['count'] > 0
        except Exception as e:
            logging.error(f"Ошибка проверки ModelKSD: {str(e)}")
            return False

    def get_last_model_ksd_pressure(self, input_data_id):
        """Получает последнее значение PressureVnkModel из ModelKSD"""
        try:
            conn = self.get_connection()
            query = """
                    SELECT TOP 1 PressureVnkModel \
                    FROM ModelKSD
                    WHERE InputData_ID = ? \
                      AND PressureVnkModel IS NOT NULL
                    ORDER BY ID DESC \
                    """
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df.iloc[0]['PressureVnkModel'] if not df.empty else None
        except Exception as e:
            logging.error(f"Ошибка получения давления ModelKSD: {str(e)}")
            return None

    def insert_research_class(self, input_data_id, klass):
        """Вставляет класс исследования"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            sql = "INSERT INTO researchClass (InputData_ID, Klass_issledovaniya) VALUES (?, ?)"
            cursor.execute(sql, (input_data_id, klass))
            conn.commit()

            logging.info(f"Класс исследования сохранен для записи ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки класса исследования: {str(e)}")
            raise

    def insert_pressure_last_point(self, input_data_id, pressure):
        """Вставляет давление на последнюю точку"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            sql = "INSERT INTO pressureLastPoint (InputData_ID, PressureLastPoint) VALUES (?, ?)"
            cursor.execute(sql, (input_data_id, pressure))
            conn.commit()

            logging.info(f"Давление сохранено для записи ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки давления: {str(e)}")
            raise

    def insert_success(self, input_data_id, uspeshnost):
        """Вставляет данные об успешности"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Вставляем только те поля, которые есть в таблице
            sql = "INSERT INTO success (InputData_ID, Uspeshnost) VALUES (?, ?)"
            cursor.execute(sql, (input_data_id, uspeshnost))
            conn.commit()

            logging.info(f"Успешность сохранена для записи ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки успешности: {str(e)}")
            raise

    def insert_density(self, input_data_id, density_zab, density_pl):
        """Вставляет данные о плотности"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            sql = "INSERT INTO density (InputData_ID, density_zab, density_pl) VALUES (?, ?, ?)"
            cursor.execute(sql, (input_data_id, density_zab, density_pl))
            conn.commit()

            logging.info(f"Плотность сохранена для записи ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки плотности: {str(e)}")
            raise

    def insert_estimated_time(self, input_data_id, durat):
        """Вставляет расчетное время"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            sql = "INSERT INTO estimatedTime (InputData_ID, Durat) VALUES (?, ?)"
            cursor.execute(sql, (input_data_id, durat))
            conn.commit()

            logging.info(f"Расчетное время сохранено для записи ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки расчетного времени: {str(e)}")
            raise

    def insert_amendments(self, input_data_id, amendments_dict):
        """Вставляет поправки в таблицу amendments"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Проверяем какие поля действительно есть в таблице
            existing_fields = self.get_table_columns('amendments')

            # Вставляем только существующие поля
            for field_name, value in amendments_dict.items():
                if field_name in existing_fields and value is not None:
                    # Проверяем, есть ли уже запись для этого InputData_ID
                    check_sql = f"SELECT COUNT(*) FROM amendments WHERE InputData_ID = ? AND {field_name} IS NOT NULL"
                    cursor.execute(check_sql, (input_data_id,))
                    exists = cursor.fetchone()[0] > 0

                    if exists:
                        # Обновляем существующую запись
                        update_sql = f"UPDATE amendments SET {field_name} = ? WHERE InputData_ID = ?"
                        cursor.execute(update_sql, (value, input_data_id))
                    else:
                        # Вставляем новую запись
                        insert_sql = f"INSERT INTO amendments (InputData_ID, {field_name}) VALUES (?, ?)"
                        cursor.execute(insert_sql, (input_data_id, value))

            conn.commit()
            logging.info(f"Поправки сохранены для записи ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки поправок: {str(e)}")
            raise

    def get_parameters(self, input_data_id):
        """Получает параметры исследования для указанной записи"""
        try:
            conn = self.get_connection()
            query = "SELECT * FROM calculatedParameters WHERE InputData_ID = ?"
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df
        except Exception as e:
            logging.error(f"Ошибка получения параметров: {str(e)}")
            return pd.DataFrame()

    def get_research_params(self, main_data_id):
        """Алиас для get_parameters (для совместимости)"""
        return self.get_parameters(main_data_id)

    def get_amendments(self, input_data_id):
        """Получает поправки для указанной записи"""
        try:
            conn = self.get_connection()
            query = "SELECT * FROM amendments WHERE InputData_ID = ?"
            df = pd.read_sql(query, conn, params=[input_data_id])

            if not df.empty:
                # Возвращаем первую запись (предполагаем одну запись на InputData_ID)
                return df.iloc[0].to_dict()
            else:
                # Возвращаем пустой словарь с правильными ключами
                columns = self.get_table_columns('amendments')
                return {col: None for col in columns if col != 'ID'}

        except Exception as e:
            logging.error(f"Ошибка получения поправок: {str(e)}")
            return {}


    @staticmethod
    def convert_parameter_value(self, value):
        """Конвертирует значение параметра в подходящий тип для БД"""
        if value is None:
            return None

        if isinstance(value, (int, float)):
            return value

        if isinstance(value, str):
            # Пробуем преобразовать строку в число
            try:
                # Убираем пробелы и запятые (для чисел с разделителями)
                clean_value = value.replace(' ', '').replace(',', '.')
                return float(clean_value)
            except ValueError:
                # Если не число, оставляем как строку
                return value.strip()

        # Для других типов преобразуем в строку
        return str(value)

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

    def save_amendments(self):
        """Сохраняет поправки из полей ввода GUI в базу данных"""
        try:
            # Получаем ID основной записи
            main_data_id = getattr(self, 'current_main_data_id', None)
            if main_data_id is None:
                last_record = self.get_last_record()
                if not last_record.empty:
                    main_data_id = last_record.iloc[0]['ID']
                else:
                    messagebox.showerror("Ошибка", "Нет основной записи для привязки поправок")
                    return

            # Собираем поправки из полей ввода GUI
            amendments_data = {}

            # Поправки Рпл (первая модель)
            amendments_data['поправка с насоса на ВНК Рпл'] = self.ppl_entries[0].get()
            amendments_data['поправка с насоса на ВДП Рпл'] = self.ppl_entries[1].get()
            amendments_data['поправка с насоса на ГНК Рпл'] = self.ppl_entries[2].get()

            # Поправки Рзаб (первая модель)
            amendments_data['поправка с насоса на ВНК Рзаб'] = self.pzab_entries[0].get()
            amendments_data['поправка с насоса на ВДП Рзаб'] = self.pzab_entries[1].get()
            amendments_data['поправка с насоса на ГНК Рзаб'] = self.pzab_entries[2].get()

            # Поправки Рпл (вторая модель)
            amendments_data['поправка с насоса на ВНК Рпл'] = self.ppl2_entries[0].get()
            amendments_data['поправка с насоса на ВДП Рпл'] = self.ppl2_entries[1].get()
            amendments_data['поправка с насоса на ГНК Рпл'] = self.ppl2_entries[2].get()

            # Поправки Рзаб (вторая модель)
            amendments_data['поправка с насоса на ВНК Рзаб'] = self.pzab2_entries[0].get()
            amendments_data['поправка с насоса на ВДП Рзаб'] = self.pzab2_entries[1].get()
            amendments_data['поправка с насоса на ГНК Рзаб'] = self.pzab2_entries[2].get()

            # Дополнительные поправки
            amendments_data['ВНК_Рпл_3'] = self.vnkp_pl3_entry.get()
            amendments_data['ВНК_Рпл_4'] = self.vnkp_pl4_entry.get()

            # Фильтруем пустые значения
            amendments_data = {k: v for k, v in amendments_data.items() if v and v.strip()}

            if not amendments_data:
                messagebox.showinfo("Информация", "Нет данных поправок для сохранения")
                return

            # Сохраняем поправки в базу
            self.insert_amendments(main_data_id, amendments_data)
            messagebox.showinfo("Успех", "Поправки успешно сохранены в базу данных")

            # Очищаем поля после сохранения
            self.clear_amendments_fields()

        except Exception as e:
            logging.error(f"Ошибка сохранения поправок: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка при сохранении поправок: {str(e)}")

    def clear_amendments_fields(self):
        """Очищает поля ввода поправок после сохранения"""
        # Очищаем все поля поправок
        for entry in self.ppl_entries + self.pzab_entries + self.ppl2_entries + self.pzab2_entries:
            entry.delete(0, 'end')

        self.vnkp_pl3_entry.delete(0, 'end')
        self.vnkp_pl4_entry.delete(0, 'end')

    def get_float_value(self, value):
        """Конвертирует значение в float, возвращает None если пустое или не число"""
        if not value:
            return None
        try:
            return float(value.replace(',', '.'))
        except ValueError:
            return None

    def update_main_data(self, record_id, update_data):
        """Обновляет данные в соответствующих таблицах"""
        try:
            # Сохраняем успешность
            if 'Uspeshnost' in update_data and update_data['Uspeshnost'] is not None:
                self.insert_success(record_id, update_data['Uspeshnost'])

            # Сохраняем класс исследования
            if 'Klass_issledovaniya' in update_data and update_data['Klass_issledovaniya'] is not None:
                self.insert_research_class(record_id, update_data['Klass_issledovaniya'])

            # Сохраняем давление
            if 'PressureLastPoint' in update_data and update_data['PressureLastPoint'] is not None:
                self.insert_pressure_last_point(record_id, update_data['PressureLastPoint'])

            # Сохраняем расчетное время
            if 'Durat' in update_data and update_data['Durat'] is not None:
                self.insert_estimated_time(record_id, update_data['Durat'])

            # Сохраняем плотность
            if 'density_zab' in update_data or 'density_pl' in update_data:
                density_zab = update_data.get('density_zab')
                density_pl = update_data.get('density_pl')
                self.insert_density(record_id, density_zab, density_pl)

            logging.info(f"Все данные для записи {record_id} обновлены")
            return True

        except Exception as e:
            logging.error(f"Ошибка обновления данных: {str(e)}")
            raise

    def get_success(self, input_data_id):
        """Получает успешность для записи"""
        try:
            conn = self.get_connection()
            query = "SELECT Uspeshnost FROM success WHERE InputData_ID = ? ORDER BY ID DESC"
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df
        except Exception as e:
            logging.error(f"Ошибка получения успешности: {str(e)}")
            return pd.DataFrame()

    def get_research_class(self, input_data_id):
        """Получает класс исследования для записи"""
        try:
            conn = self.get_connection()
            query = "SELECT Klass_issledovaniya FROM researchClass WHERE InputData_ID = ? ORDER BY ID DESC"
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df
        except Exception as e:
            logging.error(f"Ошибка получения класса исследования: {str(e)}")
            return pd.DataFrame()

    def get_pressure_last_point(self, input_data_id):
        """Получает давление для записи"""
        try:
            conn = self.get_connection()
            query = "SELECT pressureLP FROM pressureLastPoint WHERE InputData_ID = ? ORDER BY ID DESC"
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df
        except Exception as e:
            logging.error(f"Ошибка получения давления: {str(e)}")
            return pd.DataFrame()

    def get_estimated_time(self, input_data_id):
        """Получает расчетное время для записи"""
        try:
            conn = self.get_connection()
            query = "SELECT Durat FROM estimatedTime WHERE InputData_ID = ? ORDER BY ID DESC"
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df
        except Exception as e:
            logging.error(f"Ошибка получения расчетного времени: {str(e)}")
            return pd.DataFrame()

    def get_density(self, input_data_id):
        """Получает плотность для записи"""
        try:
            conn = self.get_connection()
            query = "SELECT density_zab, density_pl FROM density WHERE InputData_ID = ? ORDER BY ID DESC"
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df
        except Exception as e:
            logging.error(f"Ошибка получения плотности: {str(e)}")
            return pd.DataFrame()

    def check_table_structure(self, table_name):
        """Проверяет реальную структуру таблицы"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute(f"SELECT TOP 1 * FROM {table_name}")
            columns = [column[0] for column in cursor.description]

            print(f"Реальная структура таблицы {table_name}:")
            for col in columns:
                print(f"  - {col}")

            return columns

        except Exception as e:
            print(f"Ошибка проверки структуры {table_name}: {e}")
            return []

    def add_foreign_keys_to_tables(self):
        """Добавляет поля для связи с InputData во все таблицы"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Для каждой таблицы добавляем поле InputData_ID
            tables = [
                'success',
                'researchClass',
                'pressureLastPoint',
                'estimatedTime',
                'density',
                'calculatedParameters',
                'amendments'
            ]

            for table_name in tables:
                try:
                    # Проверяем, есть ли уже поле связи
                    cursor.execute(f"SELECT TOP 1 * FROM {table_name}")
                    columns = [column[0] for column in cursor.description]

                    if 'InputData_ID' not in columns:
                        # Добавляем поле для связи
                        cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN InputData_ID INTEGER")
                        print(f"✓ Добавлено поле InputData_ID в таблицу {table_name}")
                    else:
                        print(f"✓ Поле InputData_ID уже есть в таблице {table_name}")

                except Exception as e:
                    print(f"✗ Ошибка добавления поля в {table_name}: {e}")

            conn.commit()
            cursor.close()
            print("Все поля для связи добавлены!")

        except Exception as e:
            print(f"Ошибка: {e}")

    def calculate_pressure_values(self, input_data_id, pressure_last_point=None):
        """Рассчитывает значения для calculatedPressure"""
        try:
            # Получаем необходимые данные
            vid_issledovaniya = self.get_vid_issledovaniya(input_data_id)
            pzab_vnk = None
            ppl_vnk = None

            # Логика расчета PzabVnk
            if vid_issledovaniya and any(x in vid_issledovaniya.upper() for x in ['КСД', 'АДД']):
                last_pressure = self.get_last_pressure_vnk(input_data_id)
                pzab_vnk = last_pressure
            else:
                p_param = self.get_calculated_parameter(input_data_id, 'P @ dt=0')
                pzab_vnk = p_param

            # Логика расчета PplVnk
            has_model_vnk = self.has_model_vnk_data(input_data_id)
            has_model_ksd = self.has_model_ksd_data(input_data_id)

            if has_model_vnk:
                ppl_vnk = self.get_last_model_vnk_pressure(input_data_id)
            elif vid_issledovaniya and any(x in vid_issledovaniya.upper() for x in ['КСД', 'АДД']):
                ppl_vnk = self.get_last_model_ksd_pressure(input_data_id)
            elif vid_issledovaniya and 'ГРП' not in vid_issledovaniya.upper():
                initial_pressure = self.get_calculated_parameter(input_data_id, 'Начальное пластовое давление')
                ppl_vnk = initial_pressure
            else:
                ppl_vnk = pressure_last_point  # Используем переданное значение

            # Получаем поправки
            amendments = self.get_amendments(input_data_id)

            # Используем поправки из соответствующей модели
            if has_model_vnk or ('КПД' in str(vid_issledovaniya or '')):
                # Используем основные поправки (первая модель)
                amend_vnk_ppl = amendments.get('amendVnkPpl', 0)
                amend_vdp_ppl = amendments.get('amendVdpPpl', 0)
                amend_gnk_ppl = amendments.get('amendGnkPpl', 0)
                amend_vnk_pzab = amendments.get('amendVnkPzab', 0)
                amend_vdp_pzab = amendments.get('amendVdpPzab', 0)
                amend_gnk_pzab = amendments.get('amendGnkPzab', 0)
            else:
                # Используем вторую модель поправок
                amend_vnk_ppl = amendments.get('amendVnkPpl2', amendments.get('amendVnkPpl', 0))
                amend_vdp_ppl = amendments.get('amendVdpPpl2', amendments.get('amendVdpPpl', 0))
                amend_gnk_ppl = amendments.get('amendGnkPpl2', amendments.get('amendGnkPpl', 0))
                amend_vnk_pzab = amendments.get('amendVnkPzab2', amendments.get('amendVnkPzab', 0))
                amend_vdp_pzab = amendments.get('amendVdpPzab2', amendments.get('amendVdpPzab', 0))
                amend_gnk_pzab = amendments.get('amendGnkPzab2', amendments.get('amendGnkPzab', 0))

            # Рассчитываем остальные параметры
            calculated_data = {
                'PplVnk': ppl_vnk,
                'PzabVnk': pzab_vnk,
                'PplGlubZam': ppl_vnk - amend_vnk_ppl if ppl_vnk is not None else None,
                'PplVdp': (ppl_vnk - amend_vnk_ppl) + amend_vdp_ppl if ppl_vnk is not None else None,
                'PplGnk': (ppl_vnk - amend_vnk_ppl) + amend_gnk_ppl if ppl_vnk is not None else None,
                'PzabGlubZam': pzab_vnk - amend_vnk_pzab if pzab_vnk is not None else None,
                'PzabVdp': (pzab_vnk - amend_vnk_pzab) + amend_vdp_pzab if pzab_vnk is not None else None,
                'PzabGnk': (pzab_vnk - amend_vnk_pzab) + amend_gnk_pzab if pzab_vnk is not None else None
            }

            # Фильтруем None значения
            calculated_data = {k: v for k, v in calculated_data.items() if v is not None}

            # Сохраняем рассчитанные данные
            if calculated_data:
                self.insert_calculated_pressure(input_data_id, calculated_data)

            return calculated_data

        except Exception as e:
            logging.error(f"Ошибка расчета давлений: {str(e)}")
            raise

    def get_table_columns(self, table_name):
        """Получает список колонок таблицы"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute(f"SELECT TOP 1 * FROM {table_name}")
            return [column[0] for column in cursor.description]
        except Exception as e:
            logging.error(f"Ошибка получения колонок {table_name}: {str(e)}")
            return []

    def get_vid_issledovaniya(self, input_data_id):
        """Получает вид исследования из InputData"""
        try:
            conn = self.get_connection()
            query = "SELECT Vid_issledovaniya FROM InputData WHERE ID = ?"
            df = pd.read_sql(query, conn, params=[input_data_id])
            return df.iloc[0]['Vid_issledovaniya'] if not df.empty else None
        except Exception as e:
            logging.error(f"Ошибка получения вида исследования: {str(e)}")
            return None


    def insert_calculated_pressure(self, input_data_id, calculated_data):
        """Вставляет рассчитанные давления в calculatedPressure"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Фильтруем только существующие поля
            existing_fields = self.get_table_columns('calculatedPressure')
            valid_data = {k: v for k, v in calculated_data.items() if k in existing_fields and v is not None}

            if valid_data:
                # Добавляем InputData_ID в поля и значения
                fields = ['InputData_ID'] + list(valid_data.keys())
                values = [input_data_id] + list(valid_data.values())
                placeholders = ['?'] * len(values)

                sql = f"INSERT INTO calculatedPressure ({', '.join(fields)}) VALUES ({', '.join(placeholders)})"
                cursor.execute(sql, values)
                conn.commit()

                logging.info(f"Рассчитанные давления сохранены для записи ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки рассчитанных давлений: {str(e)}")
            raise

    def get_calculated_parameter(self, input_data_id, param_name):
        """Получает значение параметра из calculatedParameters"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute("""
                           SELECT Val
                           FROM calculatedParameters
                           WHERE InputData_ID = ?
                             AND calcParam = ?
                           """, (input_data_id, param_name))

            result = cursor.fetchone()
            return result[0] if result else None

        except Exception as e:
            print(f"Ошибка получения параметра {param_name}: {e}")
            return None

    def update_calculate_table_improved(self, input_data_id):
        """Улучшенная версия обновления calculate"""
        try:
            import math

            conn = self.get_connection()
            cursor = conn.cursor()

            # Получаем все необходимые параметры
            params = {}
            param_names = [
                'P @ dt=0', 'Delta Q', 'h', 'k', 'вязкость', 'µгаза', 'Phi',
                'Общая сжимаемость сt', 'Тип флюида'
            ]

            for param_name in param_names:
                params[param_name] = self.get_calculated_parameter(input_data_id, param_name)

            # Получаем данные из InputData
            cursor.execute("""
                           SELECT Obvodnennost, Vremya_issledovaniya
                           FROM InputData
                           WHERE ID = ?
                           """, (input_data_id,))
            input_data = cursor.fetchone()
            obvodnennost, vremya_issledovaniya = input_data if input_data else (None, None)

            # 1. P1_zab_vnk
            p1_zab_vnk = params['P @ dt=0']

            # 2. P2_zab_vnk - последнее значение из PressureVNK
            cursor.execute("""
                           SELECT TOP 1 PressureVnk
                           FROM PressureVNK
                           WHERE InputData_ID = ?
                           ORDER BY ID DESC
                           """, (input_data_id,))
            p2_zab_vnk = cursor.fetchone()
            p2_zab_vnk = p2_zab_vnk[0] if p2_zab_vnk else None

            # 3. Pday - разница давлений за сутки
            cursor.execute("""
                           SELECT MAX(PressureVnk) - MIN(PressureVnk)
                           FROM PressureVNK
                           WHERE InputData_ID = ?
                             AND Dat >= DATEADD('d', -1, (SELECT MAX(Dat) FROM PressureVNK WHERE InputData_ID = ?))
                           """, (input_data_id, input_data_id))
            pday = cursor.fetchone()
            pday = pday[0] if pday else None

            # 4. P_pl_vnk - значение из calculatedPressure
            cursor.execute("""
                           SELECT PplVnk
                           FROM calculatedPressure
                           WHERE InputData_ID = ?
                           """, (input_data_id,))
            p_pl_vnk = cursor.fetchone()
            p_pl_vnk = p_pl_vnk[0] if p_pl_vnk else None

            # 5. delta - разница с предыдущим исследованием (упрощенно)
            delta = None  # Нужна дополнительная логика для предыдущих исследований

            # 6. productivity - модуль отношения Delta Q к разнице давлений
            cursor.execute("""
                           SELECT Val
                           FROM calculatedParameters
                           WHERE InputData_ID = ?
                             AND calcParam = 'Delta Q'
                           """, (input_data_id,))
            delta_q = cursor.fetchone()
            delta_q = delta_q[0] if delta_q else None

            productivity = abs(delta_q / (p1_zab_vnk - p_pl_vnk)) if delta_q and p1_zab_vnk and p_pl_vnk else None

            # 7. Delta Q
            delta_q_value = delta_q

            # 8. Qoil = Delta Q*(1-Obvodnennost/100)*Vremya_issledovaniya/24
            cursor.execute("""
                           SELECT Obvodnennost, Vremya_issledovaniya
                           FROM InputData
                           WHERE ID = ?
                           """, (input_data_id,))
            input_data = cursor.fetchone()

            qoil = None
            if input_data and delta_q:
                obvodnennost, vremya = input_data
                qoil = delta_q * (1 - obvodnennost / 100) * vremya / 24

            # 9. Kh/Mu
            fluid_type = str(params['Тип флюида'] or '').lower()
            kh_mu = None

            if all(params.get(key) for key in ['h', 'k']):
                if 'газ' in fluid_type and params['µгаза']:
                    kh_mu = (params['h'] * 100 * params['k'] / 1000) / params['µгаза']
                elif params['вязкость']:
                    kh_mu = (params['h'] * 100 * params['k'] / 1000) / params['вязкость']

            # 10. Rinv_Ppl1
            rinv_ppl1 = None
            if (vremya_issledovaniya and all(params.get(key) for key in ['k', 'Phi', 'Общая сжимаемость сt'])):
                if 'газ' in fluid_type and params['µгаза']:
                    rinv_ppl1 = 0.037 * math.sqrt(
                        (params['k'] * vremya_issledovaniya) /
                        (params['Phi'] * params['µгаза'] * params['Общая сжимаемость сt'])
                    )
                elif params['вязкость']:
                    rinv_ppl1 = 0.037 * math.sqrt(
                        (params['k'] * vremya_issledovaniya) /
                        (params['Phi'] * params['вязкость'] * params['Общая сжимаемость сt'])
                    )

            # Вставляем данные
            cursor.execute("""
                           INSERT INTO calculate (P1_zab_vnk, P2_zab_vnk, Pday, P_pl_vnk, delta, productivity,
                                                  Delta_Q, Qoil, Kh_Mu, Rinv_Ppl1, InputData_ID)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                           """, (p1_zab_vnk, p2_zab_vnk, pday, p_pl_vnk, None, None,
                                 params['Delta Q'], None, kh_mu, rinv_ppl1, input_data_id))

            conn.commit()

        except Exception as e:
            print(f"Ошибка обновления calculate: {e}")
            raise

    def check_column_types(self, table_name):
        """Проверяет точные типы данных столбцов в таблице"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Для Access можно попробовать такой запрос
            cursor.execute(f"SELECT TOP 1 * FROM {table_name}")
            description = cursor.description

            print(f"\n=== ТИПЫ ДАННЫХ ТАБЛИЦЫ {table_name} ===")
            for column in description:
                col_name = column[0]
                col_type_code = column[1]  # Код типа данных ODBC
                col_type_name = self.get_odbc_type_name(col_type_code)
                print(f"{col_name}: {col_type_name} (код: {col_type_code})")

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

    def add_sort_order_to_model_vnk(self):
        """Добавляет поле для сортировки в таблицу ModelVNK"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Проверяем, есть ли уже поле sort_order
            cursor.execute("SELECT TOP 1 * FROM ModelVNK")
            columns = [column[0] for column in cursor.description]

            if 'sort_order' not in columns:
                # Добавляем поле для сортировки
                cursor.execute("ALTER TABLE ModelVNK ADD COLUMN sort_order INTEGER")
                print("Добавлено поле sort_order в ModelVNK")
            else:
                print("Поле sort_order уже существует")

            conn.commit()

        except Exception as e:
            print(f"Ошибка добавления поля сортировки: {e}")

    def update_calculate_table(self, input_data_id):
        """Обновляет расчетную таблицу calculate"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # 1. P1_zab_vnk - значение из calculatedParameters где calcParam = "P @ dt=0"
            cursor.execute("""
                           SELECT Val
                           FROM calculatedParameters
                           WHERE InputData_ID = ?
                             AND calcParam = 'P @ dt=0'
                           """, (input_data_id,))
            p1_zab_vnk = cursor.fetchone()
            p1_zab_vnk = p1_zab_vnk[0] if p1_zab_vnk else None

            # 2. P2_zab_vnk - последнее значение из PressureVNK
            cursor.execute("""
                           SELECT TOP 1 PressureVnk
                           FROM PressureVNK
                           WHERE InputData_ID = ?
                           ORDER BY ID DESC
                           """, (input_data_id,))
            p2_zab_vnk = cursor.fetchone()
            p2_zab_vnk = p2_zab_vnk[0] if p2_zab_vnk else None

            # 3. Pday - разница давлений за сутки
            cursor.execute("""
                           SELECT MAX(PressureVnk) - MIN(PressureVnk)
                           FROM PressureVNK
                           WHERE InputData_ID = ?
                             AND Dat >= DATEADD('d', -1, (SELECT MAX(Dat) FROM PressureVNK WHERE InputData_ID = ?))
                           """, (input_data_id, input_data_id))
            pday = cursor.fetchone()
            pday = pday[0] if pday else None

            # 4. P_pl_vnk - значение из calculatedPressure
            cursor.execute("""
                           SELECT PplVnk
                           FROM calculatedPressure
                           WHERE InputData_ID = ?
                           """, (input_data_id,))
            p_pl_vnk = cursor.fetchone()
            p_pl_vnk = p_pl_vnk[0] if p_pl_vnk else None

            # 5. delta - разница между текущим P_pl_vnk и предыдущим исследованием
            delta = None
            if p_pl_vnk:
                # Ищем предыдущее исследование для этой же скважины
                cursor.execute("""
                               SELECT i.ID, p.PplVnk
                               FROM InputData i
                                        INNER JOIN prevData p ON i.ID = p.InputData_ID
                               WHERE i.Skvazhina = (SELECT Skvazhina
                                                    FROM InputData
                                                    WHERE ID = ?)
                                 AND i.ID != ?
                               ORDER BY i.Data_issledovaniya DESC
                               """, (input_data_id, input_data_id))

                prev_research = cursor.fetchone()
                if prev_research:
                    prev_ppl_vnk = prev_research[1]
                    delta = p_pl_vnk - prev_ppl_vnk
                    print(f"Delta calculated: {p_pl_vnk} - {prev_ppl_vnk} = {delta}")

            # 6. productivity - модуль отношения Delta Q к разнице давлений
            cursor.execute("""
                           SELECT Val
                           FROM calculatedParameters
                           WHERE InputData_ID = ?
                             AND calcParam = 'Delta Q'
                           """, (input_data_id,))
            delta_q = cursor.fetchone()
            delta_q = delta_q[0] if delta_q else None

            productivity = abs(delta_q / (p1_zab_vnk - p_pl_vnk)) if delta_q and p1_zab_vnk and p_pl_vnk else None

            # 7. Delta Q
            delta_q_value = delta_q

            # 8. Qoil = Delta Q*(1-Obvodnennost/100)*Vremya_issledovaniya/24
            cursor.execute("""
                           SELECT Obvodnennost, Vremya_issledovaniya
                           FROM InputData
                           WHERE ID = ?
                           """, (input_data_id,))
            input_data = cursor.fetchone()

            qoil = None
            if input_data and delta_q:
                obvodnennost, vremya = input_data
                qoil = delta_q * (1 - obvodnennost / 100) * vremya / 24

            # 9. Kh/Mu - гидропроводность
            cursor.execute("""
                           SELECT Val
                           FROM calculatedParameters
                           WHERE InputData_ID = ?
                             AND calcParam IN ('h', 'k', 'вязкость', 'µгаза')
                           """, (input_data_id,))
            params = cursor.fetchall()

            # Получаем тип флюида
            cursor.execute("""
                           SELECT Val
                           FROM calculatedParameters
                           WHERE InputData_ID = ?
                             AND calcParam = 'Тип флюида'
                           """, (input_data_id,))
            fluid_type = cursor.fetchone()
            fluid_type = fluid_type[0].lower() if fluid_type else ''

            kh_mu = None
            if params and len(params) >= 3:
                if 'газ' in fluid_type:
                    # Для газа: Kh/Mu = (h*100 * k/1000) / µгаза
                    kh_mu = (params[0][0] * 100 * params[1][0] / 1000) / params[3][0] if len(params) >= 4 else None
                else:
                    # Для жидкости: Kh/Mu = (h*100 * k/1000) / вязкость
                    kh_mu = (params[0][0] * 100 * params[1][0] / 1000) / params[2][0]

            # 10. Rinv_Ppl1
            cursor.execute("""
                           SELECT Val
                           FROM calculatedParameters
                           WHERE InputData_ID = ?
                             AND calcParam IN ('k', 'Phi', 'Вязкость', 'µгаза', 'Общая сжимаемость сt')
                           """, (input_data_id,))
            rinv_params = cursor.fetchall()

            rinv_ppl1 = None
            if input_data and rinv_params and len(rinv_params) >= 4:
                vremya = input_data[1]  # Vremya_issledovaniya
                if 'газ' in fluid_type:
                    # Для газа
                    rinv_ppl1 = 0.037 * math.sqrt((rinv_params[0][0] * vremya) /
                                                  (rinv_params[1][0] * rinv_params[3][0] * rinv_params[4][0]))
                else:
                    # Для жидкости
                    rinv_ppl1 = 0.037 * math.sqrt((rinv_params[0][0] * vremya) /
                                                  (rinv_params[1][0] * rinv_params[2][0] * rinv_params[4][0]))

            # Вставляем или обновляем запись
            cursor.execute("""
                           INSERT INTO calculate (P1_zab_vnk, P2_zab_vnk, Pday, P_pl_vnk, delta, productivity,
                                                  Delta_Q, Qoil, Kh_Mu, Rinv_Ppl1, InputData_ID)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                           """, (p1_zab_vnk, p2_zab_vnk, pday, p_pl_vnk, delta, productivity,
                                 delta_q_value, qoil, kh_mu, rinv_ppl1, input_data_id))

            conn.commit()
            print("Таблица calculate обновлена")

        except Exception as e:
            print(f"Ошибка обновления calculate: {e}")
            raise

    def insert_prev_data(self, data_dict):
        """Вставляет данные в таблицу prevData"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            sql = """
                  INSERT INTO prevData (PplVnk, PzabVnk, DataRes, Water, Q, Kprod, Smeh, Heff, Kgidr, InputData_ID)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) \
                  """

            cursor.execute(sql, (
                data_dict.get('PplVnk'),
                data_dict.get('PzabVnk'),
                data_dict.get('DataRes'),
                data_dict.get('Water'),
                data_dict.get('Q'),
                data_dict.get('Kprod'),
                data_dict.get('Smeh'),
                data_dict.get('Heff'),
                data_dict.get('Kgidr'),
                data_dict.get('InputData_ID')
            ))

            conn.commit()
            logging.info(f"Данные предыдущего исследования сохранены для ID: {data_dict.get('InputData_ID')}")

        except Exception as e:
            logging.error(f"Ошибка вставки в prevData: {str(e)}")
            raise

# db = AccessDatabase()
# db.add_sort_order_to_model_vnk()