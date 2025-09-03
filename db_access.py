import pyodbc
import os
from utils import logger, logging
import pandas as pd
from datetime import datetime, time, timedelta
from tkinter import messagebox
import math
import sys


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
        """Возвращает последнюю запись из InputData в правильном формате"""
        try:
            conn = self.get_connection()
            query = "SELECT TOP 1 * FROM InputData ORDER BY ID DESC"
            df = pd.read_sql(query, conn)

            if df.empty:
                return pd.DataFrame()

            # Убедимся, что все колонки присутствуют
            expected_columns = ['ID', 'Company', 'Localoredenie', 'Skvazhina', 'VNK',
                                'Data_issledovaniya', 'Plast', 'Interval_perforacii',
                                'Tip_pribora', 'Glubina_ustanovki_pribora', 'Interpretator',
                                'Data_interpretacii', 'Vremya_issledovaniya', 'Obvodnennost',
                                'Nalicie_paktera', 'Data_GRP', 'Vid_issledovaniya', 'created_date']

            # Добавляем отсутствующие колонки с значениями None
            for col in expected_columns:
                if col not in df.columns:
                    df[col] = None

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
            'amendments', 'InputData', 'prevData',
            'researchClass', 'success', 'density', 'estimatedTime',
            'ModelVNK', 'PressureVNK', 'pressureLastPoint', 'TextParameters',
            'ModelKSD', 'calculate'
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

            # Очищаем предыдущие данные
            cursor.execute("DELETE FROM ModelVNK WHERE InputData_ID = ?", (input_data_id,))
            conn.commit()

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
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM ModelVNK WHERE InputData_ID = ?", (input_data_id,))
            result = cursor.fetchone()
            return result[0] > 0 if result else False
        except Exception as e:
            logging.error(f"Ошибка проверки ModelVNK: {str(e)}")
            return False

    def get_last_model_vnk_pressure(self, input_data_id):
        """Получает последнее значение PressureVnkModel из ModelVNK"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Простой запрос без сложных условий
            cursor.execute("""
                           SELECT TOP 1 PressureVnkModel
                           FROM ModelVNK
                           WHERE InputData_ID = ?
                           ORDER BY Dat DESC
                           """, (input_data_id,))

            result = cursor.fetchone()
            return result[0] if result else None

        except Exception as e:
            logging.error(f"Ошибка получения давления ModelVNK: {str(e)}")
            return None

    def insert_pressure_vnk(self, input_data_id, data_list):
        """Самый быстрый метод для Access"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Очищаем предыдущие данные
            cursor.execute("DELETE FROM PressureVNK WHERE InputData_ID = ?", (input_data_id,))
            conn.commit()

            if not data_list:
                return

            # Отключаем индексы и триггеры для ускорения
            try:
                cursor.execute("ALTER TABLE PressureVNK DISABLE INDEX ALL")
            except:
                pass  # Игнорируем если нельзя отключить индексы

            # Вставляем данные с отключенным авто-коммитом
            conn.autocommit = False
            total_inserted = 0

            for data in data_list:
                try:
                    cursor.execute(
                        "INSERT INTO PressureVNK (Dat, PressureVnk, InputData_ID) VALUES (?, ?, ?)",
                        (data.get('Dat'), data.get('PressureVnk'), input_data_id)
                    )
                    total_inserted += 1

                    # Коммитим каждые 1000 записей
                    if total_inserted % 1000 == 0:
                        conn.commit()
                        logging.info(f"Вставлено {total_inserted} записей...")

                except Exception as e:
                    logging.warning(f"Ошибка вставки записи: {e}")
                    continue

            # Финальный коммит
            conn.commit()

            # Включаем индексы обратно
            try:
                cursor.execute("ALTER TABLE PressureVNK ENABLE INDEX ALL")
            except:
                pass

            conn.autocommit = True
            logging.info(f"Данные PressureVNK сохранены для ID: {input_data_id}, всего записей: {total_inserted}")

        except Exception as e:
            conn.rollback()
            logging.error(f"Ошибка вставки PressureVNK: {str(e)}")
            raise
        finally:
            # Гарантируем, что авто-коммит включен
            try:
                conn.autocommit = True
            except:
                pass

    def insert_pressure_vnk_fastest(self, input_data_id, data_list):
        """Самый быстрый метод через CSV и ODBC bulk insert"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Очищаем предыдущие данные
            cursor.execute("DELETE FROM PressureVNK WHERE InputData_ID = ?", (input_data_id,))
            conn.commit()

            if not data_list:
                return

            # Создаем временный CSV файл
            import tempfile
            import csv

            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tmp_file:
                writer = csv.writer(tmp_file, delimiter=';')
                writer.writerow(['Dat', 'PressureVnk', 'InputData_ID'])

                for data in data_list:
                    writer.writerow([
                        data.get('Dat'),
                        data.get('PressureVnk'),
                        input_data_id
                    ])

                tmp_path = tmp_file.name

            try:
                # Используем ODBC bulk insert
                cursor.execute(f"""
                    INSERT INTO PressureVNK (Dat, PressureVnk, InputData_ID)
                    SELECT * FROM [Text;HDR=YES;DATABASE={os.path.dirname(tmp_path)}].[{os.path.basename(tmp_path)}]
                """)
                conn.commit()

            finally:
                # Удаляем временный файл
                os.unlink(tmp_path)

            logging.info(f"Данные PressureVNK сохранены через bulk insert для ID: {input_data_id}")

        except Exception as e:
            conn.rollback()
            logging.error(f"Ошибка bulk insert PressureVNK: {str(e)}")
            raise

    def insert_pressure_vnk_text_file(self, input_data_id, data_list):
        """Быстрая вставка через текстовый файл с Schema.ini"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Очищаем предыдущие данные
            cursor.execute("DELETE FROM PressureVNK WHERE InputData_ID = ?", (input_data_id,))
            conn.commit()

            if not data_list:
                return

            # Создаем директорию для временных файлов
            temp_dir = os.path.join(os.path.dirname(self.db_path), "temp_import")
            os.makedirs(temp_dir, exist_ok=True)

            # Создаем Schema.ini для правильного формата
            schema_ini = os.path.join(temp_dir, "Schema.ini")
            with open(schema_ini, 'w', encoding='utf-8') as f:
                f.write("""[pressure_data.txt]
    Format=Delimited(;)
    ColNameHeader=True
    DateTimeFormat=dd.mm.yyyy hh:nn:ss
    DecimalSymbol=.
    Col1=Dat DateTime
    Col2=PressureVnk Float
    Col3=InputData_ID Integer
    """)

            # Создаем data file
            data_file = os.path.join(temp_dir, "pressure_data.txt")
            with open(data_file, 'w', encoding='utf-8') as f:
                f.write("Dat;PressureVnk;InputData_ID\n")
                for data in data_list:
                    f.write(f"{data.get('Dat')};{data.get('PressureVnk')};{input_data_id}\n")

            # Импортируем через SQL
            cursor.execute(f"""
                INSERT INTO PressureVNK (Dat, PressureVnk, InputData_ID)
                SELECT * FROM [Text;DATABASE={temp_dir}].[pressure_data.txt]
            """)
            conn.commit()

            # Очищаем временные файлы
            try:
                os.remove(schema_ini)
                os.remove(data_file)
                os.rmdir(temp_dir)
            except:
                pass

            logging.info(f"Данные PressureVNK сохранены через text import для ID: {input_data_id}")

        except Exception as e:
            conn.rollback()
            logging.error(f"Ошибка text import PressureVNK: {str(e)}")
            raise

    def get_last_pressure_vnk(self, input_data_id):
        """Получает последнее значение PressureVnk из PressureVNK"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute("""
                           SELECT TOP 1 PressureVnk
                           FROM PressureVNK
                           WHERE InputData_ID = ?
                           ORDER BY Dat DESC
                           """, (input_data_id,))

            result = cursor.fetchone()
            return result[0] if result else None

        except Exception as e:
            logging.error(f"Ошибка получения последнего давления: {str(e)}")
            return None

    def insert_model_ksd(self, input_data_id, data_list):
        """Вставляет данные в ModelKSD с сохранением порядка"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Очищаем предыдущие данные
            cursor.execute("DELETE FROM ModelKSD WHERE InputData_ID = ?", (input_data_id,))
            conn.commit()

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
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM ModelKSD WHERE InputData_ID = ?", (input_data_id,))
            result = cursor.fetchone()
            return result[0] > 0 if result else False
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
        """Сохраняет поправки в соответствующие таблицы"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Разделяем поправки по таблицам
            amendments_1 = {}  # Для amendments (пласт 1)
            amendments_2 = {}  # Для amendments2 (пласт 2)
            amendments_3 = {}  # Для amendments3 (пласт 3)
            amendments_4 = {}  # Для amendments4 (пласт 4)

            for field_name, value in amendments_dict.items():
                if field_name.endswith('2'):
                    # Пласт 2 -> amendments2
                    clean_name = field_name.replace('2', '')
                    amendments_2[clean_name] = value
                elif field_name.endswith('3'):
                    # Пласт 3 -> amendments3
                    clean_name = field_name.replace('3', '')
                    amendments_3[clean_name] = value
                elif field_name.endswith('4'):
                    # Пласт 4 -> amendments4
                    clean_name = field_name.replace('4', '')
                    amendments_4[clean_name] = value
                else:
                    # Пласт 1 -> amendments
                    amendments_1[field_name] = value

            # Сохраняем в соответствующие таблицы
            if amendments_1:
                self._save_to_amendments_table('amendments', input_data_id, amendments_1)
            if amendments_2:
                self._save_to_amendments_table('amendments2', input_data_id, amendments_2)
            if amendments_3:
                self._save_to_amendments_table('amendments3', input_data_id, amendments_3)
            if amendments_4:
                self._save_to_amendments_table('amendments4', input_data_id, amendments_4)

        except Exception as e:
            logging.error(f"Ошибка сохранения поправок: {str(e)}")
            raise

    def _save_to_amendments_table(self, table_name, input_data_id, amendments_data):
        """Сохраняет поправки в указанную таблицу"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Очищаем старые данные
            cursor.execute(f"DELETE FROM {table_name} WHERE InputData_ID = ?", (input_data_id,))

            # Формируем SQL запрос
            fields = ['InputData_ID']
            values = [input_data_id]
            placeholders = ['?']

            for field_name, value in amendments_data.items():
                fields.append(field_name)
                values.append(value)
                placeholders.append('?')

            sql = f"INSERT INTO {table_name} ({', '.join(fields)}) VALUES ({', '.join(placeholders)})"
            cursor.execute(sql, values)
            conn.commit()

            logging.info(f"Поправки сохранены в {table_name} для ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка сохранения в {table_name}: {str(e)}")
            raise

    def get_all_amendments(self, input_data_id):
        """Получает все поправки из всех таблиц"""
        amendments = {}

        # Загружаем из всех таблиц
        tables = ['amendments', 'amendments2', 'amendments3', 'amendments4']
        suffixes = ['', '2', '3', '4']

        for table, suffix in zip(tables, suffixes):
            table_data = self.get_amendments_from_table(input_data_id, table)
            for key, value in table_data.items():
                amendments[key + suffix] = value

        return amendments

    def get_amendments_from_table(self, input_data_id, table_name):
        """Получает поправки из конкретной таблицы"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute(f"SELECT * FROM {table_name} WHERE InputData_ID = ?", (input_data_id,))
            result = cursor.fetchone()

            if result:
                # Получаем описание таблицы для имен колонок
                cursor.execute(f"SELECT * FROM {table_name} WHERE 1=0")
                columns = [column[0] for column in cursor.description]

                return {col: val for col, val in zip(columns, result) if col != 'InputData_ID'}
            return {}

        except Exception as e:
            logging.error(f"Ошибка получения данных из {table_name}: {str(e)}")
            return {}

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

    def get_amendments2(self, input_data_id):
        """Получает поправки из таблицы amendments2"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute("SELECT * FROM amendments2 WHERE InputData_ID = ?", (input_data_id,))
            result = cursor.fetchone()

            if result:
                return {
                    'amendVnkPpl2': result[0],
                    'amendVdpPpl2': result[1],
                    'amendGnkPpl2': result[2],
                    'amendVnkPzab2': result[3],
                    'amendVdpPzab2': result[4],
                    'amendGnkPzab2': result[5]
                }
            return {}

        except Exception as e:
            logging.error(f"Ошибка получения amendments2: {str(e)}")
            return {}

    def get_amendments3(self, input_data_id):
        """Получает поправки из таблицы amendments3"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute("SELECT * FROM amendments3 WHERE InputData_ID = ?", (input_data_id,))
            result = cursor.fetchone()

            if result:
                return {
                    'amendVnkPpl3': result[0],
                    'amendVdpPpl3': result[1],
                    'amendGnkPpl3': result[2],
                    'amendVnkPzab3': result[3],
                    'amendVdpPzab3': result[4],
                    'amendGnkPzab3': result[5]
                }
            return {}

        except Exception as e:
            logging.error(f"Ошибка получения amendments3: {str(e)}")
            return {}

    def get_amendments4(self, input_data_id):
        """Получает поправки из таблицы amendments4"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute("SELECT * FROM amendments4 WHERE InputData_ID = ?", (input_data_id,))
            result = cursor.fetchone()

            if result:
                return {
                    'amendVnkPpl4': result[0],
                    'amendVdpPpl4': result[1],
                    'amendGnkPpl4': result[2],
                    'amendVnkPzab4': result[3],
                    'amendVdpPzab4': result[4],
                    'amendGnkPzab4': result[5]
                }
            return {}

        except Exception as e:
            logging.error(f"Ошибка получения amendments4: {str(e)}")
            return {}

    def update_damping_table_headers(self, input_data_id):
        """Обновляет заголовки в dampingTable с именами пластов"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Получаем данные из InputData
            cursor.execute("SELECT Plast FROM InputData WHERE ID = ?", (input_data_id,))
            plast_result = cursor.fetchone()
            plast_value = plast_result[0] if plast_result else ""

            # Разделяем пласты по запятым
            plasts = [p.strip() for p in str(plast_value).split(',')] if plast_value else []

            # Создаем временную таблицу с правильными заголовками
            cursor.execute("SELECT * FROM dampingTable WHERE 1=0")
            columns = [column[0] for column in cursor.description]

            # Заменяем номера пластов на реальные имена
            new_columns = columns[:2]  # Дата, Длительность

            for i, plast in enumerate(plasts[:4]):
                plast_num = i + 1
                new_columns.extend([
                    f'Рпл на ВНК пласта {plast}, кгс/см2',
                    f'Рпл на ВДП пласта {plast}, кгс/см2'
                ])

            # Создаем временную таблицу
            temp_table_sql = "CREATE TABLE dampingTable_temp ("
            temp_table_sql += "[Дата] DATETIME, [Длительность, час] FLOAT, "

            for i, plast in enumerate(plasts[:4]):
                temp_table_sql += f"[Рпл на ВНК пласта {plast}, кгс/см2] FLOAT, "
                temp_table_sql += f"[Рпл на ВДП пласта {plast}, кгс/см2] FLOAT, "

            temp_table_sql = temp_table_sql.rstrip(', ') + ")"

            cursor.execute("DROP TABLE IF EXISTS dampingTable_temp")
            cursor.execute(temp_table_sql)

            # Копируем данные
            if plasts:
                insert_sql = "INSERT INTO dampingTable_temp SELECT * FROM dampingTable"
                cursor.execute(insert_sql)

            # Заменяем оригинальную таблицу
            cursor.execute("DROP TABLE dampingTable")
            cursor.execute("ALTER TABLE dampingTable_temp RENAME TO dampingTable")

            conn.commit()
            logging.info("Заголовки dampingTable обновлены с именами пластов")

        except Exception as e:
            logging.error(f"Ошибка обновления заголовков: {str(e)}")
            raise

    def create_and_fill_damping_table(self, input_data_id):
        """Создает и заполняет таблицу dampingTable с фиксированными длительностями"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Очищаем таблицу перед заполнением
            cursor.execute("DELETE FROM dampingTable")

            # Получаем данные из InputData
            cursor.execute("SELECT Plast FROM InputData WHERE ID = ?", (input_data_id,))
            plast_result = cursor.fetchone()
            plast_value = plast_result[0] if plast_result else ""

            # Разделяем пласты по запятым
            plasts = [p.strip() for p in str(plast_value).split(',')] if plast_value else []

            # Получаем поправки из всех таблиц
            amendments = self.get_amendments(input_data_id) or {}
            amendments2 = self.get_amendments2(input_data_id) or {}
            amendments3 = self.get_amendments3(input_data_id) or {}
            amendments4 = self.get_amendments4(input_data_id) or {}

            # Получаем данные из ModelVNK
            cursor.execute("""
                           SELECT Dat, PressureVnkModel
                           FROM ModelVNK
                           WHERE InputData_ID = ?
                             AND PressureVnkModel IS NOT NULL
                           ORDER BY Dat
                           """, (input_data_id,))
            model_data = cursor.fetchall()

            if not model_data:
                logging.warning("Нет данных в ModelVNK для заполнения dampingTable")
                return

            # Фиксированные длительности (в часах)
            durations = [0, 24, 48, 72, 96, 120, 144, 168, 192, 216, 240, 480, 600, 720]

            # Создаем данные для таблицы
            damping_data = []
            start_date = datetime.now().date()

            for i, duration in enumerate(durations):
                row_data = {
                    'Дата': start_date + timedelta(hours=duration),
                    'Длительность, час': duration
                }

                # Находим соответствующее давление из ModelVNK для этой длительности
                # Ищем ближайшее значение по времени
                target_time = timedelta(hours=duration)
                closest_pressure = None

                for dat, pressure in model_data:
                    if isinstance(dat, datetime):
                        time_diff = abs(
                            (dat - datetime.combine(start_date, time.min)).total_seconds() / 3600 - duration)
                        if closest_pressure is None or time_diff < closest_pressure[0]:
                            closest_pressure = (time_diff, pressure)

                if closest_pressure:
                    base_pressure = closest_pressure[1]
                else:
                    base_pressure = model_data[0][1]  # Первое значение по умолчанию

                # Базовое давление из ModelVNK для пласта 1
                base_pressure_vnk = base_pressure

                # Расчет для пласта 1
                vnk1 = base_pressure_vnk
                vdp1 = vnk1 - (amendments.get('amendVnkPpl', 0) - amendments.get('amendVdpPpl', 0))

                row_data['Рпл на ВНК пласта 1, кгс/см2'] = vnk1
                row_data['Рпл на ВДП пласта 1, кгс/см2'] = vdp1

                # Расчет для пласта 2 (если есть)
                if len(plasts) >= 2:
                    vnk2 = vnk1 + (amendments2.get('amendVnkPpl2', 0) - amendments.get('amendVnkPpl', 0))
                    vdp2 = vdp1 + (amendments2.get('amendVdpPpl2', 0) - amendments.get('amendVdpPpl', 0))

                    row_data['Рпл на ВНК пласта 2, кгс/см2'] = vnk2
                    row_data['Рпл на ВДП пласта 2, кгс/см2'] = vdp2

                # Расчет для пласта 3 (если есть)
                if len(plasts) >= 3:
                    vnk3 = vnk1 + (amendments3.get('amendVnkPpl3', 0) - amendments.get('amendVnkPpl', 0))
                    vdp3 = vdp1 + (amendments3.get('amendVdpPpl3', 0) - amendments.get('amendVdpPpl', 0))

                    row_data['Рпл на ВНК пласта 3, кгс/см2'] = vnk3
                    row_data['Рпл на ВДП пласта 3, кгс/см2'] = vdp3

                # Расчет для пласта 4 (если есть)
                if len(plasts) >= 4:
                    vnk4 = vnk1 + (amendments4.get('amendVnkPpl4', 0) - amendments.get('amendVnkPpl', 0))
                    vdp4 = vdp1 + (amendments4.get('amendVdpPpl4', 0) - amendments.get('amendVdpPpl', 0))

                    row_data['Рпл на ВНК пласта 4, кгс/см2'] = vnk4
                    row_data['Рпл на ВДП пласта 4, кгс/см2'] = vdp4

                damping_data.append(row_data)

            # Вставляем данные в таблицу
            for row in damping_data:
                sql = """
                      INSERT INTO dampingTable
                      (Дата, Длительность, час,
                       Рпл на ВНК пласта 1, кгс/см2, Рпл на ВДП пласта 1, кгс/см2,
                       Рпл на ВНК пласта 2, кгс/см2, Рпл на ВДП пласта 2, кгс/см2,
                       Рпл на ВНК пласта 3, кгс/см2, Рпл на ВДП пласта 3, кгс/см2,
                       Рпл на ВНК пласта 4, кгс/см2, Рпл на ВДП пласта 4, кгс/см2)
                      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) \
                      """

                params = (
                    row['Дата'],
                    row['Длительность, час'],
                    row.get('Рпл на ВНК пласта 1, кгс/см2'),
                    row.get('Рпл на ВДП пласта 1, кгс/см2'),
                    row.get('Рпл на ВНК пласта 2, кгс/см2'),
                    row.get('Рпл на ВДП пласта 2, кгс/см2'),
                    row.get('Рпл на ВНК пласта 3, кгс/см2'),
                    row.get('Рпл на ВДП пласта 3, кгс/см2'),
                    row.get('Рпл на ВНК пласта 4, кгс/см2'),
                    row.get('Рпл на ВДП пласта 4, кгс/см2')
                )

                cursor.execute(sql, params)

            conn.commit()
            logging.info(f"Таблица dampingTable заполнена для ID: {input_data_id}")

            # Проверяем результат
            cursor.execute("SELECT Дата, Длительность, час FROM dampingTable ORDER BY Длительность, час")
            results = cursor.fetchall()
            logging.info(f"Создано записей: {len(results)}")
            for date, duration in results:
                logging.info(f"Дата: {date}, Длительность: {duration} часов")

        except Exception as e:
            logging.error(f"Ошибка заполнения dampingTable: {str(e)}")
            raise

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
            conn = self.get_connection()
            cursor = conn.cursor()

            # Очищаем предыдущие расчеты
            cursor.execute("DELETE FROM calculatedPressure WHERE InputData_ID = ?", (input_data_id,))
            conn.commit()

            # Получаем необходимые данные
            vid_issledovaniya = self.get_vid_issledovaniya(input_data_id)
            pzab_vnk = None
            ppl_vnk = None

            # Логика расчета PzabVnk
            vid_upper = (vid_issledovaniya or '').upper()
            if vid_issledovaniya and any(x in vid_upper for x in ['КСД', 'АДД']):
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
            elif vid_issledovaniya and any(x in vid_upper for x in ['КСД', 'АДД']):
                ppl_vnk = self.get_last_model_ksd_pressure(input_data_id)
            elif vid_issledovaniya and 'ГРП' in vid_upper:
                initial_pressure = self.get_calculated_parameter(input_data_id, 'Начальное пластовое давление')
                ppl_vnk = initial_pressure
            else:
                ppl_vnk = pressure_last_point  # Используем переданное значение

            # Получаем поправки (только основные)
            amendments = self.get_amendments(input_data_id) or {}

            # Используем основные поправки из таблицы amendments
            amend_vnk_ppl = amendments.get('amendVnkPpl', 0) or 0
            amend_vdp_ppl = amendments.get('amendVdpPpl', 0) or 0
            amend_gnk_ppl = amendments.get('amendGnkPpl', 0) or 0
            amend_vnk_pzab = amendments.get('amendVnkPzab', 0) or 0
            amend_vdp_pzab = amendments.get('amendVdpPzab', 0) or 0
            amend_gnk_pzab = amendments.get('amendGnkPzab', 0) or 0

            # Рассчитываем параметры с проверкой на None
            calculated_data = {}

            if ppl_vnk is not None:
                calculated_data.update({
                    'PplVnk': ppl_vnk,
                    'PplGlubZam': ppl_vnk - amend_vnk_ppl,
                    'PplVdp': (ppl_vnk - amend_vnk_ppl) + amend_vdp_ppl,
                    'PplGnk': (ppl_vnk - amend_vnk_ppl) + amend_gnk_ppl
                })

            if pzab_vnk is not None:
                calculated_data.update({
                    'PzabVnk': pzab_vnk,
                    'PzabGlubZam': pzab_vnk - amend_vnk_pzab,
                    'PzabVdp': (pzab_vnk - amend_vnk_pzab) + amend_vdp_pzab,
                    'PzabGnk': (pzab_vnk - amend_vnk_pzab) + amend_gnk_pzab
                })

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

            # Очищаем предыдущие расчеты
            cursor.execute("DELETE FROM calculate WHERE InputData_ID = ?", (input_data_id,))

            # 1. P1_zab_vnk - значение из calculatedParameters
            p1_zab_vnk = self.get_calculated_parameter(input_data_id, 'P @ dt=0')

            # 2. P2_zab_vnk - последнее значение из PressureVNK
            p2_zab_vnk = self.get_last_pressure_vnk(input_data_id)

            # 3. Pday - разница давлений за сутки (исправленный синтаксис для Access)
            cursor.execute("""
                           SELECT MAX(PressureVnk) - MIN(PressureVnk)
                           FROM PressureVNK
                           WHERE InputData_ID = ?
                             AND Dat >= (SELECT MAX(Dat) FROM PressureVNK WHERE InputData_ID = ?) - 1
                           """, (input_data_id, input_data_id))
            result = cursor.fetchone()
            pday = result[0] if result else None

            # 4. P_pl_vnk
            p_pl_vnk = self.get_last_model_vnk_pressure(input_data_id)

            # 5. delta - разница с предыдущим исследованием
            delta = None
            # if p_pl_vnk:
            #     # Получаем текущую скважину
            #     cursor.execute("SELECT Skvazhina FROM InputData WHERE ID = ?", (input_data_id,))
            #     result = cursor.fetchone()
            #     current_skvazhina = result[0] if result else None
            #
            #     if current_skvazhina:
            #         # Ищем предыдущее исследование для этой скважины
            #         cursor.execute("""
            #                        SELECT TOP 1 p.PplVnk
            #                        FROM InputData i
            #                                 INNER JOIN prevData p ON i.ID = p.InputData_ID
            #                        WHERE i.Skvazhina = ?
            #                        ORDER BY i.Data_issledovaniya DESC
            #                        """, (current_skvazhina,))
            #
            #         result = cursor.fetchone()
            #         if result:
            #             prev_ppl_vnk = result[0]
            #             delta = p_pl_vnk - prev_ppl_vnk

            # 6. productivity
            delta_q = self.get_calculated_parameter(input_data_id, 'Delta Q')
            productivity = abs(delta_q / (p1_zab_vnk - p_pl_vnk)) if all([delta_q, p1_zab_vnk, p_pl_vnk]) else None

            # 7. Delta Q
            delta_q_value = delta_q

            # 8. Qoil
            cursor.execute("""
                           SELECT Obvodnennost, Vremya_issledovaniya
                           FROM InputData
                           WHERE ID = ?
                           """, (input_data_id,))
            result = cursor.fetchone()
            qoil = None
            if result and delta_q:
                obvodnennost, vremya = result
                qoil = delta_q * (1 - obvodnennost / 100) * vremya / 24 if obvodnennost and vremya else None

            # 9. Kh/Mu - гидропроводность
            kh_mu = None
            h = self.get_calculated_parameter(input_data_id, 'h')
            k = self.get_calculated_parameter(input_data_id, 'k')
            viscosity = self.get_calculated_parameter(input_data_id, 'вязкость')
            mu_gas = self.get_calculated_parameter(input_data_id, 'µгаза')

            fluid_type = self.get_calculated_parameter(input_data_id, 'Тип флюида')
            fluid_type = str(fluid_type or '').lower()

            if all([h, k]):
                if 'газ' in fluid_type and mu_gas:
                    kh_mu = (h * 100 * k / 1000) / mu_gas
                elif viscosity:
                    kh_mu = (h * 100 * k / 1000) / viscosity

            # 10. Rinv_Ppl1
            rinv_ppl1 = None
            phi = self.get_calculated_parameter(input_data_id, 'Phi')
            ct = self.get_calculated_parameter(input_data_id, 'Общая сжимаемость сt')

            if all([k, phi, vremya]) and result:
                vremya = result[1]  # Vremya_issledovaniya
                if 'газ' in fluid_type and mu_gas and ct:
                    rinv_ppl1 = 0.037 * math.sqrt((k * vremya) / (phi * mu_gas * ct))
                elif viscosity and ct:
                    rinv_ppl1 = 0.037 * math.sqrt((k * vremya) / (phi * viscosity * ct))

            # Вставляем данные в таблицу calculate
            cursor.execute("""
                           INSERT INTO calculate
                           (P1_zab_vnk, P2_zab_vnk, Pday, P_pl_vnk, delta, productivity,
                            Delta_Q, Qoil, Kh_Mu, Rinv_Ppl1, InputData_ID)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                           """, (p1_zab_vnk, p2_zab_vnk, pday, p_pl_vnk, delta, productivity,
                                 delta_q_value, qoil, kh_mu, rinv_ppl1, input_data_id))

            conn.commit()
            logging.info("Таблица calculate обновлена")

        except Exception as e:
            logging.error(f"Ошибка обновления calculate: {e}")
            raise

    def insert_prev_data(self, data_dict):
        """Вставляет данные в таблицу prevData с правильными типами"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute("DELETE FROM prevData WHERE InputData_ID = ?", (data_dict['InputData_ID'],))
            conn.commit()

            # Преобразуем данные согласно структуре таблицы
            sql = """
                  INSERT INTO prevData
                  (PplVnk, PzabVnk, DataRes, Water, Q, Kprod, Smeh, Heff, Kgidr, InputData_ID)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) \
                  """

            # Функция для извлечения числа из строки (для поля Smeh)
            def extract_number(value):
                if value is None:
                    return None
                if isinstance(value, (int, float)):
                    return float(value)
                if isinstance(value, str):
                    # Ищем число в строке (например, из "/-8.24" извлекаем -8.24)
                    import re
                    # Ищем числа с дробной частью и знаками
                    match = re.search(r'[-+]?\d*\.?\d+', value.replace(',', '.'))
                    if match:
                        try:
                            return float(match.group())
                        except ValueError:
                            return None
                    return None
                return None

            # Подготавливаем параметры с правильными типами
            params = (
                float(data_dict['PplVnk']) if data_dict['PplVnk'] is not None else None,
                float(data_dict['PzabVnk']) if data_dict['PzabVnk'] is not None else None,
                data_dict['DataRes'],  # Дата как есть (должна быть datetime)
                float(data_dict['Water']) if data_dict['Water'] is not None else None,
                float(data_dict['Q']) if data_dict['Q'] is not None else None,
                float(data_dict['Kprod']) if data_dict['Kprod'] is not None else None,
                extract_number(data_dict['Smeh']),  # Извлекаем число из строки
                float(data_dict['Heff']) if data_dict['Heff'] is not None else None,
                float(data_dict['Kgidr']) if data_dict['Kgidr'] is not None else None,
                int(data_dict['InputData_ID'])
            )

            # Отладочная информация
            logging.info("Параметры для вставки в prevData:")
            for i, param in enumerate(params):
                logging.info(f"  Параметр {i}: {param} (type: {type(param).__name__})")

            cursor.execute(sql, params)
            conn.commit()
            logging.info(f"Данные предыдущего исследования сохранены для ID: {data_dict['InputData_ID']}")

        except Exception as e:
            logging.error(f"Ошибка вставки в prevData: {str(e)}")
            raise

    def import_previous_research_data(self, input_data_id, field_name, well_name):
        """Импортирует данные предыдущего исследования из Excel файла"""
        try:
            from GUI_Claudi_ITC import table_prev_path
            logging.info(f"Начинаем импорт для field: {field_name}, well: {well_name}")

            # Формируем путь к файлу
            field_name_clean = field_name.replace(" ", "_").capitalize()
            previous_data_file = f'Итоговая таблица_{field_name_clean}.xlsx'
            previous_data_path = table_prev_path(previous_data_file)

            logging.info(f"Ищем файл: {previous_data_path}")

            if not os.path.exists(previous_data_path):
                logging.warning(f"Файл предыдущих данных не найден: {previous_data_path}")
                return False

            # Читаем Excel файл
            logging.info("Читаем Excel файл...")
            final_table_df = pd.read_excel(previous_data_path, skiprows=11)

            well_num = well_name.split()[0] if well_name else ''
            logging.info(f"Ищем данные для скважины: {well_num}")

            # Проверяем наличие колонки 'Скважина'
            if 'Скважина' not in final_table_df.columns:
                logging.error("Колонка 'Скважина' не найдена в файле")
                return False

            # Фильтруем данные по номеру скважины
            final_table_df['Скважина'] = final_table_df['Скважина'].astype(str).str.strip()
            final_table_df = final_table_df.dropna(subset=['Скважина'])
            filtered_data = final_table_df[final_table_df['Скважина'] == well_num]

            logging.info(f"Найдено записей для скважины {well_num}: {len(filtered_data)}")

            if filtered_data.empty:
                logging.warning(f"Не найдено данных для скважины {well_num}")
                return False

            # Обработка данных из файла
            pd.set_option('mode.use_inf_as_na', True)

            # Проверяем наличие колонки 'Дата испытания'
            if 'Дата испытания' not in filtered_data.columns:
                logging.error("Колонка 'Дата испытания' не найдена")
                return False

            filtered_data.loc[:, 'Дата испытания'] = filtered_data['Дата испытания'].apply(
                lambda x: datetime.strptime(x, "%d.%m.%Y") if isinstance(x, str) else x
            )

            latest_entry = filtered_data.loc[filtered_data['Дата испытания'].idxmax()]

            # Функция для безопасного преобразования значений
            def safe_convert(value, default=None):
                if pd.isna(value) or value is None:
                    return default
                try:
                    if isinstance(value, (int, float)):
                        return float(value)
                    elif isinstance(value, str):
                        # Пробуем преобразовать строку в число
                        return float(value.replace(',', '.'))
                    else:
                        return default
                except (ValueError, TypeError):
                    return default

            # Сохраняем данные в таблицу prevData с правильными типами
            prev_data = {
                'PplVnk': safe_convert(latest_entry.get('Рпл  на ВНК, кгс/см2')),
                'PzabVnk': safe_convert(latest_entry.get('Рзаб  на ВНК, кгс/см2')),
                'DataRes': latest_entry.get('Дата испытания'),
                'Water': safe_convert(latest_entry.get('% воды')),
                'Q': safe_convert(latest_entry.get('Qж/Qг, м3/сут   ')),
                'Kprod': safe_convert(latest_entry.get('Кпрод. м3/сут*кгс/см2')),
                'Smeh': latest_entry.get('Скин-фактор механич./интегр.'),  # Теперь число!
                'Heff': safe_convert(latest_entry.get('Нэф., м. ')),
                'Kgidr': safe_convert(latest_entry.get('Кгидр., Д*см/сПз')),
                'InputData_ID': input_data_id
            }

            logging.info(f"Данные для сохранения: {prev_data}")

            # Сохраняем в базу данных
            self.insert_prev_data(prev_data)
            logging.info("Данные предыдущего исследования сохранены в базу")

            return True

        except Exception as e:
            logging.error(f"Ошибка импорта данных предыдущего исследования: {str(e)}", exc_info=True)
            return False

    # def convert_date_for_access(date_value):
    #     """Конвертирует дату в формат, понятный Access"""
    #     if isinstance(date_value, datetime):
    #         return date_value
    #     elif isinstance(date_value, str):
    #         try:
    #             return datetime.strptime(date_value, "%d.%m.%Y")
    #         except ValueError:
    #             return None
    #     elif isinstance(date_value, pd.Timestamp):
    #         return date_value.to_pydatetime()
    #     return None

    def get_previous_research_data(self, input_data_id):
        """Получает данные предыдущего исследования для отчета"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute("""
                           SELECT PplVnk,
                                  PzabVnk,
                                  DataRes,
                                  Water,
                                  Q,
                                  Kprod,
                                  Smeh,
                                  Heff,
                                  Kgidr
                           FROM prevData
                           WHERE InputData_ID = ?
                           """, (input_data_id,))

            result = cursor.fetchone()
            if result:
                return {
                    'Рпл  на ВНК, кгс/см2': result[0],
                    'Рзаб  на ВНК, кгс/см2': result[1],
                    'Дата испытания': result[2].strftime("%d.%m.%Y") if result[2] else '',
                    '% воды': result[3],
                    'Qж/Qг, м3/сут': result[4],
                    'Кпрод. м3/сут*кгс/см2': result[5],
                    'Скин-фактор механич./интегр.': result[6],
                    'Нэф., м.': result[7],
                    'Кгидр., Д*см/сПз': result[8]
                }
            return {}

        except Exception as e:
            logging.error(f"Ошибка получения данных предыдущего исследования: {str(e)}")
            return {}

    def get_field_name(self, input_data_id):
        """Получает название месторождения по ID записи"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT Localoredenie FROM InputData WHERE ID = ?", (input_data_id,))
            result = cursor.fetchone()
            return result[0] if result else None
        except Exception as e:
            logging.error(f"Ошибка получения названия месторождения: {str(e)}")
            return None

    def get_well_name(self, input_data_id):
        """Получает название скважины по ID записи"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT Skvazhina FROM InputData WHERE ID = ?", (input_data_id,))
            result = cursor.fetchone()
            return result[0] if result else None
        except Exception as e:
            logging.error(f"Ошибка получения названия скважины: {str(e)}")
            return None

    def check_prevdata_structure(self):
        """Проверяет реальную структуру таблицы prevData"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Получаем информацию о колонках
            cursor.execute("SELECT TOP 1 * FROM prevData")
            description = cursor.description

            print("=== СТРУКТУРА TAБЛИЦЫ prevData ===")
            for column in description:
                col_name = column[0]
                col_type = column[1]  # Код типа данных
                print(f"{col_name}: тип кода {col_type}")

        except Exception as e:
            print(f"Ошибка проверки структуры: {e}")

    def clear_related_data(self, input_data_id):
        """Очищает все связанные данные для указанного InputData_ID"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Список таблиц для очистки
            tables = [
                'success',
                'researchClass',
                'pressureLastPoint',
                'estimatedTime',
                'density',
                'amendments',
                'prevData'  # Добавляем prevData
            ]

            for table in tables:
                try:
                    cursor.execute(f"DELETE FROM {table} WHERE InputData_ID = ?", (input_data_id,))
                    logging.info(f"Очищена таблица {table} для ID: {input_data_id}")
                except Exception as e:
                    logging.warning(f"Не удалось очистить {table}: {str(e)}")

            conn.commit()

        except Exception as e:
            logging.error(f"Ошибка очистки связанных данных: {str(e)}")
            raise

    def create_damping_table_structure(self):
        """Создает структуру таблицы dampingTable с поддержкой до 4 пластов"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Проверяем, существует ли таблица
            cursor.execute("SELECT COUNT(*) FROM MSysObjects WHERE Name='dampingTable' AND Type=1")
            exists = cursor.fetchone()[0] > 0

            if not exists:
                # Создаем таблицу
                sql = """
                    CREATE TABLE dampingTable (
                        [Дата] DATETIME,
                        [Длительность, час] FLOAT,
                        [Рпл на ВНК пласта 1, кгс/см2] FLOAT,
                        [Рпл на ВДП пласта 1, кгс/см2] FLOAT,
                        [Рпл на ВНК пласта 2, кгс/см2] FLOAT,
                        [Рпл на ВДП пласта 2, кгс/см2] FLOAT,
                        [Рпл на ВНК пласта 3, кгс/см2] FLOAT,
                        [Рпл на ВДП пласта 3, кгс/см2] FLOAT,
                        [Рпл на ВНК пласта 4, кгс/см2] FLOAT,
                        [Рпл на ВДП пласта 4, кгс/см2] FLOAT
                    )      
                """
                cursor.execute(sql)
                conn.commit()
                logging.info("Таблица dampingTable создана")

        except Exception as e:
            logging.error(f"Ошибка создания dampingTable: {str(e)}")
            raise


# db = AccessDatabase()
# db.add_sort_order_to_model_vnk()
# db.get_last_pressure_vnk('61')
# db.check_prevdata_structure()
# db.import_previous_research_data('72', 'Чаяндинское', '1057G')
