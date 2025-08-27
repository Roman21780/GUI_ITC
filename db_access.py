import pyodbc
import os
from utils import logger, logging
import pandas as pd
from datetime import datetime
from tkinter import messagebox

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
            df = pd.read_sql(query, conn)
            return df
        except Exception as e:
            logging.error(f"Ошибка получения последней записи: {str(e)}")
            return pd.DataFrame()

    def clear_data(self):
        """Очищает все данные из таблиц базы данных."""
        conn = self.get_connection()
        cursor = conn.cursor()
        tables = [
            'dampingTable', 'calculatedPressure', 'calculatedParameters',
            'amendments', 'Parameters', 'InputData',
            'researchClass', 'success', 'density', 'estimatedTime',
            'ModelVNK', 'PressureVNK', 'pressureLastPoint'
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

    def insert_calculated_parameters(self, main_data_id, params_dict):
        """Вставляет параметры в calculatedParameters."""
        conn = self.get_connection()
        cursor = conn.cursor()
        sql = "INSERT INTO calculatedParameters (MainDataID, ParamName, ParamValue) VALUES (?, ?, ?)"
        try:
            for param_name, param_value in params_dict.items():
                cursor.execute(sql, (main_data_id, param_name, str(param_value)))
            conn.commit()
            logger.info("Параметры успешно добавлены в calculatedParameters.")
        except pyodbc.Error as e:
            conn.rollback()
            logger.error(f"Ошибка при вставке в calculatedParameters: {e}")
            raise

    def insert_model_vnk(self, main_data_id, model_data):
        """Вставляет данные модели в ModelVNK."""
        conn = self.get_connection()
        cursor = conn.cursor()
        sql = "INSERT INTO ModelVNK (MainDataID, ModelData) VALUES (?, ?)"
        try:
            # Преобразуем данные модели в строку JSON для хранения
            import json
            model_json = json.dumps(model_data, ensure_ascii=False)
            cursor.execute(sql, (main_data_id, model_json))
            conn.commit()
            logger.info("Данные модели успешно добавлены в ModelVNK.")
        except pyodbc.Error as e:
            conn.rollback()
            logger.error(f"Ошибка при вставке в ModelVNK: {e}")
            raise

    def insert_pressure_vnk(self, main_data_id, pressure_data):
        """Вставляет данные давления в PressureVNK."""
        conn = self.get_connection()
        cursor = conn.cursor()
        sql = "INSERT INTO PressureVNK (MainDataID, PressureData) VALUES (?, ?)"
        try:
            import json
            pressure_json = json.dumps(pressure_data, ensure_ascii=False)
            cursor.execute(sql, (main_data_id, pressure_json))
            conn.commit()
            logger.info("Данные давления успешно добавлены в PressureVNK.")
        except pyodbc.Error as e:
            conn.rollback()
            logger.error(f"Ошибка при вставке в PressureVNK: {e}")
            raise

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

    def insert_amendments_from_gui(self, input_data_id, amendments_data):
        """Вставляет поправки из GUI окошек в таблицу amendments"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Маппинг русских названий поправок на английские (для базы)
            correction_mapping = {
                'поправка с насоса на ВНК Рпл': 'pump_to_vnk_ppl',
                'поправка с насоса на ВДП Рпл': 'pump_to_vdp_ppl',
                'поправка с насоса на ГНК Рпл': 'pump_to_gnk_ppl',
                'поправка с насоса на ВНК Рзаб': 'pump_to_vnk_pzab',
                'поправка с насоса на ВДП Рзаб': 'pump_to_vdp_pzab',
                'поправка с насоса на ГНК Рзаб': 'pump_to_gnk_pzab',
                'ВНК_Рпл_3': 'vnk_ppl_3',
                'ВНК_Рпл_4': 'vnk_ppl_4'
            }

            for correction_type_ru, value in amendments_data.items():
                if value is not None:
                    # Конвертируем русское название в английское
                    correction_type_en = correction_mapping.get(correction_type_ru, correction_type_ru)

                    cursor.execute(
                        "INSERT INTO amendments (InputData_ID, Correction_type, Value, created_date) VALUES (?, ?, ?, ?)",
                        (input_data_id, correction_type_en, float(value), datetime.now())
                    )

            conn.commit()
            logging.info(f"Поправки из GUI сохранены для записи ID: {input_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки поправок из GUI: {str(e)}")
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

    def get_all_amendments(self, input_data_id):
        """Получает все поправки для указанной записи с русскими названиями"""
        try:
            conn = self.get_connection()
            query = "SELECT * FROM amendments WHERE InputData_ID = ? ORDER BY ID"
            df = pd.read_sql(query, conn, params=[input_data_id])

            # Обратный маппинг для отображения русских названий
            reverse_mapping = {
                'pump_to_vnk_ppl': 'поправка с насоса на ВНК Рпл',
                'pump_to_vdp_ppl': 'поправка с насоса на ВДП Рпл',
                'pump_to_gnk_ppl': 'поправка с насоса на ГНК Рпл',
                'pump_to_vnk_pzab': 'поправка с насоса на ВНК Рзаб',
                'pump_to_vdp_pzab': 'поправка с насоса на ВДП Рзаб',
                'pump_to_gnk_pzab': 'поправка с насоса на ГНК Рзаб',
                'vnk_ppl_3': 'ВНК_Рпл_3',
                'vnk_ppl_4': 'ВНК_Рпл_4'
            }

            if not df.empty:
                df['Correction_type_ru'] = df['Correction_type'].map(
                    lambda x: reverse_mapping.get(x, x)
                )

            return df

        except Exception as e:
            logging.error(f"Ошибка получения поправок: {str(e)}")
            return pd.DataFrame()

    def insert_research_params(self, input_data_id, section, params):
        """Вставляет параметры исследования в calculatedParameters"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            for param_name, param_value in params.items():
                if param_value is not None:
                    cursor.execute(
                        "INSERT INTO calculatedParameters (InputData_ID, Section, Param_name, Param_value, created_date) VALUES (?, ?, ?, ?, ?)",
                        (input_data_id, section, param_name, str(param_value), datetime.now())
                    )

            conn.commit()
            logging.info(f"Параметры исследования сохранены для секции {section}")

        except Exception as e:
            logging.error(f"Ошибка вставки параметров исследования: {str(e)}")
            raise

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
            self.insert_amendments_from_gui(main_data_id, amendments_data)
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