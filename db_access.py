import pyodbc
import os
from utils import logger, logging
import pandas as pd
import datetime
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
        """
        Возвращает последнюю запись из main_data с обработкой ошибок.

        Returns:
            pandas.DataFrame: DataFrame с последней записью или пустой DataFrame при ошибке
        """
        try:
            conn = self.get_connection()

            # Более надежный запрос
            query = """
                    SELECT TOP 1 *
                    FROM main_data
                    WHERE id IS NOT NULL
                    ORDER BY id DESC, created_date DESC \
                    """

            df = pd.read_sql(query, conn)

            if not df.empty:
                logging.info(f"Получена последняя запись с ID: {df.iloc[0]['id']}")
            else:
                logging.warning("Таблица main_data пуста")

            return df

        except pyodbc.Error as e:
            logging.error(f"Ошибка базы данных при получении последней записи: {str(e)}")
            return pd.DataFrame()

        except Exception as e:
            logging.error(f"Неожиданная ошибка при получении последней записи: {str(e)}")
            return pd.DataFrame()

    def clear_data(self):
        """Очищает все данные из таблиц базы данных."""
        conn = self.get_connection()
        cursor = conn.cursor()
        tables = [
            'dampingTable', 'calculatedPressure', 'calculatedParameters',
            'research_params', 'amendments', 'Parameters', 'InputData',
            'researchClass', 'success', 'density', 'estimatedTime',
            'ModelVNK', 'PressureVNK'
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
        conn = self.get_connection()
        cursor = conn.cursor()

        sql = """
              INSERT INTO InputData (Company, Localoredenie, Skvazhina, VNK, Data_issledovaniya, Plast, \
                                     Interval_perforacii, Tip_pribora, Glubina_ustanovki_pribora, Interpretator, \
                                     Data_interpretacii, Vremya_issledovaniya, Obvodnennost, Nalicie_paktera, Data_GRP, \
                                     Vid_issledovaniya) \
              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) \
              """

        params = (
            data_dict.get('Company'), data_dict.get('Localoredenie'), data_dict.get('Skvazhina'),
            data_dict.get('VNK'), data_dict.get('Data_issledovaniya'), data_dict.get('Plast'),
            data_dict.get('Interval_perforacii'), data_dict.get('Tip_pribora'),
            data_dict.get('Glubina_ustanovki_pribora'), data_dict.get('Interpretator'),
            data_dict.get('Data_interpretacii'), data_dict.get('Vremya_issledovaniya'),
            data_dict.get('Obvodnennost'), data_dict.get('Nalicie_paktera'),
            data_dict.get('Data_GRP'), data_dict.get('Vid_issledovaniya')
        )

        try:
            cursor.execute(sql, params)
            conn.commit()
            logger.info("Данные успешно добавлены в InputData.")
            return cursor.execute("SELECT @@IDENTITY").fetchval()
        except pyodbc.Error as e:
            conn.rollback()
            logger.error(f"Ошибка при вставке в InputData: {e}")
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

    def insert_research_class(self, main_data_id, class_value):
        """Вставляет класс исследования в researchClass."""
        conn = self.get_connection()
        cursor = conn.cursor()
        sql = "INSERT INTO researchClass (MainDataID, ClassValue) VALUES (?, ?)"
        try:
            cursor.execute(sql, (main_data_id, float(class_value)))
            conn.commit()
            logger.info("Класс исследования успешно добавлен в researchClass.")
        except pyodbc.Error as e:
            conn.rollback()
            logger.error(f"Ошибка при вставке в researchClass: {e}")
            raise

    def insert_success(self, main_data_id, success_value):
        """Вставляет успешность в success."""
        conn = self.get_connection()
        cursor = conn.cursor()
        sql = "INSERT INTO success (MainDataID, SuccessValue) VALUES (?, ?)"
        try:
            cursor.execute(sql, (main_data_id, float(success_value)))
            conn.commit()
            logger.info("Успешность исследования успешно добавлена в success.")
        except pyodbc.Error as e:
            conn.rollback()
            logger.error(f"Ошибка при вставке в success: {e}")
            raise

    def insert_density(self, main_data_id, density_zab, density_pl):
        """Вставляет плотность в density."""
        conn = self.get_connection()
        cursor = conn.cursor()
        sql = "INSERT INTO density (MainDataID, DensityZab, DensityPl) VALUES (?, ?, ?)"
        try:
            cursor.execute(sql, (main_data_id, float(density_zab), float(density_pl)))
            conn.commit()
            logger.info("Плотность успешно добавлена в density.")
        except pyodbc.Error as e:
            conn.rollback()
            logger.error(f"Ошибка при вставке в density: {e}")
            raise

    def insert_estimated_time(self, main_data_id, time_value):
        """Вставляет расчетное время в estimatedTime."""
        conn = self.get_connection()
        cursor = conn.cursor()
        sql = "INSERT INTO estimatedTime (MainDataID, TimeValue) VALUES (?, ?)"
        try:
            cursor.execute(sql, (main_data_id, float(time_value)))
            conn.commit()
            logger.info("Расчетное время успешно добавлено в estimatedTime.")
        except pyodbc.Error as e:
            conn.rollback()
            logger.error(f"Ошибка при вставке в estimatedTime: {e}")
            raise

    def insert_amendments(self, main_data_id, amendments_dict):
        """Вставляет поправки в таблицу amendments"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Маппинг русских названий на английские (для consistency)
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

            for correction_type_ru, value in amendments_dict.items():
                if value is not None:
                    # Конвертируем русское название в английское
                    correction_type_en = correction_mapping.get(correction_type_ru, correction_type_ru)

                    cursor.execute(
                        "INSERT INTO amendments (main_data_id, correction_type, value, created_date) VALUES (?, ?, ?, ?)",
                        (main_data_id, correction_type_en, value, datetime.now())
                    )

            conn.commit()
            logging.info(f"Поправки сохранены для записи ID: {main_data_id}")

        except Exception as e:
            logging.error(f"Ошибка вставки поправок: {str(e)}")
            raise


    def get_parameters(self, main_data_id):
        """Получает параметры исследования для указанной записи"""
        try:
            conn = self.get_connection()
            query = "SELECT * FROM research_params WHERE main_data_id = ?"
            df = pd.read_sql(query, conn, params=[main_data_id])
            return df
        except Exception as e:
            logging.error(f"Ошибка получения параметров: {str(e)}")
            return pd.DataFrame()

    def get_research_params(self, main_data_id):
        """Алиас для get_parameters (для совместимости)"""
        return self.get_parameters(main_data_id)

    def get_amendments(self, main_data_id):
        """Получает поправки для указанной записи"""
        try:
            conn = self.get_connection()
            query = "SELECT * FROM amendments WHERE main_data_id = ? ORDER BY correction_type"
            df = pd.read_sql(query, conn, params=[main_data_id])
            return df
        except Exception as e:
            logging.error(f"Ошибка получения поправок: {str(e)}")
            return pd.DataFrame()

    def insert_research_params(self, main_data_id, section, params):
        """Вставляет параметры исследования (исправленная версия)"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            for param_name, param_value in params.items():
                if param_value is not None:
                    cursor.execute(
                        "INSERT INTO research_params (main_data_id, section, param_name, param_value, created_date) VALUES (?, ?, ?, ?, ?)",
                        (main_data_id, section, param_name, str(param_value), datetime.now())
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
        """Сохраняет поправки из полей ввода в базу данных"""
        try:
            # Получаем ID основной записи
            main_data_id = getattr(self, 'current_main_data_id', None)
            if main_data_id is None:
                last_record = self.get_last_record()
                if not last_record.empty:
                    main_data_id = last_record.iloc[0]['id']
                else:
                    messagebox.showerror("Ошибка", "Нет основной записи для привязки поправок")
                    return

            # Собираем поправки из полей ввода с русскими названиями
            amendments_dict = {}

            # Поправки Рпл (давление пластовое)
            amendments_dict['поправка с насоса на ВНК Рпл'] = self.get_float_value(self.ppl_entries[0].get())
            amendments_dict['поправка с насоса на ВДП Рпл'] = self.get_float_value(self.ppl_entries[1].get())
            amendments_dict['поправка с насоса на ГНК Рпл'] = self.get_float_value(self.ppl_entries[2].get())

            # Поправки Рзаб (давление забойное)
            amendments_dict['поправка с насоса на ВНК Рзаб'] = self.get_float_value(self.pzab_entries[0].get())
            amendments_dict['поправка с насоса на ВДП Рзаб'] = self.get_float_value(self.pzab_entries[1].get())
            amendments_dict['поправка с насоса на ГНК Рзаб'] = self.get_float_value(self.pzab_entries[2].get())

            # Поправки Рпл_2 (вторая модель)
            amendments_dict['поправка с насоса на ВНК Рпл'] = self.get_float_value(self.ppl2_entries[0].get())
            amendments_dict['поправка с насоса на ВДП Рпл'] = self.get_float_value(self.ppl2_entries[1].get())
            amendments_dict['поправка с насоса на ГНК Рпл'] = self.get_float_value(self.ppl2_entries[2].get())

            # Поправки Рзаб_2 (вторая модель)
            amendments_dict['поправка с насоса на ВНК Рзаб'] = self.get_float_value(self.pzab2_entries[0].get())
            amendments_dict['поправка с насоса на ВДП Рзаб'] = self.get_float_value(self.pzab2_entries[1].get())
            amendments_dict['поправка с насоса на ГНК Рзаб'] = self.get_float_value(self.pzab2_entries[2].get())

            # Дополнительные поправки
            amendments_dict['ВНК_Рпл_3'] = self.get_float_value(self.vnkp_pl3_entry.get())
            amendments_dict['ВНК_Рпл_4'] = self.get_float_value(self.vnkp_pl4_entry.get())

            # Фильтруем None значения
            amendments_dict = {k: v for k, v in amendments_dict.items() if v is not None}

            # Сохраняем поправки
            self.insert_amendments(main_data_id, amendments_dict)
            messagebox.showinfo("Успех", "Поправки успешно сохранены в таблицу amendments")

        except Exception as e:
            logging.error(f"Ошибка сохранения поправок: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка при сохранении поправок: {str(e)}")

    def get_float_value(self, value):
        """Конвертирует значение в float, возвращает None если пустое или не число"""
        if not value:
            return None
        try:
            return float(value.replace(',', '.'))
        except ValueError:
            return None