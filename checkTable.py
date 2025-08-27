from db_access import AccessDatabase


# Проверим все таблицы
db = AccessDatabase()
db.check_table_structure("success")
db.check_table_structure("researchClass")
db.check_table_structure("pressureLastPoint")
db.check_table_structure("estimatedTime")
db.check_table_structure("density")