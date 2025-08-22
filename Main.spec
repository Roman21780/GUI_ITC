from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

a = Analysis(
    ['GUI_Claudi.py'],   # Главный скрипт программы
    pathex=['.'],        # Путь к исходным файлам
    binaries=[('C:\\Program Files\\Python312\\python312.dll', '.')],
    datas=[
        ('Report.xlsx', '.'),
        ('GRP.docx', '.'),
        ('KVD_ID.docx', '.'),
        ('KVD_For_Killing.docx', '.'),
        ('КВД для глушения_prev.docx', '.'),
        ('KVD_Zapolyarka.docx', '.'),
        ('KVD_NNG.docx', '.'),
        ('KVD_Khantos.docx', '.'),
        ('KPD_ID.docx', '.'),
        ('KPD.docx', '.'),
        ('KSD.docx', '.'),
        ('KVD_Orenburg.docx', '.'),
        ('KVD_Orenburg_gas.docx', '.'),
        ('KVD_Orenburg2.docx', '.'),
        ('Итоговая таблица_Чаяндинское.xlsx', '.'),
        ('Итоговая таблица_Песцовое.xlsx', '.'),
        ('Итоговая таблица_Оренбургское.xlsx', '.'),
        ('Итоговая таблица_Капитоновское.xlsx', '.'),
        ('Итоговая таблица_Царичанское+филатовское.xlsx', '.'),
        ('Итоговая таблица_Западно-Таркосалинское.xlsx', '.'),
        ('Helper.xlsm', '.'),
        ('text_templates.json', '.'),
    ],
    hiddenimports=['win32timezone', 'pandas', 'openpyxl', 'docx'],     # Если нужны скрытые импорты, добавьте их здесь
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,    # Файлы не будут упакованы в один исполняемый файл
    name='GDIS_ITC_v2',    # Имя исполняемого файла
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False    # Убирает консольное окно (для GUI-приложений)
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,    # Все дополнительные файлы будут включены в папку
    strip=False,
    upx=True,
    upx_exclude=[],
    name='GDIS_welltest_v2'    # Имя папки с результатами сборки
)
