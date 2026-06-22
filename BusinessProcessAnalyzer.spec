# -*- mode: python ; coding: utf-8 -*-
# Спецификация для сборки BusinessProcessAnalyzer в .exe

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['f:\\cursr\\kursovaya'],
    binaries=[],
    datas=[],
    hiddenimports=[
        # PyQt5 модули
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'PyQt5.QtPrintSupport',
        'PyQt5.sip',
        # Pandas и зависимости
        'pandas',
        'pandas._libs',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.skiplist',
        'pandas.core.frame',
        'pandas.io.excel',
        'pandas.io.excel._openpyxl',
        # Openpyxl для чтения Excel
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.workbook',
        # Numpy (зависимость pandas)
        'numpy',
        'numpy.core._methods',
        'numpy.lib.format',
        # Модули приложения
        'excel_parser',
        'sequence_analyzer',
        'statistics_analyzer',
        'outlier_analyzer',
        'process_insights',
        'report_exporter',
        'ui_main',
        # Стандартные модули
        'json',
        're',
        'datetime',
        'collections',
        'statistics',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Исключаем ненужные модули для уменьшения размера
        'matplotlib',
        'scipy',
        'tkinter',
        'test',
        'unittest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='BusinessProcessAnalyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Без консольного окна (GUI приложение)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Можно добавить путь к .ico файлу
)
