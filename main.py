 #!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Главный файл запуска приложения анализа бизнес-процессов из Excel
"""
import sys
import os

# Добавляем текущую директорию в путь для корректного импорта модулей
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

def main():
    """Главная функция запуска приложения"""
    try:
        from ui_main import main as ui_main
        ui_main()
    except ImportError as e:
        print(f"Ошибка импорта модулей: {e}")
        print("Убедитесь, что установлены все необходимые зависимости:")
        print("pip install PyQt5 pandas openpyxl")
        sys.exit(1)
    except Exception as e:
        print(f"Ошибка запуска приложения: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()