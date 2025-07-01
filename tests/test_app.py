#!/usr/bin/env python3
"""
Скрипт для тестирования локального приложения перед деплоем.
Проверяет наличие всех необходимых файлов, директорий и зависимостей.
"""

import os
import sys
import importlib
from pathlib import Path
import pkg_resources

# Список необходимых файлов и директорий
REQUIRED_FILES = [
    'app.py',
    'wsgi.py',
    'formatting_checker.py',
    'comment_utils.py',
    'requirements.txt',
]

REQUIRED_DIRS = [
    'templates',
    'static',
    'uploads',
]

# Список необходимых модулей
REQUIRED_MODULES = [
    'flask',
    'docx',
    'lxml',
]

def test_dependencies():
    """Проверяет, что все необходимые зависимости установлены"""
    missing_packages = []
    
    try:
        installed_packages = {pkg.key: pkg.version for pkg in pkg_resources.working_set}
        
        for module_name in REQUIRED_MODULES:
            try:
                # Пробуем импортировать модуль
                importlib.import_module(module_name)
                version = installed_packages.get(module_name, "Установлен, версия не определена")
                print(f"✓ {module_name} - {version}")
            except ImportError:
                missing_packages.append(module_name)
                print(f"✗ {module_name} - не установлен")
    except Exception as e:
        print(f"Ошибка при проверке зависимостей: {e}")
        missing_packages = REQUIRED_MODULES  # Предполагаем, что не можем проверить
    
    return missing_packages

def test_file_structure():
    """Проверяет наличие всех необходимых файлов и директорий"""
    missing_files = []
    missing_dirs = []
    
    # Проверка файлов
    for file_name in REQUIRED_FILES:
        if not os.path.isfile(file_name):
            missing_files.append(file_name)
            print(f"✗ {file_name} - файл не найден")
        else:
            print(f"✓ {file_name} - файл найден, {os.path.getsize(file_name)/1024:.2f} КБ")
    
    # Проверка директорий
    for dir_name in REQUIRED_DIRS:
        if not os.path.isdir(dir_name):
            missing_dirs.append(dir_name)
            print(f"✗ {dir_name} - директория не найдена")
        else:
            file_count = sum(1 for _ in Path(dir_name).rglob('*') if _.is_file())
            print(f"✓ {dir_name} - директория найдена, {file_count} файлов")
    
    return missing_files, missing_dirs

def run_tests():
    """Запускает все тесты и проверки"""
    print("=" * 50)
    print("Тестирование приложения перед деплоем")
    print("=" * 50)
    
    print("\n1. Проверка структуры файлов и директорий:")
    missing_files, missing_dirs = test_file_structure()
    
    print("\n2. Проверка зависимостей:")
    missing_packages = test_dependencies()
    
    # Выводим результаты
    print("\n" + "=" * 50)
    print("Результаты тестирования:")
    
    if not missing_files and not missing_dirs and not missing_packages:
        print("\n✓ Все проверки пройдены успешно! Приложение готово к деплою.")
    else:
        print("\n✗ Обнаружены проблемы, которые нужно исправить перед деплоем:")
        
        if missing_files:
            print(f"\n  Отсутствующие файлы ({len(missing_files)}):")
            for file in missing_files:
                print(f"  - {file}")
        
        if missing_dirs:
            print(f"\n  Отсутствующие директории ({len(missing_dirs)}):")
            for directory in missing_dirs:
                print(f"  - {directory}")
                if directory == 'uploads':
                    print("    Директория 'uploads' может быть создана автоматически.")
        
        if missing_packages:
            print(f"\n  Отсутствующие пакеты ({len(missing_packages)}):")
            for package in missing_packages:
                print(f"  - {package}")
            print("\n  Установите их с помощью команды:")
            print(f"  pip install {' '.join(missing_packages)}")
    
    print("\n" + "=" * 50)

if __name__ == "__main__":
    run_tests() 