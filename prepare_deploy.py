#!/usr/bin/env python3
"""
Скрипт для подготовки приложения к деплою.
"""

import os
import sys
import shutil
import zipfile
from pathlib import Path
import time

# Список файлов и директорий для включения в архив
FILES_TO_INCLUDE = [
    'app.py',
    'wsgi.py',
    'formatting_checker.py',
    'formatting_utils.py',
    'comment_utils.py',
    'requirements.txt',
    'README.md',
    'templates',
    'static',
]

def create_deploy_archive(archive_name="web_app_deploy.zip"):
    """Создает архив для деплоя на хостинг"""
    
    print(f"Подготовка архива для деплоя: {archive_name}")
    
    # Проверяем наличие основных файлов
    for file in FILES_TO_INCLUDE:
        if not os.path.exists(file):
            if file == 'uploads':
                # Uploads может отсутствовать, создаем пустую директорию
                os.makedirs('uploads', exist_ok=True)
                print(f"Создана директория: uploads")
            else:
                print(f"ПРЕДУПРЕЖДЕНИЕ: Файл или директория не найдены: {file}")
    
    # Создаем архив
    try:
        with zipfile.ZipFile(archive_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Добавляем файлы
            for item in FILES_TO_INCLUDE:
                if os.path.isfile(item):
                    zipf.write(item)
                    print(f"Добавлен файл: {item}")
                elif os.path.isdir(item):
                    for root, dirs, files in os.walk(item):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = file_path  # Имя внутри архива
                            zipf.write(file_path, arcname)
                    print(f"Добавлена директория: {item}")
            
            # Создаем пустую директорию uploads
            # Я решил просто создать пустой файл-маркер, чтобы директория появилась
            marker_path = os.path.join('uploads', '.keep')
            with open(marker_path, 'w') as f:
                f.write('# Эта директория используется для загруженных файлов\n')
            zipf.write(marker_path, marker_path)
            os.remove(marker_path)  # Удаляем временный файл
            
        print(f"Архив успешно создан: {archive_name}")
        print(f"Размер архива: {os.path.getsize(archive_name) / 1024:.2f} КБ")
        return True
        
    except Exception as e:
        print(f"Ошибка при создании архива: {e}")
        return False

if __name__ == "__main__":
    # Можно указать имя архива в аргументах командной строки
    archive_name = "web_app_deploy.zip"
    if len(sys.argv) > 1:
        archive_name = sys.argv[1]
        
    create_deploy_archive(archive_name) 