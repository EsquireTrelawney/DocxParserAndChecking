import sys
import os
import logging

# Настраиваем логирование (полезно для отладки на PythonAnywhere)
logging.basicConfig(stream=sys.stderr, level=logging.INFO)

# Получаем абсолютный путь к директории приложения
path = os.path.dirname(os.path.abspath(__file__))
if path not in sys.path:
    sys.path.insert(0, path)
    logging.info(f"Added {path} to sys.path")

# Создаем директорию для загрузок, если её нет
uploads_dir = os.path.join(path, 'uploads')
if not os.path.exists(uploads_dir):
    try:
        os.makedirs(uploads_dir)
        logging.info(f"Created uploads directory at {uploads_dir}")
    except Exception as e:
        logging.error(f"Failed to create uploads directory: {e}")
else:
    logging.info(f"Uploads directory exists at {uploads_dir}")

try:
    from app import app as application
    logging.info("Successfully imported Flask application")
except Exception as e:
    logging.error(f"Error importing Flask application: {e}")
    raise

# Проверяем наличие основных файлов
required_modules = ['formatting_checker.py', 'comment_utils.py', 'formatting_utils.py']
for module in required_modules:
    module_path = os.path.join(path, module)
    if os.path.exists(module_path):
        logging.info(f"Found module: {module}")
    else:
        logging.error(f"Missing required module: {module}") 