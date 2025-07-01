"""
Утилиты для работы с форматированием документов DOCX.
"""

from docx.shared import Pt, Cm, RGBColor 
from docx.enum.text import WD_ALIGN_PARAGRAPH

def _get_style_attr(style_obj, attr_path):
    """Вспомогательная функция для безопасного получения атрибута из цепочки стилей."""
    current_style = style_obj
    while current_style:
        obj = current_style
        found = True
        for attr_name in attr_path.split('.'):
            if hasattr(obj, attr_name):
                obj = getattr(obj, attr_name)
                if obj is None and attr_name != attr_path.split('.')[-1]:
                    found = False
                    break
            else:
                found = False
                break
        if found and obj is not None:
            return obj
        current_style = current_style.base_style if hasattr(current_style, 'base_style') else None
    return None


def get_effective_first_line_indent_obj(para):
    """
    Функция, чтобы получать отступ первой строки с учетом наследования стилей.
    """
    # 1. Прямое форматирование параграфа
    if para.paragraph_format and para.paragraph_format.first_line_indent is not None:
        return para.paragraph_format.first_line_indent
    
    # 2. Стиль параграфа и его базовые стили
    if para.style:
        return _get_style_attr(para.style, 'paragraph_format.first_line_indent')
        
    return None

def get_first_line_indent_cm(para):
    """
    Получает отступ первой строки в сантиметрах, с учетом стилей.
    """
    indent_obj = get_effective_first_line_indent_obj(para)
    if indent_obj and hasattr(indent_obj, 'cm'):
        try:
            return indent_obj.cm
        except:
            if hasattr(indent_obj, 'pt') and indent_obj.pt is not None:
                return indent_obj.pt * 0.0352778
            return 0.0
    return 0.0

def get_effective_alignment(para):
    """
    Получает "эффективное" (т.е. видимое пользователю в Word самом) выравнивание с учетом наследования стилей.
    """
    # 1. Прямое форматирование параграфа
    if para.paragraph_format and para.paragraph_format.alignment is not None:
        return para.paragraph_format.alignment
    
    # 2. Стиль параграфа и его базовые стили
    if para.style:
        alignment = _get_style_attr(para.style, 'paragraph_format.alignment')
        if alignment is not None:
            return alignment
            
    return WD_ALIGN_PARAGRAPH.LEFT

def get_run_font_name(run, para_style=None):
    """Решил добавить функцию получения имени шрифта для фрагмента текста."""
    if run.font.name:
        return run.font.name
    if run.style and run.style.font and run.style.font.name:
        return run.style.font.name
    if para_style:
        font_name = _get_style_attr(para_style, 'font.name')
        if font_name:
            return font_name
    return None 

def get_run_font_size_pt(run, para_style=None):
    """Получает размер шрифта для run в пунктах, учитывая стили."""
    if run.font.size is not None and hasattr(run.font.size, 'pt'):
        return run.font.size.pt
    if run.style and run.style.font and run.style.font.size is not None and hasattr(run.style.font.size, 'pt'):
        return run.style.font.size.pt
    if para_style:
        size = _get_style_attr(para_style, 'font.size')
        if size is not None and hasattr(size, 'pt'):
            return size.pt
    return None

def get_run_font_color_rgb(run, para_style=None):
    """Здесь получаю цвет шрифта для проверки черного текста."""
    if run.font.color and run.font.color.rgb is not None:
        return run.font.color.rgb
    if run.style and run.style.font and run.style.font.color and run.style.font.color.rgb is not None:
        return run.style.font.color.rgb
    if para_style:
        color = _get_style_attr(para_style, 'font.color.rgb')
        if color is not None:
            return color
    return None

def get_run_bold_status(run, para_style=None):
    """тут получаю статус полужирности для run, учитывая стили."""
    if run.bold is not None: 
        return run.bold
    if run.style and run.style.font and run.style.font.bold is not None:
        return run.style.font.bold
    if para_style:
        bold_status = _get_style_attr(para_style, 'font.bold')
        if bold_status is not None:
            return bold_status
    return None