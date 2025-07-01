#!/usr/bin/env python3.13
# -*- coding: utf-8 -*-
"""
Скрипт для анализа документов DOCX, чтобы выявить проблемы с определением типов элементов и их форматирования. Сгенерирован с LLM
"""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import sys
import os

# Импортируем функции из нашего модуля
from formatting_checker import (
    is_in_table, is_main_heading, is_introduction_heading, is_bibliography_heading,
    is_section_heading, is_subsection_heading, is_figure_caption, is_table_title,
    is_bibliography_item, is_list_item, is_appendix_heading, get_paragraph_type
)

from formatting_utils import (
    get_effective_first_line_indent_obj, get_effective_alignment, get_first_line_indent_cm
)

def analyze_document(docx_path):
    """
    Анализирует документ DOCX и выводит информацию о типах элементов и их форматировании.
    """
    print(f"Анализ документа: {os.path.basename(docx_path)}")
    
    try:
        doc = Document(docx_path)
        
        # Счетчик параграфов
        paragraphs = [p for p in doc.paragraphs if p.text.strip()]
        print(f"Всего параграфов: {len(paragraphs)}")
        
        # Определяем, находимся ли мы в разделе библиографии
        in_bibliography_section = False
        previous_para_type = None
        
        # Первый проход - определяем разделы
        for i, para in enumerate(paragraphs):
            if is_bibliography_heading(para):
                in_bibliography_section = True
                break
        
        # Сбрасываем флаг для второго прохода
        in_bibliography_section = False
        
        # Второй проход - анализ параграфов
        for i, para in enumerate(paragraphs):
            # Обновляем флаг раздела библиографии
            if is_bibliography_heading(para):
                in_bibliography_section = True
            
            # Получаем тип параграфа
            para_type = get_paragraph_type(para, doc, in_bibliography_section, previous_para_type)
            previous_para_type = para_type
            
            # Анализ форматирования
            first_line_indent = "Не установлен"
            alignment = "Не установлено"
            font_info = "Не установлен, Не установлен, Жирный: Нет"
            line_spacing = "Не установлен"
            
            # Получаем прямо установленный отступ первой строки
            if hasattr(para, 'paragraph_format') and para.paragraph_format:
                if hasattr(para.paragraph_format, 'first_line_indent') and para.paragraph_format.first_line_indent:
                    indent_cm = 0
                    try:
                        indent_cm = para.paragraph_format.first_line_indent.cm
                    except:
                        # Если не удалось получить в см, используем значение в pt
                        if hasattr(para.paragraph_format.first_line_indent, 'pt'):
                            indent_cm = para.paragraph_format.first_line_indent.pt / 28.35  # Примерное преобразование pt в см
                    
                    first_line_indent = f"{indent_cm:.2f} см"
            
            # Получаем эффективный отступ первой строки
            try:
                effective_indent_obj = get_effective_first_line_indent_obj(para)
                effective_indent = 0
                if effective_indent_obj:
                    try:
                        effective_indent = effective_indent_obj.cm
                    except:
                        # Если не удалось получить в см, используем значение в pt
                        if hasattr(effective_indent_obj, 'pt'):
                            effective_indent = effective_indent_obj.pt / 28.35
                effective_first_line_indent = f"{effective_indent:.2f} см"
            except:
                effective_first_line_indent = "Не удалось определить"
            
            # Получаем прямо установленное выравнивание
            if hasattr(para, 'paragraph_format') and para.paragraph_format:
                if hasattr(para.paragraph_format, 'alignment') and para.paragraph_format.alignment:
                    if para.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                        alignment = "По центру"
                    elif para.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                        alignment = "По ширине"
                    elif para.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                        alignment = "По правому краю"
            
            # Получаем эффективное выравнивание
            try:
                effective_alignment = get_effective_alignment(para)
                if effective_alignment == WD_ALIGN_PARAGRAPH.CENTER:
                    effective_alignment = "По центру"
                elif effective_alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                    effective_alignment = "По ширине"
                elif effective_alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                    effective_alignment = "По правому краю"
                else:
                    effective_alignment = "По левому краю"
            except:
                effective_alignment = "Не удалось определить"
            
            # Получаем информацию о шрифте из первого run
            if para.runs:
                run = para.runs[0]
                font_name = "Не установлен"
                font_size = "Не установлен"
                is_bold = "Нет"
                
                if hasattr(run, 'font') and run.font:
                    if hasattr(run.font, 'name') and run.font.name:
                        font_name = run.font.name
                    if hasattr(run.font, 'size') and run.font.size:
                        try:
                            size_pt = run.font.size.pt
                            font_size = f"{size_pt} пт"
                        except:
                            font_size = "Не удалось определить"
                    if hasattr(run, 'bold') and run.bold:
                        is_bold = "Да"
                        
                font_info = f"{font_name}, {font_size}, Жирный: {is_bold}"
            
            # Получаем межстрочный интервал
            if hasattr(para, 'paragraph_format') and para.paragraph_format:
                if hasattr(para.paragraph_format, 'line_spacing') and para.paragraph_format.line_spacing:
                    try:
                        line_spacing = f"{para.paragraph_format.line_spacing:.2f}"
                    except:
                        line_spacing = "Не удалось определить"
            
            # Вывод информации о параграфе
            print("-" * 80)
            # Ограничиваем длину текста для вывода
            display_text = para.text[:50] + "..." if len(para.text) > 50 else para.text
            print(f"Параграф {i}: {display_text}")
            print(f"  Тип: {para_type}")
            print(f"  Стиль: {para.style.name if hasattr(para, 'style') and para.style else 'Не установлен'}")
            print(f"  Прямой отступ первой строки: {first_line_indent}")
            print(f"  Эффективный отступ первой строки: {effective_first_line_indent}")
            print(f"  Прямое выравнивание: {alignment}")
            print(f"  Эффективное выравнивание: {effective_alignment}")
            print(f"  Шрифт (в первом run): {font_info}")
            print(f"  Межстрочный интервал: {line_spacing}")
            
    except Exception as e:
        print(f"Ошибка при анализе документа: {e}")
        import traceback
        traceback.print_exc()

def main():
    if len(sys.argv) < 2:
        print("Использование: python analyze_docx.py <путь_к_docx_файлу>")
        return
    
    docx_path = sys.argv[1]
    if not os.path.exists(docx_path):
        print(f"Файл не найден: {docx_path}")
        return
    
    analyze_document(docx_path)

if __name__ == "__main__":
    main() 