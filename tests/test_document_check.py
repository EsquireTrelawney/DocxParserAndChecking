from formatting_checker import check_document_formatting_final
import os

test_file = "test_document_for_normcontrol.docx"
if not os.path.exists(test_file):
    print(f"Файл {test_file} не найден!")
    exit(1)

print(f"Проверяем документ: {test_file}")
comments = check_document_formatting_final(test_file)

print(f"\nНайдено {len(comments)} замечаний:")
comments_by_type = {}

for idx, comment, author in comments:
    # Я решил группировать комментарии для удобства анализа
    found_category = False
    for category in ["подзаголовк", "заголовок", "точк", "регистр", "поля", "отступ"]:
        if category in comment.lower():
            if category not in comments_by_type:
                comments_by_type[category] = []
            comments_by_type[category].append((idx, comment))
            found_category = True
            break
    
    if not found_category:
        if "другое" not in comments_by_type:
            comments_by_type["другое"] = []
        comments_by_type["другое"].append((idx, comment))
    
    print(f"[{idx}] {comment}")

print("\n=== Группировка комментариев по типам ===")
for category, items in comments_by_type.items():
    print(f"\n--- {category.upper()} ({len(items)}) ---")
    for idx, comment in items:
        print(f"[{idx}] {comment}") 