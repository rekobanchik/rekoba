import os
from win32com.client import Dispatch, constants

def convert_doc_to_docx(folder_path):
    # Ініціалізація Word Application
    word = Dispatch("Word.Application")
    word.Visible = False
    converted_files = []  # Список для відслідковування сконвертованих файлів

    # Рекурсивний пошук файлів
    for root, dirs, files in os.walk(folder_path):
        print(f"Обробляємо директорію: {root}")  # Виводимо поточну директорію
        for file in files:
            if file.endswith(".doc") and not file.startswith("~$"):  # Виключення тимчасових файлів
                doc_path = os.path.join(root, file)
                docx_path = doc_path + "x"  # Заміна розширення на .docx
                
                try:
                    print(f"Конвертуємо файл: {doc_path}")
                    doc = word.Documents.Open(doc_path)
                    doc.SaveAs(docx_path, constants.wdFormatXMLDocument)  # Збереження у .docx
                    doc.Close()
                    converted_files.append(docx_path)  # Додаємо файл до списку
                    print(f"Конвертовано: {doc_path} -> {docx_path}")
                except Exception as e:
                    print(f"Помилка з файлом {doc_path}: {e}")

    word.Quit()  # Закриття Word
    return converted_files

def convert_xls_to_xlsx(folder_path):
    # Ініціалізація Excel Application
    excel = Dispatch("Excel.Application")
    excel.Visible = False
    converted_files = []  # Список для відслідковування сконвертованих файлів

    # Рекурсивний пошук файлів
    for root, dirs, files in os.walk(folder_path):
        print(f"Обробляємо директорію: {root}")  # Виводимо поточну директорію
        for file in files:
            if file.endswith(".xls") and not file.startswith("~$"):  # Виключення тимчасових файлів
                xls_path = os.path.join(root, file)
                xlsx_path = xls_path + "x"  # Заміна розширення на .xlsx
                
                try:
                    print(f"Конвертуємо файл: {xls_path}")
                    workbook = excel.Workbooks.Open(xls_path)
                    workbook.SaveAs(xlsx_path, FileFormat=51)  # 51 - формат .xlsx
                    workbook.Close()
                    converted_files.append(xlsx_path)  # Додаємо файл до списку
                    print(f"Конвертовано: {xls_path} -> {xlsx_path}")
                except Exception as e:
                    print(f"Помилка з файлом {xls_path}: {e}")

    excel.Quit()  # Закриття Excel
    return converted_files

def convert_files_in_folder():
    # Запитуємо у користувача шлях до директорії
    folder_to_scan = input("Введіть шлях до директорії, де знаходяться файли для конвертації: ")

    # Перевіряємо, чи існує вказана директорія
    if not os.path.isdir(folder_to_scan):
        print("Помилка: Вказаний шлях не є директорією або вона не існує!")
        return

    print(f"Починаємо обробку директорії: {folder_to_scan}")

    converted_docs = convert_doc_to_docx(folder_to_scan)
    converted_xls = convert_xls_to_xlsx(folder_to_scan)
    
    print("\nСконвертовані файли:")
    print("Документи .doc -> .docx:")
    if converted_docs:
        print("\n".join(converted_docs))
    else:
        print("Не знайдено файлів для конвертації .doc.")
        
    print("Таблиці .xls -> .xlsx:")
    if converted_xls:
        print("\n".join(converted_xls))
    else:
        print("Не знайдено файлів для конвертації .xls.")

# Використання
convert_files_in_folder()