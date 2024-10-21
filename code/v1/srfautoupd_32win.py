import os
import win32com.client as win32
from datetime import datetime

def convert_doc_to_docx(input_path, output_path):
    """
    Конвертирует .doc файл в .docx.
    """
    word = win32.Dispatch("Word.Application")
    try:
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=16)  # 16 - формат .docx
        doc.Close()
    except Exception as e:
        print(f"Ошибка при конвертации {input_path}: {e}")
    finally:
        word.Quit()

def replace_markers_with_word(doc_path, output_path, replacements):
    """
    Заменяет маркеры в .docx файле с помощью win32com.
    """
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(doc_path)
        
        # Проходим по всем маркерам и заменяем их в документе
        for placeholder, replacement in replacements.items():
            # Выполняем поиск и замену
            word.Selection.Find.Execute(
                FindText=placeholder, 
                ReplaceWith=replacement, 
                Replace=2  # wdReplaceAll
            )
        
        # Сохраняем изменения
        doc.SaveAs(output_path)
        doc.Close()
    except Exception as e:
        print(f"Ошибка при замене маркеров в файле {doc_path}: {e}")
    finally:
        word.Quit()

def process_files_with_template(input_folder, template_path, output_folder):
    """
    Основная функция обработки:
    1. Конвертация .doc в .docx.
    2. Замена маркеров в шаблоне.
    3. Сохранение готового документа.
    """
    # Получение текущей даты
    current_date = datetime.now().strftime("%d.%m.%Y")
    replacements = {"{{date}}": current_date}
    
    # Создаем выходную папку, если её нет
    os.makedirs(output_folder, exist_ok=True)
    
    for filename in os.listdir(input_folder):
        if filename.endswith(".doc"):
            input_path = os.path.join(input_folder, filename)
            temp_docx_path = os.path.join(input_folder, filename.replace(".doc", ".docx"))
            output_path = os.path.join(output_folder, "new_" + filename.replace(".doc", ".docx"))
            
            # Конвертация .doc в .docx
            print(f"Конвертация файла {filename} в .docx...")
            convert_doc_to_docx(input_path, temp_docx_path)
            
            # Замена маркеров в .docx файле с помощью Word
            print(f"Замена маркеров в файле {filename}...")
            replace_markers_with_word(temp_docx_path, output_path, replacements)
            
            # Удаление временного .docx файла
            os.remove(temp_docx_path)
            print(f"Обработан и сохранен файл: {output_path}")

# Указываем пути к папкам и файлам
input_folder = r"C:\Users\demchenko\Desktop\SRFauto\test\old_docs"
template_path = r"C:\Users\demchenko\Desktop\SRFauto\test\template\3.docx"
output_folder = r"C:\Users\demchenko\Desktop\SRFauto\test\new_docs"

# Запуск обработки
process_files_with_template(input_folder, template_path, output_folder)

# Добавляем проверку и вывод даты в правильном формате
current_date = datetime.now().strftime("%d.%m.%Y")
print(f"Текущая дата: {current_date}")



