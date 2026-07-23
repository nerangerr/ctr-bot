import docx
import PyPDF2
import os

def test_docx(file_path):
    try:
        doc = docx.Document(file_path)
        text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        print(f"DOCX: извлечено {len(text)} символов")
        print(f"Текст: {text[:200]}")
        return text
    except Exception as e:
        print(f"DOCX ошибка: {e}")
        return ""

def test_pdf(file_path):
    try:
        text = ""
        with open(file_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text += (page.extract_text() or "") + "\n"
        print(f"PDF: извлечено {len(text)} символов")
        print(f"Текст: {text[:200]}")
        return text
    except Exception as e:
        print(f"PDF ошибка: {e}")
        return ""

# Тестируем
file_path = input("Введите полный путь к файлу: ")
if not os.path.exists(file_path):
    print("Файл не найден!")
else:
    if file_path.lower().endswith(".docx"):
        test_docx(file_path)
    elif file_path.lower().endswith(".pdf"):
        test_pdf(file_path)
    else:
        print("Неподдерживаемый формат")