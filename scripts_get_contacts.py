import os
import docx
import PyPDF2
import openpyxl
import pdfplumber
import requests
import pandas as pd
import json


# 1. Поиск файлов на сетевом диске
def find_files(directory, extensions, limit=3):
    """Функция для поиска файлов с заданными расширениями в первых 'limit' папках."""
    file_list = []
    folder_count = 0
    for root, dirs, files in os.walk(directory):
        if folder_count >= limit:
            break
        folder_count += 1
        for file in files:
            # Пропускаем временные файлы, начинающиеся с '~$'
            if file.startswith('~$'):
                continue
            if file.lower().endswith(extensions):
                file_list.append(os.path.join(root, file))
    return file_list


# 2. Чтение файлов
def read_docx(file_path):
    doc = docx.Document(file_path)
    return "\n".join([paragraph.text for paragraph in doc.paragraphs])


def read_pdf(file_path):
    text = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text


def read_excel(file_path):
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    sheet = workbook.active
    text = ""
    for row in sheet.iter_rows(values_only=True):
        text += " ".join([str(cell) for cell in row if cell]) + "\n"
    return text


def read_file(file_path):
    if file_path.endswith('.docx'):
        return read_docx(file_path)
    elif file_path.endswith('.pdf'):
        return read_pdf(file_path)
    elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        return read_excel(file_path)
    else:
        return "Unsupported file format"


# 3. Взаимодействие с нейросетью Llama для извлечения контактов
def extract_contacts(text):
    url = "http://127.0.0.1:1234/v1/chat/completions"
    headers = {"Content-Type": "application/json"}

    # Формируем тело запроса
    data = {
        "model": "lmstudio-community/Meta-Llama-3.1-8B-Instruct-GGUF/Meta-Llama-3.1-8B-Instruct-Q8_0.gguf",
        "messages": [
            {"role": "system",
             "content": "Extract contact information including full name, company, phone number, and email from the provided text."},
            {"role": "user", "content": text}
        ],
        "temperature": 0.8,
        "max_tokens": -1,
        "stream": False
    }

    # Отправляем запрос к Llama
    response = requests.post(url, headers=headers, data=json.dumps(data))

    if response.status_code == 200:
        # Обрабатываем JSON-ответ
        response_data = response.json()
        if "choices" in response_data and len(response_data["choices"]) > 0:
            return response_data["choices"][0]["message"]["content"]
    else:
        print(f"Ошибка при взаимодействии с нейросетью: {response.status_code}")
        return ""


# 4. Парсинг результатов, извлечённых из модели
def parse_contact_info(contacts_text):
    # Ищем ключевые слова, такие как "Имя", "Фамилия", "Компания", "Телефон", и извлекаем информацию
    full_name = None
    company = None
    phone = None
    email = None

    for line in contacts_text.split("\n"):
        if "Имя:" in line:
            full_name = line.split(":")[1].strip()
        elif "Компания:" in line:
            company = line.split(":")[1].strip()
        elif "Телефон:" in line:
            phone = line.split(":")[1].strip()
        elif "Email:" in line:
            email = line.split(":")[1].strip()

    return full_name, company, phone, email


# 5. Сохранение результатов в Excel файл
def save_to_excel(data, output_file):
    df = pd.DataFrame(data, columns=["Организация", "Контактные данные", "ФИО", "Должность", "Телефон", "Email"])
    df.to_excel(output_file, index=False)


# Основная логика скрипта
def process_files(directory, extensions, output_file, folder_limit=3):
    # Шаг 1: Поиск файлов
    files = find_files(directory, extensions, limit=folder_limit)

    # Шаг 2: Чтение и обработка каждого файла
    extracted_data = []
    for file in files:
        print(f"Обработка файла: {file}")

        # Чтение файла
        text = read_file(file)

        # Извлечение контактов через Llama
        contacts_text = extract_contacts(text)

        # Парсинг полученной информации
        full_name, company, phone, email = parse_contact_info(contacts_text)

        # Добавление данных для сохранения
        organization = company if company else "Неизвестно"
        extracted_data.append([
            organization,
            text[:500],  # Можно взять часть текста как контактные данные
            full_name if full_name else "Не найдено",
            "Не найдено",  # Должность не указана явно
            phone if phone else "Не найдено",
            email if email else "Не найдено"
        ])

    # Шаг 3: Сохранение в Excel
    save_to_excel(extracted_data, output_file)
    print(f"Данные сохранены в {output_file}")


# Пример использования:
if __name__ == "__main__":
    directory = r'\\192.168.99.15\04 Commercial\ЗАКАЗЧИКИ'  # Сетевой диск
    extensions = ('.docx', '.doc', '.pdf', '.xls', '.xlsx')  # Поддерживаемые расширения
    output_file = "contacts_test.xlsx"  # Файл для записи результатов

    # Запуск обработки для первых 3 папок
    process_files(directory, extensions, output_file, folder_limit=3)
