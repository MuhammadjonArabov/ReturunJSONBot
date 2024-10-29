import pdfplumber
import pandas as pd
from docx import Document
import re

def extract_contents_from_pdf(pdf_file):
    """PDF fayldan barcha ma'lumotlarni sarlavha va matnlar bo'yicha o'qib, JSON formatga o'tkazish"""
    extracted_data = {}
    with pdfplumber.open(pdf_file) as pdf:
        current_section = None
        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            for line in lines:
                # Sarlavha bilan matnni ajratish uchun shart
                if re.match(r'^\d+\.\s', line):  # misol: "1.1 Bo'lim nomi"
                    current_section = line.strip()
                    extracted_data[current_section] = {"text": ""}
                elif current_section:
                    extracted_data[current_section]["text"] += line + "\n"
    return extracted_data


def extract_text_from_docx(docx_file):
    """DOCX fayldan barcha ma'lumotlarni sarlavha va matnlar bo'yicha o'qib, JSON formatga o'tkazish"""
    extracted_data = {}
    current_section = None
    doc = Document(docx_file)
    for paragraph in doc.paragraphs:
        line = paragraph.text.strip()
        if re.match(r'^\d+\.\s', line):  # misol: "1.1 Bo'lim nomi"
            current_section = line
            extracted_data[current_section] = {"text": ""}
        elif current_section:
            extracted_data[current_section]["text"] += line + "\n"
    return extracted_data


def extract_text_from_xlsx(xlsx_file):
    """XLSX fayldan barcha ma'lumotlarni sarlavha va matnlar bo'yicha o'qib, JSON formatga o'tkazish"""
    extracted_data = {}
    df = pd.read_excel(xlsx_file)
    current_section = None
    for _, row in df.iterrows():
        for cell in row:
            cell_text = str(cell).strip()
            if re.match(r'^\d+\.\s', cell_text):  # misol: "1.1 Bo'lim nomi"
                current_section = cell_text
                extracted_data[current_section] = {"text": ""}
            elif current_section and cell_text:
                extracted_data[current_section]["text"] += cell_text + "\n"
    return extracted_data


def extract_text(file, mime_type):
    """Fayl turiga qarab tegishli funksiyani chaqiradi"""
    if mime_type == "application/pdf":
        return extract_contents_from_pdf(file)
    elif mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return extract_text_from_docx(file)
    elif mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        return extract_text_from_xlsx(file)
    else:
        raise ValueError("Qo'llab-quvvatlanmaydigan fayl turi")
