import pdfplumber
from aiogram import Router, types, F
from aiogram import Bot
from docx import Document
import pandas as pd
import json
import re

router = Router()


def clean_section_title(title):
    cleaned_title = re.sub(r'\.+\s*\d*$', '', title).strip()
    return cleaned_title


def add_subsection(data, section_number, section_title):
    section_levels = section_number.split('.')

    if len(section_levels) == 1:
        data[section_number] = {"title": section_title, "sections": {}}
    elif len(section_levels) == 2:
        parent_section = section_levels[0]
        if parent_section not in data:
            data[parent_section] = {"title": f"Bo'lim {parent_section}", "sections": {}}
        data[parent_section]["sections"][section_number] = {"title": section_title, "subsections": {}}
    elif len(section_levels) == 3:
        parent_section = section_levels[0] + '.' + section_levels[1]
        if parent_section not in data:
            data[parent_section] = {"title": f"Bo'lim {parent_section}", "sections": {}}
        data[parent_section]["sections"][section_number] = {"title": section_title, "subsections": {}}


async def extract_contents_from_pdf(pdf_file):
    extracted_data = {}
    found_index = False

    with pdfplumber.open(pdf_file) as pdf:
        for page_num in range(min(15, len(pdf.pages))):
            page = pdf.pages[page_num]
            text = page.extract_text()
            if text:
                lines = text.split("\n")
                for line in lines:
                    if "Mundarija" in line or "Содержание" in line or "Table of Contents" in line or "Оглавление" in line or "Contents" in line:
                        found_index = True
                    if found_index:
                        match = re.match(r'^(\d+(\.\d+)*)(\s+.*)$', line.strip())
                        if match:
                            section_number, section_title = match.groups()[0], match.groups()[2].strip()
                            cleaned_title = clean_section_title(section_title)
                            add_subsection(extracted_data, section_number, cleaned_title)

    return extracted_data if extracted_data else {"error": "Mundarija topilmadi"}


async def extract_text_from_docx(docx_file):
    extracted_data = {}
    found_index = False
    doc = Document(docx_file)

    for paragraph in doc.paragraphs:
        line = paragraph.text.strip()
        if "Mundarija" in line or "Содержание" in line or "Table of Contents" in line or "Оглавление" in line or "Contents" in line:
            found_index = True
        if found_index and line:
            match = re.match(r'^(\d+(\.\d+)*)(\s+.*)$', line)
            if match:
                section_number, section_title = match.groups()[0], match.groups()[2].strip()
                cleaned_title = clean_section_title(section_title)
                add_subsection(extracted_data, section_number, cleaned_title)

    return extracted_data if extracted_data else {"error": "Mundarija topilmadi"}


async def extract_text_from_xlsx(xlsx_file):
    extracted_data = {}
    found_index = False
    df = pd.read_excel(xlsx_file)

    for index, row in df.iterrows():
        row_data = row.tolist()
        for cell in row_data:
            if isinstance(cell, str) and (
                    "Mundarija" in cell or "Содержание" in cell or "Table of Contents" in cell or "Оглавление" in cell or "Contents" in cell):
                found_index = True
            if found_index and isinstance(cell, str):
                match = re.match(r'^(\d+(\.\d+)*)(\s+.*)$', cell.strip())
                if match:
                    section_number, section_title = match.groups()[0], match.groups()[2].strip()
                    cleaned_title = clean_section_title(section_title)
                    add_subsection(extracted_data, section_number, cleaned_title)

    return extracted_data if extracted_data else {"error": "Mundarija topilmadi"}


async def extract_text(file, mime_type):
    if mime_type == "application/pdf":
        return await extract_contents_from_pdf(file)
    elif mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return await extract_text_from_docx(file)
    elif mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        return await extract_text_from_xlsx(file)
    else:
        raise ValueError("Qo'llab-quvvatlanmaydigan fayl turi")


@router.message(F.document)
async def handle_document_upload(message: types.Message, bot: Bot):
    if message.document:
        file = await bot.get_file(message.document.file_id)
        downloaded_file = await bot.download_file(file.file_path)

        try:
            valid_mime_types = [
                "application/pdf",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ]

            if message.document.mime_type not in valid_mime_types:
                await message.reply(
                    "Noto'g'ri fayl turi yuklandi. Iltimos, PDF, DOCX yoki XLSX formatdagi faylni yuklang.")
                return

            extracted_contents = await extract_text(downloaded_file, message.document.mime_type)
            json_data = json.dumps(extracted_contents, ensure_ascii=False, indent=4)

            text_length = len(json_data)
            start_index = 0
            while start_index < text_length:
                end_index = start_index + 4096
                text_chunk = json_data[start_index:end_index]
                await message.reply(text_chunk)

                start_index = end_index

            if start_index >= text_length:
                await message.reply("Barcha mundarija yuborildi.")
        except Exception as e:
            await message.reply(f"Xatolik yuz berdi: {str(e)}")
    else:
        await message.reply(
            "Iltimos, PDF, DOCX yoki XLSX formatdagi faylni yuboring va faylni JSON formatiga aylantirib beraman.")


@router.message(F.text)
async def handle_text_message(message: types.Message):
    await message.reply(
        "Iltimos, PDF, DOCX yoki XLSX formatdagi faylni yuboring va faylni JSON formatiga aylantirib beraman.")
