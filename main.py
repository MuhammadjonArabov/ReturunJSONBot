import asyncio
import logging
import os
import json
import pdfplumber
import pandas as pd
from docx import Document
from aiogram import Bot, Dispatcher, Router, types, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from dotenv import load_dotenv
import re

load_dotenv()

logger = logging.getLogger(__name__)
router = Router()
TOKEN = os.getenv("BOT_TOKEN")

async def extract_contents_from_pdf(pdf_file):
    """PDF fayldan barcha ma'lumotlarni bo‘lim va pastki bo‘limlar bilan ajratib, JSON formatida qaytarish"""
    extracted_data = {}
    current_section = None
    current_subsection = None
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            for line in lines:
                line = line.strip()
                # Sarlavhalarni ajratish sharti
                if re.match(r'^\d+\.\s', line):  # Asosiy bo‘lim (misol: "2. Bo‘lim nomi")
                    current_section = line
                    current_subsection = None
                    extracted_data[current_section] = {"title": line, "sections": {}, "text": ""}
                elif re.match(r'^\d+\.\d+\s', line):  # Pastki bo‘lim (misol: "2.1 Pastki bo‘lim nomi")
                    current_subsection = line
                    extracted_data[current_section]["sections"][current_subsection] = {"title": line, "text": ""}
                elif current_section:
                    # Agar bo‘lim va pastki bo‘limlar mavjud bo‘lsa, matnni kiritish
                    if current_subsection:
                        extracted_data[current_section]["sections"][current_subsection]["text"] += line + "\n"
                    else:
                        extracted_data[current_section]["text"] += line + "\n"
    return extracted_data

async def extract_text_from_docx(docx_file):
    """DOCX fayldan barcha ma'lumotlarni bo‘lim va pastki bo‘limlar bilan ajratib, JSON formatida qaytarish"""
    extracted_data = {}
    current_section = None
    current_subsection = None
    doc = Document(docx_file)
    for paragraph in doc.paragraphs:
        line = paragraph.text.strip()
        # Bo‘limlarni ajratish sharti
        if re.match(r'^\d+\.\s', line):
            current_section = line
            current_subsection = None
            extracted_data[current_section] = {"title": line, "sections": {}, "text": ""}
        elif re.match(r'^\d+\.\d+\s', line):
            current_subsection = line
            extracted_data[current_section]["sections"][current_subsection] = {"title": line, "text": ""}
        elif current_section:
            if current_subsection:
                extracted_data[current_section]["sections"][current_subsection]["text"] += line + "\n"
            else:
                extracted_data[current_section]["text"] += line + "\n"
    return extracted_data

async def extract_text_from_xlsx(xlsx_file):
    """XLSX fayldan barcha ma'lumotlarni bo‘lim va pastki bo‘limlar bilan ajratib, JSON formatida qaytarish"""
    extracted_data = {}
    current_section = None
    current_subsection = None
    df = pd.read_excel(xlsx_file)
    for _, row in df.iterrows():
        for cell in row:
            cell_text = str(cell).strip()
            if re.match(r'^\d+\.\s', cell_text):
                current_section = cell_text
                current_subsection = None
                extracted_data[current_section] = {"title": cell_text, "sections": {}, "text": ""}
            elif re.match(r'^\d+\.\d+\s', cell_text):
                current_subsection = cell_text
                extracted_data[current_section]["sections"][current_subsection] = {"title": cell_text, "text": ""}
            elif current_section:
                if current_subsection:
                    extracted_data[current_section]["sections"][current_subsection]["text"] += cell_text + "\n"
                else:
                    extracted_data[current_section]["text"] += cell_text + "\n"
    return extracted_data

async def extract_text(file, mime_type):
    """Fayl turiga qarab tegishli funksiyani chaqiradi"""
    if mime_type == "application/pdf":
        return await extract_contents_from_pdf(file)
    elif mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return await extract_text_from_docx(file)
    elif mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        return await extract_text_from_xlsx(file)
    else:
        raise ValueError("Qo'llab-quvvatlanmaydigan fayl turi")

@router.message(CommandStart())
async def command_start_handler(message: types.Message) -> None:
    await message.answer(f"Hello, {message.from_user.full_name}!")

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
                await message.reply("Barcha matn JSON formatida yuborildi.")
        except Exception as e:
            await message.reply(f"Xatolik yuz berdi: {str(e)}")
    else:
        await message.reply(
            "Iltimos, PDF, DOCX yoki XLSX formatdagi faylni yuboring va faylni JSON formatiga aylantirib beraman.")

@router.message(F.text)
async def handle_text_message(message: types.Message):
    await message.reply(
        "Iltimos, PDF, DOCX yoki XLSX formatdagi faylni yuboring va faylni JSON formatiga aylantirib beraman.")

async def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(filename)s:%(lineno)d #%(levelname)-8s "
               "[%(asctime)s] - %(name)s - %(message)s",
    )

    logger.info("Starting bot")
    bot_token = os.getenv("BOT_TOKEN")
    if not bot_token:
        logger.error("Bot tokeni topilmadi. .env faylini tekshiring.")
        return

    bot: Bot = Bot(token=bot_token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp: Dispatcher = Dispatcher()
    dp.include_router(router)

    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        logger.info("Bot stopped")
