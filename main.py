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


def extract_toc_from_text(lines):
    """Mundarijadan bo‘lim va pastki bo‘limlarni, ularning sahifa raqamlarini ajratib olish"""
    toc = {}
    for line in lines:
        # Satrda .... bilan tugaydigan qismni sarlavha sifatida olish
        if '.......' in line:
            title = line.split('.......')[0].strip()
            page_number = re.search(r'\d+$', line)
            if page_number:
                page_number = page_number.group()
                toc[title] = int(page_number)
    return toc


async def extract_contents_from_pdf(pdf_file):
    """PDF fayldan barcha ma'lumotlarni mundarija asosida bo‘limlarga ajratib olish"""
    extracted_data = {}
    with pdfplumber.open(pdf_file) as pdf:
        toc_detected = False
        toc = {}
        for page_num, page in enumerate(pdf.pages, 1):
            lines = page.extract_text().split("\n")

            # Agar mundarija hali topilmagan bo'lsa, uni aniqlash
            if not toc_detected:
                toc = extract_toc_from_text(lines)
                if toc:
                    toc_detected = True
                    continue  # Mundarija topilganidan keyin, sahifani keyin o‘qish

            # Mundarijadagi har bir sarlavha va pastki sarlavha bo'yicha tegishli matnni olish
            for title, start_page in toc.items():
                if start_page == page_num:
                    current_section = title
                    extracted_data[title] = {
                        "title": title,
                        "sections": {},
                        "text": "",
                        "length": 0
                    }

            # Sahifadagi matnni hozirgi bo‘limga kiritish
            if current_section in extracted_data:
                extracted_data[current_section]["text"] += page.extract_text() + "\n"
                extracted_data[current_section]["length"] = len(extracted_data[current_section]["text"])

    return extracted_data


async def extract_text_from_docx(docx_file):
    """DOCX fayldan barcha ma'lumotlarni mundarija asosida ajratib olish"""
    extracted_data = {}
    doc = Document(docx_file)
    lines = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    toc = extract_toc_from_text(lines)
    current_section = None

    for line in lines:
        if current_section in toc and line.startswith(current_section):
            extracted_data[current_section] = {
                "title": line,
                "text": "",
                "sections": {},
                "length": 0
            }
        elif current_section:
            extracted_data[current_section]["text"] += line + "\n"
            extracted_data[current_section]["length"] = len(extracted_data[current_section]["text"])

    return extracted_data


async def extract_text_from_xlsx(xlsx_file):
    """XLSX fayldan barcha ma'lumotlarni mundarija asosida ajratib olish"""
    extracted_data = {}
    df = pd.read_excel(xlsx_file)
    lines = [str(cell).strip() for _, row in df.iterrows() for cell in row if str(cell).strip()]
    toc = extract_toc_from_text(lines)
    current_section = None

    for line in lines:
        if current_section in toc and line.startswith(current_section):
            extracted_data[current_section] = {
                "title": line,
                "text": "",
                "sections": {},
                "length": 0
            }
        elif current_section:
            extracted_data[current_section]["text"] += line + "\n"
            extracted_data[current_section]["length"] = len(extracted_data[current_section]["text"])

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
