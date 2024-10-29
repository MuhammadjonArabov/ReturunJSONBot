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

def clean_title(title: str) -> str:
    """Remove unwanted characters from titles."""
    return re.sub(r'\.+\s*\d+$', '', title).strip()  # Remove trailing dots and page numbers

async def extract_contents_from_pdf(pdf_file):
    """Extract sections and subsections from a PDF file and return as JSON."""
    extracted_data = {}
    current_section = None
    current_subsection = None
    section_counter = 1

    section_pattern = re.compile(r'^(Глава \d+|Быстрый старт|Начало работы|Учет денежных средств).*?\d+$')
    subsection_pattern = re.compile(r'^\d+\.\d+\s.*?\d+$')

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            lines = page.extract_text().split("\n")
            for line in lines:
                line = line.strip()

                # Detect main sections
                section_match = section_pattern.match(line)
                if section_match:
                    clean_title_text = clean_title(line)
                    current_section = str(section_counter)  # Use section_counter for key
                    extracted_data[current_section] = {
                        "title": clean_title_text,
                        "sections": {}
                    }
                    section_counter += 1  # Increment the section counter
                    current_subsection = None  # Reset current subsection
                    continue

                # Detect subsections
                subsection_match = subsection_pattern.match(line)
                if subsection_match:
                    if current_section:
                        clean_subtitle = clean_title(line)
                        current_subsection = f"{current_section}.{clean_subtitle.split('.')[0]}"  # Create subsection key
                        extracted_data[current_section]["sections"][current_subsection] = {
                            "title": clean_subtitle,
                            "subsections": {}  # Initialize subsections as empty
                        }
                    continue

    return extracted_data

async def extract_text_from_docx(docx_file):
    """Extract sections and subsections from a DOCX file."""
    extracted_data = {}
    current_section = None
    current_subsection = None
    section_counter = 1
    doc = Document(docx_file)

    for paragraph in doc.paragraphs:
        line = paragraph.text.strip()
        if re.match(r'^\d+\.\s', line):
            current_section = str(section_counter)  # Use section counter for key
            extracted_data[current_section] = {
                "title": clean_title(line),
                "sections": {}
            }
            section_counter += 1  # Increment section counter
            current_subsection = None
        elif re.match(r'^\d+\.\d+\s', line):
            if current_section:
                clean_subtitle = clean_title(line)
                current_subsection = f"{current_section}.{clean_subtitle.split('.')[0]}"
                extracted_data[current_section]["sections"][current_subsection] = {
                    "title": clean_subtitle,
                    "subsections": {}
                }
        elif current_section:
            if current_subsection:
                # Here we might add subsections or related data if needed
                pass  # Currently, no specific logic for subsections in DOCX

    return extracted_data

async def extract_text_from_xlsx(xlsx_file):
    """Extract sections and subsections from an XLSX file."""
    extracted_data = {}
    current_section = None
    current_subsection = None
    section_counter = 1
    df = pd.read_excel(xlsx_file)

    for _, row in df.iterrows():
        for cell in row:
            cell_text = str(cell).strip()
            if re.match(r'^\d+\.\s', cell_text):
                current_section = str(section_counter)  # Use section counter for key
                extracted_data[current_section] = {
                    "title": clean_title(cell_text),
                    "sections": {}
                }
                section_counter += 1
                current_subsection = None
            elif re.match(r'^\d+\.\d+\s', cell_text):
                if current_section:
                    clean_subtitle = clean_title(cell_text)
                    current_subsection = f"{current_section}.{clean_subtitle.split('.')[0]}"
                    extracted_data[current_section]["sections"][current_subsection] = {
                        "title": clean_subtitle,
                        "subsections": {}
                    }
            elif current_section:
                if current_subsection:
                    # Similar to DOCX, we can add logic for subsections if needed
                    pass  # No logic for subsections in XLSX

    return extracted_data

async def extract_text(file, mime_type):
    """Calls the appropriate extraction function based on the file type."""
    if mime_type == "application/pdf":
        return await extract_contents_from_pdf(file)
    elif mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return await extract_text_from_docx(file)
    elif mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        return await extract_text_from_xlsx(file)
    else:
        raise ValueError("Unsupported file type")

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
                await message.reply("Invalid file type. Please upload a PDF, DOCX, or XLSX file.")
                return

            extracted_contents = await extract_text(downloaded_file, message.document.mime_type)
            json_data = json.dumps(extracted_contents, ensure_ascii=False, indent=4)

            # Send JSON in chunks
            text_length = len(json_data)
            start_index = 0
            while start_index < text_length:
                end_index = start_index + 4096
                text_chunk = json_data[start_index:end_index]
                await message.reply(text_chunk)
                start_index = end_index

            await message.reply("All text has been sent in JSON format.")
        except Exception as e:
            await message.reply(f"An error occurred: {str(e)}")
    else:
        await message.reply("Please upload a PDF, DOCX, or XLSX file to convert to JSON.")

@router.message(F.text)
async def handle_text_message(message: types.Message):
    await message.reply("Please upload a PDF, DOCX, or XLSX file to convert to JSON.")

async def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(filename)s:%(lineno)d #%(levelname)-8s "
               "[%(asctime)s] - %(name)s - %(message)s",
    )

    logger.info("Starting bot")
    bot_token = os.getenv("BOT_TOKEN")
    if not bot_token:
        logger.error("Bot token not found. Check your .env file.")
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
