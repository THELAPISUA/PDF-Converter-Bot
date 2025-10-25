import asyncio
import logging
import sys

import openpyxl
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

from aiogram import Bot, Dispatcher, F, html
from aiogram.filters import CommandStart, Command
from aiogram.types import Message, Document, FSInputFile
import os

TOKEN = os.getenv("TOKEN")
dp = Dispatcher()
bot = Bot(token=TOKEN)


def get_cell_color(cell):
    """Extract RGB color from cell using direct attribute access"""
    try:
        fill = cell.fill
        if not fill or not hasattr(fill, 'fgColor'):
            return None

        fg_color = fill.fgColor

        # Access the raw RGB value directly from the object's __dict__
        # This bypasses openpyxl's validation
        if fg_color and hasattr(fg_color, '__dict__'):
            color_dict = fg_color.__dict__

            # Try to get RGB value
            if 'rgb' in color_dict and color_dict['rgb']:
                rgb_val = color_dict['rgb']
                if rgb_val and rgb_val not in ['00000000', None]:
                    return rgb_val

            # Try to get theme + tint
            if 'theme' in color_dict and color_dict['theme'] is not None:
                theme = color_dict['theme']
                tint = color_dict.get('tint', 0)

                # Standard Excel theme colors
                theme_colors = {
                    0: 'FF000000',  # dk1
                    1: 'FFFFFFFF',  # lt1
                    2: 'FF1F497D',  # dk2 (blue)
                    3: 'FFEEECE1',  # lt2 (tan)
                    4: 'FF4F81BD',  # accent1 (blue)
                    5: 'FFC0504D',  # accent2 (red)
                    6: 'FF9BBB59',  # accent3 (green)
                    7: 'FF8064A2',  # accent4 (purple)
                    8: 'FF4BACC6',  # accent5 (aqua)
                    9: 'FFF79646',  # accent6 (orange)
                }

                if theme in theme_colors:
                    return theme_colors[theme]

            # Try indexed color
            if 'indexed' in color_dict and color_dict['indexed']:
                return None  # Skip indexed colors for now

    except Exception as e:
        pass

    return None


def parse_xlsx(filename: str):
    try:
        wb = openpyxl.load_workbook(filename)
    except FileNotFoundError:
        print("ERROR", "000", "-", "failed to open file")
        exit(-1)

    ws = wb.active

    data = []
    data_color = []

    for row in ws.iter_rows():
        row_data = []
        row_data_color = []
        for cell in row:
            cell_value = cell.value if cell.value is not None else ""
            row_data.append(str(cell_value))

            color_code = get_cell_color(cell)
            row_data_color.append(color_code)

        data_color.append(row_data_color)
        data.append(row_data)

    return data, data_color


def hex_to_reportlab_color(color_value):
    """Convert Excel color to ReportLab Color object"""
    if not color_value:
        return None

    if isinstance(color_value, str):
        if color_value in ["00000000", "00", "None"]:
            return None

        if len(color_value) == 8:
            hex_color = color_value[2:]
        elif len(color_value) == 6:
            hex_color = color_value
        else:
            return None

        try:
            r = int(hex_color[0:2], 16) / 255.0
            g = int(hex_color[2:4], 16) / 255.0
            b = int(hex_color[4:6], 16) / 255.0
            return colors.Color(r, g, b)
        except (ValueError, IndexError):
            return None

    return None


def make_pdf(filename: str, data):
    doc = SimpleDocTemplate(filename, pagesize=letter)

    # data[0] — text
    # data[1] — color
    table = Table(data[0])

    # Start with basic grid style
    style_commands = [
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]

    # Add background colors
    colors_applied = 0
    for row in range(len(data[1])):
        for col in range(len(data[1][row])):
            bg_color_value = data[1][row][col]
            bg_color = hex_to_reportlab_color(bg_color_value)
            if bg_color:
                style_commands.append(('BACKGROUND', (col, row), (col, row), bg_color))
                colors_applied += 1

    print(f"Colors applied: {colors_applied}")

    color_draw = TableStyle(style_commands)
    table.setStyle(color_draw)
    doc.build([table])


@dp.message(CommandStart())
async def start(m: Message):
    await m.answer(
        "👋 Вітаю в *Excel → PDF Converter Bot!*\n\n"
        "Завантаж свій `.xlsx` файл, і бот автоматично створить охайний `.pdf`.\n"
        "Готовий спробувати? Надішли перший файл! ⚡",
        parse_mode="Markdown"
    )


@dp.message(F.document)
async def convert(m: Message, bot: Bot):
    file_id = m.document.file_id

    file = await bot.get_file(file_id)
    os.makedirs("assets/temp", exist_ok=True)
    input_path = f"assets/temp/{m.from_user.id}.xlsx"
    output_path = f"assets/temp/{m.from_user.id}.pdf"

    await bot.download_file(file.file_path, input_path)

    make_pdf(output_path, parse_xlsx(input_path))

    pdf = FSInputFile(output_path)
    await m.answer_document(pdf, caption="✅ Ось твій готовий PDF-файл!")


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    asyncio.run(main())