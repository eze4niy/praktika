import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.types import Message, ContentType
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.filters import Command
from aiogram.enums import ParseMode
import pandas as pd
import os

API_TOKEN = '7126313019:AAG0Y7FYj9ftjs6s4TSbb0VxvbESEAqj59c'

bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

async def start_command(message: Message):

    await message.answer("Отправьте Excel файл.")

async def handle_file(message: Message):
    document = message.document

    if not document.file_name.endswith('.xlsx'):
        await message.answer("Пожалуйста отправьте файл в формате Excel (.xlsx).")
        return

    file_path = os.path.join(UPLOAD_DIR, document.file_name)

    try:
        with open(file_path, 'wb') as f:
            file = await bot.download(document.file_id, destination=f)

        try:
            data = pd.read_excel(file_path, sheet_name='Worksheet')

            try:
                month_data = data.iloc[1:, [1, 5, 4]]
                month_data = month_data.rename(columns={month_data.columns[0]: 'ФИО преподавателя', month_data.columns[1]: 'Проверено', month_data.columns[2]: 'Получено'})

                month_data = month_data.dropna()
                month_data['Проверено'] = pd.to_numeric(month_data['Проверено'], errors='coerce')
                month_data['Получено'] = pd.to_numeric(month_data['Получено'], errors='coerce')
                month_data['% проверки'] = (month_data['Проверено'] / month_data['Получено']) * 100

                low_check_month = month_data[month_data['% проверки'] < 75]
            except Exception as e:
                low_check_month = pd.DataFrame()

            try:
                week_data = data.iloc[1:, [1, 10, 9]]  # Столбцы для анализа
                week_data = week_data.rename(columns={week_data.columns[0]: 'ФИО преподавателя', week_data.columns[1]: 'Проверено', week_data.columns[2]: 'Получено'})

                week_data = week_data.dropna()
                week_data['Проверено'] = pd.to_numeric(week_data['Проверено'], errors='coerce')
                week_data['Получено'] = pd.to_numeric(week_data['Получено'], errors='coerce')
                week_data['% проверки'] = (week_data['Проверено'] / week_data['Получено']) * 100

                low_check_week = week_data[week_data['% проверки'] < 75]
            except Exception as e:
                low_check_week = pd.DataFrame()

            response = ""

            if low_check_month.empty:
                response += "Все преподаватели проверили более 75% домашних заданий за месяц.\n\n"
            else:
                response += "Преподаватели с низким процентом проверки домашних заданий за месяц:\n"
                for i in range(len(low_check_month)):
                    response += f"\n- <b>{low_check_month.iloc[i]['ФИО преподавателя']}</b>: {low_check_month.iloc[i]['% проверки']:.2f}% ({int(low_check_month.iloc[i]['Проверено'])} из {int(low_check_month.iloc[i]['Получено'])})"
                response += "\n\n"

            if low_check_week.empty:
                response += "Все преподаватели проверили более 75% домашних заданий за неделю."
            else:
                response += "Преподаватели с низким процентом проверки за неделю:\n"
                for i in range(len(low_check_week)):
                    response += f"\n- <b>{low_check_week.iloc[i]['ФИО преподавателя']}</b>: {low_check_week.iloc[i]['% проверки']:.2f}% ({int(low_check_week.iloc[i]['Проверено'])} из {int(low_check_week.iloc[i]['Получено'])})"

            await message.answer(response, parse_mode=ParseMode.HTML)

        except Exception as e:
            await message.answer(f"Произошла ошибка при обработке файла: {e}")

    finally:

        pass

async def main():
    dp.message.register(start_command, Command(commands=["start"]))
    dp.message.register(handle_file, lambda message: message.content_type == ContentType.DOCUMENT)

    dp.startup.register(startup_handler)
    await dp.start_polling(bot)

async def startup_handler():
    print("Бот успешно запущен.")

if __name__ == '__main__':
    asyncio.run(main())

