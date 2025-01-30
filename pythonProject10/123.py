import telebot
import pandas as pd
import re
from io import BytesIO

TOKEN = '7916970390:AAF-tCrH9GLvsaJtj7N3vr4cykhzK6YGMDU'
bot = telebot.TeleBot(TOKEN)

storage = {}

def check_file(file):
    try:
        xlsx = pd.ExcelFile(file, engine='openpyxl')
        sheet = xlsx.parse(sheet_name=0)

        bad_rows = []
        for i, row in sheet.iterrows():
            topic = row['Тема урока']
            if pd.isna(topic) or topic.strip() == '' or not re.match(r'^Урок №\d+\. Тема:\s?.+', str(topic).strip()):
                bad_rows.append((i, topic))
        return sheet, bad_rows
    except Exception as e:
        return None, e

@bot.message_handler(commands=['start'])
def welcome(msg):
    bot.send_message(
        msg.chat.id,
        "/start - начало работы\n/save - сохранить файл\nОтправьте Excel файл и если понадобится укажите строку и новую тему в формате: строка: Урок № Тема."
    )

@bot.message_handler(content_types=['document'])
def file_handler(msg):
    try:
        doc_id = msg.document.file_id
        doc_info = bot.get_file(doc_id)
        dl_file = bot.download_file(doc_info.file_path)

        temp_buffer = BytesIO(dl_file)
        lesson_data, issues = check_file(temp_buffer)

        if issues is None:
            bot.send_message(msg.chat.id, f"Ошибка файла: {lesson_data}")
        else:
            storage[msg.chat.id] = lesson_data
            response = "Файл загружен."

            if issues:
                response += "\nОшибки найдены в строках:\n"
                for idx, bad_topic in issues:
                    response += f"Строка {idx + 1}: {bad_topic}\n"
            else:
                response += "\nВсе хорошо."

            bot.send_message(msg.chat.id, response)
    except Exception as e:
        bot.send_message(msg.chat.id, f"Что-то пошло не так: {e}")

@bot.message_handler(func=lambda m: ":" in m.text)
def change_topic(msg):
    try:
        if msg.chat.id not in storage:
            bot.send_message(msg.chat.id, "Загрузите файл перед изменениями.")
            return

        lessons = storage[msg.chat.id]
        try:
            ln, topic = msg.text.split(":", 1)
            ln = int(ln.strip()) - 1
            topic = topic.strip()

            if not re.match(r'^Урок №\d+\. Тема:\s?.+', topic):
                bot.send_message(msg.chat.id, "Неверный формат темы.")
                return

            if 0 <= ln < len(lessons):
                lessons.at[ln, 'Тема урока'] = topic
                bot.send_message(msg.chat.id, f"Обновлено: строка {ln + 1}, тема: {topic}")
            else:
                bot.send_message(msg.chat.id, "Строка вне диапазона.")
        except ValueError:
            bot.send_message(msg.chat.id, "Формат: строка: тема.")
    except Exception as e:
        bot.send_message(msg.chat.id, f"Ошибка обновления: {e}")

@bot.message_handler(commands=['save'])
def save_handler(msg):
    try:
        if msg.chat.id not in storage:
            bot.send_message(msg.chat.id, "Нечего сохранять.")
            return

        lessons = storage[msg.chat.id]
        save_name = f"result_{msg.chat.id}.xlsx"
        with BytesIO() as out:
            lessons.to_excel(out, index=False, engine='openpyxl')
            out.seek(0)
            bot.send_document(msg.chat.id, out, visible_file_name=save_name)

        bot.send_message(msg.chat.id, "Сохранено.")
    except Exception as e:
        bot.send_message(msg.chat.id, f"Ошибка сохранения: {e}")

bot.polling()
