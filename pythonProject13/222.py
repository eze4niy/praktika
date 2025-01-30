import telebot
import pandas as pd
import os

TOKEN = '7126313019:AAG0Y7FYj9ftjs6s4TSbb0VxvbESEAqj59c'
bot = telebot.TeleBot(TOKEN)

FOLDER = "files"
if not os.path.exists(FOLDER):
    os.makedirs(FOLDER)

@bot.message_handler(commands=['start'])
def start(msg):
    bot.send_message(msg.chat.id, "Отправь Excel файл.")

@bot.message_handler(content_types=['document'])
def file_handler(msg):
    doc = msg.document
    if not doc.file_name.endswith('.xlsx'):
        bot.send_message(msg.chat.id, "Ошибка, нужен файл .xlsx")
        return

    path = os.path.join(FOLDER, doc.file_name)
    file_info = bot.get_file(doc.file_id)
    file_bytes = bot.download_file(file_info.file_path)

    with open(path, 'wb') as f:
        f.write(file_bytes)

    try:
        df = pd.read_excel(path, sheet_name='Worksheet')

        try:
            month = df.iloc[1:, [1, 5, 4]]
            month.columns = ['Преподаватель', 'Сделано', 'Всего']
            month = month.dropna()
            month['Сделано'] = pd.to_numeric(month['Сделано'], errors='coerce')
            month['Всего'] = pd.to_numeric(month['Всего'], errors='coerce')
            month['Процент'] = (month['Сделано'] / month['Всего']) * 100
            low_month = month[month['Процент'] < 75]
        except:
            low_month = pd.DataFrame()

        try:
            week = df.iloc[1:, [1, 10, 9]]
            week.columns = ['Преподаватель', 'Сделано', 'Всего']
            week = week.dropna()
            week['Сделано'] = pd.to_numeric(week['Сделано'], errors='coerce')
            week['Всего'] = pd.to_numeric(week['Всего'], errors='coerce')
            week['Процент'] = (week['Сделано'] / week['Всего']) * 100
            low_week = week[week['Процент'] < 75]
        except:
            low_week = pd.DataFrame()

        text = ""
        if low_month.empty:
            text += "За месяц все хорошо.\n\n"
        else:
            text += "Проблемы за месяц:\n"
            for i in range(len(low_month)):
                row = low_month.iloc[i]
                text += f"\n- {row['Преподаватель']}: {row['Процент']:.2f}% ({int(row['Сделано'])} из {int(row['Всего'])})"
            text += "\n\n"

        if low_week.empty:
            text += "За неделю все хорошо."
        else:
            text += "Проблемы за неделю:\n"
            for i in range(len(low_week)):
                row = low_week.iloc[i]
                text += f"\n- {row['Преподаватель']}: {row['Процент']:.2f}% ({int(row['Сделано'])} из {int(row['Всего'])})"

        bot.send_message(msg.chat.id, text)

    except Exception as e:
        bot.send_message(msg.chat.id, f"Ошибка: {e}")

    finally:
        try:
            os.remove(path)
        except Exception as e:
            print(f"Файл не удален: {e}")

bot.polling(none_stop=True)
