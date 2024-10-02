import telebot
import sqlite3
from datetime import datetime

API_TOKEN = '7898118766:AAFvT99mQjvmKdpe6XZuYhI-zuFaIdYKdqk'
bot = telebot.TeleBot(API_TOKEN)

def connect_db():
    return sqlite3.connect('BdTrainingCenter.db')

def fetch_training_plans(month, plan_type):
    conn = connect_db()
    cursor = conn.cursor()
    
    # Get the first and last day of the selected month
    start_date = f"{datetime.now().year}-{month:02d}-01"
    if month == 12:
        end_date = f"{datetime.now().year + 1}-01-01"
    else:
        end_date = f"{datetime.now().year}-{month + 1:02d}-01"
    
    # Correct table names
    if plan_type == 'PB':
        cursor.execute("SELECT * FROM plane_PB WHERE start_date >= ? AND start_date < ?", (start_date, end_date))
    elif plan_type == 'OT':
        cursor.execute("SELECT * FROM plane_OT WHERE start_date >= ? AND start_date < ?", (start_date, end_date))
    
    plans = cursor.fetchall()
    conn.close()
    return plans

@bot.message_handler(commands=['start'])
def start(message):
    keyboard = [
        [telebot.types.InlineKeyboardButton("Январь", callback_data=1)],
        [telebot.types.InlineKeyboardButton("Февраль", callback_data=2)],
        [telebot.types.InlineKeyboardButton("Март", callback_data=3)],
        [telebot.types.InlineKeyboardButton("Апрель", callback_data=4)],
        [telebot.types.InlineKeyboardButton("Май", callback_data=5)],
        [telebot.types.InlineKeyboardButton("Июнь", callback_data=6)],
        [telebot.types.InlineKeyboardButton("Июль", callback_data=7)],
        [telebot.types.InlineKeyboardButton("Август", callback_data=8)],
        [telebot.types.InlineKeyboardButton("Сентябрь", callback_data=9)],
        [telebot.types.InlineKeyboardButton("Октябрь", callback_data=10)],
        [telebot.types.InlineKeyboardButton("Ноябрь", callback_data=11)],
        [telebot.types.InlineKeyboardButton("Декабрь", callback_data=12)]
    ]
    
    markup = telebot.types.InlineKeyboardMarkup(keyboard)    
    bot.send_message(message.chat.id, "Выберите месяц для заполнения информации о прошедшем обучении:", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.isdigit())
def callback_query(call):
    month = int(call.data)
    markup = telebot.types.InlineKeyboardMarkup()
    markup.add(telebot.types.InlineKeyboardButton(text="ПБ", callback_data=f"plan_PB_{month}"),
               telebot.types.InlineKeyboardButton(text="ОТ", callback_data=f"plan_OT_{month}"))
    
    bot.send_message(call.message.chat.id, "Выберите тип плана:", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith('plan_'))
def select_plan(call):
    plan_type, month = call.data.split('_')[1], int(call.data.split('_')[2])
    plans = fetch_training_plans(month, plan_type)
    plan_type_name = 'ПБ' if plan_type=='PB' else 'ОТ'
    
    if not plans:
        bot.send_message(call.message.chat.id, f"Не было запланировано обучений в этом месяцедля типа {plan_type_name}.")
        return
    
    ask_training(call.message.chat.id, plans, 0, plan_type)

def ask_training(chat_id, plans, index, plan_type):
    if index >= len(plans):
        bot.send_message(chat_id, "Все запланированные обучения были внесены в базу.")
        return
    
    employee_id, curse_num, date ,start_date, end_date = plans[index]
    
    # Fetch employee name
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT full_name FROM employee WHERE ID = ?", (employee_id,))
    employee_name = cursor.fetchone()
    
    if employee_name:
        employee_name = employee_name[0]
        plan_type_name = 'ПБ' if plan_type=='PB' else 'ОТ'
        question_text = f"Прошел ли {employee_name} в этом месяце обучение {plan_type_name} {curse_num}?"
        
        markup = telebot.types.InlineKeyboardMarkup()
        markup.add(telebot.types.InlineKeyboardButton("Да", callback_data=f"yes_{index}_{curse_num}_{plan_type}"),
                   telebot.types.InlineKeyboardButton("Нет", callback_data=f"no_{index}_{curse_num}_{plan_type}"))
        
        bot.send_message(chat_id, question_text, reply_markup=markup)
    else:
        bot.send_message(chat_id, "Ошибка: имя сотрудника не найдено.")
    
    conn.close()

@bot.callback_query_handler(func=lambda call: call.data.startswith('yes_'))
def handle_yes(call):
    index, curse_num, plan_type = call.data.split('_')[1:]
    index = int(index)
    bot.send_message(call.message.chat.id, "Введите номер подтверждающего документа:")
    bot.register_next_step_handler(call.message, handle_document_number, index, curse_num, plan_type)

@bot.callback_query_handler(func=lambda call: call.data.startswith('no_'))
def handle_no(call):
    index, curse_num, plan_type = call.data.split('_')[1:]
    index = int(index)
    
    markup = telebot.types.InlineKeyboardMarkup()
    markup.add(telebot.types.InlineKeyboardButton("Неявка", callback_data=f"not_show_{index}_{curse_num}_{plan_type}"),
               telebot.types.InlineKeyboardButton("Не сдано", callback_data=f"not_submitted_{index}_{curse_num}_{plan_type}"))
    
    bot.send_message(call.message.chat.id, "Выберите причину:", reply_markup=markup)

def handle_document_number(message, index, curse_num, plan_type):
    document_number = message.text
    conn = connect_db()
    cursor = conn.cursor()

    table_name = 'plane_PB' if plan_type == 'PB' else 'plane_OT'
    cursor.execute(f"SELECT * FROM {table_name} LIMIT 1 OFFSET ?", (index,))
    plane_details = cursor.fetchone()

    if plane_details:
        employee_id, curse_num, date, start_date, end_date = plane_details
        
        cursor.execute(
            "INSERT INTO training_report_PB (ID_employee, curse_num, details_document, result, date, start_date, end_date) VALUES (?, ?, ?, 1,date('now'), ?, ?)",
            (employee_id, curse_num, document_number, start_date, end_date)
        )
        cursor.execute("SELECT full_name FROM employee WHERE ID = ?", (employee_id,))
        employee_name = cursor.fetchone()[0]
        conn.commit()
        bot.send_message(message.chat.id, f"Информация об обучении успешно внесена для {employee_name}.")
        
        ask_training(message.chat.id, fetch_training_plans(month=int(start_date.split('-')[1]), plan_type=plan_type), index + 1, plan_type)

    conn.close()

@bot.callback_query_handler(func=lambda call: call.data.startswith('not_'))
def handle_reason(call):
    reason_type = call.data.split('_')[0]
    index, curse_num, plan_type = call.data.split('_')[2:]
    index = int(index)

    conn = connect_db()
    cursor = conn.cursor()

    table_name = 'plane_PB' if plan_type == 'PB' else 'plane_OT'
    cursor.execute(f"SELECT * FROM {table_name} LIMIT 1 OFFSET ?", (index,))
    plane_details = cursor.fetchone()

    if plane_details:
        employee_id, curse_num, date, start_date, end_date = plane_details
        cursor.execute(
            "INSERT INTO training_report_PB (ID_employee, curse_num, details_document, result,date, start_date, end_date) VALUES (?, ?, NULL, 0,date('now'), ?, ?)",
            (employee_id, curse_num, start_date, end_date)
        )
        cursor.execute("SELECT full_name FROM employee WHERE ID = ?", (employee_id,))
        employee_name = cursor.fetchone()[0]
        conn.commit()
        
        reason_msg = "Неявка записана" if reason_type == "no_show" else "Не сдача записана"
        bot.send_message(call.message.chat.id, f"{reason_msg} для {employee_name}.")
        
        ask_training(call.message.chat.id, fetch_training_plans(month=int(start_date.split('-')[1]), plan_type=plan_type), index + 1, plan_type)

    conn.close()

if __name__ == "__main__":
    bot.polling(none_stop=True)
