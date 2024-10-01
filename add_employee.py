import sqlite3
import telebot
from telebot import types

API_TOKEN = '7898118766:AAFvT99mQjvmKdpe6XZuYhI-zuFaIdYKdqk'
bot = telebot.TeleBot(API_TOKEN)

user_states = {}

def get_professions():
    conn = sqlite3.connect('BdTrainingCenter.db')
    cursor = conn.cursor()
    cursor.execute("SELECT ID, profession FROM code_profession")
    professions = cursor.fetchall()
    conn.close()
    return professions

def add_employee(full_name, profession_id, division, email):
    conn = sqlite3.connect('BdTrainingCenter.db')
    cursor = conn.cursor()

    cursor.execute("SELECT MAX(ID) FROM employee")
    last_id = cursor.fetchone()[0]

    if last_id is None:
        last_id = 0

    new_id = last_id + 1

    cursor.execute("INSERT INTO employee (ID, full_name, profession, division, email) VALUES (?, ?, ?, ?, ?)",
                   (new_id, full_name, profession_id, division, email))
    
    conn.commit()
    conn.close()
    
    return new_id

# Start adding a new employee
@bot.message_handler(commands=['add_employee'])
def start_add_employee(message):
    user_states[message.chat.id] = {'step': 1}
    bot.reply_to(message, "Введите ФИО сотрудника")

@bot.message_handler(func=lambda message: message.chat.id in user_states)
def handle_employee_info(message):
    state = user_states[message.chat.id]

    if state['step'] == 1:
        state['full_name'] = message.text
        
        # Fetch professions and create buttons
        professions = get_professions()
        markup = types.InlineKeyboardMarkup()
        
        for prof_id, prof_name in professions:
            markup.add(types.InlineKeyboardButton(prof_name, callback_data=f'profession_{prof_id}'))
        
        bot.send_message(message.chat.id, "Выберите профессию сотрудника", reply_markup=markup)

    elif state['step'] == 2:
        state['division'] = 1
        state['email'] = message.text
        new_id = add_employee(state['full_name'], state['profession_id'], state['division'], state['email'])
        
        bot.reply_to(message, f"Новый сотрудник добавлен с ID: {new_id}")
        del user_states[message.chat.id] 

@bot.callback_query_handler(func=lambda call: call.data.startswith('profession_'))
def handle_profession_selection(call):
    chat_id = call.message.chat.id
    prof_id = call.data.split('_')[1]
    user_states[chat_id]['profession_id'] = int(prof_id)
    bot.answer_callback_query(call.id)
    bot.send_message(chat_id, "Профессия выбрана! \nВведите email:")
    user_states[chat_id]['step'] = 2

bot.polling()
