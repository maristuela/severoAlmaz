import telebot
from telebot import types

# Замените 'YOUR_API_TOKEN' на токен вашего бота
API_TOKEN = '7526701581:AAEj5YAQJ8B-_VaKAEaGjdtw4ckL_aA1u-A'
bot = telebot.TeleBot(API_TOKEN)

# Начальная команда
@bot.message_handler(commands=['start'])
def send_welcome(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    button = types.KeyboardButton("Поделиться номером", request_contact=True)
    markup.add(button)
    bot.send_message(message.chat.id, "Привет! Нажми на кнопку ниже, чтобы поделиться своим номером:", reply_markup=markup)

# Обработка получения контакта
@bot.message_handler(content_types=['contact'])
def handle_contact(message):
    phone_number = message.contact.phone_number
    bot.send_message(message.chat.id, f"Спасибо! Ваш номер телефона: {phone_number}")

# Запуск бота
bot.polling(none_stop=True)
