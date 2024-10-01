import telebot
from apscheduler.schedulers.background import BackgroundScheduler
import logging
from datetime import datetime
import global_algoritm

# Включаем логирование
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

# Токен вашего бота
TOKEN = '7526701581:AAEj5YAQJ8B-_VaKAEaGjdtw4ckL_aA1u-A'
bot = telebot.TeleBot(TOKEN)

# Список чатов (замените на ваши ID)
chat_ids = ['5211807364']  # Замените на реальные ID чатов

# Функция, отправляющая сообщение
def send_monthly_message():
    message = "Это ежемесячное сообщение!"
    for chat_id in chat_ids:
        try:
            bot.send_message(chat_id, message)
        except Exception as e:
            logger.error(f"Ошибка при отправке сообщения в чат {chat_id}: {e}")

# Команда /start
@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Привет! Я бот для рассылки.")

def main():
    # Устанавливаем расписление
    scheduler = BackgroundScheduler()
    # Рассылка будет происходить 1-го числа каждого месяца в 10:00
    scheduler.add_job(send_monthly_message, 'cron', day='1', hour='10', minute = '0')
    scheduler.add_job(global_algoritm.update, 'cron', month='1', day='1', hour='0', minute='0')
    scheduler.start()

    # Запускаем бота
    bot.polling()

if __name__ == '__main__':
    main()
