import telebot
from apscheduler.schedulers.background import BackgroundScheduler
import logging
import global_algoritm
from datetime import datetime





# Определите вашу функцию
def send_yearly_message():
    print("Отправляем годовое сообщение!")

# Создайте экземпляр планировщика
scheduler = BackgroundScheduler()

# Добавьте задачу без круглых скобок
scheduler.add_job(global_algoritm.update, 'cron', month='9', day='29', hour='11', minute='48')

# Запустите планировщик
scheduler.start()

# Приложение будет работать, пока его не остановят
try:
    while True:
        pass  # Держите приложение запущенным
except (KeyboardInterrupt, SystemExit):
    scheduler.shutdown()