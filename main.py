import pandas as pd
import datetime
import logging
import time
from telegram import Bot
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, CallbackQueryHandler, CallbackContext

# Настройки Telegram бота
TOKEN = '6865466938:AAFKS2iOI3w3JtJ-sLUo11n58kT-SGEr8GI'
CHAT_ID = '-698861002'
ALLOWED_USERS = [376492213, 250362710, 246280009]  # Замените на реальные идентификаторы пользователей

# Создаем объект бота
bot = Bot(token=TOKEN)

# Чтение Excel файла
df = pd.read_excel('software_expiry_dates.xlsx')

# Текущая дата
today = datetime.datetime.today()

# Настройка логгирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

DEFAULT_NOTIFICATION_PERIOD = 30  # Период уведомлений по умолчанию, если не задан другой

def is_user_allowed(update):
    user_id = update.message.from_user.id if update.message else update.callback_query.from_user.id
    return user_id in ALLOWED_USERS

def restricted(func):
    def wrapped(update, context, *args, **kwargs):
        if not is_user_allowed(update):
            update.message.reply_text("Вам не разрешено использовать этого бота.") if update.message else update.callback_query.message.reply_text("Вам не разрешено использовать этого бота.")
            return
        return func(update, context, *args, **kwargs)
    return wrapped

@restricted
def add_software_start(update, context):
    update.message.reply_text("Введите данные в формате: 'Имя продукта, Дата окончания (ДД.ММ.ГГГГ)'")
    context.user_data['expecting'] = 'add_software'

@restricted
def delete_software_start(update, context):
    update.message.reply_text("Введите имя продукта для удаления:")
    context.user_data['expecting'] = 'delete_software'

@restricted
def set_notification_period(update, context):
    try:
        days = int(context.args[0])
        if days > 0:
            context.user_data['notification_period'] = days
            update.message.reply_text(f"Период уведомлений установлен на {days} дней")
        else:
            update.message.reply_text("Пожалуйста, введите корректное количество дней.")
    except (IndexError, ValueError):
        update.message.reply_text("Пожалуйста, введите корректное количество дней.")

def check_expiry(context):
    global today
    today = datetime.datetime.today()
    for index, row in df.iterrows():
        product_name = row['Имя продукта']
        expiry_date = row['Дата окончания']

        if isinstance(expiry_date, pd.Timestamp):
            expiry_date = expiry_date.to_pydatetime()

        days_left = (expiry_date - today).days

        notification_period = context.user_data.get('notification_period', DEFAULT_NOTIFICATION_PERIOD)
        if days_left <= notification_period:
            message = f'Уведомление: {product_name} истекает через {days_left} дней ({expiry_date.strftime("%d.%m.%Y")})'
            try:
                bot.send_message(chat_id=CHAT_ID, text=message, timeout=120)
            except Exception as e:
                logger.error(f"Ошибка при отправке сообщения: {e}")

@restricted
def show_all_software(update, context):
    expiring_software = df[df['Дата окончания'].dt.year == today.year]

    if expiring_software.empty:
        update.message.reply_text('Нет софта, истекающего в этом году.') if update.message else update.callback_query.message.reply_text('Нет софта, истекающего в этом году.')
    else:
        for index, row in expiring_software.iterrows():
            product_name = row['Имя продукта']
            expiry_date = row['Дата окончания']
            days_left = (expiry_date - today).days
            message = f'{product_name} истекает {expiry_date.strftime("%d.%m.%Y")} ({days_left} дней осталось)'
            try:
                bot.send_message(chat_id=CHAT_ID, text=message, timeout=120)
            except telegram.error.TimedOut:
                logger.warning("Таймаут при отправке сообщения. Повторная попытка через 5 секунд...")
                time.sleep(5)
                bot.send_message(chat_id=CHAT_ID, text=message, timeout=30)
            except Exception as e:
                logger.error(f"Ошибка при отправке сообщения: {e}")

# Настройка команд и кнопок
updater = Updater(TOKEN, use_context=True)
dispatcher = updater.dispatcher
dispatcher.add_handler(CommandHandler('soft_this_month', restricted(lambda update, context: check_expiring_software(context, month_offset=0))))
dispatcher.add_handler(CommandHandler('soft_next_month', restricted(lambda update, context: check_expiring_software(context, month_offset=1))))
dispatcher.add_handler(CommandHandler('soft_all', restricted(show_all_software)))
dispatcher.add_handler(CommandHandler('add_software', restricted(add_software_start)))
dispatcher.add_handler(CommandHandler('delete_software', restricted(delete_software_start)))
dispatcher.add_handler(CommandHandler('set_notification_period', restricted(set_notification_period), pass_args=True, pass_update_queue=True, pass_job_queue=True, pass_user_data=True))
dispatcher.add_handler(CommandHandler('check_expiry', restricted(lambda update, context: check_expiry(context))))

# Запуск бота
updater.start_polling()
updater.idle()
