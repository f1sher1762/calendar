import pandas as pd
import datetime
import logging
import time
import os
from telegram import Bot
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, CallbackQueryHandler, CallbackContext, JobQueue

# Настройки Telegram бота
TOKEN = '6865466938:AAFKS2iOI3w3JtJ-sLUo11n58kT-SGEr8GI'
CHAT_ID = '-4157087994'
ALLOWED_USERS = [376492213]  # Замените на реальные идентификаторы пользователей


# Создаем объект бота
bot = Bot(token=TOKEN)

# Чтение Excel файла
df = pd.read_excel('software_expiry_dates.xlsx')

# Текущая дата
today = datetime.datetime.today()

# Настройка логгирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Файл для хранения дат отправленных уведомлений
notified_dates_file = 'notified_dates.txt'

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

def log_message(user_id):
    pass

def add_software_start(update, context):
    update.message.reply_text("Введите данные в формате: 'Имя продукта, Дата окончания (ДД.ММ.ГГГГ)'")
    context.user_data['expecting'] = 'add_software'

def add_software(update, context):
    if 'expecting' in context.user_data and context.user_data['expecting'] == 'add_software':
        try:
            text = update.message.text
            product_name, expiry_date_str = text.split(',')
            expiry_date = datetime.datetime.strptime(expiry_date_str.strip(), '%d.%m.%Y')
            new_row = pd.DataFrame([[product_name.strip(), expiry_date]], columns=['Имя продукта', 'Дата окончания'])
            global df
            df = pd.concat([df, new_row], ignore_index=True)
            df.to_excel('software_expiry_dates.xlsx', index=False)
            update.message.reply_text(f"Запись '{product_name.strip()}' с датой окончания {expiry_date_str.strip()} добавлена.")
            context.user_data['expecting'] = None
        except Exception as e:
            update.message.reply_text("Произошла ошибка при добавлении записи. Убедитесь, что формат данных правильный.")
    elif 'expecting' in context.user_data and context.user_data['expecting'] == 'delete_software':
        delete_software(update, context)

def delete_software_start(update, context):
    update.message.reply_text("Введите имя продукта, который нужно удалить:")
    context.user_data['expecting'] = 'delete_software'

def delete_software(update, context):
    if 'expecting' in context.user_data and context.user_data['expecting'] == 'delete_software':
        product_name = update.message.text.strip()
        global df
        if product_name in df['Имя продукта'].values:
            df = df[df['Имя продукта'] != product_name]
            df.to_excel('software_expiry_dates.xlsx', index=False)
            update.message.reply_text(f"Запись '{product_name}' удалена.")
            context.user_data['expecting'] = None
        else:
            update.message.reply_text(f"Запись '{product_name}' не найдена.")
            context.user_data['expecting'] = None

@restricted
def button(update, context):
    query = update.callback_query
    query.answer()

    if query.data == 'this_month':
        check_expiring_software(query, context, month_offset=0)
    elif query.data == 'next_month':
        check_expiring_software(query, context, month_offset=1)
    elif query.data == 'all_software':
        show_all_software(query, context)

@restricted
def check_expiring_software(update, context, month_offset=0):
    now = datetime.datetime.now()
    month_start = datetime.datetime(now.year, now.month + month_offset, 1)
    next_month_start = datetime.datetime(now.year, now.month + month_offset + 1, 1)

    expiring_software = df[(df['Дата окончания'] >= month_start) & (df['Дата окончания'] < next_month_start)]

    if expiring_software.empty:
        update.message.reply_text('Нет софта, истекающего в выбранный период.') if update.message else update.callback_query.message.reply_text('Нет софта, истекающего в выбранный период.')
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

def check_expiry(context: CallbackContext):
    global today
    today = datetime.datetime.today()
    notified_dates = load_notified_dates()

    for index, row in df.iterrows():
        product_name = row['Имя продукта']
        expiry_date = row['Дата окончания']

        if isinstance(expiry_date, pd.Timestamp):
            expiry_date = expiry_date.to_pydatetime()

        days_left = (expiry_date - today).days

        if days_left <= 56 and expiry_date not in notified_dates:
            message = f'Уведомление: {product_name} истекает через {days_left} дней ({expiry_date.strftime("%d.%m.%Y")})'
            try:
                bot.send_message(chat_id=CHAT_ID, text=message, timeout=120)
                notified_dates.append(expiry_date)
                save_notified_dates(notified_dates)
            except telegram.error.TimedOut:
                logger.warning("Таймаут при отправке сообщения. Повторная попытка через 5 секунд...")
                time.sleep(5)
                bot.send_message(chat_id=CHAT_ID, text=message, timeout=30)
            except Exception as e:
                logger.error(f"Ошибка при отправке сообщения: {e}")

def load_notified_dates():
    if os.path.exists(notified_dates_file):
        with open(notified_dates_file, 'r') as f:
            dates = f.read().splitlines()
            return [datetime.datetime.strptime(date, '%Y-%m-%d') for date in dates]
    return []

def save_notified_dates(dates):
    with open(notified_dates_file, 'w') as f:
        for date in dates:
            f.write(date.strftime('%Y-%m-%d') + '\n')

# Настройка команд и кнопок
updater = Updater(TOKEN, use_context=True)
dispatcher = updater.dispatcher
dispatcher.add_handler(CommandHandler('soft_this_month', lambda update, context: check_expiring_software(update, context, month_offset=0)))
dispatcher.add_handler(CommandHandler('soft_next_month', lambda update, context: check_expiring_software(update, context, month_offset=1)))
dispatcher.add_handler(CommandHandler('all_software', show_all_software))
dispatcher.add_handler(CommandHandler('add_software', add_software_start))
dispatcher.add_handler(CommandHandler('delete_software', delete_software_start))
dispatcher.add_handler(CallbackQueryHandler(button))
dispatcher.add_handler(MessageHandler(Filters.text & ~Filters.command, add_software))

# Настройка JobQueue для ежедневной проверки
job_queue = updater.job_queue
job_queue.run_daily(check_expiry, time=datetime.time(hour=9, minute=0, second=0))

# Запуск бота
updater.start_polling()
updater.idle()