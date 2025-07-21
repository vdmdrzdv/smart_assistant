from datetime import datetime, timedelta
import json
import os
import threading
import time
import pandas as pd
import pytz
import telebot

from excel_handler import extract_data

USERS_FILE = 'users_data.json'
TIMEZONE = pytz.timezone('Asia/Yekaterinburg')
TOKEN = '7955516321:AAGKWegG3O70jPCVk_3cQrw5wcHrfA_27o4'
USER_STATES = {}
bot = telebot.TeleBot(TOKEN)
running = True
excel_data = []


class UserState:
    WAITING_FOR_LAST_NAME = 1
    WAITING_FOR_FIRST_NAME = 2
    WAITING_FOR_MIDDLE_NAME = 3


def load_users_data():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, 'r', encoding='utf-8') as file:
            return json.load(file)
    return {}


def save_users_data(data):
    with open(USERS_FILE, 'w', encoding='utf-8') as file:
        json.dump(data, file, ensure_ascii=False, indent=4)


def check_upcoming_events():
    """Проверяет события и отправляет уведомления"""
    now = datetime.now()
    some_weeks_later = now + timedelta(weeks=2)

    upcoming_events = excel_data[
        (excel_data['Дата отгрузки (отправки)'] >= pd.to_datetime(now)) &
        (excel_data['Дата отгрузки (отправки)'] <= pd.to_datetime(some_weeks_later))
    ]

    for _, event in upcoming_events.iterrows():
        try:
            if event['telegram_id'] is None:
                continue
            user_id = str(event['telegram_id'])

            date = event['Дата отгрузки (отправки)'].strftime('%d.%m.%Y')
            order = event['Номер договора']
            counterparty = event['Контрагент']

            message = (
                f"⏰ Информацию по срокам поставки {order} {counterparty} на {now.day}.{now.month}.{now.year}!\n"
                f"{date} ожидается отгрузка товара.\n"
                f"Дебеторская задолжность - {'Нет' if event['Дней просрочки'] == 0 else 'Да'}.\n"
                f"Оплата - {'Нет' if event['Оплачено'] == 0 else 'Да'}.\n"
                f"Необходимо проверить статус заказа."
            )
            if event['Дней просрочки'] > 0:
                message += "\n Внимание! Контрагент имеет задолженность."
            bot.send_message(user_id, message)

        except Exception as e:
            print(
                f"Ошибка при отправке уведомления пользователю {user_id}: {e}")


def background_task():
    """Фоновая задача, проверяющая события"""
    while running:
        try:
            print(f"[{datetime.now(TIMEZONE)}] Проверяю предстоящие события...")
            check_upcoming_events()
            time.sleep(15)

        except Exception as e:
            print(f"Ошибка в фоновом потоке: {e}")
            time.sleep(60)


@bot.message_handler(commands=['start'])
def handle_start(message):
    user_id = str(message.from_user.id)
    users_data = load_users_data()

    if user_id in users_data:
        bot.send_message(
            message.chat.id,
            f"С возвращением, {users_data[user_id]['last_name']} "
            f"{users_data[user_id]['first_name']} "
            f"{users_data[user_id].get('middle_name', '')}!"
        )
    else:
        USER_STATES[user_id] = {'state': UserState.WAITING_FOR_LAST_NAME}
        bot.send_message(
            message.chat.id,
            "Добро пожаловать! Для регистрации нам нужно узнать ваши данные.\n"
            "Пожалуйста, введите вашу фамилию:"
        )


@bot.message_handler(commands=['help'])
def send_help(message):
    help_text = """
        Доступные команды:
        /start - начать работу с ботом
        /help - получить справку
    """
    bot.send_message(message.chat.id, help_text)


@bot.callback_query_handler(func=lambda call: True)
def callback_inline(call):
    if call.data == "test":
        bot.send_message(call.message.chat.id, "Вы нажали inline кнопку!")


@bot.message_handler(content_types=['text'])
def handle_text(message):
    user_id = str(message.from_user.id)

    # Пропускаем сообщения, если пользователь не в процессе регистрации
    if user_id not in USER_STATES:
        return

    current_state = USER_STATES[user_id]['state']
    text = message.text.strip()

    if current_state == UserState.WAITING_FOR_LAST_NAME:
        USER_STATES[user_id]['last_name'] = text
        USER_STATES[user_id]['state'] = UserState.WAITING_FOR_FIRST_NAME
        bot.send_message(message.chat.id, "Теперь введите ваше имя:")

    elif current_state == UserState.WAITING_FOR_FIRST_NAME:
        USER_STATES[user_id]['first_name'] = text
        USER_STATES[user_id]['state'] = UserState.WAITING_FOR_MIDDLE_NAME
        bot.send_message(
            message.chat.id,
            "Введите ваше отчество (если нет, отправьте '-'):"
        )

    elif current_state == UserState.WAITING_FOR_MIDDLE_NAME:
        middle_name = text if text != '-' else ''

        # Сохраняем данные пользователя
        users_data = load_users_data()
        users_data[user_id] = {
            'last_name': USER_STATES[user_id]['last_name'],
            'first_name': USER_STATES[user_id]['first_name'],
            'middle_name': middle_name,
            'username': message.from_user.username,
            'registration_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        save_users_data(users_data)

        # Удаляем состояние пользователя
        del USER_STATES[user_id]

        # Формируем приветственное сообщение
        full_name = f"{users_data[user_id]['last_name']} " \
                    f"{users_data[user_id]['first_name']}"
        if middle_name:
            full_name += f" {users_data[user_id]['middle_name']}"

        bot.send_message(
            message.chat.id,
            f"Спасибо, {full_name}! Регистрация завершена.\n"
            f"Теперь вы можете пользоваться ботом."
        )


# Функция для сопоставления ФИО и добавления Telegram ID
def add_telegram_id_to_df(df):
    users_data = load_users_data()
    fio_to_id = {}
    for user_id, user_info in users_data.items():
        last_name = user_info.get('last_name', '')
        first_name = user_info.get('first_name', '')
        middle_name = user_info.get('middle_name', '')
        fio = f"{last_name} {first_name} {middle_name}".strip()
        fio_to_id[fio] = user_id

    def find_telegram_id(fio):
        if fio in fio_to_id:
            return fio_to_id[fio]
        return None

    df['telegram_id'] = df['Менеджер'].apply(find_telegram_id)
    return df


def main():
    try:
        if not os.path.exists(USERS_FILE):
            with open(USERS_FILE, 'w', encoding='utf-8') as file:
                json.dump({}, file)

        df = extract_data('data.xlsx')
        global excel_data
        excel_data = add_telegram_id_to_df(df)
        bg_thread = threading.Thread(target=background_task)
        bg_thread.daemon = True
        bg_thread.start()
        print("Бот запущен...")
        bot.infinity_polling()
    except:
        time.sleep(5)
        main()


if __name__ == '__main__':
    main()
