from datetime import datetime, timedelta
import json
import os
import threading
import time
import pandas as pd
import pytz
import telebot
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton

from Templates_1_tg import create_msg_file
from excel_handler import create_pivot_table, extract_data
from sales_analysis import get_trend

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


def get_message(company, type, trend):
    keyboard = InlineKeyboardMarkup()
    msg = f'👤 <b>Контрагент:</b> {company}\n' \
          f'📋 <b>Тип:</b> {type}\n'
    if trend == -1:
        msg += f'📉 <b>Состояние:</b> Падение объема закупок\n' \
               f'🛎️ <b>Рекомендация:</b> Свяжитесь с клиентом, уточните причины снижения и обсудите индивидуальные условия.'
        keyboard.add(InlineKeyboardButton("Письмо клиенту", callback_data="create_volume_down"))
        keyboard.add(InlineKeyboardButton("Все данные", callback_data="excel_down"))
    elif trend == 0:
        msg += f'📊 <b>Состояние:</b> Объем закупок стабильный\n' \
               f'🛎️ <b>Рекомендация:</b> Стабильность - признак мастерства.'
    elif trend == 1:
        msg += f'📈 <b>Состояние:</b> Рост объема закупок\n' \
               f'🛎️ <b>Рекомендация:</b> Отметьте рост объёмов и обсудите с клиентом эксклюзивные предложения.'
        keyboard.add(InlineKeyboardButton("Письмо клиенту", callback_data="create_volume_up"))
        keyboard.add(InlineKeyboardButton("Все данные", callback_data="excel_up"))
    elif trend == -2:
        msg += f'⏰ <b>Состояние:</b> Есть неоплаченные счета\n' \
               f'🛎️ <b>Рекомендация:</b> Уточните причину задержки оплаты.'
        keyboard.add(InlineKeyboardButton("Письмо клиенту", callback_data="create_unpaid"))
        keyboard.add(InlineKeyboardButton("Все данные", callback_data="excel_not_sale"))
    keyboard.add(InlineKeyboardButton("Обработан", callback_data="skip"))
    return msg, keyboard


def get_not_sale(not_sale_df):
    now = datetime.now()
    some_weeks_later = now + timedelta(weeks=1)

    upcoming_events = not_sale_df[
        (not_sale_df['Дата отгрузки (отправки)'] >= pd.to_datetime(now)) &
        (not_sale_df['Дата отгрузки (отправки)'] <= pd.to_datetime(some_weeks_later))
    ]
    if len(upcoming_events.index) > 0:
        return True
    else:
        return False


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


@bot.message_handler(commands=['summary'])
def send_summary(message):
    pivot = create_pivot_table(excel_data)
    df = add_telegram_id_to_df(pivot)
    df = df[df['telegram_id'] == str(message.from_user.id)]
    df = get_trend(df)
    for _, row in df.iterrows():
        company_df = excel_data[excel_data['Контрагент'] == row['Контрагент']]
        type_company = company_df['Тип контрагента'].iloc[0]
        not_sale_df = company_df[company_df['Оплачено'] == 0]
        if get_not_sale(not_sale_df):
            msg, keyboard = get_message(row['Контрагент'], type_company, -2)
            bot.send_message(message.chat.id, msg, parse_mode='HTML', reply_markup=keyboard)
        elif isinstance(row['Тренд'], int):
            msg, keyboard = get_message(row['Контрагент'], type_company, row['Тренд'])
            bot.send_message(message.chat.id, msg, parse_mode='HTML', reply_markup=keyboard)


@bot.callback_query_handler(func=lambda call: True)
def handle_callback(call):
    """Обработчик нажатий на инлайн-кнопки"""
    chat_id = call.message.chat.id

    try:
        bot.answer_callback_query(call.id)
        if call.data == "create_msg":
            letter_type = "shipping"
        elif call.data == "create_unpaid":
            letter_type = "unpaid"
        elif call.data == "create_volume_down":
            letter_type = "volume_down"
        elif call.data == "create_volume_up":
            letter_type = "volume_up"
        elif call.data == "create_overdue":
            letter_type = "overdue"
        elif call.data == "skip":
            bot.delete_message(chat_id=call.message.chat.id, message_id=call.message.message_id)
            return
        elif call.data == "excel_up":
            file_path = 'test_data_template.xlsx'
            try:
                with open(file_path, 'rb') as file:
                    bot.send_document(chat_id=call.message.chat.id, document=file)
            except Exception as e:
                bot.send_message(call.message.chat.id, "Произошла ошибка при отправке файла.")
            return
        elif call.data == "excel_down":
            file_path = 'test_data_template.xlsx'
            try:
                with open(file_path, 'rb') as file:
                    bot.send_document(chat_id=call.message.chat.id, document=file)
            except Exception as e:
                bot.send_message(call.message.chat.id, "Произошла ошибка при отправке файла.")
            return
        elif call.data == "excel_not_sale":
            file_path = 'test_data_template.xlsx'
            try:
                with open(file_path, 'rb') as file:
                    bot.send_document(chat_id=call.message.chat.id, document=file)
            except Exception as e:
                bot.send_message(call.message.chat.id, "Произошла ошибка при отправке файла.")
            return

        # Параметры для письма
        params = {
            "Контрагент": "ООО Ромашка",
            "Менеджер": "Иван Иванов",
            "Договор_контрагента": "Договор продажи № 290 от 14.04.2025",
            "letter_type": letter_type
        }
        if letter_type == "overdue":
            params["Спецификации"] = "Спецификация № 123"
            params["Дата_отгрузки"] = "2025-07-22"

        # Вызов функции с тестовыми данными
        file_path = create_msg_file(**params)
        if file_path:
            try:
                with open(file_path, 'rb') as file:
                    bot.send_document(chat_id=chat_id, document=file,
                                      caption=f"{os.path.basename(file_path)}")
            except Exception as e:
                print(f"Ошибка при отправке в Telegram: {str(e)}")
                bot.send_message(chat_id, "Произошла ошибка при отправке файла.")
        else:
            bot.send_message(chat_id, "Не удалось создать файл.")
    except Exception as e:
        print(f"Ошибка в handle_callback: {str(e)}")
        bot.send_message(chat_id, f"Произошла ошибка: {str(e)}")


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
        # bg_thread = threading.Thread(target=background_task)
        # bg_thread.daemon = True
        # bg_thread.start()
        print("Бот запущен...")
        bot.infinity_polling()
    except:
        time.sleep(5)
        main()


if __name__ == '__main__':
    main()
