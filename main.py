from datetime import datetime, timedelta
import json
import os
import threading
import re
import time
import pandas as pd
import pytz
import telebot
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
from telebot.types import ReplyKeyboardMarkup, KeyboardButton


from Templates_1_tg import clean_filename, create_msg_file
from excel_handler import create_pivot_table, extract_data
from sales_analysis import get_trend
from email_sender import send_email

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
            time.sleep(60)

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

        # Изменение: добавлен company в callback_data для передачи Контрагент (Садовников Николай Сергеевич)
        keyboard.add(InlineKeyboardButton("Письмо клиенту", callback_data=f"create_volume_down|{company}"))
        keyboard.add(InlineKeyboardButton("Все данные", callback_data=f"excel_down|{company}"))
    elif trend == 0:
        msg += f'📊 <b>Состояние:</b> Объем закупок стабильный\n' \
               f'🛎️ <b>Рекомендация:</b> Стабильность - признак мастерства.'
    elif trend == 1:
        msg += f'📈 <b>Состояние:</b> Рост объема закупок\n' \
               f'🛎️ <b>Рекомендация:</b> Отметьте рост объёмов и обсудите с клиентом эксклюзивные предложения.'

        # Изменение: добавлен company в callback_data для передачи Контрагент (Садовников Николай Сергеевич)
        keyboard.add(InlineKeyboardButton("Письмо клиенту", callback_data=f"create_volume_up|{company}"))
        keyboard.add(InlineKeyboardButton("Все данные", callback_data=f"excel_up|{company}"))
    elif trend == -2:
        msg += f'⏰ <b>Состояние:</b> Есть неоплаченные счета\n' \
               f'🛎️ <b>Рекомендация:</b> Уточните причину задержки оплаты.'

        # Изменение: добавлен company в callback_data для передачи Контрагент (Садовников Николай Сергеевич)
        keyboard.add(InlineKeyboardButton("Письмо клиенту", callback_data=f"create_unpaid|{company}"))
        keyboard.add(InlineKeyboardButton("Все данные", callback_data=f"excel_not_sale|{company}"))
    keyboard.add(InlineKeyboardButton("Обработан", callback_data=f"skip|{company}"))
    return msg, keyboard


def get_main_keyboard():
    keyboard = ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    keyboard.add(
        KeyboardButton("📊 Сводка по контрагентам"),
        KeyboardButton("ℹ️ Помощь"),
        KeyboardButton("💰 Задолженности")
    )
    return keyboard


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
            f"{users_data[user_id].get('middle_name', '')}!",
            reply_markup=get_main_keyboard()
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
        /summary - получить сводку по контрагентам
        /debt - узнать ДЗ по контрагентам
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
            send_email("dlyashkolisusu@gmail.com", "Данные контрагента", msg)
        elif isinstance(row['Тренд'], int):
            msg, keyboard = get_message(row['Контрагент'], type_company, row['Тренд'])
            bot.send_message(message.chat.id, msg, parse_mode='HTML', reply_markup=keyboard)
            send_email("dlyashkolisusu@gmail.com", "Данные контрагента", msg)


@bot.message_handler(commands=['debt'])
def send_debt(message):
    user_id = str(message.from_user.id)
    df = excel_data.copy()

    # Фильтруем по Telegram ID
    df = df[df['telegram_id'] == user_id]

    # Фильтруем строки, где ДЗ > 0
    debtors = df[df['ДЗ'] > 0]

    if debtors.empty:
        bot.send_message(message.chat.id, "✅ У ваших контрагентов нет задолженности.")
        return

    # Группируем по контрагенту и выводим суммы задолженностей
    grouped = (
        debtors.groupby(['Контрагент', 'Тип контрагента'], as_index=False)['ДЗ']
        .sum()
    )

    for _, row in grouped.iterrows():
        msg = (
            f"👤 <b>Контрагент:</b> {row['Контрагент']}\n"
            f"📋 <b>Тип:</b> {row['Тип контрагента']}\n"
            f"💰 <b>Задолженность:</b> {row['ДЗ']:.2f} руб.\n"
            f"🛎️ <b>Рекомендация:</b> Cвязаться с клиентом для уточнения сроков оплаты."
        )
        bot.send_message(message.chat.id, msg, parse_mode='HTML')


@bot.callback_query_handler(func=lambda call: True)
def handle_callback(call):
    """Обработчик нажатий на инлайн-кнопки"""
    chat_id = call.message.chat.id

    try:
        bot.answer_callback_query(call.id)
        # Изменение: Извлечение Контрагент из callback_data вместо текста сообщения (Садовников Николай Сергеевич)
        callback_parts = call.data.split('|')
        action = callback_parts[0]
        Контрагент = callback_parts[1] if len(callback_parts) > 1 else None

        if not Контрагент:
            bot.send_message(chat_id, "Ошибка: Контрагент не указан.")
            return

        # Изменение: Фильтрация excel_data для получения данных по Контрагент (Садовников Николай Сергеевич)
        company_df = excel_data[excel_data['Контрагент'] == Контрагент]
        if company_df.empty:
            bot.send_message(chat_id, f"Данные для контрагента {Контрагент} не найдены.")
            return

        # Изменение: Извлечение параметров из excel_data для передачи в create_msg_file (Садовников Николай Сергеевич)
        row = company_df.iloc[0]
        Менеджер = row['Менеджер']
        Договор_контрагента = row['Номер договора']
        Дата_отгрузки = row['Дата отгрузки (отправки)'].strftime(
            '%d.%m.%Y') if pd.notna(row['Дата отгрузки (отправки)']) else ""
        Спецификации = "Спецификация № Unknown"  # Заглушка, так как Спецификации отсутствует в excel_data

        # Изменение: Создание временного Excel-файла с данными сводной таблицы для Контрагент только для писем (Садовников Николай Сергеевич)
        temp_file_path = os.path.join(
            os.getcwd(), f"temp_data_{clean_filename(Контрагент)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        pivot_table = create_pivot_table(excel_data)
        company_pivot = pivot_table[pivot_table['Контрагент'] == Контрагент]
        company_pivot.to_excel(temp_file_path, index=False)

        # Обработка действий callback
        if action == "create_msg":
            letter_type = "shipping"
        elif action == "create_unpaid":
            letter_type = "unpaid"
        elif action == "create_volume_down":
            letter_type = "volume_down"
        elif action == "create_volume_up":
            letter_type = "volume_up"
        elif action == "create_overdue":
            letter_type = "overdue"
        elif action == "skip":
            bot.delete_message(chat_id=call.message.chat.id, message_id=call.message.message_id)
            return
        elif action == "excel_up":
            file_path = 'test_data_template.xlsx'
            try:
                with open(file_path, 'rb') as file:
                    bot.send_document(chat_id=call.message.chat.id, document=file)
            except Exception as e:
                bot.send_message(call.message.chat.id, "Произошла ошибка при отправке файла.")
            return
        elif action == "excel_down":
            file_path = 'test_data_template.xlsx'
            try:
                with open(file_path, 'rb') as file:
                    bot.send_document(chat_id=call.message.chat.id, document=file)
            except Exception as e:
                bot.send_message(call.message.chat.id, "Произошла ошибка при отправке файла.")
            return
        elif action == "excel_not_sale":
            file_path = 'test_data_template.xlsx'
            try:
                with open(file_path, 'rb') as file:
                    bot.send_document(chat_id=call.message.chat.id, document=file)
            except Exception as e:
                bot.send_message(call.message.chat.id, "Произошла ошибка при отправке файла.")
            return

        # Параметры для письма
        params = {
            "Контрагент": Контрагент,
            "Менеджер": Менеджер,
            "Договор_контрагента": Договор_контрагента,
            "letter_type": letter_type,
            "excel_file_path": temp_file_path
        }
        if letter_type == "overdue":
            params["Спецификации"] = Спецификации
            params["Дата_отгрузки"] = Дата_отгрузки

        # Изменение: Вызов create_msg_file с параметрами из excel_data (Садовников Николай Сергеевич)
        file_path = create_msg_file(**params)
        if file_path:
            try:
                with open(file_path, 'rb') as file:
                    bot.send_document(chat_id=chat_id, document=file,
                                      caption=f"{os.path.basename(file_path)}")
                    # Изменение: Очистка временного Excel-файла после отправки (Садовников Николай Сергеевич)
                    os.remove(temp_file_path)
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

    text = message.text.strip()

    # --- Обработка кастомных кнопок ---
    if text == "📊 Сводка по контрагентам":
        send_summary(message)
        return
    elif text == "ℹ️ Помощь":
        send_help(message)
        return
    elif text == "💰 Задолженности":
        send_debt(message)
        return

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
            f"Теперь вы можете пользоваться ботом.",
            reply_markup=get_main_keyboard()
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
