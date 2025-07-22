import win32com.client
import pythoncom
import os
from datetime import datetime, date
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
import re


def clean_filename(filename):
    invalid_chars = r'[<>:"/\\|?*]'
    return re.sub(invalid_chars, '_', filename)


def create_msg_file(Контрагент="Клиент", Менеджер="Имя Менеджера", Договор_контрагента="", Спецификации="", Дата_отгрузки="", save_path="C:\\Helper\\emails", letter_type="shipping"):
    try:
        # Инициализация COM в текущем потоке
        pythoncom.CoInitialize()

        # Подключение к Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 означает создание нового письма

        # Текущая дата для письма о просрочке
        current_date = date.today().strftime("%d.%m.%Y")

        # Установка параметров письма в зависимости от типа
        if letter_type == "shipping":
            mail.Subject = f"Уведомление от ТМК о готовности к отгрузке для {Контрагент}"
            mail.To = f"{Контрагент}@example.com"
            mail.HTMLBody = f"""
            Уважаемый клиент {Контрагент},<br><br>
            Настоящим уведомляем Вас о том, что по {Договор_контрагента}, продукция готова к отгрузке. <br>
            
            <br><br>
            С уважением,<br>
            Ваш менеджер по продажам,<br>
            {Менеджер}<br>
            Трубная металлургическая компания
            """
            file_prefix = "Уведомление_отгрузка_ТМК"
            attachment_path = 'test_data_template.xlsx'
            if os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
            else:
                print(f"Файл {attachment_path} не найден")
        elif letter_type == "unpaid":
            mail.Subject = f"Уведомление от ТМК о неоплаченных счетах для {Контрагент}"
            mail.To = f"{Контрагент}@example.com"
            mail.HTMLBody = f"""
            Уважаемый клиент {Контрагент},<br><br>
            Настоящим уведомляем Вас о том, что по {Договор_контрагента}, у вас есть неоплаченные счета. <br>
            Мы готовы рассмотреть удобный формат урегулирования ситуации и предложить варианты реструктуризации задолженности при необходимости.
            <br><br>
            С уважением,<br>
            Ваш менеджер по продажам,<br>
            {Менеджер}<br>
            Трубная металлургическая компания
            """
            file_prefix = "Уведомление_неоплаченные_счета_ТМК"
            attachment_path = 'test_data_template.xlsx'
            if os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
            else:
                print(f"Файл {attachment_path} не найден")
        elif letter_type == "volume_down":
            mail.Subject = f"Уведомление от ТМК о падении объема закупок для {Контрагент}"
            mail.To = f"{Контрагент}@example.com"
            mail.HTMLBody = f"""
            Уважаемый клиент {Контрагент},<br><br>
            Обратили внимание на снижение объема ваших закупок в последнее время. Хотел бы уточнить, есть ли сложности или причины, с которыми мы можем помочь.
            Мы готовы предложить вам специальные условия или провести совместную сессию, чтобы обсудить возможные решения и перспективы дальнейшего сотрудничества
            <br><br>
            С уважением,<br>
            Ваш менеджер по продажам,<br>
            {Менеджер}<br>
            Трубная металлургическая компания
            """
            file_prefix = "Уведомление_падение_объема_закупок_ТМК"
            attachment_path = 'test_data_template.xlsx'
            if os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
            else:
                print(f"Файл {attachment_path} не найден")
        elif letter_type == "volume_up":
            mail.Subject = f"Уведомление от ТМК о росте объема закупок для {Контрагент}"
            mail.To = f"{Контрагент}@example.com"
            mail.HTMLBody = f"""
            Уважаемый клиент {Контрагент},<br><br>
            Рады отметить, что объем ваших закупок за последний период значительно вырос. Благодарим вас за доверие!
            Мы хотим предложить вам дополнительные условия сотрудничества, персональные предложения и расширение ассортимента.
            <br><br>
            С уважением,<br>
            Ваш менеджер по продажам,<br>
            {Менеджер}<br>
            Трубная металлургическая компания
            """
            file_prefix = "Уведомление_рост_объема_закупок_ТМК"
            attachment_path = 'test_data_template.xlsx'
            if os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
            else:
                print(f"Файл {attachment_path} не найден")
        elif letter_type == "overdue":
            mail.Subject = f"Уведомление от ТМК о просрочке по счетам для {Контрагент}"
            mail.To = f"{Контрагент}@example.com"
            mail.HTMLBody = f"""
            Уважаемый клиент {Контрагент},<br><br>
            Напоминаю, что по состоянию на {current_date} у Вас имеется неоплаченный счёт(а) по {Договор_контрагента}, {Спецификации}, отгрузка от {Дата_отгрузки}.
            Просим в кратчайшие сроки произвести оплату согласно условиям договора. В случае, если оплата уже произведена, пожалуйста, проигнорируйте это письмо или направьте подтверждение платежа в ответ.
            <br><br>
            С уважением,<br>
            Ваш менеджер по продажам,<br>
            {Менеджер}<br>
            Трубная металлургическая компания
            """
            file_prefix = "Уведомление_просрочка_по_счетам_ТМК"
            attachment_path = 'test_data_template.xlsx'
            if os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
            else:
                print(f"Файл {attachment_path} не найден")
        # Создание папки для сохранения, если она не существует
        if not os.path.exists(save_path):
            os.makedirs(save_path)

        # Формирование имени файла с использованием имени контрагента и текущей даты
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        clean_контрагент = clean_filename(Контрагент)
        file_name = f"{file_prefix}_{clean_контрагент}_{timestamp}.msg"
        full_path = os.path.join(save_path, file_name)

        # Сохранение письма как .msg
        mail.SaveAs(full_path)
        print(f"Письмо сохранено как {full_path}")
        return full_path
    except Exception as e:
        print(f"Произошла ошибка при создании файла: {str(e)}")
        return None
    finally:
        # Освобождение COM-ресурсов
        pythoncom.CoUninitialize()


def create_inline_keyboard():
    keyboard = InlineKeyboardMarkup()
    keyboard.add(InlineKeyboardButton("Готовность к отгрузке", callback_data="create_msg"))
    keyboard.add(InlineKeyboardButton("Неоплаченные счета", callback_data="create_unpaid"))
    keyboard.add(InlineKeyboardButton("Падение объема закупок", callback_data="create_volume_down"))
    keyboard.add(InlineKeyboardButton("Рост объема закупок", callback_data="create_volume_up"))
    keyboard.add(InlineKeyboardButton("Есть просрочка по счетам", callback_data="create_overdue"))
    return keyboard
