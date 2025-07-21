import win32com.client
import os


def create_outlook_draft(Контрагент="Клиент", Менеджер="Имя Менеджера", Договор_контрагента ="",save_path="C:\\Helper\\emails"):
    try:
        # Подключение к Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 означает создание нового письма

        # Установка параметров письма
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

         # Создание папки для сохранения, если она не существует
        if not os.path.exists(save_path):
            os.makedirs(save_path)

        # Формирование имени файла с использованием имени контрагента и текущей даты
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"Уведомление_отгрузка_ТМК_{Контрагент}_{timestamp}.msg"
        full_path = os.path.join(save_path, file_name)

        # Сохранение письма как .msg
        mail.SaveAs(full_path)

        print(f"Письмо сохранено как {full_path}")

    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")

if __name__ == "__main__":
    # Пример вызова функции с тестовыми данными
    create_outlook_draft(
        Контрагент="ООО Ромашка",
        Менеджер="Иван Иванов",
        Договор_контрагента="Договор продажи № 290 от 14.04.2025"
    )
