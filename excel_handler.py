import pandas as pd
import warnings

# Сопоставление строк с месяцами и их числовыми значениями
month_map = {
    "Январь": "01",
    "Февраль": "02",
    "Март": "03",
    "Апрель": "04",
    "Май": "05",
    "Июнь": "06",
    "Июль": "07",
    "Август": "08",
    "Сентябрь": "09",
    "Октябрь": "10",
    "Ноябрь": "11",
    "Декабрь": "12"
}


def rename_month_columns(columns):
    new_columns = []
    for col in columns:
        for name, num in month_map.items():
            if name in col:
                # Преобразуем, например, "Январь 2025" -> "01.01.2025"
                year = col.split()[-1]
                new_col = f"01.{num}.{year}"
                new_columns.append(new_col)
                break
        else:
            new_columns.append(col)  # не месячная колонка
    return new_columns


def extract_data(file_path, sheet_name="TDSheet"):
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(file_path, sheet_name=sheet_name)

        required_columns = [
            "Спецификации",
            "Дата отгрузки (отправки)",
            "Контрагент",
            "Тип контрагента",
            "Дней просрочки",
            "Менеджер",
            "Номер договора",
            "Оплачено",
            "Январь 2025",
            "Февраль 2025",
            "Март 2025",
            "Апрель 2025",
            "Май 2025",
            "Июнь 2025",
            "ДЗ"
        ]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            raise ValueError(f"Отсутствуют обязательные колонки: {', '.join(missing_columns)}")

        result = df[required_columns].copy()
        result["Дата отгрузки (отправки)"] = pd.to_datetime(result["Дата отгрузки (отправки)"], errors='coerce')
        result["Дней просрочки"] = pd.to_numeric(result["Дней просрочки"], errors='coerce')
        result["Оплачено"] = pd.to_numeric(result["Оплачено"], errors='coerce')
        result["ДЗ"] = pd.to_numeric(result["ДЗ"], errors='coerce')

        return result

    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")
        return None


def create_pivot_table(df):
    # Предположим, df — это исходный датафрейм
    month_columns = [
        "Январь 2025", "Февраль 2025", "Март 2025", "Апрель 2025",
        "Май 2025", "Июнь 2025"
    ]

    # Сводная таблица: сумма купленных тонн по каждому контрагенту
    pivot_table = df.groupby(["Менеджер", "Контрагент"])[month_columns].sum().reset_index()
    pivot_table.columns = rename_month_columns(pivot_table.columns)

    return pivot_table.sort_values(by="Контрагент", ascending=True).reset_index(drop=True)
