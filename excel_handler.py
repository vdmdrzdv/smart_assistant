import pandas as pd
import warnings


def extract_data(file_path, sheet_name=0):
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(file_path, sheet_name=sheet_name)

        required_columns = [
            "Дата отгрузки (отправки)",
            "Контрагент",
            "Дней просрочки",
            "Менеджер",
            "Номер договора",
            "Оплачено"
        ]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            raise ValueError(f"Отсутствуют обязательные колонки: {', '.join(missing_columns)}")

        result = df[required_columns].copy()
        result["Дата отгрузки (отправки)"] = pd.to_datetime(result["Дата отгрузки (отправки)"], errors='coerce')
        result["Дней просрочки"] = pd.to_numeric(result["Дней просрочки"], errors='coerce')
        result["Оплачено"] = pd.to_numeric(result["Оплачено"], errors='coerce')
        return result

    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")
        return None
