import pandas as pd
from sklearn.linear_model import LinearRegression


def analyze_trend(company_df):
    if len(company_df) < 2:
        return 'Недостаточно данных'

    # Преобразуем даты в числовой формат для регрессии
    company_df = company_df.sort_values('Месяц')
    X = company_df['Месяц'].map(lambda d: d.toordinal()).values.reshape(-1, 1)
    y = company_df['Сумма'].values

    model = LinearRegression()
    model.fit(X, y)

    slope = model.coef_[0]

    if slope > 0:
        return 1
    elif slope < 0:
        return -1
    else:
        return 0


def get_trend(df):
    df_melted = df.melt(id_vars='Контрагент', var_name='Месяц', value_name='Сумма')
    df_melted['Месяц'] = pd.to_datetime(df_melted['Месяц'], format='%d.%m.%Y', errors='coerce')
    df_melted = df_melted.dropna(subset=['Месяц'])
    df_melted['Сумма'] = pd.to_numeric(df_melted["Сумма"], errors='coerce')
    df_melted = df_melted[df_melted['Сумма'] > 0]

    # Анализ для каждой компании
    trends = []
    for company in df['Контрагент']:
        company_data = df_melted[df_melted['Контрагент'] == company]
        trend = analyze_trend(company_data)
        trends.append(trend)
    new_df = df.copy()
    new_df.loc[:, 'Тренд'] = trends

    return new_df
