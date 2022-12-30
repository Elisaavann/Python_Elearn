import pandas as pd

file_name = 'vacancies_dif_currencies.csv'
df = pd.read_csv(file_name)
curr = df['salary_currency'].unique()
req_curr = []
for c in curr:
    if list(df['salary_currency']).count(c) > 5000:
        req_curr.append(c)
del req_curr[0]
df.sort_values(by='published_at', inplace=True)
first_date = list(df['published_at'])[0]
second_date = list(df['published_at'])[-1]
print(req_curr, first_date, second_date)
data = pd.DataFrame(columns=['date', 'BYR', 'USD', 'EUR', 'KZT', 'UAH'])
for year in range(2003, 2023):
    for month in range(1, 13):
        date = f'01/0{month}/{year}' if month < 10 else f'01/{month}/{year}'
        new_row = {'date': date}
        res = f'http://www.cbr.ru/scripts/XML_daily.asp?date_req={date}'
        values_cur = pd.read_xml(res, encoding='cp1251')
        for c in req_curr:
            if len(values_cur[values_cur['CharCode'] == c]['Value'].values) != 0:
                value = '.'.join(str(values_cur[values_cur['CharCode'] == c]['Value'].values[0]).split(','))
                new_row[c] = float(value) / int(values_cur[values_cur['CharCode'] == c]['Nominal'].values[0])
        data = pd.concat([data, pd.DataFrame.from_records([new_row])], axis=0, ignore_index=True)
        if date == '01/12/2022':
            break
data.to_csv('dataencies.csv', index=False)
print(data)