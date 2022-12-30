import pandas as pd 

file_name = 'vacancies_dif_currencies.csv'
df = pd.read_csv(file_name)
curr = df['salary_currency'].unique()
cur_dict = {}
for c in curr:
    value = list(df['salary_currency']).count(c)
    if value > 5000:
        cur_dict[c] = value
print(cur_dict)