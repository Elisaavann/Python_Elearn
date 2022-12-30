import pandas as pd
import sqlite3
from sqlalchemy import create_engine

connection = sqlite3.connect('python_vac.db')
engine = create_engine('sqlite:///D:\\LIZOK\\_Практика PY\\Pyton 2 курс\\Тема3_5 Базы Данных\\python_vac.db')
df = pd.read_csv('data_currencies.csv')
df.to_sql('currencies', con=engine, index=False)