import pandas as pd
from datetime import datetime, date

df = pd.read_excel(r'new.xlsx')
raw_date = list(df['samir'],dtype=str)
str_remark = map(str, raw_remark)
pizza=[]
for n in raw_date:

    pizza.append(n.strftime("%d-%m-%Y"))

print (pizza)