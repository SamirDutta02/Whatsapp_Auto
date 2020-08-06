import pandas as pd
df1 = pd.read_excel('date.xlsx',na_filter= False,dtype=str)
raw_performa = list(df1.loc[:, 'Date'])

print(raw_performa)
