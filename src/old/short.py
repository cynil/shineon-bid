import pandas as pd
from pandas import DataFrame
from utils import save_json

df = DataFrame(pd.read_excel('o.xlsx', usecols='B'))
title = df['信息标题']
pool = []

for index, item in enumerate(title):
    try:
        if len(item) < 10:
            pool.append(item)
    except:
        print(index, item)

save_json('short.json', pool)