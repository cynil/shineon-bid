import pandas as pd
import jieba
from utils import save_json
from pandas import DataFrame
from collections import Counter

df = DataFrame(pd.read_excel('o.xlsx', usecols='B'))
title = df['信息标题']
word_pool = []

for index, item in enumerate(title):
    try:
        word_split = list(jieba.cut(item))
        for index, word in enumerate(word_split):
            word_pool.append(word)
    except:
        print(index, item)

result = [item[0] for item in Counter(word_pool).most_common(400)]
save_json('junk2.json', result)