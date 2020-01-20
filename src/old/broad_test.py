import jieba
import json
from utils import save_json
import pandas as pd
from pandas import DataFrame


broad_word = json.load(open('broad.json', encoding='utf-8'))

junk_word = json.load(open('junk.json', encoding='utf-8'))

def too_broad(s):
    s = list(jieba.cut(str(s)))
    count = 0
    
    for i in range(len(s)):
        if ((s[i] in broad_word) or (s[i] in junk_word)) and len(s) < 8:
            count = count + 1
    return count / len(s) > 0.95

df = DataFrame(pd.read_excel('o.xlsx', usecols='B'))
title = df['信息标题']

broad_2 = []
for i in range(len(title)):
    if too_broad(title[i]):
        broad_2.append(title[i])

save_json('broad_titles.json', broad_2)