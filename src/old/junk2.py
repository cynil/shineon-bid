import pandas as pd
from pandas import DataFrame
import json
import jieba

junk = json.load(open('junk.json', encoding='utf-8'))
proxim = []
#相似度
def diff(s1, s2):
    shortli = list(jieba.cut(s1))
    longli = list(jieba.cut(s2))
    count = 0
    remains = []

    if(len(shortli) > len(longli)):
        tmp = shortli
        shortli = longli
        longli = tmp

    shortlen = len(shortli)

    for index, word in enumerate(shortli):
        if word in junk:
            shortlen = shortlen - 1
        elif word in longli:
            count = count + 1
        else:
            remains.append(word)
    return (count, shortlen, remains)


df = DataFrame(pd.read_excel('o.xlsx',  ))
title = df['信息标题']
SEARCH_RANGE = 20
TOO_SHORT = 4

for i, item in enumerate(title):
    try:
        start = max(0, i - SEARCH_RANGE)
        if i > 0:
            for j, t in enumerate(title[start:i]):
                _diff = diff(item, t)
                if (0 < _diff[1] - _diff[0] < 3) and (_diff[1] >= TOO_SHORT):
                    proxim.append({
                        "short_len": _diff[1],
                        "same_count": _diff[0],
                        "remains": _diff[2],
                        "diff": _diff[0] / _diff[1],
                        "item1": item,
                        "item2": t
                    })
    except:
        print(i, item)

def save_json(fname, data):
    f = open(fname, 'w', encoding='utf-8')
    f.write(json.dumps(data,ensure_ascii=False))
    f.close()

save_json('proxim.json', proxim)