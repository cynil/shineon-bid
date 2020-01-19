import pandas as pd
from pandas import DataFrame
import json
import jieba
import codecs

junk = json.load(open('junk.json', encoding='utf-8'))

#相似度
def diff(s1, s2):
    shortli = list(jieba.cut(s1))
    longli = list(jieba.cut(s2))
    count = 0

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
    return (shortlen - count, count / shortlen)




#重复断言算法
#1. 计算文本相似度。
#2. 标题词数过短的是不正常标题。看招标编号，招标编号相同，接受重复断言。
#3. 文本长度正常：【相似度比值在0.76以上0.9以下且相似度差值在两个以内】 或 【相似度比值大于0.9】的，足够相似，这样的通常是同一系列招标的不同包号，拒绝重复断言。
#4. 相似度比值等于1.0的，接受重复断言。
one_series = []
dupe = []
def assert_dupe():
    df = DataFrame(pd.read_excel('cut.xlsx'))
    title = df['信息标题']

    for index, item in enumerate(title):
        start = max(0, index - 10)
        if index > 0:
            for i,t in enumerate(title[start:index]):
                if (0 < diff(item, t)[0] < 3 and 0.76 < diff(item, t)[1] < 0.9) or (0.9 < diff(item, t)[1] < 1.0):
                    one_series.append({
                        "id": index,
                        "重复": "拒绝",
                        "diff": diff(item, t),
                        "item1": item,
                        "item2": t
                    })
                
                if diff(item, t)[1] == 1.0:
                    dupe.append({
                        "id": index,
                        "重复":"接受",
                        "diff": diff(item, t),
                        "item1": item,
                        "item2": t
                    })
assert_dupe()

def save_as_json(fname, data):
    f = open(fname, 'w', encoding='utf-8')
    f.write(json.dumps(data,ensure_ascii=False))
    f.close()

save_as_json('one_series.json', one_series)
save_as_json('dupe.json', dupe)