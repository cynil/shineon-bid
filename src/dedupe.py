import pandas as pd
from pandas import DataFrame
from eq import chn_eq, str_eq

#TOO_SHORT = 10
#SEARCH_RANGE=20

def dedupe(xlsx_name):
    df = DataFrame(pd.read_excel(xlsx_name))
    df['duped'] = False

    data = df.values

    for i in range(len(data) - 1):
        if i > 0:
            start = max(0, i - 20)
            print(i, start)
            for j, test_item in enumerate(df[start:i]):
                print(item)
                if len(item['信息标题']) < 10:
                    if str_eq(item['招标编号'], test_item['招标编号']):
                        item['duped'] = True
                elif len(item['信息标题']) >= 10:
                    if chn_eq(item['信息标题'], test_item['信息标题']):
                        item['duped'] = True
    
    df.to_excel('processed_' + xlsx_name)

dedupe('o_2000.xlsx')