import pandas as pd
from pandas import DataFrame
from eq import chn_eq, str_eq

#TOO_SHORT = 10
#SEARCH_RANGE=20

def dedupe(xlsx_name):
    data = DataFrame(pd.read_excel(xlsx_name)).values
    #data2 = 

    for i in range(1, len(data)):
        print(data[i])
        start = max(0, i - 20)
        for j in range(start, i):
                if chn_eq(data[i][1], data[j][1]):
                    data[i][7] = j
                if str_eq(data[i][3], data[j][3]):
                    data[i][8] = j
    DataFrame(data).to_excel('processed_' + xlsx_name)

dedupe('o_2000.xlsx')