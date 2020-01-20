import jieba
import json

junk_word = json.load(open('./old/junk.json', encoding='utf-8'))

def chn_eq(c1, c2):
    if c1 == c2:
        return True
    shortli = list(jieba.cut(c1))
    longli = list(jieba.cut(c2))
    count = 0

    if(len(shortli) > len(longli)):
        tmp = shortli
        shortli = longli
        longli = tmp

    shortlen = len(shortli)

    for index, word in enumerate(shortli):
        if word in junk_word:
            shortlen = shortlen - 1
        elif word in longli:
            count = count + 1 
    return count == shortlen

def str_eq(s1, s2):
    return (s1 in s2) or (s2 in s1)