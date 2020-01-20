import jieba
import json

#招标笼统词
junk_word = json.load(open('./old/junk.json', encoding='utf-8'))

#行业笼统词
broad_word = json.load(open('./old/broad.json', encoding='utf-8'))

def too_broad(s):
    s = list(jieba.cut(str(s)))
    count = 0
    
    for i in range(len(s)):
        if (s[i] in broad_word) or (s[i] in junk_word):
            count = count + 1
    return count / len(s) > 0.95

def chn_eq(c1, c2):
    c1 = str(c1)
    c2 = str(c2)

    if c1 == c2:
        return True
    if too_broad(c1) or too_broad(c2):
        return False

    if(len(c1) > len(c2)):
        tmp = c1
        c1 = c2
        c2 = tmp
        
    shortli = list(jieba.cut(c1))
    longli = list(jieba.cut(c2))
    count = 0

    shortlen = len(shortli)
    equal = []
    remain = []
    for index, word in enumerate(shortli):
        if word in junk_word:
            shortlen = shortlen - 1
        elif word in longli:
            count = count + 1
            equal.append(word)
        else:
            remain.append(word)
    return count == shortlen

def str_eq(s1, s2):
    s1 = str(s1)
    s2 = str(s2)
    return s1 != 'nan' and s2 != 'nan' and ((s1 in s2) or (s2 in s1))