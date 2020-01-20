import json

def save_json(fname, data):
    f = open(fname, 'w', encoding='utf-8')
    f.write(json.dumps(data, ensure_ascii=False))
    f.close()