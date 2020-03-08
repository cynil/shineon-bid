#coding:utf-8
#@COPYRIGHT CYNI 2020
import re
import os
import time
import pandas as pd
from pandas import read_excel, DataFrame
#utils -> raw2num -> eq -> command -> main

def myint(s):
    try:
        return int(s)
    except:
        return 0
def assign(s, pos, char):
    if int(pos) > len(s) - 1: return s
    s = s[:pos] + str(char) + s[pos+1:]
    return s
def only_contains(s, legal):
    for w in str(s):
        if not w in legal:
            return False
    return True
def purify(s, legal):
    s = str(s)
    r = ''
    for w in s:
        if re.match(legal, w) is not None:
            r += w
    return r

def num2AZ(num):#0 => A
    tup = divmod(num, 26)
    tup = (tup[0] - 1, 26) if tup[1] == 0 else tup
    return (chr(tup[0]+64) + chr(tup[1]+64)).replace('@','')
def AZ2num(AZ):#A => 0
    if len(AZ) < 2:
        return ord(AZ[0]) - 65
    elif len(AZ) == 2:
        return (ord(AZ[0]) - 64) * 26 + ord(AZ[1]) - 65
def AZs2num(AZs):#A:D,G => A,B,C,D,G
    r = set()
    AZs = AZs.upper()
    if AZs == '': return False
    try:
        for i in range(len(AZs)):
            if not (AZs[i] in ':,' or 64 < ord(AZs[i]) < 91):
                return False
            if AZs[i] in ':,':
                try:
                    if AZs[i+1] in ':,': return False
                except:
                    pass
        AZs = AZs.split(',')
        for c in AZs:
            try:
                if ':' in c:
                    l = min(AZ2num(c.split(':')[0]), AZ2num(c.split(':')[1]))
                    h = max(AZ2num(c.split(':')[0]), AZ2num(c.split(':')[1]))
                    for j in range(l, h + 1):
                        r.add(j)
                else:
                    r.add(AZ2num(c))
            except:
                pass
        return list(r)
    except:
        return False

def AZs2col(dname, raw):
    cols = AZs2num(raw)
    if not cols: return []

    for i in range(len(cols)):
        try:
            cols[i] = dname.columns[cols[i]]
        except:
            return []
    return cols

def cmd2num(cmd):#1,3:6 => 1,3,4,5,6
    num = set()
    for token in cmd.split(','):
        if ':' in token:
            t1 = int(token.split(':')[0])
            t2 = int(token.split(':')[1])
            l = min(t1, t2)
            h = max(t1, t2)
            for i in range(l, h + 1):
                num.add(i)
        else:
            num.add(int(token))
    return list(num)

def printhead(li):
    l = []
    for i in range(len(li)):
        l.append('[' + num2AZ(i+1) + ']' + str(li[i]))
    print(l)
def printlist(li):
    for i in range(len(li)):
        if '[' in str(li[i]):
            print(li[i])
        else:
            print('[' + str(i) + ']', li[i])
def promptmenu(menu, msg='请选择要进行的操作（输入序号）：'):
    time.sleep(1)
    print('-------------')
    printlist(menu)
    usercommand = input(msg)
    return usercommand

#将不同格式金额转换为数字格式
def chn2num(raw):#:string
    _0to9 = '零壹贰叁肆伍陆柒捌玖'
    scale = '元拾佰仟'
    scale_ = '元万亿兆'
    if raw == '': return 0
    for char in raw:
        i = raw.index(char)
        if i < len(raw) - 1:
            if char in _0to9[1:] and raw[i+1] in _0to9[1:]:
                return '[0] Illegal Input: ' + char + raw[i+1]
    jiao = 0
    fen = 0

    try:
        if '分' in raw:
            fen = _0to9.index(raw[raw.index('分') - 1])
            raw = raw[:-2]
        if '角' in raw :
            jiao = _0to9.index(raw[raw.index('角') - 1])
            raw = raw[:-2]
        
        if raw in '零': return jiao * 0.1 + fen * 0.01#空或零
        if not '元' in raw: raw += '元'

        flag = 0
        integer = '0000000000000000000000000'
        for i in range(len(raw)-1,-1,-1):
            if raw[i] in _0to9[1:]:
                value = _0to9.index(raw[i])
                if raw[i+1] in scale:
                    pos = scale.index(raw[i+1]) + flag * 4
                elif raw[i+1] in scale_:
                    pos = flag * 4
                integer = assign(integer, pos, value)
            elif raw[i] in scale:#元十百千
                pos = scale.index(raw[i]) + flag * 4
                try:
                    if raw[i-1] in scale and raw[i] != '元':
                        integer = assign(integer, pos, 1)
                except:
                        integer = assign(integer, pos, 1)
            elif raw[i] in scale_:#万亿兆
                flag = scale_.index(raw[i])
        return int(integer[::-1]) + jiao * 0.1 + fen * 0.01
    except:
        return 'Illegal Input: ' + raw

def purify_me(s):
    s = str(s)
    legal='0123456789.零壹贰叁肆伍陆柒捌玖拾佰仟万亿兆五百千角分元圆'
    nan = '省市县区乡镇街道路'
    r = ''
    s = str(s)
    for char in s:
        if char in nan: return ''
        if char in legal:
            r += char
    return r.replace('五','伍').replace('百','佰').replace('千','仟').replace('圆','元')

def myfloat(s):
    s = str(s)
    try:
        return float(s)
    except:
        if re.match('^[\d\.]+$', s) == None: return 0
        s = s.split('.')
        if len(s) > 1:
            return myfloat(s[0] + '.' + s[1])    
        return int(s[0])

def raw2num(raw):
    raw = purify_me(raw)
    if raw == '': return 0
    if re.match('^[\d\.]+元?$', raw):
        result = raw.replace('元','')

    elif re.match('^([\d\.]+)万元?.*$', raw):
        r = myfloat(raw.split('万')[0])
        result = r if r > 1e6 else r * 1e4

    elif re.search('[零壹贰叁肆伍陆柒捌玖]+', raw): #汉字[\u4e00-\u9fa5]
        arab = '0123456789.'
        num = ''
        for char in raw:
            if char in arab:
                num += char
        if len(num) > 0:
            result = num
        else:
            result = chn2num(raw)#On Error Return 'Illegal Input: ' + raw
    else:
        result = raw
    result = myfloat(result)
    if result < 5e2:
        result *= 1e4
    elif result > 1e9:
        result = 0
    return result

junk_word = set(["公告", "采购", "项目", "-", "中标", "设备", "）", "开标", "（", "成交", "结果", "中选", "公示", "招标", "、", "题", "及", "系统", "2019", "的", "关于", "电子化", "更正", "竞价", "合同", "有限公司", "评标", "供应商", "变更", "建设", "项目", "年", "公开", "招标", "流标", "服务", "竞争性", "平台", "(", ")", "]", "[", "：", "\"", "改造", "工程", "询价", "和", "磋商", "”", "“", "技术", "谈判", "编号", "来源", "等", "包", "咨询", "服务", "与", "购置", "》", "《", "【", "】", "类", "区", "公司", "编号", "月", "交易", "候选人", "日", "_", "代理", "网上", "2018", "本级", "全", "标段", "验收报告", "验收", "预", "分公司", "年度", "有限责任", ":", "废标", "设计", "号", "维护", "之", "管理", "[竞]"])
broad_word = set(["集成", "设备", "录播", "教室", "系统", "媒体", "记录", "智慧", "施工", "直播", "融", "服务", "空调", "中心", "室", "校园", "电视台", "等", "广播", "监控", "改造", "通", "一体机", "班班", "、", "竞价", "电脑", "打印机", "高清", "全", "装修", "监理", "政府", "及", "宣传", "播出", "询价", "车辆", "计算机", "设施", "电视", "维修", "摄像机", "办公设备", "视频", "平台", "教育", "网络", "录像机", "新", "制作", "学院", "工程", "台式", "购置", "购买", "办公", "网上", "多媒体", "课堂", "演播室", "创客", "电台", "调音台", "慕课", "硬盘", "其他", "专用设备", "移动", "竞争性", "演播厅", "新闻", "器材", "实验室", "配套", "【", "工程施工", "标段", "会议室", "中学", "教学设备", "采编", "一批", "碳粉", "实训室", "直", "教学", "的", "协议", "电视机", "法庭", "(", "次", ")", "便携式", "精品", "拍摄", "联想", "软件", "广维", "小学", "播", "节目", "版权", "合同", "LED", "宣传费", "升级", "变更", "备件", "中央", "运营", "合作", "交互", "服务器", "】", "磋商", "安装", "关于", "保险", "云", "室内", "耗材", "设计", "硬件", "单一", "来源", "微课", "物业管理", "谈判", "公司", "媒资", "紧急", "多功能", "显示屏", "家具", "信息化", "会议", "服务", "资源", "无线", "应急", "桌椅", "同步", "材料", "：", "自动", "音乐", "融媒", "工作", "“", "”", "广告", "技术", "间", "定点", "预", "维护", "保险合同", "服务费", "美", "实验", "课程", "开发", "存储设备", "供货", "专题片", "空间", "HP", "配件", "维", "保费", "激光打印机", "加油", "租赁", "一套", "扩容", "装备", " ", "（", "）", "UPS", "电源", "投影机", "复印机", "摄录", "编", "教学系统", "新建", "便携", "机", "学校", "资源库", "视频会议", "在线", "监管", "数字", "模拟", "通讯", "计划", "项", "文化", "数字化", "线上", "验证", "测试计划", "非编", "督导", "远程", "初步设计", "验收", "思政微", "课", "智能化", "新闻宣传", "发电机组", "汽油", "机动车辆", "台式机", "采访", "办公桌", "供应商", "手机", "台", "网站", "都市", "频道", "劳务", "派遣", "区域", "发布", "科技", "信息中心", "点播", "虚拟", "更新", "总", "承包", "用", "中央空调", "空调机", "购置税", "RCS", "标准", "许可", "一笔", "订单", "广播系统", "需", "供电", ",", "社区", "城域网", "电子化"])
pool_a = broad_word | junk_word
def wash(s, pool):
    t = str(s)
    for word in pool:
        if word in t:
            t = t.replace(word, '')
    return t

def dejunk(s):
    s = str(s)
    if len(s) < 10: return ''
    if len(wash(s, pool_a)) < 2: return ''
    return wash(s, junk_word)

def eq(l1,l2):
    cap1 = set(dejunk(l1))
    cap2 = set(dejunk(l2))
    if cap1 and cap2:
        if cap1 <= cap2 or cap2 <= cap1: return True
    
    return False

#command.py
def fetch_data(incoming, all):
    file_to_read = []
    failed = []
    data = None
    if incoming == '':
        file_to_read = all
    elif only_contains(incoming, ':0123456789,'):
        for no in cmd2num(incoming):
            try:
                file_to_read.append(all[no])
            except:
                if (str(no) + '.xls') in all:
                    file_to_read.append(str(no) + '.xls')
                elif (str(no) + '.xlsx') in all:
                    file_to_read.append(str(no) + '.xlsx')
                else:
                    failed.append(no)
    else:
        for fn in incoming.split(','):
            if (str(fn) + '.xls') in all:
                file_to_read.append(str(fn) + '.xls')
            elif (str(fn) + '.xlsx') in all:
                file_to_read.append(str(fn) + '.xlsx')
            else:
                failed.append(fn)

    for file in file_to_read:
        try:
            print('正在读取：' + file)
            df = DataFrame(read_excel(file, header=1, encoding='utf-8'))
            if not '行业关键词' in df.columns.values:#Unnamed: 0
                df = DataFrame(read_excel(file, header=0, encoding='utf-8'))
            if data is None:
                data = df
            else:
                data = pd.concat([data, df], sort=False,ignore_index=True)
        except:
            failed.append('[' + str(all.index(file)) + '] ' + file)

    if len(failed) > 0:
        print('以下文件读取失败：')
        printlist(failed)

    return data

def drop_cols(dname, cols):
    dname.drop(cols,axis=1,inplace=True)

def view_n(dname, N):
    printhead(dname.columns.values)
    for i in range(N):
        irow = []
        for cell in list(dname.iloc[i]):
            cell = cell if len(str(cell)) < 15 else str(cell)[:15] + '...'
            irow.append(str(cell))
        print('[' + str(i) + ']', irow)

def col2num(dname, cols):#[1,2,3]
    for col in cols:
        pos = dname.columns.get_loc(col)
        arab = dname[col].map(lambda x: raw2num(x))
        dname.insert(pos + 1, col + '_数值', arab)

def dedupe(dname):
    if not '中标金额' in dname.columns.values:
        dname['cashnum'] = 0
        for colname in dname.columns.values:
            if re.match('中标金额\d$', colname):
                dname['cashnum'] += dname[colname].map(lambda x: raw2num(x))
        if not dname['cashnum'].any():
            return None
    else:
        dname['cashnum'] = dname['中标金额'].map(lambda x: raw2num(x))
    if not 'retain' in dname.columns.values:
        dname['retain'] = 'retain'
    print('开始第一遍去重：')
    dname = dedupe_by_num(dname)
    print('开始第二遍去重：')
    dname = dedupe_by_sig(dname)
    print('开始第三遍去重：')
    dname = dedupe_by_cap0(dname)
    print('开始第四遍去重：')
    dname = dedupe_by_cap(dname)
    dname = dname.drop(['cashnum', 'retain'], axis=1)
    dname.sort_index(inplace=True)
    return dname

def dedupe_by_num(dname):
    len0 = dname.shape[0]
    dname['num'] = dname['招标编号'].map(lambda x: purify(x, '[a-zA-Z\d\u4e00-\u9fa5]'))
    dname.sort_values(by=['num'], na_position='last', inplace=True)

    for i in range(dname.shape[0]):
        i_row = dname.iloc[i]
        i_id = i_row.name
        i_num = i_row['num']
        i_cash = i_row['cashnum']
        
        if i > 0 and len(str(i_num)) > 6:
            j_row = dname.iloc[i-1]
            j_id = j_row.name
            j_num = j_row['num']
            j_cash = j_row['cashnum']

            if i_num == j_num:
                if i_cash != 0 and j_cash == 0:
                    dname.loc[j_id, 'retain'] = 'discard'
                elif i_cash == 0 and j_cash != 0:
                    dname.loc[i_id, 'cashnum'] = j_cash
                    dname.loc[i_id, 'retain'] = 'discard'
                else:
                    dname.loc[i_id, 'retain'] = 'discard'
        if i % 200 == 0: print('[按招标编号去重已完成: ' + str(round(i * 100 / len0, 2)) + '%]')

    dname = dname[dname['retain'].isin(['retain'])]
    dname = dname.drop(['num'], axis=1)
    print('共去除重复项:' + str(len0 - dname.shape[0]) + '项')
    return dname

def dedupe_by_sig(dname):
    len0 = dname.shape[0]
    dname['strcash'] = dname['cashnum'].map(lambda x: str(x))
    dname.sort_values(by=['strcash','行业关键词','所在市','信息标题'],ascending=False, na_position='last', inplace=True)

    for i in range(dname.shape[0]):
        i_row = dname.iloc[i]
        i_id = i_row.name
        i_cash = i_row['cashnum']
        i_kw = i_row['行业关键词']
        i_capt = i_row['信息标题']
        i_shi = i_row['所在市']
        
        if i > 0 and i_cash > 0:
            j_row = dname.iloc[i-1]
            j_id = j_row.name
            j_cash = j_row['cashnum']
            j_kw = j_row['行业关键词']
            j_capt = j_row['信息标题']
            j_shi = j_row['所在市']

            def is_eq():
                if i_cash > 0 and i_cash == j_cash:
                    if i_kw == j_kw and i_shi == j_shi and eq(i_capt, j_capt):
                        return True
                return False

            if is_eq():
                if i_cash != 0 and j_cash == 0:
                    dname.loc[j_id, 'retain'] = 'discard'
                elif i_cash == 0 and j_cash != 0:
                    dname.loc[i_id, 'cashnum'] = j_cash
                    dname.loc[i_id, 'retain'] = 'discard'
                else:
                    dname.loc[i_id, 'retain'] = 'discard'
        if i % 200 == 0: print('[按关键词去重已完成: ' + str(round(i * 100 / len0, 2)) + '%]')

    dname = dname[dname['retain'].isin(['retain'])]
    dname = dname.drop(['strcash'], axis=1)
    print('共去除重复项：' + str(len0 - dname.shape[0]) + '项')
    return dname

def dedupe_by_cap0(dname):
    len0 = dname.shape[0]
    dname['信息标题'] = dname['信息标题'].map(lambda x: str(x))
    dname.sort_values(by=['信息标题'],ascending=False, na_position='last', inplace=True)

    for i in range(dname.shape[0]):
        i_row = dname.iloc[i]
        i_id = i_row.name
        i_cash = i_row['cashnum']
        i_capt = i_row['信息标题']

        if i > 0:
            for j in range(max(i-10,0),i,1):
                j_row = dname.iloc[j]
                j_id = j_row.name
                j_cash = j_row['cashnum']
                j_capt = j_row['信息标题']

                if eq(i_capt, j_capt):
                    if i_cash > 0 and j_cash == 0:
                        if dname.loc[j_id,'retain'] == 'discard':
                                dname.loc[i_id,'retain'] = 'discard'
                        else:
                            dname.loc[j_id,'retain'] = 'discard'
                    else:
                        dname.loc[i_id,'retain'] = 'discard'
                    break
        if i % 200 == 0: print('[按信息标题去重已完成: ' + str(round(i * 100 / len0, 2)) + '%]')

    dname = dname[dname['retain'].isin(['retain'])]
    print('共去除重复项：' + str(len0 - dname.shape[0]) + '项')
    return dname

def dedupe_by_cap(dname):
    cnt = 0
    len0 = dname.shape[0]
    dname['信息标题'] = dname['信息标题'].map(lambda x: str(x))
    dname.sort_values(by=['信息标题'], na_position='last', inplace=True)
    for name, group in dname.groupby(['所在市', '行业关键词']):
        for i in range(group.shape[0]):
            if i > 0:
                i_row = group.iloc[i]
                i_id = i_row.name
                i_cash = i_row['cashnum']
                i_capt = i_row['信息标题']

                for j in range(i-1,-1,-1):
                    j_row = group.iloc[j]
                    j_id = j_row.name
                    j_cash = j_row['cashnum']
                    j_capt = j_row['信息标题']

                    if eq(i_capt, j_capt):
                        #前无金额后有金额：若前是discard，说明之前的循环已经有
                        #有金额的“行”把前改为discard了，也即已经找到了有金额的行了，
                        #这时就算当前行有金额也不必保留了。
                        if i_cash > 0 and j_cash == 0:
                            if dname.loc[j_id,'retain'] == 'discard':
                                    dname.loc[i_id,'retain'] = 'discard'
                            else:
                                dname.loc[j_id,'retain'] = 'discard'
                        else:
                            dname.loc[i_id,'retain'] = 'discard'
                        break
            cnt += 1
            if cnt % 200 == 0: print('[深度去重已完成: ' + str(round(cnt * 100 / len0, 2)) + '%]')
    dname = dname[dname['retain'].isin(['retain'])]
    print('共去除重复项:' + str(len0 - dname.shape[0]) + '项')
    return dname
    
def save(dname, fname='output.xlsx'):
    fname = fname.lower()
    if not '.xls' in fname:
        fname += '.xlsx'
    print('正在保存：' + fname)
    dname.to_excel(fname, engine='xlsxwriter')

#main.py
menu = [
    '导入数据',#fetch_data ✔
    '删除冗余列',#drop_cols✔
    '查看前N行',#view_n    ✔
    '金额转数值',#raw2num  ✔
    '去重',#dedupe        ✔
    '保存并关闭',#save      ✔
    '退出'#exit            ✔
]

def listcwd(ext):
    l = []
    for file in os.listdir(os.getcwd()):
        if ext in file and not '@@' in file:
            l.append(file)
    return l

DATA = None
print('--------------------------')
print('\n')
print('EXCEL招标数据处理工具 V1.1')
print('\n')
print('--------------------------')

usercommand = -1
while usercommand != '6':
    usercommand = promptmenu(menu)
    while (not usercommand in '0123456') and usercommand != '':
        usercommand = promptmenu(menu, '输入有误，请输入要进行的操作：')

    if usercommand == '0':#导入数据
        if DATA is not None:
            save_ = input('打开新数据前，是否要保存当前数据？(Y/N)\n:>')
            if save_ == 'Y':
                save(DATA, input('请输入文件名：\n'))

        printlist(listcwd('.xls'))
        info = input('请选择要处理的文件：\n:>')
        DATA_ = fetch_data(info, listcwd('.xls'))
        if DATA_ is None:
            print('导入数据为空！')
        else:
            DATA = DATA_
            view_n(DATA, 3)
            print('------------------')
            print('共导入数据', DATA.shape[0],'行')

    elif usercommand in '12345' and usercommand != '':
        if DATA is None:
            print('无数据，无法进行该项操作！')
        else:
            if usercommand == '1':#删除冗余列
                printhead(DATA.columns.values)
                info = input('请输入要进行删除冗余列操作的列序号：\n:>')
                while len(AZs2col(DATA, info)) == 0:
                    info = input('输入有误，请输入要进行删除冗余列操作的列序号：\n:>')
                drop_cols(DATA, AZs2col(DATA, info))
                printhead(DATA.columns.values)
            elif usercommand == '2':#查看前N行
                info = input('请输入要查看的行数：\n:>')
                while myint(info) == 0:
                    info = input('输入有误，请输入要查看的行数：\n:>')
                view_n(DATA, myint(info))
            elif usercommand == '3':#金额转数值
                printhead(DATA.columns.values)
                info = input('请输入要进行金额转数值操作的列：\n:>')
                while len(AZs2col(DATA, info)) == 0:
                    info = input('输入有误，请输入要进行金额转数值操作的列序号：\n:>')
                col2num(DATA, AZs2col(DATA, info))
                view_n(DATA, 5)
            elif usercommand == '4':# 去重
                len0 = DATA.shape[0]
                DATA_ = dedupe(DATA)
                if DATA_ is None:
                    print('缺少关键字段，无法进行去重操作！')
                else:
                    DATA = DATA_
                    print('去重完成，共去除重复项：', len0 - DATA.shape[0],'项')
            elif usercommand == '5':# 保存
                save(DATA, input('请输入文件名：\n'))
                DATA = None
                print('文件已成功保存！')

if DATA is not None:
    save_ = input('退出前是否要保存当前数据？(Y/N)\n:>')
    if save_ == 'Y':
        save(DATA, input('请输入文件名：\n'))
        print('文件已成功保存！')

ok = input('按任意键退出')