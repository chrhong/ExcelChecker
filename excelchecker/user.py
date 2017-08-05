#! -*- coding: utf-8 -*-
from __init__ import *
from easyexcel import *

def index_count_range(price_value):
    #数量区间：
    #   单价 < 100，        [0, 2000]
    #   100 < 单价 < 500，  [0, 1000]
    #   单价 > 500，        [0, 500]
    count_value = int(price_value)
    if count_value <= 100:
        count_max = 2000
    elif count_value <= 500:
        count_max = 1000
    else:
        count_max = 500
    return count_max

def index_sell_count_check(sht, row, dbi, params_list):
    col = params_list[0]
    col1 = params_list[1]

    count_value = int(sht.getCell(row, col))
    price_value = float(sht.getCell(row, col1))
    count_max = index_count_range(price_value)

    if count_value > count_max:
        return ERROR
    else:
        return NORMAL

def index_sell_price_check(sht, row, dbi, params_list):
    db = sht.xls.database

    price_value = float(sht.getCell(row, params_list[0]))
    price_min = float(db[dbi][1].strip('[]').split(',')[0])
    price_max = float(db[dbi][1].strip('[]').split(',')[1])
    if price_value < price_min or price_value > price_max:
        return ERROR
    else:
        return NORMAL

def sale_table_check(xls, key):
    sht = xls.Sheet(xls, 'Sheet1')
    db_key = key

    sale_rule = {
        "TitleLine" : "E:开票索引",
        "NoneKey" : "Mark", #ignore/mark/correct/delete
        "Add" : [],
        "Keep" : [],
        "CheckItem" : ['C','B:C'],
        "CheckMethod" : [index_sell_price_check, index_sell_count_check],
    }

    sht.dupCombColumn(sale_rule, db_key)
    return sht

def buy_table_check(xls, key):
    sht = xls.Sheet(xls, 'Sheet1')
    db_key = key

    buy_rule = {
        "TitleLine" : "F:零件号码",
        "NoneKey" : "Mark", #ignore/mark/correct/delete
        "Add" : [],
        "Keep" : [],
        "CheckItem" : ['K','I:K'],
        "CheckMethod" : [index_sell_price_check, index_sell_count_check],
    }

    sht.dupCombColumn(buy_rule, db_key)
    return sht

def stock_table_check(xls, key):
    sht = xls.Sheet(xls, 'Sheet2')
    db_key = key

    stock_rule = {
        "TitleLine" : "A:关键字",
        "NoneKey" : "Fix:C", #Ignore/Mark/Fix/Delete
        "Add" : ['D','E','F','G','H'],
        "Keep" : ['B','C'], #Keep not None
        "CheckItem" : [],
        "CheckMethod" : []
    }

    sht.dupCombColumn(stock_rule, db_key)
    return sht


excel_handler = {
    '库存表' : stock_table_check,
    '销售表' : sale_table_check,
    '进项表' : buy_table_check
}

def database_key_get(xls, type_list, check_type, source_file):
    db = []
    db_key = []

    if not type_list[check_type] == '库存表':
        #product.db read, stock do not need db
        db = xls.dbInit(source_file[:source_file.rfind('/')+1] + 'product.db')
        for li in db : db_key.append(li[0])

    return db_key

def userChecker(source_file, type_list, check_type):
    try:
        #should be called in each thread, otherwise, execl open failed.
        pythoncom.CoInitialize()
        # eprint(source_file)
        eprint('[INFO] 开始 %s 检查...' % type_list[check_type])
        start = time.clock()
        xls = EasyExcel(source_file)
        db_key = database_key_get(xls, type_list, check_type, source_file)
        sht = excel_handler[type_list[check_type]](xls, db_key)
        xls.save()
        end = time.clock()
        eprint("[INFO] %s 检查完成, 耗时: %f 秒" % (type_list[check_type], end - start))
        eprint("[INFO] 错误: %d 个, 无法匹配: %d 个" % (sht.getStatistic(ERROR), sht.getStatistic(UNKNOW)))
        xls.close()
    except:
        RaiseException()

def gui_mainloop():
    root = EasyGUI("ExcelChecker v0.2", '500x205', UNRESIZEBLE)
    root.Entry("Source file", 48, (10,20))
    root.LogWindow(("white","black"), (48,5), (10,100), UNEDITABLE)
    r1 = root.Radiobutton("进项表", 1, (10,60))
    r2 = root.Radiobutton("销售表", 2, (110,60))
    r3 = root.Radiobutton("库存表", 3, (210,60))
    root.Button('打开', (8,1), (410,17), root.Browser)
    root.Button('检查', (8,4), (410,100), lambda:root.Thread(userChecker))
    root.mainloop()

def oneClick_mainloop():
    Registry('ExcelChecker.exe')

    try:
        if len(sys.argv) < 2:
            print("[Error] argv error! Use demo file in debug mode!")
            #debug file
            source_file = u"D:/userdata/chrhong/Desktop/Python/X/去库存索引表.1.xlsx"
        else:
            source_file = ' '.join(sys.argv[1:]).strip("\'\"")
            source_file = source_file.decode('gbk').replace('\\', '/').replace('\"', '')

        print(source_file)

        if u'进项' in source_file:
            file_type = 'buy'
        elif u'索引' in source_file:
            file_type = 'index'
        elif u'库存' in source_file:
            file_type = 'stock'
        else:
            print("[ERROR] %s not supported !" % source_file)
            print(u"[INFO] Valid file: '进项', '索引', '库存' !")
            RaiseException(HANG)

        cost_db = []
        cost_key = []
        if not file_type == 'stock':
            #product.db read, stock do not need db
            cost_db = xls.dbInit(source_file[:source_file.rfind('/')+1] + 'product.db')
            for li in cost_db:
                cost_key.append(li[0])

        print('[INFO] start %s_file checking...' % file_type)
        start = time.clock()
        xls = easyExcel(source_file)
        excel_handler[file_type](xls,cost_db,cost_key)
        xls.save()
        xls.close()
        end = time.clock()
        print("[INFO] %s_table_check: %f s" % (file_type, end - start))
    except:
        RaiseException(HANG)