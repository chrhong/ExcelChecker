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

    if 0 == sht.getStatistic(ERROR):
        sht.dumpColumns('sale', ('E','B','D'), (VALUE_TYPE_TEXT, VALUE_TYPE_NUM, VALUE_TYPE_NUM))

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

    if 0 == sht.getStatistic(ERROR):
        sht.dumpColumns('buy', ('F','I','N'), (VALUE_TYPE_TEXT, VALUE_TYPE_NUM, VALUE_TYPE_NUM))

    return sht

def stock_table_check(xls, key):
    sht = xls.Sheet(xls, 'Sheet2')
    db_key = key

    stock_rule = {
        "TitleLine" : "A:关键字:合计", #keyColumn:begin:end
        "NoneKey" : "Fix:C", #Ignore/Mark/Fix/Delete
        "Add" : ['D','E','F','G','H'],
        "Keep" : ['B','C'], #Keep not None
        "CheckItem" : [],
        "CheckMethod" : []
    }

    sht.dupCombColumn(stock_rule, db_key)

    if 0 == sht.getStatistic(ERROR):
        #generate new combined excel
        modify_rule = {
            "key" : stock_rule["TitleLine"],
            "swap" : [('D','E'), ('J', 'K')],
            "swapCell" : [('5-D','5-E'), ('5-J','5-K')],
            "set" : [(), ()],
            "replace" : ["sale:H:I", "buy:F:G"],
            "titles" : ["开票索引", "零件号码"]
        }
        # sht.setColumnsFormatText(('J', 'K'))
        sht.modifyColData(modify_rule, "new_stock_file.xlsx")

    return sht

def test(xls, key):
    npfile_abs = "D:\\home\\X\\TestDoc\\npdata\\buy.npy"
    npdata = numpy.load(npfile_abs)
    npdatakey = list(npdata[0])
    for i in npdatakey:
        print(i.decode('unicode-escape'))

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
    root = EasyGUI("ExcelChecker v1.0", '500x205', UNRESIZEBLE)
    root.Entry("Source file", 48, (10,20))
    root.LogWindow(("white","black"), (48,5), (10,100), UNEDITABLE)
    root.Radiobutton("进项表", 1, (10,60))
    root.Radiobutton("销售表", 2, (110,60))
    root.Radiobutton("库存表", 3, (210,60))
    root.Button('打开', (8,1), (410,17), root.Browser)
    root.Button('检查', (8,4), (410,100), lambda:root.Thread(userChecker))
    root.mainloop()

def oneClick_mainloop():
    Registry('ExcelChecker.exe')

    try:
        if len(sys.argv) < 2:
            eprint("[Error] argv error! Use demo file in debug mode!")
            #debug file
            source_file = "test/部分进项表.xlsx"
        else:
            if IS_PYTHON2:
                source_file = ' '.join(sys.argv[1:]).strip("\'\"").decode('gbk')
            else:
                source_file = ' '.join(sys.argv[1:]).strip("\'\"")
            source_file = source_file.replace('\\', '/').replace('\"', '')

        eprint(source_file)

        type_list = ['进项表', '销售表', '库存表']
        check_type = len(type_list)

        for typei in type_list:
            if typei in source_file:
                check_type = type_list.index(typei)

        if check_type == len(type_list):
            eprint("[ERROR] %s not supported !" % source_file)
            eprint("[INFO] Valid file: %s !" % type_list.join(', '))
            RaiseException(HANG)

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
        RaiseException(HANG)