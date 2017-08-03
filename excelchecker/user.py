#! -*- coding: utf-8 -*-
from __init__ import *
from easyexcel import *

DBDB = []
DBKEY = []

cn_pattern = re.compile(u'[\u4e00-\u9fa5]+')
def contain_cn(word):
    global cn_pattern
    match = cn_pattern.search(str(word).decode('utf8'))

    return match

def hang_up_to_watch_errors():
    eprint(traceback.format_exc())
    eprint("Stop to see what error occurs!!!")
    # eprint("You can Exit with 'Ctrl + C'...")
    # while(1):
    #     pass
    sys.exit()

def regist_contextmenu(workpath, keyword):
    workpath = workpath.replace('/', '\\')
    eprint('---')
    eprint(workpath)

    target_icon = workpath + "\\" + keyword
    target_cmd = workpath + "\\" + keyword + " \"%1\""

    try:
        #Not check if it is already exist, since we can overwrite it
        subkey = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,"*\\shell")
        targetkey = winreg.CreateKey(subkey,keyword.split('.')[0])
        winreg.SetValueEx(targetkey, "Icon", 0,  winreg.REG_SZ, target_icon)
        winreg.SetValue(targetkey, "command",  winreg.REG_SZ, target_cmd)
        winreg.CloseKey(subkey)
    except OSError:
        eprint("[Error] Contextmenu register failed!")
        hang_up_to_watch_errors()
        return False
    else:
        return True

def tool_env_check():
    if getattr(sys, 'frozen', False):
        root_path = os.path.dirname(sys.executable)
    elif __file__:
        root_path = os.path.dirname(__file__)

    config_path = root_path + "/" + "installed"

    try:
        with open(config_path, 'r') as f:
            eprint("[INFO] tool is already registered to contextmemu")
    except:
        regist_contextmenu(root_path, 'ExcelChecker.exe')
        with open(config_path, 'a') as f:
            eprint("[INFO] register tool to contextmemu")
        sys.exit(0)

def get_cost_database(database_file):
    db = []
    try:
        with open(database_file, 'r') as fd:
            for line in fd.readlines():
                if line.strip().startswith('#') or '' == line.strip():
                    continue
                line = ''.join(line.strip().split(' '))
                line_list = line.split(';')
                db.append(line_list)
        # eprint(db[0][1].strip('[]').split(',')[0])
        return db
    except:
        eprint("[ERROR] %s not exist" % database_file)
        hang_up_to_watch_errors()


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

def note_cell_init(xls, sheet):
    #Note Cell check
    sht = xls.Sheet(xls, sheet)
    value1 = sht.getCell(1, 'A')
    if "NOTE:" == value1:
        eprint("NOTE Cell is already exist!")
    else:
        eprint("Create NOTE Cell")
        sht.inserRow(1)
        sht.inserRow(2)
        sht.setCell(1,'A',"NOTE:")
        sht.setCell(1,'B',"ERROR")
        sht.markCell(1,'B',ERROR)
        sht.setCell(1,'C',"NORMAL")
        sht.markCell(1,'C',NORMAL)
        sht.setCell(1,'D',"NOT IN DB")
        sht.markCell(1,'D',UNKNOW)

def index_table_check(xls,db,key):
    note_cell_init(xls, 'Sheet1')

    sht = xls.Sheet(xls,'Sheet1')
    index_tuple = sht.getColumn('E')
    index_list = str(index_tuple).replace('.0','').replace('u\'','').replace('\'','').strip('(),').split(',), (')
    # eprint index_list
    # for i in index_list:
    #     eprint i.decode('unicode-escape') #中文打印

    valid_list = []
    i = len(index_list)
    while i > 0:
        #List  index range, [0,len()-1]
        #Excel index range, [1,len()]
        i = i - 1
        if u'开票索引' == index_list[i].decode('unicode-escape'):
            eprint("Meet title line to exit")
            break

        if 'None' == index_list[i]:
            sht.markCell(i+1, 'E', ERROR)
        else:
            sht.markCell(i+1, 'E', NORMAL)
            if index_list[i] in valid_list:
                eprint(sht.getCell(i+1, 'E'))
                sht.deleteRow(i+1)
            else:
                valid_list.append(index_list[i])
                try:
                    db_i = key.index(index_list[i])
                    #销售单价检查
                    price_value = float(sht.getCell(i+1, 'C'))
                    price_min = float(db[db_i][1].strip('[]').split(',')[0])
                    price_max = float(db[db_i][1].strip('[]').split(',')[1])
                    if price_value < price_min or price_value > price_max:
                        sht.markCell(i+1, 'C', ERROR)
                    else:
                        sht.markCell(i+1, 'C', NORMAL)
                    #销售数量检查
                    count_value = int(sht.getCell(i+1, 'B'))
                    count_min = 0
                    count_max = index_count_range(price_value)
                    if count_value > count_max:
                        sht.markCell(i+1, 'B', ERROR)
                    else:
                        sht.markCell(i+1, 'B', NORMAL)
                except:
                    sht.markCell(i+1, 'E', UNKNOW)
                    eprint("%s is not in the datebase" % index_list[i].decode('unicode-escape'))

    return

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
    global DBDB
    db = DBDB

    price_value = float(sht.getCell(row, params_list[0]))
    price_min = float(db[dbi][1].strip('[]').split(',')[0])
    price_max = float(db[dbi][1].strip('[]').split(',')[1])
    if price_value < price_min or price_value > price_max:
        return ERROR
    else:
        return NORMAL

def index_check_new(xls,db,key):
    sht = xls.Sheet(xls, 'Sheet1')
    db_key = key

    index_rule = {
        "KeyColumn" : 'E',
        "TitleLine" : "开票索引",
        "NoneKey" : "mark", #ignore/mark/correct/delete
        "Add" : [],
        "Keep" : [],
        "CheckItem" : ['C','B:C'],
        "CheckMethod" : [index_sell_price_check, index_sell_count_check],
    }

    sht.dupCombColumn('E', index_rule, db_key)


def buy_table_check(xls,db,key):
    note_cell_init(xls, 'Sheet1')

    sht = xls.Sheet(xls, 'Sheet1')
    num_tuple = sht.getColumn('F')
    num_list = str(num_tuple).replace('.0','').replace('u\'','').replace('\'','').strip('(),').split(',), (')
    #eprint num_list

    valid_list = []
    i = len(num_list)
    while i > 0:
        #List  index range, [0,len()-1]
        #Excel index range, [1,len()]
        i = i - 1
        if u'零件号码' == num_list[i].decode('unicode-escape'):
            eprint("Meet title line to exit")
            break

        if 'None' == num_list[i]:
            sht.markCell(i+1, 'F', ERROR)
        else:
            sht.markCell(i+1, 'F', NORMAL)
            if num_list[i] in valid_list:
                eprint(sht.getCell(i+1, 'F'))
                sht.deleteRow(i+1)
            else:
                valid_list.append(num_list[i])
                try:
                    db_i = key.index(num_list[i])
                    #销售单价检查
                    price_value = float(sht.getCell(i+1, 'K'))
                    price_min = float(db[db_i][1].strip('[]').split(',')[0])
                    price_max = float(db[db_i][1].strip('[]').split(',')[1])
                    if price_value < price_min or price_value > price_max:
                        sht.markCell(i+1, 'K', ERROR)
                    else:
                        sht.markCell(i+1, 'K', NORMAL)
                    #销售数量检查
                    count_value = int(sht.getCell(i+1, 'I'))
                    count_min = 0
                    count_max = index_count_range(price_value)
                    if count_value > count_max:
                        sht.markCell(i+1, 'I', ERROR)
                    else:
                        sht.markCell(i+1, 'I', NORMAL)
                except:
                    sht.markCell(i+1, 'F', UNKNOW)
                    eprint("%s is not in the datebase" % num_list[i].decode('unicode-escape'))

    return

def stock_table_check(xls,db,key):
    # 销售成本=IF(D6+F6-H6=0,E6+G6,ROUND(H6*((E6+G6)/(D6+F6)),2))
    # Have new method need to be taken in use ?
    sht = xls.Sheet(xls, 'Sheet2')
    key_tuple = sht.getColumn('A')
    key_list = str(key_tuple).replace('.0','').replace('u\'','').replace('\'','').strip('(),').split(',), (')

    valid_key = []
    i = len(key_list) - 1 #remove last 'Summary' line
    while i > 0:
        #List  index range, [0,len()-1]
        #Excel index range, [1,len()]
        i = i - 1
        key_str = key_list[i]

        if u'关键字' == key_str.decode('unicode-escape'):
            eprint("Meet title line to exit")
            break

        if 'None' == key_str:
            sht.setCell(i+1, 'A', sht.getCell(i+1, 'C'))

        if key_str in valid_key:
            first_index = key_list[i+1:].index(key_str) + i + 1

            col = ord('D')
            while col <= ord('H'):
                sht.cellAdd(i+1, chr(col), first_index+1, chr(col))
                col = col + 1

            if  'None' == sht.getCell(first_index+1, 'B'):
                sht.setCell(first_index+1, 'B', sht.getCell(i+1, 'B'))
            if  'None' == sht.getCell(first_index+1, 'C'):
                sht.setCell(first_index+1, 'C', sht.getCell(i+1, 'C'))

            eprint(sht.getCell(i+1, 'A'))
            sht.deleteRow(i+1)
            key_list.pop(i) #update list
        else:
            valid_key.append(key_str)

    return

excel_handler = {
    '库存表' : stock_table_check,
    '索引表' : index_check_new,
    '进项表' : buy_table_check
}

def Checker(source_file, check_type):
    type_list = ["","进项表","索引表","库存表"]
    try:
        #should be called in each thread, otherwise, execl open failed.
        pythoncom.CoInitialize()

        if source_file == "" or check_type == 0:
            eprint("[ERROR]未选择文件 或者 未选择表格类型")
            return

        source_file = source_file.replace('\\', '/').replace('\"', '')
        eprint(source_file)

        cost_db = []
        cost_key = []
        if not type_list[check_type] == '库存表':
            #product.db read, stock do not need db
            global DBDB, DBKEY
            cost_db = get_cost_database(source_file[:source_file.rfind('/')+1] + 'product.db')
            for li in cost_db:
                cost_key.append(li[0])
            DBDB = cost_db
            DBKEY = cost_key

        eprint('[INFO] 开始 %s 检查...' % type_list[check_type])
        start = time.clock()
        xls = EasyExcel(source_file)
        excel_handler[type_list[check_type]](xls,cost_db,cost_key)
        xls.save()
        end = time.clock()
        eprint("[INFO] %s 检查完成, 耗时: %f 秒" % (type_list[check_type], end - start))
        # eprint("[INFO] 错误: %d 个, 无法匹配: %d 个" % (xls.ERROR_COUNT, xls.NONE_COUNT))
        xls.close()
    except:
        hang_up_to_watch_errors()