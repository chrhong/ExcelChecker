#coding:utf-8

from win32com.client import Dispatch
import win32com.client as win32
import _winreg as winreg
import time
import re
import os
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

cn_pattern = re.compile(u'[\u4e00-\u9fa5]+')

def contain_cn(word):
    global cn_pattern
    match = cn_pattern.search(str(word).decode('utf8'))

    return match

def hang_up_to_watch_errors():
    print("Stop to see what error occurs!!!")
    print("You can Exit with 'Ctrl + C'...")
    while(1):
        pass

    sys.exit()


class easyExcel:
    """A utility to make it easier to get at Excel.    Remembering
    to save the data is your problem, as is    error handling.
    Operates on one workbook at a time."""
    NORMAL = 0
    ERROR = 3
    GOOD = 4
    GREY = 15
    #Open file or create file
    def __init__(self, filename=None):
        self.xlApp = win32.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''
    def save(self, newfilename=None):
        if newfilename: #save to new file if file name is given
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()
    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp
    def getCell(self, sheet, row, col):
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value
    def setCell(self, sheet, row, col, value):
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value
    def setCellformat(self, sheet, row, col):
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        # sht.Cells(row, col).Font.Size = 24#字体大小
        # sht.Cells(row, col).Font.Bold = True#是否黑体
        # sht.Cells(row, col).Name = "Arial"#字体类型
        sht.Cells(row, col).Interior.ColorIndex = 3#表格背景 RED = 3
        # #sht.Range("A1").Borders.LineStyle = xlDouble
        # sht.Cells(row, col).BorderAround(1,4)#表格边框
        # sht.Rows(3).RowHeight = 30#行高
        # sht.Cells(row, col).HorizontalAlignment = -4131 #水平居中xlCenter
        # sht.Cells(row, col).VerticalAlignment = -4160 #
    def markCell(self, sheet, row, col, color):
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Interior.ColorIndex = color

    def inserRow(self,sheet,row):
        sht = self.xlBook.Worksheets(sheet)
        sht.Rows(row).Insert(1)
    def deleteRow(self, sheet, row):
        sht = self.xlBook.Worksheets(sheet)
        sht.Rows(row).Delete()#删除行
        #sht.Columns(row).Delete()#删除列
    def getColumn(self, sheet, col):
        sht = self.xlBook.Worksheets(sheet)
        nrows = sht.UsedRange.Rows.Count
        return sht.Range(sht.Cells(1, col), sht.Cells(nrows, col)).Value
    def getColumnS(self, sheet, col):
        data_list = []
        sht = self.xlBook.Worksheets(sheet)
        nrows = sht.UsedRange.Rows.Count
        i = 1
        while i<=nrows:
            print sht.Cells(i, col).Value
            data_list.append(sht.Cells(i, col).Value)
            i = i + 1
        return data_list
    def getRange(self, sheet, row1, col1, row2, col2):  #获得一块区域的数据，返回为一个二维元组
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):  #插入图片
        "Insert a picture in sheet"
        sht = self.xlBook.Worksheets(sheet)
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)
    def cpSheet(self, before):  #复制工作表
        "copy sheet"
        shts = self.xlBook.Worksheets
        shts(1).Copy(None,shts(1))
    def CellAdd(self,value1,value2):
        ret_val = 0
        if value1 is None:
            ret_val = value2
        elif value2 is None:
            ret_val = value1
        else:
            ret_val = value1 + value2
        return ret_val


def regist_contextmenu(workpath, keyword):
    workpath = workpath.replace('/', '\\')
    print('---')
    print(workpath)

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
        print("[Error] Contextmenu register failed!")
        hang_up_to_watch_errors()
        return False
    else:
        return True

def tool_env_check():
    if getattr(sys, 'frozen', False):
        root_path = os.path.dirname(sys.executable)
    elif __file__:
        root_path = os.path.dirname(__file__)

    config_path = root_path + "/" + "enabled"

    try:
        with open(config_path, 'r') as f:
            print("[INFO] tool is already registered to contextmemu")
    except:
        regist_contextmenu(root_path, 'ExcelChecker.exe')
        with open(config_path, 'a') as f:
            print("[INFO] register tool to contextmemu")
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
        # print(db[0][1].strip('[]').split(',')[0])
        return db
    except:
        print("[ERROR] %s not exist" % database_file)
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
    value1 = xls.getCell(sheet,1,'A')
    if "NOTE:" == value1:
        print("NOTE Cell is already exist!")
    else:
        print("Create NOTE Cell")
        xls.inserRow(sheet,1)
        xls.inserRow(sheet,2)
        xls.setCell(sheet,1,'A',"NOTE:")
        xls.setCell(sheet,1,'B',"ERROR")
        xls.markCell(sheet,1,'B',xls.ERROR)
        xls.setCell(sheet,1,'C',"NORMAL")
        xls.markCell(sheet,1,'C',xls.NORMAL)
        xls.setCell(sheet,1,'D',"NOT IN DB")
        xls.markCell(sheet,1,'D',xls.GREY)

def index_table_check(xls,db,key):
    note_cell_init(xls, 'Sheet1')

    sht = xls.xlBook.Worksheets('Sheet1')
    nrows = sht.UsedRange.Rows.Count
    index_tuple = sht.Range(sht.Cells(1, 'E'), sht.Cells(nrows, 'E')).Value
    index_list = str(index_tuple).replace('.0','').replace('u\'','').replace('\'','').strip('(),').split(',), (')
    # print index_list
    # for i in index_list:
    #     print i.decode('unicode-escape') #中文打印

    valid_list = []
    i = len(index_list)
    while i > 0:
        #List  index range, [0,len()-1]
        #Excel index range, [1,len()]
        i = i - 1
        if u'开票索引' == index_list[i].decode('unicode-escape'):
            print("Meet title line to exit")
            break

        if 'None' == index_list[i]:
            xls.markCell('Sheet1',i+1,'E',xls.ERROR)
        else:
            xls.markCell('Sheet1', i+1, 'E', xls.NORMAL)
            if index_list[i] in valid_list:
                print(xls.getCell('Sheet1',i+1,'E'))
                xls.deleteRow('Sheet1', i+1)
            else:
                valid_list.append(index_list[i])
                try:
                    db_i = key.index(index_list[i])
                    #销售单价检查
                    price_value = float(xls.getCell('Sheet1',i+1,'C'))
                    price_min = float(db[db_i][1].strip('[]').split(',')[0])
                    price_max = float(db[db_i][1].strip('[]').split(',')[1])
                    if price_value < price_min or price_value > price_max:
                        xls.markCell('Sheet1',i+1,'C',xls.ERROR)
                    else:
                        xls.markCell('Sheet1',i+1,'C',xls.NORMAL)
                    #销售数量检查
                    count_value = int(xls.getCell('Sheet1',i+1,'B'))
                    count_min = 0
                    count_max = index_count_range(price_value)
                    if count_value > count_max:
                        xls.markCell('Sheet1',i+1,'B',xls.ERROR)
                    else:
                        xls.markCell('Sheet1',i+1,'B',xls.NORMAL)
                except:
                    xls.markCell('Sheet1',i+1,'E',xls.GREY)
                    print("%s is not in the datebase" % index_list[i].decode('unicode-escape'))

    return


def buy_table_check(xls,db,key):
    note_cell_init(xls, 'Sheet1')

    sht = xls.xlBook.Worksheets('Sheet1')
    nrows = sht.UsedRange.Rows.Count
    num_tuple = sht.Range(sht.Cells(1, 'F'), sht.Cells(nrows, 'F')).Value
    num_list = str(num_tuple).replace('.0','').replace('u\'','').replace('\'','').strip('(),').split(',), (')
    #print num_list

    valid_list = []
    i = len(num_list)
    while i > 0:
        #List  index range, [0,len()-1]
        #Excel index range, [1,len()]
        i = i - 1
        if u'零件号码' == num_list[i].decode('unicode-escape'):
            print("Meet title line to exit")
            break

        if 'None' == num_list[i]:
            xls.markCell('Sheet1',i+1,'F',xls.ERROR)
        else:
            xls.markCell('Sheet1', i+1, 'F', xls.NORMAL)
            if num_list[i] in valid_list:
                print(xls.getCell('Sheet1',i+1,'F'))
                xls.deleteRow('Sheet1', i+1)
            else:
                valid_list.append(num_list[i])
                try:
                    db_i = key.index(num_list[i])
                    #销售单价检查
                    price_value = float(xls.getCell('Sheet1',i+1,'K'))
                    price_min = float(db[db_i][1].strip('[]').split(',')[0])
                    price_max = float(db[db_i][1].strip('[]').split(',')[1])
                    if price_value < price_min or price_value > price_max:
                        xls.markCell('Sheet1',i+1,'K',xls.ERROR)
                    else:
                        xls.markCell('Sheet1',i+1,'K',xls.NORMAL)
                    #销售数量检查
                    count_value = int(xls.getCell('Sheet1',i+1,'I'))
                    count_min = 0
                    count_max = index_count_range(price_value)
                    if count_value > count_max:
                        xls.markCell('Sheet1',i+1,'I',xls.ERROR)
                    else:
                        xls.markCell('Sheet1',i+1,'I',xls.NORMAL)
                except:
                    xls.markCell('Sheet1',i+1,'F',xls.GREY)
                    print("%s is not in the datebase" % num_list[i].decode('unicode-escape'))

    return


def stock_table_check(xls,db,key):
    # 销售成本=IF(D6+F6-H6=0,E6+G6,ROUND(H6*((E6+G6)/(D6+F6)),2))
    # Have new method need to be taken in use ?
    sht = xls.xlBook.Worksheets('Sheet2')
    nrows = sht.UsedRange.Rows.Count
    key_tuple = sht.Range(sht.Cells(1, 'A'), sht.Cells(nrows, 'A')).Value
    key_list = str(key_tuple).replace('.0','').replace('u\'','').replace('\'','').strip('(),').split(',), (')

    valid_key = []
    i = len(key_list) - 1 #remove last 'Summary' line
    while i > 0:
        #List  index range, [0,len()-1]
        #Excel index range, [1,len()]
        i = i - 1
        key_str = key_list[i]

        if u'关键字' == key_str.decode('unicode-escape'):
            print("Meet title line to exit")
            break

        if 'None' == key_str:
            sht.Cells(i+1, 'A').Value = sht.Cells(i+1, 'C').Value

        if key_str in valid_key:
            first_index = key_list[i+1:].index(key_str) + i + 1

            col = ord('D')
            while col <= ord('H'):
                sht.Cells(first_index+1, chr(col)).Value = xls.CellAdd(sht.Cells(first_index+1, chr(col)).Value, sht.Cells(i+1, chr(col)).Value)
                col = col + 1

            if  'None' == sht.Cells(first_index+1, 'B').Value:
                sht.Cells(first_index+1, 'B').Value = sht.Cells(i+1, 'B').Value
            if  'None' == sht.Cells(first_index+1, 'C').Value:
                sht.Cells(first_index+1, 'C').Value = sht.Cells(i+1, 'C').Value

            print(sht.Cells(i+1, 'A').Value)
            sht.Rows(i+1).Delete()
            key_list.pop(i) #update list
        else:
            valid_key.append(key_str)

    return

excel_handler = {
    'stock' : stock_table_check,
    'index' : index_table_check,
    'buy'   : buy_table_check
}

if __name__ == "__main__":
    try:
        tool_env_check()

        if len(sys.argv) < 2:
            print("[Error] argv error! Use demo file in debug mode!")
            #debug file
            source_file = u"D:/userdata/chrhong/Desktop/Python/X/去库存索引表.1.xlsx"
        else:
            source_file = ' '.join(sys.argv[1:]).strip("\'\"")
            source_file = source_file.decode('gbk').replace('\\', '/').replace('\"', '')

        print source_file

        if u'进项' in source_file:
            file_type = 'buy'
        elif u'索引' in source_file:
            file_type = 'index'
        elif u'库存' in source_file:
            file_type = 'stock'
        else:
            print("[ERROR] %s not supported !" % source_file)
            print(u"[INFO] Valid file: '进项', '索引', '库存' !")
            hang_up_to_watch_errors()

        cost_db = []
        cost_key = []
        if not file_type == 'stock':
            #product.db read, stock do not need db
            cost_db = get_cost_database(source_file[:source_file.rfind('/')+1] + 'product.db')
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
        hang_up_to_watch_errors()

