#coding:utf-8

from win32com.client import Dispatch
import win32com.client as win32
import time
import re

"""
Excel structure:
[1] Log : Log Cell : 
[2] Title : "WHAT A FUCK"  报关数量  报关单价  报关金额  开票索引
[3] Data :  汽车配件(凸轮轴)	15	  221.48	3322.2	 5267994
"""
EXCEL_OFFSET = 2
DATA_OFFSET = EXCEL_OFFSET + 1

cn_pattern = re.compile(u'[\u4e00-\u9fa5]+')

def contain_cn(word):
    global cn_pattern
    match = cn_pattern.search(str(word))

    return match

class easyExcel:
    """A utility to make it easier to get at Excel.    Remembering
    to save the data is your problem, as is    error handling.
    Operates on one workbook at a time."""
    ERROR = 3
    NORMAL = 0
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

    def getColumnS(self, sheet, col):
        data_list = []
        sht = self.xlBook.Worksheets(sheet)
        nrows = sht.UsedRange.Rows.Count
        i = 1
        while i<=nrows:
            data_list.append(sht.Cells(i, col).Value)
            i = i + 1
        return data_list

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

def none_index_check(index_tuple):
    index_list = []
    i=0
    for item in index_tuple:
        i = i + 1 #the num is for Excel, start from 1
        if i <= EXCEL_OFFSET: #skip data before title
            continue

        item_str = str(item).strip('(),')
        index_list.append(item_str)
        if item_str == 'None':
            xls.markCell('sheet1',i,'E',xls.ERROR)
        else:
            xls.markCell('sheet1',i,'E',xls.NORMAL)

    return index_list


def dup_index_check(xls, index_list):
    # delete invalid duplicate index
    l1 = index_list
    l2 = []
    i = len(l1)
    while i > 0:
        i = i - 1 #list index, [0,len(l1)-1]
        if 'None' == l1[i] or (l1[i][1] >= u'\u4e00' and l1[i][1] <= u'\u9fa5'):
            #print(i)
            #print(xls.getCell('Sheet1',i+DATA_OFFSET,'E'))
            continue

        if l1[i] in l2:
            print(xls.getCell('Sheet1',i+DATA_OFFSET,'E'))
            xls.deleteRow('Sheet1', i+DATA_OFFSET)
        else:
            l2.append(l1[i])

    return
def get_database(product_db):
    with open('C:/Users/hongg/Desktop/Python/X/product.db','r',encoding='UTF-8') as fd:
        for line in fd.readlines():
            if line.strip().startswith('#') or '' == line.strip():
                continue
            line = ''.join(line.strip().split(' '))
            line_list = line.split(';')
            product_db.append(line_list)

    print(product_db)


def index_table_check(xls):
    #Log Cell check
    log_cell = xls.getCell('Sheet1', 1, 1)
    if "Log Cell" == log_cell:
        print("Log Cell is already exist!")
    else:
        print("Create Log Cell")
        xls.inserRow('Sheet1', 1)
        xls.setCell('Sheet1', 1, 1,"Log Cell:")

    sht = xls.xlBook.Worksheets('Sheet1')
    nrows = sht.UsedRange.Rows.Count
    index_tuple = sht.Range(sht.Cells(1, 'E'), sht.Cells(nrows, 'E')).Value
    index_list = str(index_tuple).replace('\'','').strip('(),').split(',), (')
    valid_list = []

    i = len(index_list)
    while i > 0:
        #List  index range, [0,len()-1]
        #Excel index range, [1,len()]
        i = i - 1
        if '品名' == index_list[i]:
            print("Meet title line to exit")
            break
        elif 'None' == index_list[i]:
            xls.markCell('Sheet1', i+1, 'E', xls.ERROR)
        elif contain_cn(index_list[i]):
            xls.markCell('Sheet1', i+1, 'E', xls.NORMAL)
            continue
        else: #4995266, duplicate is not removed
            xls.markCell('Sheet1', i+1, 'E', xls.NORMAL)
            if index_list[i] in valid_list:
                print(xls.getCell('Sheet1', i+1, 'E'))
                xls.deleteRow('Sheet1', i+1)
            else:
                valid_list.append(index_list[i])

    return


def stock_table_check(xls):
    # 销售成本=IF(D6+F6=0,0,ROUND(H6*((E6+G6)/(D6+F6)),2))
    # Have new method need to be taken in use ?
    sht = xls.xlBook.Worksheets('Sheet2')
    nrows = sht.UsedRange.Rows.Count
    key_tuple = sht.Range(sht.Cells(1, 'A'), sht.Cells(nrows, 'A')).Value
    key_list = str(key_tuple).replace('\'','').strip('(),').split(',), (')

    valid_key = []
    i = len(key_list) - 1 #remove last 'Summary' line
    while i > 0:
        #List  index range, [0,len()-1]
        #Excel index range, [1,len()]
        i = i - 1
        key_str = key_list[i]

        if '关键字' == key_str:
            print("Meet title line to exit")
            break
        if 'None' == key_str:
            sht.Cells(i+1, 'A').Value = sht.Cells(i+1, 'C').Value
        elif contain_cn(key_str):
            continue
        elif key_str in valid_key:
            first_index = key_list[i+1:].index(key_str) + i + 1

            col = ord('D')
            while col <= ord('H'):
                sht.Cells(first_index+1, chr(col)).Value = xls.CellAdd(sht.Cells(first_index+1, chr(col)).Value, sht.Cells(i+1, chr(col)).Value)
                col = col + 1

            print(sht.Cells(i+1, 'A').Value)
            sht.Rows(i+1).Delete()
            key_list.pop(i) #update list
        else:
            valid_key.append(key_str)

    return


if __name__ == "__main__":
    #product.db
    # product_db = []
    # get_database(product_db)

    """
    xls = easyExcel('C:/Users/hongg/Desktop/Python/X/捷恩比17年5月库存.1.xlsx')
    start = time.clock()
    stock_table_check(xls)
    xls.save()
    xls.close()
    end = time.clock()
    print("stock_table_check: %f s" % (end - start))
    """

    # index_tuple = xls.getColumn('Sheet1', 'E')

    # index_list = none_index_check(index_tuple)
    # dup_index_check(xls, index_list)

    xls = easyExcel('C:/Users/hongg/Desktop/Python/X/去库存索引表.1.xlsx')
    start = time.clock()
    index_table_check(xls)
    xls.save()
    xls.close()
    end = time.clock()
    print("index_table_check: %f s" % (end - start))