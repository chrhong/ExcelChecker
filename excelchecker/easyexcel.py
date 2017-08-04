#! -*- coding: utf-8 -*-
from __init__ import *
from icon import img

#EasyExcel constants
NORMAL = BLANK = 0
GOOD = GREEN = 4
ERROR = RED = 3
UNKNOW = GREY = 15
RESIZEBLE = EDITABLE = True
UNRESIZEBLE = UNEDITABLE = False

# CheckRule = {
#     "TitleLine" : "title"
#     "NoneKey" : "mark", #ignore/mark/correct/delete
#     "Add" : ['A','B','C','D'],
#     "Keep" : "last",
#     "CheckItem" : ['A:C','B'],
#     "CheckMethod" : [func1, func2]
# }

class EasyLog:
    """
    To manage print functions
    """
    def __init__(self):
        self.print_list = [print,]
    def registerPrintCb(self, print_func):
        if print_func not in self.print_list:
            self.print_list.append(print_func)
    def removePrintCb(self, print_func):
        if print_func not in self.print_list:
            self.print_list.pop(print_func)
    def lprint(self, log_str):
        for print_func in self.print_list:
            print_func(log_str)

#constants
easyLog = EasyLog()
eprint = easyLog.lprint

class EasyExcel:
    """
    To handle Excel easier.
    Remember to save the data is your problem.
    Operate on one workbook at one time.
    """
    def __init__(self, filename=None):
        """Open given file or create a new file"""
        self.xlApp = win32.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def save(self, newfilename=None):
        """save to new file if file name is given"""
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()
    def close(self):
        """close the open file"""
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def cpSheet(self, src_st, dst_st):
        """make a sheet copy"""
        shts = self.xlBook.Worksheets
        shts(src_st).Copy(None,shts(src_st))
        shts(shts(src_st).index + 1).Name = dst_st

    class Sheet:
        """
        To handle one sheet.
        """
        def __init__(self, xls, sheet):
            """Open a sheet or create a sheet"""
            self.taglist = [NORMAL,GOOD,ERROR,UNKNOW]
            self.statistic = [0,0,0,0]
            self.xlSheet = xls.xlBook.Worksheets(sheet)
            self.xls = xls
            if self.xlSheet == None:
                self.xlSheet = xls.xlBook.Worksheets.Add()
                self.xlSheet.Name = sheet
                self.xlSheet.Activate()

        def getCell(self, row, col):
            """Get value of one cell"""
            return self.xlSheet.Cells(row, col).Value

        def setCell(self, row, col, value):
            """set value of one cell"""
            self.xlSheet.Cells(row, col).Value = value

        def setCellformat(self, sheet, row, col):
            "format performance of one cell"
            sht = self.xlSheet
            # sht.Cells(row, col).Font.Size = 12            #字体大小
            sht.Cells(row, col).Font.Bold = True          #是否粗体
            sht.Cells(row, col).Name = "Arial"            #字体类型
            sht.Cells(row, col).Interior.ColorIndex = 3   #表格背景 RED = 3
            # sht.Range("A1").Borders.LineStyle = xlDouble
            sht.Cells(row, col).BorderAround(1,4)         #表格边框
            sht.Rows(row).RowHeight = 30                    #行高
            # sht.Cells(row, col).HorizontalAlignment = self.xls.xlRight
            sht.Cells(row, col).VerticalAlignment = -4135

        def markCell(self, row, col, tag):
            """mark a cell with color/tag, and collect statistics"""
            self.xlSheet.Cells(row, col).Interior.ColorIndex = tag
            tag_index = self.taglist.index(tag)
            self.statistic[tag_index] = self.statistic[tag_index] + 1

        def getStatistic(self, tag):
            """get a tag's statistics"""
            tag_index = self.taglist.index(tag)
            return self.statistic[tag_index]

        def inserRow(self, row):
            self.xlSheet.Rows(row).Insert(1)
        def deleteRow(self, row):
            self.xlSheet.Rows(row).Delete()

        def inserCol(self, col):
            self.xlSheet.Columns(col).Insert(1)
        def deleteCol(self, col):
            self.xlSheet.Columns(col).Delete()

        def getRow(self, row):
            sht = self.xlSheet
            ncols = sht.UsedRange.Columns.Count
            return sht.Range(sht.Cells(row, 1), sht.Cells(row, ncols)).Value

        def getColumn(self, col):
            sht = self.xlSheet
            nrows = sht.UsedRange.Rows.Count
            return sht.Range(sht.Cells(1, col), sht.Cells(nrows, col)).Value

        def getRange(self, row1, col1, row2, col2):
            """return a 2d array (i.e. tuple of tuples)"""
            sht = self.xlSheet
            return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

        def addPicture(self, pictureName, Left, Top, Width, Height):
            "Insert a picture in sheet"
            sht = self.xlSheet
            sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

        def cellAdd(self, row1, col1, row2, col2):
            sht = self.xlSheet
            cv1 = sht.Cells(row1, col1).Value
            cv2 = sht.Cells(row2, col2).Value
            cv1 = 0 if cv1 is None else cv1
            cv2 = 0 if cv2 is None else cv2
            sht.Cells(row2, col2).Value = cv1 + cv2

        def __isNotEmpty(self, row, col, value):
            sht = self
            if 'None' == value:
                sht.markCell(row, col, ERROR)
                return False
            else:
                sht.markCell(row, col, NORMAL)
                return True

        def __checkCheckRow(self, row, dbi, combRule):
            targetL = combRule["CheckItem"]
            methodL = combRule["CheckMethod"]
            keyCol = combRule["KeyColumn"]

            target_len = len(targetL)
            i = 0
            while i < target_len:
                paramsL = targetL[i].split(':')
                tag = methodL[i](self, row, dbi, paramsL)
                col = paramsL[0]
                self.markCell(row, col, tag)
                i = i + 1

        def __handleDupRow(self, dupRow, fstRow, combRule):
            sht = self
            add_list = combRule["Add"]
            keep_list = combRule["Keep"]

            if not keep_list and not add_list:
                sht.deleteRow(dupRow)
                return False

            for acol in add_list:
                sht.cellAdd(dupRow, chr(acol), fstRow, chr(acol))

            for kcol in keep_list:
                if  'None' == sht.getCell(fstRow, kcol):
                    sht.setCell(fstRow, kcol, sht.getCell(dupRow, kcol))

            sht.deleteRow(dupRow)
            return True

        def dupCombColumn(self, col, combRule, db_key):
            sht = self
            data_tuple = sht.getColumn(col)
            data_list = str(data_tuple).replace('.0','').replace('u\'','').replace('\'','').strip('(),').split(',), (')

            title_line = combRule["TitleLine"]

            valid_list = []
            i = len(data_list)
            while i > 0:
                #List  index range, [0,len()-1]
                #Excel index range, [1,len()]
                i = i - 1
                if title_line == data_list[i]:#python 2.x .decode('unicode-escape')
                    eprint("Meet title line to end check")
                    break

                if self.__isNotEmpty(i+1, col, data_list[i]):
                    if data_list[i] in valid_list:
                        eprint(sht.getCell(i+1, col))
                        first_index = data_list[i+1:].index(data_list[i]) + i + 1
                        if self.__handleDupRow(i+1, first_index+1, combRule):
                            data_list.pop(i) #update list
                            self.__checkCheckRow(first_index+1, data_list[i], combRule)
                    else:
                        valid_list.append(data_list[i])
                        try:
                            dbi = db_key.index(data_list[i])
                        except:
                            self.markCell(i+1, col, UNKNOW)
                        else:
                            self.__checkCheckRow(i+1, dbi, combRule)

            return

class EasyGUI():
    def __init__(self, title, size, resizeble):
        self.gui = Tk()
        self.fileEntry = None
        self.radioValue = None
        self.radioValueList = [""]
        self.radioList = [None]*5
        self.logWindow = None
        self.logFlush = 0

        tmp = open("tmp.ico","wb+")
        tmp.write(base64.b64decode(img))
        tmp.close()
        self.gui.iconbitmap("tmp.ico")
        os.remove("tmp.ico")

        self.gui.title(title)
        self.gui.geometry(size)
        # self.gui.iconbitmap('ExcelChecker.ico')
        self.gui.resizable(resizeble, resizeble)
        easyLog.registerPrintCb(self.WriteLogWindow)

    def Entry(self, title, length, place):
        file_font = ("微软雅黑", 10, "normal")
        file_location = StringVar()
        self.fileEntry = Entry(self.gui, text=title, font=file_font, textvariable=file_location, borderwidth=2, width=length)
        self.fileEntry.place(x=place[0], y=place[1], anchor=NW)
    def Browser(self):
        if not self.fileEntry:
            eprint("file entry is not init! cannot continue！")
            return
        file_path = askopenfilename()
        if not file_path is None:
            self.fileEntry.delete(0,END)
            self.fileEntry.insert(index=0, string=file_path)
            eprint("打开文件:\n"+file_path)

    def Button(self, title, size, place, action):
        button_font = ("微软雅黑", 10, "bold")
        open_button = Button(self.gui, fg="black", text=title, font=button_font, width=size[0], height=size[1], borderwidth=2, command=action)
        open_button.place(x=place[0], y=place[1], anchor=NW)

    def Thread(self, main_action):
        th=threading.Thread(target=main_action,args=(self.fileEntry.get(),self.radioValue))
        th.setDaemon(True)
        th.start()
        eprint("start action thread !")
        # th.join()
    def __radio_select(self, radio_var):
        self.radioValue = radio_var.get()
        selection = "你选择了" + self.radioValueList[self.radioValue]
        eprint(selection)
    def Radiobutton(self, title, value, place):
        radio_font = ("微软雅黑", 12, "normal")
        radio_var = IntVar()
        eprint(value)
        tempRadio = Radiobutton(self.gui, text=title, font=radio_font, variable=radio_var, value=value, command=lambda: self.__radio_select(radio_var))
        tempRadio.place(x=place[0], y=place[1], anchor=NW)
        if title not in self.radioList:
            self.radioValueList.insert(value,title)
        return tempRadio
    def LogWindow(self, color, size, place, editable):
        log_font = ("微软雅黑", 10, "normal")
        self.logWindow = Text(self.gui, fg=color[0], bg=color[1], font=log_font, relief=SUNKEN, width=size[0], height=size[1])
        self.logWindow.place(x=place[0], y=place[1], anchor=NW)
        if not editable:
            self.logWindow.bind("<KeyPress>", lambda e:"break")
    def WriteLogWindow(self, log_str):
        self.logFlush = self.logFlush + 1
        log_str = str(log_str)
        self.logWindow.insert(END, log_str + '\n')
        self.logWindow.see(END)
        if self.logFlush >= 10:
            self.logFlush = 0
            self.logWindow.update()

    def mainloop(self):
        self.gui.mainloop()