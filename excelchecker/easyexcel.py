#! -*- coding: utf-8 -*-
from __init__ import *
from icon import img

HANG = True
NOHANG = False

cn_pattern = re.compile(u'[\u4e00-\u9fa5]+')
def contain_cn(word):
    global cn_pattern
    match = cn_pattern.search(str(word).decode('utf8'))
    return match

def NP_FILENAME(key_str):
    return key_str + '.npy'

def RaiseException(hang=NOHANG):
    eprint(traceback.format_exc())
    eprint("Stop to see what error occurs!!!")
    if hang:
        eprint("You can Exit with 'Ctrl + C'...")
        while(1):
            pass
    sys.exit()


class Registry:
    """
    Windows Registry
    """
    def __init__(self, keyword):
        self.key = keyword
        if getattr(sys, 'frozen', False):
            root_path = os.path.dirname(sys.executable)
        elif __file__:
            root_path = os.path.dirname(__file__)

        config_path = root_path + "/" + "installed"

        try:
            with open(config_path, 'r') as f:
                eprint("[INFO] tool is already registered to contextmemu")
        except:
            self.__create(root_path, self.key)
            with open(config_path, 'a') as f:
                eprint("[INFO] register tool to contextmemu")
            sys.exit(0)

    def __create(self, workpath, keyword):
        self.key = keyword
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
            RaiseException(HANG)
            return False
        else:
            return True


class EasyLog:
    """
    To manage print functions
    """
    def __init__(self):
        #print in python 2.x is not function, cannot put into list
        self.print_list = []
    def registerPrintCb(self, print_func):
        if print_func not in self.print_list:
            self.print_list.append(print_func)
    def removePrintCb(self, print_func):
        if print_func not in self.print_list:
            self.print_list.pop(print_func)
    def lprint(self, log):
        log_str = str(log)
        if IS_PYTHON2:
            print(log_str).decode('utf-8') #print in python 2.x
        else:
            print(log_str)

        for print_func in self.print_list:
            print_func(log_str)

#constants
easyLog = EasyLog()
eprint = easyLog.lprint

# EasyExcel_CheckRule = {
#     "TitleLine" : "E:开票索引",
#     "NoneKey" : "Mark", #Ignore/Mark/Fix/Delete
#     "Add" : [],
#     "Keep" : [],
#     "CheckItem" : ['C','B:C'],
#     "CheckMethod" : [index_sell_price_check, index_sell_count_check],
# }

#EasyExcel constants
NORMAL = BLANK = 0
GOOD = GREEN = 4
ERROR = RED = 3
UNKNOW = GREY = 15
RESIZEBLE = EDITABLE = True
UNRESIZEBLE = UNEDITABLE = False
NUMPY_FOLDER = 'npdata'

class EasyExcel:
    """
    To handle Excel easier.
    Remember to save the data is your problem.
    Operate on one workbook at one time.
    """
    def __init__(self, filename=None):
        """Open given file or create a new file"""
        self.xlApp = win32.Dispatch('Excel.Application')
        self.database = []
        self.workpath = ""
        if filename:
            self.filename = filename
            self.workpath = filename[:filename.rfind('/')+1]
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def dbInit(self, db_file):
        db = self.database
        try:
            with open(db_file, 'r') as fd:
                for line in fd.readlines():
                    if line.strip().startswith('#') or '' == line.strip():
                        continue
                    line = ''.join(line.strip().split(' '))
                    line_list = line.split(';')
                    db.append(line_list)
            # eprint(db[0][1].strip('[]').split(',')[0])
            return db
        except:
            eprint("[ERROR] %s not exist" % db_file)
            RaiseException()

    def save(self, newfilename=None):
        """save to new file if file name is given"""
        if newfilename:
            # self.filename = newfilename
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
        def setCells(self, xytuple, valuetuple):
            xy_num = len(xytuple)
            i = 0
            while i < xy_num:
                row = int(xytuple[i].split('-')[0])
                col = xytuple[i].split('-')[1]
                value = valuetuple[i]
                self.setCell(row, col, value)
                i = i + 1

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

        def setColumnsFormatText(self, col_list):
            for col in col_list:
                self.xlSheet.Columns(col).NumberFormatLocal = "@"

        def markCell(self, row, col, tag):
            """mark a cell with color/tag, and collect statistics"""
            self.xlSheet.Cells(row, col).Interior.ColorIndex = tag
            tag_index = self.taglist.index(tag)
            self.statistic[tag_index] = self.statistic[tag_index] + 1

        def getStatistic(self, tag):
            """get a tag's statistics"""
            tag_index = self.taglist.index(tag)
            return self.statistic[tag_index]

        def insertRow(self, row):
            self.xlSheet.Rows(row).Insert(1)
        def deleteRow(self, row):
            self.xlSheet.Rows(row).Delete()

        def insertCol(self, col, data=None):
            insertData = 1 if data is None else data
            self.xlSheet.Columns(col).Insert(insertData)
        def deleteCol(self, col):
            self.xlSheet.Columns(col).Delete()

        def getRow(self, row):
            sht = self.xlSheet
            ncols = sht.UsedRange.Columns.Count
            return sht.Range(sht.Cells(row, 1), sht.Cells(row, ncols)).Value

        def getColumn(self, col):
            sht = self.xlSheet
            nrows = sht.UsedRange.Rows.Count
            data_tuple = sht.Range(sht.Cells(1, col), sht.Cells(nrows, col)).Value
            #this is paticaularly designed for the project, not sure if it will suitable for others
            data_list = str(data_tuple).replace('.0','').replace('u\'','').replace('\'','').strip('(),').split(',), (')
            return data_list
        def swapColumns(self, col1_tuple, col2_tuple):
            """
            swap two columns:
            Eg: Input ('B', 'C'), ('D', 'E')
            The B<->D, C<->E

            Note: MUST BE col1_tuple[i] < col2_tuple[i]
            """
            sht = self.xlSheet
            col_num = len(col1_tuple)
            i = 0
            while i < col_num:
                #.Copy() can only called once one time
                #Make two column copy together will mix up value
                # colCopy = sht.Columns(col1_tuple[i]).Value#Copy()
                # sht.Columns(col2_tuple[i]).Insert(colCopy)
                # sht.Columns(col1_tuple[i]).Delete()

                # colCopy = sht.Columns(col2_tuple[i]).Value#Copy()
                # sht.Columns(col1_tuple[i]).Insert(colCopy)
                # sht.Columns(chr(ord(col2_tuple[i]) + 1)).Delete()

                colCopy = self.getColumn(col1_tuple[i])
                self.insertCol(col2_tuple[i], colCopy)
                self.deleteCol(col1_tuple[i])

                colCopy = self.getColumn(col2_tuple[i])
                self.insertCol(col1_tuple[i], colCopy)
                self.deleteCol(chr(ord(col2_tuple[i]) + 1))

                i = i + 1

        def swapCells(self, cell1_tuple, cell2_tuple):
            """
            swap two cells:
            Eg: Input ('X1-Y1', 'X3-Y3'), ('X2-Y2', 'X4-Y4')
            Then [X1,Y1]<->[X3,Y3], [X2,Y2]<->[X4,Y4]
            """
            cell_num = len(cell1_tuple)
            i = 0
            while i < cell_num:
                row1 = int(cell1_tuple[i].split('-')[0])
                col1 = cell1_tuple[i].split('-')[1]
                row2 = int(cell2_tuple[i].split('-')[0])
                col2 = cell2_tuple[i].split('-')[1]
                cell1_value = self.getCell(row1, col1)
                print(cell1_value)
                self.setCell(row1, col1, self.getCell(row2, col2))
                self.setCell(row2, col2, cell1_value)
                i = i + 1

        def getSheet(self):
            sht = self.xlSheet
            ncol = sht.UsedRange.Columns.Count
            nrow = sht.UsedRange.Rows.Count
            return self.getRange(1,1,nrow,ncol)

        def getRange(self, row1, col1, row2, col2):
            """return a 2d array (i.e. tuple of tuples)"""
            sht = self.xlSheet
            return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
        def setRange(self, start_row, start_col, data_array):
            """
            set a 2d array, must be a 2d array
            Col set to col, row set to row
            """
            sht = self.xlSheet
            end_row = start_row + len(data_array) - 1
            end_col = chr(ord(start_col) + len(data_array[0]) - 1) #only support col < Z
            sht.Range(sht.Cells(start_row, start_col), sht.Cells(end_row, end_col)).Value = data_array

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

        def __isNotEmpty(self, row, col, value, NoneKey_list):
            sht = self
            mode = NoneKey_list[0]

            if 'None' == value:
                if mode == "Mark":
                    sht.markCell(row, col, ERROR)
                elif mode == "Fix":
                    fix_col = NoneKey_list[1]
                    sht.setCell(row, col, sht.getCell(row, fix_col))
                    sht.markCell(row, col, NORMAL)
                    return True #need dup check again after fixed
                elif mode == "Ignore":
                    pass
                elif mode == "Delete":
                    sht.deleteRow(row)
                    #not update datalist, would this cause issues ?
                else:
                    eprint("ERROR: unknown mode detect %s, mark error anyway" % mode)
                    sht.markCell(row, col, ERROR)

                return False
            else:
                sht.markCell(row, col, NORMAL)
                return True

        def __checkCheckRow(self, row, dbi, combRule):
            targetL = combRule["CheckItem"]
            methodL = combRule["CheckMethod"]
            keyCol = combRule["TitleLine"].split(':')[0]

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
                sht.cellAdd(dupRow, acol, fstRow, acol)

            for kcol in keep_list:
                if  'None' == sht.getCell(fstRow, kcol):
                    sht.setCell(fstRow, kcol, sht.getCell(dupRow, kcol))

            sht.deleteRow(dupRow)
            return True

        def dupCombColumn(self, combRule, db_key):
            sht = self
            title_line = combRule["TitleLine"].split(':')
            NoneKey_list = combRule["NoneKey"].split(':')
            col = title_line[0]
            data_list = sht.getColumn(col)

            valid_list = []
            i = len(data_list)
            while i > 0:
                #List  index range, [0,len()-1]
                #Excel index range, [1,len()]
                i = i - 1
                if title_line[1] == data_list[i]:#python 2.x .decode('unicode-escape')
                    eprint("Meet title line to end check")
                    break

                if self.__isNotEmpty(i+1, col, data_list[i], NoneKey_list):
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
                elif NoneKey_list[0] == 'Delete':
                    data_list.pop(i) #update list
            return

        def __get_npfilename(self,npfile):
            work_path = self.xls.workpath
            npdata_list = []
            expect_npfile = work_path + NUMPY_FOLDER + '/' + npfile + ".npy"
            if os.path.exists(work_path + NUMPY_FOLDER):
                for i in os.listdir(work_path + NUMPY_FOLDER):
                    npdata_list.append(i)
            else:
                os.makedirs(work_path + NUMPY_FOLDER)
                return expect_npfile

            i = 0
            while True:
                if i == 0:
                    temp_file = npfile + ".npy"
                else:
                    temp_file = npfile + str(i) + ".npy"

                if temp_file in npdata_list:
                    i = i +  1
                else:
                    return work_path + NUMPY_FOLDER + '/' + temp_file

        def dumpColumns(self, npfile_name, col_tuple):
            npfile_abs = self.__get_npfilename(npfile_name)
            datalist = []
            for col in col_tuple:
                datalist.append(self.getColumn(col))
            numpy.save(npfile_abs, datalist)

        def __npyEnvCheck(self, replist, npdata_list):
            work_path = self.xls.workpath
            if os.path.exists(work_path + NUMPY_FOLDER):
                for i in os.listdir(work_path + NUMPY_FOLDER):
                    npdata_list.append(i)
            else:
                return False

            for item in replist:
                npyfile = item.split(':')[0] + ".npy"
                if not npyfile in npdata_list:
                    return False

            return True
        def __oneNpyReplace(self, keylist, collist, npyname, titles=None):
            npfile_abs = self.xls.workpath + NUMPY_FOLDER + '/' + npyname
            keycol = keylist[0]

            xls_keylist = self.getColumn(keycol)
            if len(keylist) == 3:
                xls_keylist.pop()
            iRow = len(xls_keylist) + 1 #insert above the row

            npdata = numpy.load(npfile_abs)
            npdatakey = list(npdata[0])
            for item in npdatakey:
                if titles and item in titles:
                    continue #ignore titles
                indexOfnp = npdatakey.index(item)
                if item in xls_keylist:
                    row = xls_keylist.index(item) + 1
                else:
                    self.insertRow(iRow) #insert above the row
                    self.setCell(iRow, keycol, item)
                    row = iRow
                    xls_keylist.append(item)
                    iRow = iRow + 1
                i = 1
                for col in collist:
                    self.setCell(row, col, npdata.T[indexOfnp][i])
                    i = i + 1
                    

        def __npyReplace(self, keylist, rulelist, npdata_list, titles=None):
            rule_len = len(rulelist)
            npyname = rulelist[0] + ".npy"
            i = 0
            while True:
                if npyname in npdata_list:
                    self.__oneNpyReplace(keylist, rulelist[1:], npyname, titles)
                    i = i + 1
                    npyname = rulelist[0] + str(i) + ".npy"
                else:
                    break
            return

        def modifyColData(self, modify_rule, newfilename=None):
            keylist = modify_rule["key"].split(':')
            swaplist = modify_rule['swap']
            swapcells_list = modify_rule['swapCell']
            set_list = modify_rule['set']
            replist = modify_rule['replace']
            #this is temp designed for keywords different in other data.npy
            titles = modify_rule['titles']

            newfilename = newfilename if newfilename else "new.xlsx"
            newfile = self.xls.workpath + newfilename
            newfile = newfile.replace('/', '\\')
            npdata_list = []
            eprint('[INFO] prepare to combine numpy data...check env...')
            if self.__npyEnvCheck(replist, npdata_list):
                eprint('[INFO] begin to combine...')
                self.swapColumns(swaplist[0], swaplist[1])
                self.swapCells(swapcells_list[0], swapcells_list[1])
                self.setCells(set_list[0], set_list[1])
                for item in replist:
                    self.__npyReplace(keylist, item.split(':'), npdata_list, titles)
                # xls.save("D:\\home\\ExcelChecker\\new.xlsx")
                self.xls.save(newfile)
                eprint("[INFO] save result to file: %s" % (newfile))
            else:
                eprint('[ERROR] env checking failed!')
            return


class EasyGUI():
    def __init__(self, title, size, resizeble):
        self.gui = Tk()
        self.fileEntry = None
        self.radioValue = None
        self.radioVar = IntVar()
        self.radioList = [""]
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
        source_file = self.fileEntry.get()
        source_file = source_file.replace('\\', '/').replace('\"', '')
        check_type = self.radioValue

        if source_file == "" or check_type == 0:
            eprint("[ERROR]未选择文件 或者 未选择表格类型")
            return

        th=threading.Thread(target=main_action,args=(source_file, self.radioList, check_type))
        th.setDaemon(True)
        th.start()
        eprint("start action thread !")
        # th.join()
    def __radio_select(self, radio_var):
        self.radioValue = radio_var.get()
        selection = "你选择了" + self.radioList[self.radioValue]
        eprint(selection)
    def Radiobutton(self, title, value, place):
        radio_font = ("微软雅黑", 12, "normal")
        tempRadio = Radiobutton(self.gui, text=title, font=radio_font, variable=self.radioVar, value=value, command=lambda:self.__radio_select(self.radioVar))
        tempRadio.place(x=place[0], y=place[1], anchor=NW)
        if title not in self.radioList:
            self.radioList.insert(value, title)
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