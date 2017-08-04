#! -*- coding: utf-8 -*-
from __init__ import *
from user import Checker
from easyexcel import *

def run():
    root = EasyGUI("ExcelChecker v0.2", '500x205', RESIZEBLE)
    root.Entry("Source file", 48, (10,20))
    root.LogWindow(("white","black"), (48,5), (10,100), UNEDITABLE)
    r1 = root.Radiobutton("进项表", 1, (10,60))
    r2 = root.Radiobutton("销售表", 2, (110,60))
    r3 = root.Radiobutton("库存表", 3, (210,60))
    root.Button('打开', (8,1), (410,17), root.Browser)
    root.Button('检查', (8,4), (410,100), lambda:root.Thread(Checker))
    root.mainloop()
