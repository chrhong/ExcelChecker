#! -*- coding: utf-8 -*-
from __init__ import *
from user import Checker
from icon import img
from easyexcel import *

CHECK_TYPE = 0
LOG_FLUSH_COUNT = 0

def printTBox(log_str):
    global LOG_FLUSH_COUNT
    LOG_FLUSH_COUNT = LOG_FLUSH_COUNT + 1

    log_str = str(log_str)
    LOG_BOX.insert(END,log_str + '\n')
    LOG_BOX.see(END)
    if LOG_FLUSH_COUNT >= 10:
        LOG_FLUSH_COUNT = 0
        LOG_BOX.update()
    # print(log_str)#python 2.x .decode('utf-8')

def type_select(type_var):
    global CHECK_TYPE
    CHECK_TYPE = type_var.get()
    type_list = ["","进项表","索引表","库存表"]
    selection = "你选择了" + type_list[CHECK_TYPE]
    eprint(selection)

def browse(source_file):
    file_path = askopenfilename()
    if not file_path is None:
        source_file.delete(0,END)
    source_file.insert(index=0, string=file_path)
    eprint("打开文件:\n"+file_path)

def back_job(file_name):
    th=threading.Thread(target=Checker,args=(file_name,CHECK_TYPE))
    th.setDaemon(True)
    th.start()
    # th.join()
    # eprint("back thread stopped!")

def run():
    root = Tk()

    tmp = open("tmp.ico","wb+")
    tmp.write(base64.b64decode(img))
    tmp.close()
    root.iconbitmap("tmp.ico")
    os.remove("tmp.ico")

    root.title("ExcelChecker v0.2")
    root.geometry('500x205')
    # root.iconbitmap('D:\\userdata\\chrhong\\Desktop\\Python\\X\\ExcelChecker.ico')
    root.resizable(False, False)
    file_location = StringVar()

    type_var = IntVar()

    file_font = ("微软雅黑", 10, "normal")
    source_file = Entry(root, text = 'Source file:', font=file_font, textvariable=file_location, borderwidth=2, width = 48)
    source_file.place(x=10, y=20, anchor=NW)

    log_box_font = file_font
    global LOG_BOX
    LOG_BOX = Text(root, fg="white", bg="black", font=log_box_font, relief=SUNKEN, width=48, height=5)
    LOG_BOX.place(x=10, y=100, anchor=NW)
    LOG_BOX.bind("<KeyPress>", lambda e:"break")

    easyLog.registerPrintCb(printTBox)

    radio_font = ("微软雅黑", 12, "normal")
    R1 = Radiobutton(root, text="进项表", font=radio_font, variable=type_var, value=1, command=lambda: type_select(type_var))
    R1.place(x=10, y=60, anchor=NW)
    R2 = Radiobutton(root, text="索引表", font=radio_font, variable=type_var, value=2, command=lambda: type_select(type_var))
    R2.place(x=110, y=60, anchor=NW)
    R3 = Radiobutton(root, text="库存表", font=radio_font, variable=type_var, value=3, command=lambda: type_select(type_var))
    R3.place(x=210, y=60, anchor=NW)

    button_font = ("微软雅黑", 10, "bold")
    button_browse = Button(root, fg="black", text='打开', font=button_font, width=8, borderwidth=2, command=lambda: browse(source_file))
    button_browse.place(x=410, y=17, anchor=NW)

    check_browse = Button(root, fg="black", bg="green", text='检查', font=button_font, width=8, height=4, borderwidth=3, command=lambda: back_job(source_file.get()))
    check_browse.place(x=410, y=100, anchor=NW)

    root.mainloop()
