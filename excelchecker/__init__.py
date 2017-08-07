#__init__.py

import time
import re
import os
import sys
import base64
import numpy
import threading
import traceback
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
import pythoncom
import win32com.client as win32

IS_PYTHON2 = True if sys.version_info.major < 3 else False

if IS_PYTHON2: #python 2.x
    import _winreg as winreg
    reload(sys)
    sys.setdefaultencoding('utf-8')
else:          #python 3.x
    import winreg as winreg
