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
import winreg as winreg #python 2.x _winreg

reload(sys)
sys.setdefaultencoding('utf-8')
