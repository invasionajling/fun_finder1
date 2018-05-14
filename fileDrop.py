import wx
import requests
from bs4 import BeautifulSoup
import os
import shutil
import re
import pypyodbc
import datetime
import webbrowser
import wx.html
import threading

class FileDropTarget(wx.FileDropTarget):
    def __init__(self, window,window2):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.window2 = window2

    def OnDropFiles(self, x, y, filePath):
        filename_raw = os.path.basename(filePath[0]).split('.')[0]
        self.window.SetLabel(str(filePath[0]))


        pattern = r'[a-zA-Z][a-zA-Z][a-zA-Z]?[a-zA-Z]?-?\d\d\d'
        re_fanhao_raw = re.findall(pattern, filename_raw)
        if re_fanhao_raw != []:
            print(re_fanhao_raw)
            re_fanhao = re_fanhao_raw[-1]
            re_re_fanhao = re.findall(r'-', re_fanhao)
            print(re_re_fanhao)
            if re_re_fanhao == []:
                filename = re_fanhao[:-3] + '-' + re_fanhao[-3:]
            else:
                filename = re_fanhao
        else:
            filename = filename_raw
        self.window2.SetValue(str(filename))
