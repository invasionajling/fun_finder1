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
from javin import PanelJavIn
from westin import PanelWestIn
from out import PanelOut
from downloadspider import PanelSpider

class Mywin(wx.Frame):
    def __init__(self, parent, title):
        super(Mywin, self).__init__(parent, title=title, size=(1140, 720))

        nb = wx.Notebook(self)
        panel_jav_in = PanelJavIn(nb)
        panel_west_in = PanelWestIn(nb)
        panel_out = PanelOut(nb)
        panel_spider=PanelSpider(nb)
        nb.AddPage(panel_jav_in,'日本入库')
        nb.AddPage(panel_west_in, '欧美入库')
        nb.AddPage(panel_out, '观看')
        nb.AddPage(panel_spider, '下载')
        self.Centre()
        self.Show()
        self.Fit()
app = wx.App()
Mywin(None, 'fun for ajling')
app.MainLoop()