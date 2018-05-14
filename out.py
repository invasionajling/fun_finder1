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
import random

class CommentDialog(wx.Dialog):
    def __init__(self, parent, title,fanhao,movie_type):
        super(CommentDialog, self).__init__(parent, title=fanhao)
        panel = wx.Panel(self)
        self.conSizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer = wx.GridSizer(3,1,5,5)
        self.fanhao = fanhao
        #初始化获取信息
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()
        self.score = 0
        # 得分这个小框
        boxScore = wx.BoxSizer(wx.HORIZONTAL)

        self.scoreList = ['0','1', '2', '3', '4', '5']
        self.scorebox = wx.RadioBox(self, label='得分', choices=self.scoreList,
                               majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        self.scorebox.Bind(wx.EVT_RADIOBOX, self.OnRadiogroupScore)
        boxScore.Add(self.scorebox, 0, wx.ALL | wx.EXPAND, 5)
        self.sizer.Add(boxScore)
        # sql_sel = "SELECT genre FROM av_genres WHERE type =1 "
        # cur.execute(sql_sel)
        # genre_info = cur.fetchall()
        # print(genre_info)

        # # 主题这个小框
        # zhuti = wx.BoxSizer(wx.HORIZONTAL)
        #
        # self.listZhuti = ['']
        # self.getListZhuti()
        # self.comboZhuti = wx.ComboBox(self, choices=self.listZhuti, size=(100,10),style=wx.CB_READONLY)
        # self.comboZhuti.SetSelection(0)
        #
        # self.zhuti_scoreList = ['-2', '-1', '0', '1', '2', '3']
        # self.zhuti_scorebox = wx.RadioBox(self, label='主题', choices=self.zhuti_scoreList,
        #                        majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        # zhuti.Add(self.comboZhuti, 0, wx.ALL | wx.EXPAND, 20)
        # zhuti.Add(self.zhuti_scorebox, 0, wx.ALL | wx.EXPAND, 5)
        # self.sizer.Add(zhuti)

        # # 角色这个小框
        # Juese = wx.BoxSizer(wx.HORIZONTAL)
        #
        # self.listJuese = ['']
        # self.getListJuese()
        # self.comboJuese = wx.ComboBox(self, choices=self.listJuese, size=(100,10),style=wx.CB_READONLY)
        # self.comboJuese.SetSelection(0)
        #
        # self.Juese_scoreList = ['-2', '-1', '0', '1', '2', '3']
        # self.Juese_scorebox = wx.RadioBox(self, label='角色', choices=self.Juese_scoreList,
        #                                   majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        # Juese.Add(self.comboJuese, 0, wx.ALL | wx.EXPAND, 20)
        # Juese.Add(self.Juese_scorebox, 0, wx.ALL | wx.EXPAND, 5)
        # self.sizer.Add(Juese)

        # # 服装这个小框
        # Fuzhuang = wx.BoxSizer(wx.HORIZONTAL)
        #
        # self.listFuzhuang = ['']
        # self.getListFuzhuang()
        # self.comboFuzhuang = wx.ComboBox(self, choices=self.listFuzhuang, size=(100,10),style=wx.CB_READONLY)
        # self.comboFuzhuang.SetSelection(0)
        #
        # self.Fuzhuang_scoreList = ['-2', '-1', '0', '1', '2', '3']
        # self.Fuzhuang_scorebox = wx.RadioBox(self, label='服装', choices=self.Fuzhuang_scoreList,
        #                                   majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        # Fuzhuang.Add(self.comboFuzhuang, 0, wx.ALL | wx.EXPAND, 20)
        # Fuzhuang.Add(self.Fuzhuang_scorebox, 0, wx.ALL | wx.EXPAND, 5)
        # self.sizer.Add(Fuzhuang)
        #
        # # 体型这个小框
        # Tixing = wx.BoxSizer(wx.HORIZONTAL)
        #
        # self.listTixing = ['']
        # self.getListTixing()
        # self.comboTixing = wx.ComboBox(self, choices=self.listTixing, size=(100,10),style=wx.CB_READONLY)
        # self.comboTixing.SetSelection(0)
        #
        # self.Tixing_scoreList = ['-2', '-1', '0', '1', '2', '3']
        # self.Tixing_scorebox = wx.RadioBox(self, label='体型', choices=self.Tixing_scoreList,
        #                                      majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        # Tixing.Add(self.comboTixing, 0, wx.ALL | wx.EXPAND,20)
        # Tixing.Add(self.Tixing_scorebox, 0, wx.ALL | wx.EXPAND, 5)
        # self.sizer.Add(Tixing)
        #
        # # 行为这个小框
        # Xingwei = wx.BoxSizer(wx.HORIZONTAL)
        #
        # self.listXingwei = ['']
        # self.getListXingwei()
        # self.comboXingwei = wx.ComboBox(self, choices=self.listXingwei, size=(100,10),style=wx.CB_READONLY)
        # self.comboXingwei.SetSelection(0)
        #
        # self.Xingwei_scoreList = ['-2', '-1', '0', '1', '2', '3']
        # self.Xingwei_scorebox = wx.RadioBox(self, label='行为', choices=self.Xingwei_scoreList,
        #                                    majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        # Xingwei.Add(self.comboXingwei, 0, wx.ALL | wx.EXPAND, 20)
        # Xingwei.Add(self.Xingwei_scorebox, 0, wx.ALL | wx.EXPAND, 5)
        # self.sizer.Add(Xingwei)
        #
        # # 玩法这个小框
        # Wanfa = wx.BoxSizer(wx.HORIZONTAL)
        #
        # self.listWanfa = ['']
        # self.getListWanfa()
        # self.comboWanfa = wx.ComboBox(self, choices=self.listWanfa, size=(100,10),style=wx.CB_READONLY)
        # self.comboWanfa.SetSelection(0)
        #
        # self.Wanfa_scoreList = ['-2', '-1', '0', '1', '2', '3']
        # self.Wanfa_scorebox = wx.RadioBox(self, label='玩法', choices=self.Wanfa_scoreList,
        #                                     majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        # Wanfa.Add(self.comboWanfa, 0, wx.ALL | wx.EXPAND, 20)
        # Wanfa.Add(self.Wanfa_scorebox, 0, wx.ALL | wx.EXPAND, 5)
        # self.sizer.Add(Wanfa)
        #
        # # 其他这个小框
        # Qita = wx.BoxSizer(wx.HORIZONTAL)
        #
        # self.listQita = ['']
        # self.getListQita()
        # self.comboQita = wx.ComboBox(self, choices=self.listQita, size=(100,10),style=wx.CB_READONLY)
        # self.comboQita.SetSelection(0)
        #
        # self.Qita_scoreList = ['-2', '-1', '0', '1', '2', '3']
        # self.Qita_scorebox = wx.RadioBox(self, label='其他', choices=self.Qita_scoreList,
        #                                   majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        # Qita.Add(self.comboQita, 0, wx.ALL | wx.EXPAND, 20)
        # Qita.Add(self.Qita_scorebox, 0, wx.ALL | wx.EXPAND, 5)
        # self.sizer.Add(Qita)



        #如果是无码或者有码视频
        # if movie_type ==0 or movie_type==1:
        #     sql_sel = "SELECT * FROM av_record WHERE fanhao = '" + self.fanhao + "'"
        #     cur.execute(sql_sel)
        #     av_info = cur.fetchall()
        #     print(av_info)
        #
        #     #获取各类别数量
        #     if av_info[0][6] is None:
        #         av_genres = ''
        #     else:
        #         av_genres = av_info[0][6]
        #     new_av_genres = av_genres.split(',')
        #     del new_av_genres[-1]
        #     print(new_av_genres)
        #     self.dict_genres = {}
        #     for av_genre in new_av_genres:
        #         if av_genre in self.listZhuti:
        #             self.comboZhuti.SetValue(av_genre)
        #         if av_genre in self.listJuese:
        #             self.comboJuese.SetValue(av_genre)
        #         if av_genre in self.listFuzhuang:
        #             self.comboFuzhuang.SetValue(av_genre)
        #         if av_genre in self.listTixing:
        #             self.comboTixing.SetValue(av_genre)
        #         if av_genre in self.listXingwei:
        #             self.comboXingwei.SetValue(av_genre)
        #         if av_genre in self.listWanfa:
        #             self.comboWanfa.SetValue(av_genre)
        #         if av_genre in self.listQita:
        #             self.comboQita.SetValue(av_genre)
        #
        #     # 获取各演员数量
        #     if av_info[0][7] is None:
        #         av_stars = ''
        #     else:
        #         av_stars = av_info[0][7]
        #         new_av_stars = av_stars.split(',')
        #         del new_av_stars[-1]
        #         print(new_av_stars)
        #         self.dict_stars = {}
        #         for av_star in new_av_stars:
        #             self.dict_stars[av_star] = '0'
        #             starList = ['-2', '-1', '0', '1', '2', '3']
        #             starbox = wx.RadioBox(self, label=av_star, choices=starList,
        #                                    majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        #             # 获取图片
        #             path = 'G:/fun_finder/stars/' + av_star + '.jpg'
        #             if os.path.exists(path) == False:
        #                 print('先获取图片')
        #                 def mdb_conn(password=""):
        #                     # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接
        #
        #                     str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        #                     conn = pypyodbc.win_connect_mdb(str)
        #                     return conn
        #
        #                 conn = mdb_conn()
        #                 cur = conn.cursor()
        #
        #                 # 获取各种信息
        #                 sql_sel = "SELECT * FROM av_stars WHERE star= '" + av_star + "'"
        #                 cur.execute(sql_sel)
        #                 star_info = cur.fetchall()
        #                 print(star_info)
        #                 star_name = star_info[0][1]
        #                 star_pid = star_info[0][3]
        #                 try:
        #                     avatar_requ = requests.get(
        #                     'https://pics.javbus.com/actress/' + str(star_pid) + '_a.jpg',timeout=4)
        #                     print('正在下载图片' + star_name)
        #                     string = 'stars/' + star_name + '.jpg'
        #                     fp = open(string, 'wb')
        #                     # 创建文件
        #                     fp.write(avatar_requ.content)
        #                     fp.close()
        #                     print('图片已经下载' + star_name)
        #                     image = wx.Image(path, wx.BITMAP_TYPE_JPEG)
        #                     image = image.Scale(50, 50)
        #                 except:
        #                     image = wx.Image('G:/fun_finder/stars/nowprinting.gif', wx.BITMAP_TYPE_GIF)
        #                     image = image.Scale(50, 50)
        #             else:
        #                 image = wx.Image(path, wx.BITMAP_TYPE_JPEG)
        #                 try:
        #                     image = image.Scale(50, 50)
        #                 except:
        #                     image = wx.Image('G:/fun_finder/stars/nowprinting.gif', wx.BITMAP_TYPE_GIF)
        #                     image = image.Scale(50, 50)
        #             bmp = wx.StaticBitmap(self, bitmap=image.ConvertToBitmap(),size=(120,50))
        #             star_box=wx.BoxSizer(wx.HORIZONTAL)
        #             star_box.Add(bmp,0,wx.ALL,10)
        #             star_box.Add(starbox,0,wx.ALL,6)
        #             self.sizer.Add(star_box)
        #             starbox.Bind(wx.EVT_RADIOBOX, lambda evt, avStar=av_star: self.OnRadiogroupStar(evt, avStar))


        #     self.conSizer.Add(self.sizer)
        #
        #     hbox = wx.BoxSizer(wx.HORIZONTAL)
        #     self.lastPlayed = wx.CheckBox(self, label='刚看完')
        #     self.btnComment = wx.Button(self, label='评分', size=(100, 30))
        #     self.btnComment.Bind(wx.EVT_BUTTON, self.onClickComment)
        #     hbox.Add(self.lastPlayed,0,wx.TOP,5)
        #     hbox.Add(self.btnComment, 0, wx.ALIGN_CENTER_HORIZONTAL, 10)
        #     self.conSizer.Add(hbox,0,wx.ALL|wx.ALIGN_CENTER_HORIZONTAL,10)
        # elif movie_type == 2:
        #     sql_sel = "SELECT * FROM west_record WHERE movie_name = '" + self.fanhao + "'"
        #     cur.execute(sql_sel)
        #     av_info = cur.fetchall()
        #     print(av_info)
        #
        #     # 获取各类别数量
        #     if av_info[0][5] is None:
        #         av_genres = ''
        #     else:
        #         av_genres = av_info[0][5]
        #     new_av_genres = av_genres.split(',')
        #     del new_av_genres[-1]
        #     print(new_av_genres)
        #     self.dict_genres = {}
        #     row_genre = 0
        #     for av_genre in new_av_genres:
        #         self.dict_genres[av_genre] = '0'
        #         genreList = ['-2', '-1', '0', '1', '2', '3']
        #         genrebox = wx.RadioBox(self, label=av_genre, choices=genreList,
        #                                majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        #         self.sizer.Add(genrebox,0,wx.ALIGN_RIGHT)
        #         row_genre = row_genre + 1
        #         genrebox.Bind(wx.EVT_RADIOBOX, lambda evt, avGenre=av_genre: self.OnRadiogroupGenre(evt, avGenre))
        #
        #     # 获取各演员数量
        #     if av_info[0][6] is None:
        #         av_stars = ''
        #     else:
        #         av_stars = av_info[0][6]
        #         new_av_stars = av_stars.split(',')
        #         del new_av_stars[-1]
        #         print(new_av_stars)
        #         self.dict_stars = {}
        #         for av_star in new_av_stars:
        #             self.dict_stars[av_star] = '0'
        #             starList = ['-2', '-1', '0', '1', '2', '3']
        #             starbox = wx.RadioBox(self, label=av_star, choices=starList,
        #                                   majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        #             # 获取图片
        #             path = 'G:/fun_finder/stars/' + av_star + '.jpg'
        #             if os.path.exists(path) == False:
        #                 print('先获取图片')
        #
        #                 def mdb_conn(password=""):
        #                     # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接
        #
        #                     str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        #                     conn = pypyodbc.win_connect_mdb(str)
        #                     return conn
        #
        #                 conn = mdb_conn()
        #                 cur = conn.cursor()
        #
        #                 # 获取各种信息
        #                 sql_sel = "SELECT * FROM av_stars WHERE star= '" + av_star + "'"
        #                 cur.execute(sql_sel)
        #                 star_info = cur.fetchall()
        #                 star_name = star_info[0][1]
        #                 star_pid = star_info[0][3]
        #                 headers = {'content-type': 'application/json',
        #                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
        #
        #                 star_requ=requests.get(star_pid,headers=headers).content
        #                 print(star_requ)
        #                 avatar_soup =  BeautifulSoup(star_requ, 'html.parser', from_encoding='utf-8')
        #                 avatar_url = avatar_soup.find('img',class_='noborder')
        #                 print(avatar_url)
        #                 if avatar_url is not None:
        #                     avatar_url = avatar_url.get('src')
        #                     avatar_requ = requests.get(avatar_url,headers=headers)
        #                     print('正在下载图片' + star_name)
        #                     string = 'stars/' + star_name + '.jpg'
        #                     fp = open(string, 'wb')
        #                     # 创建文件
        #                     fp.write(avatar_requ.content)
        #                     fp.close()
        #                     print('图片已经下载' + star_name)
        #             image = wx.Image(path, wx.BITMAP_TYPE_JPEG)
        #             try:
        #                 image = image.Scale(50, 50)
        #             except:
        #                 image = wx.Image('G:/fun_finder/stars/nowprinting.gif', wx.BITMAP_TYPE_GIF)
        #                 image = image.Scale(50, 50)
        #             bmp = wx.StaticBitmap(self, bitmap=image.ConvertToBitmap())
        #             star_box = wx.BoxSizer(wx.HORIZONTAL)
        #             star_box.Add(bmp, 0, wx.ALL, 7)
        #             star_box.Add(starbox)
        #             self.sizer.Add(star_box)
        #             starbox.Bind(wx.EVT_RADIOBOX, lambda evt, avStar=av_star: self.OnRadiogroupStar(evt, avStar))

            # 获取导演数量
            # if av_info[0][2] == '':
            #     self.score_director = '0'
            # else:
            #     self.av_director = av_info[0][2]
            #     directorList = ['-2', '-1', '0', '1', '2', '3']
            #     self.directorbox = wx.RadioBox(self, label=self.av_director, choices=directorList,
            #                                    majorDimension=1, style=wx.RA_SPECIFY_ROWS)
            #     self.directorbox.Bind(wx.EVT_RADIOBOX, self.OnRadiogroupDirector)
            #     self.sizer.Add(self.directorbox,0,wx.ALIGN_RIGHT )
            #     self.Bind(wx.EVT_RADIOBUTTON, self.OnRadiogroupDirector)

        self.conSizer.Add(self.sizer)
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.lastPlayed = wx.CheckBox(self, label='刚看完')
        self.btnComment = wx.Button(self, label='评分', size=(100, 30))
        self.btnComment.Bind(wx.EVT_BUTTON, self.onClickComment)
        hbox.Add(self.lastPlayed, 0, wx.TOP, 5)
        hbox.Add(self.btnComment, 0, wx.ALIGN_CENTER_HORIZONTAL, 10)
        self.conSizer.Add(hbox, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)

        self.SetSizer(self.conSizer)
        self.Fit()
        self.Center()

    def OnRadiogroupScore(self, e):
        self.score = self.scorebox.GetStringSelection()
        self.score = str(int(self.score)*100)
        print(self.score)

    def onClickComment(self, event):

        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 更新播放时间
        if self.lastPlayed.GetValue() == True:
            now = datetime.datetime.now()
            now = now.strftime('%Y-%m-%d %H:%M:%S')
            sql_sel = "UPDATE av_record SET last='" + str(now) + "' WHERE fanhao = '" + self.fanhao + "'"
            cur.execute(sql_sel)
            conn.commit()
            print('更新了播放时间')


        # 更新评分
        sql_sel = "UPDATE av_record SET score=0+'" + self.score + "' WHERE fanhao = '" + self.fanhao + "'"
        print(sql_sel)
        cur.execute(sql_sel)
        conn.commit()
        wx.MessageBox("评分更新完毕", "Message", wx.OK | wx.ICON_INFORMATION)



class PanelOut(wx.Panel):
    def __init__(self, parent):
        """Constructor"""
        wx.Panel.__init__(self, parent=parent,size=(1140, 760))
        self.hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.leftGrid = wx.FlexGridSizer(6,4,10,10)
        self.rightBox = wx.BoxSizer(wx.VERTICAL)


        #初始化变量
        self.selectedTypeAV = '日本'
        self.selectedDirector = '全部'
        self.selectedFaxing = '全部'
        self.selectedZhizuo = '全部'
        self.selectedXilie = '全部'
        self.selectedGenre= '全部'
        self.selectedStar = '全部'
        self.selectedEnjoyed = '全部'
        self.flag = 0
        self.typeShow = 0

        #常用这个小框
        self.OnClickShow(wx.EVT_BUTTON,self.typeShow)
        self.btnShowByScore = wx.Button(self,label='看推荐')
        self.btnShowByScore.Bind(wx.EVT_BUTTON,lambda evt, typeShow = 0: self.OnClickShow(evt, typeShow))
        self.btnShowByTime = wx.Button(self, label='看最新')
        self.btnShowByTime.Bind(wx.EVT_BUTTON,lambda evt, typeShow = 1: self.OnClickShow(evt, typeShow))
        self.btnShowByRandom = wx.Button(self, label='随机看')
        self.btnShowByRandom.Bind(wx.EVT_BUTTON,lambda evt, typeShow = 2: self.OnClickShow(evt, typeShow))

        #self.btnMore = wx.Button(self,label='更多')
        #self.btnMore.Bind(wx.EVT_BUTTON,self.OnClickMore)
        self.btnUpdateJav = wx.Button(self, label='更新日本')
        self.btnUpdateJav.Bind(wx.EVT_BUTTON, self.OnClickUpdateJav)
        self.btnUpdateWest = wx.Button(self, label='更新欧美')
        self.btnUpdateWest.Bind(wx.EVT_BUTTON, self.OnClickUpdateWest)
        play = wx.StaticBox(self, -1, '常用:')
        playSizer = wx.StaticBoxSizer(play, wx.VERTICAL)
        playSizer.Add(self.btnShowByScore, 0, wx.ALL | wx.EXPAND, 5)
        playSizer.Add(self.btnShowByTime, 0, wx.ALL | wx.EXPAND, 5)
        playSizer.Add(self.btnShowByRandom, 0, wx.ALL | wx.EXPAND, 5)


        #筛选这个小框
        shuaixuan = wx.StaticBox(self,-1,'筛选')
        shuaixuanSizer = wx.StaticBoxSizer(shuaixuan, wx.VERTICAL)

        #AV类型选择
        self.listTypeAV = ['日本', '欧美']
        self.comboType = wx.ComboBox(shuaixuan, choices=self.listTypeAV, style=wx.CB_READONLY|wx.CB_DROPDOWN)
        self.comboType.SetSelection(0)
        self.comboType.Bind(wx.EVT_COMBOBOX, self.OnChoiceTypeAV)

        #导演选择
        self.listDirector=['全部']
        self.getListDirector()
        self.comboDirector = wx.ComboBox(shuaixuan, choices=self.listDirector, style=wx.CB_READONLY)
        self.comboDirector.SetSelection(0)
        self.comboDirector.Bind(wx.EVT_COMBOBOX, self.OnChoiceDirector)

        # 发行商选择
        self.listFaxing = ['全部']
        self.getListFaxing()
        self.comboFaxing = wx.ComboBox(shuaixuan, choices=self.listFaxing, style=wx.CB_READONLY)
        self.comboFaxing.SetSelection(0)
        self.comboFaxing.Bind(wx.EVT_COMBOBOX, self.OnChoiceFaxing)

        # 制作选择
        self.listZhizuo = ['全部']
        self.getListZhizuo()
        self.comboZhizuo = wx.ComboBox(shuaixuan, choices=self.listZhizuo, style=wx.CB_READONLY)
        self.comboZhizuo.SetSelection(0)
        self.comboZhizuo.Bind(wx.EVT_COMBOBOX, self.OnChoiceZhizuo)

        # 系列选择
        self.listXilie = ['全部']
        self.getListXilie()
        self.comboXilie = wx.ComboBox(shuaixuan, choices=self.listXilie, style=wx.CB_READONLY)
        self.comboXilie.SetSelection(0)
        self.comboXilie.Bind(wx.EVT_COMBOBOX, self.OnChoiceXilie)

        # 类型选择
        self.listGenre = ['全部']
        self.getListGenre()
        self.comboGenre = wx.ComboBox(shuaixuan, choices=self.listGenre, style=wx.CB_READONLY)
        self.comboGenre.SetSelection(0)
        self.comboGenre.Bind(wx.EVT_COMBOBOX, self.OnChoiceGenre)

        # 演员选择
        self.listStars = ['全部']
        self.getListStars()
        self.comboStars = wx.ComboBox(shuaixuan, choices=self.listStars, style=wx.CB_READONLY)
        self.comboStars.SetSelection(0)
        self.comboStars.Bind(wx.EVT_COMBOBOX, self.OnChoiceStar)

        # 是否看过选择
        self.listEnjoyed = ['全部','看过','没看过']
        self.comboEnjoyed = wx.ComboBox(shuaixuan, choices=self.listEnjoyed, style=wx.CB_READONLY)
        self.comboEnjoyed.SetSelection(0)
        self.comboEnjoyed.Bind(wx.EVT_COMBOBOX, self.OnChoiceEnjoyed)

        shuaixuanSizer.Add(self.comboType, 0, wx.ALL | wx.EXPAND, 5)
        shuaixuanSizer.Add(self.comboDirector, 0, wx.ALL | wx.EXPAND, 5)
        shuaixuanSizer.Add(self.comboFaxing, 0, wx.ALL | wx.EXPAND, 5)
        shuaixuanSizer.Add(self.comboZhizuo, 0, wx.ALL | wx.EXPAND, 5)
        shuaixuanSizer.Add(self.comboXilie, 0, wx.ALL | wx.EXPAND, 5)
        shuaixuanSizer.Add(self.comboGenre, 0, wx.ALL | wx.EXPAND, 5)
        shuaixuanSizer.Add(self.comboStars, 0, wx.ALL | wx.EXPAND, 5)
        shuaixuanSizer.Add(self.comboEnjoyed, 0, wx.ALL | wx.EXPAND, 5)


        update = wx.StaticBox(self,-1,'更新')
        updateSizer = wx.StaticBoxSizer(update, wx.VERTICAL)
        updateSizer.Add(self.btnUpdateJav, 0, wx.ALL | wx.EXPAND, 5)
        updateSizer.Add(self.btnUpdateWest, 0, wx.ALL | wx.EXPAND, 5)


        self.rightBox.Add(playSizer, 0, wx.EXPAND|wx.TOP, 0)
        self.rightBox.Add(shuaixuanSizer, 0, wx.EXPAND | wx.TOP, 30)
        self.rightBox.Add(updateSizer, 0, wx.EXPAND | wx.TOP, 30)
        #self.rightBox.Add(self.btnMore, 0, wx.EXPAND | wx.TOP, 50)

        self.hbox.Add(self.leftGrid, 6, wx.ALL, 10)
        self.hbox.Add(self.rightBox, 5, wx.ALL,10)
        self.SetSizer(self.hbox)

    #播放
    def OnClickPlay(self,event,fanhao):
        rootdir = 'G:/fun_finder/资源/' + fanhao
        for filenames in os.walk(rootdir):
            if 'cover.jpg' in filenames[2]:
                filenames[2].remove('cover.jpg')
            if 'cover_front.jpg' in filenames[2]:
                filenames[2].remove('cover_front.jpg')
            if 'cover_back.jpg' in filenames[2]:
                filenames[2].remove('cover_back.jpg')

            if len(filenames[2]) > 1:
                os.startfile(rootdir)
            else:
                os.startfile(rootdir + '/'+ filenames[2][0])
                print(filenames)

    #进行评论
    def OnComment(self,event,fanhao,movie_type):
        '''
        text_obj = wx.TextDataObject()
        wx.TheClipboard.Open()
        if wx.TheClipboard.IsOpened() or wx.TheClipboard.Open():
            text_obj.SetText(fanhao)
            wx.TheClipboard.SetData(text_obj)
            wx.TheClipboard.Close()
        wx.MessageBox("已经复制到剪贴板", "Message", wx.OK | wx.ICON_INFORMATION)
        '''
        a = CommentDialog(self, "评分",fanhao,movie_type).ShowModal()
    #显示评论
    def OnShowComment(self,event,fanhao):
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 获取评论信息
        sql_sel = "SELECT comment FROM av_record WHERE fanhao = '" + str(fanhao) + "'"
        cur.execute(sql_sel)
        comment_info_cur =cur.fetchall()
        print(comment_info_cur)
        comment_info = comment_info_cur[0][0]
        if comment_info is None:
            wx.MessageBox('暂时无评论', "Message", wx.OK | wx.ICON_INFORMATION)
        else:
            wx.MessageBox(comment_info, "Message", wx.OK | wx.ICON_INFORMATION)

        cur.close()
        conn.close()

    #更新日本的评分
    def OnClickUpdateJav(self,event):
        # 通过所有番号
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 获取所有番号
        sql_sel = "SELECT fanhao FROM av_record"
        table_name='av_record'
        cur.execute(sql_sel)
        fanhaos = cur.fetchall()
        for i in range(len(fanhaos)):
            fanhao = fanhaos[i][0]
            print('-------------------正在执行:' + fanhao + '-----------------------')
            self.updateScore(fanhao,table_name)

        #处理通假符号
        sql_sel = "SELECT fanhao,genre,stars FROM av_record"
        table_name = 'av_record'
        cur.execute(sql_sel)
        genres = cur.fetchall()
        for i in range(len(genres)):
            genre = genres[i][1]
            fanhao = genres[i][0]
            star = genres[i][2]
            print('-------------------正在执行:' + genre + '-----------------------')
            genredone = self.SQLLIKETRANS(genre)
            if star is None:
                stardone=''
            else:
                stardone = self.SQLLIKETRANS(star)
            sql_update = "UPDATE av_record SET star_chr ='" + str(stardone) + "'WHERE fanhao ='" + str(fanhao) + "'"
            print(sql_update)
            cur.execute(sql_update)
            conn.commit()
            sql_update = "UPDATE av_record SET genre_chr ='" + str(genredone) + "'WHERE fanhao ='" + str(fanhao) + "'"
            cur.execute(sql_update)
            conn.commit()

        cur.close()
        conn.close()
        wx.MessageBox("数据更新完毕", "Message", wx.OK | wx.ICON_INFORMATION)

    # 更新欧美的评分
    def OnClickUpdateWest(self, event):
        # 通过所有番号
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 获取所有番号
        sql_sel = "SELECT ID FROM west_record"
        table_name = 'west_record'
        cur.execute(sql_sel)
        fanhaos = cur.fetchall()
        for i in range(len(fanhaos)):
            fanhao = fanhaos[i][0]
            print('-------------------正在执行:' + str(fanhao) + '-----------------------')
            self.updateScore(fanhao, table_name)
        cur.close()
        conn.close()
        wx.MessageBox("数据更新完毕", "Message", wx.OK | wx.ICON_INFORMATION)

    #输出更多个人喜好页面
    def output_html(self):

        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        fout = open("index.html", "w")
        fout.write("<!DOCTYPE html>")
        fout.write("<html>")
        fout.write("<head>")
        fout.write("<title>fun for ajling</title>")
        fout.write("<meta name='viewport' content='width=device-width, initial-scale=1.0'>")
        fout.write("<link href='bootstrap/css/bootstrap.min.css' rel='stylesheet' media='screen'>")
        fout.write("<link href='bootstrap/css/fun.css' rel='stylesheet' media='screen'>")
        fout.write("</head>")
        fout.write("<body>")
        fout.write("<div class='container' style='margin-top:40px'>")
        fout.write("<div class='row''>")
        fout.write("<div class='span8'>")
        fout.write("<h4>最喜欢的star</h4>")

        # 获取TOP10 stars
        sql_sel_star = "SELECT TOP 10 * FROM av_stars ORDER BY num DESC"
        cur.execute(sql_sel_star)
        star_infos = cur.fetchall()
        print(star_infos)

        for star_info in star_infos:
            fout.write("<div class='media'>")
            fout.write("<a class='pull-left avatar' target='_blank' href='https://www.javbus.com/star/" + str(
                star_info[3]) + "'>")
            fout.write("<img src='https://pics.javbus.com/actress/" + str(
                star_info[3]) + "_a.jpg' class='img-circle'></a>")
            fout.write("<div class='media-body''>")
            fout.write("<h4 class='media-heading'>" + str(star_info[1]) + "</h4>")
            fout.write("</div>")
            fout.write("</div>")
        fout.write("</div>")
        fout.write("<div class='span4''>")


        fout.write("<h4>最喜欢的类型</h4>")
        # 获取TOP10类型
        sql_sel = "SELECT TOP 10 * FROM av_genres ORDER BY num DESC"
        cur.execute(sql_sel)
        genre_infos = cur.fetchall()
        print(genre_infos)
        for genre_info in genre_infos:
            fout.write("<a class='label label-success' target='_blank' href='https://www.javbus.com/genre/" + str(
                genre_info[3]) + "' style='margin-left:20px'>" + str(genre_info[1]) + "</a>")

        fout.write("<h4 style='margin-top:40px'>最喜欢的导演</h4>")
        # 获取TOP10类型
        sql_sel = "SELECT TOP 5 * FROM av_director ORDER BY num DESC"
        cur.execute(sql_sel)
        director_infos = cur.fetchall()
        print(director_infos)
        for director_info in director_infos:
            if director_info[1] != '':
                fout.write("<a target='_blank' href='https://www.javbus.com/director/" + str(
                    director_info[3]) + "' style='margin-left:20px'>" + str(director_info[1]) + "</a><br>")

        fout.write("<h4 style='margin-top:40px'>最喜欢的系列</h4>")
        # 获取TOP10类型
        sql_sel = "SELECT TOP 5 * FROM av_xilie ORDER BY num DESC"
        cur.execute(sql_sel)
        xilie_infos = cur.fetchall()
        print(xilie_infos)
        for xilie_info in xilie_infos:
            if xilie_info[1] != '':
                fout.write("<a target='_blank' href='https://www.javbus.com/series/" + str(
                    xilie_info[3]) + "' style='margin-left:20px'>" + str(xilie_info[1]) + "</a><br>")

        fout.write("</div>")
        fout.write("</div>")

        fout.write("<script src='http://code.jquery.com/jquery.js'></script>")
        fout.write("<script src='bootstrap/js/bootstrap.min.js''></script>")
        fout.write("</body>")
        fout.write("</html>")

        cur.close()
        conn.close()

    #更新某个番号的得分，根据传入的表名通过不同的规则来判断
    def updateScore(self,fanhao,table_name):
        # 通过番号获取各种信息
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 获取各种信息
        if table_name == 'av_record':
            sql_sel = "SELECT * FROM av_record WHERE fanhao = '" + fanhao + "'"
            cur.execute(sql_sel)
            av_info = cur.fetchall()
            print(av_info)
            # 获取导演数量
            if av_info[0][2] == '' or av_info[0][2] is None:
                num_director = 0
            else:
                av_director = av_info[0][2]
                sql_sel = "SELECT num FROM av_director WHERE director = '" + av_director + "'"
                cur.execute(sql_sel)
                raw_direcotr = cur.fetchall()
                num_director = raw_direcotr[0][0]
            print('导演数量：' + str(num_director))

            # 获取发行商数量
            if av_info[0][3] == '' or av_info[0][3] is None:
                num_faxing = 0
            else:
                av_faxing = av_info[0][3]
                sql_sel = "SELECT num FROM av_faxing WHERE faxing = '" + av_faxing + "'"
                cur.execute(sql_sel)
                raw_faxing = cur.fetchall()
                num_faxing = raw_faxing[0][0]
            print('发行商数量：' + str(num_faxing))

            # 获取制作商数量
            if av_info[0][4] == '' or av_info[0][4] is None:
                num_zhizuo = 0
            else:
                av_zhizuo = av_info[0][4]
                sql_sel = "SELECT num FROM av_zhizuo WHERE zhizuo = '" + av_zhizuo + "'"
                cur.execute(sql_sel)
                num_zhizuo = cur.fetchall()[0][0]
            print('制作商数量：' + str(num_zhizuo))

            # 获取系列数量
            if av_info[0][5] == '' or av_info[0][5] is None:
                num_xilie = 0
            else:
                av_xilie = av_info[0][5]
                sql_sel = "SELECT num FROM av_xilie WHERE xilie = '" + av_xilie + "'"
                cur.execute(sql_sel)
                num_xilie = cur.fetchall()[0][0]
            print('系列数量：' + str(num_xilie))

            # 获取各类别数量
            if av_info[0][6] is None:
                av_genres = ''
            else:
                av_genres = av_info[0][6]
            new_av_genres = av_genres.split(',')
            del new_av_genres[-1]
            print(new_av_genres)
            num_genres = 0
            for av_genre in new_av_genres:
                sql_sel = "SELECT num FROM av_genres WHERE genre = '" + av_genre + "'"
                cur.execute(sql_sel)
                num_genre = cur.fetchall()[0][0]
                num_genres = num_genres + num_genre
                print(av_genre + '类别数量：' + str(num_genre))

            # 获取各演员数量
            if av_info[0][7] is None:
                av_stars = ''
            else:
                av_stars = av_info[0][7]
            new_av_stars = av_stars.split(',')
            del new_av_stars[-1]
            print(new_av_stars)
            num_stars = 0
            for av_star in new_av_stars:
                sql_sel = "SELECT num FROM av_stars WHERE star = '" + av_star + "'"
                cur.execute(sql_sel)
                raw_star = cur.fetchall()
                num_star = raw_star[0][0]
                num_stars = num_stars + num_star
                print(av_star + '数量：' + str(num_star))

            # 获取播放时间
            if av_info[0][8] is None or av_info[0][8] == '0':
                num_days = 120
            else:
                last_str = av_info[0][8]
                last = datetime.datetime.strptime(last_str, '%Y-%m-%d %H:%M:%S')
                now_str = datetime.datetime.now()
                now_str = now_str.strftime('%Y-%m-%d %H:%M:%S')
                now = datetime.datetime.strptime(now_str, '%Y-%m-%d %H:%M:%S')
                date_num_days = now - last
                num_days = date_num_days.days

            # 获取编辑得分
            if av_info[0][10] == '' or av_info[0][10] is None:
                num_otherscore = 0
            else:
                num_otherscore = float(av_info[0][10])*10
            print('编辑得分：' + str(num_zhizuo))
            # score = num_director + num_faxing + num_zhizuo + num_xilie + num_stars + num_genres * 3 - (
            #                                                                                           120 - num_days) * 10 + num_otherscore
            score = num_director + num_faxing + num_zhizuo + num_xilie + num_stars* 4 + num_genres * 2 + num_otherscore* 2

            sql_update = "UPDATE av_record SET score ='" + str(score) + "' WHERE fanhao ='" + fanhao + "'"
            cur.execute(sql_update)
            conn.commit()
            print('该番号更新得分为' + str(score))
            cur.close()
            conn.close()
        elif table_name == 'av_recommend':
            sql_sel = "SELECT * FROM av_recommend WHERE fanhao = '" + fanhao + "'"
            cur.execute(sql_sel)
            av_info = cur.fetchall()
            print(av_info)
            # 获取导演数量
            if av_info[0][2] == '':
                num_director = 0
            else:
                av_director = av_info[0][2]
                sql_sel = "SELECT num FROM av_director WHERE director = '" + av_director + "'"
                cur.execute(sql_sel)
                director_info=cur.fetchall()
                if len(director_info)==0:
                    num_director = 0
                else:
                    num_director = director_info[0][0]
            print('导演数量：' + str(num_director))

            # 获取制作商数量
            if av_info[0][3] == '':
                num_zhizuo = 0
            else:
                av_zhizuo = av_info[0][3]
                sql_sel = "SELECT num FROM av_zhizuo WHERE zhizuo = '" + av_zhizuo + "'"
                cur.execute(sql_sel)
                zhizuo_info = cur.fetchall()
                if len(zhizuo_info) == 0:
                    num_zhizuo = 0
                else:
                    num_zhizuo = zhizuo_info[0][0]
            print('制作商数量：' + str(num_zhizuo))


            # 获取发行商数量
            if av_info[0][4] == '':
                num_faxing = 0
            else:
                av_faxing = av_info[0][4]
                sql_sel = "SELECT num FROM av_faxing WHERE faxing = '" + av_faxing + "'"
                cur.execute(sql_sel)
                faxing_info = cur.fetchall()
                if len(faxing_info) == 0:
                    num_faxing = 0
                else:
                    num_faxing = faxing_info[0][0]
            print('发行商数量：' + str(num_faxing))



            # 获取系列数量
            if av_info[0][5] == '':
                num_xilie = 0
            else:
                av_xilie = av_info[0][5]
                sql_sel = "SELECT num FROM av_xilie WHERE xilie = '" + av_xilie + "'"
                cur.execute(sql_sel)
                xilie_info=cur.fetchall()
                if len(xilie_info) == 0:
                    num_xilie = 0
                else:
                    num_xilie =xilie_info[0][0]
            print('系列数量：' + str(num_xilie))

            # 获取各类别数量
            if av_info[0][6] is None:
                av_genres = ''
            else:
                av_genres = av_info[0][6]
            new_av_genres = av_genres.split(',')
            del new_av_genres[-1]
            print(new_av_genres)
            num_genres = 0
            for av_genre in new_av_genres:
                sql_sel = "SELECT num FROM av_genres WHERE genre = '" + av_genre + "'"
                cur.execute(sql_sel)
                genre_info = cur.fetchall()
                if len(genre_info) == 0:
                    num_genre = 0
                else:
                    num_genre =genre_info[0][0]
                num_genres = num_genres + num_genre
                print(av_genre + '类别数量：' + str(num_genre))

            # 获取各演员数量
            if av_info[0][7] is None:
                av_stars = ''
            else:
                av_stars = av_info[0][7]
            new_av_stars = av_stars.split(',')
            del new_av_stars[-1]
            print(new_av_stars)
            num_stars = 0
            for av_star in new_av_stars:
                sql_sel = "SELECT num FROM av_stars WHERE star = '" + av_star + "'"
                cur.execute(sql_sel)
                star_info = cur.fetchall()
                if len(star_info) == 0:
                    num_star = 0
                else:
                    num_star = star_info[0][0]
                num_stars = num_stars + num_star
                print(av_star + '数量：' + str(num_star))


            score = num_director + num_faxing + num_zhizuo + num_xilie + num_stars*2 + num_genres * 3
            sql_update = "UPDATE av_recommend SET score ='" + str(score) + "' WHERE fanhao ='" + fanhao + "'"
            cur.execute(sql_update)
            conn.commit()
            print('该番号更新得分为' + str(score))
            cur.close()
            conn.close()
        else:
            sql_sel = "SELECT * FROM west_record WHERE ID = " + str(fanhao) + ""
            print(sql_sel)
            cur.execute(sql_sel)
            av_info = cur.fetchall()
            print(av_info)
            # 获取导演数量
            if av_info[0][2] == '':
                num_director = 0
            else:
                av_director = av_info[0][2]
                sql_sel = "SELECT num FROM av_director WHERE director = '" + av_director + "'"
                cur.execute(sql_sel)
                director_info = cur.fetchall()
                if len(director_info) == 0:
                    num_director = 0
                else:
                    num_director = director_info[0][0]
            print('导演数量：' + str(num_director))

            # 获取制作商数量
            if av_info[0][3] == '':
                num_zhizuo = 0
            else:
                av_zhizuo = av_info[0][3]
                sql_sel = "SELECT num FROM av_zhizuo WHERE zhizuo = '" + av_zhizuo + "'"
                cur.execute(sql_sel)
                zhizuo_info = cur.fetchall()
                if len(zhizuo_info) == 0:
                    num_zhizuo = 0
                else:
                    num_zhizuo = zhizuo_info[0][0]
            print('制作商数量：' + str(num_zhizuo))


            # 获取系列数量
            if av_info[0][4] == '':
                num_xilie = 0
            else:
                av_xilie = av_info[0][4]
                sql_sel = "SELECT num FROM av_xilie WHERE xilie = '" + av_xilie + "'"
                cur.execute(sql_sel)
                xilie_info = cur.fetchall()
                if len(xilie_info) == 0:
                    num_xilie = 0
                else:
                    num_xilie = xilie_info[0][0]
            print('系列数量：' + str(num_xilie))

            # 获取各类别数量
            if av_info[0][5] is None:
                av_genres = ''
            else:
                av_genres = av_info[0][5]
            new_av_genres = av_genres.split(',')
            del new_av_genres[-1]
            print(new_av_genres)
            num_genres = 0
            for av_genre in new_av_genres:
                sql_sel = "SELECT num FROM av_genres WHERE genre = '" + av_genre + "'"
                cur.execute(sql_sel)
                genre_info = cur.fetchall()
                if len(genre_info) == 0:
                    num_genre = 0
                else:
                    num_genre = genre_info[0][0]
                num_genres = num_genres + num_genre
                print(av_genre + '类别数量：' + str(num_genre))

            # 获取各演员数量
            if av_info[0][6] is None:
                av_stars = ''
            else:
                av_stars = av_info[0][6]
            new_av_stars = av_stars.split(',')
            del new_av_stars[-1]
            print(new_av_stars)
            num_stars = 0
            for av_star in new_av_stars:
                sql_sel = "SELECT num FROM av_stars WHERE star = '" + av_star + "'"
                cur.execute(sql_sel)
                star_info = cur.fetchall()
                if len(star_info) == 0:
                    num_star = 0
                else:
                    num_star = star_info[0][0]
                num_stars = num_stars + num_star
                print(av_star + '数量：' + str(num_star))

            # 获取播放时间
            if av_info[0][7] is None or av_info[0][7] == '0':
                num_days = 120
            else:
                last_str = av_info[0][7]
                last = datetime.datetime.strptime(last_str, '%Y-%m-%d %H:%M:%S')
                now_str = datetime.datetime.now()
                now_str = now_str.strftime('%Y-%m-%d %H:%M:%S')
                now = datetime.datetime.strptime(now_str, '%Y-%m-%d %H:%M:%S')
                date_num_days = now - last
                num_days = date_num_days.days

            score = num_director + num_zhizuo + num_xilie + num_stars * 2 + num_genres * 3- ( 120 - num_days) * 4
            sql_update = "UPDATE west_record SET score ='" + str(score) + "' WHERE ID =" + str(fanhao) + ""
            cur.execute(sql_update)
            conn.commit()
            print('该番号更新得分为' + str(score))
            cur.close()
            conn.close()

    #打开更多个人喜好页面
    def OnClickMore(self,event):
        self.output_html()
        url = 'G:/fun_finder/index.html'
        webbrowser.open(url, new=0, autoraise=True)

    #获取导演列表
    def getListDirector(self):
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 导演列表
        random.seed()
        sql_sel = "SELECT director FROM av_director ORDER BY num DESC"
        cur.execute(sql_sel)
        infos = cur.fetchall()
        print(infos)
        for info in infos:
            nameDirector = info[0]
            self.listDirector.append(nameDirector)

    #获取发行商列表
    def getListFaxing(self):
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 导演列表
        random.seed()
        sql_sel = "SELECT faxing FROM av_faxing ORDER BY num DESC"
        cur.execute(sql_sel)
        infos = cur.fetchall()
        print(infos)
        for info in infos:
            nameFaxing = info[0]
            self.listFaxing.append(nameFaxing)

    # 获取制作商列表
    def getListZhizuo(self):
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 导演列表
        random.seed()
        sql_sel = "SELECT zhizuo FROM av_zhizuo ORDER BY num DESC"
        cur.execute(sql_sel)
        infos = cur.fetchall()
        print(infos)
        for info in infos:
            nameZhizuo = info[0]
            self.listZhizuo.append(nameZhizuo)

    # 获取系列列表
    def getListXilie(self):
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 导演列表
        random.seed()
        sql_sel = "SELECT xilie FROM av_xilie ORDER BY num DESC"
        cur.execute(sql_sel)
        infos = cur.fetchall()
        print(infos)
        for info in infos:
            nameXilie = info[0]
            self.listXilie.append(nameXilie)

    # 获取类型列表
    def getListGenre(self):
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接
            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 导演列表
        random.seed()
        sql_sel = "SELECT genre FROM av_genres ORDER BY num DESC"
        cur.execute(sql_sel)
        infos = cur.fetchall()
        print(infos)
        for info in infos:
            nameGenre = info[0]
            self.listGenre.append(nameGenre)

    # 获取演员列表
    def getListStars(self):
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接
            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 导演列表
        random.seed()
        sql_sel = "SELECT star FROM av_stars ORDER BY num DESC"
        cur.execute(sql_sel)
        infos = cur.fetchall()
        print(infos)
        for info in infos:
            nameStar = info[0]
            self.listStars.append(nameStar)

    #设定选择的AV类型
    def OnChoiceTypeAV(self, event):
        self.selectedTypeAV = self.comboType.GetStringSelection()
        print(self.selectedTypeAV)

    # 设定选择的导演
    def OnChoiceDirector(self, event):
        self.selectedDirector = self.comboDirector.GetStringSelection()
        print(self.selectedDirector)

    # 设定选择的发行商
    def OnChoiceFaxing(self, event):
        self.selectedFaxing = self.comboFaxing.GetStringSelection()
        print(self.selectedFaxing)

    # 设定选择的制作商
    def OnChoiceZhizuo(self, event):
        self.selectedZhizuo = self.comboZhizuo.GetStringSelection()
        print(self.selectedZhizuo)

    # 设定选择的系列
    def OnChoiceXilie(self, event):
        self.selectedXilie = self.comboXilie.GetStringSelection()
        print(self.selectedXilie)

    # 设定选择的类别
    def OnChoiceGenre(self, event):
        self.selectedGenre = self.comboGenre.GetStringSelection()
        print(self.selectedGenre)

    # 设定选择的演员
    def OnChoiceStar(self, event):
        self.selectedStar = self.comboStars.GetStringSelection()
        print(self.selectedStar)

    # 设定选择是否看过
    def OnChoiceEnjoyed(self, event):
        self.selectedEnjoyed = self.comboEnjoyed.GetStringSelection()
        print(self.selectedEnjoyed)


    # 显示视频
    def OnClickShow(self, event,typeShow):
        self.leftGrid.Clear(True)

        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 获取TOP20番号
        #如果是日本 那么 type=0和1，如果是欧美，那么type=2
        if self.selectedTypeAV == '日本':
            sqlTypeAV = 'av_record'

            if self.selectedDirector == '全部':
                sqlDirector = ''
            else:
                sqlDirector = "and director = '" + self.selectedDirector + "'"

            if self.selectedFaxing == '全部':
                sqlFaxing = ''
            else:
                sqlFaxing = "and faxing = '" + self.selectedFaxing + "'"

            if self.selectedZhizuo == '全部':
                sqlZhizuo = ''
            else:
                sqlZhizuo = "and zhizuo = '" + self.selectedZhizuo + "'"

            if self.selectedXilie == '全部':
                sqlXilie = ''
            else:
                sqlXilie = "and xilie = '" + self.selectedXilie + "'"

            if self.selectedGenre == '全部':
                sqlGenre = ''
            else:
                print(self.SQLLIKETRANS(self.selectedGenre))
                sqlGenre = "and genre_chr like '%" + self.selectedGenre + "%'"

            if self.selectedStar == '全部':
                sqlStar = ''
            else:
                sqlStar = "and star_chr like '%" + self.selectedStar + "%'"

            if self.selectedEnjoyed == '全部':
                sqlEnjoyed = ''
            elif self.selectedEnjoyed == '看过':
                sqlEnjoyed = "and last <> '0'"
            else:
                sqlEnjoyed = "and last = '0'"

            if self.flag == 0:
                sqlFlag ='DESC'
            else:
                sqlFlag =''

            if typeShow == 0:
                sqlOrderBy = 'score'
            elif typeShow == 1:
                sqlOrderBy = 'id'
            elif typeShow == 2:
                random.seed()
                sqlOrderBy ='Rnd( '+ str(random.random()) + ' - ID)'
            sql_sel = "SELECT TOP 20 * FROM " + sqlTypeAV + " WHERE ID <>0 " + sqlDirector + " " + sqlFaxing + "" + sqlZhizuo + "" + sqlXilie + " " + sqlGenre + " " + sqlStar + " " + sqlEnjoyed + " ORDER BY " + sqlOrderBy + " " + sqlFlag + ""
            print(sql_sel)
            cur.execute(sql_sel)
            infos = cur.fetchall()

            for info in infos:
                self.item_box = wx.BoxSizer(wx.VERTICAL)
                image = wx.Image('G:/fun_finder/资源/' + info[1] + '/cover.jpg', wx.BITMAP_TYPE_JPEG)
                w = image.GetWidth()
                h = image.GetHeight()
                image = image.Scale(201, 135)
                self.bmp = wx.StaticBitmap(self, bitmap=image.ConvertToBitmap())
                self.lblFanhao = wx.StaticText(self, -1, info[1], size=(100, 20))
                self.lblScore = wx.StaticText(self, -1, '得分：' + str(info[9]), style=wx.ALIGN_RIGHT, size=(100, 20))

                self.actionBox = wx.BoxSizer(wx.HORIZONTAL)
                self.btnPlay = wx.Button(self, label="播放", size=(100, 30))
                self.btnPlay.Bind(wx.EVT_BUTTON, lambda evt, fanhao=info[1]: self.OnClickPlay(evt, fanhao))
                self.btnComment = wx.Button(self, label='评分', size=(50, 30))
                self.btnComment.Bind(wx.EVT_BUTTON,
                                     lambda evt, fanhao=info[1], movie_type=0: self.OnComment(evt, fanhao, movie_type))
                self.btnShowComment = wx.Button(self, label='看评论', size=(50, 30))
                self.btnShowComment.Bind(wx.EVT_BUTTON, lambda evt, fanhao=info[1]: self.OnShowComment(evt, fanhao))
                self.infoBox = wx.BoxSizer(wx.HORIZONTAL)
                self.infoBox.Add(self.lblFanhao, 0, wx.EXPAND, 0)
                self.infoBox.AddStretchSpacer(1)
                self.infoBox.Add(self.lblScore, 0, wx.EXPAND, 0)
                self.actionBox.Add(self.btnPlay, 0, wx.ALL, 0)
                self.actionBox.Add(self.btnComment, 0, wx.ALL, 0)
                self.actionBox.Add(self.btnShowComment, 0, wx.ALL, 0)
                self.item_box.Add(self.bmp, 0, wx.RIGHT, 0)
                self.item_box.Add(self.infoBox, 0, wx.TOP, 5)
                self.item_box.Add(self.actionBox, 1, wx.TOP, 5)
                self.leftGrid.Add(self.item_box, 0, wx.BOTTOM, 10)
        else:
            sqlTypeAV= 'west_record'
            if self.selectedDirector == '全部':
                sqlDirector = ''
            else:
                sqlDirector = "and director = '" + self.selectedDirector + "'"


            if self.selectedZhizuo == '全部':
                sqlZhizuo = ''
            else:
                sqlZhizuo = "and studios = '" + self.selectedZhizuo + "'"

            if self.selectedXilie == '全部':
                sqlXilie = ''
            else:
                sqlXilie = "and series = '" + self.selectedXilie + "'"

            if self.selectedGenre == '全部':
                sqlGenre = ''
            else:
                print(self.SQLLIKETRANS(self.selectedGenre))
                sqlGenre = "and catagorys like '%" + self.selectedGenre + "%'"

            if self.selectedStar == '全部':
                sqlStar = ''
            else:
                sqlStar = "and stars like '%" + self.selectedStar + "%'"

            if self.selectedEnjoyed == '全部':
                sqlEnjoyed = ''
            else:
                sqlEnjoyed = "and last = '0'"

            if self.flag == 0:
                sqlFlag ='DESC'
            else:
                sqlFlag =''

            if typeShow == 0:
                sqlOrderBy = 'score'
            elif typeShow == 1:
                sqlOrderBy = 'id'
            elif typeShow == 2:
                random.seed()
                sqlOrderBy ='Rnd( '+ str(random.random()) + ' - ID)'
            sql_sel = "SELECT TOP 20 * FROM " + sqlTypeAV + " WHERE ID <>0 " + sqlDirector + "" + sqlZhizuo + "" + sqlXilie + " " + sqlGenre + " " + sqlStar + " " + sqlEnjoyed + " ORDER BY " + sqlOrderBy + " " + sqlFlag + ""
            print(sql_sel)
            cur.execute(sql_sel)
            infos = cur.fetchall()
            print(infos)
            for info in infos:
                self.item_box = wx.BoxSizer(wx.VERTICAL)
                image_back = wx.Image('G:/fun_finder/资源/' + info[1] + '/cover_back.jpg', wx.BITMAP_TYPE_JPEG)
                image_back = image_back.Scale(95, 135)
                self.bmp_back = wx.StaticBitmap(self, bitmap=image_back.ConvertToBitmap())
                image_front = wx.Image('G:/fun_finder/资源/' + info[1] + '/cover_front.jpg', wx.BITMAP_TYPE_JPEG)
                image_front = image_front.Scale(95, 135)
                self.bmp_front = wx.StaticBitmap(self, bitmap=image_front.ConvertToBitmap())
                self.lblFanhao = wx.StaticText(self, -1, info[1], size=(100, 20))
                self.lblScore = wx.StaticText(self, -1, '得分：' + str(info[8]), style=wx.ALIGN_RIGHT, size=(100, 20))
                self.imageBox = wx.BoxSizer(wx.HORIZONTAL)
                self.imageBox.Add(self.bmp_back, 1)
                self.imageBox.Add(self.bmp_front, 1)
                self.actionBox = wx.BoxSizer(wx.HORIZONTAL)
                self.btnPlay = wx.Button(self, label="播放", size=(100, 30))
                self.btnPlay.Bind(wx.EVT_BUTTON, lambda evt, fanhao=info[1]: self.OnClickPlay(evt, fanhao))
                self.btnComment = wx.Button(self, label='评分', size=(50, 30))
                self.btnComment.Bind(wx.EVT_BUTTON,
                                     lambda evt, fanhao=info[1], movie_type=2: self.OnComment(evt, fanhao, movie_type))
                self.infoBox = wx.BoxSizer(wx.HORIZONTAL)
                self.infoBox.Add(self.lblFanhao, 0, wx.EXPAND, 0)
                self.infoBox.AddStretchSpacer(1)
                self.infoBox.Add(self.lblScore, 0, wx.EXPAND, 0)
                self.actionBox.Add(self.btnPlay, 1, wx.ALL, 0)
                self.actionBox.Add(self.btnComment, 1, wx.ALL, 0)
                self.item_box.Add(self.imageBox, 0, wx.RIGHT, 0)
                self.item_box.Add(self.infoBox, 0, wx.TOP, 5)
                self.item_box.Add(self.actionBox, 1, wx.TOP, 5)
                self.leftGrid.Add(self.item_box, 0, wx.BOTTOM, 10)

        cur.close()
        conn.close()
        self.Layout()
        self.Refresh()
        self.Update()
        self.leftGrid.Layout()
        if self.flag == 0:
            self.flag = 1
        else:
            self.flag = 0

    def SQLLIKETRANS(self,jn):
        jn = jn.replace('ゴ','')
        jn = jn.replace('ガ', '')
        jn = jn.replace('ギ', '')
        jn = jn.replace('グ', '')
        jn = jn.replace('ゲ', '')
        jn = jn.replace('ザ', '')
        jn = jn.replace('ジ', '')
        jn = jn.replace('ズ', '')
        jn = jn.replace('ヅ', '')
        jn = jn.replace('デ', '')
        jn = jn.replace('ド', '')
        jn = jn.replace('ポ', '')
        jn = jn.replace('ベ', '')
        jn = jn.replace('プ', '')
        jn = jn.replace('ビ', '')
        jn = jn.replace('パ', '')
        jn = jn.replace('ヴ', '')
        jn = jn.replace('ボ', '')
        jn = jn.replace('ペ', '')
        jn = jn.replace('ブ', '')
        jn = jn.replace('ピ', '')
        jn = jn.replace('バ', '')
        jn = jn.replace('ヂ', '')
        jn = jn.replace('ダ', '')
        jn = jn.replace('ゾ', '')
        jn = jn.replace('ゼ', '')
        return jn
