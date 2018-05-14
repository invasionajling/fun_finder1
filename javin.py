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
from fileDrop import FileDropTarget

class PanelJavIn(wx.Panel):
    def __init__(self, parent):
        """Constructor"""
        wx.Panel.__init__(self, parent=parent,size=(1140, 550))

        conBox = wx.BoxSizer(wx.VERTICAL)
        hbox1 =wx.BoxSizer(wx.HORIZONTAL)
        l1 = wx.StaticText(self, label="番号：")
        self.txt_fanhao = wx.TextCtrl(self)
        self.btn_search = wx.Button(self,label="搜索")
        self.btn_search.Bind(wx.EVT_BUTTON, self.OnClickSearch)
        self.btn_searchStar = wx.Button(self, label="查演员")
        self.btn_searchStar.Bind(wx.EVT_BUTTON, self.OnClickSearchStar)
        self.btn_searchScore = wx.Button(self, label="查编辑得分")
        self.btn_searchScore.Bind(wx.EVT_BUTTON, self.OnClickSearchScore)
        self.btn_searchComment = wx.Button(self, label="查评论")
        self.btn_searchComment.Bind(wx.EVT_BUTTON, self.OnClickSearchComment)
        self.lblStates = wx.StaticText(self, label="运行状态：", size=(100,20))

        hbox1.Add(l1,0,wx.ALL,13)
        hbox1.Add(self.txt_fanhao, 0, wx.ALL,10)
        hbox1.Add(self.btn_search, 0, wx.ALL, 7)
        hbox1.Add(self.btn_searchStar, 0, wx.ALL, 7)
        hbox1.Add(self.btn_searchScore, 0, wx.ALL, 7)
        hbox1.Add(self.btn_searchComment, 0, wx.ALL, 7)
        hbox1.Add(self.lblStates, 0, wx.ALL, 13)

        hbox2 = wx.BoxSizer(wx.HORIZONTAL)

        self.lblDrag = wx.StaticText(self, label="", style=wx.ALIGN_CENTRE | wx.ST_NO_AUTORESIZE,
                                     size=(700, 523))
        self.lblDrag.SetBackgroundColour('#cccccc')


        cover = wx.StaticBox(self, -1, '封面:')
        self.coverSizer = wx.StaticBoxSizer(cover, wx.VERTICAL)
        self.coverSizer.Add(self.lblDrag, 0, wx.ALL | wx.EXPAND, 10)
        hbox2.Add(self.coverSizer, 0, wx.ALL, 10)


        hbox2_right = wx.BoxSizer(wx.VERTICAL)
        font = wx.Font(18, wx.SWISS, wx.NORMAL, wx.BOLD)
        self.lblName = wx.StaticText(self, label="名称：", style=wx.ALIGN_CENTRE)
        self.lblName.SetFont(font)
        self.lblDirector = wx.StaticText(self, label="导演：", style=wx.ALIGN_CENTRE)
        self.lblZhizuo = wx.StaticText(self, label="制作：", style=wx.ALIGN_CENTRE)
        self.lblFaxing = wx.StaticText(self, label="发行：", style=wx.ALIGN_CENTRE)
        self.lblXilie = wx.StaticText(self, label="系列：", style=wx.ALIGN_CENTRE)
        self.lblGenres = wx.StaticText(self, label="类型：", style=wx.ALIGN_CENTRE)
        self.lblStars = wx.StaticText(self, label="演员：", style=wx.ALIGN_CENTRE)
        self.lblOtherscore = wx.StaticText(self, label="编辑得分：", style=wx.ALIGN_CENTRE)
        self.lblComment = wx.StaticText(self, label="评论：", style=wx.ALIGN_LEFT|wx.ST_NO_AUTORESIZE,size=(500,120))
        self.btn_updateStar = wx.Button(self, label="更新演员")
        self.btn_updateStar.Bind(wx.EVT_BUTTON, self.OnClickUpdateStar)
        lblList = ['日本有码', '日本无码']
        self.rbox = wx.RadioBox(self, label='类型', pos=(736, 64), choices=lblList,
                                majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        self.rbox.Bind(wx.EVT_RADIOBOX, self.onRadioBox)
        self.btn_inStroe = wx.Button(self, label="入库", size=(382, 65))
        self.btn_inStroe.Bind(wx.EVT_BUTTON, self.OnClickStore)
        self.btn_updateScore = wx.Button(self, label="更新编辑得分")
        self.btn_updateScore.Bind(wx.EVT_BUTTON, self.OnClickUpdateScore)
        self.btn_updateComment = wx.Button(self, label="更新评论")
        self.btn_updateComment.Bind(wx.EVT_BUTTON, self.OnClickUpdateComment)
        hbox2_right.Add(self.lblName,0,wx.BOTTOM|wx.TOP,10)
        hbox2_right.Add(self.lblDirector, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblZhizuo, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblFaxing, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblXilie, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblGenres, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblStars, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblOtherscore, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblComment, 0, wx.BOTTOM, 30)
        hbox2_right.Add(self.rbox, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.btn_inStroe, 0, wx.BOTTOM, 10)
        btnBox = wx.BoxSizer(wx.HORIZONTAL)
        btnBox.Add(self.btn_updateStar, 0, wx.RIGHT, 10)
        btnBox.Add(self.btn_updateScore, 0, wx.RIGHT, 10)
        btnBox.Add(self.btn_updateComment,0,wx.RIGHT,10)
        hbox2_right.Add(btnBox, 0, wx.ALL, 0)
        hbox2.Add(hbox2_right, 0, wx.ALL, 10)

        conBox.Add(hbox1,0,wx.ALL,0)
        conBox.Add(hbox2, 0, wx.ALL, 0)
        self.SetSizer(conBox)



        self.fanhao = ''
        self.fanhaoPlus=''
        self.av_director=''
        self.pid_director=''
        self.av_faxing=''
        self.pid_faxing=''
        self.av_zhizuo=''
        self.pid_zhizuo=''
        self.av_xilie=''
        self.pid_xilie=''
        self.av_genres=[]
        self.new_av_genres=''
        self.pid_genres=[]
        self.av_stars=[]
        self.new_av_stars=''
        self.pid_stars=[]
        self.other_socre = ''
        self.comment = ''
        self.type = 0

        dropTarget = FileDropTarget(self.lblDrag,self.txt_fanhao)
        self.lblDrag.SetDropTarget(dropTarget)

        self.fanhao = self.txt_fanhao.GetValue()
        self.txt_fanhao.SetValue(self.fanhao)

    def OnClickSearch(self, event):
        self.fanhao = ''
        self.fanhaoPlus = ''
        self.av_director = ''
        self.pid_director = ''
        self.av_faxing = ''
        self.pid_faxing = ''
        self.av_zhizuo = ''
        self.pid_zhizuo = ''
        self.av_xilie = ''
        self.pid_xilie = ''
        self.av_genres = []
        self.new_av_genres = ''
        self.pid_genres = []
        self.av_stars = []
        self.new_av_stars = ''
        self.pid_stars = []
        self.other_socre='0'
        self.comment=''
        self.fanhao=self.txt_fanhao.GetValue()
        self.lblStates.SetLabel('正在搜索番号：'+self.fanhao)
        if self.fanhao is None:
            return
        else:
            headers = {'content-type': 'application/json',
                       'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
            requ = requests.get('https://www.javbus.com/' + self.fanhao,  headers=headers).content


            #requ = requests.get('https://www.javbus.com/' + self.fanhao,timeout=4,headers=headers).content
            requ2 = requests.get('http://www.j15av.com/cn/vl_searchbyid.php?list&keyword='+self.fanhao, headers=headers).content
        self.lblStates.SetLabel('请求成功，正在处理内容。。。')
        # 获取标签内容
        soup = BeautifulSoup(requ, 'html.parser', from_encoding='utf-8')
        soup2 = BeautifulSoup(requ2, 'html.parser', from_encoding='utf-8')
        # 先判断是否能正常搜索到内容
        is404 = soup.find('h4', {'style': 'font-size:36px;'})
        if is404 is not None:
            wx.MessageBox("该番号查不到", "Message", wx.OK | wx.ICON_INFORMATION)
            return
        cover_url = soup.find('a', class_='bigImage').get('href')

        #cover_url = soup2.find('img', id='video_jacket_img').get('src')
        print(cover_url)
        # 创建文件夹
        if os.path.exists('资源/' + self.fanhao):
            print('文件夹已经存在')
            self.lblStates.SetLabel('文件夹已经存在')
        else:
            os.makedirs('资源/' + self.fanhao)
            print('创建文件夹成功')
            self.lblStates.SetLabel('创建文件夹成功')

        # 获取图片
        if os.path.exists('资源/' + self.fanhao+'/cover.jpg'):
            string = '资源/' + self.fanhao + '/cover.jpg'
            image = wx.Image(string, wx.BITMAP_TYPE_JPEG)
            self.bmp = wx.StaticBitmap(self, bitmap=image.ConvertToBitmap(), size=(700, 523), pos=(25, 88))
        else:
            try:
                self.lblStates.SetLabel('正在获取图片')
                pic = requests.get(cover_url, headers=headers)
            except:
                print('【错误】当前图片无法下载')
                self.lblStates.SetLabel('请重试')
                wx.MessageBox("获取图片超时", "Message", wx.OK | wx.ICON_INFORMATION)
                return
            self.lblStates.SetLabel('获取信息成功')
            string = '资源/' + self.fanhao + '/cover.jpg'
            fp = open(string, 'wb')
            # 创建文件
            fp.write(pic.content)
            fp.close()
            image = wx.Image(string, wx.BITMAP_TYPE_JPEG)
            self.bmp = wx.StaticBitmap(self, bitmap=image.ConvertToBitmap(), size=(700, 523),pos=(25,88))

        # 获取名字
        soup_name = soup.find('h3')
        soup_name = soup_name.string
        self.lblName.SetLabel('名称：' + soup_name)


        # 获取导演
        soup_director = soup.find('a', href=re.compile(r'director'))
        if soup_director is None:
            self.av_director = ''
            self.pid_director = ''
        else:
            self.av_director = soup_director.string
            self.pid_director = soup_director.get('href').split('/')[-1]
        self.lblDirector.SetLabel('导演：'+ self.av_director)


        # 获取制作商
        soup_zhizuo = soup.find('a', href=re.compile(
            r'https://www.javbus.com/studio|https://www.javbus.com/uncensored/studio'))
        if soup_zhizuo is None:
            self.av_zhizuo = ''
            self.pid_zhizuo = ''
        else:
            self.av_zhizuo = soup_zhizuo.string
            self.pid_zhizuo = soup_zhizuo.get('href').split('/')[-1]
        self.lblZhizuo.SetLabel('制作商: ' + self.av_zhizuo)

        # 获取发行商
        soup_faxing = soup.find('a', href=re.compile(
            r'https://www.javbus.com/label|https://www.javbus.com/uncensored/label'))
        if soup_faxing is None:
            self.av_faxing = ''
            self.pid_faxing = ''
        else:
            self.av_faxing = soup_faxing.string
            self.pid_faxing = soup_faxing.get('href').split('/')[-1]
        self.lblFaxing.SetLabel('发行商： ' + self.av_faxing)

        # 获取系列
        soup_xilie = soup.find('a', href=re.compile(
            r'https://www.javbus.com/series|https://www.javbus.com/uncensored/series'))
        if soup_xilie is None:
            self.av_xilie = ''
            self.pid_xilie = ''
        else:
            self.av_xilie = soup_xilie.string
            self.pid_xilie = soup_xilie.get('href').split('/')[-1]
        self.lblXilie.SetLabel('系列：' + self.av_xilie)

        # 获取类别
        soup_genres = soup.find_all('a', href=re.compile(
            r'https://www.javbus.com/genre/.|https://www.javbus.com/uncensored/genre.'))
        del soup_genres[-1];
        if soup_genres is None:
            self.av_genres = ['无类别']
            self.pid_genres = ['']
        else:
            for soup_genre in soup_genres[2:]:
                self.av_genres.append(soup_genre.string)
                pid_genre = soup_genre.get('href').split('/')[-1]
                self.pid_genres.append(pid_genre)

        for av_genre in self.av_genres:
            av_genre = av_genre + ','
            self.new_av_genres = av_genre + self.new_av_genres
        print(self.new_av_genres)
        self.lblGenres.SetLabel('类别:'+self.new_av_genres)


        # 获取演员
        soup_stars = soup.find_all('span', {'onmouseover': re.compile(r'hoverdiv')})
        print(soup_stars)
        if soup_stars ==[]:
            self.OnClickSearchStar(event)
        else:
            for soup_star in soup_stars:
                soup_star_txt = soup_star.get_text()
                soup_star_txt = soup_star_txt.replace('\n', '')
                pid_star = soup_star.a.get('href').split('/')[-1]
                self.av_stars.append(soup_star_txt)
                self.pid_stars.append(pid_star)

        for av_star in self.av_stars:
            av_star = av_star + ','
            self.new_av_stars = av_star + self.new_av_stars
        self.lblStars.SetLabel('演员：'+self.new_av_stars)

        #获取编辑得分
        self.OnClickSearchScore(event)

        # 获取评论
        self.OnClickSearchComment(event)

    def OnClickStore(self,event):

        if self.fanhao =='':
            wx.MessageBox("请勿乱点", "Message", wx.OK | wx.ICON_INFORMATION)
            return

        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 把所有数据插入av_record表
        sql_sel = "SELECT COUNT(*) FROM av_record WHERE fanhao = '" + self.fanhao + "'"
        cur.execute(sql_sel)
        num_fanhao = cur.fetchall()[0][0]
        if num_fanhao == 1:
            wx.MessageBox("该番号已存在", "Message", wx.OK | wx.ICON_INFORMATION)
            print('该番号已存在！！')
            return
        else:
            sql_insert = "INSERT INTO av_record(fanhao,director,faxing,zhizuo,xilie,genre,stars,last,score,other_score,comment,type)  VALUES ('" + self.fanhao + "','" + self.av_director + "','" + self.av_faxing + "','" + self.av_zhizuo + "','" + self.av_xilie + "','" \
                         + self.new_av_genres + "','" + self.new_av_stars + "','0','1','"+self.other_score+"','"+self.comment+"','"+str(self.type)+"')"
            cur.execute(sql_insert)
            conn.commit()
            print('增加新番号成功！！')

        # 把所有数据插入av_director表
        if self.av_director is not None:
            sql_sel = "SELECT COUNT(director) FROM av_director WHERE director = '" + self.av_director + "'"
            cur.execute(sql_sel)
            num_director = cur.fetchall()[0][0]
            if num_director == 1:
                #sql_insert = "UPDATE av_director SET num = num+1 WHERE director ='" + self.av_director + "'"
                print('导演已存在！！')
            else:
                sql_insert = "INSERT INTO av_director(director,num,pid) VALUES ('" + self.av_director + "','1','" + self.pid_director + "')"
                print('新导演添加！！')
                cur.execute(sql_insert)
                conn.commit()

        # 把所有数据插入av_faxing表
        if self.av_faxing is not None:
            sql_sel = "SELECT COUNT(faxing) FROM av_faxing WHERE faxing = '" + self.av_faxing + "'"
            cur.execute(sql_sel)
            num_faxing = cur.fetchall()[0][0]
            if num_faxing == 1:
                #sql_insert = "UPDATE av_faxing SET num = num+1 WHERE faxing ='" + self.av_faxing + "'"
                print('发行商已存在！！！')
            else:
                sql_insert = "INSERT INTO av_faxing(faxing,num,pid) VALUES ('" + self.av_faxing + "','1','" + self.pid_faxing + "')"
                print('新发行商添加！！！')
                cur.execute(sql_insert)
                conn.commit()

        # 把所有数据插入av_zhizuo表
        if self.av_zhizuo is not None:
            sql_sel = "SELECT COUNT(zhizuo) FROM av_zhizuo WHERE zhizuo = '" + self.av_zhizuo + "'"
            cur.execute(sql_sel)
            num_zhizuo = cur.fetchall()[0][0]
            if num_zhizuo == 1:
                #sql_insert = "UPDATE av_zhizuo SET num = num+1 WHERE zhizuo ='" + self.av_zhizuo + "'"
                print('制作商已存在！！！')
            else:
                sql_insert = "INSERT INTO av_zhizuo(zhizuo,num,pid) VALUES ('" + self.av_zhizuo + "','1','" + self.pid_zhizuo + "')"
                print('新制作商添加！！！')
                cur.execute(sql_insert)
                conn.commit()

        # 把所有数据插入av_xilie表
        if self.av_xilie is not None:
            sql_sel = "SELECT COUNT(xilie) FROM av_xilie WHERE xilie = '" + self.av_xilie + "'"
            cur.execute(sql_sel)
            num_zhizuo = cur.fetchall()[0][0]
            if num_zhizuo == 1:
                #sql_insert = "UPDATE av_xilie SET num = num+1 WHERE xilie ='" + self.av_xilie + "'"
                print('该系列已存在！！！')
            else:
                sql_insert = "INSERT INTO av_xilie(xilie,num,pid) VALUES ('" + self.av_xilie + "','1','" + self.pid_xilie + "')"
                print('新系列添加！！！')
                cur.execute(sql_insert)
                conn.commit()

        # 把所有数据插入av_genre表
        for i in range(len(self.av_genres)):
            sql_sel = "SELECT COUNT(genre) FROM av_genres WHERE genre = '" + self.av_genres[i] + "'"
            cur.execute(sql_sel)
            num_genre = cur.fetchall()[0][0]
            if num_genre == 1:
                #sql_insert = "UPDATE av_genres SET num = num+1 WHERE genre ='" + self.av_genres[i] + "'"
                print('该类别已存在')
            else:
                sql_insert = "INSERT INTO av_genres(genre,num,pid) VALUES ('" + self.av_genres[i] + "','1','" + \
                             self.pid_genres[i] + "')"
                print(sql_insert)
                cur.execute(sql_insert)
                conn.commit()

        # 把所有数据插入av_stars表
        for j in range(len(self.av_stars)):
            sql_sel = "SELECT COUNT(star) FROM av_stars WHERE star = '" + self.av_stars[j] + "'"
            cur.execute(sql_sel)
            num_star = cur.fetchall()[0][0]
            if num_star == 1:
                #sql_insert = "UPDATE av_stars SET num = num+1 WHERE star ='" + self.av_stars[j] + "'"
                print('该演员已存在！！！')
            else:
                sql_insert = "INSERT INTO av_stars(star,num,pid) VALUES ('" + self.av_stars[j] + "','1','" + self.pid_stars[
                    j] + "')"

                print('新演员添加！！！')
                cur.execute(sql_insert)
                conn.commit()

        cur.close()
        conn.close()

        #转移文件
        self.fanhaoPlus = self.lblDrag.GetLabel()
        filename_old = os.path.basename(self.fanhaoPlus)
        filename_new = self.fanhao + '.'+filename_old.split('.')[-1]
        try:
            print(self.fanhaoPlus)
            shutil.move(self.fanhaoPlus, 'G:/fun_finder/资源/' + self.fanhao + '/' + filename_old)
            os.rename('G:/fun_finder/资源/'+self.fanhao+'/'+filename_old, 'G:/fun_finder/资源/'+self.fanhao+'/'+filename_new)
        except:
            print('文件不存在')
        wx.MessageBox("入库成功", "Message", wx.OK | wx.ICON_INFORMATION)

    def OnClickSearchStar(self,event):
        self.fanhao = self.txt_fanhao.GetValue()
        self.lblStates.SetLabel('正在搜索番号：' + self.fanhao)
        self.av_stars = []
        self.new_av_stars = ''
        self.pid_stars = []
        self.lblStars.SetLabel('演员：')
        if self.fanhao is None:
            return
        else:
            headers = {'content-type': 'application/json',
                       'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
            s = requests.session()
            s.keep_alive = False
            try:
                requ = requests.get('http://www.j18ib.com/cn/vl_searchbyid.php?list&keyword='+self.fanhao, headers=headers,timeout=2).content
            except:
                wx.MessageBox("获取网址超时", "Message", wx.OK | wx.ICON_INFORMATION)
                return

        self.lblStates.SetLabel('请求成功，正在处理内容。。。')
        # 获取标签内容
        soup = BeautifulSoup(requ, 'html.parser', from_encoding='utf-8')
        # 先判断是否能正常搜索到内容
        is404 = soup.find('em', {'style': 'text-align: center; color: #C0C0C0;'})
        if is404 is not None:
            wx.MessageBox("该番号查不到", "Message", wx.OK | wx.ICON_INFORMATION)
            return

        print(soup)
        soup_stars = soup.find_all('span', class_='star')
        print('彦彦彪')
        print(soup_stars)
        if soup_stars ==[]:
            wx.MessageBox("没有演员更新", "Message", wx.OK | wx.ICON_INFORMATION)
            return
        else:
            for soup_star in soup_stars:
                soup_star_txt = soup_star.get_text()
                print(soup_star_txt)
                self.av_stars.append(soup_star_txt)
                print(self.av_stars)

        for av_star in self.av_stars:
            av_star = av_star + ','
            self.new_av_stars = av_star + self.new_av_stars
        self.lblStars.SetLabel('演员：'+self.new_av_stars)

        for av_star in self.av_stars:
            requ = requests.get('https://www.javbus.com/searchstar/' + av_star, headers=headers, timeout=3).content
            soup = BeautifulSoup(requ, 'html.parser', from_encoding='utf-8')
            pid_star = soup.find('a', class_='avatar-box text-center')
            if pid_star is None:
                pid_star=''
            else:
                pid_star = pid_star.get('href').split('/')[-1]
            self.pid_stars.append(pid_star)
            self.lblStates.SetLabel('更新演员pid：'+pid_star)
            print(self.pid_stars)

        wx.MessageBox("演员信息已更新", "Message", wx.OK | wx.ICON_INFORMATION)

    def OnClickUpdateStar(self,event):
        if self.fanhao =='':
            wx.MessageBox("请勿乱点", "Message", wx.OK | wx.ICON_INFORMATION)
            return

        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 把所有数据插入av_record表
        sql_sel = "SELECT COUNT(*) FROM av_record WHERE fanhao = '" + self.fanhao + "'"
        cur.execute(sql_sel)
        num_fanhao = cur.fetchall()[0][0]
        if num_fanhao == 1:
            print('该番号存在，可以更新数据')
            sql_insert = "UPDATE av_record SET stars='" + self.new_av_stars + "' WHERE fanhao = '" + self.fanhao + "'"
            cur.execute(sql_insert)
            conn.commit()
            print('演员数据更新成功！！')

        # 把所有数据插入av_stars表
        for j in range(len(self.av_stars)):
            sql_sel = "SELECT COUNT(star) FROM av_stars WHERE star = '" + self.av_stars[j] + "'"
            cur.execute(sql_sel)
            num_star = cur.fetchall()[0][0]
            if num_star == 1:
                #sql_insert = "UPDATE av_stars SET num = num+1 WHERE star ='" + self.av_stars[j] + "'"
                print('该演员已存在！！！')
            else:
                sql_insert = "INSERT INTO av_stars(star,num,pid) VALUES ('" + self.av_stars[j] + "','1','" + self.pid_stars[j] + "')"
                print('新演员添加！！！')
            cur.execute(sql_insert)
            conn.commit()

        cur.close()
        conn.close()
        wx.MessageBox("演员数据已入库", "Message", wx.OK | wx.ICON_INFORMATION)
        self.lblStates.SetLabel('入库成功')

    def OnClickSearchScore(self,event):
        self.fanhao = self.txt_fanhao.GetValue()
        self.lblStates.SetLabel('正在查编辑得分:'+self.fanhao)
        headers = {'content-type': 'application/json',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}

        try:
            s = requests.session()
            s.keep_alive = False
            requ_other = requests.get('http://www.ja14b.com/cn/vl_searchbyid.php?keyword=' + self.fanhao,
                                      headers=headers).content
        except:
            wx.MessageBox("获取网址超时", "Message", wx.OK | wx.ICON_INFORMATION)
            return
        self.lblStates.SetLabel('请求编辑得分成功，正在获取内容')
        # 获取标签内容
        soup_other = BeautifulSoup(requ_other, 'html.parser', from_encoding='utf-8')
        # 先判断是否能正常搜索到内容
        is404 = soup_other.find('p', {'style': 'text-align: center; color: #C0C0C0;'})
        if is404 is not None:
            wx.MessageBox("该番号查不到", "Message", wx.OK | wx.ICON_INFORMATION)
            return

        isNone = soup_other.find('div',id='badalert')
        if isNone is not None:
            wx.MessageBox("该番号查不到", "Message", wx.OK | wx.ICON_INFORMATION)
            return

        isDouble = soup_other.find('div', class_=re.compile(r'videos'))
        if isDouble is not None:
            print(isDouble)
            url = isDouble.a.get('href')
            url = 'http://www.ja14b.com/cn'+url[1:]
            print(url)
            requ_double=requests.get(url,headers=headers, timeout=2).content
            soup_double = BeautifulSoup(requ_double, 'html.parser', from_encoding='utf-8')
            other_score = soup_double.find('span', class_='score')
            self.other_score = other_score.string[1:-1]
            self.lblOtherscore.SetLabel('编辑得分:' + self.other_score)
            print(other_score)
            self.lblStates.SetLabel('已经获取编辑得分')
            #wx.MessageBox("已经获取编辑得分", "Message", wx.OK | wx.ICON_INFORMATION)
        else:
            other_score = soup_other.find('span', class_='score')
            print(other_score)
            if other_score.string is None:
                self.other_score = '0'
            else:
                self.other_score = other_score.string[1:-1]
            self.lblOtherscore.SetLabel('编辑得分:'+self.other_score)
            print(other_score)
            self.lblStates.SetLabel('已经获取编辑得分')
            #wx.MessageBox("已经获取编辑得分", "Message", wx.OK | wx.ICON_INFORMATION)

    def OnClickUpdateScore(self,event):
        if self.fanhao =='':
            wx.MessageBox("请勿乱点", "Message", wx.OK | wx.ICON_INFORMATION)
            return

        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 把所有数据插入av_record表
        sql_sel = "SELECT COUNT(*) FROM av_record WHERE fanhao = '" + self.fanhao + "'"
        cur.execute(sql_sel)
        num_fanhao = cur.fetchall()[0][0]
        if num_fanhao == 1:
            print('该番号存在，可以更新数据')
            sql_insert = "UPDATE av_record SET other_score='" + self.other_score + "' WHERE fanhao = '" + self.fanhao + "'"
            cur.execute(sql_insert)
            conn.commit()
            print('编辑得分数据更新成功！！')

        cur.close()
        conn.close()
        wx.MessageBox("编辑得分数据已入库", "Message", wx.OK | wx.ICON_INFORMATION)
        self.lblStates.SetLabel('入库成功')

    def OnClickSearchComment(self,event):
        self.fanhao = self.txt_fanhao.GetValue()
        self.lblStates.SetLabel('正在查评论:' + self.fanhao)
        headers = {'content-type': 'application/json',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}

        try:
            s = requests.session()
            s.keep_alive = False
            requ_comment = requests.get('http://www.ja14b.com/cn/vl_searchbyid.php?keyword=' + self.fanhao,
                                      headers=headers, timeout=2).content
        except:
            wx.MessageBox("获取网址超时", "Message", wx.OK | wx.ICON_INFORMATION)
            return
        self.lblStates.SetLabel('请求评论成功，正在获取内容')
        # 获取标签内容
        soup_comment = BeautifulSoup(requ_comment, 'html.parser', from_encoding='utf-8')
        # 先判断是否能正常搜索到内容
        is404 = soup_comment.find('p', {'style': 'text-align: center; color: #C0C0C0;'})
        if is404 is not None:
            wx.MessageBox("该番号查不到", "Message", wx.OK | wx.ICON_INFORMATION)
            return

        isNone = soup_comment.find('div', id='badalert')
        if isNone is not None:
            wx.MessageBox("该番号查不到", "Message", wx.OK | wx.ICON_INFORMATION)
            return

        self.comment = ''
        isDouble = soup_comment.find('div', class_=re.compile(r'videos'))
        if isDouble is not None:
            print(isDouble)
            url = isDouble.a.get('href')
            url = 'http://www.ja14b.com/cn' + url[1:]
            print(url)
            requ_double = requests.get(url, headers=headers, timeout=2).content
            soup_double = BeautifulSoup(requ_double, 'html.parser', from_encoding='utf-8')
            comments = soup_double.find_all('textarea', class_='hidden')
            for comment_new in comments:
                comment_new = comment_new.string
                if self.isComment(comment_new) == True:
                    print(comment_new)
                    self.comment = self.comment + comment_new + '\n'
            self.lblComment.SetLabel('评论:' + self.comment)
            print(self.lblComment)
            self.lblStates.SetLabel('已经获取评论')
            wx.MessageBox("获取评论成功", "Message", wx.OK | wx.ICON_INFORMATION)
        else:
            comments = soup_comment.find_all('textarea', class_='hidden')
            for comment_new in comments:
                comment_new = comment_new.string
                if self.isComment(comment_new) == True:
                    print(comment_new)
                    self.comment = self.comment + comment_new + '\n'
            self.lblComment.SetLabel('评论:' + self.comment)
            print(self.lblComment)
            self.lblStates.SetLabel('已经获取评论')
            wx.MessageBox("获取评论成功", "Message", wx.OK | wx.ICON_INFORMATION)

    def isComment(self,comment_new):

        pattern = re.compile(r'magnet|url|ed2k|thunder|http')
        search = pattern.search(comment_new)
        if search:
            print('去除无效评论')
            return False
        return True

    def OnClickUpdateComment(self,event):
        if self.fanhao =='':
            wx.MessageBox("请勿乱点", "Message", wx.OK | wx.ICON_INFORMATION)
            return

        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 把所有数据插入av_record表
        sql_sel = "SELECT COUNT(*) FROM av_record WHERE fanhao = '" + self.fanhao + "'"
        cur.execute(sql_sel)
        num_fanhao = cur.fetchall()[0][0]
        if num_fanhao == 1:
            print('该番号存在，可以更新数据')
            sql_insert = "UPDATE av_record SET comment='" + self.comment + "' WHERE fanhao = '" + self.fanhao + "'"
            cur.execute(sql_insert)
            conn.commit()
            print('评论更新成功！！')

        cur.close()
        conn.close()
        wx.MessageBox("评论数据已入库", "Message", wx.OK | wx.ICON_INFORMATION)
        self.lblStates.SetLabel('入库成功')

    def onRadioBox(self, e):
        if self.rbox.GetStringSelection() == '日本有码':
            self.type = 0
        elif self.rbox.GetStringSelection() == '日本无码':
            self.type = 1