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

class PanelWestIn(wx.Panel):
    def __init__(self, parent):
        """Constructor"""
        wx.Panel.__init__(self, parent=parent,size=(1140, 550))

        conBox = wx.BoxSizer(wx.VERTICAL)
        hbox1 =wx.BoxSizer(wx.HORIZONTAL)
        l1 = wx.StaticText(self, label="名称：")
        self.txt_fanhao = wx.TextCtrl(self)
        self.btn_search = wx.Button(self,label="搜索")
        self.btn_search.Bind(wx.EVT_BUTTON, self.OnClickSearch)
        self.lblStates = wx.StaticText(self, label="运行状态：", size=(100,20))

        hbox1.Add(l1,0,wx.ALL,13)
        hbox1.Add(self.txt_fanhao, 0, wx.ALL,10)
        hbox1.Add(self.btn_search, 0, wx.ALL, 7)
        hbox1.Add(self.lblStates, 0, wx.ALL, 13)

        hbox2 = wx.BoxSizer(wx.HORIZONTAL)

        self.lblDrag = wx.StaticText(self, label="", style=wx.ALIGN_CENTRE | wx.ST_NO_AUTORESIZE,
                                     size=(740, 523))
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
        self.lblStudio = wx.StaticText(self, label="制作：", style=wx.ALIGN_CENTRE)
        self.lblSeries = wx.StaticText(self, label="系列：", style=wx.ALIGN_CENTRE)
        self.lblGenres = wx.StaticText(self, label="类型：", style=wx.ALIGN_CENTRE)
        self.lblStars = wx.StaticText(self, label="演员：", style=wx.ALIGN_CENTRE)
        self.btn_inStroe = wx.Button(self, label="入库", size=(382, 65))
        self.btn_inStroe.Bind(wx.EVT_BUTTON, self.OnClickStore)
        hbox2_right.Add(self.lblName,0,wx.BOTTOM|wx.TOP,10)
        hbox2_right.Add(self.lblDirector, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblStudio, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblSeries, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblGenres, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.lblStars, 0, wx.BOTTOM, 10)
        hbox2_right.Add(self.btn_inStroe, 0, wx.BOTTOM, 10)
        btnBox = wx.BoxSizer(wx.HORIZONTAL)
        hbox2_right.Add(btnBox, 0, wx.ALL, 0)
        hbox2.Add(hbox2_right, 0, wx.ALL, 10)

        conBox.Add(hbox1,0,wx.ALL,0)
        conBox.Add(hbox2, 0, wx.ALL, 0)
        self.SetSizer(conBox)



        self.name = ''
        self.fanhaoPlus=''
        self.av_director=''
        self.pid_director=''
        self.av_studio=''
        self.pid_studio=''
        self.av_series=''
        self.pid_series=''
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
        self.name = ''
        self.fanhaoPlus = ''
        self.av_director = ''
        self.pid_director = ''
        self.av_studio = ''
        self.pid_studio = ''
        self.av_series = ''
        self.pid_series = ''
        self.av_genres = []
        self.new_av_genres = ''
        self.pid_genres = []
        self.av_stars = []
        self.new_av_stars = ''
        self.pid_stars = []
        self.other_socre=''
        self.comment=''
        self.name=self.txt_fanhao.GetValue()
        self.lblStates.SetLabel('正在搜索：'+self.name)
        if self.name is None:
            return
        else:
            headers = {'content-type': 'application/json',
                       'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}

            try:
                name_url = self.name.replace(' ','+')
                print(name_url)
                requ = requests.get('http://www.data18.com/search/?k=' + name_url, timeout=2,headers=headers).content

            except:
                wx.MessageBox("获取网址超时", "Message", wx.OK | wx.ICON_INFORMATION)
                return

        self.lblStates.SetLabel('请求成功，正在处理内容。。。')
        # 获取标签内容
        soup = BeautifulSoup(requ, 'html.parser', from_encoding='utf-8')
        # 先判断是否能正常搜索到内容
        is404 = soup.find(text='0 Results')
        if is404 is not None:
            wx.MessageBox("该电影查不到", "Message", wx.OK | wx.ICON_INFORMATION)
            return

        new_url_soup = soup.find('a', href=re.compile('http://www.data18.com/movies/\d'))
        #通过movie标签查找
        if new_url_soup is not None:
            try:
                new_url = new_url_soup.get('href')
                requ_movie = requests.get(new_url, timeout=2, headers=headers).content
            except:
                wx.MessageBox("获取新网址超时", "Message", wx.OK | wx.ICON_INFORMATION)
                return

            soup_movie = BeautifulSoup(requ_movie, 'html.parser', from_encoding='utf-8')
            # 获取名字
            soup_name = soup_movie.find('h1')
            soup_name = soup_name.string
            self.lblName.SetLabel('名称：' + soup_name)
            self.name = soup_name

            # 创建文件夹
            if os.path.exists('资源/' + self.name):
                print('文件夹已经存在')
                self.lblStates.SetLabel('文件夹已经存在')
            else:
                os.makedirs('资源/' + self.name)
                print('创建文件夹成功')
                self.lblStates.SetLabel('创建文件夹成功')

            # 获取图片
            cover_url = soup_movie.find_all('a', class_='grouped_elements')
            cover_front_url = cover_url[0].get('href')
            cover_back_url = cover_url[1].get('href')
            print(cover_front_url,cover_back_url)
            try:
                self.lblStates.SetLabel('正在获取图片')
                pic_front = requests.get(cover_front_url, timeout=3)
                pic_back = requests.get(cover_back_url, timeout=3)
            except:
                print('【错误】当前图片无法下载')
                self.lblStates.SetLabel('请重试')
                wx.MessageBox("获取图片超时", "Message", wx.OK | wx.ICON_INFORMATION)
                return
            self.lblStates.SetLabel('获取信息成功')
            string_front = '资源/' + self.name + '/cover_front.jpg'
            fp = open(string_front, 'wb')
            # 创建文件
            fp.write(pic_front.content)
            fp.close()

            string_back = '资源/' + self.name + '/cover_back.jpg'
            fp = open(string_back, 'wb')
            # 创建文件
            fp.write(pic_back.content)
            fp.close()

            image_front = wx.Image(string_front, wx.BITMAP_TYPE_JPEG)
            self.bmp = wx.StaticBitmap(self, bitmap=image_front.ConvertToBitmap(), pos=(394, 88), size=(370, 523))
            image_back = wx.Image(string_back, wx.BITMAP_TYPE_JPEG)
            self.bmp = wx.StaticBitmap(self, bitmap=image_back.ConvertToBitmap(), pos=(24, 88), size=(370, 523))

            # 获取制作商
            soup_studio = soup_movie.find('u').string
            print(soup_studio)
            if soup_studio is None:
                self.av_studio = ''
                self.pid_studio = ''
            else:
                self.av_studio = soup_studio.string
                self.pid_studio = 'http://www.data18.com/sites/' + self.av_studio
            self.lblStudio.SetLabel('制作商: ' + self.av_studio)

            # 获取导演
            soup_director_raw = soup_movie.find('b', text='Director:')
            soup_director = soup_director_raw.find_next('a').string
            if soup_director is None:
                self.av_director = ''
                self.pid_director = ''
            else:
                self.av_director = soup_director.string
                self.pid_director = soup_director_raw.find_next('a').get('href')
            self.lblDirector.SetLabel('导演：' + self.av_director)

            # 获取系列
            soup_series_raw = soup_movie.find('b',text='Serie:')
            soup_series = soup_series_raw.find_next('a').string
            print(soup_series)
            if soup_series is None:
                self.av_series = ''
                self.pid_series = ''
            else:
                self.av_series = soup_series.string
                self.pid_series = soup_series_raw.find_next('a').get('href')
            self.lblSeries.SetLabel('系列：' + self.av_series)


            # 获取类别
            soup_genres = soup_movie.find('b',text='Categories:').find_next_siblings('a')
            print(soup_genres)
            if soup_genres is None:
                self.av_genres = ['无类别']
                self.pid_genres = ['']
            else:
                for soup_genre in soup_genres:
                    self.av_genres.append(soup_genre.string)
                    pid_genre = soup_genre.get('href')
                    self.pid_genres.append(pid_genre)

            for av_genre in self.av_genres:
                av_genre = av_genre + ','
                self.new_av_genres = av_genre + self.new_av_genres
            print(self.new_av_genres)
            print(self.pid_genres)
            self.lblGenres.SetLabel('类别:' + self.new_av_genres)

            # 获取演员
            soup_stars = soup_movie.find_all('p',class_='line1',style='align: center;')
            print(soup_stars)
            for soup_star in soup_stars:
                soup_star_txt = soup_star.a.string
                pid_star = soup_star.a.get('href')
                self.av_stars.append(soup_star_txt)
                self.pid_stars.append(pid_star)

            print(self.av_stars)
            for av_star in self.av_stars:
                av_star = av_star + ','
                self.new_av_stars = av_star + self.new_av_stars
            self.lblStars.SetLabel('演员：' + self.new_av_stars)


        #通过content标签查找
        else:
            new_url = soup.find('p', class_='line1').a.get('href')
            try:
                requ_movie = requests.get(new_url, timeout=2, headers=headers).content
            except:
                wx.MessageBox("获取新网址超时", "Message", wx.OK | wx.ICON_INFORMATION)
                return

            soup_movie = BeautifulSoup(requ_movie, 'html.parser', from_encoding='utf-8')

            # 获取名字
            soup_name = soup_movie.find('h1')
            soup_name = soup_name.string
            self.lblName.SetLabel('名称：' + soup_name)
            self.name = soup_name

            cover_url = soup_movie.find('div', id='moviewrap').img.get('src')
            print(cover_url)

            # 创建文件夹
            if os.path.exists('资源/' + self.name):
                print('文件夹已经存在')
                self.lblStates.SetLabel('文件夹已经存在')
            else:
                os.makedirs('资源/' + self.name)
                print('创建文件夹成功')
                self.lblStates.SetLabel('创建文件夹成功')

            # 获取图片
            try:
                self.lblStates.SetLabel('正在获取图片')
                pic = requests.get(cover_url, timeout=3)
            except:
                print('【错误】当前图片无法下载')
                self.lblStates.SetLabel('请重试')
                wx.MessageBox("获取图片超时", "Message", wx.OK | wx.ICON_INFORMATION)
                return
            self.lblStates.SetLabel('获取信息成功')
            string = '资源/' + self.name + '/cover_front.jpg'
            fp = open(string, 'wb')
            # 创建文件
            fp.write(pic.content)
            fp.close()
            image = wx.Image(string, wx.BITMAP_TYPE_JPEG)
            self.bmp = wx.StaticBitmap(self, bitmap=image.ConvertToBitmap(), pos=(75, 170))


            # 获取制作商
            soup_studio = soup_movie.find('u').string
            print(soup_studio)
            if soup_studio is None:
                self.av_studio = ''
                self.pid_studio = ''
            else:
                self.av_studio = soup_studio.string
                self.pid_studio = 'http://www.data18.com/sites/'+self.av_studio
            self.lblStudio.SetLabel('制作商: ' + self.av_studio)


            # 获取系列
            soup_series_raw = soup_movie.find(text='Site:')
            soup_series = soup_series_raw.find_next('a').string
            print(soup_series)
            if soup_series is None:
                self.av_series = ''
                self.pid_series = ''
            else:
                self.av_series = soup_series.string
                self.pid_series = soup_series_raw.find_next('a').get('href')
            self.lblSeries.SetLabel('系列：' + self.av_series)



            # 获取导演
            soup_director = soup.find('a', href=re.compile(r'director'))
            if soup_director is None:
                self.av_director = ''
                self.pid_director = ''
            else:
                self.av_director = soup_director.string
                self.pid_director = soup_director.get('href').split('/')[-1]
            self.lblDirector.SetLabel('导演：'+ self.av_director)


            # 获取类别
            soup_genres = soup_movie.find('div',style='margin-top: 3px;').find_all('a')
            print(soup_genres)
            if soup_genres is None:
                self.av_genres = ['无类别']
                self.pid_genres = ['']
            else:
                for soup_genre in soup_genres:
                    self.av_genres.append(soup_genre.string)
                    pid_genre = soup_genre.get('href')
                    self.pid_genres.append(pid_genre)

            for av_genre in self.av_genres:
                av_genre = av_genre + ','
                self.new_av_genres = av_genre + self.new_av_genres
            print(self.new_av_genres)
            print(self.pid_genres)
            self.lblGenres.SetLabel('类别:'+self.new_av_genres)


            # 获取演员
            soup_stars = soup_movie.find('b', text='Starring:').find_next_siblings('a')
            print(soup_stars)
            for soup_star in soup_stars:
                soup_star_txt = soup_star.get_text()
                pid_star = soup_star.get('href')
                self.av_stars.append(soup_star_txt)
                self.pid_stars.append(pid_star)

            for av_star in self.av_stars:
                av_star = av_star + ','
                self.new_av_stars = av_star + self.new_av_stars
            self.lblStars.SetLabel('演员：'+self.new_av_stars)


    def OnClickStore(self,event):

        if self.name =='':
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
        sql_sel = "SELECT COUNT(*) FROM west_record WHERE movie_name = '" + self.name + "'"
        cur.execute(sql_sel)
        num_fanhao = cur.fetchall()[0][0]
        if num_fanhao == 1:
            wx.MessageBox("该电影已存在", "Message", wx.OK | wx.ICON_INFORMATION)
            print('该电影已存在！！')
            return
        else:
            sql_insert = "INSERT INTO west_record (movie_name,director,studio,series,catagorys,stars,last,score) VALUES ('" + self.name + "','" + self.av_director + "','" + self.av_studio + "','" + self.av_series + "','" \
                         + self.new_av_genres + "','" + self.new_av_stars + "','0','1')"
            cur.execute(sql_insert)
            conn.commit()
            print('增加新番号成功！！')

        # 把所有数据插入av_director表
        if self.av_director is not None:
            sql_sel = "SELECT COUNT(director) FROM av_director WHERE director = '" + self.av_director + "'"
            cur.execute(sql_sel)
            num_director = cur.fetchall()[0][0]
            if num_director == 1:
                print('导演已存在！！')
            else:
                sql_insert = "INSERT INTO av_director(director,num,pid) VALUES ('" + self.av_director + "','1','" + self.pid_director + "')"
                print('新导演添加！！')
                cur.execute(sql_insert)
                conn.commit()

        # 把所有数据插入av_zhizuo表
        if self.av_studio is not None:
            sql_sel = "SELECT COUNT(zhizuo) FROM av_zhizuo WHERE zhizuo = '" + self.av_studio + "'"
            cur.execute(sql_sel)
            num_zhizuo = cur.fetchall()[0][0]
            if num_zhizuo == 1:
                print('制作商已存在！！！')
            else:
                sql_insert = "INSERT INTO av_zhizuo(zhizuo,num,pid) VALUES ('" + self.av_studio + "','1','" + self.pid_studio + "')"
                print('新制作商添加！！！')
                cur.execute(sql_insert)
                conn.commit()

        # 把所有数据插入av_xilie表
        if self.av_series is not None:
            sql_sel = "SELECT COUNT(xilie) FROM av_xilie WHERE xilie = '" + self.av_series + "'"
            cur.execute(sql_sel)
            num_zhizuo = cur.fetchall()[0][0]
            if num_zhizuo == 1:
                #sql_insert = "UPDATE av_xilie SET num = num+1 WHERE xilie ='" + self.av_xilie + "'"
                print('该系列已存在！！！')
            else:
                sql_insert = "INSERT INTO av_xilie(xilie,num,pid) VALUES ('" + self.av_series + "','1','" + self.pid_series + "')"
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
        filename_new = self.name + '.'+filename_old.split('.')[-1]
        try:
            print(self.fanhaoPlus)
            shutil.move(self.fanhaoPlus, 'G:/fun_finder/资源/' + self.name + '/' + filename_old)
            os.rename('G:/fun_finder/资源/'+self.name+'/'+filename_old, 'G:/fun_finder/资源/'+self.name+'/'+filename_new)
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

            try:
                requ = requests.get('http://www.ja14b.com/cn/vl_searchbyid.php?list&keyword='+self.fanhao, headers=headers,timeout=2).content
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

        soup_stars = soup.find_all('span', class_='star')
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
        sql_sel = "SELECT COUNT(*) FROM av_record WHERE ID = '" + self.fanhao + "'"
        cur.execute(sql_sel)
        num_fanhao = cur.fetchall()[0][0]
        if num_fanhao == 1:
            print('该番号存在，可以更新数据')
            sql_insert = "UPDATE av_record SET stars='" + self.new_av_stars + "' WHERE ID = '" + self.fanhao + "'"
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
            requ_other = requests.get('http://www.ja14b.com/cn/vl_searchbyid.php?keyword=' + self.fanhao,
                                      headers=headers, timeout=2).content
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
            if other_score is None:
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
        sql_sel = "SELECT COUNT(*) FROM av_record WHERE ID = '" + self.fanhao + "'"
        cur.execute(sql_sel)
        num_fanhao = cur.fetchall()[0][0]
        if num_fanhao == 1:
            print('该番号存在，可以更新数据')
            sql_insert = "UPDATE av_record SET other_score='" + self.other_score + "' WHERE ID = '" + self.fanhao + "'"
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
        sql_sel = "SELECT COUNT(*) FROM av_record WHERE ID = '" + self.fanhao + "'"
        cur.execute(sql_sel)
        num_fanhao = cur.fetchall()[0][0]
        if num_fanhao == 1:
            print('该番号存在，可以更新数据')
            sql_insert = "UPDATE av_record SET comment='" + self.comment + "' WHERE ID = '" + self.fanhao + "'"
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
