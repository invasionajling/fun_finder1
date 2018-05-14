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

class PanelSpider(wx.Panel):
    def __init__(self, parent):
        """Constructor"""
        wx.Panel.__init__(self, parent=parent,size=(1140, 550))
        self.hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.leftGrid = wx.BoxSizer(wx.VERTICAL)
        self.rightBox = wx.BoxSizer(wx.VERTICAL)

        self.btnTuijian = wx.Button(self, label='查看推荐下载')
        self.btnTuijian.Bind(wx.EVT_BUTTON, self.OnClickTuijian)
        self.btnUpdateXiazai = wx.Button(self, label='更新')
        self.btnUpdateXiazai.Bind(wx.EVT_BUTTON, self.OnClickUpdateXiazai)
        self.btnSpider1 = wx.Button(self, label='JAVLIB无码爬虫')
        self.btnSpider1.Bind(wx.EVT_BUTTON, self.OnClickSpider1)
        self.btnSpider2 = wx.Button(self, label='JAVBUS有码爬虫')
        self.btnSpider2.Bind(wx.EVT_BUTTON, self.OnClickSpider2)


        download = wx.StaticBox(self, -1, '下载:')
        downloadSizer = wx.StaticBoxSizer(download, wx.VERTICAL)
        downloadSizer.Add(self.btnTuijian, 0, wx.ALL | wx.EXPAND, 5)
        downloadSizer.Add(self.btnUpdateXiazai, 0, wx.ALL | wx.EXPAND, 5)

        spider = wx.StaticBox(self,-1,'爬虫')
        spiderSizer = wx.StaticBoxSizer(spider, wx.VERTICAL)
        spiderSizer.Add(self.btnSpider1, 0, wx.ALL | wx.EXPAND, 5)
        spiderSizer.Add(self.btnSpider2, 0, wx.ALL | wx.EXPAND, 5)

        self.leftGrid.Add(downloadSizer, 0, wx.EXPAND | wx.TOP, 10)
        self.rightBox.Add(spiderSizer, 0, wx.EXPAND | wx.TOP, 10)
        self.hbox.Add(self.leftGrid, 2, wx.ALL, 10)
        self.hbox.Add(self.rightBox, 1, wx.ALL, 10)
        self.SetSizer(self.hbox)

    # 打开推荐下载页面
    def OnClickTuijian(self, event):
        fanhaos = []

        '''
        i=1
        while len(fanhaos) <= 20:
            try:
                url_gaofen='http://www.ja14b.com/cn/vl_bestrated.php?&mode=2&page='+str(i)
                requ_gaofen = requests.get(url_gaofen,headers=headers, timeout=2).content
            except:
                wx.MessageBox("获取网址超时", "Message", wx.OK | wx.ICON_INFORMATION)
                return

            soup_gaofen = BeautifulSoup(requ_gaofen, 'html.parser', from_encoding='utf-8')
            fanhaos_gaofen_raw=soup_gaofen.find_all('div',class_='id')

            fanhaos_gaofen=[]
            for fanhao in fanhaos_gaofen_raw:
                fanhao = fanhao.string
                if self.isDownload(fanhao) == True:
                    fanhaos_gaofen.append(fanhao)
                    print(fanhaos_gaofen)

            try:
                url_xiangyao = 'http://www.ja14b.com/cn/vl_mostwanted.php?&mode=2&page='+str(i)
                requ_xiangyao = requests.get(url_xiangyao,headers=headers, timeout=2).content
            except:
                wx.MessageBox("获取网址超时", "Message", wx.OK | wx.ICON_INFORMATION)
                return

            soup_xiangyao = BeautifulSoup(requ_xiangyao, 'html.parser', from_encoding='utf-8')
            fanhaos_xiangyao_raw=soup_xiangyao.find_all('div',class_='id')

            fanhaos_xiangyao=[]
            for fanhao in fanhaos_xiangyao_raw:
                fanhao = fanhao.string
                if self.isDownload(fanhao) == True:
                    fanhaos_xiangyao.append(fanhao)
                    print(fanhaos_xiangyao)
            fanhaos_raw = fanhaos_gaofen + fanhaos_xiangyao
            fanhaos = list(set(fanhaos_raw))
            fanhaos.sort(key=fanhaos_raw.index)
            i = i +1

        '''

        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        fout = open("tuijian.html", "w")
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
        fout.write(
            "<ul class ='nav nav-tabs'><li  class='active'><a href='#youma' data-toggle='tab'>有码推荐</a></li><li><a href='#wuma' data-toggle='tab'>无码推荐</a></li></ul>")
        fout.write("<div class ='tab-content'><div class ='tab-pane active' id='youma'>")
        fout.write("<div class='row'>")
        # 获取TOP100有码番号
        sql_sel = "SELECT TOP 200 fanhao,href FROM av_recommend WHERE type=0 ORDER BY score DESC"
        cur.execute(sql_sel)

        fanhaos = cur.fetchall()
        print(fanhaos)
        for fanhao_raw in fanhaos:
            fanhao = fanhao_raw[0]
            href = fanhao_raw[1]
            if self.isDownload(fanhao) == True:
                fout.write("<a class='span4' href='https://www.javbus.com/" + fanhao + "' target='_blank'>")
                fout.write("<img src='" + href + "'>")
                fout.write("<p>" + fanhao + "</p>")
                fout.write("</a>")
        fout.write("</div>")
        fout.write("</div>")

        fout.write("<div class ='tab-pane' id='wuma'>")
        fout.write("<div class='row'>")
        # 获取TOP100无码番号
        sql_sel = "SELECT TOP 200 fanhao,href FROM av_recommend WHERE type=1 ORDER BY score DESC"
        cur.execute(sql_sel)
        fanhaos = cur.fetchall()
        for fanhao_raw in fanhaos:
            print(fanhao_raw)
            fanhao = fanhao_raw[0]
            href = fanhao_raw[1]
            if self.isDownload(fanhao) == True:
                fout.write("<a class='span4' href='https://www.javbus.com/" + fanhao + "' target='_blank'>")
                fout.write("<img src='" + href + "'>")
                fout.write("<p>" + fanhao + "</p>")
                fout.write("</a>")
        fout.write("</div>")
        fout.write("</div>")
        cur.close()
        conn.close()
        fout.write("</div>")
        fout.write("</div>")

        fout.write("</div>")
        fout.write("<script src='http://code.jquery.com/jquery.js'></script>")
        fout.write("<script src='bootstrap/js/bootstrap.min.js''></script>")
        fout.write("</body>")
        fout.write("</html>")

        url = 'G:/fun_finder/tuijian.html'
        webbrowser.open(url, new=0, autoraise=True)

    # 判断是否已经入库和已经下载
    def isDownload(self, fanhao):
        # 判断是否在库中
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 把所有数据插入av_record表
        sql_sel = "SELECT COUNT(*) FROM av_record WHERE fanhao = '" + fanhao + "'"
        cur.execute(sql_sel)
        num_fanhao = cur.fetchall()[0][0]
        cur.close()
        conn.close()
        if num_fanhao == 1:
            print(fanhao + ':该番号已存在！！')
            return False

        # 判断是否在下载队列中
        rootdir = 'G:/日本'
        for filenamess in os.walk(rootdir):
            filenames = filenamess[1] + filenamess[2]
            break
        filename = ''
        for i in range(len(filenames)):
            filename = filename + filenames[i]

        filename = filename.lower()
        fanhao = fanhao.lower()

        pattern = re.compile(fanhao)
        search = pattern.search(filename)
        if search:
            print('已在下载队列中')
            return False

        fanhao_without = fanhao.replace('-', '')
        pattern_without = re.compile(fanhao_without)
        search_without = pattern_without.search(filename)
        if search_without:
            print('已在下载队列中')
            return False

        return True

    # 判断是否在推荐表中
    def isInRecommand(self, fanhao):
        # 判断是否在库中
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 把所有数据插入av_record表
        sql_sel = "SELECT COUNT(*) FROM av_recommend WHERE fanhao = '" + fanhao + "'"
        cur.execute(sql_sel)
        num_fanhao = cur.fetchall()[0][0]
        cur.close()
        conn.close()
        if num_fanhao == 1:
            print(fanhao + ':该番号已存在！！')
            return False

        return True

        # 更新下载的得分

    def OnClickUpdateXiazai(self, event):
        # 通过所有番号
        def mdb_conn(password=""):
            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
            conn = pypyodbc.win_connect_mdb(str)
            return conn

        conn = mdb_conn()
        cur = conn.cursor()

        # 获取所有番号
        sql_sel = "SELECT fanhao FROM av_recommend"
        table_name = 'av_recommend'
        cur.execute(sql_sel)
        fanhaos = cur.fetchall()
        for i in range(len(fanhaos)):
            fanhao = fanhaos[i][0]
            print('-------------------正在执行:' + fanhao + '-----------------------')
            self.updateScore(fanhao, table_name)
        cur.close()
        conn.close()
        wx.MessageBox("数据更新完毕", "Message", wx.OK | wx.ICON_INFORMATION)

    # 更新某个番号的得分，根据传入的表名通过不同的规则来判断
    def updateScore(self, fanhao, table_name):
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
            sql_sel = "SELECT * FROM av_record WHERE ID = '" + fanhao + "'"
            cur.execute(sql_sel)
            av_info = cur.fetchall()
            print(av_info)
            # 获取导演数量
            if av_info[0][1] == '':
                num_director = 0
            else:
                av_director = av_info[0][1]
                sql_sel = "SELECT num FROM av_director WHERE director = '" + av_director + "'"
                cur.execute(sql_sel)
                raw_direcotr = cur.fetchall()
                num_director = raw_direcotr[0][0]
            print('导演数量：' + str(num_director))

            # 获取发行商数量
            if av_info[0][2] == '':
                num_faxing = 0
            else:
                av_faxing = av_info[0][2]
                sql_sel = "SELECT num FROM av_faxing WHERE faxing = '" + av_faxing + "'"
                cur.execute(sql_sel)
                raw_faxing = cur.fetchall()
                num_faxing = raw_faxing[0][0]
            print('发行商数量：' + str(num_faxing))

            # 获取制作商数量
            if av_info[0][3] == '':
                num_zhizuo = 0
            else:
                av_zhizuo = av_info[0][3]
                sql_sel = "SELECT num FROM av_zhizuo WHERE zhizuo = '" + av_zhizuo + "'"
                cur.execute(sql_sel)
                num_zhizuo = cur.fetchall()[0][0]
            print('制作商数量：' + str(num_zhizuo))

            # 获取系列数量
            if av_info[0][4] == '':
                num_xilie = 0
            else:
                av_xilie = av_info[0][4]
                sql_sel = "SELECT num FROM av_xilie WHERE xilie = '" + av_xilie + "'"
                cur.execute(sql_sel)
                num_xilie = cur.fetchall()[0][0]
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
                num_genre = cur.fetchall()[0][0]
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
                raw_star = cur.fetchall()
                num_star = raw_star[0][0]
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

            # 获取编辑得分
            if av_info[0][9] == '' or av_info[0][9] is None:
                num_otherscore = 0
            else:
                num_otherscore = float(av_info[0][9]) * 10
            print('编辑得分：' + str(num_zhizuo))
            score = num_director + num_faxing + num_zhizuo + num_xilie + num_stars + num_genres * 3 - (
                                                                                                          120 - num_days) * 10 + num_otherscore
            sql_update = "UPDATE av_record SET score ='" + str(score) + "' WHERE ID ='" + fanhao + "'"
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
                xilie_info = cur.fetchall()
                if len(xilie_info) == 0:
                    num_xilie = 0
                else:
                    num_xilie = xilie_info[0][0]
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
                    num_genre = genre_info[0][0]
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

            score = num_director + num_faxing + num_zhizuo + num_xilie + num_stars * 2 + num_genres * 3
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

            score = num_director + num_zhizuo + num_xilie + num_stars * 2 + num_genres * 3 - (120 - num_days) * 4
            sql_update = "UPDATE west_record SET score ='" + str(score) + "' WHERE ID =" + str(fanhao) + ""
            cur.execute(sql_update)
            conn.commit()
            print('该番号更新得分为' + str(score))
            cur.close()
            conn.close()

    # 启动爬虫1号
    def OnClickSpider1(self, evnet):
        t1 = threading.Thread(target=self.Spider1)
        t1.start()

        # 启动爬虫1号

    def OnClickSpider2(self, evnet):
        t1 = threading.Thread(target=self.Spider2)
        t1.start()

    # 爬虫1号有码脚本
    def Spider1(self):
        root_url = 'http://www.j18ib.com/cn/'
        headers = {'content-type': 'application/json',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}

        #处理首页最热门的内容
        # print("正在抓取首页最热门的内容")
        # index_root = requests.get(root_url, headers=headers).content
        # index_soup = BeautifulSoup(index_root, 'html.parser', from_encoding='utf-8')
        # index_fanhao = index_soup.find_all('div', class_='id')
        # for new_fanhao in index_fanhao:
        #     fanhao = new_fanhao.get_text()
        #     print('正在获取:' + fanhao + '的相关信息')
        #     if self.isDownload(fanhao) == True and self.isInRecommand(fanhao) == True:
        #         print(fanhao + '该番号可以获取信息')
        #         av_director = ''
        #         av_zhizuo = ''
        #         av_faxing = ''
        #         av_xilie = ''
        #         av_genres = []
        #         av_stars = []
        #         new_av_genres = ''
        #         new_av_stars = ''
        #         av_href = ''
        #
        #         requ = requests.get('https://www.javbus.com/' + fanhao, headers=headers).content
        #         print(requ + '该番号可以获取信息')
        #         soup = BeautifulSoup(requ, 'html.parser', from_encoding='utf-8')
        #         # 获取封面地址
        #         soup_href = soup.find('a', class_='bigImage')
        #         if soup_href is None:
        #             av_href = ''
        #         else:
        #             av_href = soup_href.get('href')
        #         print('获取封面地址成功！！')
        #         # 获取演员
        #         soup_stars = soup.find_all('span', {'onmouseover': re.compile(r'hoverdiv')})
        #         print(soup_stars)
        #         if soup_stars == []:
        #             av_stars = ''
        #         else:
        #             for soup_star in soup_stars:
        #                 soup_star_txt = soup_star.get_text()
        #                 soup_star_txt = soup_star_txt.replace('\n', '')
        #                 av_stars.append(soup_star_txt)
        #
        #         for av_star in av_stars:
        #             av_star = av_star + ','
        #             new_av_stars = av_star + new_av_stars
        #         print('获取演员成功！！')
        #         # 获取导演
        #         soup_director = soup.find('a', href=re.compile(r'director'))
        #         if soup_director is None:
        #             av_director = ''
        #         else:
        #             av_director = soup_director.string
        #         print('获取导演成功！！')
        #         # 获取制作商
        #         soup_zhizuo = soup.find('a', href=re.compile(
        #             r'https://www.javbus.com/studio|https://www.javbus.com/uncensored/studio'))
        #         if soup_zhizuo is None:
        #             av_zhizuo = ''
        #         else:
        #             av_zhizuo = soup_zhizuo.string
        #         print('获取制作商成功！！')
        #         # 获取发行商
        #         soup_faxing = soup.find('a', href=re.compile(
        #             r'https://www.javbus.com/label|https://www.javbus.com/uncensored/label'))
        #         if soup_faxing is None:
        #             av_faxing = ''
        #         else:
        #             av_faxing = soup_faxing.string
        #         print('获取发行商成功！！')
        #         # 获取系列
        #         soup_xilie = soup.find('a', href=re.compile(
        #             r'https://www.javbus.com/series|https://www.javbus.com/uncensored/series'))
        #         if soup_xilie is None:
        #             av_xilie = ''
        #         else:
        #             av_xilie = soup_xilie.string
        #         print('获取系列成功！！')
        #         # 获取类别
        #         soup_genres = soup.find_all('a', href=re.compile(
        #             r'https://www.javbus.com/genre/.|https://www.javbus.com/uncensored/genre.'))
        #         del soup_genres[-1];
        #         if soup_genres is None:
        #             av_genres = ['无类别']
        #         else:
        #             for soup_genre in soup_genres[2:]:
        #                 av_genres.append(soup_genre.string)
        #
        #         for av_genre in av_genres:
        #             av_genre = av_genre + ','
        #             new_av_genres = av_genre + new_av_genres
        #         print('获取类别成功！！')
        #
        #         # 把各种内容放到数据库中
        #         def mdb_conn(password=""):
        #             # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接
        #
        #             str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        #             conn = pypyodbc.win_connect_mdb(str)
        #             return conn
        #
        #         conn = mdb_conn()
        #         cur = conn.cursor()
        #         # 把所有数据插入av_record表
        #
        #         sql_insert = "INSERT INTO av_recommend (fanhao,director,zhizuo,faxing,xilie,genre,star,type,href) VALUES ('" + fanhao + "','" + av_director + "','" + av_zhizuo + "','" + av_faxing + "','" + av_xilie + "','" \
        #                      + new_av_genres + "','" + new_av_stars + "',0,'" + av_href + "')"
        #         cur.execute(sql_insert)
        #         conn.commit()
        #         print('增加新番号成功！！')
        #         cur.close()
        #         conn.close()


        # 处理全部评价最高的内容
        # count = 1
        # while count <= 10:
        #     topall_root = requests.get(root_url+'vl_bestrated.php?&mode=2&page='+ str(count), headers=headers).content
        #     topall_soup = BeautifulSoup(topall_root, 'html.parser', from_encoding='utf-8')
        #     topall_fanhao = topall_soup.find_all('div', class_='id')
        #     print("正在抓取第%d页:%s" % (count, topall_root + str(count)))
        #     for new_fanhao in topall_fanhao:
        #         fanhao = new_fanhao.get_text()
        #         print('正在获取:' + fanhao + '的相关信息')
        #         if self.isDownload(fanhao) == True and self.isInRecommand(fanhao) == True:
        #             print(fanhao + '该番号可以获取信息')
        #             av_director = ''
        #             av_zhizuo = ''
        #             av_faxing = ''
        #             av_xilie = ''
        #             av_genres = []
        #             av_stars = []
        #             new_av_genres = ''
        #             new_av_stars = ''
        #             av_href = ''
        #
        #             requ = requests.get('https://www.javbus.com/' + fanhao, headers=headers).content
        #             soup = BeautifulSoup(requ, 'html.parser', from_encoding='utf-8')
        #             # 获取封面地址
        #             soup_href = soup.find('a', class_='bigImage')
        #             if soup_href is None:
        #                 av_href = ''
        #             else:
        #                 av_href = soup_href.get('href')
        #             print('获取封面地址成功！！')
        #             # 获取演员
        #             soup_stars = soup.find_all('span', {'onmouseover': re.compile(r'hoverdiv')})
        #             print(soup_stars)
        #             if soup_stars == []:
        #                 av_stars = ''
        #             else:
        #                 for soup_star in soup_stars:
        #                     soup_star_txt = soup_star.get_text()
        #                     soup_star_txt = soup_star_txt.replace('\n', '')
        #                     av_stars.append(soup_star_txt)
        #
        #             for av_star in av_stars:
        #                 av_star = av_star + ','
        #                 new_av_stars = av_star + new_av_stars
        #             print('获取演员成功！！')
        #             # 获取导演
        #             soup_director = soup.find('a', href=re.compile(r'director'))
        #             if soup_director is None:
        #                 av_director = ''
        #             else:
        #                 av_director = soup_director.string
        #             print('获取导演成功！！')
        #             # 获取制作商
        #             soup_zhizuo = soup.find('a', href=re.compile(
        #                 r'https://www.javbus.com/studio|https://www.javbus.com/uncensored/studio'))
        #             if soup_zhizuo is None:
        #                 av_zhizuo = ''
        #             else:
        #                 av_zhizuo = soup_zhizuo.string
        #             print('获取制作商成功！！')
        #             # 获取发行商
        #             soup_faxing = soup.find('a', href=re.compile(
        #                 r'https://www.javbus.com/label|https://www.javbus.com/uncensored/label'))
        #             if soup_faxing is None:
        #                 av_faxing = ''
        #             else:
        #                 av_faxing = soup_faxing.string
        #             print('获取发行商成功！！')
        #             # 获取系列
        #             soup_xilie = soup.find('a', href=re.compile(
        #                 r'https://www.javbus.com/series|https://www.javbus.com/uncensored/series'))
        #             if soup_xilie is None:
        #                 av_xilie = ''
        #             else:
        #                 av_xilie = soup_xilie.string
        #             print('获取系列成功！！')
        #             # 获取类别
        #             soup_genres = soup.find_all('a', href=re.compile(
        #                 r'https://www.javbus.com/genre/.|https://www.javbus.com/uncensored/genre.'))
        #             del soup_genres[-1];
        #             if soup_genres is None:
        #                 av_genres = ['无类别']
        #             else:
        #                 for soup_genre in soup_genres[2:]:
        #                     av_genres.append(soup_genre.string)
        #
        #             for av_genre in av_genres:
        #                 av_genre = av_genre + ','
        #                 new_av_genres = av_genre + new_av_genres
        #             print('获取类别成功！！')
        #             # 把各种内容放到数据库中
        #             def mdb_conn(password=""):
        #                 # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接
        #
        #                 str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        #                 conn = pypyodbc.win_connect_mdb(str)
        #                 return conn
        #
        #             conn = mdb_conn()
        #             cur = conn.cursor()
        #             # 把所有数据插入av_record表
        #
        #             sql_insert = "INSERT INTO av_recommend (fanhao,director,zhizuo,faxing,xilie,genre,star,type,href) VALUES ('" + fanhao + "','" + av_director + "','" + av_zhizuo + "','" + av_faxing + "','" + av_xilie + "','" \
        #                          + new_av_genres + "','" + new_av_stars + "',0,'" + av_href + "')"
        #             cur.execute(sql_insert)
        #             conn.commit()
        #             print('增加新番号成功！！')
        #             cur.close()
        #             conn.close()
        #     count = count + 1

        # 处理上个月评价最高的内容
        # count = 1
        # while count <= 10:
        #     toplast_root = requests.get(root_url + 'vl_bestrated.php?&mode=1&page=' + str(count),
        #                                headers=headers).content
        #     toplast_soup = BeautifulSoup(toplast_root, 'html.parser', from_encoding='utf-8')
        #     toplast_fanhao = toplast_soup.find_all('div', class_='id')
        #     for new_fanhao in toplast_fanhao:
        #         fanhao = new_fanhao.get_text()
        #         print('正在获取:' + fanhao + '的相关信息')
        #         if self.isDownload(fanhao) == True and self.isInRecommand(fanhao) == True:
        #             print(fanhao + '该番号可以获取信息')
        #             av_director = ''
        #             av_zhizuo = ''
        #             av_faxing = ''
        #             av_xilie = ''
        #             av_genres = []
        #             av_stars = []
        #             new_av_genres = ''
        #             new_av_stars = ''
        #             av_href = ''
        #
        #             requ = requests.get('https://www.javbus.com/' + fanhao, headers=headers).content
        #             soup = BeautifulSoup(requ, 'html.parser', from_encoding='utf-8')
        #             # 获取封面地址
        #             soup_href = soup.find('a', class_='bigImage')
        #             if soup_href is None:
        #                 av_href = ''
        #             else:
        #                 av_href = soup_href.get('href')
        #             print('获取封面地址成功！！')
        #             # 获取演员
        #             soup_stars = soup.find_all('span', {'onmouseover': re.compile(r'hoverdiv')})
        #             print(soup_stars)
        #             if soup_stars == []:
        #                 av_stars = ''
        #             else:
        #                 for soup_star in soup_stars:
        #                     soup_star_txt = soup_star.get_text()
        #                     soup_star_txt = soup_star_txt.replace('\n', '')
        #                     av_stars.append(soup_star_txt)
        #
        #             for av_star in av_stars:
        #                 av_star = av_star + ','
        #                 new_av_stars = av_star + new_av_stars
        #             print('获取演员成功！！')
        #             # 获取导演
        #             soup_director = soup.find('a', href=re.compile(r'director'))
        #             if soup_director is None:
        #                 av_director = ''
        #             else:
        #                 av_director = soup_director.string
        #             print('获取导演成功！！')
        #             # 获取制作商
        #             soup_zhizuo = soup.find('a', href=re.compile(
        #                 r'https://www.javbus.com/studio|https://www.javbus.com/uncensored/studio'))
        #             if soup_zhizuo is None:
        #                 av_zhizuo = ''
        #             else:
        #                 av_zhizuo = soup_zhizuo.string
        #             print('获取制作商成功！！')
        #             # 获取发行商
        #             soup_faxing = soup.find('a', href=re.compile(
        #                 r'https://www.javbus.com/label|https://www.javbus.com/uncensored/label'))
        #             if soup_faxing is None:
        #                 av_faxing = ''
        #             else:
        #                 av_faxing = soup_faxing.string
        #             print('获取发行商成功！！')
        #             # 获取系列
        #             soup_xilie = soup.find('a', href=re.compile(
        #                 r'https://www.javbus.com/series|https://www.javbus.com/uncensored/series'))
        #             if soup_xilie is None:
        #                 av_xilie = ''
        #             else:
        #                 av_xilie = soup_xilie.string
        #             print('获取系列成功！！')
        #             # 获取类别
        #             soup_genres = soup.find_all('a', href=re.compile(
        #                 r'https://www.javbus.com/genre/.|https://www.javbus.com/uncensored/genre.'))
        #             del soup_genres[-1];
        #             if soup_genres is None:
        #                 av_genres = ['无类别']
        #             else:
        #                 for soup_genre in soup_genres[2:]:
        #                     av_genres.append(soup_genre.string)
        #
        #             for av_genre in av_genres:
        #                 av_genre = av_genre + ','
        #                 new_av_genres = av_genre + new_av_genres
        #             print('获取类别成功！！')
        #
        #             # 把各种内容放到数据库中
        #             def mdb_conn(password=""):
        #                 # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接
        #
        #                 str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        #                 conn = pypyodbc.win_connect_mdb(str)
        #                 return conn
        #
        #             conn = mdb_conn()
        #             cur = conn.cursor()
        #             # 把所有数据插入av_record表
        #
        #             sql_insert = "INSERT INTO av_recommend (fanhao,director,zhizuo,faxing,xilie,genre,star,type,href) VALUES ('" + fanhao + "','" + av_director + "','" + av_zhizuo + "','" + av_faxing + "','" + av_xilie + "','" \
        #                          + new_av_genres + "','" + new_av_stars + "',0,'" + av_href + "')"
        #             cur.execute(sql_insert)
        #             conn.commit()
        #             print('增加新番号成功！！')
        #             cur.close()
        #             conn.close()
        #     count = count + 1

        # 处理全部最想要的的内容
        # count = 1
        # while count <= 10:
        #     mostwantedall_root = requests.get(root_url + 'vl_mostwanted.php?&mode=2&page=' + str(count),
        #                                 headers=headers).content
        #     mostwantedall_soup = BeautifulSoup(mostwantedall_root, 'html.parser', from_encoding='utf-8')
        #     mostwantedall_fanhao = mostwantedall_soup.find_all('div', class_='id')
        #     for new_fanhao in mostwantedall_fanhao:
        #         fanhao = new_fanhao.get_text()
        #         print('正在获取:' + fanhao + '的相关信息')
        #         if self.isDownload(fanhao) == True and self.isInRecommand(fanhao) == True:
        #             print(fanhao + '该番号可以获取信息')
        #             av_director = ''
        #             av_zhizuo = ''
        #             av_faxing = ''
        #             av_xilie = ''
        #             av_genres = []
        #             av_stars = []
        #             new_av_genres = ''
        #             new_av_stars = ''
        #             av_href = ''
        #
        #             requ = requests.get('https://www.javbus.com/' + fanhao, headers=headers).content
        #             soup = BeautifulSoup(requ, 'html.parser', from_encoding='utf-8')
        #             # 获取封面地址
        #             soup_href = soup.find('a', class_='bigImage')
        #             if soup_href is None:
        #                 av_href = ''
        #             else:
        #                 av_href = soup_href.get('href')
        #             print('获取封面地址成功！！')
        #             # 获取演员
        #             soup_stars = soup.find_all('span', {'onmouseover': re.compile(r'hoverdiv')})
        #             print(soup_stars)
        #             if soup_stars == []:
        #                 av_stars = ''
        #             else:
        #                 for soup_star in soup_stars:
        #                     soup_star_txt = soup_star.get_text()
        #                     soup_star_txt = soup_star_txt.replace('\n', '')
        #                     av_stars.append(soup_star_txt)
        #
        #             for av_star in av_stars:
        #                 av_star = av_star + ','
        #                 new_av_stars = av_star + new_av_stars
        #             print('获取演员成功！！')
        #             # 获取导演
        #             soup_director = soup.find('a', href=re.compile(r'director'))
        #             if soup_director is None:
        #                 av_director = ''
        #             else:
        #                 av_director = soup_director.string
        #             print('获取导演成功！！')
        #             # 获取制作商
        #             soup_zhizuo = soup.find('a', href=re.compile(
        #                 r'https://www.javbus.com/studio|https://www.javbus.com/uncensored/studio'))
        #             if soup_zhizuo is None:
        #                 av_zhizuo = ''
        #             else:
        #                 av_zhizuo = soup_zhizuo.string
        #             print('获取制作商成功！！')
        #             # 获取发行商
        #             soup_faxing = soup.find('a', href=re.compile(
        #                 r'https://www.javbus.com/label|https://www.javbus.com/uncensored/label'))
        #             if soup_faxing is None:
        #                 av_faxing = ''
        #             else:
        #                 av_faxing = soup_faxing.string
        #             print('获取发行商成功！！')
        #             # 获取系列
        #             soup_xilie = soup.find('a', href=re.compile(
        #                 r'https://www.javbus.com/series|https://www.javbus.com/uncensored/series'))
        #             if soup_xilie is None:
        #                 av_xilie = ''
        #             else:
        #                 av_xilie = soup_xilie.string
        #             print('获取系列成功！！')
        #             # 获取类别
        #             soup_genres = soup.find_all('a', href=re.compile(
        #                 r'https://www.javbus.com/genre/.|https://www.javbus.com/uncensored/genre.'))
        #             del soup_genres[-1];
        #             if soup_genres is None:
        #                 av_genres = ['无类别']
        #             else:
        #                 for soup_genre in soup_genres[2:]:
        #                     av_genres.append(soup_genre.string)
        #
        #             for av_genre in av_genres:
        #                 av_genre = av_genre + ','
        #                 new_av_genres = av_genre + new_av_genres
        #             print('获取类别成功！！')
        #
        #             # 把各种内容放到数据库中
        #             def mdb_conn(password=""):
        #                 # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接
        #
        #                 str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        #                 conn = pypyodbc.win_connect_mdb(str)
        #                 return conn
        #
        #             conn = mdb_conn()
        #             cur = conn.cursor()
        #             # 把所有数据插入av_record表
        #
        #             sql_insert = "INSERT INTO av_recommend (fanhao,director,zhizuo,faxing,xilie,genre,star,type,href) VALUES ('" + fanhao + "','" + av_director + "','" + av_zhizuo + "','" + av_faxing + "','" + av_xilie + "','" \
        #                          + new_av_genres + "','" + new_av_stars + "',0,'" + av_href + "')"
        #             cur.execute(sql_insert)
        #             conn.commit()
        #             print('增加新番号成功！！')
        #             cur.close()
        #             conn.close()
        #     count = count + 1

        # 处理上个月最想要的的内容
        count = 1
        while count <= 10:
            mostwantedlast_root = requests.get(root_url + 'vl_mostwanted.php?&mode=1&page=' + str(count),
                                              headers=headers).content
            mostwantedlast_soup = BeautifulSoup(mostwantedlast_root, 'html.parser', from_encoding='utf-8')
            mostwantedlast_fanhao = mostwantedlast_soup.find_all('div', class_='id')
            for new_fanhao in mostwantedlast_fanhao:
                fanhao = new_fanhao.get_text()
                print('正在获取:' + fanhao + '的相关信息')
                if self.isDownload(fanhao) == True and self.isInRecommand(fanhao) == True:
                    print(fanhao + '该番号可以获取信息')
                    av_director = ''
                    av_zhizuo = ''
                    av_faxing = ''
                    av_xilie = ''
                    av_genres = []
                    av_stars = []
                    new_av_genres = ''
                    new_av_stars = ''
                    av_href = ''

                    requ = requests.get('https://www.javbus.com/' + fanhao, headers=headers).content
                    soup = BeautifulSoup(requ, 'html.parser', from_encoding='utf-8')
                    # 获取封面地址
                    soup_href = soup.find('a', class_='bigImage')
                    if soup_href is None:
                        av_href = ''
                    else:
                        av_href = soup_href.get('href')
                    print('获取封面地址成功！！')
                    # 获取演员
                    soup_stars = soup.find_all('span', {'onmouseover': re.compile(r'hoverdiv')})
                    print(soup_stars)
                    if soup_stars == []:
                        av_stars = ''
                    else:
                        for soup_star in soup_stars:
                            soup_star_txt = soup_star.get_text()
                            soup_star_txt = soup_star_txt.replace('\n', '')
                            av_stars.append(soup_star_txt)

                    for av_star in av_stars:
                        av_star = av_star + ','
                        new_av_stars = av_star + new_av_stars
                    print('获取演员成功！！')
                    # 获取导演
                    soup_director = soup.find('a', href=re.compile(r'director'))
                    if soup_director is None:
                        av_director = ''
                    else:
                        av_director = soup_director.string
                    print('获取导演成功！！')
                    # 获取制作商
                    soup_zhizuo = soup.find('a', href=re.compile(
                        r'https://www.javbus.com/studio|https://www.javbus.com/uncensored/studio'))
                    if soup_zhizuo is None:
                        av_zhizuo = ''
                    else:
                        av_zhizuo = soup_zhizuo.string
                    print('获取制作商成功！！')
                    # 获取发行商
                    soup_faxing = soup.find('a', href=re.compile(
                        r'https://www.javbus.com/label|https://www.javbus.com/uncensored/label'))
                    if soup_faxing is None:
                        av_faxing = ''
                    else:
                        av_faxing = soup_faxing.string
                    print('获取发行商成功！！')
                    # 获取系列
                    soup_xilie = soup.find('a', href=re.compile(
                        r'https://www.javbus.com/series|https://www.javbus.com/uncensored/series'))
                    if soup_xilie is None:
                        av_xilie = ''
                    else:
                        av_xilie = soup_xilie.string
                    print('获取系列成功！！')
                    # 获取类别
                    soup_genres = soup.find_all('a', href=re.compile(
                        r'https://www.javbus.com/genre/.|https://www.javbus.com/uncensored/genre.'))
                    del soup_genres[-1];
                    if soup_genres is None:
                        av_genres = ['无类别']
                    else:
                        for soup_genre in soup_genres[2:]:
                            av_genres.append(soup_genre.string)

                    for av_genre in av_genres:
                        av_genre = av_genre + ','
                        new_av_genres = av_genre + new_av_genres
                    print('获取类别成功！！')

                    # 把各种内容放到数据库中
                    def mdb_conn(password=""):
                        # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

                        str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
                        conn = pypyodbc.win_connect_mdb(str)
                        return conn

                    conn = mdb_conn()
                    cur = conn.cursor()
                    # 把所有数据插入av_record表

                    sql_insert = "INSERT INTO av_recommend (fanhao,director,zhizuo,faxing,xilie,genre,star,type,href) VALUES ('" + fanhao + "','" + av_director + "','" + av_zhizuo + "','" + av_faxing + "','" + av_xilie + "','" \
                                 + new_av_genres + "','" + new_av_stars + "',0,'" + av_href + "')"
                    cur.execute(sql_insert)
                    conn.commit()
                    print('增加新番号成功！！')
                    cur.close()
                    conn.close()
            count = count + 1
        wx.MessageBox("爬虫爬取结束", "Message", wx.OK | wx.ICON_INFORMATION)
    # 爬虫2号无码脚本
    def Spider2(self):
        root_url = 'https://www.javbus.com/uncensored/page/'
        headers = {'content-type': 'application/json',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
        count = 1
        while count <= 2:
            print("正在抓取第%d页:%s" % (count, root_url + str(count)))
            try:
                re_root = requests.get(root_url + str(count), headers=headers).content
                soup = BeautifulSoup(re_root, 'html.parser', from_encoding='utf-8')
                new_fanhaos = soup.find_all('a', class_='movie-box')
                for new_fanhao in new_fanhaos:
                    new_url = new_fanhao.get('href')
                    print('正在处理:' + new_url)
                    re_new = requests.get(new_url, headers=headers).content
                    soup_new = BeautifulSoup(re_new, 'html.parser', from_encoding='utf-8')
                    fanhao = new_url.split('/')[-1]
                    print('正在获取:' + fanhao + '的相关信息')
                    if self.isDownload(fanhao) == True and self.isInRecommand(fanhao) == True:
                        print(fanhao + '该番号可以获取信息')
                        av_director = ''
                        av_zhizuo = ''
                        av_faxing = ''
                        av_xilie = ''
                        av_genres = []
                        av_stars = []
                        new_av_genres = ''
                        new_av_stars = ''
                        av_href = ''

                        # 获取封面地址
                        soup_href = soup_new.find('a', class_='bigImage')
                        if soup_href is None:
                            av_href = ''
                        else:
                            av_href = soup_href.get('href')

                        # 获取演员
                        soup_stars = soup_new.find_all('span', {'onmouseover': re.compile(r'hoverdiv')})
                        print(soup_stars)
                        if soup_stars == []:
                            av_stars = ''
                        else:
                            for soup_star in soup_stars:
                                soup_star_txt = soup_star.get_text()
                                soup_star_txt = soup_star_txt.replace('\n', '')
                                av_stars.append(soup_star_txt)

                        for av_star in av_stars:
                            av_star = av_star + ','
                            new_av_stars = av_star + new_av_stars

                        # 获取导演
                        soup_director = soup_new.find('a', href=re.compile(r'director'))
                        if soup_director is None:
                            av_director = ''
                        else:
                            av_director = soup_director.string

                        # 获取制作商
                        soup_zhizuo = soup_new.find('a', href=re.compile(
                            r'https://www.javbus.com/studio|https://www.javbus.com/uncensored/studio'))
                        if soup_zhizuo is None:
                            av_zhizuo = ''
                        else:
                            av_zhizuo = soup_zhizuo.string

                        # 获取发行商
                        soup_faxing = soup_new.find('a', href=re.compile(
                            r'https://www.javbus.com/label|https://www.javbus.com/uncensored/label'))
                        if soup_faxing is None:
                            av_faxing = ''
                        else:
                            av_faxing = soup_faxing.string

                        # 获取系列
                        soup_xilie = soup_new.find('a', href=re.compile(
                            r'https://www.javbus.com/series|https://www.javbus.com/uncensored/series'))
                        if soup_xilie is None:
                            av_xilie = ''
                        else:
                            av_xilie = soup_xilie.string

                        # 获取类别
                        soup_genres = soup_new.find_all('a', href=re.compile(
                            r'https://www.javbus.com/genre/.|https://www.javbus.com/uncensored/genre.'))
                        del soup_genres[-1];
                        if soup_genres is None:
                            av_genres = ['无类别']
                        else:
                            for soup_genre in soup_genres[2:]:
                                av_genres.append(soup_genre.string)

                        for av_genre in av_genres:
                            av_genre = av_genre + ','
                            new_av_genres = av_genre + new_av_genres

                        # 把各种内容放到数据库中
                        def mdb_conn(password=""):
                            # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

                            str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
                            conn = pypyodbc.win_connect_mdb(str)
                            return conn

                        conn = mdb_conn()
                        cur = conn.cursor()

                        # 把所有数据插入av_record表

                        sql_insert = "INSERT INTO av_recommend (fanhao,director,zhizuo,faxing,xilie,genre,star,type,href) VALUES ('" + fanhao + "','" + av_director + "','" + av_zhizuo + "','" + av_faxing + "','" + av_xilie + "','" \
                                     + new_av_genres + "','" + new_av_stars + "',1,'" + av_href + "')"
                        cur.execute(sql_insert)
                        conn.commit()
                        print('增加新番号成功！！')
                        cur.close()
                        conn.close()
            except:
                continue
            count = count + 1
        wx.MessageBox("爬虫爬取结束", "Message", wx.OK | wx.ICON_INFORMATION)