# *-* coding:utf-8 *-*
import shutil
import sys
import requests
from bs4 import BeautifulSoup
import os
import re
import pypyodbc

def fileDone(fanhao,cover_url):

    #创建文件夹
    if os.path.exists('资源/'+fanhao):
        print('文件夹已经存在')
        pass
    else:
        os.makedirs('资源/'+fanhao)
        print('创建文件夹成功')

    #获取图片
    try:
        print('正在获取图片：')
        pic = requests.get(cover_url, timeout=2)
        print('获取图片成功')
    except requests.exceptions.ConnectionError:
        print('【错误】当前图片无法下载')

    string = '资源/'+ fanhao + '/cover.jpg'
    fp = open(string, 'wb')
    # 创建文件
    fp.write(pic.content)
    fp.close()
    print('图片下载成功')
    try:
        shutil.move(fanhaoPlus, '资源/'+fanhao + '/' + fanhaoPlus)
    except:
        print('文件不存在')

    print('文件全部处理成功')

def db_add(fanhao,av_director,pid_director,av_faxing,pid_faxing,av_zhizuo,pid_zhizuo,av_xilie,pid_xilie,av_genres,pid_genres,av_stars,pid_stars):

    new_av_genres =''
    for av_genre in av_genres:
        av_genre = av_genre + ','
        new_av_genres =av_genre+new_av_genres

    new_av_stars = ''
    for av_star in av_stars:
        av_star = av_star + ','
        new_av_stars = av_star + new_av_stars

    print (new_av_genres)
    print(new_av_stars)
    def mdb_conn(password=""):
    #功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

        str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        conn = pypyodbc.win_connect_mdb(str)
        return conn

    conn = mdb_conn()
    cur = conn.cursor()

    #把所有数据插入av_record表
    sql_sel ="SELECT COUNT(*) FROM av_record WHERE ID = '"+fanhao+"'"
    cur.execute(sql_sel)
    num_fanhao = cur.fetchall()[0][0]
    if num_fanhao == 1:
        print('该番号已存在！！')
    else:
        sql_insert = "INSERT INTO av_record VALUES ('"+ fanhao +"','"+ av_director+"','"+av_faxing+"','"+av_zhizuo+"','"+av_xilie+"','"\
                     +new_av_genres+"','"+new_av_stars+"','0','1')"
        cur.execute(sql_insert)
        conn.commit()
        print('增加新番号成功！！')


    # 把所有数据插入av_director表
    if av_director is not None:
        sql_sel = "SELECT COUNT(director) FROM av_director WHERE director = '" + av_director + "'"
        cur.execute(sql_sel)
        num_director = cur.fetchall()[0][0]
        if num_director == 1:
            sql_insert = "UPDATE av_director SET num = num+1 WHERE director ='" + av_director + "'"
            print('导演已存在，导演加分成功！！')
        else:
            sql_insert = "INSERT INTO av_director(director,num,pid) VALUES ('" + av_director + "','1','"+pid_director+"')"
            print('新导演添加！！')
        cur.execute(sql_insert)
        conn.commit()

    # 把所有数据插入av_faxing表
    if av_faxing is not None:
        sql_sel = "SELECT COUNT(faxing) FROM av_faxing WHERE faxing = '" + av_faxing + "'"
        cur.execute(sql_sel)
        num_faxing = cur.fetchall()[0][0]
        if num_faxing == 1:
            sql_insert = "UPDATE av_faxing SET num = num+1 WHERE faxing ='" +av_faxing + "'"
            print('发行商已存在，发行商加分成功！！！')
        else:
            sql_insert = "INSERT INTO av_faxing(faxing,num,pid) VALUES ('"+ av_faxing+"','1','"+pid_faxing+"')"
            print('新发行商添加！！！')
        cur.execute(sql_insert)
        conn.commit()

    # 把所有数据插入av_zhizuo表
    if av_zhizuo is not None:
        sql_sel = "SELECT COUNT(zhizuo) FROM av_zhizuo WHERE zhizuo = '" + av_zhizuo + "'"
        cur.execute(sql_sel)
        num_zhizuo = cur.fetchall()[0][0]
        if num_zhizuo == 1:
            sql_insert = "UPDATE av_zhizuo SET num = num+1 WHERE zhizuo ='" + av_zhizuo + "'"
            print('制作商已存在，制作商加分成功！！！')
        else:
            sql_insert = "INSERT INTO av_zhizuo(zhizuo,num,pid) VALUES ('" + av_zhizuo + "','1','"+pid_zhizuo+"')"
            print('新制作商添加！！！')
        cur.execute(sql_insert)
        conn.commit()

    # 把所有数据插入av_xilie表
    if av_xilie is not None:
        sql_sel = "SELECT COUNT(xilie) FROM av_xilie WHERE xilie = '" + av_xilie + "'"
        cur.execute(sql_sel)
        num_zhizuo = cur.fetchall()[0][0]
        if num_zhizuo == 1:
            sql_insert = "UPDATE av_xilie SET num = num+1 WHERE xilie ='" + av_xilie + "'"
            print('该系列已存在，系列加分成功！！！')
        else:
            sql_insert = "INSERT INTO av_xilie(xilie,num,pid) VALUES ('" + av_xilie + "','1','"+pid_xilie+"')"
            print('新系列添加！！！')
        cur.execute(sql_insert)
        conn.commit()

    # 把所有数据插入av_genre表
    for i in range(len(av_genres)):
        sql_sel = "SELECT COUNT(genre) FROM av_genres WHERE genre = '" + av_genres[i] + "'"
        cur.execute(sql_sel)
        num_genre = cur.fetchall()[0][0]
        if num_genre == 1:
            sql_insert = "UPDATE av_genres SET num = num+1 WHERE genre ='" + av_genres[i] + "'"
            print('该类别已存在，类别加分成功')
        else:
            sql_insert = "INSERT INTO av_genres(genre,num,pid) VALUES ('" + av_genres[i] + "','1','"+pid_genres[i] +"')"
            print(sql_insert)
        cur.execute(sql_insert)
        conn.commit()

    # 把所有数据插入av_stars表
    for j in range(len(av_stars)):
        sql_sel = "SELECT COUNT(star) FROM av_stars WHERE star = '" + av_stars[j] + "'"
        cur.execute(sql_sel)
        num_star = cur.fetchall()[0][0]
        if num_star == 1:
            sql_insert = "UPDATE av_stars SET num = num+1 WHERE star ='" + av_stars[j] + "'"
            print('该演员已存在，演员加分成功！！！')
        else:
            sql_insert = "INSERT INTO av_stars(star,num,pid) VALUES ('" + av_stars[j] + "','1','"+pid_stars[j] +"')"

            print('新演员添加！！！')
        cur.execute(sql_insert)
        conn.commit()


    cur.close()
    conn.close()

def db_delete(fanhao,av_director,av_faxing,av_zhizuo,av_xilie,av_genres,av_stars):

    new_av_genres =''
    for av_genre in av_genres:
        av_genre = av_genre + ','
        new_av_genres =av_genre+new_av_genres

    new_av_stars = ''
    for av_star in av_stars:
        av_star = av_star + ','
        new_av_stars = av_star + new_av_stars

    print (new_av_genres)
    print(new_av_stars)
    def mdb_conn(password=""):
    #功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

        str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        conn = pypyodbc.win_connect_mdb(str)
        return conn

    conn = mdb_conn()
    cur = conn.cursor()

    #删除av_record表中该番号数据
    sql_sel ="SELECT COUNT(*) FROM av_record WHERE ID = '"+fanhao+"'"
    cur.execute(sql_sel)
    num_fanhao = cur.fetchall()[0][0]
    if num_fanhao == 1:
        print('该番号已存在,即将删除')
        sql_delete = "DELETE FROM av_record WHERE ID = '" + fanhao + "'"
        cur.execute(sql_delete)
        conn.commit()
    else:
        print('该番号不存在，无法删除')


    #av_director表中导演扣一分
    sql_sel = "SELECT COUNT(director) FROM av_director WHERE director = '" + av_director + "'"
    cur.execute(sql_sel)
    num_director = cur.fetchall()[0][0]
    if num_director == 1:
        print('导演已存在')
        sql_insert = "UPDATE av_director SET num = num-1 WHERE director ='" + av_director + "'"
        print('导演扣分成功')
        cur.execute(sql_insert)
        conn.commit()
    else:
        print('导演不存在')


    # av_faxing表中发行商扣一分
    sql_sel = "SELECT COUNT(faxing) FROM av_faxing WHERE faxing = '" + av_faxing + "'"
    cur.execute(sql_sel)
    num_faxing = cur.fetchall()[0][0]
    if num_faxing == 1:
        print('发行商已存在')
        sql_insert = "UPDATE av_faxing SET num = num-1 WHERE faxing ='" +av_faxing+ "'"
        print('发行商扣分成功')
        cur.execute(sql_insert)
        conn.commit()
    else:
        print('发行商不存在')


    # 把所有数据插入av_zhizuo表
    sql_sel = "SELECT COUNT(zhizuo) FROM av_zhizuo WHERE zhizuo = '" + av_zhizuo + "'"
    cur.execute(sql_sel)
    num_zhizuo = cur.fetchall()[0][0]
    if num_zhizuo == 1:
        print('制作商已存在')
        sql_insert = "UPDATE av_zhizuo SET num = num-1 WHERE zhizuo ='" + av_zhizuo + "'"
        print('制作商扣分成功')
        cur.execute(sql_insert)
        conn.commit()
    else:
        print('制作商不存在')


    # 把所有数据插入av_xilie表
    sql_sel = "SELECT COUNT(xilie) FROM av_xilie WHERE xilie = '" + av_xilie + "'"
    cur.execute(sql_sel)
    num_zhizuo = cur.fetchall()[0][0]
    if num_zhizuo == 1:
        print('该系列已存在')
        sql_insert = "UPDATE av_xilie SET num = num-1 WHERE xilie ='" + av_xilie + "'"
        print('系列扣分成功')
        cur.execute(sql_insert)
        conn.commit()
    else:
        print('系列不存在')


    # 把所有数据插入av_genre表
    for table_av_genre in av_genres:
        sql_sel = "SELECT COUNT(genre) FROM av_genres WHERE genre = '" + table_av_genre + "'"
        cur.execute(sql_sel)
        num_genre = cur.fetchall()[0][0]
        if num_genre == 1:
            print('该类别已存在')
            sql_insert = "UPDATE av_genres SET num = num-1 WHERE genre ='" + table_av_genre + "'"
            print('类别扣分成功')
            cur.execute(sql_insert)
            conn.commit()
        else:
            print('类别不存在')


    # 把所有数据插入av_stars表
    for table_av_star in av_stars:
        sql_sel = "SELECT COUNT(star) FROM av_stars WHERE star = '" + table_av_star + "'"
        cur.execute(sql_sel)
        num_genre = cur.fetchall()[0][0]
        if num_genre == 1:
            print('该演员已存在')
            sql_insert = "UPDATE av_stars SET num = num-1 WHERE star ='" + table_av_star + "'"
            print('演员扣分成功')
            cur.execute(sql_insert)
            conn.commit()
        else:
            print('演员不存在')

    cur.close()
    conn.close()

def getInfo(fanhao):
    #请求内容
    print('正在请求：' + fanhao)
    headers = {'content-type': 'application/json',
               'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}

    try:
        requ = requests.get('https://www.javbus.com/'+fanhao,timeout=1,headers=headers).content
    except:
        print("获取网址超时")
        return


    #获取标签内容
    soup = BeautifulSoup(requ, 'html.parser', from_encoding='utf-8')
    #先判断是否能正常搜索到内容
    is404 = soup.find('h4',{'style':'font-size:36px;'})
    if is404 is not None:
        print('------该番号不存在------')
        return

    cover_url = soup.find('a',class_='bigImage').get('href')

    #获取导演
    soup_director = soup.find('a',href=re.compile(r'director'))
    if soup_director is None:
        av_director = ''
        pid_director=''
    else:
        av_director =soup_director.string
        pid_director = soup_director.get('href').split('/')[-1]
    print('导演: ' + av_director + pid_director)

    #获取制作商
    soup_zhizuo = soup.find('a',href=re.compile(r'https://www.javbus.com/studio|https://www.javbus.com/uncensored/studio'))
    if soup_zhizuo is None:
        av_zhizuo = ''
        pid_zhizuo =''
    else:
        av_zhizuo =soup_zhizuo.string
        pid_zhizuo = soup_zhizuo.get('href').split('/')[-1]
    print('制作商: ' + av_zhizuo+pid_zhizuo)

    #获取发行商
    soup_faxing = soup.find('a',href=re.compile(r'https://www.javbus.com/label|https://www.javbus.com/uncensored/label'))
    if soup_faxing is None:
        av_faxing = ''
        pid_faxing = ''
    else:
        av_faxing =soup_faxing.string
        pid_faxing = soup_faxing.get('href').split('/')[-1]
    print('发行商： ' + av_faxing + pid_faxing)
    # 获取系列
    soup_xilie = soup.find('a', href=re.compile(r'https://www.javbus.com/series|https://www.javbus.com/uncensored/series'))
    if soup_xilie is None:
        av_xilie = ''
        pid_xilie=''
    else:
        av_xilie = soup_xilie.string
        pid_xilie = soup_xilie.get('href').split('/')[-1]
    print('系列： ' + av_xilie +pid_xilie)
    # 获取类别
    av_genres=[]
    pid_genres=[]
    soup_genres = soup.find_all('a', href=re.compile(r'https://www.javbus.com/genre/.|https://www.javbus.com/uncensored/genre.'))
    del soup_genres[-1];
    if soup_genres is None:
        av_genres = ['无类别']
        pid_genres = ['']
    else:
        for soup_genre in soup_genres[2:]:
            av_genres.append(soup_genre.string)
            pid_genre=soup_genre.get('href').split('/')[-1]
            pid_genres.append(pid_genre)
    # 获取演员
    av_stars = []
    pid_stars =[]
    soup_stars = soup.find_all('span', {'onmouseover':re.compile(r'hoverdiv')})
    if soup_stars is None:
        av_stars = ['无演员']
        pid_stars = ['']
    else:
        for soup_star in soup_stars:
            soup_star_txt = soup_star.get_text()
            soup_star_txt = soup_star_txt.replace('\n','')
            pid_star = soup_star.a.get('href').split('/')[-1]
            av_stars.append(soup_star_txt)
            pid_stars.append(pid_star)

    fileDone(fanhao,cover_url)
    db_add(fanhao,av_director,pid_director,av_faxing,pid_faxing,av_zhizuo,pid_zhizuo,av_xilie,pid_xilie,av_genres,pid_genres,av_stars,pid_stars)
    #db_delete(fanhao, av_director, av_faxing, av_zhizuo, av_xilie, av_genres, av_stars)


if __name__ == '__main__':
    # 获取番号名称
    fanhaoPlus = sys.argv[1]
    #fanhaoPlus = '032916_270.mp4'
    fanhao = fanhaoPlus.split('.')[0]
    print(fanhaoPlus)

    #获取番号信息并入库
    getInfo(fanhao)


    os.system("pause")
