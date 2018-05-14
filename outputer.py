import pypyodbc
import codecs
import os
import datetime

def output_html():

    def mdb_conn(password=""):
        # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

        str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        conn = pypyodbc.win_connect_mdb(str)
        return conn

    conn = mdb_conn()
    cur = conn.cursor()

    #fout = codecs.open('index.html', 'w', 'utf-8')

    fout = open("index.html","w")
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
    fout.write("<h4>推荐观看</h4>")
    # 获取TOP10番号
    sql_sel = "SELECT TOP 10 * FROM av_record ORDER BY score DESC"
    cur.execute(sql_sel)
    infos = cur.fetchall()
    print(infos)

    for info in infos:
        print(info[5])
        fout.write("<div class='media'>")
        fout.write("<a class='cover pull-left'>")
        fout.write("<img src='G:/fun_finder/资源/" + info[0] +"/cover.jpg'>")
        fout.write("</a>")
        fout.write("<div class='media-body'>")
        fout.write("<h3 class='media-heading'>" + info[0] +"</h3>")
        fout.write("<p>演员：" + info[6] +"</p>")
        fout.write("<p>类别：" + info[5] +"</p>")
        print(info[5])
        fout.write("<h3 class='text-error bold'>" + str(info[8]) +"</h3>")
        fout.write("</div>")
        fout.write("</div>")
    fout.write("</div>")

    fout.write("<div class='span4''>")
    fout.write("<h4>最喜欢的star</h4>")

    # 获取TOP10 stars
    sql_sel_star = "SELECT TOP 5 * FROM av_stars ORDER BY num DESC"
    cur.execute(sql_sel_star)
    star_infos = cur.fetchall()
    print(star_infos)

    for star_info in star_infos:
        fout.write("<div class='media'>")
        fout.write("<a class='pull-left avatar' target='_blank' href='https://www.javbus.com/star/" + str(star_info[3]) +"'>")
        fout.write("<img src='https://pics.javbus.com/actress/" + str(star_info[3]) +"_a.jpg' class='img-circle'></a>")
        fout.write("<div class='media-body''>")
        fout.write("<h4 class='media-heading'>" + str(star_info[1]) +"</h4>")
        fout.write("</div>")
        fout.write("</div>")

    fout.write("<h4 style='margin-top:40px'>最喜欢的类型</h4>")
    # 获取TOP10类型
    sql_sel = "SELECT TOP 10 * FROM av_genres ORDER BY num DESC"
    cur.execute(sql_sel)
    genre_infos = cur.fetchall()
    print(genre_infos)
    for genre_info in genre_infos:
        fout.write("<a class='label label-success' href='https://www.javbus.com/genre/" + str(genre_info[3]) +"' style='margin-left:20px'>" + str(genre_info[1]) +"</a>")

    fout.write("<h4 style='margin-top:40px'>最喜欢的导演</h4>")
    # 获取TOP10类型
    sql_sel = "SELECT TOP 5 * FROM av_director ORDER BY num DESC"
    cur.execute(sql_sel)
    director_infos = cur.fetchall()
    print(director_infos)
    for director_info in director_infos:
        if director_info[1] != '':
            fout.write("<a  href='https://www.javbus.com/director/" + str(
            director_info[3]) + "' style='margin-left:20px'>" + str(director_info[1]) + "</a><br>")

    fout.write("<h4 style='margin-top:40px'>最喜欢的系列</h4>")
    # 获取TOP10类型
    sql_sel = "SELECT TOP 5 * FROM av_xilie ORDER BY num DESC"
    cur.execute(sql_sel)
    xilie_infos = cur.fetchall()
    print(xilie_infos)
    for xilie_info in xilie_infos:
        if xilie_info[1] != '':
            fout.write("<a  href='https://www.javbus.com/series/" + str(
                xilie_info[3]) + "' style='margin-left:20px'>" + str(xilie_info[1]) + "</a><br>")

    fout.write("</div>")
    fout.write("</div>")

    fout.write("<script src='http://code.jquery.com/jquery.js'></script>")
    fout.write("<script src='bootstrap/js/bootstrap.min.js''></script>")
    fout.write("</body>")
    fout.write("</html>")

    cur.close()
    conn.close()

def updateScore(fanhao):
    #通过番号获取各种信息
    def mdb_conn(password=""):
    #功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

        str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        conn = pypyodbc.win_connect_mdb(str)
        return conn

    conn = mdb_conn()
    cur = conn.cursor()


    #获取各种信息
    sql_sel = "SELECT * FROM av_record WHERE ID = '" + fanhao + "'"
    cur.execute(sql_sel)
    av_info = cur.fetchall()
    print(av_info)

    #获取导演数量
    if av_info[0][1] == '':
        num_director = 0
    else:
        av_director = av_info[0][1]
        sql_sel = "SELECT num FROM av_director WHERE director = '" + av_director + "'"
        cur.execute(sql_sel)
        num_director = cur.fetchall()[0][0]
    print('导演数量：' + str(num_director))

    # 获取发行商数量
    if av_info[0][2] == '':
        num_faxing=0
    else:
        av_faxing = av_info[0][2]
        sql_sel = "SELECT num FROM av_faxing WHERE faxing = '" + av_faxing + "'"
        cur.execute(sql_sel)
        num_faxing = cur.fetchall()[0][0]
    print('发行商数量：' + str(num_faxing))

    # 获取制作商数量
    if av_info[0][3] == '':
        num_zhizuo=0
    else:
        av_zhizuo = av_info[0][3]
        sql_sel = "SELECT num FROM av_zhizuo WHERE zhizuo = '" + av_zhizuo + "'"
        cur.execute(sql_sel)
        num_zhizuo = cur.fetchall()[0][0]
    print('制作商数量：' + str(num_zhizuo))

    # 获取系列数量
    if av_info[0][4] == '':
        num_xilie=0
    else:
        av_xilie = av_info[0][4]
        sql_sel = "SELECT num FROM av_xilie WHERE xilie = '" + av_xilie + "'"
        cur.execute(sql_sel)
        num_xilie = cur.fetchall()[0][0]
    print('系列数量：' + str(num_xilie))


    # 获取各类别数量
    if av_info[0][5] is None:
        av_genres=''
    else:
        av_genres = av_info[0][5]
    new_av_genres =av_genres.split(',')
    del new_av_genres[-1]
    print(new_av_genres)
    num_genres = 0
    for av_genre in new_av_genres:
        sql_sel = "SELECT num FROM av_genres WHERE genre = '" + av_genre + "'"
        cur.execute(sql_sel)
        num_genre = cur.fetchall()[0][0]
        num_genres = num_genres + num_genre
        print(av_genre +'类别数量：' + str(num_genre))

    # 获取各演员数量
    if av_info[0][6] is None:
        av_stars=''
    else:
        av_stars = av_info[0][6]
    new_av_stars = av_stars.split(',')
    del new_av_stars[-1]
    print(new_av_stars)
    num_stars = 0
    for av_star in new_av_stars:
        sql_sel = "SELECT num FROM av_stars WHERE star = '" + av_star + "'"
        cur.execute(sql_sel)
        num_star = cur.fetchall()[0][0]
        num_stars = num_stars + num_star
        print(av_star + '数量：' + str(num_star))


     # 获取播放时间
    if av_info[0][7] is None or av_info[0][7]=='0':
        num_days = 240
    else:
        last_str = av_info[0][7]
        last = datetime.datetime.strptime(last_str, '%Y-%m-%d %H:%M:%S')
        now_str = datetime.datetime.now()
        now_str = now_str.strftime('%Y-%m-%d %H:%M:%S')
        now = datetime.datetime.strptime(now_str, '%Y-%m-%d %H:%M:%S')
        date_num_days = now-last
        num_days=date_num_days.days

    # 更新得分 目前权重都为1
    score = num_director + num_faxing + num_zhizuo + num_xilie + num_stars + num_genres - (240-num_days)
    sql_update = "UPDATE av_record SET score ='" + str(score) + "' WHERE ID ='" + fanhao + "'"
    cur.execute(sql_update)
    conn.commit()
    print('该番号更新得分为' + str(score))

    cur.close()
    conn.close()

def updateAllScore():
    # 通过所有番号
    def mdb_conn(password=""):
        # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

        str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        conn = pypyodbc.win_connect_mdb(str)
        return conn

    conn = mdb_conn()
    cur = conn.cursor()

    # 获取所有番号
    sql_sel = "SELECT ID FROM av_record"
    cur.execute(sql_sel)
    fanhaos = cur.fetchall()
    cur.close()
    conn.close()
    for i in range(len(fanhaos)):
        fanhao = fanhaos[i][0]
        print('-------------------正在执行:' +fanhao+'-----------------------')
        updateScore(fanhao)


updateAllScore()
print('---------------数据全部更新完毕，开始输出html-----------------')
output_html()
os.system('pause')





























