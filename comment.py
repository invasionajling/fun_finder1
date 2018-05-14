import pypyodbc
import datetime
import os

def updateScore(fanhao,level):
    def mdb_conn(password=""):
        # 功能：创建数据库连接 :param db_name: 数据库名称 :param db_name: 数据库密码，默认为空 :return: 返回数据库连接

        str = 'driver={Microsoft Access Driver (*.mdb)};PWD' + password + ";DBQ=fun.mdb"
        conn = pypyodbc.win_connect_mdb(str)
        return conn

    conn = mdb_conn()
    cur = conn.cursor()

    # 获取各种信息
    sql_sel = "SELECT * FROM av_record WHERE ID = '" + fanhao + "'"
    cur.execute(sql_sel)
    av_info = cur.fetchall()
    print(av_info)

    # 更新播放时间
    now = datetime.datetime.now()
    now = now.strftime('%Y-%m-%d %H:%M:%S')
    sql_sel = "UPDATE av_record SET last='" + str(now) + "' WHERE ID = '" + fanhao + "'"
    cur.execute(sql_sel)
    conn.commit()
    print('播放时间更新：')


    # 获取导演数量
    if av_info[0][1] == '':
        num_director = 0
    else:
        av_director = av_info[0][1]
        sql_sel = "UPDATE av_director SET num=num+'"+level+"' WHERE director = '" + av_director + "'"
        cur.execute(sql_sel)
        conn.commit()
        print('导演数量更新：')

    # 获取发行商数量
    if av_info[0][2] == '':
        num_faxing = 0
    else:
        av_faxing = av_info[0][2]
        sql_sel = "UPDATE av_faxing SET num=num+'" + level + "' WHERE faxing = '" + av_faxing + "'"
        cur.execute(sql_sel)
        conn.commit()
        print('发行商数量更新：')

    # 获取制作商数量
    if av_info[0][3] == '':
        num_zhizuo = 0
    else:
        av_zhizuo = av_info[0][3]
        sql_sel = "UPDATE av_zhizuo SET num=num+'" + level + "' WHERE zhizuo = '" + av_zhizuo + "'"
        cur.execute(sql_sel)
        conn.commit()
        print('制作商数量更新：')

    # 获取系列数量
    if av_info[0][4] == '':
        num_xilie = 0
    else:
        av_xilie = av_info[0][4]
        sql_sel = "UPDATE av_xilie SET num=num+'" + level + "' WHERE xilie = '" + av_xilie + "'"
        cur.execute(sql_sel)
        conn.commit()
        print('系列数量更新：')


    # 获取各类别数量
    if av_info[0][5] is None:
        av_genres = ''
    else:
        av_genres = av_info[0][5]
    new_av_genres = av_genres.split(',')
    del new_av_genres[-1]
    print(new_av_genres)
    for av_genre in new_av_genres:
        sql_sel = "UPDATE av_genres SET num=num+'" + level + "' WHERE genre = '" + av_genre + "'"
        cur.execute(sql_sel)
        conn.commit()
        print(av_genre + '类别数量更新：')

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
        sql_sel = "UPDATE  av_stars SET num=num+'" + level + "' WHERE star = '" + av_star + "'"
        cur.execute(sql_sel)
        conn.commit()
        print(av_star + '数量更新：')

    cur.close()
    conn.close()

fanhao = input("番号: ")
level = input("评级1-5:")

print(fanhao,level)
updateScore(fanhao,level)
os.system('pause')