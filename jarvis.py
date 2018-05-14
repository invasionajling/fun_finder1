#coding=utf8
import itchat

name_gao = '@e5ae51fca683ca34ea7dd9f90ce4e8c4'


@itchat.msg_register(itchat.content.TEXT)
def reply(msg):
    return msg['Text']

itchat.auto_login(hotReload=True)
itchat.run()