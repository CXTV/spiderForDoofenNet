#include libs
import urllib.request as ur
import urllib.parse as up
import tkinter as tk
import base64
import hashlib

#全局变量
conf = {
    "username":"",
    "password":"",
    "cookie":"",
    "class":""
    }


#多分网的密码上传加密算法
def pwdEncoder(pwd):
    tmpPwd = ["","",""]
    md5Tmp = ["","",""]
    
    oldPwd = base64.b64encode(bytes(pwd,"utf-8")).decode()
    
    tmpPwd[0] = oldPwd[0]
    tmpPwd[1] = oldPwd[1]
    tmpPwd[2] = oldPwd[2:]

    md5Pwd = hashlib.md5(pwd.encode("utf-8")).hexdigest()

    md5Tmp[0] = md5Pwd[1]
    md5Tmp[1] = md5Pwd[3]
    md5Tmp[2] = md5Pwd[4:7]

    newPwd = str(tmpPwd[0]) + str(md5Tmp[0]) +\
             str(tmpPwd[1]) + str(md5Tmp[1]) + \
             str(md5Tmp[2]) + str(tmpPwd[2]) + str(len(tmpPwd[2]))

    return base64.b64encode(bytes(newPwd,"utf-8")).decode()#返回Str


#定义登录方法
def logIn():
    feedBack.config(text = "登录成功!")#不知道为什么无法显示

    username = userInput.get()
    password = pwdEncoder(pwdInput.get())
    #读取输入并加密

    data_send = {"username":username,"password":password}
    data_send =str(data_send).encode("utf-8")
    #生成Body

    header_send = {
        'Host': 'www.doofen.com',
        'Connection': 'keep-alive',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Origin': 'http://www.doofen.com',
        'X-Requested-With': 'XMLHttpRequest',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36',
        'Content-Type': 'application/json;charset=UTF-8',
        'Referer': 'http://www.doofen.com/doofen/login.html?',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9'
        }#定义Head

    request = ur.Request(url = "http://www.doofen.com/doofen/sys/login", data = data_send, headers = header_send)
    response = ur.urlopen(request)
    #发送登录请求

    if response.read().decode().find("\"success\":true")!=-1 :
        logedIn()
        conf["username"] = username
        conf["password"] = password
    else:
        feedBack.config(text = "登录失败！请检查账号密码")

#登录成功
def logedIn():
    feedBack.config(text = "登录成功!")
    userInput.grid_remove()
    pwdInput.grid_remove()
    


#生成登录窗体
root = tk.Tk()
root.title = "多分网整理工具"

tk.Label(root,text = "登入您的多分网账号:").grid(row = 0)
tk.Label(root,text = "手机号:").grid(row = 1,sticky = "W")
tk.Label(root,text = "密码:").grid(row = 2,sticky = "W")

feedBack = tk.Label(root,text = "",fg = "Blue")
feedBack.grid(row = 4)

userInput = tk.Entry(root)
pwdInput = tk.Entry(root)
userInput.grid(row = 1,column = 1)
pwdInput.grid(row = 2,column = 1)

tk.Button(root,text = "登录",command = logIn).grid(row = 3)

root.mainloop()
