
#include libs
import urllib.request as ur
import urllib.parse as up
import tkinter as tk
from tkinter import ttk
import base64
import hashlib
import json

#全局变量
subjectId = "1"
stuId = "8510011600000686"
examId = "2001523842908308"
conf = {
    "username":"",
    "password":"",
    "jsessionId":"",
    "classId":"",
    "classes":""
    }
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

    request = ur.Request(url = "http://www.doofen.com/doofen/sys/login", data = data_send, headers = header_send)
    response = ur.urlopen(request)
    #发送登录请求

    tmpCoo = str(response.info())
    conf["jsessionId"] = tmpCoo[tmpCoo.find("JSESSIONID=")+11:tmpCoo.find("; Path")]
    #获取Session通信码

    response = response.read().decode()#package=>string
    
    if response.find("\"success\":true")!=-1 :
        conf["classes"] = json.loads(response)["data"]["actInfo"]["tchRole"]
        conf["username"] = username
        conf["password"] = password
        logedIn()
    else:
        feedBack.config(text = "登录失败！请检查账号密码")


#登录成功
def logedIn():
    feedBack.config(text = "登录成功!")
    
    #绘制主界面容器
    frmMain = tk.Frame(root)

    tk.Label(frmMain,text = "选择班级:").grid(column = 0,row =0)
    className = tk.StringVar()
    classChoose = ttk.Combobox(frmMain,textvariable = className)
    classChoose.grid(column = 1,row = 0)
    classChoose["state"] = "readonly"
    classChoose.bind("<<ComboboxSelected>>", lambda _ : getStuid(className.get()))#这个参数传递折腾死我了

    tk.Label(frmMain,text = "选择考试:").grid(column = 0,row = 1)
    testName = tk.StringVar()
    testChoose = ttk.Combobox(frmMain,textvariable = testName)
    testChoose.grid(column = 1,row = 1)
    testChoose["state"] = "readonly"

    tk.Label(frmMain,text = "选择科目:").grid(column = 0,row = 2)
    subChoose = tk.Listbox(frmMain,selectmode = tk.MULTIPLE)
    subChoose.grid(column = 1,row = 2)

    tk.Button(frmMain,text = "开始抓取", command = getContent).grid(row = 3,column = 2)


    #加载班级数据
    tmpClass = []
    for classNo in conf["classes"]:
        tmpClass.append(classNo.split("|")[0])
    classChoose["values"] = tuple(tmpClass)
    classChoose.current(0)

    for sub in ["语文","数学","英语","物理","化学","生物"]:  
        subChoose.insert(tk.END,sub)

    #应用新界面
    frmLog.grid_remove()
    frmMain.grid(column = 0)
    
#获取stuid
def getStuid(className):
    #print(className)
    conf["classId"] = "851001" + className[className.find("1"):]#获得链接

    stuUrl = "http://www.doofen.com/doofen/851001/cls/" + conf["classId"] + "/stu/list"
    response = ur.urlopen(stuUrl)#发送数据包

    students = []
    for person in json.loads(response.read().decode()):
        tmp = {"id":person["stuId"],"name":person["stuName"]}
        students.append(tmp)
    #移植对象
    #print(students)


#获取内容
def getContent():
    for student in students:
        file = open("/test/" + student["name"] + ".txt","a+")

        for subjectId in subjects:
            header_send["Cookie"] = "JSESSIONID=" + conf["jsessionId"]
            #数据包头

            url = "http://www.doofen.com/doofen/851001/report/subjectDatas?rId=" + \
                subjectId + "_" + examId + "_" + stuId
            #数据包地址

            request = ur.Request(url = url, headers = header_send)
            response = ur.urlopen(request).read().decode()
            
            dataObj = json.loads(response)


#生成主窗体
root = tk.Tk()
root.title("多分网整理工具")

#绘制登录容器
frmLog = tk.Frame(root,width = 50,height = 30)
frmLog.grid(column = 0)

tk.Label(frmLog,text = "登入您的多分网账号:").grid(row = 0)
tk.Label(frmLog,text = "手机号:").grid(row = 1,sticky = "W")
tk.Label(frmLog,text = "密码:").grid(row = 2,sticky = "W")#3个文本

feedBack = tk.Label(frmLog,text = "",fg = "Blue")
feedBack.grid(row = 4)#反馈文本

userInput = tk.Entry(frmLog)
pwdInput = tk.Entry(frmLog,show = "*")
userInput.grid(row = 1,column = 1)
pwdInput.grid(row = 2,column = 1)#输入框*2

tk.Button(frmLog,text = "登录",command = logIn).grid(row = 3)

root.mainloop()
