#
# 导入库
#
import urllib.request as ur
import tkinter as tk
import tkinter.messagebox as tm
from tkinter import ttk
from base64 import b64encode
from hashlib import md5
import json
from traceback import format_exc
import tkinter.filedialog as tf
import webbrowser
# 标准库

import win32com.client  # Use "pip install pypiwin32" to get it

#
# 全局变量
#
thisVision = 2335
conf = {
    "username":"",
    "password":"",
    "remember":0,
    "jsessionId":"",
    "classId":"",
    "classes":[],
    "students":{},
    "subjects":[],
    "examId":"",
    "exams":[],
    "outputPath":"",
    "inputFile":"",
    "crashed":0,
    "showExcel":0
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
    }  # 定义Head

#
# 读取配置数据
#
def readconf():
    global conf

    file = open(r"spiderDoofen.conf", "r")
    content = file.read()
    conf = json.loads(content.replace("'","\""))
    file.close()

    if conf["crashed"]:
        if tm.askyesno("继续","程序上次运行时意外崩溃。是否从断点继续生成？"):
            userInput.insert(0,conf["username"])
            pwdInput.insert(0,conf["password"])
            remember.set(1)
            log(str(conf))
            login()
            getcontent()
    # 断点继续

    if conf["remember"]:
        userInput.insert(0,conf["username"])
        pwdInput.insert(0,conf["password"])
        remember.set(1)
    # 记住密码


#
# 检查更新
#
def checkupdate():
    log("Start Checking For Updates...")
    response = ur.urlopen("http://p7zz4jl0d.bkt.clouddn.com/update")
    lateVision = json.loads(response.read().decode())
    log("Present Vision: " + str(thisVision) + " ,Lastest Vision: " + str(lateVision))
    # 发包解包

    if lateVision > thisVision:  # 询问下载
        if tm.askyesno("更新", "有可用的更新，是否现在下载？"):
            webbrowser.open("http://p7zz4jl0d.bkt.clouddn.com/" + str(lateVision) + ".exe")

#
# 多分网的密码上传加密算法<<http://www.doofen.com/doofen/assets/scripts/login/login.js
#
def pwdencoder(pwd):
    tmpPwd = ["","",""]
    md5Tmp = ["","",""]
    
    oldPwd = b64encode(bytes(pwd,"utf-8")).decode()
    
    tmpPwd[0] = oldPwd[0]
    tmpPwd[1] = oldPwd[1]
    tmpPwd[2] = oldPwd[2:]

    md5Pwd = md5(pwd.encode("utf-8")).hexdigest()

    md5Tmp[0] = md5Pwd[1]
    md5Tmp[1] = md5Pwd[3]
    md5Tmp[2] = md5Pwd[4:7]

    newPwd = str(tmpPwd[0]) + str(md5Tmp[0]) +\
             str(tmpPwd[1]) + str(md5Tmp[1]) +\
             str(md5Tmp[2]) + str(tmpPwd[2]) + str(len(tmpPwd[2]))

    return b64encode(bytes(newPwd,"utf-8")).decode()  # 返回Str

#
# 定义登录方法
#
def login():
    log(str(conf))
    username = userInput.get()
    password = pwdencoder(pwdInput.get())
    # 读取输入并加密

    data_send = {"username":username,"password":password}
    data_send = str(data_send).encode("utf-8")
    # 生成Body

    request = ur.Request(
        url = "http://www.doofen.com/doofen/sys/login",
        data = data_send,
        headers = header_send)
    response = ur.urlopen(request)
    # 发送登录请求

    tmpCoo = str(response.info())
    conf["jsessionId"] = tmpCoo[tmpCoo.find("JSESSIONID=")+11:tmpCoo.find("; Path")]
    # 获取Session通信码
    response = response.read().decode()  # package=>string
    
    if response.find("\"success\":true") != -1 :
        conf["classes"] = json.loads(response)["data"]["actInfo"]["tchRole"]
        conf["username"] = username
        conf["password"] = pwdInput.get()
        logedin()
    else:
        feedBack.config(text = "登录失败！请检查账号密码")
    return


#
# 登录成功
#
def logedin():
    log("Logged In As " + conf["username"] + " ,JSESSIONID=" + conf["jsessionId"])
    conf["remember"] = remember.get()
    
    # 加载班级列表
    tmpClass = []
    for classNo in conf["classes"]:
        tmpClass.append(classNo.split("|")[0])
    classChoose["values"] = tuple(tmpClass)
    classChoose.current(0)

    for sub in ["语文","数学","英语","物理","化学","历史","地理","政治","生物"]:
        subChoose.insert(tk.END,sub)  # 加载学科列表

    # 应用新界面
    frmLog.grid_remove()
    frmMain.grid(column = 0)

    if conf["crashed"]:
        return
    else:
        confwrite()
        classload()  # 载入班级数据
        return
    
#
# 加载班级内容
#
def classload(*arg):
    log("Start Loading Class \"" + className.get() + "\".")

    conf["students"].clear()
    conf["exams"].clear()
    examChoose["values"] = ("",)
    examChoose.current(0)  # 清空学生、考试数据

    # 获取学生id
    conf["classId"] = "851001" + className.get()[className.get().find("1"):]  # 获得链接
    stuUrl = "http://www.doofen.com/doofen/851001/cls/" + conf["classId"] + "/stu/list"
    response = ur.urlopen(stuUrl)  # 发送数据包

    resRead = response.read()
    log("Students List Package Received Of " + str(len(resRead)) + " Bytes.")

    for person in json.loads(resRead.decode()):
        conf["students"][str(person["stuId"])] = {"name":person["stuName"]}
        # 写入到conf
    
    if conf["students"] == {}:  # 处理班级错误
        log("Students List Loading Error With An Empty List.",1)
        tm.showwarning("警告", "班级\" " + className.get() + " \"的学生数据为空。")

    log("Students List Loaded with " + str(len(conf["students"])) + " students.")

    # 获取考试名称
    try:
        stuUrl = "http://www.doofen.com/doofen/851001/examsit/student/studentRptData?s="\
                + list(conf["students"].keys())[0] + "&p=0&r=3"
        response = ur.urlopen(stuUrl)  # 发送数据包
        
        resRead = response.read()
        log("Exams List Package Received.")

        for exam in json.loads(resRead.decode()):  # 遍历试卷列表
            if conf["exams"].count(str(exam["examId"])) == 0:  # 寻找新的id
                conf["exams"].append(str(exam["examId"]))  # 插入新的考试id

        tmpExams = conf["exams"]
        examChoose["values"] = tuple(tmpExams)
        examChoose.current(0)
        # 写入试卷列表

        log("Exams List Loaded.\n" + str(conf["exams"]))

    except IndexError:  # 处理班级错误
        tm.showwarning("警告", "班级\" " + className.get() + " \"没有学生数据，不能读取考试列表。")
        log("Exams List Loading Error With No Student Found.",1)

childrenDict = {}
def getchildren(fatherDict):  # 遍历字典中所有键-值
    for i in range(len(fatherDict)):
        childKey = list(fatherDict.keys())[i]
        childDict = fatherDict[childKey]
        if isinstance(childDict,dict):
            getchildren(childDict)
        else:
            childrenDict[childKey] = childDict

#
# 选择模板文件
#
def selectfile():
    filename = tf.askopenfilename(filetypes=[("Excel 文件","*.xl;*.xls;*.xlm;*.xlsx;*.xlsm;*.xlsb")])  
    inputFileName.set(filename)

#
# 选择输出目录
#
def selectpath():
    filepath = tf.askdirectory()
    outputPathName.set(filepath)

#
# 主任务——下载管理数据
#
def getcontent():
    global conf
    global header_send

    if not conf["crashed"]:
        log("Start Getting Contents. Checking Values...")

        # 检查并载入配置数据
        log("Students List Loaded With A Length Of " + str(len(conf["students"])) + ".")

        conf["subjects"].clear()
        for sub in subChoose.curselection():
            conf["subjects"].append({"id":str(sub + 1),"name":subChoose.get(sub)})
        log("Subjects List Loaded As " + str(conf["subjects"]) + ".")

        conf["examId"] = examName.get()
        log("ExamId Loaded As " + str(conf["examId"]) + ".")

        conf["inputFile"] = inputFileName.get()
        log("Input File Loaded As " + conf["inputFile"])

        conf["outputPath"] = outputPathName.get()
        log("Output Path Loaded As " + conf["outputPath"])

        conf["showExcel"] = showExcel.get()
        if conf["showExcel"]:log("Show Excel.")
        else:log("Hide Excel.")

        if (conf["students"] == {} or 
            conf["subjects"] == [] or 
            conf["examId"] == "" or 
            conf["inputFile"] == "" or 
            conf["outputPath"] == ""):
                log("Incomplete Arguments.Stop Running.",1)
                tm.showwarning("警告", "设置不完整或值无效。")
                return
            # 处理数据缺失


        log("Checking Done.\n\tStart Getting Content...Please Wait...")
        tm.showinfo("提示", "抓取可能需要十余分钟时间，在程序提示完成前，请不要点击或关闭本程序。\n" +
                    "如果系统弹出了\"无响应\"提示框，请忽视。（除非已经真的很久没动静了）")

        # 考试全科大表(获取班级排名)
        url = "http://www.doofen.com/doofen/851001/rpt100/1001?clsId=" + \
                    conf["classId"] + "&examId=" + conf["examId"]
        # 数据包地址

        request = ur.Request(url = url, headers = header_send)
        response = ur.urlopen(request).read().decode()
        log("Sheet Of Full Subjects Received In " + str(len(response)) + " Bytes.")
    
        objTmp = json.loads(response)[0]["stuScore"]
    
        for i in range(len(objTmp)):
            stuTmp = conf["students"][str(objTmp[i]["stuId"])]
            stuTmp["mainScore"] = objTmp[i]["stuMixScore"]
            stuTmp["mainGradeRank"] = objTmp[i]["stuMixRank"]
            stuTmp["mainClassRank"] = i + 1
    # 数据检查结束
    else:
        log("Continue To Get Contents...")
        log("Loaded " + str(len(conf["students"])) + " Students.")

    # 唤起Excel
    try:
        app = win32com.client.Dispatch('Excel.Application')
        app.Visible = conf["showExcel"]
        app.DisplayAlerts=0
    except:
        log("Failed To Call Excel Process.Stop Running.\n" + format_exc(),2)
        tm.showerror("错误", "无法唤起 \"Excel.Application\",抓取停止\n" +
                     "可能是由于未安装Excel。如果你确信这是个bug，请将窗口下方完整的日志发送到源码页面。")
        return
    else:
        log("Started Excel.Application.")

    conf["crashed"] = 1  # 记录断点

    for stuId in conf["students"].copy():  # 遍历学生
        student = conf["students"][stuId]
        
        # 打开模板文档
        try:template = app.Workbooks.Open(conf["inputFile"])
        except:
            log("Failed To Open File: \"" + conf["inputFile"] + "\".Stop Running.\n" + format_exc(),2)
            tm.showerror("错误", "无法打开选择的模板文档，抓取停止\n请尝试手动重启Excel")
            return
        #else:log("Opened " + conf["inputFile"])

        for keys in student:
            try:app.ActiveSheet.Cells.Replace(keys, student[keys],1)
            except:log("While Replacing,\n" + format_exc(),2)

        table = str.maketrans("/", "\\")
        writePath = conf["outputPath"].translate(table) + "\\" + student["name"] + ".xlsx"

        try:app.ActiveWorkBook.SaveAs(writePath)  # 另存为
        except:
            log("Failed To Write Out File: \"" + writePath + "\".Stop Running.\n" + format_exc(),2)
            tm.showerror("错误", "无法保存文档，抓取停止\n请确认目录存在")
            return
        else:log("Wrote Out File " + writePath)

        # 遍历科目
        for subject in conf["subjects"]:
            header_send["Cookie"] = "JSESSIONID=" + conf["jsessionId"]
            # 数据包头

            url = "http://www.doofen.com/doofen/851001/report/subjectDatas?rId=" + \
                subject["id"] + "_" + conf["examId"] + "_" + stuId
            # 数据包地址

            request = ur.Request(url = url, headers = header_send)
            response = ur.urlopen(request).read().decode()
            dataObj = json.loads(response)  # 解析数据包

            childrenDict.clear()
            getchildren(dataObj)
            # 遍历字典对象

            # 整理错题
            wrongInfo = childrenDict["wrongItemStatInfo"]
            childrenDict["wrongItemStatInfo"] = ""
            # 获取错题单行
            wrongStart = app.ActiveSheet.Cells.Find(What = "wrongStart" + str(subject["id"]), LookAt=1 )
            wrongEnd = app.ActiveSheet.Cells.Find(What = "wrongEnd" + str(subject["id"]), LookAt = 1 )
            
            for each in wrongInfo:
                app.ActiveSheet.Rows(wrongStart.Row).Insert(2,2)
                app.ActiveSheet.Range(wrongStart,wrongEnd).Copy()
                app.ActiveSheet.Range(wrongStart,wrongEnd).Offset(0,1).Select()
                
                app.ActiveSheet.Paste()

                app.Selection.Replace("wrongStart" + str(subject["id"]), " ",1)
                app.Selection.Replace("wrongEnd" + str(subject["id"]), " ",1)

                for i in range(len(each)):
                    childKey = list(each.keys())[i]
                    childVal = each[childKey]
                    app.Selection.Replace(childKey + str(subject["id"]), str(childVal),1)

            app.ActiveSheet.Rows(wrongStart.Row).Delete()

            for i in range(len(childrenDict)):
                childKey = list(childrenDict.keys())[i]
                childVal = childrenDict[childKey]
                try:app.ActiveSheet.Cells.Replace(childKey + str(subject["id"]), str(childVal),1)
                except:log("While Replacing,\n" + format_exc(),2)

            log("Subject " + subject["name"] +  " of " +
               student["name"] + " Loaded With " + str(len(childrenDict)) + " Items.")

            del conf["students"][stuId]
            confwrite()
            # 写入配置文件

        app.ActiveWorkbook.Save()
        app.ActiveWorkbook.Close()

    conf["crashed"] = 0
    confwrite()
    # 写入配置文件
    tm.showinfo("提示", "抓取完成！")

#
# 配置文件写入
#
def confwrite():
    file = open(r"spiderDoofen.conf","w")
    file.write(str(conf))
    file.close()

#
# 日志生成函数
#
def log(string,type = 0):
    typeList = ["\n[Info] ","\n[Warn] ","\n[Error] ",""]
    runLog.config(state = tk.NORMAL)
    runLog.insert(tk.END, typeList[type] + string)
    runLog.config(state = tk.DISABLED)



#
# 生成主窗体
#
root = tk.Tk()
root.title("多分网整理工具")

# 绘制登录容器
frmLog = tk.Frame(root)

tk.Label(frmLog,text = "登入您的多分网账号:").grid(row = 0, columnspan = 2)
tk.Label(frmLog,text = "手机号:").grid(row = 1,sticky = "W")
tk.Label(frmLog,text = "密码:").grid(row = 2,sticky = "W")  # 初始化文本对象

feedBack = tk.Label(frmLog,text = "",fg = "Blue")
feedBack.grid(row = 4, columnspan = 2)  # 反馈文本

userInput = tk.Entry(frmLog)
pwdInput = tk.Entry(frmLog,show = "*")
userInput.grid(row = 1,column = 1)
pwdInput.grid(row = 2,column = 1)  # 初始化输入框*2

remember = tk.IntVar()
remCheck = tk.Checkbutton(frmLog, text = "记住我", variable = remember)
remCheck.grid(row = 3, column = 0, columnspan = 2)

tk.Button(frmLog,text = "登录",command = login).grid(row = 4, columnspan = 2)  # 登录按钮


# 绘制主界面容器
#
# column
#        0         1           2          3
#row┌────────────────────┐
# 0 │选择班级┌───┬┐┌───┬┐选择模板│
#   │        └───┴┘└───┴┘        │
# 1 │选择考试┌───┬┐┌───┬┐选择路径│
#   │        └───┴┘└───┴┘        │
#   │        ┌────┐                    │
#   │        │xxx     │     ┌────┐   │
# 2 │选择科目│xxx     │     │开始抓取│   │
#   │        │xxx     │     └────┘   │
#   │        │xxx     │                    │
# 3 │        └────┘     ☑显示Excel    │
#   │ ┌─────────────────┐ │
#   │ │log:                              │ │
# 4 │ │    xxx                           │ │
#   │ └─────────────────┘ │
#   └────────────────────┘
#

frmMain = tk.Frame(root)

tk.Label(frmMain,text = "选择班级").grid(column = 0,row =0)
# 初始化'选择班级'Combobox对象
className = tk.StringVar()
classChoose = ttk.Combobox(frmMain,textvariable = className)
classChoose.grid(column = 1,row = 0)
classChoose["state"] = "readonly"
classChoose.bind("<<ComboboxSelected>>",classload)

tk.Label(frmMain,text = "选择考试").grid(column = 0,row = 1)
# 初始化'选择考试'Combobox对象
examName = tk.StringVar()
examChoose = ttk.Combobox(frmMain,textvariable = examName)
examChoose.grid(column = 1,row = 1)
examChoose["state"] = "readonly"

tk.Label(frmMain,text = "选择科目").grid(column = 0,row = 2,rowspan = 2)
# 初始化'选择科目'Listbox对象
subChoose = tk.Listbox(frmMain,selectmode = tk.MULTIPLE)
subChoose.grid(column = 1,row = 2, rowspan = 2)

# 初始化'模板文件'Entry对象
inputFileName = tk.StringVar()
inputFile = tk.Entry(frmMain, textvariable = inputFileName, state = "readonly")
inputFile.grid(column = 2, row = 0, padx = 3)
tk.Button(frmMain, text = "选择模板", command = selectfile).grid(row = 0, column = 3)

# 初始化'输出目录'Entry对象
outputPathName = tk.StringVar()
outputPath = tk.Entry(frmMain, textvariable = outputPathName, state = "readonly")
outputPath.grid(column = 2, row = 1, padx = 3)
tk.Button(frmMain, text = "选择输出路径", command = selectpath).grid(row = 1, column = 3)

# 显示Excel选框
showExcel = tk.IntVar()
showExcelCheck = tk.Checkbutton(frmMain, text = "显示Excel", variable = showExcel)
showExcelCheck.grid(column = 2, row = 3, columnspan = 2)

# 初始化日志Text对象
runLog = tk.Text(frmMain,height = 10,width = 62, state = tk.DISABLED)
runLog.grid(column = 0,row = 4, columnspan = 4)

startButton = tk.Button(frmMain,text = "开始抓取", command = getcontent, width = 15, height = 4)
startButton.grid(row = 2,column = 2, columnspan = 2)


frmLog.grid(column = 0)  # 显示登录容器


log("-" *50 + "\n" + "-" *50 +
    "\n欢迎使用多分网数据整理工具！\n" +
    "源码地址: https://github.com/CaptainMorch/spiderForDoofenNet \n" +
    "软件使用说明: http://t.cn/Ru0oFnA \n" +
    "作者:Captain_Morch \nFor My Best Class23 :)\n" +
    "-" *50 + "\n" + "-" *50 + "\n广告:\n=>[作者的其他项目]Minecraft还原一中 http://www.mcgyyz.cn\n" +
    "=>此广告位长期不招租\n" +
    "-" *50 + "\n" + "-" *50 ,3)

checkupdate()  # 检查更新
readconf()  # 读取配置文件

root.mainloop()
