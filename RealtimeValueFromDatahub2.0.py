#!/usr/bin/env python
#coding=utf-8
# %% [markdown]
# # 1.利用DDE获取DataHub数据
# ### 1.1.需抓取的tagname清单以TXT文档形式维护filename = "./tagname.txt"
# ### 1.2.获取的value抓换为数组并构造成JSON
# 

# %%
from win32com.client import Dispatch
from datetime import datetime 
import time
import json
import requests
import schedule
import psutil
import os
from apscheduler.schedulers.blocking import BlockingScheduler
import pythoncom

# %% [markdown]
# ### 1.1.打开tagname文件并保存为数组

# %%
def Opentxt(file):
    taglist = []
    with open(file,'r') as txt:
        for line in txt.readlines():
            if "Root." in line:
                taglist.append(line.replace("\n",".Value"))
                result=','.join(taglist)
    return result

# %% [markdown]
# #### 获取实时时间函数

# %%
def timefunc():
    SpotTime = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    return SpotTime

# %% [markdown]
# #### 构造JSON形式类

# %%
class Tag:  
    def __init__(self, tagname, timestamp,value):  
        self.tagname = tagname.replace(".value","")  
        self.curtime = time.strftime("%H:%M:%S", time.localtime())
        self.curdate = time.strftime("%Y-%m-%d", time.localtime())
        self.tagvalue=value

# %% [markdown]
# ### 1.2.获取数据并转换为JSON格式

# %%
def ddefunc(datahubname,topic,filename):
    pythoncom.CoInitialize()
    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = 0 #隐藏
    xlApp.Application.DisplayAlerts = 0 #禁止弹出会话
    nChan = xlApp.Application.DDEInitiate(datahubname, topic) #datahub名称
    arrname = Opentxt(filename).split(",")   #tagname
    timestamp = timefunc()    #timestamp
    ValueResult = []
    Jsonlist = []
    for i in arrname:
        repi = i.replace(".value","")
        DDEVALUE = xlApp.DDErequest(nChan, i)
        
        if not DDEVALUE[0]:
            pass
        else:    
            ValueResult.append(Tag(i.replace(".Value",""),timestamp,str(DDEVALUE[0])))
    for tag in ValueResult:
        Jsonlist.append(eval(json.dumps(tag.__dict__)))
    return Jsonlist

# %% [markdown]
# # 2.webapi函数
# ### 2.1.返回信息日志
# ### 2.2.API函数构造
# %% [markdown]
# #### 日志文件生成及放回信息状态
# %% [markdown]
# def log(logtxt):
#     logtm = time.strftime("%Y%m%d""%H%M%S", time.localtime())
#     file = r".\log" + logtm + '.log'
#     with open(file,'w') as f:
#         f.write(logtxt)
#     # print("日志写入成功！")

# %%
def log(logtxt):
    logyear = time.strftime("%Y", time.localtime())
    logmoth = time.strftime("%m", time.localtime())
    logday = time.strftime("%d", time.localtime())
    logtm = time.strftime("%H%M%S", time.localtime())
    file_path = '{}/{}/{}'.format(logyear,logmoth,logday)  # 此处也可以使用nowTime或hourTime，看你想使用哪种格式了。
    print(file_path)
    # 判断文件夹是否已存在
    isExists = os.path.exists(file_path)
    if not isExists:
        os.makedirs(file_path )        
    file = os.getcwd()+"\\" + logyear +"\\" + logmoth +"\\" + logday +"\\"+ logtm + '.log'
#     file = os.getcwd()+"\\" + logyear +"\\"+ logtm + '.log'
    print(file)
    with open(file,'w') as f:
        f.write(logtxt)
    # print("日志写入成功！")


# %%
def webapi(Jsonlist,url):
    r_json = json.dumps(Jsonlist)
#     print(r_json)
    body = {"action":"create","info":r_json}
    result = requests.post(url,body)
    logtxt=result.text
    log(logtxt)

# %% [markdown]
# # 3.主函数构造
# ### 传入名称 文档地址 api接口地址

# %%
def main1():
    datahubname = "JSPIMSTEST"
    topic = "JSPIMSTEST"
    filename = r".\tagname.txt"
    RTvalue = ddefunc(datahubname,topic,filename)
#     print(RTvalue)
#     url = "www.baidu.com"
    url = "https://safety.ccpgp.com.cn/wuwei/API/PIMStoWWYTZDWXY_API.jsp"
    webapi(RTvalue,url)
    print("*****"*5)
    print(datetime.now().strftime("%Y/%m/%d %H:%M:%S"))
    print ('当前进程的内存使用：',psutil.Process(os.getpid()).memory_info().rss)
    print ('当前进程的内存使用：%.4f GB' % (psutil.Process(os.getpid()).memory_info().rss / 1024 / 1024 / 1024) )
    print("*****"*5)

# %% [markdown]
# # 4.定时任务

# %%
# if __name__ == "__main__":
#     schedule.every(5).seconds.do(main) 
#     while True:
#         schedule.run_pending()   # 运行所有可以运行的任务
# #         time.sleep(1)


# %%
def ramused():
    curr_pid = os.getpid()
    currApp = psutil.Process(curr_pid)
    currApp_ramused = currApp.memory_full_info()
    usedram = currApp_ramused.uss / 1024. / 1024. / 1024.
    return usedram


def job1():
    main1()


def main():
    scheduler = BlockingScheduler()
    scheduler.add_job(job1, 'interval', seconds=5)
    try:
        scheduler.start()
    except Exception as err:
        print(err)
        return


if __name__ == '__main__':
    main()


