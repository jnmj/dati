from selenium import webdriver
from pyhooked import Hook, KeyboardEvent
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains

import traceback
import ctypes
import os
import sys
import time
import win32gui
import win32com.client
import win32con
from PIL import Image, ImageGrab
from aip import AipOcr
from selenium.webdriver.common.keys import Keys

'''
优酷,百万英雄（手机）   0
UC（手机）              1
一直播（手机）          2   
好看视频（模拟器）      3
贴吧（模拟器）          4
NOW视频（模拟器）       5
冲顶大会（模拟器）      6
'''

# 配置
data_directory = 'screenshots'

Rect = [(0, 300, 1080, 1350), (0, 400, 1080, 1300), (0, 450, 1080, 1400), (0, 170, 480, 520), (0, 220, 480, 600), (0, 250, 480, 600), (0, 170, 480, 490)]

keyVM = {'F2': '', 'F4' : 'SS1', 'F6' : 'SS2' }

keyRect = {'F2' : Rect[0], 'F4' : Rect[0], 'F6' : Rect[3]}

search_engine = 'http://www.baidu.com'

# ocr_engine = 'baidu'
ocr_engine = 'baidu'

### baidu orc
app_id = '10725701'
app_key = 'j47XknezW95aFRkPDPfGMGGv'
app_secret = 'mf2YXzpDg0wQuSGomnuNpSo1Atqc1Uwl'

### 0 表示普通识别
### 1 表示精确识别
api_version = 0


### 0 不搜索答案
### 1 搜索答案
searchAns = 0

class RECT(ctypes.Structure):
    _fields_ = [('left', ctypes.c_long),
            ('top', ctypes.c_long),
            ('right', ctypes.c_long),
            ('bottom', ctypes.c_long)]
    def __str__(self):
        return str((self.left, self.top, self.right, self.bottom))

def get_file_content(path):
    with open(path, 'rb') as fp:
        return fp.read()
        
#截屏和裁剪
def getCutImage(key):
    if keyVM[key]=='':
        os.system("adb shell /system/bin/screencap -p /sdcard/screenshot.png")
        fullPicPath = os.path.join(data_directory, time.strftime('%m_%d_%H_%M_%S',time.localtime(time.time()))+'_full.png')
        os.system("adb pull /sdcard/screenshot.png "+fullPicPath)
        fullImage = Image.open(fullPicPath)
        #fullImage.show()
        cutImage = fullImage.crop(keyRect[key])
        cutPicPath = os.path.join(data_directory, time.strftime('%m_%d_%H_%M_%S',time.localtime(time.time()))+'_cut.png')
        cutImage.save(cutPicPath)
        return get_file_content(cutPicPath)
    else:
        '''
        vmName = keyVM[key]
        window = win32gui.FindWindow(None, vmName)
        win32gui.ShowWindow(window, win32con.SW_RESTORE)
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(window)
        time.sleep(1)
        fullRect = win32gui.GetWindowRect(window)
        fullImage = ImageGrab.grab(fullRect)
        #fullImage.show()
        fullImage.save(os.path.join(data_directory, time.strftime('%m_%d_%H_%M_%S',time.localtime(time.time()))+'_full.png'))
        '''
        fullImage = Image.open('./screenshots/01_28_22_11_43_full.png')
        cutImage = fullImage.crop(keyRect[key])
        cutPicPath = os.path.join(data_directory,time.strftime('%m_%d_%H_%M_%S', time.localtime(time.time())) + '_cut.png')
        cutImage.save(cutPicPath)
        return get_file_content(cutPicPath)
        
#识别
def recognize(cutImage):
    
    client = AipOcr(appId=app_id, apiKey=app_key, secretKey=app_secret)
    client.setConnectionTimeoutInMillis(3 * 1000)

    options = {}
    options["language_type"] = "CHN_ENG"

    if api_version == 1:
        result = client.basicAccurate(cutImage, options)
    else:
        result = client.basicGeneral(cutImage, options)

    if "error_code" in result:
        print("百度OCR识别出错，按热键重新开始")
        return []
    OCRRet = [words["words"] for words in result["words_result"]]
    print('OCR结果：')
    print(OCRRet)
    QA = ['']*4
    i=0
    for i in range(len(OCRRet)):
        if '?' in OCRRet[i] or '？' in OCRRet[i]:
            break
    QA[0] = "".join(OCRRet[:(i + 1)])
    global isComplete
    if(len(OCRRet)-1-i < 3):
        isComplete=False
    else:
        isComplete=True
        QA[1] = OCRRet[-3]
        QA[2] = OCRRet[-2]
        QA[3] = OCRRet[-1]
    #keyword = "".join([e.strip("\r\n") for e in keyword if e])
    #print(QA)
    return QA
   
#预处理识别结果
def preProcess(QA):
    global isComplete
    for char, repl in [("“", ""), ("”", ""), (" ", ""), ("\t", ""), ("《", ""), ("》", ""),('"',""),('(',""),(')',""),('（',""),('）',"")]:
        QA[0] = QA[0].replace(char, repl)
        if isComplete==True:
            QA[1] = QA[1].replace(char, repl)
            QA[2] = QA[2].replace(char, repl)
            QA[3] = QA[3].replace(char, repl)

    if r"." in QA[0]:
        QA[0] = QA[0].split(r".", 1)[1]
    tag = QA[0].find("?")
    if tag != -1:
        QA[0] = QA[0][:tag]
    tag = QA[0].find("？")
    if tag != -1:
        QA[0] = QA[0][:tag]
    #keywords = keyword.split(" ")
    #keyword = "".join([e.strip("\r\n") for e in keywords if e])
    #if len(keyword) < 1:
    #    print("没识别出来，按热键重新开始")
    #    return ''
    if isComplete==True:
        for char in [r'.', r':', r'：']:
            if char in QA[1]:
                QA[1]=QA[1].split(char, 1)[1]
            if char in QA[2]:
                QA[2]=QA[2].split(char, 1)[1]
            if char in QA[3]:
                QA[3]=QA[3].split(char, 1)[1]
    print('问题：'+ QA[0])
    if isComplete==True:
        print('A：' + QA[1])
        print('B：' + QA[2])
        print('C：' + QA[3])
    else:
        print('答案不完整')
    return QA
    
def work(key):
    
    print('开始处理')

    try:
        cutImage = getCutImage(key)
        #cutImage = get_file_content('./screenshots/01_27_13_41_26_cut.png')
    except Exception as e:
        print('获取图片异常')
        traceback.print_exc()
        print('\n按热键开始')
        return

    try:
        QA = recognize(cutImage)
    except Exception as e:
        print('识别图片异常')
        traceback.print_exc()
        print('\n按热键开始')
        return

    try:
        QA = preProcess(QA)
    except Exception as e:
        print('处理识别结果异常')
        traceback.print_exc()
        print('\n按热键开始')
        return

    #浏览器搜索
    try:
        lastTitle=browser.title
        global lastQuestion
        if lastQuestion == QA[0]:
            print('同样的问题')
            print('\n按热键开始')
            return

        lastQuestion=(QA[0])
        elems[0].clear()
        elems[0].send_keys(QA[0])
        elems[0].send_keys(Keys.RETURN)
        WebDriverWait(browser, 2, 0.1).until_not(lambda x: browser.title==lastTitle)
        WebDriverWait(browser, 2, 0.1).until(lambda x: browser.execute_script("return document.readyState")=="complete")
        global isComplete
        if isComplete == True:
            for i in range(1,4):
                js= 'var re = new RegExp('+'"'+QA[i]+'"'+',"gmi");' \
                    'var h3s = document.getElementsByTagName("h3");' \
                    'for (var i=0;i<h3s.length;i++){' \
                        'var title = h3s[i].getElementsByTagName("a")[0];' \
                        'title.innerHTML = title.innerHTML.replace(re,\'<span style="background-color:#ff0">\'+'+'"'+QA[i]+'"'+'+\'</span>\');' \
                    '}' \
                    'var r1 = document.getElementsByClassName("newTimeFactor_before_abs");' \
                    'for (var i=0;i<r1.length;i++){' \
                        'r1[i].innerHTML="";' \
                    '}' \
                    'var r2 = document.getElementsByClassName("f13");' \
                    'for (var i=0;i<r2.length;i++){' \
                        'r2[i].innerHTML="";' \
                    '}' \
                    'var abstracts = document.getElementsByClassName("c-abstract");' \
                    'for (var i=0;i<abstracts.length;i++){' \
                        'abstracts[i].innerHTML = abstracts[i].innerHTML.replace(re,\'<span style="background-color:#ff0">\'+'+'"'+QA[i]+'"'+'+\'</span>\');' \
                    '}'
                #print(js)
                browser.execute_script(js)

        if isComplete==True and searchAns==1:
            for j in range(1,4):
                browser.switch_to.window(handles[j])
                elems[j].clear()
                elems[j].send_keys(QA[j])
                elems[j].send_keys(Keys.RETURN)
            browser.switch_to.window(handles[0])
        
        
        print('搜索完成')
    except Exception as e:
        print('搜索异常')
        traceback.print_exc()

    print('\n按热键开始')

def handle_events(args):
    if args.current_key in keyVM and isinstance(args, KeyboardEvent) and args.event_type == 'key down':
        work(args.current_key)
    elif args.current_key == 'Q' and args.event_type == 'key down':
        hk.stop()           #无法正常退出
        print('程序退出')



if __name__ == "__main__":

    ctypes.windll.user32.SetProcessDPIAware(2)

    isComplete = False    
    lastQuestion = ''
    
    browser = webdriver.Chrome('./tools/chromedriver.exe')
    browser.get(search_engine)

    elems = [0]*4
    elems[0] = browser.find_element_by_id("kw")
    
    if searchAns==1:
        newWindow = 'window.open("https://www.baidu.com");'
        for i in range(1, 4):
            browser.execute_script(newWindow)
            handles = browser.window_handles
            browser.switch_to.window(handles[-1])
            elems[i] = browser.find_element_by_id("kw")
        handles = browser.window_handles
        browser.switch_to.window(handles[0])
       
    print('Chrome已启动')
   
    os.system("adb start-server")
    print('adb服务已启动')

    print('\n按热键开始')
    hk = Hook()
    hk.handler = handle_events
    hk.hook()





