# -*- coding: utf-8 -*-
"""
Created on Tue Nov 23 20:10:41 2021

@author: User
"""

import datetime
import time
import calendar
from colorama import init, Fore  #文字顏色套件

#google語音
from gtts import gTTS
from playsound import playsound
import re

#speak to text
import speech_recognition
from win32com.client import Dispatch

import requests

import pandas_read_xml as pdx

def speak(text):
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(text)


def output(text):
    print(text)
    speak(text)
    

def getNews(topic=''):
    
    # ews api
    main_url = "https://newsapi.org/v2/top-headlines?"+"country=tw&"+"q="+topic+"&"+"apiKey=b3de4680862849e5943c091018b181b9"
  
    # fetching data in json format
    open_page = requests.get(main_url).json()
  
    # getting all articles in a string article
    article = open_page["articles"]
  
    # empty list which will 
    # contain all trending news
    results = []
    
    for ar in article:
        results.append({'title':ar["title"],'des':ar["description"],'url':ar["url"]})

    return results
def markTheDay(day, cal, color):
    cal_f = cal[:cal.find(str(day))]
    cal = cal[cal.find(str(day))+len(str(day)):]
    init(autoreset=True)
    print(cal_f + color + str(day) + Fore.RESET, end = '')
    return cal

def ToDoList():
    #待辦事項
    TODO_list = []	#存取待辦事項，格式:[名稱, [年, 月, 日], 相距天數，幾天前提醒,提醒日日期]
    #讀取txt
    path = 'ToDoList.txt'
    f = open(path, 'a+')
    f.close()
    f = open(path, 'r')
    list_history = f.readlines()
    list_data = []
    '''
    for i in list_history:
        a=re.findall(r'\w+',i) 
        list_data.extend(a)
    '''
    for line in list_history:
        year_todo = line[:line.find('/')]
        month_todo = line[line.find('/')+1:line.find('/',line.find('/')+1)]
        day_todo = line[line.find('/',line.find('/')+1)+1:line.find(' ')]
        TODO_name = line[line.find(' ')+1:line.find(':')]
        
        TODO_date=[int(year_todo),int(month_todo),int(day_todo)]
        #計算相差天數
        time_dis =datetime.date(TODO_date[0], TODO_date[1], TODO_date[2]) - datetime.date.today()
        
        TODO_hint_date = line[line.find(' ',line.find(' ')+1)+1:line.find('\n')]
        TODO_hint_date = datetime.date.fromisoformat(TODO_hint_date)
        
        TODO_hint = TODO_hint_date - datetime.date.today()
        
        TODO_list.append([TODO_name, TODO_date, time_dis.days, TODO_hint.days, TODO_hint_date])
        
    #擷取日期和事項

    date = datetime.date.today()	#取得今天日期
    date_str = str(date)	#傳換成字串
    #取出年月日
    year = int(date_str[:date_str.find("-")])
    month = int(date_str[date_str.find("-")+1:date_str.find("-",date_str.find("-")+1)])
    day = int(date_str[date_str.find("-",date_str.find("-")+1)+1:])  #先轉int將0消掉


    while True:
        
        print("有甚麼待辦事項需要我幫你紀錄的嗎?")
        speak("有甚麼待辦事項需要我幫你紀錄的嗎?")
        #speech("有甚麼待辦事項需要我幫你紀錄的嗎?")
        #s=input()
        s = Voice_To_Text()
        if '沒' in s:
            break
        if '年' in s:
            year_todo = s[s.find('年')-4:s.find('年')]
        else:
            year_todo = datetime.date.today().year
            
        if '月' in s:
            if(s.find('月') > 1):
                month_todo = s[s.find('月')-2:s.find('月')]
                if not (month_todo[0] in ['1','2','3','4','5','6','7','8','9']) :
                    month_todo = month_todo[-1]
            else:
                month_todo = s[s.find('月')-1:s.find('月')]
        else:
            month_todo = datetime.date.today().month
        if '日' in s:
            if s.find('日')>1:
                day_todo = s[s.find('日')-2:s.find('日')]
                if not (day_todo[0] in ['1','2','3','4','5','6','7','8','9']) :
                    day_todo = day_todo[-1]
            else:
                day_todo = s[s.find('日')-1:s.find('日')]
        elif '號' in s:
            if s.find('號'):
                day_todo = s[s.find('號')-2:s.find('號')]
                if not (day_todo[0] in ['1','2','3','4','5','6','7','8','9']) :
                    day_todo = day_todo[-1]
            else:
                day_todo = s[s.find('號')-1:s.find('號')]
        else:
            
            print("不好意思我聽不懂，麻煩你再說一次")
            speak("不好意思我聽不懂，麻煩你再說一次")
            #speech("不好意思我聽不懂，麻煩你再說一次")
            continue
        if '要' in s:
            TODO_name = s[s.find('要')+1:]
        else:
            
            print("不好意思我聽不懂，麻煩你再說一次")
            speak("不好意思我聽不懂，麻煩你再說一次")
            #speech("不好意思我聽不懂，麻煩你再說一次")
        TODO_date=[int(year_todo),int(month_todo),int(day_todo)]
        #計算相差天數
        time_dis = datetime.date(TODO_date[0], TODO_date[1], TODO_date[2]) - datetime.date.today()
        #TODO_hint = int(input("請問要在幾天前提醒您呢?"))
        print('請問要在幾天前提醒您呢?')
        speak('請問要在幾天前提醒您呢?')
        #speech('請問要在幾天前提醒您呢?')
        unknow = True;
        while(unknow):
            s = Voice_To_Text()
            if '天' in s:
                s = re.sub('二','2',s)
                s = re.sub('兩','2',s)
                if len(re.findall('[0-9]+(?=天)', s)) ==0:
                    continue;
                else:
                    TODO_hint = int(re.findall('[0-9]+(?=天)', s)[0])
                '''
                if s[:s.find('天')] == "兩":
                    TODO_hint = 2
                else:
                    TODO_hint = int(s[:s.find('天')])
                '''
                unknow = False
            else:
                output("不好意思我聽不懂，麻煩你再說一次")
                #speak("不好意思我聽不懂，麻煩你再說一次")
                #speech("不好意思我聽不懂，麻煩你再說一次")
        # 算出提醒日的日期
        TODO_hint_date = datetime.date(TODO_date[0], TODO_date[1], TODO_date[2]) - datetime.timedelta(days=TODO_hint)
        TODO_list.append([TODO_name, TODO_date, time_dis.days, TODO_hint, TODO_hint_date])

    TODO_list.append(['Today', [year, month, day], 0,-1,-1])
    #依照日期排序
    TODO_list.sort(key = lambda TODO_list:[TODO_list[1][2]])
    TODO_list.sort(key = lambda TODO_list:[TODO_list[1][1]])
    TODO_list.sort(key = lambda TODO_list:[TODO_list[1][0]])

    cal = calendar.month(year, month)
    # 標示日期會標到年分故先輸出月曆到Sunday 
    cal_f = cal[:cal.find('Su')]
    # 存取尚未輸出的部分
    cal = cal[cal.find(str('Su')):]
    print(cal_f, end = '')

    # 將當月的待辦事項日期用紅色標示，當天則用黃色
    for i in range(len(TODO_list)):
        if TODO_list[i][1][0] == year:
            if TODO_list[i][1][1] == month:
                if TODO_list[i][0] == 'Today':
                    cal = markTheDay(TODO_list[i][1][2], cal, Fore.GREEN)
                else:
                    cal = markTheDay(TODO_list[i][1][2], cal, Fore.RED)
            elif TODO_list[i][1][1] > month:
                break
        elif TODO_list[i][1][0] > year:
            break
    print(cal)
    f.close()
    f = open(path, 'w')
    # 輸出待辦清單
    for i in range(len(TODO_list)):
        if TODO_list[i][2] != 0:
            print(str(TODO_list[i][1][0]) + '/' + 
                  str(TODO_list[i][1][1]) + '/' + 
                  str(TODO_list[i][1][2]) + ' ' +
                  TODO_list[i][0], end = '', file=f)
        print(str(TODO_list[i][1][0]) + '/' +
    		  str(TODO_list[i][1][1]) + '/' + 
    		  str(TODO_list[i][1][2]) + ' ' +
    		  TODO_list[i][0], end = '')
            
        if(TODO_list[i][2] < 0):
            print(':' + str(abs(TODO_list[i][2])) + '天前', end = '', file=f)
            print(':' + str(abs(TODO_list[i][2])) + '天前')
            print(' '+str(TODO_hint_date), file=f)
        elif(TODO_list[i][2] > 0):
            print(':' + str(abs(TODO_list[i][2])) + '天後', end = '', file=f)
            print(':' + str(abs(TODO_list[i][2])) + '天後')
            print(' '+str(TODO_hint_date), file=f)
        elif(TODO_list[i][2] == 0):
            print('')
        
    print('')

    #輸出到txt
    f.close()
    # 提醒功能
    for i in range(len(TODO_list)):
      if(TODO_list[i][4]==datetime.date.today()):
    	  text = "提醒您，"+str(TODO_list[i][2])+"天後要"+str(TODO_list[i][0])
    	  print(text)
    	  speak(text)
          #speech(text)

    print('')
            
    print("待辦事項已經幫您紀錄完畢囉~")
    speak("待辦事項已經幫您紀錄完畢囉")

def speech(text):
    file_name = str(int(time.time()))                
    speech = gTTS(text = text, lang = language, slow = False)
    speech.save(file_name + '.mp3')
    playsound(file_name + '.mp3')

#在日曆上標示日期



    
def Voice_To_Text():
    r = speech_recognition.Recognizer()
    with speech_recognition.Microphone() as source: 
        r.adjust_for_ambient_noise(source)
        print("請開始說話:")       
        audio = r.listen(source)
        
    try:
        Text = r.recognize_google(audio, language="zh-TW")     
              
    except speech_recognition.UnknownValueError:
        Text = "無法翻譯"
    except speech_recognition.RequestError as e:
        Text = "無法翻譯{0}".format(e)

    print(Text)          
    return Text

def listen():
    text = Voice_To_Text()
    while(True):
        if '無法翻譯' in text:
            text = Voice_To_Text()
            continue
        else:
            break
    return text

def getCovidNews():
  df = pdx.read_xml('https://www.mohw.gov.tw/rss-16-1.html')
  df = df.pipe(pdx.flatten)
  df = df.pipe(pdx.flatten)
  df = df.pipe(pdx.flatten)
  df = df.pipe(pdx.flatten)
  dataList = []
  for i in range(len(df)):
    
    dataList.append([df['rss|channel|item|title'][i],df['rss|channel|item|description'][i],df['rss|channel|pubDate'][i]])
  return dataList

def covidInfo():
  date = datetime.date.today()
  d = getCovidNews()
  for i in range(len(d)):
    if ('新增' in d[i][0] and 'COVID-19' in d[i][0] and '確定病例' in d[i][0]):
      index = i
      break
  newsDay = int(d[index][1][d[i][1].find('今(')+2:d[index][1].find('日')-1])
  if(newsDay == date.day):
    timeStr = '今天'
  elif(newsDay == date.day-1):
    timeStr = '昨天'
  else:
    timeStr = str(newsDay)+'日'
  totalCase = d[index][1][d[index][1].find('國內新增')+4:d[index][1].find('例')]
  localCase = d[index][1][d[index][1].find('分別為')+3:d[index][1].find('例本土個案')]
  outsideCase = d[index][1][d[index][1].find('及')+1:d[index][1].find('例境外移入')]
  localSexCaseStr = d[index][1][d[index][1].find('新增本土個案'):d[index][1].find(',')]
  localSexCaseM = localSexCaseStr[localSexCaseStr.find('為')+1:localSexCaseStr.find('例男性')]
  localSexCaseW = localSexCaseStr[localSexCaseStr.find('、')+1:localSexCaseStr.find('例女性')]
  outsideSexCaseStr = d[index][1][d[index][1].find('新增境外移入'):d[index][1].find(',')]
  outsideSexCaseM = outsideSexCaseStr[outsideSexCaseStr.find('為')+1:outsideSexCaseStr.find('例男性')]
  outsideSexCaseW = outsideSexCaseStr[outsideSexCaseStr.find('、')+1:outsideSexCaseStr.find('例女性')]
  contryFrom = d[index][1][d[index][1].find('分別自'):d[index][1].find('。',d[index][1].find('分別自'))]
  return {'totalCase':totalCase,'localCase':localCase,'outsideCase':outsideCase,'localSexCaseM':localSexCaseM,'localSexCaseW':localSexCaseW,'outsideSexCaseM':outsideSexCaseM,'outsideSexCaseW':outsideSexCaseW,'contryFrom':contryFrom,'timeStr':timeStr}
  date = datetime.date.today()
  d = getCovidNews()
  for i in range(len(d)):
    if ('新增' in d[i][0] and 'COVID-19' in d[i][0] and '確定病例' in d[i][0]):
      index = i
      break
  newsDay = int(d[index][1][d[i][1].find('今(')+2:d[index][1].find('日')-1])
  if(newsDay == date.day):
    timeStr = '今天'
  elif(newsDay == date.day-1):
    timeStr = '昨天'
  else:
    timeStr = str(newsDay)+'日'
  totalCase = d[index][1][d[index][1].find('國內新增')+4:d[index][1].find('例')]
  localCase = d[index][1][d[index][1].find('分別為')+3:d[index][1].find('例本土個案')]
  outsideCase = d[index][1][d[index][1].find('及')+1:d[index][1].find('例境外移入')]
  localSexCaseStr = d[index][1][d[index][1].find('新增本土個案'):d[index][1].find(',')]
  localSexCaseM = localSexCaseStr[localSexCaseStr.find('為')+1:localSexCaseStr.find('例男性')]
  localSexCaseW = localSexCaseStr[localSexCaseStr.find('、')+1:localSexCaseStr.find('例女性')]
  outsideSexCaseStr = d[index][1][d[index][1].find('新增境外移入'):d[index][1].find(',')]
  outsideSexCaseM = outsideSexCaseStr[outsideSexCaseStr.find('為')+1:outsideSexCaseStr.find('例男性')]
  outsideSexCaseW = outsideSexCaseStr[outsideSexCaseStr.find('、')+1:outsideSexCaseStr.find('例女性')]
  contryFrom = d[index][1][d[index][1].find('分別自'):d[index][1].find('。',d[index][1].find('分別自'))]
  return {'timeStr':timeStr,'totalCase':totalCase,'localCase':localCase,'outsideCase':outsideCase,'localSexCaseM':localSexCaseM,'localSexCaseW':localSexCaseW,'outsideSexCaseM':outsideSexCaseM,'outsideSexCaseW':outsideSexCaseW,'contryFrom':contryFrom}

#情感
with open('D:/NUK/AI語意/期中專題/ntusd-negative.txt', mode='r', encoding='utf-8') as f:
    negs = f.readlines()
with open('D:/NUK/AI語意/期中專題/ntusd-positive.txt', mode='r', encoding='utf-8') as f:
    poss = f.readlines()

language = 'zh'

#Intro
print("歡迎回家!")
print("今天過得如何呀?")
#s = input()
speak("歡迎回家!")
speak("今天過得如何呀?")
#speech("歡迎回家!")
#speech("今天過得如何呀?")
s = Voice_To_Text()
neg = [] #負面詞句List
pos = []
x = 0
isNotNeg = True
for i in negs:
    a=re.findall(r'\w+',i) 
    neg.extend(a)
for i in poss:
    a=re.findall(r'\w+',i) 
    pos.extend(a)
    

for i in range(len(neg)):
    if neg[i] in s.strip():
        print('秀秀~讓我幫你分攤一些工作吧!')
        speak('秀秀 讓我幫你分攤一些工作吧!')
        #speech('秀秀 讓我幫你分攤一些工作吧!')
        isNotNeg = False
        break

if(isNotNeg):
    for i in range(len(pos)):
        if pos[i] in s.strip():
            print('恭喜~真是開心的一天!那讓我幫你分攤一些工作吧!讓你天天開心!')
            speak('恭喜 真是開心的一天!那讓我幫你分攤一些工作吧!讓你天天開心!')
            #speech('恭喜 真是開心的一天!那讓我幫你分攤一些工作吧!讓你天天開心!')
            break

 

while(True):
    output("我可以幫您紀錄待辦事項、播報頭條新聞、提供COVID-19疫情資訊")
    output("請問需要什麼服務呢?")
    s = Voice_To_Text()
    if '事項' in s:
        ToDoList()
    elif '新聞' in s:
        news = getNews()
        output('這邊將為您播報本日頭條新聞')
        for i in range(3):
            output('第'+str(i+1)+'則:'+news[i]['title'])
        output('請問需要詳細內容嗎?')
        s = listen()
        if(not('不' in s)):
            while(True):
                output('請問需要第幾則的內容?')
                s = listen()
                s = re.sub('一','1',s)
                s = re.sub('二','2',s)
                s = re.sub('三','3',s)
                num = re.findall('\d', s)
                output(news[int(num[0])-1]['des'])
                output('請問需要新聞連結嗎?')
                s = listen()
                if(not('不' in s)):
                    output('這裡是新聞連結:')
                    print(news[int(num[0])-1]['url'])
                output('請問還需要其他則的內容嗎')
                s = listen()
                if('不' in s):
                    break
        output('請問您有想要搜尋特定主題的內容嗎?')
        s = listen()
        if(not('不' in s)):
            while True:
                output('請問想要搜尋什麼主題')
                s = listen()
                news = getNews(s)
                output('這邊為您播報'+s+'的新聞')
                if(len(news)==0):
                    output('抱歉我找不到'+s+'的新聞')
                    output('請問還需要其他主題的內容嗎')
                    s = listen()
                    if('不' in s):
                        break
                    else:
                        continue
                if(len(news)<3):
                    n = len(news)
                else:
                    n=3
                for i in range(n):
                    output('第'+str(i+1)+'則:'+news[i]['title'])
                output('請問需要詳細內容嗎?')
                s = listen()
                if(not('不' in s)):
                    while(True):
                        output('請問需要第幾則的內容?')
                        s = listen()
                        s = re.sub('一','1',s)
                        s = re.sub('二','2',s)
                        s = re.sub('三','3',s)
                        num = re.findall('\d', s)
                        output(news[int(num[0])-1]['des'])
                        output('請問需要新聞連結嗎?')
                        s = listen()
                        if(not('不' in s)):
                            output('這裡是新聞連結:')
                            print(news[int(num[0])-1]['url'])
                        output('請問還需要其他則的內容嗎')
                        s = listen()
                        if('不' in s):
                            break
                output('請問還需要搜尋其他主題嗎')
                s = listen()
                if('不' in s):
                    break
    #elif '天氣' in s:
        
    elif 'COVID-19' in s or '疫情' in s:
        covidData = covidInfo()
        output(covidData['timeStr']+'新增'+covidData['totalCase']+'例COVID-19確定病例')
        while(True):
            output('請問想更了解境外還是境內?')
            s = Voice_To_Text()
            if '外' in s:
                output(covidData['timeStr']+"境外移入"+covidData['outsideCase']+'例')
                output("請問需要更多資訊嗎")
                s = listen()
                if(not('不' in s)):
                    output('境外移入個案為'+covidData['outsideSexCaseM']+'位男性和'+covidData['outsideSexCaseW']+'位女性，'+covidData['contryFrom'])
                output('請問需要境內案例的資訊嗎?')
                s = listen()
                if(not('不' in s)):
                    output(covidData['timeStr']+"境內案例共"+covidData['localCase']+'例')
                    output("請問需要更多資訊嗎")
                    s = listen()
                    if(not('不' in s)):
                        output('境內個案為'+covidData['localSexCaseM']+'位男性和'+covidData['localSexCaseW']+'位女性')
            elif '內' in s:
                output(covidData['timeStr']+"境內案例共"+covidData['localCase']+'例')
                output("請問需要更多資訊嗎")
                s = listen()
                if(not('不' in s)):
                    output('境內個案為'+covidData['localSexCaseM']+'位男性和'+covidData['localSexCaseW']+'位女性')
                output('請問需要境外案例的資訊嗎?')
                s = listen()
                if(not('不' in s)):
                    output(covidData['timeStr']+"境外移入"+covidData['outsideCase']+'例')
                    output("請問需要更多資訊嗎")
                    s = listen()
                    if(not('不' in s)):
                        output('境外移入個案為'+covidData['outsideSexCaseM']+'位男性和'+covidData['outsideSexCaseW']+'位女性，'+covidData['contryFrom']) 
            else:
                continue
            break
    else:
        continue
    output('請問還需要其他服務嗎?')
    s =listen()
    if('不' in s):
        break;
output('好的，我們下次見!')