import urllib
import urllib.request
import http.cookiejar
import string
import re
from bs4 import BeautifulSoup
import time


#-----------------下面是模拟登陆部分----------------
#登录的主页面
hosturl = 'http://ielab.buct.edu.cn' #自己填写
#post数据接收和处理的页面（我们要向这个页面发送我们构造的Post数据）
posturl = 'http://ielab.buct.edu.cn/BBS/login.aspx' #从数据包中分析出，处理post请求的url
#设置一个cookie处理器，它负责从服务器下载cookie到本地，并且在发送请求时带上本地的cookie
cj = http.cookiejar.LWPCookieJar()
cookie_support = urllib.request.HTTPCookieProcessor(cj)
opener = urllib.request.build_opener(cookie_support, urllib.request.HTTPHandler)
urllib.request.install_opener(opener)
#打开登录主页面（他的目的是从页面下载cookie，这样我们在再送post数据时就有cookie了，否则发送不成功）
h = urllib.request.urlopen(hosturl)
#构造header，一般header至少要包含一下两项。这两项是从抓到的包里分析得出的。
headers = {'User-Agent' : 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:14.0) Gecko/20100101 Firefox/14.0.1',
           'Referer' : 'http://ielab.buct.edu.cn/BBS/login.aspx?postusername=&=&selectedtemplateid=4'}
#构造Post数据，他也是从抓大的包里分析得出的。
postData = {'op' : 'dmlogin',
            'f' : 'st',
            'username' : '测试员1', #你的用户名
            'password' : '123456', #你的密码，密码可能是明文传输也可能是密文，如果是密文需要调用相应的加密算法加密
            'question' : 0,
            'answer' :None,
            'selectedtemplateid': 4,
            'login': None,
            'rmbr' : 'true',   #特有数据，不同网站可能不同
            'tmp' : '0.7306424454308195'  #特有数据，不同网站可能不同

            }

#需要给Post数据编码
postData = urllib.parse.urlencode(postData).encode('utf-8')
#通过urllib2提供的request方法来向指定Url发送我们构造的数据，并完成登录过程
request = urllib.request.Request(posturl, postData, headers)
#print(request)
response = urllib.request.urlopen(request)
#text = (response.read()).decode('utf-8')
#print(text)
#-----------------上面是模拟登陆-----------------------


#-----------------下面是获取指定网页内容----------------

page = urllib.request.urlopen("http://ielab.buct.edu.cn/BBS/")
soup = BeautifulSoup(page)
page.close()
h1_a = (soup.find('h1')).find('a')
hyperlink = hosturl + h1_a.get('href')
base_href = hyperlink.rstrip(".aspx")
#container for soup
L1 = []
#container for address,5 pages may be enough to our forum
L2=  [hyperlink,base_href+"-2.aspx",base_href+"-3.aspx",base_href+"-4.aspx",base_href+"-5.aspx"]
for add in L2:
    page = urllib.request.urlopen(add)
    soup = BeautifulSoup(page)
    page.close()
    #考虑到后面去重，所以这里不作处理。
    #该论坛有些特点:帖子地址里面的数字有蹊跷
    L1.append(soup)

#-----------------文件操作
#r,w,a
save_path="E:/Documents/study/研究生/实验室报告/names.txt" 
f_obj = open(save_path,'a')
#----------------日期和时间
## dd/mm/yyyy格式
f_obj.write(time.strftime("%Y/%m/%d")+"\t"+(time.strftime("%H:%M:%S"))+"\r\n")

L3 = []
for each in L1:
    for a in each.findAll('cite'):
        b = a.find('span')
        if b != None :
            L3.append(b.text)
                      
#list 去重
L3 = set(L3)
##for each in L3:
##    print(each)
##    f_obj.write(each+"\r\n")
##f_obj.close()


import  xdrlib ,sys
import xlrd
import xlwt3

#读名单
listNames = []
try:
    data = xlrd.open_workbook("E:/Documents/study/研究生/实验室报告/实验室研究生标准名单.xls")
except Exception as e:
        print (str(e))   
table = data.sheets()[0]
nrows = table.nrows
ncols = table.ncols
for i in range(1,nrows):
   listNames.append(table.cell(i,0).value)

#根据爬到的名单和标准名单比对
distPath = "E:/Documents/study/研究生/实验室报告/"
fileName = time.strftime("%Y-%m-%d")+"文献下载名单.xls"
wbk=xlwt3.Workbook()
sheet = wbk.add_sheet('sheet1',True)
sheet.write(0, 0, "研究生")
sheet.write(0, 1, "已下载")
for i in range(1,nrows):
    sheet.write(i,0,listNames[i-1])
    if listNames[i-1] in L3:
        sheet.write(i,1,"√")
wbk.save(distPath+fileName)





