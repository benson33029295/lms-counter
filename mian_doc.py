import urllib.request as req
import bs4
import requests
import os
from time import sleep
from tqdm import tqdm, trange


path = os.getcwd()
print("working directory is ", path)


head={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36"
        ,"Cookie":"_ga=GA10.3.371810122.1599381523; optimizelyEndUserId=oeu1603884149889r0.0832646509054058; cookie_locale=zh-tw; cookiesession1=51DDE42DSPPMS83GW68FQJ0UEI9OE146; PHPSESSID=32m42brvlin169qa4huneq1tj1; cookie_account=407010024; cookie_passwd=49792a1c23379b068908f3391cf86b62"}


firstpos="https://lms.ndmctsgh.edu.tw/home.php"


def find_html(pos):
    request=req.Request(pos,headers=head)
    with req.urlopen(request) as response:
        data=response.read().decode("utf-8")
    return data




######class_docpage_course_doc
##### main code starts form here ######
data0=find_html(firstpos)

###抓取課程名稱並列印
global classname_list
global dictionary
global posdic
global chosen_url
global data
root0=bs4.BeautifulSoup(data0,"html.parser")
classtitles=root0.find_all("div",class_="mnuItem")
classname_list=[]
posdic={}
num=0
for classtitle in classtitles:
    num=num+1
    classname_list.append(str(num)+". "+classtitle.a.string)
    posdic[str(num)]="https://lms.ndmctsgh.edu.tw/"+str(classtitle.a["href"])
for a in classname_list:
    print(a)
classchosen=input("請問您要更新哪一門課呢?")
chosen_url=posdic[classchosen]
print("正在更新"+classname_list[int(classchosen)-1]+"\nurl:"+chosen_url)

###################

global doc_url
data_chosen=find_html(chosen_url)
rootc=bs4.BeautifulSoup(data_chosen,"html.parser")
titles=rootc.find("div",class_="Edoc")
doc_url="https://lms.ndmctsgh.edu.tw/"+titles.a["href"]
print(doc_url)


global course_list
data=find_html(doc_url)

root=bs4.BeautifulSoup(data,"html.parser")
titles=root.find_all("div",class_="Econtent")
course_list=[]
for title in titles:
    if title.a!=None:
        course_list.append("https://lms.ndmctsgh.edu.tw"+title.a['href'])

print(course_list)

global doc_list
doc_list=[]
for a in course_list:
    url2=a
    data2=find_html(url2)
    root2=bs4.BeautifulSoup(data2,"html.parser")
    docs=root2.find_all("a",target="_blank")
    for n in docs:
        if "pptx"  in str(n) :
            doc_list.append(n)
        else:
            if "ppt"  in str(n) :
                doc_list.append(n)
            if "pdf" in str(n):
                doc_list.append(n)
            if "docx" in str(n) :
                doc_list.append(n)

print("檔案目錄如下:")
print(doc_list)



global ndl
global filename
filenames=[]
ndl=[]
for a in doc_list:
    if doc_list.index(a)%2!=0:
        ndl.append("https://lms.ndmctsgh.edu.tw"+a['href'])
        filenames.append(a.string)
print("ndl 如下---------------------------------------")
print(ndl)
print("filenames如下----------------------------------------------")
print(filenames)
if filenames==[]:
    print("該門課程沒有任何檔案可供下載")

#下載地址
def downloadone(adr,fn):
    Download_address=adr
    #把下載地址傳送給requests模組
    f=requests.get(Download_address,headers=head)
    #下載檔案
    with open(fn,"wb") as code:
         code.write(f.content)


from os import listdir



files = listdir(path)

global shortened_fi_list
shortened_fi_list=[]

with tqdm(total=100) as pbar:
    for filename in filenames:
        if ".pptx" in filename:
            shortened_filename=filename[:-5]
        if ".ppt" in filename:
            shortened_filename=filename[:-4]
        if ".docx" in filename:
            shortened_filename=filename[:-4]
        if ".pdf" in  filename:
            shortened_filename=filename[:-3]
        if shortened_filename in shortened_fi_list:
            print(str(shortened_filename)+"  already exists")
            
        address=ndl[filenames.index(filename)]
        downloadone(address,filename)
        print(shortened_filename+"successfully downloaded")
        sleep(0.1)
        pbar.update(100/len(filenames))
    





