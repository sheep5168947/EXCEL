from os import write
import os
import pandas as pd
from pandas.core.frame import DataFrame
from splinter.browser import Browser
from bs4 import BeautifulSoup
import time


def webscrabe(string,account,password,index):
    borwser = Browser("chrome")

    borwser.visit("https://aca.nuk.edu.tw/Aca/login.asp")
    borwser.find_by_name("Account").first.fill(account)
    borwser.find_by_name("Password").first.fill(password)
    borwser.find_by_name("B1").click()
    borwser.visit("https://aca.nuk.edu.tw/SN2/Basic/Student.asp")
    borwser.find_by_name("No").first.fill(string)
    borwser.find_by_name("B3").click()
    name = borwser.find_by_xpath("//tbody[1]/tr[2]/td[3]").text
    department = borwser.find_by_xpath("//tbody[1]/tr[2]/td[5]").text
    typename(index,name,department)
    borwser.quit



def major():
    account = input("輸入你的教務系統帳號：")  # 輸入你的教務系統帳號
    password = input("輸入你的教務系統密碼：")  # 輸入你的教務系統密碼

    
    for i in df.index:
        string = str(df['學號'][i])
        string = string[0:8]
        if(string == "NAN"):
            continue
        else:
            webscrabe(string,account,password,i)
        # print(string)
def typename(i,name,department):
    df['姓名'][i] = name
    df['系所'][i] = department
    
def grade():
    for i in df.index:

        string = str(df['學號'][i])
        
        # print(st[0])
        if(string[0] == 'M'):
            if(string[3] == '9'):
                df['年級'][i] = "碩一"
            elif(string[3] == '8'):
                df['年級'][i] = "碩二"
            else:
                df['年級'][i] = "其他"
        elif(string[0] == 'L'):
            df['年級'][i] = "其他"
        elif(string[0] == 'n'):
            df['年級'][i] = "其他"
        else:
            if(string[3] == '9'):
                df['年級'][i] = "大一"
            elif(string[3] == '8'):
                df['年級'][i] = "大二"
            elif(string[3] == '7'):
                df['年級'][i] = "大三"
            elif(string[3] == '6'):
                df['年級'][i] = "大四"
            else:
                df['年級'][i] = "其他"

# excelPath = input("輸入你的excel名稱(excel位置要跟執行檔同一個資料夾且需要副檔名為xlsx): ")
# df = pd.read_excel(excelPath)
# grade()
# major()

# DataFrame(df).to_excel('test2.xlsx',sheet_name='活動類_學生', index=False, header=True)