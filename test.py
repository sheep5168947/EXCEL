from PyQt5 import QtWidgets, QtGui, QtCore
from excel_ui import Ui_MainWindow
import sys
from os import write
import os
import pandas as pd
from pandas.core.frame import DataFrame
from splinter.browser import Browser
from bs4 import BeautifulSoup
import time



class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.major)
        self.ui.pushButton_2.clicked.connect(self.close)
        self.ui.pushButton_3.clicked.connect(self.clean)
        self.ui.textEdit_2.setPlaceholderText("請正確輸入且不可為白")
        self.ui.textEdit.setPlaceholderText("請正確輸入且不可為白")
        self.ui.textEdit_3.setPlaceholderText("請正確輸入且不可為白")
        # self.ui.pushButton.setEnabled(false)
    def clean(self):
        self.ui.textEdit.setPlainText("")
        self.ui.textEdit_2.setPlainText("")
        self.ui.textEdit_3.setPlainText("")
    def close(self):
        self.close()
    def webscrabe(self,string,account,password,index):
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
        self.typename(index,name,department)
        borwser.quit

    def major(self):
        # account = input("輸入你的教務系統帳號：")  # 輸入你的教務系統帳號
        # password = input("輸入你的教務系統密碼：")  # 輸入你的教務系統密碼
        
        if(self.ui.textEdit.toPlainText() == "" or self.ui.textEdit_2.toPlainText() == "" or self.ui.textEdit_3.toPlainText() == ""):
            print("nope")
        else:
            df = pd.read_excel(self.ui.textEdit_3.toPlainText())
            account = self.ui.textEdit.toPlainText()
            password = self.ui.textEdit_2.toPlainText()
            for i in df.index:
                string = str(df['學號'][i])
                string = string[0:8]
                if(string == "NAN"):
                    continue
                else:
                    self.webscrabe(string,account,password,i)
                    self.grade()
            # print(string)
    def typename(self,i,name,department):
        df['姓名'][i] = name
        df['系所'][i] = department
        
    def grade(self):
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
            
if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
    df = pd