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
    def close(self):
        self.close()
    def major(self):
        # account = input("輸入你的教務系統帳號：")  # 輸入你的教務系統帳號
        # password = input("輸入你的教務系統密碼：")  # 輸入你的教務系統密碼

        if(self.ui.textEdit)
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
            


if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())