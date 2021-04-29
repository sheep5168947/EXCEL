from os import write
import pandas as pd
from pandas.core.frame import DataFrame

df = pd.read_excel("學生EP欄位1100309.xlsx",sheet_name="活動類_學生")
# writer = pd.ExcelWriter("學生EP欄位1100309.xlsx")
# DataFrame.to_excel(writer,sheet_name="活動類_學生")
# workbook = writer.book
# worksheet = writer.sheets['活動類_學生']
# writer.save
for i in range(1,538):

    string = str(df['學號'][i])
    # print(st[0])
    if(string[0] == 'M'):
        if(string[3] == '9'):
            df['年級'][i] = "碩一"
        elif(string[3] =='8'):
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

    # print(df['學號'][i])
    # print(df.values[i])

print(df)
DataFrame(df).to_excel('學生EP欄位1100309.xlsx', sheet_name='活動類_學生', index=False, header=True)