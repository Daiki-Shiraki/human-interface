import os
import re

def GetOutputFileName(file_path):
    text = os.path.basename(file_path)
    text = re.findall(r"\d+", text)
    text = "_".join(text)+'.xlsx'
    return text

def Number2String(number):
    if(isinstance(number, str)):
        print('一部の学生IDがstr型で入力されました：' + number)
        number = number.replace(" ", '')
        number = int(number)
    return str(number)
