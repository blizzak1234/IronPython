# -*- coding: utf-8 -*-

from model.group import Group
import random
import string
import os.path #для работы с путями до файлов
import getopt # для использования опций из командной строки
import sys #для получения доступа к этим опциям
import time

import clr #для работы с виртуальной машиной .Net
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

# описание получение опций из оф. документации питона
try:
    opts, args = getopt.getopt(sys.argv[1:], "n:f:", ["number of groups", "file"])
except getopt.GetoptError as err:
    getopt.usage()
    sys.exit(2)

# указываем дефольные значения
n = 2
f = "data/groups.xlsx"

for o, a in opts:
    if o == "-n": #если название опции равно -n
        n = int(a) #значит в ней задается количество групп
    elif o == "-f":
        f = a

def random_string(prefix, maxlen):
    symbols = string.ascii_letters + string.digits
    return prefix + "".join([random.choice(symbols) for i in range(random.randrange(maxlen))])


testdata = [Group(name="")] + [
    Group(name=random_string("name", 10))
    for i in range(n) #цикл для генерации случайных данных. 5 раз. то есть все данные для теста будут состоять из одной пустой группы и 5 групп со случ. данными

]

# сохраняем сгенерированные данные в фаил
file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", f)  # путь к файлу

# открываем фаил




excel = Excel.ApplicationClass()
excel.visible = True

workbook = excel.Workbooks.Add()
sheet = workbook.Activesheet

for i in range(len(testdata)):
    sheet.Range["A%s" % (i+1)].Value2 = testdata[i].name

workbook.SaveAs(file)

time.sleep(10)

excel.Quit()
