import xlrd
import xlwt
import os, shutil
from Staff import Staff

doneDirPath = os.path.join(os.getcwd(), '文件池', '已完成')
if not os.path.exists(doneDirPath):
    os.makedirs(doneDirPath)

list = os.listdir('文件池')
myWorkbook = xlwt.Workbook()
mySheet = myWorkbook.add_sheet('New Staff')
for i in range(0, len(list)):
    path = os.path.join(os.getcwd(), '文件池', list[i])
    print(path)

    if os.path.isfile(path) and os.path.splitext(list[i])[-1] == ".xlsx":
        print('test')
        book = xlrd.open_workbook(path)
        sh = book.sheet_by_index(0)
        staff = Staff(sh)
        staff.writeToSheet(mySheet, i)

        # shutil.move(path, os.path.join(doneDirPath, list[i]))


myWorkbook.save('结果.xls')