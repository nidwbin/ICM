import os
import csv
import xlsxwriter

path = 'D:/Data/CSV/'
savePath = 'D:/Data/'
startNumber = 1
endNumber = 1000
head = ['DATE', 'LATITUDE', 'LONGITUDE', 'DEPTH', 'PRS', 'TMP']
startRow = 4
endRow = 5


def save(name, datalist):
    data = csv.reader(open(name))
    tmpList = [0, 0, 0, 0, 0, 0]
    cash = tmpList.copy()
    tmp = 0
    for j in data:
        tmp += 1
        if j[0] != 'END_DATA':
            if tmp == 11:
                tmpList[0] = int(j[0].replace('DATE = ', ''))
            elif tmp == 13:
                tmpList[1] = float(j[0].replace('LATITUDE = ', ''))
            elif tmp == 14:
                tmpList[2] = float(j[0].replace('LONGITUDE = ', ''))
            elif tmp == 15:
                t = j[0].replace('DEPTH = ', '')
                if t != '':
                    tmpList[3] = float(t)
                else:
                    tmpList[3] = -99.0
            elif tmp > 17:
                tmpList[4] = float(j[0])
                tmpList[5] = float(j[2])
                if tmpList[5] != -999 and cash != tmpList:
                    datalist.append(tmpList.copy())
                    cash = tmpList.copy()


def fun(num):
    workbook = xlsxwriter.Workbook(savePath + num + '.xlsx')
    sheet = workbook.add_worksheet('data')
    if os.path.exists(path):
        sheet.write_row('A1', head)
        datalist = []
        for i in range(startNumber, endNumber):
            name = path + num + '_%02d' % i + '_ct1.csv'
            if os.path.exists(name):
                save(name, datalist)
            else:
                print('Not found \"' + name + '\"')
            name = path + num + '_%03d' % i + '_ct1.csv'
            if os.path.exists(name):
                save(name, datalist)
            else:
                print('Not found \"' + name + '\"')
        datalist.sort()
        for i in range(2, len(datalist) + 1):
            sheet.write_row('A' + str(i), datalist[i - 1])
    else:
        print("Error path")
    workbook.close()


if __name__ == '__main__':
    for i in range(startRow, endRow + 1):
        for j in range(1, 6):
            fun(str(i) + '-' + str(j))
