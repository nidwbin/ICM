import os
import csv
import xlsxwriter

path = 'D:/Data/CSV/'
startNumber = 1
endNumber = 150
head = ['DATE', 'LATITUDE', 'LONGITUDE', 'DEPTH', 'PRS', 'TMP']
filename = '1-3'
workbook = xlsxwriter.Workbook(path + filename +'.xlsx')
sheet = workbook.add_worksheet('data')

if __name__ == '__main__':
    if os.path.exists(path):
        sheet.write_row('A1', head)
        count = 1
        for i in range(startNumber, endNumber):
            name = path + filename + '_%03d' % i + '_ct1.csv'
            if os.path.exists(name):
                data = csv.reader(open(name))
                dataList = [0, 0, 0, 0, 0, 0]
                cash = dataList.copy()
                tmp = 0
                for j in data:
                    tmp += 1
                    if j[0] != 'END_DATA':
                        if tmp == 11:
                            dataList[0] = int(j[0].replace('DATE = ', ''))
                        elif tmp == 13:
                            dataList[1] = float(j[0].replace('LATITUDE = ', ''))
                        elif tmp == 14:
                            dataList[2] = float(j[0].replace('LONGITUDE = ', ''))
                        elif tmp == 15:
                            t = j[0].replace('DEPTH = ', '')
                            if t != '':
                                dataList[3] = float(t)
                            else:
                                dataList[3] = -99.0
                        elif tmp > 17:
                            dataList[4] = float(j[0])
                            dataList[5] = float(j[2])
                            if cash != dataList:
                                count += 1
                                sheet.write_row('A' + str(count), dataList)
                                cash = dataList.copy()
            else:
                print('Not found \"' + name + '\"')
    else:
        print("Error path")
    workbook.close()
