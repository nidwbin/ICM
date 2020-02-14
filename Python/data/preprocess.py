import os
import csv
import xlsxwriter

path = 'D:/Data/CSV/'
startNumber = 1
endNumber = 150
head = ['DATE', 'LATITUDE', 'LONGITUDE', 'DEPTH', 'PRS', 'TMP']
workbook = xlsxwriter.Workbook(path + 'new_data.xlsx')
sheet = workbook.add_worksheet('data')

if __name__ == '__main__':
    if os.path.exists(path):
        sheet.write_row('A1', head)
        for i in range(startNumber, endNumber):
            name = path + 'data_from_SeaDataNet_North-Sea_TS_QCed_V1.1_' + '%03d' % i + '_ct1.csv'
            if os.path.exists(name):
                data = csv.reader(open(name))
                dataList = [0, 0, 0, 0, 0, 0]
                tmp = 0
                count = 1
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
                            dataList[3] = float(j[0].replace('DEPTH = ', ''))
                        elif tmp > 17:
                            dataList[4] = float(j[0])
                            dataList[5] = float(j[2])
                            count += 1
                            sheet.write_row('A' + str(count), dataList)
            else:
                print('Not found \"' + name + '\"')
    else:
        print("Error path")
    workbook.close()
