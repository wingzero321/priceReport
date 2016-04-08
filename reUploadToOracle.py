# -*- coding: utf-8 -*-

import sys
import cx_Oracle
import csv
import time
import datetime
import os
import xlsxFormatSetting as settings
import downloadUtil

reload(sys)
sys.setdefaultencoding('utf8')
reload(sys)

def downloadFile(dataTime):
    readFilename = 'price_Malibu_' + dataTime + '_new.csv'
    print 'start: download file ' + readFilename
    downloadUtil.download(readFilename)
    print 'download success!'

def uploadToOracle(dataTime):
    readFilePath = settings.filePath + '\\data\\price_Malibu_' + dataTime + '_new.csv'
    if os.path.exists(readFilePath):
        print 'Begin upload File: ' + readFilePath
    else:
        print 'No such File: ' + readFilePath
        return
    # readFilePath = 'D:\\work\\priceReport\\data\\price_Malibu_2016-03-18_new.csv'
    readFile = file(readFilePath, 'r')
    reader = csv.reader(readFile)

    conn = cx_Oracle.connect('admin/admin@10.0.0.233:15233/ORCL')
    c = conn.cursor()

    lineNum = 0

    sql = 'delete from price_report where SUBSTR(P_DATE, 0, 10) = \'' + dataTime + '\''
    c.execute(sql)
    conn.commit()

    for line in reader:
        if lineNum > 0:
            values = []
            valueStr = ''
            for cell in line:
                values.append(cell)
                if len(valueStr) == 0:
                    valueStr += ':1'
                else:
                    valueStr += ',:1'
            values[1] = int(round(float(values[1]), -2))
            c.execute('insert into price_report values(' + valueStr + ')', values)
        lineNum += 1
        print lineNum
    conn.commit()
    c.close()
    conn.close()
    print 'upload File: ' + readFilePath + " Success!"

if __name__ == '__main__':
    if len(sys.argv) == 1:
        localtime = time.strftime('%Y-%m-%d',time.localtime(time.time()))
        downloadFile(localtime)
        uploadToOracle(localtime)
    elif len(sys.argv) == 4 and sys.argv[1] == 'full':
        startTime = datetime.datetime.strptime(sys.argv[2], '%Y-%m-%d')
        endTime = time.mktime(time.strptime(sys.argv[3], '%Y-%m-%d'))
        while time.mktime(startTime.timetuple()) <= endTime:
            dataTime = startTime.strftime('%Y-%m-%d')
            uploadToOracle(dataTime)
            uploadToOracle(dataTime)
            startTime += datetime.timedelta(days = 1)
    else:
        for i in range(1, len(sys.argv)):
            uploadToOracle(sys.argv[i])
            uploadToOracle(sys.argv[i])
