# -*- coding: utf-8 -*-

import sys
import time
from openpyxl.reader.excel import load_workbook
from openpyxl.writer.excel import ExcelWriter
import xlsxFormatSetting as settings
import xlsxUtil
from xlsxUtil import *

reload(sys)
sys.setdefaultencoding('utf8')
reload(sys)

def markValue(series,date):
    filePath = settings.filePath+'\\report\\'+series.encode('gbk')+'-报价日报_'.encode('gbk')+\
               date+'.xlsx'
    print filePath

    wb = load_workbook(filename=filePath)
    sheet = wb.get_sheet_by_name('报表说明')
    util = XlsxUtil(sheet)

    colNum = 2

    markValues = settings.markValue[series]

    sheet = wb.get_sheet_by_name('网站报价均值分析')
    util = XlsxUtil(sheet)
    util.singalAnalysisOfNetworkOffer(series, colNum, markValues, 'all')
    util.singalAnalysisOfNetworkOffer(series, colNum, markValues, 'null')
    util.singalAnalysisOfNetworkOffer(series, colNum, markValues, 'all_autohome')
    util.singalAnalysisOfNetworkOffer(series, colNum, markValues, 'null_autohome')

    sheet = wb.get_sheet_by_name('大区报价分析')
    util = XlsxUtil(sheet)
    util.singalAnalysisOfAreaPrice(series, colNum, markValues, 'all')
    util.singalAnalysisOfAreaPrice(series, colNum, markValues, 'all_autohome')

    sheet = wb.get_sheet_by_name('省份报价分析')
    util = XlsxUtil(sheet)
    util.singalAnalysisOfProvincesOffer(series, colNum, markValues, 'all')
    util.singalAnalysisOfProvincesOffer(series, colNum, markValues, 'all_autohome')

    print '保存文件中...'.encode('gbk')
    wb.save(filePath)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        date = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    elif len(sys.argv) == 2:
        date = sys.argv[1]
    for series in settings.carseries:
        print series.decode('UTF8').encode('gbk')
        markValue(series,date)
        print '填色'.encode('gbk')

