# -*- coding: utf-8 -*-

import os
import sys
from openpyxl.styles import *
from style import Style
import xlsxFormatSetting as settings

os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
reload(sys)
sys.setdefaultencoding('utf8')
reload(sys)

class XlsxUtil:
    def __init__(self, sheet):
        self.sheet = sheet

    def createReportInstruction(self, dateList, colNum, rowNum):
        colHeight = {'title': 21, 'thead': 19.5, 'content': 35.25}

        thead = ['分析报表', '分析时间段', '筛选条件', '分析目的']
        self.mergeRowCell('报表使用说明', colNum, rowNum, len(thead), colHeight['title'], Style('标题'))

        rowNum += 1
        self.setRow(colNum, rowNum, thead, Style('行标题'), colHeight['thead'])

        if len(dateList) == 1:
            timeStr = dateList[0].strftime('%m-%d')
        else:
            timeStr = dateList[0].strftime('%m-%d') + '~' + dateList[len(dateList)-1].strftime('%m-%d')

        rowNum += 1
        row = ['网站经销商数据变化分析', timeStr, '无', '分析网站报价经销商数变化情况，发现数量变化突跳']
        self.setRow(colNum, rowNum, row, Style('内容'), colHeight['content'])

        rowNum += 1
        row = ['车型网站报价均值分析', timeStr, '无', '分析网站之间价格差异']
        self.setRow(colNum, rowNum, row, Style('内容'), colHeight['content'])

        rowNum += 1
        row = ['大区报价分析', timeStr, '4S店,直营店,卫星店', '分析大区之间的异常报价与显著差异']
        self.setRow(colNum, rowNum, row, Style('内容'), colHeight['content'])

        rowNum += 1
        row = ['省份报价分析', timeStr, '4S店,直营店,卫星店', '分析大区下省份之间的异常报价与显著差异']
        self.setRow(colNum, rowNum, row, Style('内容'), colHeight['content'])

        rowNum += 1
        row = ['报价信息', timeStr, '4S店,直营店,卫星店', '一周内的报价详细信息']
        self.setRow(colNum, rowNum, row, Style('内容'), colHeight['content'])

        return rowNum+1

    def createReportParameterSetting(self, conn, series, dateList, colNum, rowNum):
        colHeight = {'title': 24.75, 'thead': 25.5, 'content': 25.5}

        thead = ['范围', '车型', '平均值 网络报价', '平均值 MSRP', '平均值 差价', '标杆值', '报警原则']
        self.mergeRowCell('报表参数设置【（参数-报价）应用于非大区分析报表 - 突出显示】', colNum, rowNum,
                          len(thead), colHeight['title'], Style('标题'))

        rowNum += 1
        self.setRow(colNum, rowNum, thead, Style(), colHeight['thead'])

        #最后两个单元格加粗,字体改为宋体
        self.setCellBold(self.sheet[self.parseCellID(colNum+len(thead)-2, rowNum)])
        self.setCellSongTi(self.sheet[self.parseCellID(colNum+len(thead)-2, rowNum)])
        self.setCellBold(self.sheet[self.parseCellID(colNum+len(thead)-1, rowNum)])
        self.setCellSongTi(self.sheet[self.parseCellID(colNum+len(thead)-1, rowNum)])
        #标杆值一栏改变颜色
        self.setCellBgColor(self.sheet[self.parseCellID(colNum+len(thead)-2, rowNum)], 'FFC000')

        dateList[0].strftime('%m-%d')
        values = []
        sqlTimeStr = ""
        for i in range(len(dateList)):
            if i == 0:
                sqlTimeStr += ' and ( '
            else:
                sqlTimeStr += ' or '
            values.append(dateList[i].strftime('%Y-%m-%d'))
            sqlTimeStr += " crawl_date = :1 "
        sqlTimeStr += ' ) '
        values.append(series)


        cursor = conn.cursor()
        #result_all 查询4S店,直营店,卫星店的信息
        sql = 'select CARTYPE_NAME, avg(PRICE) PRICE, avg(MSRP) MSRP, avg(msrp)-avg(price) DIFFERENTIAL from price_report ' \
              'where IS_4S != \'#N/A\' ' + sqlTimeStr + ' and carseries_name = :1 group by CARTYPE_NAME'
        cursor.execute(sql, values)
        result_all = cursor.fetchall()
        #result_all 查询未匹配经销商名称的信息
        sql = 'select CARTYPE_NAME, avg(PRICE) PRICE, avg(MSRP) MSRP, avg(msrp)-avg(price) DIFFERENTIAL from price_report ' \
              'where IS_4S = \'#N/A\' ' + sqlTimeStr + ' and carseries_name = :1  group by CARTYPE_NAME'
        cursor.execute(sql, values)
        result_NA = cursor.fetchall()
        cursor.close()

        rowNum += 1
        if len(result_all) + len(result_NA) <= 0:
            return -1
        else:
            self.mergeColCell('MSRP减报价均值,低于标杆值报红色预警', colNum+len(thead)-1, rowNum,
                              len(result_all)+len(result_NA), Style())

            if len(result_all) > 0:
                self.mergeColCell('4S店,直营店,卫星店', colNum, rowNum, len(result_all), Style())
                for row in result_all:
                    rowValues = []
                    for cell in row:
                        rowValues.append(cell)
                    if settings.markValue.has_key(series) and settings.markValue[series].has_key(rowValues[0]):
                        rowValues.append(settings.markValue[series][rowValues[0]])
                    else:
                        rowValues.append(settings.markValue['默认值'])
                    self.setRow(colNum+1, rowNum, rowValues, Style(), colHeight['content'])
                    #标杆值改背景色
                    self.setCellBgColor(self.sheet[self.parseCellID(colNum+1+len(row), rowNum)], 'FFC000')
                    rowNum += 1

            if len(result_NA) > 0:
                self.mergeColCell('未匹配经销商名称', colNum, rowNum, len(result_NA), Style())
                for row in result_NA:
                    rowValues = []
                    for cell in row:
                        rowValues.append(cell)
                    rowValues.append(2000)
                    self.setRow(colNum+1, rowNum, rowValues, Style(), colHeight['content'])
                    self.setCellBgColor(self.sheet[self.parseCellID(colNum+1+len(row), rowNum)], 'FFC000')
                    rowNum += 1
        return rowNum

    def createReportsAnalysisOfWebsiteDealer(self, conn, series, dateList, colNum, rowNum, type):
        colHeight = {'title': 21, 'thead': 22.5, 'content': 22.5}

        if type == 'all':
            titleType = '4S店,直营店,卫星店'
        else:
            titleType = '未匹配经销商名称'
        self.mergeRowCell(titleType+' - 网站经销商数量分析', colNum, rowNum, len(dateList)+2, colHeight['title'], Style('标题'))

        rowNum += 1
        self.mergeRowCell('报价日期 日期', colNum+1, rowNum, len(dateList)+1, colHeight['title'], Style('行标题2'))

        rowNum += 1
        thead = ['网站来源']
        for cell in dateList:
            thead.append(cell.strftime('%m月%d日'))
        thead.append('总计')
        for i in range(9-len(thead)):
            thead.append('')
        self.setRow(colNum, rowNum, thead, Style(), colHeight['thead'])

        cursor = conn.cursor()
        values = [series]
        wheresql = ' where carseries_name = :1 '
        if type == 'all':
            wheresql += ' and IS_4S != \'#N/A\' '
        else:
            wheresql += ' and IS_4S = \'#N/A\' '
        wheresql += ' and ( '
        for i in range(len(dateList)):
            if i != 0:
                wheresql += ' or '
            values.append(dateList[i].strftime('%Y-%m-%d'))
            wheresql += ' crawl_date = :1 '
        wheresql += ' ) '
        sql = 'select SOURCE, crawl_date, COUNT(*) COUNT ' \
              'from ( select distinct SOURCE, crawl_date, NVL(dealerName,agency_name) from price_report ' + wheresql + \
              ' ) group by SOURCE, crawl_date'
        cursor.execute(sql, values)
        result = cursor.fetchall()
        cursor.close()
        if len(result) <= 0:
            return -1
        resultMap = {}
        for cell in result:
            resultMap[cell[0] + ' ' + cell[1]] = cell[2]

        rowNum += 1
        length = 0
        for rowValue in settings.websiteDealer:
            row = [rowValue]
            hasData = False
            for colValue in dateList:
                key = rowValue + ' ' + colValue.strftime('%Y-%m-%d')
                if resultMap.has_key(key):
                    row.append(resultMap[key])
                    hasData = True
                else:
                    row.append('')
            if hasData:
                row.append('=IF(ISERR(AVERAGE(' + self.parseCellID(colNum+1, rowNum) + ':' +
                           self.parseCellID(colNum+len(row)-1, rowNum) + ')), ,AVERAGE(' +
                           self.parseCellID(colNum+1, rowNum) + ':' +
                           self.parseCellID(colNum+len(row)-1, rowNum) + '))')
                self.setRow(colNum, rowNum, row, Style(), colHeight['content'])
                rowNum += 1
                length += 1
        row = ['总计']
        for i in range(len(dateList)+1):
            row.append('=IF(ISERR(AVERAGE(' + self.parseCellID(colNum+i+1, rowNum-length) + ':' +
                        self.parseCellID(colNum+i+1, rowNum-1) + ')), ,AVERAGE(' +
                        self.parseCellID(colNum+i+1, rowNum-length) + ':' +
                        self.parseCellID(colNum+i+1, rowNum-1) + '))')
        self.setRow(colNum, rowNum, row, Style(), colHeight['content'])

        return rowNum

    def createAnalysisOfNetworkOffer(self, conn, series, dateList, colNum, rowNum, type):
        colHeight = {'title': 21, 'thead': 13.5, 'content': 13.5}

        if type == 'all':
            titleType = '4S店,直营店,卫星店'
            title_end = ''
        elif type == 'null':
            titleType = '未匹配经销商名称'
            title_end = ''
        elif type == 'all_autohome':
            titleType = '4S店,直营店,卫星店'
            title_end = '(汽车之家与易车)'
        else:
            titleType = '未匹配经销商名称'
            title_end = '(汽车之家与易车)'
        self.mergeRowCell(titleType+' - 网站报价均值分析'+title_end, colNum, rowNum,
                          len(dateList)+3, colHeight['title'], Style('标题'))

        rowNum += 1
        self.mergeRowCell('报价日期 日期', colNum+2, rowNum, len(dateList)+1, colHeight['title'], Style('行标题2'))

        rowNum += 1
        thead = ['车型', '网站来源']
        for i in range(len(dateList)):
            thead.append(dateList[i].strftime('%m月%d日'))
        thead.append('总计')
        self.setRow(colNum, rowNum, thead, Style(), colHeight['thead'])

        cursor = conn.cursor()
        values = [series]
        wheresql = ' where carseries_name = :1 '
        if type == 'all' or type == 'all_autohome':
            wheresql += ' and IS_4S != \'#N/A\' '
        else:
            wheresql += ' and IS_4S = \'#N/A\' '
        wheresql += ' and ( '
        for i in range(len(dateList)):
            if i != 0:
                wheresql += ' or '
            values.append(dateList[i].strftime('%Y-%m-%d'))
            wheresql += ' crawl_date = :1 '
        wheresql += ' ) '
        if type == 'all_autohome' or type == 'null_autohome':
            wheresql += ' and ( SOURCE like \'汽车之家%\' or  SOURCE like \'易车%\' ) '

        sql = 'select cartype_name, source, crawl_date, avg(msrp)-avg(price) differential ' \
              'from ( select cartype_name, source, price, msrp, crawl_date ' \
                        'from price_report ' + wheresql + ' ) group by cartype_name, crawl_date, source'
        cursor.execute(sql, values)
        result = cursor.fetchall()

        sql = 'select cartype_name, source, avg(msrp)-avg(price) differential ' \
              'from ( select cartype_name, source, price, msrp ' \
                        'from price_report ' + wheresql + ' ) group by cartype_name, source'
        cursor.execute(sql, values)
        resultAll = cursor.fetchall()
        cursor.close()
        if len(result) <= 0:
            return -1
        resultMap = {}
        for cell in result:
            if resultMap.has_key(cell[0]):
                (resultMap[cell[0]])[cell[1]+' '+cell[2]] = cell[3]
            else:
                resultMap[cell[0]] = {}
                (resultMap[cell[0]])[cell[1]+' '+cell[2]] = cell[3]
        resultAllMap = {}
        for cell in resultAll:
            if resultAllMap.has_key(cell[0]):
                (resultAllMap[cell[0]])[cell[1]] = cell[2]
            else:
                resultAllMap[cell[0]] = {}
                (resultAllMap[cell[0]])[cell[1]] = cell[2]

        rowNum += 1
        for category in resultMap:
            categoryMap = resultMap[category]
            length = 0
            for rowValue in settings.websiteDealer:
                row = [rowValue]
                hasData = False
                for colValue in dateList:
                    categoryMapKey = rowValue + ' ' + colValue.strftime('%Y-%m-%d')
                    if categoryMap.has_key(categoryMapKey):
                        row.append(categoryMap[categoryMapKey])
                        hasData = True
                    else:
                        row.append('')
                if hasData:
                    row.append(resultAllMap[category][rowValue])
                    self.setRow(colNum+1, rowNum, row, Style(), colHeight['content'])
                    length += 1
                    rowNum += 1
            self.mergeColCell(category, colNum, rowNum-length, length, Style())

        return rowNum

    def createAnalysisOfAreaPrice(self, conn, series, dateList, colNum, rowNum, type):
        colHeight = {'title': 21, 'thead': 13.5, 'content': 13.5}

        if type == 'all':
            title_end = ''
        elif type == 'all_autohome':
            title_end = '(汽车之家与易车)'
        self.mergeRowCell('销售大区报价分析'+title_end, colNum, rowNum, len(dateList)+3, colHeight['title'], Style('标题'))

        rowNum += 1
        self.mergeRowCell('报价日期 日期', colNum+2, rowNum, len(dateList)+1, colHeight['title'], Style('行标题2'))

        rowNum += 1
        thead = ['车型', '销售大区']
        for i in range(len(dateList)):
            thead.append(dateList[i].strftime('%m月%d日'))
        thead.append('总计')
        self.setRow(colNum, rowNum, thead, Style(), colHeight['thead'])

        cursor = conn.cursor()
        values = [series]
        wheresql = ' where carseries_name = :1 '
        wheresql += ' and ( '
        for i in range(len(dateList)):
            if i != 0:
                wheresql += ' or '
            values.append(dateList[i].strftime('%Y-%m-%d'))
            wheresql += ' crawl_date = :1 '
        wheresql += ' ) '
        if type == 'all_autohome':
            wheresql += ' and ( SOURCE like \'汽车之家%\' or  SOURCE like \'易车%\' ) '
        sql = 'select cartype_name, sales_area, crawl_date, avg(msrp)-avg(price) differential ' \
              'from ( select cartype_name, sales_area, price, msrp, crawl_date ' \
                        'from price_report ' + wheresql + ' ) group by cartype_name, crawl_date, sales_area'
        cursor.execute(sql, values)
        result = cursor.fetchall()
        sql = 'select cartype_name, sales_area, avg(msrp)-avg(price) differential ' \
              'from ( select cartype_name, sales_area, price, msrp ' \
                        'from price_report ' + wheresql + ' ) group by cartype_name, sales_area'
        cursor.execute(sql, values)
        resultAll = cursor.fetchall()
        cursor.close()
        if len(result) <= 0:
            return -1
        resultMap = {}
        for cell in result:
            if resultMap.has_key(cell[0]):
                (resultMap[cell[0]])[cell[1]+' '+cell[2]] = cell[3]
            else:
                resultMap[cell[0]] = {}
                (resultMap[cell[0]])[cell[1]+' '+cell[2]] = cell[3]
        resultAllMap = {}
        for cell in resultAll:
            if resultAllMap.has_key(cell[0]):
                (resultAllMap[cell[0]])[cell[1]] = cell[2]
            else:
                resultAllMap[cell[0]] = {}
                (resultAllMap[cell[0]])[cell[1]] = cell[2]

        rowNum += 1
        for category in resultMap:
            categoryMap = resultMap[category]
            #一共有雪佛兰1~8区和无大区分类数据，这里需要且只需要雪佛兰1~8区的数据
            self.mergeColCell(category, colNum, rowNum, 8, Style())
            for rowValue in settings.area:
                row = [rowValue]
                for colValue in dateList:
                    categoryMapKey = rowValue + ' ' + colValue.strftime('%Y-%m-%d')
                    if categoryMap.has_key(categoryMapKey):
                        row.append(categoryMap[categoryMapKey])
                    else:
                        row.append('')
                if len(row) > 1:
                    row.append(resultAllMap[category][rowValue])
                    self.setRow(colNum+1, rowNum, row, Style(), colHeight['content'])
                    rowNum += 1

        return rowNum

    def createTableAnalysisOfProvincesOffer(self, conn, series, dateList, colNum, rowNum, type):
        cursor = conn.cursor()
        values = [series]
        wheresql = ' where carseries_name = :1 '
        wheresql += ' and ( '
        for i in range(len(dateList)):
            if i != 0:
                wheresql += ' or '
            values.append(dateList[i].strftime('%Y-%m-%d'))
            wheresql += ' crawl_date = :1 '
        wheresql += ' ) '
        if type == 'all_autohome':
            wheresql += ' and ( SOURCE like \'汽车之家%\' or  SOURCE like \'易车%\' ) '
        sql = 'select DISTINCT cartype_name from price_report ' + wheresql
        cursor.execute(sql, values)
        result = cursor.fetchall()
        if len(result) <= 0:
            return -1
        cartypeList = []
        for i in range(len(result)):
            cartypeList.append(result[i][0])

        colHeight = {'title': 21, 'thead': 20.25, 'content': 13.5}

        if type == 'all':
            title_end = ''
        elif type == 'all_autohome':
            title_end = '(汽车之家与易车)'
        self.mergeRowCell('销售大区报价分析'+title_end, colNum, rowNum,
                          len(cartypeList)+2, colHeight['title'], Style('标题'))

        rowNum += 1
        self.mergeRowCell('车型', colNum+2, rowNum, len(cartypeList), colHeight['title'], Style('行标题2'))

        rowNum += 1
        thead = ['销售大区', '省份']
        for i in range(len(cartypeList)):
            thead.append(cartypeList[i])
        self.setRow(colNum, rowNum, thead, Style(), colHeight['thead'])

        sql = 'select sales_area, province, cartype_name, avg(msrp)-avg(price) differential ' \
              'from ( select cartype_name, sales_area, province, price, msrp, crawl_date ' \
                        'from price_report ' + wheresql + ' ) group by cartype_name, sales_area, province'
        cursor.execute(sql, values)
        result = cursor.fetchall()
        cursor.close()
        if len(result) <= 0:
            return -1
        resultMap = {}
        for cell in result:
            if resultMap.has_key(cell[0]):
                (resultMap[cell[0]])[cell[1]+' '+cell[2]] = cell[3]
            else:
                resultMap[cell[0]] = {}
                (resultMap[cell[0]])[cell[1]+' '+cell[2]] = cell[3]

        rowNum += 1
        for category in settings.area:
            categoryMap = resultMap[category]
            length = 0
            for province in settings.province[category]:
                row = [province]
                for cartype in cartypeList:
                    categoryMapKey = province + ' ' + cartype
                    if categoryMap.has_key(categoryMapKey):
                        row.append(categoryMap[categoryMapKey])
                    else:
                        row.append('')
                if len(row) > 1:
                    self.setRow(colNum+1, rowNum, row, Style(), colHeight['content'])
                    length += 1
                    rowNum += 1
            self.mergeColCell(category, colNum, rowNum-length, length, Style())

        return rowNum

    def createTableAnalysisOfDetailedQuotation(self, conn, series, dateList, colNum, rowNum):
        colHeight = {'title': 13.5, 'thead': 13.5, 'content': 13.5}


        fieldnameList = settings.fieldList
        style = Style()
        style.setFont('宋体', 11, True, 'FFFFFF')
        style.setPatternFill('5B9BD5')
        self.setRow(colNum, rowNum, fieldnameList, style, colHeight['title'])

        fieldnameStr = ''
        for i in range(len(fieldnameList)):
            if len(fieldnameStr) > 0:
                fieldnameStr += ' , '
            fieldnameStr += fieldnameList[i]
        cursor = conn.cursor()
        values = [series]
        wheresql = ' where carseries_name = :1 '
        values.append(dateList[len(dateList)-1].strftime('%Y-%m-%d'))
        wheresql += ' and crawl_date = :1 '
        sql = 'select ' + fieldnameStr + ' from price_report ' + wheresql
        cursor.execute(sql, values)
        result = cursor.fetchall()
        cursor.close()
        if len(result) <= 0:
            return -1
        print '    写入'.encode('gbk')+str(len(result))+'行数据'.encode('gbk')
        odd_style = Style()
        odd_style.setFont('宋体', 11, False, '000000')
        even_style = Style()
        even_style.setFont('宋体', 11, False, '000000')
        even_style.setPatternFill('DDEBF7')

        rowNum += 1
        for row in result:
            if rowNum % 2 == 1:
                self.setRow(colNum, rowNum, row, odd_style, colHeight['content'])
                print '      第'.encode('gbk')+str(rowNum-1)+'数据'.encode('gbk')
            else:
                self.setRow(colNum, rowNum, row, even_style, colHeight['content'])
                print '      第'.encode('gbk')+str(rowNum-1)+'数据'.encode('gbk')
            rowNum += 1
        return rowNum

    def findMarkValue(self, colNum):
        rowNum = self.findCellRowNumByValue(colNum, '报表参数设置【（参数-报价）应用于非大区分析报表 - 突出显示】')

        if rowNum != -1:
            #标题的第二行开始是数据
            rowNum += 2
            markValues = {'IS4S': {}, 'NOT4S': {}}
            while rowNum <= self.sheet.max_row:
                if self.sheet[self.parseCellID(colNum, rowNum)].value == '4S店,直营店,卫星店':
                    cartype = self.sheet[self.parseCellID(colNum+1, rowNum)].value
                    markValue = self.sheet[self.parseCellID(colNum+5, rowNum)].value
                    markValues['IS4S'][cartype] = markValue
                elif self.sheet[self.parseCellID(colNum, rowNum)].value == '未匹配经销商名称':
                    cartype = self.sheet[self.parseCellID(colNum+1, rowNum)].value
                    markValue = self.sheet[self.parseCellID(colNum+5, rowNum)].value
                    markValues['NOT4S'][cartype] = markValue
                rowNum += 1
            return markValues
        return None

    def singalAnalysisOfNetworkOffer(self, colNum, markValues, type):
        if type == 'all':
            title = '4S店,直营店,卫星店 - 网站报价均值分析'
        elif type == 'null':
            title = '未匹配经销商名称 - 网站报价均值分析'
        elif type == 'all_autohome':
            title = '4S店,直营店,卫星店 - 网站报价均值分析(汽车之家与易车)'
        else:
            title = '未匹配经销商名称 - 网站报价均值分析（汽车之家与易车）'
        rowNum = self.findCellRowNumByValue(colNum, title)

        if rowNum != -1:
            rowNum += 3
            while self.sheet[self.parseCellID(colNum, rowNum)].value is not None:
                cartype = self.sheet[self.parseCellID(colNum, rowNum)].value
                for i in range(colNum+2, self.sheet.max_column):
                    cell = self.sheet[self.parseCellID(i, rowNum)]
                    if (cell.value is not None) and cell.value < markValues[cartype]:
                        self.setCellBgColor(cell, settings.markvalueColor)
                rowNum += 1

    def singalAnalysisOfAreaPrice(self, colNum, markValues, type):
        if type == 'all':
            title = '销售大区报价分析'
        elif type == 'all_autohome':
            title = '销售大区报价分析(汽车之家与易车)'
        rowNum = self.findCellRowNumByValue(colNum, title)

        if rowNum != -1:
            rowNum += 3
            while self.sheet[self.parseCellID(colNum, rowNum)].value is not None:
                cartype = self.sheet[self.parseCellID(colNum, rowNum)].value
                for i in range(colNum+2, self.sheet.max_column):
                    cell = self.sheet[self.parseCellID(i, rowNum)]
                    if (cell.value is not None) and cell.value < markValues[cartype]:
                        self.setCellBgColor(cell, settings.markvalueColor)
                rowNum += 1

    def singalAnalysisOfProvincesOffer(self, colNum, markValues, type):
        if type == 'all':
            title = '销售大区报价分析'
        elif type == 'all_autohome':
            title = '销售大区报价分析(汽车之家与易车)'
        rowNum = self.findCellRowNumByValue(colNum, title)

        if rowNum != -1:
            rowNum += 2
            colNum += 2
            while self.sheet[self.parseCellID(colNum, rowNum)].value is not None:
                cartype = self.sheet[self.parseCellID(colNum, rowNum)].value
                for i in range(34):
                    cell = self.sheet[self.parseCellID(colNum, rowNum+1+i)]
                    if (cell.value is not None) and cell.value < markValues[cartype]:
                        self.setCellBgColor(cell, settings.markvalueColor)
                colNum += 1

    def findCellRowNumByValue(self, colNum,value):
        rowLength = self.sheet.max_row
        rowNum = -1
        for i in range(rowLength):
            if self.sheet[self.parseCellID(colNum, i+1)].value == value:
                rowNum = i+1
                break
        return rowNum

    def setRowWidth(self):
        key = self.sheet.title.encode('utf-8')
        if settings.width.has_key(key):
            width = settings.width[key]
        else:
            width = [8.38]
        for i in range(len(width)):
            self.sheet.column_dimensions[self.parseCellRowID(i+1)].width = width[i]

    def mergeRowCell(self, titleName, colNum, rowNum, length, height, style):
        self.sheet.row_dimensions[rowNum].height = height
        self.sheet.merge_cells(self.parseCellID(colNum, rowNum)+':'+self.parseCellID(colNum+(length-1), rowNum))
        for i in range(length):
            cell = self.sheet[self.parseCellID(colNum+i, rowNum)]
            self.setCell(cell, titleName, style)

    def mergeColCell(self, cellName, colNum, rowNum, length, style):
        self.sheet.merge_cells(self.parseCellID(colNum, rowNum)+':'+self.parseCellID(colNum, rowNum+(length-1)))
        for i in range(length):
            cell = self.sheet[self.parseCellID(colNum, rowNum+i)]
            self.setCell(cell, cellName, style)

    #setRow 单元格使用统一格式
    def setRow(self, colNum, rowNum, row, style, height):
        self.sheet.row_dimensions[rowNum].height = height
        for i in range(len(row)):
            cell = self.sheet[self.parseCellID(colNum+i, rowNum)]
            self.setCell(cell, row[i], style)

    def setRowBgColor(self, colNum, rowNum, row, color):
        for i in range(len(row)):
            cell = self.sheet[self.parseCellID(colNum+i, rowNum)]
            self.setCellBgColor(cell, color)

    def setCell(self, cell, cellValue, style):
        cell.value = cellValue
        cell.font = Font(name=style.font['family'],
                         size=style.font['size'],
                         bold=style.font['blod'],
                         color=style.font['color'])
        cell.fill = PatternFill(patternType=style.patternFill['patternType'],
                                start_color=style.patternFill['color'],
                                end_color=style.patternFill['color'])
        cell.alignment = Alignment(horizontal=style.alignment['horizontal'],
                                   vertical=style.alignment['vertical'],
                                   wrap_text=style.alignment['wrap_text'])
        cell.border = Border(left=Side(border_style=style.border['style'], color=style.border['color']),
                             right=Side(border_style=style.border['style'], color=style.border['color']),
                             top=Side(border_style=style.border['style'],  color=style.border['color']),
                             bottom=Side(border_style=style.border['style'], color=style.border['color']))
        cell.number_format = style.number_format['number']


    def setCellBold(self, cell):
        font = cell.font
        cell.font = Font(name=font.name,
                         size=font.size,
                         bold=True,
                         color=font.color)

    def setCellSongTi(self, cell):
        font = cell.font
        cell.font = Font(name='宋体',
                         size=font.size,
                         bold=font.bold,
                         color=font.color)

    def setCellBgColor(self, cell, color):
        cell.fill = PatternFill(patternType='solid',
                                start_color='FF'+color,
                                end_color='FF'+color)

    def parseCellID(self, colNum, rowNum):
        if isinstance(colNum, int) and isinstance(rowNum, int):
            id = ''
            while colNum >= 1:
                id = chr((colNum-1) % 26 + 65) + id
                colNum = (colNum-1) / 26
            id += str(rowNum)
            return id
        else:
            return None

    def parseCellRowID(self, colNum):
        if isinstance(colNum, int):
            id = ''
            while colNum >= 1:
                id = chr((colNum-1) % 26 + 65) + id
                colNum = (colNum-1) / 26
            return id
        else:
            return None

if __name__ == '__main__':
    util = XlsxUtil(None)
    print util.parseCellRowID(1)
