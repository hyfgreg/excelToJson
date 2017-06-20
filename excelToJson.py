import xlrd
from datetime import datetime,date,time
import pymongo

Table = 'BusTimetable'+datetime.today().strftime('%Y-%m-%d')
client = pymongo.MongoClient('localhost')
db = client['yidong']

def readData(filename):
    data = xlrd.open_workbook(filename)
    return data

def parseTime(data):
    return time(*data[3:]).strftime('%H:%M:%S')

def parseJiading(data,uid):

    lineID = '121'
    routeSeqUp = '1211'
    routeSeqDown = '1212'

    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols

    timetable = []

    nameRowUp = 22
    nameRowDown = 1

    rowsUp = [i for i in range(23, 41)]
    rowsDown = [i for i in range(2,21)]

    nameUp = table.row_values(nameRowUp)
    nameDown = table.row_values(nameRowDown)

    #处理上行的
    round = 1
    for i in rowsUp:
        mid = 1
        for j in range(ncols):
            t = xlrd.xldate_as_tuple(table.cell(i,j).value,data.datemode)
            node = {
                'UID':uid,
                'MID':str(mid),
                'DeptTime':parseTime(t),
                'RoundNum':lineID+str(round),
                'LineID':lineID,
                'RouteSeq':routeSeqUp,
                'StationName':nameUp[j]
            }
            mid += 1
            uid += 1
            timetable.append(node)
            # print(node)
        round += 1

    #处理下行的
    round = 1
    for i in rowsDown:
        mid = 1
        for j in range(ncols):
            t = xlrd.xldate_as_tuple(table.cell(i, j).value, data.datemode)
            node = {
                'UID':str(uid),
                'MID': str(mid),
                'DeptTime': parseTime(t),
                'RoundNum': lineID + str(round),
                'LineID': lineID,
                'RouteSeq': routeSeqDown,
                'StationName': nameDown[j]
            }
            mid += 1
            uid += 1
            timetable.append(node)
            #rint(node)
        round += 1

    return timetable,uid

def parseZhongshan(data,uid):
    lineID = '101'
    routeSeqUp = '1011'
    routeSeqDown = '1012'

    table = data.sheets()[1]
    nrows = table.nrows
    ncols = table.ncols

    timetable = []

    nameRowUp = 1
    nameRowDown = 8

    rowsUp = [i for i in range(2, 8)]
    rowsDown = [i for i in range(9, 15)]

    nameUp = table.row_values(nameRowUp)
    nameDown = table.row_values(nameRowDown)

    # 处理上行的
    round = 1
    for i in rowsUp:
        mid = 1
        for j in range(ncols):
            t = xlrd.xldate_as_tuple(table.cell(i, j).value, data.datemode)
            node = {
                'UID': uid,
                'MID': str(mid),
                'DeptTime': parseTime(t),
                'RoundNum': lineID + str(round),
                'LineID': lineID,
                'RouteSeq': routeSeqUp,
                'StationName': nameUp[j]
            }
            mid += 1
            uid += 1
            timetable.append(node)
            # print(node)
        round += 1

    # 处理下行的
    round = 1
    for i in rowsDown:
        mid = 1
        for j in range(ncols):
            t = xlrd.xldate_as_tuple(table.cell(i, j).value, data.datemode)
            node = {
                'UID': str(uid),
                'MID': str(mid),
                'DeptTime': parseTime(t),
                'RoundNum': lineID + str(round),
                'LineID': lineID,
                'RouteSeq': routeSeqDown,
                'StationName': nameDown[j]
            }
            mid += 1
            uid += 1
            timetable.append(node)
            # rint(node)
        round += 1

    return timetable, uid

def parseData(data,sheetNum,uid,lineID,nameRowUp,nameRowDown,rowsUp,rowsDown):
    routeSeqUp = lineID+'1'
    routeSeqDown = lineID+'2'

    table = data.sheets()[sheetNum]
    nrows = table.nrows
    ncols = table.ncols
    #指定某列的行数：

    #len(sheet.col_values(XXXX))
    #指定某行的列数：
    upName = table.row_values(nameRowUp)
    downName = table.row_values(nameRowDown)

    while '' in upName:
        upName.remove('')

    while '' in downName:
        downName.remove('')

    ncolsUp = len(upName)
    ncolsDown = len(downName)


    timetable = []

    # rowsUp = [i for i in range(2, 8)]
    # rowsDown = [i for i in range(9, 15)]

    nameUp = table.row_values(nameRowUp)
    nameDown = table.row_values(nameRowDown)

    # 处理上行的
    round = 1
    for i in rowsUp:
        mid = 1
        for j in range(ncolsUp):
            t = xlrd.xldate_as_tuple(table.cell(i, j).value, data.datemode)
            node = {
                'UID': uid,
                'MID': str(mid),
                'DeptTime': parseTime(t),
                'RoundNum': lineID + str(round),
                'LineID': lineID,
                'RouteSeq': routeSeqUp,
                'StationName': nameUp[j]
            }
            mid += 1
            uid += 1
            timetable.append(node)
            # print(node)
        round += 1

    # 处理下行的
    round = 1
    for i in rowsDown:
        mid = 1
        for j in range(ncolsDown):
            t = xlrd.xldate_as_tuple(table.cell(i, j).value, data.datemode)
            node = {
                'UID': str(uid),
                'MID': str(mid),
                'DeptTime': parseTime(t),
                'RoundNum': lineID + str(round),
                'LineID': lineID,
                'RouteSeq': routeSeqDown,
                'StationName': nameDown[j]
            }
            mid += 1
            uid += 1
            timetable.append(node)
            # rint(node)
        round += 1

    return timetable, uid

def dataToFinalFormat(data):
    return {'network_timetable_edbus':data}

def dataToMongo(data,Table):
    if Table and db[Table].insert(data):
        print('保存成功')
    else:
        print('保存失败')

if __name__ == '__main__':
    filename = '专线时刻表20170613.xlsx'
    data = readData(filename)


    #parseData(data, sheetNum, uid, lineID, nameRowUp, nameRowDown, rowsUp, rowsDown):




    # timetable,uid1=parseJiading(data,1)
    slData = [data,0,1,'121',22,1,[i for i in range(23, 41)],[i for i in range(2, 21)]]
    timetable1,uid1 = parseData(*slData)

    zsData = [data,1,uid1,'101',1,8,[i for i in range(2, 8)],[i for i in range(9, 15)]]
    timetable2, uid2 = parseData(*zsData)
    # timetable2,uid2 = parseZhongshan(data,uid1)

    jdbData = [data,2,uid2,'41',1,10,[i for i in range(2,10)],[i for i in range(11,19)]]
    timetable3, uid3 = parseData(*jdbData)

    hqData = [data,3,uid3,'21',1,28,[i for i in range(2,26)],[i for i in range(29,52)]]
    timetable4,uid4 = parseData(*hqData)

    timetable = []
    timetable.extend(timetable1)
    timetable.extend(timetable2)
    timetable.extend(timetable3)
    timetable.extend(timetable4)


    rowData = dataToFinalFormat(timetable)
    # print(rowData)

    # print(timetable2==timetable2t)
    # print(timetable2t)
    # print(len(timetable2t))
    dataToMongo(rowData,Table)


    # print(dataToFinalFormat(timetable))

    # rowData = parseJiading(data)
    # for item in rowData[1:21]:
    #     print(item)