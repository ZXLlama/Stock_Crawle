import xml.etree.ElementTree as ET
import requests
import time
from openpyxl import Workbook
def fs(sheet,data,row): #填寫工作表中的內容(fill sheet)
    for column, value in enumerate(data,1): #重複讀取資料
        sheet.cell(row = row,column = column,value = value) #將資料放置在行(row)和列(column)上，其格子填寫value值
def rd(startYear,startMonth,endYear,endMonth,day = "01"): #returnStrDateList 組合日期字串
    result = []
    if startYear == endYear:
        for month in range(startMonth,endMonth+1):
            month = str(month)
            if len(month) == 1:
                month = "0" + month
            result.append(str(startYear),month,day)
        return result
    for year in range(startYear,endYear+1):
        if year == startYear: #當年份=起始年份時的月份
            for month in range(startMonth,13):
                month = str(month)
                if len(month) == 1:
                    month = "0" + month
                result.append(str(year)+month+day)

        elif year == endYear: #當年份=結束年份時的月份
            for month in range(1,endMonth+1):
                month = str(month)
                if len(month) == 1:
                    month = "0" + month
                result.append(str(year)+month+day)

        else: #當年份=除起始年份及結束年分時的月份
            for month in range(1,13):
                month = str(month)
                if len(month) == 1:
                    month = "0" + month
                result.append(str(year)+month+day)

    return result
tree = ET.parse('data.xml') #解析XML檔案1
root = tree.getroot() #解析XML檔案2
paramsDict = {} #建立空字典(主要使用之字典)
for child in root: # 迭代每個元素
    tag_name = child.tag #獲取tag名稱
    tag_value = child.text #獲取tag值 
    paramsDict[tag_name] = tag_value  #儲存至字典
Fields = ["日期","成交股數","成交金額","開盤價","最高價","最低價","收盤價","漲跌價差","成交筆數"] #每列標題
wb = Workbook() #建立Excel檔案
sheet = wb.active
sheet.title = "Fields" #工作表名稱
fs(sheet,Fields,1) #在第一行寫入列標題
startYear,startMonth = int(paramsDict["startYear"]),int(paramsDict["startMonth"]) #讀取字典中的起始日期並轉換為正整數
endtYear,endMonth = int(paramsDict["endYear"]),int(paramsDict["endMonth"]) #讀取字典中的結束日期並轉換為正整數
yl = rd(startYear,startMonth,endtYear,endMonth) #年份清單(YearList)
row = 2 #從第二行開始
url = 'https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY' #台灣證交所網址前綴
for YearMonth in yl: #重複抓取資料
    rq = rq = requests.get(url,params={
    "response":"json",
    "date":YearMonth,
    "stockNo":paramsDict["stockNo"], #股票代碼(亞泥1102)
})
    jsonData = rq.json() #把資料轉換成json檔
    dailyPriceList = jsonData.get("data",[])
    for dailyPrice in dailyPriceList:
        fs(sheet,dailyPrice,row)
        row += 1 #換行
        print(row)
    time.sleep(0.25) #等一秒避免被拒絕訪問
name = paramsDict["excelName"]
wb.save(name + ".xlsx")