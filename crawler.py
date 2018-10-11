import requests
from bs4 import BeautifulSoup
import urllib.parse
from datetime import datetime
import time
import calendar
import xlwt

## class_氣象資料類別屬性 ##
class WeatherData:
    #def add_DateTime(self, Date):
    #    self.DateTime.append(Date)
    def __init__(self, ObsTime, StnPres, SeaPres, Temperature, Td, RH, WS, WD, WSGust, WDGust, Precp, PrecpHour, SunShine, GlobalRad, Visb): #init宣告
        #self.DateTime = [] #設定日期為變數不行宣告在全域，不然給值時會覆蓋所有created class
        self.ObsTime = ObsTime
        self.StnPres = StnPres
        self.SeaPres = SeaPres
        self.Temperature = Temperature
        self.Td = Td
        self.RH = RH
        self.WS = WS
        self.WD = WD
        self.WSGust = WSGust
        self.WDGust = WDGust
        self.Precp = Precp
        self.PrecpHour = PrecpHour
        self.SunShine = SunShine
        self.GlobalRad = GlobalRad
        self.Visb = Visb
 
## def_給年月起始與終點，輸出起始日至最後一次之所有日期(終點預設為目前時間)
def DateStrings(BeginYear, BeginMonth, EndYear = int(datetime.now().strftime("%Y")) ,EndMonth = int(datetime.now().strftime("%m"))):
    # 給定dict空間Dates，i計算共幾天供dict當key使用
    i = 0
    Dates = {}
    FMT = "%Y-%m-%d"

    # 迴圈從起始年跑到結束年    
    for Year in range(BeginYear,EndYear + 1): # Python range包頭不包尾，故+1

        # 依年份設定下個迴圈的起始月份與結束月份
        if (Year == BeginYear and Year != EndYear):            
            RunBeginMonth = BeginMonth
            RunEndMonth = 12
        elif (Year == EndYear and Year != BeginYear):
            RunBeginMonth = 1
            RunEndMonth = EndMonth
        elif (Year == EndYear and Year == BeginYear):
            RunBeginMonth = BeginMonth
            RunEndMonth = EndMonth
        else:
            RunBeginMonth = 1
            RunEndMonth = 12
        
        # 依給定月份範圍跑迴圈
        for Mon in range(RunBeginMonth, RunEndMonth + 1): # Python range包頭不包尾，故+1
            # 取得該年特定月份的天數
            StartWeek, DaysInMonth = calendar.monthrange(Year, Mon)
            for Date in range(1, DaysInMonth + 1):
                DateTime = str(Year) + "-{0:02d}".format(Mon) + "-" + "{0:02d}".format(Date)
                i=i+1                
                Dates[i] = DateTime
                 # 檢查時間差，<=0代表含當日
                tdelta = datetime.now() - datetime.strptime( DateTime, FMT )
                # 時間差<=0代表含當日，但當日資料要明天才會出來，故設定<=1
                if (tdelta.days <= 1 ): 
                    break
            if (tdelta.days <= 1 ):
                break
        if (tdelta.days <= 1 ):
            break
    return Dates 

## def_給定網址抓cwb data ##
def get_cwb_data(url):
    response = requests.get(url).text
    soup = BeautifulSoup(response, 'lxml') #把一些html標籤文字解釋成文字, ex:&nbsp會轉換成空白建
    CrawlingData = soup.find_all('td')    

    # Table標題
    #DataField=0
    #title = []
    #Title = soup.find_all('tr','second_tr')
    #for item in Title:
    #    print(item)
    #    meta = item.find('th','br')
    #    print(meta)
    #    #title.append(item.getText().strip())
    #    #print(title[DataField])    
    #    DataField = DataField + 1
    
    # 建立爬蟲資料的空間
    WeatherData_Day = []
    CrawledData = []
    
    try:
        for hr in range(0,24): # 24小時資料，Python range 包頭不包尾        
            ## 每小時有15項資料 ##
            for item in CrawlingData[ 4+15*hr+1 : 4+15*(hr+1) ]:  
                if item.getText().strip() != "" and item.getText().strip() != "X":
                    CrawledData.append(item.getText().strip())                
                else:
                    CrawledData.append("")
            # 將data寫入Array 
            WeatherData_Day.append( \
                            WeatherData(hr, CrawledData[0],CrawledData[1],CrawledData[2], CrawledData[3], \
                            CrawledData[4], CrawledData[5], CrawledData[6], CrawledData[7], \
                            CrawledData[8], CrawledData[9], CrawledData[10], CrawledData[11], \
                            CrawledData[12], CrawledData[13] ))   
            # 清空此小時爬到的資料，讓下個迴圈重新塞資料
            CrawledData.clear() 
            # print("--------------------------------------------------------------")
    except: 
        print("Error: 無法抓", DatePicker, "資料")
        for hr in range(0,24): # 24小時資料，Python range 包頭不包尾        
            ## 每小時有15項資料 ##
            for item in range (1, 15):
                CrawledData.append("")
            # 將data寫入Array
            WeatherData_Day.append( \
                            WeatherData(hr, CrawledData[0],CrawledData[1],CrawledData[2], CrawledData[3], \
                            CrawledData[4], CrawledData[5], CrawledData[6], CrawledData[7], \
                            CrawledData[8], CrawledData[9], CrawledData[10], CrawledData[11], \
                            CrawledData[12], CrawledData[13] ))
            # 清空此小時爬到的資料，讓下個迴圈重新塞資料
            CrawledData.clear() 

        return WeatherData_Day
    return WeatherData_Day

## Main Program ##
DateNow = datetime.now().strftime("%Y-%m-%d %H:%M")
StatName = "山佳"
Station = "C0A520"
Dict_DatesStrings = DateStrings( 2017, 10, 2017, 12 ) # 起始年月份

# 建立Excel #
file = xlwt.Workbook()
table = file.add_sheet(StatName, cell_overwrite_ok = True)
style = xlwt.XFStyle()    # 初始化Excel樣式
font = xlwt.Font()        # 為樣式創建字体
font.name = "Times New Roman"
font.height = 12*20       # 12*20, for 12 point
style.font = font
style.num_format_str = "0.00" # 寫入數字格式

# 開始爬資料 #
ExcelRow = 1 # 從第1列開始塞資料，第0列塞欄位名稱
for DatePicker in Dict_DatesStrings:
    Page_url = "http://e-service.cwb.gov.tw/HistoryDataQuery/DayDataController.do?command=viewMain&station=" + Station + "&stname=" + StatName + "&datepicker=" + Dict_DatesStrings[DatePicker]
    WeatherData_Day = get_cwb_data( Page_url )

    # time.sleep(0.1) # 爬太快不給爬
    print( Dict_DatesStrings[DatePicker] )
    # 24小時資料，Python range 包頭不包尾
    for hr in range(0,24): 
        # 把object轉成dictionary 
        Dict_WeatherData = vars( WeatherData_Day[hr] )   
        DataField=0
        for Key in Dict_WeatherData:
            if (Key == "ObsTime"):
                Dict_WeatherData[Key] = "{0} {1:02d}:00:00".format(Dict_DatesStrings[DatePicker], hr)
        
            table.write( ExcelRow, DataField, Dict_WeatherData[Key], style ) 
            DataField +=  1
        ExcelRow += 1


# Excel標題 #
DataField=0
for Key in Dict_WeatherData:
    table.write( 0, DataField, Key, style )
    DataField +=  1

# 輸出Excel存檔 #
SavePath= "C:\\Projects\\crawler\\" + StatName + ".xls"
file.save(SavePath)
print("Excel is Exported!")

