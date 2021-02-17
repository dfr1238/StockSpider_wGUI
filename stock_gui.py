import os.path as path
import os
from PySimpleGUI.PySimpleGUI import Print
import scrapy
import PySimpleGUI as sg
import configparser
import pandas as pd
import csv
from datetime import datetime
from scrapy.crawler import CrawlerProcess
from scrapy.utils.project import get_project_settings

sg.theme('DarkAmber') #設定顏色主題
sg.set_options(auto_size_buttons=True)

#全域變數
this_Year = datetime.today().year #獲取今年年份
year_List =[] #存放年份
local_Coid_CSV_List =[] #本地股號表存放
local_Coid_CSV_List_Header = [] #本地股號表標頭
season_List =['1','2','3','4'] #存放季度

#常數
config_path=os.getenv('APPDATA')+'\DSApps\StockSpider\\' #設定檔路徑
setting_ini='setting.ini'   #設定檔名稱
local_csv='local_coid.csv'  #CSV檔名稱
default_MDBNAME='theStockDB'    #預設MongoDB名稱
default_MDCDNAME='theStockData' #預設MongoDB的CD名稱
default_MDUrl='mongodb://localhost:27017'   #預設MongoDB連接Url

#setting.ini相關設定
curpath = os.path.dirname(os.path.realpath(__file__))   #目前路徑
cfgpath = os.path.join(config_path,setting_ini) #設定檔路徑
csvpath = os.path.join(config_path+local_csv)   #本地股號表路徑
conf = configparser.ConfigParser()  #創建設定檔對象

#CSV相關
dict ={'代號' : local_Coid_CSV_List} #建立空的本地股號列表
csvdf = pd.DataFrame(dict) #導入pd使用


for i in range(2000,this_Year+1): #新增從2000至今年的年份至列表中
    year_List.append('%4s' % i)

def reset_setting():#重置設定
    if not os.path.exists(config_path):
        os.makedirs(config_path)
    if(not conf.has_section('MongoDB')):
        conf.add_section('MongoDB')
    conf.set('MongoDB','MONGO_URI',default_MDUrl)
    conf.set('MongoDB','DBNAME',default_MDBNAME)
    conf.set('MongoDB','CDATANAME',default_MDCDNAME)
    conf.write(open(cfgpath, 'w'))
    sg.popup('已建立設定檔。')

def reset_csv():#重建csv檔
    csvdf.to_csv(csvpath, index=False)
    sg.popup('已建立本地股號表。')

def check_setting():#檢查設定
    if(path.exists(config_path+setting_ini)):
        Print('已檢查到設定檔。')
        conf.read(cfgpath,encoding='utf-8')
    else:
        sg.popup('未檢查到設定檔，創建中...',title='系統')
        reset_setting()


def check_local_csv():#檢查本地CSV
    if(path.exists(config_path+local_csv)):
        Print('已檢查到本地股號表。')
    else:
        sg.popup('未建立本地股號表，創建中...',title='系統')
        reset_csv()

check_setting()
check_local_csv()

#視窗設計

def set_AutoMode_Window(): #自動爬取來源
    autoMode_Layout =[
                [sg.Text('選擇股號來源')],
                [sg.Radio('從本地股號表讀入',group_id='AM_LoadMode',key='loadFromLocal'),sg.Radio('從CSV檔匯入',group_id='AM_LoadMode',key='loadFromCSV')],
                [sg.Button('確定'),sg.Button('取消')],
                [sg.Text('本地股報表位於：\n'+csvpath)]
                     ]
    return sg.Window("選擇資料來源",autoMode_Layout,margins=(20,10),finalize=True,modal=True,disable_close=True,disable_minimize=True)

def set_Main_Window(): #主視窗
    main_Layout = [ 
                [sg.Text('資料庫')],
                [sg.Button('顯示資料庫資料')],
                [sg.Text('（自動＼手動）抓取設定')],
                [sg.Combo(year_List, size=(6,5), key='SearchYear',default_value=this_Year),sg.Text('年'),sg.Combo(season_List, size=(2,5), key='SearchSeason'),sg.Text('季度'),sg.Text('（手動）查詢股號'),sg.Input(key='Manual_coid',size=(10,5))],
                [sg.Text('爬取模式'),sg.Radio('自動',group_id='SMode',default=True,key='Auto'),sg.Radio('手動',group_id='SMode',key='Manual')],
                [sg.Button('開始爬取')],
                [sg.Text('運行計算式')],
                [sg.Button('公式一'),sg.Button('公式二'),sg.Button('公式三'),sg.Button('公式四')],
                [sg.Text('其他選項')],
                [sg.Button('編輯本地股號表'),sg.Button('設定'),sg.Button('離開'),sg.Button('關於'),sg.Button('說明')] 
                    ]
    return sg.Window("股票資料抓取與運算", main_Layout, margins=(40,20), finalize=True)

def set_Local_CSV_Window(): #編輯本地股號表
    csvdf = pd.read_csv(csvpath, sep=',', engine='python', header=None)
    local_Coid_CSV_List = csvdf.values.tolist()
    local_Coid_CSV_List_Header = csvdf.iloc[0].tolist()
    local_Coid_CSV_List = csvdf[1:].values.tolist()
    local_Coid_CSV_Layout =[
        [sg.Table(values=local_Coid_CSV_List,
        headings=local_Coid_CSV_List_Header,
        auto_size_columns=False,
        display_row_numbers=False,
        num_rows=min(25,len(local_Coid_CSV_List)))],
        [sg.Button('關閉且「不保存」變更'),sg.Button('關閉且「保存」變更'),sg.Button('保存當前變更'),sg.Button('重新整理'),
        sg.Button('匯入外部股號表'),sg.Button('重置本地股號表')],
        [sg.Text(f'本地股號表CSV位於{csvpath}')]
            ]
    return sg.Window("編輯本地股號表",local_Coid_CSV_Layout,grab_anywhere=False, finalize=True)

def set_Setting_Window(): #主視窗 -> 設定
    setting_Layout = [
    [sg.Text(f'設定檔的路徑位於：{config_path+setting_ini}')],
    [sg.Text('MongoDB －你絕大多數不用更動這個選項，此選項區是關於資料庫連接有關與存放爬取資料的相關設定。')],
    [sg.Text('MongoDB 連結：\t'),sg.Input(default_text=(conf['MongoDB']['MONGO_URI']),size=(30,1),k='mDBUrI')],
    [sg.Text('MongoDB 資料庫名稱：\t'),sg.Input(default_text=(conf['MongoDB']['DBNAME']),size=(30,1),k='mDBName')],
    [sg.Text('MongoDB 集合名稱：\t'),sg.Input(default_text=(conf['MongoDB']['CDATANAME']),size=(30,1),k='mCDName')],
    [sg.Button('保存'),sg.Button('取消'),sg.Button('重置')]
                    ]
    
    return sg.Window("程式設定",setting_Layout, margins=(10,5),finalize=True,modal=True,disable_close=True,disable_minimize=True)

main_Window,setting_Window,aM_Window,local_Csv_Window = set_Main_Window(),None,None,None
print('主視窗載入完成。')

while True: #監控視窗回傳
    window,event, values = sg.read_all_windows()
    if window == main_Window: #主視窗
        if event in (sg.WIN_CLOSED,'離開'):
            break
        if event == "開始爬取":
            if(len(values['SearchYear']) and len(values['SearchSeason'])): #檢查是否有選擇年與季度
                if(values['Auto'] is True): #檢查是否使用自動模式
                    sg.popup('使用自動模式抓取')
                    aM_Window=set_AutoMode_Window()
                if(values['Manual'] is True): #檢查是否使用手動模式
                    if(str(values['Manual_coid']).isnumeric() and len(values['Manual_coid']) == 4): #檢查輸入的股號格式
                        sg.popup('使用手動模式')
                    else:
                        sg.popup_error('查詢股號欄位有誤！ \n請輸入正確的格式：四位數純數字',modal=True)
            else:
                sg.popup_error('請選擇年份與季度！')

        if event == "公式一":
            sg.popup('執行公式1')

        if event == "公式二":
            sg.popup('執行公式2')

        if event == "公式三":
            sg.popup('執行公式3')

        if event == "公式四":
            sg.popup('執行公式4')

        if event == "設定":
            setting_Window=set_Setting_Window()

        if event == "編輯本地股號表":
            sg.popup('打開編輯本地股號表')
            local_Csv_Window=set_Local_CSV_Window()

        if event == "關於":
            sg.popup('股票資訊爬蟲\n版本： 1.0\n作者：Douggy Sans\n2021年編寫',title='關於')
    
    if window == setting_Window: #主視窗 -> 設定視窗之互動
        if event in (sg.WIN_CLOSED,'取消'):
            window.close()
            setting_Window=None
        
        if event == "保存":
            conf.set('MongoDB','MONGO_URI',str(values['mDBUrI']))
            conf.set('MongoDB','DBNAME',str(values['mDBName']))
            conf.set('MongoDB','CDATANAME',str(values['mCDName']))
            sg.popup('已保存設定！',title='已保存')
            window.close()
            setting_Window=None

        if event == "重置":
            if(sg.popup_ok_cancel('是否重置設定？',title='確認重置',modal=True) == 'OK'):
                reset_setting()
                sg.popup('已重置！')
                window.close()
                setting_Window=None
    
    if window == aM_Window: #主視窗 -> 自動抓取 -> 選擇來源
        if event in (sg.WIN_CLOSED,'取消'):
            window.close()
            aM_Window=None
    
    if window == local_Csv_Window: # 主視窗 -> 編輯本地股號表
        if event in (sg.WIN_CLOSED,'關閉且「不保存」變更'):
            window.close()
            local_Csv_Window=None
        if event == "重新整理":
            window.close()
            local_Csv_Window=None
            local_Csv_Window=set_Local_CSV_Window()
    
window.close()