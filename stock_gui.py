from os import popen
import scrapy
import PySimpleGUI as sg
from datetime import datetime
from scrapy.crawler import CrawlerProcess
from scrapy.utils.project import get_project_settings

sg.theme('DarkAmber') #設定顏色主題

this_Year = datetime.today().year #獲取今年年份
year_List =[] #存放年份
season_List =['1','2','3','4'] #存放季度
for i in range(2000,this_Year+1): #新增從2000至今年的年份至列表中
    year_List.append('%4s' % i)


#視窗設計
def set_Main_Window(): #主視窗
    main_Layout = [ [sg.Text('（自動＼手動）抓取設定')],
                [sg.DropDown(year_List, size=(6,5), key='SearchYear'),sg.Text('年'),sg.DropDown(season_List, size=(2,5), key='SearchSeason'),sg.Text('季度'),sg.Text('（手動）查詢股號'),sg.InputText(key='Manual_coid',size=(10,5))],
                [sg.Text('爬取模式'),sg.Radio('自動',group_id='SMode',default=True,key='Auto'),sg.Radio('手動',group_id='SMode',key='Manual'),sg.Button('開始爬取')],
                [sg.Text('運行計算式')],
                [sg.Button('公式一'),sg.Button('公式二'),sg.Button('公式三'),sg.Button('公式四')],
                [sg.Text('其他選項')],
                [sg.Button('編輯股號表'),sg.Button('設定'),sg.Button('離開'),sg.Button('說明')] ]
    return sg.Window("股票資料抓取與運算", main_Layout, margins=(40,20), finalize=True)

def set_Setting_Window(): #主視窗 -> 設定
    setting_Layout = [
    [sg.Text('MongoDB 資料庫名稱：\t'),sg.InputText(size=(20,1))],
    [sg.Text('MongoDB 集合名稱：\t'),sg.InputText(size=(20,1))],
    [sg.Button('確定'),sg.Button('取消')]]
    return sg.Window("程式設定",setting_Layout, margins=(10,5),finalize=True,modal=True,keep_on_top=True,disable_close=True,disable_minimize=True)

main_Window,setting_Window = set_Main_Window(),None

while True: #監控視窗回傳
    window,event, values = sg.read_all_windows()
    if window == main_Window and event in (sg.WIN_CLOSED,'離開'):
        break
    
    if window == main_Window:
        if event == "開始爬取":
            if(len(values['SearchYear']) and len(values['SearchSeason'])): #檢查是否有選擇年與季度
                if(values['Auto'] is True): #檢查是否使用自動模式
                    sg.popup('使用自動模式抓取')
                if(values['Manual'] is True): #檢查是否使用手動模式
                    if(str(values['Manual_coid']).isnumeric() and len(values['Manual_coid']) == 4): #檢查輸入的股號格式
                        sg.popup('使用手動模式')
                    else:
                        sg.popup_error('查詢股號欄位有誤！ \n請輸入正確的格式：四位數純數字')
            else:
                sg.popup_error('請選擇年份與季度！')
            
    if window == main_Window:
        if event == "設定":
            setting_Window=set_Setting_Window()

    if window == setting_Window and event in (sg.WIN_CLOSED,'取消'):
        window.close()
        setting_Window=None

window.close()