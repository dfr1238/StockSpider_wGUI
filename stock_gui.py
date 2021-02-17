import PySimpleGUI as sg
from datetime import datetime

sg.theme('DarkAmber') #設定顏色主題

this_Year = datetime.today().year #獲取今年年份
year_list =[] #存放年份
season_list =[1,2,3,4] #存放季度
for i in range(2000,this_Year+1): #新增從2000至今年的年份至列表中
    year_list.append('%5d' % i)


#主視窗設計
main_Layout = [ [sg.Text('（自動＼手動）抓取設定')],
                [sg.DropDown(year_list, size=(6,5)),sg.Text('年'),sg.DropDown(season_list, size=(2,5)),sg.Text('季度')],
                [sg.Text('自動抓取資料'),sg.Button('自動爬取')],
                [sg.Text('手動輸入股號抓取資料')],
                [sg.InputText(size=(10,5)),sg.Button('單筆抓取')],
                [sg.Text('運行計算式')],
                [sg.Button('公式一'),sg.Button('公式二'),sg.Button('公式三'),sg.Button('公式四')],
                [sg.Text('其他選項')],
                [sg.Button('編輯股號表'),sg.Button('設定'),sg.Button('關閉')] ]

#創建視窗
main_Window = sg.Window("股票資料抓取與運算", main_Layout, margins=(40,20))

#從視窗讀取資料
while True:
    event, values = main_Window.read()
    if event == "關閉" or event == sg.WIN_CLOSED:
        break

main_Window.close()