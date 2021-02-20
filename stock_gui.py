import os.path as path
import os
from tkinter import Widget
from typing import Dict
import scrapy
import PySimpleGUI as sg
import numpy as np
import configparser
import pandas as pd
import csv
from datetime import datetime
from scrapy.crawler import CrawlerProcess
from scrapy.utils.project import get_project_settings
import subprocess

sg.theme('DarkAmber') #設定顏色主題
sg.set_options(auto_size_buttons=True)

#全域變數
this_Year = datetime.today().year #獲取今年年份
year_List =[] #存放年份
season_List =['1','2','3','4'] #存放季度
#資料存放
user_Coid_CSV_List =[] #暫存股號表存放
backup_Coid_pd_df =[] #備份user的df

#State
local_Coid_CSV_is_changed = False #存放資料異動狀態
local_Coid_CSV_step=0 #存放異動步驟

#常數
profile_PATH=os.getenv('APPDATA')+'\DSApps\StockSpider\\' #設定檔路徑
_file_name_setting_ini='setting.ini'   #設定檔名稱
_file_name_local_csv='local_Coid.csv'  #CSV檔名稱
default_MDBNAME='theStockDB'    #預設MongoDB名稱
default_MDCDNAME='theStockData' #預設MongoDB的CD名稱
default_MDUrl='mongodb://localhost:27017'   #預設MongoDB連接Url

#setting.ini相關設定
curpath = os.path.dirname(os.path.realpath(__file__))   #目前路徑
cfgpath = os.path.join(profile_PATH,_file_name_setting_ini) #設定檔路徑
csvpath = os.path.join(profile_PATH+_file_name_local_csv)   #本地股號表路徑
conf = configparser.ConfigParser()  #創建設定檔對象


#PD.DF設置
coid_dict ={'代號':[],'名稱':[]} #建立空的本地股號列表
coid_header=['代號','名稱']
local_csvdf = pd.DataFrame(coid_dict).astype(str) #導入本地股號表pd使用
user_df = pd.DataFrame(coid_dict).astype(str) #建立暫存本地股號表pd使用
user_df['名稱'] = user_df['名稱'].astype(str)
user_df['代號'] = user_df['代號'].astype(str)
local_csvdf['名稱'] = local_csvdf['名稱'].astype(str)
local_csvdf['代號'] = local_csvdf['代號'].astype(str)


for i in range(2000,this_Year+1): #新增從2000至今年的年份至列表中
    year_List.append('%4s' % i)

#程式初始化方法
def reset_setting():#重置設定
    if not os.path.exists(profile_PATH):
        os.makedirs(profile_PATH)
    if(not conf.has_section('MongoDB')):
        conf.add_section('MongoDB')
    conf.set('MongoDB','MONGO_URI',default_MDUrl)
    conf.set('MongoDB','DBNAME',default_MDBNAME)
    conf.set('MongoDB','CDATANAME',default_MDCDNAME)
    conf.write(open(cfgpath, 'w'))
    sg.popup('已建立設定檔。')

def reset_csv():#重建csv檔
    global user_Coid_CSV_List,local_csvdf,user_df
    local_csvdf = pd.DataFrame(coid_dict).astype(str) #導入本地股號表pd使用
    local_csvdf.to_csv(csvpath,index=False,sep=',')
    user_df=local_csvdf
    user_Coid_CSV_List=user_df.values.tolist()
    sg.popup('已建立本地股號表。')
    sg.Print('重建本地股號表')

def check_setting():#檢查設定
    if(path.exists(profile_PATH+_file_name_setting_ini)):
        sg.Print('已檢查到設定檔。')
        conf.read(cfgpath,encoding='utf-8')
    else:
        sg.popup('未檢查到設定檔，創建中...',title='系統')
        reset_setting()


def check_local_csv():#檢查本地CSV
    if(path.exists(profile_PATH+_file_name_local_csv)):
        global local_csvdf,user_Coid_CSV_List,user_df
        sg.Print('已檢查到本地股號表。')
        try:
            local_csvdf = pd.read_csv(csvpath, sep=',', engine='python')
            user_Coid_CSV_List=local_csvdf.values.tolist()
            user_df=local_csvdf
        except pd.errors.EmptyDataError:
            sg.popup('讀取本地股號表時發生錯誤！重建本地股號表...')
            reset_csv()
            check_local_csv()
    else:
        sg.popup('未建立本地股號表，創建中...',title='系統')
        reset_csv()

check_setting()
check_local_csv()

#普通應用方法

def filter_Local_CSV_Table(filter_String): #資料過濾
    global user_Coid_CSV_List
    sg.Print(filter_String)
    filter_Coid_CSV_List=[]
    filter_Coid_CSV_List=filter(lambda item: filter_String in item[0] or filter_String in item[1],user_Coid_CSV_List)
    local_Csv_Window['_local_Coid_CSV_Table'].update(values=list(filter_Coid_CSV_List))

def save_Local_CSV(): #保存至本地股號表
    global local_Coid_CSV_step,local_Coid_CSV_is_changed
    user_df.to_csv(csvpath,index=False,sep=',', header=coid_dict)
    check_local_csv()
    backup_Coid_pd_df.clear()
    load_Local_CSV_Table_CSVDF()

def load_Local_CSV_Table_USERVDF(): #從暫存PD.DF中讀入至表單中
    global user_df,user_Coid_CSV_List,local_Coid_CSV_step,local_Coid_CSV_is_changed
    user_Coid_CSV_List=user_df.values.tolist()
    local_Csv_Window['_local_Coid_CSV_Table'].update(values=user_Coid_CSV_List)

def load_Local_CSV_Table_CSVDF(): #從本地股號表PD.DF中讀入至表單中
    global local_csvdf,local_Coid_CSV_step,local_Coid_CSV_is_changed,user_Coid_CSV_List
    user_Coid_CSV_List=local_csvdf.values.tolist()
    local_Csv_Window['_local_Coid_CSV_Table'].update(values=user_Coid_CSV_List)
    local_Coid_CSV_is_changed=False
    local_Coid_CSV_step=0

def refresh_Local_CSV_Table(): #重新整理股號表
    global local_csvdf,local_Coid_CSV_step,local_Coid_CSV_is_changed
    if(local_Csv_Window != None):
        load_Local_CSV_Table_USERVDF()

def update_Local_CSV_Table(): #從本地股表重新載入
    global local_csvdf,local_Coid_CSV_step,local_Coid_CSV_is_changed
    if(local_Csv_Window != None):
        load_Local_CSV_Table_CSVDF()
    
def reset_Local_CSV_Table(): #重置本地股號表
    global local_csvdf,local_Coid_CSV_step,local_Coid_CSV_is_changed
    if(local_Csv_Window != None):
        reset_csv()
        load_Local_CSV_Table_CSVDF()

def local_CSV_Row_Edit(index): #編輯本地股號表 ->編輯單筆資料
    global local_Csv_Window,local_Coid_CSV_is_changed,user_Coid_CSV_List
    old_Coid=str(user_Coid_CSV_List[index][0])
    try:
        old_Coname=str(user_Coid_CSV_List[index][1])
    except IndexError:
        old_Coname=''
    csv_Row_Edit_Window=set_local_CSV_Edit_Row(old_Coid,old_Coname)
    while True : #監聽回傳
        event, values = csv_Row_Edit_Window.read()
        if(event == '保存'):
            print(old_Coid)
            new_Coid=str(values['COID'])
            new_Coname=str(values['CONAME'])
            print(new_Coid)
            user_df['代號'] = user_df['代號'].astype(str)
            print(user_df['代號'])
            user_df['代號'] = user_df['代號'].replace([old_Coid],[new_Coid])
            if '名稱' in user_df.columns:
                user_df['名稱'] = user_df['名稱'].replace([old_Coname],[new_Coname])
            else:
                user_df['名稱'] =""
                user_df['名稱'] = user_df['名稱'].astype(str)
                user_df['代號'] = user_df['代號'].astype(str)
                user_df.loc[index].at['名稱'] = new_Coname
                user_df.drop_duplicates()
            print(user_df['代號'])
            local_Coid_CSV_is_changed=True
            break
        if(event == '取消'):
            csv_Row_Edit_Window.close()
            csv_Row_Edit_Window=None
            break
    if(local_Coid_CSV_is_changed):
        refresh_Local_CSV_Table()
        csv_Row_Edit_Window.close()
        csv_Row_Edit_Window=None
    local_Csv_Window.make_modal()

def local_CSV_usercsvfile_import(isReplace,csv_path): #匯入動作
    global local_Coid_CSV_step,user_Coid_CSV_List
    global local_Coid_CSV_is_changed,user_df
    import_csvdf = pd.DataFrame(coid_dict).astype(str) #導入pd使用
    import_csvdf = pd.read_csv(csv_path, sep=',', engine='python')
    print('名稱' not in import_csvdf.columns)
    if '名稱' not in import_csvdf.columns:
        import_csvdf['名稱'] =""
    import_csvdf['名稱'] = import_csvdf['名稱'].astype(str)
    import_csvdf['代號'] = import_csvdf['代號'].astype(str)
    backup_Coid_pd_df.append(user_df)
    if(isReplace): #取代
        user_df=import_csvdf
    else: #加入
        df_list=[user_df,import_csvdf]
        merged_df=pd.concat(df_list,axis=0)
        merged_df = merged_df.drop_duplicates()
        user_Coid_CSV_List=merged_df.values.tolist()
        user_df=merged_df
        
    user_Coid_CSV_List=user_df.values.tolist()
    local_Coid_CSV_is_changed=True
    local_Coid_CSV_step+=1
    refresh_Local_CSV_Table()
    local_Csv_Window.make_modal()

def local_CSV_Import_usercsvfile(): #選擇外部股號表檔案
    user_CSV_File_Path = sg.popup_get_file('讀入外部股號表',no_window=True,file_types=(("外部CSV股號表","*.csv"),))
    if user_CSV_File_Path !='': #檢測到檔案
        local_Csv_imode_Window=set_local_CSV_Import_usercsvfile_mode()
        while True: #監控視窗回傳
            event, values = local_Csv_imode_Window.read()
            if event =="確定":
                if(values['ucMode_Replace'] == True):
                    Button =sg.popup_ok_cancel('取代目前已有的股號表，確定嗎？')
                    if(Button =="OK"):
                        sg.popup("使用取代模式")
                        local_CSV_usercsvfile_import(True,user_CSV_File_Path)
                        break
                else:
                    Button = sg.popup_ok_cancel('添加至目前已有的股號表，確定嗎？')
                    if(Button =="OK"):
                        sg.popup("使用增加模式")
                        local_CSV_usercsvfile_import(False,user_CSV_File_Path)
                        break
            if event =="取消":
                break
        local_Csv_imode_Window.close()
    local_Csv_imode_Window=None
    window.make_modal()

#視窗設計

def set_local_CSV_Edit_Row(coid,coname): #編輯本地股號表 ->編輯單筆資料
    Edit_Row_Layout =[
        [sg.Text('公司股號：'),sg.Input(default_text=coid,size=(5,1),k='COID')],
        [sg.Text('公司名稱：'),sg.Input(default_text=coname,size=(25,1),k='CONAME')],
        [sg.Button('保存'),sg.Button('取消')]
    ]
    return sg.Window("編輯單筆資料",Edit_Row_Layout,margins=(30,10),finalize=True,modal=True,disable_close=True,disable_minimize=True)

def set_local_CSV_Import_usercsvfile_mode(): #匯入外部股號表 -> 匯入模式
    usercsvfile_Mode_Layout =[
        [sg.Text('選擇匯入模式')],
        [sg.Radio('取代－取代整個股號表',key='ucMode_Replace',group_id='usercsvMode')],
        [sg.Radio('加入－將CSV內的股號表新增至目前已有',key='ucMode_Add',group_id='usercsvMode')],
        [sg.Button('確定'),sg.Button('取消')]
    ]
    return sg.Window("匯入模式",usercsvfile_Mode_Layout,margins=(40,20),finalize=True,modal=True,disable_close=True,disable_minimize=True)

def set_AutoMode_Window(): #主視窗 -> [自動模式] -> 自動爬取來源
    autoMode_Layout =[
                [sg.Text('選擇股號來源')],
                [sg.Radio('從本地股號表讀入',group_id='AM_LoadMode',key='loadFromLocal'),sg.Radio('從CSV檔匯入',group_id='AM_LoadMode',key='_loadFromCSV')],
                [sg.Button('確定'),sg.Button('取消')],
                [sg.Text('本地股報表位於：\n'+csvpath)]
                     ]
    return sg.Window("選擇資料來源",autoMode_Layout,margins=(20,10),finalize=True,modal=True,disable_close=True,disable_minimize=True)

def set_Main_Window(): #主視窗
    main_Layout = [ 
                [sg.Text('資料庫')],
                [sg.Button('顯示資料庫資料')],
                [sg.Text('（自動＼手動）抓取設定')],
                [sg.Combo(year_List, size=(6,5), key='_SearchYear',default_value=this_Year),sg.Text('年'),sg.Combo(season_List, size=(2,5), key='_SearchSeason'),sg.Text('季度'),sg.Text('（手動）查詢股號'),sg.Input(key='_Manual_coid',size=(10,5))],
                [sg.Text('爬取模式'),sg.Radio('自動',group_id='SMode',default=True,key='_Auto'),sg.Radio('手動',group_id='SMode',key='_Manual')],
                [sg.Button('開始爬取')],
                [sg.Text('運行計算式')],
                [sg.Button('公式一'),sg.Button('公式二'),sg.Button('公式三'),sg.Button('公式四')],
                [sg.Text('其他選項')],
                [sg.Button('編輯本地股號表'),sg.Button('設定'),sg.Button('離開'),sg.Button('關於'),sg.Button('說明')] 
                    ]
    return sg.Window("股票資料抓取與運算", main_Layout, margins=(40,20), finalize=True)

def set_Local_CSV_Window(): #主視窗 -> 編輯本地股號表
    local_Coid_CSV_Layout =[
        [sg.Text('輸入股號或公司名稱過濾'),sg.Input(size=(25,1),k='filter_data',enable_events=True)],
        [sg.Text('已變更資料，尚未保存' if local_Coid_CSV_is_changed else '原始資料'),sg.Button('復原',disabled=not(local_Coid_CSV_is_changed))],
        [sg.Table(values=user_Coid_CSV_List,
        headings=['代號','名稱'],
        auto_size_columns=False,
        display_row_numbers=False,
        num_rows=25,select_mode="browse",enable_events=True,
        key='_local_Coid_CSV_Table',right_click_menu=['右鍵',['編輯','刪除']],justification='center',bind_return_key=True)],
        [sg.Button('關閉且「不保存」變更'),sg.Button('關閉且「保存」變更')],
        [sg.Button('保存當前變更'),sg.Button('重新整理')],
        [sg.Button('匯入外部股號表'),sg.Button('重置本地股號表')],
        [sg.Text(f'本地股號表CSV位於{csvpath}')]
        ]
    return sg.Window("編輯本地股號表",local_Coid_CSV_Layout,grab_anywhere=False, finalize=True,modal=True,disable_close=True,disable_minimize=True,
    force_toplevel=True)

def set_Setting_Window(): #主視窗 -> 設定
    setting_Layout = [
    [sg.Text(f'設定檔的路徑位於：{profile_PATH+_file_name_setting_ini}')],
    [sg.Text('MongoDB －你絕大多數不用更動這個選項，此選項區是關於資料庫連接有關與存放爬取資料的相關設定。')],
    [sg.Text('MongoDB 連結：\t'),sg.Input(default_text=(conf['MongoDB']['MONGO_URI']),size=(30,1),k='mDBUrI')],
    [sg.Text('MongoDB 資料庫名稱：\t'),sg.Input(default_text=(conf['MongoDB']['DBNAME']),size=(30,1),k='mDBName')],
    [sg.Text('MongoDB 集合名稱：\t'),sg.Input(default_text=(conf['MongoDB']['CDATANAME']),size=(30,1),k='mCDName')],
    [sg.Button('保存'),sg.Button('取消'),sg.Button('重置')],
    [sg.Button('開啟設定目錄')]
                    ]
    
    return sg.Window("程式設定",setting_Layout, margins=(10,5),finalize=True,modal=True,disable_close=True,disable_minimize=True)

main_Window,setting_Window,aM_Window,local_Csv_Window,local_Csv_imode_Window = set_Main_Window(),None,None,None,None
csv_Row_Edit_Window=None
print('主視窗載入完成。')

while True: #監控視窗回傳
    window,event, values = sg.read_all_windows()
    sg.Print(f'Window:{window},event:{event},values:{values}')
    if window == main_Window: #主視窗
        if event in (sg.WIN_CLOSED,'離開'):
            break
        if event == "開始爬取":
            if(len(values['_SearchYear']) and len(values['_SearchSeason'])): #檢查是否有選擇年與季度
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
            load_Local_CSV_Table_CSVDF()

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
        if event == "開啟設定目錄":
            subprocess.Popen(f'explorer {profile_PATH}')
    
    if window == aM_Window: #主視窗 -> 自動抓取 -> 選擇來源
        if event in (sg.WIN_CLOSED,'取消'):
            window.close()
            aM_Window=None
    
    if window == local_Csv_Window: # 主視窗 -> 編輯本地股號表
        window.bind("<Double-Button-1>","編輯")
        if event == 'filter_data':
            if(values['filter_data'] == ''):
                sg.Print('Null')
                refresh_Local_CSV_Table()
            else:
                filter_Local_CSV_Table(values['filter_data'])
        if event in ('保存當前變更'):
            save_Local_CSV()
        if event in ('關閉且「不保存」變更'):
            update_Local_CSV_Table()
            window.close()
            local_Csv_Window=None
        if event in('關閉且「保存」變更'):
            print(user_df)
            save_Local_CSV()
            backup_Coid_pd_df.clear()
            window.close()
            local_Csv_Window=None
        if event == "重新整理":
            if(local_Coid_CSV_is_changed):
                if(sg.popup_ok_cancel('你股號表尚未存檔，重新整理將會喪失變更的資料並還原修改前的樣子',title='重新整理',modal=True) == 'OK'):
                    update_Local_CSV_Table()
            else:
                update_Local_CSV_Table()
            window.make_modal()
        if event == "匯入外部股號表":
            local_CSV_Import_usercsvfile()
        if event == "重置本地股號表":
            if(sg.popup_ok_cancel('是否重置股號表？',title='確認重置',modal=True) == 'OK'):
                reset_Local_CSV_Table()
                sg.popup('已重置！')
            window.make_modal()
        if event == "刪除":
            sg.popup('刪除')
        if event == "過濾":
            sg.popup("過濾")
        if event == "編輯":
            select_row=values['_local_Coid_CSV_Table']
            if select_row != []:
                select_row=int(select_row[0])
                local_CSV_Row_Edit(select_row)
    
window.close()