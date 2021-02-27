import configparser
import os
import os.path as path
import winsound
from datetime import datetime
from math import ceil as math_ceil

import pathlib

import numpy as np

import pymongo
import PySimpleGUI as sg
from pandas import DataFrame
from pandas import errors as pd_errors
from pandas import read_csv as pd_read_csv

from StockScrapyProject.StockScrapyProject.run_scraper import Scraper

theme_list = sg.theme_list()
sg.set_options(auto_size_buttons=True)

print('！！！此為運行時的控制台，關閉將會立刻關閉程式！！！')

formula_Info = ('A1=現金及約當現金\nA2=透過損益按公允價值衡量之金融資產－流動\nA3=透過其他綜合損益按公允價值衡量之金融資產－流動\nA4=按攤銷後成本衡量之金融資產－流動\nA5=避險之金融資產－流動\nA6=非流動負債合計\nA7=普通股股本\nB1=營業收入合計\nB2=營業利益（損失）\nB3=營業外收入及支出合計\nB4=稀釋每股盈餘合計(EPS)\nPrice=收盤價')

main_Window_Help='歡迎使用股票資訊與計算程式，如果你是第一次使用的話，請先到「編輯本機股號表」裡面匯入你的股號表（格式為CSV），之後即可開始使用其他功能\n如果有要做計算功能，請先用網路爬蟲功能抓取必要的財務報告與股價！且注意需要的抓取數目！\n財務報告：如果目前屬於第四季，那麼你只要抓取該年的1-4季的資料即可，但若不是的話得額外抓取去年的1-4季資料才行！\n股價資料：收盤之後啟動爬蟲抓取，將會幫你抓取今日的收盤價資訊。\n「編輯本機股號表」內可以匯入你已有的一系列股號的CSV檔，請記得匯入時請包含代號與名稱欄位\n你可以在「設定」中存取或刪除指定的資料庫，或者更改配色主題'


theme =''
# 全域變數
this_Year = datetime.today().year  # 獲取今年年份
this_month = datetime.today().month  # 獲取這個月份
this_season = math_ceil(this_month/4)  # 換算季度
year_List = []  # 存放年份
season_List = ['1', '2', '3', '4']  # 存放季度
month_List = ['1', '2', '3', '4', '5', '6',
              '7', '8', '9', '10', '11', '12']  # 存放月份
this_year_season_List = []

DBClient = pymongo.MongoClient()

search_Year = ''
search_Season = ''

for i in range(1, this_season+1):
    this_year_season_List.append(str(i))

# 資料存放
user_Coid_CSV_List = []  # 暫存股號表存放
filter_Coid_CSV_List = []  # 過濾用存放
backup_Coid_pd_df = []  # 備份user的df
DB_LIST = [] #存放資料庫列表
CODATA_LIST = [] #存放資料集列表
filter_String = ''  # 過濾字串

# State
local_Coid_CSV_is_filter = False  # 存放過濾狀態
local_Coid_CSV_is_changed = False  # 存放資料異動狀態
DB_Connect_Status = False #存放連線狀態
DB_CODATA_Exist = False #存放資料集存在狀態
DB_READY = DB_Connect_Status and DB_CODATA_Exist #存放資料庫預備狀態

# 常數
profile_PATH = os.getenv('APPDATA')+'\DSApps\StockSpider\\'  # 設定檔路徑
_file_name_setting_ini = 'setting.ini'  # 設定檔名稱
_file_name_local_csv = 'local_Coid.csv'  # CSV檔名稱
default_MDBNAME = 'theStockDB'  # 預設MongoDB名稱
default_MDCDNAME = 'theStockData'  # 預設MongoDB的CD名稱
default_MDUrl = 'mongodb://localhost:27017/?readPreference=primary&appname=StockSpiderwGUI&ssl=false'  # 預設MongoDB連接Url

# setting.ini相關設定
curpath = os.path.dirname(os.path.realpath(__file__))  # 目前路徑
cfgpath = os.path.join(profile_PATH, _file_name_setting_ini)  # 設定檔路徑
csvpath = os.path.join(profile_PATH+_file_name_local_csv)  # 本機股號表路徑
conf = configparser.ConfigParser()  # 創建設定檔對象


# PD.DF設置
coid_dict = {"代號": [], "名稱": []}  # 建立空的本機股號列表
coid_dict_type = {'代號': 'string', '名稱': 'string'}  # 建立股號列表檔案類型
local_csvdf = DataFrame(coid_dict)  # 導入本機股號表pd使用
user_df = DataFrame(coid_dict)  # 建立暫存本機股號表pd使用
import_csv_df = DataFrame(coid_dict)  # 導入pd使用
import_csv_df.astype("string")  # 設定資料類型為字串
user_df.astype("string")  # 設定資料類型為字串
local_csvdf.astype("string")  # 設定資料類型為字串


for i in range(2019, this_Year+1):  # 新增從2000至今年的年份至列表中
    year_List.append('%4s' % i)


def push_4_season_back():  # 往前推四個季度
    global this_season
    if(this_season == 4):
        return 1
    if(this_season == 3):
        return 4
    if(this_season == 2):
        return 3
    if(this_season == 1):
        return 2


# 初始化視窗設計

def set_MONOGO_List(title, check_list, inputName):
    if(len(check_list) != 0):
        List_Value = inputName
        disable_delete = False
    else:
        List_Value = ''
        disable_delete = True
    DB_List_Layout = [
        [sg.Text('請選擇要使用的項目')],
        [sg.Combo(check_list, default_value=List_Value,
                  k='MList', readonly=True, size=(25, 1))],
        [sg.Text('欄位空白將會視為新建')],
        [sg.Button('新建'), sg.Button('刪除', disabled=disable_delete),
         sg.Text('\t\t'), sg.Button('確定'), sg.Button('取消')],
    ]
    return sg.Window(title, layout=DB_List_Layout, modal=True, force_toplevel=True, disable_close=True, disable_minimize=True)
# 程式初始化方法


def check_Mongo():
    connect_Mongo(False, False, False, False)
    return DB_READY


def connect_Mongo(isInit, isCreateNewDB, isCreateCODATA, isNeedSelect):
    global DBClient, DB_Connect_Status, DB_CODATA_Exist, DB_READY, DB_LIST, CODATA_LIST
    MongoURI = str(conf['MongoDB']['MONGO_URI'])
    DBClient = pymongo.MongoClient(MongoURI)
    CreateNewDB = isCreateNewDB
    CreateCODATA = isCreateCODATA
    print(MongoURI)
    try:
        try:
            # str(conf['MongoDB']['DBNAME'])
            MongoDBName = conf.get('MongoDB', 'DBNAME')
        except configparser.NoOptionError:
            MongoDBName = ''
        try:
            # str(conf['MongoDB']['CDATANAME'])
            MongoDB_CODATA = conf.get('MongoDB', 'CDATANAME')
        except configparser.NoOptionError:
            MongoDB_CODATA = ''

        sg.popup_no_buttons('連接到 MonogoDB 中，請稍後...', non_blocking=True, grab_anywhere=False,
                            no_titlebar=True, auto_close=True, auto_close_duration=1)
        DB_LIST = DBClient.list_database_names()
        DB_LIST.remove('local')
        DB_LIST.remove('config')
        DB_LIST.remove('admin')
        #MongoDBName=MongoDBName if (MongoDBName in DB_LIST) else ''
        if(isNeedSelect):
            set_DB_Window = set_MONOGO_List('選擇資料庫', DB_LIST, MongoDBName)
            while True:
                events, values = set_DB_Window.read()

                if events == '確定':
                    MongoDBName = values['MList']
                    if(MongoDBName != ''):
                        CreateNewDB = False
                        break
                    else:
                        CreateNewDB = True
                        break

                if events == '取消':
                    set_DB_Window.close()
                    set_DB_Window = None
                    return None

                if events == '刪除' and values['MList'] != '':
                    winsound.MessageBeep(winsound.MB_ICONQUESTION)
                    Button = sg.popup_yes_no(
                        '確定要刪除 '+values['MList']+' 資料庫嗎？ 該動作無法復原。', no_titlebar=True)
                    if(Button == 'Yes'):
                        DBClient.drop_database(values['MList'])
                        DB_LIST = DBClient.list_database_names()
                        DB_LIST.remove('local')
                        DB_LIST.remove('config')
                        DB_LIST.remove('admin')
                        set_DB_Window['MList'].update(values=DB_LIST, value='')
                if events == '新建':
                    MongoDBName = ''
                    CreateNewDB = True
                    break
            set_DB_Window.close()
            set_DB_Window = None
        print(MongoDBName)
        if(MongoDBName in DB_LIST and (not CreateNewDB)):
            sg.popup_no_buttons(f'已連接到 {MongoDBName}，檢查 {MongoDB_CODATA} 是否資料集存在...', non_blocking=True,
                                grab_anywhere=False, no_titlebar=True, auto_close=True, auto_close_duration=1)
            DB_Connect_Status = True
            DB_CODATA_Exist = False
            NewCOName = ''
            db = DBClient.get_database(MongoDBName)
            CODATA_LIST = db.list_collection_names()
            try:
                CODATA_LIST.remove('init')
            except ValueError:
                None
            #MongoDB_CODATA=MongoDB_CODATA if (MongoDB_CODATA in CODATA_LIST) else ''
            if(isNeedSelect):
                set_DB_Window = set_MONOGO_List(
                    '選擇資料集', CODATA_LIST, MongoDB_CODATA)
                while True:
                    events, values = set_DB_Window.read()

                    if events == '確定':
                        MongoDB_CODATA = values['MList']
                        print(MongoDB_CODATA)
                        if(MongoDB_CODATA != ''):
                            CreateCODATA = False
                            break
                        else:
                            CreateCODATA = True
                            break

                    if events == '取消':
                        set_DB_Window.close()
                        set_DB_Window = None
                        return None

                    if events == '刪除' and values['MList'] != '':
                        CODATA_COUNT = len(CODATA_LIST)
                        winsound.MessageBeep(winsound.MB_ICONQUESTION)
                        Button = sg.popup_yes_no(
                            f'確定要刪除 '+values['MList']+f' 資料集嗎？\n該資料庫資料集數量為 {CODATA_COUNT}\n系統會自動刪除無資料集的資料庫', no_titlebar=True)
                        if(Button == 'Yes'):
                            db.drop_collection(values['MList'])
                            CODATA_LIST = db.list_collection_names()
                            set_DB_Window['MList'].update(
                                values=CODATA_LIST, value='')
                    if events == '新建':
                        MongoDB_CODATA = ''
                        CreateCODATA = True
                        break
                set_DB_Window.close()
                set_DB_Window = None
            print(MongoDB_CODATA)

            if(MongoDB_CODATA in CODATA_LIST and (not CreateCODATA)):
                DB_CODATA_Exist = True
                DB_READY = DB_Connect_Status and DB_CODATA_Exist
                conf.set('MongoDB', 'DBNAME', str(MongoDBName))
                conf.set('MongoDB', 'CDATANAME', str(MongoDB_CODATA))
                conf.write(open(cfgpath, 'w'))
                sg.SystemTray.notify(
                    'MonogoDB 已預備完成！', f'資料庫: {MongoDBName} \n資料集: {MongoDB_CODATA}\n功能初始化完成', display_duration_in_ms=300, fade_in_duration=.2)
            else:
                if(isInit or CreateCODATA):
                    if(not CreateCODATA and not isInit):
                        winsound.MessageBeep(winsound.MB_ICONQUESTION)
                        if(sg.popup_yes_no(f'在 {MongoDBName} 之資料庫中找不到 {MongoDB_CODATA} 資料集，是否創建新的資料集？', no_titlebar=True) == 'Yes'):
                            NewCOName = sg.popup_get_text(
                                message='輸入資料集的名稱', default_text=default_MDCDNAME)
                    else:
                        #sg.Print('New CODATA:'+NewCOName)
                        NewCOName = sg.popup_get_text(
                            message='輸入資料集的名稱', default_text=default_MDCDNAME)
                        if(NewCOName != None):
                            try:
                                db.create_collection(NewCOName)
                                conf.set('MongoDB', 'CDATANAME', NewCOName)

                            except pymongo.errors.CollectionInvalid:
                                winsound.MessageBeep(winsound.MB_ICONHAND)
                                sg.popup_error(f'在 {MongoDBName} 當中該資料集已存在！')
                                setting_Window.make_modal()
                            db.drop_collection('init')
                            check_Mongo()
                else:
                    sg.popup(
                        f'在 {MongoDBName} 之資料庫中找不到 {MongoDB_CODATA} 資料集，請到設定更改或建立有效的資料庫。', no_titlebar=True)
        else:
            DB_Connect_Status = False
            DB_CODATA_Exist = False
            DB_READY = DB_Connect_Status and DB_CODATA_Exist
            NewDBName = None
            if(isInit or CreateNewDB):
                if(not CreateNewDB and not isInit):
                    if(sg.popup_yes_no(f'在 MongoDB 中找不到 {MongoDBName} 資料庫，是否創建資料庫？', no_titlebar=True) == 'Yes'):
                        NewDBName = sg.popup_get_text(
                            message='輸入資料庫的名稱', default_text=default_MDBNAME)
                else:
                    NewDBName = sg.popup_get_text(
                        message='輸入資料庫的名稱', default_text=default_MDBNAME)
                print('New DB:'+str(NewDBName))
                if(NewDBName != None):
                    db = DBClient[NewDBName]
                    db.create_collection('init')
                    conf.set('MongoDB', 'DBNAME', str(NewDBName))
                    connect_Mongo(True, False, CreateCODATA, False)
                else:
                    return None
            else:
                sg.popup_ok(
                    f'在 MongoDB 中找不到 {MongoDBName} 資料庫，請到設定更改為有效的資料庫名稱。', no_titlebar=True)
    except pymongo.errors.ServerSelectionTimeoutError:
        DB_Connect_Status = False
        DB_CODATA_Exist = False
        DB_READY = DB_Connect_Status and DB_CODATA_Exist
        conf.set('MongoDB', 'DBNAME', '')
        conf.set('MongoDB', 'CDATANAME', '')
        conf.write(open(cfgpath, 'w'))
        sg.popup_ok('MongoDB 連接失敗，請確定是否有安裝 MonogoDB\n或者 MongoDB 服務是否有運行中！',
                    title='MonogoDB', no_titlebar=True)
        return None


def reset_setting():  # 重置設定
    if not os.path.exists(profile_PATH):
        os.makedirs(profile_PATH)
    if(not conf.has_section('MongoDB')):
        conf.add_section('MongoDB')
    if(not conf.has_section('System')):
        conf.add_section('System')
    conf.set('MongoDB', 'MONGO_URI', str(default_MDUrl))
    conf.set('System','Theme',str('DarkBlack1'))
    conf.write(open(cfgpath, 'w'))
    conf.read(cfgpath, encoding='utf-8')
    sg.popup('已建立設定檔。')
    connect_Mongo(True, False, False, True)


def reset_csv():  # 重建csv檔
    global user_Coid_CSV_List, local_csvdf, user_df
    local_csvdf = DataFrame(coid_dict)  # 導入本機股號表pd使用
    local_csvdf.to_csv(csvpath, index=False, sep=',')
    user_df = local_csvdf
    user_Coid_CSV_List = user_df.values.tolist()
    sg.SystemTray.notify(
        '系統', '已建立本機股號表。', display_duration_in_ms=250, fade_in_duration=.2)
    # sg.Print('重建本機股號表')


def check_setting():  # 檢查設定
    if(path.exists(profile_PATH+_file_name_setting_ini)):
        sg.SystemTray.notify(
            '系統', '已檢查到設定檔。', display_duration_in_ms=250, fade_in_duration=.2)
        conf.read(cfgpath, encoding='utf-8')

        if(conf.has_option('System', 'Theme')):
            theme = sg.theme(conf.get('System','Theme'))

        if(not conf.has_option('MongoDB', 'mongo_uri') or not conf.has_option('MongoDB', 'dbname') or not conf.has_option('MongoDB', 'cdataname')):
            winsound.MessageBeep(winsound.MB_ICONHAND)
            sg.popup_error('系統', '資料庫相關設置遺失！重置設定檔中...')
            reset_setting()
    else:
        sg.SystemTray.notify('系統', '未檢查到設定檔，創建中...',
                             display_duration_in_ms=5000, fade_in_duration=.2)
        reset_setting()


def check_local_csv():  # 檢查本機CSV
    if(path.exists(profile_PATH+_file_name_local_csv)):
        global local_csvdf, user_Coid_CSV_List, user_df
        sg.SystemTray.notify('系統', '已檢查到本機股號表。',
                             display_duration_in_ms=250, fade_in_duration=.2)
        try:
            local_csvdf = pd_read_csv(
                csvpath, sep=',', engine='python', dtype=coid_dict_type, na_filter=False)
            local_csvdf = local_csvdf
            user_Coid_CSV_List = local_csvdf.values.tolist()
            user_df = local_csvdf
        except pd_errors.EmptyDataError:
            winsound.MessageBeep(winsound.MB_ICONHAND)
            sg.popup_error('讀取本機股號表時發生錯誤！重建本機股號表...')
            reset_csv()
            check_local_csv()
    else:
        sg.popup('未建立本機股號表，創建中...', title='系統')
        reset_csv()

# 初始化函數


scrapyer = Scraper()
check_setting()
check_local_csv()
check_Mongo()

# 爬蟲調用


def call_Price_Spider(isLocal, LOAD_CSVPATH):
    global csvpath, Force_Exit_Window
    info = ''
    if(isLocal):
        scrapyer.set_PriceSpider(csvpath)
        sg.popup_no_buttons('啟用爬蟲中...', grab_anywhere=False,
                            no_titlebar=True, auto_close=True,auto_close_duration=5)
        main_Window.minimize()
        main_Window.close()
        scrapyer.run_PriceSpider()
    else:
        scrapyer.set_PriceSpider(LOAD_CSVPATH)
        sg.popup_no_buttons('啟用爬蟲中...', grab_anywhere=False,
                            no_titlebar=True, auto_close=True,auto_close_duration=5)
        main_Window.minimize()
        main_Window.close()
        scrapyer.run_PriceSpider()
    print('\n')
    winsound.MessageBeep(type=winsound.MB_OK)
    os._exit(0)


def call_Stock_Spider(isAutoMode, isLocal, LOAD_CSVPATH, M_CO_ID):
    global csvpath, Force_Exit_Window
    info = ''
    if(isAutoMode):
        if(isLocal):
            scrapyer.set_StockSpider(
                Year=search_Year, Season=search_Season, Mode='Auto', CSV=csvpath)
            sg.popup_no_buttons('啟用爬蟲中...', grab_anywhere=False,
                            no_titlebar=True, auto_close=True,auto_close_duration=5)
            main_Window.minimize()
            main_Window.close()
            scrapyer.run_StockSpider()
        else:
            scrapyer.set_StockSpider(
                Year=search_Year, Season=search_Season, Mode='Auto', CSV=LOAD_CSVPATH)
            sg.popup_no_buttons('啟用爬蟲中...', grab_anywhere=False,
                            no_titlebar=True, auto_close=True,auto_close_duration=5)
            main_Window.minimize()
            main_Window.close()
            scrapyer.run_StockSpider()
    else:
        scrapyer.set_StockSpider(
            Year=search_Year, Season=search_Season, Mode='Manual', CO_ID=M_CO_ID)
        scrapyer.run_StockSpider()
    print('\n')
    winsound.MessageBeep(type=winsound.MB_OK)
    sg.popup_ok('抓取股票財務報告完成！程式將會關閉！')
    os._exit(0)


# 普通應用方法

def local_CSV_Restore_USER_DF():  # 還原本機股號表資料狀態
    global user_df, local_Coid_CSV_is_changed, backup_Coid_pd_df, local_Csv_Window
    user_df = backup_Coid_pd_df.pop(-1)
    step = len(backup_Coid_pd_df)
    print(step)
    if(step == 0):
        local_Coid_CSV_is_changed = False
        local_Csv_Window['backup_btn'].update(
            disabled=not(local_Coid_CSV_is_changed))


def local_CSV_Backup_USER_DF():  # 備份本機股號表資料狀態
    global user_df, local_Coid_CSV_is_changed, backup_Coid_pd_df, local_Csv_Window
    backup_Coid_pd_df.append(user_df)
    step = len(backup_Coid_pd_df)
    local_Coid_CSV_is_changed = True
    print(step)
    local_Csv_Window['backup_btn'].update(
        disabled=not(local_Coid_CSV_is_changed))


def local_CSV_Row_Add(co_ID, co_Name):  # 新增資料
    global user_df
    local_CSV_Backup_USER_DF()
    new_row = {'代號': str(co_ID), '名稱': str(co_Name)}
    user_df = user_df.append(new_row, ignore_index=True)
    refresh_Local_CSV_Table()


def filter_Local_CSV_Table(filter_String):  # 資料過濾
    global user_Coid_CSV_List, filter_Coid_CSV_List, local_Csv_Window
    filter_String = str(filter_String)
    # sg.Print(filter_String)
    filter_Coid_CSV_List = []
    if(filter_String.isnumeric()):
        filter_Coid_CSV_List = filter(
            lambda x: filter_String in x[0] or filter_String in x[1], user_Coid_CSV_List)
    else:
        filter_Coid_CSV_List = filter(
            lambda x: filter_String in x[1] or filter_String in x[0], user_Coid_CSV_List)
    filter_Coid_CSV_List = list(filter_Coid_CSV_List)
    local_Csv_Window['_local_Coid_CSV_Table'].update(
        values=filter_Coid_CSV_List)


def save_Local_CSV():  # 保存至本機股號表
    global local_Coid_CSV_is_changed, user_df
    user_df.to_csv(csvpath, index=False, sep=',',
                   header=coid_dict, encoding='utf-8')
    check_local_csv()
    load_Local_CSV_Table_CSVDF()


def load_Local_CSV_Table_USERVDF():  # 從暫存PD.DF中讀入至表單中
    global user_df, user_Coid_CSV_List, local_Coid_CSV_is_changed, filter_String
    user_Coid_CSV_List = user_df.values.tolist()
    if(local_Coid_CSV_is_filter):
        filter_Local_CSV_Table(filter_String)
    else:
        local_Csv_Window['_local_Coid_CSV_Table'].update(
            values=user_Coid_CSV_List)


def load_Local_CSV_Table_CSVDF():  # 從本機股號表PD.DF中讀入至表單中
    global local_csvdf, local_Coid_CSV_is_changed, user_Coid_CSV_List
    user_Coid_CSV_List = local_csvdf.values.tolist()
    local_Csv_Window['_local_Coid_CSV_Table'].update(values=user_Coid_CSV_List)


def refresh_Local_CSV_Table():  # 重新整理股號表
    global local_csvdf, local_Coid_CSV_is_changed
    if(local_Csv_Window != None):
        load_Local_CSV_Table_USERVDF()


def update_Local_CSV_Table():  # 從本機股表重新載入
    global local_csvdf, local_Coid_CSV_is_changed
    if(local_Csv_Window != None):
        load_Local_CSV_Table_CSVDF()


def reset_Local_CSV_Table():  # 重置本機股號表
    global local_csvdf, local_Coid_CSV_is_changed
    if(local_Csv_Window != None):
        reset_csv()
        load_Local_CSV_Table_CSVDF()


def local_CSV_Row_Edit(isEdit, index):  # 編輯本機股號表 ->編輯單筆資料
    global local_Csv_Window, local_Coid_CSV_is_changed, user_Coid_CSV_List, user_df, filter_Coid_CSV_List
    if(not local_Coid_CSV_is_filter):
        old_Coid = str(user_Coid_CSV_List[index][0])
        try:
            old_Coname = str(user_Coid_CSV_List[index][1])
        except IndexError:
            old_Coname = ''
    else:
        old_Coid = str(filter_Coid_CSV_List[index][0])
        try:
            old_Coname = str(filter_Coid_CSV_List[index][1])
        except IndexError:
            old_Coname = ''
    if(isEdit):
        csv_Row_Edit_Window = set_local_CSV_Edit_Row(old_Coid, old_Coname)
        while True:  # 監聽回傳
            event, values = csv_Row_Edit_Window.read()
            if(event == '保存'):
                local_CSV_Backup_USER_DF()
                print(old_Coid)
                new_Coid = str(values['COID'])
                new_Coname = str(values['CONAME'])
                print(new_Coid)
                user_df['代號'] = user_df['代號']
                print(user_df['代號'])
                user_df['代號'] = user_df['代號'].replace([old_Coid], [new_Coid])
                if '名稱' in user_df.columns:
                    user_df['名稱'] = user_df['名稱'].replace(
                        [old_Coname], [new_Coname])
                else:
                    user_df.loc[index].at['名稱'] = new_Coname
                print(user_df['代號'])
                break
            if(event == '取消'):
                break
    else:
        csv_Row_Edit_Window = set_local_CSV_Remove_Row(old_Coid, old_Coname)
        while True:
            event, values = csv_Row_Edit_Window.read()
            if(event == '是'):
                local_CSV_Backup_USER_DF()
                user_df = user_df[user_df['代號'] != old_Coid]
                user_df = user_df[user_df['名稱'] != old_Coname]
                break
            else:
                break
    if(local_Coid_CSV_is_changed):
        refresh_Local_CSV_Table()
    csv_Row_Edit_Window.close()
    csv_Row_Edit_Window = None
    local_Csv_Window.make_modal()


def local_CSV_usercsvfile_import(isReplace, csv_path):  # 匯入動作
    global user_Coid_CSV_List
    global local_Coid_CSV_is_changed, user_df, import_csv_df
    import_csv_df = pd_read_csv(
        csv_path, sep=',', engine='python', dtype=coid_dict_type, na_filter=False)
    local_CSV_Backup_USER_DF()
    if(isReplace):  # 取代
        user_df = import_csv_df
    else:  # 加入
        user_df = user_df.append(import_csv_df)
        # df_list=[user_df,import_csv_df]
        # merged_df=pd.concat(df_list,axis=0,ignore_index=True)
        #merged_df = merged_df.drop_duplicates(subset=['代號'])
        print(user_df.dtypes, import_csv_df.dtypes)
        user_df = user_df.drop_duplicates(
            subset=['代號', '名稱'], keep='first', ignore_index=True)
        user_Coid_CSV_List = user_df.values.tolist()

    user_Coid_CSV_List = user_df.values.tolist()
    local_Coid_CSV_is_changed = True
    refresh_Local_CSV_Table()
    local_Csv_Window.make_modal()


def local_CSV_Import_usercsvfile():  # 選擇外部股號表檔案
    user_CSV_File_Path = sg.popup_get_file(
        '讀入外部股號表', no_window=True, file_types=(("外部CSV股號表", "*.csv"),))
    if user_CSV_File_Path != '':  # 檢測到檔案
        local_Csv_imode_Window = set_local_CSV_Import_usercsvfile_mode()
        while True:  # 監控視窗回傳
            event, values = local_Csv_imode_Window.read()
            if event == "確定":
                if(values['ucMode_Replace'] == True):
                    winsound.MessageBeep(winsound.MB_ICONQUESTION)
                    Button = sg.popup_ok_cancel('取代目前已有的股號表，確定嗎？')
                    if(Button == "OK"):
                        local_CSV_usercsvfile_import(True, user_CSV_File_Path)
                        break
                else:
                    winsound.MessageBeep(winsound.MB_ICONQUESTION)
                    Button = sg.popup_ok_cancel('添加至目前已有的股號表，確定嗎？')
                    if(Button == "OK"):
                        local_CSV_usercsvfile_import(False, user_CSV_File_Path)
                        break
            if event == "取消":
                break
        local_Csv_imode_Window.close()
    local_Csv_imode_Window = None
    window.make_modal()

# 資料庫管理物件


class MongoDB_Load():
    is_calcDataDF_Ready=False
    global DBClient, DB_LIST, CODATA_LIST, year_List, season_List
    default_dict={'測試欄位':['測試資料'],'測試欄位2':['測試資料'],'測試欄位3':['測試資料']}
    tableDF=DataFrame(data=default_dict)
    table_List = ['測試資料']
    table_Heading = ['測試欄位']
    year_filter_list = ['全部']
    season_filter_list = ['全部']
    year_filter_list += year_List
    season_filter_list += season_List
    mongoDB_DBName = ''
    mongoDB_CData = ''
    tableType = '請選擇要讀取的資料'
    db_Data_Newest_Year = ''
    db_Data_Newest_Season = ''
    StockDataDF=DataFrame([])
    calcDataDF=DataFrame()
    calcAnsDF=DataFrame()
    StockPriceDF=DataFrame([])
    MixDataDF=DataFrame([])
    Date=''
    db=''
    codata=''
    displayDB_Layout = []


    def filter_db_Table(self,filter_String):
        if(filter_String!=""):
            filter_list=filter(lambda x: filter_String in x[0:] ,self.table_List)
            filter_list=list(filter_list)
            print(filter_list)
            displayDB_Window['display_Table'].update(values=filter_list)
        else:
            filter_list=[]
            self.update_TableData()
            self.update_TableWithoutColChange()
    def export_Table(self):
        filename = sg.popup_get_file('選擇儲存路徑','匯出表格',default_path=f'{self.Date } 之{self.tableType} 匯出',save_as=True,file_types=(("CSV 檔","*.csv"),("Excel 檔","*.xlsx")),no_window=True)
        if(pathlib.Path(filename).suffix==".csv"):
            self.tableDF.to_csv(filename,encoding='utf-8', index=False)
        if(pathlib.Path(filename).suffix==".xlsx"):
            self.tableDF.to_excel(filename,encoding='utf-8', index=False)
    def clean_Data(self):
        self.table_List.clear()
        self.table_Heading.clear()
        pass

    def __init__(self) -> None:
        self.mongoDB_DBName = conf.get('MongoDB', 'DBNAME')
        self.mongoDB_CData = conf.get('MongoDB', 'CDATANAME')
        self.db = DBClient[self.mongoDB_DBName] #獲得現在資料庫
        self.codata = self.db[self.mongoDB_CData] #獲得現在資料庫中哪個資料集
        pass
    
    def calc_Forumla(self,coid_list,ForumlaType):
        self.calcAnsDF=DataFrame()
        start_year=int(self.db_Data_Newest_Year)
        start_season=int(self.db_Data_Newest_Season)
        current_process=0
        coid_list_max=len(coid_list)
        for coid in coid_list:
            return_FType=False
            current_process+=1
            ans_block=0
            not_manual_cancel = sg.one_line_progress_meter('計算資料..',current_process,coid_list_max,'calc',orientation='h')
            if(not not_manual_cancel and current_process < coid_list_max):
                Button = sg.popup_yes_no('確定取消？',title='手動取消')
                if(Button=='Yes'):
                    sg.popup_ok('已手動取消！')
                    main_Window.normal()
                    return False
            calc_VarData=self.calcDataDF[(self.calcDataDF["股號"]==coid) & (self.calcDataDF["年份"] == str(start_year)) & (self.calcDataDF["季度"] == str(start_season))]
            calc_VarData = calc_VarData.reset_index(drop=True)
            name=calc_VarData.iloc[0]["名稱"]
            recent_EPS=calc_VarData.iloc[0]["近四季 EPS"]
            if(ForumlaType=='公式一'):
                return_FType=True
                A1=calc_VarData.iloc[0]["A1"]
                A2=calc_VarData.iloc[0]["A2"]
                A3=calc_VarData.iloc[0]["A3"]
                A4=calc_VarData.iloc[0]["A4"]
                A5=calc_VarData.iloc[0]["A5"]
                A6=calc_VarData.iloc[0]["A6"]
                A7=calc_VarData.iloc[0]["A7"]
                if(A7 == 0.00):
                    continue
                Price=calc_VarData.iloc[0]["股價"]
                ans_block=round((((A1+A2+A3+A4+A5)-A6)/(A7/10.00)-Price),2)
                pass
            
            if(ForumlaType=='公式二'):
                return_FType=True
                B4=calc_VarData.iloc[0]["B4"]
                LastYear_B4=calc_VarData.iloc[0]["去年同期B4"]
                Price=calc_VarData.iloc[0]["股價"]
                r_EPS=calc_VarData.iloc[0]["近四季 EPS"]
                if(LastYear_B4 == 0.00 or r_EPS == 0.00):
                    continue

                ans_block=round((((B4-LastYear_B4)/LastYear_B4)*100.0-Price/r_EPS),2)
                pass

            if(ForumlaType== '公式三'):
                return_FType=True
                B2=calc_VarData.iloc[0]["B2"]
                LastYear_B2=calc_VarData.iloc[0]["去年同期B2"]
                Price=calc_VarData.iloc[0]["股價"]
                r_EPS=calc_VarData.iloc[0]["近四季 EPS"]
                if(LastYear_B2 == 0.00 or r_EPS == 0.00):
                    continue
                ans_block=round((((B2-LastYear_B2)/LastYear_B2)*100.00-Price/r_EPS),2)
                pass
            if(ForumlaType == '公式四'):
                return_FType=True
                calc_block_1=0.00
                calc_block_1=calc_VarData.iloc[0]["B2"]
                calc_block_1-=calc_VarData.iloc[0]["去年同期B2"]
                bool_block=calc_block_1 > 1.00
                calc_block_2=0.00
                calc_block_2=calc_VarData.iloc[0]["B3"]
                calc_block_2-=calc_VarData.iloc[0]["去年同期B3"]
                bool_block2=calc_block_2 < 1.00
                if(bool_block and bool_block2):
                    ans_block='符合'
                else:
                    continue
            if(ForumlaType == '公式五'):
                calc_block_1=0.00
                calc_block_1=calc_VarData.iloc[0]["B2"]
                calc_block_1-=calc_VarData.iloc[0]["去年同期B2"]
                bool_block=calc_block_1 > 1.00
                if(bool_block):
                    ans_block='符合'
                else:
                    continue
                pass
            if(ForumlaType == '公式六'):
                calc_block_1=0.00
                calc_block_1=calc_VarData.iloc[0]["B3"]
                calc_block_1-=calc_VarData.iloc[0]["去年同期B3"]
                bool_block=calc_block_1 < 1.00
                if(bool_block):
                    ans_block='符合'
                else:
                    continue
                pass
            dict= {"股號":str(coid),"名稱":name,"年份":str(start_year),"季度":str(start_season),"公式":ForumlaType,"答案":ans_block,"近四季 EPS":recent_EPS}
            cols=['股號','名稱','年份','季度','答案','近四季 EPS']
            self.calcAnsDF = self.calcAnsDF.append(dict, ignore_index=True)
            self.calcAnsDF = self.calcAnsDF[cols]
            self.calcAnsDF = self.calcAnsDF.replace([np.inf, -np.inf], np.nan).dropna(axis=0)
            if(return_FType):
                pass
                #print(self.calcAnsDF.dtypes)
            #print(self.calcAnsDF)
        return True

    def get_calc_Formula_var(self,coid_list):
        if(self.is_calcDataDF_Ready):
            return True
        #print(coid_list)
        start_year=int(self.db_Data_Newest_Year)
        start_season=int(self.db_Data_Newest_Season)
        run_year=start_year
        run_season=start_season
        end_year=start_year
        end_year-=1
        current_process=0
        coid_list_max=len(coid_list)
        for coid in coid_list: #目前股號
            #print(coid)
            current_process+=1
            not_manual_cancel = sg.one_line_progress_meter('從資料中抓取變數...',current_process,coid_list_max,'calc',orientation='h')
            if(not not_manual_cancel and current_process < coid_list_max):
                Button = sg.popup_yes_no('確定取消？',title='手動取消')
                if(Button=='Yes'):
                    sg.popup_ok('已手動取消！')
                    main_Window.normal()
                    return False
            global local_csvdf
            run_year=start_year
            count_Season=1
            run_season=start_season
            getStockData_StartTime=self.StockDataDF[(self.StockDataDF["股號"]==coid) & (self.StockDataDF["年份"] == str(start_year)) & (self.StockDataDF["季度"] == str(start_season))]
            getStockData_LastYear=self.StockDataDF[(self.StockDataDF["股號"]==coid) & (self.StockDataDF["年份"] == str(start_year-1)) & (self.StockDataDF["季度"] == str(start_season))]
            getStockPriceData=self.StockPriceDF[ ( self.StockPriceDF["股號"]==coid ) & ( self.StockPriceDF["收盤日"] == self.Date )]
            recent_EPS=0.0
            getStockData_StartTime = getStockData_StartTime.reset_index(drop=True)
            getStockData_LastYear = getStockData_LastYear.reset_index(drop=True)
            getStockPriceData = getStockPriceData.reset_index(drop=True)
            try:
                TEST=getStockData_LastYear.iloc[0]["B3"]
            except IndexError:
                sg.Print(f'抓不到股號 {coid} 在 {start_year-1} 之第 {start_season} 季度中的資料，請確定是否有抓取該年該季的資料！')
                continue
            try:
                name=getStockPriceData.iloc[0]["公司縮寫"]
            except IndexError:
                sg.Print(f'抓不到 {coid} 在 {self.Date}的股價資料！')
                continue
            A1=getStockData_StartTime.iloc[0]["A1"]
            A2=getStockData_StartTime.iloc[0]["A2"]
            A3=getStockData_StartTime.iloc[0]["A3"]
            A4=getStockData_StartTime.iloc[0]["A4"]
            A5=getStockData_StartTime.iloc[0]["A5"]
            A6=getStockData_StartTime.iloc[0]["A6"]
            A7=getStockData_StartTime.iloc[0]["A7"]
            B1=getStockData_StartTime.iloc[0]["B1"]
            B2=getStockData_StartTime.iloc[0]["B2"]
            B3=getStockData_StartTime.iloc[0]["B3"]
            last_year_B3=getStockData_LastYear.iloc[0]["B3"]
            last_year_B2=getStockData_LastYear.iloc[0]["B2"]
            last_year_B4=0.0
            B4=0.0
            if(start_season!=4):
                B4=getStockData_StartTime.iloc[0]["B4"]
                last_year_B4=getStockData_LastYear.iloc[0]["B4"]
            else:
                temp_S1_S3_SUM=0.0
                B4=getStockData_StartTime.iloc[0]["B4"]
                last_year_B4=getStockData_LastYear.iloc[0]["B4"]
                #print('年',run_year,'季度',run_season)
                for season in range(3,0,-1): #今年
                    temp_StockData=self.StockDataDF[(self.StockDataDF["股號"]==coid) & (self.StockDataDF["年份"] == str(start_year)) & (self.StockDataDF["季度"] == str(season))]
                    temp_StockData = temp_StockData.reset_index(drop=True)
                    temp_S1_S3_SUM+=temp_StockData.iloc[0]["B4"]
                #print(temp_S1_S3_SUM)
                B4-=temp_S1_S3_SUM
                B4=round(B4,2)
                temp_S1_S3_SUM=0.0
                #print('年',run_year,'季度',run_season)
                for season in range(3,0,-1): #去年
                    temp_StockData=self.StockDataDF[(self.StockDataDF["股號"]==coid) & (self.StockDataDF["年份"] == str(start_year-1)) & (self.StockDataDF["季度"] == str(season))]
                    temp_StockData = temp_StockData.reset_index(drop=True)
                    temp_S1_S3_SUM+=temp_StockData.iloc[0]["B4"]
                last_year_B4-=temp_S1_S3_SUM
                last_year_B4=round(last_year_B4,2)
            try:
                Price=getStockPriceData.iloc[0]["收盤價"]
            except IndexError:
                sg.Print(f'抓不到 {coid} 的收盤價，請確認股價資料的正確性，必要時請更新本機股號表！')
                Price=0.0
            is_run_season_exist=True
            while run_year >= end_year: #目前年份
                temp_S1_S3_SUM=0.0
                if(count_Season > 4):
                    #print('Out Year')
                    break
                while run_season >= 1 and count_Season <= 4: #目前季度
                    #print(f'COID: {coid} RUNY: {run_year} RUNS : {run_season} countS: {count_Season}')
                    getStockData_PerSeason=self.StockDataDF[(self.StockDataDF["股號"]==coid) & (self.StockDataDF["年份"] == str(run_year)) & (self.StockDataDF["季度"] == str(run_season))]
                    getStockData_PerSeason = getStockData_PerSeason.reset_index(drop=True)
                    if(run_season!=4):
                        recent_EPS+=getStockData_PerSeason.iloc[0]["B4"]
                        #print('近四季',recent_EPS)
                    else:
                        #print('年',run_year,'季度',run_season)
                        for season in range(3,0,-1):
                            temp_StockData=self.StockDataDF[(self.StockDataDF["股號"]==coid) & (self.StockDataDF["年份"] == str(run_year)) & (self.StockDataDF["季度"] == str(season))]
                            temp_StockData = temp_StockData.reset_index(drop=True)
                            print(f'年份： {run_year} 季度: {season}')
                            print(temp_StockData)
                            try:
                                temp_S1_S3_SUM+=temp_StockData.iloc[0]["B4"]
                            except IndexError:
                                sg.Print(f'在讀入 {coid} 股號時之 {run_year} 年的第 {run_season} 季度時發生錯誤\n請確定有抓取該股號當年的第 {season} 季度的財務報告！')
                                is_run_season_exist=False
                                break
                            #print(temp_S1_S3_SUM)
                        if(not(is_run_season_exist)):
                            break
                    if(not(is_run_season_exist)):
                        break
                    S4_EPS=round(getStockData_PerSeason.iloc[0]["B4"],2)
                    #print('第四季EPS',S4_EPS)
                    #print('前三季總和',temp_S1_S3_SUM)
                    S4_EPS-=temp_S1_S3_SUM
                    #print('第四季EPS-前三季EPS',round(S4_EPS,2))
                    recent_EPS+=round(S4_EPS,2)
                    recent_EPS=round(recent_EPS,2)
                    print('近四季',recent_EPS)
                    run_season-=1
                    count_Season+=1
                    #print(f'{getStockData_PerSeason}')
                run_season=4
                run_year-=1
                #print('Out Season')
            if(not(is_run_season_exist)):
                 continue
            dict={'股號':str(coid),'名稱':name,'年份':str(self.db_Data_Newest_Year),'季度':str(self.db_Data_Newest_Season),'A1':A1,'A2':A2,'A3':A3,'A4':A4,'A5':A5,'A6':A6,'A7':A7,'B1':B1,'B2':B2,'B3':B3,'B4':B4,'去年同期B2':last_year_B2,'去年同期B4':last_year_B4,'去年同期B3':last_year_B3,'股價':Price,'近四季 EPS':recent_EPS}
            self.calcDataDF = self.calcDataDF.append(dict, ignore_index=True,sort=False)
            cols=['股號','名稱','年份','季度','A1','A2','A3','A4','A5','A6','A7','B1','B2','B3','B4','去年同期B2','去年同期B3','去年同期B4','股價','近四季 EPS']
            self.calcDataDF = self.calcDataDF[cols]
            print(f'當年當季資料：{getStockData_StartTime}\n去年同期資料：{getStockData_LastYear}\n今日股價：{Price}\n近四季EPS：{recent_EPS}')
        print(self.calcDataDF)
        print(self.calcDataDF.dtypes)
        self.is_calcDataDF_Ready=True
        return True

    def init_calc(self,FormulaType):
        coid_list=[]
        if(self.load_MixData(True)):
            if(not(self.is_calcDataDF_Ready,True)):
                self.calcDataDF=DataFrame()
            self.calcAnsDF=DataFrame()
            self.tableType = FormulaType
            self.Date = datetime.today().strftime("%Y-%m-%d")
            #self.Date = "2021-02-23"
            coid_list=self.StockDataDF["股號"].drop_duplicates().tolist()
            self.db_Data_Newest_Year = self.StockDataDF["年份"].max()
            self.db_Data_Newest_Season = self.StockDataDF.loc[self.StockDataDF["年份"]==self.db_Data_Newest_Year,"季度"].max()
            print(f'年份：{self.db_Data_Newest_Year} 季度：{self.db_Data_Newest_Season}')
            if(self.get_calc_Formula_var(coid_list)):
                coid_list=self.calcDataDF["股號"].drop_duplicates().tolist()
                if(self.calc_Forumla(coid_list,FormulaType)):
                    print('True')
                    self.calcAnsDF.sort_values(by='答案',ascending=True)
                    self.tableDF = self.calcAnsDF
                    self.load_AnsTable()
                    return True
                else:
                    sg.popup_error('讀取時發生意外！已中斷操作！')
                    main_Window.normal()
            else:
                sg.popup_error('讀取時發生意外！已中斷操作！')
                main_Window.normal()
    def load_AnsTable(self):
        self.update_TableData()
        pass
    def update_Window(self):
        global displayDB_Window
        self.update_TableData()
        displayDB_Window.close()
        displayDB_Window=None
        displayDB_Window=self.set_display_DB_Data()
        displayDB_Window.bring_to_front()

    def update_TableWithoutColChange(self):
        displayDB_Window['display_Table'].update(values=self.table_List,num_rows=min(30,len(self.table_List)))

    def update_TableData(self):
        self.table_List=self.tableDF.values.tolist()
        self.table_Heading=list(self.tableDF.head())

    def load_StockPriceTable(self):
        is_Exist=True
        is_Exist=self.load_StockPriceData()
        self.update_TableData()
        return is_Exist

    def load_StockDataTable(self):
        is_Exist=True
        is_Exist=self.load_StockData()
        self.update_TableData()
        return is_Exist

    def load_from_DB(self,query_slot,query):
        temp_list=[]
        current_process=0
        max_process=self.codata.count_documents(filter=query)
        for prt in self.codata.find(query, query_slot):
            current_process+=1
            not_manual_cancel = sg.one_line_progress_meter('從資料庫讀入資料..',current_process,max_process,'calc',orientation='h')
            if(not not_manual_cancel and current_process < max_process):
                Button = sg.popup_yes_no('確定取消？',title='手動取消')
                if(Button=='Yes'):
                    sg.popup_ok('已手動取消！')
                    return False
            temp_list.append(prt)
        self.tableDF=DataFrame(temp_list)
        return True

    def load_StockPriceData(self):
        query_slot = {"_id": 0,"DATA_TYPE":0} #過濾輸出query
        query = {"DATA_TYPE": "股價資料"} #過濾搜尋query
        if(not self.load_from_DB(query_slot,query)):
            return False
        try:
            self.tableDF=self.tableDF[['CO_ID','SYear','SDate','CO_SHORT_NAME','Price','SUB_DATA_TYPE']]
        except KeyError:
            sg.popup('找不到股價資料')
            return False
        self.tableType ='股價資料'
        self.clean_Data()
        self.tableDF=self.tableDF.rename(columns={"CO_ID":"股號","SYear":"年份","SDate":"收盤日","CO_SHORT_NAME":"公司縮寫","Price":"收盤價","SUB_DATA_TYPE":"隸屬交易所"})
        #print(self.tableDF)
        self.tableDF=self.tableDF.dropna()
        self.StockPriceDF=self.tableDF
        print(self.StockPriceDF)
        return True

    def load_MixData(self,isCalc):
        if(self.is_calcDataDF_Ready and isCalc):
            return True
        else:
            return (self.load_StockData() and self.load_StockPriceData())
        pass

    def load_StockData(self):
        query_slot = {"_id": 0,"DATA_TYPE":0,"SUB_DATA_TYPE":0} #過濾輸出query
        query = {"DATA_TYPE": "財務報告"} #過濾搜尋query
        self.load_from_DB(query_slot,query)
        try:
            self.tableDF=self.tableDF[['CO_ID','SYear','SSeason','CO_FULL_NAME','A1','A2','A3','A4','A5','A6','A7','B1','B2','B3','B4']]
        except KeyError:
            sg.popup('找不到財務報告資料')
            return False
        self.tableType = '財務報告'
        self.clean_Data() #清空列表
        self.tableDF=self.tableDF.rename(columns={"CO_ID":"股號","SYear":"年份","SSeason":"季度","CO_FULL_NAME":"公司全名"})
        self.tableDF=self.tableDF.dropna()
        print(self.tableDF)
        self.StockDataDF=self.tableDF
        return True
        

    def init_MongoDB(self):
        self.Date = datetime.today().strftime("%Y-%m-%d")
        self.mongoDB_DBName = conf.get('MongoDB', 'DBNAME')
        self.mongoDB_CData = conf.get('MongoDB', 'CDATANAME')
        self.update_TableData()
        #self.load_StockDataTable()

    def sort_table(self,key,order1,order2):
        order_com1=False
        order_com2=False
        if (order1=='由大到小'):
            order_com1=False
        else:
            order_com1=True
        if (order2=='由大到小'):
            order_com2=False
        else:
            order_com2=True

        self.tableDF=self.tableDF.sort_values(by=key,ascending=[order_com1,order_com2],axis=0)
        self.update_TableData()
        self.update_TableWithoutColChange()
        pass
    def set_display_DB_Data(self):
        displayDB_Layout = [
            [sg.Text('目前顯示'), sg.Text(self.tableType, k='display_Type')],
            [sg.Table(values=self.table_List, auto_size_columns=False, def_col_width=10,justification="left",
                      headings=self.table_Heading, num_rows=min(30,len(self.table_List)), select_mode="extended",
                      enable_events=True, key='display_Table', bind_return_key=True, vertical_scroll_only=False)],
            [sg.Text('過濾條件\t'), sg.Text('過濾數值'), sg.Input(k='Input_Filter', size=(35, 1),enable_events=True)],
            [sg.Text('排序資料'), sg.Combo(self.table_Heading, default_value=self.table_Heading[0], k='Order_Data_1', size=(10, 1),readonly=True,enable_events=True),sg.Combo(self.table_Heading, default_value=self.table_Heading[1], k='Order_Data_2', size=(10, 1),readonly=True,enable_events=True)],
            [sg.Text('順序類型'), sg.Combo(['由大到小', '由小到大'], default_value='由大到小', k='Order_Type_1', size=(
                10, 1),readonly=True,enable_events=True),sg.Combo(['由大到小', '由小到大'], default_value='由大到小', k='Order_Type_2', size=(
                10, 1),readonly=True,enable_events=True)],
            [sg.Text('動作\t'), sg.Button('匯出'), sg.Button('關閉'),
             sg.Button('讀取財務報告'), sg.Button('讀取股價資料'),sg.Button(
            '查閱欄位變數')],
        ]
        return sg.Window("顯示資料庫資料", displayDB_Layout, resizable=True, margins=(5, 5), finalize=True, modal=False, element_justification="center")


# 視窗設計

def set_Force_Exit():
    Force_Exit_Layout = [
        [sg.Text('爬蟲運行中...')],
    ]
    return sg.Window("動作中...", Force_Exit_Layout, margins=(20, 10), finalize=True, modal=True, no_titlebar=True)


def set_local_CSV_Remove_Row(coid, coname):  # 編輯本機股號表 -> 刪除單筆資料
    Remove_Row_Layout = [
        [sg.Text(f'刪除 {coid} {coname} ？')],
        [sg.Button('是'), sg.Button('否')]
    ]
    return sg.Window("刪除單筆資料", Remove_Row_Layout, margins=(30, 10), finalize=True, modal=True, no_titlebar=True)


def set_local_CSV_Add_Row():  # 編輯本機股號表 -> 新增單筆資料
    add_Row_Layout = [
        [sg.Text('公司股號：'), sg.Input(default_text='', size=(5, 1), k='COID')],
        [sg.Text('公司名稱：'), sg.Input(default_text='', size=(25, 1), k='CONAME')],
        [sg.Button('保存'), sg.Button('取消')]
    ]
    return sg.Window("新增單筆資料", add_Row_Layout, margins=(30, 10), finalize=True, modal=True, no_titlebar=True)


def set_local_CSV_Edit_Row(coid, coname):  # 編輯本機股號表 -> 編輯單筆資料
    edit_Row_Layout = [
        [sg.Text('公司股號：'), sg.Input(default_text=coid, size=(5, 1), k='COID')],
        [sg.Text('公司名稱：'), sg.Input(
            default_text=coname, size=(25, 1), k='CONAME')],
        [sg.Button('保存'), sg.Button('取消')]
    ]
    return sg.Window("編輯單筆資料", edit_Row_Layout, margins=(30, 10), finalize=True, modal=True, no_titlebar=True)


def set_local_CSV_Import_usercsvfile_mode():  # 匯入外部股號表 -> 匯入模式
    usercsvfile_Mode_Layout = [
        [sg.Text('選擇匯入模式')],
        [sg.Radio('取代－取代整個股號表', key='ucMode_Replace', group_id='usercsvMode')],
        [sg.Radio('加入－將CSV內的股號表新增至目前已有', key='ucMode_Add',
                  group_id='usercsvMode')],
        [sg.Button('確定'), sg.Button('取消')]
    ]
    return sg.Window("匯入模式", usercsvfile_Mode_Layout, margins=(40, 20), finalize=True, modal=True, no_titlebar=True)


def set_AutoMode_Window():  # 多筆模式 -> 自動爬取來源
    autoMode_Layout = [
        [sg.Text('選擇股號來源')],
        [sg.Radio('從本機股號表讀入', group_id='AM_LoadMode', key='_loadFromLocal', default=True), sg.Radio(
            '從CSV檔匯入', group_id='AM_LoadMode', key='_loadFromCSV')],
        [sg.Button('確定'), sg.Button('取消')],
        [sg.Text('本機股報表位於：\n'+csvpath)]
    ]
    return sg.Window("選擇資料來源", autoMode_Layout, margins=(20, 10), finalize=True, modal=True, no_titlebar=True)


def set_manual_Spider_Stock_Window():  # 爬取模式 -> 單筆模式
    manual_Stock_Spider_Layout = [
        [sg.Text('輸入股號：'), sg.Input(size=(5, 1), k='_Manual_coid')],
        [sg.Combo(year_List, size=(6, 5), key='_StartSearchYear', default_value=this_Year, enable_events=True, readonly=True), sg.Text('年'), sg.Combo(
            this_year_season_List, size=(2, 5), key='_StartSearchSeason', default_value=this_season, enable_events=True, readonly=True), sg.Text('季度')],
        [sg.Button('確定'), sg.Button('返回')]
    ]
    return sg.Window('單筆爬取財務報表', manual_Stock_Spider_Layout, margins=(30, 10), finalize=True, modal=True, no_titlebar=True)


def set_auto_Spider_Stock_Window():  # 爬取模式 -> 多筆模式
    auto_Stock_Spider_Layout = [
        [sg.Text('年度與季度')],
        [sg.Combo(year_List, size=(6, 5), key='_StartSearchYear', default_value=this_Year, enable_events=True, readonly=True), sg.Text('年'), sg.Combo(
            this_year_season_List, size=(2, 5), key='_StartSearchSeason', default_value=this_season, enable_events=True, readonly=True), sg.Text('季度')],
        [sg.Button('確定'), sg.Button('返回')]
    ]
    return sg.Window('多筆爬取財務報表', auto_Stock_Spider_Layout, margins=(30, 10), finalize=True, modal=True, no_titlebar=True)


def set_Spider_Stock_Select_Mode_Window():  # 主視窗 -> 爬取模式
    stock_Spider_Layout = [
        [sg.Text('選擇爬取模式：')],
        [sg.Radio('多筆－調用CSV檔批次抓取', group_id='SMode', default=True, key='_Auto'), sg.Radio(
            '單筆－輸入單個股號抓取', group_id='SMode', key='_Manual')],
        [sg.Button('確定'), sg.Button('返回')]
    ]
    return sg.Window('爬取財務報表模式', stock_Spider_Layout, margins=(30, 10), finalize=True, modal=True, no_titlebar=True)


def set_Main_Window():  # 主視窗
    main_Layout = [
        [sg.Text('[資料庫]')],
        [sg.Button('存取資料庫', disabled=(not DB_READY),tooltip='存取資料庫內已有的資料：財務報告、股價資料，可以按順序排列後匯出報表檔。'),
         sg.Button('連接資料庫', visible=(not DB_READY),tooltip='資料庫屬於離線狀態，點擊即可重新嘗試連線。')],
        [sg.Text('[網路爬蟲]')],
        [sg.Button('開始爬取財務報告', disabled=((not DB_READY) or len(user_Coid_CSV_List)==0),tooltip='到 TWSE 爬取一系列或單筆股號的財務報告')],
        [sg.Button('開始爬取股價資料', disabled=((not DB_READY) or len(user_Coid_CSV_List)==0),tooltip='到 TWSE與 TPEX 爬取一系列的收盤價資訊。')],
        [sg.Text('[運行計算式]')],
        [sg.Combo(['公式一', '公式二', '公式三', '公式四', '公式五', '公式六'],
                  default_value='公式一', k='Combo_Formula', size=(8, 1), readonly=True, enable_events=True),
         sg.Text('公式詳情'), sg.Text('[ ( A1 + A2 + A3 + A4 + A5 ) - A6 ] / ( A7 / 10 ) - Price', k='Combo_Formula_Full', size=(65, 1))],
        [sg.Button('計算', disabled=(not DB_READY),tooltip='將資料庫內的財務報告與股價資料讀入後以選擇的公式進行運算。'), sg.Button(
            '查閱公式變數', disabled=(not DB_READY),tooltip='瞭解公式內的變數意義。')],
        [sg.Text('其他選項')],
        [sg.Button('編輯本機股號表',tooltip='編輯、刪除股號表中的股號，獲釋匯入外部股號表至本機股號表。'), sg.Button('設定',tooltip='設定資料庫連線，編輯或刪除資料庫，主題等相關設定。'), sg.Button(
            '離開'), sg.Button('關於',tooltip='DS'), sg.Button('說明',tooltip='獲取該頁面的說明')]
    ]
    return sg.Window("股票資料抓取與運算", main_Layout, margins=(40, 20), finalize=True)


def set_Local_CSV_Window():  # 主視窗 -> 編輯本機股號表
    local_Coid_CSV_Layout = [
        [sg.Text('輸入股號或公司名稱過濾'), sg.Input(size=(25, 1),
                                          k='filter_data', enable_events=True), sg.Button('清除過濾')],
        [sg.Text('（若輸入中文後列表沒更新，請輕按Shift。）')],
        [sg.Button('復原', disabled=not(local_Coid_CSV_is_changed), k='backup_btn')],
        [sg.Table(values=user_Coid_CSV_List,
                  headings=['代號', '名稱'],
                  auto_size_columns=False,
                  display_row_numbers=False,
                  num_rows=25, select_mode="browse", enable_events=False,
                  key='_local_Coid_CSV_Table', right_click_menu=['右鍵', ['編輯', '刪除']], justification='center', bind_return_key=True)],
        [sg.Button('新增單筆資料',tooltip='輸入單筆股號與名稱。')],
        [sg.Button('關閉且「不保存」變更',tooltip='不做任何變更後關閉視窗。'), sg.Button('關閉且「保存」變更',tooltip='保存並關閉。')],
        [sg.Button('保存當前變更',tooltip='保存目前變更。'), sg.Button('重新整理',tooltip='從原始的資料庫中重新載入。')],
        [sg.Button('匯入外部股號表',tooltip='匯入外部的股號表，可以加入或取代目前現有之本機股號表'), sg.Button('重置本機股號表',tooltip='清空本機股號表')],
        [sg.Text(f'本機股號表CSV位於{csvpath}')]
    ]
    return sg.Window("編輯本機股號表", local_Coid_CSV_Layout, grab_anywhere=False, finalize=True, modal=True, disable_close=False, disable_minimize=False,
                     force_toplevel=True, no_titlebar=False)


def set_Setting_Window():  # 主視窗 -> 設定
    setting_Layout = [
        [sg.Text(f'設定檔的路徑位於：{profile_PATH+_file_name_setting_ini}')],
        [sg.Text('MongoDB －你絕大多數不用更動這個選項，此選項區是關於資料庫連接有關與存放爬取資料的相關設定。')],
        [sg.Text('MongoDB 連結－設定資料庫的位置與登入方法等')],
        [sg.Input(default_text=(conf['MongoDB']['MONGO_URI']),
                  size=(80, 1), k='mDBUrI')],
        [sg.Text('MongoDB 資料庫 － 選擇要存取的資料庫')],
        [sg.Combo(DB_LIST, default_value=(conf['MongoDB']['DBNAME']), size=(
            30, 1), k='mDBName', readonly=True, enable_events=True,disabled=(not DB_READY))],
        [sg.Text('MongoDB 資料集 － 選擇上述資料庫中要存取的資料集')],
        [sg.Combo(CODATA_LIST, default_value=(conf['MongoDB']['CDATANAME']), size=(
            30, 1), k='mCDName', readonly=True,disabled=(not DB_READY))],
        [sg.Text('介面主題'),sg.Combo(theme_list,default_value=sg.theme(),size=(20,1),readonly=True,k='mTheme',enable_events=True)],
        [sg.Button('保存'), sg.Button('取消'), sg.Button('重置')],
        [sg.Button('開啟設定目錄'), sg.Button('管理資料庫與資料集',disabled=(not DB_READY))]
    ]

    return sg.Window("程式設定", setting_Layout, margins=(10, 5), finalize=True, modal=True, disable_close=True, disable_minimize=True, no_titlebar=True)


MDB_Load = None

main_Window, setting_Window, aM_Window, local_Csv_Window, local_Csv_imode_Window = set_Main_Window(
), None, None, None, None
csv_Row_Edit_Window, csv_Row_Add_Window = None, None
Spider_Stock_Select_Mode_Window, Spider_Stock_Price_Window = None, None
auto_Spider_Stock_Window, manual_Spider_Stock_Window = None, None
Force_Exit_Window = None
displayDB_Window = None

main_Window.bring_to_front()
if(DB_READY):
    scrapyer.change_Project_Setting(str(conf.get('MongoDB', 'MONGO_URI')), str(
        conf.get('MongoDB', 'DBNAME')), str(conf.get('MongoDB', 'cdataname')))
    MDB_Load = MongoDB_Load()

print('主視窗載入完成。')
while True:  # 監控視窗回傳
    window, event, values = sg.read_all_windows()
    # sg.Print(f'Window:{window},event:{event},values:{values}')
    if window == main_Window:  # 主視窗
        if event == '說明':
            sg.popup_ok(main_Window_Help,'主視窗說明',no_titlebar=True)
        if event == 'Combo_Formula':
            if values['Combo_Formula'] == '公式一':
                window['Combo_Formula_Full'].update(
                    value="[ ( A1 + A2 + A3 + A4 + A5 ) - A6 ] / ( A7 / 10 ) - Price")
            if values['Combo_Formula'] == '公式二':
                window['Combo_Formula_Full'].update(
                    value="｛( B4 - 去年同期的 B4 ) / ( 去年同期的 B4 ) ｝x 100 - Price / 近四季 EPS")
            if values['Combo_Formula'] == '公式三':
                window['Combo_Formula_Full'].update(
                    value="｛( B2 - 去年同期 B2 ) / 去年同期 B2｝x 100 - Price / 近四季 EPS")
            if values['Combo_Formula'] == '公式四':
                window['Combo_Formula_Full'].update(
                    value="( B2 - 去年同期 B2 ) > 1 or ( B3- 去年同期 B3 ) < 1")
            if values['Combo_Formula'] == '公式五':
                window['Combo_Formula_Full'].update(
                    value="( B2 - 去年同期 B2 ) > 1")
            if values['Combo_Formula'] == '公式六':
                window['Combo_Formula_Full'].update(
                    value="( B3- 去年同期 B3 ) < 1")
        if event in (sg.WIN_CLOSED, '離開'):
            break
        if event == "開始爬取股價資料":
            if(check_Mongo()):
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
                sg.popup(
                    '由於Scrapy框架的天生限制。\n在執行完一個爬蟲之後程式將會自動關閉，手動開啟後得以進行下一個爬蟲作業。', title='注意')
                Spider_Stock_Price_Window = set_AutoMode_Window()
                window.minimize()
        if event == "開始爬取財務報告":
            print(len(user_Coid_CSV_List))
            if(check_Mongo()):
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
                sg.popup(
                    '由於Scrapy框架的天生限制。\n在執行完一個爬蟲之後程式將會自動關閉，手動開啟後得以進行下一個爬蟲作業。', title='注意')
                Spider_Stock_Select_Mode_Window = set_Spider_Stock_Select_Mode_Window()
                window.minimize()
        if event == "存取資料庫":
            if(check_Mongo()):
                main_Window.minimize()
                MDB_Load.init_MongoDB()
                displayDB_Window = MDB_Load.set_display_DB_Data()

        if event == "連接資料庫":
            window.close()
            connect_Mongo(False, False, False, True)
            if(DB_READY):
                MDB_Load = MongoDB_Load()
                scrapyer.change_Project_Setting(str(conf.get('MongoDB', 'MONGO_URI')), str(
                    conf.get('MongoDB', 'DBNAME')), str(conf.get('MongoDB', 'cdataname')))
            main_Window = None
            main_Window = set_Main_Window()

        if event == "計算":
            main_Window.minimize()
            if(MDB_Load.init_calc(str(values['Combo_Formula']))):
                displayDB_Window = MDB_Load.set_display_DB_Data()
            else:
                main_Window.normal()

        if event == "設定":
            window.minimize()
            setting_Window = set_Setting_Window()

        if event == "編輯本機股號表":
            window.minimize()
            local_Csv_Window = set_Local_CSV_Window()
            load_Local_CSV_Table_CSVDF()

        if event == "查閱公式變數":
            sg.popup(formula_Info,'公式變數參考，你可以變更公式時移動視窗來參考。',no_titlebar=True,grab_anywhere=True,non_blocking=True)
        if event == "關於":
            sg.popup('股票資訊爬蟲\n版本： 1.0\n作者：Douggy Sans\n2021年編寫', title='關於')

    if window == displayDB_Window:

        if event == "Input_Filter":
            Order_Data=[str(values['Order_Data_1']),str(values['Order_Data_2'])]
            MDB_Load.sort_table(Order_Data,str(values['Order_Type_1']),str(values['Order_Type_2']))
            MDB_Load.filter_db_Table(values['Input_Filter'])

        if event == "匯出":
            MDB_Load.export_Table()
        if event in (sg.WIN_CLOSED, '關閉'):
            displayDB_Window.close()
            displayDB_Window = None
            main_Window.normal()

        if event == "Order_Data_1" or event == "Order_Data_2"   or event == "Order_Type_1" or event == "Order_Type_2":
            Order_Data=[str(values['Order_Data_1']),str(values['Order_Data_2'])]
            MDB_Load.sort_table(Order_Data,str(values['Order_Type_1']),str(values['Order_Type_2']))
            MDB_Load.filter_db_Table(values['Input_Filter'])

        if event == "查閱欄位變數":
            sg.popup(formula_Info,'公式變數參考，你可以變更公式時移動視窗來參考。',no_titlebar=True,grab_anywhere=True,non_blocking=True)           
        if event == "讀取財務報告":
            if(MDB_Load.load_StockDataTable()):
                MDB_Load.update_Window()

        if event == "讀取股價資料":
            if(MDB_Load.load_StockPriceTable()):
                MDB_Load.update_Window()

    if window == setting_Window:  # 主視窗 -> 設定視窗之互動
        window.bring_to_front()
        if event in (sg.WIN_CLOSED, '取消'):
            sg.theme(conf.get('System','Theme'))
            window.close()
            setting_Window = None
            main_Window.normal()

        if event == 'mDBName':
            print(values['mDBName'])
            db = DBClient.get_database(str(values['mDBName']))
            codata_list = db.list_collection_names()
            print(codata_list)
            setting_Window['mCDName'].update(
                value=codata_list[0], values=codata_list)
        if event == '管理資料庫與資料集':
            connect_Mongo(False, True, True, True)
            setting_Window.disable()
            setting_Window.make_modal()
            DB_LIST = DBClient.list_database_names()
            DB_LIST.remove('local')
            DB_LIST.remove('config')
            DB_LIST.remove('admin')
            db = DBClient.get_database(str(values['mDBName']))
            #db = DBClient.get_database(str(conf.get('MongoDB', 'DBNAME')))
            db.drop_collection('init')
            CODATA_LIST = db.list_collection_names()

            if(len(CODATA_LIST) == 0):
                mCDNAME_value = ''
            else:
                mCDNAME_value = CODATA_LIST[len(CODATA_LIST)-1]

            if(len(DB_LIST) == 0):
                mDBName_value = ''
            else:
                mDBName_value = DB_LIST[len(DB_LIST)-1]

            setting_Window['mCDName'].update(
                value=mCDNAME_value, values=CODATA_LIST)
            setting_Window['mDBName'].update(
                value=mDBName_value, values=DB_LIST)
            setting_Window.enable()
            setting_Window.make_modal()
        if event == "mTheme":
            theme = sg.theme(str(values['mTheme']))
            setting_Window.close()
            setting_Window=set_Setting_Window()
        if event == "保存":
            print("Save"+str(values['mDBUrI']),
                  str(values['mDBName']), str(values['mCDName']))
            conf.set('MongoDB', 'DBNAME', str(values['mDBName']))
            conf.set('MongoDB', 'cdataname', str(values['mCDName']))
            conf.set('System','Theme',str(values['mTheme']))
            check_Mongo()
            if(DB_READY):
                conf.set('MongoDB', 'MONGO_URI', str(values['mDBUrI']))
                conf.write(open(cfgpath, 'w'))
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                sg.popup('已保存設定！', title='已保存')
                scrapyer.change_Project_Setting(str(conf.get('MongoDB', 'MONGO_URI')), str(
                    conf.get('MongoDB', 'DBNAME')), str(conf.get('MongoDB', 'cdataname')))
                window.close()
                setting_Window = None
                main_Window.close()
                main_Window=set_Main_Window()
            else:
                setting_Window.make_modal()

        if event == "重置":
            winsound.MessageBeep(winsound.MB_ICONQUESTION)
            if(sg.popup_ok_cancel('是否重置設定？', title='確認重置', modal=True) == 'OK'):
                reset_setting()
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                sg.popup('已重置！')
                theme = sg.theme(str(values['mTheme']))
                main_Window.close()
                main_Window = set_Main_Window()
                setting_Window.close()
        if event == "開啟設定目錄":
            os.startfile(profile_PATH)

    if window == Spider_Stock_Price_Window:  # 主視窗 -> 股價爬蟲 -> 選擇來源
        window.bring_to_front()
        if event in (sg.WIN_CLOSED, '取消'):
            window.close()
            Spider_Stock_Price_Window = None
            main_Window.normal()
        if event == '確定':
            if(values['_loadFromLocal']):
                window.close()
                Spider_Stock_Price_Window = None
                call_Price_Spider(True, '')
            else:
                custom_Price_Spider_CSVPATH = sg.popup_get_file(
                    '讀入外部股號表', no_window=True, file_types=(("外部CSV股號表", "*.csv"),))
                if(custom_Price_Spider_CSVPATH != ''):
                    window.close()
                    Spider_Stock_Price_Window = None
                    call_Price_Spider(False, custom_Price_Spider_CSVPATH)
                else:
                    window.close()
                    Spider_Stock_Price_Window = None

    if window == aM_Window:  # 主視窗 -> 財務批次爬蟲 -> 選擇來源
        window.bring_to_front()
        if event in (sg.WIN_CLOSED, '取消'):
            window.close()
            aM_Window = None
            main_Window.normal()
        if event == '確定':
            if(values['_loadFromLocal']):
                window.close()
                aM_Window = None
                call_Stock_Spider(True, True, '', '')
            else:
                custom_Stock_Spider_CSVPATH = sg.popup_get_file(
                    '讀入外部股號表', no_window=True, file_types=(("外部CSV股號表", "*.csv"),))
                if(custom_Stock_Spider_CSVPATH != ''):
                    window.close()
                    aM_Window = None
                    call_Stock_Spider(
                        True, False, custom_Stock_Spider_CSVPATH, '')
                else:
                    window.close()
                    aM_Window = None

    if window == csv_Row_Add_Window:  # 編輯本機股號表 -> 新增單筆資料
        if event == '保存':
            co_id = values['COID']
            co_name = values['CONAME']
            if((co_id or co_name != '') and (co_id != co_name)):
                local_CSV_Row_Add(co_id, co_name)
                csv_Row_Add_Window.close()
                csv_Row_Add_Window = None
                local_Csv_Window.make_modal()
            elif(co_id == co_name):
                winsound.MessageBeep(winsound.MB_ICONHAND)
                sg.popup_error('股號與公司名稱重複！')
                window.make_modal()
            else:
                winsound.MessageBeep(winsound.MB_ICONHAND)
                sg.popup_error('負號與公司欄位請勿留空！')
                window.make_modal()

        if event == '取消':
            csv_Row_Add_Window.close()
            csv_Row_Add_Window = None
            local_Csv_Window.make_modal()

    if window == local_Csv_Window:  # 主視窗 -> 編輯本機股號表
        window['_local_Coid_CSV_Table'].bind("bind_return_key", "編輯")
        if event == 'backup_btn':
            local_CSV_Restore_USER_DF()
            refresh_Local_CSV_Table()
        if event == '新增單筆資料':
            local_Coid_CSV_is_filter = False
            window['filter_data'].update('')
            refresh_Local_CSV_Table()
            csv_Row_Add_Window = set_local_CSV_Add_Row()
        if event == '清除過濾':
            local_Coid_CSV_is_filter = False
            window['filter_data'].update('')
            refresh_Local_CSV_Table()
        if event == 'filter_data':
            if(values['filter_data'] == ''):
                local_Coid_CSV_is_filter = False
                # sg.Print('Null')
                refresh_Local_CSV_Table()
            else:
                local_Coid_CSV_is_filter = True
                filter_String = values['filter_data']
                filter_Local_CSV_Table(filter_String)
        if event == '保存當前變更':
            save_Local_CSV()
        if event in ('關閉且「不保存」變更', sg.WIN_CLOSED):
            update_Local_CSV_Table()
            window.close()
            local_Csv_Window = None
            main_Window.close()
            main_Window=set_Main_Window()
        if event == '關閉且「保存」變更':
            print(user_df)
            save_Local_CSV()
            backup_Coid_pd_df.clear()
            window.close()
            local_Csv_Window = None
            main_Window.close()
            main_Window=set_Main_Window()
        if event == "重新整理":
            if(local_Coid_CSV_is_changed):
                winsound.MessageBeep(winsound.MB_ICONQUESTION)
                if(sg.popup_yes_no('你股號表尚未存檔，重新整理將會喪失變更的資料並還原修改前的樣子，確定嗎？', title='重新整理', modal=True, no_titlebar=True) == 'Yes'):
                    update_Local_CSV_Table()
            else:
                update_Local_CSV_Table()
            window.make_modal()
        if event == "匯入外部股號表":
            local_CSV_Import_usercsvfile()
        if event == "重置本機股號表":
            winsound.MessageBeep(winsound.MB_ICONQUESTION)
            if(sg.popup_yes_no('是否重置股號表？', title='確認重置', modal=True) == 'Yes'):
                local_CSV_Backup_USER_DF()
                reset_Local_CSV_Table()
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                sg.popup('已重置！')
            window.make_modal()
        if event == "刪除":
            select_row = values['_local_Coid_CSV_Table']
            if select_row != []:
                select_row = int(select_row[0])
                local_CSV_Row_Edit(False, select_row)
            else:
                winsound.MessageBeep(winsound.MB_ICONHAND)
                sg.popup('請先選擇有效的項目進行來編輯')
            window.make_modal()
        if event == "編輯" or event == '_local_Coid_CSV_Table':
            select_row = values['_local_Coid_CSV_Table']
            if select_row != []:
                select_row = int(select_row[0])
                local_CSV_Row_Edit(True, select_row)
            else:
                winsound.MessageBeep(winsound.MB_ICONHAND)
                sg.popup('請先選擇有效的項目進行來編輯')
            window.make_modal()

    if window == manual_Spider_Stock_Window:  # 爬取模式選擇 -> 單筆模式
        if event == "_StartSearchYear":
            if(values['_StartSearchYear'] != this_Year):
                window['_StartSearchSeason'].update(
                    values=season_List, set_to_index=0)
            else:
                window['_StartSearchSeason'].update(
                    values=this_year_season_List, set_to_index=0)
        if event == '確定':
            window.close()
            search_Year = str(values['_StartSearchYear'])
            search_Season = str(values['_StartSearchSeason'])
            manual_Spider_Stock_Window = None
            call_Stock_Spider(False, False, '', str(values['_Manual_coid']))
        if event == '返回':
            window.close()
            manual_Spider_Stock_Window = None

    if window == auto_Spider_Stock_Window:  # 爬取模式選擇 -> 批次模式
        if event == "_StartSearchYear":
            if(values['_StartSearchYear'] != this_Year):
                window['_StartSearchSeason'].update(
                    values=season_List, set_to_index=0)
            else:
                window['_StartSearchSeason'].update(
                    values=this_year_season_List, set_to_index=0)
        if event == '確定':
            window.close()
            search_Year = str(values['_StartSearchYear'])
            search_Season = str(values['_StartSearchSeason'])
            auto_Spider_Stock_Window = None
            aM_Window = set_AutoMode_Window()
        if event == '返回':
            window.close()
            auto_Spider_Stock_Window = None

    if window == Spider_Stock_Select_Mode_Window:  # 主視窗 ->爬取模式選擇
        if event == '確定':
            if(values['_Auto']):
                window.close()
                Spider_Stock_Select_Mode_Window = None
                auto_Spider_Stock_Window = set_auto_Spider_Stock_Window()
            else:
                window.close()
                Spider_Stock_Select_Mode_Window = None
                manual_Spider_Stock_Window = set_manual_Spider_Stock_Window()
        if event == '返回':
            window.close()
            Spider_Stock_Select_Mode_Window = None
            main_Window.normal()
window.close()