import os.path as path
import os
import PySimpleGUI as sg
import configparser
import pandas as pd
from datetime import datetime

from StockScrapyProject.StockScrapyProject.run_scraper import Scraper
import math
import pymongo

sg.theme('DarkAmber')  # 設定顏色主題
sg.set_options(auto_size_buttons=True)


# 全域變數
this_Year = datetime.today().year  # 獲取今年年份
this_month = datetime.today().month  # 獲取這個月份
this_season = math.ceil(this_month/4)  # 換算季度
year_List = []  # 存放年份
season_List = ['1', '2', '3', '4']  # 存放季度
this_year_season_List = []

DBClient = pymongo.MongoClient()
DB_Connect_Status = False
DB_CODATA_Exist = False
DB_READY = DB_Connect_Status and DB_CODATA_Exist
DB_LIST = []
CODATA_LIST = []

search_Year = ''
search_Season = ''

for i in range(1, this_season+1):
    this_year_season_List.append(str(i))

# 資料存放
user_Coid_CSV_List = []  # 暫存股號表存放
filter_Coid_CSV_List = []  # 過濾用存放
backup_Coid_pd_df = []  # 備份user的df
filter_String = ''  # 過濾字串

# State
local_Coid_CSV_is_filter = False  # 存放過濾狀態
local_Coid_CSV_is_changed = False  # 存放資料異動狀態

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
csvpath = os.path.join(profile_PATH+_file_name_local_csv)  # 本地股號表路徑
conf = configparser.ConfigParser()  # 創建設定檔對象


# PD.DF設置
coid_dict = {"代號": [], "名稱": []}  # 建立空的本地股號列表
coid_dict_type = {'代號': 'string', '名稱': 'string'}  # 建立股號列表檔案類型
local_csvdf = pd.DataFrame(coid_dict)  # 導入本地股號表pd使用
user_df = pd.DataFrame(coid_dict)  # 建立暫存本地股號表pd使用
import_csv_df = pd.DataFrame(coid_dict)  # 導入pd使用
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
                  k='MList', readonly=True, size=(12, 1))],
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
                    break

                if events == '刪除' and values['MList'] != '':
                    Button = sg.popup_yes_no(
                        '確定要刪除 '+values['MList']+' 資料庫嗎？', no_titlebar=True)
                    if(Button == 'Yes'):
                        DBClient.drop_database(values['MList'])
                        DB_LIST = DBClient.list_database_names()
                        DB_LIST.remove('local')
                        DB_LIST.remove('config')
                        DB_LIST.remove('admin')
                        set_DB_Window['MList'].update(values=DB_LIST, value='')
                if events == '新建':
                    MongoDBName = '#'
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

                    if events == '取消':
                        set_DB_Window.close()
                        set_DB_Window = None
                        return None

                    if events == '確定':
                        MongoDB_CODATA = values['MList']
                        if(MongoDB_CODATA != ''):
                            CreateCODATA = False
                            breakpoint
                        else:
                            CreateCODATA = True
                            break
                    if events == '刪除' and values['MList'] != '':
                        Button = sg.popup_yes_no(
                            f'確定要刪除 '+values['MList']+' 資料集嗎？', no_titlebar=True)
                        if(Button == 'Yes'):
                            db.drop_collection(values['MList'])
                            CODATA_LIST = db.list_collection_names()
                            set_DB_Window['MList'].update(
                                values=CODATA_LIST, value='')
                    if events == '新建':
                        MongoDB_CODATA = '#'
                        CreateCODATA = True
                        break
                set_DB_Window.close()
                set_DB_Window = None
            print(MongoDB_CODATA)

            if(MongoDB_CODATA in CODATA_LIST and (not CreateCODATA) or isNeedSelect):
                DB_CODATA_Exist = True
                DB_READY = DB_Connect_Status and DB_CODATA_Exist
                conf.set('MongoDB', 'DBNAME', str(MongoDBName))
                conf.set('MongoDB', 'CDATANAME', str(MongoDB_CODATA))
                conf.write(open(cfgpath, 'w'))
                sg.SystemTray.notify(
                    'MonogoDB 已預備完成！', f'資料庫: {MongoDBName} \n資料集: {MongoDB_CODATA}\n功能初始化完成', display_duration_in_ms=1000, fade_in_duration=.2)
            else:
                if(isInit or CreateCODATA):
                    if(not CreateCODATA and not isInit):
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
                                sg.popup(f'在 {MongoDBName} 當中該資料集已存在！')
                                setting_Window.make_modal()
                            db.drop_collection('init')
                            connect_Mongo(True, False, False, False)
                else:
                    sg.popup(
                        f'在 {MongoDBName} 之資料庫中找不到 {MongoDB_CODATA} 資料集，請到設定更改為有效的資料庫名稱。', no_titlebar=True)
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
    conf.set('MongoDB', 'MONGO_URI', str(default_MDUrl))
    conf.write(open(cfgpath, 'w'))
    conf.read(cfgpath, encoding='utf-8')
    sg.popup('已建立設定檔。')
    connect_Mongo(True, False, False, True)


def reset_csv():  # 重建csv檔
    global user_Coid_CSV_List, local_csvdf, user_df
    local_csvdf = pd.DataFrame(coid_dict)  # 導入本地股號表pd使用
    local_csvdf.to_csv(csvpath, index=False, sep=',')
    user_df = local_csvdf
    user_Coid_CSV_List = user_df.values.tolist()
    sg.SystemTray.notify(
        '系統', '已建立本地股號表。', display_duration_in_ms=1000, fade_in_duration=.2)
    # sg.Print('重建本地股號表')


def check_setting():  # 檢查設定
    if(path.exists(profile_PATH+_file_name_setting_ini)):
        sg.SystemTray.notify(
            '系統', '已檢查到設定檔。', display_duration_in_ms=250, fade_in_duration=.2)
        conf.read(cfgpath, encoding='utf-8')
        if(not conf.has_option('MongoDB', 'mongo_uri') or not conf.has_option('MongoDB', 'dbname') or not conf.has_option('MongoDB', 'cdataname')):
            sg.popup_error('系統', '資料庫相關設置遺失！重置設定檔中...')
            reset_setting()
    else:
        sg.SystemTray.notify('系統', '未檢查到設定檔，創建中...',
                             display_duration_in_ms=5000, fade_in_duration=.2)
        reset_setting()


def check_local_csv():  # 檢查本地CSV
    if(path.exists(profile_PATH+_file_name_local_csv)):
        global local_csvdf, user_Coid_CSV_List, user_df
        sg.SystemTray.notify('系統', '已檢查到本地股號表。',
                             display_duration_in_ms=250, fade_in_duration=.2)
        try:
            local_csvdf = pd.read_csv(
                csvpath, sep=',', engine='python', dtype=coid_dict_type, na_filter=False)
            local_csvdf = local_csvdf
            user_Coid_CSV_List = local_csvdf.values.tolist()
            user_df = local_csvdf
        except pd.errors.EmptyDataError:
            sg.popup('讀取本地股號表時發生錯誤！重建本地股號表...')
            reset_csv()
            check_local_csv()
    else:
        sg.popup('未建立本地股號表，創建中...', title='系統')
        reset_csv()

# 初始化函數


scrapyer = Scraper()
check_setting()
check_local_csv()
connect_Mongo(False, False, False, False)

# 爬蟲調用


def call_Price_Spider(isLocal, LOAD_CSVPATH):
    global csvpath, Force_Exit_Window
    info = ''
    if(isLocal):
        scrapyer.set_PriceSpider(csvpath)
        sg.popup_no_buttons('啟用爬蟲中...', grab_anywhere=False,
                            no_titlebar=True, auto_close=True)
        main_Window.close()
        Force_Exit_Window = set_Force_Exit()
        scrapyer.run_PriceSpider()
    else:
        scrapyer.set_PriceSpider(LOAD_CSVPATH)
        sg.popup_no_buttons('啟用爬蟲中...', grab_anywhere=False,
                            no_titlebar=True, auto_close=True)
        main_Window.close()
        Force_Exit_Window = set_Force_Exit()
        scrapyer.run_PriceSpider()
    print('\n')
    sg.popup_ok('抓取股價資料完成！程式將會關閉！')
    os._exit(0)


def call_Stock_Spider(isAutoMode, isLocal, LOAD_CSVPATH, M_CO_ID):
    global csvpath, Force_Exit_Window
    info = ''
    if(isAutoMode):
        if(isLocal):
            scrapyer.set_StockSpider(
                Year=search_Year, Season=search_Season, Mode='Auto', CSV=csvpath)
            sg.popup_no_buttons('連接到資料庫中，請稍後...', non_blocking=True,
                                grab_anywhere=False, no_titlebar=True, auto_close=True)
            main_Window.close()
            Force_Exit_Window = set_Force_Exit()
            scrapyer.run_StockSpider()
        else:
            scrapyer.set_StockSpider(
                Year=search_Year, Season=search_Season, Mode='Auto', CSV=LOAD_CSVPATH)
            sg.popup_no_buttons('連接到資料庫中，請稍後...', non_blocking=True,
                                grab_anywhere=False, no_titlebar=True, auto_close=True)
            main_Window.close()
            Force_Exit_Window = set_Force_Exit()
            scrapyer.run_StockSpider()
    else:
        scrapyer.set_StockSpider(
            Year=search_Year, Season=search_Season, Mode='Manual', CO_ID=M_CO_ID)
        scrapyer.run_StockSpider()
    print('\n')
    sg.popup_ok('抓取股票財務報告完成！程式將會關閉！')
    os._exit(0)


# 普通應用方法

def local_CSV_Restore_USER_DF():  # 還原本地股號表資料狀態
    global user_df, local_Coid_CSV_is_changed, backup_Coid_pd_df, local_Csv_Window
    user_df = backup_Coid_pd_df.pop(-1)
    step = len(backup_Coid_pd_df)
    print(step)
    if(step == 0):
        local_Coid_CSV_is_changed = False
        local_Csv_Window['backup_btn'].update(
            disabled=not(local_Coid_CSV_is_changed))


def local_CSV_Backup_USER_DF():  # 備份本地股號表資料狀態
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


def save_Local_CSV():  # 保存至本地股號表
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


def load_Local_CSV_Table_CSVDF():  # 從本地股號表PD.DF中讀入至表單中
    global local_csvdf, local_Coid_CSV_is_changed, user_Coid_CSV_List
    user_Coid_CSV_List = local_csvdf.values.tolist()
    local_Csv_Window['_local_Coid_CSV_Table'].update(values=user_Coid_CSV_List)


def refresh_Local_CSV_Table():  # 重新整理股號表
    global local_csvdf, local_Coid_CSV_is_changed
    if(local_Csv_Window != None):
        load_Local_CSV_Table_USERVDF()


def update_Local_CSV_Table():  # 從本地股表重新載入
    global local_csvdf, local_Coid_CSV_is_changed
    if(local_Csv_Window != None):
        load_Local_CSV_Table_CSVDF()


def reset_Local_CSV_Table():  # 重置本地股號表
    global local_csvdf, local_Coid_CSV_is_changed
    if(local_Csv_Window != None):
        reset_csv()
        load_Local_CSV_Table_CSVDF()


def local_CSV_Row_Edit(isEdit, index):  # 編輯本地股號表 ->編輯單筆資料
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
    import_csv_df = pd.read_csv(
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
                    Button = sg.popup_ok_cancel('取代目前已有的股號表，確定嗎？')
                    if(Button == "OK"):
                        sg.popup("使用取代模式")
                        local_CSV_usercsvfile_import(True, user_CSV_File_Path)
                        break
                else:
                    Button = sg.popup_ok_cancel('添加至目前已有的股號表，確定嗎？')
                    if(Button == "OK"):
                        sg.popup("使用增加模式")
                        local_CSV_usercsvfile_import(False, user_CSV_File_Path)
                        break
            if event == "取消":
                break
        local_Csv_imode_Window.close()
    local_Csv_imode_Window = None
    window.make_modal()

# 視窗設計


def set_Force_Exit():
    Force_Exit_Layout = [
        [sg.Text('爬蟲運行中...')],
    ]
    return sg.Window("動作中...", Force_Exit_Layout, margins=(20, 10), finalize=True, modal=True, no_titlebar=True)


def set_local_CSV_Remove_Row(coid, coname):  # 編輯本地股號表 -> 刪除單筆資料
    Remove_Row_Layout = [
        [sg.Text(f'刪除 {coid} {coname} ？')],
        [sg.Button('是'), sg.Button('否')]
    ]
    return sg.Window("刪除單筆資料", Remove_Row_Layout, margins=(30, 10), finalize=True, modal=True, no_titlebar=True)


def set_local_CSV_Add_Row():  # 編輯本地股號表 -> 新增單筆資料
    add_Row_Layout = [
        [sg.Text('公司股號：'), sg.Input(default_text='', size=(5, 1), k='COID')],
        [sg.Text('公司名稱：'), sg.Input(default_text='', size=(25, 1), k='CONAME')],
        [sg.Button('保存'), sg.Button('取消')]
    ]
    return sg.Window("新增單筆資料", add_Row_Layout, margins=(30, 10), finalize=True, modal=True, no_titlebar=True)


def set_local_CSV_Edit_Row(coid, coname):  # 編輯本地股號表 -> 編輯單筆資料
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
        [sg.Radio('從本地股號表讀入', group_id='AM_LoadMode', key='_loadFromLocal', default=True), sg.Radio(
            '從CSV檔匯入', group_id='AM_LoadMode', key='_loadFromCSV')],
        [sg.Button('確定'), sg.Button('取消')],
        [sg.Text('本地股報表位於：\n'+csvpath)]
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
        [sg.Combo(year_List, size=(6, 5), key='_StartSearchYear', default_value=this_Year-1, enable_events=True, readonly=True), sg.Text('年'), sg.Combo(
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
        [sg.Button('存取資料庫', disabled=(not DB_READY)),
         sg.Button('連接資料庫', visible=(not DB_READY))],
        [sg.Text('[網路爬蟲]')],
        [sg.Button('開始爬取財務報告', disabled=(not DB_READY))],
        [sg.Button('開始爬取股價資料', disabled=(not DB_READY))],
        [sg.Text('[運行計算式]')],
        [sg.Button('公式一'), sg.Button('公式二'),
         sg.Button('公式三'), sg.Button('公式四')],
        [sg.Text('其他選項')],
        [sg.Button('編輯本地股號表'), sg.Button('設定'), sg.Button(
            '離開'), sg.Button('關於'), sg.Button('說明')]
    ]
    return sg.Window("股票資料抓取與運算", main_Layout, margins=(40, 20), finalize=True)


def set_Local_CSV_Window():  # 主視窗 -> 編輯本地股號表
    local_Coid_CSV_Layout = [
        [sg.Text('輸入股號或公司名稱過濾'), sg.Input(size=(25, 1),
                                          k='filter_data', enable_events=True), sg.Button('清除過濾')],
        [sg.Text('（若輸入中文後列表沒更新，請輕按Shift。）')],
        [sg.Button('復原', disabled=not(
            local_Coid_CSV_is_changed), k='backup_btn')],
        [sg.Table(values=user_Coid_CSV_List,
                  headings=['代號', '名稱'],
                  auto_size_columns=False,
                  display_row_numbers=False,
                  num_rows=25, select_mode="browse", enable_events=False,
                  key='_local_Coid_CSV_Table', right_click_menu=['右鍵', ['編輯', '刪除']], justification='center', bind_return_key=True)],
        [sg.Button('新增單筆資料')],
        [sg.Button('關閉且「不保存」變更'), sg.Button('關閉且「保存」變更')],
        [sg.Button('保存當前變更'), sg.Button('重新整理')],
        [sg.Button('匯入外部股號表'), sg.Button('重置本地股號表')],
        [sg.Text(f'本地股號表CSV位於{csvpath}')]
    ]
    return sg.Window("編輯本地股號表", local_Coid_CSV_Layout, grab_anywhere=False, finalize=True, modal=True, disable_close=True, disable_minimize=True,
                     force_toplevel=True)


def set_Setting_Window():  # 主視窗 -> 設定
    setting_Layout = [
        [sg.Text(f'設定檔的路徑位於：{profile_PATH+_file_name_setting_ini}')],
        [sg.Text('MongoDB －你絕大多數不用更動這個選項，此選項區是關於資料庫連接有關與存放爬取資料的相關設定。')],
        [sg.Text('MongoDB 連結－設定資料庫的位置與登入方法等')],
        [sg.Input(default_text=(conf['MongoDB']['MONGO_URI']),
                  size=(30, 1), k='mDBUrI')],
        [sg.Text('MongoDB 資料庫 － 選擇要存取的資料庫')],
        [sg.Combo(DB_LIST, default_value=(conf['MongoDB']['DBNAME']), size=(
            30, 1), k='mDBName', readonly=True, enable_events=True)],
        [sg.Text('MongoDB 資料集 － 選擇上述資料庫中要存取的資料集')],
        [sg.Combo(CODATA_LIST, default_value=(conf['MongoDB']['CDATANAME']), size=(
            30, 1), k='mCDName', readonly=True)],
        [sg.Button('保存'), sg.Button('取消'), sg.Button('重置')],
        [sg.Button('開啟設定目錄'), sg.Button('管理資料庫與資料集')]
    ]

    return sg.Window("程式設定", setting_Layout, margins=(10, 5), finalize=True, modal=True, disable_close=True, disable_minimize=True)


main_Window, setting_Window, aM_Window, local_Csv_Window, local_Csv_imode_Window = set_Main_Window(
), None, None, None, None
csv_Row_Edit_Window, csv_Row_Add_Window = None, None
Spider_Stock_Select_Mode_Window, Spider_Stock_Price_Window = None, None
auto_Spider_Stock_Window, manual_Spider_Stock_Window = None, None
Force_Exit_Window = None
main_Window.bring_to_front()
if(DB_READY):
    scrapyer.change_Project_Setting(str(conf.get('MongoDB', 'MONGO_URI')), str(
        conf.get('MongoDB', 'DBNAME')), str(conf.get('MongoDB', 'cdataname')))

print('主視窗載入完成。')
while True:  # 監控視窗回傳
    window, event, values = sg.read_all_windows()
    # sg.Print(f'Window:{window},event:{event},values:{values}')
    if window == main_Window:  # 主視窗
        if event in (sg.WIN_CLOSED, '離開'):
            break
        if event == "開始爬取股價資料":
            if(check_Mongo()):
                sg.popup(
                    '由於Scrapy框架的天生限制。\n在執行完一個爬蟲之後程式將會自動關閉，手動開啟後得以進行下一個爬蟲作業。', title='注意')
                Spider_Stock_Price_Window = set_AutoMode_Window()
        if event == "開始爬取財務報告":
            if(check_Mongo()):
                sg.popup(
                    '由於Scrapy框架的天生限制。\n在執行完一個爬蟲之後程式將會自動關閉，手動開啟後得以進行下一個爬蟲作業。', title='注意')
                Spider_Stock_Select_Mode_Window = set_Spider_Stock_Select_Mode_Window()
        if event == "存取資料庫":
            if(check_Mongo()):
                sg.popup('存取資料庫')

        if event == "連接資料庫":
            window.close()
            connect_Mongo(False, False, False, True)
            main_Window = None
            main_Window = set_Main_Window()
        if event == "公式一":
            sg.popup('執行公式1')

        if event == "公式二":
            sg.popup('執行公式2')

        if event == "公式三":
            sg.popup('執行公式3')

        if event == "公式四":
            sg.popup('執行公式4')

        if event == "設定":
            setting_Window = set_Setting_Window()

        if event == "編輯本地股號表":
            local_Csv_Window = set_Local_CSV_Window()
            load_Local_CSV_Table_CSVDF()

        if event == "關於":
            sg.popup('股票資訊爬蟲\n版本： 1.0\n作者：Douggy Sans\n2021年編寫', title='關於')

    if window == setting_Window:  # 主視窗 -> 設定視窗之互動
        window.bring_to_front()
        if event in (sg.WIN_CLOSED, '取消'):
            window.close()
            setting_Window = None

        if event == 'mDBName':
            print(values['mDBName'])
            db = DBClient.get_database(str(values['mDBName']))
            codata_list = db.list_collection_names()
            print(codata_list)
            setting_Window['mCDName'].update(
                value=codata_list[0], values=codata_list)
        if event == '管理資料庫與資料集':
            connect_Mongo(False, True, True, True)
            window.close()
            setting_Window = None
            setting_Window = set_Setting_Window()
        if event == "保存":
            print("Save"+str(values['mDBUrI']),
                  str(values['mDBName']), str(values['mCDName']))
            if(check_Mongo()):
                conf.set('MongoDB', 'MONGO_URI', str(values['mDBUrI']))
                conf.set('MongoDB', 'DBNAME', str(values['mDBName']))
                conf.set('MongoDB', 'cdataname', str(values['mCDName']))
                conf.write(open(cfgpath, 'w'))
                check_Mongo()
                sg.popup('已保存設定！', title='已保存')
                scrapyer.change_Project_Setting(str(conf.get('MongoDB', 'MONGO_URI')), str(
                    conf.get('MongoDB', 'DBNAME')), str(conf.get('MongoDB', 'cdataname')))
                window.close()
                setting_Window = None
            else:
                setting_Window.bring_to_front()
                setting_Window.make_modal()

        if event == "重置":
            if(sg.popup_ok_cancel('是否重置設定？', title='確認重置', modal=True) == 'OK'):
                reset_setting()
                sg.popup('已重置！')
                window.close()
                setting_Window = None
        if event == "開啟設定目錄":
            os.startfile(profile_PATH)

    if window == Spider_Stock_Price_Window:  # 主視窗 -> 股價爬蟲 -> 選擇來源
        window.bring_to_front()
        if event in (sg.WIN_CLOSED, '取消'):
            window.close()
            Spider_Stock_Price_Window = None
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

    if window == csv_Row_Add_Window:  # 編輯本地股號表 -> 新增單筆資料
        if event == '保存':
            co_id = values['COID']
            co_name = values['CONAME']
            if((co_id or co_name != '') and (co_id != co_name)):
                local_CSV_Row_Add(co_id, co_name)
                csv_Row_Add_Window.close()
                csv_Row_Add_Window = None
                local_Csv_Window.make_modal()
            elif(co_id == co_name):
                sg.popup_error('股號與公司名稱重複！')
                window.make_modal()
            else:
                sg.popup_error('負號與公司欄位請勿留空！')
                window.make_modal()

        if event == '取消':
            csv_Row_Add_Window.close()
            csv_Row_Add_Window = None
            local_Csv_Window.make_modal()

    if window == local_Csv_Window:  # 主視窗 -> 編輯本地股號表
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
        if event in ('保存當前變更'):
            save_Local_CSV()
        if event in ('關閉且「不保存」變更'):
            update_Local_CSV_Table()
            window.close()
            local_Csv_Window = None
        if event in ('關閉且「保存」變更'):
            print(user_df)
            save_Local_CSV()
            backup_Coid_pd_df.clear()
            window.close()
            local_Csv_Window = None
        if event == "重新整理":
            if(local_Coid_CSV_is_changed):
                if(sg.popup_ok_cancel('你股號表尚未存檔，重新整理將會喪失變更的資料並還原修改前的樣子', title='重新整理', modal=True, no_titlebar=True) == 'OK'):
                    update_Local_CSV_Table()
            else:
                update_Local_CSV_Table()
            window.make_modal()
        if event == "匯入外部股號表":
            local_CSV_Import_usercsvfile()
        if event == "重置本地股號表":
            if(sg.popup_ok_cancel('是否重置股號表？', title='確認重置', modal=True) == 'OK'):
                local_CSV_Backup_USER_DF()
                reset_Local_CSV_Table()
                sg.popup('已重置！')
            window.make_modal()
        if event == "刪除":
            select_row = values['_local_Coid_CSV_Table']
            if select_row != []:
                select_row = int(select_row[0])
                local_CSV_Row_Edit(False, select_row)
            else:
                sg.popup('請先選擇有效的項目進行來編輯')
            window.make_modal()
        if event == "編輯" or event == '_local_Coid_CSV_Table':
            select_row = values['_local_Coid_CSV_Table']
            if select_row != []:
                select_row = int(select_row[0])
                local_CSV_Row_Edit(True, select_row)
            else:
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
window.close()
