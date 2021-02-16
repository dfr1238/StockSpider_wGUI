import csv

print('上櫃資料集\n')
with open('.\上櫃.csv',newline='') as csvfile_Lc:
    rows = csv.DictReader(csvfile_Lc)
    for row in rows:
        if(row['代號'].isnumeric()):
            print(row['代號'],row['名稱'])

csvfile_Lc.close()

print('上市資料集\n')
with open('.\上市.csv',newline='') as csvfile_Lc:
    rows = csv.DictReader(csvfile_Lc)
    for row in rows:
        if(row['代號'].isnumeric()):
            print(row['代號'],row['名稱'])

csvfile_Lc.close()