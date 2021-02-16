import csv

file='.\上市_urlA.csv'
print(file)
with open(file,newline='',encoding="utf-8") as csvfile_Lc:
    rows = csv.DictReader(csvfile_Lc)
    for row in rows:
        if(row['代號'].isnumeric() & (len(row['代號']) !=6)):
            print(row['代號'])

csvfile_Lc.close()