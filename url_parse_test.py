import urllib.parse as urlParse
from urllib.parse import parse_qs

url ='https://mops.twse.com.tw/server-java/t164sb01?step=1&CO_ID=9188&SYEAR=2019&SSEASON=3&REPORT_ID=C'
parsed = urlParse.urlparse(url)
company_Id=parse_qs(parsed.query)['CO_ID'] #獲取網址股號
report_ID=parse_qs(parsed.query)['REPORT_ID'] #獲取回報ID
print(f"CO_ID: {company_Id[0]},RID: {report_ID[0]}")