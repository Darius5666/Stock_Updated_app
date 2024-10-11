import requests
import json
import pygsheets
import datetime
import time
# 授權

gc = pygsheets.authorize(service_file='/etc/secrets/Key.json')

# 打開 Google Sheet
sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1p524CpGfpYusHkXGM4WR__lzqEClDsuQk35EWyCaINE/edit?gid=0#gid=0')
wk = sht.worksheet_by_title("工作表1")

# 取得 API 數據
url = "https://openapi.taifex.com.tw/v1/MarketDataOfMajorInstitutionalTradersDetailsOfFuturesContractsBytheDate"
re = requests.get(url=url).json()
target_time = "09:00:00"
current_time = datetime.datetime.now().strftime("%H:%M:%S")
updated=0
while True:
# 迭代處理 API 數據

        while current_time != target_time:
            current_time = datetime.datetime.now().strftime("%H:%M:%S")
            time.sleep(1)
        for item in re:
            if item['ContractCode'] == '臺股期貨' and item['Item'] == '外資及陸資':
                day = item["Date"]
                col = wk.get_col(1, include_empty=False)
                
                # 取得之前的數據 (B列的最後一個值)
                
                pastB = wk.cell(f"B{len(col)}").value
                curB_value = abs(int(item['OpenInterest(Net)']))  # 目前B欄的數值
                curC=wk.cell(f"C{len(col)+1}").value
                curD=wk.cell(f"D{len(col)+1}").value
                # 計算增量
                if pastB!="未平倉餘額多空淨值口數":
                    delta = int(item['OpenInterest(Net)']) +int(pastB)
                else:
                    delta=0
                    pastB=1
                
                # 計算百分比變化
                percent_change = f"{round((delta / int(pastB)) * 100, 2)}%"

                # 更新 A, B, C, D 欄的值
                wk.update_value(f"A{len(col) + 1}", day[0:4] + "\\" + day[4:6] + "\\" + day[6:8])
                wk.update_value(f"B{len(col) + 1}", curB_value)

                wk.update_value(f"C{len(col) + 1}", str(delta))
                wk.update_value(f"D{len(col) + 1}", percent_change)   
                current_time = datetime.datetime.now().strftime("%H:%M:%S")
                break
