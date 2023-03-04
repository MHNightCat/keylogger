import keyboard
import json
import datetime
import os
import pandas as pd

# 設置 JSON 檔案路徑和檔案名稱
file_path = './'
today_str = datetime.datetime.now().strftime("%Y-%m-%d")
json_file = os.path.join(file_path, today_str + '.json')

# 如果昨天的檔案存在，就讀取檔案中的計數信息
yesterday_str = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
yesterday_json_file = os.path.join(file_path, yesterday_str + '.json')
if os.path.exists(yesterday_json_file):
    with open(yesterday_json_file, 'r') as f:
        yesterday_counts = json.load(f)
    # 轉換為DataFrame
    df = pd.DataFrame(yesterday_counts.items(), columns=["Key", yesterday_str])
    # 將DataFrame儲存為Excel檔案
    excel_file = os.path.join(file_path, yesterday_str + '.xlsx')
    df.to_excel(excel_file, index=False)
    print(f"{yesterday_json_file} converted to {excel_file}")

# 如果今天的檔案存在，就讀取檔案中的計數信息
if os.path.exists(json_file):
    with open(json_file, 'r') as f:
        counts = json.load(f)
else:
    counts = {}
print(f"Loaded counts from {json_file}")

# 定義計數器
total_count = counts.get("total_count", 0)

# 鍵盤事件回調函數
def on_event(event):
    global total_count
    key = event.name
    counts[key] = counts.get(key, 0) + 1
    total_count += 1
    counts["total_count"] = total_count
    with open(json_file, 'w') as f:
        json.dump(counts, f, sort_keys=True)

keyboard.on_press(on_event)
keyboard.on_release(on_event)

# 等待鍵盤事件
keyboard.wait()