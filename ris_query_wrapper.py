"""
村里查詢包裝腳本
使用 python_embedded 的 Python 執行查詢，避免模組衝突
"""
import sys
import json
import os

# 🔥 將 _internal 目錄加入模組搜尋路徑
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# 🔥 找到並加入 python_embedded 的路徑
base_dir = os.path.dirname(script_dir)  # 從 _internal 往上一層
python_embedded_paths = [
    os.path.join(base_dir, 'python_embedded', 'Lib', 'site-packages'),
    os.path.join(base_dir, 'python_embedded', 'Lib'),
]
for path in python_embedded_paths:
    if os.path.exists(path) and path not in sys.path:
        sys.path.insert(0, path)

# 從命令列參數獲取查詢資訊
if len(sys.argv) < 5:
    print(json.dumps({"error": "參數不足"}))
    sys.exit(1)

city = sys.argv[1]
district = sys.argv[2]
road = sys.argv[3]
number = sys.argv[4]
number_part = sys.argv[5] if len(sys.argv) > 5 else number

# 🔥 強制 stdout 立即輸出（使用環境變量）
os.environ['PYTHONUNBUFFERED'] = '1'

# 定義一個立即輸出的 print 函數
def flush_print(msg):
    """立即輸出訊息（不使用緩衝）"""
    print(msg, flush=True)

# 導入查詢模組
from ris_query_selenium import RISQueryBot

# 執行查詢
try:
    bot = RISQueryBot(headless=False, message_callback=flush_print)
    result = bot.query_by_address(city, district, road, number, number_part)
    bot.close()

    if result:
        # 輸出 JSON 格式結果
        print("===RESULT_START===")
        print(json.dumps(result, ensure_ascii=False))
        print("===RESULT_END===")
    else:
        print(json.dumps({"error": "查詢失敗"}))

except Exception as e:
    print(json.dumps({"error": str(e)}))
    import traceback
    traceback.print_exc()
