import psutil
import subprocess
import threading

# 🔥 Windows 高 DPI 支援：必須在 import tkinter 之前設定
import ctypes
try:
    # 告訴 Windows 此程式支援高 DPI
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        # 舊版 Windows 的 fallback
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import font, scrolledtext, filedialog, messagebox, ttk
import tkinter.simpledialog as simpledialog
import re
from ctypes import wintypes
import os
import time
import json
from datetime import datetime

# 🔥 提前建立 root 並顯示 splash，讓使用者啟動瞬間就看到視窗（避免「黑屏 3 秒沒反應」感）
# 後續所有 widget / 函式定義照常建在這個 root 上
root = tk.Tk()
root.title("地籍資料查詢系統 - 載入中...")
# 簡單置中的 splash 視窗
try:
    _sw = root.winfo_screenwidth()
    _sh = root.winfo_screenheight()
    root.geometry(f"360x120+{(_sw - 360) // 2}+{(_sh - 120) // 2}")
except Exception:
    root.geometry("360x120+800+400")
_splash_frame = tk.Frame(root, bg='#1976D2')
_splash_frame.pack(fill=tk.BOTH, expand=True)
tk.Label(_splash_frame, text="📋 地籍資料查詢系統",
         font=("微軟正黑體", 14, "bold"), bg='#1976D2', fg='white').pack(pady=(20, 5))
tk.Label(_splash_frame, text="載入中，請稍候...",
         font=("微軟正黑體", 11), bg='#1976D2', fg='white').pack()
root.update()  # 強制立即顯示

# 版本資訊
VERSION = "1.0.9"
BUILD_DATE = "2026-05-18"

# 🔥 處理 PyInstaller 打包後的路徑問題
import sys

# 🔥 設定工作目錄為主程式執行檔所在目錄（不是 _internal）
if getattr(sys, 'frozen', False):
    # 打包環境：使用 .exe 所在目錄
    PROGRAM_DIR = os.path.dirname(sys.executable)
else:
    # 開發環境：使用 main.py 所在目錄
    PROGRAM_DIR = os.path.dirname(os.path.abspath(__file__))

# 切換工作目錄到主程式目錄
os.chdir(PROGRAM_DIR)
print(f"[DEBUG] 工作目錄設定為: {PROGRAM_DIR}")
print(f"[DEBUG] 當前工作目錄: {os.getcwd()}")

# 🔥 定義路徑輔助函數（主程式版本）
def get_data_json_path():
    """取得 data.json 的完整路徑"""
    return os.path.join(PROGRAM_DIR, 'data.json')

def get_work_folder(folder_name):
    """取得工作資料夾的完整路徑"""
    return os.path.join(PROGRAM_DIR, folder_name)

def get_resource_path(relative_path):
    """
    獲取資源的絕對路徑，適用於開發環境和 PyInstaller 打包後的環境
    """
    try:
        # PyInstaller 會創建臨時資料夾，並將路徑存儲在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        # 如果不是打包環境，使用當前工作目錄
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 🔥 嘗試修復 tkinter 輸入問題
if sys.platform == 'win32':
    # 設置環境變數，可能修復 IME 問題
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    # 嘗試禁用 IME
    try:
        import locale
        locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
    except:
        pass

# 🔥 docx 載入很慢（1-2 秒），main.py 不直接用，移到 data_editor.py 內 lazy 載入
# 保留可用性檢查供其他模組查詢
try:
    import importlib.util
    DOCX_AVAILABLE = importlib.util.find_spec("docx") is not None
except Exception:
    DOCX_AVAILABLE = False
if not DOCX_AVAILABLE:
    print("警告：未安裝 python-docx，無法生成 Word 報告")

# 導入配置管理模組
from config_manager import open_config_manager, load_config, save_config

# 導入資料編輯器模組
import data_editor

# 設定 data_editor 模組需要的全域變數和函數
def setup_data_editor_module():
    """設定 data_editor 模組需要的全域變數"""
    data_editor.load_final_data = load_final_data
    data_editor.save_final_data = save_final_data
    data_editor.load_contacts_from_excel = load_contacts_from_excel
    data_editor.load_urban_planning_data = load_urban_planning_data
    data_editor.get_city_from_district = get_city_from_district
    data_editor.extract_district_from_address = extract_district_from_address
    data_editor.calculate_total_land_area = calculate_total_land_area
    data_editor.convert_transcript_to_final_data = convert_transcript_to_final_data
    data_editor.update_message = update_message
    data_editor.show_large_message = show_large_message
    data_editor.show_large_yesno = show_large_yesno
    data_editor.show_large_yesnocancel = show_large_yesnocancel
    data_editor.root = root
    data_editor.choose_json_file = choose_json_file  # 🔥 加入 choose_json_file 函數
    # 傳遞函數而非路徑，讓 data_editor 可以動態取得最新路徑
    data_editor.get_data_final_path = get_data_final_path
    data_editor.get_selected_json_path = get_selected_json_path
    data_editor.get_screenshot_path = get_screenshot_path
    # report_type 會在執行時動態設定，所以初始化為 None
    data_editor.report_type = None


# === JSON 讀取設定（預設 output，可被使用者覆蓋） ===
# 🔥 修正：使用絕對路徑
if getattr(sys, 'frozen', False):
    # 打包環境
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # 開發環境
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

JSON_DIR_DEFAULT = os.path.join(BASE_DIR, "output")
selected_json_path = None
json_dir = JSON_DIR_DEFAULT

# 🔥 從設定檔讀取上次選擇的 JSON 路徑
try:
    config = load_config()
    if config and 'selected_json' in config:
        saved_path = config['selected_json']
        # 🔥 修正：如果路徑不存在，嘗試在 output 目錄中尋找同名檔案
        if saved_path:
            if os.path.exists(saved_path):
                selected_json_path = saved_path
                print(f"[啟動] 已載入上次選擇的JSON檔案：{os.path.basename(saved_path)}")
            else:
                # 嘗試在當前 output 目錄中尋找同名檔案
                filename = os.path.basename(saved_path)
                potential_path = os.path.join(JSON_DIR_DEFAULT, filename)
                if os.path.exists(potential_path):
                    selected_json_path = potential_path
                    print(f"[啟動] 在output目錄找到JSON檔案：{filename}")
                else:
                    print(f"[警告] 上次選擇的JSON檔案不存在：{filename}")
except Exception as e:
    print(f"[警告] 讀取設定檔失敗：{e}")

# === 資料累積相關 ===
def get_data_final_path():
    """
    取得 data_final.json 的完整路徑
    位置：建立的資料夾/4.其他相關/data_final.json
    如果 data.json 不存在，則回傳程式目錄下的路徑
    """
    try:
        # 🔥 使用 get_data_json_path() 確保讀取正確位置的 data.json
        data_json_path = get_data_json_path()
        if os.path.exists(data_json_path):
            with open(data_json_path, 'r', encoding='utf-8') as f:
                data_list = json.load(f)
                if data_list:
                    first_data = data_list[0]
                    area = first_data.get('area', '')
                    section = first_data.get('section', '')
                    lot_number = first_data.get('lot_number', '')
                    folder_name = f"{area}{section}-{lot_number}"

                    # 🔥 使用 get_work_folder() 確保在主程式目錄下建立資料夾
                    other_dir = get_work_folder(os.path.join(folder_name, "4.其他相關"))
                    if not os.path.exists(other_dir):
                        os.makedirs(other_dir)

                    return os.path.join(other_dir, "data_final.json")
    except Exception as e:
        print(f"取得 data_final.json 路徑時發生錯誤：{e}")

    # 🔥 如果無法取得資料夾，回傳程式目錄下的路徑（向後相容）
    return os.path.join(PROGRAM_DIR, "data_final.json")

DATA_FINAL_PATH = get_data_final_path()

def get_selected_json_path():
    """取得當前已選擇的 JSON 檔案路徑"""
    return selected_json_path

def get_data_json_path():
    """
    取得 data.json 的完整路徑
    優先順序：
    1. 建立的資料夾/4.其他相關/data.json（如果資料夾存在）
    2. 程式目錄下的 data.json（向後相容或首次建立）
    """
    # 🔥 使用絕對路徑，確保在任何工作目錄下都能正確找到
    base_data_json = os.path.join(PROGRAM_DIR, 'data.json')

    try:
        # 先檢查程式目錄下是否有 data.json（用於取得資料夾名稱）
        if os.path.exists(base_data_json):
            with open(base_data_json, 'r', encoding='utf-8') as f:
                data_list = json.load(f)
                if data_list:
                    first_data = data_list[0]
                    area = first_data.get('area', '')
                    section = first_data.get('section', '')
                    lot_number = first_data.get('lot_number', '')

                    if area and section and lot_number:
                        folder_name = f"{area}{section}-{lot_number}"
                        # 🔥 使用 PROGRAM_DIR 為基礎的絕對路徑
                        other_dir = os.path.join(PROGRAM_DIR, folder_name, "4.其他相關")
                        target_path = os.path.join(other_dir, "data.json")

                        # 如果目標資料夾存在且有 data.json，使用該檔案
                        if os.path.exists(target_path):
                            return target_path
    except Exception as e:
        print(f"取得 data.json 路徑時發生錯誤：{e}")

    # 預設回傳程式目錄下的路徑（絕對路徑）
    return base_data_json

def save_data_json(data_list):
    """
    儲存 data.json 到正確位置
    會同時儲存到：
    1. 程式目錄下（向後相容）
    2. 建立的資料夾/4.其他相關/（如果資料夾存在）
    """
    # 🔥 使用絕對路徑
    base_data_json = os.path.join(PROGRAM_DIR, 'data.json')

    try:
        # 儲存到程式目錄（使用絕對路徑）
        with open(base_data_json, 'w', encoding='utf-8') as f:
            json.dump(data_list, f, ensure_ascii=False, indent=2)

        # 如果有資料，嘗試同時儲存到建立的資料夾
        if data_list:
            first_data = data_list[0]
            area = first_data.get('area', '')
            section = first_data.get('section', '')
            lot_number = first_data.get('lot_number', '')

            if area and section and lot_number:
                folder_name = f"{area}{section}-{lot_number}"
                # 🔥 使用 PROGRAM_DIR 為基礎的絕對路徑
                other_dir = os.path.join(PROGRAM_DIR, folder_name, "4.其他相關")

                # 確保資料夾存在
                if not os.path.exists(other_dir):
                    os.makedirs(other_dir)

                target_path = os.path.join(other_dir, "data.json")
                with open(target_path, 'w', encoding='utf-8') as f:
                    json.dump(data_list, f, ensure_ascii=False, indent=2)
                print(f"[OK] 已同步儲存到 {target_path}")

        return True
    except Exception as e:
        print(f"儲存 data.json 時發生錯誤：{e}")
        return False

# === 通訊錄載入函數 ===
def load_contacts_from_excel(file_path='通訊錄.xlsx'):
    """
    從通訊錄 Excel 檔案中載入聯絡人資料，依店別分組

    回傳:
        dict: {
            'stores': ['總部', '至聖店', ...],  # 店別列表
            'contacts_by_store': {
                '總部': [{'name': '姓名', 'phone': '電話'}, ...],
                '至聖店': [...]
            }
        }
    """
    result = {
        'stores': [],
        'contacts_by_store': {}
    }

    try:
        import pandas as pd
        from base_dir_helper import get_base_dir
        import os

        # 🔥 使用 base_dir 來定位通訊錄檔案
        if not os.path.isabs(file_path):
            base_dir = get_base_dir()
            file_path = os.path.join(base_dir, file_path)

        # 讀取 Excel
        df = pd.read_excel(file_path, sheet_name=0)

        current_store = None  # 當前店別

        # 遍歷所有行
        for idx, row in df.iterrows():
            # 取得各欄位值
            store_col = row.get('店別', '')  # 第1欄：店別
            title_col = row.get('Unnamed: 1', '')  # 第2欄：職稱
            name = row.get('Unnamed: 2', '')  # 第3欄：姓名
            phone = row.get('Unnamed: 3', '')  # 第4欄：電話
            remark = row.get('Unnamed: 4', '')  # 第5欄：備註

            # 檢查是否為店別標題行
            if pd.notna(store_col) and isinstance(store_col, str) and store_col.strip():
                store_name = store_col.strip()
                # 🔥 過濾掉「店別」標題行和「總公司管理部」
                if store_name != '店別' and store_name != '總公司管理部':
                    current_store = store_name
                    if current_store not in result['stores']:
                        result['stores'].append(current_store)
                        result['contacts_by_store'][current_store] = []
                elif store_name == '總公司管理部':
                    # 設為 None，後續聯絡人不會被加入
                    current_store = None

            # 檢查是否為有效聯絡人
            if name and phone and isinstance(name, str) and isinstance(phone, str):
                # 排除標題行和特殊標記
                if ('＜' not in name and '＞' not in name and
                    '姓名' not in name and '店別' not in name and
                    '：' not in name):  # 排除「經理：唐凱華」這類格式

                    # 🔥 檢查備註欄位是否有任何文字（離職、成仁、離退等）
                    has_remark = pd.notna(remark) and isinstance(remark, str) and remark.strip()

                    # 🔥 也檢查職稱欄是否包含備註
                    title_has_remark = False
                    if pd.notna(title_col) and isinstance(title_col, str):
                        if '離職' in title_col or '成仁' in title_col or '離退' in title_col:
                            title_has_remark = True

                    # 🔥 過濾特定職稱（行政、會計、業助）
                    title_str = str(title_col).strip() if pd.notna(title_col) else ''
                    is_filtered_title = title_str in ['行政', '會計', '業助']

                    # 🔥 總部只保留莊明昇
                    is_allowed = True
                    if current_store == '總部':
                        clean_name_temp = name.strip()
                        if clean_name_temp != '莊明昇':
                            is_allowed = False

                    if not has_remark and not title_has_remark and not is_filtered_title and is_allowed:
                        # 清理姓名和電話
                        clean_name = name.strip()
                        clean_phone = str(phone).strip()

                        # 確保電話號碼有效（以0開頭，長度合理）
                        if clean_phone.startswith('0') and len(clean_phone) >= 9:
                            if current_store:
                                result['contacts_by_store'][current_store].append({
                                    'name': clean_name,
                                    'phone': clean_phone
                                })

        # 統計總人數
        total = sum(len(contacts) for contacts in result['contacts_by_store'].values())
        print(f"[OK] 成功載入 {len(result['stores'])} 個店別，共 {total} 位聯絡人")
        return result

    except FileNotFoundError:
        print(f"[警告] 找不到通訊錄檔案：{file_path}")
        return result
    except Exception as e:
        print(f"[錯誤] 讀取通訊錄失敗：{e}")
        import traceback
        traceback.print_exc()
        return result

# === 取得截圖檔案路徑（根據 data.json 第一筆） ===
def get_screenshot_path():
    """
    根據 data.json 第一筆資料，取得截圖檔案應該儲存的路徑

    回傳:
        str: 截圖檔案完整路徑（例如：大寮區義堂段-358-21/4.其他相關/gui_map_screenshot.png）
        如果無法取得資料夾資訊，則回傳 None
    """
    try:
        # 從 data.json 取得第一筆資料，建立資料夾路徑
        if os.path.exists('data.json'):
            with open('data.json', 'r', encoding='utf-8') as f:
                data_list = json.load(f)
            if data_list and len(data_list) > 0:
                first_data = data_list[0]
                area = first_data.get('area', '')
                section = first_data.get('section', '')
                lot_number = first_data.get('lot_number', '')

                if area and section and lot_number:
                    base_folder = f"{area}{section}-{lot_number}"
                    screenshot_path = os.path.join(base_folder, "4.其他相關", "gui_map_screenshot.png")

                    # 確保 4.其他相關 資料夾存在
                    other_folder = os.path.join(base_folder, "4.其他相關")
                    os.makedirs(other_folder, exist_ok=True)

                    return screenshot_path

        return None
    except Exception as e:
        print(f"[錯誤] 取得截圖路徑失敗：{e}")
        return None

# === 讀取都市計畫分區 JSON 資料 ===
def load_urban_planning_data(target_lots=None):
    """
    從 data.json 第一筆對應的資料夾下的「4.其他相關」讀取都市計畫分區 JSON 資料

    功能：
    - 根據 data.json 第一筆建立資料夾路徑（例如：大寮區義堂段-358-21）
    - 搜尋該資料夾下「4.其他相關」的所有「分區查詢資料-*.json」和「都市計畫圖街廓資訊-*.json」檔案
    - 🔥 只讀取符合 target_lots 指定地段地號的檔案
    - 讀取每個檔案的「使用分區」、「建蔽率」、「容積率」欄位
    - 去重後整理成清單（例如：使用分區：農業區、工業區；建蔽率：60、70；容積率：180、240）

    參數:
        target_lots: list of dict，例如 [{'地段': '大寮區義堂段', '地號': '358-21'}, ...]
                     如果為 None，則讀取所有檔案

    回傳:
        dict: {
            'zones': ['農業區', '工業區'],  # 去重後的使用分區列表
            'building_coverage_ratios': ['60', '70'],  # 去重後的建蔽率列表
            'floor_area_ratios': ['180', '240'],  # 去重後的容積率列表
            'details': [  # 每個 JSON 的詳細資料
                {
                    'file': '分區查詢資料-大寮區義堂段-358-21.json',
                    '地段': '大寮區義堂段',
                    '地號': '358-21',
                    '使用分區': '農業區',
                    '建蔽率': '60',
                    '容積率': '180',
                    '土地面積': '1234.56',
                    '公告現值': '12345',
                    '公告地價': '6789',
                    '附註': ''
                },
                ...
            ]
        }
        如果找不到檔案則回傳 None
    """
    try:
        import re
        import glob

        # 🔥 搜尋當前目錄下所有包含「4.其他相關\都市計畫資料彙整.json」的資料夾
        json_path = None

        # 方法1：從 data.json 取得第一筆資料，建立資料夾路徑
        if os.path.exists('data.json'):
            with open('data.json', 'r', encoding='utf-8') as f:
                data_list = json.load(f)
            if data_list and len(data_list) > 0:
                first_data = data_list[0]
                area = first_data.get('area', '')
                section = first_data.get('section', '')
                lot_number = first_data.get('lot_number', '')

                if area and section and lot_number:
                    base_folder = f"{area}{section}-{lot_number}"
                    candidate_path = os.path.join(base_folder, "4.其他相關", "都市計畫資料彙整.json")

                    if os.path.exists(candidate_path):
                        json_path = candidate_path
                        print(f"[OK] 找到都市計畫資料彙整檔案（方法1）：{json_path}")

        # 🔧 方法2：只在沒有指定 target_lots 時才搜尋其他資料夾
        # 避免讀取到其他案件的都市計畫資料
        if not json_path and not target_lots:
            print("[提示] 方法1失敗，搜尋當前目錄下所有符合的檔案...")
            pattern = os.path.join("*", "4.其他相關", "都市計畫資料彙整.json")
            matching_files = glob.glob(pattern)

            if matching_files:
                # 使用第一個找到的檔案
                json_path = matching_files[0]
                print(f"[OK] 找到都市計畫資料彙整檔案（方法2）：{json_path}")
                if len(matching_files) > 1:
                    print(f"[提示] 找到多個檔案，使用第一個：{json_path}")

        if not json_path:
            print(f"[提示] 找不到都市計畫資料彙整檔案")
            return None

        # 🔥 讀取統一的 JSON 檔案
        with open(json_path, 'r', encoding='utf-8') as f:
            combined_data = json.load(f)

        # 取得兩類資料
        zone_data_list = combined_data.get('分區查詢資料', [])
        urban_plan_list = combined_data.get('都市計畫圖街廓資訊', [])

        print(f"[OK] 讀取到 {len(zone_data_list)} 筆分區查詢資料")
        print(f"[OK] 讀取到 {len(urban_plan_list)} 筆都市計畫圖街廓資訊")

        # 🔥 建立目標地段地號的集合，用於快速比對
        target_set = None
        if target_lots:
            target_set = set()
            for lot in target_lots:
                section = lot.get('地段', '')
                lot_num = lot.get('地號', '')
                if section and lot_num:
                    target_set.add(f"{section}{lot_num}")
            print(f"[檢查] 目標地段地號：{target_set}")

        # 處理所有資料
        all_details = []
        all_zones = set()  # 使用集合自動去重
        all_building_coverage = set()  # 建蔽率去重集合
        all_floor_area = set()  # 容積率去重集合

        def extract_numeric_value(value_str):
            """從字串中提取數字部分（例如：'60%' -> '60'，'60' -> '60'）"""
            if not value_str:
                return ''
            # 移除百分號、空白等，只保留數字和小數點
            cleaned = re.sub(r'[^\d.]', '', str(value_str))
            return cleaned if cleaned else ''

        def normalize_lot_number(lot_num):
            """
            標準化地號格式，移除前導零
            例如：'0193-0000' -> '193-0000'，'0193' -> '193'，'193' -> '193'
            """
            if not lot_num:
                return ''
            # 處理母號-子號格式
            if '-' in lot_num:
                parts = lot_num.split('-')
                # 移除母號的前導零
                mother = parts[0].lstrip('0') or '0'
                # 移除子號的前導零
                child = parts[1].lstrip('0') or '0'
                # 如果子號為0，只返回母號
                if child == '0':
                    return mother
                return f"{mother}-{child}"
            else:
                # 只有母號，移除前導零
                return lot_num.lstrip('0') or '0'

        # 處理所有資料（分區查詢資料 + 都市計畫圖街廓資訊）
        all_data = zone_data_list + urban_plan_list

        # 🔥 建立標準化的目標地段地號集合
        normalized_target_set = None
        if target_lots:
            normalized_target_set = set()
            for lot in target_lots:
                section = lot.get('地段', '')
                lot_num = lot.get('地號', '')
                if section and lot_num:
                    normalized_lot = normalize_lot_number(lot_num)
                    normalized_target_set.add(f"{section}{normalized_lot}")
            print(f"[檢查] 標準化後的目標地段地號：{normalized_target_set}")

        # 🔥 第一輪：嘗試比對地段地號（使用標準化後的地號）
        matched_data = []
        if normalized_target_set:
            for data_dict in all_data:
                json_section = data_dict.get('地段', '').strip()
                json_lot = data_dict.get('地號', '').strip()

                if not json_section and not json_lot:
                    continue

                # 🔥 標準化地號後比對
                normalized_json_lot = normalize_lot_number(json_lot)
                json_key = f"{json_section}{normalized_json_lot}"

                if json_key in normalized_target_set:
                    matched_data.append(data_dict)
                    print(f"     [✓ 匹配] 地段='{json_section}' 地號='{json_lot}' (標準化: {normalized_json_lot})")

        # 🔥 如果沒有匹配的資料
        if not matched_data:
            if target_set:
                # 🔧 有指定目標地段地號但沒有匹配，表示這筆土地沒有都市計畫資料
                # 不應該載入其他案件的資料
                print(f"[提示] 沒有找到匹配的地段地號，此土地可能非都市計畫土地")
                return None
            else:
                # 沒有提供目標地段地號，載入所有資料
                print(f"[提示] 未提供目標地段地號，載入所有資料")
                matched_data = all_data

        # 處理匹配的資料
        for data_dict in matched_data:
            try:
                # 過濾空資料
                json_section = data_dict.get('地段', '').strip()
                json_lot = data_dict.get('地號', '').strip()

                if not json_section and not json_lot:
                    print(f"     [跳過] 空資料")
                    continue

                print(f"     [✓ 載入] 地段='{json_section}' 地號='{json_lot}'")

                # 🔥 收集使用分區（修正欄位名稱：使用分區整編名稱 → 使用分區完整名稱）
                zone = data_dict.get('使用分區', '') or data_dict.get('使用分區完整名稱', '') or data_dict.get('使用分區別名', '')
                if zone:
                    all_zones.add(zone)
                    print(f"        使用分區={zone}")

                # 收集建蔽率
                building_coverage = data_dict.get('建蔽率', '')
                if building_coverage:
                    cleaned_value = extract_numeric_value(building_coverage)
                    if cleaned_value:
                        all_building_coverage.add(cleaned_value)
                        print(f"        建蔽率={cleaned_value}")

                # 收集容積率
                floor_area = data_dict.get('容積率', '')
                if floor_area:
                    cleaned_value = extract_numeric_value(floor_area)
                    if cleaned_value:
                        all_floor_area.add(cleaned_value)
                        print(f"        容積率={cleaned_value}")

                all_details.append(data_dict)

            except Exception as e:
                print(f"[警告] 處理資料時失敗：{e}")
                continue

        if not all_details:
            print("[提示] 無法讀取任何符合條件的都市計畫資料")
            return None

        # 整理結果
        result = {
            'zones': sorted(list(all_zones)),  # 排序後的去重列表
            'building_coverage_ratios': sorted(list(all_building_coverage), key=lambda x: float(x) if x else 0),  # 按數值排序
            'floor_area_ratios': sorted(list(all_floor_area), key=lambda x: float(x) if x else 0),  # 按數值排序
            'details': all_details
        }

        print(f"[OK] 整理完成（共 {len(all_details)} 筆匹配資料）")
        print(f"     使用分區 ({len(result['zones'])} 種): {', '.join(result['zones'])}")
        print(f"     建蔽率 ({len(result['building_coverage_ratios'])} 種): {', '.join(result['building_coverage_ratios'])}")
        print(f"     容積率 ({len(result['floor_area_ratios'])} 種): {', '.join(result['floor_area_ratios'])}")

        return result

    except Exception as e:
        print(f"[錯誤] 讀取都市計畫分區資料失敗：{e}")
        import traceback
        traceback.print_exc()
        return None

# === 台灣行政區對照表（縣市 → 行政區列表） ===
CITY_DISTRICTS = {
    "台南市": [
        "中西區", "東區", "南區", "北區", "安平區", "安南區",
        "永康區", "歸仁區", "新化區", "左鎮區", "玉井區", "楠西區",
        "南化區", "仁德區", "關廟區", "龍崎區", "官田區", "麻豆區",
        "佳里區", "西港區", "七股區", "將軍區", "學甲區", "北門區",
        "新營區", "後壁區", "白河區", "東山區", "六甲區", "下營區",
        "柳營區", "鹽水區", "善化區", "大內區", "山上區", "新市區", "安定區"
    ],
    "屏東縣": [
        "屏東市", "潮州鎮", "東港鎮", "恆春鎮", "萬丹鄉", "長治鄉",
        "麟洛鄉", "九如鄉", "里港鄉", "鹽埔鄉", "高樹鄉", "萬巒鄉",
        "內埔鄉", "竹田鄉", "新埤鄉", "枋寮鄉", "新園鄉", "崁頂鄉",
        "林邊鄉", "南州鄉", "佳冬鄉", "琉球鄉", "車城鄉", "滿州鄉",
        "枋山鄉", "三地門鄉", "霧台鄉", "瑪家鄉", "泰武鄉", "來義鄉",
        "春日鄉", "獅子鄉", "牡丹鄉"
    ],
    "高雄市": [
        "楠梓區", "左營區", "鼓山區", "三民區", "鹽埕區", "前金區",
        "新興區", "苓雅區", "前鎮區", "旗津區", "小港區", "鳳山區",
        "林園區", "大寮區", "大樹區", "大社區", "仁武區", "鳥松區",
        "岡山區", "橋頭區", "燕巢區", "田寮區", "阿蓮區", "路竹區",
        "湖內區", "茄萣區", "永安區", "彌陀區", "梓官區", "旗山區",
        "美濃區", "六龜區", "甲仙區", "杉林區", "內門區", "茂林區",
        "桃源區", "那瑪夏區"
    ]
}

def get_city_from_district(district_name):
    """
    根據行政區名稱查詢所屬縣市

    參數:
        district_name: 行政區名稱（如「鳥松區」、「屏東市」）

    回傳:
        (縣市名稱, 是否為重複區名) 的 tuple
        如果找不到則回傳 (None, False)
    """
    if not district_name:
        return None, False

    found_cities = []
    for city, districts in CITY_DISTRICTS.items():
        if district_name in districts:
            found_cities.append(city)

    if not found_cities:
        return None, False
    elif len(found_cities) == 1:
        return found_cities[0], False  # 唯一配對
    else:
        # 有重複區名（如：中正區、東區等）
        return found_cities[0], True  # 回傳第一個配對，並標記為重複

def extract_district_from_address(address_string):
    """
    從地址字串中提取行政區

    參數:
        address_string: 包含地址的字串

    回傳:
        行政區名稱（含「區」、「市」、「鎮」、「鄉」）
    """
    if not address_string:
        return None

    # 正規化形似字，避免 Unicode 差異造成 regex 無法匹配
    # 例：謄本中的 巿(U+5DFF) 與 市(U+5E02) 外觀相同但 code point 不同
    normalized = (address_string
                  .replace(chr(0x5dff), chr(0x5e02))   # 巿 → 市
                  .replace(chr(0x81fa), chr(0x53f0))   # 臺 → 台
                  )

    # 正則：尋找「XXX區」、「XXX市」、「XXX鎮」、「XXX鄉」
    pattern = r'([^\s]{1,4}(?:區|市|鎮|鄉))'
    match = re.search(pattern, normalized)

    return match.group(1) if match else None

def calculate_total_land_area(json_data):
    """
    計算總地坪（坪）
    
    參數:
        json_data: 從 JSON 讀取的資料（包含標示部和所有權部）
    
    回傳:
        總地坪（坪，字串格式，保留2位小數）
    """
    try:
        import re
        
        # 取得標示部面積（平方公尺）
        標示部 = json_data.get("標示部", {})
        
        # 標示部可能是 dict 或 list
        if isinstance(標示部, list):
            if not 標示部:
                return ""
            標示部 = 標示部[0] if isinstance(標示部[0], dict) else {}
        elif not isinstance(標示部, dict):
            標示部 = {}
        
        面積_str = 標示部.get("面積", "")
        
        if not 面積_str:
            return ""
        
        # 移除單位，提取數字
        面積_str = str(面積_str).replace("平方公尺", "").replace(",", "").strip()
        
        try:
            標示部面積 = float(面積_str)
        except:
            return ""
        
        # 取得所有權部
        所有權部 = json_data.get("所有權部", [])
        
        if not isinstance(所有權部, list):
            所有權部 = [所有權部] if 所有權部 else []
        
        if not 所有權部:
            # 🔥 如果沒有所有權部，用標示部面積 × 0.3025
            總地坪_坪 = 標示部面積 * 0.3025
            return f"{總地坪_坪:.2f}"
        
        # 計算總面積（平方公尺）
        總面積_平方公尺 = 0
        
        for 所有權人 in 所有權部:
            if not isinstance(所有權人, dict):
                continue
            
            權利範圍 = 所有權人.get("權利範圍", "")
            
            if not 權利範圍:
                continue
            
            # 解析權利範圍
            # 可能的格式：
            # "全部 1分之1"
            # "4分之1"
            # "10000分之1234"
            
            分子 = 1
            分母 = 1
            
            # 先移除"全部"等文字
            權利範圍 = 權利範圍.replace("全部", "").strip()
            
            # 匹配 "X分之Y" 格式
            match = re.search(r'(\d+)\s*分之\s*(\d+)', 權利範圍)
            
            if match:
                分母 = int(match.group(1))
                分子 = int(match.group(2))
            else:
                # 如果沒有匹配到，假設是全部
                分子 = 1
                分母 = 1
            
            # 計算該所有權人的面積
            if 分母 > 0:
                該筆面積 = 標示部面積 * (分子 / 分母)
                總面積_平方公尺 += 該筆面積
                
                print(f"  所有權人：{所有權人.get('所有權人', '未知')}")
                print(f"    權利範圍：{權利範圍} ({分子}/{分母})")
                print(f"    該筆面積：{該筆面積:.2f} 平方公尺")
        
        # 🔥 換算為坪（平方公尺 × 0.3025）
        總地坪_坪 = 總面積_平方公尺 * 0.3025
        
        print(f"\n總面積：{總面積_平方公尺:.2f} 平方公尺")
        print(f"總地坪：{總地坪_坪:.2f} 坪")
        
        return f"{總地坪_坪:.2f}"
        
    except Exception as e:
        print(f"計算總地坪時發生錯誤：{e}")
        import traceback
        traceback.print_exc()
        return ""

def load_final_data():
    """載入最終資料"""
    data_final_path = get_data_final_path()
    if os.path.exists(data_final_path):
        try:
            with open(data_final_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_final_data(data):
    """儲存最終資料"""
    try:
        data_final_path = get_data_final_path()
        with open(data_final_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        update_message(f"[完成] 已更新 {data_final_path}")
    except Exception as e:
        update_message(f"[錯誤] 儲存資料失敗: {e}")

def check_data_consistency(transcript_file=None):
    """
    檢查 data.json 和 transcript_data 的資料一致性

    參數：
        transcript_file: 指定要檢查的 transcript JSON 文件路徑（可選）
                        如果不指定，則使用最新的檔案

    返回：
        (is_consistent, warning_message, data_json_info, transcript_info, is_order_issue)
        is_order_issue: True=順序問題, False=完全不同地段, None=無問題或無法判斷
    """
    try:
        data_json_info = None
        transcript_info = None

        # 1. 檢查 data.json
        data_json_path = get_data_json_path()
        if os.path.exists(data_json_path):
            with open(data_json_path, 'r', encoding='utf-8') as f:
                data_list = json.load(f)
            if data_list and len(data_list) > 0:
                first_data = data_list[0]
                data_json_info = {
                    'area': first_data.get('area', ''),
                    'section': first_data.get('section', ''),
                    'lot_number': first_data.get('lot_number', ''),
                    'count': len(data_list)
                }

        # 2. 檢查 transcript_data JSON
        import glob

        # 🔧 如果指定了文件，使用指定的文件；否則使用最新的
        if transcript_file and os.path.exists(transcript_file):
            latest_transcript = transcript_file
        else:
            # 🔥 修正：使用絕對路徑搜尋（支援打包環境）
            output_dir = os.path.join(BASE_DIR, "output")
            transcript_files = glob.glob(os.path.join(output_dir, "trans_*.json")) + glob.glob(os.path.join(output_dir, "transcript_data_*.json"))
            if transcript_files:
                # 取最新的檔案
                latest_transcript = max(transcript_files, key=os.path.getmtime)
            else:
                latest_transcript = None

        if latest_transcript:

            with open(latest_transcript, 'r', encoding='utf-8') as f:
                transcript_data = json.load(f)

            if transcript_data and len(transcript_data) > 0:
                # 找第一筆土地謄本
                first_land = None
                for item in transcript_data:
                    if "土地謄本" in item.get("謄本類型", ""):
                        first_land = item
                        break

                if first_land:
                    basic_info = first_land.get("基本資訊", {})
                    land_number = basic_info.get("地號建號", "")

                    # 解析地號（例如：湖內區大湖段 2589-0008地號 / 屏東市花田段 1196-0000地號）
                    # 🔥 不只「區」，也要支援「鄉」「鎮」「市」結尾的鄉鎮市區
                    import re
                    match = re.search(r'([^\s]+(?:區|鄉|鎮|市))([^\s]+段)\s*(\d+-\d+)', land_number)
                    if match:
                        transcript_info = {
                            'area': match.group(1),
                            'section': match.group(2),
                            'lot_number': match.group(3),
                            'count': len([t for t in transcript_data if "土地謄本" in t.get("謄本類型", "")]),
                            'file': os.path.basename(latest_transcript),
                            'file_path': latest_transcript
                        }

        # 3. 比對結果
        if data_json_info and transcript_info:
            # 正規化地號格式（去除前導零，並處理簡化格式）
            def normalize_lot(lot):
                """
                正規化地號格式
                - "193" → "193-0"
                - "0193-0000" → "193-0"
                - "358-21" → "358-21"
                🔥 遇到非標準格式（如 land.py 合併查詢產生的 "1173,1174,1194"）
                   不再 raise ValueError，改成回傳原字串，讓比對自然不相等而不影響其他筆。
                """
                if not lot:
                    return ''
                try:
                    if '-' in lot:
                        parts = lot.split('-')
                        return f"{int(parts[0])}-{int(parts[1])}"
                    else:
                        return f"{int(lot)}-0"
                except (ValueError, IndexError):
                    return str(lot)

            data_lot = normalize_lot(data_json_info['lot_number'])
            transcript_lot = normalize_lot(transcript_info['lot_number'])

            if (data_json_info['area'] == transcript_info['area'] and
                data_json_info['section'] == transcript_info['section'] and
                data_lot == transcript_lot):
                return (True, None, data_json_info, transcript_info, None)
            else:
                # 不一致！先判斷是「順序問題」還是「完全不同地段」
                is_order_issue = False

                # 檢查 transcript_info 的地號是否存在於 data.json 中
                if os.path.exists('data.json'):
                    with open('data.json', 'r', encoding='utf-8') as f:
                        data_list = json.load(f)

                    for item in data_list:
                        item_lot = normalize_lot(item.get('lot_number', ''))
                        if (item.get('area') == transcript_info['area'] and
                            item.get('section') == transcript_info['section'] and
                            item_lot == transcript_lot):
                            is_order_issue = True
                            break

                if is_order_issue:
                    # 這是順序問題
                    warning = f"""
【資料不一致警告】

發現 data.json 和謄本結構化資料的第一筆地號不同：

📄 data.json（第一筆）：
   {data_json_info['area']}{data_json_info['section']} {data_json_info['lot_number']}
   （共 {data_json_info['count']} 筆）

📄 {transcript_info['file']}（第一筆土地謄本）：
   {transcript_info['area']}{transcript_info['section']} {transcript_info['lot_number']}
   （共 {transcript_info['count']} 筆土地謄本）

⚠️ 可能的影響：
   • 物件編輯器會根據 data.json 第一筆尋找稅額檔案
   • 後續流程依賴順序對齊時可能找不到對應記錄

💡 建議：依 data.json 順序整批重排謄本（避免後續處理混亂）
"""
                else:
                    # 完全不同地段
                    warning = f"""
【資料不一致警告】

發現 data.json 和謄本結構化資料完全不同地段：

📄 data.json（第一筆）：
   {data_json_info['area']}{data_json_info['section']} {data_json_info['lot_number']}
   （共 {data_json_info['count']} 筆）

📄 {transcript_info['file']}（第一筆土地謄本）：
   {transcript_info['area']}{transcript_info['section']} {transcript_info['lot_number']}
   （共 {transcript_info['count']} 筆土地謄本）

⚠️ 這不是順序問題！兩者毫無關聯！

💡 建議：取消修改，重新使用【地籍便民系統】建立 data.json
   這樣才能為新地段建立正確的基礎資料結構
"""

                return (False, warning, data_json_info, transcript_info, is_order_issue)

        # 如果任一檔案不存在，返回提示
        if not data_json_info:
            return (True, "data.json 不存在或為空", None, None, None)
        if not transcript_info:
            # 🔥 即使沒有謄本資料，仍應返回 data_json_info 以便更新標題
            return (True, "找不到謄本結構化資料", data_json_info, None, None)

        return (True, None, data_json_info, transcript_info, None)

    except Exception as e:
        return (True, f"檢查時發生錯誤：{e}", None, None, None)

def fix_data_json_order(transcript_info):
    """
    根據 transcript_info，調整 data.json 的順序
    將符合的地號移到第一位
    """
    try:
        if not os.path.exists('data.json'):
            return False, "data.json 不存在"

        with open('data.json', 'r', encoding='utf-8') as f:
            data_list = json.load(f)

        if not data_list:
            return False, "data.json 為空"

        # 正規化地號格式
        def normalize_lot(lot):
            if '-' in lot:
                parts = lot.split('-')
                return f"{int(parts[0])}-{int(parts[1])}"
            return str(int(lot))

        target_lot = normalize_lot(transcript_info['lot_number'])

        # 找到符合的項目
        target_idx = None
        for i, item in enumerate(data_list):
            item_lot = normalize_lot(item.get('lot_number', ''))
            if (item.get('area') == transcript_info['area'] and
                item.get('section') == transcript_info['section'] and
                item_lot == target_lot):
                target_idx = i
                break

        if target_idx is None:
            return False, f"在 data.json 中找不到 {transcript_info['area']}{transcript_info['section']} {transcript_info['lot_number']}"

        if target_idx == 0:
            return True, "data.json 的第一筆已經是正確的地號，無需調整"

        # 移到第一位
        data_list.insert(0, data_list.pop(target_idx))

        # 儲存
        save_data_json(data_list)

        return True, f"已將 {transcript_info['lot_number']} 移到 data.json 的第一位"

    except Exception as e:
        return False, f"調整失敗：{e}"

def fix_transcript_json_order(data_info, transcript_info):
    """
    依 data.json 的完整順序重排謄本結構化 JSON。

    - 對 data.json 每一筆，找出謄本中對應的記錄，依序排在前面
    - 謄本中無對應的記錄（例如建物謄本或多餘的土地謄本）按原相對順序附在後面
    """
    try:
        import glob
        import re as _re

        # 🔧 優先使用目前已載入的謄本 JSON；若無則退回「最新修改時間」的檔案
        latest_transcript = None
        if transcript_info and transcript_info.get('file_path') and os.path.exists(transcript_info['file_path']):
            latest_transcript = transcript_info['file_path']
        else:
            output_dir = os.path.join(BASE_DIR, "output")
            transcript_files = glob.glob(os.path.join(output_dir, "trans_*.json")) + glob.glob(os.path.join(output_dir, "transcript_data_*.json"))
            if not transcript_files:
                return False, "找不到謄本結構化 JSON 檔案"
            latest_transcript = max(transcript_files, key=os.path.getmtime)

        # 讀取謄本
        with open(latest_transcript, 'r', encoding='utf-8') as f:
            transcript_data = json.load(f)

        if not transcript_data:
            return False, "謄本結構化 JSON 為空"

        # 讀取 data.json 完整列表
        data_json_path = get_data_json_path()
        if not os.path.exists(data_json_path):
            return False, "data.json 不存在"
        with open(data_json_path, 'r', encoding='utf-8') as f:
            data_list = json.load(f)
        if not data_list:
            return False, "data.json 為空"

        # 正規化地號格式
        def normalize_lot(lot):
            lot = str(lot or '').strip().replace('地號', '').replace('建號', '')
            try:
                if '-' in lot:
                    parts = lot.split('-')
                    return f"{int(parts[0])}-{int(parts[1])}"
                if len(lot) >= 8:
                    return f"{int(lot[:4])}-{int(lot[4:])}"
                return f"{int(lot)}-0"
            except (ValueError, IndexError):
                return lot

        # 1. data.json 的 key 序列（依 data.json 順序）
        data_keys = []
        for d in data_list:
            a = (d.get('area', '') or '').strip()
            s = (d.get('section', '') or '').strip()
            l = normalize_lot(d.get('lot_number', ''))
            if a and s and l:
                data_keys.append((a, s, l))

        # 2. 為每筆謄本算出 key（無法解析則為 None）
        # 2-a. 為每筆謄本算出 key
        #   土地謄本: ('land', area, section, norm_lot)
        #   建物謄本: ('building', section, norm_lot)  — 用「建物坐落地號」配對到對應土地
        #     注意：建物坐落地號通常沒有「區/鄉/鎮/市」前綴（如「大港段二小段  1506-0000」）
        #     所以建物的 key 只用 section + lot，配對時忽略 area
        transcript_keys = []
        for t in transcript_data:
            type_ = t.get("謄本類型", "")
            basic = t.get("基本資訊", {}) or {}
            land_number = basic.get("地號建號", "")
            if "土地謄本" in type_:
                m = _re.search(r'([^\s]+(?:區|鄉|鎮|市))([^\s]+段)\s*(\d+-?\d*)', land_number)
                transcript_keys.append(
                    ('land', m.group(1), m.group(2), normalize_lot(m.group(3))) if m else None
                )
            elif "建物謄本" in type_:
                mark = t.get("標示部", {}) or {}
                location = mark.get("建物坐落地號", "") or ""
                # 配對第一個「段+地號」（多筆坐落時只取首筆作為主要對應）
                mb = _re.search(r'([^\s]+段)\s*(\d+-?\d*)', location)
                transcript_keys.append(
                    ('building', mb.group(1), normalize_lot(mb.group(2))) if mb else None
                )
            else:
                transcript_keys.append(None)

        # 3. 排序策略：「土地全部在前（依 data.json 順序），建物全部在後（依原順序）」
        #    這樣後續統計處理較單純，不需要考慮土地建物交錯
        used = [False] * len(transcript_data)
        land_order = []      # 依 data.json 順序的土地
        unmatched_lands = [] # 不在 data.json 的土地（依原順序附在土地區段尾端）
        building_order = []  # 所有建物（依原順序，附在最末段）

        # 先依 data.json 順序放對應的土地
        for dkey in data_keys:
            d_area, d_section, d_norm_lot = dkey
            for i, tk in enumerate(transcript_keys):
                if used[i] or tk is None or tk[0] != 'land':
                    continue
                if tk[1] == d_area and tk[2] == d_section and tk[3] == d_norm_lot:
                    land_order.append(i)
                    used[i] = True
                    break

        # 依原順序掃一遍：未配對的土地接在後面，建物全部丟到 building_order
        for i, tk in enumerate(transcript_keys):
            if used[i]:
                continue
            if tk is not None and tk[0] == 'land':
                unmatched_lands.append(i)
                used[i] = True
            elif tk is not None and tk[0] == 'building':
                building_order.append(i)
                used[i] = True

        # 完全無法解析類型的記錄（None）放最後保底
        leftovers = [i for i in range(len(transcript_data)) if not used[i]]

        new_order = land_order + unmatched_lands + building_order + leftovers

        matched_lands = len(land_order)
        unmatched_land_count = len(unmatched_lands)
        building_count = len(building_order)
        if matched_lands == 0:
            return False, f"在謄本資料中找不到任何 data.json 對應的地號（共 {len(data_keys)} 筆）"

        # 4. 已是正確順序就不寫檔
        if new_order == list(range(len(transcript_data))):
            _parts = [f"土地 {matched_lands} 筆對齊"]
            if building_count:
                _parts.append(f"建物 {building_count} 筆")
            return True, f"順序已與 data.json 一致，無需調整（{'、'.join(_parts)}）"

        # 5. 套用新順序
        transcript_data = [transcript_data[i] for i in new_order]

        # 儲存（加入檔案佔用處理和重試機制）
        import time
        max_retries = 3
        retry_delay = 0.5  # 秒

        for attempt in range(max_retries):
            try:
                # 先寫入暫存檔，確保資料完整性
                temp_file = latest_transcript + '.tmp'
                with open(temp_file, 'w', encoding='utf-8') as f:
                    json.dump(transcript_data, f, ensure_ascii=False, indent=2)

                # 驗證暫存檔內容是否有效
                with open(temp_file, 'r', encoding='utf-8') as f:
                    verify_data = json.load(f)
                    if len(verify_data) != len(transcript_data):
                        raise ValueError("暫存檔驗證失敗：資料筆數不符")

                # 替換原檔案
                if os.path.exists(latest_transcript):
                    backup_file = latest_transcript + '.bak'
                    # 移除舊備份
                    if os.path.exists(backup_file):
                        os.remove(backup_file)
                    # 原檔案改名為備份
                    os.rename(latest_transcript, backup_file)

                # 暫存檔改名為正式檔案
                os.rename(temp_file, latest_transcript)

                # 刪除備份（成功後）
                if os.path.exists(backup_file):
                    os.remove(backup_file)

                _parts = [f"土地 {matched_lands} 筆依 data.json 排序"]
                if unmatched_land_count:
                    _parts.append(f"另 {unmatched_land_count} 筆土地不在 data.json 中（接於後）")
                if building_count:
                    _parts.append(f"建物 {building_count} 筆統一置於最末")
                return True, f"已重排謄本：{'，'.join(_parts)}（{os.path.basename(latest_transcript)}）"

            except PermissionError as pe:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    continue
                else:
                    # 清理暫存檔
                    if os.path.exists(temp_file):
                        try:
                            os.remove(temp_file)
                        except:
                            pass
                    return False, f"檔案被佔用，無法寫入：{pe}\n請確認沒有其他程式開啟此檔案，然後重試"

            except Exception as write_error:
                # 清理暫存檔
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except:
                        pass
                # 嘗試從備份恢復
                if os.path.exists(backup_file) and not os.path.exists(latest_transcript):
                    try:
                        os.rename(backup_file, latest_transcript)
                    except:
                        pass
                raise write_error

        return False, "儲存失敗：超過重試次數"

    except json.JSONDecodeError as je:
        return False, f"JSON 格式錯誤：{je}\n檔案可能已損壞，請檢查：{latest_transcript}"

    except Exception as e:
        import traceback
        return False, f"調整失敗：{e}\n{traceback.format_exc()}"

def update_final_data(new_data):
    """更新最終資料（合併新資料）"""
    final_data = load_final_data()
    final_data.update(new_data)
    save_final_data(final_data)

# 連續模式相關變數
continuous_mode_active = False
continuous_mode_queue = []
continuous_mode_index = 0
continuous_mode_timer = None
current_running_button = None  # 🔥 追蹤當前正在執行的按鈕

def get_process_children(parent_pid, include_parent=True):
    try:
        parent = psutil.Process(parent_pid)
        children = parent.children(recursive=True)
        if include_parent:
            children.insert(0, parent)
        return children
    except psutil.NoSuchProcess:
        return []

def kill_subprocess_and_children(process):
    if not process:
        return
    try:
        parent_pid = process.pid
        children = get_process_children(parent_pid, include_parent=False)
        chrome_count = 0
        for child in children:
            try:
                if 'chrome' in child.name().lower():
                    child.terminate()
                    chrome_count += 1
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        if chrome_count > 0:
            print(f"已終止 {chrome_count} 個相關的 Chrome 進程", flush=True)
            time.sleep(0.5)
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=2)
                print("程序已終止", flush=True)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait()
                print("程序已強制終止", flush=True)
        for child in children:
            try:
                if child.is_running():
                    child.kill()
            except:
                pass
    except Exception as e:
        print(f"終止進程時發生錯誤: {e}", flush=True)

def read_output(process):
    global is_running, current_running_button
    while is_running:
        try:
            output = process.stdout.readline()
            if output:
                output_str = output.strip()
                # 🔥 再次檢查 is_running，避免顯示關閉後的輸出
                if not is_running:
                    break

                # 🔥 檢查是否為 GUI 預填指令
                if output_str.startswith("__GUI_PREFILL__:"):
                    prefill_value = output_str.split(":", 1)[1] if ":" in output_str else ""
                    # 在 GUI 輸入框中預填值
                    root.after(0, lambda: prefill_input_field(prefill_value))
                    continue  # 不顯示這個特殊標記

                # 🔥 檢查是否為進度回報
                # 格式：__PROGRESS__:current/total:elapsed_sec:eta_sec
                if output_str.startswith("__PROGRESS__:"):
                    try:
                        payload = output_str[len("__PROGRESS__:"):]
                        parts = payload.split(":")
                        c_str, t_str = parts[0].split("/")
                        c_val = int(c_str)
                        t_val = int(t_str)
                        e_val = int(parts[1]) if len(parts) > 1 else 0
                        eta_val = int(parts[2]) if len(parts) > 2 else 0
                        root.after(0, lambda c=c_val, t=t_val, e=e_val, a=eta_val: handle_progress_message(c, t, e, a))
                    except Exception:
                        pass
                    continue  # 不顯示這個特殊標記

                # 🔥 偵測子程序輸出的長分隔線，自動換成 auto-fit divider，避免不同螢幕寬度下換行
                _stripped = output_str.strip()
                if _stripped:
                    # 整行同一字元，長度 ≥ 10 → 視為分隔線
                    if len(_stripped) >= 10 and len(set(_stripped)) == 1 and _stripped[0] in "=*＊-─━─":
                        update_divider(_stripped[0])
                        continue
                    # 「===   標題文字   ===」型置中標題
                    _m_center = re.match(r'^([=*＊\-─━─]{3,})\s+(.+?)\s+([=*＊\-─━─]{3,})$', _stripped)
                    if _m_center and _m_center.group(1)[0] == _m_center.group(3)[0]:
                        update_centered_divider(_m_center.group(2), _m_center.group(1)[0])
                        continue

                update_message(output_str)
                if "已完成執行" in output_str:
                    update_message("[完成] 該程序正常結束")
                    # 🔥 加上視覺分隔線
                    update_message("")
                    update_divider("＊")
                    update_message("")

                    # 🔥 如果是 land.py 完成，立即更新視窗標題
                    if current_running_button == start_land_button:
                        root.after(100, update_title_from_data_json)  # 延遲 100ms 確保 data.json 寫入完成

                    # 🔥 恢復按鈕狀態
                    if continuous_mode_active:
                        # 連續模式：恢復當前按鈕文字並禁用
                        if current_running_button:
                            original_text = current_running_button.cget("text").replace("關閉 ", "")
                            current_running_button.config(text=original_text, state=tk.DISABLED)
                            current_running_button = None
                    else:
                        # 非連續模式：正常啟用所有按鈕
                        enable_all_buttons()
                        for button in all_buttons:
                            if (button != real_estate_price_button and
                                button != transcript_button and
                                button != continuous_mode_button and
                                hasattr(button, 'cget') and
                                button.cget("text").startswith("關閉")):
                                original_text = button.cget("text").replace("關閉 ", "")
                                button.config(text=original_text)

                    is_running = False

                    # 🔥 子程式正常結束後，隱藏進度條
                    root.after(0, hide_task_progress)

                    # 檢查是否在連續模式中
                    if continuous_mode_active:
                        schedule_next_script()
                    break
            else:
                poll_result = process.poll()
                if poll_result is not None:
                    # 🔥 只在正常結束時顯示訊息（不是被手動關閉）
                    if is_running:
                        update_message(f"🔧 子程序已結束，返回碼: {poll_result}")
                        # 🔥 加上視覺分隔線
                        update_message("")
                        update_divider("＊")
                        update_message("")
                    is_running = False

                    # 🔥 如果是 land.py 完成，立即更新視窗標題
                    if current_running_button == start_land_button:
                        root.after(100, update_title_from_data_json)  # 延遲 100ms 確保 data.json 寫入完成

                    # 🔥 恢復按鈕狀態
                    if continuous_mode_active:
                        # 連續模式：恢復當前按鈕文字並禁用
                        if current_running_button:
                            original_text = current_running_button.cget("text").replace("關閉 ", "")
                            current_running_button.config(text=original_text, state=tk.DISABLED)
                            current_running_button = None
                    else:
                        # 非連續模式：正常啟用所有按鈕
                        enable_all_buttons()
                        for button in all_buttons:
                            if (button != real_estate_price_button and
                                button != transcript_button and
                                button != continuous_mode_button and
                                hasattr(button, 'cget') and
                                button.cget("text").startswith("關閉")):
                                original_text = button.cget("text").replace("關閉 ", "")
                                button.config(text=original_text)

                    # 🔥 電子謄本結構化 結束且成功 → 自動觸發讀取 JSON（兩種模式皆然）
                    # 連續模式下 handle_read_json_function 會自行呼叫 schedule_next_script，
                    # 所以本處需用 auto_trigger_json 旗標避免下一支腳本被跑兩次
                    auto_trigger_json = (current_running_button == structured_transcript_button
                                         and poll_result == 0)
                    if auto_trigger_json:
                        root.after(500, handle_read_json_function)

                    # 🔥 子程式結束（無論正常或異常），隱藏進度條
                    root.after(0, hide_task_progress)

                    # 檢查是否在連續模式中（若已自動觸發讀取 JSON 則交由它接手 schedule_next_script）
                    if continuous_mode_active and not auto_trigger_json:
                        schedule_next_script()
                    break
                else:
                    time.sleep(0.1)
        except Exception as e:
            # 🔥 只在非手動關閉的情況下顯示錯誤
            if is_running:
                update_message(f"[錯誤] 讀取輸出時發生錯誤: {str(e)}")
            is_running = False
            # 🔥 發生錯誤，隱藏進度條
            root.after(0, hide_task_progress)
            # 🔥 錯誤時也要恢復按鈕
            if continuous_mode_active and current_running_button:
                original_text = current_running_button.cget("text").replace("關閉 ", "")
                current_running_button.config(text=original_text, state=tk.DISABLED)
                current_running_button = None
            else:
                enable_all_buttons()
            break

def get_available_work_area():
    """取得可用的工作區域（考慮 DPI 縮放）"""
    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", wintypes.LONG),
            ("top", wintypes.LONG),
            ("right", wintypes.LONG),
            ("bottom", wintypes.LONG)
        ]

    # 使用 DPI 感知的方法取得工作區
    try:
        # 嘗試使用 GetMonitorInfo（支援 DPI 感知）
        from ctypes import POINTER, WINFUNCTYPE
        MONITORINFO_CLASS = type('MONITORINFO', (ctypes.Structure,), {
            '_fields_': [
                ("cbSize", wintypes.DWORD),
                ("rcMonitor", RECT),
                ("rcWork", RECT),
                ("dwFlags", wintypes.DWORD)
            ]
        })

        hMonitor = ctypes.windll.user32.MonitorFromWindow(0, 1)  # MONITOR_DEFAULTTOPRIMARY
        monitor_info = MONITORINFO_CLASS()
        monitor_info.cbSize = ctypes.sizeof(MONITORINFO_CLASS)

        if ctypes.windll.user32.GetMonitorInfoW(hMonitor, ctypes.byref(monitor_info)):
            return monitor_info.rcWork
    except Exception:
        pass

    # Fallback: 使用舊方法
    SPI_GETWORKAREA = 0x0030
    rect = RECT()
    ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
    return rect

ANSI_COLOR_MAP = {
    '30': 'black', '31': 'red', '32': 'green', '33': 'yellow',
    '34': 'blue', '35': 'magenta', '36': 'cyan', '37': 'white',
    '90': 'gray', '91': 'red', '92': 'green', '93': 'yellow',
    # 🔥 94（bright blue）改用 'skyblue'，原本 'blue' 在黑底上太深難讀
    '94': 'skyblue', '95': 'magenta', '96': 'cyan', '97': 'white',
    '40': 'black', '41': 'red', '42': 'green', '43': 'yellow',
    '44': 'blue', '45': 'magenta', '46': 'cyan', '47': 'white',
}

# 🔥 root 已在最頂端建立（splash 用），這裡不再重新建立
# 移除 splash 內容，準備建主介面
try:
    _splash_frame.destroy()
except Exception:
    pass

# 🔥 取得 DPI 縮放比例
try:
    # 取得主螢幕的 DPI
    hdc = ctypes.windll.user32.GetDC(0)
    dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
    ctypes.windll.user32.ReleaseDC(0, hdc)
    dpi_scale = dpi / 96.0  # 96 DPI = 100%
except Exception as e:
    dpi_scale = 1.0

# Tkinter 回報的是邏輯像素（受 DPI 影響）
screen_width_logical = root.winfo_screenwidth()
screen_height_logical = root.winfo_screenheight()

# 啟動時嘗試載入設定
try:
    cfg = load_config()
    json_dir = cfg.get("json_dir", JSON_DIR_DEFAULT)
    selected_json_path = cfg.get("selected_json")
except Exception:
    json_dir = JSON_DIR_DEFAULT
    selected_json_path = None

work_area = get_available_work_area()

# 🔥 智能視窗大小計算：統一使用物理像素（work_area 已經是物理像素）
available_width = work_area.right - work_area.left
available_height = work_area.bottom - work_area.top
screen_width = available_width  # 使用物理像素的螢幕寬度
screen_height = available_height  # 使用物理像素的螢幕高度

# 🔥 視窗配置：主程式固定佔 1/3，Chrome 佔 2/3（使用物理像素）
# 維持 1/3 比例，不因 DPI 而擴大（避免與網頁重疊）
window_width = screen_width // 3

# 高度使用工作區完整高度（扣除邊距，稍微留空避免切齊工作列）
window_height = available_height - 40

# 計算位置：靠右側（物理像素）
window_x = work_area.right - window_width
window_y = work_area.top

# 🔥 重要：SetProcessDpiAwareness(2) 使 tkinter 使用物理像素，不需要轉換！
# 直接使用物理像素
window_width_for_tk = window_width
window_height_for_tk = window_height
window_x_for_tk = window_x
window_y_for_tk = window_y

# 🔥 計算建議的 Chrome 視窗大小（供其他程式參考）
# 確保 Chrome + main 視窗不重疊，並留 20px 間距（使用物理像素）
chrome_max_width = available_width - window_width - 20

# 🔥 儲存視窗配置到設定檔（全部使用物理像素，供其他程式讀取）
try:
    config_data = {
        'dpi_scale': dpi_scale,
        'screen_width': screen_width,  # 物理像素
        'screen_height': screen_height,  # 物理像素
        'work_area': {  # 物理像素
            'left': work_area.left,
            'top': work_area.top,
            'right': work_area.right,
            'bottom': work_area.bottom
        },
        'main_window': {  # 物理像素
            'width': window_width,
            'height': window_height,
            'x': window_x,
            'y': window_y
        }
    }
    import json
    config_path = os.path.join(PROGRAM_DIR, 'window_config.json')
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, indent=2)
except Exception as e:
    pass


# 🔥 使用物理像素設定 tkinter 視窗（因為已啟用 SetProcessDpiAwareness(2)）
root.geometry(f"{window_width_for_tk}x{window_height_for_tk}+{window_x_for_tk}+{window_y_for_tk}")

# 🔥 根據 DPI 和視窗寬度決定標題格式（避免高 DPI 下標題被截斷）
BASE_TITLE = "地籍資料查詢系統"  # 修改基礎名稱
if dpi_scale >= 1.75:
    # 175%: 極簡格式（Windows 控制按鈕佔用很多空間）
    TITLE_FORMAT = f"地籍系統 v{{version}}"  # 縮短系統名稱
    TITLE_WITH_LOCATION = f"地籍系統 - {{location}}"  # 只顯示地號，不顯示版本
elif dpi_scale >= 1.5:
    # 150%: 簡短格式
    TITLE_FORMAT = f"{BASE_TITLE} v{{version}}"
    TITLE_WITH_LOCATION = f"{BASE_TITLE} - {{location}}"  # 省略版本號
else:
    # 100-125%: 完整格式
    TITLE_FORMAT = f"{BASE_TITLE} v{{version}}"
    TITLE_WITH_LOCATION = f"{BASE_TITLE} v{{version}} - {{location}}"

root.title(TITLE_FORMAT.format(version=VERSION))

# 🔥 動態計算 UI 配置（自動適應任何解析度和 DPI）
def calculate_ui_config(window_width, dpi_scale):
    """根據視窗寬度和 DPI 動態計算最佳 UI 配置

    🔥 關鍵洞察：
    - window_width 是物理像素（因為 SetProcessDpiAwareness(2)）
    - tkinter font size 會被 DPI 放大！font_size=16 在 175% DPI 下實際是 16×1.75=28px
    - 實測：tkinter 按鈕寬度 ≈ width × font_size × 0.87（係數約 0.85-0.9）
    """

    # 🔥 實測係數：按鈕實際像素寬度 ≈ width × font_size × 0.87 × dpi_scale
    CHAR_COEF = 0.87  # 實測得到的係數

    SIDE_PADDING = 10  # 視窗左右邊距（縮小以容納更多）
    MIN_COLUMN_GAP = 2  # 最小欄間距（共 2 個間距）

    # 可用於三欄的總寬度（物理像素）
    available_width = window_width - SIDE_PADDING * 2 - MIN_COLUMN_GAP * 2
    column_width = available_width / 3

    # 🔥 按鈕需要顯示的最長文字是「謄本與土增稅」= 6 個中文字
    # width=12 可以容納 6 個中文字
    TARGET_WIDTH = 12  # 按鈕 width 參數

    # 🔥 公式推導（考慮 DPI 放大）：
    # 按鈕像素寬度 = TARGET_WIDTH × font_size × CHAR_COEF × dpi_scale
    # 要讓按鈕放進欄寬：TARGET_WIDTH × font_size × CHAR_COEF × dpi_scale ≤ column_width
    # 所以：font_size ≤ column_width / (TARGET_WIDTH × CHAR_COEF × dpi_scale)
    max_font = column_width / (TARGET_WIDTH * CHAR_COEF * dpi_scale)
    button_font_size = int(max_font)
    button_font_size = max(9, min(16, button_font_size))

    # 按鈕字元寬度
    button_width = TARGET_WIDTH

    # 🔥 計算實際三欄佔用寬度（物理像素），確保不超出
    pixel_per_button = button_width * button_font_size * CHAR_COEF * dpi_scale
    actual_three_buttons_width = pixel_per_button * 3

    # 🔥 如果三欄超出可用寬度，縮減按鈕寬度
    while actual_three_buttons_width > available_width and button_width > 8:
        button_width -= 1
        pixel_per_button = button_width * button_font_size * CHAR_COEF * dpi_scale
        actual_three_buttons_width = pixel_per_button * 3

    # 計算剩餘空間給欄距
    remaining_space = available_width - actual_three_buttons_width
    column_padx = int(remaining_space / 4)  # 分給兩個間隙
    column_padx = max(0, min(10, column_padx))

    # 🔥 子頁面計算：子頁面也是三欄佈局，按鈕文字更長（如「全國地政電子謄本」= 8 個中文字 = 16 字元單位）
    SUB_TARGET_WIDTH = 16  # 子頁面按鈕需要 16 字元單位
    SUB_MIN_WIDTH = 14     # 🔥 按鈕寬度下限（可縮到 14 以保持較大字體）

    # 🔥 子頁面：優先保持字體 >= 9，寧可縮減按鈕寬度
    sub_page_font_size = 14  # 從最大開始
    sub_button_width = SUB_TARGET_WIDTH

    # 嘗試找到合適的字體和按鈕寬度組合
    while sub_page_font_size >= 9:
        sub_pixel_per_button = sub_button_width * sub_page_font_size * CHAR_COEF * dpi_scale
        actual_sub_three_buttons_width = sub_pixel_per_button * 3

        if actual_sub_three_buttons_width <= available_width:
            break  # 找到合適的組合

        # 先縮減按鈕寬度（但不低於下限）
        if sub_button_width > SUB_MIN_WIDTH:
            sub_button_width -= 1
        else:
            # 按鈕寬度已達下限，縮減字體並重設按鈕寬度
            sub_page_font_size -= 1
            sub_button_width = SUB_TARGET_WIDTH

    # 輸入框字體
    input_font_size = int(button_font_size * 1.25)
    input_font_size = max(14, min(20, input_font_size))

    # 訊息字體
    message_font_size = max(9, min(15, button_font_size - 1))

    return {
        'button_font_size': button_font_size,
        'sub_page_font_size': sub_page_font_size,
        'input_font_size': input_font_size,
        'message_font_size': message_font_size,
        'button_width': button_width,
        'sub_button_width': sub_button_width,
        'column_padx': column_padx
    }

# 計算 UI 配置
ui_config = calculate_ui_config(window_width_for_tk, dpi_scale)
button_font_size = ui_config['button_font_size']
sub_page_font_size = ui_config['sub_page_font_size']
input_font_size = ui_config['input_font_size']
message_font_size = ui_config['message_font_size']

# 調試輸出
print(f"[UI配置] 視窗={window_width_for_tk}px, DPI={dpi_scale:.2f}")
print(f"[UI配置] 按鈕字體={button_font_size}, 子頁字體={sub_page_font_size}, 輸入字體={input_font_size}")

custom_font = font.Font(family="Microsoft JhengHei", size=button_font_size)  # 主頁面按鈕字體
sub_page_font = font.Font(family="Microsoft JhengHei", size=sub_page_font_size)  # 子頁面按鈕字體
input_font = font.Font(family="Microsoft JhengHei", size=input_font_size)  # 輸入框字體

# === JSON 狀態列 ===
status_frame = tk.Frame(root, bg='#F0F0F0', relief=tk.SUNKEN, bd=1, height=35)
status_frame.pack(side=tk.TOP, fill=tk.X)
status_frame.pack_propagate(False)  # 🔥 防止子元件撐大 frame

json_status_label = tk.Label(
    status_frame,
    text="JSON: 未選擇",
    font=("Microsoft JhengHei", max(8, message_font_size - 2)),  # 🔥 根據訊息框字體動態調整
    bg='#F0F0F0',
    fg='#333333',
    anchor='w',
    padx=8,  # 🔥 減少內距
    pady=3   # 🔥 減少內距
)
json_status_label.pack(fill=tk.X)

frame_top = tk.Frame(root)
frame_top.pack(side=tk.TOP, fill=tk.X, pady=(10, 0))

# 🔥 先定義 frame_bottom（先 pack 在底部）
frame_bottom = tk.Frame(root, height=60)
frame_bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
frame_bottom.pack_propagate(False)

# 🔥 再定義 frame_middle（填滿中間剩餘空間）
# 使用物理像素計算（因為 tkinter 在 DPI 感知模式下使用物理像素）
calculated_message_height = window_height_for_tk - 200
message_height = max(250, min(600, calculated_message_height))

frame_middle = tk.Frame(root, height=message_height)
frame_middle.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=10)  # 🔥 改回 BOTH + expand
# 不要 pack_propagate(False)，讓它可以自動調整

message_box = scrolledtext.ScrolledText(
    frame_middle, wrap=tk.WORD, state=tk.DISABLED,
    font=("Microsoft JhengHei", message_font_size),  # 🔥 使用動態字體大小
    bg='black', fg='white', insertbackground='white'
)
message_box.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
message_box.config(spacing3=1)

# 🔥 長任務進度條（預設隱藏，子程式送出 __PROGRESS__ 訊號才會出現）
# 放在主介面按鈕列與訊息區之間，給跑長的批次任務（例如 finance_tax 222 筆）使用
task_progress_frame = tk.Frame(root, bg='#E3F2FD')
task_progress_label = tk.Label(
    task_progress_frame,
    text="",
    font=("Microsoft JhengHei", 11),
    bg='#E3F2FD',
    fg='#0D47A1',
    anchor='w'
)
task_progress_label.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(4, 2))
task_progress_bar = ttk.Progressbar(
    task_progress_frame,
    orient='horizontal',
    mode='determinate',
    length=100,
    maximum=100
)
task_progress_bar.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(0, 6))
# 初始不 pack task_progress_frame 本身，等收到進度訊號時才顯示

input_field = tk.Entry(frame_bottom, font=input_font, fg="gray")
input_field.insert(0, "請在此輸入...")
def on_entry_click(event):
    if input_field.get() == "請在此輸入...":
        input_field.delete(0, "end")
        input_field.config(fg="black")
def on_focus_out(event):
    if input_field.get() == "":
        input_field.insert(0, "請在此輸入...")
        input_field.config(fg="gray")
input_field.bind("<FocusIn>", on_entry_click)
input_field.bind("<FocusOut>", on_focus_out)
input_field.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

# 🔥 使用動態計算的按鈕寬度和間距
BUTTON_WIDTH = ui_config['button_width']
SUB_PAGE_BUTTON_WIDTH = ui_config['sub_button_width']
COLUMN_PADX = ui_config['column_padx']

# 調試輸出
print(f"[UI配置] 按鈕寬度={BUTTON_WIDTH}, 子頁按鈕={SUB_PAGE_BUTTON_WIDTH}, 欄距={COLUMN_PADX}")

# 主頁面三欄
left_button_frame = tk.Frame(frame_top)
middle_button_frame = tk.Frame(frame_top)
right_button_frame = tk.Frame(frame_top)
left_button_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, anchor="e")
middle_button_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, expand=True, anchor="center")
right_button_frame.pack(side=tk.RIGHT, padx=COLUMN_PADX, anchor="w")

# 實價登錄頁面
real_estate_price_page = tk.Frame(frame_top)
real_estate_left_frame = tk.Frame(real_estate_price_page)
real_estate_middle_frame = tk.Frame(real_estate_price_page)
real_estate_right_frame = tk.Frame(real_estate_price_page)
real_estate_left_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, anchor="e")
real_estate_middle_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, expand=True, anchor="center")
real_estate_right_frame.pack(side=tk.RIGHT, padx=COLUMN_PADX, anchor="w")

# 謄本專區頁面
transcript_page = tk.Frame(frame_top)
transcript_left_frame = tk.Frame(transcript_page)
transcript_middle_frame = tk.Frame(transcript_page)
transcript_right_frame = tk.Frame(transcript_page)
transcript_left_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, anchor="e")
transcript_middle_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, expand=True, anchor="center")
transcript_right_frame.pack(side=tk.RIGHT, padx=COLUMN_PADX, anchor="w")

# 🔥 功能分頁（案件管理相關功能）
utility_page = tk.Frame(frame_top)
utility_left_frame = tk.Frame(utility_page)
utility_middle_frame = tk.Frame(utility_page)
utility_right_frame = tk.Frame(utility_page)
utility_left_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, anchor="e")
utility_middle_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, expand=True, anchor="center")
utility_right_frame.pack(side=tk.RIGHT, padx=COLUMN_PADX, anchor="w")

# 🔥 土地使用分區分頁
land_use_page = tk.Frame(frame_top)
land_use_left_frame = tk.Frame(land_use_page)
land_use_middle_frame = tk.Frame(land_use_page)
land_use_right_frame = tk.Frame(land_use_page)
land_use_left_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, anchor="e")
land_use_middle_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, expand=True, anchor="center")
land_use_right_frame.pack(side=tk.RIGHT, padx=COLUMN_PADX, anchor="w")

# 🔥 使用執照查詢分頁（取代原本「使用執照查詢」單一按鈕，現在進入次頁面選縣市）
building_license_page = tk.Frame(frame_top)
building_license_left_frame = tk.Frame(building_license_page)
building_license_middle_frame = tk.Frame(building_license_page)
building_license_right_frame = tk.Frame(building_license_page)
building_license_left_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, anchor="e")
building_license_middle_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, expand=True, anchor="center")
building_license_right_frame.pack(side=tk.RIGHT, padx=COLUMN_PADX, anchor="w")

# 🔥 為所有子頁面設定固定高度，與主頁面對齊
for page in [transcript_page, real_estate_price_page, utility_page, land_use_page, building_license_page]:
    page.config(height=200)  # 固定所有頁面高度

# 🔥 修改按鈕樣式 - 統一所有按鈕使用相同寬度（包含有顏色的按鈕）
button_style = {
    'font': custom_font, 'width': BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'borderwidth': 2, 'bg': '#E0E0E0', 'fg': 'black'
}

# 🔥 子頁面按鈕樣式（使用較寬的寬度和較小的字體）
sub_page_button_style = {
    'font': sub_page_font, 'width': SUB_PAGE_BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'borderwidth': 2, 'bg': '#E0E0E0', 'fg': 'black'
}

# 🔥 地籍便民系統的特殊樣式（橙色背景）
# 統一使用 custom_font（不加粗），確保所有按鈕視覺寬度一致
priority_button_style = {
    'font': custom_font,  # 🔥 改用統一字體，不再加粗
    'width': BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'borderwidth': 3, 'bg': '#FF8C00', 'fg': 'white',  # 橙色背景，白色文字
    'activebackground': '#FF7F00', 'activeforeground': 'white'  # 滑鼠懸停效果
}

# 連續模式按鈕樣式（淺藍色）
continuous_button_style = {
    'font': custom_font,  # 🔥 改用統一字體，不再加粗
    'width': BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'borderwidth': 3, 'bg': '#87CEEB', 'fg': 'black',
    'activebackground': '#77BEDD', 'activeforeground': 'black'
}

is_running = False
process = None

# 連續模式相關函數
def select_continuous_mode_scripts():
    """選擇連續模式要執行的腳本"""
    global continuous_mode_queue
    
    # 定義所有可執行的腳本（按主頁面按鈕順序：左排→中排→右排，各排由上至下，遇分頁依序展開）
    all_scripts = [
        # ===== 左排 =====
        # 1. 地籍便民系統
        ("地籍便民系統 (依地段地號建立目錄)", "land.py"),
        # 2. 謄本與土增稅（分頁內容）
        ("【謄本與土增稅】全國地政電子謄本", "hinet.py"),
        ("【謄本與土增稅】全功能地籍資料查詢", "qpt_hinet.py"),
        ("【謄本與土增稅】電子謄本結構化", "GUI_transcript_pdf 1141021-01.py"),
        ("【功能】讀取結構化JSON", "read_json"),
        ("【謄本與土增稅】財政部土增稅", "finance_tax.py"),
        ("【謄本與土增稅】高雄市土增稅估算", "kaohsiung_tax.py"),
        # 3. 使用執照查詢（高雄市單獨 / 屏東縣 + 台南市共用 NBUPIC）
        ("【使用執照查詢】高雄市", "nlma.py"),
        ("【使用執照查詢】屏東縣", "nlma_nbupic.py:屏東縣"),
        ("【使用執照查詢】台南市", "nlma_nbupic.py:臺南市"),
        # 4. 重要設施查詢
        ("重要設施查詢", "gbmap.py"),
        # ===== 中排 =====
        # 1. 土地使用分區（分頁內容）
        ("【土地使用分區】國土測繪圖資", "nlscmaps.py"),
        ("【土地使用分區】高雄都市計畫", "urbangis.py"),
        ("【土地使用分區】全國土地使用分區", "luz.py"),
        # 2. V523 系統
        ("V523 系統 (僅高雄市)", "V523.py"),
        # 3. 國土分區查詢
        ("國土分區查詢系統", "tcdmap.py"),
        # 4. 地質雲查詢
        ("地質雲查詢", "geologycloud.py"),
        # ===== 右排 =====
        # 1. 實價登錄（分頁內容）
        ("【實價登錄】內政部", "interior_price.py"),
        ("【實價登錄】104實價網", "price104.py"),
        # 2. 編輯物件資料
        ("【最終步驟】編輯物件資料", "edit_data")
    ]
    
    # 創建選擇對話框CC
    selection_window = tk.Toplevel(root)
    selection_window.title("連續模式 - 選擇執行腳本")

    # 🔥 根據 DPI 動態調整視窗大小和字體
    # 高 DPI 需要更寬的視窗來容納按鈕
    if dpi_scale >= 1.75:
        dialog_width = 600  # 175%：固定 600px
    elif dpi_scale >= 1.5:
        dialog_width = 550  # 150%：固定 550px
    else:
        dialog_width = 500  # 100%/125%：固定 500px
    dialog_height = window_height_for_tk  # 與主程式同高
    selection_window.geometry(f"{dialog_width}x{dialog_height}")
    selection_window.transient(root)
    selection_window.grab_set()

    # 置中顯示
    # 計算位置，讓對話框與主頁面併排
    dialog_x = root.winfo_x() - dialog_width - 20  # 主頁面左邊
    dialog_y = root.winfo_y()
    selection_window.geometry(f"+{dialog_x}+{dialog_y}")

    # 🔥 標題字體根據 DPI 調整
    title_font_size = max(12, int(14 / dpi_scale * 1.2))
    tk.Label(selection_window, text="請選擇要執行的程式：",
         font=("Microsoft JhengHei", title_font_size, "bold")).pack(pady=(10, 4))

    # 🔥 操作提醒（黃底，提示部分子程式需手動互動）
    warn_font_size = max(9, int(11 / dpi_scale * 1.2))
    warn_text = (
        "⚠ 請隨時注意程式運行狀況\n"
        "部分子程式需手動操作（輸入驗證碼、調整選項、按 Enter 等）"
    )
    tk.Label(selection_window, text=warn_text,
             font=("Microsoft JhengHei", warn_font_size),
             bg='#FFF3CD', fg='#856404',
             justify=tk.LEFT, padx=10, pady=6,
             relief=tk.SOLID, borderwidth=1).pack(fill=tk.X, padx=15, pady=(0, 8))

    # 建立勾選框架
    check_frame = tk.Frame(selection_window)
    check_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
    
    # 新增滾動條
    canvas = tk.Canvas(check_frame)
    scrollbar = tk.Scrollbar(check_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # 存儲勾選框變數
    check_vars = []

    # 🔥 使用執照查詢三縣市為互斥群組（每次只能選一個），可從 config 讀預設
    LICENSE_QUERY_SCRIPTS = {"nlma.py", "nlma_nbupic.py:屏東縣", "nlma_nbupic.py:臺南市"}
    try:
        _cfg_for_default = load_config() or {}
        license_default = _cfg_for_default.get("license_query_default", "nlma.py")
    except Exception:
        license_default = "nlma.py"
    license_vars = []  # 收集互斥群組的 var，建好後綁 trace
    license_header_inserted = False

    # 建立勾選框
    # 🔥 勾選框字體根據 DPI 調整
    checkbox_font_size = max(10, int(14 / dpi_scale * 1.2))
    for i, (name, script) in enumerate(all_scripts):
        is_license = script in LICENSE_QUERY_SCRIPTS

        # 🔥 第一個 license 項目前插入群組標題（非勾選元件）
        if is_license and not license_header_inserted:
            tk.Label(scrollable_frame, text="── 使用執照查詢（擇一） ──",
                     font=("Microsoft JhengHei", checkbox_font_size, "bold"),
                     fg='#555').pack(anchor='w', pady=(6, 2))
            license_header_inserted = True

        var = tk.BooleanVar()
        # 預設選取
        if is_license and script == license_default:
            var.set(True)
        check_vars.append((var, name, script))

        # 顯示文字：license 項目縮排 + 去掉【使用執照查詢】前綴
        if is_license:
            display_name = "    " + name.replace("【使用執照查詢】", "")
            license_vars.append(var)
        else:
            display_name = name

        cb = tk.Checkbutton(scrollable_frame, text=display_name, variable=var,
                   font=("Microsoft JhengHei", checkbox_font_size))
        cb.pack(anchor='w', pady=2)

    # 🔥 互斥邏輯：當 license 群組中某項變 True，其他自動變 False
    def _make_mutex(this_var, group):
        def _cb(*_):
            if this_var.get():
                for other in group:
                    if other is not this_var and other.get():
                        other.set(False)
        return _cb
    for v in license_vars:
        v.trace_add('write', _make_mutex(v, license_vars))
    
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # 綁定滑鼠滾輪事件（修正：綁定到 canvas 和 scrollable_frame 及其子元件）
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    # 綁定到 canvas 和 scrollable_frame
    canvas.bind("<MouseWheel>", _on_mousewheel)
    scrollable_frame.bind("<MouseWheel>", _on_mousewheel)

    # 綁定到所有勾選框（當滑鼠在勾選框上時也能滾動）
    def bind_mousewheel_to_children(widget):
        widget.bind("<MouseWheel>", _on_mousewheel)
        for child in widget.winfo_children():
            bind_mousewheel_to_children(child)

    bind_mousewheel_to_children(scrollable_frame)
    
    # 按鈕框架
    button_frame = tk.Frame(selection_window)
    button_frame.pack(pady=10)
    
    selected_scripts = []
    
    def on_confirm():
        nonlocal selected_scripts
        selected_scripts = [(name, script) for var, name, script in check_vars if var.get()]
        if not selected_scripts:
            messagebox.showwarning("警告", "請至少選擇一個腳本！")
            return
        # 🔥 儲存使用者選的「使用執照查詢」城市為下次預設
        chosen_license = next((s for v, n, s in check_vars
                               if s in LICENSE_QUERY_SCRIPTS and v.get()), None)
        if chosen_license:
            try:
                _cfg = load_config() or {}
                _cfg["license_query_default"] = chosen_license
                save_config(_cfg, show_message=False)
            except Exception:
                pass
        selection_window.destroy()
    
    def on_cancel():
        nonlocal selected_scripts
        selected_scripts = []
        selection_window.destroy()
    
    def select_all():
        # 🔥 license 群組是互斥的，跳過讓使用者原本選的城市保留
        for var, _, script in check_vars:
            if script in LICENSE_QUERY_SCRIPTS:
                continue
            var.set(True)

    def deselect_all():
        for var, _, _ in check_vars:
            var.set(False)

    def select_buy():
        # 買賣：全選除了"全功能地籍資料查詢"、"高雄市土增稅估算"、"全國土地使用分區"
        # 🔥 license 群組是互斥的，跳過保留使用者選擇
        for var, name, script in check_vars:
            if script in LICENSE_QUERY_SCRIPTS:
                continue
            if ("全功能地籍資料查詢" not in name and
                "高雄市土增稅估算" not in name and
                "全國土地使用分區" not in name):
                var.set(True)
            else:
                var.set(False)

    def select_rent():
        # 租賃：全選除了"全功能地籍資料查詢"、"土地增值稅"、"電子謄本結構化"、"讀取結構化JSON"（包含編輯物件資料）
        # 🔥 license 群組是互斥的，跳過保留使用者選擇
        for var, name, script in check_vars:
            if script in LICENSE_QUERY_SCRIPTS:
                continue
            if ("全功能地籍資料查詢" not in name and
                "土地增值稅" not in name and
                "電子謄本結構化" not in name and
                "讀取結構化JSON" not in name):
                var.set(True)
            else:
                var.set(False)
    
    # 快速選擇按鈕
    quick_frame = tk.Frame(button_frame)
    quick_frame.pack(pady=5)

    # 🔥 按鈕字體根據 DPI 調整
    btn_font_size = max(9, int(12 / dpi_scale * 1.2))
    btn_width = max(6, int(8 / dpi_scale * 1.2))

    tk.Button(quick_frame, text="買 賣", command=select_buy,
            font=("Microsoft JhengHei", btn_font_size), bg='#90EE90', width=btn_width).pack(side=tk.LEFT, padx=5)
    tk.Button(quick_frame, text="租 賃", command=select_rent,
            font=("Microsoft JhengHei", btn_font_size), bg='#FFB6C1', width=btn_width).pack(side=tk.LEFT, padx=5)
    tk.Button(quick_frame, text="全 選", command=select_all,
            font=("Microsoft JhengHei", btn_font_size), bg='#E0E0E0', width=btn_width).pack(side=tk.LEFT, padx=5)
    tk.Button(quick_frame, text="全不選", command=deselect_all,
            font=("Microsoft JhengHei", btn_font_size), bg='#E0E0E0', width=btn_width).pack(side=tk.LEFT, padx=5)

    # 分隔線
    separator = tk.Frame(button_frame, height=2, bg='gray')
    separator.pack(fill=tk.X, pady=10)

    # 確定取消按鈕
    action_frame = tk.Frame(button_frame)
    action_frame.pack()

    action_btn_width = max(8, int(10 / dpi_scale * 1.2))
    tk.Button(action_frame, text="確定", command=on_confirm,
            font=("Microsoft JhengHei", btn_font_size), bg='#4CAF50', fg='white', width=action_btn_width).pack(side=tk.LEFT, padx=10)
    tk.Button(action_frame, text="取消", command=on_cancel,
            font=("Microsoft JhengHei", btn_font_size), bg='#F44336', fg='white', width=action_btn_width).pack(side=tk.LEFT, padx=10)
    
    # 等待對話框關閉
    root.wait_window(selection_window)
    
    return selected_scripts

def start_continuous_mode():
    """開始連續模式"""
    global continuous_mode_active, continuous_mode_queue, continuous_mode_index
    
    if is_running:
        update_message("有程式正在運行中，無法啟動連續模式")
        return
    
    # ⭐ 立即顯示注意事項（移到最前面）
    update_divider()
    update_message("【連續模式】")
    update_divider()
    update_message("注意事項：")
    update_message("① 程式將依選擇按順序執行")
    update_message("② 部分程式過程中仍需與使用者配合操作")
    update_message("③ 請注意輸入相關資訊的提示")
    update_message("④ 每個程式完成後會自動等待5秒執行下一個")
    update_divider()
    update_message("請在彈出的視窗中選擇要執行的腳本...")
    update_message("")
    
    # 強制立即更新 UI
    root.update_idletasks()
    root.update()
    
    # 選擇要執行的腳本
    selected_scripts = select_continuous_mode_scripts()
    
    if not selected_scripts:
        update_message("連續模式已取消")
        return
    
    continuous_mode_active = True
    continuous_mode_queue = selected_scripts
    continuous_mode_index = 0
    
    # 更新按鈕狀態
    continuous_mode_button.config(text="停止連續模式", bg='#FF6B6B', fg='white')
    
    # 禁用其他按鈕
    disable_other_buttons(continuous_mode_button)
    
    # ⭐ 確認後才顯示執行清單
    update_message("連續模式開始！")
    update_message(f"將執行 {len(continuous_mode_queue)} 個腳本：")
    for i, (name, _) in enumerate(continuous_mode_queue, 1):
        update_message(f"   {i}. {name}")
    update_message("")
    
    # 開始執行第一個腳本
    execute_next_script()

def stop_continuous_mode():
    """停止連續模式"""
    global continuous_mode_active, continuous_mode_timer, current_running_button

    continuous_mode_active = False

    if continuous_mode_timer:
        root.after_cancel(continuous_mode_timer)
        continuous_mode_timer = None

    # 🔥 恢復所有按鈕狀態，包括正在執行的按鈕
    if current_running_button:
        original_text = current_running_button.cget("text").replace("關閉 ", "")
        current_running_button.config(text=original_text)
        current_running_button = None

    # 恢復按鈕狀態
    continuous_mode_button.config(text="連續模式", **continuous_button_style)
    enable_all_buttons()

    update_message("連續模式已停止")

    # 開啟建立的資料夾
    open_created_folder()

def execute_next_script():
    """執行下一個腳本"""
    global continuous_mode_index
    
    if not continuous_mode_active or continuous_mode_index >= len(continuous_mode_queue):
        # 連續模式結束
        update_message("連續模式執行完成！")
        stop_continuous_mode()
        return
    
    name, script = continuous_mode_queue[continuous_mode_index]
    update_message(f"正在執行第 {continuous_mode_index + 1}/{len(continuous_mode_queue)} 個：{name}")
    
    # 執行腳本
    start_subprocess_internal(script)
    continuous_mode_index += 1

def schedule_next_script():
    """安排下一個腳本執行"""
    global continuous_mode_timer
    
    if not continuous_mode_active:
        return
    
    if continuous_mode_index >= len(continuous_mode_queue):
        stop_continuous_mode()
        return
    
    # 5秒後執行下一個腳本
    update_message("5秒後執行下一個腳本...")
    continuous_mode_timer = root.after(5000, execute_next_script)

def load_data_json_from_folder():
    """從資料夾的 4.其他相關 載入 data.json 到程式目錄"""
    global json_status_label  # 🔥 宣告為全域變數，才能更新狀態列
    try:
        # 使用檔案對話框選擇 data.json 檔案
        from tkinter import filedialog
        import glob

        # 🔥 預設開啟程式目錄，方便找到其他地段資料夾
        initial_dir = os.getcwd()

        data_json_path = filedialog.askopenfilename(
            title="選擇案件的 data.json 檔案（在 4.其他相關 資料夾內）",
            initialdir=initial_dir,
            filetypes=[("JSON files", "data.json"), ("All files", "*.*")]
        )

        if not data_json_path:
            update_message("[取消] 未選擇檔案")
            return

        if not os.path.exists(data_json_path):
            show_large_message(
                "找不到檔案",
                f"選擇的檔案不存在"
            )
            return

        # 讀取並複製到程式目錄
        import shutil
        shutil.copy2(data_json_path, "data.json")

        # 讀取第一筆資料顯示資訊
        with open("data.json", 'r', encoding='utf-8') as f:
            data_list = json.load(f)
            if data_list:
                first = data_list[0]
                info = f"{first.get('area', '')}{first.get('section', '')} {first.get('lot_number', '')}"
                update_message(f"[成功] 已載入 data.json：{info}")
                update_message(f"  共 {len(data_list)} 筆資料")

                # 🔥 先更新狀態列顯示（在顯示對話框之前）
                lot_display = simplify_lot_number(f"{first.get('area', '')}{first.get('section', '')}-{first.get('lot_number', '')}")
                json_status_label.config(
                    text=f"JSON: {lot_display}",
                    fg='#006400'
                )
                # 🔥 強制更新 UI，確保狀態列立即顯示
                root.update_idletasks()

                # 🔥 更新主視窗標題
                update_title_from_data_json()

                # 🔥 最後才顯示對話框（這樣狀態列已經更新完成）
                show_large_message(
                    "載入成功",
                    f"已從資料夾載入 data.json\n\n案件：{info}\n筆數：{len(data_list)}"
                )
            else:
                update_message("[警告] data.json 中沒有資料")

    except Exception as e:
        update_message(f"[錯誤] 載入 data.json 失敗：{e}")
        show_large_message("載入失敗", f"無法載入 data.json\n\n錯誤：{e}")

def open_created_folder():
    """開啟建立的資料夾，並確保標準子目錄結構存在"""
    try:
        # 讀取 data.json 的第一筆資料來確定資料夾名稱
        if os.path.exists('data.json'):
            with open('data.json', 'r', encoding='utf-8') as f:
                data_list = json.load(f)
                if data_list:
                    first_data = data_list[0]
                    area = first_data.get('area', '')
                    section = first_data.get('section', '')
                    lot_number = first_data.get('lot_number', '')
                    folder_name = f"{area}{section}-{lot_number}"
                    
                    # 建立標準子目錄結構
                    subdirectories = [
                        "0.謄本",
                        "1.基本資料",
                        "2.照片", 
                        "3.行情",
                        "4.其他相關",
                        "5.委託",
                        "6.成交"
                    ]
                    
                    # 確保主資料夾存在
                    if not os.path.exists(folder_name):
                        os.makedirs(folder_name)
                        update_message(f"建立主資料夾：{folder_name}")
                    
                    # 建立子目錄
                    created_dirs = []
                    for subdir in subdirectories:
                        subdir_path = os.path.join(folder_name, subdir)
                        if not os.path.exists(subdir_path):
                            os.makedirs(subdir_path)
                            created_dirs.append(subdir)
                    
                    if created_dirs:
                        update_message(f"建立子目錄：{', '.join(created_dirs)}")
                    else:
                        update_message("所有子目錄已存在")
                    
                    # 在1.基本資料下建立png子目錄（如果不存在）
                    png_dir = os.path.join(folder_name, "1.基本資料", "png")
                    if not os.path.exists(png_dir):
                        os.makedirs(png_dir)
                        update_message("建立png子目錄")
                    
                    # 開啟資料夾
                    if os.path.exists(folder_name):
                        os.startfile(folder_name)
                        update_message(f"已開啟資料夾：{folder_name}")
                    else:
                        update_message(f"找不到資料夾：{folder_name}")
                        
    except Exception as e:
        update_message(f"開啟資料夾時發生錯誤：{e}")

def import_archive():
    """匯入封存資料夾或ZIP檔案"""
    try:
        update_divider()
        update_message("開始匯入封存...")

        # 選擇檔案或資料夾
        from tkinter import filedialog
        import zipfile
        import shutil
        import re

        # 建立選擇對話框
        choice_dialog = tk.Toplevel(root)
        choice_dialog.title("選擇匯入來源")
        choice_dialog.geometry("400x320")  # 🔥 增加高度避免按鈕被壓扁
        choice_dialog.transient(root)
        choice_dialog.grab_set()

        # 置中顯示
        screen_width = choice_dialog.winfo_screenwidth()
        screen_height = choice_dialog.winfo_screenheight()
        x = (screen_width - 400) // 2
        y = (screen_height - 320) // 2
        choice_dialog.geometry(f"400x320+{x}+{y}")

        tk.Label(choice_dialog, text="請選擇匯入來源類型：",
                font=("Microsoft JhengHei", 14, "bold")).pack(pady=20)

        import_type = tk.StringVar(value="folder")

        tk.Radiobutton(choice_dialog, text="封存資料夾",
                      variable=import_type, value="folder",
                      font=("Microsoft JhengHei", 12)).pack(pady=10, anchor="w", padx=50)

        tk.Radiobutton(choice_dialog, text="ZIP 壓縮檔",
                      variable=import_type, value="zip",
                      font=("Microsoft JhengHei", 12)).pack(pady=10, anchor="w", padx=50)

        btn_frame = tk.Frame(choice_dialog)
        btn_frame.pack(pady=20)

        choice_selected = [False]

        def on_confirm():
            choice_selected[0] = True
            choice_dialog.destroy()

        def on_cancel():
            choice_dialog.destroy()

        tk.Button(btn_frame, text="確定", command=on_confirm,
                 font=("Microsoft JhengHei", 12, "bold"),
                 bg='#4CAF50', fg='white', width=10, height=2).pack(side=tk.LEFT, padx=10)

        tk.Button(btn_frame, text="取消", command=on_cancel,
                 font=("Microsoft JhengHei", 12),
                 bg='#F44336', fg='white', width=10, height=2).pack(side=tk.LEFT, padx=10)

        choice_dialog.wait_window()

        if not choice_selected[0]:
            update_message("[取消] 使用者取消選擇")
            return

        selected_type = import_type.get()
        source_path = None
        temp_extract_path = None

        # 根據類型選擇檔案或資料夾
        if selected_type == "zip":
            source_path = filedialog.askopenfilename(
                title="選擇 ZIP 封存檔案",
                filetypes=[("ZIP 檔案", "*.zip"), ("所有檔案", "*.*")]
            )
            if not source_path:
                update_message("[取消] 未選擇檔案")
                return

            update_message(f"選擇的 ZIP：{os.path.basename(source_path)}")

            # 解壓縮到臨時目錄
            temp_extract_path = os.path.join(os.getcwd(), "_temp_extract")
            if os.path.exists(temp_extract_path):
                shutil.rmtree(temp_extract_path)
            os.makedirs(temp_extract_path)

            update_message("正在解壓縮...")

            # 建立進度條視窗
            progress_window = tk.Toplevel(root)
            progress_window.title("匯入進度")

            # 🔥 根據 DPI 調整視窗大小
            prog_width = int(500 / dpi_scale * 1.2)
            prog_height = int(150 / dpi_scale * 1.2)
            progress_window.geometry(f"{prog_width}x{prog_height}")
            progress_window.transient(root)
            progress_window.grab_set()

            # 置中顯示
            screen_width = progress_window.winfo_screenwidth()
            screen_height = progress_window.winfo_screenheight()
            x = (screen_width - prog_width) // 2
            y = (screen_height - prog_height) // 2
            progress_window.geometry(f"{prog_width}x{prog_height}+{x}+{y}")

            # 🔥 根據 DPI 調整字體
            prog_font_size = max(10, int(12 / dpi_scale * 1.2))
            progress_label = tk.Label(progress_window, text="正在解壓縮...",
                                     font=("Microsoft JhengHei", prog_font_size))
            progress_label.pack(pady=15)

            progress_bar = ttk.Progressbar(progress_window, length=int(400 / dpi_scale * 1.2), mode='determinate')
            progress_bar.pack(pady=10)

            detail_font_size = max(8, int(9 / dpi_scale * 1.2))
            progress_detail = tk.Label(progress_window, text="",
                                      font=("Microsoft JhengHei", detail_font_size), fg="gray")
            progress_detail.pack(pady=5)

            progress_window.update()

            # 解壓縮 ZIP 檔案並顯示進度
            with zipfile.ZipFile(source_path, 'r') as zip_ref:
                file_list = zip_ref.namelist()
                total_files = len(file_list)
                progress_bar['maximum'] = total_files

                for idx, file in enumerate(file_list, 1):
                    zip_ref.extract(file, temp_extract_path)
                    progress_bar['value'] = idx
                    progress_detail.config(text=f"{idx}/{total_files} 個檔案")
                    if idx % 10 == 0 or idx == total_files:
                        progress_window.update()

            progress_window.destroy()

            # 找到解壓縮後的資料夾
            extracted_items = os.listdir(temp_extract_path)
            if len(extracted_items) == 1 and os.path.isdir(os.path.join(temp_extract_path, extracted_items[0])):
                source_folder = os.path.join(temp_extract_path, extracted_items[0])
            else:
                source_folder = temp_extract_path

            archive_name = os.path.basename(source_folder)

        else:  # folder
            source_folder = filedialog.askdirectory(title="選擇封存資料夾")
            if not source_folder:
                update_message("[取消] 未選擇資料夾")
                return

            archive_name = os.path.basename(source_folder)
            update_message(f"選擇的資料夾：{archive_name}")

        # 從封存名稱提取原始資料夾名稱
        # 格式：114.10.27左營區福山段-539,左營區重惠街106號7樓 → 左營區福山段-539
        # 或：114.10.27左營區福山段-539 → 左營區福山段-539

        # 移除日期部分（民國年.月.日）
        pattern = r'^\d+\.\d+\.\d+(.+?)(?:,.*)?$'
        match = re.match(pattern, archive_name)

        if match:
            original_name = match.group(1)
        else:
            # 如果格式不符，直接使用原名稱
            original_name = archive_name

        update_message(f"原始資料夾名稱：{original_name}")

        # 目標路徑
        dest_folder = os.path.join(os.getcwd(), original_name)

        # 檢查目標是否已存在
        if os.path.exists(dest_folder):
            result = show_large_yesno(
                "資料夾已存在",
                f"工作區已存在相同名稱的資料夾：\n{original_name}\n\n是否要覆蓋？\n\n（原有資料將被刪除）"
            )

            if not result:
                update_message("[取消] 使用者取消覆蓋")
                # 清理臨時目錄
                if temp_extract_path and os.path.exists(temp_extract_path):
                    shutil.rmtree(temp_extract_path)
                return

            # 刪除舊資料夾
            update_message("正在刪除舊資料夾...")
            try:
                shutil.rmtree(dest_folder)
                update_message("[刪除] 已刪除舊資料夾")
            except Exception as e:
                update_message(f"[錯誤] 刪除失敗：{e}")
                show_large_message("錯誤", f"無法刪除舊資料夾，可能被其他程式占用：\n\n{e}")
                # 清理臨時目錄
                if temp_extract_path and os.path.exists(temp_extract_path):
                    shutil.rmtree(temp_extract_path)
                return

        # 複製資料夾到工作區
        update_message("正在複製資料夾到工作區...")

        # 建立進度條視窗（如果是從 ZIP 解壓，進度視窗已經關閉，需要重新建立）
        progress_window = tk.Toplevel(root)
        progress_window.title("匯入進度")

        # 🔥 根據 DPI 調整視窗大小
        prog_width = int(500 / dpi_scale * 1.2)
        prog_height = int(150 / dpi_scale * 1.2)
        progress_window.geometry(f"{prog_width}x{prog_height}")
        progress_window.transient(root)
        progress_window.grab_set()

        # 置中顯示
        screen_width = progress_window.winfo_screenwidth()
        screen_height = progress_window.winfo_screenheight()
        x = (screen_width - prog_width) // 2
        y = (screen_height - prog_height) // 2
        progress_window.geometry(f"{prog_width}x{prog_height}+{x}+{y}")

        # 🔥 根據 DPI 調整字體
        prog_font_size = max(10, int(12 / dpi_scale * 1.2))
        progress_label = tk.Label(progress_window, text="正在複製資料夾...",
                                 font=("Microsoft JhengHei", prog_font_size))
        progress_label.pack(pady=15)

        progress_bar = ttk.Progressbar(progress_window, length=int(400 / dpi_scale * 1.2), mode='determinate')
        progress_bar.pack(pady=10)

        detail_font_size = max(8, int(9 / dpi_scale * 1.2))
        progress_detail = tk.Label(progress_window, text="",
                                  font=("Microsoft JhengHei", detail_font_size), fg="gray")
        progress_detail.pack(pady=5)

        progress_window.update()

        # 計算總檔案數
        total_files = sum([len(files) for _, _, files in os.walk(source_folder)])
        progress_bar['maximum'] = total_files

        # 使用自訂複製函數來顯示進度
        copied_files = [0]

        def copy_with_progress(src, dst, *, follow_symlinks=True):
            shutil.copy2(src, dst, follow_symlinks=follow_symlinks)
            copied_files[0] += 1
            progress_bar['value'] = copied_files[0]
            progress_detail.config(text=f"{copied_files[0]}/{total_files} 個檔案")
            if copied_files[0] % 10 == 0 or copied_files[0] == total_files:
                progress_window.update()

        shutil.copytree(source_folder, dest_folder, copy_function=copy_with_progress)

        progress_window.destroy()
        update_message(f"[完成] 已複製到：{dest_folder}")

        # 載入 data.json
        data_json_path = os.path.join(dest_folder, "4.其他相關", "data.json")
        if os.path.exists(data_json_path):
            shutil.copy2(data_json_path, "data.json")
            update_message("[完成] 已載入 data.json")

            # 讀取並顯示資訊
            with open('data.json', 'r', encoding='utf-8') as f:
                data_list = json.load(f)
                if data_list:
                    first_data = data_list[0]
                    area = first_data.get('area', '')
                    section = first_data.get('section', '')
                    lot_number = first_data.get('lot_number', '')
                    update_message(f"物件資訊：{area}{section}-{lot_number}")
        else:
            update_message("[警告] 未找到 data.json")

        # 驗證並顯示 data_final.json 的資訊
        data_final_path = os.path.join(dest_folder, "4.其他相關", "data_final.json")
        if os.path.exists(data_final_path):
            try:
                with open(data_final_path, 'r', encoding='utf-8') as f:
                    data_final = json.load(f)
                update_message("[完成] 已載入物件資料 (data_final.json)")

                # 顯示重要資訊
                transaction_type = data_final.get('交易類型', '未設定')
                building_address = data_final.get('建物門牌', '')
                land_mark = data_final.get('土地標示', '')

                update_message(f"  交易類型：{transaction_type}")
                if building_address:
                    update_message(f"  建物門牌：{building_address}")
                if land_mark:
                    update_message(f"  土地標示：{land_mark}")

                # 根據交易類型顯示價格資訊
                if transaction_type == "租件":
                    rent = data_final.get('租金', '')
                    if rent:
                        update_message(f"  租金：{rent} 萬元/月")
                else:  # 售件
                    total_price = data_final.get('總價', '')
                    if total_price:
                        update_message(f"  總價：{total_price} 萬元")

            except Exception as e:
                update_message(f"[警告] 讀取 data_final.json 時發生錯誤：{e}")
        else:
            update_message("[警告] 未找到 data_final.json")
            update_message("  物件資料可能需要重新建立")

        # 清理臨時目錄
        if temp_extract_path and os.path.exists(temp_extract_path):
            shutil.rmtree(temp_extract_path)
            update_message("[清理] 已刪除臨時檔案")

        update_divider()

        # 準備詳細的匯入完成訊息
        completion_msg = f"封存已成功匯入！\n\n資料夾：{original_name}\n\n"

        # 加入物件資訊（如果有）
        if os.path.exists(data_final_path):
            try:
                with open(data_final_path, 'r', encoding='utf-8') as f:
                    data_final = json.load(f)
                transaction_type = data_final.get('交易類型', '未設定')
                completion_msg += f"交易類型：{transaction_type}\n"

                building_address = data_final.get('建物門牌', '')
                if building_address:
                    completion_msg += f"建物門牌：{building_address}\n"

                completion_msg += "\n"
            except:
                pass

        completion_msg += "現在可以開始編輯物件資料。"

        show_large_message("匯入完成", completion_msg)

        # 🔥 重要：更新 selected_json_path，讓狀態列顯示正確
        global selected_json_path
        selected_json_path = data_json_path  # 更新為匯入的 JSON 路徑

        # 更新設定檔，儲存新的路徑
        config = load_config()
        config['selected_json_path'] = selected_json_path
        save_config(config)

        # 更新狀態列顯示
        update_json_status_display()

    except Exception as e:
        update_message(f"[錯誤] 匯入封存失敗：{e}")
        import traceback
        update_message(traceback.format_exc())
        show_large_message("錯誤", f"匯入封存失敗：\n\n{e}")

def process_mortgage_data(mortgages, data_final):
    """處理貸款資訊（他項權利）

    Args:
        mortgages: 他項權利列表
        data_final: 要更新的資料字典
    """
    if not mortgages:
        return

    # 🔥 銀行關鍵字對照表（模糊比對用）
    bank_keywords = {
        '臺灣銀行': '台銀',
        '台灣銀行': '台銀',
        '合作金庫': '合庫',
        '第一商業銀行': '一銀',
        '第一銀行': '一銀',
        '華南商業銀行': '華銀',
        '華南銀行': '華銀',
        '彰化商業銀行': '彰銀',
        '彰化銀行': '彰銀',
        '台北富邦': '富邦',
        '富邦商業銀行': '富邦',
        '富邦銀行': '富邦',
        '國泰世華商業銀行': '國泰',
        '國泰世華': '國泰',
        '中國信託商業銀行': '中信',
        '中國信託': '中信',
        '玉山商業銀行': '玉山',
        '玉山銀行': '玉山',
        '台新國際商業銀行': '台新',
        '台新銀行': '台新',
        '兆豐國際商業銀行': '兆豐',
        '兆豐銀行': '兆豐',
        '土地銀行': '土銀',
        '郵局': '郵局',
        '中華郵政': '郵局',
    }

    # 🔥 先去除共同擔保的重複項（使用證明書字號去重）
    seen_certificates = {}  # {證明書字號: mortgage}
    unique_mortgages = []

    for mortgage in mortgages:
        cert_no = mortgage.get("證明書字號", "")
        if cert_no and cert_no != "（空白）":
            # 有證明書字號，用它來判斷是否為共同擔保
            if cert_no not in seen_certificates:
                seen_certificates[cert_no] = mortgage
                unique_mortgages.append(mortgage)
            # 如果證明書字號已存在，跳過（共同擔保，不重複計算）
        else:
            # 沒有證明書字號，保留（可能是獨立的抵押權）
            unique_mortgages.append(mortgage)

    # 收集銀行和金額
    bank_dict = {}  # {簡稱: [金額1, 金額2, ...]}

    for mortgage in unique_mortgages:
        bank_name = mortgage.get("權利人", "")
        amount_str = mortgage.get("擔保債權總金額", "")

        # 🔥 模糊比對：只要包含關鍵字就匹配
        simplified_bank = None
        for keyword, short_name in bank_keywords.items():
            if keyword in bank_name:
                simplified_bank = short_name
                break

        # 🔥 如果找不到對照，執行簡化規則
        if not simplified_bank:
            # 判斷機構類型
            if '農會' in bank_name or '漁會' in bank_name or '信用合作社' in bank_name:
                # 農會、漁會、信合社：保留完整名稱
                cleaned_name = bank_name.replace('股份有限公司', '')
                simplified_bank = cleaned_name
            elif '銀行' in bank_name:
                # 一般銀行：移除後綴，保留主要名稱
                cleaned_name = bank_name.replace('股份有限公司', '').replace('商業銀行', '').replace('銀行', '')
                # 如果清理後還有內容，加上「銀行」；否則直接使用原名
                if cleaned_name and len(cleaned_name) >= 2:
                    simplified_bank = cleaned_name + '銀行'
                else:
                    simplified_bank = bank_name.replace('股份有限公司', '')
            else:
                # 其他情況（可能是自然人或其他機構）：移除公司後綴
                simplified_bank = bank_name.replace('股份有限公司', '')
        
        # 提取金額數字
        if amount_str:
            import re
            numbers = re.findall(r'[\d,]+', amount_str)
            if numbers:
                amount = int(numbers[0].replace(',', ''))
                
                if simplified_bank not in bank_dict:
                    bank_dict[simplified_bank] = []
                bank_dict[simplified_bank].append(amount)
    
    # 匯整結果
    if bank_dict:
        # 貸款銀行：去重後列出
        unique_banks = list(bank_dict.keys())
        data_final["貸款銀行"] = "、".join(unique_banks)
        
        # 設定金額：轉換為萬為單位
        bank_totals = {}
        for bank, amounts in bank_dict.items():
            total_yuan = sum(amounts)
            total_wan = total_yuan / 10000
            bank_totals[bank] = total_wan
        
        # 總計（萬）
        grand_total_wan = sum(bank_totals.values())

        # 🔥 格式：先顯示總計，有多筆才列明細
        if len(bank_totals) == 1:
            # 只有一筆，直接顯示總計（加上「萬元正」）
            data_final["設定金額"] = f"{grand_total_wan:,.0f}萬元正"
        else:
            # 多筆：總計 + 明細
            details = "、".join([f"{bank}{amount:,.0f}萬" for bank, amount in bank_totals.items()])
            data_final["設定金額"] = f"{grand_total_wan:,.0f}萬元正"
            data_final["設定金額明細"] = details

def convert_transcript_to_final_data(transcript_json_path):
    """
    將結構化 JSON 轉換成 data_final.json

    Args:
        transcript_json_path: 結構化 JSON 檔案路徑
    """
    global selected_json_path  # 🔥 加入全域變數宣告

    try:
        update_divider()
        update_message("開始轉換結構化 JSON...")
        
        # 讀取結構化 JSON
        with open(transcript_json_path, 'r', encoding='utf-8') as f:
            transcript_data = json.load(f)
        
        if not transcript_data:
            update_message("[錯誤] 結構化 JSON 為空")
            return

        # 🔥 先分離土地謄本和建物謄本（支援第一類和第二類）
        land_transcripts = [t for t in transcript_data if "土地謄本" in t.get("謄本類型", "")]
        building_transcripts = [t for t in transcript_data if "建物謄本" in t.get("謄本類型", "")]

        update_message(f"找到 {len(land_transcripts)} 筆土地謄本")
        update_message(f"找到 {len(building_transcripts)} 筆建物謄本")

        # 🔥 從 data.json 讀取縣市資訊（需驗證是否同一物件）
        city_name = None
        city_source = ""
        is_ambiguous = False  # 標記是否為重複區名

        try:
            if os.path.exists('data.json'):
                with open('data.json', 'r', encoding='utf-8') as f:
                    data_list = json.load(f)

                if data_list and len(data_list) > 0:
                    first_data = data_list[0]

                    # 取得 data.json 的地段資訊（用於 fallback 訊息顯示）
                    data_json_section = first_data.get('section', '')
                    data_json_lot = first_data.get('lot_number', '')
                    data_json_city = first_data.get('city', '')

                    # 🔥 比對地段地號（任一筆 data.json × 任一筆謄本相符即驗證通過）
                    # 修正：之前只比對首筆對首筆，順序不同就誤判「不符」
                    import re as _re

                    def _norm_lot(s):
                        """正規化地號：'193' or '0193-0000' or '193-0' → '193-0'"""
                        s = str(s or '').strip().replace('地號', '')
                        try:
                            if '-' in s:
                                a, b = s.split('-', 1)
                                return f"{int(a)}-{int(b)}"
                            if len(s) >= 8:
                                return f"{int(s[:4])}-{int(s[4:])}"
                            return f"{int(s)}-0"
                        except (ValueError, IndexError):
                            return s

                    match_info = None  # (data_section, data_lot)
                    if land_transcripts:
                        for d in data_list:
                            d_section = (d.get('section', '') or '').strip()
                            d_lot_raw = (d.get('lot_number', '') or '').strip()
                            if not (d_section and d_lot_raw):
                                continue
                            d_lot_norm = _norm_lot(d_lot_raw)
                            for t in land_transcripts:
                                t_地號 = t.get("基本資訊", {}).get("地號建號", "")
                                m = _re.search(r'([^\s]+段)\s*(\d+-?\d*)', t_地號)
                                if not m:
                                    continue
                                if d_section == m.group(1) and d_lot_norm == _norm_lot(m.group(2)):
                                    match_info = (d_section, d_lot_raw)
                                    break
                            if match_info:
                                break

                    if match_info:
                        city_name = data_json_city
                        city_source = "data.json（已驗證相符）"
                        update_message(f"  [OK] 驗證通過：data.json 與謄本為同一物件")
                        update_message(f"    地段：{match_info[0]} | 地號：{match_info[1]} | 縣市：{city_name}")
                    elif land_transcripts:
                        update_message(f"  [警告] data.json 與謄本地號不符，不使用 data.json 的縣市")
                        update_message(f"    data.json[0]: {data_json_section} {data_json_lot}")
                        first_t = land_transcripts[0].get("基本資訊", {}).get("地號建號", "")
                        update_message(f"    謄本[0]: {first_t}")
        except Exception as e:
            update_message(f"  [警告] 讀取 data.json 失敗：{e}")

        # 🔥 如果 data.json 驗證失敗，從謄本的「區」推斷縣市
        if not city_name and land_transcripts:
            transcript_地號 = land_transcripts[0].get("基本資訊", {}).get("地號建號", "")

            # 提取行政區
            district = extract_district_from_address(transcript_地號)

            if district:
                city_name, is_ambiguous = get_city_from_district(district)

                if city_name:
                    city_source = f"從謄本推斷（{district}）"
                    if is_ambiguous:
                        city_source += " [警告]可能不準確（重複區名）"
                        update_message(f"  [警告] 從謄本推斷縣市：{city_name}（警告：{district} 存在於多個縣市）")
                    else:
                        update_message(f"  [OK] 從謄本推斷縣市：{city_name}（行政區：{district}）")

        # 🔥 fallback：從 data.json 的 city 欄位取得（即使 section/lot 驗證失敗）
        if not city_name:
            try:
                if os.path.exists('data.json'):
                    with open('data.json', 'r', encoding='utf-8') as f:
                        data_list_fb = json.load(f)
                    if data_list_fb:
                        fallback_city = data_list_fb[0].get('city', '')
                        if fallback_city:
                            city_name = fallback_city
                            city_source = "data.json（未驗證 section/lot，作為 fallback）"
                            update_message(f"  [OK] 使用 data.json city 欄位作為 fallback：{city_name}")
            except Exception:
                pass

        # 🔥 最終預設值（不再硬寫高雄市，改為空字串並警告）
        if not city_name:
            city_name = ""
            city_source = "無法推斷（請手動確認縣市）"
            update_message(f"  [警告] 無法推斷縣市，土地標示將不含縣市前綴")

        update_message(f"  📍 縣市來源：{city_source}")
        
        # 初始化 data_final
        data_final = {
            "物件編號": "",
            "契約編號": "",
            "案名": "",
            "土地標示": "",
            "建物門牌": "",
            "型式": "",
            "用途": "",
            "交易類型": "售件",
            "總價": "",
            "含車位價格": "",
            "租金": "",
            "押金": "二個月",
            "租期": "",
            "整地期": "",
            "調幅": "",
            "管理費": "",
            "格局選項": [],
            "稅賦": {
                "所得稅": "",
                "補充保費": "",
                "營業稅": "",
                "房屋稅": "",
                "地價稅": ""
            },
            "總地坪": "",
            "主建物坪數": "",
            "附屬建物坪數": "",
            "公共設施坪數": "",
            "總建坪": "",
            "謄本登記總面積": "",
            "增建面積": "",
            "竣工日期": "",
            "建築結構": "",
            "座向": "",
            "主建物": {},
            "附屬建物": {},
            "公共設施": {},
            "格局": {"房": "", "廳": "", "衛": ""},
            "土地單價": "",
            "建物單價": "",
            "公告現值": "",
            "自用增值稅": "",
            "一般增值稅": "",
            "所有權型態": "",
            "所有權人": "",
            "權利範圍": "",
            "貸款銀行": "",
            "設定金額": "",
            "現況": "",
            "鑰匙": "",
            "車位": {"有無": "", "型式": "", "編號": ""},
            "電梯": "",
            "管理費": "",
            "面臨道路": "",
            "面寬": "",
            "深度": "",
            "都市土地使用分區": "",
            "建蔽率": "",
            "容積率": "",
            "訴求重點": "",
            "承辦人": "",
            "電話": ""
        }
        
        # === 處理土地謄本 ===
        if land_transcripts:
            # 可能有多筆土地，需要整合地號
            all_land_ids = []
            total_land_area_m2 = 0
            
            # 🔥 新增：收集所有貸款資訊
            all_mortgages = []
            
            for land_data in land_transcripts:
                # 收集地號
                land_id = land_data.get("基本資訊", {}).get("地號建號", "")
                if land_id:
                    all_land_ids.append(land_id)
                
                # 🔥 累加面積（需考慮權利範圍）
                area_str = land_data.get("標示部", {}).get("面積", "")
                if area_str:
                    area_m2 = float(area_str.replace(",", "").replace("平方公尺", ""))
                    
                    # 🔥 取得該筆土地的所有權人權利範圍
                    owners = land_data.get("所有權部", [])
                    for owner in owners:
                        權利範圍 = owner.get("權利範圍", "")
                        
                        if 權利範圍:
                            # 解析權利範圍（例如：10000分之195）
                            分子 = 1
                            分母 = 1
                            
                            import re
                            權利範圍 = 權利範圍.replace("全部", "").strip()
                            match = re.search(r'(\d+)\s*分之\s*(\d+)', 權利範圍)
                            
                            if match:
                                分母 = int(match.group(1))
                                分子 = int(match.group(2))
                            
                            # 🔥 計算實際擁有面積
                            if 分母 > 0:
                                實際面積 = area_m2 * (分子 / 分母)
                                total_land_area_m2 += 實際面積
                                
                                update_message(f"  土地面積: {area_m2:.2f} 平方公尺")
                                update_message(f"  權利範圍: {權利範圍} ({分子}/{分母})")
                                update_message(f"  實際擁有: {實際面積:.2f} 平方公尺")
                        else:
                            # 沒有權利範圍，假設全部
                            total_land_area_m2 += area_m2
                
                # 🔥 收集他項權利（貸款）
                mortgages = land_data.get("他項權利部", [])
                if mortgages:
                    all_mortgages.extend(mortgages)
            
            # 🔥 整合土地標示（依 data.json 順序、簡化地號格式）
            # 例：'高雄市旗津區旗津段 859,887,446地號'（0446-0000→446、0446-0002→446-2）
            if all_land_ids:
                import re as _re

                def _norm_lot(s):
                    """正規化（比對用）：'0446-0000'→'446-0'、'446'→'446-0'"""
                    s = str(s or '').strip().replace('地號', '').replace('建號', '')
                    try:
                        if '-' in s:
                            a, b = s.split('-', 1)
                            return f"{int(a)}-{int(b)}"
                        if len(s) >= 8:
                            return f"{int(s[:4])}-{int(s[4:])}"
                        return f"{int(s)}-0"
                    except (ValueError, IndexError):
                        return s

                def _simplify_lot(s):
                    """簡化（顯示用）：'0446-0000'→'446'、'0446-0002'→'446-2'"""
                    s = str(s or '').strip()
                    try:
                        if '-' in s:
                            a, b = s.split('-', 1)
                            a_int, b_int = int(a), int(b)
                            return f"{a_int}" if b_int == 0 else f"{a_int}-{b_int}"
                        if len(s) >= 8:
                            a_int, b_int = int(s[:4]), int(s[4:])
                            return f"{a_int}" if b_int == 0 else f"{a_int}-{b_int}"
                        return str(int(s))
                    except (ValueError, IndexError):
                        return s

                def _parse_land_id(land_id):
                    """'旗津區旗津段 0446-0000地號' → ('旗津區', '旗津段', '0446-0000')"""
                    m = _re.search(r'([^\s]+(?:區|鄉|鎮|市))([^\s]+段)\s*(\d+-?\d*)', land_id)
                    return (m.group(1), m.group(2), m.group(3)) if m else (None, None, None)

                # 讀取 data.json 順序
                data_order_keys = []
                try:
                    if os.path.exists('data.json'):
                        with open('data.json', 'r', encoding='utf-8') as _f:
                            _dl = json.load(_f)
                        for _d in _dl:
                            _a = (_d.get('area', '') or '').strip()
                            _s = (_d.get('section', '') or '').strip()
                            _l = _norm_lot(_d.get('lot_number', ''))
                            if _a and _s and _l:
                                data_order_keys.append((_a, _s, _l))
                except Exception:
                    pass

                # 依 data.json 順序排序（未匹配附在後面）
                def _sort_key(land_id):
                    a, s, l = _parse_land_id(land_id)
                    if a and s and l:
                        k = (a, s, _norm_lot(l))
                        if k in data_order_keys:
                            return data_order_keys.index(k)
                    return len(data_order_keys) + 1

                ordered = sorted(all_land_ids, key=_sort_key)

                # 組合輸出
                first_a, first_s, _ = _parse_land_id(ordered[0])
                if first_a and first_s:
                    lots = []
                    for lid in ordered:
                        _, _, l = _parse_land_id(lid)
                        if l:
                            lots.append(_simplify_lot(l))
                    data_final["土地標示"] = f"{city_name}{first_a}{first_s} {','.join(lots)}地號"
                else:
                    # 無法解析，保留原始字串
                    data_final["土地標示"] = f"{city_name}{ordered[0]}"
            
            # 🔥 總地坪（已考慮權利範圍）
            if total_land_area_m2 > 0:
                total_land_ping = total_land_area_m2 * 0.3025
                data_final["總地坪"] = f"{total_land_ping:.2f}"
                update_message(f"  總實際面積: {total_land_area_m2:.2f} 平方公尺")
                update_message(f"  總地坪: {total_land_ping:.2f} 坪")
            
            # 取第一筆土地的其他資訊
            land_data = land_transcripts[0]
            
            # 公告現值
            land_value = land_data.get("標示部", {}).get("公告土地現值", "")
            if "元/平方公尺" in land_value:
                value = land_value.split("元/平方公尺")[0].split()[-1]
                data_final["公告現值"] = value.replace(",", "")
            
            # 🔥 非都市土地使用地類別
            使用分區 = land_data.get("標示部", {}).get("使用分區", "")
            使用地類別 = land_data.get("標示部", {}).get("使用地類別", "")
            
            # 檢查是否為都市土地（兩者都是「（空白）」）
            if 使用分區 == "（空白）" and 使用地類別 == "（空白）":
                # 都市土地，不填入
                data_final["非都市土地使用地類別"] = ""
                update_message("  土地類別: 都市土地")
            else:
                # 非都市土地，組合兩個欄位
                parts = []
                if 使用分區 and 使用分區 != "（空白）":
                    parts.append(使用分區)
                if 使用地類別 and 使用地類別 != "（空白）":
                    parts.append(使用地類別)
                
                if parts:
                    data_final["非都市土地使用地類別"] = " / ".join(parts)
                    update_message(f"  土地類別: 非都市土地 - {data_final['非都市土地使用地類別']}")
                else:
                    data_final["非都市土地使用地類別"] = ""

            # 所有權人資訊
            owners = land_data.get("所有權部", [])
            if owners:
                owner = owners[0]
                data_final["所有權人"] = owner.get("所有權人", "")
                data_final["權利範圍"] = owner.get("權利範圍", "")
            
            # 🔥 處理貸款資訊（優先使用土地他項）
            process_mortgage_data(all_mortgages, data_final)

        # === 處理建物謄本（多建號時：樓層加建號註記、門牌串接、總坪數加總）===
        if building_transcripts:
            import re as _re_b

            def _simplify_bld_id(full_id):
                """從 '新興區新興段四小段 03330-000建號' 提取 '3330' 或 '3330-1'"""
                m = _re_b.search(r'(\d+)-(\d+)\s*建號', full_id)
                if m:
                    a, b = int(m.group(1)), int(m.group(2))
                    return f"{a}" if b == 0 else f"{a}-{b}"
                m2 = _re_b.search(r'(\d+)', full_id)
                return str(int(m2.group(1))) if m2 else full_id

            def _split_address(addr):
                """把門牌切成 (路名前綴, 門牌號碼)。
                以最後一個「段/弄/巷/路/街/大道」標記為分界。
                例：'高雄市三民區建國路1號' → ('高雄市三民區建國路', '1號')
                    '高雄市三民區建國路1段100號5樓' → ('高雄市三民區建國路1段', '100號5樓')
                """
                best_end = -1
                for marker in ['段', '弄', '巷', '大道', '路', '街']:
                    idx = addr.rfind(marker)
                    if idx >= 0:
                        end = idx + len(marker)
                        if end > best_end:
                            best_end = end
                if best_end > 0:
                    return addr[:best_end], addr[best_end:]
                return addr, ""

            def _consolidate_addresses(addresses):
                """合併共同前綴的門牌：['建國路1號', '建國路3號'] → '建國路1號、3號'"""
                if not addresses:
                    return ""
                if len(addresses) == 1:
                    return addresses[0]
                groups = []  # [(prefix, [numparts])]，保留出現順序
                for addr in addresses:
                    prefix, numpart = _split_address(addr)
                    for grp in groups:
                        if grp[0] == prefix:
                            if numpart and numpart not in grp[1]:
                                grp[1].append(numpart)
                            break
                    else:
                        groups.append((prefix, [numpart] if numpart else []))
                parts = []
                for prefix, numparts in groups:
                    if not numparts:
                        parts.append(prefix)
                    elif len(numparts) == 1:
                        parts.append(prefix + numparts[0])
                    else:
                        parts.append(prefix + "、".join(numparts))
                return "｜".join(parts)

            multi_bld = len(building_transcripts) > 1

            all_addresses = []
            total_registered_area = 0.0
            floor_data = {}
            main_building_total = 0.0
            附屬建物明細 = {}
            附屬總坪數 = 0.0
            共有部分明細 = {}
            公設總坪數 = 0.0
            parking_numbers = []

            for building_data in building_transcripts:
                building_id_full = building_data.get("基本資訊", {}).get("地號建號", "")
                bld_num = _simplify_bld_id(building_id_full)
                標示部 = building_data.get("標示部", {}) or {}

                # === 建物門牌 ===
                address = 標示部.get("建物門牌", "")
                if address:
                    district = ""
                    for suffix in ["區", "鄉", "鎮", "市"]:
                        if suffix in building_id_full:
                            district = building_id_full.split(suffix)[0] + suffix
                            break
                    full_addr = f"{city_name}{district}{address}" if (district and city_name) else address
                    if full_addr and full_addr not in all_addresses:
                        all_addresses.append(full_addr)

                # === 謄本登記總面積（加總所有建物）===
                registered_total = 標示部.get("總面積", "")
                if registered_total:
                    try:
                        area_m2 = float(registered_total.replace(",", "").replace("平方公尺", ""))
                        total_registered_area += area_m2 * 0.3025
                    except (ValueError, TypeError):
                        pass

                # === 主建物各層 ===
                for i in range(20):
                    floor_key = "層次" if i == 0 else f"層次{i}"
                    area_key = "層次面積" if i == 0 else f"層次面積{i}"
                    floor_name = 標示部.get(floor_key, "")
                    floor_area = 標示部.get(area_key, "")
                    if floor_name and floor_area:
                        try:
                            area_m2 = float(floor_area.replace(",", "").replace("平方公尺", ""))
                            area_ping = area_m2 * 0.3025
                            key = f"{floor_name}（{bld_num}）" if multi_bld else floor_name
                            floor_data[key] = f"{area_ping:.2f}"
                            main_building_total += area_ping
                        except (ValueError, TypeError):
                            pass

                # === 附屬建物 ===
                for i in range(20):
                    type_key = "附屬建物用途" if i == 0 else f"附屬建物用途{i}"
                    area_key = "附屬建物面積" if i == 0 else f"附屬建物面積{i}"
                    附屬類型 = 標示部.get(type_key, "")
                    附屬面積 = 標示部.get(area_key, "")
                    if 附屬類型 and 附屬面積:
                        try:
                            area_m2 = float(附屬面積.replace(",", "").replace("平方公尺", ""))
                            area_ping = area_m2 * 0.3025
                            key = f"{附屬類型}（{bld_num}）" if multi_bld else 附屬類型
                            附屬建物明細[key] = f"{area_ping:.2f}"
                            附屬總坪數 += area_ping
                        except (ValueError, TypeError):
                            pass

                # === 公共設施 ===
                for i in range(20):
                    building_key = "共有部分建號" if i == 0 else f"共有部分建號{i}"
                    area_key = "共有部分面積" if i == 0 else f"共有部分面積{i}"
                    range_key = "權利範圍" if i == 0 else f"權利範圍{i}"
                    共有建號 = 標示部.get(building_key, "")
                    共有面積 = 標示部.get(area_key, "")
                    權利範圍 = 標示部.get(range_key, "")
                    if 共有面積 and 權利範圍:
                        try:
                            area_m2 = float(共有面積.replace(",", "").replace("平方公尺", ""))
                            ratio = 1.0
                            if "分之" in 權利範圍:
                                parts = 權利範圍.replace("分之", "/").split("/")
                                if len(parts) == 2:
                                    try:
                                        ratio = int(parts[1]) / int(parts[0])
                                    except Exception:
                                        pass
                            公設坪數 = area_m2 * ratio * 0.3025
                            base_key = 共有建號 if 共有建號 else "公共設施"
                            key = f"{base_key}（{bld_num}持分）" if multi_bld else base_key
                            共有部分明細[key] = {
                                "面積": f"{公設坪數:.2f}",
                                "權利範圍": 權利範圍
                            }
                            公設總坪數 += 公設坪數
                        except (ValueError, TypeError):
                            pass

                # === 停車位編號 ===
                for i in range(20):
                    key = "含停車位編號" if i == 0 else f"含停車位編號{i}"
                    parking_value = 標示部.get(key, "")
                    if parking_value:
                        match = _re_b.match(r'^\s*([A-Za-z0-9\-]+(?:號)?)', parking_value)
                        if match:
                            parking_num = match.group(1).strip().rstrip(',').strip()
                            if parking_num and parking_num not in parking_numbers:
                                parking_numbers.append(parking_num)
                                update_message(f"  找到停車位編號: {parking_num}（建號 {bld_num}）")

            # 元資料（取第一筆建物）：竣工日期、建築結構、地上層數
            primary_building = building_transcripts[0]
            primary_標示部 = primary_building.get("標示部", {}) or {}
            data_final["竣工日期"] = primary_標示部.get("建築完成日期", "")
            data_final["建築結構"] = primary_標示部.get("主要建材", "")
            層數 = primary_標示部.get("層數", "")
            if 層數:
                m = _re_b.search(r'(\d+)', 層數)
                if m:
                    data_final["地上層數"] = str(int(m.group(1)))
            data_final["地下層數"] = ""

            # 寫入彙整後的資料（門牌合併共同前綴）
            data_final["建物門牌"] = _consolidate_addresses(all_addresses)
            if all_addresses:
                update_message(f"  建物門牌: {data_final['建物門牌']}")
            if total_registered_area:
                data_final["謄本登記總面積"] = f"{total_registered_area:.2f}"
            data_final["主建物"] = floor_data
            data_final["主建物坪數"] = f"{main_building_total:.2f}"
            data_final["附屬建物"] = 附屬建物明細
            data_final["附屬建物坪數"] = f"{附屬總坪數:.2f}"
            data_final["公共設施"] = 共有部分明細
            data_final["公共設施坪數"] = f"{公設總坪數:.2f}"
            data_final["總建坪"] = f"{main_building_total + 附屬總坪數 + 公設總坪數:.2f}"

            if parking_numbers:
                data_final["車位"]["編號"] = "、".join(parking_numbers)
                if not data_final["車位"]["有無"]:
                    data_final["車位"]["有無"] = "有，公設內"
                update_message(f"  車位編號（共 {len(parking_numbers)} 筆）: {data_final['車位']['編號']}")

            if multi_bld:
                update_message(f"  [建物彙整] 共 {len(building_transcripts)} 筆建號，樓層/附屬已加上建號註記")

        # 🔥 如果土地謄本沒有貸款資訊，才使用建物他項（用第一個有貸款的建物）
        if not data_final.get("貸款銀行") and building_transcripts:
            for _bld in building_transcripts:
                building_mortgages = _bld.get("他項權利部", [])
                if building_mortgages:
                    process_mortgage_data(building_mortgages, data_final)
                    break

        # 儲存 data_final.json
        data_final_path = get_data_final_path()
        with open(data_final_path, 'w', encoding='utf-8') as f:
            json.dump(data_final, f, ensure_ascii=False, indent=2)

        update_message(f"[完成] 已成功轉換並儲存到 {data_final_path}")

        # 🔥 更新 selected_json_path 全域變數並儲存到設定檔
        selected_json_path = transcript_json_path

        # 🔥 儲存到設定檔，以便下次啟動時記住（不顯示單獨的彈窗）
        try:
            config = load_config() or {}
            config['selected_json_path'] = transcript_json_path
            save_config(config, show_message=False)  # 🔥 不顯示單獨的「配置已保存」彈窗
            update_message(f"[OK] 已更新並儲存當前選擇的JSON檔案：{os.path.basename(transcript_json_path)}")
        except Exception as e:
            update_message(f"[警告] 儲存設定檔失敗：{e}")

        update_divider()
        update_message("[資訊] 建物面積分析：")
        update_message(f"  • 主建物: {data_final['主建物坪數']} 坪")
        update_message(f"  • 附屬建物: {data_final['附屬建物坪數']} 坪")
        update_message(f"  • 公共設施: {data_final['公共設施坪數']} 坪")
        update_message(f"  • 實際總建坪: {data_final['總建坪']} 坪")
        update_message(f"  • 謄本登記總面積: {data_final['謄本登記總面積']} 坪（參考）")
        update_divider()
        update_message("📋 已提取的基本資訊：")
        update_message(f"  • 土地標示: {data_final['土地標示']}")
        update_message(f"  • 建物門牌: {data_final['建物門牌']}")
        update_message(f"  • 總地坪: {data_final['總地坪']} 坪")
        update_message(f"  • 所有權人: {data_final['所有權人']}")
        update_message(f"  • 貸款銀行: {data_final['貸款銀行']}")
        update_divider()
        
    except Exception as e:
        update_message(f"[錯誤] 轉換失敗: {e}")
        import traceback
        update_message(traceback.format_exc())

def toggle_continuous_mode():
    """切換連續模式"""
    if continuous_mode_active:
        stop_continuous_mode()
    else:
        start_continuous_mode()

def handle_read_json_function():
    """處理讀取結構化JSON功能"""
    global selected_json_path, json_dir, continuous_mode_active

    update_message("執行讀取結構化JSON功能...")

    try:
        # 🔥 修正：使用絕對路徑指向 output 目錄
        if getattr(sys, 'frozen', False):
            # 打包環境：使用執行檔所在目錄
            base_dir = os.path.dirname(sys.executable)
        else:
            # 開發環境：使用腳本所在目錄
            base_dir = os.path.dirname(os.path.abspath(__file__))

        output_dir = os.path.join(base_dir, "output")
        latest_json = None
        latest_time = 0
        
        if os.path.exists(output_dir):
            for file in os.listdir(output_dir):
                if file.endswith('.json'):
                    file_path = os.path.join(output_dir, file)
                    file_time = os.path.getmtime(file_path)
                    if file_time > latest_time:
                        latest_time = file_time
                        latest_json = file_path
        
        if latest_json:
            # 顯示檔案資訊
            file_name = os.path.basename(latest_json)
            file_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(latest_time))
            file_size = os.path.getsize(latest_json)
            
            update_message(f"找到最新結構化JSON檔案：")
            update_message(f"  檔案名稱：{file_name}")
            update_message(f"  修改時間：{file_time}")
            update_message(f"  檔案大小：{file_size} bytes")
            
            # 讀取並顯示基本資訊
            try:
                with open(latest_json, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                if isinstance(data, list) and data:
                    first_record = data[0]
                    if "基本資訊" in first_record:
                        lot_info = first_record["基本資訊"].get("地號建號", "")
                        update_message(f"  地號資訊：{lot_info}")
                    update_message(f"  記錄筆數：{len(data)}")
            except Exception as e:
                update_message(f"讀取檔案內容時發生錯誤：{e}")
            
            # 自動確認使用這個檔案（連續模式中）
            if continuous_mode_active:
                update_message("連續模式中，自動使用此檔案")
                selected_json_path = latest_json
                json_dir = os.path.dirname(latest_json)

                # 更新設定
                try:
                    cfg = load_config()
                except Exception:
                    cfg = {}
                cfg["json_dir"] = json_dir
                cfg["selected_json"] = selected_json_path
                try:
                    save_config(cfg, show_message=False)  # 🔥 不顯示訊息
                except Exception:
                    pass

                # 更新狀態列顯示
                update_json_status_display()
                update_message("結構化JSON檔案已成功載入")

                # 🔥 自動轉換為物件調查表資料（與手動載入流程一致）
                update_divider()
                update_message("正在自動轉換為物件調查表資料...")
                convert_transcript_to_final_data(latest_json)
                update_divider()

                # 🔥 連續模式中檢查並提示修復 data.json 與謄本資料順序（傳入選擇的文件）
                update_message("正在檢查 data.json 與謄本資料的順序一致性...")
                is_consistent, warning, data_info, transcript_info, is_order_issue = check_data_consistency(transcript_file=latest_json)

                if not is_consistent and data_info and transcript_info:
                    update_message("")
                    update_divider()
                    update_message("【發現順序不一致】")
                    update_message(f"  data.json 第一筆：{data_info['area']}{data_info['section']} {data_info['lot_number']}")
                    update_message(f"  謄本第一筆：{transcript_info['area']}{transcript_info['section']} {transcript_info['lot_number']}")
                    update_message("")
                    update_message("💡 建議：以 data.json 為準，調整謄本結構化 JSON 的順序")
                    update_message("")
                    update_message("將在 10 秒後自動調整謄本結構化 JSON 順序...")
                    update_message("如需調整 data.json，請立即按 Ctrl+C 中斷程式")
                    update_divider()
                    update_message("")

                    # 倒數計時 10 秒
                    for i in range(10, 0, -1):
                        update_message(f"倒數 {i} 秒...")
                        root.update()
                        time.sleep(1)

                    # 自動執行修復（調整謄本結構化 JSON）
                    update_message("")
                    update_message("開始調整謄本結構化 JSON 順序...")
                    success, message = fix_transcript_json_order(data_info, transcript_info)
                    if success:
                        update_message(f"[完成] {message}")
                    else:
                        update_message(f"[警告] 修復失敗：{message}")
                elif is_consistent:
                    update_message("✓ data.json 與謄本資料順序一致")
                elif warning:
                    update_message(f"[提示] {warning}")
            else:
                # 非連續模式時詢問使用者
                prompt_text = (
                    f"是否載入此結構化 JSON 檔案？\n\n"
                    f"{file_name}\n{file_time}"
                    f"<HR>"
                    f"• 是（Y）：此 JSON 對應目前 data.json 的地號\n"
                    f"           → 載入後可接著做後續查詢/估算\n\n"
                    f"• 否（N）：只是單獨處理 PDF（與 data.json 無關）\n"
                    f"           → 不需載入"
                )
                result = show_large_yesno("確認", prompt_text)
                if result:
                    selected_json_path = latest_json
                    json_dir = os.path.dirname(latest_json)

                    # 同步儲存到 config（與手動載入一致）
                    try:
                        cfg = load_config() or {}
                    except Exception:
                        cfg = {}
                    cfg["json_dir"] = json_dir
                    cfg["selected_json"] = selected_json_path
                    try:
                        save_config(cfg, show_message=False)
                    except Exception:
                        pass

                    update_json_status_display()
                    update_message("結構化JSON檔案已成功載入")

                    # 🔥 帶入時自動依 data.json 順序重排（不彈窗）
                    # 注意：不能依「一致性檢查」決定是否重排，因為一致性只看第一筆土地，
                    #       無法偵測「建物排在土地之前」這種需要重排的情況。
                    #       直接呼叫 fix，fix 內部會判斷需不需要實際寫檔。
                    try:
                        _ic, _wm, _di, _ti, _io = check_data_consistency(transcript_file=latest_json)
                        if _di and _ti and (_ic or _io):
                            _ok, _msg = fix_transcript_json_order(_di, _ti)
                            if _ok and "無需調整" not in _msg:
                                # 真的重排了才顯示
                                update_message(f"[OK] {_msg}")
                                update_json_status_display()
                            elif not _ok:
                                update_message(f"[警告] 自動重排失敗：{_msg}")
                            # 「無需調整」→ 靜默，不顯示
                        elif _di and _ti and not _ic and not _io:
                            update_message("[警告] data.json 與謄本完全不同地段（可能是不同案件），不會自動重排")
                    except Exception as _e:
                        update_message(f"[警告] 一致性檢查失敗: {_e}")

                    # 🔥 跳出勾選對話框讓使用者選擇要載入的記錄與所有權人，再轉換物調資料
                    update_divider()
                    update_message("請於勾選視窗選擇要載入的記錄...")
                    _ok2, _msg2 = _apply_filter_and_convert(latest_json)
                    if _ok2:
                        update_divider()
                        update_message("[完成] JSON 已載入並轉換完成")
                        update_message("   • 可執行「土地增值稅」程式計算稅額")
                        update_message("   • 可點擊「編輯物件資料」補充其他資訊")
                        update_divider()
                    else:
                        update_message(f"[取消] {_msg2}")
                else:
                    update_message("使用者取消載入結構化JSON檔案")
        else:
            update_message("在output目錄中未找到結構化JSON檔案")
            if not continuous_mode_active:
                # 非連續模式時提供手動選擇
                choose_json_file()
    
    except Exception as e:
        update_message(f"讀取結構化JSON功能執行時發生錯誤：{e}")
    
    # 連續模式中需要觸發下一個腳本
    if continuous_mode_active:
        # 模擬子程序結束，觸發下一個腳本
        root.after(2000, schedule_next_script)  # 2秒後執行下一個

def update_json_status_display():
    """更新JSON狀態顯示"""
    if selected_json_path:
        filename = os.path.basename(selected_json_path)
        try:
            with open(selected_json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            lot_display = ""
            if isinstance(data, list) and len(data) > 0:
                first_record = data[0]
                if "基本資訊" in first_record:
                    lot_info = first_record["基本資訊"].get("地號建號", "")
                    lot_display = simplify_lot_number(lot_info)

            # 🔥 從檔名提取日期時間（格式：YYYYMMDD_HHMM）
            import re
            datetime_match = re.search(r'(\d{8}_\d{4})', filename)
            datetime_str = datetime_match.group(1) if datetime_match else ""

            if lot_display and datetime_str:
                json_status_label.config(text=f"JSON: {datetime_str} | {lot_display}", fg='#006400')
            elif lot_display:
                json_status_label.config(text=f"JSON: {lot_display}", fg='#006400')
            else:
                json_status_label.config(text=f"JSON: {datetime_str if datetime_str else filename}", fg='#006400')
        except:
            json_status_label.config(text=f"JSON: {filename}", fg='#006400')

def start_subprocess_internal(script_name):
    """內部函數，用於連續模式執行腳本"""
    global is_running, process, selected_json_path, json_dir, current_running_button

    # 🔥 解析 nlma_nbupic.py:城市 格式（連續模式專用識別碼，因同一支腳本對應屏東/臺南兩個按鈕）
    nbupic_city = None
    if script_name.startswith("nlma_nbupic.py:"):
        nbupic_city = script_name.split(':', 1)[1]
        script_name = "nlma_nbupic.py"

    # 特殊處理讀取JSON功能
    if script_name == "read_json":
        handle_read_json_function()
        return

    # 🔥 特殊處理編輯物件資料功能
    if script_name == "edit_data":
        update_message("正在開啟【編輯物件資料】視窗...")
        # 切換到謄本頁面
        if continuous_mode_active:
            show_transcript_page()
        # 打開編輯視窗
        try:
            data_editor.open_data_editor()
            update_message("【編輯物件資料】視窗已開啟，請完成編輯後繼續...")
        except Exception as e:
            update_message(f"開啟編輯視窗時發生錯誤：{e}")
        # 連續模式中觸發下一個腳本
        if continuous_mode_active:
            root.after(2000, schedule_next_script)
        return

    # 🔥 如果是子頁面程式，自動切換到對應頁面
    if continuous_mode_active and script_name in script_to_page_map:
        page = script_to_page_map[script_name]
        if page == "real_estate":
            show_real_estate_price_page()
            update_message(f"→ 自動切換到【實價登錄】頁面")
        elif page == "transcript":
            show_transcript_page()
            update_message(f"→ 自動切換到【謄本與土增稅】頁面")
        elif page == "land_use":
            show_land_use_page()
            update_message(f"→ 自動切換到【土地使用分區】頁面")
        elif page == "building_license":
            show_building_license_page()
            update_message(f"→ 自動切換到【使用執照查詢】頁面")

    # 🔥 找到對應的按鈕並更新狀態
    if nbupic_city:
        # NBUPIC 依城市決定按鈕（script_to_button_map 無法區分）
        button = building_license_pt_button if nbupic_city == "屏東縣" else building_license_tn_button
    else:
        button = script_to_button_map.get(script_name)
    if button and continuous_mode_active:
        current_running_button = button
        original_text = button.cget("text")
        button.config(text=f"關閉 {original_text}", state=tk.NORMAL)

    # 🔥 尋找最新版本的電子謄本結構化
    def find_latest_transcript_exe(base_dir):
        """在指定目錄尋找最新版本的電子謄本結構化.exe"""
        import glob
        import re

        # 搜尋所有版本（支援 電子謄本結構化.exe 和 電子謄本結構化v*.exe）
        pattern = os.path.join(base_dir, "電子謄本結構化*.exe")
        exe_files = glob.glob(pattern)

        # 也搜尋子資料夾中的 exe（onedir 模式）
        pattern_subdir = os.path.join(base_dir, "電子謄本結構化*", "電子謄本結構化*.exe")
        exe_files += glob.glob(pattern_subdir)

        if not exe_files:
            return None

        # 解析版本號並排序
        def extract_version(filepath):
            filename = os.path.basename(filepath)
            # 支援格式：電子謄本結構化.exe, 電子謄本結構化v1.0.exe, 電子謄本結構化_v1.0.2a.exe
            match = re.search(r'[v_](\d+)\.(\d+)(?:\.(\d+))?([a-zA-Z])?', filename)
            if match:
                major = int(match.group(1))
                minor = int(match.group(2))
                patch = int(match.group(3)) if match.group(3) else 0
                suffix = match.group(4).lower() if match.group(4) else ''
                suffix_val = ord(suffix) if suffix else 0
                return (major, minor, patch, suffix_val)
            # 沒有版本號的視為 0.0.0
            return (0, 0, 0, 0)

        exe_files.sort(key=extract_version, reverse=True)
        return exe_files[0]

    try:
        # 🔥 取得腳本的正確路徑（支援 PyInstaller 打包）
        script_path = get_resource_path(script_name)

        # 🔥 修正：在打包環境中使用嵌入式 Python 執行子程式
        if getattr(sys, 'frozen', False):
            # 🔥 特殊處理：GUI_transcript_pdf 是獨立的 exe（自動尋找最新版本）
            if script_name == "GUI_transcript_pdf 1141021-01.py":
                base_dir = os.path.dirname(sys.executable)
                latest_exe = find_latest_transcript_exe(base_dir)

                if latest_exe:
                    cmd = [latest_exe, '--base-dir', base_dir]
                    update_message(f"啟動電子謄本結構化：{os.path.basename(latest_exe)}")
                else:
                    update_message("[錯誤] 找不到電子謄本結構化.exe")
                    is_running = False
                    if button and continuous_mode_active:
                        button.config(text=original_text, state=tk.DISABLED)
                        current_running_button = None
                    if continuous_mode_active:
                        schedule_next_script()
                    return
            else:
                # 打包環境：使用 python_embedded 資料夾中的 Python
                # 🔥 使用 pythonw.exe（不顯示 CMD 視窗）
                # 🔥 hinet.py 和 qpt_hinet.py 已改用 GUI 對話框，不需要 CMD
                python_embedded = os.path.join(os.path.dirname(sys.executable), 'python_embedded', 'pythonw.exe')
                if os.path.exists(python_embedded):
                    cmd = [python_embedded, script_path]
                else:
                    # 如果找不到嵌入式 Python，顯示錯誤訊息
                    update_message("[錯誤] 找不到 Python 執行環境，請確認打包是否完整")
                    is_running = False
                    if button and continuous_mode_active:
                        button.config(text=original_text, state=tk.DISABLED)
                        current_running_button = None
                    if continuous_mode_active:
                        schedule_next_script()
                    return
        else:
            # 開發環境：使用系統 Python
            cmd = [sys.executable, script_path]

        # 🔥 傳遞基準目錄給子程式（data.json 和工作資料夾的位置）
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        cmd += ['--base-dir', base_dir]

        # 稅務腳本帶入 JSON 參數
        if script_name in ('kaohsiung_tax.py', 'finance_tax.py'):
            if selected_json_path:
                cmd += ['--json', selected_json_path]
            else:
                cmd += ['--json-dir', json_dir or JSON_DIR_DEFAULT]

        # 🔥 nlma_nbupic.py：連續模式中依識別碼帶入 --city
        if script_name == "nlma_nbupic.py" and nbupic_city:
            cmd += ['--city', nbupic_city]
            update_message(f"[NBUPIC] 將查詢「{nbupic_city}」")

        # 🔥 設定環境變數以隔離 Python 路徑
        env = os.environ.copy()
        # 🔥 強制子程式 stdout/stderr 用 UTF-8（避免 Windows CP950 與 Popen encoding='utf-8' 不一致）
        env['PYTHONIOENCODING'] = 'utf-8'
        cwd = None
        if getattr(sys, 'frozen', False):
            # 打包環境：設定環境變數只使用 python_embedded
            python_embedded_dir = os.path.join(os.path.dirname(sys.executable), 'python_embedded')
            internal_dir = os.path.join(os.path.dirname(sys.executable), '_internal')
            env['PYTHONHOME'] = python_embedded_dir
            # 🔥 設定 PYTHONPATH 讓子程式能 import：
            # 1. _internal 中的模組（base_dir_helper等）
            # 2. python_embedded 中的套件（PyPDF2等）
            python_site_packages = os.path.join(python_embedded_dir, 'Lib', 'site-packages')
            env['PYTHONPATH'] = internal_dir + os.pathsep + python_site_packages
            # 禁用用戶 site-packages
            env['PYTHONNOUSERSITE'] = '1'
            # 🔥 設定工作目錄為 _internal，讓子程式能找到同目錄的模組
            # 雖然 cwd 在 _internal，但 --base-dir 參數會告訴子程式主程式目錄在哪
            cwd = internal_dir
            print(f"[DEBUG] 子程式工作目錄: {cwd}")
            print(f"[DEBUG] 子程式 BASE_DIR 參數: {base_dir}")

        # 🔥 隱藏子程式的 CMD 視窗（Windows）
        startupinfo = None
        if sys.platform == 'win32':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

        process = subprocess.Popen(
            cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace',  # errors='replace' 容錯：非 UTF-8 byte 變 ?
            bufsize=1, universal_newlines=True, env=env, cwd=cwd,
            startupinfo=startupinfo
        )
        is_running = True
        threading.Thread(target=read_output, args=(process,), daemon=True).start()
    except Exception as e:
        update_message(f"啟動程序失敗: {str(e)}")
        is_running = False
        # 🔥 恢復按鈕狀態
        if button and continuous_mode_active:
            button.config(text=original_text, state=tk.DISABLED)
            current_running_button = None
        if continuous_mode_active:
            schedule_next_script()

def web_open(url):
    import webbrowser
    webbrowser.open(url)

# 頁面切換
def show_real_estate_price_page():
    left_button_frame.pack_forget()
    middle_button_frame.pack_forget()
    right_button_frame.pack_forget()
    transcript_page.pack_forget()
    utility_page.pack_forget()
    land_use_page.pack_forget()
    building_license_page.pack_forget()
    real_estate_price_page.pack(side=tk.TOP, fill=tk.X, pady=(10, 0))

    # 🔥 如果在連續模式中，禁用子頁面的所有按鈕（除了返回按鈕和正在執行的按鈕）
    if continuous_mode_active:
        for button in real_estate_page_buttons:
            if button == real_estate_back_button or button == current_running_button:
                button.config(state=tk.NORMAL)
            else:
                button.config(state=tk.DISABLED)

def show_transcript_page():
    left_button_frame.pack_forget()
    middle_button_frame.pack_forget()
    right_button_frame.pack_forget()
    real_estate_price_page.pack_forget()
    utility_page.pack_forget()
    land_use_page.pack_forget()
    building_license_page.pack_forget()
    transcript_page.pack(side=tk.TOP, fill=tk.X, pady=(10, 0))

    # 🔥 如果在連續模式中，禁用子頁面的所有按鈕（除了返回按鈕和正在執行的按鈕）
    if continuous_mode_active:
        for button in transcript_page_buttons:
            if button == transcript_back_button or button == current_running_button:
                button.config(state=tk.NORMAL)
            else:
                button.config(state=tk.DISABLED)

def show_utility_page():
    """顯示功能分頁（案件管理相關功能）"""
    left_button_frame.pack_forget()
    middle_button_frame.pack_forget()
    right_button_frame.pack_forget()
    real_estate_price_page.pack_forget()
    transcript_page.pack_forget()
    land_use_page.pack_forget()
    building_license_page.pack_forget()
    utility_page.pack(side=tk.TOP, fill=tk.X, pady=(10, 0))

    # 🔥 如果在連續模式中，禁用子頁面的所有按鈕（除了返回按鈕）
    if continuous_mode_active:
        for button in utility_page_buttons:
            if button == utility_back_button:
                button.config(state=tk.NORMAL)
            else:
                button.config(state=tk.DISABLED)

def show_land_use_page():
    """顯示土地使用分區分頁"""
    left_button_frame.pack_forget()
    middle_button_frame.pack_forget()
    right_button_frame.pack_forget()
    real_estate_price_page.pack_forget()
    transcript_page.pack_forget()
    utility_page.pack_forget()
    building_license_page.pack_forget()
    land_use_page.pack(side=tk.TOP, fill=tk.X, pady=(10, 0))

    # 🔥 如果在連續模式中，禁用子頁面的所有按鈕（除了返回按鈕和正在執行的按鈕）
    if continuous_mode_active:
        for button in land_use_page_buttons:
            if button == land_use_back_button or button == current_running_button:
                button.config(state=tk.NORMAL)
            else:
                button.config(state=tk.DISABLED)

def show_building_license_page():
    """顯示使用執照查詢分頁（讓使用者選縣市）"""
    left_button_frame.pack_forget()
    middle_button_frame.pack_forget()
    right_button_frame.pack_forget()
    real_estate_price_page.pack_forget()
    transcript_page.pack_forget()
    utility_page.pack_forget()
    land_use_page.pack_forget()
    building_license_page.pack(side=tk.TOP, fill=tk.X, pady=(10, 0))

    # 🔥 連續模式中禁用子頁面非返回按鈕
    if continuous_mode_active:
        for button in building_license_page_buttons:
            if button == building_license_back_button or button == current_running_button:
                button.config(state=tk.NORMAL)
            else:
                button.config(state=tk.DISABLED)

def show_main_page():
    real_estate_price_page.pack_forget()
    transcript_page.pack_forget()
    utility_page.pack_forget()
    land_use_page.pack_forget()
    building_license_page.pack_forget()
    left_button_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, anchor="e")
    middle_button_frame.pack(side=tk.LEFT, padx=COLUMN_PADX, expand=True, anchor="center")
    right_button_frame.pack(side=tk.RIGHT, padx=COLUMN_PADX, anchor="w")

    # 🔥 如果在連續模式中，恢復主頁面的按鈕狀態（停止連續模式、頁面切換按鈕和正在執行的按鈕可用）
    if continuous_mode_active:
        for button in main_page_buttons:
            # 允許：停止連續模式按鈕、頁面切換按鈕、正在執行的按鈕
            if button not in [continuous_mode_button, real_estate_price_button, transcript_button, utility_button, land_use_zone_button, current_running_button]:
                button.config(state=tk.DISABLED)
            else:
                button.config(state=tk.NORMAL)

# 互斥/啟動
def toggle_buttons(button, original_text, script_name):
    global is_running
    if is_running:
        exit_subprocess(button)
        # 🔥 恢復地籍便民系統的特殊樣式
        if button == start_land_button:
            button.config(text=original_text, **priority_button_style)
        else:
            button.config(text=original_text)
        enable_all_buttons()
    else:
        button.config(text=f"關閉 {original_text}")
        # 🔥 立即禁用其他按鈕（讓使用者看到「正在執行」的狀態，包含對話框期間）
        disable_other_buttons(button)
        start_subprocess(button, script_name)
        # 🔥 若 start_subprocess 內部因「取消」而沒實際啟動，會自行 enable_all_buttons + 還原文字

def disable_other_buttons(active_button):
    # 🔥 如果在連續模式中，保持頁面切換按鈕可用
    if continuous_mode_active:
        for button in all_buttons:
            # 允許：停止連續模式按鈕、頁面切換按鈕
            if button not in [active_button, real_estate_price_button, transcript_button]:
                button.config(state=tk.DISABLED)
    else:
        # 非連續模式：禁用除了當前執行按鈕外的所有按鈕
        for button in all_buttons:
            if button != active_button:
                button.config(state=tk.DISABLED)

def enable_all_buttons():
    for button in all_buttons:
        button.config(state=tk.NORMAL)

# 🔥 主頁面按鈕（重新分配）
# 左邊 4 個 - 🔥 地籍便民系統使用特殊樣式，謄本與土增稅在其下方
start_land_button = tk.Button(left_button_frame, text="地籍便民系統",
                             command=lambda: toggle_buttons(start_land_button, "地籍便民系統", "land.py"),
                             **priority_button_style)  # 🔥 使用特殊樣式

# 🔥 謄本與土增稅按鈕（紫色）- 移至左側，在地籍便民系統下方
transcript_button_style = {
    'font': custom_font,
    'width': BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'bg': '#9C27B0', 'fg': 'white',  # 紫色
    'activebackground': '#7B1FA2', 'activeforeground': 'white'
}
transcript_button = tk.Button(left_button_frame, text="謄本與土增稅", command=show_transcript_page, **transcript_button_style)

start_land_use_button = tk.Button(left_button_frame, text="使用執照查詢", command=show_building_license_page, **button_style)
start_important_facility_button = tk.Button(left_button_frame, text="重要設施查詢", command=lambda: toggle_buttons(start_important_facility_button, "重要設施查詢", "gbmap.py"), **button_style)

# 中間 4 個 - 🔥 國土測繪圖資改為土地使用分區分頁按鈕，V523 系統移至此處
# 🔥 土地使用分區：淡紅色
land_use_zone_button_style = {**button_style, 'bg': '#FFCDD2', 'fg': 'black',
                              'activebackground': '#EF9A9A', 'activeforeground': 'black'}
land_use_zone_button = tk.Button(middle_button_frame, text="土地使用分區", command=show_land_use_page, **land_use_zone_button_style)
start_V523_button = tk.Button(middle_button_frame, text="V523 系統", command=lambda: toggle_buttons(start_V523_button, "V523 系統", "V523.py"), **button_style)
start_land_partition_button = tk.Button(middle_button_frame, text="國土分區查詢", command=lambda: toggle_buttons(start_land_partition_button, "國土分區查詢", "tcdmap.py"), **button_style)
start_mapping_button = tk.Button(middle_button_frame, text="地質雲查詢", command=lambda: toggle_buttons(start_mapping_button, "地質雲查詢", "geologycloud.py"), **button_style)

# 右邊 4 個 - 🔥 編輯物件資料移至此處（實價登錄下方）
# 🔥 實價登錄：淡奶油色 #FFF3CD
real_estate_price_button_style = {**button_style, 'bg': '#FFF3CD', 'fg': 'black',
                                   'activebackground': '#FFE99B', 'activeforeground': 'black'}
real_estate_price_button = tk.Button(right_button_frame, text="實價登錄", command=show_real_estate_price_page, **real_estate_price_button_style)

# 🔥 編輯物件資料按鈕（淺紫色）
edit_data_button_style = {
    'font': custom_font,
    'width': BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'bg': '#CE93D8', 'fg': 'black',  # 淺紫色
    'activebackground': '#BA68C8', 'activeforeground': 'black'
}
edit_object_data_button = tk.Button(right_button_frame, text="編輯物件資料", command=data_editor.open_data_editor, **edit_data_button_style)

continuous_mode_button = tk.Button(right_button_frame, text="連續模式",
                                  command=toggle_continuous_mode,
                                  **continuous_button_style)

# 🔥 功能分頁按鈕（淺綠色，黑色字）
utility_button_style = {
    'font': custom_font,
    'width': BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'bg': '#81C784', 'fg': 'black',  # 淺綠色，黑色字
    'activebackground': '#66BB6A', 'activeforeground': 'black'
}

utility_button = tk.Button(right_button_frame, text="功能分頁",
                          command=show_utility_page,
                          **utility_button_style)

# 實價登錄頁面按鈕
interior_button = tk.Button(real_estate_left_frame, text="內政部", command=lambda: toggle_buttons(interior_button, "內政部", "interior_price.py"), **sub_page_button_style)
price104_button = tk.Button(real_estate_middle_frame, text="104實價網", command=lambda: toggle_buttons(price104_button, "104實價網", "price104.py"), **sub_page_button_style)
real_estate_back_button = tk.Button(real_estate_right_frame, text="回至主頁", command=show_main_page, **sub_page_button_style)

# 🔥 功能分頁按鈕（案件管理相關功能）
def open_work_folder():
    """開啟工作資料夾"""
    try:
        if os.path.exists(get_data_json_path()):
            with open(get_data_json_path(), 'r', encoding='utf-8') as f:
                data_list = json.load(f)
                if data_list:
                    first_data = data_list[0]
                    area = first_data.get('area', '')
                    section = first_data.get('section', '')
                    lot_number = first_data.get('lot_number', '')
                    folder_name = f"{area}{section}-{lot_number}"

                    if os.path.exists(folder_name):
                        os.startfile(folder_name)
                        update_message(f"已開啟資料夾：{folder_name}")
                    else:
                        update_message(f"[警告] 資料夾不存在：{folder_name}")
                        messagebox.showwarning("警告", f"資料夾不存在：\n{folder_name}\n\n請先執行相關程式建立資料夾。")
                else:
                    update_message("[錯誤] data.json 是空的")
                    messagebox.showerror("錯誤", "data.json 是空的，無法確定資料夾名稱")
        else:
            update_message("[錯誤] 找不到 data.json")
            messagebox.showerror("錯誤", "找不到 data.json，無法確定資料夾名稱")
    except Exception as e:
        update_message(f"[錯誤] 開啟資料夾失敗：{e}")
        messagebox.showerror("錯誤", f"開啟資料夾失敗：\n{e}")

# 功能分頁 - 左側按鈕（載入案件、讀取結構化JSON）
utility_load_data_button = tk.Button(utility_left_frame, text="載入案件 data.json",
                                     command=load_data_json_from_folder,
                                     **sub_page_button_style)
# 🔥 讀取結構化JSON：原本在謄本專區，移到功能分頁左欄
# 用 lambda 延後 choose_json_file 的查名（該函式定義在後面）
read_json_button = tk.Button(utility_left_frame, text="讀取結構化JSON",
                             command=lambda: choose_json_file(),
                             **sub_page_button_style)

# 功能分頁 - 中間按鈕（匯入/匯出封存）
utility_import_archive_button = tk.Button(utility_middle_frame, text="匯入封存",
                                          command=import_archive,
                                          **sub_page_button_style)
utility_export_archive_button = tk.Button(utility_middle_frame, text="匯出封存",
                                          command=data_editor.export_archive,
                                          **sub_page_button_style)

# 功能分頁 - 右側按鈕（修改帳密、開啟工作資料夾、回至主頁）
# 🔥 修改帳密：開啟 Windows 認證管理員的帳密管理介面
def _open_credential_manager():
    try:
        from credential_helper import open_credential_manager as _ocm
        _ocm(parent=root)
    except Exception as _e:
        messagebox.showerror("錯誤", f"開啟帳密管理失敗：\n{_e}")

utility_credential_button = tk.Button(utility_right_frame, text="修改帳密",
                                      command=_open_credential_manager,
                                      **sub_page_button_style)
# 🔥 開啟工作資料夾：原本在左欄，移到右欄回至主頁上方
utility_open_folder_button = tk.Button(utility_right_frame, text="開啟工作資料夾",
                                       command=open_work_folder,
                                       **sub_page_button_style)
utility_back_button = tk.Button(utility_right_frame, text="回至主頁",
                                command=show_main_page,
                                **sub_page_button_style)

# 功能分頁按鈕列表
utility_page_buttons = [
    utility_load_data_button, read_json_button,
    utility_import_archive_button, utility_export_archive_button,
    utility_credential_button, utility_open_folder_button, utility_back_button
]

# 🔥 土地使用分區分頁按鈕
# 左側：國土測繪圖資（淡藍色）、高雄都市計畫（淡綠色）
land_use_nlscmaps_button_style = {**sub_page_button_style, 'bg': '#BBDEFB', 'fg': 'black',
                                   'activebackground': '#90CAF9', 'activeforeground': 'black'}
land_use_urbangis_button_style = {**sub_page_button_style, 'bg': '#C8E6C9', 'fg': 'black',
                                   'activebackground': '#A5D6A7', 'activeforeground': 'black'}
land_use_nlscmaps_button = tk.Button(land_use_left_frame, text="國土測繪圖資",
                                      command=lambda: toggle_buttons(land_use_nlscmaps_button, "國土測繪圖資", "nlscmaps.py"),
                                      **land_use_nlscmaps_button_style)
land_use_urbangis_button = tk.Button(land_use_left_frame, text="高雄都市計畫",
                                      command=lambda: toggle_buttons(land_use_urbangis_button, "高雄都市計畫", "urbangis.py"),
                                      **land_use_urbangis_button_style)

# 中間：全國土地使用分區
land_use_luz_button = tk.Button(land_use_middle_frame, text="全國土地使用分區",
                                 command=lambda: toggle_buttons(land_use_luz_button, "全國土地使用分區", "luz.py"),
                                 **sub_page_button_style)

# 右側：回至主頁
land_use_back_button = tk.Button(land_use_right_frame, text="回至主頁",
                                  command=show_main_page,
                                  **sub_page_button_style)

# 土地使用分區分頁按鈕列表
land_use_page_buttons = [
    land_use_nlscmaps_button, land_use_urbangis_button,
    land_use_luz_button, land_use_back_button
]

# 🔥 使用執照查詢分頁按鈕
#    高雄市 → nlma.py（既有，buildmis.kcg.gov.tw Vue 站）
#    屏東縣 / 台南市 → nlma_nbupic.py（NBUPIC 雲端共用站，9 縣市同一套）
#    啟動時透過 NLMA_NBUPIC_CITY_FOR_BUTTON 字典告訴 start_subprocess 要傳 --city
NLMA_NBUPIC_CITY_FOR_BUTTON = {}  # id(button) → 縣市名稱

building_license_kh_button = tk.Button(building_license_left_frame, text="高雄市",
                                       command=lambda: toggle_buttons(building_license_kh_button, "高雄市", "nlma.py"),
                                       **sub_page_button_style)
building_license_pt_button = tk.Button(building_license_middle_frame, text="屏東縣",
                                       command=lambda: toggle_buttons(building_license_pt_button, "屏東縣", "nlma_nbupic.py"),
                                       **sub_page_button_style)
building_license_tn_button = tk.Button(building_license_middle_frame, text="台南市",
                                       command=lambda: toggle_buttons(building_license_tn_button, "台南市", "nlma_nbupic.py"),
                                       **sub_page_button_style)
building_license_back_button = tk.Button(building_license_right_frame, text="回主選單",
                                          command=show_main_page,
                                          **sub_page_button_style)

# 為兩個 NBUPIC 城市按鈕登記要傳的 --city 參數
NLMA_NBUPIC_CITY_FOR_BUTTON[id(building_license_pt_button)] = "屏東縣"
NLMA_NBUPIC_CITY_FOR_BUTTON[id(building_license_tn_button)] = "臺南市"

# 使用執照查詢分頁按鈕列表
building_license_page_buttons = [
    building_license_kh_button, building_license_pt_button,
    building_license_tn_button, building_license_back_button
]

# ── 謄本專區相關函數 ──
def choose_json_file():
    """選擇結構化 JSON（同時記住路徑並轉換為物調資料）"""
    global selected_json_path, json_dir

    try:
        initial = os.path.abspath(json_dir if json_dir else JSON_DIR_DEFAULT)
    except Exception:
        initial = os.path.abspath(JSON_DIR_DEFAULT)

    path = filedialog.askopenfilename(
        initialdir=initial,
        title="選擇結構化 JSON",
        filetypes=[("JSON files","*.json")]
    )

    if path:
        selected_json_path = path
        json_dir = os.path.dirname(path)

        # 儲存設定
        try:
            cfg = load_config()
        except Exception:
            cfg = {}
        cfg["json_dir"] = json_dir
        cfg["selected_json"] = selected_json_path
        try:
            save_config(cfg, show_message=False)  # 🔥 讀取JSON時不顯示訊息
        except Exception:
            pass

        update_message(f"已選擇 JSON：{selected_json_path}")

        # 解析並顯示簡化的地號資訊
        try:
            with open(selected_json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            lot_display = ""
            if isinstance(data, list) and len(data) > 0:
                first_record = data[0]
                if "基本資訊" in first_record:
                    lot_info = first_record["基本資訊"].get("地號建號", "")
                    lot_display = simplify_lot_number(lot_info)

            # 更新狀態列
            filename = os.path.basename(selected_json_path)

            # 🔥 從檔名提取日期時間（格式：YYYYMMDD_HHMM）
            import re
            datetime_match = re.search(r'(\d{8}_\d{4})', filename)
            datetime_str = datetime_match.group(1) if datetime_match else ""

            if lot_display and datetime_str:
                json_status_label.config(
                    text=f"JSON: {datetime_str} | {lot_display}",
                    fg='#006400'
                )
            elif lot_display:
                json_status_label.config(
                    text=f"JSON: {lot_display}",
                    fg='#006400'
                )
            else:
                json_status_label.config(
                    text=f"JSON: {datetime_str if datetime_str else filename}",
                    fg='#006400'
                )
        except Exception as e:
            filename = os.path.basename(selected_json_path)
            json_status_label.config(
                text=f"JSON: {filename}",
                fg='#006400'
            )
            update_message(f"[警告] 解析地號資訊失敗: {e}")

        # 🔥 先自動依 data.json 順序重排謄本（不彈窗），再轉換物調資料
        # 即使一致性檢查 OK 也要 call fix，才能處理「建物排在土地前」這類 layout 問題
        try:
            is_consistent, warning_msg, data_info, transcript_info, is_order_issue = check_data_consistency(transcript_file=path)
            if data_info and transcript_info and (is_consistent or is_order_issue):
                success, message = fix_transcript_json_order(data_info, transcript_info)
                if success and "無需調整" not in message:
                    update_divider()
                    update_message(f"[OK] {message}")
                    update_json_status_display()
                    update_divider()
                elif not success:
                    update_divider()
                    update_message(f"[警告] 自動重排失敗：{message}")
                    update_divider()
                # 「無需調整」→ 靜默，不顯示分隔線與訊息
            elif data_info and transcript_info and not is_consistent and not is_order_issue:
                update_divider()
                update_message("[警告] data.json 與謄本完全不同地段（可能是不同案件）")
                update_message(f"  data.json[0]: {data_info['area']}{data_info['section']} {data_info['lot_number']}")
                update_message(f"  謄本[0]: {transcript_info['area']}{transcript_info['section']} {transcript_info['lot_number']}")
                update_message("  → 不會自動重排；如需切換案件請手動載入正確的 data.json")
                update_divider()
        except Exception as e:
            update_message(f"[警告] 一致性檢查失敗: {e}")

        # 🔥 跳出勾選對話框讓使用者選擇要載入的記錄與所有權人，再轉換物調資料
        update_divider()
        update_message("請於勾選視窗選擇要載入的記錄...")
        _ok2, _msg2 = _apply_filter_and_convert(path)
        if _ok2:
            update_divider()
            update_message("[完成] JSON 已載入並轉換完成")
            update_message("   • 可執行「土地增值稅」程式計算稅額")
            update_message("   • 可點擊「編輯物件資料」補充其他資訊")
            update_divider()
        else:
            update_message(f"[取消] {_msg2}")

    else:
        update_message("未選擇檔案")

# ── 謄本專區頁面按鈕 ──
# 【左欄 - 資料來源】
# 🔥 全國地政電子謄本：土黃色按鈕；點擊後立即顯示「模組加載中」提示，避免使用者誤以為沒反應
hinet_button_style = {
    'font': sub_page_font, 'width': SUB_PAGE_BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'borderwidth': 2, 'bg': '#B8860B', 'fg': 'white',  # 土黃色（DarkGoldenrod）
    'activebackground': '#A07509', 'activeforeground': 'white'
}
hinet_button = tk.Button(
    transcript_left_frame, text="全國地政電子謄本",
    command=lambda: (
        update_message("歡迎使用【地籍電子謄本】小程式，模組加載中...")
        or toggle_buttons(hinet_button, "全國地政電子謄本", "hinet.py")
    ),
    **hinet_button_style
)
qpt_hinet_button = tk.Button(transcript_left_frame, text="全功能地籍資料查詢", command=lambda: toggle_buttons(qpt_hinet_button, "全功能地籍資料查詢", "qpt_hinet.py"), **sub_page_button_style)

# 🔥 電子謄本結構化按鈕（深藍色）
structured_transcript_button_style = {
    'font': sub_page_font, 'width': SUB_PAGE_BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'borderwidth': 2, 'bg': '#1565C0', 'fg': 'white',  # 深藍色
    'activebackground': '#0D47A1', 'activeforeground': 'white'
}
structured_transcript_button = tk.Button(
    transcript_left_frame, text="電子謄本結構化",
    command=lambda: toggle_buttons(structured_transcript_button, "電子謄本結構化", "GUI_transcript_pdf 1141021-01.py"),
    **structured_transcript_button_style
)

# 【中欄 - 資料處理與土增稅】
# 🔥 讀取結構化JSON 已移到「功能分頁」左欄
# 🔥 財政部土增稅：淡金色按鈕
finance_button_style = {
    'font': sub_page_font, 'width': SUB_PAGE_BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'borderwidth': 2, 'bg': '#E6C547', 'fg': 'black',  # 淡金色
    'activebackground': '#D4B237', 'activeforeground': 'black'
}
finance_button = tk.Button(transcript_middle_frame, text="財政部土增稅",
    command=lambda: toggle_buttons(finance_button, "財政部土增稅", "finance_tax.py"), **finance_button_style)
kaohsiung_tax_button = tk.Button(transcript_middle_frame, text="高雄土增稅估算",
    command=lambda: toggle_buttons(kaohsiung_tax_button, "高雄土增稅估算", "kaohsiung_tax.py"), **sub_page_button_style)

# 【右欄 - 設定/編輯/返回】
# 🔥 暫時註釋帳號設定按鈕（未來可能會用到）
# config_manage_button = tk.Button(transcript_right_frame, text="帳號設定", command=lambda: open_config_manager(root, message_box, update_message), **sub_page_button_style)

# 🔥 地籍圖處理工具 - 動態尋找最新版本
def launch_cadastral_map_tool():
    """啟動地籍圖處理工具（自動尋找最新版本）"""
    import glob
    import subprocess
    import re

    try:
        # 🔥 取得程式所在目錄（與電子謄本結構化相同的方式）
        if getattr(sys, 'frozen', False):
            # 打包後：使用 exe 所在目錄
            base_dir = os.path.dirname(sys.executable)
        else:
            # 開發環境：使用程式檔案所在目錄
            base_dir = os.path.dirname(os.path.abspath(__file__))

        # 在程式所在目錄搜尋所有版本
        pattern = os.path.join(base_dir, "地籍圖處理工具*.exe")
        exe_files = glob.glob(pattern)

        if not exe_files:
            update_message("[錯誤] 找不到地籍圖處理工具")
            messagebox.showerror("錯誤", "找不到地籍圖處理工具\n\n請確認程式與地籍圖處理工具在同一目錄")
            return

        # 解析版本號並排序，找出最新版本
        def extract_version(filepath):
            """從檔名提取版本號用於比較"""
            filename = os.path.basename(filepath)
            # 支援格式：v0.9.1, v0.8.5a, _v0.7A 等
            match = re.search(r'[v_](\d+)\.(\d+)(?:\.(\d+))?([a-zA-Z])?', filename)
            if match:
                major = int(match.group(1))
                minor = int(match.group(2))
                patch = int(match.group(3)) if match.group(3) else 0
                suffix = match.group(4).lower() if match.group(4) else ''
                # 轉換為可比較的數字（suffix 轉換為 ASCII 值，空字串為 0）
                suffix_val = ord(suffix) if suffix else 0
                return (major, minor, patch, suffix_val)
            return (0, 0, 0, 0)

        # 按版本號排序，取最新的
        exe_files.sort(key=extract_version, reverse=True)
        latest_exe = exe_files[0]

        update_message(f"啟動地籍圖處理工具：{os.path.basename(latest_exe)}")
        subprocess.Popen([latest_exe], shell=True)

    except Exception as e:
        update_message(f"[錯誤] 啟動地籍圖處理工具失敗：{e}")
        messagebox.showerror("錯誤", f"啟動地籍圖處理工具失敗：\n{e}")

# 🔥 地籍圖處理工具按鈕（青色/藍綠色）
cadastral_map_tool_button_style = {
    'font': sub_page_font, 'width': SUB_PAGE_BUTTON_WIDTH, 'height': 1, 'relief': tk.RAISED,
    'borderwidth': 2, 'bg': '#00897B', 'fg': 'white',  # 青色/藍綠色
    'activebackground': '#00695C', 'activeforeground': 'white'
}
cadastral_map_tool_button = tk.Button(
    transcript_right_frame,
    text="地籍圖處理工具",
    command=launch_cadastral_map_tool,
    **cadastral_map_tool_button_style
)

transcript_back_button = tk.Button(transcript_right_frame, text="回至主頁", command=show_main_page, **sub_page_button_style)

# 🔥 腳本名稱到按鈕的映射（用於連續模式）
script_to_button_map = {}  # 將在按鈕定義後初始化

def simplify_lot_number(lot_string):
    """簡化地號顯示：鳳山區鳳青段 0169-0000地號 -> 鳳山區鳳青段 169"""
    import re
    
    # 匹配格式：XX區XX段 0000-0000地號 或 XX區XX段 0000地號
    match = re.search(r'(.+?段)\s*(\d+)(?:-(\d+))?(?:地號|建號)?', lot_string)
    if match:
        district_section = match.group(1)  # 鳳山區鳳青段
        main_num = int(match.group(2))     # 169 (去除前導0)
        sub_num_str = match.group(3)      # 可能是 None 或 "0000"
        
        if sub_num_str:
            sub_num = int(sub_num_str)
            if sub_num == 0:
                return f"{district_section} {main_num}"
            else:
                return f"{district_section} {main_num}-{sub_num}"
        else:
            return f"{district_section} {main_num}"
    
    return lot_string  # 如果格式不符，返回原字串

# 主頁面所有按鈕 - 🔥 更新：移除 start_nlscmaps_button, start_planning_button，加入 land_use_zone_button, edit_object_data_button
main_page_buttons = [
    start_land_use_button, start_important_facility_button, start_V523_button,
    start_land_button, land_use_zone_button,
    start_land_partition_button, start_mapping_button, real_estate_price_button, transcript_button,
    continuous_mode_button, utility_button, edit_object_data_button
]
real_estate_page_buttons = [interior_button, price104_button, real_estate_back_button]
transcript_page_buttons = [
    hinet_button, qpt_hinet_button, structured_transcript_button,
    finance_button, kaohsiung_tax_button,
    # config_manage_button,  # 🔥 已註釋
    cadastral_map_tool_button, transcript_back_button
]
all_buttons = main_page_buttons + real_estate_page_buttons + transcript_page_buttons + land_use_page_buttons + building_license_page_buttons

# 🔥 初始化腳本名稱到按鈕的映射 - 更新：nlscmaps.py, urbangis.py 改用分頁按鈕，新增 luz.py
script_to_button_map = {
    "land.py": start_land_button,
    "nlma.py": building_license_kh_button,  # 🔥 改：nlma.py 對應到次頁面的「高雄市」按鈕
    "gbmap.py": start_important_facility_button,
    "V523.py": start_V523_button,
    "nlscmaps.py": land_use_nlscmaps_button,  # 🔥 改用土地使用分區分頁中的按鈕
    "urbangis.py": land_use_urbangis_button,  # 🔥 改用土地使用分區分頁中的按鈕
    "luz.py": land_use_luz_button,  # 🔥 新增：全國土地使用分區
    "tcdmap.py": start_land_partition_button,
    "geologycloud.py": start_mapping_button,
    "interior_price.py": interior_button,
    "price104.py": price104_button,
    "hinet.py": hinet_button,
    "qpt_hinet.py": qpt_hinet_button,
    "GUI_transcript_pdf 1141021-01.py": structured_transcript_button,
    "read_json": read_json_button,
    "finance_tax.py": finance_button,
    "kaohsiung_tax.py": kaohsiung_tax_button,
}

# 🔥 腳本名稱到頁面的映射（用於連續模式自動切換頁面）
script_to_page_map = {
    # 實價登錄子頁面
    "interior_price.py": "real_estate",
    "price104.py": "real_estate",
    # 謄本與土增稅子頁面
    "hinet.py": "transcript",
    "qpt_hinet.py": "transcript",
    "GUI_transcript_pdf 1141021-01.py": "transcript",
    "read_json": "transcript",
    "finance_tax.py": "transcript",
    "kaohsiung_tax.py": "transcript",
    "edit_data": "transcript",
    # 🔥 土地使用分區子頁面
    "nlscmaps.py": "land_use",
    "urbangis.py": "land_use",
    "luz.py": "land_use",
    # 🔥 使用執照查詢子頁面
    "nlma.py": "building_license",
    "nlma_nbupic.py": "building_license",
}

# 🔥 主頁面排列（新順序）
# 左邊 4 個：地籍便民系統、謄本與土增稅、使用執照查詢、重要設施查詢
start_land_button.pack(pady=3, anchor="e")
transcript_button.pack(pady=3, anchor="e")  # 🔥 移至左側
start_land_use_button.pack(pady=3, anchor="e")
start_important_facility_button.pack(pady=3, anchor="e")

# 中間 4 個 - 🔥 土地使用分區取代國土測繪圖資，V523 系統移至此處
land_use_zone_button.pack(pady=3, anchor="center")
start_V523_button.pack(pady=3, anchor="center")
start_land_partition_button.pack(pady=3, anchor="center")
start_mapping_button.pack(pady=3, anchor="center")

# 右邊 4 個：實價登錄、編輯物件資料、連續模式、功能分頁
real_estate_price_button.pack(pady=3, anchor="w")
edit_object_data_button.pack(pady=3, anchor="w")  # 🔥 編輯物件資料移至主頁右側
continuous_mode_button.pack(pady=3, anchor="w")
utility_button.pack(pady=3, anchor="w")  # 🔥 功能分頁按鈕

# 實價登錄頁面排列
interior_button.pack(pady=3, anchor="e")
price104_button.pack(pady=3, anchor="center")
real_estate_back_button.pack(pady=3, anchor="w")

# 🔥 功能分頁排列
# 左欄：載入案件、讀取結構化JSON
utility_load_data_button.pack(pady=3, anchor="e")
read_json_button.pack(pady=3, anchor="e")
# 中欄：匯入封存、匯出封存
utility_import_archive_button.pack(pady=3, anchor="center")
utility_export_archive_button.pack(pady=3, anchor="center")
# 右欄：修改帳密、開啟工作資料夾、回至主頁
utility_credential_button.pack(pady=3, anchor="w")
utility_open_folder_button.pack(pady=3, anchor="w")
utility_back_button.pack(pady=3, anchor="w")

# 🔥 土地使用分區分頁排列
# 左欄：國土測繪圖資、高雄都市計畫
land_use_nlscmaps_button.pack(pady=3, anchor="e")
land_use_urbangis_button.pack(pady=3, anchor="e")
# 中欄：全國土地使用分區
land_use_luz_button.pack(pady=3, anchor="center")
# 右欄：回至主頁
land_use_back_button.pack(pady=3, anchor="w")

# 🔥 使用執照查詢分頁排列
# 左欄：高雄市
building_license_kh_button.pack(pady=3, anchor="e")
# 中欄：屏東縣、台南市
building_license_pt_button.pack(pady=3, anchor="center")
building_license_tn_button.pack(pady=3, anchor="center")
# 右欄：回主選單
building_license_back_button.pack(pady=3, anchor="w")

# 謄本專區頁面排列
# 【左欄 - 資料來源】
hinet_button.pack(pady=3, anchor="w")
qpt_hinet_button.pack(pady=3, anchor="w")
structured_transcript_button.pack(pady=3, anchor="w")

# 【中欄 - 資料處理與土增稅】
# 🔥 讀取結構化JSON 已移到「功能分頁」左欄
finance_button.pack(pady=3, anchor="center")
kaohsiung_tax_button.pack(pady=3, anchor="center")

# 【右欄 - 設定/編輯/返回】
# config_manage_button.pack(pady=3, anchor="e")  # 🔥 已註釋
cadastral_map_tool_button.pack(pady=3, anchor="e")  # 🔥 地籍圖處理工具
transcript_back_button.pack(pady=3, anchor="e")

def start_subprocess(button, script_name):
    global is_running, process, selected_json_path, json_dir, current_running_button

    # 🔥 設定當前執行的按鈕（用於追蹤執行狀態和更新標題）
    current_running_button = button

    try:
        # 🔥 取得腳本的正確路徑（支援 PyInstaller 打包）
        script_path = get_resource_path(script_name)

        # 🔥 尋找最新版本的電子謄本結構化（與連續模式共用邏輯）
        def find_latest_transcript_exe_local(base_dir):
            """在指定目錄尋找最新版本的電子謄本結構化.exe"""
            import glob
            import re

            pattern = os.path.join(base_dir, "電子謄本結構化*.exe")
            exe_files = glob.glob(pattern)
            pattern_subdir = os.path.join(base_dir, "電子謄本結構化*", "電子謄本結構化*.exe")
            exe_files += glob.glob(pattern_subdir)

            if not exe_files:
                return None

            def extract_version(filepath):
                filename = os.path.basename(filepath)
                match = re.search(r'[v_](\d+)\.(\d+)(?:\.(\d+))?([a-zA-Z])?', filename)
                if match:
                    major = int(match.group(1))
                    minor = int(match.group(2))
                    patch = int(match.group(3)) if match.group(3) else 0
                    suffix = match.group(4).lower() if match.group(4) else ''
                    suffix_val = ord(suffix) if suffix else 0
                    return (major, minor, patch, suffix_val)
                return (0, 0, 0, 0)

            exe_files.sort(key=extract_version, reverse=True)
            return exe_files[0]

        # 🔥 修正：在打包環境中使用嵌入式 Python 執行子程式
        if getattr(sys, 'frozen', False):
            # 🔥 特殊處理：GUI_transcript_pdf 是獨立的 exe（自動尋找最新版本）
            if script_name == "GUI_transcript_pdf 1141021-01.py":
                base_dir = os.path.dirname(sys.executable)
                latest_exe = find_latest_transcript_exe_local(base_dir)

                if latest_exe:
                    cmd = [latest_exe]
                    update_message(f"啟動電子謄本結構化：{os.path.basename(latest_exe)}")
                else:
                    update_message("[錯誤] 找不到電子謄本結構化.exe")
                    is_running = False
                    enable_all_buttons()
                    return
            else:
                # 打包環境：使用 python_embedded 資料夾中的 Python
                # 🔥 使用 pythonw.exe（不顯示 CMD 視窗）
                # 🔥 hinet.py 和 qpt_hinet.py 已改用 GUI 對話框，不需要 CMD
                python_embedded = os.path.join(os.path.dirname(sys.executable), 'python_embedded', 'pythonw.exe')
                if os.path.exists(python_embedded):
                    cmd = [python_embedded, script_path]
                else:
                    # 如果找不到嵌入式 Python，顯示錯誤訊息
                    update_message("[錯誤] 找不到 Python 執行環境，請確認打包是否完整")
                    is_running = False
                    enable_all_buttons()
                    return
        else:
            # 開發環境：使用系統 Python
            cmd = [sys.executable, script_path]

        # 🔥 傳遞基準目錄給子程式（data.json 和工作資料夾的位置）
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        cmd += ['--base-dir', base_dir]

        # 🔥 nlma_nbupic.py：依被點擊的按鈕帶入對應的 --city（屏東縣 / 臺南市 / 未來其他 NBUPIC 縣市）
        if script_name == "nlma_nbupic.py":
            city_for_btn = NLMA_NBUPIC_CITY_FOR_BUTTON.get(id(button))
            if city_for_btn:
                cmd += ['--city', city_for_btn]
                update_message(f"[NBUPIC] 將查詢「{city_for_btn}」")

        # 🔥 全國土地使用分區：開多選對話框讓使用者勾選要執行哪幾筆
        if script_name == "luz.py":
            try:
                data_json_path = get_data_json_path()
                if os.path.exists(data_json_path):
                    with open(data_json_path, 'r', encoding='utf-8') as f:
                        luz_data = json.load(f)
                    if isinstance(luz_data, list) and len(luz_data) > 1:
                        intro = (
                            f"偵測到 data.json 共 {len(luz_data)} 筆地號。\n"
                            f"請勾選要執行的筆次（預設全選）。可用下方按鈕或快捷鍵：\n"
                            f"  全選(A)  全不選(N)  只第一筆(F)  Enter=執行  Esc=取消"
                        )
                        selected_indices = show_lot_selection_dialog(
                            "【土地使用分區】選擇要執行的地號",
                            intro,
                            luz_data
                        )
                        if selected_indices is None:
                            update_message("[已取消] 未啟動【全國土地使用分區】")
                            is_running = False
                            original_text = button.cget("text").replace("關閉 ", "")
                            button.config(text=original_text)
                            enable_all_buttons()
                            return
                        if len(selected_indices) == 0:
                            update_message("[已取消] 未選任何筆，未啟動")
                            is_running = False
                            original_text = button.cget("text").replace("關閉 ", "")
                            button.config(text=original_text)
                            enable_all_buttons()
                            return

                        idx_str = ",".join(str(i) for i in selected_indices)
                        cmd += ['--indices', idx_str]
                        if len(selected_indices) == len(luz_data):
                            update_message(f"[全部] 將執行全部 {len(luz_data)} 筆")
                        elif len(selected_indices) == 1:
                            only_one = luz_data[selected_indices[0]]
                            update_message(
                                f"[單筆] 只執行第 {selected_indices[0]+1} 筆："
                                f"{only_one.get('area','')}{only_one.get('section','')} {only_one.get('lot_number','')}"
                            )
                        else:
                            update_message(f"[自選] 將執行 {len(selected_indices)} / {len(luz_data)} 筆")
            except Exception as e:
                update_message(f"[警告] 讀取 data.json 判斷查詢範圍失敗：{e}")

        # 🔥 地籍便民系統：若 data.json 已有資料，詢問是否接續之前的查詢
        if script_name == "land.py":
            try:
                data_json_path = get_data_json_path()
                if os.path.exists(data_json_path):
                    with open(data_json_path, 'r', encoding='utf-8') as f:
                        existing_data = json.load(f)
                    if isinstance(existing_data, list) and existing_data:
                        first = existing_data[0]
                        last = existing_data[-1]
                        first_desc = f"{first.get('city','')} {first.get('area','')}{first.get('section','')} {first.get('lot_number','')}"
                        last_desc = f"{last.get('city','')} {last.get('area','')}{last.get('section','')} {last.get('lot_number','')}"
                        count = len(existing_data)

                        # 只有一筆時，第一筆等於最後一筆，不重複顯示
                        if count == 1:
                            detail_text = f"目前僅有 1 筆：{first_desc}"
                        else:
                            detail_text = (
                                f"共 {count} 筆\n"
                                f"第一筆（決定資料夾名稱）：{first_desc}\n"
                                f"最後一筆（上次停在此）：{last_desc}"
                            )

                        choice = show_large_yesnocancel(
                            "接續之前的查詢？",
                            f"偵測到 data.json 已有紀錄：\n\n"
                            f"{detail_text}\n\n"
                            f"接續查詢 → 新資料會加在最後一筆之後\n"
                            f"重新開始 → data.json 將被新的第一筆取代\n"
                            f"取消操作 → 不啟動地籍便民系統",
                            yes_text="接續查詢",
                            no_text="重新開始",
                            cancel_text="取消操作"
                        )
                        if choice is None:
                            # 取消 → 不啟動 land.py，還原按鈕文字
                            update_message("[取消] 使用者取消啟動地籍便民系統")
                            is_running = False
                            # 還原按鈕文字（去掉 toggle_buttons 在啟動前加的「關閉 」前綴）
                            if button:
                                try:
                                    cur_text = button.cget("text")
                                    if cur_text.startswith("關閉 "):
                                        restored = cur_text.replace("關閉 ", "", 1)
                                        # 地籍便民系統有特殊橙色樣式
                                        if button == start_land_button:
                                            button.config(text=restored, **priority_button_style)
                                        else:
                                            button.config(text=restored)
                                except Exception:
                                    pass
                            if continuous_mode_active:
                                if current_running_button:
                                    current_running_button = None
                                schedule_next_script()
                            else:
                                enable_all_buttons()
                            return
                        elif choice is True:
                            cmd += ['--resume']
                            update_message(f"[接續模式] 保留原有 {count} 筆（上次停在 {last_desc}）")
                        else:
                            update_message("[重新開始] 新的第一筆查詢會取代既有 data.json")
            except Exception as e:
                update_message(f"[警告] 讀取 data.json 判斷接續狀態失敗：{e}")

        # 稅務腳本帶入 JSON 參數
        if script_name in ('kaohsiung_tax.py', 'finance_tax.py'):
            if selected_json_path:
                cmd += ['--json', selected_json_path]
            else:
                cmd += ['--json-dir', json_dir or JSON_DIR_DEFAULT]

        # 🔥 設定環境變數以隔離 Python 路徑
        env = os.environ.copy()
        # 🔥 強制子程式 stdout/stderr 用 UTF-8（避免 Windows CP950 與 Popen encoding='utf-8' 不一致）
        env['PYTHONIOENCODING'] = 'utf-8'
        cwd = None
        if getattr(sys, 'frozen', False):
            # 打包環境：設定環境變數只使用 python_embedded
            python_embedded_dir = os.path.join(os.path.dirname(sys.executable), 'python_embedded')
            internal_dir = os.path.join(os.path.dirname(sys.executable), '_internal')
            env['PYTHONHOME'] = python_embedded_dir
            # 🔥 設定 PYTHONPATH 讓子程式能 import：
            # 1. _internal 中的模組（base_dir_helper等）
            # 2. python_embedded 中的套件（PyPDF2等）
            python_site_packages = os.path.join(python_embedded_dir, 'Lib', 'site-packages')
            env['PYTHONPATH'] = internal_dir + os.pathsep + python_site_packages
            # 禁用用戶 site-packages
            env['PYTHONNOUSERSITE'] = '1'
            # 🔥 設定工作目錄為 _internal，讓子程式能找到同目錄的模組
            # 雖然 cwd 在 _internal，但 --base-dir 參數會告訴子程式主程式目錄在哪
            cwd = internal_dir
            print(f"[DEBUG] 子程式工作目錄: {cwd}")
            print(f"[DEBUG] 子程式 BASE_DIR 參數: {base_dir}")

        # 🔥 隱藏子程式的 CMD 視窗（Windows）
        startupinfo = None
        if sys.platform == 'win32':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

        process = subprocess.Popen(
            cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace',  # errors='replace' 容錯：非 UTF-8 byte 變 ?
            bufsize=1, universal_newlines=True, env=env, cwd=cwd,
            startupinfo=startupinfo
        )
        is_running = True
        threading.Thread(target=read_output, args=(process,), daemon=True).start()
    except Exception as e:
        update_message(f"[錯誤] 啟動程序失敗: {str(e)}")
        is_running = False
        enable_all_buttons()

def exit_subprocess(button):
    global is_running, process, current_running_button
    # 🔥 先將 is_running 設為 False，讓 read_output 線程停止讀取新輸出
    is_running = False

    # 🔥 手動關閉時也要隱藏進度條
    hide_task_progress()

    if process:
        # update_message("正在終止該程序及其相關進程...")
        # 🔥 給 read_output 線程一點時間來停止
        time.sleep(0.1)
        kill_subprocess_and_children(process)

        # 🔥 顯示手動關閉的結束訊息
        update_message("🔧 子程序已結束，返回碼: 手動關閉")
        update_message("")
        update_divider("＊")
        update_message("")

        # 🔥 如果是 land.py 手動關閉，也要更新視窗標題
        if current_running_button == start_land_button:
            root.after(100, update_title_from_data_json)

        try:
            downloads_dir = os.path.join(os.getcwd(), "downloads")
            if os.path.exists(downloads_dir):
                for file in os.listdir(downloads_dir):
                    if file.startswith("AnalysisReport") and file.endswith(".pdf"):
                        try:
                            os.remove(os.path.join(downloads_dir, file))
                        except Exception:
                            pass
        except Exception:
            pass
    original_text = button.cget("text").replace("關閉 ", "")

    # 🔥 根據是否在連續模式來決定按鈕狀態
    if continuous_mode_active:
        # 連續模式：恢復按鈕文字並禁用，然後跳到下一個腳本
        button.config(text=original_text, state=tk.DISABLED)
        current_running_button = None
        update_message("程序已終止，將跳過此腳本並執行下一個...")
        # 觸發下一個腳本
        schedule_next_script()
    else:
        # 非連續模式：正常啟用所有按鈕
        button.config(text=original_text)
        enable_all_buttons()
        # update_message("程序已完全終止，按鈕狀態已重置")

def _format_duration(seconds):
    """秒數轉可讀格式：1h 23m / 45m 12s / 30s"""
    seconds = max(0, int(seconds))
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    if h > 0:
        return f"{h}h {m:02d}m"
    if m > 0:
        return f"{m}m {s:02d}s"
    return f"{s}s"

def show_task_progress():
    """顯示進度條（插到訊息區上方）"""
    try:
        task_progress_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(2, 2), before=frame_middle)
    except Exception:
        pass

def hide_task_progress():
    """隱藏進度條"""
    try:
        task_progress_frame.pack_forget()
    except Exception:
        pass

def update_task_progress(current, total, elapsed_sec, eta_sec):
    """根據子程式送來的數字更新進度條"""
    try:
        if total <= 0:
            return
        pct = (current / total) * 100
        task_progress_bar['value'] = pct
        elapsed_str = _format_duration(elapsed_sec)
        eta_str = _format_duration(eta_sec) if eta_sec > 0 else "估算中"
        task_progress_label.config(
            text=f"進度：{current}/{total} ({pct:.1f}%)    已跑 {elapsed_str}    預估剩餘 {eta_str}"
        )
    except Exception as e:
        print(f"[進度更新錯誤] {e}")

def handle_progress_message(current, total, elapsed, eta):
    """處理子程式傳來的 __PROGRESS__ 訊號"""
    if current == 0:
        show_task_progress()
    update_task_progress(current, total, elapsed, eta)
    # 完成後 3 秒自動隱藏
    if total > 0 and current >= total:
        root.after(3000, hide_task_progress)

def update_divider(char="=", margin=2):
    """依 message_box 當前寬度自動填滿一行的分隔線。
    避免硬寫死「'=' * 39」在不同 DPI/視窗寬度下換行。
    """
    global message_box
    try:
        import tkinter.font as _tkfont
        # 🔥 先強制 tkinter 完成 idle 任務，盡量讓 widget 已 rendered，winfo_width 才會正確
        try:
            message_box.update_idletasks()
        except Exception:
            pass
        f = _tkfont.Font(font=message_box.cget('font'))
        char_w = f.measure(char) or 8
        pixel_w = message_box.winfo_width()
        if pixel_w <= 1:
            # 視窗還沒完成佈局 → 用合理預設值（避免用 cget('width') 拿到預設的 80）
            cnt = 30
        else:
            cnt = max(10, pixel_w // char_w - margin)
        update_message(char * cnt)
    except Exception:
        update_message(char * 30)


def update_centered_divider(text, char="=", margin=2):
    """置中標題：「===   財政部土地增值稅試算自動化   ===」依視窗寬度自動填滿一行。"""
    global message_box
    try:
        import tkinter.font as _tkfont
        f = _tkfont.Font(font=message_box.cget('font'))
        char_w = f.measure(char) or 8
        text_pixel_w = f.measure(text)
        pixel_w = message_box.winfo_width()
        if pixel_w <= 1:
            update_message(text)
            return
        margin_px = char_w * margin
        spacer_px = char_w * 6  # 文字兩側各空 3 個字寬
        side_pixels = max(0, (pixel_w - text_pixel_w - spacer_px - margin_px) // 2)
        side_chars = max(1, side_pixels // char_w)
        line = char * side_chars + "   " + text + "   " + char * side_chars
        update_message(line)
    except Exception:
        update_message(text)


def update_message(message):
    global message_box, input_field
    parts = re.split(r'(\033\[[0-9;]*m)', message)
    current_fg = 'white'; current_bg = 'black'
    message_box.config(state=tk.NORMAL)
    # 常用 256-color 對照（橘色等）
    _XTERM256 = {
        208: 'orange', 202: '#FF5F00', 196: '#FF0000',
        46: '#00FF00', 21: '#0000FF', 51: '#00FFFF',
        201: '#FF00FF', 226: '#FFFF00',
    }
    for part in parts:
        if part.startswith('\033['):
            codes = part[2:-1].split(';')
            # 🔥 支援 38;5;N（256-color foreground）與 48;5;N（background）
            i = 0
            while i < len(codes):
                code = codes[i]
                if code.isdigit():
                    code_int = int(code)
                    # 256-color 序列：38;5;N 或 48;5;N
                    if code_int in (38, 48) and i + 2 < len(codes) and codes[i+1] == '5' and codes[i+2].isdigit():
                        color = _XTERM256.get(int(codes[i+2]), 'white')
                        if code_int == 38:
                            current_fg = color
                        else:
                            current_bg = color
                        i += 3
                        continue
                    if 30 <= code_int <= 37 or 90 <= code_int <= 97:
                        current_fg = ANSI_COLOR_MAP.get(str(code_int), 'white')
                    elif 40 <= code_int <= 47:
                        current_bg = ANSI_COLOR_MAP.get(str(code_int), 'black')
                    elif code_int == 0:
                        current_fg = 'white'; current_bg = 'black'
                i += 1
        else:
            tag_name = f"{current_fg}_{current_bg}"
            if tag_name not in message_box.tag_names():
                message_box.tag_configure(tag_name, foreground=current_fg, background=current_bg)
            message_box.insert(tk.END, part, tag_name)
    if not message.endswith("\n"):
        message_box.insert(tk.END, "\n")
    message_box.yview(tk.END)
    message_box.config(state=tk.DISABLED)
    message_box.update_idletasks()

    # 🔥 自動將焦點移到輸入框（當檢測到需要輸入的關鍵字時）
    input_keywords = ['請輸入', '輸入', '請選擇', '請問', '選擇', '請提供',
                      '請確認', '確認', '(Y/N)', '(y/n)', '[Y/N]', '[y/n]',
                      '請按', 'Enter', '繼續', '是否']
    if any(keyword in message for keyword in input_keywords):
        try:
            # 清除提示文字
            if input_field.get() == "請在此輸入...":
                input_field.delete(0, "end")
                input_field.config(fg="black")
            # 🔥 強制將主視窗和輸入框獲得焦點
            # 先將主視窗置頂並獲得焦點
            root.lift()
            root.focus_force()
            # 再將焦點強制設定到輸入框（使用 focus_force 而非 focus_set）
            input_field.focus_force()
        except Exception:
            pass  # 如果發生錯誤（例如視窗已關閉），靜默忽略

# 🔥🔥🔥 在這裡加入新函數 🔥🔥🔥
def show_large_message(title, message):
    """顯示大型訊息對話框（支援捲動）"""
    msg_window = tk.Toplevel()
    msg_window.title(title)

    # 🔥 根據 DPI 動態調整視窗尺寸
    width = int(700 / dpi_scale * 1.2)

    # 🔥 根據內容行數自適應高度，設定合理的上下限
    line_count = message.count('\n') + 1
    # 每行約 30 像素（含字級間距），加上標題、按鈕、padding 約 220 像素
    estimated_height = int((line_count * 30 + 220) / dpi_scale * 1.2)
    min_height = int(380 / dpi_scale * 1.2)
    max_height = int(800 / dpi_scale * 1.2)  # 最高 800 像素，避免內容被切
    height = max(min_height, min(estimated_height, max_height))

    msg_window.geometry(f"{width}x{height}")
    msg_window.resizable(True, True)  # 🔥 允許調整大小
    msg_window.transient(root)  # 🔥 修復：指定父視窗
    msg_window.grab_set()

    # 🔥 置中顯示
    msg_window.update_idletasks()
    x = (msg_window.winfo_screenwidth() - width) // 2
    y = (msg_window.winfo_screenheight() - height) // 2
    msg_window.geometry(f"{width}x{height}+{x}+{y}")

    # 🔥 標題（根據 DPI 調整字體）
    title_font_size = max(14, int(22 / dpi_scale * 1.2))
    title_label = tk.Label(msg_window, text=title,
                          font=("Microsoft JhengHei", title_font_size, "bold"),
                          fg="#1976D2",
                          pady=20)
    title_label.pack()

    # 🔥 先放按鈕框架在底部（確保按鈕一定顯示）
    button_frame = tk.Frame(msg_window)
    button_frame.pack(side=tk.BOTTOM, pady=20)

    button_font_size = max(12, int(14 / dpi_scale * 1.2))
    ok_btn = tk.Button(button_frame,
                      text="確定",
                      command=msg_window.destroy,
                      font=("Microsoft JhengHei", button_font_size, "bold"),
                      bg='#4CAF50',
                      fg='white',
                      width=15,
                      height=2,
                      cursor="hand2")
    ok_btn.pack()

    # 🔥 訊息內容（使用 ScrolledText 支援捲動）- 填充剩餘空間
    content_frame = tk.Frame(msg_window)
    content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))

    # 🔥 使用 ScrolledText 取代 Label，根據 DPI 調整字體
    content_font_size = max(11, int(13 / dpi_scale * 1.2))
    text_widget = scrolledtext.ScrolledText(
        content_frame,
        font=("Microsoft JhengHei", content_font_size),
        wrap=tk.WORD,
        bg="#F5F5F5",
        relief=tk.FLAT,
        padx=15,  # 🔥 從 10 改為 15
        pady=15   # 🔥 從 10 改為 15
    )
    text_widget.pack(fill=tk.BOTH, expand=True)
    text_widget.insert("1.0", message)
    text_widget.config(state=tk.DISABLED)  # 設為唯讀

    # 🔥 讓按鈕可以用 Enter 鍵關閉
    msg_window.bind('<Return>', lambda e: msg_window.destroy())
    msg_window.bind('<Escape>', lambda e: msg_window.destroy())

    msg_window.focus()

def show_large_yesno(title, message):
    """顯示大字體的是/否對話框

    參數:
        title: 視窗標題
        message: 訊息內容

    回傳:
        True: 使用者點選「是」
        False: 使用者點選「否」
    """
    result = [False]  # 用列表儲存結果，以便在內部函數修改

    dialog = tk.Toplevel(root)
    dialog.title(title)

    # 🔥 根據訊息長度動態調整視窗大小
    msg_lines = message.count('\n') + 1
    base_height = 150
    line_height = 30
    button_height = 100
    dialog_height = min(base_height + (msg_lines * line_height) + button_height, 600)

    # 置中顯示
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width - 780) // 2
    y = (screen_height - dialog_height) // 2
    dialog.geometry(f"780x{dialog_height}+{x}+{y}")

    dialog.transient(root)
    dialog.grab_set()

    # 🔥 支援 '<HR>' 標記插入真正的水平分隔線（貼齊彈窗寬度）
    if '<HR>' in message:
        parts = message.split('<HR>')
        for i, part in enumerate(parts):
            tk.Label(dialog, text=part.strip('\n'),
                     font=("Microsoft JhengHei", 16),
                     wraplength=750,
                     justify=tk.LEFT,
                     padx=20, pady=8).pack(fill=tk.X)
            if i < len(parts) - 1:
                tk.Frame(dialog, height=1, bg='#888').pack(fill=tk.X, padx=20, pady=4)
    else:
        # 訊息文字（大字體）- 🔥 不使用 expand，固定高度
        msg_label = tk.Label(dialog, text=message,
                            font=("Microsoft JhengHei", 16),
                            wraplength=750,
                            justify=tk.LEFT,
                            padx=20, pady=20)
        msg_label.pack(fill=tk.X)

    # 按鈕區 - 🔥 固定在底部
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(side=tk.BOTTOM, pady=20)

    def on_yes():
        result[0] = True
        dialog.destroy()

    def on_no():
        result[0] = False
        dialog.destroy()
    
    # 是按鈕
    yes_btn = tk.Button(btn_frame, text="是(Y)", 
                       font=("Microsoft JhengHei", 14),
                       command=on_yes,
                       width=10, height=2,
                       bg='#4CAF50', fg='white')
    yes_btn.pack(side=tk.LEFT, padx=10)
    
    # 否按鈕
    no_btn = tk.Button(btn_frame, text="否(N)", 
                      font=("Microsoft JhengHei", 14),
                      command=on_no,
                      width=10, height=2,
                      bg='#F44336', fg='white')
    no_btn.pack(side=tk.LEFT, padx=10)
    
    # 鍵盤快捷鍵
    dialog.bind('y', lambda e: on_yes())
    dialog.bind('Y', lambda e: on_yes())
    dialog.bind('n', lambda e: on_no())
    dialog.bind('N', lambda e: on_no())
    dialog.bind('<Return>', lambda e: on_yes())  # Enter = 是
    dialog.bind('<Escape>', lambda e: on_no())   # Esc = 否
    
    dialog.focus_set()
    dialog.wait_window()

    return result[0]


def show_record_filter_dialog(transcript_data):
    """彈出勾選對話框：讓使用者選擇要載入的記錄與所有權人。

    返回:
        - 已過濾的 transcript_data list（使用者確認載入的部分）
        - None：使用者取消（不應載入）
    """
    import tkinter as tk
    from tkinter import ttk

    if not transcript_data:
        return transcript_data

    dialog = tk.Toplevel(root)
    dialog.title("選擇要載入的資料")
    dialog.geometry("960x780")
    dialog.transient(root)
    dialog.grab_set()
    dialog.update_idletasks()
    sw = dialog.winfo_screenwidth()
    sh = dialog.winfo_screenheight()
    dialog.geometry(f"960x780+{(sw - 960) // 2}+{(sh - 780) // 2}")

    result = {"data": None}

    # 標題
    tk.Label(dialog, text="選擇要載入的記錄與所有權人",
             font=("微軟正黑體", 17, "bold"), fg="#1976D2").pack(pady=(18, 4))
    tk.Label(dialog, text="未勾選的項目不會載入；原 JSON 檔案不會被修改",
             font=("微軟正黑體", 12), fg="#888").pack(pady=(0, 10))

    # 控制列
    ctrl = tk.Frame(dialog)
    ctrl.pack(fill=tk.X, padx=18, pady=(0, 6))

    # 可捲動內容區
    body = tk.Frame(dialog, bd=1, relief=tk.SOLID)
    body.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

    canvas = tk.Canvas(body, highlightthickness=0)
    sb = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas)
    inner_window = canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=sb.set)
    canvas.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    def _on_inner_resize(_e):
        canvas.configure(scrollregion=canvas.bbox("all"))
    inner.bind("<Configure>", _on_inner_resize)

    def _on_canvas_resize(e):
        canvas.itemconfigure(inner_window, width=e.width)
    canvas.bind("<Configure>", _on_canvas_resize)

    def _on_mousewheel(e):
        canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
    canvas.bind("<Enter>", lambda _: canvas.bind_all("<MouseWheel>", _on_mousewheel))
    canvas.bind("<Leave>", lambda _: canvas.unbind_all("<MouseWheel>"))

    record_vars = []
    owner_vars = []  # 與 record_vars 同長度，每個 element 是 list of BooleanVar
    # 🔥 父子雙向同步用的旗標，避免 trace 互相觸發造成無窮迴圈
    _syncing = {"flag": False}

    for ri, record in enumerate(transcript_data):
        type_ = record.get("謄本類型", "")
        basic = record.get("基本資訊", {}) or {}
        lot_or_building = basic.get("地號建號", "")

        rvar = tk.BooleanVar(value=True)
        record_vars.append(rvar)

        if "土地" in type_:
            icon = "🟫"
            bg = "#FFF9E7"
        elif "建物" in type_:
            icon = "🏢"
            bg = "#E3F2FD"
        else:
            icon = "📄"
            bg = "#F5F5F5"

        rframe = tk.Frame(inner, bg=bg)
        rframe.pack(fill=tk.X, padx=4, pady=(8, 0))
        tk.Checkbutton(rframe, text=f"{icon}  {type_}  ｜  {lot_or_building}",
                       variable=rvar, bg=bg,
                       font=("微軟正黑體", 14, "bold"),
                       anchor="w").pack(fill=tk.X, padx=4, pady=6)

        owners = record.get("所有權部", []) or []
        this_owner_vars = []
        for owner in owners:
            ovar = tk.BooleanVar(value=True)
            this_owner_vars.append(ovar)
            name = owner.get("所有權人", "")
            ratio = owner.get("權利範圍", "")
            seq = owner.get("登記次序", "")
            other = (owner.get("其他登記事項", "") or "").strip()
            # 只保留有意義的「其他登記事項」（過濾「（空白）」）
            if other in ("", "（空白）", "(空白)"):
                other = ""
            # 主行：姓名（權利範圍）｜登記次序 XXXX
            head = f"{name}（{ratio}）" if ratio else name
            if seq:
                head += f"  ｜登記次序 {seq}"
            # 子層加左邊內距讓勾選框視覺上比父層內縮
            tk.Checkbutton(inner, text=f"└─  {head}",
                           variable=ovar,
                           font=("微軟正黑體", 12), fg="#333",
                           anchor="w").pack(fill=tk.X, padx=(48, 4))
            # 副行：其他登記事項（再內縮一些，與所有權人主行對齊）
            if other:
                shown = other if len(other) <= 80 else (other[:77] + "...")
                tk.Label(inner, text=shown,
                         font=("微軟正黑體", 11), fg="#1976D2",
                         anchor="w", justify="left").pack(fill=tk.X, padx=(96, 4))
        owner_vars.append(this_owner_vars)

        # 🔥 父子雙向連動（共用 _syncing 旗標避免迴圈）
        def _make_parent_to_children(_rvar, _ovars):
            def _cb(*_args):
                if _syncing["flag"]:
                    return
                _syncing["flag"] = True
                try:
                    v = _rvar.get()
                    for _ov in _ovars:
                        _ov.set(v)
                finally:
                    _syncing["flag"] = False
            return _cb

        def _make_children_to_parent(_rvar, _ovars):
            def _cb(*_args):
                if _syncing["flag"]:
                    return
                if not _ovars:
                    return
                _syncing["flag"] = True
                try:
                    states = [_ov.get() for _ov in _ovars]
                    # 🔥 語意：父勾選框 = 「這筆記錄是否要被載入」
                    #   任何一個子勾選 → 父必須勾起（否則整筆 record 被過濾掉，連帶其他被勾的子也沒了）
                    #   全部子未勾選 → 父也取消
                    if any(states):
                        _rvar.set(True)
                    else:
                        _rvar.set(False)
                finally:
                    _syncing["flag"] = False
            return _cb

        rvar.trace_add('write', _make_parent_to_children(rvar, this_owner_vars))
        _child_cb = _make_children_to_parent(rvar, this_owner_vars)
        for _ov in this_owner_vars:
            _ov.trace_add('write', _child_cb)

    # 按鈕區
    def _select_all():
        _syncing["flag"] = True
        try:
            for rv in record_vars:
                rv.set(True)
            for ovs in owner_vars:
                for ov in ovs:
                    ov.set(True)
        finally:
            _syncing["flag"] = False

    def _deselect_all():
        _syncing["flag"] = True
        try:
            for rv in record_vars:
                rv.set(False)
            for ovs in owner_vars:
                for ov in ovs:
                    ov.set(False)
        finally:
            _syncing["flag"] = False

    tk.Button(ctrl, text="全選", command=_select_all,
              font=("微軟正黑體", 12), padx=14, pady=4).pack(side=tk.LEFT, padx=5)
    tk.Button(ctrl, text="全不選", command=_deselect_all,
              font=("微軟正黑體", 12), padx=14, pady=4).pack(side=tk.LEFT, padx=5)
    tk.Label(ctrl, text=f"共 {len(transcript_data)} 筆記錄",
             font=("微軟正黑體", 12), fg="#666").pack(side=tk.RIGHT, padx=5)

    def _on_confirm():
        filtered = []
        for ri, record in enumerate(transcript_data):
            if not record_vars[ri].get():
                continue
            new_record = dict(record)  # shallow copy；下層 list/dict 引用即可
            if "所有權部" in new_record and isinstance(new_record["所有權部"], list):
                kept = [new_record["所有權部"][oi]
                        for oi, ov in enumerate(owner_vars[ri]) if ov.get()]
                new_record["所有權部"] = kept
            filtered.append(new_record)
        result["data"] = filtered
        dialog.destroy()

    def _on_cancel():
        # result["data"] 保持 None
        dialog.destroy()

    btns = tk.Frame(dialog)
    btns.pack(pady=(10, 18))
    tk.Button(btns, text="✓ 確認載入", command=_on_confirm,
              font=("微軟正黑體", 15, "bold"), bg="#4CAF50", fg="white",
              padx=30, pady=10, cursor="hand2").pack(side=tk.LEFT, padx=12)
    tk.Button(btns, text="✗ 取消", command=_on_cancel,
              font=("微軟正黑體", 15), bg="#f44336", fg="white",
              padx=30, pady=10, cursor="hand2").pack(side=tk.LEFT, padx=12)

    dialog.protocol("WM_DELETE_WINDOW", _on_cancel)
    dialog.bind("<Escape>", lambda _: _on_cancel())
    dialog.wait_window()

    return result["data"]


def _apply_filter_and_convert(latest_json):
    """讀檔 → 跳勾選對話框 → 把過濾結果寫到 temp → 轉換 → 清掉 temp。

    返回 (success: bool, message: str)
    """
    try:
        with open(latest_json, 'r', encoding='utf-8') as f:
            all_records = json.load(f)
    except Exception as e:
        return False, f"無法讀取謄本 JSON：{e}"

    filtered = show_record_filter_dialog(all_records)
    if filtered is None:
        return False, "使用者取消載入"
    if not filtered:
        return False, "未勾選任何記錄，不載入"

    total = len(all_records)
    kept = len(filtered)
    # 全選 → 直接用原檔，避免不必要的 temp 檔
    if kept == total and all(
        len(r.get("所有權部", [])) == len(f.get("所有權部", []))
        for r, f in zip(all_records, filtered)
    ):
        update_message(f"[載入] 全部 {kept} 筆記錄、所有權人未過濾")
        convert_transcript_to_final_data(latest_json)
        return True, "全部載入"

    # 🔥 有過濾 → 寫到該案件「4.其他相關」資料夾，作為新的選定 JSON
    # 後續 土增稅、編輯物件資料 等程序都用這個過濾後版本
    try:
        # 從 data.json 第一筆推得案件資料夾
        data_json_path = get_data_json_path()
        with open(data_json_path, 'r', encoding='utf-8') as f:
            data_list = json.load(f)
        first = data_list[0] if data_list else {}
        area = first.get('area', '')
        section = first.get('section', '')
        lot_number = first.get('lot_number', '')
        folder_name = f"{area}{section}-{lot_number}"
        other_dir = get_work_folder(os.path.join(folder_name, "4.其他相關"))
        os.makedirs(other_dir, exist_ok=True)

        # 檔名：原檔名 + _filtered 後綴
        base, ext = os.path.splitext(os.path.basename(latest_json))
        filtered_name = f"{base}_filtered{ext}"
        filtered_path = os.path.join(other_dir, filtered_name)

        with open(filtered_path, 'w', encoding='utf-8') as f:
            json.dump(filtered, f, ensure_ascii=False, indent=2)
        update_message(f"[載入] 過濾後 {kept}/{total} 筆 → 已存至 {filtered_path}")
    except Exception as e:
        return False, f"寫入過濾後 JSON 失敗：{e}"

    # convert 會把全域 selected_json_path 設成 filtered_path（這正是我們要的）
    convert_transcript_to_final_data(filtered_path)
    # 確保全域與 config 都指向 filtered_path（雙保險，避免後續流程拿到原始檔）
    global selected_json_path
    selected_json_path = filtered_path
    try:
        cfg = load_config() or {}
        cfg['selected_json'] = filtered_path
        cfg['selected_json_path'] = filtered_path
        save_config(cfg, show_message=False)
    except Exception:
        pass
    try:
        update_json_status_display()
    except Exception:
        pass
    return True, f"已載入過濾後 {kept}/{total} 筆（檔案：{filtered_name}）"


def show_large_yesnocancel(title, message, yes_text="是(Y)", no_text="否(N)", cancel_text="取消(C)"):
    """顯示大字體的是/否/取消對話框

    參數:
        title: 視窗標題
        message: 訊息內容
        yes_text/no_text/cancel_text: 自訂按鈕文字（預設「是(Y)」「否(N)」「取消(C)」）

    回傳:
        True: 使用者點選「是」
        False: 使用者點選「否」
        None: 使用者點選「取消」
    """
    result = [None]  # 用列表儲存結果

    dialog = tk.Toplevel(root)
    dialog.title(title)

    # 🔥 根據訊息長度動態調整視窗大小
    msg_lines = message.count('\n') + 1
    base_height = 150
    line_height = 30
    button_height = 100
    dialog_height = min(base_height + (msg_lines * line_height) + button_height, 600)

    # 置中顯示
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width - 780) // 2
    y = (screen_height - dialog_height) // 2
    dialog.geometry(f"780x{dialog_height}+{x}+{y}")
    
    dialog.transient(root)
    dialog.grab_set()
    
    # 訊息文字（大字體）- 🔥 不使用 expand，固定高度
    msg_label = tk.Label(dialog, text=message,
                        font=("Microsoft JhengHei", 16),
                        wraplength=750,
                        justify=tk.LEFT,
                        padx=20, pady=20)
    msg_label.pack(fill=tk.X)

    # 按鈕區 - 🔥 固定在底部
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(side=tk.BOTTOM, pady=20)
    
    def on_yes():
        result[0] = True
        dialog.destroy()
    
    def on_no():
        result[0] = False
        dialog.destroy()
    
    def on_cancel():
        result[0] = None
        dialog.destroy()
    
    # 動態寬度：依按鈕文字長度估計（含 padding）
    _btn_w = max(8, max(len(yes_text), len(no_text), len(cancel_text)) + 2)

    # 是按鈕
    yes_btn = tk.Button(btn_frame, text=yes_text,
                       font=("Microsoft JhengHei", 14),
                       command=on_yes,
                       width=_btn_w, height=2,
                       bg='#4CAF50', fg='white')
    yes_btn.pack(side=tk.LEFT, padx=10)

    # 否按鈕
    no_btn = tk.Button(btn_frame, text=no_text,
                      font=("Microsoft JhengHei", 14),
                      command=on_no,
                      width=_btn_w, height=2,
                      bg='#F44336', fg='white')
    no_btn.pack(side=tk.LEFT, padx=10)

    # 取消按鈕
    cancel_btn = tk.Button(btn_frame, text=cancel_text,
                          font=("Microsoft JhengHei", 14),
                          command=on_cancel,
                          width=_btn_w, height=2,
                          bg='#9E9E9E', fg='white')
    cancel_btn.pack(side=tk.LEFT, padx=10)
    
    # 鍵盤快捷鍵
    dialog.bind('y', lambda e: on_yes())
    dialog.bind('Y', lambda e: on_yes())
    dialog.bind('n', lambda e: on_no())
    dialog.bind('N', lambda e: on_no())
    dialog.bind('c', lambda e: on_cancel())
    dialog.bind('C', lambda e: on_cancel())
    dialog.bind('<Return>', lambda e: on_yes())    # Enter = 是
    dialog.bind('<Escape>', lambda e: on_cancel()) # Esc = 取消
    
    # 點擊關閉視窗 = 取消
    dialog.protocol("WM_DELETE_WINDOW", on_cancel)

    dialog.focus_set()
    dialog.wait_window()

    return result[0]

def show_lot_selection_dialog(title, intro, entries):
    """多選對話框：讓使用者勾選要執行的地號筆次。

    參數:
        title: 視窗標題
        intro: 上方說明文字
        entries: list of dicts，每筆含 city/area/section/lot_number

    回傳:
        list of 0-based indices：使用者選的筆次
        None：使用者按取消或關閉視窗
    """
    result = {"indices": None}

    dialog = tk.Toplevel(root)
    dialog.title(title)

    # 視窗大小與置中
    screen_w = dialog.winfo_screenwidth()
    screen_h = dialog.winfo_screenheight()
    w = 720
    h = min(680, screen_h - 100)
    x = (screen_w - w) // 2
    y = (screen_h - h) // 2
    dialog.geometry(f"{w}x{h}+{x}+{y}")
    dialog.transient(root)
    dialog.grab_set()

    # 說明區
    intro_label = tk.Label(dialog, text=intro,
                          font=("Microsoft JhengHei", 12),
                          justify=tk.LEFT, anchor='w',
                          padx=15, pady=10,
                          wraplength=w - 40)
    intro_label.pack(side=tk.TOP, fill=tk.X)

    # 清單區（Listbox + Scrollbar）
    list_frame = tk.Frame(dialog)
    list_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=15, pady=5)

    scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
    listbox = tk.Listbox(list_frame,
                         selectmode=tk.MULTIPLE,
                         font=("Microsoft JhengHei", 11),
                         yscrollcommand=scrollbar.set,
                         activestyle='none')
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # 填入項目
    for i, e in enumerate(entries, start=1):
        line = f"  {i:>3}.  {e.get('city','')}  {e.get('area','')}{e.get('section','')}  {e.get('lot_number','')}"
        listbox.insert(tk.END, line)

    # 預設全選
    listbox.selection_set(0, tk.END)

    # 已選數狀態列
    status_var = tk.StringVar(value=f"已選 {len(entries)} / {len(entries)} 筆")
    status_label = tk.Label(dialog, textvariable=status_var,
                           font=("Microsoft JhengHei", 11, "bold"),
                           fg='#1976D2', anchor='w',
                           padx=15, pady=6)
    status_label.pack(side=tk.TOP, fill=tk.X)

    def _update_status(event=None):
        status_var.set(f"已選 {len(listbox.curselection())} / {len(entries)} 筆")

    listbox.bind('<<ListboxSelect>>', _update_status)

    # 按鈕區
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(side=tk.BOTTOM, pady=12)

    def select_all():
        listbox.selection_set(0, tk.END)
        _update_status()

    def select_none():
        listbox.selection_clear(0, tk.END)
        _update_status()

    def select_first():
        listbox.selection_clear(0, tk.END)
        if entries:
            listbox.selection_set(0)
            listbox.see(0)
        _update_status()

    def on_ok():
        sel = list(listbox.curselection())
        if not sel:
            if not show_large_yesno(
                "未選任何筆",
                "目前沒有勾選任何筆，按「是」視為取消、不啟動；\n按「否」回去繼續勾選。"
            ):
                return
            # 視為取消
            result["indices"] = None
        else:
            result["indices"] = sel
        dialog.destroy()

    def on_cancel():
        result["indices"] = None
        dialog.destroy()

    btn_style = {"font": ("Microsoft JhengHei", 11), "width": 11, "height": 1}

    tk.Button(btn_frame, text="全選(A)", command=select_all,
              bg='#ECEFF1', fg='black', **btn_style).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_frame, text="全不選(N)", command=select_none,
              bg='#ECEFF1', fg='black', **btn_style).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_frame, text="只第一筆(F)", command=select_first,
              bg='#ECEFF1', fg='black', **btn_style).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_frame, text="執行", command=on_ok,
              bg='#4CAF50', fg='white', **btn_style).pack(side=tk.LEFT, padx=15)
    tk.Button(btn_frame, text="取消", command=on_cancel,
              bg='#F44336', fg='white', **btn_style).pack(side=tk.LEFT, padx=3)

    # 快捷鍵：A/N/F 快速選擇，Enter=執行，Esc=取消
    dialog.bind('a', lambda e: select_all())
    dialog.bind('A', lambda e: select_all())
    dialog.bind('n', lambda e: select_none())
    dialog.bind('N', lambda e: select_none())
    dialog.bind('f', lambda e: select_first())
    dialog.bind('F', lambda e: select_first())
    dialog.bind('<Return>', lambda e: on_ok())
    dialog.bind('<Escape>', lambda e: on_cancel())
    dialog.protocol("WM_DELETE_WINDOW", on_cancel)

    dialog.focus_set()
    dialog.wait_window()

    return result["indices"]

def prefill_input_field(value):
    """
    將值預填到 GUI 輸入框中
    用於子程式自動辨識驗證碼後，預填到主程式的輸入框
    """
    global input_field
    try:
        # 清除當前內容
        input_field.delete(0, tk.END)
        # 填入預填值
        input_field.insert(0, value)
        # 設定為正常文字顏色（非灰色 placeholder）
        input_field.config(fg="black")
        # 讓輸入框獲得焦點，方便使用者確認或修改
        input_field.focus_set()
        # 選中所有文字，方便修改
        input_field.select_range(0, tk.END)
        input_field.icursor(tk.END)
        update_message(f"[自動填入] 已將驗證碼預填到輸入框: {value}")
    except Exception as e:
        update_message(f"[錯誤] 預填輸入框失敗: {e}")

def send_input(event=None):
    global process, is_running
    if not is_running:
        update_message("[警告] 沒有運行中的子程序，無法發送輸入")
        return
    if not process:
        update_message("[警告] 沒有程序實例")
        return
    poll_result = process.poll()
    if poll_result is not None:
        update_message(f"[警告] 該程序已結束")
        is_running = False
        enable_all_buttons()
        return
    try:
        user_input = input_field.get()
        # 🔥 如果是 placeholder，當作空字串
        if user_input == "請在此輸入...":
            user_input = ""
        input_field.delete(0, tk.END)
        # 🔥 重新加上 placeholder
        input_field.insert(0, "請在此輸入...")
        input_field.config(fg="gray")
        if process.stdin and not process.stdin.closed:
            process.stdin.write(user_input + '\n')
            process.stdin.flush()
            display_input = user_input if user_input.strip() else "(空白或預設)"
            update_message(f"[完成] 成功發送: {display_input}")
        else:
            stdin_status = "不存在" if not process.stdin else ("已關閉" if process.stdin.closed else "可用")
            update_message(f"[警告] 輸入管道狀態: {stdin_status}")
    except (OSError, ValueError, BrokenPipeError) as e:
        update_message(f"[錯誤] 發送輸入失敗: {str(e)}")
        update_message("程序可能已經異常結束")
        is_running = False
        enable_all_buttons()
        for button in all_buttons:
            if (button != real_estate_price_button and 
                button != transcript_button and 
                hasattr(button, 'cget') and 
                button.cget("text").startswith("關閉")):
                original_text = button.cget("text").replace("關閉 ", "")
                button.config(text=original_text)
    except Exception as e:
        update_message(f"[錯誤] 發送輸入時發生未預期的錯誤: {str(e)}")

def on_closing():
    global is_running, process
    if is_running and process:
        for button in all_buttons:
            if hasattr(button, 'cget') and button.cget("text").startswith("關閉"):
                exit_subprocess(button)
                break
    root.destroy()

input_field.bind("<Return>", send_input)
root.protocol("WM_DELETE_WINDOW", on_closing)

# 🔥 初始化 JSON 狀態顯示（含地號資訊）
if selected_json_path:
    filename = os.path.basename(selected_json_path)
    try:
        with open(selected_json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        lot_display = ""
        if isinstance(data, list) and len(data) > 0:
            first_record = data[0]
            if "基本資訊" in first_record:
                lot_info = first_record["基本資訊"].get("地號建號", "")
                lot_display = simplify_lot_number(lot_info)
        
        # 🔥 從檔名提取日期時間（格式：YYYYMMDD_HHMM）
        import re
        datetime_match = re.search(r'(\d{8}_\d{4})', filename)
        datetime_str = datetime_match.group(1) if datetime_match else ""

        if lot_display and datetime_str:
            json_status_label.config(text=f"JSON: {datetime_str} | {lot_display}", fg='#006400')
        elif lot_display:
            json_status_label.config(text=f"JSON: {lot_display}", fg='#006400')
        else:
            json_status_label.config(text=f"JSON: {datetime_str if datetime_str else filename}", fg='#006400')
    except:
        json_status_label.config(text=f"JSON: {filename}", fg='#006400')
else:
    json_status_label.config(text=f"JSON: 自動選擇（{json_dir}）", fg='#FF8C00')

# 🔥 新增：程式啟動時的提醒訊息
def show_startup_reminder():
    update_divider()
    update_message("🔥 【重要提醒】不動產營業員作業便捷工具集")
    update_divider()
    update_message("📍 首次使用請先執行【地籍便民系統】(\033[38;5;208m橙色按鈕\033[0m)")
    update_message("📁 將自動建立地段、地號資料夾作為儲存目錄")
    update_message("📋 建議執行順序：")
    update_message("   1️地籍便民系統 → 建立【資料夾】")
    update_message("   2️【謄本與土增稅】：")
    update_message("          全國地政電子謄本 → 電子謄本結構化")
    update_message("   3️土增稅查詢")
    update_message("   4️使用分區查詢")
    update_message("   5️其他查詢功能 → 依需求使用")
    update_message("   6️編輯物件資料【物件調查表】")
    update_divider()
    update_message("💡 \033[38;5;208m橙色按鈕\033[0m為優先執行功能，請注意！")
    update_message("")

# === 資料一致性檢查（啟動時自動執行）===
def check_and_warn_data_consistency():
    """在啟動時檢查資料一致性，如有問題則顯示警告視窗"""
    # 🔥 先把診斷訊息收進 buffer，正常時靜默通過，有問題時才一次秀出來
    diag_lines = []
    try:
        diag_lines.append(f"[啟動] 檢查 data.json 一致性中...")
        diag_lines.append(f"  PROGRAM_DIR = {PROGRAM_DIR}")
        _data_path = get_data_json_path()
        diag_lines.append(f"  data.json 路徑 = {_data_path}")
        diag_lines.append(f"  data.json 存在 = {os.path.exists(_data_path)}")
        diag_lines.append(f"  selected_json_path = {selected_json_path}")
    except Exception as _diag_e:
        print(f"[診斷訊息收集失敗] {_diag_e}")

    try:
        # 🔥 使用已儲存的 selected_json_path（如果有的話）
        is_consistent, warning_msg, data_info, transcript_info, is_order_issue = check_data_consistency(selected_json_path)
        try:
            diag_lines.append(f"  data_info = {data_info}")
            diag_lines.append(f"  transcript_info = {bool(transcript_info)}（True=有解析到, False=無）")
            diag_lines.append(f"  is_consistent = {is_consistent}")
            diag_lines.append(f"  warning_msg = {warning_msg}")
        except Exception:
            pass

        # 🔥 如果 check_data_consistency 沒有找到 transcript_info，但 selected_json_path 存在
        # 嘗試手動解析 selected_json_path 來構建 transcript_info（用於警告對話框）
        if not transcript_info and selected_json_path and os.path.exists(selected_json_path):
            try:
                with open(selected_json_path, 'r', encoding='utf-8') as f:
                    transcript_data = json.load(f)

                if transcript_data and len(transcript_data) > 0:
                    # 找第一筆土地謄本
                    first_land = None
                    for item in transcript_data:
                        if "土地謄本" in item.get("謄本類型", ""):
                            first_land = item
                            break

                    if first_land:
                        basic_info = first_land.get("基本資訊", {})
                        land_number = basic_info.get("地號建號", "")

                        # 解析地號（例如：湖內區大湖段 2589-0008地號）
                        import re
                        match = re.search(r'([^\s]+區|[^\s]+鄉|[^\s]+鎮|[^\s]+市)([^\s]+段)\s*(\d+-\d+)', land_number)
                        if match:
                            transcript_info = {
                                'area': match.group(1),
                                'section': match.group(2),
                                'lot_number': match.group(3),
                                'count': len([t for t in transcript_data if "土地謄本" in t.get("謄本類型", "")]),
                                'file': os.path.basename(selected_json_path),
                                'file_path': selected_json_path
                            }

                            # 🔥 重新檢查一致性
                            if data_info:
                                def normalize_lot(lot):
                                    if '-' in lot:
                                        parts = lot.split('-')
                                        return f"{int(parts[0])}-{int(parts[1])}"
                                    else:
                                        return f"{int(lot)}-0"

                                data_lot = normalize_lot(data_info['lot_number'])
                                transcript_lot = normalize_lot(transcript_info['lot_number'])

                                if not (data_info['area'] == transcript_info['area'] and
                                        data_info['section'] == transcript_info['section'] and
                                        data_lot == transcript_lot):
                                    # 不一致！判斷是順序問題還是完全不同地段
                                    is_order_issue_new = False
                                    if os.path.exists('data.json'):
                                        with open('data.json', 'r', encoding='utf-8') as f:
                                            data_list = json.load(f)
                                        for item in data_list:
                                            item_lot = normalize_lot(item.get('lot_number', ''))
                                            if (item.get('area') == transcript_info['area'] and
                                                item.get('section') == transcript_info['section'] and
                                                item_lot == transcript_lot):
                                                is_order_issue_new = True
                                                break

                                    is_consistent = False
                                    is_order_issue = is_order_issue_new
                                    if is_order_issue:
                                        warning_msg = f"""
【資料不一致警告】

發現 data.json 和謄本結構化資料的第一筆地號不同：

📄 data.json（第一筆）：
   {data_info['area']}{data_info['section']} {data_info['lot_number']}
   （共 {data_info['count']} 筆）

📄 {transcript_info['file']}（第一筆土地謄本）：
   {transcript_info['area']}{transcript_info['section']} {transcript_info['lot_number']}
   （共 {transcript_info['count']} 筆土地謄本）

⚠️ 可能的影響：
   • 土地增值稅程式會根據第一筆建立資料夾
   • 物件編輯器會根據 data.json 第一筆尋找稅額檔案
   • 這可能導致找不到正確的 tax_result.json

💡 建議：調整 data.json 的順序，使第一筆與謄本資料一致
"""
                                    else:
                                        warning_msg = f"""
【資料不一致警告】

發現 data.json 和謄本結構化資料完全不同地段：

📄 data.json（第一筆）：
   {data_info['area']}{data_info['section']} {data_info['lot_number']}
   （共 {data_info['count']} 筆）

📄 {transcript_info['file']}（第一筆土地謄本）：
   {transcript_info['area']}{transcript_info['section']} {transcript_info['lot_number']}
   （共 {transcript_info['count']} 筆土地謄本）

⚠️ 這不是順序問題！兩者毫無關聯！

💡 建議：取消修改，重新使用【地籍便民系統】建立 data.json
   這樣才能為新地段建立正確的基礎資料結構
"""
            except Exception as e:
                print(f"[警告] 解析 selected_json_path 失敗：{e}")

        # 🔥 全部檢查跑完後再決定要不要秀診斷訊息：正常時靜默通過，有問題時才一次秀出
        if (not is_consistent) or warning_msg:
            for _ln in diag_lines:
                update_message(_ln)

        # 更新視窗標題以顯示 data.json 的資訊
        if data_info:
            location = f"{data_info['area']}{data_info['section']}-{data_info['lot_number']}"
            root.title(TITLE_WITH_LOCATION.format(version=VERSION, location=location))

        # 🔥 更新綠色 JSON 標籤以顯示結構化 JSON 的資訊（用於與標題比對）
        # 優先順序：1. 本地謄本資料 2. selected_json_path 3. 提示訊息
        if transcript_info:
            filename = transcript_info.get('file', 'transcript_data.json')
            # 🔥 從檔名提取日期時間（格式：YYYYMMDD_HHMM）
            import re
            datetime_match = re.search(r'(\d{8}_\d{4})', filename)
            datetime_str = datetime_match.group(1) if datetime_match else filename

            lot_display = f"{transcript_info['area']}{transcript_info['section']}-{transcript_info['lot_number']}"
            json_status_label.config(text=f"JSON: {datetime_str} | {lot_display}", fg='#006400')
        elif selected_json_path and os.path.exists(selected_json_path):
            # 如果本地沒有謄本資料，但有已載入的 JSON，嘗試顯示它的資訊
            try:
                with open(selected_json_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list) and len(data) > 0 and "基本資訊" in data[0]:
                    lot_info = data[0]["基本資訊"].get("地號建號", "")
                    from config_manager import simplify_lot_number
                    lot_display = simplify_lot_number(lot_info)

                    import re
                    filename = os.path.basename(selected_json_path)
                    datetime_match = re.search(r'(\d{8}_\d{4})', filename)
                    datetime_str = datetime_match.group(1) if datetime_match else filename

                    json_status_label.config(text=f"JSON: {datetime_str} | {lot_display}", fg='#006400')
                else:
                    json_status_label.config(text=f"JSON: {os.path.basename(selected_json_path)}", fg='#006400')
            except:
                json_status_label.config(text=f"JSON: {os.path.basename(selected_json_path)}", fg='#006400')
        elif data_info:
            # 如果沒有謄本資料，但有 data.json，顯示提示
            json_status_label.config(text="JSON: 尚未產生結構化謄本資料", fg='#FF8C00')
        else:
            # 完全沒有資料
            json_status_label.config(text=f"JSON: 自動選擇（{json_dir}）", fg='#FF8C00')

        # 🔥 自動依 data.json 順序重排謄本（不彈窗）
        # 即使一致性檢查 OK 也要 call fix，才能處理「建物排在土地前」這類 layout 問題
        if data_info and transcript_info and (is_consistent or is_order_issue):
            success, message = fix_transcript_json_order(data_info, transcript_info)
            if success and "無需調整" not in message:
                update_divider()
                update_message(f"[OK] {message}")
                update_divider()
            elif not success:
                update_divider()
                update_message(f"[警告] 自動重排失敗：{message}")
                update_divider()
            # 「無需調整」→ 靜默，不顯示
        elif data_info and transcript_info and not is_consistent and not is_order_issue:
            # 完全不同地段，僅警告不自動處理
            update_divider()
            update_message("[警告] data.json 與謄本完全不同地段（可能是不同案件）")
            update_message(f"  data.json[0]: {data_info['area']}{data_info['section']} {data_info['lot_number']}")
            update_message(f"  謄本[0]: {transcript_info['area']}{transcript_info['section']} {transcript_info['lot_number']}")
            update_message("  → 不會自動重排；如需切換案件請手動載入正確的 data.json")
            update_divider()

    except Exception as e:
        # 檢查失敗不影響程式啟動，但把錯誤秀到訊息區方便除錯
        import traceback as _tb
        _detail = _tb.format_exc()
        print(f"資料一致性檢查失敗：{e}\n{_detail}")
        try:
            update_message(f"[警告] 啟動時讀取 data.json 失敗：{e}")
            update_message("（視窗標題不會顯示地號，但其他功能不受影響）")
        except Exception:
            pass

def update_title_from_data_json():
    """
    從 data.json 讀取資料並更新視窗標題
    當 land.py 執行完成後，即時更新標題而不需重啟程式
    """
    try:
        # 讀取 data.json
        data_json_path = get_data_json_path()
        if os.path.exists(data_json_path):
            with open(data_json_path, 'r', encoding='utf-8') as f:
                data_list = json.load(f)

            if data_list and len(data_list) > 0:
                first_data = data_list[0]
                area = first_data.get('area', '')
                section = first_data.get('section', '')
                lot_number = first_data.get('lot_number', '')

                # 更新視窗標題
                location = f"{area}{section}-{lot_number}"
                root.title(TITLE_WITH_LOCATION.format(version=VERSION, location=location))
                print(f"[標題更新] {location}")
            else:
                # data.json 存在但沒有資料
                root.title(TITLE_FORMAT.format(version=VERSION))
        else:
            # data.json 不存在
            root.title(TITLE_FORMAT.format(version=VERSION))

    except Exception as e:
        print(f"[警告] 更新視窗標題失敗：{e}")

# 延遲顯示提醒訊息和資料一致性檢查
root.after(100, show_startup_reminder)
root.after(500, check_and_warn_data_consistency)  # 在提醒訊息之後執行

# 🔥 如果啟動時已載入 JSON，更新狀態列顯示
if selected_json_path:
    root.after(200, update_json_status_display)

setup_data_editor_module()

root.mainloop()