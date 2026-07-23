# nlma 使用執照查詢
import sys
import os

# 🔥 打包環境修正：將 _internal 加入 sys.path
if getattr(sys, 'frozen', False):
    # 如果是從打包的 EXE 執行
    internal_dir = os.path.dirname(os.path.abspath(__file__))
    if internal_dir not in sys.path:
        sys.path.insert(0, internal_dir)
elif '__file__' in globals():
    # 開發環境
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)


import os
import sys
import atexit  # 添加atexit模块，用于注册程序退出时的回调函数
# 清理畫面並顯示歡迎訊息
os.system('cls')
sys.stdout.reconfigure(line_buffering=True)

# 🔥 line_buffering 只在換行才 flush；input(提示) 的提示沒換行會顯示不全 → 覆寫 input 先 flush 提示
import builtins as _builtins
_orig_input = _builtins.input
def _flushing_input(prompt=''):
    if prompt:
        try:
            sys.stdout.write(prompt)
            sys.stdout.flush()
        except Exception:
            pass
    return _orig_input()
_builtins.input = _flushing_input

print("歡迎使用【全國建築執照存根查詢】自動化小程式~\n模組載入中...", flush=True)
import json
import time
import re
import base64
import logging
import numpy as np
import ctypes
import keyboard  # 用於系統級鍵盤事件
from ctypes import wintypes
from PIL import Image
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window, apply_page_zoom

# 🔥 基準目錄設定（data.json 和工作資料夾的位置）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

# 設定日誌級別，屏蔽系統錯誤訊息
logging.getLogger("selenium").setLevel(logging.WARNING)
logging.getLogger("PIL").setLevel(logging.WARNING)

# 禁止 FutureWarning
import warnings
warnings.filterwarnings("ignore", category=FutureWarning)

# 重定向標準錯誤輸出
sys.stderr = open(os.devnull, "w")

# 全局變數存儲查詢類型描述
CURRENT_QUERY_DESCRIPTION = None
# 新增全局變數來存儲預選的查詢類型
PRESELECTED_QUERY_TYPE = None
# 添加全局變數，用於控制程式完全退出
SHOULD_EXIT_PROGRAM = False

# 初始化 driver
driver = None
# 🔥 修正：使用 BASE_DIR 而非 os.getcwd()，避免打包後路徑跑到 _internal
TMP_PDF_PATH = os.path.join(BASE_DIR, "tmp_pdf")
if not os.path.exists(TMP_PDF_PATH):
    os.makedirs(TMP_PDF_PATH)

def safe_input(prompt="", check_interval=1.0):
    """
    安全的 input 函數，會定期檢查瀏覽器是否還存活
    如果瀏覽器被關閉，會拋出異常以終止等待

    Args:
        prompt: 提示訊息
        check_interval: 檢查瀏覽器狀態的間隔（秒）
    """
    import threading
    import queue
    global driver

    # 創建一個 queue 來傳遞輸入結果
    result_queue = queue.Queue()
    input_received = threading.Event()

    def input_thread():
        """執行緒中執行 input"""
        try:
            user_input = input(prompt)
            result_queue.put(('success', user_input))
        except (KeyboardInterrupt, EOFError) as e:
            result_queue.put(('interrupt', str(e)))
        finally:
            input_received.set()

    # 啟動輸入執行緒
    thread = threading.Thread(target=input_thread, daemon=True)
    thread.start()

    # 定期檢查瀏覽器狀態，直到收到輸入或發現瀏覽器關閉
    while not input_received.is_set():
        if driver:
            try:
                # 嘗試獲取當前 URL 來檢查連接是否還存在
                _ = driver.current_url
            except Exception:
                # driver 連接已斷開（瀏覽器被關閉）
                print("\n偵測到瀏覽器已關閉，程式將退出", flush=True)
                raise SystemExit("瀏覽器已關閉")

        # 等待一小段時間後再次檢查
        input_received.wait(timeout=check_interval)

    # 獲取輸入結果
    status, value = result_queue.get()
    if status == 'interrupt':
        print("\n程式被中斷", flush=True)
        raise SystemExit("使用者中斷")

    return value

def get_available_work_area():
    SPI_GETWORKAREA = 0x0030
    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", wintypes.LONG),
            ("top", wintypes.LONG),
            ("right", wintypes.LONG),
            ("bottom", wintypes.LONG)
        ]
    rect = RECT()
    ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
    return rect

def calculate_browser_window_position():
    """從 window_config.json 算出瀏覽器視窗（邏輯像素）：2/3 物理寬 ÷ DPI，位置在主程式對側。
    與其他子程式一致；避免在高DPI機器把物理像素當邏輯像素而開太大（約 3/4）。
    回傳 (x, y, width, height)，皆為邏輯像素。"""
    bx, by, bw, bh = 0, 0, 1024, 1024
    try:
        cfg_path = os.path.join(BASE_DIR, 'window_config.json')
        if os.path.exists(cfg_path):
            with open(cfg_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            wa = cfg.get('work_area', {})
            mw = cfg.get('main_window', {})
            if wa and mw:
                dpi = cfg.get('dpi_scale') or 1.0
                work_left = wa.get('left', 0)
                work_top = wa.get('top', 0)
                work_right = wa.get('right', 1920)
                work_bottom = wa.get('bottom', 1080)
                main_x = mw.get('x', 1280)
                main_w = mw.get('width', 640)
                chrome_w_phys = (work_right - work_left) * 2 // 3      # 物理像素的 2/3
                if (main_x + main_w / 2) > (work_left + work_right) / 2:
                    chrome_x_phys = work_left                          # 主程式在右 → 瀏覽器放左
                else:
                    chrome_x_phys = work_right - chrome_w_phys         # 主程式在左 → 瀏覽器放右
                chrome_h_phys = work_bottom - work_top - 20
                if chrome_w_phys < 800:
                    chrome_w_phys = 800
                if chrome_h_phys < 600:
                    chrome_h_phys = 600
                bw = int(chrome_w_phys / dpi)                          # 🔥 除以 DPI → 邏輯像素
                bh = int(chrome_h_phys / dpi) + 32                     # +標題欄高度
                bx = int(chrome_x_phys / dpi)
                by = int(work_top / dpi)
                return bx, by, bw, bh
    except Exception as e:
        print(f"[視窗] 讀 window_config 失敗，改用螢幕推算：{e}", flush=True)
    # 後備（無設定檔時）：用工作區推算 2/3
    try:
        work_area = get_available_work_area()
        bw = (work_area.right - work_area.left) * 2 // 3
        bh = work_area.bottom - work_area.top - 20
        bx = work_area.left
        by = work_area.top
    except Exception:
        pass
    return bx, by, bw, bh


def init_driver():
    options = webdriver.ChromeOptions()
    prefs = {
        "savefile.default_directory": TMP_PDF_PATH,
        "download.prompt_for_download": False,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-logging")
    # 啟用 GPU 渲染診斷
    options.add_argument("--enable-logging")
    options.add_argument("--v=1")  # 設置日誌等級為 VERBOSE（詳細）
    options.add_argument("--log-level=3")  # 僅顯示錯誤日誌

    # 計算瀏覽器視窗位置和大小（邏輯像素）
    browser_x, browser_y, browser_width, browser_height = calculate_browser_window_position()
    print(f"瀏覽器視窗設定：位置({browser_x}, {browser_y})，大小({browser_width}x{browser_height})（邏輯像素）", flush=True)

    global driver
    driver = create_chrome_driver(options=options)

    # 🔥 啟動後再設定視窗大小/位置（不用 --window-size 啟動參數：避免高DPI下把物理像素當邏輯像素而過大、或DPI改變時崩潰）
    try:
        driver.set_window_size(browser_width, browser_height)
        driver.set_window_position(browser_x, browser_y)
    except Exception as _e:
        print(f"[視窗] 設定大小失敗：{_e}", flush=True)
    try:
        verify_and_fix_chrome_window(driver)
    except Exception:
        pass

def select_query_method():
    """
    讓使用者選擇查詢方式，而不是自動依序嘗試
    """
    print("\n===== 全國建築執照存根查詢 - 選擇查詢方式 =====", flush=True)
    print("請選擇查詢方式：", flush=True)
    print("[1] 建築地號查詢", flush=True)
    print("[2] 建築地址查詢", flush=True)
    print("[3] 手動查詢", flush=True)
    
    while True:
        choice = input("請輸入選項編號(1-3)：").strip()
        if choice in ["1", "2", "3"]:
            return choice
        print("無效選項，請重新輸入(1-3)", flush=True)

def select_data_for_lot_query():
    """
    從data.json中選擇要查詢的地段地號
    如果只有一筆資料，自動選擇；如果有多筆，讓使用者選擇
    """
    with open(get_data_json_path(), 'r', encoding='utf-8') as file:
        data_list = json.load(file)

    if not data_list:
        print("data.json中沒有資料。", flush=True)
        return None

    # 🔥 檢查是否為高雄市，只有高雄市才能執行建築執照查詢
    first_entry = data_list[0]
    city = first_entry.get('city', '')
    if city != '高雄市':
        print(f"\033[96m目前查詢是「{city}」，非【高雄市】！\033[0m", flush=True)
        print(f"\033[96m未來有機會時，再增加其他縣市，敬請期待~~\033[0m", flush=True)
        return None
    
    # 🔥 新增：如果只有一筆資料，直接自動選擇
    if len(data_list) == 1:
        selected_data = data_list[0]
        city = selected_data.get('city', 'N/A')
        area = selected_data.get('area', 'N/A')
        section = selected_data.get('section', 'N/A')
        lot_number = selected_data.get('lot_number', 'N/A')
        coordinates = selected_data.get('coordinates', '')
        coord_info = f" 座標: {coordinates}" if coordinates else ""

        print(f"\033[96mdata.json中只有1筆資料，自動選擇：{city}{area} {section} {lot_number}{coord_info}\033[0m", flush=True)
        return [selected_data]  # 🔥 返回列表格式，保持一致性
    
    # 原有的多筆資料選擇邏輯
    print("\n===== 建築地號查詢 - 選擇地段地號 =====", flush=True)
    print(f"在data.json中共找到 {len(data_list)} 筆資料：", flush=True)
    
    # 顯示所有資料
    for idx, data in enumerate(data_list, 1):
        city = data.get('city', 'N/A')
        area = data.get('area', 'N/A')
        section = data.get('section', 'N/A')
        lot_number = data.get('lot_number', 'N/A')
        
        # 格式化顯示每一筆資料
        data_info = f"[{idx}] {city}{area} {section} {lot_number}"
        if 'coordinates' in data:
            data_info += f" 座標: {data['coordinates']}"
        print(data_info, flush=True)
    
    # 🔥 用戶選擇（支援複選，以逗號分隔，支援 'a' 代表全選）
    print("\n請選擇要查詢的地段地號（輸入序號，可用逗號分隔多個，例如：1,3,5，或輸入 'a' 全選）：", flush=True)
    while True:
        try:
            user_input = input().strip().lower()

            # 🔥 處理全選
            if user_input == 'a' or user_input == 'all':
                choices = list(range(1, len(data_list) + 1))
                print(f"\n已選擇全部 {len(choices)} 個項目", flush=True)
            else:
                # 🔥 處理複選：分割逗號並解析多個序號
                choices_str = [s.strip() for s in user_input.split(',')]
                choices = []

                # 驗證所有輸入的序號
                for choice_str in choices_str:
                    choice = int(choice_str)
                    if 1 <= choice <= len(data_list):
                        choices.append(choice)
                    else:
                        print(f"序號 {choice} 超出範圍，請輸入1-{len(data_list)}之間的數字", flush=True)
                        choices = []  # 清空，重新輸入
                        break

            # 如果所有序號都有效
            if choices:
                # 顯示所有選擇的項目
                print(f"\n你選擇了 {len(choices)} 個項目：", flush=True)
                selected_data_list = []
                for choice in choices:
                    selected_data = data_list[choice-1]
                    city = selected_data.get('city', 'N/A')
                    area = selected_data.get('area', 'N/A')
                    section = selected_data.get('section', 'N/A')
                    lot_number = selected_data.get('lot_number', 'N/A')
                    coordinates = selected_data.get('coordinates', '')
                    coord_info = f" 座標: {coordinates}" if coordinates else ""
                    print(f"  [{choice}] {city}{area} {section} {lot_number}{coord_info}", flush=True)
                    selected_data_list.append(selected_data)

                # 🔥 返回選擇的資料列表（支援多筆查詢）
                return selected_data_list
        except ValueError:
            print("請輸入有效的數字（可用逗號分隔多個序號，或輸入 'a' 全選）", flush=True)

def input_address_for_query():
    """
    讓用戶輸入建築地址，包含完整的驗證和重試機制，並從data.json自動載入city和area
    """
    max_attempts = 3  # 最多允許嘗試次數
    attempt = 0
    
    # 嘗試從data.json讀取第一筆資料的city和area
    default_city = None
    default_area = None
    try:
        with open(get_data_json_path(), 'r', encoding='utf-8') as file:
            data_list = json.load(file)
            if data_list and len(data_list) > 0:
                first_data = data_list[0]
                default_city = first_data.get('city')
                default_area = first_data.get('area')

                # 🔥 檢查是否為高雄市，只有高雄市才能執行建築執照查詢
                if default_city != '高雄市':
                    print(f"\033[96m目前查詢是「{default_city}」，非【高雄市】！\033[0m", flush=True)
                    print(f"\033[96m未來有機會時，再增加其他縣市，敬請期待~~\033[0m", flush=True)
                    return None

                print(f"已從data.json載入縣市：{default_city}，區域：{default_area}", flush=True)
    except Exception as e:
        print(f"讀取data.json時發生錯誤: {e}", flush=True)
    
    while attempt < max_attempts:
        print("\n===== 建築地址查詢 =====", flush=True)
        
        # 顯示預設的縣市區信息，讓用戶只需輸入後續地址
        if default_city and default_area:
            print(f"使用預設縣市區：{default_city}{default_area}", flush=True)
            print("請輸入後續地址（例如：建國路三段260號）：", flush=True)
            street_address = input().strip()
            
            # 生成完整地址
            address = f"{default_city}{default_area}{street_address}"
            city = default_city
            area = default_area
        else:
            print("請輸入完整建築地址（例如：高雄市鳳山區建國路三段260號）：", flush=True)
            address = input().strip()
            
            # 驗證地址是否符合基本格式
            if not address:
                print("地址不能為空，請重新輸入", flush=True)
                attempt += 1
                continue
                
            # 驗證是否包含縣市和區域
            if not ("市" in address and "區" in address):
                print("地址格式不完整，請包含縣市和區域（例如：高雄市鳳山區...）", flush=True)
                retry = input("是否要重新輸入？(Y/N)：").strip().upper()
                if retry != 'N':
                    attempt += 1
                    continue
            
            # 解析地址以獲取city和area
            city = None
            area = None
            
            if "市" in address and "區" in address:
                city_parts = address.split("市", 1)
                city = city_parts[0] + "市"
                area_parts = city_parts[1].split("區", 1)
                area = area_parts[0] + "區"
                
                print(f"已解析出縣市：{city}，區域：{area}", flush=True)
            else:
                print("無法解析出縣市區資訊，請重新輸入", flush=True)
                attempt += 1
                continue
        
        # 進一步驗證地址元素（路/街、號等）
        if not any(keyword in address for keyword in ["路", "街"]):
            print("地址似乎缺少路名或街名", flush=True)
            retry = input("是否要重新輸入？(Y/N)：").strip().upper()
            if retry != 'N':
                attempt += 1
                continue
                
        if "號" not in address:
            print("地址似乎缺少門牌號碼", flush=True)
            print("是否要繼續使用此地址？(Y/N)：[Y]", flush=True)
            retry = input().strip().upper()
            if retry != '' and retry != 'Y':
                attempt += 1
                continue
        
        # 確認地址
        print(f"您輸入的完整地址是：{address}", flush=True)
        print("確認使用此地址？(Y/N)：[Y]", flush=True)
        confirm = input().strip().upper()
        if confirm == '' or confirm == 'Y':
            return {
                "city": city,
                "area": area,
                "address": address
            }
        else:
            print("已取消，請重新輸入", flush=True)
            attempt += 1
    
    # 超過最大嘗試次數
    print(f"已超過最大嘗試次數({max_attempts})，程序將退出", flush=True)
    return None

def display_and_select_data():
    """
    顯示data.json中的所有資料，並讓使用者選擇要執行的項目
    返回：選中的資料列表
    """
    with open(get_data_json_path(), 'r', encoding='utf-8') as file:
        data_list = json.load(file)
    
    if not data_list:
        print("data.json中沒有資料。", flush=True)
        return []
    
    print("\n===== 全國建築執照存根查詢 - 資料選擇 =====", flush=True)
    print(f"在data.json中共找到 {len(data_list)} 筆資料：", flush=True)
    
    # 顯示所有資料
    for idx, data in enumerate(data_list, 1):
        city = data.get('city', 'N/A')
        area = data.get('area', 'N/A')
        section = data.get('section', 'N/A')
        lot_number = data.get('lot_number', 'N/A')
        address = data.get('address', 'N/A')
        
        # 格式化顯示每一筆資料
        data_info = f"[{idx}] {city}{area} {section} {lot_number}"
        if 'coordinates' in data:
            data_info += f" 座標: {data['coordinates']}"
        print(data_info, flush=True)
        if address:
            print(f"    地址: {address}", flush=True)
    
    # 用戶選擇
    print("\n請選擇要執行的項目編號（多個項目請用逗號分隔，輸入'all'執行全部）：", flush=True)
    user_input = input().strip()
    
    # 若用戶選擇全部執行
    if user_input.lower() == 'all':
        print("將執行所有項目", flush=True)
        return data_list
    
    # 否則，解析用戶選擇的項目
    selected_indices = []
    try:
        # 解析用戶輸入，支援逗號分隔和範圍選擇(如1-3)
        for part in user_input.split(','):
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                selected_indices.extend(range(start, end + 1))
            elif part:
                selected_indices.append(int(part))
    except ValueError:
        print("輸入格式錯誤，請輸入數字或數字範圍。將執行所有項目。", flush=True)
        return data_list
    
    # 驗證選擇的索引是否有效
    selected_indices = [idx for idx in selected_indices if 1 <= idx <= len(data_list)]
    if not selected_indices:
        print("未選擇有效項目，將退出執行。", flush=True)
        return []
    
    # 返回選中的資料
    selected_data = [data_list[idx-1] for idx in selected_indices]
    print(f"將執行 {len(selected_data)} 個選定項目", flush=True)
    return selected_data



def main_program():
    """
    主程式循環，允許用戶連續查詢多個地址或地號
    """
    global SHOULD_EXIT_PROGRAM
    SHOULD_EXIT_PROGRAM = False

    # 🔥 在程式開始時檢查是否為高雄市
    try:
        with open(get_data_json_path(), 'r', encoding='utf-8') as file:
            data_list = json.load(file)
            if data_list and len(data_list) > 0:
                first_entry = data_list[0]
                city = first_entry.get('city', '')
                if city != '高雄市':
                    print(f"\033[96m目前查詢是「{city}」，非【高雄市】！\033[0m", flush=True)
                    print(f"\033[96m未來有機會時，再增加其他縣市，敬請期待~~\033[0m", flush=True)
                    return  # 直接退出程式
    except Exception as e:
        print(f"讀取data.json時發生錯誤: {e}", flush=True)
        return

    try:
        continue_querying = True

        while continue_querying and not SHOULD_EXIT_PROGRAM:
            try:
                # 🔥 使用已完善的 test() 函數（支援多筆資料查詢）
                test()

                # 如果設置了退出訊號，就直接退出循環
                if SHOULD_EXIT_PROGRAM:
                    print("用戶請求完全退出程式。", flush=True)
                    break

                """詢問用戶是否繼續查詢，返回布爾值"""
                print("請選擇後續操作：", flush=True)
                print("[1] 繼續查詢建築地號", flush=True)
                print("[2] 繼續查詢建築地址", flush=True)
                print("[3] 繼續手動查詢", flush=True)
                print("[0] 結束程式", flush=True)
                print("請輸入選項編號(0-3)：", flush=True)

                try:
                    choice = safe_input().strip()
                except (EOFError, OSError):
                    # 標準輸入已關閉，直接退出
                    continue_querying = False
                    break

                if choice == "0":
                    continue_querying = False
                    print("程式結束，感謝使用！", flush=True)
                elif choice in ["1", "2", "3"]:
                    # 繼續查詢，下一次循環將啟動新的查詢
                    global PRESELECTED_QUERY_TYPE
                    PRESELECTED_QUERY_TYPE = choice
                    print(f"將繼續進行{'建築地號' if choice == '1' else '建築地址' if choice == '2' else '手動'}查詢...", flush=True)
                else:
                    print("無效選項，程式將結束。", flush=True)
                    continue_querying = False

            except (EOFError, OSError):
                # 標準輸入/輸出已關閉，靜默退出
                continue_querying = False
            except Exception as e:
                try:
                    print(f"查詢過程中發生錯誤：{e}", flush=True)
                    retry = safe_input("是否重新開始查詢？(Y/N)：").strip().upper()
                    if retry != 'Y':
                        continue_querying = False
                except (EOFError, OSError):
                    continue_querying = False

    except (EOFError, OSError):
        # 標準輸入/輸出已關閉，靜默退出
        pass
    except Exception as e:
        try:
            print(f"程式執行過程中發生錯誤：{e}", flush=True)
        except (EOFError, OSError):
            pass
    finally:
        try:
            print("程式已完成執行，感謝使用！", flush=True)
        except (EOFError, OSError):
            pass

def run_query_session():
    """執行一次完整的查詢流程，包括初始化瀏覽器和關閉"""
    global driver
    try:
        # 1. 詢問查詢類型
        query_type = select_query_type_with_preselected()

        # 如果用戶選擇退出（query_type 為 None），直接返回
        if query_type is None:
            return

        # 2. 根據查詢類型獲取必要的資料
        selected_data = None
        address_data = None

        if query_type == "1":  # 建築地號查詢
            selected_data = select_data_for_lot_query()
            if not selected_data:
                print("未選擇有效的地段地號資料，本次查詢結束。", flush=True)
                return
        
        elif query_type == "2":  # 建築地址查詢
            address_data = input_address_for_query()
            if not address_data:
                print("未獲取有效的地址資料，本次查詢結束。", flush=True)
                return
        
        # 3. 初始化瀏覽器並打開網頁
        init_driver()
        print("正在開啟網頁...", flush=True)
        driver.get("https://cloudbm.nlma.gov.tw/CPTL/cpt0407m.do?")
        print("網頁已開啟，開始查詢資料", flush=True)

        # 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
        verify_and_fix_chrome_window(driver)

        # 🔥 自動判斷是否需要縮放頁面（解決高 DPI 或小螢幕下的排版問題）
        # 根據 DPI 決定縮放次數：150% → 2次(80%)，175% → 4次(67%)
        zoom_count = 0
        try:
            config_path = os.path.join(BASE_DIR, 'window_config.json')
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                current_dpi = config.get('dpi_scale', 1.0)
                screen_width = config.get('screen_width', 1920)
                logical_width = screen_width / current_dpi

                # 🔥 根據 DPI 和螢幕解析度決定縮放次數
                if current_dpi >= 1.75:
                    zoom_count = 4  # 175% 或更高：縮放到 67%（按 4 次）
                    target_zoom = "67%"
                elif current_dpi >= 1.5:
                    zoom_count = 2  # 150%：縮放到 80%
                    target_zoom = "80%"
                elif current_dpi >= 1.25 and logical_width < 1200:
                    zoom_count = 2  # 小螢幕 + 125%：縮放到 80%
                    target_zoom = "80%"
                else:
                    zoom_count = 0  # 100%、125%（大螢幕）：不縮放
                    target_zoom = "100%"

        except Exception:
            pass  # 靜默處理偵測配置錯誤

        # 🔥 使用者在「Chrome縮放設定」面板調過的值優先（zoom_config.json）；沒調過才用上面的 DPI 預設
        try:
            _zcp = os.path.join(BASE_DIR, 'zoom_config.json')
            if os.path.exists(_zcp):
                _zov = json.load(open(_zcp, encoding='utf-8')).get('overrides', {}).get('nlma')
                if _zov is not None:
                    zoom_count = int(_zov)
        except Exception:
            pass

        apply_page_zoom(driver, zoom_count)  # 共用：拉前景→Ctrl+0→縮放→驗證

        # 4. 設置查詢結果和狀態變量
        all_results = {}  # 用於儲存所有有效結果
        status = None
        results = None
        
        try:
            # 5. 根據查詢類型執行對應的查詢流程
            base_directory = setup_directory(query_type, selected_data, address_data)

            if query_type == "1":  # 建築地號查詢
                # 🔥 selected_data 現在是列表，取第一筆資料
                data_item = selected_data[0] if isinstance(selected_data, list) else selected_data
                city = data_item.get('city')
                area = data_item.get('area')
                section = data_item.get('section')
                lot_number = data_item.get('lot_number')

                print(f"\033[96m開始處理資料：city={city}, area={area}, section={section}, lot_number={lot_number}\033[0m", flush=True)
                select_city(city)

                status, results = perform_search("4", city, area, section, lot_number)
            
            elif query_type == "2":  # 建築地址查詢
                city = address_data.get('city', '')
                area = address_data.get('area', '')
                address = address_data.get('address', '')
                
                print(f"開始處理地址：{address}", flush=True)
                select_city(city)
                
                status, results = perform_search("2", city, area, None, None, address)
            
            elif query_type == "3":  # 手動查詢
                print("自動選擇高雄市並進入手動查詢模式...", flush=True)
                select_city("高雄市")
                
                status, results = perform_search("3", None, None, None, None)
            
            # 6. 處理查詢結果
            process_results(status, results, base_directory, all_results)
            
        except Exception as e:
            print(f"查詢過程中發生錯誤：{e}", flush=True)
        finally:
            # 7. 關閉瀏覽器
            print("查詢完成，關閉瀏覽器", flush=True)
            if driver:
                try:
                    driver.quit()
                except Exception as e:
                    # 瀏覽器可能已被手動關閉，靜默處理
                    pass
                finally:
                    driver = None
    
    except Exception as e:
        print(f"執行查詢流程時發生錯誤：{e}", flush=True)


def select_query_type_with_preselected():
    """獲取查詢類型，支持預選的類型"""
    global PRESELECTED_QUERY_TYPE
    
    # 如果有預選類型則使用預選
    if PRESELECTED_QUERY_TYPE:
        query_type = PRESELECTED_QUERY_TYPE
        print(f"\n使用預選查詢類型：{'建築地號查詢' if query_type == '1' else '建築地址查詢' if query_type == '2' else '手動查詢'}", flush=True)
        # 使用後清除預選值，避免影響下次正常選擇
        PRESELECTED_QUERY_TYPE = None
        return query_type
    
    # 否則正常詢問用戶
    print("\n===== 全國建築執照存根查詢 - 選擇查詢方式 =====", flush=True)
    print("請選擇查詢方式：", flush=True)
    print("[1] 建築地號查詢", flush=True)
    print("[2] 建築地址查詢", flush=True)
    print("[3] 手動查詢", flush=True)
    print("[0] 退出本程式", flush=True)
    print("請輸入選項編號(0-3)：", flush=True)

    query_type = ""
    while query_type not in ["0", "1", "2", "3"]:
        query_type = input().strip()
        if query_type not in ["0", "1", "2", "3"]:
            print("無效選項，請重新輸入(0-3)", flush=True)

    # 如果用戶選擇退出
    if query_type == "0":
        global SHOULD_EXIT_PROGRAM
        SHOULD_EXIT_PROGRAM = True
        print("程式結束，感謝使用！", flush=True)
        return None

    return query_type

def setup_directory(query_type, selected_data, address_data):
    """根據查詢類型設置儲存目錄"""
    # 首先嘗試從 data.json 中讀取第一組資料作為預設命名
    default_directory = None
    try:
        with open(get_data_json_path(), 'r', encoding='utf-8') as file:
            data_list = json.load(file)
            if data_list and len(data_list) > 0:
                first_data = data_list[0]
                # 🔥 使用 get_work_folder 確保在主程式目錄下建立
                default_directory = get_work_folder(
                    f"{first_data['area']}{first_data['section']}-{first_data['lot_number']}"
                )
    except Exception as e:
        print(f"讀取 data.json 時發生錯誤: {e}", flush=True)

    if query_type == "1":  # 建築地號查詢
        # 🔥 selected_data 現在是列表，取第一筆資料
        data_item = selected_data[0] if isinstance(selected_data, list) else selected_data
        # 🔥 使用 get_work_folder 確保在主程式目錄下建立
        base_directory = get_work_folder(
            f"{data_item['area']}{data_item['section']}-{data_item['lot_number']}"
        )
    elif query_type == "2":  # 建築地址查詢
        # 使用 data.json 中的第一組資料作為目錄
        if default_directory:
            base_directory = default_directory
        else:
            # 如果無法讀取 data.json，則使用原有的命名方式
            city = address_data.get('city', '')
            area = address_data.get('area', '')

            if not city or not area:
                base_directory = get_work_folder("地址查詢結果")
            else:
                base_directory = get_work_folder(f"{city}{area}-地址查詢")
            print(f"警告: 無法使用 data.json 中的資料作為目錄，使用備用命名: {base_directory}", flush=True)
    else:  # 手動查詢
        # 使用 data.json 中的第一組資料作為目錄
        if default_directory:
            base_directory = default_directory
        else:
            base_directory = get_work_folder("手動查詢結果")
    
    os.makedirs(base_directory, exist_ok=True)
    return base_directory

# 修改process_results函数中创建选项的部分
def process_results(status, results, base_directory, all_results):
    """處理查詢結果並允許用戶選擇下載項目"""
    if status == "manual_exit":
        print("用戶手動退出查詢。", flush=True)
    
    elif status == "success" and results:
        print(f"查詢成功，共找到 {len(results)} 筆資料。", flush=True)

        # 修改这一行，使用数字索引而非字母
        all_results.update({generate_option(i + 1): result for i, result in enumerate(results)})

        # 🔥 只有一筆資料時，自動處理，不詢問使用者
        if len(all_results) == 1:
            print("只有一筆資料，自動處理中...", flush=True)
            for option, result in all_results.items():
                try:
                    click_and_save_pdf(result, base_directory)
                except Exception as e:
                    print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)
            print("處理完成。", flush=True)
            return

        # 顯示結果並讓用戶選擇
        while results:  # 只有在有結果時才進入選擇循環
            print_results_with_options(all_results)
            print("請輸入選擇的選項（例如：1, 3 或 2-3 範圍，可混用 1,3-5；a/all 全選）：", flush=True)
            _raw = input().strip()

            # 🔥 處理"全選"（支援 all, a, A）
            if _raw.lower() in ("all", "a"):
                print("開始處理所有選項...", flush=True)
                for option, result in all_results.items():
                    try:
                        click_and_save_pdf(result, base_directory)
                    except Exception as e:
                        print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)

                print("所有選項處理完成。", flush=True)
                break

            # 處理單個/多個/範圍選項（支援逗號分隔與 2-3 範圍，可混用）
            else:
                _expanded = []
                for _part in _raw.split(","):
                    _part = _part.strip()
                    if not _part:
                        continue
                    if "-" in _part:
                        try:
                            _a, _b = [int(x.strip()) for x in _part.split("-", 1)]
                            if _a > _b:
                                _a, _b = _b, _a
                            _expanded.extend(str(n) for n in range(_a, _b + 1))
                        except ValueError:
                            _expanded.append(_part)  # 解析失敗保留原值，下面會判無效
                    else:
                        _expanded.append(_part)
                # 去重保序
                _seen = set()
                selected_options = []
                for _o in _expanded:
                    if _o not in _seen:
                        _seen.add(_o)
                        selected_options.append(_o)

                all_options_valid = True
                for option in selected_options:
                    option = option.strip()
                    if option in all_results:
                        try:
                            click_and_save_pdf(all_results[option], base_directory)
                        except Exception as e:
                            print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)
                    else:
                        print(f"無效的選項：{option}", flush=True)
                        all_options_valid = False
                        break

                if all_options_valid:
                    print("選定的選項處理完成。", flush=True)
                    break
    
    elif status == "no_data":
        print("查詢無資料。", flush=True)
    
    else:
        print("查詢失敗或中斷。", flush=True)

def test():
    """
    主要執行函數，先詢問查詢類型，然後依據類型執行不同流程
    支持連續查詢模式，每次查詢完成後詢問用戶是否繼續
    """
    global PRESELECTED_QUERY_TYPE, driver
    try:
        continue_querying = True
        
        while continue_querying:
            # 1. 詢問用戶要使用哪種查詢類型，如果有預選類型則使用預選
            if 'PRESELECTED_QUERY_TYPE' in globals() and PRESELECTED_QUERY_TYPE:
                query_type = PRESELECTED_QUERY_TYPE
                query_name = '建築地號查詢' if query_type == '1' else '建築地址查詢' if query_type == '2' else '手動查詢'
                print(f"\n使用預選查詢類型：{query_name}", flush=True)
                # 使用後清除預選值，避免影響下次正常選擇
                PRESELECTED_QUERY_TYPE = None
            else:
                print("\n===== 全國建築執照存根查詢 - 選擇查詢方式 =====", flush=True)
                print("請選擇查詢方式：", flush=True)
                print("[1] 建築地號查詢", flush=True)
                print("[2] 建築地址查詢", flush=True)
                print("[3] 手動查詢", flush=True)
                
                # 確保用戶輸入有效的查詢類型
                query_type = ""
                while query_type not in ["1", "2", "3"]:
                    print("請輸入選項編號(1-3)：", flush=True)
                    query_type = safe_input().strip()
                    if query_type not in ["1", "2", "3"]:
                        print("無效選項，請重新輸入(1-3)", flush=True)
            
            # 2. 根據查詢類型初始化必要的資料
            selected_data = None
            address_data = None
            
            # 3. 針對不同查詢類型，獲取所需資料
            if query_type == "1":  # 建築地號查詢
                selected_data = select_data_for_lot_query()
                if not selected_data:
                    print("未選擇有效的地段地號資料。", flush=True)

                    # 詢問是否繼續其他查詢
                    print("\n===== 查詢已終止 =====", flush=True)
                    if not ask_continue_query():
                        break
                    continue
            
            elif query_type == "2":  # 建築地址查詢
                address_data = input_address_for_query()
                if not address_data:
                    print("未獲取有效的地址資料。", flush=True)
                    
                    # 詢問是否繼續其他查詢
                    print("\n===== 查詢已終止 =====", flush=True)
                    if not ask_continue_query():
                        break
                    continue
            
            # 4. 初始化瀏覽器並打開網頁
            init_driver()
            print("正在開啟網頁...", flush=True)
            driver.get("https://cloudbm.nlma.gov.tw/CPTL/cpt0407m.do?")
            print("網頁已開啟，開始查詢資料", flush=True)

            try:
                # 6. 根據查詢類型執行對應的查詢流程
                if query_type == "1":  # 建築地號查詢
                    # 🔥 使用 data.json 中的第一組資料作為儲存目錄（統一三種查詢方式）
                    try:
                        with open(get_data_json_path(), 'r', encoding='utf-8') as file:
                            data_list = json.load(file)
                            if data_list and len(data_list) > 0:
                                first_data = data_list[0]
                                base_directory = get_work_folder(
                                    f"{first_data['area']}{first_data['section']}-{first_data['lot_number']}"
                                )
                            else:
                                # 如果 data.json 為空，使用 selected_data 的第一筆
                                first_item = selected_data[0] if isinstance(selected_data, list) else selected_data
                                base_directory = get_work_folder(
                                    f"{first_item['area']}{first_item['section']}-{first_item['lot_number']}"
                                )
                    except Exception as e:
                        first_item = selected_data[0] if isinstance(selected_data, list) else selected_data
                        base_directory = get_work_folder(
                            f"{first_item['area']}{first_item['section']}-{first_item['lot_number']}"
                        )
                    os.makedirs(base_directory, exist_ok=True)

                    # 🔥 循環處理所有選擇的資料
                    for data_index, data_item in enumerate(selected_data, 1):
                        print(f"\n\033[96m{'='*60}\033[0m", flush=True)
                        print(f"\033[96m正在處理第 {data_index}/{len(selected_data)} 筆資料\033[0m", flush=True)
                        print(f"\033[96m{'='*60}\033[0m\n", flush=True)

                        # 5. 設置查詢結果和狀態變量
                        all_results = {}  # 用於儲存所有有效結果
                        status = None
                        results = None

                        # 選擇縣市並執行地號查詢
                        city = data_item.get('city')
                        area = data_item.get('area')
                        section = data_item.get('section')
                        lot_number = data_item.get('lot_number')

                        print(f"\033[96m開始處理資料：city={city}, area={area}, section={section}, lot_number={lot_number}\033[0m", flush=True)
                        select_city(city)

                        # 只執行建築地號查詢（對應原代碼中的類型"4"）
                        status, results = perform_search("4", city, area, section, lot_number)

                        # 處理查詢結果
                        if status == "manual_exit":
                            print("用戶手動退出查詢。", flush=True)
                            break  # 退出資料循環

                        elif status == "success" and results:
                            print(f"查詢成功，共找到 {len(results)} 筆資料。", flush=True)

                            # 將結果添加到全局結果集
                            all_results.update({str(i + 1): result for i, result in enumerate(results)})

                            # 🔥 只有一筆資料時，自動處理，不詢問使用者
                            if len(all_results) == 1:
                                print("只有一筆資料，自動處理中...", flush=True)
                                for option, result in all_results.items():
                                    try:
                                        click_and_save_pdf(result, base_directory)
                                    except Exception as e:
                                        print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)
                                print("處理完成。", flush=True)
                            else:
                                # 顯示結果並讓用戶選擇
                                while results:  # 只有在有結果時才進入選擇循環
                                    print_results_with_options(all_results)
                                    print("請輸入選擇的選項（例如：1, 3 或 2-3 範圍，可混用 1,3-5；a/all 全選）：", flush=True)
                                    selected_options = _expand_sel(safe_input())

                                    # 🔥 處理"全選"（支援 all, a, A）
                                    first_option = selected_options[0].strip().lower()
                                    if first_option in ("all", "a"):
                                        print("開始處理所有選項...", flush=True)
                                        for option, result in all_results.items():
                                            try:
                                                click_and_save_pdf(result, base_directory)
                                            except Exception as e:
                                                print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)

                                        print("所有選項處理完成。", flush=True)
                                        break

                                    # 處理單個或多個選項
                                    else:
                                        all_options_valid = True
                                        for option in selected_options:
                                            option = option.strip()
                                            if option in all_results:
                                                try:
                                                    click_and_save_pdf(all_results[option], base_directory)
                                                except Exception as e:
                                                    print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)
                                            else:
                                                print(f"無效的選項：{option}", flush=True)
                                                all_options_valid = False
                                                break

                                        if all_options_valid:
                                            print("選定的選項處理完成。", flush=True)
                                            break

                        elif status == "no_data":
                            print("查詢無資料。", flush=True)

                        else:
                            print("查詢失敗或中斷。", flush=True)

                        # 如果還有下一筆資料，自動繼續查詢（不需詢問）
                        if data_index < len(selected_data):
                            print(f"\n還有 {len(selected_data) - data_index} 筆資料待查詢，自動繼續下一筆...", flush=True)
                            time.sleep(1)  # 稍微延遲，讓使用者看到訊息

                            # 🔥 重新載入查詢頁面，確保網頁回到初始狀態
                            print("重新載入查詢頁面...", flush=True)
                            driver.get("https://cloudbm.nlma.gov.tw/CPTL/cpt0407m.do?")
                            time.sleep(2)  # 等待頁面載入
                            # 直接 continue 到下一輪迴圈

                elif query_type == "2":  # 建築地址查詢
                    # 設置查詢結果和狀態變量
                    all_results = {}
                    status = None
                    results = None

                    city = address_data.get('city', '')
                    area = address_data.get('area', '')
                    address = address_data.get('address', '')

                    # 🔥 使用 data.json 中的第一組資料作為儲存目錄（統一三種查詢方式）
                    try:
                        with open(get_data_json_path(), 'r', encoding='utf-8') as file:
                            data_list = json.load(file)
                            if data_list and len(data_list) > 0:
                                first_data = data_list[0]
                                base_directory = get_work_folder(
                                    f"{first_data['area']}{first_data['section']}-{first_data['lot_number']}"
                                )
                            else:
                                base_directory = get_work_folder("地址查詢結果")
                    except Exception as e:
                        base_directory = get_work_folder("地址查詢結果")

                    os.makedirs(base_directory, exist_ok=True)

                    # 選擇縣市
                    print(f"開始處理地址：{address}", flush=True)
                    select_city(city)

                    # 使用原始的perform_search函數，但確保使用正確的參數
                    # 注意這裡傳入的查詢類型是"2"表示建築地址查詢
                    status, results = perform_search("2", city, area, None, None, address)

                    # 處理查詢結果
                    if status == "manual_exit":
                        print("用戶手動退出查詢。", flush=True)

                    elif status == "success" and results:
                        print(f"查詢成功，共找到 {len(results)} 筆資料。", flush=True)

                        # 將結果添加到全局結果集
                        all_results.update({str(i + 1): result for i, result in enumerate(results)})

                        # 🔥 只有一筆資料時，自動處理，不詢問使用者
                        if len(all_results) == 1:
                            print("只有一筆資料，自動處理中...", flush=True)
                            for option, result in all_results.items():
                                try:
                                    click_and_save_pdf(result, base_directory)
                                except Exception as e:
                                    print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)
                            print("處理完成。", flush=True)
                        else:
                            # 顯示結果並讓用戶選擇
                            while results:
                                print_results_with_options(all_results)
                                print("請輸入選擇的選項（例如：1, 3 或 2-3 範圍，可混用 1,3-5；a/all 全選）：", flush=True)
                                selected_options = _expand_sel(safe_input())

                                # 🔥 處理"全選"（支援 all, a, A）
                                first_option = selected_options[0].strip().lower()
                                if first_option in ("all", "a"):
                                    print("開始處理所有選項...", flush=True)
                                    for option, result in all_results.items():
                                        try:
                                            click_and_save_pdf(result, base_directory)
                                        except Exception as e:
                                            print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)

                                    print("所有選項處理完成。", flush=True)
                                    break

                                # 處理單個或多個選項
                                else:
                                    all_options_valid = True
                                    for option in selected_options:
                                        option = option.strip()
                                        if option in all_results:
                                            try:
                                                click_and_save_pdf(all_results[option], base_directory)
                                            except Exception as e:
                                                print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)
                                        else:
                                            print(f"無效的選項：{option}", flush=True)
                                            all_options_valid = False
                                            break

                                    if all_options_valid:
                                        print("選定的選項處理完成。", flush=True)
                                        break

                    elif status == "no_data":
                        print("查詢無資料。", flush=True)

                    else:
                        print("查詢失敗或中斷。", flush=True)

                elif query_type == "3":  # 手動查詢
                    # 設置查詢結果和狀態變量
                    all_results = {}
                    status = None
                    results = None

                    # 🔥 使用 data.json 中的第一組資料作為儲存目錄（與選項1、2統一）
                    try:
                        with open(get_data_json_path(), 'r', encoding='utf-8') as file:
                            data_list = json.load(file)
                            if data_list and len(data_list) > 0:
                                first_data = data_list[0]
                                base_directory = get_work_folder(
                                    f"{first_data['area']}{first_data['section']}-{first_data['lot_number']}"
                                )
                            else:
                                base_directory = get_work_folder("手動查詢結果")
                    except Exception as e:
                        base_directory = get_work_folder("手動查詢結果")

                    os.makedirs(base_directory, exist_ok=True)

                    # 自動選擇高雄市（不再詢問用戶）
                    print("自動選擇高雄市並進入手動查詢模式...", flush=True)
                    select_city("高雄市")

                    # 執行手動查詢（對應原代碼中的類型"3"）
                    status, results = perform_search("3", None, None, None, None)

                    # 處理查詢結果
                    if status == "manual_exit":
                        print("用戶手動退出查詢。", flush=True)

                    elif status == "success" and results:
                        print(f"查詢成功，共找到 {len(results)} 筆資料。", flush=True)

                        # 將結果添加到全局結果集
                        all_results.update({str(i + 1): result for i, result in enumerate(results)})

                        # 🔥 只有一筆資料時，自動處理，不詢問使用者
                        if len(all_results) == 1:
                            print("只有一筆資料，自動處理中...", flush=True)
                            for option, result in all_results.items():
                                try:
                                    click_and_save_pdf(result, base_directory)
                                except Exception as e:
                                    print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)
                            print("處理完成。", flush=True)
                        else:
                            # 顯示結果並讓用戶選擇
                            while results:  # 只有在有結果時才進入選擇循環
                                print_results_with_options(all_results)
                                print("請輸入選擇的選項（例如：1, 3 或 2-3 範圍，可混用 1,3-5；a/all 全選）：", flush=True)
                                selected_options = _expand_sel(safe_input())

                                # 🔥 處理"全選"（支援 all, a, A）
                                first_option = selected_options[0].strip().lower()
                                if first_option in ("all", "a"):
                                    print("開始處理所有選項...", flush=True)
                                    for option, result in all_results.items():
                                        try:
                                            click_and_save_pdf(result, base_directory)
                                        except Exception as e:
                                            print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)

                                    print("所有選項處理完成。", flush=True)
                                    break

                                # 處理單個或多個選項
                                else:
                                    all_options_valid = True
                                    for option in selected_options:
                                        option = option.strip()
                                        if option in all_results:
                                            try:
                                                click_and_save_pdf(all_results[option], base_directory)
                                            except Exception as e:
                                                print(f"處理選項 {option} 時發生錯誤：{e}", flush=True)
                                        else:
                                            print(f"無效的選項：{option}", flush=True)
                                            all_options_valid = False
                                            break

                                    if all_options_valid:
                                        print("選定的選項處理完成。", flush=True)
                                        break

                    elif status == "no_data":
                        print("查詢無資料。", flush=True)

                    else:
                        print("查詢失敗或中斷。", flush=True)

            except Exception as e:
                print(f"查詢過程中發生錯誤：{e}", flush=True)
            finally:
                # 關閉瀏覽器
                if driver:
                    print("本次查詢完成，關閉瀏覽器", flush=True)
                    driver.quit()
                    driver = None
            
            # 8. 詢問是否繼續其他查詢
            print("\n===== 查詢已完成 =====", flush=True)
            if not ask_continue_query():
                break
    
    except Exception as e:
        print(f"執行過程中發生錯誤: {e}", flush=True)
    finally:
        # 確保瀏覽器正確關閉
        if driver:
            print("關閉瀏覽器", flush=True)
            driver.quit()
        notify_main_program()

def ask_continue_query():
    """詢問用戶是否繼續查詢，返回布爾值"""
    print("請選擇後續操作：", flush=True)
    print("[1] 繼續查詢建築地號", flush=True)
    print("[2] 繼續查詢建築地址", flush=True)
    print("[3] 繼續手動查詢", flush=True)
    print("[0] 結束程式", flush=True)
    print("請輸入選項編號(0-3)：", flush=True)
    
    choice = input().strip()
    
    if choice == "0":
        print("程式結束，感謝使用！", flush=True)
        return False
    elif choice in ["1", "2", "3"]:
        # 設置全局預選的查詢類型
        global PRESELECTED_QUERY_TYPE
        PRESELECTED_QUERY_TYPE = choice
        print(f"將繼續進行{'建築地號' if choice == '1' else '建築地址' if choice == '2' else '手動'}查詢...", flush=True)
        return True
    else:
        print("無效選項，程式將結束。", flush=True)
        return False

# 新增全局變數來存儲預選的查詢類型
PRESELECTED_QUERY_TYPE = None


def select_city(city_name):
    try:
        city_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{city_name}')]"))
        )
        city_link.click()
        print(f"\033[96m已成功選擇縣市: {city_name}\033[0m", flush=True)

        all_windows = driver.window_handles
        if all_windows: #加入了這個，以免選單在切換時找不到任何選單造成錯誤。
            if len(all_windows) > 1:
                driver.switch_to.window(all_windows[-1])
                print("已切換到新視窗", flush=True)

                # 🔥 新視窗也需要縮放，讀取 DPI 設定並執行縮放
                try:
                    config_path = os.path.join(BASE_DIR, 'window_config.json')
                    if os.path.exists(config_path):
                        with open(config_path, 'r', encoding='utf-8') as f:
                            config = json.load(f)
                        current_dpi = config.get('dpi_scale', 1.0)
                        screen_width = config.get('screen_width', 1920)
                        logical_width = screen_width / current_dpi

                        # 🔥 根據 DPI 和螢幕解析度決定縮放次數
                        if current_dpi >= 1.75:
                            zoom_count = 4  # 175%：縮放到 67%（按 4 次）
                        elif current_dpi >= 1.5:
                            zoom_count = 2  # 150%：縮放到 80%（按 2 次）
                        elif current_dpi >= 1.25 and logical_width < 1200:
                            zoom_count = 2  # 小螢幕 + 125%：縮放到 80%
                        else:
                            zoom_count = 0  # 100%、125%（大螢幕）：不縮放

                        # 🔥 使用者覆寫優先（zoom_config.json）
                        try:
                            _zcp2 = os.path.join(BASE_DIR, 'zoom_config.json')
                            if os.path.exists(_zcp2):
                                _zov2 = json.load(open(_zcp2, encoding='utf-8')).get('overrides', {}).get('nlma')
                                if _zov2 is not None:
                                    zoom_count = int(_zov2)
                        except Exception:
                            pass

                        if zoom_count > 0:
                            time.sleep(0.5)  # 等待新視窗完全載入
                        apply_page_zoom(driver, zoom_count)  # 共用：拉前景→Ctrl+0→縮放→驗證
                except Exception:
                    pass  # 靜默處理縮放錯誤
        # print(f"當前視窗網址: {driver.current_url}", flush=True)
            else:
             print("沒有發現新視窗，仍然在原視窗中", flush=True)
    except TimeoutException:
        print(f"無法找到指定的縣市連結: {city_name}", flush=True)
        
def select_query_type(query_type_value):
    """選擇查詢類型 radio。
    🔥 新版高雄站使用 id="QType1" ~ "QType6"。
       對照：1=執照號碼  2=起造人姓名  3=建築地址  4=建築地號  5=發照日期  6=掛號號碼
    🔥 重點：QType radio 被 CSS 隱藏（美化過的 radio 設計），所以用 presence_of_element_located
       而非 element_to_be_clickable，並用 JS click 繞過可見性檢查。
    🔥 點完還要等對應的 searchbbox{NN} div 從 display:none 變 display:block。
    """
    if query_type_value not in ("1", "2", "3", "4", "5", "6"):
        print(f"無效的查詢類型：{query_type_value}", flush=True)
        return

    # 等頁面 DOM 載入完成
    try:
        WebDriverWait(driver, 15).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except TimeoutException:
        print("[警告] 頁面載入超時，但繼續嘗試...", flush=True)

    elem_id = f"QType{query_type_value}"
    box_class = f"searchbbox0{query_type_value}"  # 對應的表單 div class

    try:
        # 用 presence_of_element_located（只看 DOM 存在，不管視覺可見）
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, elem_id))
        )
        radio = driver.find_element(By.ID, elem_id)
        # JS click 繞過 Selenium 的「element 必須可見才能點」限制
        driver.execute_script("arguments[0].click();", radio)
        print(f"已選擇查詢類型：{query_type_value}（#{elem_id}）", flush=True)

        # 🔥 等對應的 searchbbox 區塊變 visible（display:block）才算真的切換完成
        try:
            WebDriverWait(driver, 5).until(
                lambda d: 'block' in (
                    d.find_element(By.CSS_SELECTOR, f"div.{box_class}").get_attribute("style") or ""
                )
            )
        except TimeoutException:
            print(f"[警告] 等待 .{box_class} 顯示超時，仍繼續操作", flush=True)
    except TimeoutException:
        print(f"無法選擇查詢類型 {query_type_value} 的按鈕（找不到 #{elem_id}）", flush=True)
    except Exception as e:
        print(f"選擇查詢類型 {query_type_value} 時發生錯誤：{e}", flush=True)

def select_district(city, area):
    """選擇縣市+地區下拉。
    🔥 新版高雄站用 id="dist" 的單一下拉，選項格式為「高雄市三民區」（縣市與區合併）。
    🔥 等真實的縣市選項載入（不只是 placeholder「-- 請選擇 --」），且每次都重新抓
       元素以避免 Vue 重新 render 後抓到 stale element。
    """
    district_name = f"{city}{area}"

    def _real_options(d):
        """回傳排除 placeholder 後的選項（過濾掉「請選擇」「-- --」這類）"""
        try:
            opts = Select(d.find_element(By.ID, "dist")).options
            return [o.text.strip() for o in opts
                    if o.text.strip() and '請選擇' not in o.text and o.text.strip() != '--']
        except Exception:
            return []

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "dist"))
        )

        # 🔥 等到真正的縣市選項出現（至少 1 個非 placeholder）
        try:
            WebDriverWait(driver, 15).until(lambda d: len(_real_options(d)) > 0)
        except TimeoutException:
            print("[警告] 等待 dist 真實選項載入超時，目前僅有 placeholder", flush=True)

        # 重新抓元素，避免 Vue 在等待期間重新 render 造成 stale
        dist_select = Select(driver.find_element(By.ID, "dist"))
        options = [opt.text for opt in dist_select.options]
        if district_name in options:
            dist_select.select_by_visible_text(district_name)
            print(f"\033[96m已選擇地區：{district_name}\033[0m", flush=True)
        else:
            real = [o for o in options
                    if o.strip() and '請選擇' not in o and o.strip() != '--']
            print(f"無法找到地區：{district_name}", flush=True)
            print(f"  真實選項共 {len(real)} 個（前 5 個示例）：{real[:5]}", flush=True)

    except TimeoutException:
        print("無法找到或選擇縣市+地區的下拉選單（找不到 #dist）", flush=True)

def select_section(section_name):
    def _real_section_options(d):
        try:
            opts = Select(d.find_element(By.ID, "section")).options
            return [o.text.strip() for o in opts
                    if o.text.strip() and '請選擇' not in o.text and o.text.strip() != '--']
        except Exception:
            return []

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "section"))
        )

        # 🔥 等真實段名選項載入（排除 placeholder）
        try:
            WebDriverWait(driver, 15).until(lambda d: len(_real_section_options(d)) > 0)
        except TimeoutException:
            print("[警告] 等待 section 真實選項載入超時", flush=True)

        # 重新抓元素，避免 stale
        section_select = Select(driver.find_element(By.ID, "section"))
        options = [opt.text for opt in section_select.options]
        if section_name in options:
            section_select.select_by_visible_text(section_name)
            print(f"\033[96m已選擇地段：{section_name}\033[0m", flush=True)
            return True  # 正確跳出
        else:
            print(f"無法找到地段：{section_name}，可用選項為：", flush=True)
            while True: #加入了 While 無限迴圈，強制使用者在這裏輸入。
                for index, option in enumerate(options): #使用 enumerate 添加 index 
                    print(f"{index}. {option}", flush=True)  #印出index 和 option
                
                selected_input = input(f"請根據選項編號手動輸入要選擇的地段選項（0-{len(options) - 1}), 輸入【q】退出當前查詢，或輸入【0】完全退出程式：").strip().lower()
                if selected_input == 'q':
                    print("您選擇退出當前查詢。", flush=True)
                    return False
                elif selected_input == '0':
                    global SHOULD_EXIT_PROGRAM
                    SHOULD_EXIT_PROGRAM = True
                    print("您選擇完全退出程式。", flush=True)
                    return False
                
                if selected_input.isdigit():
                    selected_index = int(selected_input)
                    if 0 <= selected_index < len(options): #確認選擇選項正確
                        section_select.select_by_index(selected_index)
                        print(f"您選擇的地段是: {options[selected_index]}", flush =True)
                        return True
                    else :
                        print("您輸入的數字不在此選單中，請重新輸入:", flush =True);
                
                else:
                    print(f"您輸入的內容錯誤，請重新輸入", flush =True)


    except TimeoutException as e:
        print(f"無法找到或選擇地段的下拉選單，發生錯誤{e}", flush=True)
        return False  #當錯誤發生或是使用者想要手動輸入，一律 return False
       
def fill_lot_number(lot_number):
    try:
        mother_number_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "road_no1"))
        )
        child_number_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "road_no2"))
        )
        
        if "-" in lot_number:
            mother_number, child_number = lot_number.split("-")
            mother_number_input.clear()
            mother_number_input.send_keys(mother_number.strip())
            child_number_input.clear()
            child_number_input.send_keys(child_number.strip())
            print(f"\033[96m已填入母號：{mother_number} 和子號：{child_number}\033[0m", flush=True)
        else:
            mother_number_input.clear()
            mother_number_input.send_keys(lot_number.strip())
            print(f"\033[96m已填入母號：{lot_number}\033[0m", flush=True)

    except TimeoutException:
        print("無法找到母號或子號的輸入框", flush=True)

def handle_no_data_popup():
    """
    處理查無資料的彈窗，並顯示對應的查詢描述。
    改進版：縮短等待時間，增加轉場處理
    """
    global CURRENT_QUERY_DESCRIPTION  # 使用全局變數
    try:
        # 先檢查是否有結果表格出現（這表示查詢成功）
        try:
            results_table = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CLASS_NAME, "licstable"))
            )
            # 如果找到結果表格，滾動以顯示結果並返回False（表示查詢成功）
            scroll_to_show_results(results_table)
            return False
        except TimeoutException:
            # 如果沒有找到結果表格，繼續檢查是否有"查無資料"彈窗
            pass
        
        # 使用較短的等待時間先嘗試定位彈窗
        popup = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "swal2-popup"))
        )
        message_element = popup.find_element(By.ID, "swal2-html-container")
        message_text = message_element.text.strip()

        # 🔥 處理空白彈出視窗
        if not message_text or not message_element.is_displayed():
            print("偵測到空白警告彈出視窗，關閉並重試...", flush=True)
            ok_button = popup.find_element(By.CLASS_NAME, "swal2-confirm")
            ok_button.click()
            time.sleep(1)
            return "blank"  # 返回 "blank" 表示空白彈窗（需要重新點擊查詢）

        if "查無資料" in message_text:
            query_desc = CURRENT_QUERY_DESCRIPTION or "未知查詢類型"
            print(f"\033[93m[{query_desc}] 查無資料\033[0m", flush=True)
            print("⏰ 彈窗將停留 3 秒供您查看...", flush=True)

            # 🔥 使用 JavaScript 停止 SweetAlert2 的自動關閉計時器
            try:
                driver.execute_script("""
                    if (window.Swal && window.Swal.stopTimer) {
                        window.Swal.stopTimer();
                    }
                """)
            except:
                pass

            # 🔥 使用明確的等待並輸出倒數計時
            import time as time_module
            for i in range(3, 0, -1):
                print(f"⏰ {i} 秒後將關閉彈窗...", flush=True)
                time_module.sleep(1)

            ok_button = popup.find_element(By.CLASS_NAME, "swal2-confirm")
            ok_button.click()
            print("關閉訊息", flush=True)
            time.sleep(1)
            return True  # 表示查無資料
    except TimeoutException:
        # 如果短時間內沒有找到彈窗，再次檢查是否有結果表格
        try:
            results_table = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "licstable"))
            )
            # 如果找到結果表格，滾動以顯示結果並返回False（表示查詢成功）
            scroll_to_show_results(results_table)
            return False
        except TimeoutException:
            # 如果還是沒有找到結果表格，使用較長時間等待彈窗
            try:
                popup = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, "swal2-popup"))
                )
                message_element = popup.find_element(By.ID, "swal2-html-container")
                message_text = message_element.text.strip()

                # 🔥 處理空白彈出視窗
                if not message_text or not message_element.is_displayed():
                    print("偵測到空白警告彈出視窗，關閉並重試...", flush=True)
                    ok_button = popup.find_element(By.CLASS_NAME, "swal2-confirm")
                    ok_button.click()
                    time.sleep(1)
                    return "blank"  # 返回 "blank" 表示空白彈窗（需要重新點擊查詢）

                if "查無資料" in message_text:
                    query_desc = CURRENT_QUERY_DESCRIPTION or "未知查詢類型"
                    print(f"\033[93m[{query_desc}] 查無資料\033[0m", flush=True)
                    print("⏰ 彈窗將停留 3 秒供您查看...", flush=True)

                    # 🔥 使用 JavaScript 停止 SweetAlert2 的自動關閉計時器
                    try:
                        driver.execute_script("""
                            if (window.Swal && window.Swal.stopTimer) {
                                window.Swal.stopTimer();
                            }
                        """)
                    except:
                        pass

                    # 🔥 使用明確的等待並輸出倒數計時
                    import time as time_module
                    for i in range(3, 0, -1):
                        print(f"⏰ {i} 秒後將關閉彈窗...", flush=True)
                        time_module.sleep(1)

                    ok_button = popup.find_element(By.CLASS_NAME, "swal2-confirm")
                    ok_button.click()
                    print("關閉訊息", flush=True)
                    time.sleep(1)
                    return True  # 表示查無資料
            except TimeoutException:
                # 如果長時間等待後仍找不到彈窗，可能是系統正在處理大量數據
                print("未找到查詢結果或彈窗，可能正在處理數據，請稍等...", flush=True)
                # 最後一次嘗試查找結果表格
                try:
                    results_table = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "licstable"))
                    )
                    scroll_to_show_results(results_table)
                    return False
                except TimeoutException:
                    print("查詢超時，未能顯示結果。", flush=True)
    
    return False  # 預設返回False表示查詢成功或未知狀態

def scroll_to_show_results(results_table):
    """滾動頁面以顯示查詢結果表格"""
    try:
        driver.execute_script("arguments[0].scrollIntoView(true);", results_table)
        print("已滾動頁面以顯示查詢結果。", flush=True)

        # 嘗試統計結果數量並顯示
        rows = results_table.find_elements(By.CLASS_NAME, "datarow")
        print(f"共找到 {len(rows)} 筆資料", flush=True)

        # 🔥 如果結果數量為 0，檢查是否有「查無資料」彈窗
        if len(rows) == 0:
            print("偵測到空結果表格，檢查是否有彈窗...", flush=True)
            time.sleep(0.5)  # 短暫等待彈窗出現
            try:
                popup = driver.find_element(By.CLASS_NAME, "swal2-popup")
                if popup.is_displayed():
                    message_element = popup.find_element(By.ID, "swal2-html-container")
                    message_text = message_element.text.strip()

                    # 🔥 處理空白彈出視窗
                    if not message_text or not message_element.is_displayed():
                        print("偵測到空白警告彈出視窗，關閉...", flush=True)
                        ok_button = popup.find_element(By.CLASS_NAME, "swal2-confirm")
                        ok_button.click()
                        return

                    if "查無資料" in message_text:
                        query_desc = CURRENT_QUERY_DESCRIPTION or "未知查詢類型"
                        print(f"\033[93m[{query_desc}] 查無資料\033[0m", flush=True)
                        print("⏰ 彈窗將停留 3 秒供您查看...", flush=True)

                        # 倒數計時
                        import time as time_module
                        for i in range(3, 0, -1):
                            print(f"⏰ {i} 秒後將關閉彈窗...", flush=True)
                            time_module.sleep(1)

                        ok_button = popup.find_element(By.CLASS_NAME, "swal2-confirm")
                        ok_button.click()
                        print("關閉訊息", flush=True)
            except:
                pass

        # 如果結果數量大於50，顯示提示信息
        if len(rows) > 50:
            print("注意：查詢結果較多，請稍候查詢結果...", flush=True)

    except Exception as e:
        print(f"滾動顯示結果時發生錯誤：{e}", flush=True)

def handle_verification_error_popup(silent=False):
    """
    處理驗證碼錯誤彈窗

    Args:
        silent: 如果為 True，不輸出訊息（用於清理時）

    Returns:
        "error" - 驗證碼錯誤（需重新辨識驗證碼）
        "blank" - 空白彈窗（需重新點擊查詢按鈕）
        None - 沒有彈窗
    """
    try:
        popup = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "swal2-popup"))
        )
        message_element = popup.find_element(By.ID, "swal2-html-container")
        message_text = message_element.text.strip()

        # 處理驗證碼錯誤
        if "驗證碼錯誤" in message_text:
            if not silent:
                print("驗證碼錯誤", flush=True)
            ok_button = popup.find_element(By.CLASS_NAME, "swal2-confirm")
            ok_button.click()
            if not silent:
                print("已點擊彈窗上的 OK 按鈕", flush=True)
            return "error"
        elif not message_text or not message_element.is_displayed():
            # 空白彈出視窗（訊息被隱藏或為空）
            if not silent:
                print("偵測到空白警告彈出視窗，準備重新查詢...", flush=True)
            ok_button = popup.find_element(By.CLASS_NAME, "swal2-confirm")
            ok_button.click()
            if not silent:
                print("已關閉空白彈出視窗", flush=True)
            return "blank"
    except TimeoutException:
        pass
    return None

def click_query_button():
    try:
        # 🔥 先滾動到驗證碼區域（讓視覺上看得到），新站 canvas id="myCanvas"
        #    若找不到（網站結構改動）就退回滾動到查詢按鈕本身，且靜默失敗不印 stack trace
        try:
            canvas = driver.find_element(By.ID, "myCanvas")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", canvas)
            time.sleep(0.3)
        except Exception:
            try:
                btn = driver.find_element(By.ID, "btnLogin")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                time.sleep(0.3)
            except Exception:
                pass  # 靜默失敗，scroll 只是讓人看得到，失敗不影響功能
        
        query_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "btnLogin"))
        )
        try:
            query_button.click()
            print("已點擊查詢按鈕，準備驗證碼程序1...", flush=True)

            # 🔥 點擊後立即注入 JavaScript 監聽器，阻止彈窗自動關閉
            try:
                driver.execute_script("""
                    // 覆蓋 Swal.fire 方法，自動停止計時器
                    if (window.Swal) {
                        const originalFire = window.Swal.fire;
                        window.Swal.fire = function(...args) {
                            const result = originalFire.apply(this, args);
                            if (window.Swal.stopTimer) {
                                window.Swal.stopTimer();
                            }
                            return result;
                        };
                    }
                """)
            except:
                pass

        except ElementClickInterceptedException:
            print("查詢按鈕被彈窗遮擋，嘗試關閉彈窗後再點擊", flush=True)
            popup_result = handle_no_data_popup()
            # 如果是空白彈窗，已經關閉了，繼續點擊查詢按鈕
            query_button.click()
            print("已點擊查詢按鈕，準備驗證碼程序2...", flush=True)

            # 🔥 同樣在這裡也注入監聽器
            try:
                driver.execute_script("""
                    if (window.Swal) {
                        const originalFire = window.Swal.fire;
                        window.Swal.fire = function(...args) {
                            const result = originalFire.apply(this, args);
                            if (window.Swal.stopTimer) {
                                window.Swal.stopTimer();
                            }
                            return result;
                        };
                    }
                """)
            except:
                pass

    except TimeoutException:
        print("無法找到查詢按鈕", flush=True)

def _try_get_captcha_from_playlink():
    """🔥 高雄新站捷徑：從「語音播放」連結 #playcaptcha 的 href 直接取出驗證碼答案
    （網站把答案寫在 google_translate_tts 的 q= 參數裡，原本是給螢幕閱讀器用）。

    成功回傳驗證碼字串、失敗回傳 None；不丟例外，外層可安全 fall back 到 OCR。
    """
    try:
        link = driver.find_element(By.ID, "playcaptcha")
        href = link.get_attribute("href") or ""
        # q 參數可能被包在引號或 URL-encode 過的引號裡
        m = re.search(r'[?&]q=(?:["\']|%22|%27)*([A-Za-z0-9]+)', href)
        if m:
            code = m.group(1)
            if 4 <= len(code) <= 10:  # 合理長度過濾
                return code
    except Exception:
        pass
    return None


def get_verification_code():
    """
    擷取驗證碼。優先從 #playcaptcha 連結直接取，失敗再退回 RapidOCR 辨識 canvas。
    """
    # 🔥 優先用 playcaptcha href 捷徑（不用 OCR、瞬間取得、永遠正確）
    quick_code = _try_get_captcha_from_playlink()
    if quick_code:
        print(f"✓ 已從 #playcaptcha 連結取得驗證碼：{quick_code}", flush=True)
        return quick_code

    print("[退回] playcaptcha 捷徑失敗，改用 OCR 辨識...", flush=True)

    try:
        # 首先確認我們在哪種查詢模式下
        print("檢查當前查詢模式...", flush=True)
        current_query_type = CURRENT_QUERY_DESCRIPTION
        print(f"當前查詢模式: {current_query_type}", flush=True)

        # 確保頁面已經完全載入
        print("等待頁面完全載入...", flush=True)
        time.sleep(2)  # 增加等待時間，確保頁面元素已完全載入
        
        # 確保驗證碼區域可見
        print("尋找驗證碼輸入框...", flush=True)
        try:
            code_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "inputCode"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", code_input)
            print("已滾動到驗證碼輸入框", flush=True)
            time.sleep(1)  # 等待滾動完成
        except Exception as e:
            print(f"無法找到驗證碼輸入框: {e}", flush=True)
        
        # 再次尋找Canvas元素 - 使用較寬鬆的定位方式
        print("尋找驗證碼Canvas元素...", flush=True)
        canvas = None
        try:
            # 嘗試使用原始XPath
            canvas = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div[10]/div[2]/canvas"))
            )
            print("使用原始XPath成功找到Canvas", flush=True)
        except:
            try:
                # 嘗試使用更簡單的定位方式 - 所有Canvas元素
                canvases = driver.find_elements(By.TAG_NAME, "canvas")
                if canvases:
                    canvas = canvases[0]  # 假設第一個Canvas是驗證碼
                    print(f"找到 {len(canvases)} 個Canvas元素，使用第一個", flush=True)
                else:
                    # 嘗試其他可能的XPath或CSS選擇器
                    try:
                        canvas = driver.find_element(By.CSS_SELECTOR, "canvas")
                        print("使用CSS選擇器找到Canvas", flush=True)
                    except:
                        print("無法找到任何Canvas元素", flush=True)
            except Exception as e:
                print(f"使用替代方法查找Canvas失敗: {e}", flush=True)
        
        # 如果找到Canvas，使用 JavaScript 獲取圖片數據（不受縮放影響）
        if canvas:
            print("使用JavaScript從Canvas提取圖片數據...", flush=True)
            try:
                # 🔥 使用 JavaScript 直接從 Canvas 獲取 base64 圖片數據
                # 這種方法不受瀏覽器縮放影響，獲取的是原始圖片
                canvas_base64 = driver.execute_script(
                    "return arguments[0].toDataURL('image/png').substring(22);",
                    canvas
                )

                # 將 base64 轉換為圖片
                import base64
                canvas_png = base64.b64decode(canvas_base64)
                image = Image.open(BytesIO(canvas_png))
                print(f"Canvas圖片提取成功，尺寸: {image.size}", flush=True)
                
                # 不再保存除錯圖片（減少磁碟寫入）

                # 使用 RapidOCR 辨識驗證碼 (延遲載入以提升啟動速度)
                print("初始化 RapidOCR（驗證碼優化參數）...", flush=True)
                from rapidocr_helper import create_rapidocr_reader, rapidocr_readtext
                # 🔧 針對驗證碼優化的參數：
                # - 降低 det_db_thresh 提高文字偵測敏感度
                # - 降低 det_db_box_thresh 接受更多候選框
                # - 降低 text_score 接受更多識別結果
                # - 提高 max_side_len 保持圖片清晰度
                # - 關閉 use_angle_cls 加快速度（驗證碼通常不旋轉）
                reader = create_rapidocr_reader(
                    use_angle_cls=False,        # 驗證碼不需要方向校正
                    use_text_score=True,
                    text_score=0.3,             # 降低置信度閾值（接受更多結果）
                    det_db_thresh=0.2,          # 降低二值化閾值（更敏感）
                    det_db_box_thresh=0.3,      # 降低文字框閾值（提高檢出率）
                    det_db_unclip_ratio=1.8,    # 擴大文字框範圍
                    max_side_len=1280           # 提高圖片解析度
                )
                print("開始辨識驗證碼...", flush=True)
                # 🔥 使用 detail=1 獲取完整結果（包含座標），以便根據 x 座標排序
                result_with_coords = rapidocr_readtext(reader, np.array(image), detail=1)
                print(f"RapidOCR原始結果: {result_with_coords}", flush=True)

                if not result_with_coords:
                    print("未辨識到任何文字", flush=True)
                    return None

                # 🔥 根據文字框的 x 座標（左上角）排序，確保從左到右的順序
                # result_with_coords 格式: [[[x1,y1],[x2,y2],[x3,y3],[x4,y4]], text, confidence]
                sorted_results = sorted(result_with_coords, key=lambda item: item[0][0][0])  # 按第一個點的 x 座標排序

                # 提取排序後的文字
                sorted_texts = [item[1] for item in sorted_results]
                print(f"識別的驗證碼列表(按x座標排序): {sorted_texts}", flush=True)

                # 獲取辨識結果並去除空白與非預期字符
                captcha_text = ''.join(sorted_texts).replace(" ", "").strip()
                captcha_text = ''.join(filter(str.isalnum, captcha_text))  # 僅保留字母和數字
                print(f"識別的驗證碼為: {captcha_text}", flush=True)
                
                return captcha_text if captcha_text else None
            except Exception as e:
                print(f"Canvas截圖或辨識過程出錯: {e}", flush=True)
        
        # 如果直接截取Canvas失敗，改為截取整個頁面然後裁剪（在記憶體中處理，不保存檔案）
        print("直接截取Canvas失敗，嘗試整頁截圖方法...", flush=True)
        try:
            # 如果Canvas在前面已找到，使用其位置裁剪
            if canvas:
                canvas_location = canvas.location
                canvas_size = canvas.size
                print(f"Canvas位置: {canvas_location}, 大小: {canvas_size}", flush=True)

                # 🔥 直接取得截圖的 bytes，在記憶體中處理（不保存到磁碟）
                screenshot_bytes = driver.get_screenshot_as_png()
                full_img = Image.open(BytesIO(screenshot_bytes))

                left = canvas_location['x']
                top = canvas_location['y']
                right = left + canvas_size['width']
                bottom = top + canvas_size['height']

                # 裁剪驗證碼區域（在記憶體中）
                captcha_img = full_img.crop((left, top, right, bottom))
                print(f"已在記憶體中裁剪驗證碼圖片", flush=True)

                # 使用RapidOCR辨識 (延遲載入以提升啟動速度)
                from rapidocr_helper import create_rapidocr_reader, rapidocr_readtext
                reader = create_rapidocr_reader(
                    use_angle_cls=False,
                    use_text_score=True,
                    text_score=0.3,
                    det_db_thresh=0.2,
                    det_db_box_thresh=0.3,
                    det_db_unclip_ratio=1.8,
                    max_side_len=1280
                )
                result = rapidocr_readtext(reader, np.array(captcha_img), detail=0)
                print(f"裁剪圖片辨識結果: {result}", flush=True)

                # 處理辨識結果
                captcha_text = ''.join(result).replace(" ", "").strip()
                captcha_text = ''.join(filter(str.isalnum, captcha_text))  # 僅保留字母和數字
                print(f"裁剪圖片處理後的驗證碼: {captcha_text}", flush=True)

                return captcha_text if captcha_text else None
            else:
                print("無法找到Canvas位置來裁剪整頁截圖", flush=True)
        except Exception as e:
            print(f"整頁截圖方法失敗: {e}", flush=True)
        
        # 如果都失敗了，告知使用者需要手動輸入
        print("所有辨識方法均失敗，將提示使用者手動輸入", flush=True)
        return None
        
    except Exception as e:
        print(f"驗證碼擷取過程中發生未預期錯誤: {e}", flush=True)
        return None
    
def fill_verification_code(code):
    try:
        # 診斷：顯示原始驗證碼值（包含任何隱藏字元）
        print(f"[診斷] 準備填入驗證碼: {repr(code)} (長度: {len(code)})", flush=True)

        # 確保驗證碼已經過 strip 處理
        code = code.strip()
        print(f"[診斷] Strip 後驗證碼: {repr(code)} (長度: {len(code)})", flush=True)

        # 等待驗證碼輸入框出現
        code_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "inputCode"))
        )

        # 確保元素可見且可互動
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "inputCode"))
        )

        # 捲動到元素位置確保可見
        driver.execute_script("arguments[0].scrollIntoView(true);", code_input)
        time.sleep(0.5)

        # 清空並填入驗證碼
        code_input.clear()
        time.sleep(0.2)  # 等待清空完成
        code_input.send_keys(code)
        time.sleep(0.2)  # 等待填入完成

        # 驗證是否填入成功
        filled_value = code_input.get_attribute('value')
        print(f"[診斷] 填入後欄位值: {repr(filled_value)} (長度: {len(filled_value)})", flush=True)

        if filled_value == code:
            print(f"✓ 驗證碼 [{code}]: 已成功填入驗證碼欄位", flush=True)
        else:
            print(f"⚠ 驗證碼填入異常 - 預期: {repr(code)}, 實際: {repr(filled_value)}", flush=True)
            # 重試一次
            code_input.clear()
            time.sleep(0.2)
            code_input.send_keys(code)
            time.sleep(0.2)
            filled_value = code_input.get_attribute('value')
            print(f"✓ 重試後驗證碼: {repr(filled_value)}", flush=True)

        # 🔥 等待網頁 JavaScript 完成驗證碼欄位的處理
        # 這很重要！網頁可能需要時間來驗證輸入的驗證碼格式
        print("[診斷] 等待網頁處理驗證碼...", flush=True)
        time.sleep(1.5)  # 給網頁足夠時間處理驗證碼

    except TimeoutException:
        print("✗ 錯誤：無法找到驗證碼輸入框（Timeout）", flush=True)
    except Exception as e:
        print(f"✗ 錯誤：填入驗證碼時發生異常 - {e}", flush=True)

def attempt_verification():
    global SHOULD_EXIT_PROGRAM
    max_attempts = 3
    print(f"開始驗證邏輯，最多嘗試 {max_attempts} 次。", flush=True)
    time.sleep(3)
    for attempt in range(1, max_attempts + 1):
        try:
            code = get_verification_code()
            print(f"\033[93m驗證碼第 {attempt} 次嘗試：{code}\033[0m", flush=True)

            if not code:
                print(f"驗證碼辨識失敗（第 {attempt}/{max_attempts} 次）", flush=True)
                # 如果是前3次自動嘗試，直接 continue 重試
                if attempt < max_attempts:
                    print("重新嘗試辨識...", flush=True)
                    try:
                        # 重新載入頁面
                        driver.refresh()
                        time.sleep(3)
                        print("頁面已重新載入", flush=True)
                    except Exception as e:
                        print(f"重新載入頁面失敗: {e}", flush=True)
                    continue  # 跳過本次，進入下一次自動嘗試
                else:
                    # 第3次也失敗，進入手動模式
                    print(f"前 {max_attempts} 次自動辨識均失敗，無法取得驗證碼", flush=True)
                    # 不需要在這裡手動輸入，後面會進入無限重試的手動模式
                    break

            fill_verification_code(code)
            click_query_button()

            # 後續的驗證邏輯保持不變... #驗證邏輯跑完再進行點擊。

            # 1. 先檢查是否有驗證碼錯誤或空白彈窗
            time.sleep(3)
            popup_result = handle_verification_error_popup()

            if popup_result == "error":
                # 驗證碼錯誤，需重新辨識驗證碼
                print(f"驗證碼錯誤，第 {attempt} 次失敗。", flush=True)
                continue
            elif popup_result == "blank":
                # 空白彈窗，重新點擊查詢按鈕
                print("偵測到空白彈窗，自動重新點擊查詢按鈕...", flush=True)
                time.sleep(1)
                click_query_button()
                time.sleep(3)  # 等待查詢結果
                # 繼續後續的查詢結果檢查

            # 2. 檢查是否有查無資料彈窗，或者查詢結果
            no_data_result = handle_no_data_popup()
            if no_data_result == "blank":
                # 空白彈窗，重新點擊查詢按鈕
                print("偵測到空白彈窗（來自 handle_no_data_popup），自動重新點擊查詢按鈕...", flush=True)
                time.sleep(1)
                click_query_button()
                time.sleep(3)  # 等待查詢結果
                # 重新檢查查詢結果
                no_data_result = handle_no_data_popup()

            if no_data_result == True:
                print("查無資料。", flush=True)
                return "no_data"
            elif no_data_result == False:
                # 如果沒有彈窗，可能是查詢成功或正在處理中
                try:
                    # 等待結果表格出現，表示查詢成功
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "licstable"))
                    )
                    return "success"
                except TimeoutException:
                    # 使用較長時間再次檢查是否有"查無資料"彈窗
                    time.sleep(5)
                    no_data_result = handle_no_data_popup()
                    if no_data_result == "blank":
                        # 空白彈窗，重新點擊查詢按鈕
                        print("偵測到空白彈窗（第二次檢查），自動重新點擊查詢按鈕...", flush=True)
                        time.sleep(1)
                        click_query_button()
                        time.sleep(3)  # 等待查詢結果
                        # 重新檢查查詢結果
                        no_data_result = handle_no_data_popup()

                    if no_data_result == True:
                        print("查無資料。", flush=True)
                        return "no_data"
                    else:
                        # 如果依然沒有彈窗也沒有結果表格，可能是系統正在處理
                        print("未偵測到明確回應，再次檢查...", flush=True)
                        time.sleep(3)

                        # 最後一次等待結果
                        try:
                            WebDriverWait(driver, 15).until(
                                EC.presence_of_element_located((By.CLASS_NAME, "licstable"))
                            )
                            return "success"
                        except TimeoutException:
                            print(f"查詢超時（第 {attempt}/{max_attempts} 次），將重試", flush=True)
           
        except Exception as e:
            print(f"驗證過程中出現錯誤：{e}", flush=True)
            
    print("所有驗證嘗試均失敗。", flush=True)
    print("進入手動驗證模式...", flush=True)
    print("正在清理頁面狀態...", flush=True)

    # 進入手動模式前，先處理可能殘留的彈窗（但不重新載入頁面，以保留已填入的資料）
    try:
        handle_verification_error_popup(silent=True)  # 靜默關閉可能的錯誤彈窗
    except Exception as e:
        print(f"清理彈窗失敗: {e}", flush=True)

    # 第3次失敗後，進入無限循環手動模式，直到成功或使用者選擇離開
    manual_attempt = 1
    while True:
        print(f"\n=== 手動驗證第 {manual_attempt} 次 ===", flush=True)

        # 自動辨識驗證碼並帶入主程式輸入框供使用者確認
        print("正在自動辨識驗證碼...", flush=True)
        auto_code = get_verification_code()

        if auto_code:
            print(f"\033[93m自動辨識到驗證碼: {auto_code}\033[0m", flush=True)
            print("請檢查驗證碼是否正確：", flush=True)
            print("  - 若正確，請直接按【ENTER】送出", flush=True)
            print("  - 若需修改，請輸入正確的驗證碼後按【ENTER】", flush=True)
            print("  - 輸入【ESC】可離開程式", flush=True)

            # 發送特殊標記給主程式，通知預填驗證碼到 GUI 輸入框
            print(f"__GUI_PREFILL__:{auto_code}", flush=True)

            # 讓使用者在主程式輸入框中確認或修改
            print(f"驗證碼 [{auto_code}]: ", flush=True)
            user_input = safe_input().strip()

            if user_input.upper() == "ESC":
                return "manual_exit"
            elif user_input:  # 使用者輸入了修改內容
                code = user_input
            else:  # 使用者直接按ENTER，使用自動辨識的驗證碼
                code = auto_code
        else:
            # 自動辨識失敗，請求手動輸入
            print("自動辨識失敗，請手動輸入驗證碼或按 ESC 離開程式：", flush=True)
            print("驗證碼: ", flush=True)
            code = safe_input().strip()

            if code.upper() == "ESC":
                return "manual_exit"

        # 將最終確認的驗證碼填入網頁
        fill_verification_code(code)
        click_query_button()  # 送出查詢

        # 檢查結果
        # 1. 先檢查是否有驗證碼錯誤或空白彈窗
        time.sleep(3)
        popup_result = handle_verification_error_popup()

        if popup_result == "error":
            # 驗證碼錯誤，需重新辨識驗證碼
            print(f"驗證碼錯誤，第 {manual_attempt} 次手動嘗試失敗。", flush=True)
            manual_attempt += 1
            continue  # 重新循環，再次嘗試
        elif popup_result == "blank":
            # 空白彈窗，重新點擊查詢按鈕
            print("偵測到空白彈窗，自動重新點擊查詢按鈕...", flush=True)
            time.sleep(1)
            click_query_button()
            time.sleep(3)  # 等待查詢結果
            # 繼續後續的查詢結果檢查

        # 2. 檢查是否有查無資料彈窗
        no_data_result = handle_no_data_popup()
        if no_data_result == "blank":
            # 空白彈窗，重新點擊查詢按鈕
            print("偵測到空白彈窗（來自 handle_no_data_popup），自動重新點擊查詢按鈕...", flush=True)
            time.sleep(1)
            click_query_button()
            time.sleep(3)  # 等待查詢結果
            # 重新檢查查詢結果
            no_data_result = handle_no_data_popup()

        if no_data_result == True:
            print("查無資料。", flush=True)
            return "no_data"

        # 3. 等待結果表格出現
        try:
            WebDriverWait(driver, 25).until(
                EC.presence_of_element_located((By.CLASS_NAME, "licstable"))
            )
            print(f"✅ 驗證成功！（手動第 {manual_attempt} 次）", flush=True)
            return "success"
        except TimeoutException:
            # 再次檢查是否有彈窗
            time.sleep(5)
            no_data_result = handle_no_data_popup()
            if no_data_result == "blank":
                # 空白彈窗，重新點擊查詢按鈕
                print("偵測到空白彈窗（Timeout後檢查），自動重新點擊查詢按鈕...", flush=True)
                time.sleep(1)
                click_query_button()
                time.sleep(3)  # 等待查詢結果
                # 重新檢查查詢結果
                no_data_result = handle_no_data_popup()

            if no_data_result == True:
                print("查無資料。", flush=True)
                return "no_data"
            else:
                print(f"查詢超時，第 {manual_attempt} 次手動嘗試可能失敗。", flush=True)
                manual_attempt += 1
                continue  # 重新循環，再次嘗試


def fill_license_number(license_number):
    """
    填寫執照號碼到查詢欄位。
    """
    try:
        pass
        # print(f"嘗試執行：{license_number}", flush=True)
        # license_input = WebDriverWait(driver, 5).until(
        #     EC.presence_of_element_located((By.ID, "license_number"))
        # )
        # driver.execute_script("arguments[0].scrollIntoView(true);", license_input)  # 確保可見
        # license_input.clear()
        # license_input.send_keys(license_number)
        # print(f"已成功填寫執照號碼：{license_number}", flush=True)
    except TimeoutException:
        print("無法找到", flush=True)

def scroll_until_element_visible(element, max_attempts=10):
    for attempt in range(max_attempts):
        if element.is_displayed():
            return True
        driver.execute_script("window.scrollBy(0, 200);")  # 每次滾動 200 像素
        time.sleep(0.5)  # 等待頁面載入
    return False

def perform_search(query_type, city, area, section, lot_number, address=None, license_number=None):
    global CURRENT_QUERY_DESCRIPTION
    global SHOULD_EXIT_PROGRAM  # 增加這一行

    query_type_descriptions = {
        "4": "建築地號",
        "2": "建築地址",
        "3": "手動輸入"
    }
    CURRENT_QUERY_DESCRIPTION = query_type_descriptions.get(query_type, "未知查詢類型")

    print(f"執行查詢類型：{CURRENT_QUERY_DESCRIPTION} ({query_type})", flush=True)

    result = "no_data"
    results = []

    # 🔥 內部 mode 編號對應到新站實際 radio 的 ID 編號（OLD 站順序與 NEW 站不同）
    #    NEW 站：QType1=執照號碼 / QType2=起造人 / QType3=建築地址 / QType4=建築地號
    #    OLD 站順序不同，所以這裡要做轉換
    NEW_QTYPE_FOR_MODE = {
        "4": "4",   # 建築地號 → 新站 QType4
        "2": "3",   # 建築地址 → 新站 QType3（OLD 站是 span[2]，NEW 站變 QType3）
    }

    if query_type in NEW_QTYPE_FOR_MODE:
        select_query_type(NEW_QTYPE_FOR_MODE[query_type])
    elif query_type == "3":
        # 手動模式：不預先點任何 radio，讓使用者在網頁上自行選擇要用哪種查詢
        print("[手動模式] 不預先點選任何查詢類型，請於網頁上自行選擇", flush=True)
    else:
        # 其他未預期的 query_type，仍嘗試直接用其值當 radio 編號
        select_query_type(query_type)

    if query_type == "4":  # 地號查詢
      print("執行地號查詢邏輯...", flush=True)
      select_district(city, area)
      section_result = select_section(section) #加入判斷式，讓程式在沒有選取成功就直接結束
      if not section_result:
            return "no_data", results
    
      fill_lot_number(lot_number)  #如果執行完以上都沒錯，再執行地號輸入。
      result = attempt_verification()

      if result =="success":  # 成功的話就會解析和跳出， 把資料連同 result 回傳。
         results = parse_search_results()
         return result, results #直接return結果，讓test決定後續動作。

       #如果是失敗的話什麼也不做
       

    elif query_type == "2":  # 建築地址查詢
        print("執行地址查詢邏輯...", flush=True)
        result = perform_address_search(city, area, address)  # 回傳值會直接到 test  裡面處理判斷式
    elif query_type == "3":  # 手動查詢
        print("執行手動查詢...", flush=True)
        print("\033[93m【注意】：\n輸入【q】可退出當前查詢，回主選單\n輸入【0】完全退出子程式\033[0m", flush=True)
        if not license_number:
            print("\033[93m請於網頁輸入相關資料，驗證碼不需輸入\n完成後請按下【ENTER】執行查詢\033[0m", flush=True)
            user_input = safe_input().strip().lower()  # 轉為小寫處理
            if user_input == 'q':
                return "manual_exit", []
            elif user_input == '0':
                SHOULD_EXIT_PROGRAM = True
                return "manual_exit", []
            elif user_input:  # 如果用戶輸入了其他內容而不是直接按ENTER
                print("請直接按ENTER繼續，或輸入q/0退出。", flush=True)
                return "manual_exit", []  # 因為輸入了不正確的內容，退出當前查詢

        while True:
            result = attempt_verification()
            
            if result == "manual_exit" or result == "success":
                break
            elif result == "no_data":
                print("\033[93m[手動輸入] 查無資料，請繼續於網頁修改資料，輸入【q】退出當前查詢，或輸入【0】完全退出程式\033[0m", flush=True)
                print("\033[93m請繼續修改資料後按【ENTER】，輸入【q】退出當前查詢，或輸入【0】完全退出程式：\033[0m", flush=True)
                user_input = safe_input().strip()
                if user_input == 'q':
                    return "manual_exit", []
                elif user_input == '0':
                    SHOULD_EXIT_PROGRAM = True
                    return "manual_exit", []


    # 若查詢成功，解析查詢結果 
    if result == "success":  #如果到這邊流程都正確，就解析後回傳。
      results = parse_search_results()  

    return result, results #不管哪一種都回傳資料，由 Test 函數判斷後續。



def check_exit_command(user_input):
    """檢查用戶輸入是否為退出命令"""
    if not user_input:
        return False
    user_input = user_input.strip().lower()  # 轉為小寫
    # 'q'表示退出當前查詢，'0'表示完全退出程式
    if user_input == '0':
        global SHOULD_EXIT_PROGRAM
        SHOULD_EXIT_PROGRAM = True
        return True
    return user_input.upper() == 'q'  # 返回當前查詢的退出狀態

def parse_search_results():
    try:
        # 定位表格
        results_table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "licstable"))
        )
        
        # 確保結果表格可見
        scroll_to_show_results(results_table)
        
        rows = results_table.find_elements(By.CLASS_NAME, "datarow")
        
        if not rows:
            print("查詢結果為空。", flush=True)
            return []

        valid_results = []
        # print("原始查詢結果如下：", flush=True)
        for index, row in enumerate(rows):
            try:
                # 獲取當前行的所有欄位
                columns = row.find_elements(By.TAG_NAME, "td")

                # 🔥 執照明細的連結機制會因網站改版而不同，需掃整列同時支援兩種：
                #    (A) 舊版/他縣市：第一欄 <a href="http...">（用 GET 跳轉）
                #    (B) 新版高雄 buildmis：<form action=... method=POST>+CSRF hidden+<button class=link-btn>，
                #        target=_blank 開新分頁。這種沒有 <a>，且需 POST+CSRF，只能「點擊按鈕」讓瀏覽器自帶 token 送出。
                link_element = None   # 可點擊/跳轉的元素
                link_href = None      # 記錄用途：GET 網址 或 表單 action
                link_mode = None      # 'get'（<a href>）或 'click'（POST 送出按鈕）
                license_number = columns[0].text.strip()
                # (A) 先找一般 <a href>
                for a in row.find_elements(By.TAG_NAME, "a"):
                    href = (a.get_attribute("href") or "").strip()
                    if href.lower().startswith(("http://", "https://")):
                        link_element, link_href, link_mode = a, href, 'get'
                        if a.text.strip():
                            license_number = a.text.strip()
                        break
                # (B) 沒 <a> → 找 POST 表單的送出按鈕
                if link_element is None:
                    btns = row.find_elements(
                        By.CSS_SELECTOR, "form button, button.link-btn, form input[type='submit']")
                    if btns:
                        btn = btns[0]
                        link_element, link_mode = btn, 'click'
                        try:
                            form = btn.find_element(By.XPATH, "./ancestor::form")
                            link_href = (form.get_attribute("action") or "").strip() or "(POST表單)"
                        except Exception:
                            link_href = "(POST表單)"
                        if btn.text.strip():
                            license_number = btn.text.strip()
                # (C) 兩種都沒有 → 真的無連結，印診斷 HTML 供日後追
                if link_element is None:
                    link_href = "無"
                    try:
                        print(f"[診斷] 此列找不到連結，列HTML(前600字)：{(row.get_attribute('outerHTML') or '')[:600]}", flush=True)
                    except Exception:
                        pass

                # 提取其他欄位數據
                issue_date = columns[4].text.strip() if len(columns) > 4 else "N/A"
                status = columns[5].text.strip() if len(columns) > 5 else "N/A"

                # 收集有效結果 - 这里修改为使用数字作为选项索引
                valid_results.append(
                    {
                        "option": str(index + 1),  # 使用数字索引，从1开始
                        "license_number": license_number,
                        "issue_date": issue_date,
                        "status": status,
                        "link_element": link_element,
                        "link_href": link_href,
                        "link_mode": link_mode,
                    }
                )
            except Exception as e:
                print(f"解析第 {index + 1} 行時發生錯誤：{e}", flush=True)
                continue

        return valid_results

    except TimeoutException:
        print("未找到查詢結果表格，可能無資料或網頁結構變更。", flush=True)
    except Exception as e:
        print(f"解析查詢結果時發生錯誤：{e}", flush=True)

    return []

# 修改print_results_with_options函数，以更好的方式显示数字选项
def print_results_with_options(all_results):
    """
    列印所有查詢結果並附加選項標記。
    包含"使"字的資料以亮藍色顯示。
    """
    if not all_results:
        print("沒有可用的查詢結果。", flush=True)
        return

    print("\n查詢結果如下：", flush=True)
    for option, result in all_results.items():
        license_number = result.get("license_number", "N/A")
        issue_date = result.get("issue_date", "N/A")
        status = result.get("status", "N/A")

        # 🔥 檢查執照號碼中是否包含"使"字
        # 如果包含，使用亮藍色 (\033[96m) 顯示整行
        result_text = f"({option}) {license_number} | {issue_date} | {status}"
        if "使" in license_number:
            print(f"\033[96m{result_text}\033[0m", flush=True)  # 亮藍色
        else:
            print(result_text, flush=True)  # 正常顏色


def scroll_to_element(element):
    """滾動頁面直到元素可見"""
    try:
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(1)  # 等待滾動完成
        return True
    except Exception as e:
        print(f"滾動到元素時發生錯誤：{e}", flush=True)
        return False


def save_page_as_pdf(filename):
    # 先產生 PDF 位元組（與寫檔分開，寫檔失敗時才能改存替代檔而不用重印）
    try:
        pdf_settings = {"landscape": False, "printBackground": True}
        pdf_data = driver.execute_cdp_cmd("Page.printToPDF", pdf_settings)
        time.sleep(1)
        data = base64.b64decode(pdf_data["data"])
    except Exception as e:
        print(f"\033[91m產生 PDF 時發生錯誤：{e}\033[0m", flush=True)
        return

    try:
        with open(filename, "wb") as f:
            f.write(data)
        print(f"已成功保存 PDF：\n\033[93m{filename}\033[0m", flush=True)
    except PermissionError:
        # 🔥 原檔很可能正被開啟(PDF 檢視器鎖住)導致無法覆寫 → 改存不衝突的檔名，確保不漏檔
        base, ext = os.path.splitext(filename)
        alt = f"{base}_{time.strftime('%Y%m%d_%H%M%S')}{ext}"
        try:
            with open(alt, "wb") as f:
                f.write(data)
            print(f"\033[93m原檔可能正被開啟(鎖住)無法覆寫，已改存為：\n{alt}\033[0m", flush=True)
            print("\033[93m提示：若要更新原檔，請先關閉正在檢視的同名 PDF 再重查。\033[0m", flush=True)
        except Exception as e2:
            print(f"\033[91m保存 PDF 失敗（原檔被鎖、替代檔也寫入失敗）：{e2}\033[0m", flush=True)
            print("\033[93m請先關閉正在開啟的同名 PDF，再重新查詢。\033[0m", flush=True)
    except Exception as e:
        print(f"\033[91m保存 PDF 時發生錯誤：{e}\033[0m", flush=True)


def select_address_query_type():
    """選擇 '建築地址' 的查詢類型。
    🔥 新版高雄站「建築地址」是 QType3，且 radio 被 CSS 隱藏，需用 presence + JS click。
    """
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "QType3"))
        )
        radio = driver.find_element(By.ID, "QType3")
        driver.execute_script("arguments[0].click();", radio)
        print("已選擇查詢類型：建築地址（#QType3）", flush=True)

        # 等 searchbbox03 切換到 display:block
        try:
            WebDriverWait(driver, 5).until(
                lambda d: 'block' in (
                    d.find_element(By.CSS_SELECTOR, "div.searchbbox03").get_attribute("style") or ""
                )
            )
        except TimeoutException:
            print("[警告] 等待建築地址表單顯示超時，仍繼續操作", flush=True)
    except TimeoutException:
        print("無法選擇查詢類型：建築地址（找不到 #QType3）", flush=True)
    except Exception as e:
        print(f"選擇建築地址查詢類型時發生錯誤：{e}", flush=True)


def perform_address_search(city, area, address=None):
    """進行建築地址查詢，根據地址解析填入相關欄位，並執行驗證"""
    # 確保選中「建築地址」查詢類型
    try:
        select_address_query_type()
    except Exception as e:
        print(f"選擇建築地址查詢類型時發生錯誤：{e}", flush=True)
        return "no_data"

    # 自動生成 base_address
    base_address = f"{city}{area}" if city and area else ""

    # 提示輸入建築地址，允許修改
    if not address:
        print(f"請輸入接續之建築地址：\033[93m{base_address}\033[0m", flush=True)
        user_input = input()
        if user_input.strip():
            # 如果使用者輸入不包含 base_address，則補上
            if not user_input.startswith(base_address):
                address = base_address + user_input
            else:
                address = user_input
        else:
            address = base_address  # 使用者直接按 Enter 時，採用 base_address

    print(f"輸入的建築地址為：{address}", flush=True)
    
    # 驗證地址格式
    if not city or not area:
        print("警告：地址缺少縣市或區域信息", flush=True)
        retry = input("是否要放棄此次查詢？(Y/N)：").strip().upper()
        if retry == 'Y':
            return "no_data"
            
    address_data = parse_address(address)  # 解析地址
    
    try:
        # 使用select_city_area函數選擇縣市區
        success = select_city_area(address_data["city_area"])
        if not success:
            print("無法選擇縣市區，是否要放棄此次查詢？", flush=True)
            retry = input("放棄請輸入Y，繼續請輸入N：").strip().upper()
            if retry == 'Y':
                return "no_data"
                
        # 填入地址欄位
        fill_address_fields(address_data)
        
        # 嘗試進行查詢
        result = attempt_verification()
        return result
        
    except Exception as e:
        print(f"查詢過程中發生錯誤：{e}", flush=True)
        retry = input("是否要重試？(Y/N)：").strip().upper()
        if retry == 'Y':
            # 重新開始地址查詢
            print("重新開始地址查詢...", flush=True)
            return perform_address_search(city, area, address)
        else:
            return "no_data"

def parse_address(address):
    """
    解析地址為字典格式，包含縣市區、路街段、巷、弄、號等多種組合。
    """
    address_data = {
        "city_area": None,
        "road": None,
        "segment": None,
        "lane": None,
        "alley": None,
        "number": None
    }
    
    # 提取縣市區（市、區）
    if "市" in address and "區" in address:
        city_area_split = address.split("市", 1)
        address_data["city_area"] = f"{city_area_split[0]}市{city_area_split[1].split('區', 1)[0]}區"
        remaining_address = city_area_split[1].split("區", 1)[1]
    else:
        remaining_address = address

    # 依次解析地址結構
    remaining_address = parse_road_and_segment(remaining_address, address_data)
    remaining_address = parse_lane(remaining_address, address_data)
    remaining_address = parse_alley(remaining_address, address_data)
    parse_number(remaining_address, address_data)

    return address_data

def select_city_area(city_area):
    """選擇 '縣市區' 下拉選單的項目，使用多種匹配策略"""
    try:
        city_area_select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "addradr"))
        )
        city_area_select = Select(city_area_select_element)
        
        # 確保下拉選單已載入選項
        WebDriverWait(driver, 5).until(
            lambda d: len(city_area_select.options) > 1
        )
        
        options = [opt.text for opt in city_area_select.options]
        
        # 策略1: 直接精確匹配
        if city_area in options:
            city_area_select.select_by_visible_text(city_area)
            print(f"已選擇縣市區（精確匹配）：{city_area}", flush=True)
            return True
        
        # 策略2: 在選項中查找包含city_area的項
        for opt in options:
            if opt != "-- 請選擇 --" and city_area in opt:
                city_area_select.select_by_visible_text(opt)
                print(f"已選擇縣市區（包含匹配）：{opt}", flush=True)
                return True
                
        # 策略3: 提取區名，然後在選項中查找
        if "區" in city_area:
            district = city_area.split("區")[0] + "區"
            for opt in options:
                if opt != "-- 請選擇 --" and district in opt:
                    city_area_select.select_by_visible_text(opt)
                    print(f"已選擇縣市區（區名匹配）：{opt}", flush=True)
                    return True
        
        # 策略4: 基於關鍵詞的匹配
        if "市" in city_area and "區" in city_area:
            keywords = city_area.split("市")[0].strip()  # 例如 "高雄"
            district = city_area.split("區")[0].split("市")[-1].strip()  # 例如 "鳳山"
            
            best_match = None
            best_count = 0
            
            for opt in options:
                if opt == "-- 請選擇 --":
                    continue
                
                match_count = 0
                if keywords in opt:
                    match_count += 1
                if district in opt:
                    match_count += 1
                
                if match_count > best_count:
                    best_count = match_count
                    best_match = opt
            
            if best_match:
                city_area_select.select_by_visible_text(best_match)
                print(f"已選擇縣市區（關鍵詞匹配）：{best_match}", flush=True)
                return True
        
        # 策略5: 手動選擇
        print(f"無法自動匹配縣市區：{city_area}", flush=True)
        print("可用的縣市區選項有：", flush=True)
        for i, opt in enumerate(options):
            if opt != "-- 請選擇 --":
                print(f"[{i}] {opt}", flush=True)
        
        user_choice = input("請輸入選項編號（或輸入ESC跳過）：").strip()
        if user_choice.upper() == "ESC":
            return False
        
        try:
            choice_index = int(user_choice)
            if 0 <= choice_index < len(options):
                selected_option = options[choice_index]
                city_area_select.select_by_visible_text(selected_option)
                print(f"已選擇縣市區（手動選擇）：{selected_option}", flush=True)
                return True
            else:
                print("無效的選項編號", flush=True)
                return False
        except ValueError:
            print("請輸入有效的數字", flush=True)
            return False
    except TimeoutException:
        print("無法找到縣市區下拉選單", flush=True)
        return False
    except Exception as e:
        print(f"選擇縣市區時發生錯誤：{e}", flush=True)
        return False

# 修改parse_road_and_segment函數以處理"路x段"格式
def parse_road_and_segment(remaining_address, address_data):
    """
    處理路、段、街等解析，支持多種名稱組合。
    """
    road_keywords = ["路", "街"]
    segment_keyword = "段"
    
    # 特殊處理：匹配"x路y段"的模式（例如"建國路三段"）
    for keyword in road_keywords:
        if keyword in remaining_address and segment_keyword in remaining_address:
            road_index = remaining_address.find(keyword)
            segment_index = remaining_address.find(segment_keyword)
            
            # 確認"段"在"路"之後
            if road_index < segment_index:
                # 提取完整的"x路y段"（包括數字）
                address_data["road"] = remaining_address[:segment_index + 1]
                remaining_address = remaining_address[segment_index + 1:]
                address_data["segment"] = ""  # 清空segment，因已包含在road中
                return remaining_address
    
    # 標準處理：分別處理路名和段名
    for keyword in road_keywords:
        if keyword in remaining_address:
            road_parts = remaining_address.split(keyword, 1)
            address_data["road"] = road_parts[0] + keyword
            remaining_address = road_parts[1]
            break

    # 檢查是否有段
    if segment_keyword in remaining_address:
        segment_parts = remaining_address.split(segment_keyword, 1)
        address_data["segment"] = segment_parts[0] + segment_keyword
        remaining_address = segment_parts[1]

    return remaining_address

def parse_lane(remaining_address, address_data):
    """處理巷的解析。"""
    if "巷" in remaining_address:
        lane_split = remaining_address.split("巷", 1)
        # 🔥 只保留數字，不加「巷」字（網頁欄位只需要數字）
        address_data["lane"] = lane_split[0]
        remaining_address = lane_split[1]

    return remaining_address


def parse_alley(remaining_address, address_data):
    """處理弄的解析。"""
    if "弄" in remaining_address:
        alley_split = remaining_address.split("弄", 1)
        # 🔥 只保留數字，不加「弄」字（網頁欄位只需要數字）
        address_data["alley"] = alley_split[0]
        remaining_address = alley_split[1]

    return remaining_address


def parse_number(remaining_address, address_data):
    """處理號的解析，確保只保留數字部分。"""
    if "號" in remaining_address:
        number = remaining_address.split("號", 1)[0].strip()
        address_data["number"] = number
    else:
        address_data["number"] = remaining_address.strip()

def fill_address_fields(address_data):
    """依照解析結果填入地址欄位，支援各種組合"""
    try:
        # 1. 填入路名
        if address_data["road"]:
            road_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "addrad2"))
            )
            road_input.clear()
            road_input.send_keys(address_data["road"])
            print(f"已填入路名：{address_data['road']}", flush=True)

        # 2. 填入巷（或段，如果有）
        lane_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "addrad3"))
        )
        lane_input.clear()
        
        if address_data["segment"] and address_data["segment"].strip():
            # 如果有段，填入段名
            lane_input.send_keys(address_data["segment"])
            print(f"已填入段名：{address_data['segment']}", flush=True)
        elif address_data["lane"]:
            # 否則填入巷名
            lane_input.send_keys(address_data["lane"])
            print(f"已填入巷名：{address_data['lane']}", flush=True)

        # 3. 填入弄（如果有）
        if address_data["alley"]:
            alley_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "addrad4"))
            )
            alley_input.clear()
            alley_input.send_keys(address_data["alley"])
            print(f"已填入弄名：{address_data['alley']}", flush=True)

        # 4. 填入號碼
        if address_data["number"]:
            number_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "addrad5"))
            )
            number_input.clear()
            number_input.send_keys(address_data["number"])
            print(f"已填入號碼：{address_data['number']}", flush=True)

        print("已填入所有地址欄位", flush=True)

    except TimeoutException:
        print("無法找到地址的欄位", flush=True)
    except Exception as e:
        print(f"填入地址欄位時發生錯誤：{e}", flush=True)


# 修改generate_option函数，使其返回数字而非字母
def generate_option(index):
    """
    根据索引生成选项编号。
    索引从1开始，直接返回数字序号字符串。
    """
    return str(index)

def _expand_sel(raw):
    """把選擇輸入展開：支援逗號分隔與 2-3 範圍（可混用 1,3-5），保留 a/all 等非數字原樣。
    回傳 list；空輸入回 ['']（讓後續判為無效，維持原行為）。"""
    out = []
    for part in str(raw).split(","):
        part = part.strip()
        if not part:
            continue
        bits = part.split("-", 1)
        if "-" in part and len(bits) == 2 and bits[0].strip().isdigit() and bits[1].strip().isdigit():
            a, b = int(bits[0].strip()), int(bits[1].strip())
            if a > b:
                a, b = b, a
            out.extend(str(n) for n in range(a, b + 1))
        else:
            out.append(part)
    seen = set()
    res = []
    for o in out:
        if o not in seen:
            seen.add(o)
            res.append(o)
    return res or ['']

def _nlma_zoom_count():
    """依 window_config 的 DPI + zoom_config 覆寫，算出 nlma 要縮幾次（給跳轉後的新分頁用）。"""
    zc = 0
    try:
        _cfg = os.path.join(BASE_DIR, 'window_config.json')
        if os.path.exists(_cfg):
            c = json.load(open(_cfg, encoding='utf-8'))
            dpi = c.get('dpi_scale', 1.0) or 1.0
            sw = c.get('screen_width', 1920)
            lw = sw / dpi
            if dpi >= 1.75:
                zc = 4
            elif dpi >= 1.5:
                zc = 2
            elif dpi >= 1.25 and lw < 1200:
                zc = 2
            else:
                zc = 0
    except Exception:
        pass
    try:
        _zp = os.path.join(BASE_DIR, 'zoom_config.json')
        if os.path.exists(_zp):
            ov = json.load(open(_zp, encoding='utf-8')).get('overrides', {}).get('nlma')
            if ov is not None:
                zc = int(ov)
    except Exception:
        pass
    return zc

def click_and_save_pdf(result, base_directory):
    opened_new = False
    origin_window = None
    try:
        link_element = result.get("link_element")
        link_mode = result.get("link_mode")
        link_href = result.get("link_href")
        license_number = result.get("license_number") or ""

        # 沒有可用的連結元素/模式（例如該列真的無連結）→ 乾淨略過，不去 driver.get("無") 爆 invalid argument
        if link_element is None or link_mode not in ("get", "click"):
            print(f"此筆沒有可下載的執照連結（{license_number or '無編號'}），略過下載。", flush=True)
            return

        origin_window = driver.current_window_handle
        before = list(driver.window_handles)

        if link_mode == "get":
            # (A) 舊版/他縣市：<a href> 直接 GET 跳轉（同分頁）
            print(f"準備跳轉到鏈接：{link_href}", flush=True)
            driver.get(link_href)
        else:
            # (B) 新版高雄 buildmis：點擊 POST 表單送出按鈕，瀏覽器自帶 CSRF 送出，target=_blank 開新分頁
            print(f"準備點擊送出以開啟執照明細（POST）：{link_href}", flush=True)
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link_element)
            except Exception:
                pass
            try:
                link_element.click()
            except Exception:
                driver.execute_script("arguments[0].click();", link_element)
            # 等新分頁出現並切過去（若沒開新分頁，視為同分頁導航）
            try:
                WebDriverWait(driver, 12).until(lambda d: len(d.window_handles) > len(before))
                new_window = [h for h in driver.window_handles if h not in before][0]
                driver.switch_to.window(new_window)
                opened_new = True
                print("已切換到執照明細分頁。", flush=True)
            except TimeoutException:
                print("未偵測到新分頁，改於當前分頁擷取。", flush=True)

        # 先確保 DOM 框架載入完成（取代原本猜測的 target-table 類名）
        try:
            WebDriverWait(driver, 10).until(
                lambda d: d.execute_script("return document.readyState") == "complete")
        except Exception:
            pass

        # 🔥 明細頁是 SPA：框架載入後，表格內容才由 AJAX 抓回來填入。
        #    只等 readyState 就列印，會得到「有標題、值全空」的 PDF。
        #    改以「執照號碼是否出現在頁面文字」判斷資料已填入(明細頁最上方會顯示使用執照號碼)。
        _full = re.sub(r'\s+', '', license_number or '')
        _m = re.search(r'(\d{3,}(?:-\d+)?)號?', license_number or '')
        _core = _m.group(1) if _m else ''   # 例：00865-01，靜態模板不會有這串，出現即代表資料已載入

        def _detail_data_ready(d):
            try:
                body = re.sub(r'\s+', '', d.execute_script(
                    "return document.body ? document.body.innerText : ''") or "")
            except Exception:
                return False
            if _core and _core in body:
                return True
            return bool(_full) and _full in body

        if _core or _full:
            try:
                WebDriverWait(driver, 20).until(_detail_data_ready)
                print("執照明細資料已載入。", flush=True)
            except TimeoutException:
                print("等待執照明細資料逾時，仍嘗試擷取當前內容。", flush=True)
        # 資料出現後再給短暫緩衝，讓表格/樣式渲染穩定再列印
        time.sleep(1.5)

        # 套用縮放（與其他頁面一致）
        try:
            apply_page_zoom(driver, _nlma_zoom_count())
        except Exception:
            pass

        # 建立目標目錄
        pdf_directory = os.path.join(base_directory, "1.基本資料")
        os.makedirs(pdf_directory, exist_ok=True)

        # 設置檔名
        sanitized_name = (license_number or "執照").replace("/", "_").replace("\\", "_").strip()
        filename = f"02_使用執照-{sanitized_name}.pdf"
        filepath = os.path.join(pdf_directory, filename)

        # 保存頁面為 PDF
        save_page_as_pdf(filepath)

    except Exception as e:
        print(f"處理選項時發生錯誤：{e}", flush=True)
    finally:
        # 收尾：新分頁關掉並切回原結果視窗；同分頁 GET 則 back 回結果頁（讓後續多筆仍可續處理）
        try:
            if opened_new:
                driver.close()
                if origin_window:
                    driver.switch_to.window(origin_window)
            elif result.get("link_mode") == "get":
                driver.back()
        except Exception:
            pass

# 注册退出回调函数，确保程序无论如何退出都会发送通知
def notify_main_program():
    print("子程式已完成執行", flush=True)
    sys.stdout.flush()  # 確保消息被立即發送
# 注册退出函数，确保程序退出时一定会发送通知
atexit.register(notify_main_program)

if __name__ == '__main__':
    try:
        main_program()  # 使用新的主程式循環功能
    except KeyboardInterrupt:
        try:
            print("\n程式被使用者中斷", flush=True)
        except (EOFError, OSError):
            pass
    except SystemExit:
        try:
            print("\n程式正常退出", flush=True)
        except (EOFError, OSError):
            pass
    except (EOFError, OSError):
        # 標準輸入/輸出已關閉，靜默退出
        pass
    except Exception as e:
        try:
            print(f"程式執行過程中發生錯誤: {e}", flush=True)
        except (EOFError, OSError):
            pass
    finally:
        # 確保關閉 driver（如果還存在的話）
        try:
            if 'driver' in globals() and driver:
                driver.quit()
        except:
            pass
