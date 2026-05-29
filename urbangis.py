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
os.system('cls')
print("歡迎使用【高雄市都市計畫】自動化小程式，模組載入中...", flush=True)
import json
import time
import re
import ctypes
import keyboard  # 用於系統級鍵盤事件
from PIL import Image
from fpdf import FPDF
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException
import sys
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window


# 🔥 自訂例外：用於從 step 3/4 早跳出（道路用地不需要點擊+擷取建蔽容積）
class _SkipUrbanPlanQuery(Exception):
    """道路用地時，跳過點擊紅框與都市計畫圖資擷取，直接到截圖步驟"""
    pass

# 🔥 基準目錄設定（data.json 和工作資料夾的位置）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

# 🔥 自動偵測 DPI 縮放比例
def get_dpi_scale():
    """偵測系統 DPI 縮放比例"""
    try:
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, hdc)
        dpi_scale = dpi / 96.0  # 96 DPI = 100%
        return dpi_scale
    except Exception as e:
        print(f"[WARNING] 無法偵測 DPI，使用預設值 1.0: {e}")
        return 1.0

# 🔥 視窗大小設定：從 main.py 生成的設定檔讀取，自動填滿剩餘空間
urbangis_window_width = 1024
urbangis_window_height = 1024
urbangis_window_x = 0
urbangis_window_y = 0

try:
    config_path = os.path.join(BASE_DIR, 'window_config.json')
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            window_config = json.load(f)

        # 取得工作區域和 main 視窗資訊
        work_area = window_config.get('work_area', {})
        main_window = window_config.get('main_window', {})

        if work_area and main_window:
            # 🔥 取得 DPI 縮放比例（優先使用配置檔，否則自動偵測）
            dpi_scale = window_config.get('dpi_scale')
            if dpi_scale is None:
                dpi_scale = get_dpi_scale()

            # 取得主程式視窗資訊
            main_x = main_window.get('x', 1280)
            main_y = main_window.get('y', 0)
            main_w = main_window.get('width', 640)
            main_h = main_window.get('height', 1020)

            # 根據主程式位置，智能計算 Chrome 視窗的位置和大小
            work_left = work_area.get('left', 0)
            work_top = work_area.get('top', 0)
            work_right = work_area.get('right', 1920)
            work_bottom = work_area.get('bottom', 1080)

            # 計算 Chrome 視窗大小：固定佔螢幕 2/3 寬度
            screen_width_physical = work_right - work_left
            chrome_width = (screen_width_physical * 2) // 3

            # 判斷主程式在螢幕的哪一側，決定 Chrome 位置
            main_window_center_x = main_x + main_w / 2
            screen_center_x = (work_left + work_right) / 2

            if main_window_center_x > screen_center_x:
                # 主程式在右側，Chrome 放左側
                chrome_x = work_left
            else:
                # 主程式在左側，Chrome 放右側
                chrome_x = work_right - chrome_width

            # 高度配合主程式高度（與主程式同高）
            chrome_y = main_y
            chrome_height = main_h

            # 確保視窗大小合理
            if chrome_width < 800:
                chrome_width = 800
            if chrome_height < 600:
                chrome_height = 600

            # 轉換為 Chrome 的邏輯像素（Selenium 使用邏輯像素）
            chrome_width_logical = int(chrome_width / dpi_scale)
            chrome_x_logical = int(chrome_x / dpi_scale)
            chrome_y_logical = int(chrome_y / dpi_scale)

            # 🔥 Chrome 視窗高度需要加上標題欄高度（約 32px 邏輯像素）
            # 這樣整個視窗（包含標題欄）才會和主程式一樣高
            CHROME_TITLEBAR_HEIGHT = 32
            chrome_height_logical = int(chrome_height / dpi_scale) + CHROME_TITLEBAR_HEIGHT

            urbangis_window_width = chrome_width_logical
            urbangis_window_height = chrome_height_logical
            urbangis_window_x = chrome_x_logical
            urbangis_window_y = chrome_y_logical
        else:
            # 舊版設定檔
            pass
    else:
        # 設定檔不存在
        pass
except Exception as e:
    print(f"[WARNING] 讀取視窗設定失敗: {e}，使用預設值")

# 讀取 JSON 數據
with open(get_data_json_path(), 'r', encoding='utf-8') as f:
    data_list = json.load(f)

if not data_list:
    print("data.json 是空的，請檢查資料。", flush=True)
    exit()

# 取得第一組資料，建立基礎目錄結構
first_entry = data_list[0]

# 🔥 檢查是否為高雄市，只有高雄市才能執行高雄都市計畫
city = first_entry.get('city', '')
if city != '高雄市':
    print(f"\033[96m目前查詢是「{city}」，非【高雄市】！\033[0m", flush=True)
    print(f"\033[96m高雄都市計畫系統僅適用於高雄市。\033[0m", flush=True)
    exit(1)

base_region = first_entry['area']
base_section = first_entry['section']
base_lot_number = first_entry['lot_number']

# 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
base_directory = get_work_folder(f"{base_region}{base_section}-{base_lot_number}")
basic_data_dir = os.path.join(base_directory, "1.基本資料")
other_data_dir = os.path.join(base_directory, "4.其他相關")
png_dir = os.path.join(basic_data_dir, "png")

# 建立目錄結構
os.makedirs(png_dir, exist_ok=True)
os.makedirs(other_data_dir, exist_ok=True)
print(f"已建立目錄 {base_directory}", flush=True)

# 初始化 WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
# 🔥 不在啟動參數設定視窗大小，改為啟動後再調整（避免 DPI 改變時崩潰）
# options.add_argument(f"--window-size={window_width},{window_height}")
# options.add_argument("--window-position=0,0")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
# 🔥 允許地理位置（避免網站在被拒絕時呼叫 alert 卡住 Selenium）；通知拒絕
options.add_experimental_option("prefs", {
    "profile.default_content_setting_values.geolocation": 1,
    "profile.default_content_setting_values.notifications": 2,
})
driver = create_chrome_driver(options=options)

# 🔥 啟動後立即設定視窗大小和位置（避免在啟動參數設定導致 DPI 改變時崩潰）
try:
    driver.set_window_size(urbangis_window_width, urbangis_window_height)
    driver.set_window_position(urbangis_window_x, urbangis_window_y)
except Exception as e:
    print(f"[WARNING] 視窗設定失敗: {e}，繼續執行")

url = "https://urbangis.kcg.gov.tw/"


# 函式：初始化頁面
def initialize_page(driver, url):
    driver.get(url)
    # 🔥 網頁載入後重新設定視窗大小（避免被網站重置）
    try:
        driver.set_window_size(urbangis_window_width, urbangis_window_height)
        driver.set_window_position(urbangis_window_x, urbangis_window_y)
    except:
        pass
    # 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
    verify_and_fix_chrome_window(driver)
    driver.implicitly_wait(10)

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

            if zoom_count > 0:
                print(f"[頁面縮放] 螢幕={screen_width}px, DPI={int(current_dpi * 100)}%, 邏輯寬度={logical_width:.0f}px → 縮放至 {target_zoom}", flush=True)
            else:
                print(f"[頁面縮放] 螢幕={screen_width}px, DPI={int(current_dpi * 100)}%, 邏輯寬度={logical_width:.0f}px → 不需要縮放", flush=True)
    except Exception as e:
        print(f"[頁面縮放] 偵測配置時發生錯誤: {e}，使用預設不縮放", flush=True)

    if zoom_count > 0:
        try:
            # 🔥 使用 keyboard 庫發送系統級的 Ctrl + - 按鍵（真正的瀏覽器縮放）
            # 先點擊頁面確保 Chrome 視窗有焦點
            body = driver.find_element(By.TAG_NAME, 'body')
            body.click()
            time.sleep(0.5)

            # 🔥 根據 DPI 決定按幾次 Ctrl + -
            print(f"[頁面縮放] 正在使用系統鍵盤縮放（按 {zoom_count} 次）...", flush=True)
            for i in range(zoom_count):
                keyboard.press_and_release('ctrl+-')  # 系統級按鍵
                time.sleep(0.5)

            print(f"[頁面縮放] ✓ 已使用系統鍵盤縮放至 {target_zoom}", flush=True)
        except Exception as e:
            print(f"[頁面縮放] 設定縮放時發生錯誤: {e}", flush=True)
    else:
        print("[頁面縮放] 螢幕配置正常，不需要縮放", flush=True)

    # 關閉所有「關閉」按鈕的彈出視窗
    try:
        # 等待頁面載入
        time.sleep(2)

        # 使用 JavaScript 點擊所有「關閉」按鈕
        driver.execute_script("""
            var closeButtons = document.querySelectorAll('button');
            for (var i = 0; i < closeButtons.length; i++) {
                if (closeButtons[i].textContent.trim() === '關閉') {
                    closeButtons[i].click();
                    console.log('點擊了關閉按鈕');
                }
            }
        """)
        # print("已點擊所有關閉按鈕", flush=True)
        time.sleep(1)  # 等待關閉動畫完成
    except Exception as e:
        print(f"關閉視窗時發生錯誤: {e}", flush=True)

    # 關閉鷹眼圖
    try:
        close_eye_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "collapse_overviewMapBtn"))
        )
        driver.execute_script("arguments[0].click();", close_eye_button)
        # print("關閉了鷹眼圖", flush=True)
        time.sleep(0.5)  # 等待關閉動畫完成
    except TimeoutException:
        print("鷹眼圖沒有出現或無法點擊", flush=True)

import re
import csv

# 函式：擷取分區查詢彈出視窗的資料（不含建蔽率/容積率）
def extract_zone_data(driver, area='', section=''):
    """
    從分區查詢彈出視窗中擷取土地資訊
    🔥 注意：分區查詢彈窗不包含建蔽率/容積率，這兩項只在都市計畫圖街廓資訊中才有
    🔥 修改：返回資料字典，不再直接存檔

    參數:
        driver: Selenium WebDriver
        area: 行政區（例如：'大寮區'）
        section: 地段（例如：'仁德段'）
    """
    try:
        # 等待分區查詢彈出視窗出現
        time.sleep(2)

        # 檢查彈出視窗是否存在
        try:
            popup_elements = driver.find_elements(By.CLASS_NAME, "esri-popup__main-container")

            if popup_elements:
                for idx, elem in enumerate(popup_elements):
                    pass
        except Exception as e:
            pass

        # 使用 JavaScript 擷取整個頁面的文字內容
        page_text = driver.execute_script("return document.body.innerText;")

        # 解析資料
        data_dict = {
            '地段': '',
            '地號': '',
            '使用分區': '',
            '土地面積': '',
            '公告現值': '',
            '公告地價': '',
            '附註': ''
        }

        # 使用正則表達式從整個頁面文字中提取資料
        lines = page_text.split('\n')

        for line in lines:
            # 移除所有空白字元（包括全形空格）以便比對
            line_no_space = re.sub(r'[\s　]+', '', line)  # 移除所有半形和全形空格

            # 檢查是否包含冒號（半形或全形）
            has_colon = ':' in line or '：' in line

            # 尋找「地段」
            if '地' in line_no_space and '段' in line_no_space and has_colon:
                # 使用冒號分割（支援全形和半形）
                value = re.split('[：:]', line, 1)[-1].strip()
                if value and '地號' not in value and '使用' not in value:
                    data_dict['地段'] = value
                    print(f"   找到地段: {value}", flush=True)

            # 尋找「地號」
            elif '地' in line_no_space and '號' in line_no_space and has_colon:
                value = re.split('[：:]', line, 1)[-1].strip()
                if value and '地段' not in value and '使用' not in value:
                    data_dict['地號'] = value
                    print(f"   找到地號: {value}", flush=True)

            # 尋找「使用分區」
            elif '使用' in line_no_space and '分區' in line_no_space and has_colon:
                value = re.split('[：:]', line, 1)[-1].strip()
                if value:
                    data_dict['使用分區'] = value
                    print(f"   找到使用分區: {value}", flush=True)

            # 尋找「土地面積」
            elif '土地' in line_no_space and '面積' in line_no_space and has_colon:
                value = re.split('[：:]', line, 1)[-1].strip()
                if value:
                    data_dict['土地面積'] = value
                    print(f"   找到土地面積: {value}", flush=True)

            # 尋找「公告現值」
            elif '公告' in line_no_space and '現值' in line_no_space and has_colon:
                value = re.split('[：:]', line, 1)[-1].strip()
                if value:
                    data_dict['公告現值'] = value
                    print(f"   找到公告現值: {value}", flush=True)

            # 尋找「公告地價」
            elif '公告' in line_no_space and '地價' in line_no_space and has_colon:
                value = re.split('[：:]', line, 1)[-1].strip()
                if value:
                    data_dict['公告地價'] = value
                    print(f"   找到公告地價: {value}", flush=True)

            # 尋找「附註」
            elif '附' in line_no_space and '註' in line_no_space and has_colon:
                value = re.split('[：:]', line, 1)[-1].strip()
                if value:
                    data_dict['附註'] = value
                    print(f"   找到附註: {value}", flush=True)

        # 🔥 如果網頁中沒有抓到地段，或地段不完整，使用傳入的參數補全
        if area and section:
            full_section = f"{area}{section}"
            # 如果抓到的地段只有段名，沒有行政區，則補上行政區
            if data_dict['地段'] and data_dict['地段'] == section:
                data_dict['地段'] = full_section
                print(f"   [修正] 補全地段為: {full_section}", flush=True)
            # 如果完全沒抓到地段，使用完整地段
            elif not data_dict['地段']:
                data_dict['地段'] = full_section
                print(f"   [修正] 設定地段為: {full_section}", flush=True)

        # 🔥 只在資料完整時才顯示（避免顯示空資料）
        print(f"   地段: {'✓ ' + data_dict['地段'] if data_dict['地段'] else '✗ 未取得'}", flush=True)
        print(f"   地號: {'✓ ' + data_dict['地號'] if data_dict['地號'] else '✗ 未取得'}", flush=True)
        print(f"   使用分區: {'✓ ' + data_dict['使用分區'] if data_dict['使用分區'] else '✗ 未取得'}", flush=True)
        print(f"   土地面積: {'✓ ' + data_dict['土地面積'] if data_dict['土地面積'] else '✗ 未取得'}", flush=True)
        print(f"   公告現值: {'✓ ' + data_dict['公告現值'] if data_dict['公告現值'] else '✗ 未取得'}", flush=True)
        print(f"   公告地價: {'✓ ' + data_dict['公告地價'] if data_dict['公告地價'] else '✗ 未取得'}", flush=True)

        if data_dict['地號'] and data_dict['使用分區']:
            print(f"✓ 已擷取分區查詢資料:\n   地段={data_dict['地段']}, 地號={data_dict['地號']}, 使用分區={data_dict['使用分區']}", flush=True)
        elif data_dict['地號']:
            print(f"⚠️ 部分擷取分區查詢資料:\n   地段={data_dict['地段']}, 地號={data_dict['地號']}, 使用分區=（未取得）", flush=True)
        else:
            pass

        return data_dict

    except Exception as e:
        print(f"[ERROR] 擷取分區查詢資料時發生錯誤: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return None

# 更新 select_section_with_retry 函數，當無法自動選擇到地段時提示手動選擇
def select_section_with_retry(section_select, section_name, max_retries=3):
    # 清理段名中的括弧和數字，例如「內惟段四小段(0123)」變成「內惟段四小段」
    clean_section_name = re.sub(r"\(.*?\)", "", section_name).strip()
    
    for attempt in range(max_retries):
        try:
            options = section_select.options
            # 嘗試精確匹配已清理的段名
            for option in options:
                option_text_cleaned = re.sub(r"\(.*?\)", "", option.text).strip()
                if option_text_cleaned == clean_section_name:
                    section_select.select_by_visible_text(option.text)
                    print(f"成功選擇地段: {option.text}", flush=True)
                    return True
            
            # 若未找到精確匹配,列印所有選項的 Unicode 碼點供診斷
            print(f"未找到精確匹配的地段: {clean_section_name}", flush=True)
            print(f"目標地段 Unicode: {[f'{c}(U+{ord(c):04X})' for c in clean_section_name]}", flush=True)
            print("網頁中可用的地段選項:", flush=True)
            for i, option in enumerate(options, 1):
                option_text_cleaned = re.sub(r"\(.*?\)", "", option.text).strip()
                unicode_repr = [f'{c}(U+{ord(c):04X})' for c in option_text_cleaned]
                print(f"  {i}. {option_text_cleaned} - Unicode: {unicode_repr}", flush=True)

        except StaleElementReferenceException:
            # print(f"嘗試重新取得選項 (第 {attempt + 1} 次)", flush=True)
            time.sleep(1)
    
    # 如果無法找到精確匹配的地段，提示使用者手動選擇
    print(f"多次嘗試後仍無法自動選擇地段: {clean_section_name}", flush=True)
    print("請手動選擇正確的地段，然後按 Enter 繼續...", flush=True)
    input()
    return True


# 函式：處理查詢後的操作，保存 PNG
def finalize_query(driver, region, section, lot_number, png_dir, png_files, identifiers):
    try:
        # 確保圖形<path>元素加載完成
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "path"))
        )
        print("圖形<path>已出現，地圖已完成載入", flush=True)
    except TimeoutException:
        print("等待圖形<path>元素時超時，無法確認地圖是否載入完成", flush=True)
    


    # 關閉定位查詢和查詢結果視窗
    try:
        driver.execute_script("""
            var buttons = document.querySelectorAll('button[title="關閉"]');
            for (var i = 0; i < buttons.length; i++) {
                buttons[i].click();
            }
        """)
        print("已關閉定位查詢和查詢結果視窗", flush=True)
        time.sleep(1)
    except Exception as e:
        print(f"關閉視窗時發生錯誤: {e}", flush=True)

    # **新增：等待使用者調整網頁 (倒數計時 - 使用 print)**
    print("\033[93m請在接下來的 10 秒內，\n手動調整網頁縮放、移動地圖，以達到最佳截圖效果\033[0m", flush=True)
    countdown(10)
    
    # 定義 PNG 檔名
    png_filename = f"07_都市計畫使用分區-{region}{section}-{lot_number}.png"
    screenshot_path = os.path.join(png_dir, png_filename)
    
    # 保存截圖
    try:
        driver.save_screenshot(screenshot_path)
        print(f"\033[93m網頁截圖已保存為 {screenshot_path}\033[0m", flush=True)
        png_files.append(screenshot_path)  # 收集 PNG 路徑
        identifiers.append(f"{region}{section}-{lot_number}")  # 收集標識符
    except Exception as e:
        print(f"截圖時發生錯誤: {e}", flush=True)

# 函式：合併所有 PNG 成為一個 PDF
def create_combined_pdf(png_files, identifiers, pdf_dir):
    if not png_files:
        print("沒有找到任何 PNG 檔案，無法生成 PDF。", flush=True)
        return

    # 使用新命名規則合併標識符
    area_section_dict = {}
    for identifier in identifiers:
        # 僅分割前兩個值，其他部分當作 lot_number
        parts = identifier.split('-')
        region = parts[0]
        lot = '-'.join(parts[1:])  # 將後面的部分重新合併成地號
        
        if region in area_section_dict:
            area_section_dict[region].append(lot)
        else:
            area_section_dict[region] = [lot]

    # 建立檔名，將每個區域的地號以「、」分隔，並用「+」連接不同地區地段
    pdf_filename = "07_都市計畫使用分區-" + "+".join(
        [f"{region}-{','.join(lots)}" for region, lots in area_section_dict.items()]
    ) + ".pdf"

    pdf_path = os.path.join(pdf_dir, pdf_filename)

    try:
        pdf = FPDF(orientation='L', unit='mm', format='A4')  # 橫向 PDF
        for png in png_files:
            pdf.add_page()
            pdf.image(png, x=0, y=0, w=297, h=210)  # 調整大小至 A4 橫向
        pdf.output(pdf_path)
        print(f"\033[93mPDF 已經儲存至 {pdf_path}\033[0m", flush=True)
    except Exception as e:
        print(f"將 PNG 轉成 PDF 時發生錯誤: {e}", flush=True)

def countdown(duration):
    for i in range(duration, 0, -1):
        # 使用 print 將倒數數字輸出到 stdout
        print(f"倒數：{i} 秒", flush=True)  # 將數字轉換為字串
        time.sleep(1)
    print("倒數結束", flush=True)

# 函式:都市計畫圖資詳細查詢流程
def query_urban_planning_details(driver, area, section, lot_number, png_dir, png_files=None, identifiers=None):
    """
    在截圖完成後,執行都市計畫圖資詳細查詢流程:
    1. 關閉分區查詢的訊息彈窗
    2. 點擊比例尺選單,選擇 2500
    3. 點擊紅框區域,顯示都市計畫圖訊息
    4. 擷取都市計畫圖訊息資料（返回資料，不再存檔）
    5. 點擊比例尺選單,選擇 1000
    6. 截圖並存檔 PNG/PDF（同時加入 png_files/identifiers，供合併 PDF 使用）

    🔥 修改：返回都市計畫圖資料字典，不再直接存檔
    """
    try:
        print("\n===== 都市計畫圖資詳細查詢流程 =====", flush=True)

        # 步驟 1: 關閉分區查詢的訊息彈窗
        print("步驟 1: 關閉分區查詢訊息彈窗...", flush=True)
        try:
            # 使用開發者工具中找到的正確選擇器
            driver.execute_script("""
                // 方法1: 找 div.titleButton.close (從開發者工具確認的選擇器)
                var closeButtons = document.querySelectorAll('div.titleButton.close');
                console.log('找到 div.titleButton.close 數量:', closeButtons.length);
                for (var i = 0; i < closeButtons.length; i++) {
                    closeButtons[i].click();
                    console.log('已點擊 div.titleButton.close', i);
                }

                // 方法2: 找標題為「關閉」的按鈕
                var titleCloseButtons = document.querySelectorAll('button[title="關閉"]');
                console.log('找到 button[title="關閉"] 數量:', titleCloseButtons.length);
                for (var i = 0; i < titleCloseButtons.length; i++) {
                    titleCloseButtons[i].click();
                    console.log('已點擊 button[title="關閉"]', i);
                }

                // 方法3: 找所有包含 titleButton 和 close 類別的元素
                var allCloseButtons = document.querySelectorAll('.titleButton.close, [class*="titleButton"][class*="close"]');
                console.log('找到 titleButton close 數量:', allCloseButtons.length);
                for (var i = 0; i < allCloseButtons.length; i++) {
                    allCloseButtons[i].click();
                    console.log('已點擊 titleButton close', i);
                }
            """)
            print("✓ 已關閉分區查詢訊息彈窗", flush=True)
            time.sleep(1)
        except Exception as e:
            print(f"關閉訊息彈窗時發生錯誤: {e}", flush=True)
            print("請手動關閉彈窗，然後按 Enter 繼續...", flush=True)
            input()

        # 步驟 2: 點擊比例尺選單,選擇 2500
        print("步驟 2: 設定地圖比例尺為 2500...", flush=True)
        try:
            scale_select_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "selScale"))
            )
            Select(scale_select_element).select_by_value("2500")
            print("✓ 已自動設定比例尺為 2500", flush=True)
            time.sleep(2)  # 等待地圖重新渲染
        except TimeoutException:
            print("⚠ 無法自動設定比例尺", flush=True)
            print("請手動點擊比例尺選單並選擇 2500,然後按 Enter 繼續...", flush=True)
            input()

        # 步驟 3: 點擊紅框區域,顯示都市計畫圖訊息（帶驗證和自動調整）
        print("步驟 3: 點擊地圖紅框區域以顯示都市計畫圖訊息...", flush=True)

        # 🔥 保存第一步的使用分區資料用於驗證
        expected_zone = None
        if len(all_zone_data) > 0:
            expected_zone = all_zone_data[-1].get('使用分區', '')
            print(f"📋 預期的使用分區（來自第一步）: {expected_zone}", flush=True)

        # 🔥 道路用地：跳過第二次點擊與建蔽容積擷取（道路本來就沒有建蔽率/容積率）
        is_road_only = bool(expected_zone) and "道路用地" in expected_zone

        click_success = False
        manual_click_needed = False

        try:
            # 🔥 道路用地：標記成功並跳出，不執行點擊邏輯
            if is_road_only:
                print("ℹ 偵測到「道路用地」，跳過第二次點擊（無建蔽容積資料可抓）", flush=True)
                click_success = True
                raise _SkipUrbanPlanQuery

            from selenium.webdriver.common.action_chains import ActionChains
            from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window

            # 先用 JavaScript 找到紅色邊框的 path 元素
            red_path_element = driver.execute_script("""
                // 尋找所有 path 元素
                var allPaths = document.querySelectorAll('path');
                console.log('找到 path 元素數量:', allPaths.length);

                // 尋找紅色邊框的 path (stroke="rgb(255, 0, 0)" 且 stroke-width="5")
                for (var i = 0; i < allPaths.length; i++) {
                    var path = allPaths[i];
                    var stroke = path.getAttribute('stroke');
                    var strokeWidth = path.getAttribute('stroke-width');
                    var fillOpacity = path.getAttribute('fill-opacity');

                    // 檢查是否為紅色邊框
                    if (stroke === 'rgb(255, 0, 0)' && strokeWidth === '5' && fillOpacity === '0') {
                        console.log('找到紅色邊框 path:', i);
                        return path;
                    }
                }
                return null;
            """)

            if red_path_element:
                actions = ActionChains(driver)
                size = red_path_element.size

                # 🔥 定義多個點擊位置嘗試順序
                click_positions = [
                    ("右邊界中點", size['width'] // 2, 0),           # 原始位置
                    ("右邊界向中心1/2", size['width'] // 4, 0),       # 向中心調整一半
                    ("中心點", 0, 0),                                  # 正中心
                    ("左邊界向中心1/2", -(size['width'] // 4), 0),    # 向左一半
                ]

                for attempt, (position_name, x_offset, y_offset) in enumerate(click_positions, 1):
                    print(f"嘗試 {attempt}/4: 點擊紅框{position_name}位置...", flush=True)

                    try:
                        # 點擊紅框
                        actions.move_to_element_with_offset(red_path_element, x_offset, y_offset).click().perform()

                        # 🔥 等待彈窗出現（增加等待時間並檢查）
                        print(f"   等待彈窗出現...", flush=True)
                        time.sleep(3)  # 等待3秒讓彈窗完全加載

                        # 🔥 使用多種方法檢查彈窗是否出現
                        popup_appeared = False
                        for wait_attempt in range(8):  # 最多等待8秒
                            # 方法1: 檢查頁面文本
                            page_text = driver.execute_script("return document.body.innerText;")

                            # 方法2: 檢查是否有彈窗元素
                            popup_elements = driver.execute_script("""
                                var popups = document.querySelectorAll('.popup, .modal, [class*="popup"], [class*="modal"]');
                                return popups.length;
                            """)

                            if ('使用分區整編名稱' in page_text or
                                '前往分區條文申請' in page_text or
                                '計畫區名稱' in page_text or
                                popup_elements > 0):
                                popup_appeared = True
                                print(f"   ✓ 彈窗已出現（等待了 {wait_attempt + 1} 秒）", flush=True)
                                break
                            time.sleep(1)

                        if not popup_appeared:
                            print(f"   ⚠️ 等待8秒後彈窗仍未出現，跳過此位置", flush=True)
                            continue

                        # 🔥 驗證：檢查彈出的數據是否為都市計畫資訊
                        # 使用更強大的方式擷取彈窗內的文字（包含HTML結構）
                        page_text = driver.execute_script("return document.body.innerText;")
                        popup_html = driver.execute_script("""
                            var popup = document.querySelector('.esriPopupWrapper');
                            return popup ? popup.innerHTML : '';
                        """)

                        # 合併兩種來源的文字
                        lines = page_text.split('\n')

                        # 🔥 調試：打印部分頁面內容
                        relevant_lines = [line for line in lines if '計畫區' in line or '建蔽率' in line or 'MAPKEY' in line or '使用分區' in line]
                        if relevant_lines:
                            print(f"   [調試] innerText找到: {relevant_lines[:8]}", flush=True)

                        # 也從HTML中解析（更可靠）
                        if popup_html and '使用分區完整名稱' in popup_html:
                            print(f"   [調試] HTML中找到'使用分區完整名稱'欄位", flush=True)

                        # 🔥 改進的驗證方式：優先檢查"使用分區完整名稱"，備用方案檢查其他欄位
                        has_plan_data = False
                        plan_region = None
                        zone_full_name = None  # 使用分區完整名稱
                        zone_alias = None
                        building_coverage = None
                        detected_zone = None  # 最終檢測到的使用分區

                        # 先嘗試從HTML中解析（更可靠）
                        if popup_html:
                            # 使用正則表達式從HTML中提取資料
                            import re

                            # 提取使用分區完整名稱
                            match = re.search(r'使用分區完整名稱[：:\s]*</td><td[^>]*>([^<]+)</td>', popup_html)
                            if match:
                                zone_full_name = match.group(1).strip()
                                detected_zone = zone_full_name
                                has_plan_data = True
                                print(f"   ✓ 從HTML提取到使用分區完整名稱: {zone_full_name}", flush=True)

                            # 提取計畫區名稱
                            match = re.search(r'計畫區名稱[：:\s]*</td><td[^>]*>([^<]+)</td>', popup_html)
                            if match:
                                plan_region = match.group(1).strip()
                                has_plan_data = True

                            # 提取使用分區別名
                            match = re.search(r'使用分區別名[：:\s]*</td><td[^>]*>([^<]+)</td>', popup_html)
                            if match:
                                zone_alias = match.group(1).strip()

                            # 提取建蔽率
                            match = re.search(r'建蔽率[^：:]*[：:\s]*</td><td[^>]*>([^<]+)</td>', popup_html)
                            if match:
                                building_coverage = match.group(1).strip()
                                has_plan_data = True

                        # 備用方案：從innerText解析
                        if not has_plan_data:
                            for line in lines:
                                has_colon = ':' in line or '：' in line

                                # 檢查"使用分區完整名稱"
                                if '使用分區完整名稱' in line and has_colon:
                                    zone_full_name = re.split('[：:]', line, 1)[-1].strip()
                                    detected_zone = zone_full_name
                                    has_plan_data = True

                                # 檢查"計畫區名稱"
                                elif '計畫區名稱' in line and has_colon:
                                    plan_region = re.split('[：:]', line, 1)[-1].strip()
                                    has_plan_data = True

                                # 檢查"使用分區別名"
                                elif '使用分區別名' in line and has_colon:
                                    zone_alias = re.split('[：:]', line, 1)[-1].strip()
                                    if not detected_zone:
                                        detected_zone = zone_alias

                                # 檢查"建蔽率"
                                elif '建蔽率' in line and has_colon:
                                    building_coverage = re.split('[：:]', line, 1)[-1].strip()
                                    has_plan_data = True

                        # 🔥 驗證邏輯
                        if has_plan_data:
                            print(f"   ✓ 檢測到都市計畫資料:", flush=True)
                            print(f"     計畫區: {plan_region}", flush=True)
                            print(f"     使用分區完整名稱: {zone_full_name}", flush=True)
                            print(f"     使用分區別名: {zone_alias}", flush=True)
                            print(f"     建蔽率: {building_coverage}", flush=True)

                            # 🔥 驗證是否與第一步的資料匹配
                            if expected_zone and detected_zone:
                                # 直接比對完整名稱
                                if zone_full_name and expected_zone in zone_full_name:
                                    print(f"✅ 驗證成功！使用分區完全匹配: {zone_full_name}", flush=True)
                                    click_success = True
                                    break
                                elif zone_full_name and zone_full_name in expected_zone:
                                    print(f"✅ 驗證成功！使用分區包含匹配: {zone_full_name}", flush=True)
                                    click_success = True
                                    break
                                # 嘗試透過別名關聯（備用方案）
                                elif zone_alias:
                                    is_related = False
                                    if '第一種住宅區' in expected_zone:
                                        is_related = '住一' in zone_alias or '住1' in zone_alias
                                    elif '第二種住宅區' in expected_zone:
                                        is_related = '住二' in zone_alias or '住2' in zone_alias
                                    elif '第三種住宅區' in expected_zone:
                                        is_related = '住三' in zone_alias or '住3' in zone_alias
                                    elif '第四種住宅區' in expected_zone:
                                        is_related = '住四' in zone_alias or '住4' in zone_alias
                                    elif '商業區' in expected_zone:
                                        is_related = '商' in zone_alias
                                    elif '工業區' in expected_zone:
                                        is_related = '工' in zone_alias

                                    if is_related:
                                        print(f"✅ 驗證成功！分區別名匹配: 預期={expected_zone}, 別名={zone_alias}", flush=True)
                                        click_success = True
                                        break
                                    else:
                                        print(f"⚠️ 分區不完全匹配: 預期={expected_zone}, 檢測={detected_zone}", flush=True)
                                        print(f"   但有都市計畫資料，嘗試下一個位置", flush=True)
                                else:
                                    print(f"⚠️ 無法驗證分區匹配，嘗試下一個位置", flush=True)
                            else:
                                # 沒有預期值可驗證，但有都市計畫資料就接受
                                print(f"✅ 檢測到都市計畫資料（無第一步資料可驗證）", flush=True)
                                click_success = True
                                break
                        else:
                            print(f"❌ 未檢測到都市計畫資料", flush=True)
                            # 關閉錯誤的彈窗
                            try:
                                driver.execute_script("""
                                    var closeButtons = document.querySelectorAll('div.titleButton.close, button[title="關閉"]');
                                    for (var i = 0; i < closeButtons.length; i++) {
                                        closeButtons[i].click();
                                    }
                                """)
                                time.sleep(1)
                            except:
                                pass

                    except Exception as e:
                        print(f"點擊位置 {position_name} 時發生錯誤: {e}", flush=True)

                if not click_success:
                    print("⚠️ 所有自動點擊位置都驗證失敗", flush=True)
                    manual_click_needed = True

            else:
                print("⚠ 未找到紅色邊框", flush=True)
                manual_click_needed = True

        except _SkipUrbanPlanQuery:
            pass  # 道路用地跳過，不視為錯誤
        except Exception as e:
            print(f"自動點擊紅框時發生錯誤: {e}", flush=True)
            manual_click_needed = True

        # 如果自動點擊失敗，請求手動點擊
        if manual_click_needed:
            print("=" * 39, flush=True)
            print("🔴 需要手動操作！", flush=True)
            print(f"請手動點擊地圖上的紅框區域（使用分區應為: {expected_zone}）", flush=True)
            print("然後按 Enter 繼續...", flush=True)
            print("=" * 39, flush=True)
            input()

        # 步驟 4: 擷取都市計畫圖街廓資訊並存到 JSON
        print("步驟 4: 擷取都市計畫圖街廓資訊...", flush=True)
        try:
            # 🔥 道路用地：跳過資料擷取與比對
            if is_road_only:
                print("ℹ 道路用地，跳過建蔽容積資料擷取與比對", flush=True)
                urban_plan_data = None
                raise _SkipUrbanPlanQuery

            time.sleep(2)  # 等待彈窗出現

            # 使用 JavaScript 擷取整個頁面的文字內容
            page_text = driver.execute_script("return document.body.innerText;")

            # 🔥 改為收集多筆街廓資訊（用 MAPKEY 作為分隔標識）
            all_urban_plan_entries = []  # 收集所有街廓資訊
            current_entry = None  # 當前正在解析的資料

            def create_new_entry():
                """建立新的街廓資訊資料結構"""
                return {
                    '地段': f"{area}{section}",
                    '地號': lot_number,
                    '計畫區名稱': '',
                    '使用分區完整名稱': '',
                    '使用分區別名': '',
                    '建蔽率': '',
                    '容積率': '',
                    '參考面積': '',
                    '附帶條件': '',
                    'MAPKEY': ''
                }

            # 解析文字內容
            lines = page_text.split('\n')
            for line in lines:
                has_colon = ':' in line or '：' in line

                # 🔥 遇到「計畫區名稱」表示新的一筆資料開始
                if '計畫區名稱' in line and has_colon:
                    # 如果有前一筆資料，先儲存
                    if current_entry and current_entry.get('MAPKEY'):
                        all_urban_plan_entries.append(current_entry)
                    # 建立新的資料結構
                    current_entry = create_new_entry()
                    value = re.split('[：:]', line, 1)[-1].strip()
                    if value:
                        current_entry['計畫區名稱'] = value
                        print(f"   找到計畫區名稱: {value}", flush=True)

                # 以下欄位需要 current_entry 存在
                elif current_entry is not None:
                    # 尋找「使用分區完整名稱」
                    if '使用分區完整名稱' in line and has_colon:
                        value = re.split('[：:]', line, 1)[-1].strip()
                        if value:
                            current_entry['使用分區完整名稱'] = value
                            print(f"   找到使用分區完整名稱: {value}", flush=True)

                    # 尋找「使用分區別名」
                    elif '使用分區別名' in line and has_colon:
                        value = re.split('[：:]', line, 1)[-1].strip()
                        if value:
                            current_entry['使用分區別名'] = value
                            print(f"   找到使用分區別名: {value}", flush=True)

                    # 尋找「建蔽率」（文字）
                    elif '建蔽率' in line and has_colon:
                        value = re.split('[：:]', line, 1)[-1].strip()
                        # 移除「(文字)」等標註
                        value = re.sub(r'\(.*?\)', '', value).strip()
                        if value:
                            current_entry['建蔽率'] = value
                            print(f"   找到建蔽率: {value}", flush=True)

                    # 尋找「容積率」（文字）
                    elif '容積率' in line and has_colon:
                        value = re.split('[：:]', line, 1)[-1].strip()
                        # 移除「(文字)」等標註
                        value = re.sub(r'\(.*?\)', '', value).strip()
                        if value:
                            current_entry['容積率'] = value
                            print(f"   找到容積率: {value}", flush=True)

                    # 尋找「參考面積」
                    elif '參考面積' in line and has_colon:
                        value = re.split('[：:]', line, 1)[-1].strip()
                        # 移除「(平方公尺)」等標註
                        value = re.sub(r'\(.*?\)', '', value).strip()
                        if value:
                            current_entry['參考面積'] = value
                            print(f"   找到參考面積: {value}", flush=True)

                    # 尋找「附帶條件」
                    elif '附帶條件' in line and has_colon:
                        value = re.split('[：:]', line, 1)[-1].strip()
                        if value:
                            current_entry['附帶條件'] = value
                            print(f"   找到附帶條件: {value}", flush=True)

                    # 尋找「MAPKEY」
                    elif 'MAPKEY' in line and has_colon:
                        value = re.split('[：:]', line, 1)[-1].strip()
                        if value:
                            current_entry['MAPKEY'] = value
                            print(f"   找到MAPKEY: {value}", flush=True)

            # 🔥 最後一筆資料也要儲存
            if current_entry and current_entry.get('MAPKEY'):
                all_urban_plan_entries.append(current_entry)

            # 🔥 回傳所有街廓資訊（列表形式）
            if all_urban_plan_entries:
                print(f"✓ 已擷取 {len(all_urban_plan_entries)} 筆都市計畫圖街廓資訊", flush=True)
                urban_plan_data = all_urban_plan_entries  # 回傳列表
            else:
                print(f"⚠ 未找到都市計畫圖街廓資訊", flush=True)
                urban_plan_data = None

        except _SkipUrbanPlanQuery:
            pass  # 道路用地跳過，不視為錯誤
        except Exception as e:
            print(f"擷取都市計畫圖街廓資訊時發生錯誤: {e}", flush=True)
            import traceback
            traceback.print_exc()
            urban_plan_data = None

        # 步驟 5: 點擊比例尺選單,選擇 1000
        print("步驟 5: 設定地圖比例尺為 1000...", flush=True)
        try:
            scale_select_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "selScale"))
            )
            Select(scale_select_element).select_by_value("1000")
            print("✓ 已自動設定比例尺為 1000", flush=True)
            time.sleep(2)  # 等待地圖重新渲染
        except TimeoutException:
            print("⚠ 無法自動設定比例尺", flush=True)
            print("請手動點擊比例尺選單並選擇 1000,然後按 Enter 繼續...", flush=True)
            input()

        # 步驟 6: 截圖並存檔 PNG/PDF
        print("步驟 6: 截圖並存檔...", flush=True)

        # **新增：等待使用者調整網頁 (倒數計時 - 使用 print)**
        print("\033[93m請在接下來的 10 秒內，\n手動調整網頁縮放、移動地圖，以達到最佳截圖效果\033[0m", flush=True)
        countdown(10)

        # 定義 PNG 檔名
        png_filename = f"07_都市計畫圖-{area}{section}-{lot_number}.png"
        screenshot_path = os.path.join(png_dir, png_filename)

        # 保存截圖（只存 PNG，PDF 由最後 create_combined_pdf 統一合併）
        try:
            driver.save_screenshot(screenshot_path)
            print(f"✓ 都市計畫圖截圖已保存為 {screenshot_path}", flush=True)

            # 🔥 加入合併 PDF 用的清單（與「使用分區」截圖一同合併輸出）
            # 只 append PNG 路徑；identifiers 不再重覆 append，避免 PDF 檔名出現 859,859,...
            if png_files is not None:
                png_files.append(screenshot_path)

        except Exception as e:
            print(f"截圖時發生錯誤: {e}", flush=True)

        print("===== 都市計畫圖資詳細查詢流程完成 =====\n", flush=True)

        # 🔥 返回都市計畫圖資料
        return urban_plan_data

    except Exception as e:
        print(f"都市計畫圖資詳細查詢流程發生錯誤: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return None

# 初始化頁面
initialize_page(driver, url)

# 列表來收集所有 PNG 路徑和標識符
all_png_files = []
all_identifiers = []

# 🔥 列表來收集所有都市計畫資料
all_zone_data = []  # 收集所有「分區查詢資料」
all_urban_plan_data = []  # 收集所有「都市計畫圖街廓資訊」
skipped_plan_queries = []  # 🔥 記錄跳過都市計畫圖資查詢的資料

# 迴圈處理每一筆資料
for data in data_list:
    area, section, lot_number = data['area'], data['section'], data['lot_number']
    print(f"\033[96m處理資料: 區域 = {area}, 段 = {section}, 地號 = {lot_number}\033[0m", flush=True)
    
    # 重新載入頁面並初始化
    initialize_page(driver, url)

    # 土地使用分區查詢流程
    try:
        land_use_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@title='土地使用分區查詢']"))
        )
        driver.execute_script("arguments[0].click();", land_use_button)
        print("點擊了【土地使用分區查詢】", flush=True)
    except TimeoutException:
        print("土地使用分區查詢按鈕無法點擊", flush=True)
        continue

    # 行政區選擇
    try:
        dist_select_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "DIST"))
        )
        Select(dist_select_element).select_by_visible_text(area)
        print(f"已選擇行政區: {area}", flush=True)
    except (TimeoutException, StaleElementReferenceException) as e:
        print(f"選擇行政區時發生錯誤: {e}", flush=True)
        continue

    # 地段選擇
    try:
        sect_select_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "SECT"))
        )
        sect_select = Select(sect_select_element)
        if not select_section_with_retry(sect_select, section):
            print("地段選擇失敗，跳過該查詢", flush=True)
            continue
        # print(f"已選擇地段: {section}")
    except (TimeoutException, StaleElementReferenceException, NoSuchElementException) as e:
        print(f"選擇地段時發生錯誤: {e}", flush=True)
        continue

    # 填入地號
    try:
        lot_input_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "NUMBER"))
        )
        lot_input_element.clear()
        lot_input_element.send_keys(lot_number)
        print(f"已填入地號: {lot_number}", flush=True)
    except TimeoutException:
        print("填入地號時發生錯誤", flush=True)
        continue

    # 點擊查詢按鈕
    try:
        search_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and @class='btn btn-info btn-block']"))
        )
        driver.execute_script("arguments[0].click();", search_button)
        print("成功點擊查詢按鈕", flush=True)
    except TimeoutException:
        print("查詢按鈕無法點擊，跳過該查詢", flush=True)
        continue
    time.sleep(8)
    # 點擊查詢結果項目並加入左向右、右向左的比對
    try:
        target_title = f"{area}{section}{lot_number}"
        # print(f"嘗試點擊查詢結果項目: {target_title}", flush=True)

        # 列印查詢結果項目
        result_items = driver.find_elements(By.XPATH, "//ul[@id='searchResult']/li")
        result_found = False

        # 初次嘗試精確匹配「行政區 + 地段 + 地號」
        for item in result_items:
            title_text = item.get_attribute("title")
            # print(f"查詢結果項目: {title_text}", flush=True)
            if title_text == target_title:
                result_item = item
                result_found = True
                break
  
        # 若無法找到完整匹配，進行模糊匹配
        if not result_found:
            print(f"無法找到完整匹配，嘗試模糊比對: 以『行政區』開始並以『地號』結尾", flush=True)
            for item in result_items:
                title_text = item.get_attribute("title")
                if title_text.startswith(area) and title_text.endswith(lot_number):
                    result_item = item
                    result_found = True
                    print(f"模糊比對找到匹配: {title_text}", flush=True)
                    break

        # 若無法找到任何匹配結果，提示手動操作
        if not result_found:
            print(f"無法匹配到查詢結果項目: {target_title}", flush=True)
            print("請手動點擊目標項目，然後按 Enter 繼續...", flush=True)
            input()

        # 若找到匹配結果，嘗試點擊該項目
        else:
            try:
                WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, f"//ul[@id='searchResult']/li[@title='{result_item.get_attribute('title')}']"))
                )
                driver.execute_script("arguments[0].click();", result_item)
                print(f"成功點擊查詢結果: {result_item.get_attribute('title')}", flush=True)
            except TimeoutException:
                print(f"查詢結果項目無法點擊: {result_item.get_attribute('title')}", flush=True)
                print("請手動點擊目標項目，然後按 Enter 繼續...", flush=True)
                input()

    except Exception as e:
        print(f"點擊查詢結果項目時發生錯誤: {e}", flush=True)
        print("請手動點擊目標項目，然後按 Enter 繼續...", flush=True)
        input()

    # **新增：擷取土地使用分區查詢的彈出視窗資料**
    print("正在擷取土地使用分區資料(第一次)...", flush=True)

    # 🔥 在擷取資料前重新設定視窗大小（確保視窗位置正確）
    try:
        driver.set_window_size(urbangis_window_width, urbangis_window_height)
        driver.set_window_position(urbangis_window_x, urbangis_window_y)
    except:
        pass

    zone_data_1 = extract_zone_data(driver, area, section)
    if zone_data_1:
        all_zone_data.append(zone_data_1)
        # 🔥 只有在成功取得使用分區時才顯示
        if zone_data_1.get('使用分區'):
            print(f"✓ 成功擷取分區查詢資料(第一次): 使用分區={zone_data_1.get('使用分區')}", flush=True)
        else:
            print("⚠️ 第一次查詢：未取得使用分區資料", flush=True)
    else:
        print("❌ 警告：第一次分區查詢資料擷取失敗！", flush=True)

    # 關閉分區查詢和查詢結果
    try:
        driver.execute_script("""
            var buttons = document.querySelectorAll('button[title="關閉"]');
            for (var i = 0; i < buttons.length; i++) {
                buttons[i].click();
            }
        """)
        print("已關閉分區查詢和查詢結果視窗", flush=True)
    except Exception as e:
        print(f"關閉分區查詢視窗時發生錯誤: {e}", flush=True)

    # 定位查詢流程
    try:
        # 先檢查按鈕是否存在
        toggle_buttons = driver.find_elements(By.XPATH, "//button[@type='button' and @aria-label='Toggle navigation']")

        if toggle_buttons:
            for idx, btn in enumerate(toggle_buttons):
                pass

        # 🔥 如果按鈕被隱藏，先用 JavaScript 讓它顯示
        if toggle_buttons and not toggle_buttons[0].is_displayed():
            driver.execute_script("arguments[0].style.display = 'block';", toggle_buttons[0])
            driver.execute_script("arguments[0].style.visibility = 'visible';", toggle_buttons[0])
            time.sleep(0.5)

        # 等待按鈕可點擊（增加等待時間）
        toggle_button = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//button[@type='button' and @aria-label='Toggle navigation']"))
        )

        # 🔥 強制使用 JavaScript 點擊（不管按鈕是否顯示）
        try:
            driver.execute_script("arguments[0].click();", toggle_button)
            print("✓ 點擊了【功能表按鈕】(JavaScript)", flush=True)
        except Exception as e1:
            raise

        time.sleep(2)  # 🔥 增加等待時間，確保選單完全展開

        # 🔥 檢查功能表是否成功展開
        try:
            menu_element = driver.find_element(By.ID, "navbarSupportedContent")
            menu_display = driver.execute_script("return window.getComputedStyle(arguments[0]).display;", menu_element)
            menu_classes = menu_element.get_attribute("class")

            # 🔥 如果功能表沒有展開，強制展開
            if menu_display == 'none' or 'show' not in menu_classes:
                driver.execute_script("arguments[0].style.display = 'block';", menu_element)
                driver.execute_script("arguments[0].classList.add('show');", menu_element)
                time.sleep(1)
        except Exception as e:
            pass

    except TimeoutException as e:
        print(f"[ERROR] 功能表按鈕無法點擊 - TimeoutException: {e}", flush=True)

        # 嘗試截圖以協助診斷
        try:
            screenshot_path = os.path.join(png_dir, f"debug_menu_button_fail_{lot_number}.png")
            driver.save_screenshot(screenshot_path)
        except:
            pass

        continue
    except Exception as e:
        print(f"[ERROR] 功能表按鈕點擊時發生未預期錯誤: {e}", flush=True)
        continue

    try:
        # 先檢查所有可能的定位查詢按鈕
        locate_buttons = driver.find_elements(By.XPATH, "//div[contains(@class, 'function_xySearch_btn')]")

        for idx, btn in enumerate(locate_buttons):
            pass

        # 嘗試找到並點擊定位查詢按鈕
        locate_search_button = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//div[@class='m-2 font-weight-bold hoverHand function_xySearch_btn' and text()='定位查詢']"))
        )

        # 🔥 如果按鈕不可見，強制顯示
        if not locate_search_button.is_displayed():
            driver.execute_script("arguments[0].style.display = 'block';", locate_search_button)
            driver.execute_script("arguments[0].style.visibility = 'visible';", locate_search_button)
            driver.execute_script("arguments[0].style.opacity = '1';", locate_search_button)
            time.sleep(0.5)

        driver.execute_script("arguments[0].click();", locate_search_button)
        print("✓ 點擊了【定位查詢】", flush=True)

        driver.execute_script("document.getElementById('navbarSupportedContent').style.display = 'none';")
        print("✓ 已隱藏功能表選單", flush=True)
        time.sleep(1)

    except TimeoutException as e:
        print(f"[ERROR] 定位查詢按鈕無法點擊 - TimeoutException: {e}", flush=True)

        # 檢查功能表是否已展開
        try:
            menu_element = driver.find_element(By.ID, "navbarSupportedContent")
            menu_display = driver.execute_script("return window.getComputedStyle(arguments[0]).display;", menu_element)
        except:
            pass

        # 截圖診斷
        try:
            screenshot_path = os.path.join(png_dir, f"debug_locate_button_fail_{lot_number}.png")
            driver.save_screenshot(screenshot_path)
        except:
            pass

        continue
    except Exception as e:
        print(f"[ERROR] 定位查詢按鈕點擊時發生未預期錯誤: {e}", flush=True)
        continue

    try:
        dropdown = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "funcItemList"))
        )
        Select(dropdown).select_by_visible_text("地籍")
        print("點擊了'地籍'", flush=True)
    except TimeoutException:
        print("地籍選項無法點擊", flush=True)
        continue

    # 選擇行政區和地段、地號
    try:
        # 行政區選擇部分
        # print(f"開始選擇行政區: {area}", flush=True)
        area_select_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "DIST05"))
        )
        area_select = Select(area_select_element)
        area_select.select_by_visible_text(area)
        print(f"定位查詢 - 已選擇行政區: {area}", flush=True)

        # 地段選擇部分
        # print(f"開始選擇地段: {section}", flush=True)
        section_select_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "SECT05"))
        )
        section_select = Select(section_select_element)
        if not select_section_with_retry(section_select, section):
            print("地段選擇失敗", flush=True)
            continue

        # 填入地號部分
        try:
            # print(f"開始填入地號: {lot_number}", flush=True)
            lot_input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "NUMBER05"))
            )
            
            # 清除並填入地號，重試機制以確保填入成功
            for attempt in range(3):
                lot_input_element.clear()
                lot_input_element.send_keys(lot_number)
                time.sleep(1.5)  # 等待輸入穩定
                
                # 確認地號已成功填入
                filled_lot_number = lot_input_element.get_attribute("value")
                if filled_lot_number == lot_number:
                    print(f"成功填入地號: {lot_number}", flush=True)
                    break
                else:
                    print(f"地號填入失敗，重試中 ({attempt + 1}/3)...", flush=True)
            else:
                print(f"地號填入失敗，無法確認地號: {lot_number}", flush=True)
                continue

        except TimeoutException:
            print("地號欄位載入超時", flush=True)
            continue

    except TimeoutException:
        print("行政區、地段或地號欄位載入超時", flush=True)
        continue

    # 點擊定位查詢的「查詢」按鈕
    try:
        # print("試著點擊定位查詢的查詢按鈕", flush=True)

        # 確認查詢按鈕是否出現並可點擊
        search_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[@id='UBA030100_TOP']//button[@id='searchBtn' and contains(@class, 'btn-info')]")
            )
        )
        
        # 使用 JavaScript 點擊查詢按鈕
        driver.execute_script("arguments[0].click();", search_button)
        print("成功點擊定位查詢按鈕")
    except TimeoutException:
        print("定位查詢的查詢按鈕未出現或無法點擊，重試中...")
        continue  # 若按鈕無法點擊則跳過當前查詢
    except Exception as e:
        print(f"定位查詢按鈕操作過程中發生錯誤: {e}", flush=True)

    # 點擊查詢結果項目並加入手動點擊提示
    try:
        # 等待查詢結果區塊出現
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "UBA030100_BOTTOM"))
        )

        # 格式化查詢結果的識別資訊
        target_title = f"{area}{section}{lot_number}"
        # print(f"嘗試點擊查詢結果項目: {target_title}", flush=True)

        # 取得所有查詢結果的選項並列印
        result_items = driver.find_elements(By.XPATH, "//div[@id='UBA030100_BOTTOM']//ul[@id='searchResult']/li")
        # print("查詢結果選項:", flush=True)
        result_found = False
        for item in result_items:
            title_text = item.get_attribute("title")
            # print(f"查詢結果項目: {title_text}", flush=True)
            if title_text == target_title:
                result_item = item
                result_found = True
                break

        if not result_found:
            # 若無法比對到結果，暫停並提示使用者手動點擊
            print(f"無法自動匹配到查詢結果項目: {target_title}", flush=True)
            print("請手動點擊目標項目，然後按 Enter 繼續...", flush=True)
            input()
        else:
            # 確保查詢結果項目可點擊
            try:
                WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, f"//div[@id='UBA030100_BOTTOM']//ul[@id='searchResult']/li[@title='{target_title}']"))
                )
                driver.execute_script("arguments[0].click();", result_item)
                print(f"成功點擊查詢結果: {target_title}", flush=True)
            except TimeoutException:
                print(f"查詢結果項目 {target_title} 無法點擊", flush=True)
                print("請手動點擊目標項目，然後按 Enter 繼續...", flush=True)
                input()

    except Exception as e:
        print(f"點擊查詢結果項目時發生錯誤: {e}")
        print("請手動點擊目標項目，然後按 Enter 繼續...", flush=True)
        input()

    # **新增：擷取定位查詢的彈出視窗資料**
    print("正在擷取定位查詢資料(第二次)...", flush=True)
    zone_data_2 = extract_zone_data(driver, area, section)

    # 🔥 檢查是否為「都市計畫外」
    is_outside_urban_plan = False
    if zone_data_2:
        all_zone_data.append(zone_data_2)
        zone_name = zone_data_2.get('使用分區', '')

        # 🔥 判斷是否為「都市計畫外」
        if '都市計畫外' in zone_name or '非都市' in zone_name:
            is_outside_urban_plan = True
            print(f"✓ 成功擷取分區查詢資料(第二次): 使用分區={zone_name}", flush=True)
            print(f"\033[93m⚠️ 偵測到「{zone_name}」，將跳過都市計畫圖資詳細查詢流程\033[0m", flush=True)
        elif zone_name:
            print(f"✓ 成功擷取分區查詢資料(第二次): 使用分區={zone_name}", flush=True)
        else:
            print("⚠️ 第二次查詢：未取得使用分區資料", flush=True)
    else:
        print("❌ 警告：第二次分區查詢資料擷取失敗！", flush=True)

    # 呼叫 finalize_query 並傳入當前資料
    finalize_query(driver, area, section, lot_number, png_dir, all_png_files, all_identifiers)

    # **新增:都市計畫圖資詳細查詢流程(在截圖完成後執行)**
    # 🔥 只有在非「都市計畫外」時才執行
    if not is_outside_urban_plan:
        urban_plan_data = query_urban_planning_details(
            driver, area, section, lot_number, png_dir,
            png_files=all_png_files
        )
        if urban_plan_data:
            # 🔥 urban_plan_data 現在是列表，展開加入
            if isinstance(urban_plan_data, list):
                all_urban_plan_data.extend(urban_plan_data)
            else:
                all_urban_plan_data.append(urban_plan_data)
    else:
        # 🔥 記錄跳過的資料，稍後在 PDF 生成後顯示
        skipped_plan_queries.append(f"{area}{section}-{lot_number}")

# 迴圈結束後，生成合併的 PDF
create_combined_pdf(all_png_files, all_identifiers, basic_data_dir)

# 🔥 顯示跳過都市計畫圖資查詢的訊息
if skipped_plan_queries:
    print(f"\033[93m⏩ 已跳過都市計畫圖資詳細查詢流程\033[0m\n", flush=True)

# 🔥 統一儲存所有都市計畫資料為一個 JSON 檔案
print("\n===== 儲存都市計畫資料 =====", flush=True)
try:
    # 🔥 過濾掉空資料：地號為空就過濾掉（地號是必要欄位）
    filtered_zone_data = []
    for zone_data in all_zone_data:
        # 🔥 地號必須有值
        if zone_data.get('地號', '').strip():
            filtered_zone_data.append(zone_data)
        else:
            print(f"   [過濾] 跳過空的分區查詢資料（地號為空）", flush=True)

    # 🔥 過濾掉空的都市計畫圖街廓資訊：地號為空就過濾掉（地號是必要欄位）
    filtered_urban_plan_data = []
    for urban_data in all_urban_plan_data:
        # 🔥 地號必須有值
        if urban_data.get('地號', '').strip():
            filtered_urban_plan_data.append(urban_data)
        else:
            print(f"   [過濾] 跳過空的都市計畫圖街廓資訊（地號為空）", flush=True)

    combined_data = {
        '分區查詢資料': filtered_zone_data,
        '都市計畫圖街廓資訊': filtered_urban_plan_data
    }

    json_filename = "都市計畫資料彙整.json"
    json_path = os.path.join(other_data_dir, json_filename)

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(combined_data, f, ensure_ascii=False, indent=2)

    print(f"✓ 已儲存都市計畫資料至: {json_path}", flush=True)
    print(f"   分區查詢資料: {len(filtered_zone_data)} 筆（已過濾 {len(all_zone_data) - len(filtered_zone_data)} 筆空資料）", flush=True)
    print(f"   都市計畫圖街廓資訊: {len(filtered_urban_plan_data)} 筆（已過濾 {len(all_urban_plan_data) - len(filtered_urban_plan_data)} 筆空資料）", flush=True)
except Exception as e:
    print(f"儲存都市計畫資料時發生錯誤: {e}", flush=True)
    import traceback
    traceback.print_exc()

# print("查詢完成，關閉瀏覽器", flush=True)
driver.quit()

def notify_main_program():
    print("高雄都市計畫已完成執行", flush=True)

if __name__ == "__main__":
    try:
        # 子程式的主要邏輯
        # 在這裡執行子程式的操作，例如表單填寫和資料處理
        pass  # print("執行高雄都市計畫主要邏輯", flush=True)

    except Exception as e:
        print(f"高雄都市計畫執行過程中發生錯誤: {e}", flush=True)
    finally:
        # 在結尾通知主程式
        notify_main_program()