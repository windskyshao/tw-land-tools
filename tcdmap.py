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
print("歡迎使用【國土分區查詢系統】自動化小程式，模組載入中...", flush=True)
import json
import time
import ctypes
import keyboard  # 用於系統級鍵盤事件
from fpdf import FPDF
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, NoSuchElementException
from webdriver_helper import create_chrome_driver

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
    except Exception:
        return 1.0

# 🔥 視窗大小設定：從 main.py 生成的設定檔讀取，自動填滿剩餘空間
tcdmap_window_width = 1024
tcdmap_window_height = 1024
tcdmap_window_x = 0
tcdmap_window_y = 0

try:
    config_path = os.path.join(BASE_DIR, 'window_config.json')
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            window_config = json.load(f)

        work_area = window_config.get('work_area', {})
        main_window = window_config.get('main_window', {})

        if work_area and main_window:
            # 取得 DPI 縮放比例
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

            # 高度使用完整工作區域高度
            chrome_y = work_top
            chrome_height = work_bottom - work_top - 20

            # 確保視窗大小合理
            if chrome_width < 800:
                chrome_width = 800
            if chrome_height < 600:
                chrome_height = 600

            # 轉換為 Chrome 的邏輯像素
            chrome_width_logical = int(chrome_width / dpi_scale)
            chrome_x_logical = int(chrome_x / dpi_scale)
            chrome_y_logical = int(chrome_y / dpi_scale)

            # 🔥 Chrome 視窗高度需要加上標題欄高度（約 32px 邏輯像素）
            # 這樣整個視窗（包含標題欄）才會和主程式一樣高
            CHROME_TITLEBAR_HEIGHT = 32
            chrome_height_logical = int(chrome_height / dpi_scale) + CHROME_TITLEBAR_HEIGHT

            tcdmap_window_width = chrome_width_logical
            tcdmap_window_height = chrome_height_logical
            tcdmap_window_x = chrome_x_logical
            tcdmap_window_y = chrome_y_logical
except Exception as e:
    pass

# 讀取 data.json 文件
with open(get_data_json_path(), 'r', encoding='utf-8') as f:
    data_list = json.load(f)

if not data_list:
    print("data.json 是空的，請檢查資料。", flush=True)
    exit()

# 取得第一組資料，建立基礎目錄結構
first_entry = data_list[0]
base_region = first_entry['area']
base_section = first_entry['section']
base_lot_number = first_entry['lot_number']
# 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
base_directory = get_work_folder(f"{base_region}{base_section}-{base_lot_number}")
basic_data_dir = os.path.join(base_directory, "1.基本資料")
png_dir = os.path.join(basic_data_dir, "png")

# 建立目錄結構
os.makedirs(png_dir, exist_ok=True)
print(f"已建立目錄 {base_directory}", flush=True)

# 初始化 WebDriver
options = webdriver.ChromeOptions()
# 🔥 不在啟動參數設定視窗大小，改為啟動後再調整（避免 DPI 改變時崩潰）
# options.add_argument("--window-size=1024,1024")
driver = create_chrome_driver(options=options)

# 🔥 啟動後立即設定視窗大小和位置
try:
    driver.set_window_size(tcdmap_window_width, tcdmap_window_height)
    driver.set_window_position(tcdmap_window_x, tcdmap_window_y)
except Exception as e:
    print(f"[WARNING] 視窗設定失敗: {e}，繼續執行")

# 定義替換字典：data.json 中的段名到網頁顯示的段名
section_name_map = {
    "磚子磘段": "磚子段",
    # 可以加入更多映射對應
}


def normalize_cjk(text: str) -> str:
    """
    正規化中文形似字，避免 Unicode 不同導致比對失敗。
    例：巿(U+5DFF 布匹) -> 市(U+5E02 城市)，臺 -> 台
    """
    return (text
            .replace(chr(0x5dff), chr(0x5e02))   # 巿 -> 市
            .replace(chr(0x81fa), chr(0x53f0))   # 臺 -> 台
            )

# 定義函數來生成指定城市的地圖網址
# 🔥 2026-03: 網站已從 up.tcd.gov.tw 永久移轉至 nluz.nlma.gov.tw，URL 格式新增 version 參數
def get_city_url(city_name):
    # 🔥 CN 參數放前面，符合官網 JS 的 URL 格式：?CN=縣市&version=版本
    # 🔥 2026-03: 各縣市多已進入報部版階段，改用報部版
    return f"https://nluz.nlma.gov.tw/ex3smap/default.aspx?CN={city_name}&version=報部版"

# 函數：檢查並關閉模態窗口
def ensure_modal_closed():
    try:
        # 🔥 縮短等待時間，彈窗通常很快就會出現
        agree_button = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary[data-bs-dismiss='modal']"))
        )
        agree_button.click()
        print("已點擊「我同意並進入系統」按鈕。", flush=True)
    except (TimeoutException, NoSuchElementException):
        print("模態窗口已關閉或不存在。", flush=True)
    except ElementNotInteractableException:
        print("「我同意並進入系統」按鈕不可點擊。", flush=True)

# 函數：關閉彈出表格視窗
def close_popup_window():
    try:
        close_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".esri-popup__button[title='關閉']"))
        )
        close_button.click()
        print("已關閉彈出表格視窗。", flush=True)
    except TimeoutException:
        print("彈出表格視窗未顯示或已關閉。", flush=True)

# 等待頁面加載完畢，直到 "讀取中" 提示消失
# 🔥 2026-03: 網站改版後 loader id 已從 "loader" 改為 "loading"
def wait_for_loader_to_disappear():
    try:
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, "loading"))
        )
        print("頁面已加載完成，'讀取中'提示消失。", flush=True)
    except TimeoutException:
        print("等待 '讀取中' 提示消失超時，可能仍有遮擋。", flush=True)

# 計算最相似的段名選項
def find_best_matching_section(target_section, options):
    # 1. 優先完全相符
    for option in options:
        if option.text == target_section:
            return option

    # 2. 用字元交集相似度（不計位置，看共有幾個字）
    target_chars = set(target_section)
    best_option = None
    highest_similarity = 0

    for option in options:
        option_text = option.text
        option_chars = set(option_text)
        # 共同字元數（不重複計算）
        common = len(target_chars & option_chars)
        if common > highest_similarity:
            highest_similarity = common
            best_option = option

    # 3. 相似度門檻：必須共用超過 1 個字（避免只靠「段」字湊合）
    #    例：「花園段」需至少共用 2 個字才算有效匹配
    min_threshold = max(2, len(target_section) // 2)
    if highest_similarity < min_threshold:
        return None  # 沒有足夠相似的選項，回傳 None → 跳過此筆

    return best_option

# 列表來收集所有 PNG 路徑和標識符
all_png_files = []
all_identifiers = []

# 記錄上一筆資料的縣市
previous_city = None

# 遍歷 data.json 中的每一組數據
for data in data_list:
    city = data['city']
    area = data['area']
    section = data['section']
    lot_number = data['lot_number']
    coordinates = data['coordinates']

    print(f"處理資料: 城市 = {city}, 區域 = {area}, 段 = {section}, 地號 = {lot_number}, 座標 = {coordinates}", flush=True)

    # 取得替換後的段名，用於選擇操作
    section_to_use = section_name_map.get(section, section)

    # 判斷是否需要重新載入頁面
    if city != previous_city:
        city_url = get_city_url(city)
        driver.get(city_url)
        # 🔥 網頁載入後重新設定視窗大小（避免被網站重置）
        try:
            driver.set_window_size(tcdmap_window_width, tcdmap_window_height)
            driver.set_window_position(tcdmap_window_x, tcdmap_window_y)
        except:
            pass
        print(f"已打開 {city} 的地圖頁面", flush=True)

        # 🔥 自動判斷是否需要縮放頁面
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
                    zoom_count = 5  # 175%：按 5 次
                elif current_dpi >= 1.5:
                    zoom_count = 4  # 150%：按 4 次
                elif current_dpi >= 1.25:
                    # 🔥 小螢幕 125% 需要更多縮放
                    if logical_width < 1200:
                        zoom_count = 4  # 小螢幕 125%：按 4 次
                    else:
                        zoom_count = 3  # 大螢幕 125%：按 3 次
                else:
                    zoom_count = 1  # 100%：按 1 次
        except Exception:
            pass

        if zoom_count > 0:
            try:
                body = driver.find_element(By.TAG_NAME, 'body')
                body.click()
                time.sleep(0.2)

                # 🔥 快速連續縮放，不需要每次都等待
                for _ in range(zoom_count):
                    keyboard.press_and_release('ctrl+-')
                    time.sleep(0.1)  # 縮短等待時間
            except Exception:
                pass

        ensure_modal_closed()
    else:
        close_popup_window()  # 關閉上次查詢的結果視窗

    # 🔥 等待 cada_townname 真正載入（有選項才算完成），最多等 20 秒
    try:
        WebDriverWait(driver, 20).until(
            lambda d: len(d.find_element(By.ID, "cada_townname")
                          .find_elements(By.TAG_NAME, "option")) > 0
        )
    except Exception:
        pass

    # 選擇 "鄉鎮市區"
    try:
        town_dropdown = driver.find_element(By.ID, "cada_townname")
        town_options = town_dropdown.find_elements(By.TAG_NAME, "option")
        print(f"[診斷] cada_townname 選項數量: {len(town_options)}，前3筆: {[o.text for o in town_options[:3]]}", flush=True)
        for option in town_options:
            if normalize_cjk(option.text) == normalize_cjk(area):
                option.click()
                print(f"已選擇鄉鎮市區: {area}", flush=True)
                break
        else:
            available = [o.text for o in town_options]
            print(f"找不到鄉鎮市區: {area}，可用選項: {available}", flush=True)
    except Exception as e:
        print(f"無法操作鄉鎮市區下拉: {e}", flush=True)

    # 🔥 等待段名下拉載入（選完鄉鎮市區後 cada_secname 會更新）
    try:
        WebDriverWait(driver, 10).until(
            lambda d: len(d.find_element(By.ID, "cada_secname")
                          .find_elements(By.TAG_NAME, "option")) > 0
        )
    except Exception:
        pass

    # 選擇 "段名" 使用最佳匹配邏輯
    section_found = False
    try:
        section_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "cada_secname"))
        )
        section_dropdown.click()
        time.sleep(1)  # 等待選項載入

        # 使用最佳匹配邏輯
        best_option = find_best_matching_section(section_to_use, section_dropdown.find_elements(By.TAG_NAME, "option"))
        if best_option:
            best_option.click()
            print(f"已選擇段名: {best_option.text}", flush=True)
            section_found = True
        else:
            print(f"找不到段名: {section}", flush=True)

    except TimeoutException:
        print(f"無法找到或選擇段名: {section}", flush=True)

    # 如果未找到段名，跳過該筆資料
    if not section_found:
        print(f"找不到段名: {section}，跳過此筆資料。", flush=True)
        previous_city = city  # 保持 previous_city 不變
        continue

    # 輸入 "地號"
    try:
        lot_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "cada_num"))
        )
        lot_input.clear()
        lot_input.send_keys(lot_number)
        print(f"已輸入地號: {lot_number}", flush=True)
    except TimeoutException:
        print(f"無法找到或輸入地號: {lot_number}", flush=True)

    # 點擊「查詢」按鈕
    try:
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "cada_search"))
        )
        search_button.click()
        print("已點擊查詢按鈕。", flush=True)
    except TimeoutException:
        print("無法找到或點擊查詢按鈕。", flush=True)

    # 等待查詢結果載入完成
    time.sleep(5)

    # 截圖並保存
    filename_base = f"08_國土分區-{area}{section}-{lot_number}"
    screenshot_path = os.path.join(png_dir, f"{filename_base}.png")
    driver.save_screenshot(screenshot_path)
    print(f"\033[93m已截圖並保存為 {screenshot_path}\033[0m", flush=True)

    # 收集 PNG 路徑和標識符
    all_png_files.append(screenshot_path)
    all_identifiers.append(f"{area}{section}-{lot_number}")

    # 更新 previous_city 為當前 city
    previous_city = city

# 合併所有 PNG 成為一個 PDF
def create_combined_pdf(png_files, identifiers, pdf_dir):
    if not png_files:
        print("沒有找到任何 PNG 檔案，無法生成 PDF。", flush=True)
        return

    area_section_dict = {}
    for identifier in identifiers:
        parts = identifier.split('-')
        region = parts[0]
        lot = '-'.join(parts[1:])
        area_section_dict.setdefault(region, []).append(lot)

    pdf_filename = "08_國土分區-" + "+".join(
        [f"{region}-{','.join(lots)}" for region, lots in area_section_dict.items()]
    ) + ".pdf"
    pdf_path = os.path.join(pdf_dir, pdf_filename)

    try:
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        for png in png_files:
            pdf.add_page()
            pdf.image(png, x=0, y=0, w=297, h=210)
        pdf.output(pdf_path)
        print(f"\033[93mPDF 已經儲存至 {pdf_path}\033[0m", flush=True)
    except Exception as e:
        print(f"將 PNG 轉成 PDF 時發生錯誤: {e}", flush=True)

# 呼叫 create_combined_pdf 並生成合併的 PDF
create_combined_pdf(all_png_files, all_identifiers, basic_data_dir)

# 關閉瀏覽器
driver.quit()
print("關閉瀏覽器", flush=True)
print("國土分區查詢已完成執行", flush=True)  #  <--- 在這裡加入通知訊息