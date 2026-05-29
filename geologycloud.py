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
import json
import time
import ctypes
import keyboard  # 用於系統級鍵盤事件
from PIL import Image
from fpdf import FPDF
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window

# 🔥 基準目錄設定（data.json 和工作資料夾的位置）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

# 清除終端機並顯示歡迎訊息
os.system('cls')
print("歡迎使用【地質雲】自動化小程式，模組載入中...", flush=True)

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
geology_window_width = 1024
geology_window_height = 1024
geology_window_x = 0
geology_window_y = 0

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

            geology_window_width = chrome_width_logical
            geology_window_height = chrome_height_logical
            geology_window_x = chrome_x_logical
            geology_window_y = chrome_y_logical
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
print(f"基礎目錄結構已建立於 {base_directory}", flush=True)

# 設定 WebDriver
options = webdriver.ChromeOptions()
# options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36")
# 🔥 不在啟動參數設定視窗大小，改為啟動後再調整（避免 DPI 改變時崩潰）
# options.add_argument("--window-size=1024,1024")
# options.add_argument("--window-position=0,0")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# 使用 ChromeDriverManager 安裝 ChromeDriver 並應用選項
driver = create_chrome_driver(options=options)

# 🔥 啟動後立即設定視窗大小和位置
try:
    driver.set_window_size(geology_window_width, geology_window_height)
    driver.set_window_position(geology_window_x, geology_window_y)
except Exception as e:
    print(f"[WARNING] 視窗設定失敗: {e}，繼續執行")

# 指定要打開的網址
url = "https://www.geologycloud.tw/map/liquefaction/zh-tw"
time.sleep(2)
driver.get(url)
# 🔥 網頁載入後重新設定視窗大小（避免被網站重置）
try:
    driver.set_window_size(geology_window_width, geology_window_height)
    driver.set_window_position(geology_window_x, geology_window_y)
except:
    pass

# 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
verify_and_fix_chrome_window(driver)

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
            zoom_count = 2  # 175%：按 2 次
        elif current_dpi >= 1.5:
            zoom_count = 1  # 150%：按 1 次
        elif current_dpi >= 1.25 and logical_width < 1200:
            zoom_count = 1  # 小螢幕 + 125%：按 1 次
        else:
            zoom_count = 0  # 100%、125%（大螢幕）：不縮放
except Exception:
    pass

if zoom_count > 0:
    try:
        body = driver.find_element(By.TAG_NAME, 'body')
        body.click()
        time.sleep(0.5)

        for _ in range(zoom_count):
            keyboard.press_and_release('ctrl+-')
            time.sleep(0.3)

        # 縮放後等待頁面穩定
        time.sleep(1)
    except Exception:
        pass

driver.implicitly_wait(10)

# 點擊 "我知道了" 按鈕（有縮放時跳過，因為按鈕可能已消失）
if zoom_count == 0:
    try:
        know_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btn_Soillique_OK"))
        )
        know_button.click()
        print("已點擊「我知道了」按鈕。", flush=True)
    except TimeoutException:
        print("無法找到「我知道了」按鈕。", flush=True)

# 列表來收集所有 PNG 路徑和標識符
all_png_files = []
all_identifiers = []
all_data_entries = []  # 儲存所有處理過的資料項目

# 遍歷 data.json 中的每一組數據
for index, data in enumerate(data_list):
    city = data['city']
    area = data['area']
    section = data['section']
    lot_number = data['lot_number']
    coordinates = data['coordinates']

    print(f"處理資料 {index+1}/{len(data_list)}: 城市 = {city}, 區域 = {area}, 段 = {section}, 地號 = {lot_number}, 座標 = {coordinates}", flush=True)

    try:
        # 輸入座標到搜尋欄位
        search_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#map-geocoding-input input[type='text']"))
        )
        search_input.clear()
        search_input.send_keys(coordinates)

        # 模擬滑鼠懸停到搜尋圖示
        search_icon = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#map-geocoding-input .fa-search"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(search_icon).perform()
        time.sleep(0.5)

        # 點擊搜尋按鈕
        search_icon.click()
        print("已點擊搜尋按鈕。", flush=True)

        # 等待結果列表項目顯示並點擊
        list_item = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#map-geocoding-list .list-group-item"))
        )
        list_item.click()
        print("已點擊結果列表中的項目。", flush=True)

        # 等待搜尋結果並截圖
        time.sleep(2)  # 等待搜尋結果
        filename_base = f"09_地質雲-{area}{section}-{lot_number}"
        screenshot_path = os.path.join(png_dir, f"{filename_base}.png")
        driver.save_screenshot(screenshot_path)
        print(f"\033[93m已截圖並保存為 {screenshot_path}\033[0m", flush=True)

        # 收集 PNG 路徑和標識符
        all_png_files.append(screenshot_path)
        all_identifiers.append(f"{area}{section}-{lot_number}")
        all_data_entries.append(data)

    except TimeoutException:
        print(f"處理座標 {coordinates} 時發生超時錯誤。", flush=True)

# 關閉瀏覽器
driver.quit()
print("查詢完成，關閉瀏覽器", flush=True)

# 合併所選 PNG 成為一個 PDF
def create_combined_pdf(png_files, identifiers, pdf_dir):
    if not png_files:
        print("沒有找到任何 PNG 檔案，無法生成 PDF。", flush=True)
        return

    # 使用新命名規則合併標識符
    area_section_dict = {}
    for identifier in identifiers:
        parts = identifier.split('-')
        region = parts[0]
        lot = '-'.join(parts[1:])  # 將後面的部分重新合併成地號

        if region in area_section_dict:
            area_section_dict[region].append(lot)
        else:
            area_section_dict[region] = [lot]

    # 建立 PDF 檔名，將每個區域的地號以「、」分隔，並用「+」連接不同地區地段
    pdf_filename = "09_地質雲-" + "+".join(
        [f"{region}-{','.join(lots)}" for region, lots in area_section_dict.items()]
    ) + ".pdf"

    pdf_path = os.path.join(pdf_dir, pdf_filename)

    # 🔥 檢查檔案是否已存在且被占用
    if os.path.exists(pdf_path):
        try:
            # 嘗試刪除舊檔案
            os.remove(pdf_path)
        except PermissionError:
            print(f"\033[91m警告：PDF 檔案正在被其他程式使用，請關閉該檔案後按 Enter 繼續...\033[0m", flush=True)
            print(f"檔案路徑：{pdf_path}", flush=True)
            input()
            # 再次嘗試刪除
            try:
                os.remove(pdf_path)
            except Exception as e:
                print(f"仍無法刪除舊檔案: {e}", flush=True)
                return None

    try:
        pdf = FPDF(orientation='L', unit='mm', format='A4')  # 橫向 PDF
        for png in png_files:
            pdf.add_page()
            pdf.image(png, x=0, y=0, w=297, h=210)  # 調整大小至 A4 橫向

        # 嘗試寫入 PDF，如果失敗則重試
        max_retries = 3
        for attempt in range(max_retries):
            try:
                pdf.output(pdf_path)
                print(f"\033[93mPDF 已經儲存至 {pdf_path}\033[0m", flush=True)
                return pdf_path
            except PermissionError as e:
                if attempt < max_retries - 1:
                    print(f"\033[91m寫入 PDF 失敗（嘗試 {attempt + 1}/{max_retries}），請關閉該檔案後按 Enter 重試...\033[0m", flush=True)
                    print(f"檔案路徑：{pdf_path}", flush=True)
                    input()
                else:
                    print(f"將 PNG 轉成 PDF 時發生錯誤: {e}", flush=True)
                    print(f"\033[91m提示：請檢查是否有 PDF 閱讀器正在開啟此檔案\033[0m", flush=True)
                    return None
    except Exception as e:
        print(f"將 PNG 轉成 PDF 時發生錯誤: {e}", flush=True)
        return None

# 詢問使用者要合併哪些地號並生成合併的 PDF
def ask_user_for_selection(all_data_entries, all_png_files, all_identifiers, basic_data_dir):
    # 如果只有一筆資料，直接生成 PDF
    if len(all_data_entries) == 1:
        print("\n只有一筆資料，直接生成 PDF...", flush=True)
        return create_combined_pdf(all_png_files, all_identifiers, basic_data_dir)

    # 多筆資料時才詢問使用者
    print("\n請選擇要合併為 PDF 的地號：", flush=True)
    print("0. 全部選取", flush=True)
    
    # 預設選擇最後一個
    default_selection = len(all_data_entries) - 1
    
    # 顯示所有可選項目
    for i, data in enumerate(all_data_entries):
        area = data['area']
        section = data['section']
        lot_number = data['lot_number']
        default_mark = " (預設)" if i == default_selection else ""
        print(f"{i+1}. {area}{section}-{lot_number}{default_mark}", flush=True)

    print("\n請輸入您要選擇的項目編號（多個選項以逗號分隔，例如：1,3,5），或直接按 Enter 使用預設選擇：", flush=True)
    selection = input().strip()
    
    selected_indices = []
    
    if not selection:  # 使用者直接按 Enter，選擇預設項目
        selected_indices = [default_selection]
        print(f"使用預設選擇：{all_data_entries[default_selection]['area']}{all_data_entries[default_selection]['section']}-{all_data_entries[default_selection]['lot_number']}", flush=True)
    elif selection == "0":  # 全部選擇
        selected_indices = list(range(len(all_data_entries)))
        print("已選擇所有項目", flush=True)
    else:
        try:
            # 解析用戶輸入，將編號轉換為索引（減 1）
            selected_indices = [int(idx.strip()) - 1 for idx in selection.split(',')]
            # 檢查索引是否有效
            for idx in selected_indices:
                if idx < 0 or idx >= len(all_data_entries):
                    print(f"警告：編號 {idx+1} 超出範圍，已忽略")
                    selected_indices.remove(idx)
            
            if not selected_indices:  # 如果沒有有效選擇，則使用預設
                selected_indices = [default_selection]
                print(f"沒有有效選擇，使用預設選擇：{all_data_entries[default_selection]['area']}{all_data_entries[default_selection]['section']}-{all_data_entries[default_selection]['lot_number']}")
            else:
                print(f"已選擇 {len(selected_indices)} 個項目")
        except ValueError:
            # 輸入格式錯誤，使用預設選擇
            selected_indices = [default_selection]
            print(f"輸入格式不正確，使用預設選擇：{all_data_entries[default_selection]['area']}{all_data_entries[default_selection]['section']}-{all_data_entries[default_selection]['lot_number']}")
    
    # 獲取選定項目的 PNG 和標識符
    selected_png_files = [all_png_files[i] for i in selected_indices]
    selected_identifiers = [all_identifiers[i] for i in selected_indices]
    
    # 生成 PDF
    return create_combined_pdf(selected_png_files, selected_identifiers, basic_data_dir)

# 呼叫詢問使用者並生成合併的 PDF
pdf_path = ask_user_for_selection(all_data_entries, all_png_files, all_identifiers, basic_data_dir)

def notify_main_program():
    print("地質雲已完成執行", flush=True)
    # if pdf_path:
        # print(f"生成的 PDF 文件路徑：{pdf_path}", flush=True)

if __name__ == "__main__":
    try:
        # 子程式的主要邏輯
        # 在這裡執行子程式的操作，例如表單填寫和資料處理
        # print("執行地質雲主要邏輯", flush=True)
        pass
    except Exception as e:
        print(f"地質雲執行過程中發生錯誤: {e}", flush=True)
    finally:
        # 在結尾通知主程式
        notify_main_program()