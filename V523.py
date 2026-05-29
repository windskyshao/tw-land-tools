#V523地產資訊網
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
import re
import shutil
import json
import time
import keyboard  # 用於系統級鍵盤事件
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window

# 🔥 基準目錄設定（data.json 和工作資料夾的位置）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

# 清理畫面並顯示歡迎訊息
os.system('cls')
sys.stdout.reconfigure(line_buffering=True)
print("歡迎使用【V523地產資訊網】自動化小程式，模組載入中...", flush=True)



# 設定（根據縣市動態取出帳號密碼，從 Windows 認證管理員）
# 帳密已不再寫在程式碼中，由 credential_helper 處理
USERNAME = None  # 將在 test() 中根據縣市從 keyring 取出
PASSWORD = None
current_dir = os.getcwd()

# global driver 變數
driver: webdriver.Chrome
current_window_handle: str

# 定義 TMP_PDF_PATH
TMP_PDF_PATH = os.path.join(os.getcwd(), "tmp_pdf")

# 確保 TMP_PDF_PATH 資料夾存在
if not os.path.exists(TMP_PDF_PATH):
    os.makedirs(TMP_PDF_PATH)

# 🔥 讀取視窗配置
print("\n" + "="*60, flush=True)
print("🔍 V523 視窗配置偵測", flush=True)
print("="*60, flush=True)

try:
    config_path = os.path.join(BASE_DIR, 'window_config.json')
    print(f"[配置] 檢查配置檔: {config_path}", flush=True)

    if os.path.exists(config_path):
        print("[配置] ✓ 找到 window_config.json", flush=True)
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        dpi_scale = config.get('dpi_scale', 1.0)
        work_area = config['work_area']
        main_window = config['main_window']

        print(f"[配置] DPI 縮放: {dpi_scale}x (即 {int(dpi_scale * 100)}%)", flush=True)
        print(f"[配置] 螢幕尺寸: {config.get('screen_width')}x{config.get('screen_height')}", flush=True)

        work_left = work_area['left']
        work_top = work_area['top']
        work_right = work_area['right']
        work_bottom = work_area['bottom']

        main_x = main_window['x']
        main_y = main_window['y']
        main_w = main_window['width']
        main_h = main_window['height']

        # 🔥 計算 Chrome 視窗大小：填滿剩餘空間（與其他子程式一致）
        screen_width_physical = work_right - work_left

        print(f"[視窗計算] 螢幕寬度: {screen_width_physical}px, DPI縮放: {dpi_scale}x", flush=True)
        print(f"[視窗計算] 主程式位置: x={main_x}, 寬度={main_w}", flush=True)

        # 根據主程式位置決定 Chrome 視窗位置和寬度
        screen_center_x = work_left + (work_right - work_left) // 2

        if main_x < screen_center_x:
            # 主程式在左側，Chrome 放右側，填滿剩餘空間
            chrome_x = main_x + main_w
            chrome_width = work_right - chrome_x
            print(f"[視窗計算] 主程式在左側 → Chrome放右側填滿剩餘空間", flush=True)
            print(f"[視窗計算] Chrome起點: x={chrome_x}, 寬度={chrome_width}px", flush=True)
        else:
            # 主程式在右側，Chrome 放左側，填滿剩餘空間
            chrome_x = work_left
            chrome_width = main_x - work_left
            print(f"[視窗計算] 主程式在右側 → Chrome放左側填滿剩餘空間", flush=True)
            print(f"[視窗計算] Chrome起點: x={chrome_x}, 寬度={chrome_width}px", flush=True)

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

        print(f"[視窗計算] 邏輯像素: 寬度={chrome_width_logical}, 高度={chrome_height_logical}", flush=True)
        print(f"[視窗計算] 邏輯位置: x={chrome_x_logical}, y={chrome_y_logical}", flush=True)

        v523_window_width = chrome_width_logical
        v523_window_height = chrome_height_logical
        v523_window_x = chrome_x_logical
        v523_window_y = chrome_y_logical

        print("="*60, flush=True)
        print(f"✓ 視窗配置完成: {chrome_width_logical}x{chrome_height_logical} @ ({chrome_x_logical}, {chrome_y_logical})", flush=True)
        print("="*60 + "\n", flush=True)
    else:
        # 如果沒有配置檔，使用預設值
        print("[配置] ✗ window_config.json 不存在", flush=True)
        print("[配置] 使用預設視窗大小: 1024x1024", flush=True)
        print("="*60 + "\n", flush=True)
        v523_window_width = 1024
        v523_window_height = 1024
        v523_window_x = 0
        v523_window_y = 0
except Exception as e:
    # 出錯時使用預設值
    print(f"[配置] ✗ 讀取配置檔時發生錯誤: {e}", flush=True)
    print("[配置] 使用預設視窗大小: 1024x1024", flush=True)
    print("="*60 + "\n", flush=True)
    v523_window_width = 1024
    v523_window_height = 1024
    v523_window_x = 0
    v523_window_y = 0

def make_directory_structure(city, section, lot_number):
    folder_name = f"{city}{section}-{lot_number}"
    # 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
    folder_path = os.path.join(get_work_folder(folder_name), "1.基本資料")
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def init_driver():
    import tempfile
    import random

    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")

    # 🔥 不在啟動參數設定視窗大小，改為啟動後再調整（避免 DPI 改變時崩潰）
    # options.add_argument("--window-size=1024,1024")
    # options.add_argument("--window-position=0,0")
    options.add_argument("--log-level=3")  # 設定日誌級別為錯誤

    # 🔥 使用臨時的 user-data-dir，完全避免密碼彈窗（與 land.py 相同）
    user_data_dir = os.path.join(tempfile.gettempdir(), f"chrome_v523_{int(time.time())}_{random.randint(1000, 9999)}")
    options.add_argument(f"--user-data-dir={user_data_dir}")

    # 🔥 完全禁用密碼管理器彈窗
    options.add_argument("--disable-save-password-bubble")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--remote-debugging-port=0")
    options.add_argument("--disable-infobars")

    # 🔥 忽略 SSL 憑證錯誤（V523 網站憑證可能過期或無效）
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors")

    # 合併原有設定和新增的密碼管理設定
    prefs = {
        "printing.print_preview_sticky_settings.appState": json.dumps({
            "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2
        }),
        "savefile.default_directory": TMP_PDF_PATH,  # 自動儲存的路徑
        "download.prompt_for_download": False,       # 禁止下載時提示

        # 禁用密碼儲存和自動填入
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "profile.default_content_setting_values.notifications": 2,  # 禁用通知
        "profile.default_content_settings.popups": 0,  # 允許彈窗（但我們會用其他方式控制）
        "autofill.profile_enabled": False,  # 禁用自動填入

        # 🔥 禁用 Google 密碼外洩檢查（防止「變更你的密碼」彈窗）
        "profile.password_manager_leak_detection": False,
        "safebrowsing.enabled": False,  # 禁用安全瀏覽（包含密碼外洩檢查）

        # 🔥 額外加強：完全禁用密碼管理相關功能
        "profile.password_manager_enabled": False,
        "credentials_enable_service": False,
        "password_manager_enabled": False,
        "profile.content_settings.exceptions.auto_select_certificate": {},
        "profile.default_content_setting_values.automatic_downloads": 1
    }

    options.add_experimental_option("prefs", prefs)
    # 🔥 移除 --kiosk-printing 以避免直接送到實體印表機
    # options.add_argument("--kiosk-printing")  # 此參數會導致直接列印而非另存 PDF
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    options.add_experimental_option('useAutomationExtension', False)

    global driver
    driver = create_chrome_driver(options=options)

    # 🔥 啟動後立即設定視窗大小和位置
    try:
        driver.set_window_size(v523_window_width, v523_window_height)
        driver.set_window_position(v523_window_x, v523_window_y)
    except Exception as e:
        print(f"[WARNING] 視窗設定失敗: {e}，繼續執行")

    # 🔥 設定頁面載入超時時間（60秒）和隱式等待（10秒）
    driver.set_page_load_timeout(60)
    driver.implicitly_wait(5)

def load_popup_coordinates():
    """載入已儲存的彈窗座標"""
    coords_file = os.path.join(current_dir, 'google_popup_coords.json')
    if os.path.exists(coords_file):
        try:
            with open(coords_file, 'r', encoding='utf-8') as f:
                coords = json.load(f)
            # print(f"[載入] 已儲存的彈窗座標: {coords}", flush=True)
            return coords
        except Exception as e:
            print(f"[錯誤] 載入座標檔案失敗: {e}", flush=True)
    return None

def save_popup_coordinates(x, y):
    """儲存彈窗座標"""
    coords_file = os.path.join(current_dir, 'google_popup_coords.json')
    try:
        coords = {'x': x, 'y': y}
        with open(coords_file, 'w', encoding='utf-8') as f:
            json.dump(coords, f, ensure_ascii=False, indent=2)
        print(f"[儲存] 新座標已記錄: ({x}, {y})", flush=True)
        return True
    except Exception as e:
        print(f"[錯誤] 儲存座標失敗: {e}", flush=True)
        return False

def handle_google_security_popup():
    """智慧型處理 Google 安全彈窗"""
    import pyautogui
    import threading

    time.sleep(0.5)  # 🔥 縮短等待時間從 2 秒到 0.5 秒

    # 先嘗試使用已儲存的座標自動點擊
    saved_coords = load_popup_coordinates()

    if saved_coords:
        # print("=" * 60, flush=True)
        # print("🤖 嘗試使用已記錄的座標自動關閉彈窗...", flush=True)
        # print(f"   座標: ({saved_coords['x']}, {saved_coords['y']})", flush=True)
        # print("=" * 60, flush=True)

        try:
            pyautogui.click(saved_coords['x'], saved_coords['y'])
            time.sleep(0.3)  # 🔥 縮短等待時間從 1 秒到 0.3 秒

            # 檢查是否成功（簡單方式：等待一下看看）
            # print("✓ 自動點擊完成", flush=True)
            time.sleep(0.5)  # 🔥 縮短等待時間從 2 秒到 0.5 秒
            return
        except Exception as e:
            print(f"❌ 自動點擊失敗: {e}", flush=True)
            print("   將切換到手動模式...", flush=True)

    # 2. 如果沒有儲存的座標，或自動點擊失敗，進入手動模式
    print("\n" + "=" * 60, flush=True)
    print("⚠️  手動處理模式", flush=True)
    print("=" * 60, flush=True)
    print("如果出現 Google 帳號安全檢查彈窗：", flush=True)
    print("  1. 請點擊【不用了】或【關閉】按鈕", flush=True)
    print("  2. 點擊後，請將滑鼠停留在該按鈕上不要移動", flush=True)
    print("  3. 然後按鍵盤的 Enter 鍵", flush=True)
    print("     → 程式會自動記錄該座標，下次就能自動點擊！", flush=True)
    print("\n如果沒有彈窗，直接按 Enter 繼續...", flush=True)
    print("=" * 60, flush=True)

    # 建立一個標記來記錄是否按下 Enter
    enter_pressed = threading.Event()

    def wait_for_enter():
        input()
        enter_pressed.set()

    # 啟動等待 Enter 的執行緒
    enter_thread = threading.Thread(target=wait_for_enter)
    enter_thread.daemon = True
    enter_thread.start()

    # 等待使用者按下 Enter
    enter_thread.join()

    # 獲取滑鼠當前位置
    try:
        mouse_x, mouse_y = pyautogui.position()

        # 判斷滑鼠是否移動過（如果座標不是預設位置，就記錄）
        # 假設使用者有移動到按鈕上
        if mouse_x > 100 or mouse_y > 100:  # 簡單判斷不是在螢幕左上角
            print(f"\n📍 偵測到滑鼠位置: ({mouse_x}, {mouse_y})", flush=True)
            print("要儲存此座標供下次使用嗎？(y/n): ", flush=True, end='')

            save_choice = input().strip().lower()
            if save_choice == 'y' or save_choice == '':
                if save_popup_coordinates(mouse_x, mouse_y):
                    print("✓ 座標已儲存！下次執行時會自動點擊此位置", flush=True)
            else:
                print("⊗ 座標未儲存", flush=True)
        else:
            print("ℹ️  未偵測到有效的滑鼠位置，繼續執行...", flush=True)
    except Exception as e:
        print(f"[錯誤] 獲取滑鼠位置失敗: {e}", flush=True)

    print("=" * 60, flush=True)
    time.sleep(1)

def check_network_connectivity():
    """檢查網路連線和 V523 網站狀態"""
    import socket
    import urllib.request

    print("\n" + "="*60, flush=True)
    print("🔍 網路連線診斷", flush=True)
    print("="*60, flush=True)

    # 1. 檢查基本網路連線
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        print("✓ 網路連線正常", flush=True)
    except OSError:
        print("✗ 無法連接網路，請檢查網路連線", flush=True)
        return False

    # 2. 檢查 DNS 解析
    try:
        socket.gethostbyname("www.v523.com.tw")
        print("✓ DNS 解析正常", flush=True)
    except socket.gaierror:
        print("✗ DNS 解析失敗，無法解析 www.v523.com.tw", flush=True)
        return False

    # 3. 嘗試連接 V523 網站
    try:
        socket.create_connection(("www.v523.com.tw", 80), timeout=10)
        print("✓ V523 網站可連接", flush=True)
    except (socket.timeout, OSError) as e:
        print(f"✗ 無法連接到 V523 網站: {e}", flush=True)
        print("   可能原因：", flush=True)
        print("   1. 網站維護中", flush=True)
        print("   2. 網站伺服器暫時無法訪問", flush=True)
        print("   3. 防火牆阻擋", flush=True)
        print("   4. ISP 封鎖", flush=True)
        return False

    print("="*60, flush=True)
    print("✓ 網路診斷完成，可以嘗試連線\n", flush=True)
    return True

def login():
    # 🔥 先進行網路連線診斷
    if not check_network_connectivity():
        print("\n❌ 網路連線診斷失敗，請檢查網路後重新執行", flush=True)
        print("建議：", flush=True)
        print("  1. 檢查您的網路連線", flush=True)
        print("  2. 嘗試用瀏覽器手動訪問 http://www.v523.com.tw/", flush=True)
        print("  3. 如果手動也無法訪問，可能網站維護中，請稍後再試", flush=True)
        raise Exception("網路連線診斷失敗")

    # 🔥 增加重試機制，最多嘗試 3 次
    max_retries = 3
    for attempt in range(max_retries):
        try:
            print(f"正在連線到 V523 網站... (嘗試 {attempt + 1}/{max_retries})", flush=True)
            driver.get("http://www.v523.com.tw/")
            print("✓ 網站載入成功", flush=True)
            # 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
            verify_and_fix_chrome_window(driver)
            time.sleep(1)  # 🔥 縮短等待時間，加快登入速度
            break
        except Exception as e:
            print(f"✗ 連線失敗: {e}", flush=True)
            if attempt < max_retries - 1:
                print(f"等待 5 秒後重試...", flush=True)
                time.sleep(5)
            else:
                print("已達最大重試次數，連線失敗", flush=True)
                print("\n建議檢查項目：", flush=True)
                print("  1. 網站是否正在維護", flush=True)
                print("  2. 您的 IP 是否被限制", flush=True)
                print("  3. 稍後再試", flush=True)
                raise Exception(f"無法連線到 V523 網站，已重試 {max_retries} 次")

    # 切換到 iframe
    try:
        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[2]/div[1]/iframe"))
        )
        driver.switch_to.frame(iframe)
        # print("✓ 已切換到登入 iframe", flush=True)
    except TimeoutException:
        print("✗ 無法找到 iframe，請檢查頁面結構", flush=True)
        return

    # 輸入帳號和密碼
    try:
        username = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "username"))
        )
        username.clear()  # 清空輸入框
        username.send_keys(USERNAME)
        # print("已成功輸入帳號", flush=True)

        password = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "password2"))
        )
        password.clear()  # 清空密碼框
        password.send_keys(PASSWORD)
        # print("已成功輸入密碼", flush=True)

        login_button = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.ID, "but"))
        )
        login_button.click()
        # print("已成功點擊登入按鈕", flush=True)
        time.sleep(0.5)  # 等待登入完成

        # 回到主框架
        driver.switch_to.default_content()

        # 🔥 登入完成後稍作等待（後續由 open_query_page() 負責導航和頁面設定）
        time.sleep(0.5)

    except TimeoutException:
        print("登入頁面的元素無法找到，請檢查 ID 或頁面結構是否有變更", flush=True)

def open_query_page():
    """
    導航到地籍查詢頁面，並根據 DPI 和螢幕高度自動設定頁面縮放
    """
    # 🔥 導航到地籍查詢頁面（避免停留在登入後的首頁誤觸廣告）
    driver.get("http://www.v523.com.tw/landform-new/index.jsp")
    print("✓ 已導航到地籍分區頁面", flush=True)

    # 🔥 智慧型處理 Google 安全彈窗
    handle_google_security_popup()

    # 🔥 等待頁面完全載入
    time.sleep(2)

    # 🔥 自動判斷是否需要縮放頁面（解決高 DPI 或小螢幕下的排版問題）
    # 根據 DPI 決定縮放次數：150% → 2次(80%)，175% → 3次(75%)
    zoom_count = 0
    try:
        config_path = os.path.join(BASE_DIR, 'window_config.json')
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)

            current_dpi = config.get('dpi_scale', 1.0)
            screen_height = config.get('screen_height', 1080)

            # 🔥 根據 DPI 決定縮放次數
            if current_dpi >= 1.75:
                zoom_count = 4  # 175% 或更高：縮放到 67%（按 4 次）
            elif current_dpi >= 1.5:
                zoom_count = 2  # 150%：縮放到 80%（按 2 次）
            else:
                zoom_count = 0  # 100%、125%：不縮放

    except Exception:
        pass  # 偵測配置失敗時靜默處理

    if zoom_count > 0:
        try:
            # 🔥 使用 keyboard 庫發送系統級的 Ctrl + - 按鍵（真正的瀏覽器縮放）
            # 先點擊頁面確保 Chrome 視窗有焦點
            body = driver.find_element(By.TAG_NAME, 'body')
            body.click()
            time.sleep(0.5)

            # 🔥 根據 DPI 決定按幾次 Ctrl + -
            for _ in range(zoom_count):
                keyboard.press_and_release('ctrl+-')  # 系統級按鍵
                time.sleep(0.5)
        except Exception:
            pass  # 靜默處理縮放錯誤

    return zoom_count > 0  # 返回是否設定了縮放，供 finally 區塊使用

def select_combo_option(select_id, target_text):
    """
    🔥 通用函數：選擇 combo-select 自訂下拉選單的選項
    V523 網站使用自訂的 combo-select 元件，需要：
    1. 點擊 combo-select 容器展開下拉選單
    2. 點擊對應的 option-item
    """
    try:
        # 找到 select 元素的父容器 (combo-select)
        combo_container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f"#{select_id}"))
        ).find_element(By.XPATH, "./ancestor::div[contains(@class,'combo-select')]")

        # 🔥 先點擊 combo-input 或 combo-arrow 來展開下拉選單
        try:
            # 嘗試點擊輸入框或箭頭
            click_target = combo_container.find_element(By.CSS_SELECTOR, ".combo-input, .combo-arrow")
            driver.execute_script("arguments[0].click();", click_target)
        except:
            # 如果找不到，直接點擊容器
            driver.execute_script("arguments[0].click();", combo_container)

        time.sleep(0.8)  # 等待下拉選單展開

        # 找到並點擊對應的 option-item
        dropdown = combo_container.find_element(By.CSS_SELECTOR, "ul.combo-dropdown")

        # 🔥 等待下拉選單可見
        WebDriverWait(driver, 5).until(
            EC.visibility_of(dropdown)
        )

        options = dropdown.find_elements(By.CSS_SELECTOR, "li.option-item")

        # 🔥 Debug: 顯示找到的選項
        print(f"[DEBUG] 找到 {len(options)} 個選項", flush=True)

        for option in options:
            option_text = option.text.strip()
            if option_text == target_text:
                # 🔥 確保選項可見後再點擊
                driver.execute_script("arguments[0].scrollIntoView(true);", option)
                time.sleep(0.2)
                driver.execute_script("arguments[0].click();", option)
                time.sleep(0.5)
                return True

        # 🔥 如果精確匹配失敗，顯示可用選項供 debug
        available_options = [opt.text.strip() for opt in options[:5]]
        print(f"[WARNING] 找不到選項: {target_text}", flush=True)
        print(f"[DEBUG] 可用選項(前5個): {available_options}", flush=True)
        return False

    except Exception as e:
        # 如果自訂下拉選單失敗，嘗試直接操作原生 select
        print(f"[INFO] combo-select 失敗，嘗試直接操作 select: {e}", flush=True)
        try:
            select_element = driver.find_element(By.ID, select_id)

            # 🔥 使用 JavaScript 直接選擇並觸發相關事件
            result = driver.execute_script("""
                var select = arguments[0];
                var targetText = arguments[1];
                var found = false;
                for (var i = 0; i < select.options.length; i++) {
                    if (select.options[i].text === targetText) {
                        select.selectedIndex = i;
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                        select.dispatchEvent(new Event('input', { bubbles: true }));
                        found = true;
                        break;
                    }
                }
                return found;
            """, select_element, target_text)

            if result:
                print(f"[INFO] 直接操作 select 成功", flush=True)
                time.sleep(0.5)
                return True
            else:
                print(f"[WARNING] 在 select 中也找不到選項: {target_text}", flush=True)
                return False

        except Exception as e2:
            print(f"[ERROR] 直接操作 select 也失敗: {e2}", flush=True)
            return False

def fill_form_from_data(city, area, section, lot_number):
    try:
        print("開始處理資料: 城市 = {}, 區域 = {}, 段 = {}, 地號 = {}".format(city, area, section, lot_number), flush=True)

        # 🔥 選擇縣市（使用新的 combo-select 方式）
        try:
            if select_combo_option("county", city):
                print(f"已成功選擇縣市: {city}", flush=True)
                time.sleep(1.5)  # 等待鄉鎮市區選項載入
            else:
                print(f"選擇縣市失敗: {city}", flush=True)
                return
        except Exception as e:
            print(f"選擇縣市時發生錯誤: {e}", flush=True)
            return

        # 🔥 選擇鄉鎮市區
        try:
            if select_combo_option("borough", area):
                print(f"已成功選擇鄉鎮市區: {area}", flush=True)
                time.sleep(1.5)  # 等待地段選項載入
            else:
                print(f"選擇鄉鎮市區失敗: {area}", flush=True)
                return
        except Exception as e:
            print(f"選擇鄉鎮市區時發生錯誤: {e}", flush=True)
            return

        # 選擇地段和小段
        # print(f"開始處理地段: {section}", flush=True)
        if "小段" in section:
            match = re.match(r"(.*段)(.*小段)", section)
            if match:
                main_section = match.group(1)
                sub_section = match.group(2)
                print(f"已解析出地段: {main_section}, 小段: {sub_section}", flush=True)
                select_section(main_section)
                select_sub_section(sub_section)
            else:
                print(f"無法解析地段和小段: {section}", flush=True)
        else:
            select_section(section)

        # 輸入地號
        try:
            # print(f"嘗試輸入地號: {lot_number}", flush=True)
            lot_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "name"))
            )
            lot_input.clear()
            lot_input.send_keys(lot_number)
            print(f"已輸入地號: {lot_number}", flush=True)
        except Exception as e:
            print(f"輸入地號時發生錯誤: {e}", flush=True)
            return

    except Exception as e:
        print(f"填寫表單過程中發生未知錯誤: {e}", flush=True)

def select_section(section):
    """選擇地段（使用 combo-select 方式）"""
    try:
        # 🔥 使用最佳匹配邏輯找到正確的地段名稱
        section_select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "segment"))
        )
        options = section_select_element.find_elements(By.TAG_NAME, "option")
        best_option = find_best_matching_section(section, options)

        if best_option:
            best_option_text = best_option.text
            # 使用 combo-select 方式選擇
            if select_combo_option("segment", best_option_text):
                print(f"已成功選擇地段: {best_option_text}", flush=True)
                time.sleep(1)  # 等待小段選項載入
            else:
                print(f"選擇地段失敗: {best_option_text}", flush=True)
        else:
            print(f"找不到匹配的地段: {section}", flush=True)

    except Exception as e:
        print(f"選擇地段時發生錯誤: {e}", flush=True)

def select_sub_section(sub_section):
    """選擇小段（使用 combo-select 方式）"""
    try:
        print(f"嘗試選擇小段: {sub_section}", flush=True)
        # 🔥 使用 combo-select 方式選擇
        if select_combo_option("subSegment", sub_section):
            print(f"已成功選擇小段: {sub_section}", flush=True)
            time.sleep(0.5)
        else:
            print(f"選擇小段失敗: {sub_section}", flush=True)
    except Exception as e:
        print(f"在選擇小段時發生錯誤: {e}", flush=True)


def find_best_matching_section(target_section, options):
    """
    找到最佳匹配的段或小段選項
    """
    best_option = None
    highest_similarity = 0

    for option in options:
        option_text = option.text.strip()  # 取得選項文字
        option_value = option.get_attribute("data-value")  # 取得 data-value 屬性

        # 直接比較文字是否完全相符
        if target_section == option_text:
            return option

        # 如果 target_section 為數值，檢查是否與 data-value 相符
        if target_section == option_value:
            return option

        # 進行部分匹配計算
        similarity = max(
            len([c for c in zip(target_section, option_text) if c[0] == c[1]]),
            len([c for c in zip(target_section[::-1], option_text[::-1]) if c[0] == c[1]])
        )

        if similarity > highest_similarity:
            highest_similarity = similarity
            best_option = option

    return best_option


def query_location():
    try:
        print("[DEBUG] 開始尋找【位置分區】按鈕...", flush=True)

        # 🔥 使用多種方式尋找「位置分區查詢」按鈕
        location_button = None

        # 方法1: 使用 onclick 屬性尋找（找到包含 onclick 的 div）
        try:
            location_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(@onclick,'funPosition')]"))
            )
            print("[DEBUG] ✓ 方法1: 透過 onclick 找到按鈕", flush=True)
        except:
            pass

        # 方法2: 使用文字內容尋找
        if not location_button:
            try:
                location_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//font[contains(text(),'位置分區')]"))
                )
                print("[DEBUG] ✓ 方法2: 透過文字找到按鈕", flush=True)
            except:
                pass

        # 方法3: 使用原本的絕對 XPath（備用）
        if not location_button:
            try:
                location_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[1]/div/div[1]/div/div/div[11]/div/div[1]/div/font"))
                )
                print("[DEBUG] ✓ 方法3: 透過絕對 XPath 找到按鈕", flush=True)
            except:
                pass

        if not location_button:
            print("[DEBUG] ❌ 所有方法都找不到【位置分區】按鈕", flush=True)
            return

        # 滾動到按鈕可見
        driver.execute_script("arguments[0].scrollIntoView(true);", location_button)
        time.sleep(0.5)  # 等待滾動效果

        # 🔥 直接使用 JavaScript 執行 funPosition() 函數，確保真正觸發查詢
        print("[DEBUG] 嘗試直接執行 funPosition() 函數...", flush=True)
        try:
            driver.execute_script("funPosition();")
            print("【位置分區】透過 JS 函數觸發成功", flush=True)
        except Exception as js_err:
            print(f"[DEBUG] JS 函數執行失敗: {js_err}，改用點擊方式", flush=True)
            driver.execute_script("arguments[0].click();", location_button)
            print("【位置分區】點擊成功", flush=True)

        # 🔥 等待彈窗出現（重要！）
        time.sleep(2)
        print("[DEBUG] 已等待彈窗載入", flush=True)

    except TimeoutException:
        print("[DEBUG] ❌ 查詢按鈕未出現（超時）", flush=True)
    except Exception as e:
        print(f"[DEBUG] ❌ 定位【位置分區】按鈕時發生錯誤: {e}", flush=True)
        import traceback
        traceback.print_exc()

def press_ok_button():
    try:
        # 🔥 額外等待一下，確保彈窗完全渲染
        time.sleep(1)
        print("[DEBUG] 開始尋找確定按鈕...", flush=True)

        # 🔥 使用新的 zbox-popup 結構尋找彈窗
        popup = None
        try:
            # 方法1: 尋找 zbox-popup 彈窗
            popup = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".zbox-popup.zbox-popup-in"))
            )
            print("[DEBUG] ✓ 找到 zbox-popup 彈窗元素", flush=True)
        except:
            # 方法2: 備用 - 使用舊的 div[9] 結構
            try:
                popup = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/div[9]"))
                )
                print("[DEBUG] ✓ 找到 div[9] 彈窗元素", flush=True)
            except:
                print("[DEBUG] ✗ 找不到彈窗元素", flush=True)
                raise TimeoutException("找不到彈窗")

        # 🔥 點擊第一個確定按鈕（查詢確認）
        ok_button = None
        try:
            # 方法1: 使用 zbox-popup-button 類別
            ok_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".zbox-popup-button[index='0']"))
            )
            print("[DEBUG] ✓ 方法1: 找到確定按鈕 (zbox-popup-button)", flush=True)
        except:
            # 方法2: 使用文字尋找
            try:
                ok_button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'zbox-popup-button') and text()='確定']"))
                )
                print("[DEBUG] ✓ 方法2: 透過文字找到確定按鈕", flush=True)
            except:
                # 方法3: 備用舊方式
                ok_button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[9]/div/div[2]/span[1]"))
                )
                print("[DEBUG] ✓ 方法3: 透過舊 XPath 找到確定按鈕", flush=True)

        driver.execute_script("arguments[0].click();", ok_button)
        print("點擊【查詢確認】按鈕", flush=True)
        time.sleep(2)

        # 🔥 檢查是否有第二個彈窗（是否過期確認）
        try:
            second_popup = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".zbox-popup.zbox-popup-in"))
            )
            second_ok_button = second_popup.find_element(By.CSS_SELECTOR, ".zbox-popup-button[index='0']")
            driver.execute_script("arguments[0].click();", second_ok_button)
            print("點擊【是否過期確定】按鈕", flush=True)
        except:
            print("[DEBUG] 沒有第二個彈窗，繼續執行", flush=True)

        # 🔥 等待查詢完成（等待地圖上出現土地標記或載入完成）
        print("[DEBUG] 等待查詢結果載入...", flush=True)
        time.sleep(3)  # 先等待一下

        # 🔥 檢查是否有查詢結果（土地資訊面板或地圖標記）
        try:
            # 等待土地資訊面板出現，或者等待 loading 消失
            WebDriverWait(driver, 10).until(
                lambda d: d.execute_script("return typeof(landInfo) !== 'undefined' && landInfo !== null") or
                          len(d.find_elements(By.CSS_SELECTOR, ".land-info, .gm-style")) > 0
            )
            print("[DEBUG] ✓ 查詢結果已載入", flush=True)
        except:
            print("[DEBUG] 等待查詢結果超時，繼續執行", flush=True)

        time.sleep(2)  # 額外等待確保完全載入

        # 🔥 檢查視窗數量作為調試資訊
        windows_count = len(driver.window_handles)
        print(f"[DEBUG] 目前視窗數量: {windows_count}", flush=True)

    except TimeoutException:
        print("[DEBUG] ✗ 等待確定按鈕超時", flush=True)
        print("確定按鈕未出現", flush=True)
    except Exception as e:
        print(f"[DEBUG] ✗ 處理確定按鈕時發生錯誤: {e}", flush=True)
        print("確定按鈕處理失敗", flush=True)

def get_pdf_page():
    try:
        print("[DEBUG] 開始尋找【分區列印】按鈕...", flush=True)

        # 🔥 使用多種方式尋找「分區列印」按鈕
        pdf_button = None

        # 方法1: 使用 onclick 屬性尋找 (funPrintBasic(2) 是分區列印)
        try:
            pdf_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(@onclick,'funPrintBasic(2)')]"))
            )
            print("[DEBUG] ✓ 方法1: 透過 onclick funPrintBasic(2) 找到按鈕", flush=True)
        except:
            pass

        # 方法2: 使用文字內容尋找
        if not pdf_button:
            try:
                pdf_button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "//font[contains(text(),'分區列印')]"))
                )
                print("[DEBUG] ✓ 方法2: 透過文字找到按鈕", flush=True)
            except:
                pass

        # 方法3: 在 searchPanel 內尋找（舊方式備用）
        if not pdf_button:
            try:
                search_panel = driver.find_element(By.ID, "searchPanel")
                driver.execute_script("arguments[0].scrollTop = 0;", search_panel)
                time.sleep(0.3)

                scroll_step = 100
                max_scroll_height = search_panel.get_attribute("scrollHeight")
                current_scroll_position = 0
                print(f"[DEBUG] searchPanel 最大高度: {max_scroll_height}", flush=True)

                while current_scroll_position < int(max_scroll_height):
                    try:
                        pdf_button = driver.find_element(By.XPATH, "//font[contains(text(),'分區列印')]")
                        if pdf_button.is_displayed():
                            print(f"[DEBUG] ✓ 方法3: 在滾動位置 {current_scroll_position} 找到按鈕", flush=True)
                            break
                    except:
                        current_scroll_position += scroll_step
                        driver.execute_script("arguments[0].scrollTop += arguments[1];", search_panel, scroll_step)
                        time.sleep(0.2)
                        pdf_button = None
            except:
                pass

        if not pdf_button:
            print("[DEBUG] ❌ 所有方法都找不到【分區列印】按鈕", flush=True)
            return

        # 滾動到按鈕可見
        driver.execute_script("arguments[0].scrollIntoView(true);", pdf_button)
        time.sleep(0.5)

        # 🔥 直接使用 JavaScript 執行 funPrintBasic(2) 函數
        print("[DEBUG] 嘗試直接執行 funPrintBasic(2) 函數...", flush=True)
        try:
            driver.execute_script("funPrintBasic(2);")
            print("【分區列印】透過 JS 函數觸發成功", flush=True)
        except Exception as js_err:
            print(f"[DEBUG] JS 函數執行失敗: {js_err}，改用點擊方式", flush=True)
            driver.execute_script("arguments[0].click();", pdf_button)
            print("【分區列印】點擊成功", flush=True)

        time.sleep(1)  # 等待新視窗開啟（增加等待時間）

        # 🔥 檢查是否有錯誤彈窗
        try:
            error_popup = driver.find_element(By.CSS_SELECTOR, ".zbox-popup.zbox-popup-in")
            error_text = error_popup.text
            if "請先" in error_text or "查詢" in error_text:
                print(f"[DEBUG] ⚠️ 出現錯誤提示: {error_text}", flush=True)
                # 關閉錯誤彈窗
                close_btn = error_popup.find_element(By.CSS_SELECTOR, ".zbox-popup-button")
                driver.execute_script("arguments[0].click();", close_btn)
        except:
            pass  # 沒有錯誤彈窗是好事

    except Exception as e:
        print(f"[DEBUG] ❌ 操作過程中發生錯誤: {e}", flush=True)
        print(f"操作過程中發生錯誤: {e}", flush=True)
        import traceback
        traceback.print_exc()

def extract_v523_data():
    """
    從 V523 分區列印頁面抓取資料
    回傳字典包含：土地位置資訊、土地使用分區資訊、土地相關資訊、面寬資訊
    """
    try:
        data = {}

        # 使用 JavaScript 抓取表格資料（處理左右兩欄的結構）
        table_data = driver.execute_script("""
            var result = {};
            var tables = document.querySelectorAll('table');

            tables.forEach(function(table, index) {
                var rows = table.querySelectorAll('tr');
                rows.forEach(function(row) {
                    var cells = row.querySelectorAll('td');

                    // 處理每一個 cell，因為左右各有一組「項目:內容」
                    for (var i = 0; i < cells.length; i++) {
                        var cellText = cells[i].textContent.trim();

                        // 檢查是否包含冒號或：
                        if (cellText.includes('：') || cellText.includes(':')) {
                            var parts = cellText.split(/[：:]/);
                            if (parts.length >= 1) {
                                var key = parts[0].trim();
                                var value = parts.slice(1).join(':').trim();

                                if (key) {
                                    // 如果沒有內容，設為空字串
                                    result[key] = value || '';
                                }
                            }
                        } else {
                            // 如果整個 cell 都是標題（如「土地位置資訊」），跳過
                            if (cellText && !cellText.match(/資訊|列印|分區/)) {
                                // 可能是沒有冒號的項目
                                result[cellText] = '';
                            }
                        }
                    }
                });
            });

            return result;
        """)

        data['表格資料'] = table_data

        # 抓取地圖上的面寬資訊（藍色框框內的數字）
        dimensions = driver.execute_script("""
            var dims = [];
            // 尋找所有包含測量資訊的元素
            var labels = document.querySelectorAll('div[style*="blue"], div[style*="rgb(0, 0, 255)"]');
            labels.forEach(function(label) {
                var text = label.textContent.trim();
                if (text.match(/\\d+\\.\\d+\\(m\\)/)) {
                    dims.push(text);
                }
            });

            // 也嘗試從 canvas 或其他元素取得
            if (dims.length === 0) {
                var allDivs = document.querySelectorAll('div');
                allDivs.forEach(function(div) {
                    var text = div.textContent.trim();
                    if (text.match(/^\\d+\\.\\d+\\(m\\)$/)) {
                        dims.push(text);
                    }
                });
            }

            return dims;
        """)

        data['面寬資訊'] = dimensions if dimensions else []

        # print(f"[OK] 已抓取 {len(data['表格資料'])} 項表格資料", flush=True)
        # print(f"[OK] 已抓取 {len(data['面寬資訊'])} 項面寬資訊", flush=True)

        return data

    except Exception as e:
        print(f"[錯誤] 抓取 V523 資料時發生錯誤: {e}", flush=True)
        return None

def save_pdf(file_name: str, save_directory: str):
    # 🔥 等待新視窗開啟，最多等 15 秒（200ms 高頻 poll，比舊的 1 秒快很多）
    print("[DEBUG] 等待新視窗開啟...", flush=True)
    max_wait_sec = 15
    poll_interval = 0.2
    deadline = time.time() + max_wait_sec
    last_progress_print = -3
    new_window_opened = False
    all_window_handles = driver.window_handles

    while time.time() < deadline:
        all_window_handles = driver.window_handles
        if len(all_window_handles) >= 2:
            elapsed = max_wait_sec - (deadline - time.time())
            print(f"[DEBUG] ✓ 新視窗已開啟（{elapsed:.1f} 秒，共 {len(all_window_handles)} 個視窗）", flush=True)
            new_window_opened = True
            break
        elapsed_int = int(max_wait_sec - (deadline - time.time()))
        if elapsed_int >= last_progress_print + 3:
            print(f"[DEBUG] 等待新視窗... ({elapsed_int}/{max_wait_sec}秒)", flush=True)
            last_progress_print = elapsed_int
        time.sleep(poll_interval)

    if not new_window_opened:
        print(f"[DEBUG] ❌ 等待 {max_wait_sec} 秒後仍無新視窗", flush=True)
        print("無法找到新視窗，可能該地號無資料", flush=True)
        return False

    # 切換到新視窗
    driver.switch_to.window(all_window_handles[-1])
    print("[DEBUG] 已切換到新視窗", flush=True)

    # 🔥 檢查頁面是否為 404 或錯誤頁面
    try:
        current_url = driver.current_url
        print(f"[DEBUG] 當前 URL: {current_url}", flush=True)

        page_text = driver.find_element(By.TAG_NAME, "body").text

        if "404" in page_text or "Not Found" in page_text or "找不到" in page_text or "不是私人連線" in page_text:
            print("⚠️ 該地號沒有資料（404 Not Found），自動關閉視窗並跳過", flush=True)
            driver.close()
            driver.switch_to.window(all_window_handles[0])
            return False

        print("[DEBUG] ✓ 頁面正常，準備抓取資料", flush=True)

    except Exception as e:
        print(f"[DEBUG] 檢查頁面時發生錯誤: {e}", flush=True)
        driver.close()
        driver.switch_to.window(all_window_handles[0])
        return False

    # 🔥 先抓取頁面資料
    # print("正在抓取頁面資料...", flush=True)
    v523_data = extract_v523_data()

    # 儲存 JSON 到 4.其他相關 資料夾
    if v523_data:
        try:
            # 從 save_directory 推導出 4.其他相關 資料夾路徑
            # save_directory 格式: "大寮區義堂段-358-21/1.基本資料"
            base_folder = os.path.dirname(save_directory)  # "大寮區義堂段-358-21"
            other_folder = os.path.join(base_folder, "4.其他相關")
            os.makedirs(other_folder, exist_ok=True)

            # 從 file_name 取得地號作為 JSON 檔名
            # file_name 格式: "04_V523-大寮區義堂段-358-21.pdf"
            json_file_name = file_name.replace('.pdf', '.json').replace('04_V523-', 'V523資料-')
            json_file_path = os.path.join(other_folder, json_file_name)

            with open(json_file_path, 'w', encoding='utf-8') as f:
                json.dump(v523_data, f, ensure_ascii=False, indent=2)

            # print(f"[OK] 已儲存資料至：{json_file_path}", flush=True)
        except Exception as e:
            print(f"[錯誤] 儲存 JSON 資料時發生錯誤: {e}", flush=True)

    # 🔥 捲動到分區圖並放大地圖
    try:
        print("正在調整地圖顯示...", flush=True)

        # 🔥 先捲動回最上方，讓使用者看到網頁完整內容
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(2)  # 停留在最上方，讓使用者看到網頁

        # 🔥 使用原生的平滑捲動到控制按鈕位置（約頁面的 20%，讓控制按鈕出現在畫面下緣）
        driver.execute_script("""
            var targetY = document.body.scrollHeight * 0.4;
            window.scrollTo({
                top: targetY,
                behavior: 'smooth'
            });
        """)
        time.sleep(2.5)  # 🔥 增加等待時間，讓使用者能看到慢速捲動過程

        # 找到並點擊地圖攝影機控制按鈕（使用 aria-label）
        try:
            camera_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="地圖攝影機控制項"]'))
            )
            time.sleep(0.5)  # 🔥 增加延遲，讓使用者看到按鈕
            camera_button.click()
            print("已點擊地圖攝影機控制按鈕", flush=True)
            time.sleep(0.8)  # 🔥 增加延遲，讓控制按鈕展開
        except Exception as e:
            print(f"找不到地圖攝影機控制按鈕: {e}", flush=True)

        # 找到並點擊放大按鈕（+）兩次
        try:
            # 🔥 記錄當前滾動位置
            scroll_position = driver.execute_script("return window.pageYOffset;")

            # 第一次放大
            zoom_in_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="放大"]'))
            )
            time.sleep(0.5)  # 讓使用者看到放大按鈕

            # 使用 JavaScript 點擊，避免觸發滾動
            driver.execute_script("arguments[0].click();", zoom_in_button)
            time.sleep(0.8)  # 等待地圖放大動畫完成
            print("已放大地圖（第 1 次）", flush=True)

            # 🔥 恢復滾動位置
            driver.execute_script(f"window.scrollTo(0, {scroll_position});")
            time.sleep(0.3)

            # 🔥 第二次放大：重新定位按鈕元素（避免 stale element）
            zoom_in_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="放大"]'))
            )

            # 使用 JavaScript 點擊，避免觸發滾動
            driver.execute_script("arguments[0].click();", zoom_in_button)
            time.sleep(0.8)  # 等待地圖放大動畫完成
            print("已放大地圖（第 2 次）", flush=True)

            # 🔥 再次恢復滾動位置
            driver.execute_script(f"window.scrollTo(0, {scroll_position});")
            time.sleep(0.3)

            # 🔥 再次點擊控制按鈕，隱藏控制項（讓上下左右 +- 六個按鈕消失）
            try:
                camera_button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="地圖攝影機控制項"]'))
                )
                driver.execute_script("arguments[0].click();", camera_button)
                print("已隱藏地圖控制按鈕", flush=True)
                time.sleep(0.5)
            except Exception as hide_e:
                print(f"隱藏控制按鈕時發生錯誤: {hide_e}", flush=True)

        except Exception as e:
            print(f"放大地圖時發生錯誤: {e}", flush=True)

    except Exception as e:
        print(f"調整地圖顯示時發生錯誤: {e}", flush=True)
    
    print("\033[93m請調整列印內容(地圖縮放、街景視角)，完成後按【Enter】鍵以繼續...\n\033[0m", flush=True)
    input()

    try:
        # 🔥 使用 Chrome DevTools Protocol 直接生成 PDF（更可靠的方式）
        print("正在生成 PDF...", flush=True)
        result = driver.execute_cdp_cmd("Page.printToPDF", {
            "landscape": False,           # 直向
            "printBackground": True,      # 包含背景
            "scale": 1.0,                 # 縮放比例
            "paperWidth": 8.27,           # A4 寬度（英吋）
            "paperHeight": 11.69,         # A4 高度（英吋）
            "marginTop": 0,               # 上邊距
            "marginBottom": 0,            # 下邊距
            "marginLeft": 0,              # 左邊距
            "marginRight": 0,             # 右邊距
        })

        if 'data' in result:
            import base64
            pdf_data = base64.b64decode(result['data'])

            destination_pdf = os.path.join(save_directory, file_name)
            with open(destination_pdf, 'wb') as f:
                f.write(pdf_data)

            print(f"\033[93m✓ PDF 已儲存到: {destination_pdf}\033[0m", flush=True)
        else:
            print(f"❌ 生成 PDF 失敗：未取得 PDF 資料", flush=True)
            driver.close()
            driver.switch_to.window(all_window_handles[0])
            return False

    except Exception as e:
        print(f"生成 PDF 時發生錯誤: {e}", flush=True)
        driver.close()
        driver.switch_to.window(all_window_handles[0])
        return False

    # 成功完成
    driver.close()
    # print("已關閉 PDF 分頁", flush=True)
    driver.switch_to.window(all_window_handles[0])
    # print("已切換回主查詢頁面", flush=True)
    return True

def test():
    global USERNAME, PASSWORD

    # 讀取資料
    with open(get_data_json_path(), 'r', encoding='utf-8') as file:
        data_list = json.load(file)

    # 先顯示所有讀取到的資料
    print("\n" + "="*39)
    for i, data in enumerate(data_list, 1):
        print(f"{i:2d}. {data['city']} {data['area']} {data['section']} {data['lot_number']}")
    print(f"總共有 {len(data_list)} 筆資料需要處理")
    print("="*39 + "\n")

    # 🔥 檢查縣市，並依縣市從 Windows 認證管理員取出帳號密碼
    # 第一次執行時會跳對話框讓使用者輸入；之後自動讀取
    first_data = data_list[0]
    city = first_data.get('city', '')

    if city == '高雄市':
        service_key = "V523_KH"
        label = "V523 高雄市"
    elif city == '屏東縣':
        service_key = "V523_PT"
        label = "V523 屏東縣"
    else:
        print("\n" + "="*60, flush=True)
        print(f"❌ 錯誤：目前僅支援【高雄市】和【屏東縣】的資料查詢", flush=True)
        print(f"   您的資料縣市為：{city}", flush=True)
        print(f"   目前暫不支援此縣市", flush=True)
        print("="*60 + "\n", flush=True)
        return

    print(f"✓ 偵測到縣市：{city}", flush=True)

    try:
        from credential_helper import get_or_prompt_credentials
    except ImportError as _e:
        print(f"❌ 載入 credential_helper 失敗：{_e}", flush=True)
        print("   請確認 credential_helper.py 與 keyring 套件已安裝", flush=True)
        return

    USERNAME, PASSWORD = get_or_prompt_credentials(service_key, label)
    if not USERNAME or not PASSWORD:
        print(f"❌ 未取得 {label} 的帳密，程式中止", flush=True)
        return

    print("自動化作業啟動中...", flush=True)

    # 詢問是否繼續
    # user_input = input("請確認以上資料是否正確？(按 Enter 繼續，輸入 'n' 取消): ")
    # if user_input.lower() == 'n':
    #     print("程式已取消執行")
    #     return

    # 建立目錄結構
    main_save_directory = make_directory_structure(first_data['area'], first_data['section'], first_data['lot_number'])

    # 初始化瀏覽器
    init_driver()
    login()
    # 🔥 導航到地籍頁面，並自動設定縮放（在頁面載入後）
    need_zoom = open_query_page()

    try:
        # 逐筆處理資料
        for i, data in enumerate(data_list, 1):
            city = data['city']
            area = data['area']
            section = data['section']
            lot_number = data['lot_number']

            pdf_file_name = f"04_V523-{area}{section}-{lot_number}.pdf"

            # 顯示當前處理進度
            print("\n" + "-"*78)
            print(f"【處理進度】第 {i}/{len(data_list)} 筆")
            print(f"目前處理: {city} {area} {section} {lot_number}")
            print(f"檔案名稱: {pdf_file_name}")
            print("-"*78)

            # 執行處理流程
            print(f"[DEBUG] === 開始處理第 {i} 筆資料 ===", flush=True)

            print(f"[DEBUG] 步驟 1/4: 填寫表單", flush=True)
            fill_form_from_data(city, area, section, lot_number)

            print(f"[DEBUG] 步驟 2/4: 點擊位置分區", flush=True)
            query_location()

            print(f"[DEBUG] 步驟 3/4: 點擊確定按鈕", flush=True)
            press_ok_button()

            # 🔥 點擊「分區列印」按鈕（這時才會開啟新視窗）
            print(f"[DEBUG] 步驟 4/4: 點擊分區列印按鈕", flush=True)
            get_pdf_page()

            # 🔥 保存 PDF（會在這裡切換到新視窗並檢查頁面）
            print(f"[DEBUG] 步驟 5/5: 保存 PDF", flush=True)
            if not save_pdf(pdf_file_name, main_save_directory):
                print(f"\033[93m⚠ 第 {i} 筆資料無法處理（頁面載入失敗或無資料），已跳過\033[0m")
                print(f"[DEBUG] === 第 {i} 筆資料處理失敗 ===\n", flush=True)
                continue

            # 完成提示
            print(f"\033[97m✓ 第 {i} 筆資料處理完成\033[0m")

        # 全部完成提示
        # print("\n" + "="*60)
        print("\033[97m【全部處理完成】\033[0m")
        print(f"共處理了 {len(data_list)} 筆資料")
        # print("="*84)

    finally:
        # 🔥 恢復瀏覽器縮放為 100%（無論是否發生錯誤都會執行）
        if need_zoom:
            try:
                # 🔥 重要：先激活 Chrome 窗口，避免鍵盤事件作用到其他程式（如 VSCode）
                try:
                    # 確保切換回主視窗（如果有多個視窗的話）
                    driver.switch_to.window(driver.window_handles[0])
                    # 點擊 body 元素，確保 Chrome 獲得焦點
                    body = driver.find_element(By.TAG_NAME, 'body')
                    body.click()
                    time.sleep(0.5)  # 等待焦點切換完成
                except Exception:
                    pass  # 靜默處理焦點切換錯誤

                # 🔥 根據當前 DPI 決定要恢復幾次（與縮放次數相同）
                try:
                    config_path = os.path.join(BASE_DIR, 'window_config.json')
                    if os.path.exists(config_path):
                        with open(config_path, 'r', encoding='utf-8') as f:
                            config = json.load(f)
                        current_dpi = config.get('dpi_scale', 1.0)
                        if current_dpi >= 1.75:
                            restore_count = 4  # 175%：恢復 4 次（從 67% 回到 100%）
                        elif current_dpi >= 1.5:
                            restore_count = 2  # 150%：恢復 2 次（從 80% 回到 100%）
                        else:
                            restore_count = 2
                    else:
                        restore_count = 2
                except:
                    restore_count = 2

                # 使用系統鍵盤 Ctrl + = (等同於 Ctrl + +) 恢復縮放
                for _ in range(restore_count):
                    keyboard.press_and_release('ctrl+=')
                    time.sleep(0.5)
            except Exception:
                pass  # 靜默處理恢復縮放錯誤

    clear_tmp_directory()

def clear_tmp_directory():
    tmp_pdf_path = os.path.join(current_dir, 'tmp_pdf')
    if os.path.exists(tmp_pdf_path):
        try:
            shutil.rmtree(tmp_pdf_path)
            # print("成功清除所有暫存檔案及目錄", flush=True)
        except Exception as e:
            print(f"清除暫存檔案及目錄時發生錯誤: {e}", flush=True)

def notify_main_program():
    print("V523已完成執行", flush=True)

if __name__ == '__main__':
    test()  # 執行主要邏輯
    notify_main_program()  # 通知主程式

