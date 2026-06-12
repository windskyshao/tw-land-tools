print("歡迎使用【全國地政電子謄本查詢系統】\n程式模組載入中，請稍候...")
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
import io
import re
import sys
import json
import time
import base64
import fitz  # PyMuPDF
import logging
import pikepdf  # 用於 PDF 解密
import requests
import warnings
import numpy as np
import ctypes
import keyboard  # 用於系統級鍵盤事件
from PIL import Image
from getpass import getpass  # 使用 getpass 進行密碼輸入，不會顯示在控制台
from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException,TimeoutException

# 🔥 基準目錄設定（data.json 和工作資料夾的位置）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window

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
hinet_window_width = 1024
hinet_window_height = 1024
hinet_window_x = 0
hinet_window_y = 0

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

            hinet_window_width = chrome_width_logical
            hinet_window_height = chrome_height_logical
            hinet_window_x = chrome_x_logical
            hinet_window_y = chrome_y_logical
except Exception as e:
    pass

# 全域變數放在 import 之後，函數定義之前
auto_confirm_all = False

# 設置日誌
logging.getLogger('selenium').setLevel(logging.CRITICAL)
logging.basicConfig(
    filename='app.log', filemode='w',
    format='%(name)s - %(levelname)s - %(message)s', level=logging.DEBUG
)

# 抑制 PyTorch 警告
warnings.filterwarnings("ignore", category=UserWarning, module="torch")

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='surrogateescape')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='surrogateescape')
sys.stdout.reconfigure(line_buffering=True)

print("✓ 模組載入完成", flush=True)  # 🔥 主程式已先顯示「模組加載中」，這裡改為完成訊息避免重複

# 添加程序結束標記
program_should_exit = False

def check_for_quit_command():
    """檢查是否收到退出命令 - 簡化版本避免 Windows 兼容性問題"""
    global program_should_exit
    
    # 簡化版本，只檢查全域標記
    if program_should_exit:
        print("收到退出命令，程序準備結束...", flush=True)
        return True
    
    # 不進行複雜的鍵盤檢測，避免 WinError 10038
    return False

class suppress_stdout_stderr(object):
    def __enter__(self):
        self.null_fds = [os.open(os.devnull, os.O_RDWR) for _ in range(2)]
        self.save_fds = (os.dup(1), os.dup(2))
        os.dup2(self.null_fds[0], 1)
        os.dup2(self.null_fds[1], 2)

    def __exit__(self, *_):
        os.dup2(self.save_fds[0], 1)
        os.dup2(self.save_fds[1], 2)
        for fd in self.null_fds + list(self.save_fds):
            os.close(fd)


def setup_driver():
    # 設定 WebDriver
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36"
    )
    # 🔥 不在啟動參數設定視窗大小，改為啟動後再調整（避免 DPI 改變時崩潰）
    # options.add_argument("--window-size=1024,1024")
    # options.add_argument("--window-position=0,0")
    options.add_argument("--log-level=3")  # 設定日誌級別為錯誤
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    options.add_experimental_option('useAutomationExtension', False)

    # 新增這些選項來避免網絡錯誤
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")  # 停用擴充功能,加快關閉
    options.add_argument("--remote-debugging-port=0")  # 避免端口衝突

    # 🔥 完全禁用密碼管理器彈窗
    options.add_argument("--disable-save-password-bubble")

    # 禁用密碼儲存和自動填入
    prefs = {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "profile.default_content_setting_values.notifications": 2,  # 禁用通知
        "profile.default_content_settings.popups": 0,  # 允許彈窗（但我們會用其他方式控制）
        "autofill.profile_enabled": False,  # 禁用自動填入
    }
    options.add_experimental_option("prefs", prefs)

    # 🔥 使用 webdriver_helper 統一建立 WebDriver
    driver = create_chrome_driver(options)

    # 🔥 啟動後立即設定視窗大小和位置
    try:
        driver.set_window_size(hinet_window_width, hinet_window_height)
        driver.set_window_position(hinet_window_x, hinet_window_y)
        print(f"[INFO] 全國地政電子謄本視窗已設定: {hinet_window_width}x{hinet_window_height} at ({hinet_window_x},{hinet_window_y})")
    except Exception as e:
        print(f"[WARNING] 視窗設定失敗: {e}，繼續執行")

    return driver


def read_first_entry_from_json(json_file):
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if data and isinstance(data, list) and len(data) > 0:
            first_entry = data[0]
            return first_entry
        else:
            print("data.json 檔案為空或格式不正確。", flush=True)
            return None
    except Exception as e:
        print(f"讀取 data.json 檔案時發生錯誤：{e}", flush=True)
        return None


def construct_custom_directory(first_entry):
    area = first_entry.get('area', '')
    section = first_entry.get('section', '')
    lot_number = first_entry.get('lot_number', '')
    directory_name = f"{area}{section}-{lot_number}"
    return directory_name


def create_custom_directory(base_dir, custom_dir_name):
    # 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
    custom_dir_path = os.path.join(get_work_folder(custom_dir_name), "0.謄本")
    os.makedirs(custom_dir_path, exist_ok=True)
    return custom_dir_path


def preprocess_captcha_image(image):
    # 將圖片轉換為灰度
    gray_image = image.convert('L')
    # 應用二值化處理
    binary_image = gray_image.point(lambda x: 0 if x < 128 else 255, '1')
    # 轉換為 NumPy 陣列，確保類型正確
    numpy_image = np.array(binary_image)
    if numpy_image.dtype != np.uint8:
        numpy_image = numpy_image.astype(np.uint8)
    return numpy_image

def manual_captcha_input(driver):
    captcha_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "aa-captchaID"))
    )
    captcha_input.clear()
    print("請在下方手動輸入【驗證碼】：", flush=True)
    user_captcha = input().strip()
    captcha_input.send_keys(user_captcha)


def get_captcha_image_from_canvas(driver):
    # 使用 JavaScript 獲取 Canvas 的 Base64 圖片數據
    canvas_script = """
    var canvas = document.getElementById('AAAIden1');
    return canvas.toDataURL('image/png').substring(22);  // 去掉 Base64 頭部
    """
    captcha_base64 = driver.execute_script(canvas_script)

    # 將 Base64 圖片轉換為 PIL 圖片對象（記憶體中處理，不保存到磁碟）
    captcha_image_data = base64.b64decode(captcha_base64)
    captcha_image = Image.open(io.BytesIO(captcha_image_data))
    print("從 Canvas 取得驗證碼圖片（記憶體處理）", flush=True)
    return captcha_image

def handle_captcha(driver, register_type, retry_count=0):
    max_retries = 3
    captcha_image_path = None  # 🔥 驗證碼改記憶體處理後不一定有檔案，先給預設值避免 referenced before assignment
    try:
        # 檢查退出命令
        if check_for_quit_command():
            return False

        # 找到驗證碼圖片元素
        print("尋找驗證碼元素...", flush=True)

        # 先等待頁面載入完成
        time.sleep(1)

        # 嘗試多種方式定位驗證碼元素
        captcha_element = None
        try:
            # 方式1：使用 ID
            print("  嘗試使用 ID 定位...", flush=True)
            captcha_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "AAAIden1"))
            )
            print("  ✓ 使用 ID 找到驗證碼元素", flush=True)
        except Exception as e1:
            print(f"  ✗ ID 定位失敗: {str(e1)[:100]}", flush=True)
            try:
                # 方式2：使用 XPath
                print("  嘗試使用 XPath 定位...", flush=True)
                captcha_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//canvas[@id='AAAIden1']"))
                )
                print("  ✓ 使用 XPath 找到驗證碼元素", flush=True)
            except Exception as e2:
                print(f"  ✗ XPath 定位失敗: {str(e2)[:100]}", flush=True)
                # 方式3：列出頁面上所有 canvas 元素
                print("  嘗試查找所有 canvas 元素...", flush=True)
                canvases = driver.find_elements(By.TAG_NAME, "canvas")
                print(f"  頁面上共有 {len(canvases)} 個 canvas 元素", flush=True)
                for i, canvas in enumerate(canvases):
                    canvas_id = canvas.get_attribute("id")
                    print(f"    Canvas {i}: id='{canvas_id}'", flush=True)
                    if "AAA" in str(canvas_id) or "Iden" in str(canvas_id):
                        captcha_element = canvas
                        print(f"  ✓ 找到可能的驗證碼 canvas: {canvas_id}", flush=True)
                        break

        if not captcha_element:
            raise Exception("無法找到驗證碼元素")

        captcha_canvas = captcha_element

        # 滾動驗證碼到視窗中間，確保完全可見
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});",
            captcha_canvas
        )
        print("已滾動到驗證碼位置", flush=True)
        time.sleep(0.5)  # 等待滾動完成

        # 方法1：直接使用元素截圖（Selenium 4.0+）
        try:
            print("使用元素直接截圖方法...", flush=True)
            captcha_image_png = captcha_canvas.screenshot_as_png
            captcha_image = Image.open(io.BytesIO(captcha_image_png))
            print("✓ 使用元素截圖成功（記憶體處理）", flush=True)

        except Exception as element_screenshot_error:
            # 方法2：從 Canvas 提取圖片（備用方案）
            print(f"元素截圖失敗，使用 Canvas 提取方法: {element_screenshot_error}", flush=True)

            try:
                # 從 Canvas 獲取 base64 圖片數據
                canvas_script = """
                var canvas = arguments[0];
                return canvas.toDataURL('image/png').substring(22);
                """
                captcha_base64 = driver.execute_script(canvas_script, captcha_canvas)

                # 解碼圖片（記憶體中處理，不保存到磁碟）
                captcha_image_data = base64.b64decode(captcha_base64)
                captcha_image = Image.open(io.BytesIO(captcha_image_data))
                print("✓ 使用 Canvas 提取成功（記憶體處理）", flush=True)

            except Exception as canvas_error:
                print(f"Canvas 提取失敗: {canvas_error}", flush=True)
                raise Exception("所有驗證碼截圖方法都失敗")

        # 使用 RapidOCR 識別驗證碼 (延遲載入以提升啟動速度)
        print("使用 RapidOCR 識別驗證碼...", flush=True)
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

        # 將 PIL Image 轉換為 numpy 數組，直接傳給 RapidOCR（避免路徑問題）
        captcha_image_np = np.array(captcha_image)
        with suppress_stdout_stderr():
            result = rapidocr_readtext(reader, captcha_image_np, detail=0)

        # 獲取驗證碼結果
        captcha_text = ''.join(result).strip()
        print(f"\033[93m識別出的驗證碼為: {captcha_text}\033[0m", flush=True)

        # 輸入驗證碼
        if len(captcha_text) in [4, 5] and captcha_text.isalnum():
            captcha_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "aa-captchaID"))
            )
            captcha_input.clear()
            captcha_input.send_keys(captcha_text)
            time.sleep(0.5)
            print("填入驗證碼", flush=True)

            # 點擊自然人憑證小方格（第一類謄本選項）
            if register_type == '1':
                checkbox = driver.find_element(
                    By.XPATH, "/html/body/section/div/div[2]/div[3]/form/div[1]/span/input"
                )
                checkbox.click()
                print("點擊自然人憑證小方格", flush=True)

            # 刪除暫存的驗證碼圖片
            try:
                if captcha_image_path and os.path.exists(captcha_image_path):
                    os.remove(captcha_image_path)
                    print("已刪除暫存驗證碼圖片", flush=True)
            except Exception as del_error:
                print(f"刪除驗證碼圖片失敗（不影響程式執行）: {del_error}", flush=True)

            return True
        else:
            print(f"驗證碼格式錯誤: {captcha_text}", flush=True)

            # 即使失敗也嘗試刪除圖片
            try:
                if captcha_image_path and os.path.exists(captcha_image_path):
                    os.remove(captcha_image_path)
            except:
                pass

            return False

    except Exception as e:
        print(f"處理驗證碼時出錯: {e}", flush=True)

        # 嘗試清理暫存圖片
        try:
            captcha_image_path = os.path.join(os.getcwd(), "captcha_image_corrected.png")
            if os.path.exists(captcha_image_path):
                os.remove(captcha_image_path)
        except:
            pass

        if retry_count < max_retries:
            print("自動驗證碼處理失敗，重試...", flush=True)
            return handle_captcha(driver, register_type, retry_count + 1)
        else:
            print("所有自動驗證嘗試均失敗。", flush=True)
            print("進入手動驗證模式...", flush=True)

            # 🔥 無限循環手動模式
            manual_attempt = 1
            while True:
                print(f"\n=== 手動驗證第 {manual_attempt} 次 ===", flush=True)

                try:
                    # 🔧 嘗試自動辨識驗證碼（作為預填值）
                    print("正在自動辨識驗證碼...", flush=True)
                    auto_captcha = None
                    try:
                        # 重新擷取驗證碼圖片
                        captcha_canvas = driver.find_element(By.ID, "AAAIden1")
                        captcha_image_png = captcha_canvas.screenshot_as_png
                        captcha_image = Image.open(io.BytesIO(captcha_image_png))
                        captcha_image_np = np.array(captcha_image)

                        # 使用 RapidOCR 辨識
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
                        result = rapidocr_readtext(reader, captcha_image_np, detail=0)
                        auto_captcha = ''.join(result).strip()
                    except Exception as ocr_err:
                        print(f"自動辨識失敗: {ocr_err}", flush=True)

                    if auto_captcha and len(auto_captcha) in [4, 5]:
                        print(f"\033[93m自動辨識到驗證碼: {auto_captcha}\033[0m", flush=True)
                        print("請檢查驗證碼是否正確：", flush=True)
                        print("  - 若正確，請直接按【ENTER】送出", flush=True)
                        print("  - 若需修改，請輸入正確的驗證碼後按【ENTER】", flush=True)

                        # 🔥 發送特殊標記給主程式，通知預填驗證碼到 GUI 輸入框
                        print(f"__GUI_PREFILL__:{auto_captcha}", flush=True)

                        # 🔥 讓使用者在主程式輸入框中確認或修改
                        user_input = input(f"驗證碼 [{auto_captcha}]: ").strip()

                        if user_input:  # 使用者輸入了修改內容
                            manual_captcha = user_input
                        else:  # 使用者直接按ENTER，使用自動辨識的驗證碼
                            manual_captcha = auto_captcha
                    else:
                        # 自動辨識失敗，請求手動輸入
                        print("自動辨識失敗，請手動輸入驗證碼：", flush=True)
                        manual_captcha = input("驗證碼: ").strip()

                    # 填入驗證碼
                    captcha_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.NAME, "aa-captchaID"))
                    )
                    captcha_input.clear()
                    captcha_input.send_keys(manual_captcha)
                    print(f"已輸入驗證碼: {manual_captcha}", flush=True)

                    # 點擊自然人憑證小方格（第一類謄本選項）
                    if register_type == '1':
                        checkbox = driver.find_element(
                            By.XPATH, "/html/body/section/div/div[2]/div[3]/form/div[1]/span/input"
                        )
                        checkbox.click()
                        print("點擊自然人憑證小方格", flush=True)

                    return True

                except Exception as inner_e:
                    print(f"手動輸入驗證碼時出錯: {inner_e}", flush=True)
                    manual_attempt += 1
                    continue  # 🔥 重新循環，再次嘗試


def login_attempt(driver, register_type, username, password, card_cert=None, max_retries=3):
    attempt = 0
    zoom_applied = False  # 🔥 標記是否已經執行過縮放，避免重複縮放

    while attempt < max_retries:
        try:
            # 檢查退出命令
            if check_for_quit_command():
                return False
            
            print(f"第 {attempt + 1} 次嘗試登入...", flush=True)
            print("打開目標網站:【全國地政電子謄本系統】", flush=True)
            driver.get("https://ep.land.nat.gov.tw/EpaperDoc/MainNew")

            # 🔥 網頁載入後重新設定視窗大小（避免被網站重置）
            try:
                driver.set_window_size(hinet_window_width, hinet_window_height)
                driver.set_window_position(hinet_window_x, hinet_window_y)
            except:
                pass

            # 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
            verify_and_fix_chrome_window(driver)

            # 🔥 自動判斷是否需要縮放頁面（解決高 DPI 或小螢幕下的排版問題）
            # 根據 DPI 決定縮放次數：150% → 2次(80%)，175% → 4次(67%)
            # 🔥 只在第一次嘗試時執行縮放，避免重試時重複縮放
            if not zoom_applied:
                zoom_count = 0
                try:
                    import os as os_mod
                    import json as json_mod
                    config_path = os_mod.path.join(BASE_DIR, 'window_config.json')
                    if os_mod.path.exists(config_path):
                        with open(config_path, 'r', encoding='utf-8') as f:
                            config = json_mod.load(f)

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

                zoom_applied = True  # 🔥 標記已執行過縮放

            print("點擊【登入】，跳轉至【登入頁面】", flush=True)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//a[@href='/Login/AAALogin']"))
            ).click()

            print("填入【用戶識別碼】", flush=True)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "aa-uid"))
            ).send_keys(username)

            print("填入【用戶密碼】", flush=True)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "aa-passwd"))
            ).send_keys(password)

            # 處理驗證碼
            print("開始處理驗證碼...", flush=True)
            if not handle_captcha(driver, register_type):
                print("驗證碼處理失敗，清空輸入欄位，準備重試...", flush=True)
                attempt += 1
                continue

            print("點擊【認證】送出登入資料", flush=True)
            time.sleep(1)
            try:
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "submit_hn"))
                ).click()
            except Exception as e:
                print("嘗試點擊登入按鈕時發生錯誤，將重試...", flush=True)
                time.sleep(2)  # 等待頁面加載完成
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "submit_hn"))
                ).click()

            # 檢查是否有警告框
            print("檢查是否有警告框...", flush=True)
            try:
                WebDriverWait(driver, 3).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                print(f"警告框內容: {alert.text}", flush=True)
                alert.accept()
                print("警告框已接受，繼續操作...", flush=True)
                # 驗證碼失敗，進行下一次嘗試
                attempt += 1
                continue
            except TimeoutException:
                print("無警告框，繼續後續操作...", flush=True)

            # 驗證登入成功
            print("檢查是否成功登入...", flush=True)
            if register_type == '1':
                try:
                    print("自然人憑證資料輸入中...", flush=True)
                    
                    # 使用配置中的自然人憑證資訊（如果有）
                    if card_cert and card_cert.get("id_no") and card_cert.get("pin"):
                        id_no = card_cert["id_no"]
                        pin = card_cert["pin"]
                        print("使用已保存的自然人憑證資訊", flush=True)
                    else:
                        print("未找到保存的自然人憑證資訊，請手動輸入", flush=True)
                        from getpass import getpass
                        id_no = input("請輸入身分證號碼: ").strip()
                        pin = getpass("請輸入憑證密碼: ").strip()
                    
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "idNo"))
                    ).send_keys(id_no)
                    
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "pin"))
                    ).send_keys(pin)
                    
                    driver.find_element(By.CLASS_NAME, "btn.btn-primary").click()
                    print("憑證資料已送出", flush=True)
                    
                    # 如果這是新輸入的資訊，詢問是否保存（改存 Windows 認證管理員）
                    if not (card_cert and card_cert.get("id_no") and card_cert.get("pin")):
                        save_cert = input("是否保存這些自然人憑證資訊以便下次使用？(Y/N): ").strip().lower()
                        if save_cert == 'y':
                            try:
                                from credential_helper import save_credentials as ch_save
                                ch_save("hinet_cert", id_no, pin)
                                print("憑證資訊已加密儲存到 Windows 認證管理員", flush=True)
                            except Exception as e:
                                print(f"保存憑證資訊時出錯: {e}", flush=True)
                    
                except Exception as cert_error:
                    # 處理自然人憑證錯誤
                    retry, action = handle_card_reader_error(driver, cert_error)
                    if retry:
                        # 重新嘗試
                        continue
                    elif action == "2":
                        # 切換到第二類謄本
                        register_type = '2'
                        # 返回主頁重新登入
                        driver.get("https://ep.land.nat.gov.tw/EpaperDoc/MainNew")
                        return login_attempt(driver, '2', username, password, card_cert)
                    elif action == "exit":
                        # 退出程式
                        print("程式即將退出...")
                        sys.exit(0)
                    else:
                        # 處理其他錯誤
                        raise cert_error

            WebDriverWait(driver, 10).until(EC.url_contains("https://ep.land.nat.gov.tw/"))
            print("\033[93m登入成功！\033[0m", flush=True)

            # 嘗試關閉密碼儲存彈窗（如果出現）
            try:
                time.sleep(1)  # 等待彈窗出現
                # 嘗試點擊「一律不要」按鈕（Never Save）
                # Chrome 的密碼彈窗按鈕可能有不同的定位方式
                try:
                    never_button = driver.find_element(By.XPATH, "//button[contains(text(), '一律不要')]")
                    never_button.click()
                    print("已關閉密碼儲存提示", flush=True)
                except:
                    # 如果找不到中文按鈕，嘗試英文
                    try:
                        never_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Never')]")
                        never_button.click()
                        print("已關閉密碼儲存提示", flush=True)
                    except:
                        # 嘗試使用鍵盤按 ESC 關閉
                        try:
                            from selenium.webdriver.common.keys import Keys
                            driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                            print("已使用 ESC 關閉彈窗", flush=True)
                        except:
                            pass  # 彈窗不存在或無法關閉，不影響主流程
            except:
                pass  # 忽略所有錯誤，不影響登入流程

            return True

        except Exception as e:
            print(f"登入過程中出現異常: {e}", flush=True)
            # 檢查是否是讀卡機問題
            retry, action = handle_card_reader_error(driver, e, attempt)
            if retry:
                attempt += 1
                continue
            elif action == "2":
                # 切換到第二類謄本
                register_type = '2'
                # 返回主頁重新登入
                driver.get("https://ep.land.nat.gov.tw/EpaperDoc/MainNew")
                return login_attempt(driver, '2', username, password, card_cert)
            elif action == "exit":
                # 退出程式
                print("程式即將退出...")
                sys.exit(0)
            else:
                # 一般錯誤，繼續嘗試
                attempt += 1

    print("登入嘗試次數已達上限，將啟用手動登入模式", flush=True)
    return manual_login(driver, register_type, username, password)

def manual_login(driver, register_type, username, password):
    # 手動輸入驗證碼
    retry_count = 0
    max_retries = 3
    
    while retry_count < max_retries:
        try:
            # 檢查退出命令
            if check_for_quit_command():
                return False

            print("執行手動輸入【驗證碼】程序", flush=True)
            # 確保頁面重新載入到登入頁面
            driver.get("https://ep.land.nat.gov.tw/Login/AAALogin")
            
            # 清空帳號和密碼
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "aa-uid"))
            ).clear()
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "aa-passwd"))
            ).clear()
            
            # 填充帳號和密碼
            print("填入【用戶識別碼】", flush=True)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "aa-uid"))
            ).send_keys(username)
            print("填入【用戶密碼】", flush=True)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "aa-passwd"))
            ).send_keys(password)
            
            captcha_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "aa-captchaID"))
            )
            captcha_input.clear()
            print("請在下方手動輸入【驗證碼】: ", flush=True)
            user_captcha = input().strip()
            print("填入【驗證碼】", flush=True)
            time.sleep(0.5)
            captcha_input.send_keys(user_captcha)

            if register_type == '1':
                try:
                    checkbox = driver.find_element(
                        By.XPATH,
                        "/html/body/section/div/div[2]/div[3]/form/div[1]/span/input"
                    )
                    checkbox.click()
                    print(f"點擊使用自然人憑證小方格\033[1m囗\033[0m", flush=True)
                except Exception as e:
                    print(f"點擊小方格時出現錯誤: {e}", flush=True)
                    retry, action = handle_card_reader_error(driver, e, retry_count)
                    if retry:
                        retry_count += 1
                        continue
                    elif action == "2":
                        # 切換到第二類謄本
                        print("切換到第二類謄本模式...", flush=True)
                        return manual_login(driver, '2', username, password)
                    elif action == "exit":
                        # 退出程式
                        print("程式即將退出...")
                        sys.exit(0)
                    else:
                        # 其他錯誤
                        return False
                time.sleep(3)

            # 提交登入表單
            submit_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "submit_hn"))
            )
            print("點擊【認證】送出登入資料", flush=True)
            time.sleep(1)
            submit_button.click()

            if register_type == '1':
                try:
                    print("自然人憑證資料輸入中...", flush=True)
                    # 🔥 從 Windows 憑證管理員取得自然人憑證資訊（不再寫死身分證/PIN）
                    from credential_helper import get_credentials
                    _cert_id, _cert_pin = get_credentials("hinet_cert")
                    if not _cert_id or not _cert_pin:
                        print("⚠ 尚未設定自然人憑證帳密，請至「帳號設定」設定 hinet_cert", flush=True)
                        return False
                    # 清空輸入框
                    time.sleep(2)
                    print("清空欄位", flush=True)
                    WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((By.ID, "idNo"))
                    ).clear()
                    WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((By.ID, "pin"))
                    ).clear()
                    print("填入【憑證統號/身分證號】", flush=True)
                    WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((By.ID, "idNo"))
                    ).send_keys(_cert_id)
                    print("填入【憑證密碼】", flush=True)
                    WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((By.ID, "pin"))
                    ).send_keys(_cert_pin)
                    submit_buttonA = WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "btn.btn-primary"))
                    )
                    print("填入【確定送出】", flush=True)
                    submit_buttonA.click() 
                    time.sleep(2)
                except Exception as cert_error:
                    # 處理自然人憑證錯誤
                    retry, action = handle_card_reader_error(driver, cert_error, retry_count)
                    if retry:
                        retry_count += 1
                        continue
                    elif action == "2":
                        # 切換到第二類謄本
                        print("切換到第二類謄本模式...", flush=True)
                        return manual_login(driver, '2', username, password)
                    elif action == "exit":
                        # 退出程式
                        print("程式即將退出...")
                        sys.exit(0)
                    else:
                        # 處理其他錯誤
                        print(f"處理自然人憑證時出錯: {cert_error}", flush=True)
                        retry_count += 1
                        continue

            # 檢查是否登入成功
            try:
                alert = Alert(driver)
                print(f"\033[91m操作時出錯: Alert Text: {alert.text}\033[0m", flush=True)
                print("關閉錯誤彈窗", flush=True)
                alert.accept()  # 關閉彈窗
                print(f"登入失敗，請再重試一次...", flush=True)
                retry_count += 1
            except:
                try:
                    WebDriverWait(driver, 30).until(
                        EC.url_contains("https://ep.land.nat.gov.tw/")
                    )
                    print(f"\033[93m恭喜您，登入成功\033[0m", flush=True)
                    return True
                except:
                    print("無法確認登入狀態，可能登入失敗", flush=True)
                    retry_count += 1
        
        except Exception as e:
            print(f"手動登入過程中出現錯誤: {e}", flush=True)
            retry, action = handle_card_reader_error(driver, e, retry_count)
            if retry:
                retry_count += 1
                continue
            elif action == "2":
                # 切換到第二類謄本
                print("切換到第二類謄本模式...", flush=True)
                return manual_login(driver, '2', username, password)
            elif action == "exit":
                # 退出程式
                print("程式即將退出...")
                sys.exit(0)
            else:
                retry_count += 1
    
    print("手動登入嘗試次數已達上限", flush=True)
    return False
            
def extract_text_from_first_page(pdf_path):
    try:
        with fitz.open(pdf_path) as doc:
            text = doc[0].get_text() if len(doc) > 0 else ''
        return text
    except Exception as e:
        print(f"錯誤: 無法提取 PDF 檔案內容 {pdf_path}, 錯誤訊息: {e}", flush=True)
        return ''

def extract_text_from_pdf(pdf_path):
    """
    從PDF中提取所有頁面的文本
    """
    try:
        with fitz.open(pdf_path) as doc:
            text = ""
            for page in doc:
                text += page.get_text()
            return text
    except Exception as e:
        print(f"錯誤: 無法提取 PDF 檔案內容 {pdf_path}, 錯誤訊息: {e}", flush=True)
        # 如果失敗，回退到只讀取第一頁
        return extract_text_from_first_page(pdf_path)
    
def is_valid_land_number(num):
    """
    檢查是否為有效的地號格式 (XXXX-XXXX)
    """
    return bool(re.match(r'\d{4}-\d{4}$', num))

def is_valid_building_number(num):
    """
    檢查是否為有效的建號格式 (XXXXX-XXX)
    """
    return bool(re.match(r'\d{5}-\d{3}$', num))

def extract_all_numbers(text, pattern, is_valid_func):
    """
    使用正則表達式從文本中提取所有符合模式的號碼
    並過濾掉重複項和無效格式
    """
    matches = pattern.finditer(text)
    result = []
    for match in matches:
        number = match.group(1)
        if number not in result and is_valid_func(number):
            result.append(number)
    return result

def extract_all_land_numbers(text):
    """提取所有地號"""
    pattern = re.compile(r'(\d{4}-\d{4})地號')
    return extract_all_numbers(text, pattern, is_valid_land_number)

def extract_all_building_numbers(text):
    """提取所有建號"""
    pattern = re.compile(r'(\d{5}-\d{3})建號')
    return extract_all_numbers(text, pattern, is_valid_building_number)

def clean_text(text):
    return text.replace('\n', ' ').replace('\r', '')

def extract_base_section(text):
    """提取基本的行政區段名稱"""
    pattern_base = re.compile(r'(\S+(?:鄉|鎮|[市巿]|區).*?段).*?(?=\d{4}-\d{4})')
    base_match = pattern_base.search(text)
    return base_match.group(1).strip() if base_match else ''

def extract_entries_from_text(text):
    """
    將文本分割成各個土地資料區塊並提取資訊
    """
    entries = []
    
    try:
        # 檢查謄本類型
        doc_type = None
        if '土地登記第二類謄本' in text:
            doc_type = '土地登記第二類謄本'
        elif '土地登記第一類謄本' in text:
            doc_type = '土地登記第一類謄本'
        elif '建物登記第二類謄本' in text:
            doc_type = '建物登記第二類謄本'
        elif '建物登記第一類謄本' in text:
            doc_type = '建物登記第一類謄本'
            
        if not doc_type:
            return entries
            
        # 找出基準段（行政區段名稱）
        base_pattern = re.compile(r'([^\s]+(?:鄉|鎮|[市巿]|區).*?段)[^\n]*?')
        base_match = base_pattern.search(text)
        base_section = base_match.group(1).strip() if base_match else ''
        
        # 根據謄本類型選擇正確的編號格式
        if '建物' in doc_type:
            number_pattern = re.compile(r'(\d{5}-\d{3})(?=建號)')
            suffix = '建號'
        else:
            number_pattern = re.compile(r'(\d{4}-\d{4})(?=地號)')
            suffix = '地號'
            
        # 提取所有編號
        numbers = number_pattern.findall(text)
        
        # 移除重複的編號並維持順序
        seen_numbers = set()
        unique_numbers = []
        for num in numbers:
            if num not in seen_numbers:
                seen_numbers.add(num)
                unique_numbers.append(num)
                print(f"找到一筆資料編號: {num}", flush=True)
        
        if base_section and unique_numbers:
            entries.extend([(base_section, num, suffix) for num in unique_numbers])
            
    except Exception as e:
        print(f"提取資料時發生錯誤: {e}", flush=True)
        
    return entries

def format_lot_list(lot_numbers, total_count=None):
    """
    將地號/建號列表格式化為檔名友善的形式：
    - 1~3 筆：完整列出，以頓號分隔（例："157、158、232"）
    - ≥ 4 筆：取第一筆代表 + 等N筆（例："157等26筆"）

    這樣可以避免 Windows 路徑超過 260 字元限制。
    """
    if not lot_numbers:
        return ''
    count = total_count if total_count is not None else len(lot_numbers)
    if count <= 3:
        return '、'.join(lot_numbers[:count])
    return f"{lot_numbers[0]}等{count}筆"

def extract_info_enhanced(text):
    """
    改進版的謄本資訊提取函數，更精確地捕獲所有地號和建號
    支援:
    - 土地登記第一/二類謄本 (多筆地號)
    - 建物登記第一/二類謄本 (多筆建號)
    - 地籍圖謄本
    - 建物測量成果圖
    """
    try:
        # 判斷文件類型
        doc_type = None
        if '土地登記第二類謄本' in text:
            doc_type = '土地登記第二類謄本'
        elif '土地登記第一類謄本' in text:
            doc_type = '土地登記第一類謄本'
        elif '建物登記第二類謄本' in text:
            doc_type = '建物登記第二類謄本'
        elif '建物登記第一類謄本' in text:
            doc_type = '建物登記第一類謄本'
        elif '建物測量成果圖' in text:
            # 建物測量成果圖處理邏輯
            pattern段名 = re.compile(r'段名：\s*(.*?)\s')
            pattern建號 = re.compile(r'建號：\s*母建號\s*(\S+)')
            段名_match = pattern段名.search(text)
            建號_match = pattern建號.search(text)
            段名 = 段名_match.group(1) if 段名_match else ''
            建號 = 建號_match.group(1).translate(str.maketrans('〇一二三四五六七八九', '0123456789')) if 建號_match else ''
            return '建物測量成果圖', 段名, 建號
        elif '地籍圖謄本' in text:
            # 🔥 改寫：不再直接把整段「土地坐落」文字塞進檔名（多筆時會有數千字）
            # 先定位「土地坐落：」之後再搜尋段名，避免前綴被吃進來
            loc_match = re.search(r'土地坐落[：:]\s*', text)
            search_start = loc_match.end() if loc_match else 0
            section_pattern = re.compile(r'([^\s]+?(?:鄉|鎮|[市巿]|區)[^\s]*?段)')
            section_match = section_pattern.search(text, search_start)
            段名 = section_match.group(1) if section_match else ''

            地號資訊 = ''
            if 段名:
                # 從段名結束處往後抓地號列表
                after_section = text[section_match.end():]
                # 優先匹配「X,Y,Z 等N筆」的格式
                multi_match = re.match(r'\s*([\d\-\.,，\s]+?)等\s*(\d+)\s*筆', after_section)
                if multi_match:
                    raw_list = multi_match.group(1)
                    total_count = int(multi_match.group(2))
                    lot_numbers = [n.strip() for n in re.split(r'[,，\s]+', raw_list) if n.strip()]
                    if lot_numbers:
                        地號資訊 = format_lot_list(lot_numbers, total_count) + '地號'
                else:
                    # 沒有「等N筆」字樣時，抓段名後緊接的連續地號字串
                    simple_match = re.match(r'\s*([\d\-\.,，\s]+)', after_section)
                    if simple_match:
                        lot_numbers = [n.strip() for n in re.split(r'[,，\s]+', simple_match.group(1)) if n.strip()]
                        if lot_numbers:
                            地號資訊 = format_lot_list(lot_numbers, len(lot_numbers)) + '地號'

            return '地籍圖謄本', 段名, 地號資訊
            
        if not doc_type:
            return '未知類別', '', ''
            
        # 提取行政區段名稱 (基本段名)
        base_pattern = re.compile(r'([^\s]+(?:鄉|鎮|[市巿]|區).*?段)[^\n]*?')
        base_match = base_pattern.search(text)
        base_section = base_match.group(1).strip() if base_match else ''
        
        if not base_section:
            return doc_type, '', ''
            
        # 提取地號和建號
        land_numbers = []
        building_numbers = []
        
        # 提取所有地號
        land_pattern = re.compile(r'(\d{4}-\d{4})(?=地號)')
        land_matches = land_pattern.findall(text)
        for num in land_matches:
            if num not in land_numbers and is_valid_land_number(num):
                land_numbers.append(num)
                print(f"找到地號: {num}", flush=True)
        
        # 提取所有建號
        building_pattern = re.compile(r'(\d{5}-\d{3})(?=建號)')
        building_matches = building_pattern.findall(text)
        for num in building_matches:
            if num not in building_numbers and is_valid_building_number(num):
                building_numbers.append(num)
                print(f"找到建號: {num}", flush=True)
            
        # 組合結果
        land_info = ""
        building_info = ""
        
        if land_numbers:
            land_info = f"{base_section} {format_lot_list(land_numbers, len(land_numbers))}地號"
            print(f"整理後的地號列表：{land_numbers}", flush=True)

        if building_numbers:
            building_info = f"{base_section} {format_lot_list(building_numbers, len(building_numbers))}建號"
            print(f"整理後的建號列表：{building_numbers}", flush=True)
            
        result_info = ""
        if land_info and building_info:
            result_info = f"{land_info}+{building_info}"
        elif land_info:
            result_info = land_info
        elif building_info:
            result_info = building_info
            
        return doc_type, base_section, result_info
                
    except Exception as e:
        print(f"提取資訊時發生錯誤: {e}", flush=True)
        print(f"發生錯誤的文本片段: {text[:200]}...", flush=True)
        return '未知類別', '', ''

def combine_entries_for_filename(info):
    """
    將地號資訊組合成檔名
    """
    if not info['type'] or not info['numbers']:
        return '未知類別', '', ''
        
    # 用頓號連接所有地號
    numbers_str = '、'.join(info['numbers'])
    
    # 根據類型添加地號或建號後綴
    if '建物' in info['type']:
        numbers_str += '建號'
    else:
        numbers_str += '地號'
        
    return info['type'], info['region'], numbers_str

    
def decrypt_pdf(pdf_path):
    decrypted_path = pdf_path.replace('.pdf', '_decrypted.pdf')
    if os.path.exists(decrypted_path):
        # 如果已經解密過，直接返回解密後的路徑
        return decrypted_path

    try:
        with pikepdf.open(pdf_path) as pdf:
            pdf.save(decrypted_path)  # 解密後保存為新檔案
        print(f"成功解密 PDF: {pdf_path}", flush=True)
        return decrypted_path
    except pikepdf._qpdf.PasswordError:
        print(f"錯誤: 無法解密 PDF {pdf_path}，需要密碼", flush=True)
        return None
    except Exception as e:
        print(f"錯誤: 解密 PDF 過程中發生問題: {e}", flush=True)
        return None
def generate_unique_filename(directory, base_name):
    counter = 1
    new_filename = f"{base_name}.pdf"
    while os.path.exists(os.path.join(directory, new_filename)):
        new_filename = f"{base_name}_{counter}.pdf"
        counter += 1
    return new_filename


def rename_pdf(original_path, output_dir=None, auto_confirm=False, custom_dir_options=None, selected_output_dir=None):
    # 添加目錄選擇
    if auto_confirm and selected_output_dir and selected_output_dir['dir']:
        output_dir = selected_output_dir['dir']
    else:
        if custom_dir_options:
            print("請選擇保存檔案的目錄：", flush=True)
            print("1. 預設目錄（下載的謄本）", flush=True)
            print(f"2. 自訂目錄（{custom_dir_options}）", flush=True)
            print("（直接按下 Enter 選擇預設選項）", flush=True)
            while True:
                choice = input("請輸入選項編號（1 或 2，預設：1）：").strip()
                if not choice:
                    choice = '1'
                if choice == '1':
                    output_dir = get_work_folder("下載的謄本")
                    break
                elif choice == '2':
                    output_dir = custom_dir_options
                    break
                else:
                    print("輸入無效，請重新輸入。", flush=True)
        else:
            output_dir = get_work_folder("下載的謄本")

        if auto_confirm and selected_output_dir is not None:
            selected_output_dir['dir'] = output_dir

    os.makedirs(output_dir, exist_ok=True)

    # 🔥 回傳三態：(output_dir, rename_status, file_path)
    #   rename_status = 'moved'       → 已改名並搬到 output_dir
    #   rename_status = 'saved_only'  → PDF 已下載但還在暫存位置（檔名/路徑異常無法搬動）
    try:
        text = extract_text_from_first_page(original_path)

        # 如果是地籍圖謄本，進行解密
        decrypted_path = original_path
        if '地籍圖謄本' in text:
            print("檢測到地籍圖謄本，開始解密...", flush=True)
            decrypted_result = decrypt_pdf(original_path)
            if decrypted_result:
                decrypted_path = decrypted_result

                # 重新讀取解密後的全部文本
                try:
                    with fitz.open(decrypted_path) as doc:
                        text = ""
                        for page in doc:
                            text += page.get_text()
                except Exception as e:
                    print(f"讀取解密後文件時發生錯誤: {e}", flush=True)
                    text = extract_text_from_first_page(decrypted_path)
            else:
                print(f"跳過未解密的檔案: {original_path}", flush=True)
                return output_dir, 'saved_only', original_path
        else:
            # 非地籍圖謄本，讀取所有頁面的文本
            try:
                with fitz.open(original_path) as doc:
                    text = ""
                    for page in doc:
                        text += page.get_text()
            except Exception as e:
                print(f"讀取全部頁面文本時出錯: {e}", flush=True)
                # 如果讀取所有頁面失敗，回退到只讀第一頁
                text = extract_text_from_first_page(original_path)

        # 提取和清理文本
        cleaned_text = clean_text(text)
        類別, 段名, 地號和建號 = extract_info_enhanced(cleaned_text)
        
        # 顯示提取到的資訊
        print("提取到的資訊:")
        print(f"類別: {類別}")
        print(f"段名: {段名}")
        print(f"地號和建號: {地號和建號}")

        # 檔案命名邏輯
        if 類別 == '建物測量成果圖':
            if 段名 and 地號和建號:
                base_name = f"{類別}-{段名}{地號和建號}"
            elif 段名:
                base_name = f"{類別}-{段名}"
            elif 地號和建號:
                base_name = f"{類別}-{地號和建號}"
            else:
                base_name = None
        elif 類別 == '地籍圖謄本':
            # 🔥 段名+地號資訊都有就一起拼（例：地籍圖謄本-枋山鄉莿桐段157等26筆地號）
            if 段名 and 地號和建號:
                base_name = f"{類別}-{段名}{地號和建號}"
            elif 段名:
                base_name = f"{類別}-{段名}"
            else:
                base_name = None
        elif 類別:  # 其他謄本類型
            base_name = f"{類別}-{地號和建號}" if 地號和建號 else None
        else:
            base_name = None

        if base_name:
            new_filename = generate_unique_filename(output_dir, base_name)
            new_path = os.path.join(output_dir, new_filename)
            try:
                os.rename(decrypted_path, new_path)
                print(f'檔案已重命名並移動到: \n\033[93m{new_path}\033[0m', flush=True)
                return output_dir, 'moved', new_path
            except Exception as e:
                print(f"錯誤: 無法重命名或移動檔案，錯誤訊息: {e}", flush=True)
                # 嘗試複製而不是移動
                try:
                    import shutil
                    shutil.copy2(decrypted_path, new_path)
                    print(f'檔案已複製到: \n\033[93m{new_path}\033[0m', flush=True)

                    # 複製成功後，刪除原始暫存檔
                    try:
                        if os.path.exists(decrypted_path) and decrypted_path != new_path:
                            os.remove(decrypted_path)
                            print(f"已刪除暫存檔: {decrypted_path}", flush=True)
                    except Exception as del_e:
                        print(f"刪除暫存檔失敗（不影響程式執行）: {del_e}", flush=True)
                    return output_dir, 'moved', new_path

                except Exception as copy_e:
                    print(f"錯誤: 無法複製檔案，錯誤訊息: {copy_e}", flush=True)
                    # 改名 + 複製都失敗 → 檔案還在暫存位置
                    return output_dir, 'saved_only', decrypted_path
        else:
            # 無法組出檔名（資訊抽取失敗）→ 檔案仍在暫存位置
            print(f"\033[93m[警告] 無法組出檔名，檔案仍在暫存區：{decrypted_path}\033[0m", flush=True)
            return output_dir, 'saved_only', decrypted_path

    except Exception as e:
        print(f"錯誤: 無法處理檔案 {original_path}，錯誤訊息: {e}", flush=True)
        return output_dir, 'saved_only', original_path
    
def handle_card_reader_error(driver, e, retry_count=0, max_retries=3):
    """
    處理自然人憑證讀卡機相關錯誤
    """
    error_message = str(e).lower()
    # 檢查錯誤類型，判斷是否與自然人憑證相關
    is_card_reader_error = (
        "element not interactable" in error_message or
        "無法使用自然人憑證" in error_message or
        "讀卡機" in error_message or
        "idNo" in error_message or
        "pin" in error_message
    )
    
    if is_card_reader_error and retry_count < max_retries:
        print("\n\033[91m檢測到自然人憑證或讀卡機問題！\033[0m", flush=True)
        print("\n請確認：", flush=True)
        print("1. 自然人憑證讀卡機已正確連接至電腦", flush=True)
        print("2. 自然人憑證晶片卡已正確插入讀卡機", flush=True)
        print("3. 自然人憑證驅動程式已安裝", flush=True)
        print("4. 讀卡機上的指示燈是否有亮起", flush=True)
        
        while True:
            print("\n請在完成上述檢查後，輸入選項：", flush=True)
            print("1. 已修正問題，繼續執行", flush=True)
            print("2. 切換到第二類謄本（不需要自然人憑證）", flush=True)
            print("3. 退出程式", flush=True)
            
            try:
                choice = input("請輸入選項 (1/2/3): ").strip()
                
                if choice == "1":
                    print("重新嘗試自然人憑證登入...", flush=True)
                    return True, None  # 繼續執行，retry
                elif choice == "2":
                    print("切換到第二類謄本模式...", flush=True)
                    return False, "2"  # 切換到第二類謄本
                elif choice == "3":
                    print("程式即將退出...", flush=True)
                    return False, "exit"  # 退出程式
                else:
                    print("無效的選項，請輸入 1、2 或 3", flush=True)
            except Exception:
                print("輸入錯誤，請重新輸入", flush=True)
    else:
        # 不是讀卡機錯誤或已達最大重試次數
        print(f"\033[91m發生錯誤: {e}\033[0m", flush=True)
        if retry_count >= max_retries:
            print("已達最大重試次數，無法繼續", flush=True)
        return False, "error"  # 其他錯誤
    
def load_or_create_config():
    """加載或創建配置（改用 Windows 認證管理員，含舊 config.json 自動遷移）"""
    import os
    import json
    import tkinter as tk
    from tkinter import simpledialog, messagebox

    # 🔥 載入 credential_helper
    try:
        from credential_helper import get_credentials, save_credentials as ch_save
    except ImportError as _ie:
        print(f"❌ 載入 credential_helper 失敗：{_ie}", flush=True)
        return None

    # 🔥 1. 嘗試從 keyring 讀
    user, pwd = get_credentials("hinet")
    cert_id, cert_pin = get_credentials("hinet_cert")
    if user and pwd:
        config = {"username": user, "password": pwd}
        config["card_cert"] = ({"id_no": cert_id, "pin": cert_pin}
                                if cert_id and cert_pin else None)
        print("已從 Windows 認證管理員載入帳密", flush=True)
        return config

    # 🔥 2. 嘗試從舊 config.json 遷移
    config_file = os.path.join(BASE_DIR, "config.json")
    if os.path.exists(config_file):
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                old = json.load(f)
            if old.get("username") and old.get("password"):
                print(f"偵測到舊版 config.json，自動遷移到 Windows 認證管理員...", flush=True)
                ch_save("hinet", old["username"], old["password"])
                cc = old.get("card_cert")
                if isinstance(cc, dict) and cc.get("id_no") and cc.get("pin"):
                    ch_save("hinet_cert", cc["id_no"], cc["pin"])
                try:
                    os.remove(config_file)
                    print(f"已刪除舊憑證檔：{config_file}", flush=True)
                except Exception as _de:
                    print(f"無法刪除舊憑證檔（可手動刪除）：{config_file}：{_de}", flush=True)
                return old
        except Exception as e:
            print(f"讀取舊 config.json 失敗：{e}", flush=True)

    # 🔥 3. 第一次使用：跳對話框輸入
    print("需要設置帳號資訊...", flush=True)

    # 創建自訂對話框
    class ConfigDialog(tk.Toplevel):
        def __init__(self, parent):
            super().__init__(parent)
            self.title("地政電子謄本 - 帳號設定")
            self.geometry("450x450")
            self.resizable(False, False)

            # 置中顯示
            self.update_idletasks()
            x = (self.winfo_screenwidth() // 2) - (450 // 2)
            y = (self.winfo_screenheight() // 2) - (450 // 2)
            self.geometry(f'450x450+{x}+{y}')

            self.result = None
            self.create_widgets()

            # 模態對話框
            self.transient(parent)
            self.grab_set()

        def create_widgets(self):
            # 標題說明
            title_label = tk.Label(self, text="請填寫登入資訊", font=("Microsoft JhengHei", 14, "bold"))
            title_label.pack(pady=(20, 15))

            # 主框架
            main_frame = tk.Frame(self)
            main_frame.pack(padx=30, pady=10, fill=tk.BOTH, expand=True)

            # 用戶名
            tk.Label(main_frame, text="用戶名：", font=("Microsoft JhengHei", 11)).grid(row=0, column=0, sticky='e', pady=8, padx=5)
            self.username_entry = tk.Entry(main_frame, font=("Microsoft JhengHei", 11), width=25)
            self.username_entry.grid(row=0, column=1, pady=8, padx=5)
            self.username_entry.focus()

            # 密碼
            tk.Label(main_frame, text="密碼：", font=("Microsoft JhengHei", 11)).grid(row=1, column=0, sticky='e', pady=8, padx=5)
            self.password_entry = tk.Entry(main_frame, font=("Microsoft JhengHei", 11), width=25, show='*')
            self.password_entry.grid(row=1, column=1, pady=8, padx=5)

            # 分隔線
            separator = tk.Frame(main_frame, height=2, bg='#CCCCCC')
            separator.grid(row=2, column=0, columnspan=2, sticky='ew', pady=15)

            # 說明文字
            tk.Label(main_frame, text="自然人憑證資訊（選填）", font=("Microsoft JhengHei", 10), fg='#666666').grid(row=3, column=0, columnspan=2, pady=(0, 8))

            # 身分證號
            tk.Label(main_frame, text="身分證號：", font=("Microsoft JhengHei", 11)).grid(row=4, column=0, sticky='e', pady=8, padx=5)
            self.id_entry = tk.Entry(main_frame, font=("Microsoft JhengHei", 11), width=25)
            self.id_entry.grid(row=4, column=1, pady=8, padx=5)

            # 憑證密碼
            tk.Label(main_frame, text="憑證密碼：", font=("Microsoft JhengHei", 11)).grid(row=5, column=0, sticky='e', pady=8, padx=5)
            self.pin_entry = tk.Entry(main_frame, font=("Microsoft JhengHei", 11), width=25, show='*')
            self.pin_entry.grid(row=5, column=1, pady=8, padx=5)

            # 按鈕框架
            button_frame = tk.Frame(self)
            button_frame.pack(pady=20)

            # 確定按鈕
            ok_button = tk.Button(button_frame, text="確定", font=("Microsoft JhengHei", 11),
                                  width=10, command=self.ok_clicked, bg='#4CAF50', fg='white')
            ok_button.pack(side=tk.LEFT, padx=10)

            # 取消按鈕
            cancel_button = tk.Button(button_frame, text="取消", font=("Microsoft JhengHei", 11),
                                      width=10, command=self.cancel_clicked, bg='#F44336', fg='white')
            cancel_button.pack(side=tk.LEFT, padx=10)

            # 綁定 Enter 鍵
            self.bind('<Return>', lambda e: self.ok_clicked())
            self.bind('<Escape>', lambda e: self.cancel_clicked())

        def ok_clicked(self):
            username = self.username_entry.get().strip()
            password = self.password_entry.get().strip()

            if not username or not password:
                messagebox.showerror("錯誤", "用戶名和密碼為必填項目", parent=self)
                return

            id_no = self.id_entry.get().strip()
            pin = self.pin_entry.get().strip()

            self.result = {
                "username": username,
                "password": password,
                "card_cert": {
                    "id_no": id_no,
                    "pin": pin
                } if id_no and pin else None
            }
            self.destroy()

        def cancel_clicked(self):
            self.result = None
            self.destroy()

    # 創建主視窗並隱藏
    root = tk.Tk()
    root.withdraw()

    # 顯示對話框
    dialog = ConfigDialog(root)
    dialog.deiconify()        # 🔥 確保對話框顯示
    dialog.lift()             # 🔥 提升到最上層
    dialog.focus_force()      # 🔥 強制聚焦
    dialog.wait_visibility()  # 🔥 等待對話框完全顯示
    root.wait_window(dialog)  # 🔥 等待對話框關閉

    config = dialog.result
    root.destroy()

    if not config:
        print("用戶取消設定", flush=True)
        return None

    # 🔥 儲存到 Windows 認證管理員（不再寫 config.json）
    try:
        ch_save("hinet", config["username"], config["password"])
        cc = config.get("card_cert")
        if isinstance(cc, dict) and cc.get("id_no") and cc.get("pin"):
            ch_save("hinet_cert", cc["id_no"], cc["pin"])
        print("帳密已加密儲存到 Windows 認證管理員", flush=True)
    except Exception as e:
        print(f"儲存帳密時出錯: {e}", flush=True)

    return config

def manage_config():
    """管理配置文件"""
    import os
    import json
    # 🔥 使用完整的 GUI 視窗介面
    import tkinter as tk
    from tkinter import simpledialog, messagebox

    config_file = os.path.join(BASE_DIR, "config.json")

    if not os.path.exists(config_file):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("錯誤", "尚未創建配置文件", parent=root)
        root.destroy()
        return

    try:
        with open(config_file, "r", encoding="utf-8") as f:
            config = json.load(f)
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("錯誤", f"讀取配置文件時出錯: {e}", parent=root)
        root.destroy()
        return

    # 🔥 建立主視窗
    window = tk.Tk()
    window.title("配置管理")
    window.geometry("650x400")
    window.resizable(False, False)
    window.attributes('-topmost', True)  # 🔥 視窗永遠在最前面

    # 設定視窗置中
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - 650) // 2
    y = (screen_height - 400) // 2
    window.geometry(f"650x400+{x}+{y}")

    # 標題
    title_label = tk.Label(window, text="配置管理", font=("微軟正黑體", 18, "bold"))
    title_label.pack(pady=20)

    # 配置項目框架
    frame = tk.Frame(window)
    frame.pack(pady=10, padx=20, fill="both", expand=True)

    # 用戶名
    username_frame = tk.Frame(frame)
    username_frame.pack(fill="x", pady=10)
    username_label = tk.Label(username_frame, text=f"用戶名: {config.get('username', '未設置')}", font=("微軟正黑體", 12), width=35, anchor="w")
    username_label.pack(side="left")

    def modify_username():
        new_username = simpledialog.askstring("修改用戶名", "請輸入新的用戶名:", parent=window)
        if new_username:
            config["username"] = new_username.strip()
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            username_label.config(text=f"用戶名: {config.get('username', '未設置')}")
            messagebox.showinfo("成功", "用戶名已更新", parent=window)

    username_btn = tk.Button(username_frame, text="修改", command=modify_username, font=("微軟正黑體", 11), width=8)
    username_btn.pack(side="right")

    # 密碼
    password_frame = tk.Frame(frame)
    password_frame.pack(fill="x", pady=10)
    password_label = tk.Label(password_frame, text=f"密碼: {'已設置 (******)' if config.get('password') else '未設置'}", font=("微軟正黑體", 12), width=35, anchor="w")
    password_label.pack(side="left")

    def modify_password():
        new_password = simpledialog.askstring("修改密碼", "請輸入新的密碼:", show='*', parent=window)
        if new_password:
            config["password"] = new_password.strip()
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            password_label.config(text=f"密碼: {'已設置 (******)' if config.get('password') else '未設置'}")
            messagebox.showinfo("成功", "密碼已更新", parent=window)

    password_btn = tk.Button(password_frame, text="修改", command=modify_password, font=("微軟正黑體", 11), width=8)
    password_btn.pack(side="right")

    # 自然人憑證
    cert_frame = tk.Frame(frame)
    cert_frame.pack(fill="x", pady=10)
    cert_label = tk.Label(cert_frame, text=f"自然人憑證: {'已設置' if config.get('card_cert') else '未設置'}", font=("微軟正黑體", 12), width=35, anchor="w")
    cert_label.pack(side="left")

    def modify_cert():
        id_no = simpledialog.askstring("自然人憑證", "請輸入身分證號碼:", parent=window)
        if id_no:
            pin = simpledialog.askstring("自然人憑證", "請輸入憑證密碼:", show='*', parent=window)
            if pin:
                config["card_cert"] = {
                    "id_no": id_no.strip(),
                    "pin": pin.strip()
                }
                with open(config_file, "w", encoding="utf-8") as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                cert_label.config(text=f"自然人憑證: {'已設置' if config.get('card_cert') else '未設置'}")
                messagebox.showinfo("成功", "自然人憑證已更新", parent=window)

    cert_btn = tk.Button(cert_frame, text="修改", command=modify_cert, font=("微軟正黑體", 11), width=8)
    cert_btn.pack(side="right")

    # 底部按鈕框架
    bottom_frame = tk.Frame(window)
    bottom_frame.pack(pady=20)

    def clear_all():
        confirm = messagebox.askyesno("確認", "確定要清除所有配置嗎？", parent=window)
        if confirm:
            os.remove(config_file)
            messagebox.showinfo("成功", "配置已清除", parent=window)
            window.destroy()

    clear_btn = tk.Button(bottom_frame, text="清除所有配置", command=clear_all, font=("微軟正黑體", 11), width=15, bg="#ff6b6b", fg="white")
    clear_btn.pack(side="left", padx=10)

    close_btn = tk.Button(bottom_frame, text="關閉", command=window.destroy, font=("微軟正黑體", 11), width=15)
    close_btn.pack(side="left", padx=10)

    window.mainloop()

def main():
    global program_should_exit
    
    try:
        # 清理控制台（適用於 Windows 和其他系統）
        os.system('cls' if os.name == 'nt' else 'clear')
        print(f"～\033[44m【全國地政電子謄本查詢系統】自動化程式\033[0m～", flush=True)

        # 讀取配置文件
        config = load_or_create_config()

        # 🔥 檢查用戶是否取消設定
        if not config:
            print("用戶取消設定，程式退出", flush=True)
            return

        # 從配置中獲取帳號密碼
        username = config.get("username")
        password = config.get("password")
        card_cert = config.get("card_cert")
        
        # 讀取 data.json 的第一組資料，構建自訂目錄
        first_entry = read_first_entry_from_json(get_data_json_path())
        if first_entry:
            custom_dir_name = construct_custom_directory(first_entry)
            # 🔥 使用 BASE_DIR（主程式執行檔所在目錄）
            custom_dir_path = create_custom_directory(BASE_DIR, custom_dir_name)
        else:
            custom_dir_path = None  # 如果讀取失敗，無法使用自訂目錄

        print("請選擇操作選項：", flush=True)
        print("\033[93m１.送件作業\033[0m", flush=True)
        print("\033[93m２.送件、領件整體作業\033[0m", flush=True)
        print("\033[93m３.領件作業\033[0m", flush=True)
        print("\033[93m４.退出程式\033[0m", flush=True)
        print("（如需修改帳密，請回主程式 → 功能分頁 → 修改帳密）", flush=True)

        while True:
            try:
                # 檢查退出命令
                if check_for_quit_command():
                    break

                scenario = int(input())
                if scenario in [1, 2, 3, 4]:
                    print(f"您選擇了選項 {scenario}，正在執行對應操作...", flush=True)
                    break
                else:
                    print("請輸入 1、2、3 或 4 作為選項。", flush=True)
            except ValueError:
                print("請輸入有效的整數選項。", flush=True)
            except KeyboardInterrupt:
                print("\n收到中斷信號，程序即將退出...", flush=True)
                break

        if program_should_exit:
            return

        if scenario == 4:
            # 退出程式選項
            print("\033[38;5;46m感謝使用！程式即將退出...\033[0m", flush=True)
            return

        # 🔥 帳密來源已改為 Windows 認證管理員，正常情況下這裡不會缺少
        # 若真的缺少代表 load_or_create_config 流程被取消，直接導回主選單
        if not username or not password:
            print("\033[91m缺少帳號或密碼，請回主程式 → 功能分頁 → 修改帳密 設定\033[0m", flush=True)
            return main()

        driver = None
        try:
            if scenario == 1:
                # 送件作業
                county, district, section, land_number_list, gland_number, ID_number, register_type, register_options = get_user_input()
                driver = setup_driver()
                if login_attempt(driver, register_type, username, password, card_cert):
                    after_login_success(
                    driver, county, district, section, land_number_list,
                    gland_number, ID_number, register_type, register_options, scenario, custom_dir_path
                    )
                print("\033[41m【送件作業】執行完畢，祝您有美好的一天~~\033[0m", flush=True)
                return main()  # 返回主選單

            elif scenario == 2:
                # 送件、領件整體作業
                county, district, section, land_number_list, gland_number, ID_number, register_type, register_options = get_user_input()
                driver = setup_driver()
                if login_attempt(driver, register_type, username, password, card_cert):
                    after_login_success(
                        driver, county, district, section, land_number_list,
                        gland_number, ID_number, register_type, register_options, scenario, custom_dir_path
                    )
                print("\033[42m【送件&領件】作業執行完畢，祝您冒泡連連~~\033[0m", flush=True)
                return main()  # 返回主選單

            elif scenario == 3:
                # 領件作業
                register_type = "2"  # 使用第二類謄本模式作為預設
                driver = setup_driver()
                if login_attempt(driver, register_type, username, password, card_cert):
                    download_document(driver, scenario, custom_dir_options=custom_dir_path) # 這裡也要傳入
                print("\033[41m【領件作業】執行完畢，祝您心想事成~~\033[0m", flush=True)
                return main()  # 返回主選單
                
        except KeyboardInterrupt:
            print("\n收到中斷信號，程序即將退出...", flush=True)
        except Exception as e:
            print(f"程序執行過程中發生錯誤: {e}", flush=True)
        finally:
            # 確保瀏覽器正確關閉
            if driver:
                try:
                    # 🔥 快速關閉:先關閉所有視窗,再 quit
                    try:
                        main_window = driver.current_window_handle
                        for handle in driver.window_handles:
                            if handle != main_window:
                                try:
                                    driver.switch_to.window(handle)
                                    driver.close()
                                except:
                                    pass
                        driver.switch_to.window(main_window)
                    except:
                        pass

                    # 停止頁面加載,加快退出速度
                    try:
                        driver.execute_script("window.stop();")
                    except:
                        pass

                    driver.quit()
                    print("瀏覽器已關閉", flush=True)
                except Exception as e:
                    # 即使 quit 失敗也不顯示錯誤,因為可能瀏覽器已經關閉
                    pass
    
    except Exception as e:
        print(f"主程序發生嚴重錯誤: {e}", flush=True)
    finally:
        # 確保程序結束時輸出結束訊息
        print("程序準備結束，正在清理資源...", flush=True)
        sys.stdout.flush()
        sys.stderr.flush()
        
def get_user_input():
    # 提示輸入縣市
    print("請輸入縣市 (直接按 Enter 使用預設值：高雄市):", flush=True)
    county = input().replace("台", "臺")
    if not county.strip():  # 判斷輸入是否為空
        county = "高雄市"  # 如果為空，設定預設值
    print(f"您輸入的縣市是: {county}", flush=True)

    # 提示輸入地區
    print("請輸入地區:", flush=True)
    district = input()
    print(f"您輸入的地區是: {district}", flush=True)

    # 提示輸入地段
    print("請輸入地段:", flush=True)
    section = input()
    print(f"您輸入的地段是: {section}", flush=True)

    # 提示輸入地號
    print("請輸入地號 (以','為分隔為同組，以空格分隔為分批多組調閱):", flush=True)
    land_number_input = input()
    # 處理空字符串情況，確保即使沒有輸入也要建立一個列表
    if land_number_input.strip():
        land_number_list = land_number_input.split() # 將地號以空白分割為列表
        print(f"您輸入的地號是: {'、'.join(land_number_list)}", flush=True)
    else:
        land_number_list = []  # 如果沒有輸入，創建空列表
        print(f"您輸入的地號是: ", flush=True)

    # 提示輸入建號
    print("請輸入建號:", flush=True)
    gland_number = input()
    print(f"您輸入的建號是: {gland_number}", flush=True)

    # 提示輸入統一編號
    print("請輸入統一編號(一類):", flush=True)
    ID_number = input()
    print(f"您輸入的統一編號是: {ID_number}", flush=True)

    # 提示選擇謄本類型
    register_type = validate_input(
        "請選擇謄本類型：(\033[093m １.第一類謄本  ２.第二類謄本\033[0m ): ",
        ["1", "2"]
    )
    print(f"您選擇的謄本類型是: {register_type}", flush=True)

    # 提示選擇複選的謄本類型
    print("請選擇以下謄本類型（可複選，用【空白】分隔）: ", flush=True)
    print("\033[093m１.登記謄本\033[0m", flush=True)
    print("\033[093m２.地籍圖謄本\033[0m", flush=True)
    print("\033[093m３.建物測量成果圖\033[0m", flush=True)
    print("請輸入選項號碼(可複選，以【空白】分隔如：1 2 3):", flush=True)
    register_options = input().split()
    print(f"您選擇的選項是: {' '.join(register_options)}", flush=True)
    
    # 加入確認與修改機制
    while True:
        print("\n\033[96m=== 您輸入的資料摘要 ===\033[0m", flush=True)
        print(f"\033[96m1. 縣市: {county}\033[0m", flush=True)
        print(f"\033[96m2. 地區: {district}\033[0m", flush=True)
        print(f"\033[96m3. 地段: {section}\033[0m", flush=True)
        print(f"\033[96m4. 地號: {land_number_list if land_number_list else '無'}\033[0m", flush=True)
        print(f"\033[96m5. 建號: {gland_number}\033[0m", flush=True)
        print(f"\033[96m6. 統一編號: {ID_number}\033[0m", flush=True)
        print(f"\033[96m7. 謄本類型: {'第一類謄本' if register_type == '1' else '第二類謄本'}\033[0m", flush=True)
        
        # 將數字轉換為謄本類型名稱，以便更清楚顯示
        option_names = []
        for opt in register_options:
            if opt == "1":
                option_names.append("登記謄本")
            elif opt == "2":
                option_names.append("地籍圖謄本")
            elif opt == "3":
                option_names.append("建物測量成果圖")
        
        print(f"\033[96m8. 謄本項目: {', '.join(option_names)}\033[0m", flush=True)
        
        print("\n\033[93m請確認以上資料是否正確？\033[0m", flush=True)
        print("\033[93m0. 確認無誤，繼續執行\033[0m", flush=True)
        print("\033[93m1-8. 輸入數字修改對應項目\033[0m", flush=True)
        choice = input("請輸入選項: ").strip()
        
        if choice == '0':
            print("程式開始自動執行...", flush=True)
            break
        elif choice == '1':
            print("請重新輸入縣市 (直接按 Enter 使用預設值：高雄市):", flush=True)
            county = input().replace("台", "臺")
            if not county.strip():
                county = "高雄市"
            print(f"已更新縣市為: {county}", flush=True)
        elif choice == '2':
            print("請重新輸入地區:", flush=True)
            district = input()
            print(f"已更新地區為: {district}", flush=True)
        elif choice == '3':
            print("請重新輸入地段:", flush=True)
            section = input()
            print(f"已更新地段為: {section}", flush=True)
        elif choice == '4':
            print("請重新輸入地號 (以','為分隔為同組，以空格分隔為分批多組調閱):", flush=True)
            land_number_input = input()
            if land_number_input.strip():
                land_number_list = land_number_input.split()
                print(f"已更新地號為: {'、'.join(land_number_list)}", flush=True)
            else:
                land_number_list = []
                print("已清空地號輸入", flush=True)
        elif choice == '5':
            print("請重新輸入建號:", flush=True)
            gland_number = input()
            print(f"已更新建號為: {gland_number}", flush=True)
        elif choice == '6':
            print("請重新輸入統一編號(一類):", flush=True)
            ID_number = input()
            print(f"已更新統一編號為: {ID_number}", flush=True)
        elif choice == '7':
            register_type = validate_input(
                "請重新選擇謄本類型：(\033[093m １.第一類謄本  ２.第二類謄本\033[0m ): ",
                ["1", "2"]
            )
            print(f"已更新謄本類型為: {register_type}", flush=True)
        elif choice == '8':
            print("請重新選擇謄本類型（可複選，用【空白】分隔）: ", flush=True)
            print("\033[093m１.登記謄本\033[0m", flush=True)
            print("\033[093m２.地籍圖謄本\033[0m", flush=True)
            print("\033[093m３.建物測量成果圖\033[0m", flush=True)
            print("請輸入選項號碼(可複選，以【空白】分隔如：1 2 3):", flush=True)
            register_options = input().split()
            print(f"已更新謄本項目為: {' '.join(register_options)}", flush=True)
        else:
            print("\033[91m無效的選項，請輸入0-8之間的數字\033[0m", flush=True)

    return county, district, section, land_number_list, gland_number, ID_number, register_type, register_options

def validate_input(prompt, valid_choices):
    while True:
        print(prompt, flush=True)
        choice = input().strip()
        if choice in valid_choices:
            return choice
        else:
            print("輸入無效，請重新輸入。", flush=True)

def after_login_success(driver, county, district, section, land_number_list, gland_number, ID_number, register_type, register_options, scenario, custom_dir_path):
    # 成功登入並跳轉到 Index2 後的操作
    def handle_alerts(driver):
        print("檢查是否有警告框...", flush=True)
        try:
            alert = driver.switch_to.alert
            print(f"警告框內容: {alert.text}", flush=True)
            alert.accept()
            print("警告框已接受，繼續操作...", flush=True)
        except Exception:
            # 如果沒有警告框，安靜地忽略錯誤
            pass
    
    def check_time_limit():
        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//center/h3[contains(text(), '本日地所作業時間')]"))
            )
            print("\033[91m本日未在開放時間內，無法提供申請！\033[0m", flush=True)
            return False
        except TimeoutException:
            return True

    print("轉跳至【網路申領地政電子謄本作業同意書】", flush=True)

    # 等待並點擊同意方格
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ok"))
    )
    checkbox = driver.find_element(By.ID, "ok")
    time.sleep(0.5)
    checkbox.click()
    print("點擊【同意方格】", flush=True)

    # 等待並點擊同意按鈕
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "yes"))
    )
    agree_button = driver.find_element(By.ID, "yes")
    time.sleep(1)
    agree_button.click()
    print("點擊【同意按鈕】", flush=True)
    time.sleep(1)

    # 時間限制檢查放在這邊
    if not check_time_limit():
        return False

    handle_alerts(driver)

    # 直接跳轉到目標頁面
    driver.get("https://ep.land.nat.gov.tw/EpaperApply/TaiwanMap")
    print("轉跳至【台灣省-縣市選單地圖】", flush=True)
    time.sleep(1)

    # 如果選擇的是"第一類謄本"，則先點擊"請選擇申請角色"
    if register_type == '1':
        role_select = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//form/table[1]//select"))
        )
        Select(role_select).select_by_visible_text("登記名義人代理人")
        print("選擇【申請角色】：登記名義人代理人", flush=True)

    # 等待目標頁面加載完成
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/table[2]/tbody/tr/td/table[2]/tbody/tr/td[1]/form/table[2]/tbody/tr[2]/td/select"))
    )
    time.sleep(1)
    print("展開【縣市下拉選單】", flush=True)
    city_select = driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr/td/table[2]/tbody/tr/td[1]/form/table[2]/tbody/tr[2]/td/select")
    city_select.click()

    # 選擇 "縣市" 的選項，將"台"替換為"臺"
    normalized_county = county.replace("台", "臺")

    try:
        select = Select(city_select)
        time.sleep(2)
        select.select_by_visible_text(normalized_county)
        print(f"成功選擇【縣市】：{normalized_county}", flush=True)
    except Exception as e:
        print(f"【縣市】比對選擇失敗", flush=True)
        print(f"您輸入的縣市名稱『{county}』無法匹配，請手動選擇。", flush=True)

        # 等待使用者手動選擇並監控頁面變化
        WebDriverWait(driver, 60).until(
            EC.url_changes(driver.current_url)
        )
        print("手動選擇完成，程式將繼續執行後續操作...", flush=True)

    handle_alerts(driver)
    
    # 如果選擇的是"第一類謄本"，檢查是否跳轉至代理切結書頁面
    if register_type == '1':
        try:
            # 等待頁面跳轉到代理切結書頁面
            WebDriverWait(driver, 10).until(EC.url_contains("https://ep.land.nat.gov.tw/EpaperApply/ApplyAuthorize"))
            print("跳轉至【代理切結書】頁面", flush=True)

            # 點擊小方格
            checkbox = driver.find_element(By.ID, "ok")
            checkbox.click()
            print("勾選【同意方格】", flush=True)
            time.sleep(1)

            # 點擊"同意"按鈕
            agree_button = driver.find_element(By.ID, "yes")
            agree_button.click()
            print("點擊【同意按鈕】", flush=True)
        except Exception as e:
            print(f"在【代理切結書】頁面執行操作時出現錯誤: {e}", flush=True)
            return False
        
    handle_alerts(driver)

    # 等待【地區下拉選單】加載
    print("等待【地區下拉選單】加載", flush=True)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/table[2]/tbody/tr/td/table[2]/tbody/tr/td[1]/form/table/tbody/tr[2]/td/select")))
    time.sleep(3)
    print("展開【地區下拉選單】", flush=True)
    town_select = driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr/td/table[2]/tbody/tr/td[1]/form/table/tbody/tr[2]/td/select")
    town_select.click()
    time.sleep(3)

    # 選擇 "地區" 的選項
    try:
        town_select = Select(driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr/td/table[2]/tbody/tr/td[1]/form/table/tbody/tr[2]/td/select"))
        town_select.select_by_visible_text(district)
        print(f"成功選擇【地區下拉選單】：{district}", flush=True)
    except Exception as e:
        print(f"【地區】比對選擇失敗", flush=True)
        print(f"您輸入的地區名稱『{district}』無法匹配，請手動選擇。", flush=True)

        # 等待使用者手動選擇並監控頁面變化
        WebDriverWait(driver, 60).until(
            EC.url_changes(driver.current_url)
        )
        print("手動選擇完成，程式將繼續執行後續操作...", flush=True)
    time.sleep(3)
        
    handle_alerts(driver)
    
    # 檢查地號列表是否為空，如果為空，創建一個空字符串的列表項
    if not land_number_list:
        print("未輸入地號，將使用空白地號處理", flush=True)
        land_number_list = [""]  # 創建包含一個空字符串的列表
    
    # 直接傳遞所有參數給 process_single_land_number 函數
    if not process_single_land_number(land_number_list[0], section, driver, gland_number, ID_number, register_options, scenario):
        return False
        
    for index, land_number in enumerate(land_number_list[1:], 1):
        # 只有在後續地號處理時才跳轉回首頁
        # 點擊 "謄本申請" 按鈕
        print("點擊【謄本申請】按鈕", flush=True)
        driver.execute_script("clickback()")
        time.sleep(3)
        # 等待【申請用途下拉選單】加載
        print("等待【申請用途下拉選單】加載", flush=True)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "applyfor"))
        )
        # 找到下拉選單元素並點擊展開
        print("展開【申請用途下拉選單】", flush=True)
        applyfor_select = driver.find_element(By.ID, "applyfor")
        applyfor_select.click()  # 點擊以展開下拉選單
        time.sleep(3)
        select_applyfor = Select(applyfor_select)
        select_applyfor.select_by_value("01")
        print("選擇【購屋、貸款使用】成功", flush=True)

        # 等待並展開地段選單
        print("點擊【地段選單】元素", flush=True)
        session_title = driver.find_element(By.ID, "session_title")
        session_title.click()
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "session_list"))
        )
        # 檢查並處理地段選擇
        district_options = driver.find_elements(By.CLASS_NAME, "session_name")
        district_texts = [option.text for option in district_options]

        corrected_section = section  # 預設為原始輸入

        if section in district_texts:
            district_option = driver.find_element(By.XPATH, f"//div[@class='session_name' and text()='{section}']")
            district_option.click()
            print(f"成功選擇【地段】：{section}", flush=True)
        else:
            print(f"您輸入的地段『{section}』無法匹配，選單內有罕見字或沒有該地段名！", flush=True)
            print("可選的地段如下，請選相近或可能的地段：", flush=True)
            for i, option in enumerate(district_options):
                print(f"{i+1}. {option.text}", flush=True)
            # 提示使用者輸入選擇的索引
            while True:
                try:
                    print("請輸入您選擇的地段索引：", flush=True)
                    index = int(input()) - 1
                    if 0 <= index < len(district_options):
                        district_options[index].click()
                        selected_district = district_options[index].text

                        # 修正輸出
                        if section == "磚子磘段" and selected_district == "磚子段":
                            corrected_section = section
                            print(f"成功選擇【地段】：{selected_district}，轉為【{corrected_section}】", flush=True)
                        else:
                            corrected_section = selected_district
                            print(f"成功選擇【地段】：{corrected_section}", flush=True)
                        break
                    else:
                        print("索引無效，請輸入有效的索引。", flush=True)
                except ValueError:
                    print("輸入無效，請輸入有效的數字。", flush=True)
        
        # 同樣直接傳遞所有參數
        if not process_single_land_number(land_number, section, driver, gland_number, ID_number, register_options, scenario):
            break # 如果使用者取消某筆，就不再往下跑

    # 處理完所有地號，再根據 scenario 判斷是否下載
    if scenario == 2:
        download_document(driver, scenario, custom_dir_options=custom_dir_path)

    # 處理完所有地號，再根據 scenario 判斷是否下載
    if scenario == 2:
        download_document(driver, scenario, custom_dir_options=custom_dir_path)
    


def select_date(driver):
    try:
        # 等待並找到日期下拉選單
        select_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "DateFrom")))
        select_obj = Select(select_element)
        options = [option.text for option in select_obj.options]  # 提取所有選項文本

        # 列出所有日期選項
        print("\n可用的日期選項:", flush=True)
        for index, date in enumerate(options, start=1):
            print(f"{index}. {date}", flush=True)

        print("\n請選擇操作:", flush=True)
        print("數字1-N: 選擇對應的日期", flush=True)
        print("Q: 退出程式", flush=True)

        while True:
            print("請輸入選項 (Q退出程式): ", flush=True)
            user_input = input().strip()

            if user_input.lower() == "q":
                print("程式即將退出...", flush=True)
                sys.exit(0)
            
            # 嘗試轉換為數字
            try:
                choice = int(user_input) - 1
                
                # 檢查選擇是否有效
                if 0 <= choice < len(options):
                    select_obj.select_by_index(choice)  # 選擇用戶選擇的日期
                    print(f"已選擇日期: {options[choice]}", flush=True)

                    # 檢查並等待送出按鈕
                    btn_sent = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "/html/body/table[2]/tbody/tr/td[2]/form/table/tbody/tr/td[2]/div[1]/p/font/input[1]"))
                    )
                    if btn_sent:
                        btn_sent.click()  # 點擊送出按鈕
                        print("已送出日期選擇", flush=True)
                    else:
                        print("無法找到送出按鈕。", flush=True)
                    break  # 成功選擇日期，跳出循環
                else:
                    print("無效的日期編號，請重新選擇。", flush=True)
            except ValueError:
                if user_input.lower() == 'q':
                    print("程式即將退出...", flush=True)
                    sys.exit(0)
                print("輸入的不是一個有效的數字，請重新輸入。", flush=True)

    except Exception as e:
        print(f"處理日期選擇時出錯: {e}", flush=True)
        print("您可以輸入 'Q' 退出程式: ", flush=True)
        user_input = input().strip()
        if user_input.lower() == 'q':
            print("程式即將退出...", flush=True)
            sys.exit(0)

def process_single_land_number(land_number, section, driver, gland_number, ID_number, register_options, current_scenario):
    """
    處理單筆地號的函數
    現在函數接受所有需要的參數，不依賴全域變數
    """
    def find_best_matching_section(target_section, options):
        """
        改進的地段匹配函數，使用精確匹配和模糊匹配結合
        """
        # 首先嘗試完全匹配
        for option in options:
            if target_section == option.text:
                # 完全匹配
                print(f"找到完全匹配的地段：{option.text}", flush=True)
                return option
        
        # 嘗試前綴匹配（輸入的地段是選項的開頭部分）
        for option in options:
            if option.text.startswith(target_section):
                print(f"找到前綴匹配的地段：{option.text}", flush=True)
                # 確認匹配
                print(f"選擇【{option.text}】(從【{target_section}】匹配)")
                return option
        
        # 儲存最佳匹配的選項
        best_option = None
        highest_similarity = 0
        max_similarity_percentage = 0

        for option in options:
            option_text = option.text
            # 計算目標和選項的長度
            target_len = len(target_section)
            option_len = len(option_text)
            
            # 計算共同字符數
            common_chars = 0
            for c1, c2 in zip(target_section, option_text):
                if c1 == c2:
                    common_chars += 1
            
            # 計算相似度百分比（相對於較短的字符串）
            min_len = min(target_len, option_len)
            similarity_percentage = (common_chars / min_len * 100) if min_len > 0 else 0
            
            if similarity_percentage > max_similarity_percentage:
                max_similarity_percentage = similarity_percentage
                best_option = option
        
        # 要求確認與模糊匹配
        if best_option and max_similarity_percentage >= 50:
            print(f"找到最佳匹配的地段：{best_option.text}（相似度：{max_similarity_percentage:.1f}%）", flush=True)
            print(f"您輸入的地段是【{target_section}】，系統建議選擇【{best_option.text}】", flush=True)
            
            print("是否接受此建議？(Y/N)：", flush=True)
            response = input().strip().lower()
            if response == 'y' or response == '':
                return best_option
            else:
                return None
        
        return None  # 沒有找到合適的匹配
        
    # 等待下拉選單加載
    print("等待【申請用途下拉選單】加載", flush=True)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "applyfor"))
    )
    # 找到下拉選單元素並點擊展開
    print("展開【申請用途下拉選單】", flush=True)
    applyfor_select = driver.find_element(By.ID, "applyfor")
    applyfor_select.click()  # 點擊以展開下拉選單
    time.sleep(1)
    #  點擊 "購屋、貸款使用" 的選項
    option_to_select = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//select[@id='applyfor']/option[@value='01']"))
    )
    option_to_select.click()
    print("選擇【購屋、貸款使用】成功", flush=True)

    # 使用 JavaScript 強制關閉下拉選單
    driver.execute_script("arguments[0].blur();", applyfor_select)
    time.sleep(1)  # 稍作等待，確保選單已關閉

    # 等待並展開地段選單
    print("點擊【地段選單】元素", flush=True)
    session_title = driver.find_element(By.ID, "session_title")
    session_title.click()
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, "session_list"))
    )

    # 使用最佳匹配邏輯
    section_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, "session_list"))
    )
    best_option = find_best_matching_section(section, section_dropdown.find_elements(By.CLASS_NAME, "session_name"))
    
    if best_option:
        best_option.click()
        print(f"成功選擇【地段】：{best_option.text}", flush=True)
    else:
        print(f"您輸入的地段『{section}』無法匹配，選單內有罕見字或沒有該地段名！", flush=True)
        print("可選的地段如下，請選相近或可能的地段：", flush=True)
        district_options = section_dropdown.find_elements(By.CLASS_NAME, "session_name")
        for i, option in enumerate(district_options):
            print(f"{i+1}. {option.text}", flush=True)

        # 提示使用者輸入選擇的索引
        while True:
            try:
                print("請輸入您選擇的地段索引：", flush=True)
                index = int(input()) - 1
                if 0 <= index < len(district_options):
                    district_options[index].click()
                    selected_district = district_options[index].text

                    # 修正輸出
                    if section == "磚子磘段" and selected_district == "磚子段":
                        corrected_section = section
                        print(f"成功選擇【地段】：{selected_district}，轉為【{corrected_section}】", flush=True)
                    else:
                        corrected_section = selected_district
                        print(f"成功選擇【地段】：{corrected_section}", flush=True)
                    break
                else:
                    print("索引無效，請輸入有效的索引。", flush=True)
            except ValueError:
                print("輸入無效，請輸入有效的數字。", flush=True)

    time.sleep(1) # 確保選單完全關閉
    
    # 無論是自動選擇還是手動選擇，繼續後續操作
    print("加載【地號輸入框】", flush=True)
    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/table[2]/tbody/tr/td/form[2]/div[2]/table[1]/tbody/tr/td/div[3]/font/input[1]"))
    )

    print("定位【地號輸入框】", flush=True)
    land_number_input = driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr/td/form[2]/div[2]/table[1]/tbody/tr/td/div[3]/font/input[1]")
    # 處理空地號情況
    if land_number and land_number.strip():
        land_number_input.send_keys(land_number)
        print(f"成功填入【地號】：{land_number}", flush=True)
    else:
        print("地號為空，不填入地號", flush=True)
    time.sleep(2)

    print("加載【建號輸入框】", flush=True)
    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/table[2]/tbody/tr/td/form[2]/div[2]/table[1]/tbody/tr/td/div[3]/font/input[2]"))
    )

    print("定位【建號輸入框】", flush=True)
    gland_number_input = driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr/td/form[2]/div[2]/table[1]/tbody/tr/td/div[3]/font/input[2]")
    
    # 使用參數 gland_number
    if gland_number and gland_number.strip():
        gland_number_input.send_keys(gland_number.strip())
        print(f"成功填入【建號】：{gland_number}", flush=True)
    else:
        print("建號為空，不填入建號", flush=True)
    time.sleep(2)

    print("加載【統一編號輸入框】", flush=True)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@id='INPUT_015']"))
    )

    print("定位【統一編號輸入框】", flush=True)
    id_number_input = driver.find_element(By.XPATH, "//input[@id='INPUT_015']")
    
    # 使用參數 ID_number
    if ID_number and ID_number.strip():
        id_number_input.send_keys(ID_number)
        print(f"成功填入【統一編號】：{ID_number}", flush=True)
    else:
        print("統一編號為空，不填入", flush=True)
    time.sleep(2)

    # 使用參數 register_options
    # 根據使用者選擇的選項來點擊對應的選框
    for option in register_options:
        option = option.strip()  # 去掉空白字符
        if option == "1":  # 登記謄本
            print("定位【登記謄本】", flush=True)
            checkbox = driver.find_element("xpath", "//span[@onclick='LabelClick(event);']/input[@type='checkbox' and @name='INPUT_021']")
            checkbox.click()
            print("成功點擊【登記謄本】", flush=True)
            time.sleep(2)

            print("定位【無需列印地上建物建號】", flush=True)
            checkbox = driver.find_element("xpath", "//span[@onclick='LabelClick(event);']/input[@type='checkbox' and @name='INPUT_022']")
            checkbox.click()
            print("成功取消【無需列印地上建物建號】", flush=True)
            time.sleep(2)

            # 根據統一編號的值選擇對應的選項
            id_number_value = id_number_input.get_attribute('value')
            if id_number_value:  # 如果統一編號有值
                print("統一編號已輸入，選擇【所有權個人全部】選項", flush=True)
                ownership_option = driver.find_element(By.XPATH, "//input[@name='INPUT_023' and @value='0']")
            else:  # 如果統一編號沒有值
                print("統一編號未輸入，選擇【全部(土地/建物)】選項", flush=True)
                ownership_option = driver.find_element(By.XPATH, "//input[@name='INPUT_023' and @value='1']")
            ownership_option.click()

        elif option == "2":  # 地籍圖謄本
            print("定位【地籍圖謄本】", flush=True)
            checkbox = driver.find_element("xpath", "//span[@onclick='LabelClick(event);']/input[@type='checkbox' and @name='INPUT_031']")
            checkbox.click()
            print("成功點擊【地籍圖謄本】", flush=True)
            time.sleep(2)

        elif option == "3":  # 建物測量成果圖
            print("定位【建物測量成果圖】", flush=True)
            try:
                checkbox = driver.find_element("xpath", "//span[@onclick='LabelClick(event);']/input[@type='checkbox' and @name='INPUT_041']")
                checkbox.click()
                print("成功點擊【建物測量成果圖】", flush=True)
            except Exception as e:
                print(f"點擊【建物測量成果圖】時出錯: {e}", flush=True)
                # 嘗試使用JavaScript點擊
                try:
                    script = "document.querySelector(\"input[name='INPUT_041']\").click();"
                    driver.execute_script(script)
                    print("使用JavaScript成功點擊【建物測量成果圖】", flush=True)
                except Exception as js_e:
                    print(f"使用JavaScript點擊【建物測量成果圖】時出錯: {js_e}", flush=True)
            time.sleep(2)

    # 等待"新增資料"按鈕可點擊
    print("等待【新增資料】按鈕可點擊", flush=True)
    add_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "btnnew"))
    )

    # 點擊"新增資料"按鈕
    print("點擊【新增資料】按鈕", flush=True)
    add_button.click()
    print("成功點擊【新增資料】按鈕", flush=True)
    time.sleep(3)

    # 等待表格資料加載
    print("等待新增資料內容加載", flush=True)
    try:
        rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//table[@id='DataSheet']/tbody/tr"))
        )

        print("遍歷每一行，提取資料", flush=True)
        for row in rows:
            # 提取每一行的各個單元格
            cells = row.find_elements(By.TAG_NAME, "td")

            # 確保行包含足夠的單元格
            if len(cells) >= 4:
                section_in_row = cells[1].text  # 地段
                land_number_in_row = cells[2].text  # 地號
                building_number = cells[3].text  # 建號
                id_number = cells[4].text if len(cells) > 4 else ""  # 統一編號

                print("==============================\n", flush=True)

                def colored_output(label, value):
                    if value:  # 當有值時才列印
                        print(f"{label}: {value}", flush=True)

                # 根據是否有值變色並列印
                show_user_selections(register_options)
                colored_output("地段", section_in_row)
                colored_output("地號", land_number_in_row)
                colored_output("建號", building_number)
                colored_output("統一編號", id_number)

                print("\n==============================", flush=True)
            else:
                print("此行包含不足的單元格")
    except Exception as e:
        print(f"獲取表格資料時出錯: {e}", flush=True)
        print("繼續執行...", flush=True)

    confirm = check_confirmation(driver, current_scenario)

    if confirm:
        # 按下送出鈕
        submit_button = driver.find_element(By.NAME, "btnsend")
        submit_button.click()

        # 檢查是否出現錯誤訊息
        while True:
            if check_for_errors(driver):
                print("\033[93m請先刪除錯誤及修改資料，並記得【新增資料】後，\n在此輸入 Y 重新送出，或輸入 N 取消操作：\033[0m", flush=True)
                retry = input().strip().lower()
                if retry == 'y':
                    print("重新送出資料...")
                    submit_button.click()  # 再次點擊送出按鈕
                elif retry == 'n':
                    print("您已取消操作，即將退出【送件區】...")
                    return False
                else:
                    print("無效輸入，請輸入 Y 或 N。")
            else:
                print("\033[93m成功申請調閱謄本, 可至【領件區】下載\n\n\033[0m", flush=True)
                # 進行領件區的後續處理
                break
    else:
        print("您已取消操作，即將退出【送件區】...\n", flush=True)
        return False
    return True

def show_user_selections(register_options):
    type_mapping = {
        "1": "登記謄本",
        "2": "地籍圖謄本",
        "3": "建物測量成果圖"
    }
    selected_types = [type_mapping.get(option.strip()) for option in register_options if option.strip() in type_mapping]

    if selected_types:
        print(f"\033[093m您申請了: {', '.join(selected_types)}\033[0m", flush=True)
    else:
        print("您未選擇任何謄本類型", flush=True)

# 在程式碼的開頭，加入全域變數
auto_confirm_all = False

def check_confirmation(driver, scenario):
    """檢查送出資料的確認提示"""
    global auto_confirm_all
    
    if auto_confirm_all:
        print("自動確認：資料已送出", flush=True)
        return True
    
    print("以上資料是否正確？若資料有誤，或欲一次申請多筆可於網頁中更正後，再送出", flush=True)

    if scenario == 2:
        print("\033[93m若先前已申請調閱謄本，可輸入【N】進入【領件區】\033[0m", flush=True)

    # 等待使用者輸入
    import sys
    import time

    while True:
        print("請輸入  [Y]確認本筆  /  [N]取消  /  [A]全部自動同意  /  [Q]退出程式：", flush=True)
        print("\033[93m  ※[A]＝整個程式執行期間，後續所有確認(含送件的其餘筆數、之後領件作業的詢問)一律自動同意、不再詢問；\033[0m", flush=True)
        print("\033[93m    只想確認「這一筆」請按[Y]；按[A]後若要恢復詢問，需關閉再重新開啟程式。\033[0m", flush=True)
        user_input = input().strip().lower()

        if not user_input:
            # 空輸入，直接重新提示，不顯示錯誤訊息
            continue
        elif user_input == 'y':  # 使用者輸入 y 確認
            return True
        elif user_input == 'n':  # 使用者輸入 n 取消
            print("操作已取消。", flush=True)
            return False
        elif user_input == 'a': # 使用者輸入 a 代表自動同意
            auto_confirm_all = True
            print("已設定為自動同意，後續將不再詢問。", flush=True)
            return True
        elif user_input == 'q':  # 使用者輸入 q 退出
            print("程式即將退出...", flush=True)
            sys.exit(0)
        else:
            print("無效的輸入，請輸入[Y]、[N]、[A] 或 [Q]。", flush=True)

def check_for_errors(driver):
    """檢查送出後是否有錯誤訊息，包括彈出視窗和網頁錯誤訊息"""
    # 檢查彈出視窗中的錯誤訊息
    try:
        popup_element = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "msg_dialog"))
        )
        # 嘗試抓取訊息
        popup_msg = popup_element.find_element(By.TAG_NAME, "li").text.strip()

        # 判斷抓取的訊息是否有效
        if popup_msg:
            print(f"\033[91m彈出視窗訊息: {popup_msg}\033[0m", flush=True)
            return True  # 找到有效的彈出視窗訊息
    except TimeoutException:
        print("沒有發現彈出視窗訊息！", flush=True)
    except NoSuchElementException:
        print("謄本調閱中, 請耐心等待...", flush=True)
        time.sleep(10)
    # 檢查網頁中的錯誤訊息
    try:
        error_element = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.ID, "ErrorMsgText"))
        )
        error_msg = error_element.get_attribute('innerText').strip()
        print(f"\033[91m發現網頁錯誤訊息: {error_msg}\033[0m", flush=True)
        return True  # 發現網頁錯誤訊息
    except TimeoutException:
        pass

    return False  # 沒有發現任何錯誤訊息

def download_document(driver, scenario, custom_dir_options=None):
    try:
        time.sleep(1)
        if scenario != 2:
            print("正在執行點擊同意框", flush=True)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ok"))
            ).click()
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "yes"))
            ).click()
            time.sleep(0.5)

        print("正在跳轉到【領件區】", flush=True)
        driver.get("https://ep.land.nat.gov.tw/EpaperDoc/DocQuery")

        WebDriverWait(driver, 10).until(EC.url_contains("/EpaperDoc/DocQuery"))
        print("點擊成功，已跳轉到【領件區】", flush=True)

        # 🔧 加入循環，讓使用者可以在找不到項目時重選日期
        while True:
            # 選擇日期
            print("請選取日期...", flush=True)
            select_date(driver)

            # 🔥 檢查點：選完日期後、進入下載前，先確認資料是否已出現
            #   （謄本可能仍在調閱中、需要時間；使用者沒送件成功；或日期選錯了 → 可重新整理、重選日期或退出）
            _ITEM_XPATH = "//tr[contains(., '/') and .//a[contains(text(), '預覽列印')] and not(contains(., '申請日期時間'))]"
            _reselect_date = False
            while True:
                # 🔥 等清單載入後再判定：資料可能晚幾秒才渲染，避免「其實有資料卻誤判尚未出現」
                _item_count = 0
                for _i in range(10):
                    try:
                        _item_count = len(driver.find_elements(By.XPATH, _ITEM_XPATH))
                    except Exception:
                        _item_count = 0
                    if _item_count > 0:
                        break
                    # 給約 2 秒渲染後，若頁面已明確「查無記錄」就不再空等（加速真的沒資料的情況）
                    if _i >= 2:
                        try:
                            _body = driver.find_element(By.TAG_NAME, 'body').text
                            if ('查無記錄' in _body) or ('共有 0 筆' in _body) or ('共有0筆' in _body):
                                break
                        except Exception:
                            pass
                    time.sleep(1)

                # ✅ 有資料：不必多按 Enter，直接進入下載流程（後面會問起始頁）
                if _item_count > 0:
                    print(f"\033[38;5;46m✓ 此日期偵測到 {_item_count} 筆可領取謄本，直接進入下載...\033[0m", flush=True)
                    break

                # 📭 沒資料：列出選項（各自獨立一行、加顏色好辨識）
                print("", flush=True)
                print("\033[93m📭 此日期目前『尚未』出現可領取的謄本。\033[0m", flush=True)
                print("\033[90m   可能：謄本仍在調閱中、未送件成功，或日期選錯了。\033[0m", flush=True)
                print("請選擇：", flush=True)
                print("\033[38;5;46m   [R]　重新整理再查　—　謄本還在調閱中，稍候再查\033[0m", flush=True)
                print("\033[94m   [D]　重新選擇日期　—　日期可能選錯了\033[0m", flush=True)
                print("\033[91m   [Q]　退出領件　　　—　回主選單\033[0m", flush=True)
                _chk = input().strip().lower()
                if _chk in ('r', ''):  # R 或直接 Enter → 重新整理再查
                    print("重新整理頁面，重新查詢中...", flush=True)
                    driver.refresh()
                    time.sleep(3)
                    continue
                elif _chk == 'd':
                    print("重新選擇日期...", flush=True)
                    # 重載領件區頁面，確保回到乾淨的日期選擇狀態（與系統重選日期一致）
                    try:
                        driver.get("https://ep.land.nat.gov.tw/EpaperDoc/DocQuery")
                        WebDriverWait(driver, 10).until(EC.url_contains("/EpaperDoc/DocQuery"))
                    except Exception:
                        pass
                    _reselect_date = True
                    break  # 跳出檢查點 → 由外層 while 重新選日期
                elif _chk == 'q':
                    print("已退出領件作業，返回主選單。", flush=True)
                    return  # 結束 download_document，回到主選單
                else:
                    print("\033[90m（請按 R 重查、D 換日期 或 Q 退出）\033[0m", flush=True)
                    continue

            if _reselect_date:
                continue  # 回到外層 while 迴圈頂端 → 重新 select_date 選日期

            # 🔥 起始頁、是否自動下載、是否只問一次目錄 → 已移除（精簡流程）
            #   清單已含所有頁，不需起始頁；點領件本來就是要下載；
            #   目錄改為「下載前統一確認」一次決定（見 handle_preview_print）
            auto_download = True
            auto_confirm = True
            selected_output_dir = {'dir': None}

            result = handle_preview_print(
                driver,
                auto_download=auto_download,
                auto_confirm=auto_confirm,
                custom_dir_options=custom_dir_options,
                selected_output_dir=selected_output_dir
            )

            # 🔧 檢查是否需要重選日期
            if result == 'retry_date':
                print("\n" + "=" * 50, flush=True)
                print("返回日期選擇，請重新選擇其他日期", flush=True)
                print("=" * 50 + "\n", flush=True)
                # 重新載入領件區頁面
                driver.get("https://ep.land.nat.gov.tw/EpaperDoc/DocQuery")
                WebDriverWait(driver, 10).until(EC.url_contains("/EpaperDoc/DocQuery"))
                continue  # 繼續循環，重新選擇日期
            else:
                break  # 正常結束或其他情況，跳出循環

    except Exception as e:
        print(f"在下載檔案過程中發生錯誤：{e}", flush=True)

def download_pdf_with_cookies(
    driver,
    auto_confirm,
    custom_dir_options=None,
    selected_output_dir=None
):
    """
    回傳 (status, path)：
      status = 'moved'       → PDF 已下載並歸檔到 output_dir，path 為最終路徑
      status = 'saved_only'  → PDF 已下載但還在暫存位置，path 為目前檔案路徑
      status = 'failed'      → PDF 根本沒抓到，path 為 None
    """
    try:
        # 獲取嵌入的 PDF 連結
        embed_elements = driver.find_elements(By.TAG_NAME, "embed") + \
                         driver.find_elements(By.TAG_NAME, "iframe") + \
                         driver.find_elements(By.TAG_NAME, "object")

        if not embed_elements:
            print("未找到任何嵌入的 PDF 連結。", flush=True)
            return ('failed', None)

        for embed in embed_elements:
            pdf_url = embed.get_attribute("src")
            if pdf_url:
                print(f"找到 PDF 連結: {pdf_url}", flush=True)

                # 建立下載請求（加入重試機制）
                session = requests.Session()
                headers = {'User-Agent': 'Mozilla/5.0'}
                for cookie in driver.get_cookies():
                    session.cookies.set(cookie['name'], cookie['value'])

                # 🔥 加入重試機制處理 DNS 解析失敗
                max_retries = 3
                retry_delay = 2  # 秒
                response = None

                for attempt in range(max_retries):
                    try:
                        response = session.get(pdf_url, headers=headers, timeout=30)
                        response.raise_for_status()  # 檢查HTTP請求是否成功
                        break  # 成功，跳出重試循環
                    except (requests.exceptions.ConnectionError,
                            requests.exceptions.Timeout,
                            requests.exceptions.RequestException) as download_error:
                        if attempt < max_retries - 1:
                            print(f"下載失敗 (嘗試 {attempt + 1}/{max_retries}): {download_error}", flush=True)
                            print(f"等待 {retry_delay} 秒後重試...", flush=True)
                            time.sleep(retry_delay)
                            retry_delay *= 2  # 指數退避
                        else:
                            # 最後一次嘗試失敗，拋出異常
                            raise

                if response is None:
                    print("下載失敗：無法獲取回應", flush=True)
                    continue  # 嘗試下一個 embed 元素

                file_name = f"downloaded_{int(time.time())}.pdf"
                with open(file_name, 'wb') as f:
                    f.write(response.content)
                    print(f"PDF下載並保存為: {file_name}", flush=True)

                 # 提取檔案名稱資訊並重新命名
                print("提取檔案名稱資訊並重新命名...", flush=True)
                類別, 段名, 地號 = '未提取', '未提取', '未提取' #預設值先給定
                try:
                    text = extract_text_from_first_page(file_name)
                    cleaned_text = clean_text(text)
                    類別, 段名, 地號 = extract_info_enhanced(cleaned_text)
                except Exception as e:
                    print(f"提取檔案資訊時發生錯誤: {e}", flush=True)

                # 🔥 rename_pdf 現在回傳三態
                _output_dir, rename_status, final_path = rename_pdf(
                    file_name, output_dir=None,
                    auto_confirm=auto_confirm,
                    custom_dir_options=custom_dir_options,
                    selected_output_dir=selected_output_dir
                )
                return (rename_status, final_path)
            else:
                print("嵌入元素沒有 'src' 屬性。", flush=True)
        return ('failed', None)  # 迴圈結束沒有找到 PDF
    except Exception as e:
        print(f"下載 PDF 檔案時出錯: {e}", flush=True)
        return ('failed', None)
    

def collect_all_preview_items(driver, debug_mode=False, skip_diagnosis=False):
    """收集所有頁面的預覽列印項目資訊

    Args:
        driver: Selenium WebDriver
        debug_mode: 若為 True，會輸出詳細除錯資訊
        skip_diagnosis: 若為 True，跳過診斷輸出（避免重複）
    """
    all_items = []
    page_num = 1

    while True:
        try:
            if debug_mode:
                print(f"[DEBUG] 正在收集第 {page_num} 頁的項目...", flush=True)

            # 🔧 先等待頁面載入完成（等待表格出現）
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//table"))
                )
                if debug_mode:
                    print(f"[DEBUG] 表格已載入", flush=True)
            except:
                if debug_mode:
                    print(f"[DEBUG] 等待表格逾時", flush=True)

            # 使用更精確的 XPath：只選擇真正包含資料的 <tr>
            # 資料列的特徵：有申請時間格式（114/xx/xx），且有「預覽列印」連結
            # 排除表頭和說明文字列
            data_rows = driver.find_elements(By.XPATH,
                "//tr[contains(., '/') and .//a[contains(text(), '預覽列印')] and not(contains(., '申請日期時間'))]")

            # 🔧 如果主要 XPath 找不到，嘗試備用搜尋條件
            if len(data_rows) == 0:
                if debug_mode:
                    print(f"[DEBUG] 主要XPath找不到，嘗試備用條件...", flush=True)
                # 備用1: 只搜尋包含「預覽列印」文字的連結所在的 tr
                data_rows = driver.find_elements(By.XPATH,
                    "//tr[.//a[contains(text(), '預覽列印')]]")
                if debug_mode:
                    print(f"[DEBUG] 備用條件1找到 {len(data_rows)} 列", flush=True)

            if len(data_rows) == 0:
                # 備用2: 搜尋包含「預覽」的連結
                data_rows = driver.find_elements(By.XPATH,
                    "//tr[.//a[contains(text(), '預覽')]]")
                if debug_mode:
                    print(f"[DEBUG] 備用條件2找到 {len(data_rows)} 列", flush=True)

            if debug_mode:
                print(f"[DEBUG] 找到 {len(data_rows)} 列資料", flush=True)

            # 🔍 當找不到資料列時，輸出診斷資訊（只輸出一次）
            if len(data_rows) == 0 and not skip_diagnosis:
                print(f"\n[診斷] 找不到預覽列印項目，開始診斷...", flush=True)

                # 1. 檢查頁面上所有連結
                all_links = driver.find_elements(By.TAG_NAME, "a")
                print(f"[診斷] 頁面上共有 {len(all_links)} 個連結", flush=True)
                preview_links = [link for link in all_links if '預覽' in link.text or '列印' in link.text]
                print(f"[診斷] 其中包含「預覽」或「列印」的連結: {len(preview_links)} 個", flush=True)
                for idx, link in enumerate(preview_links[:10]):  # 最多顯示10個
                    try:
                        print(f"  連結{idx+1}: '{link.text}' (顯示={link.is_displayed()})", flush=True)
                    except:
                        print(f"  連結{idx+1}: (無法取得資訊)", flush=True)

                # 2. 檢查所有表格列
                all_rows = driver.find_elements(By.XPATH, "//tr")
                print(f"[診斷] 頁面上共有 {len(all_rows)} 個 <tr> 列", flush=True)

                # 3. 嘗試用較寬鬆的條件搜尋
                loose_rows = driver.find_elements(By.XPATH, "//tr[.//a]")
                print(f"[診斷] 包含連結的 <tr> 列: {len(loose_rows)} 個", flush=True)

                # 4. 顯示前幾個包含連結的列內容
                for idx, row in enumerate(loose_rows[:5]):
                    try:
                        row_text = row.text.replace('\n', ' ')[:100]
                        print(f"  列{idx+1}: {row_text}...", flush=True)
                    except:
                        pass

                # 5. 截圖保存（用於遠端除錯）
                try:
                    import os
                    screenshot_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_screenshot.png")
                    driver.save_screenshot(screenshot_path)
                    print(f"[診斷] 已保存截圖: {screenshot_path}", flush=True)
                except Exception as e:
                    print(f"[診斷] 截圖失敗: {e}", flush=True)

                print("[診斷] 診斷結束\n", flush=True)

            for i, row in enumerate(data_rows):
                try:
                    # 使用固定索引直接提取資料（根據之前的 DEBUG 輸出）
                    cells = row.find_elements(By.TAG_NAME, "td")

                    # DEBUG: 顯示每個儲存格的內容（已停用）
                    # print(f"    [DEBUG] 列 {i+1} 有 {len(cells)} 個儲存格:", flush=True)
                    # for idx, cell in enumerate(cells):
                    #     text = cell.text.strip()
                    #     if text:
                    #         display_text = text[:80] if len(text) > 80 else text
                    #         print(f"      [{idx}]: {display_text}", flush=True)

                    if len(cells) >= 10:
                        # 資料列結構（基於DEBUG輸出）：
                        # [0]=空, [1]=時間, [2]=地區, [3]=收件號, [4]=種類, [5]=張數, [6]=金額, [7]=按鈕, [8]=下載時間(optional), [9]=地號
                        apply_time = cells[1].text.strip()
                        location = cells[2].text.strip()
                        doc_number = cells[3].text.strip()
                        doc_type = cells[4].text.strip()
                        sheets = cells[5].text.strip()
                        fee = cells[6].text.strip()
                        # 🔥 狀態欄（cells[7]）：含「未下載」＝尚未下載、會扣費；否則為已領取（不再扣費）
                        status_text = cells[7].text.strip() if len(cells) > 7 else ''
                        is_pending = ('未下載' in status_text)

                        # 地號在 cells[9]
                        plot_info = cells[9].text.strip() if len(cells) > 9 else ''
                        # 移除括號內的比例尺資訊
                        if '（' in plot_info:
                            plot_info = plot_info.split('（')[0].strip()
                    else:
                        # print(f"    [警告] 列 {i+1} 儲存格數量不足", flush=True)
                        continue

                    # print(f"    提取: {apply_time} | {location} | {doc_number} | {doc_type} | {fee} | {plot_info}", flush=True)

                    # 驗證資料有效性
                    if '/' in apply_time and '電謄字' in doc_number:
                        # 找到「預覽列印」連結
                        link = row.find_element(By.XPATH, ".//a[contains(text(), '預覽列印')]")

                        item_info = {
                            'page': page_num,
                            'row_index': i,  # 在當前頁面中的行索引
                            'index': len(all_items),
                            'apply_time': apply_time,
                            'location': location,
                            'doc_number': doc_number,
                            'doc_type': doc_type,
                            'sheets': sheets,
                            'fee': fee,
                            'pending': is_pending,   # True＝未下載（會扣費）；False＝已領取過（不扣費）
                            'plot_info': plot_info,
                            'link': link,
                            'row_element': row
                        }
                        all_items.append(item_info)
                        # print(f"  - 項目 {len(all_items)}: {doc_type} ({doc_number}) {plot_info} - {fee}元", flush=True)
                    # else:
                        # print(f"    [警告] 列 {i+1} 資料格式不正確", flush=True)

                except Exception as e:
                    # print(f"  提取第 {i+1} 列資訊時出錯: {e}", flush=True)
                    continue

            # 嘗試找下一頁按鈕
            try:
                next_page_button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "//li[@class='pgNext']/a"))
                )
                if next_page_button.is_displayed() and next_page_button.is_enabled():
                    # print(f"找到下一頁按鈕，前往第 {page_num + 1} 頁...", flush=True)
                    next_page_button.click()
                    time.sleep(2)
                    page_num += 1
                else:
                    break
            except (NoSuchElementException, TimeoutException):
                # print("已到最後一頁", flush=True)
                break

        except Exception as e:
            # print(f"收集第 {page_num} 頁項目時出錯: {e}", flush=True)
            break

    return all_items


def display_items_and_get_selection(items, downloaded=None):
    """顯示所有項目並讓使用者複選（downloaded：已下載過的 index 會標記）"""
    downloaded = downloaded or set()
    if not items:
        print("沒有可下載的項目", flush=True)
        return []

    print("\n" + "=" * 39, flush=True)
    print("可下載的謄本清單：", flush=True)
    print("=" * 39, flush=True)
    print(f"{'編號'} {'地區':<12} {'謄本種類':<10} {'金額':<6} {'地號/備註'}", flush=True)
    print("-" * 63, flush=True)

    for i, item in enumerate(items, 1):
        plot_info = item.get('plot_info', '')
        # 已領取＝網頁狀態非「未下載」，或本次已下載過 → 不扣費；其餘為未下載 → 需扣費
        already = (not item.get('pending', True)) or ((i - 1) in downloaded)
        mark = "  \033[38;5;46m不扣費\033[0m" if already else "  \033[97;41m 需扣費 \033[0m"
        print(f"{i:<3} {item['location']:<12} {item['doc_type']:<10} {item['fee']:<6} {plot_info}{mark}", flush=True)

    print("=" * 39, flush=True)
    print("\n\033[96m請選擇要下載的項目：\033[0m", flush=True)
    print("  - 輸入編號，多個編號請用逗號分隔（例如：1,3,5）", flush=True)
    print("  - 輸入範圍請用減號（例如：1-5）", flush=True)
    print("  - 輸入 \033[38;5;46m'all' 或 'a'\033[0m 下載全部", flush=True)
    print("  - 直接按 \033[91mEnter\033[0m 取消下載", flush=True)

    while True:
        user_input = input("請輸入選擇：").strip()
        # print(f"[DEBUG] 收到輸入: '{user_input}'", flush=True)

        if not user_input:
            # print("[DEBUG] 空輸入，取消下載", flush=True)
            return []

        if user_input.lower() in ['all', 'a']:
            # print(f"[DEBUG] 選擇全部 {len(items)} 個項目", flush=True)
            return list(range(len(items)))

        try:
            selected_indices = []
            parts = user_input.split(',')

            for part in parts:
                part = part.strip()
                if '-' in part:
                    # 範圍選擇
                    start, end = part.split('-')
                    start = int(start.strip())
                    end = int(end.strip())
                    for idx in range(start - 1, end):
                        if 0 <= idx < len(items) and idx not in selected_indices:
                            selected_indices.append(idx)
                else:
                    # 單一選擇
                    idx = int(part) - 1
                    if 0 <= idx < len(items) and idx not in selected_indices:
                        selected_indices.append(idx)

            if selected_indices:
                selected_indices.sort()
                print(f"\n已選擇 {len(selected_indices)} 個項目：", flush=True)
                for idx in selected_indices:
                    item = items[idx]
                    print(f"  - {idx + 1}. {item['doc_type']} ({item['doc_number']}) - {item['fee']}元", flush=True)

                # 直接返回選擇的索引，不再要求二次確認
                # 用戶已經明確選擇了項目，應該直接進行下載
                # print(f"[DEBUG] 確認下載 {len(selected_indices)} 個項目", flush=True)
                return selected_indices
            else:
                print("未選擇任何有效項目，請重新輸入", flush=True)

        except ValueError:
            print("輸入格式錯誤，請重新輸入", flush=True)


def download_single_item(driver, item, item_number, total_items, auto_confirm, custom_dir_options, selected_output_dir):
    """下載單個項目

    回傳 (status, path)：
      status = 'moved'       → 下載且歸檔成功
      status = 'saved_only'  → 下載成功但未歸檔，path 是檔案目前所在位置
      status = 'failed'      → 下載失敗
    """
    try:
        print(f"\n[{item_number}/{total_items}] 正在處理：{item['doc_type']} ({item['doc_number']}) - {item['fee']}元", flush=True)

        try:
            data_rows_xpath = "//tr[contains(., '/') and .//a[contains(text(), '預覽列印')] and not(contains(., '申請日期時間'))]"
            data_rows = driver.find_elements(By.XPATH, data_rows_xpath)

            row_index = item.get('row_index', 0)
            if row_index >= len(data_rows):
                print(f"\033[91m行索引 {row_index} 超出範圍（共 {len(data_rows)} 行）\033[0m", flush=True)
                return ('failed', None)

            target_row = data_rows[row_index]
            link = target_row.find_element(By.XPATH, ".//a[contains(text(), '預覽列印')]")

            # 滾動到連結
            driver.execute_script("arguments[0].scrollIntoView();", link)
            WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.ID, "uiLockId")))
            time.sleep(1)

            # 點擊連結
            link.click()
            print(f"已點擊『預覽列印』連結", flush=True)

            main_window_handle = driver.current_window_handle
            windows = driver.window_handles

            if len(windows) > 1:
                new_window_handle = [handle for handle in windows if handle != main_window_handle][0]
                driver.switch_to.window(new_window_handle)

                # 🔥 用容器承接 download_pdf_with_cookies 的回傳，避免 try/except/finally 的 scope 問題
                dl_result = ('failed', None)

                try:
                    # 檢查是否有收費資訊
                    fee_element = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//strong[contains(text(), '本筆案件將收費')]"))
                    )
                    # 有收費資訊，點擊 yes 按鈕
                    pay_button = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.NAME, "yes"))
                    )
                    pay_button.click()

                    # 等待下載視窗
                    WebDriverWait(driver, 10).until(EC.new_window_is_opened(driver.window_handles))
                    new_windows = [handle for handle in driver.window_handles if handle != main_window_handle and handle != new_window_handle]

                    if new_windows:
                        download_window = new_windows[0]
                        driver.switch_to.window(download_window)
                        dl_result = download_pdf_with_cookies(driver, auto_confirm=auto_confirm, custom_dir_options=custom_dir_options, selected_output_dir=selected_output_dir)
                    else:
                        print("\033[91m未找到下載視窗\033[0m", flush=True)
                        return ('failed', None)

                except TimeoutException:
                    # 沒有收費資訊，直接下載
                    print("沒有找到收費資訊，直接下載...", flush=True)
                    dl_result = download_pdf_with_cookies(driver, auto_confirm=auto_confirm, custom_dir_options=custom_dir_options, selected_output_dir=selected_output_dir)

                finally:
                    # 關閉新視窗並回到主視窗
                    try:
                        driver.execute_script("window.close();")
                    except:
                        try:
                            driver.close()
                        except:
                            pass
                    driver.switch_to.window(main_window_handle)

                # 🔥 根據 dl_result 顯示正確的結果訊息
                status, path = dl_result
                if status == 'moved':
                    print(f"\033[38;5;46m✓ 下載並歸檔完成\033[0m", flush=True)
                elif status == 'saved_only':
                    print(f"\033[93m⚠ 已下載但未歸檔，檔案留在：{path}\033[0m", flush=True)
                else:
                    print(f"\033[91m✗ 下載失敗\033[0m", flush=True)
                return dl_result
            else:
                print("\033[91m未找到新視窗\033[0m", flush=True)
                return ('failed', None)

        except Exception as e:
            print(f"\033[91m找不到該項目的連結: {e}\033[0m", flush=True)
            return ('failed', None)

    except Exception as e:
        print(f"\033[91m下載項目時出錯: {e}\033[0m", flush=True)
        return ('failed', None)


def navigate_to_page(driver, target_page):
    """導航到指定頁碼"""
    try:
        # 檢查當前在哪一頁
        current_page = 1
        try:
            active_page = driver.find_element(By.XPATH, "//li[@class='active']/a")
            current_page = int(active_page.text)
        except:
            pass

        if current_page == target_page:
            return True

        # 需要翻頁
        if target_page < current_page:
            # 先回到第一頁
            try:
                first_page_link = driver.find_element(By.XPATH, "//li[@class='pgFirst']/a")
                first_page_link.click()
                time.sleep(2)
            except:
                pass

        # 從第一頁翻到目標頁
        for _ in range(target_page - 1):
            try:
                next_button = driver.find_element(By.XPATH, "//li[@class='pgNext']/a")
                next_button.click()
                time.sleep(2)
            except:
                break

        return True

    except Exception as e:
        print(f"導航到第 {target_page} 頁時出錯: {e}", flush=True)
        return False


def handle_preview_print(
    driver,
    auto_download=False,
    auto_confirm=False,
    retry_count=0,
    custom_dir_options=None,
    selected_output_dir=None
):
    max_retries = 3  # 設定最大重試次數

    try:
        # 第一步：收集所有頁面的項目資訊
        # 🔥 訊息提前印（在耗時動作前），讓使用者立即看到，避免誤以為卡住
        print("列表提取中...", flush=True)
        # 🔧 只在第一次嘗試時輸出診斷，重試時跳過診斷
        skip_diagnosis = (retry_count > 0)
        all_items = collect_all_preview_items(driver, debug_mode=False, skip_diagnosis=skip_diagnosis)

        if not all_items:
            print("找不到任何可下載的項目", flush=True)
            if retry_count < max_retries:
                # 🔧 不再額外呼叫診斷，因為 collect_all_preview_items 已經輸出過了
                print("重新整理頁面重試...", flush=True)
                driver.refresh()
                time.sleep(3)  # 減少等待時間
                return handle_preview_print(
                    driver,
                    auto_download=auto_download,
                    auto_confirm=auto_confirm,
                    retry_count=retry_count + 1,
                    custom_dir_options=custom_dir_options,
                    selected_output_dir=selected_output_dir
                )
            else:
                print("已達到最大重試次數，無法找到『預覽列印』連結。", flush=True)
                print("該日期可能沒有可下載的項目，將返回日期選擇...", flush=True)
                return 'retry_date'  # 返回特殊值，表示需要重選日期

        # print(f"\n共收集到 {len(all_items)} 個項目", flush=True)

        # 目錄分流設定（迴圈外固定不變）
        fallback_dir = get_work_folder("下載的謄本")
        case_dir = custom_dir_options  # 案件資料夾的「0.謄本」（來自 data.json 第一筆），或 None

        # 🔥 逐檔智慧分流：每一筆依「自己的區」各自歸檔（相符→案件0.謄本；否則→下載的謄本）
        def _district_of(s):
            ms = re.findall(r'[^\s]{1,3}(?:區|鄉|鎮|市)', str(s or ''))
            return ms[-1] if ms else ''
        _case_area = _district_of(os.path.basename(os.path.dirname(case_dir))) if case_dir else ''

        def _dir_for(item):
            if case_dir and _case_area and (_district_of(item.get('location', '')) in ('', _case_area)):
                return case_dir
            return fallback_dir

        downloaded = set()  # 本次已下載（領取）的 index，避免重複領件付費

        # 🔁 清單 ⇄ 下載 可循環：下載完回清單，可再下載其他、回主選單或離開
        while True:
            # 第二步：選擇要下載的項目（已下載者會標記）
            selected_indices = display_items_and_get_selection(all_items, downloaded)
            if not selected_indices:
                print("未選擇任何項目，離開領件", flush=True)
                return True

            # 第三步：金額 + 逐檔分流，下載前統一確認（下載＝領件＝會扣費！）
            force_dir = None
            _proceed = True
            while True:
                total_fee = 0     # 只累計「會扣費（未下載過）」的金額
                n_charge = 0      # 會實際扣費的筆數
                n_free = 0        # 已領取過、不再扣費的筆數
                n_case = 0
                n_fallback = 0
                for _idx in selected_indices:
                    _it = all_items[_idx]
                    _already = (not _it.get('pending', True)) or (_idx in downloaded)
                    if _already:
                        n_free += 1
                    else:
                        n_charge += 1
                        try:
                            total_fee += int(''.join(filter(str.isdigit, str(_it.get('fee', '0')))) or 0)
                        except Exception:
                            pass
                    _d = force_dir or _dir_for(_it)
                    if case_dir and _d == case_dir:
                        n_case += 1
                    else:
                        n_fallback += 1
                print("", flush=True)
                print(f"\033[93m即將下載 {len(selected_indices)} 筆：\033[0m", flush=True)
                if n_charge > 0:
                    print(f"\033[91m   ▸ {n_charge} 筆會扣費，合計 {total_fee} 元（下載即領件、會產生費用，請確認！）\033[0m", flush=True)
                if n_free > 0:
                    print(f"\033[38;5;46m   ▸ {n_free} 筆先前已領取，不會再扣費\033[0m", flush=True)
                print("\033[96m存放方式（依各筆的「區」自動分流）：\033[0m", flush=True)
                if case_dir and n_case > 0:
                    print(f"\033[96m   • {n_case} 筆（與案件「{_case_area}」相符）→ 案件「0.謄本」\033[0m", flush=True)
                if n_fallback > 0:
                    print(f"\033[96m   • {n_fallback} 筆 → 「下載的謄本」\033[0m", flush=True)
                print("請確認（項目已選定，直接 Enter 即可下載）：", flush=True)
                print("\033[38;5;46m   [Enter] 或 [Y]　確認下載\033[0m", flush=True)
                if case_dir and n_case > 0 and not force_dir:
                    print("\033[94m   [A]　全部改存到「下載的謄本」\033[0m", flush=True)
                print("\033[91m   [N]　取消，回清單重選\033[0m", flush=True)
                _c = input().strip().lower()
                if _c in ('y', ''):
                    break
                elif _c == 'a' and case_dir and not force_dir:
                    force_dir = fallback_dir
                    continue
                elif _c == 'n':
                    _proceed = False
                    break
                else:
                    print("\033[90m（請按 Enter/Y 確認、A 全部存下載的謄本 或 N 取消）\033[0m", flush=True)
                    continue
            if not _proceed:
                continue  # 取消 → 回清單重選

            auto_confirm = True

            # 第四步：依選擇逐檔下載
            print(f"\n開始下載 {len(selected_indices)} 個項目...", flush=True)
            print("=" * 80, flush=True)
            moved_count = 0
            saved_count = 0
            failed_count = 0
            staged_paths = []
            try:
                first_page_link = driver.find_element(By.XPATH, "//li[@class='pgFirst']/a")
                first_page_link.click()
                time.sleep(2)
            except:
                pass
            current_page = 1
            for i, idx in enumerate(selected_indices, 1):
                item = all_items[idx]
                # 逐檔智慧分流：依此筆的「區」決定存檔目錄
                _item_dir = force_dir or _dir_for(item)
                if selected_output_dir is not None:
                    selected_output_dir['dir'] = _item_dir
                try:
                    os.makedirs(_item_dir, exist_ok=True)
                except Exception:
                    pass
                if item['page'] != current_page:
                    print(f"\n導航到第 {item['page']} 頁...", flush=True)
                    if navigate_to_page(driver, item['page']):
                        current_page = item['page']
                    else:
                        print(f"\033[91m無法導航到第 {item['page']} 頁，跳過此項目\033[0m", flush=True)
                        failed_count += 1
                        continue
                status, path = download_single_item(driver, item, i, len(selected_indices), auto_confirm, custom_dir_options, selected_output_dir)
                if status == 'moved':
                    moved_count += 1
                    downloaded.add(idx)
                elif status == 'saved_only':
                    saved_count += 1
                    downloaded.add(idx)
                    if path:
                        staged_paths.append(path)
                else:
                    failed_count += 1

            print("\n" + "=" * 80, flush=True)
            print("下載完成！", flush=True)
            print(f"  \033[38;5;46m成功下載並歸檔：{moved_count} 個\033[0m", flush=True)
            if saved_count > 0:
                print(f"  \033[93m已下載但未歸檔：{saved_count} 個（需手動搬移）\033[0m", flush=True)
                for p in staged_paths:
                    print(f"    - {p}", flush=True)
            print(f"  \033[91m下載失敗：{failed_count} 個\033[0m", flush=True)
            print("=" * 80, flush=True)

            # 🔥 下載後選單：回清單再下載其他 / 回主選單 / 離開程式
            remaining = [j for j in range(len(all_items)) if j not in downloaded]
            print("", flush=True)
            if remaining:
                print(f"\033[96m清單還有 {len(remaining)} 筆尚未下載。接下來：\033[0m", flush=True)
                print("\033[38;5;46m   [Enter] 或 [L]　回清單，再下載其他\033[0m", flush=True)
                print("\033[94m   [M]　回主選單（送件／領件）\033[0m", flush=True)
                print("\033[91m   [Q]　離開程式\033[0m", flush=True)
                _post = input().strip().lower()
                if _post == 'q':
                    print("離開程式。", flush=True)
                    sys.exit(0)
                elif _post == 'm':
                    return True
                else:
                    continue  # 回清單（沿用記憶體清單，已下載者會標記）
            else:
                print("\033[38;5;46m清單項目已全部下載完成。\033[0m", flush=True)
                print("\033[94m   [Enter] 或 [M]　回主選單\033[0m", flush=True)
                print("\033[91m   [Q]　離開程式\033[0m", flush=True)
                _post = input().strip().lower()
                if _post == 'q':
                    print("離開程式。", flush=True)
                    sys.exit(0)
                else:
                    return True

    except Exception as e:
        print(f"處理『預覽列印』連結時發生錯誤: {e}", flush=True)
        return False


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n程序被中斷", flush=True)
    except Exception as e:
        print(f"程序異常結束: {e}", flush=True)
    finally:
        # 強制輸出結束訊息
        print("子程式已完成執行", flush=True)
        sys.stdout.flush()
        sys.stderr.flush()
        # 確保程序完全結束
        import atexit
        atexit.register(lambda: print("程序已安全退出", flush=True))