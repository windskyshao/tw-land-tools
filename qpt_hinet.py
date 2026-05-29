print("歡迎使用【全功能地籍查詢系統】\n程式模組載入中，請稍候...", flush=True)
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


import pyautogui  # 需要安装: pip install pyautogui
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
import shutil
import numpy as np
import ctypes
import keyboard  # 用於系統級鍵盤事件
from PIL import Image
from getpass import getpass  # 使用 getpass 進行密碼輸入，不會顯示在控制台
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException,TimeoutException
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window

# 🔥 基準目錄設定（data.json 和工作資料夾的位置）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

# 🔥 顏色輸出函數
def cprint(text):
    """亮藍色輸出"""
    print(f"\033[96m{text}\033[0m", flush=True)

def yprint(text):
    """亮黃色輸出"""
    print(f"\033[93m{text}\033[0m", flush=True)

def bprint(text):
    """藍色輸出（保留但不常用）"""
    print(f"\033[94m{text}\033[0m", flush=True)

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
qpt_window_width = 1024
qpt_window_height = 1024
qpt_window_x = 0
qpt_window_y = 0

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

            qpt_window_width = chrome_width_logical
            qpt_window_height = chrome_height_logical
            qpt_window_x = chrome_x_logical
            qpt_window_y = chrome_y_logical
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

# 抑制 PyTorch 和 EasyOCR 的所有警告訊息
warnings.filterwarnings("ignore", category=UserWarning, module="torch")
warnings.filterwarnings("ignore", category=FutureWarning)  # 隱藏 FutureWarning

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
sys.stdout.reconfigure(line_buffering=True)

print("歡迎使用【地籍電子謄本】小程式，模組加載中...", flush=True)


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


def setup_driver(pdf_save_dir=None):
    """
    設置WebDriver，添加自動保存PDF的预設置

    Args:
        pdf_save_dir: PDF 儲存目錄，如果為 None 則使用預設目錄
    """
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

    # 🔥 加速啟動和關閉
    options.add_argument("--disable-extensions")  # 停用擴充功能
    options.add_argument("--disable-gpu")  # 停用GPU加速(加快關閉)
    options.add_argument("--no-sandbox")  # 停用沙箱(加快啟動)

    # 🔥 完全禁用密碼管理器彈窗
    options.add_argument("--disable-save-password-bubble")

    # 添加PDF列印预設置
    if pdf_save_dir is None:
        pdf_save_dir = get_work_folder("下載的謄本")
    os.makedirs(pdf_save_dir, exist_ok=True)
    
    prefs = {
        # 🔥 完全禁用密碼管理器和自動填入
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "profile.default_content_setting_values.notifications": 2,  # 禁用通知
        "autofill.profile_enabled": False,  # 禁用自動填入
        # PDF 列印設定
        "printing.print_preview_sticky_settings.appState": json.dumps({
            "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
            "isHeaderFooterEnabled": False,
            "marginsType": 0  # 預設邊界
        }),
        "savefile.default_directory": pdf_save_dir,  # 自動儲存的路徑
        "download.default_directory": pdf_save_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": False,  # 🔥 改為 False，讓PDF直接儲存
        "printing.default_destination_selection_rules": {
            "kind": "local",
            "namePattern": "Save as PDF"
        }
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--kiosk-printing")  # 啟用自動列印
    options.add_argument("--disable-popup-blocking")  # 停用彈出視窗封鎖

    # 🔥 使用 webdriver_helper 統一建立 WebDriver
    driver = create_chrome_driver(options)

    # 🔥 啟動後立即設定視窗大小和位置
    try:
        driver.set_window_size(qpt_window_width, qpt_window_height)
        driver.set_window_position(qpt_window_x, qpt_window_y)
        print(f"[INFO] 全功能地籍視窗已設定: {qpt_window_width}x{qpt_window_height} at ({qpt_window_x},{qpt_window_y})")
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

def handle_alert_with_retry(driver, max_wait=3, max_retries=3):
    """處理可能延遲出現的警告框，增加檢測次數和等待時間"""
    for i in range(max_retries):
        try:
            # 設置更長的等待時間，以確保能捕獲警告框
            alert = WebDriverWait(driver, max_wait).until(EC.alert_is_present())
            if alert:
                alert_text = alert.text
                print(f"警告框內容: {alert_text}", flush=True)
                alert.accept()
                print("警告框已接受", flush=True)
                return True, alert_text
        except TimeoutException:
            print(f"等待警告框超時 (第 {i+1}/{max_retries} 次檢查)", flush=True)
            if i == max_retries - 1:
                # 最後一次檢查失敗
                return False, None
            # 短暫等待後再次檢查
            time.sleep(0.5)
        except Exception as e:
            print(f"檢查警告框時發生錯誤: {e}", flush=True)
            # 嘗試直接切換到警告框
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                print(f"直接獲取到警告框: {alert_text}", flush=True)
                alert.accept()
                print("警告框已接受", flush=True)
                return True, alert_text
            except:
                print("直接獲取警告框也失敗", flush=True)
            # 短暫等待後再次檢查
            time.sleep(0.5)
    
    # 所有嘗試都失敗
    return False, None

def print_error(e):
    """簡化錯誤信息輸出"""
    error_msg = str(e)
    if "Stacktrace:" in error_msg:
        error_parts = error_msg.split("Stacktrace:")
        print(f"\033[91m錯誤: {error_parts[0].strip()}\033[0m", flush=True)
    else:
        print(f"\033[91m錯誤: {error_msg}\033[0m", flush=True)


def handle_captcha(driver, register_type, retry_count=0, max_retries=3):
    """處理驗證碼，支持自動和手動輸入"""
    try:
        # 自動嘗試次數已達上限，轉為手動模式但保留在當前頁面
        if retry_count >= max_retries:
            print(f"自動驗證碼識別已失敗 {max_retries} 次，轉為手動輸入模式", flush=True)
            return manual_captcha_input_enhanced(driver, register_type)
        
        # 找到驗證碼元素
        captcha_image_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "AAAIden1"))
        )
        
        # 嘗試刷新驗證碼
        try:
            print("嘗試點擊驗證碼圖片以刷新...", flush=True)
            captcha_image_element.click()
            time.sleep(1)
        except Exception as e:
            print(f"刷新驗證碼失敗: {e}", flush=True)
        
        # 方法1：直接使用元素截圖（Selenium 4.0+）
        try:
            # print("使用元素直接截圖方法...", flush=True)
            captcha_image_png = captcha_image_element.screenshot_as_png
            captcha_image = Image.open(io.BytesIO(captcha_image_png))
            # print("✓ 使用元素截圖成功（記憶體處理）", flush=True)

        except Exception as element_screenshot_error:
            # 方法2：從 Canvas 提取圖片（備用方案）
            print(f"元素截圖失敗，使用 Canvas 提取方法: {element_screenshot_error}", flush=True)

            try:
                # 從 Canvas 獲取 base64 圖片數據
                canvas_script = """
                var canvas = arguments[0];
                return canvas.toDataURL('image/png').substring(22);
                """
                captcha_base64 = driver.execute_script(canvas_script, captcha_image_element)

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

        # 前 3 次自動送出，不詢問用戶
        if retry_count < max_retries:
            print(f"\033[93m驗證碼第 {retry_count + 1} 次嘗試：{captcha_text}\033[0m", flush=True)
        else:
            # 這個分支不會執行到，因為在函數開頭就會轉為手動模式
            print(f"\033[93m識別出的驗證碼為: {captcha_text}\033[0m", flush=True)

        # 輸入驗證碼
        captcha_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "aa-captchaID"))
        )
        captcha_input.clear()
        captcha_input.send_keys(captcha_text)
        time.sleep(0.5)
        print("填入驗證碼", flush=True)

        # 只在第一類謄本選項時才點擊自然人憑證小方格
        if register_type == '1':
            try:
                checkbox = driver.find_element(
                    By.XPATH, "/html/body/section/div/div[2]/div[3]/form/div[1]/span/input"
                )
                checkbox.click()
                print("點擊自然人憑證小方格", flush=True)
            except Exception as e:
                print(f"點擊小方格時出錯: {e}", flush=True)

        return True
    except Exception as e:
        print(f"處理驗證碼時出錯: {e}", flush=True)
        return handle_captcha(driver, register_type, retry_count + 1, max_retries)

    
def manual_captcha_input_enhanced(driver, register_type):
    """在當前頁面進行手動驗證碼輸入（支援自動辨識預填和無限重試）"""
    manual_attempt = 1
    while True:  # 🔥 無限循環
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

            # 清空驗證碼輸入框
            captcha_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "aa-captchaID"))
            )
            captcha_input.clear()

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
                    user_captcha = user_input
                else:  # 使用者直接按ENTER，使用自動辨識的驗證碼
                    user_captcha = auto_captcha
            else:
                # 自動辨識失敗，請求手動輸入
                print("\033[93m請手動輸入驗證碼（看當前網頁畫面）：\033[0m", flush=True)
                user_captcha = input().strip()

            captcha_input.send_keys(user_captcha)
            print(f"已填入驗證碼：{user_captcha}", flush=True)

            # 確保第一類謄本勾選了自然人憑證
            if register_type == '1':
                try:
                    checkbox = driver.find_element(
                        By.XPATH, "/html/body/section/div/div[2]/div[3]/form/div[1]/span/input"
                    )
                    # 檢查是否已勾選
                    if not checkbox.is_selected():
                        checkbox.click()
                        print("點擊自然人憑證小方格", flush=True)
                except Exception as e:
                    print(f"點擊小方格時出錯: {e}", flush=True)

            return True

        except Exception as e:
            print(f"手動輸入驗證碼時出錯: {e}", flush=True)
            manual_attempt += 1
            continue  # 🔥 重新循環，再次嘗試

def login_attempt(driver, register_type, username, password, card_cert=None, base_url=None, max_retries=3):
    attempt = 0
    zoom_applied = False  # 🔥 標記是否已經執行過縮放，避免重複縮放

    # 獲取當前URL，判斷縣市代碼
    current_url = driver.current_url
    print(f"當前URL: {current_url}", flush=True)
    
    # 使用傳入的base_url，如果沒有則從當前URL提取
    if not base_url:
        # 從URL提取縣市代碼（舊邏輯）
        county_code = "kc"  # 默認高雄市
        base_url = "https://pqt-kcgetw.land.nat.gov.tw"
        
        try:
            domain_parts = current_url.split('/')
            if len(domain_parts) > 2:
                domain = domain_parts[2]  # 例如 pqt-pthgetw.land.nat.gov.tw
                
                if '-' in domain and 'getw' in domain:
                    parts = domain.split('-')
                    if len(parts) > 1:
                        county_part = parts[1]
                        if 'getw' in county_part:
                            county_code = county_part.split('getw')[0]  # 例如從 pthgetw 提取 pth
                            base_url = f"https://pqt-{county_code}getw.land.nat.gov.tw"
        except Exception as e:
            print(f"提取縣市代碼時發生錯誤: {e}，使用默認高雄市代碼", flush=True)
    
    # 從base_url提取縣市代碼，用於顯示
    try:
        if "pqt-" in base_url and "getw" in base_url:
            domain_part = base_url.split('/')[2]  # 例如 pqt-pthgetw.land.nat.gov.tw
            county_code_match = re.search(r'pqt-(.+?)getw', domain_part)
            if county_code_match:
                county_code = county_code_match.group(1)
            else:
                county_code = "未知"
        else:
            county_code = "未知"
    except:
        county_code = "未知"
    
    print(f"使用縣市代碼: {county_code}, 基礎URL: {base_url}", flush=True)
    
    # 接下來是login_attempt函數的其餘部分，但在所有使用base_url的地方保持一致
    # ...

    while attempt < max_retries:
        try:
            print(f"第 {attempt + 1} 次嘗試登入...", flush=True)
            
            # 重要: 檢查當前URL，如果包含錯誤代碼或不是正確的登入頁面，則重新導航
            current_url = driver.current_url
            if "errorCode" in current_url or "aa-result=s341" in current_url or (
                "AuthScreen.jsp" not in current_url and "aaav2.hinet.net" not in current_url):
                print("當前頁面不是登入頁面或包含錯誤代碼，重新導航至系統首頁...", flush=True)
                # 先導航到系統首頁
                driver.get(base_url)

                # 網頁載入後重新設定視窗大小（避免被網站重置）
                try:
                    driver.set_window_size(qpt_window_width, qpt_window_height)
                    driver.set_window_position(qpt_window_x, qpt_window_y)
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

                time.sleep(2)

                # 點擊進入系統按鈕
                try:
                    enter_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'進入系統') or contains(text(),'登入')]"))
                    )
                    enter_button.click()
                    print("已點擊【進入系統】按鈕", flush=True)
                except Exception as e:
                    print(f"點擊【進入系統】按鈕時出錯: {e}", flush=True)
                    # 嘗試使用JavaScript執行
                    try:
                        driver.execute_script("enterSystem();")
                        print("已使用JavaScript執行進入系統", flush=True)
                    except Exception as js_e:
                        print(f"JavaScript執行失敗: {js_e}", flush=True)
                        print("請手動點擊進入系統按鈕...", flush=True)
                        # 等待用戶操作
                        WebDriverWait(driver, 60).until(
                            lambda d: "AuthScreen.jsp" in d.current_url or "aaav2.hinet.net" in d.current_url
                        )
                
                # 等待頁面加載
                time.sleep(3)
            
            # 確保我們現在在登入頁面
            new_url = driver.current_url
            print(f"跳轉到驗證頁面的 URL: {new_url}", flush=True)
            
            # 檢查是否已經在登入頁面
            if "AuthScreen.jsp" not in new_url and "aaav2.hinet.net" not in new_url:
                print("未能成功跳轉到登入頁面，嘗試直接導航...", flush=True)
                driver.get("https://aaav2.hinet.net/A1/AuthScreen.jsp")
                time.sleep(3)
            
            print("填入【用戶識別碼】", flush=True)
            try:
                username_field = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "aa-uid"))
                )
                username_field.clear()
                username_field.send_keys(username)
            except Exception as e:
                print(f"填入用戶識別碼時出錯: {e}", flush=True)
                print("頁面可能不是登入頁面，嘗試刷新...", flush=True)
                driver.refresh()
                time.sleep(3)
                username_field = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "aa-uid"))
                )
                username_field.clear()
                username_field.send_keys(username)

            print("填入【用戶密碼】", flush=True)
            password_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "aa-passwd"))
            )
            password_field.clear()
            password_field.send_keys(password)

            # 處理驗證碼
            print("開始處理驗證碼...", flush=True)
            if not handle_captcha(driver, register_type, attempt):  # 傳入當前嘗試次數
                print("驗證碼處理失敗，清空輸入欄位，準備重試...", flush=True)
                attempt += 1
                continue

            print("點擊【認證】送出登入資料", flush=True)
            time.sleep(2)  # 等待頁面加載完成，修改為2秒

            # 嘗試多種方式點擊提交按鈕
            try:
                # 方法1: 直接點擊按鈕
                submit_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "submit_hn"))
                )
                submit_button.click()
                print("點擊提交按鈕", flush=True)
            except Exception as e:
                print(f"點擊提交按鈕失敗: {e}", flush=True)
                try:
                    # 方法2: 使用JavaScript執行點擊
                    driver.execute_script("document.getElementById('submit_hn').click();")
                    print("使用JavaScript點擊提交按鈕", flush=True)
                except Exception as js_e:
                    print(f"JavaScript點擊失敗: {js_e}", flush=True)
                    try:
                        # 方法3: 使用a標籤點擊
                        driver.execute_script("document.querySelector('#submit_hn a').click();")
                        print("使用JavaScript點擊a標籤", flush=True)
                    except Exception as a_e:
                        print(f"點擊a標籤失敗: {a_e}", flush=True)
                        try:
                            # 方法4: 直接提交表單
                            driver.execute_script("document.AuthScreen.submit();")
                            print("直接提交AuthScreen表單", flush=True)
                        except Exception as form_e:
                            print(f"表單提交失敗: {form_e}", flush=True)
                            print("無法點擊提交按鈕，將重試...", flush=True)
                            attempt += 1
                            continue

            # 首先檢查是否有驗證碼錯誤的alert (這一步非常關鍵)
            time.sleep(1)  # 等待可能的彈窗
            has_alert, alert_text = handle_alert_with_retry(driver, max_wait=3, max_retries=2)
            if has_alert:
                if "驗證碼錯誤" in alert_text or "認證碼錯誤" in alert_text or "圖形認證碼錯誤" in alert_text:
                    print(f"驗證碼錯誤: {alert_text}，將重試", flush=True)
                    attempt += 1
                    continue  # 重新開始登入流程
                else:
                    print(f"遇到其他警告: {alert_text}", flush=True)

            # 根據謄本類型分別處理
            if register_type == '1':
                # 第一類謄本需要檢查自然人憑證頁面
                try:
                    time.sleep(2)  # 等待頁面加載
                    
                    # 檢查是否顯示自然人憑證輸入頁面
                    if "idNo" in driver.page_source:
                        print("檢測到自然人憑證輸入頁面", flush=True)
                        
                        # 自然人憑證資料輸入
                        print("自然人憑證資料輸入中...", flush=True)
                        
                        # 使用配置中的自然人憑證信息（如果有）
                        if card_cert and card_cert.get("id_no") and card_cert.get("pin"):
                            id_no = card_cert["id_no"]
                            pin = card_cert["pin"]
                            print("使用已保存的自然人憑證信息", flush=True)
                        else:
                            print("未找到保存的自然人憑證信息，請手動輸入", flush=True)
                            from getpass import getpass
                            id_no = input("請輸入身分證號碼: ").strip()
                            pin = getpass("請輸入憑證密碼: ").strip()
                        
                        id_field = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, "idNo"))
                        )
                        id_field.clear()
                        id_field.send_keys(id_no)
                        
                        pin_field = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, "pin"))
                        )
                        pin_field.clear()
                        pin_field.send_keys(pin)
                        
                        submit_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.CLASS_NAME, "btn.btn-primary"))
                        )
                        submit_button.click()
                        print("憑證資料已送出", flush=True)
                        
                        # 重要：檢查提交憑證信息後是否有錯誤 (很可能是驗證碼錯誤)
                        time.sleep(1)
                        has_alert, alert_text = handle_alert_with_retry(driver, max_wait=3, max_retries=2)
                        if has_alert:
                            print(f"憑證信息提交後出現警告: {alert_text}", flush=True)
                            if "驗證碼錯誤" in alert_text or "認證碼錯誤" in alert_text or "圖形認證碼錯誤" in alert_text:
                                attempt += 1
                                continue  # 重新開始登入流程
                    else:
                        print("未檢測到自然人憑證輸入頁面，可能有錯誤", flush=True)
                        
                        # 再次檢查是否有彈窗警告
                        has_alert, alert_text = handle_alert_with_retry(driver, max_wait=3, max_retries=2)
                        if has_alert:
                            print(f"發現警告: {alert_text}", flush=True)
                            attempt += 1
                            continue  # 重新開始登入流程
                except Exception as cert_error:
                    # 處理自然人憑證錯誤
                    print(f"處理自然人憑證時出錯: {cert_error}", flush=True)
                    retry, action = handle_card_reader_error(driver, cert_error)
                    if retry:
                        attempt += 1
                        continue
                    elif action == "2":
                        # 切換到第二類謄本
                        register_type = '2'
                        # 返回主頁重新登入
                        driver.get(f"{base_url}/Home")
                        return login_attempt(driver, '2', username, password, card_cert)
                    elif action == "exit":
                        # 退出程式
                        print("程式即將退出...")
                        sys.exit(0)
                    else:
                        # 處理其他錯誤，繼續嘗試
                        attempt += 1
                        continue
            
            # 處理後續頁面 (對所有謄本類型通用)
            time.sleep(3)  # 等待頁面加載
            
            # 檢查當前URL並輸出，幫助診斷問題
            current_url = driver.current_url
            print(f"當前URL (登入後): {current_url}", flush=True)
            
            # 檢查是否有錯誤代碼 - 如果有則重試
            if "errorCode" in current_url or "aa-result=s341" in current_url:
                print(f"檢測到錯誤代碼在URL中，登入失敗，將重試: {current_url}", flush=True)
                attempt += 1
                continue
            
            # 檢查代理人授權頁面
            if "agent.jsp" in current_url:
                print("檢測到代理人授權確認頁面，點擊「確定」按鈕...", flush=True)
                try:
                    confirm_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'det2_btn') and contains(@onclick, 'send_submit')]"))
                    )
                    confirm_button.click()
                    print("已點擊「確定」按鈕", flush=True)
                    time.sleep(2)  # 增加等待時間
                except Exception as e:
                    print(f"點擊確定按鈕時出錯: {e}", flush=True)
                
                # 更新當前URL
                current_url = driver.current_url
                print(f"點擊確定後的URL: {current_url}", flush=True)
            
            # 檢查個資同意頁面
            if "person.jsp" in current_url:
                print("檢測到個資同意頁面，勾選同意選項並點擊「同意」按鈕...", flush=True)
                try:
                    # 勾選同意選項
                    checkbox = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "check_doc"))
                    )
                    checkbox.click()
                    print("已勾選同意選項", flush=True)
                    
                    # 點擊同意按鈕
                    agree_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "yes_button"))
                    )
                    agree_button.click()
                    print("已點擊「同意」按鈕", flush=True)
                    time.sleep(2)  # 增加等待時間
                except Exception as e:
                    print(f"處理個資同意頁面時出錯: {e}", flush=True)
                
                # 更新當前URL
                current_url = driver.current_url
                print(f"點擊同意後的URL: {current_url}", flush=True)
            
            # 檢查是否成功登入到主頁
            current_url = driver.current_url
            if "/Home" in current_url or "land.nat.gov.tw/Home" in current_url:
                print("\033[93m登入成功！\033[0m", flush=True)
                return True
            else:
                # 再等待一段時間，看頁面是否會自動跳轉
                print("頁面尚未跳轉至主頁，再等待5秒...", flush=True)
                time.sleep(5)
                
                # 再次檢查當前URL
                current_url = driver.current_url
                print(f"再次等待後的URL: {current_url}", flush=True)
                
                if "/Home" in current_url or "land.nat.gov.tw/Home" in current_url:
                    print("\033[93m登入成功！\033[0m", flush=True)
                    return True
                else:
                    print("\033[93m無法自動確認登入狀態，請手動確認是否已成功登入。\033[0m", flush=True)
                    confirm = input("是否已成功登入？(Y/N): ").strip().lower()
                    if confirm == 'y' or confirm == '':
                        print("\033[93m用戶確認已登入成功！\033[0m", flush=True)
                        return True
                    else:
                        print("用戶確認未登入成功，將重試...", flush=True)
                        attempt += 1
                        continue

        except Exception as e:
            print(f"登入過程中出現異常: {e}", flush=True)
            # 檢查是否需要處理任何未捕獲的警告框
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                print(f"發現未處理的警告框: {alert_text}", flush=True)
                alert.accept()
                
                # 如果是驗證碼錯誤，增加嘗試次數並繼續
                if "驗證碼錯誤" in alert_text or "認證碼錯誤" in alert_text or "圖形認證碼錯誤" in alert_text:
                    print("發現驗證碼錯誤，將重試", flush=True)
                    attempt += 1
                    continue
            except:
                pass
            
            # 檢查是否是讀卡機問題
            retry, action = handle_card_reader_error(driver, e, attempt)
            if retry:
                attempt += 1
                continue
            elif action == "2":
                # 切換到第二類謄本
                register_type = '2'
                # 返回主頁重新登入
                driver.get(f"{base_url}/Home")
                return login_attempt(driver, '2', username, password, card_cert)
            elif action == "exit":
                # 退出程式
                print("程式即將退出...")
                sys.exit(0)
            else:
                # 一般錯誤，繼續嘗試
                attempt += 1
                continue

    # 當達到最大嘗試次數後的處理邏輯
    print(f"登入嘗試次數已達上限 ({max_retries} 次)，將在當前頁面進行手動驗證", flush=True)
    return manual_verification_on_current_page(driver, register_type, username, password, card_cert)


def manual_verification_on_current_page(driver, register_type, username, password, card_cert=None):
    """在當前頁面上進行手動驗證，不重新導航"""
    try:
        # First check if we're already on the login page
        current_url = driver.current_url
        print(f"手動驗證前的當前URL: {current_url}", flush=True)
        
        # If we're not on the authentication page, we need to get there
        if "AuthScreen.jsp" not in current_url and "aaav2.hinet.net" not in current_url:
            print("頁面可能已經重定向，嘗試恢復到登入頁面", flush=True)
            
            # Find and click the login button if visible
            try:
                login_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'登入') or contains(text(),'進入系統')]"))
                )
                login_button.click()
                print("已點擊登入/進入系統按鈕", flush=True)
                time.sleep(2)
            except:
                print("找不到登入按鈕，等待用戶手動操作", flush=True)
                # Give the user a chance to manually navigate
                print("\033[93m請手動導航到登入頁面，然後按Enter繼續\033[0m", flush=True)
                input()
        
        # Now we should be on the login page, or the user has manually navigated
        # Proceed with manual verification
        try:
            # Try to find and clear the username field
            username_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "aa-uid"))
            )
            username_field.clear()
            username_field.send_keys(username)
            print("已填入用戶名", flush=True)
            
            # Find and clear the password field
            password_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "aa-passwd"))
            )
            password_field.clear()
            password_field.send_keys(password)
            print("已填入密碼", flush=True)
            
            # Find the captcha field
            captcha_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "aa-captchaID"))
            )
            captcha_field.clear()
            print("\033[93m請手動輸入驗證碼：\033[0m", flush=True)
            captcha = input().strip()
            captcha_field.send_keys(captcha)
            print("已填入驗證碼", flush=True)
            
            # Check for certificate checkbox if this is a first-class document
            if register_type == '1':
                try:
                    checkbox = driver.find_element(
                        By.XPATH, "/html/body/section/div/div[2]/div[3]/form/div[1]/span/input"
                    )
                    if not checkbox.is_selected():
                        checkbox.click()
                        print("已勾選自然人憑證選項", flush=True)
                except:
                    print("找不到自然人憑證勾選框", flush=True)
            
            try:
                time.sleep(1)
                # 方法1: 直接點擊按鈕
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "submit_hn"))
                ).click()
                print("點擊提交按鈕", flush=True)
            except Exception as e:
                print(f"點擊提交按鈕失敗: {e}", flush=True)
                try:
                    # 方法2: 使用JavaScript執行點擊
                    driver.execute_script("document.getElementById('submit_hn').click();")
                    print("使用JavaScript點擊提交按鈕", flush=True)
                except Exception as js_e:
                    print(f"JavaScript點擊失敗: {js_e}", flush=True)
                    try:
                        # 方法3: 使用a標籤點擊
                        driver.execute_script("document.querySelector('#submit_hn a').click();")
                        print("使用JavaScript點擊a標籤", flush=True)
                    except Exception as a_e:
                        print(f"點擊a標籤失敗: {a_e}", flush=True)
                        try:
                            # 方法4: 直接提交表單
                            driver.execute_script("document.AuthScreen.submit();")
                            print("直接提交AuthScreen表單", flush=True)
                        except Exception as form_e:
                            print(f"表單提交失敗: {form_e}", flush=True)
                            print("無法點擊提交按鈕，將重試...", flush=True)
                            attempt += 1
            
            # Wait for processing
            time.sleep(3)
            
            # Handle certificate input if needed
            current_url = driver.current_url
            if "idNo" in driver.page_source and register_type == '1':
                print("檢測到自然人憑證輸入頁面", flush=True)
                
                # Use saved certificate info
                if card_cert and card_cert.get("id_no") and card_cert.get("pin"):
                    id_no = card_cert["id_no"]
                    pin = card_cert["pin"]
                else:
                    print("請輸入自然人憑證資訊:", flush=True)
                    id_no = input("身分證號碼: ").strip()
                    pin = input("憑證密碼: ").strip()
                
                # Enter certificate info
                try:
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "idNo"))
                    ).send_keys(id_no)
                    
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "pin"))
                    ).send_keys(pin)
                    
                    submit_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CLASS_NAME, "btn.btn-primary"))
                    )
                    submit_button.click()
                    print("已提交自然人憑證資訊", flush=True)
                except:
                    print("無法自動填寫自然人憑證資訊，請手動操作", flush=True)
                    print("\033[93m請手動填寫自然人憑證資訊並提交，然後按Enter繼續\033[0m", flush=True)
                    input()
            
            # Process subsequent pages
            time.sleep(3)
            
            # Handle agent confirmation and privacy consent pages
            for check_count in range(3):  # Check a few times in case of slow page transitions
                current_url = driver.current_url
                
                # Check for agent.jsp
                if "agent.jsp" in current_url:
                    print("檢測到代理人授權頁面，點擊確定按鈕", flush=True)
                    try:
                        confirm_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'send_submit')]"))
                        )
                        confirm_button.click()
                        print("已點擊確定按鈕", flush=True)
                    except:
                        print("無法自動點擊確定按鈕，請手動操作", flush=True)
                        print("\033[93m請手動點擊確定按鈕，然後按Enter繼續\033[0m", flush=True)
                        input()
                
                # Check for person.jsp
                if "person.jsp" in current_url:
                    print("檢測到個資同意頁面", flush=True)
                    try:
                        # Check the consent box
                        checkbox = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, "check_doc"))
                        )
                        checkbox.click()
                        print("已勾選同意選項", flush=True)
                        
                        # Click agree button
                        agree_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, "yes_button"))
                        )
                        agree_button.click()
                        print("已點擊同意按鈕", flush=True)
                    except:
                        print("無法自動處理個資同意頁面，請手動操作", flush=True)
                        print("\033[93m請手動勾選同意並點擊同意按鈕，然後按Enter繼續\033[0m", flush=True)
                        input()
                
                # Check if we've reached the home page
                if "/Home" in current_url:
                    print("\033[93m登入成功！\033[0m", flush=True)
                    return True
                
                time.sleep(2)
            
            # Final check for successful login
            current_url = driver.current_url
            if "/Home" in current_url:
                print("\033[93m登入成功！\033[0m", flush=True)
                return True
            else:
                print("\033[93m無法確認是否成功登入，請手動確認\033[0m", flush=True)
                confirm = input("是否已成功登入？(Y/N，預設Y): ").strip().lower()
                if confirm == 'n':
                    print("用戶確認登入失敗", flush=True)
                    return False
                else:
                    print("用戶確認登入成功", flush=True)
                    return True
                
        except Exception as e:
            print(f"手動驗證過程中出錯: {e}", flush=True)
            print("請完全手動操作登入流程", flush=True)
            print("\033[93m請手動完成整個登入流程，成功後按Enter繼續\033[0m", flush=True)
            input()
            
            # Check if login was successful
            current_url = driver.current_url
            if "/Home" in current_url:
                print("\033[93m登入成功！\033[0m", flush=True)
                return True
            else:
                print("\033[93m無法確認是否成功登入，請手動確認\033[0m", flush=True)
                confirm = input("是否已成功登入？(Y/N，預設Y): ").strip().lower()
                if confirm == 'n':
                    return False
                else:
                    return True
            
    except Exception as e:
        print(f"執行手動驗證時發生嚴重錯誤: {e}", flush=True)
        print("\033[93m請完全手動完成登入流程，完成後按Enter\033[0m", flush=True)
        input()
        return True  # Assume success after manual intervention
        
            
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
    user, pwd = get_credentials("qpt_hinet")
    cert_id, cert_pin = get_credentials("qpt_hinet_cert")
    if user and pwd:
        config = {"username": user, "password": pwd}
        config["card_cert"] = ({"id_no": cert_id, "pin": cert_pin}
                                if cert_id and cert_pin else None)
        print("已從 Windows 認證管理員載入帳密", flush=True)
        return config

    # 🔥 1.5 fallback：qpt_hinet 沒設時，借用 hinet 的（兩者通常共用同組 HiNet 帳密）
    user, pwd = get_credentials("hinet")
    if user and pwd:
        config = {"username": user, "password": pwd, "card_cert": None}
        print("ℹ 全功能地籍資料查詢未獨立設定，自動使用「全國地政電子謄本」的帳密", flush=True)
        return config

    # 🔥 2. 嘗試從舊 config.json 遷移
    config_file = os.path.join(BASE_DIR, "config.json")
    if os.path.exists(config_file):
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                old = json.load(f)
            if old.get("username") and old.get("password"):
                print(f"偵測到舊版 config.json，自動遷移到 Windows 認證管理員...", flush=True)
                ch_save("qpt_hinet", old["username"], old["password"])
                cc = old.get("card_cert")
                if isinstance(cc, dict) and cc.get("id_no") and cc.get("pin"):
                    ch_save("qpt_hinet_cert", cc["id_no"], cc["pin"])
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

    # 創建自訂對話框（qpt_hinet 只用第二類謄本，不需自然人憑證欄位）
    class ConfigDialog(tk.Toplevel):
        def __init__(self, parent):
            super().__init__(parent)
            self.title("全功能地籍資料查詢 - 帳號設定")
            self.geometry("450x260")
            self.resizable(False, False)

            # 置中顯示
            self.update_idletasks()
            x = (self.winfo_screenwidth() // 2) - (450 // 2)
            y = (self.winfo_screenheight() // 2) - (260 // 2)
            self.geometry(f'450x260+{x}+{y}')

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

            # 綁定 Enter 鍵：用戶名→密碼→送出（標準表單行為）
            self.username_entry.bind('<Return>', lambda e: self.password_entry.focus_set())
            self.password_entry.bind('<Return>', lambda e: self.ok_clicked())
            self.bind('<Escape>', lambda e: self.cancel_clicked())

        def ok_clicked(self):
            username = self.username_entry.get().strip()
            password = self.password_entry.get().strip()

            if not username or not password:
                messagebox.showerror("錯誤", "用戶名和密碼為必填項目", parent=self)
                return

            self.result = {
                "username": username,
                "password": password,
                "card_cert": None
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
        ch_save("qpt_hinet", config["username"], config["password"])
        cc = config.get("card_cert")
        if isinstance(cc, dict) and cc.get("id_no") and cc.get("pin"):
            ch_save("qpt_hinet_cert", cc["id_no"], cc["pin"])
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

def get_county_url(county):
    """根據縣市名稱獲取對應的入口網址"""
    # 將"台"替換為"臺"以確保正確匹配
    normalized_county = county.replace("台", "臺")
    
    # 縣市對應的URL映射表
    county_urls = {
        "高雄市": "https://pqt-kcgetw.land.nat.gov.tw/",
        "屏東縣": "https://pqt-pthgetw.land.nat.gov.tw/",
        "臺南市": "https://pqt-tainanetw.land.nat.gov.tw/",
        "嘉義市": "https://pqt-chiayietw.land.nat.gov.tw/",
        "嘉義縣": "https://pqt-cyhgetw.land.nat.gov.tw/",
        "雲林縣": "https://pqt-yunlinetw.land.nat.gov.tw/",
        "彰化縣": "https://pqt-chcgetw.land.nat.gov.tw/",
        "南投縣": "https://pqt-nantouetw.land.nat.gov.tw/",
        "臺中市": "https://pqt.taichungland.nat.gov.tw/",
        "苗栗縣": "https://pqt-miaolietw.land.nat.gov.tw/",
        "新竹市": "https://pqt-hccgetw.land.nat.gov.tw/",
        "新竹縣": "https://pqt-hchgetw.land.nat.gov.tw/",
        "基隆市": "https://pqt-klcgetw.land.nat.gov.tw/",
        "宜蘭縣": "https://pqt-e-landetw.land.nat.gov.tw/",
        "花蓮縣": "https://pqt-hletw.land.nat.gov.tw/",
        "臺東縣": "https://pqt-taitungetw.land.nat.gov.tw/",
        "澎湖縣": "https://mqt-penghuetw.land.nat.gov.tw/",
        "金門縣": "https://mqt-kinmenetw.land.nat.gov.tw/",
        "連江縣": "https://mqt-matsuetw.land.nat.gov.tw/"
    }
    
    # 檢查是否為特殊處理的縣市（臺北市、新北市、桃園市）
    special_counties = ["臺北市", "新北市", "桃園市"]
    if normalized_county in special_counties:
        print(f"\033[93m注意：{normalized_county}使用不同的查詢系統(https://www.ttt.nat.gov.tw)，本程式暫不支援。\033[0m", flush=True)
        return None
    
    # 返回縣市對應的URL，如果找不到則返回None
    return county_urls.get(normalized_county)

def navigate_to_county_system(driver, county):
    """導航到指定縣市的地政系統，並返回該縣市的基礎URL"""
    county_url = get_county_url(county)
    if not county_url:
        print(f"無法找到【{county}】的對應網址，請確認縣市名稱是否正確。", flush=True)
        # 嘗試通過地圖頁面選擇
        driver.get("https://www.land.nat.gov.tw/Home/InformationMap")
        print("已導航至【台灣省地圖】頁面，請手動選擇縣市...", flush=True)
        
        # 等待用戶手動選擇
        initial_url = driver.current_url
        WebDriverWait(driver, 120).until(
            lambda d: d.current_url != initial_url
        )
        print("檢測到頁面變化，請繼續手動操作...", flush=True)
        
        # 等待用戶選擇全方位版(光特)
        WebDriverWait(driver, 120).until(
            lambda d: "getw.land.nat.gov.tw" in d.current_url
        )
        print("已成功導航至縣市電傳系統頁面", flush=True)
        
        # 返回當前URL的基礎部分
        current_url = driver.current_url
        base_url_parts = current_url.split('/')
        if len(base_url_parts) >= 3:
            return '/'.join(base_url_parts[:3])  # 返回如 https://pqt-pthgetw.land.nat.gov.tw
        return "https://pqt-kcgetw.land.nat.gov.tw"  # 默認高雄
    
    # 直接導航到縣市系統首頁
    print(f"導航至【{county}】地政系統：{county_url}", flush=True)
    driver.get(county_url)
    
    # 等待並點擊"進入系統"按鈕
    try:
        enter_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '進入系統') or contains(@onclick, 'enterSystem')]"))
        )
        enter_button.click()
        print("已點擊【進入系統】按鈕", flush=True)
    except Exception as e:
        print(f"點擊【進入系統】按鈕時出錯: {e}", flush=True)
        print("請手動點擊【進入系統】按鈕，程式將等待...", flush=True)
        
        # 等待URL變化，判斷用戶是否已手動操作
        initial_url = driver.current_url
        WebDriverWait(driver, 60).until(
            lambda d: d.current_url != initial_url
        )
        print("檢測到頁面變化，繼續執行...", flush=True)
    
    return county_url  # 返回縣市的基礎URL


def main():
    # 清理控制台（適用於 Windows 和其他系統）
    os.system('cls' if os.name == 'nt' else 'clear')
    print(f"～\033[44m【全能地籍資料查詢系統】自動化程式\033[0m～", flush=True)

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
    print("\033[93m１.送件作業（預設）\033[0m", flush=True)
    print("\033[93m０.退出程式\033[0m", flush=True)
    print("（如需修改帳密，請回主程式 → 功能分頁 → 修改帳密）", flush=True)

    while True:
        try:
            user_input = input("請輸入選項（預設為 1）: ").strip()
            if user_input == "":
                scenario = 1
            else:
                scenario = int(user_input)

            if scenario in [0, 1]:
                print(f"您選擇了選項 {scenario}，正在執行對應操作...", flush=True)
                break
            else:
                print("請輸入 0 或 1 作為選項。", flush=True)
        except ValueError:
            print("請輸入有效的整數選項。", flush=True)

    if scenario == 0:
        # 退出程式
        print("\033[92m程式已退出，感謝使用！\033[0m", flush=True)
        import sys
        sys.exit(0)

    # 🔥 帳密來源已改為 Windows 認證管理員，正常情況下這裡不會缺少
    if not username or not password:
        print("\033[91m缺少帳號或密碼，請回主程式 → 功能分頁 → 修改帳密 設定\033[0m", flush=True)
        return main()
        
        # 更新配置
        config["username"] = username
        config["password"] = password

        with open(os.path.join(BASE_DIR, "config.json"), "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        
        print("配置已更新", flush=True)

    # 送件作業（預設）
    county, district, section, land_number_list, gland_number, ID_number, register_type, register_options, ownership_only, gland_number_list, mortgage_only = get_user_input()

    # 🔥 詢問使用者選擇 PDF 儲存模式和目錄
    pdf_save_dir = None

    if custom_dir_path:
        # 詢問儲存模式
        print("\n請選擇 PDF 儲存模式：", flush=True)
        print("\033[093m１. 批次模式：一開始問一次，所有檔案都存到同一目錄（推薦）\033[0m", flush=True)
        print("\033[093m２. 手動模式：每個檔案都個別詢問存檔位置\033[0m", flush=True)
        print("（直接按 Enter 選擇預設選項）", flush=True)

        while True:
            mode_choice = input("請輸入選項編號（1 或 2，預設：1）：").strip()
            if not mode_choice:
                mode_choice = '1'
            if mode_choice in ['1', '2']:
                break
            else:
                print("輸入無效，請重新輸入。", flush=True)

        if mode_choice == '1':
            # 批次模式：詢問一次目錄
            print("\n請選擇 PDF 儲存目錄：", flush=True)
            print("\033[093m１. 預設目錄（下載的謄本）\033[0m", flush=True)
            print(f"\033[093m２. 自訂目錄（{custom_dir_path}）\033[0m", flush=True)
            print("（直接按 Enter 選擇預設選項）", flush=True)

            while True:
                choice = input("請輸入選項編號（1 或 2，預設：1）：").strip()
                if not choice:
                    choice = '1'
                if choice == '1':
                    pdf_save_dir = get_work_folder("下載的謄本")
                    print(f"\n\033[96m✓ 批次模式 - 預設目錄：\033[0m", flush=True)
                    print(f"\033[96m{pdf_save_dir}\033[0m", flush=True)
                    break
                elif choice == '2':
                    pdf_save_dir = custom_dir_path
                    print(f"\n\033[96m✓ 批次模式 - 自訂目錄：\033[0m", flush=True)
                    print(f"\033[96m{pdf_save_dir}\033[0m", flush=True)
                    break
                else:
                    print("輸入無效，請重新輸入。", flush=True)
        else:
            # 手動模式：每次都問
            print(f"\033[92m✓ 手動模式 - 每個檔案都會詢問存檔位置\033[0m", flush=True)
            print(f"\033[93m（注意：目前此模式尚未完全實作，將使用預設目錄）\033[0m", flush=True)
            pdf_save_dir = get_work_folder("下載的謄本")
    else:
        # 沒有自訂目錄，使用預設目錄
        pdf_save_dir = get_work_folder("下載的謄本")
        print(f"\033[92m使用預設目錄：{pdf_save_dir}\033[0m", flush=True)

    driver = setup_driver(pdf_save_dir)
    
    # 顯示列印功能說明
    add_print_keyboard_shortcut(driver)
    
    # 導航到對應縣市的系統
    county_base_url = navigate_to_county_system(driver, county)
    if county_base_url:
        if login_attempt(driver, register_type, username, password, card_cert, county_base_url):
            after_login_success(
                driver, county, district, section, land_number_list,
                gland_number, ID_number, register_type, register_options,
                scenario, custom_dir_path, ownership_only, gland_number_list, mortgage_only
            )
    
    # 在查詢完成後不要退出，而是等待用戶按鍵
    cprint("\n查詢已完成。請檢查網頁結果，完成後按Enter繼續...")
    input()

    print("\033[41m【送件作業】執行完畢，祝您有美好的一天~~\033[0m", flush=True)
    print("", flush=True)  # 空行確保前面內容都輸出

    # 使用亮黃色顯示提示，更容易看到
    yprint("=" * 60)
    yprint("是否關閉瀏覽器窗口？")
    yprint("  輸入 Y 或 y：關閉瀏覽器")
    yprint("  輸入 N 或 n 或直接按 Enter：保持瀏覽器開啟")
    yprint("=" * 60)
    print("請輸入您的選擇: ", end='', flush=True)

    try:
        close_browser_input = input().strip()
        print(f"[除錯] 您輸入的原始值: '{close_browser_input}'", flush=True)
        close_browser = close_browser_input.lower()
        print(f"[除錯] 轉換為小寫後: '{close_browser}'", flush=True)
    except Exception as e:
        print(f"[錯誤] 讀取輸入時發生錯誤: {e}，預設為不關閉 (N)", flush=True)
        close_browser = 'n'

    if close_browser == 'y':
        print("正在關閉瀏覽器...", flush=True)
        try:
            # 🔥 快速關閉:先關閉所有視窗,再 quit
            try:
                # 關閉所有額外視窗
                main_window = driver.current_window_handle
                for handle in driver.window_handles:
                    if handle != main_window:
                        driver.switch_to.window(handle)
                        driver.close()
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
            print(f"關閉瀏覽器時發生錯誤: {e}", flush=True)
    else:
        # 使用者選擇不關閉瀏覽器，程式需要保持運行以維持 driver 物件
        yprint("瀏覽器將保持開啟狀態")
        yprint("您可以繼續在瀏覽器中操作")
        yprint("完成後請按 Enter 鍵結束程式...")
        print("", flush=True)

        try:
            input()  # 等待使用者按 Enter
            print("正在結束程式並關閉瀏覽器...", flush=True)
            driver.quit()
            print("程式已結束", flush=True)
        except Exception as e:
            print(f"結束時發生錯誤: {e}", flush=True)
            try:
                driver.quit()
            except:
                pass

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

    # 提示輸入地號 - 更新提示訊息
    print("請輸入地號 (以逗號','分隔為不同地號批次調閱):", flush=True)
    land_number_input = input()
    # 處理空字符串情況，確保即使沒有輸入也要建立一個列表
    if land_number_input.strip():
        # 現在以逗號分隔不同地號
        land_number_list = land_number_input.split(',')
        land_number_list = [num.strip() for num in land_number_list]  # 清除每個項目的前後空白
        print(f"您輸入的地號是: {', '.join(land_number_list)}", flush=True)
    else:
        land_number_list = []  # 如果沒有輸入，創建空列表
        print("您輸入的地號是: ", flush=True)

    # 提示輸入建號 - 支援多建號輸入
    print("請輸入建號 (以逗號','分隔為不同建號批次調閱):", flush=True)
    gland_number_input = input()
    # 處理空字符串情況，確保即使沒有輸入也要建立一個列表
    if gland_number_input.strip():
        # 現在以逗號分隔不同建號
        gland_number_list = gland_number_input.split(',')
        gland_number_list = [num.strip() for num in gland_number_list]  # 清除每個項目的前後空白
        print(f"您輸入的建號是: {', '.join(gland_number_list)}", flush=True)
    else:
        gland_number_list = []  # 如果沒有輸入，創建空列表
        print("您輸入的建號是: ", flush=True)
        
    # 為了保持與原程式的兼容性，我們同時保留 gland_number 變數（改為空字串或第一個建號值）
    gland_number = gland_number_list[0] if gland_number_list else ""

    # 提示輸入統一編號
    print("請輸入統一編號(本項尚未開放，勿填，請ENTER略過):", flush=True)
    ID_number = input()
    # print(f"您輸入的統一編號是: {ID_number}", flush=True)

    # 提示選擇謄本類型
    register_type = validate_input(
        "請選擇謄本類型：(\033[093m １.第一類謄本(尚未開放，勿選)  ２.第二類謄本\033[0m ): ",
        ["1", "2"]
    )
    print(f"您選擇的謄本類型是: {register_type}", flush=True)

    # 提示選擇複選的謄本類型 - 改進版本
    print("請選擇以下謄本類型（可複選，用【空白】分隔）: ", flush=True)
    print("\033[093m１.登記謄本 (標示部+所有權部+他項權利部)\033[0m", flush=True)
    print("\033[093m２.登記謄本 (僅所有權部)\033[0m", flush=True)
    print("\033[093m３.登記謄本 (他項權利部)\033[0m", flush=True)
    print("\033[093m４.建物測量成果圖\033[0m", flush=True)
    print("請輸入選項號碼(可複選，以【空白】分隔如：1 4，注意 1、2、3 僅能擇一):", flush=True)
    register_options_input = input().split()
    
    # 處理輸入選項，將選項2、3轉換為內部標記，供後續處理使用
    register_options = []
    ownership_only = False
    mortgage_only = False  # 新增：標記為僅查詢他項權利部

    for option in register_options_input:
        if option == "1":
            register_options.append("1")  # 完整登記謄本
        elif option == "2":
            register_options.append("1")  # 也是登記謄本，但增加一個標記
            ownership_only = True  # 標記為僅查詢所有權部
        elif option == "3":
            register_options.append("1")  # 也是登記謄本，但增加一個標記
            mortgage_only = True  # 標記為僅查詢他項權利部
        elif option == "4":
            register_options.append("3")  # 建物測量成果圖
    
    # 顯示選擇的選項
    option_names = []
    if "1" in register_options:
        if ownership_only:
            option_names.append("登記謄本 (僅所有權部)")
        elif mortgage_only:
            option_names.append("登記謄本 (他項權利部)")
        else:
            option_names.append("登記謄本 (完整)")
    if "3" in register_options:
        option_names.append("建物測量成果圖")
    
    print(f"您選擇的選項是: {' '.join(option_names)}", flush=True)
    
    # 加入確認與修改機制
    while True:
        print("\n\033[96m=== 您輸入的資料摘要 ===\033[0m", flush=True)
        print(f"\033[96m1. 縣市: {county}\033[0m", flush=True)
        print(f"\033[96m2. 地區: {district}\033[0m", flush=True)
        print(f"\033[96m3. 地段: {section}\033[0m", flush=True)
        print(f"\033[96m4. 地號: {land_number_list if land_number_list else '無'}\033[0m", flush=True)
        if gland_number_list:
            print(f"\033[96m5. 建號: {', '.join(gland_number_list)}\033[0m", flush=True)
        else:
            print(f"\033[96m5. 建號: {gland_number}\033[0m", flush=True)
        print(f"\033[96m6. 統一編號: {ID_number}\033[0m", flush=True)
        print(f"\033[96m7. 謄本類型: {'第一類謄本' if register_type == '1' else '第二類謄本'}\033[0m", flush=True)
        print(f"\033[96m8. 謄本項目: {', '.join(option_names)}\033[0m", flush=True)
        
        print("\n\033[93m請確認以上資料是否正確？\033[0m", flush=True)
        print("\033[93m0 或直接按Enter：確認無誤，繼續執行\033[0m", flush=True)
        print("\033[93m1-8. 輸入數字修改對應項目\033[0m", flush=True)
        choice = input("請輸入選項: ").strip()
        
        if choice == '0' or choice == '':
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
            print("請重新輸入地號 (以逗號','分隔為不同地號批次調閱):", flush=True)
            land_number_input = input()
            if land_number_input.strip():
                # 以逗號分隔不同地號
                land_number_list = land_number_input.split(',')
                land_number_list = [num.strip() for num in land_number_list]
                print(f"已更新地號為: {', '.join(land_number_list)}", flush=True)
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
                "請重新選擇謄本類型：(\033[093m １.第一類謄本(尚未開放，勿選)   ２.第二類謄本\033[0m ): ",
                ["1", "2"]
            )
            print(f"已更新謄本類型為: {register_type}", flush=True)
        elif choice == '8':
            print("請重新選擇謄本類型（可複選，用【空白】分隔）: ", flush=True)
            print("\033[093m１.登記謄本 (標示部+所有權部+他項權利部)\033[0m", flush=True)
            print("\033[093m２.登記謄本 (僅所有權部)\033[0m", flush=True)
            print("\033[093m３.登記謄本 (他項權利部)\033[0m", flush=True)
            print("\033[093m４.建物測量成果圖\033[0m", flush=True)
            print("請輸入選項號碼(可複選，以【空白】分隔如：1 4，注意 1、2、3 僅能擇一):", flush=True)
            register_options_input = input().split()

            # 重新處理輸入選項
            register_options = []
            ownership_only = False
            mortgage_only = False

            for option in register_options_input:
                if option == "1":
                    register_options.append("1")  # 完整登記謄本
                elif option == "2":
                    register_options.append("1")  # 也是登記謄本，但增加一個標記
                    ownership_only = True  # 標記為僅查詢所有權部
                elif option == "3":
                    register_options.append("1")  # 也是登記謄本，但增加一個標記
                    mortgage_only = True  # 標記為僅查詢他項權利部
                elif option == "4":
                    register_options.append("3")  # 建物測量成果圖

            # 更新顯示的選項名稱
            option_names = []
            if "1" in register_options:
                if ownership_only:
                    option_names.append("登記謄本 (僅所有權部)")
                elif mortgage_only:
                    option_names.append("登記謄本 (他項權利部)")
                else:
                    option_names.append("登記謄本 (完整)")
            if "3" in register_options:
                option_names.append("建物測量成果圖")
            
            print(f"已更新謄本項目為: {', '.join(option_names)}", flush=True)
        else:
            print("\033[91m無效的選項，請輸入0-8之間的數字\033[0m", flush=True)

    return county, district, section, land_number_list, gland_number, ID_number, register_type, register_options, ownership_only, gland_number_list, mortgage_only

def validate_input(prompt, valid_choices):
    while True:
        print(prompt, flush=True)
        choice = input().strip()
        if choice in valid_choices:
            return choice
        else:
            print("輸入無效，請重新輸入。", flush=True)

def handle_multiple_owners(driver, item_type, item_value, number, county, district, section):
    """
    處理所有權部有多個所有權人的情況，逐個點擊並列印詳細信息
    """
    try:
        # 檢查是否顯示所有權人列表
        owners_count_element = None
        try:
            # 嘗試查找"所有權人共X人"的文本
            owners_count_element = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//td[contains(text(), '所有權人共')]"))
            )
        except:
            # 如果找不到，則認為不是列表頁面
            print("未檢測到所有權人列表，可能是單個所有權人或其他頁面", flush=True)
            return False
        
        # 如果找到了"所有權人共X人"的文本，解析出人數
        owners_count_text = owners_count_element.text
        print(f"檢測到所有權人列表: {owners_count_text}", flush=True)
        
        # 查找所有的所有權人連結
        owner_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'treeQuery') and contains(@class, 'class_box')]")
        
        if not owner_links:
            print("未找到任何所有權人連結", flush=True)
            return False
        
        print(f"找到 {len(owner_links)} 個所有權人連結", flush=True)

        # 先列印所有權人列表頁面（總覽）
        print("列印所有權人列表總覽頁面...", flush=True)
        click_print_and_save_pdf(driver, item_type, item_value, number, county, district, section, suffix="_所有權人列表")
        time.sleep(3)  # 等待列印完成

        # 🚀 優化版：使用字典快速配對，避免雙重迴圈
        # 先建立 onclick -> 資料 的映射
        onclick_to_texts = {}  # {onclick: [text1, text2, ...]}

        for link in owner_links:
            text = link.text.strip()
            onclick = link.get_attribute("onclick")
            if onclick:
                if onclick not in onclick_to_texts:
                    onclick_to_texts[onclick] = []
                onclick_to_texts[onclick].append(text)

        # 快速配對：對每個 onclick，找出數字(登記次序)和非數字(姓名)
        owners_data = []
        for onclick, texts in onclick_to_texts.items():
            registration_order = None
            owner_name = None

            for text in texts:
                if text.isdigit():
                    registration_order = text
                else:
                    owner_name = text

            # 只有同時有登記次序和姓名才加入
            if registration_order and owner_name:
                owners_data.append({
                    "registration_order": registration_order,
                    "owner_name": owner_name,
                    "onclick": onclick
                })

        # 依登記次序排序
        owners_data.sort(key=lambda x: int(x['registration_order']))

        print(f"✓ 快速配對完成，共 {len(owners_data)} 個所有權人", flush=True)

        # 若數量超過3個，詢問使用者是否全部調閱
        selected_indices = None
        skip_all_owners = False  # 🔥 新增：標記是否略過全部

        if len(owners_data) > 3:
            print(f"\n\033[93m========================================\033[0m", flush=True)
            print(f"\033[93m檢測到 {len(owners_data)} 個所有權人\033[0m", flush=True)
            print(f"\033[93m========================================\033[0m", flush=True)

            # 顯示所有權人清單供使用者參考
            for i, owner in enumerate(owners_data, 1):
                print(f"  {i}. 登記次序={owner['registration_order']}, 姓名={owner['owner_name']}", flush=True)

            print(f"\n\033[93m請選擇操作：\033[0m", flush=True)
            print(f"\033[093m  [Enter] = 全部調閱（預設）\033[0m", flush=True)
            print(f"\033[093m  [N] = 略過全部（不調閱任何人）\033[0m", flush=True)
            print(f"\033[093m  [序號] = 只調閱指定序號（例如：1,3,5 或 1-3,5 或 2-4）\033[0m", flush=True)
            print(f"\033[93m========================================\033[0m", flush=True)

            user_input = input("\033[93m請輸入選擇: \033[0m").strip()

            # 🔥 處理輸入
            if not user_input:
                # 直接按 Enter，調閱全部
                print("\033[92m✓ 將調閱全部所有權人\033[0m", flush=True)
            elif user_input.upper() == 'N':
                # 輸入 N，略過全部
                skip_all_owners = True
                print("\033[92m✓ 已選擇略過全部所有權人\033[0m", flush=True)
            elif user_input:  # 如果有輸入序號，解析序號
                selected_indices = set()
                parts = user_input.split(',')

                for part in parts:
                    part = part.strip()
                    if '-' in part:  # 範圍選擇，例如 1-3
                        try:
                            start, end = part.split('-')
                            start = int(start.strip())
                            end = int(end.strip())
                            for i in range(start, end + 1):
                                if 1 <= i <= len(owners_data):
                                    selected_indices.add(i - 1)  # 轉換為0-based索引
                        except ValueError:
                            print(f"\033[91m無效的範圍格式: {part}，將略過\033[0m", flush=True)
                    else:  # 單一序號
                        try:
                            idx = int(part)
                            if 1 <= idx <= len(owners_data):
                                selected_indices.add(idx - 1)  # 轉換為0-based索引
                            else:
                                print(f"\033[91m序號 {idx} 超出範圍，將略過\033[0m", flush=True)
                        except ValueError:
                            print(f"\033[91m無效的序號: {part}，將略過\033[0m", flush=True)

                selected_indices = sorted(list(selected_indices))
                print(f"\033[92m✓ 將調閱以下所有權人: {[i+1 for i in selected_indices]}\033[0m", flush=True)

        # 🔥 如果選擇略過全部，直接返回 True (表示已處理,不需要再用普通方式)
        if skip_all_owners:
            print("\033[93m已略過所有權人的調閱\033[0m", flush=True)
            return True

        # 處理每個所有權人
        for idx, owner_info in enumerate(owners_data):
            # 如果有選擇性調閱，檢查當前索引是否在選擇範圍內
            if selected_indices is not None and idx not in selected_indices:
                print(f"略過第 {idx+1} 個所有權人（未選擇）", flush=True)
                continue

            registration_order = owner_info["registration_order"]
            owner_name = owner_info["owner_name"]
            onclick = owner_info["onclick"]
            
            print(f"\n處理第 {idx+1}/{len(owners_data)} 個所有權人: 登記次序={registration_order}, 姓名={owner_name}", flush=True)
            
            # 先檢查我們是否仍在列表頁面 (尋找"所有權人共X人"文本)
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, "//td[contains(text(), '所有權人共')]"))
                )
                print("目前在所有權人列表頁面", flush=True)
                
                # 尋找matching registration_order的連結
                target_link = None
                all_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'treeQuery') and contains(@class, 'class_box')]")
                
                for link in all_links:
                    if link.text.strip() == registration_order:
                        target_link = link
                        break
                
                if target_link:
                    print(f"準備點擊登記次序為 {registration_order} 的連結", flush=True)
                    
                    # 確保元素可見和可點擊
                    driver.execute_script("arguments[0].scrollIntoView(true);", target_link)
                    time.sleep(1)
                    
                    # 點擊連結
                    target_link.click()
                    print(f"已點擊連結: {registration_order}", flush=True)
                else:
                    print(f"找不到登記次序為 {registration_order} 的連結，嘗試使用onclick直接執行", flush=True)
                    try:
                        driver.execute_script(onclick)
                        print(f"已使用JavaScript執行onclick: {onclick}", flush=True)
                    except Exception as js_e:
                        print(f"執行onclick時出錯: {js_e}", flush=True)
                        continue
                
            except Exception as e:
                print(f"在列表頁面操作時出錯: {e}", flush=True)
                print("嘗試使用瀏覽器返回按鈕回到列表頁面...", flush=True)
                
                # 嘗試使用瀏覽器的返回按鈕
                try:
                    driver.back()
                    time.sleep(3)  # 等待頁面加載
                    
                    # 確認是否返回到列表頁面
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//td[contains(text(), '所有權人共')]"))
                        )
                        print("已返回所有權人列表頁面", flush=True)
                        
                        # 繼續下一個循環，將在下一次迭代中嘗試點擊
                        continue
                    except:
                        print("返回後未能找到所有權人列表，跳過剩餘所有權人", flush=True)
                        return False
                except Exception as back_e:
                    print(f"使用瀏覽器返回按鈕時出錯: {back_e}", flush=True)
                    return False
            
            # 等待詳細頁面加載
            time.sleep(3)
            
            # 檢查是否有查詢錯誤
            def check_for_query_error():
                """檢查是否有查詢錯誤訊息，並顯示錯誤內容"""
                try:
                    # 尋找紅色背景的「查詢失敗」標題
                    error_header = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, "//th[contains(@style, 'background: red')]/center/font[contains(text(), '查詢失敗')]"))
                    )
                    
                    # 找到錯誤訊息內容
                    error_message_element = driver.find_element(By.XPATH, "//td[contains(@class, 'right2')]")
                    error_message = error_message_element.text
                    
                    # 顯示錯誤詳細訊息
                    print(f"\033[91m查詢失敗！\033[0m", flush=True)
                    print(f"\033[91m{error_message}\033[0m", flush=True)
                    
                    # 提取錯誤代碼
                    error_code_match = re.search(r'錯誤代碼:\(([^)]+)\)', error_message)
                    if error_code_match:
                        error_code = error_code_match.group(1)
                        print(f"\033[91m錯誤代碼: {error_code}\033[0m", flush=True)
                    
                    return True
                except (NoSuchElementException, TimeoutException):
                    # 如果沒有找到錯誤訊息，表示沒有錯誤
                    return False
            
            # 檢查是否成功加載詳細頁面
            try:
                # 檢查是否不再顯示列表頁面的元素
                WebDriverWait(driver, 3).until_not(
                    EC.presence_of_element_located((By.XPATH, "//td[contains(text(), '所有權人共')]"))
                )
                print("成功加載所有權人詳細頁面", flush=True)
            except:
                print("可能未能成功加載詳細頁面，但將繼續嘗試處理", flush=True)
            
            # 檢查是否有查詢錯誤
            if check_for_query_error():
                print(f"查詢所有權人 {text} 時發生錯誤，跳過列印和保存。", flush=True)
                
                # 嘗試返回列表頁面
                try:
                    driver.back()
                    print("已使用瀏覽器返回功能", flush=True)
                    time.sleep(3)  # 等待返回操作完成
                    continue  # 繼續處理下一個所有權人
                except Exception as back_e:
                    print(f"使用瀏覽器返回按鈕時出錯: {back_e}", flush=True)
                    return False
            
            # 執行列印
            print(f"列印所有權人 {registration_order} {owner_name} 的詳細信息", flush=True)
            
            # 清理姓名中的星號
            cleaned_name = owner_name.replace("＊", "").replace("*", "")
            
            # 組合檔名附加信息
            owner_info_suffix = f"({registration_order}{cleaned_name})"
            
            # 修改click_print_and_save_pdf函數調用，傳入owner_info_suffix
            custom_modify_pdf(driver, item_type, item_value, number, county, district, section, owner_info_suffix)
            
            # 等待列印完成
            time.sleep(3)
            
            # 不使用返回按鈕，而是重新執行查詢來顯示列表
            print("重新執行查詢以顯示所有權人列表...", flush=True)
            try:
                # 找到「所有權部」的頁籤並點擊
                ownership_tab = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'TabbedPanelsTab') and contains(@onclick, \"treeQuery\") and contains(text(), '所有權部')]"))
                )
                ownership_tab.click()
                print("已點擊「所有權部」頁籤，重新載入列表", flush=True)
                time.sleep(3)  # 等待頁面加載
            except Exception as e:
                print(f"無法點擊所有權部頁籤: {e}", flush=True)
                print("嘗試使用其他方法返回列表...", flush=True)
                
                # 嘗試查找返回按鈕作為備選方案
                try:
                    back_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '回上一頁') or contains(text(), '返回') or contains(@class, 'back')]"))
                    )
                    back_button.click()
                    print("已點擊頁面中的返回按鈕", flush=True)
                    time.sleep(3)  # 等待返回操作完成
                except:
                    print("未找到頁面中的返回按鈕，嘗試使用瀏覽器返回", flush=True)
                    try:
                        driver.back()
                        print("已使用瀏覽器返回功能", flush=True)
                        time.sleep(3)  # 等待返回操作完成
                    except Exception as back_e:
                        print(f"使用瀏覽器返回按鈕時出錯: {back_e}", flush=True)
                        print("可能無法返回列表頁面，將跳過剩餘所有權人", flush=True)
                        return False
            
            # 確認是否已返回列表頁面
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//td[contains(text(), '所有權人共')]"))
                )
                print("已成功返回所有權人列表頁面", flush=True)
                time.sleep(2)  # 確保頁面完全加載
            except:
                print("未能確認返回列表頁面，將跳過剩餘所有權人", flush=True)
                return False
        
        return True
    
    except Exception as e:
        print(f"處理多個所有權人時出錯: {e}", flush=True)
        try:
            # 嘗試使用瀏覽器返回功能
            driver.back()
            print("已使用瀏覽器返回功能", flush=True)
            time.sleep(3)
        except:
            print("使用瀏覽器返回功能失敗", flush=True)
        return False

def custom_modify_pdf(driver, item_type, item_value, number, county, district, section, owner_info_suffix=None):
    """
    點擊列印按鈕並保存PDF，支援自定義檔名格式
    
    參數:
    - driver: WebDriver對象
    - item_type: 'L'表示土地，'B'表示建物
    - item_value: 例如 '1'(土地標示部), '03'(土地所有權部) 等
    - number: 地號或建號
    - county: 縣市
    - district: 地區
    - section: 地段
    - owner_info_suffix: 所有權人資訊的後綴，例如 "(0002蔡)"
    """
    try:
        print("嘗試點擊【列印】按鈕...", flush=True)
        
        # 構建基本檔名
        type_name = "土地" if item_type == 'L' else "建物"
        project_name = get_project_name(item_type, item_value)
        
        # 為檔名添加所有權人資訊（如果有）
        if owner_info_suffix:
            file_name = f"{county}{district}{section}-{number}-{project_name}{owner_info_suffix}"
        else:
            file_name = f"{county}{district}{section}-{number}-{project_name}"
        
        # 清理檔名中的非法字符
        file_name = file_name.replace(":", "").replace("/", "").replace("\\", "")
        file_name = file_name.replace("*", "").replace("?", "").replace("\"", "")
        file_name = file_name.replace("<", "").replace(">", "").replace("|", "")
        
        save_dir = get_work_folder("下載的謄本")
        os.makedirs(save_dir, exist_ok=True)
        
        # 記錄查詢前目錄中的文件
        files_before = set(os.listdir(save_dir))
        
        # 使用JavaScript點擊列印按鈕
        try:
            print("執行點擊列印按鈕的JavaScript...", flush=True)
            driver.execute_script("printContent();")
            print("已使用JavaScript執行列印功能", flush=True)
        except Exception as js_e:
            print(f"使用JavaScript執行列印功能時出錯: {js_e}", flush=True)
            try:
                print_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@class='icon_1']/a[contains(@onclick, 'printContent()')]"))
                )
                print_button.click()
                print("\033[93m已點擊【列印】按鈕\033[0m", flush=True)
            except Exception as e:
                print(f"點擊列印按鈕失敗: {e}", flush=True)
                return False
        
        # 等待足夠長的時間讓PDF自動保存完成
        print("等待系統自動處理列印和保存流程（5秒）...", flush=True)
        time.sleep(5)
        
        # 對比查詢前後的文件列表，找出新增的文件
        files_after = set(os.listdir(save_dir))
        new_files = files_after - files_before
        new_pdfs = [f for f in new_files if f.endswith('.pdf')]
        
        if new_pdfs:
            print(f"\n檢測到 {len(new_pdfs)} 個新的PDF文件:", flush=True)
            for pdf in new_pdfs:
                file_path = os.path.join(save_dir, pdf)
                file_size = os.path.getsize(file_path)
                print(f"- {pdf} (大小: {file_size} 字節)", flush=True)
            
            # 如果有多個新PDF，使用創建時間來判斷正確的文件
            if len(new_pdfs) > 1:
                # 對新PDF文件按創建時間排序
                pdf_info = []
                for pdf in new_pdfs:
                    file_path = os.path.join(save_dir, pdf)
                    pdf_info.append({
                        'filename': pdf,
                        'path': file_path,
                        'created_time': os.path.getctime(file_path),
                        'size': os.path.getsize(file_path)
                    })
                
                pdf_info.sort(key=lambda x: x['created_time'])
                
                # 保留第一個PDF，刪除其他的
                correct_pdf = pdf_info[0]
                print(f"\n\033[92m識別到正確的PDF: {correct_pdf['filename']} (創建時間最早)\033[0m", flush=True)
                
                # 重命名正確的PDF
                if correct_pdf['filename'] != file_name + ".pdf":
                    try:
                        new_path = os.path.join(save_dir, file_name + ".pdf")
                        if os.path.exists(new_path):
                            timestamp = time.strftime("%Y%m%d%H%M%S")
                            file_name = f"{file_name}_{timestamp}"
                            new_path = os.path.join(save_dir, file_name + ".pdf")
                        
                        os.rename(correct_pdf['path'], new_path)
                        print(f"\033[92m已將正確的PDF重命名為: {file_name}.pdf\033[0m", flush=True)
                    except Exception as rename_e:
                        print(f"重命名文件時出錯: {rename_e}", flush=True)
                
                # 刪除其他PDF
                for pdf in pdf_info[1:]:
                    try:
                        os.remove(pdf['path'])
                        print(f"\033[93m已刪除多餘的PDF: {pdf['filename']}\033[0m", flush=True)
                    except Exception as remove_e:
                        print(f"刪除文件時出錯: {remove_e}", flush=True)
            else:
                # 只有一個新PDF，直接重命名
                pdf = new_pdfs[0]
                file_path = os.path.join(save_dir, pdf)
                
                if pdf != file_name + ".pdf":
                    try:
                        new_path = os.path.join(save_dir, file_name + ".pdf")
                        if os.path.exists(new_path):
                            timestamp = time.strftime("%Y%m%d%H%M%S")
                            file_name = f"{file_name}_{timestamp}"
                            new_path = os.path.join(save_dir, file_name + ".pdf")
                        
                        os.rename(file_path, new_path)
                        yprint(f"已將PDF重命名為: {file_name}.pdf")
                    except Exception as rename_e:
                        print(f"重命名文件時出錯（嘗試複製再刪除）: {rename_e}", flush=True)
                        try:
                            shutil.copy2(file_path, new_path)
                            print(f"已成功複製為: {file_name}.pdf", flush=True)
                            try:
                                os.remove(file_path)
                                print("已刪除原始被佔用的PDF檔案", flush=True)
                            except Exception as delete_e:
                                print(f"原始PDF無法刪除（略過）: {delete_e}", flush=True)
                        except Exception as copy_e:
                            print(f"複製並重新命名仍失敗，略過該檔案: {copy_e}", flush=True)
        else:
            print("\033[91m未檢測到新的PDF文件\033[0m", flush=True)
        
        return True
    
    except Exception as e:
        print(f"\033[91m整個列印過程出錯: {e}\033[0m", flush=True)
        return False

def download_pdf_from_popup(driver, file_name_prefix):
    """
    從彈出視窗或iframe中下載PDF
    
    參數:
    - driver: WebDriver對象
    - file_name_prefix: 文件名前綴
    
    返回:
    - bool: 是否成功下載
    """
    try:
        # 儲存當前窗口句柄
        original_window = driver.current_window_handle
        
        # 檢查是否有新的視窗
        all_windows = driver.window_handles
        
        # 如果只有一個視窗，檢查是否有iframe
        if len(all_windows) == 1:
            try:
                # 查找可能包含PDF的iframe
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                if iframes:
                    print(f"找到 {len(iframes)} 個iframe，嘗試搜索PDF內容...", flush=True)
                    
                    for i, iframe in enumerate(iframes):
                        try:
                            print(f"切換到iframe {i+1}/{len(iframes)}", flush=True)
                            driver.switch_to.frame(iframe)
                            
                            # 檢查iframe中的embed或object元素
                            pdf_elements = driver.find_elements(By.TAG_NAME, "embed") + \
                                           driver.find_elements(By.TAG_NAME, "object")
                            
                            for pdf_elem in pdf_elements:
                                src = pdf_elem.get_attribute("src") or pdf_elem.get_attribute("data")
                                if src and (".pdf" in src.lower() or "pdf" in src.lower()):
                                    print(f"在iframe中找到PDF鏈接: {src}", flush=True)
                                    
                                    # 下載PDF
                                    try:
                                        response = requests.get(src, timeout=30)
                                        if response.status_code == 200:
                                            save_dir = get_work_folder("下載的謄本")
                                            os.makedirs(save_dir, exist_ok=True)
                                            
                                            save_path = os.path.join(save_dir, file_name_prefix + ".pdf")
                                            if os.path.exists(save_path):
                                                timestamp = time.strftime("%Y%m%d%H%M%S")
                                                save_path = os.path.join(save_dir, file_name_prefix + f"_{timestamp}.pdf")
                                            
                                            with open(save_path, 'wb') as f:
                                                f.write(response.content)
                                            
                                            print(f"\033[92m成功從iframe下載PDF並保存為: {os.path.basename(save_path)}\033[0m", flush=True)
                                            
                                            # 切回主文檔
                                            driver.switch_to.default_content()
                                            return True
                                        else:
                                            print(f"下載PDF失敗，HTTP狀態碼: {response.status_code}", flush=True)
                                    except Exception as dl_e:
                                        print(f"從iframe下載PDF時發生錯誤: {dl_e}", flush=True)
                            
                            # 如果在當前iframe中沒有找到PDF，切回主文檔
                            driver.switch_to.default_content()
                        except Exception as frame_e:
                            print(f"處理iframe {i+1}時出錯: {frame_e}", flush=True)
                            driver.switch_to.default_content()
            except Exception as iframe_e:
                print(f"搜索iframe時出錯: {iframe_e}", flush=True)
                driver.switch_to.default_content()
        
        # 如果有多個視窗，檢查其它視窗中的PDF
        if len(all_windows) > 1:
            print(f"找到 {len(all_windows)} 個視窗，嘗試在新視窗中查找PDF...", flush=True)
            
            for window in all_windows:
                if window != original_window:
                    try:
                        driver.switch_to.window(window)
                        print(f"已切換到新視窗，URL: {driver.current_url}", flush=True)
                        
                        # 檢查URL是否為PDF
                        current_url = driver.current_url
                        if current_url.lower().endswith('.pdf') or 'pdf' in current_url.lower():
                            print(f"檢測到PDF URL: {current_url}", flush=True)
                            
                            # 使用請求庫下載PDF
                            try:
                                response = requests.get(current_url, timeout=30)
                                if response.status_code == 200:
                                    save_dir = get_work_folder("下載的謄本")
                                    os.makedirs(save_dir, exist_ok=True)
                                    
                                    save_path = os.path.join(save_dir, file_name_prefix + ".pdf")
                                    if os.path.exists(save_path):
                                        timestamp = time.strftime("%Y%m%d%H%M%S")
                                        save_path = os.path.join(save_dir, file_name_prefix + f"_{timestamp}.pdf")
                                    
                                    with open(save_path, 'wb') as f:
                                        f.write(response.content)
                                    
                                    print(f"\033[92m成功從URL下載PDF並保存為: {os.path.basename(save_path)}\033[0m", flush=True)
                                    
                                    # 關閉當前視窗並切回原始視窗
                                    driver.close()
                                    driver.switch_to.window(original_window)
                                    return True
                                else:
                                    print(f"下載PDF失敗，HTTP狀態碼: {response.status_code}", flush=True)
                            except Exception as dl_e:
                                print(f"下載PDF時發生錯誤: {dl_e}", flush=True)
                        
                        # 檢查頁面中的PDF元素
                        pdf_elements = driver.find_elements(By.TAG_NAME, "embed") + \
                                       driver.find_elements(By.TAG_NAME, "object") + \
                                       driver.find_elements(By.TAG_NAME, "iframe")
                        
                        for pdf_elem in pdf_elements:
                            src = pdf_elem.get_attribute("src") or pdf_elem.get_attribute("data")
                            if src and (".pdf" in src.lower() or "pdf" in src.lower()):
                                print(f"在新視窗中找到PDF鏈接: {src}", flush=True)
                                
                                # 下載PDF
                                try:
                                    response = requests.get(src, timeout=30)
                                    if response.status_code == 200:
                                        save_dir = get_work_folder("下載的謄本")
                                        os.makedirs(save_dir, exist_ok=True)
                                        
                                        save_path = os.path.join(save_dir, file_name_prefix + ".pdf")
                                        if os.path.exists(save_path):
                                            timestamp = time.strftime("%Y%m%d%H%M%S")
                                            save_path = os.path.join(save_dir, file_name_prefix + f"_{timestamp}.pdf")
                                        
                                        with open(save_path, 'wb') as f:
                                            f.write(response.content)
                                        
                                        print(f"\033[92m成功從視窗中的鏈接下載PDF並保存為: {os.path.basename(save_path)}\033[0m", flush=True)
                                        
                                        # 關閉當前視窗並切回原始視窗
                                        driver.close()
                                        driver.switch_to.window(original_window)
                                        return True
                                    else:
                                        print(f"下載PDF失敗，HTTP狀態碼: {response.status_code}", flush=True)
                                except Exception as dl_e:
                                    print(f"下載PDF時發生錯誤: {dl_e}", flush=True)
                        
                        # 關閉當前視窗
                        driver.close()
                    except Exception as window_e:
                        print(f"處理新視窗時出錯: {window_e}", flush=True)
                        try:
                            driver.close()
                        except:
                            pass
            
            # 切回原始視窗
            driver.switch_to.window(original_window)
        
        return False
    
    except Exception as e:
        print(f"從彈出視窗下載PDF時發生錯誤: {e}", flush=True)
        
        # 確保返回原始視窗
        try:
            driver.switch_to.window(original_window)
        except:
            pass
        
        return False
    
def handle_building_map_download_button(driver, item_type, item_value, number, county, district, section):
    """
    建物測量成果圖 - 使用網頁中的儲存影像檔按鈕
    """
    try:
        print("開始建物測量成果圖下載處理流程...", flush=True)
        
        # 確保我們在建物測量成果圖頁籤
        try:
            print("尋找並點擊【建物測量成果圖】頁籤...", flush=True)
            tabs = driver.find_elements(By.XPATH, "//li[contains(@class, 'TabbedPanelsTab')]")
            found_tab = False
            for tab in tabs:
                tab_text = tab.text.strip()
                if '建物測量成果圖' in tab_text or '建物圖' in tab_text:
                    print(f"找到頁籤: {tab_text}", flush=True)
                    tab.click()
                    print(f"已點擊【{tab_text}】頁籤", flush=True)
                    time.sleep(3)  # 等待頁面加載
                    found_tab = True
                    break
            
            if not found_tab:
                print("未找到建物測量成果圖頁籤，可能已經在該頁籤上", flush=True)
        except Exception as e:
            print(f"處理頁籤時出錯: {e}", flush=True)
        
        # 記錄下載前的文件列表
        save_dir = get_work_folder("下載的謄本")
        os.makedirs(save_dir, exist_ok=True)
        files_before = set(os.listdir(save_dir))
        downloads_dir = os.path.expanduser("~/Downloads")  # 用戶下載目錄
        if os.path.exists(downloads_dir):
            downloads_before = set(os.listdir(downloads_dir))
        else:
            downloads_before = set()
        
        # 尋找並點擊下載按鈕 (根據提供的HTML)
        print("尋找儲存影像檔按鈕...", flush=True)
        
        # 方法1: 使用提供的class和元素名稱精確查找
        download_button = None
        try:
            # 嘗試找div.saveImage內的custom_btn元素
            download_button = driver.find_element(By.CSS_SELECTOR, "div.saveImage custom_btn")
            print("找到儲存影像檔按鈕 (使用CSS選擇器)", flush=True)
        except:
            try:
                # 嘗試找div.saveImage元素
                download_button = driver.find_element(By.CLASS_NAME, "saveImage")
                print("找到儲存影像檔容器 (使用CLASS_NAME)", flush=True)
            except:
                try:
                    # 嘗試使用XPath
                    download_button = driver.find_element(By.XPATH, "//div[@class='saveImage ol-unselectable ol-control']")
                    print("找到儲存影像檔容器 (使用XPath)", flush=True)
                except:
                    print("找不到精確的儲存影像檔按鈕，將嘗試其他方法", flush=True)
        
        # 方法2: 如果找不到精確匹配，使用更寬鬆的選擇器
        if not download_button:
            try:
                # 尋找title包含"儲存影像檔"的任何元素
                download_button = driver.find_element(By.XPATH, "//*[contains(@title, '儲存影像檔')]")
                print("找到儲存影像檔按鈕 (通過title屬性)", flush=True)
            except:
                print("通過title屬性找不到儲存影像檔按鈕", flush=True)
        
        # 方法3: 最後嘗試尋找頁面上所有可能的控制元素
        if not download_button:
            try:
                # 尋找所有ol-control類的元素
                control_elements = driver.find_elements(By.CLASS_NAME, "ol-control")
                print(f"找到 {len(control_elements)} 個控制元素", flush=True)
                
                # 檢查每個控制元素
                for elem in control_elements:
                    try:
                        # 嘗試獲取元素的title或文本
                        title = elem.get_attribute("title") or ""
                        text = elem.text or ""
                        
                        # 檢查是否與儲存相關
                        if "儲存" in title or "下載" in title or "save" in title.lower() or "download" in title.lower() or \
                           "儲存" in text or "下載" in text:
                            download_button = elem
                            print(f"找到可能的儲存按鈕: {title} {text}", flush=True)
                            break
                    except:
                        continue
            except Exception as e:
                print(f"搜索控制元素時出錯: {e}", flush=True)
        
        # 如果找到按鈕，點擊它
        if download_button:
            try:
                # 確保元素可見
                driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
                time.sleep(1)
                
                # 嘗試直接點擊
                print("嘗試點擊儲存影像檔按鈕...", flush=True)
                download_button.click()
                print("已點擊儲存影像檔按鈕", flush=True)
            except Exception as click_e:
                print(f"直接點擊失敗: {click_e}", flush=True)
                try:
                    # 嘗試使用JavaScript點擊
                    print("嘗試使用JavaScript點擊儲存影像檔按鈕...", flush=True)
                    driver.execute_script("arguments[0].click();", download_button)
                    print("已使用JavaScript點擊儲存影像檔按鈕", flush=True)
                except Exception as js_e:
                    print(f"JavaScript點擊失敗: {js_e}", flush=True)
                    # 如果按鈕內有子元素，嘗試點擊子元素
                    try:
                        child_elem = download_button.find_element(By.TAG_NAME, "*")
                        driver.execute_script("arguments[0].click();", child_elem)
                        print("已點擊儲存影像檔按鈕的子元素", flush=True)
                    except:
                        print("找不到或無法點擊儲存影像檔按鈕的子元素", flush=True)
        else:
            print("無法找到儲存影像檔按鈕，將嘗試使用鍵盤模擬右鍵操作", flush=True)
            
            # 嘗試使用鍵盤快捷鍵或右鍵操作
            try:
                # 尋找可能的圖片元素
                map_elements = driver.find_elements(By.TAG_NAME, "img")
                for img in map_elements:
                    if img.is_displayed():
                        img_width = img.get_attribute("width")
                        img_height = img.get_attribute("height")
                        
                        if img_width and img_height:
                            try:
                                w = int(img_width)
                                h = int(img_height)
                                if w > 200 and h > 200:  # 較大的圖片可能是地圖
                                    # 右鍵點擊圖片
                                    print(f"嘗試對較大的圖片執行右鍵操作 (大小: {w}x{h})", flush=True)
                                    actions = ActionChains(driver)
                                    actions.context_click(img).perform()
                                    print("已執行右鍵點擊", flush=True)
                                    break
                            except ValueError:
                                continue
            except Exception as action_e:
                print(f"執行右鍵操作失敗: {action_e}", flush=True)
        
        # 等待下載完成 (10秒)
        print("等待圖片下載完成 (10秒)...", flush=True)
        time.sleep(10)
        
        # 檢查是否有新文件
        files_after = set(os.listdir(save_dir))
        new_files = files_after - files_before
        
        # 也檢查下載目錄
        if os.path.exists(downloads_dir):
            downloads_after = set(os.listdir(downloads_dir))
            new_downloads = downloads_after - downloads_before
        else:
            new_downloads = set()
        
        # 構建目標文件名
        target_filename = f"{county}{district}{section}-{number}-建物測量成果圖.jpg"
        target_filename = target_filename.replace(":", "").replace("/", "").replace("\\", "")
        target_filename = target_filename.replace("*", "").replace("?", "").replace("\"", "")
        target_filename = target_filename.replace("<", "").replace(">", "").replace("|", "")
        target_path = os.path.join(save_dir, target_filename)
        
        # 處理下載的文件
        found_image = False
        
        # 首先檢查當前目錄中的新文件
        new_image_files = [f for f in new_files if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
        if new_image_files:
            print(f"在當前目錄找到 {len(new_image_files)} 個新圖片文件:", flush=True)
            for file in new_image_files:
                print(f"- {file}", flush=True)
            
            # 重命名最新的文件
            if len(new_image_files) > 0:
                # 按修改時間排序，獲取最新的文件
                newest_file = max(new_image_files, key=lambda f: os.path.getmtime(os.path.join(save_dir, f)))
                source_path = os.path.join(save_dir, newest_file)
                
                # 重命名文件
                try:
                    # 檢查目標文件是否已存在
                    if os.path.exists(target_path):
                        timestamp = time.strftime("%Y%m%d%H%M%S")
                        target_path = os.path.join(save_dir, f"{county}{district}{section}-{number}-建物測量成果圖_{timestamp}.jpg")
                    
                    os.rename(source_path, target_path)
                    print(f"已將文件 {newest_file} 重命名為 {os.path.basename(target_path)}", flush=True)
                    found_image = True
                except Exception as rename_e:
                    print(f"重命名文件失敗: {rename_e}", flush=True)
                    # 嘗試複製而非重命名
                    try:
                        shutil.copy2(source_path, target_path)
                        print(f"已複製文件 {newest_file} 為 {os.path.basename(target_path)}", flush=True)
                        found_image = True
                    except Exception as copy_e:
                        print(f"複製文件失敗: {copy_e}", flush=True)
        
        # 如果當前目錄沒有找到，檢查下載目錄
        if not found_image and new_downloads:
            new_download_images = [f for f in new_downloads if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
            if new_download_images:
                print(f"在下載目錄找到 {len(new_download_images)} 個新圖片文件:", flush=True)
                for file in new_download_images:
                    print(f"- {file}", flush=True)
                
                # 使用下載目錄中最新的文件
                if len(new_download_images) > 0:
                    # 按修改時間排序，獲取最新的文件
                    newest_download = max(new_download_images, key=lambda f: os.path.getmtime(os.path.join(downloads_dir, f)))
                    download_path = os.path.join(downloads_dir, newest_download)
                    
                    # 複製文件到目標目錄並重命名
                    try:
                        # 檢查目標文件是否已存在
                        if os.path.exists(target_path):
                            timestamp = time.strftime("%Y%m%d%H%M%S")
                            target_path = os.path.join(save_dir, f"{county}{district}{section}-{number}-建物測量成果圖_{timestamp}.jpg")
                        
                        shutil.copy2(download_path, target_path)
                        print(f"已複製下載文件 {newest_download} 為 {os.path.basename(target_path)}", flush=True)
                        found_image = True
                    except Exception as copy_e:
                        print(f"複製下載文件失敗: {copy_e}", flush=True)
        
        # 如果自動操作失敗，請求用戶手動操作
        if not found_image:
            print("\n\033[93m未檢測到自動下載的圖片文件，請手動操作:\033[0m", flush=True)
            print("1. 請點擊網頁右側的「儲存影像檔」按鈕", flush=True)
            print("2. 若出現右鍵選單，選擇「儲存圖片為...」", flush=True)
            print(f"3. 儲存為 {target_filename}", flush=True)
            
            # 等待用戶確認操作完成
            file_path = input("\n\033[93m請輸入儲存的文件完整路徑，或直接按Enter繼續:\033[0m ")
            
            if file_path and os.path.exists(file_path):
                # 用戶提供了文件路徑
                try:
                    # 複製文件到目標路徑
                    shutil.copy2(file_path, target_path)
                    print(f"已複製文件 {file_path} 為 {target_path}", flush=True)
                    found_image = True
                except Exception as copy_e:
                    print(f"複製文件失敗: {copy_e}", flush=True)
        
        # 最終結果報告
        if found_image:
            print(f"\n\033[92m建物測量成果圖處理完成! 檔案保存為: {os.path.basename(target_path)}\033[0m", flush=True)
            return True
        else:
            print("\n\033[93m建物測量成果圖處理未完成，未能自動儲存檔案\033[0m", flush=True)
            return False
        
    except Exception as e:
        print(f"處理建物測量成果圖時發生錯誤: {e}", flush=True)
        return False

def after_login_success(driver, county, district, section, land_number_list, gland_number, ID_number, register_type, register_options, scenario, custom_dir_path, ownership_only=False, gland_number_list=None, mortgage_only=False):
    # 成功登入並跳轉到 Home 後的操作
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

    def check_for_query_error():
        """檢查是否有查詢錯誤訊息，並顯示錯誤內容"""
        try:
            # 尋找紅色背景的「查詢失敗」標題
            error_header = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//th[contains(@style, 'background: red')]/center/font[contains(text(), '查詢失敗')]"))
            )
            
            # 找到錯誤訊息內容
            error_message_element = driver.find_element(By.XPATH, "//td[contains(@class, 'right2')]")
            error_message = error_message_element.text
            
            # 顯示錯誤詳細訊息
            print(f"\033[91m查詢失敗！\033[0m", flush=True)
            print(f"\033[91m{error_message}\033[0m", flush=True)
            
            # 提取錯誤代碼
            error_code_match = re.search(r'錯誤代碼:\(([^)]+)\)', error_message)
            if error_code_match:
                error_code = error_code_match.group(1)
                print(f"\033[91m錯誤代碼: {error_code}\033[0m", flush=True)
            
            # 常見錯誤代碼提示
            if 'X32' in error_message:
                print("\033[93m提示: 此錯誤表示查無該筆資料，請確認您輸入的查詢條件是否正確。\033[0m", flush=True)
            elif 'X43' in error_message:
                print("\033[93m提示: 此錯誤表示該建號需辦理個資隱匿作業，需洽管轄地政事務所。\033[0m", flush=True)
            
            return True
        except (NoSuchElementException, TimeoutException):
            # 如果沒有找到錯誤訊息，表示沒有錯誤
            return False

    print("轉跳至【全功能電傳系統】", flush=True)
    
    # 將建號字串轉為清單（如果外部沒有傳入 gland_number_list，就自動從 gland_number 轉換）
    if not gland_number_list and gland_number and isinstance(gland_number, str):
        gland_number_list = [num.strip() for num in gland_number.split(',') if num.strip()]

    # 時間限制檢查放在這邊
    if not check_time_limit():
        return False

    handle_alerts(driver)
    
    # 記住已經選擇過的鄉鎮市區和地段，避免重複選擇
    selected_district = None
    selected_section = None

    # 處理縣市選擇
    # print("等待【縣市下拉選單】加載", flush=True)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "country"))
    )
    
    # 選擇縣市 - 將"台"替換為"臺"以確保正確匹配
    normalized_county = county.replace("台", "臺")
    
    # 根據提供的HTML，縣市選擇器現在是class="country"
    time.sleep(2)
    print("展開【縣市下拉選單】", flush=True)
    country_select = driver.find_element(By.CLASS_NAME, "country")
    
    # 獲取縣市下拉選單中的所有選項
    options = Select(country_select).options
    
    # 尋找匹配的縣市
    county_found = False
    for option in options:
        option_text = option.text
        if normalized_county in option_text:
            # print(f"找到匹配的縣市: {option_text}", flush=True)
            Select(country_select).select_by_visible_text(option_text)
            county_found = True
            cprint(f"已選擇縣市: {option_text}")
            break
    
    if not county_found:
        print(f"未找到匹配的縣市: {normalized_county}，請檢查縣市名稱是否正確", flush=True)
        # 可以讓用戶選擇或使用默認值
        print("使用默認縣市選項", flush=True)
        if len(options) > 1:  # 確保至少有一個選項（第一個通常是"請選擇"）
            Select(country_select).select_by_index(1)  # 選擇第二個選項
    
    # 等待鄉鎮市區下拉選單更新
    time.sleep(3)
    
    # 選擇鄉鎮市區
    if district:
        try:
            # print("等待【鄉鎮市區下拉選單】加載", flush=True)
            township_select = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "township"))
            )
            
            # 獲取鄉鎮市區下拉選單中的所有選項
            township_options = Select(township_select).options
            
            # 尋找匹配的鄉鎮市區
            district_found = False
            for option in township_options:
                if district in option.text:
                    Select(township_select).select_by_visible_text(option.text)
                    district_found = True
                    selected_district = option.text  # 記住已選擇的鄉鎮市區
                    cprint(f"已選擇鄉鎮市區: {option.text}")
                    break
            
            if not district_found:
                print(f"未找到匹配的鄉鎮市區: {district}", flush=True)
                # 可以讓用戶手動選擇
                print("請從以下選項中選擇鄉鎮市區:", flush=True)
                for i, option in enumerate(township_options[1:], 1):  # 跳過第一個"請選擇"選項
                    print(f"{i}. {option.text}", flush=True)
                
                choice = input("請輸入選項編號: ").strip()
                try:
                    choice_idx = int(choice)
                    if 1 <= choice_idx <= len(township_options) - 1:
                        Select(township_select).select_by_index(choice_idx)
                        selected_district = township_options[choice_idx].text  # 記住選擇的鄉鎮市區
                        cprint(f"已選擇鄉鎮市區: {township_options[choice_idx].text}")
                    else:
                        print("選項編號無效，使用默認選項", flush=True)
                        if len(township_options) > 1:
                            Select(township_select).select_by_index(1)  # 選擇第二個選項
                            selected_district = township_options[1].text  # 記住選擇的鄉鎮市區
                except ValueError:
                    print("輸入無效，使用默認選項", flush=True)
                    if len(township_options) > 1:
                        Select(township_select).select_by_index(1)  # 選擇第二個選項
                        selected_district = township_options[1].text  # 記住選擇的鄉鎮市區
        except Exception as e:
            print(f"選擇鄉鎮市區時出錯: {e}", flush=True)
    
    # 等待地段下拉選單更新
    time.sleep(3)
    
    # 選擇地段
    if section:
        try:
            # print("等待【地段下拉選單】加載", flush=True)

            # 先檢查是否需要在文字輸入框輸入地段代碼
            sectioncode_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "sectioncode"))
            )
            
            # 使用下拉選單選擇地段
            section_select = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "section"))
            )
            
            # 等待下拉選單加載選項
            WebDriverWait(driver, 10).until(
                lambda d: len(Select(section_select).options) > 1
            )
            
            # 獲取地段下拉選單中的所有選項
            section_options = Select(section_select).options
            
            # 尋找匹配的地段
            section_found = False
            for option in section_options:
                if section in option.text:
                    Select(section_select).select_by_visible_text(option.text)
                    section_found = True
                    selected_section = option.text  # 記住已選擇的地段
                    cprint(f"已選擇地段: {option.text}")
                    break
            
            if not section_found:
                print(f"未找到匹配的地段: {section}", flush=True)
                # 列出所有可選地段
                print("請從以下選項中選擇地段:", flush=True)
                for i, option in enumerate(section_options[1:], 1):  # 跳過第一個"請選擇"選項
                    print(f"{i}. {option.text}", flush=True)
                
                choice = input("請輸入選項編號: ").strip()
                try:
                    choice_idx = int(choice)
                    if 1 <= choice_idx <= len(section_options) - 1:
                        Select(section_select).select_by_visible_text(section_options[choice_idx].text)
                        selected_section = section_options[choice_idx].text  # 記住選擇的地段
                        cprint(f"已選擇地段: {section_options[choice_idx].text}")
                    else:
                        print("選項編號無效，使用默認選項", flush=True)
                except ValueError:
                    print("輸入無效，使用默認選項", flush=True)
        except Exception as e:
            print(f"選擇地段時出錯: {e}", flush=True)
    
    # 處理土地/建物選擇和所有地號/建號的查詢
    
    # 解析要查詢的所有地號
    all_land_numbers = []
    if land_number_list:
        for land_entry in land_number_list:
            # 移除空白字符後加入列表
            entry = land_entry.strip()
            if entry:  # 確保不是空字符
                all_land_numbers.append(entry)
    
    # 列印出解析的地號，便於確認
    if all_land_numbers:
        print(f"解析後的地號列表: {', '.join(all_land_numbers)}", flush=True)
    else:
        print("沒有輸入地號", flush=True)
    
    # 建立要查詢的項目
    query_items = []
    
    # 檢查是否有選擇登記謄本（選項1）
    has_registration = "1" in register_options
    # 檢查是否有選擇地籍圖謄本（選項2）
    has_land_map = "2" in register_options
    # 檢查是否有選擇建物測量成果圖（選項3）
    has_building_map = "3" in register_options
    
    # 新增：詢問是否只查詢所有權部或他項權利部
    # ownership_only = False
    # mortgage_only = False
    # print(f"ownership_only={ownership_only}, mortgage_only={mortgage_only} ←來自主程式傳入的值", flush=True)
    if has_registration:
        if ownership_only:
            yprint("已選擇僅查詢所有權部")
        elif mortgage_only:
            cprint("已選擇僅查詢他項權利部")
        else:
            print("\033[92m將查詢所有部分（標示部、所有權部、他項權利部）\033[0m", flush=True)
    
    # 🔥 優先處理建物查詢（如果有建號列表）- 處理每個建號
    # 原因：大樓案件先調閱建號,可在建物所有權部看到土地資訊,避免直接調閱有數百個所有權人的土地謄本

    # 🔥 新增：詢問是否從建物所有權部調閱土地
    skip_land_query_from_list = False  # 標記是否跳過後續的地號查詢

    if gland_number_list and all_land_numbers:
        print("\033[93m\n========================================\033[0m", flush=True)
        print("\033[93m💡 檢測到同時有地號和建號\033[0m", flush=True)
        print("\033[93m========================================\033[0m", flush=True)
        print(f"  地號: {', '.join(all_land_numbers)}", flush=True)
        print(f"  建號: {', '.join(gland_number_list)}", flush=True)
        print("\033[93m\n調閱方式選擇：\033[0m", flush=True)
        print("\033[093m  [1] = 從建物所有權部調閱土地（推薦）\033[0m", flush=True)
        print("\033[093m      → 先調建號,再從建物所有權部的「土地登記次序」調土地\033[0m", flush=True)
        print("\033[093m      → 避免處理大量所有權人\033[0m", flush=True)
        print("\033[093m  [2] = 按原邏輯調閱（不推薦）\033[0m", flush=True)
        print("\033[093m      → 調完建號後,再調地號（可能有數百個所有權人）\033[0m", flush=True)
        print("\033[93m========================================\033[0m", flush=True)

        user_choice = input("\033[93m請輸入選擇 [1/2, 預設=1]: \033[0m").strip()

        if not user_choice or user_choice == '1':
            print("\033[92m✓ 已選擇：從建物所有權部調閱土地\033[0m", flush=True)
            skip_land_query_from_list = True
        else:
            print("\033[92m✓ 已選擇：按原邏輯調閱（建號→地號）\033[0m", flush=True)
            skip_land_query_from_list = False

    if gland_number_list:
        # 處理多個建號
        for gnum in gland_number_list:
            if gnum and gnum.strip():
                if has_registration:  # 登記謄本
                    if ownership_only:
                        # 如果只查詢所有權部，則只添加建物所有權部的查詢項目
                        query_items.append(('B', '09', gnum.strip()))  # 建物所有權部
                    elif mortgage_only:
                        # 如果只查詢他項權利部，則只添加建物他項權利部的查詢項目
                        query_items.append(('B', '0B', gnum.strip()))  # 建物他項權利部
                    else:
                        # 如果查詢所有部分，則添加所有三個部分的查詢項目
                        query_items.append(('B', '7', gnum.strip()))  # 建物標示部
                        query_items.append(('B', '09', gnum.strip()))  # 建物所有權部
                        query_items.append(('B', '0B', gnum.strip()))  # 建物他項權利部

                if has_building_map:  # 建物測量成果圖 - 放在建物登記謄本之後
                    query_items.append(('B', 'C', gnum.strip()))
    else:
        # 如果沒有傳入多建號列表，使用舊的單一建號邏輯
        if gland_number and gland_number.strip():
            if has_registration:  # 登記謄本
                if ownership_only:
                    # 如果只查詢所有權部，則只添加建物所有權部的查詢項目
                    query_items.append(('B', '09', gland_number.strip()))  # 建物所有權部
                elif mortgage_only:
                    # 如果只查詢他項權利部，則只添加建物他項權利部的查詢項目
                    query_items.append(('B', '0B', gland_number.strip()))  # 建物他項權利部
                else:
                    # 如果查詢所有部分，則添加所有三個部分的查詢項目
                    query_items.append(('B', '7', gland_number.strip()))  # 建物標示部
                    query_items.append(('B', '09', gland_number.strip()))  # 建物所有權部
                    query_items.append(('B', '0B', gland_number.strip()))  # 建物他項權利部

            if has_building_map:  # 建物測量成果圖 - 放在建物登記謄本之後
                query_items.append(('B', 'C', gland_number.strip()))

    # 🔥 之後才處理土地查詢（建號查詢完成後）
    # 如果選擇從建物所有權部調閱土地,則跳過這裡的地號查詢
    if all_land_numbers and not skip_land_query_from_list:
        for land_number in all_land_numbers:
            if has_land_map:  # 地籍圖謄本
                query_items.append(('L', 'D', land_number))

            if has_registration:  # 登記謄本
                if ownership_only:
                    # 如果只查詢所有權部，則只添加所有權部的查詢項目
                    query_items.append(('L', '03', land_number))  # 土地所有權部
                elif mortgage_only:
                    # 如果只查詢他項權利部，則只添加他項權利部的查詢項目
                    query_items.append(('L', '05', land_number))  # 土地他項權利部
                else:
                    # 如果查詢所有部分，則添加所有三個部分的查詢項目
                    query_items.append(('L', '1', land_number))  # 土地標示部
                    query_items.append(('L', '03', land_number))  # 土地所有權部
                    query_items.append(('L', '05', land_number))  # 土地他項權利部

    # 如果沒有任何查詢項目，添加一個默認的土地標示部查詢
    if not query_items:
        query_items.append(('L', '1', ''))  # 默認空地號的土地標示部查詢

    # 儲存初始查詢頁面的URL，方便後續返回使用
    initial_query_url = driver.current_url
    # print(f"保存初始查詢頁面URL: {initial_query_url}", flush=True)

    # 列印所有將要執行的查詢項目，便於用戶確認
    print("\n【查詢清單】總共 {} 項查詢:".format(len(query_items)), flush=True)
    for idx, (item_type, item_value, number) in enumerate(query_items):
        item_type_str = "建物" if item_type == 'B' else "土地"
        project_name = get_project_name(item_type, item_value)
        cprint(f"  {idx+1}. {item_type_str} {project_name} - {number}")
    print("")  # 空行
    
    # 依序執行所有查詢
    for idx, (item_type, item_value, number) in enumerate(query_items):
        try:
            print(f"\n開始第 {idx+1}/{len(query_items)} 筆查詢: {'建物' if item_type == 'B' else '土地'} - {number}", flush=True)
            
            # 確保回到查詢頁面 - 但不刷新頁面以保留已選擇的值
            if driver.current_url != initial_query_url:
                print("返回查詢頁面...", flush=True)
                driver.get(initial_query_url)
                time.sleep(3)  # 等待頁面加載
                
                # 如果回到查詢頁面後，可能需要重新選擇之前的鄉鎮市區和地段
                if selected_district:
                    try:
                        township_select = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CLASS_NAME, "township"))
                        )
                        Select(township_select).select_by_visible_text(selected_district)
                        print(f"重新選擇鄉鎮市區: {selected_district}", flush=True)
                        time.sleep(1)
                    except Exception as e:
                        print(f"重新選擇鄉鎮市區時出錯: {e}", flush=True)
                
                if selected_section:
                    try:
                        section_select = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CLASS_NAME, "section"))
                        )
                        Select(section_select).select_by_visible_text(selected_section)
                        print(f"重新選擇地段: {selected_section}", flush=True)
                        time.sleep(1)
                    except Exception as e:
                        print(f"重新選擇地段時出錯: {e}", flush=True)
            
            # 選擇土地/建物單選按鈕
            if item_type == 'L':  # 土地
                land_radio = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "RLAND"))
                )
                if not land_radio.is_selected():
                    land_radio.click()
                    print("已選擇【土地】查詢", flush=True)
                    time.sleep(2)  # 增加等待時間
            else:  # 建物
                build_radio = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "RBUILD"))
                )
                if not build_radio.is_selected():
                    build_radio.click()
                    print("已選擇【建物】查詢", flush=True)
                    time.sleep(2)  # 增加等待時間
            
            # 等待選項更新
            time.sleep(1)
            
            # 選擇查詢項目
            if item_type == 'L':  # 土地
                # 選擇土地查詢項目
                project_l_select = Select(driver.find_element(By.ID, "projectL"))
                project_l_select.select_by_value(item_value)
                print(f"已選擇【{get_project_name('L', item_value)}】", flush=True)
            else:  # 建物
                # 選擇建物查詢項目
                project_b_select = Select(driver.find_element(By.ID, "projectB"))
                project_b_select.select_by_value(item_value)
                print(f"已選擇【{get_project_name('B', item_value)}】", flush=True)
            
            # 等待選項變更生效
            time.sleep(2)
            
            # 確保輸入框已清空後再輸入新值
            try:
                # 使用JavaScript清空和重新填入值，確保跳過可能的事件攔截
                driver.execute_script("document.getElementById('number').value = '';")
                
                # 然後填入地號/建號
                number_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "number"))
                )
                number_input.clear()  # 再次確保清空
                number_input.send_keys(number)
                
                # 使用JavaScript再次確認值已設置
                driver.execute_script(f"document.getElementById('number').value = '{number}';")
                cprint(f"已輸入{'建號' if item_type == 'B' else '地號'}: {number}")
            except Exception as e:
                print(f"輸入地號/建號時出錯: {e}", flush=True)
                try:
                    # 使用另一種方式嘗試輸入
                    driver.execute_script(f"$('#number').val('{number}');")
                    print(f"使用jQuery設置{'建號' if item_type == 'B' else '地號'}: {number}", flush=True)
                except Exception as je:
                    print(f"jQuery設置值時出錯: {je}", flush=True)
            
            # 如果有統一編號且查詢所有權部，輸入統一編號
            if ID_number and ID_number.strip() and ('03' in item_value or '09' in item_value):
                try:
                    # 檢查統一編號區塊是否顯示
                    owner_div = driver.find_element(By.ID, "owner_div")
                    owner_display_style = owner_div.get_attribute("style")
                    
                    if "display: none" in owner_display_style:
                        print("統一編號區塊未顯示，嘗試啟用...", flush=True)
                        # 嘗試使用JavaScript顯示
                        driver.execute_script("$('#owner_div').show();")
                        print("已嘗試顯示統一編號區塊", flush=True)
                    
                    # 確保選擇「統一編號」選項
                    owner_radio = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "onwer"))
                    )
                    if not owner_radio.is_selected():
                        owner_radio.click()
                        print("已選擇【統一編號】選項", flush=True)
                    
                    # 輸入統一編號
                    owner_code_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "onwer_code"))
                    )
                    owner_code_input.clear()
                    owner_code_input.send_keys(ID_number.strip())
                    # 使用JavaScript再次確認值
                    driver.execute_script(f"document.getElementById('onwer_code').value = '{ID_number.strip()}';")
                    print(f"已輸入統一編號: {ID_number.strip()}", flush=True)
                except Exception as e:
                    print(f"輸入統一編號時出錯: {e}", flush=True)# 增加點擊查詢按鈕前的等待時間，確保輸入完成
            time.sleep(2)
            
            # 如果是第一類謄本，可能需要選擇謄本類型
            if register_type == '1' and ('03' in item_value or '09' in item_value):
                try:
                    # 嘗試查找並選擇第一類謄本
                    cltype_div = driver.find_element(By.ID, "cltype_div")
                    if "display: none" not in cltype_div.get_attribute("style"):
                        radio_button = driver.find_element(By.ID, "RadioGroup2_0")
                        if not radio_button.is_selected():
                            radio_button.click()
                            print("已選擇【第一類謄本】", flush=True)
                            time.sleep(1)
                except Exception as e:
                    print(f"選擇謄本類型時出錯: {e}", flush=True)
            
            # 等待可能的遮罩消失
            try:
                WebDriverWait(driver, 5).until_not(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".blockUI.blockOverlay"))
                )
            except:
                # 如果超時或找不到元素，繼續執行
                pass
            
            # 點擊查詢按鈕
            try:
                search_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'landQuery()')]"))
                )
                search_button.click()
                print("已點擊【查詢】按鈕", flush=True)
            except Exception as e:
                print(f"點擊查詢按鈕時出錯: {e}", flush=True)
                # 嘗試使用JavaScript點擊
                try:
                    driver.execute_script("landQuery();")
                    print("已使用JavaScript執行查詢", flush=True)
                except Exception as js_e:
                    print(f"JavaScript執行查詢失敗: {js_e}", flush=True)
            
            # 等待查詢結果並處理可能的彈窗
            time.sleep(2)
            
            # 處理「是否要送出一類資料」的確認彈窗
            if register_type == '1':
                try:
                    # 檢查是否有彈窗
                    has_alert, alert_text = handle_alert_with_retry(driver, max_wait=2, max_retries=2)
                    if has_alert and "一類資料" in alert_text:
                        print(f"自動接受彈窗: {alert_text}", flush=True)
                        # 在handle_alert_with_retry中已接受
                except Exception as e:
                    print(f"處理確認彈窗時出錯: {e}", flush=True)
            
            # 檢查是否有任何彈窗提示缺少輸入值
            try:
                has_alert, alert_text = handle_alert_with_retry(driver, max_wait=2, max_retries=2)
                if has_alert:
                    if "請輸入地號" in alert_text or "請輸入建號" in alert_text:
                        print(f"警告: {alert_text} - 重新嘗試填入", flush=True)
                        # 重新填入值
                        number_input = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, "number"))
                        )
                        number_input.clear()
                        number_input.send_keys(number)
                        print(f"重新填入{'建號' if item_type == 'B' else '地號'}: {number}", flush=True)
                        
                        # 再次嘗試查詢
                        try:
                            search_button = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'landQuery()')]"))
                            )
                            search_button.click()
                            print("再次點擊【查詢】按鈕", flush=True)
                        except:
                            driver.execute_script("landQuery();")
                            print("再次使用JavaScript執行查詢", flush=True)
                    elif "請選擇鄉鎮" in alert_text:
                        print(f"警告: {alert_text} - 嘗試重新選擇鄉鎮市區", flush=True)
                        # 確保鄉鎮市區已選擇
                        if selected_district:
                            try:
                                township_select = WebDriverWait(driver, 10).until(
                                    EC.presence_of_element_located((By.CLASS_NAME, "township"))
                                )
                                Select(township_select).select_by_visible_text(selected_district)
                                print(f"重新選擇鄉鎮市區: {selected_district}", flush=True)
                                time.sleep(1)
                                
                                # 再次嘗試查詢
                                time.sleep(2)
                                try:
                                    search_button = WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'landQuery()')]"))
                                    )
                                    search_button.click()
                                    print("再次點擊【查詢】按鈕", flush=True)
                                except:
                                    driver.execute_script("landQuery();")
                                    print("再次使用JavaScript執行查詢", flush=True)
                            except Exception as se:
                                print(f"重新選擇鄉鎮市區時出錯: {se}", flush=True)
                    else:
                        print(f"彈窗訊息: {alert_text}", flush=True)
            except:
                # 沒有彈窗或處理失敗，繼續執行
                pass
            
            # 等待查詢結果加載
            time.sleep(3)
            
            # 檢查是否有查詢錯誤
            if check_for_query_error():
                print(f"查詢【{get_project_name(item_type, item_value)}】時發生錯誤，跳過列印和保存。", flush=True)
                # 如果有錯誤，直接跳到下一個查詢項目
                continue
            
            # 檢查是否有錯誤訊息
            try:
                error_msg = driver.find_element(By.CLASS_NAME, "errorMsg")
                if error_msg.is_displayed():
                    print(f"查詢出錯: {error_msg.text}", flush=True)
                    # 如果有錯誤訊息，直接跳到下一個查詢項目
                    continue
            except:
                # 沒有錯誤訊息，繼續執行
                pass
            
            # 檢查是否是登記謄本查詢且查詢成功
            if "1" in register_options and (
                '1' in item_value or '03' in item_value or 
                '05' in item_value or '7' in item_value or
                '09' in item_value or '0B' in item_value
            ):
                print("\n\033[93m檢測到登記謄本查詢，開始處理...\033[0m", flush=True)
                time.sleep(3)  # 等待查詢結果完全加載
                
                # 如果是他項權利部，先檢查是否有多個權利人
                if '05' in item_value or '0B' in item_value:
                    # 嘗試處理多個他項權利
                    if handle_multiple_rights(driver, item_type, item_value, number, county, district, section):
                        print("已完成多個他項權利的處理", flush=True)
                    else:
                        # 如果不是多個權利人或處理失敗，使用普通方式列印
                        print("使用普通方式處理他項權利部", flush=True)
                        custom_modify_pdf(driver, item_type, item_value, number, county, district, section)
                # 新增: 如果是所有權部，也檢查是否有多個所有權人
                elif '03' in item_value or '09' in item_value:
                    # 嘗試處理多個所有權人
                    if handle_multiple_owners(driver, item_type, item_value, number, county, district, section):
                        print("已完成多個所有權人的處理", flush=True)

                        # 🔥 如果是建物所有權部 ('09')，檢查是否有土地登記次序資訊
                        if '09' in item_value:
                            try:
                                extract_and_query_land_registration(driver, county, district, section, custom_save_dir=custom_dir_path)
                            except Exception as land_e:
                                print(f"抓取土地登記次序時發生錯誤: {land_e}", flush=True)
                    else:
                        # 如果不是多個所有權人或處理失敗，使用普通方式列印
                        print("使用普通方式處理所有權部", flush=True)
                        custom_modify_pdf(driver, item_type, item_value, number, county, district, section)

                        # 🔥 如果是建物所有權部 ('09')，檢查是否有土地登記次序資訊
                        if '09' in item_value:
                            try:
                                extract_and_query_land_registration(driver, county, district, section, custom_save_dir=custom_dir_path)
                            except Exception as land_e:
                                print(f"抓取土地登記次序時發生錯誤: {land_e}", flush=True)
                else:
                    # 其他類型的謄本，直接列印
                    custom_modify_pdf(driver, item_type, item_value, number, county, district, section)
                
                time.sleep(2)  # 等待保存操作完成
            
            # 檢查是否是建物測量成果圖且查詢成功
            elif "3" in register_options and 'C' in item_value:
                print("\n\033[93m檢測到建物測量成果圖查詢，開始處理...\033[0m", flush=True)
                time.sleep(3)  # 等待結果加載
                
                # 檢查是否有錯誤訊息
                if not check_for_query_error():
                    # 沒有錯誤，使用新的下載按鈕處理方法
                    handle_building_map_download_button(driver, item_type, item_value, number, county, district, section)
                    time.sleep(2)  # 等待保存操作完成
                else:
                    print("建物測量成果圖查詢出錯，跳過列印", flush=True)
                    
            # 檢查是否是地籍圖謄本且查詢成功
            elif "2" in register_options and 'D' in item_value:
                print("\n\033[93m檢測到地籍圖謄本查詢，開始處理...\033[0m", flush=True)
                time.sleep(3)  # 等待結果加載
                
                # 檢查是否有錯誤訊息
                if not check_for_query_error():
                    # 沒有錯誤，執行列印操作
                    click_print_and_save_pdf(driver, item_type, item_value, number, county, district, section)
                    time.sleep(2)  # 等待保存操作完成
                else:
                    print("地籍圖謄本查詢出錯，跳過列印", flush=True)
            
            # 如果是最後一個查詢項目且scenario=2，下載所有文件
            if idx == len(query_items) - 1 and scenario == 2:
                download_document(driver, scenario, custom_dir_options=custom_dir_path)
            
        except Exception as e:
            print(f"處理查詢項目時出錯: {e}", flush=True)
            # 繼續處理下一個項目
            continue
    
    return True


def click_print_and_save_pdf(driver, item_type, item_value, number, county, district, section, suffix="", custom_save_dir=None):
    """
    只點擊列印按鈕，讓系統自動處理PDF保存過程
    同時檢查是否有查詢錯誤

    參數:
        suffix: 檔名後綴，例如 "_他項權利列表"
        custom_save_dir: 自訂儲存目錄，如果為 None 則使用預設目錄
    """
    try:
        # 首先檢查是否有查詢錯誤
        try:
            # 尋找紅色背景的「查詢失敗」標題
            error_header = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.XPATH, "//th[contains(@style, 'background: red')]/center/font[contains(text(), '查詢失敗')]"))
            )

            # 找到錯誤訊息內容
            error_message_element = driver.find_element(By.XPATH, "//td[contains(@class, 'right2')]")
            error_message = error_message_element.text

            # 顯示錯誤詳細訊息
            print(f"\033[91m查詢失敗，跳過列印！\033[0m", flush=True)
            print(f"\033[91m{error_message}\033[0m", flush=True)

            # 提取錯誤代碼
            error_code_match = re.search(r'錯誤代碼:\(([^)]+)\)', error_message)
            if error_code_match:
                error_code = error_code_match.group(1)
                print(f"\033[91m錯誤代碼: {error_code}\033[0m", flush=True)

            return False  # 有錯誤，不進行列印
        except (NoSuchElementException, TimeoutException):
            # 沒有找到錯誤訊息，表示查詢成功，繼續列印
            pass

        print("嘗試點擊【列印】按鈕...", flush=True)

        # 构建文件名
        type_name = "土地" if item_type == 'L' else "建物"
        project_name = get_project_name(item_type, item_value)
        file_name = f"{county}{district}{section}-{number}-{project_name}{suffix}"
        file_name = file_name.replace(":", "").replace("/", "").replace("\\", "")
        file_name = file_name.replace("*", "").replace("?", "").replace("\"", "")
        file_name = file_name.replace("<", "").replace(">", "").replace("|", "")

        # 🔥 使用自訂目錄或預設目錄
        if custom_save_dir:
            save_dir = custom_save_dir
            print(f"📁 使用自訂目錄: {save_dir}", flush=True)
        else:
            save_dir = get_work_folder("下載的謄本")
        os.makedirs(save_dir, exist_ok=True)
        
        # 記錄查詢前目錄中的文件
        files_before = set(os.listdir(save_dir))
        print(f"📁 列印前資料夾中有 {len(files_before)} 個檔案", flush=True)

        # 🔥 先執行 printContent() 準備列印內容，然後用 CDP 列印
        try:
            print("準備列印內容...", flush=True)

            # 先執行 printContent() 讓網頁準備好列印內容
            driver.execute_script("printContent();")
            time.sleep(2)  # 等待列印內容準備完成

            print("使用 CDP 強制列印為 PDF...", flush=True)

            # 構建完整的檔案路徑
            pdf_file_path = os.path.join(save_dir, file_name + ".pdf")

            # 使用 CDP 的 Page.printToPDF 命令
            print_options = {
                'landscape': False,
                'displayHeaderFooter': False,
                'printBackground': True,
                'preferCSSPageSize': True,
            }

            # 執行列印
            result = driver.execute_cdp_cmd('Page.printToPDF', print_options)

            # 解碼 base64 的 PDF 數據
            import base64
            pdf_data = base64.b64decode(result['data'])

            # 寫入檔案
            with open(pdf_file_path, 'wb') as f:
                f.write(pdf_data)

            print(f"\033[92m✓ PDF 已成功儲存: {file_name}.pdf\033[0m", flush=True)
            print(f"   存檔路徑: {pdf_file_path}", flush=True)
            return True

        except Exception as cdp_e:
            print(f"\033[91mCDP 列印失敗: {cdp_e}\033[0m", flush=True)
            print("嘗試使用傳統方式列印...", flush=True)

            # 備用方案：使用JavaScript點擊列印按鈕
            try:
                print("執行點擊列印按鈕的JavaScript...", flush=True)
                driver.execute_script("printContent();")
                print("已使用JavaScript執行列印功能", flush=True)
            except Exception as js_e:
                print(f"使用JavaScript執行列印功能時出錯: {js_e}", flush=True)
                try:
                    print_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[@class='icon_1']/a[contains(@onclick, 'printContent()')]"))
                    )
                    print_button.click()
                    print("\033[93m已點擊【列印】按鈕\033[0m", flush=True)
                except Exception as e:
                    print(f"點擊列印按鈕失敗: {e}", flush=True)
                    return False

        # 🔥 增加等待時間並輪詢檢查
        print("等待系統自動處理列印和保存流程...", flush=True)
        max_wait_time = 15  # 最長等待15秒
        check_interval = 1  # 每秒檢查一次
        waited = 0
        new_pdfs = []

        while waited < max_wait_time:
            time.sleep(check_interval)
            waited += check_interval

            files_after = set(os.listdir(save_dir))
            new_files = files_after - files_before

            # 🔥 支援 Chrome 預設檔名格式: "高雄市地政電傳全方位地籍資料查詢系統.pdf" 或 "(1).pdf", "(2).pdf"
            new_pdfs = []
            for f in new_files:
                if f.endswith('.pdf'):
                    # 排除 .crdownload 臨時檔案
                    full_path = os.path.join(save_dir, f)
                    if os.path.exists(full_path) and not f.endswith('.crdownload'):
                        new_pdfs.append(f)

            if new_pdfs:
                print(f"✓ 在第 {waited} 秒檢測到 {len(new_pdfs)} 個新的PDF!", flush=True)
                break
            else:
                if waited % 3 == 0:  # 每3秒輸出一次
                    print(f"⏳ 已等待 {waited} 秒，尚未檢測到PDF...", flush=True)
        
        if new_pdfs:
            print(f"\n檢測到 {len(new_pdfs)} 個新的PDF文件:", flush=True)
            for pdf in new_pdfs:
                file_path = os.path.join(save_dir, pdf)
                file_size = os.path.getsize(file_path)
                print(f"- {pdf} (大小: {file_size} 字節)", flush=True)
            
            # 如果有多個新PDF，根據您的描述，第一個是正確的
            if len(new_pdfs) > 1:
                # 對新PDF文件按創建時間排序
                pdf_info = []
                for pdf in new_pdfs:
                    file_path = os.path.join(save_dir, pdf)
                    pdf_info.append({
                        'filename': pdf,
                        'path': file_path,
                        'created_time': os.path.getctime(file_path),
                        'size': os.path.getsize(file_path)
                    })
                
                pdf_info.sort(key=lambda x: x['created_time'])
                
                # 保留第一個PDF，刪除其他的
                correct_pdf = pdf_info[0]
                print(f"\n\033[92m識別到正確的PDF: {correct_pdf['filename']} (創建時間最早)\033[0m", flush=True)
                
                # 重命名正確的PDF
                if correct_pdf['filename'] != file_name + ".pdf":
                    try:
                        new_path = os.path.join(save_dir, file_name + ".pdf")
                        if os.path.exists(new_path):
                            timestamp = time.strftime("%Y%m%d%H%M%S")
                            file_name = f"{file_name}_{timestamp}"
                            new_path = os.path.join(save_dir, file_name + ".pdf")
                        
                        os.rename(correct_pdf['path'], new_path)
                        print(f"\033[92m已将正確的PDF重命名為: {file_name}.pdf\033[0m", flush=True)
                    except Exception as rename_e:
                        print(f"重命名文件時出錯: {rename_e}", flush=True)
                
                # 刪除其他PDF
                for pdf in pdf_info[1:]:
                    try:
                        os.remove(pdf['path'])
                        print(f"\033[93m已刪除多餘的PDF: {pdf['filename']}\033[0m", flush=True)
                    except Exception as remove_e:
                        print(f"刪除文件時出錯: {remove_e}", flush=True)
            else:
                # 只有一個新PDF，直接重命名
                pdf = new_pdfs[0]
                file_path = os.path.join(save_dir, pdf)
                
                if pdf != file_name + ".pdf":
                    try:
                        new_path = os.path.join(save_dir, file_name + ".pdf")
                        if os.path.exists(new_path):
                            timestamp = time.strftime("%Y%m%d%H%M%S")
                            file_name = f"{file_name}_{timestamp}"
                            new_path = os.path.join(save_dir, file_name + ".pdf")
                        
                        os.rename(file_path, new_path)
                        print(f"\033[92m已将PDF重命名為: {file_name}.pdf\033[0m", flush=True)
                    except Exception as rename_e:
                        print(f"重命名文件時出錯（嘗試複製再刪除）: {rename_e}", flush=True)
                        try:
                            shutil.copy2(file_path, new_path)
                            print(f"已成功複製為: {file_name}.pdf", flush=True)
                            try:
                                os.remove(file_path)
                                print("已刪除原始被佔用的PDF檔案", flush=True)
                            except Exception as delete_e:
                                print(f"原始PDF無法刪除（略過）: {delete_e}", flush=True)
                        except Exception as copy_e:
                            print(f"複製並重新命名仍失敗，略過該檔案: {copy_e}", flush=True)
        else:
            print("\033[91m未檢測到新的PDF文件\033[0m", flush=True)
        
        return True
    
    except Exception as e:
        print(f"\033[91m整個列印過程出錯: {e}\033[0m", flush=True)
        return False

# 輔助函數，獲取項目名稱
def get_project_name(type_code, value_code):
    land_projects = {
        '1': '土地標示部',
        '03': '土地所有權部',
        '05': '土地他項權利部',
        'D': '地籍圖',
        'R1': '土地參考資訊'
    }
    
    build_projects = {
        '7': '建物標示部',
        '09': '建物所有權部',
        '0B': '建物他項權利部',
        'C': '建物測量成果圖',
        'R7': '建物參考資訊'
    }
    
    if type_code == 'L':
        return land_projects.get(value_code, '未知項目')
    else:
        return build_projects.get(value_code, '未知項目')
    
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
                driver.quit()
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
                    driver.quit()
                    sys.exit(0)
                print("輸入的不是一個有效的數字，請重新輸入。", flush=True)

    except Exception as e:
        print(f"處理日期選擇時出錯: {e}", flush=True)
        print("您可以輸入 'Q' 退出程式: ", flush=True)
        user_input = input().strip()
        if user_input.lower() == 'q':
            print("程式即將退出...", flush=True)
            driver.quit()
            sys.exit(0)

def process_single_land_number(land_number, section, driver, gland_number, ID_number, register_options, current_scenario, gland_number_list=None):
    """
    處理單筆地號的函數
    現在函數接受所有需要的參數，包括新增的 gland_number_list 參數
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
    
    # 建號轉清單（新增）
    if gland_number and isinstance(gland_number, str):
        gland_number_list = [num.strip() for num in gland_number.split(',') if num.strip()]
    else:
        gland_number_list = []

    # 使用參數 gland_number 或 gland_number_list（原本就有的）
    if gland_number_list and len(gland_number_list) > 0:
        first_gland_number = gland_number_list[0].strip()
        gland_number_input.send_keys(first_gland_number)
        print(f"成功填入【建號】：{first_gland_number}", flush=True)
        if len(gland_number_list) > 1:
            print(f"注意：還有 {len(gland_number_list)-1} 個建號將在後續步驟中處理", flush=True)
    elif gland_number and gland_number.strip():
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
    print("\033[93m若正確則輸入【Y】，送出申請調閱\033[0m", flush=True)

    if scenario == 2:
        print("\033[93m若先前已申請調閱謄本，可輸入【N】進入【領件區】\033[0m", flush=True)
    else:
        print("\033[93m輸入【Q】可退出程式\033[0m", flush=True)

    # 等待使用者輸入
    import sys

    while True:
        user_input = input("請輸入[Y]來確認或[N]來取消(或輸入 [A] 代表之後都自動同意, [Q] 退出程式): ").strip().lower()

        if user_input == 'y':  # 使用者輸入 y 確認
            return True
        elif user_input == 'n':  # 使用者輸入 n 取消
            print("操作已取消。", flush=True)
            return False
        elif user_input == 'a':  # 使用者輸入 a 代表自動同意
            auto_confirm_all = True
            print("已設定為自動同意，後續將不再詢問。", flush=True)
            return True
        elif user_input == 'q':  # 使用者輸入 q 退出
            print("程式即將退出...", flush=True)
            driver.quit()
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

        # print("正在跳轉到【領件區】", flush=True)
        # driver.get("https://ep.land.nat.gov.tw/EpaperDoc/DocQuery")

        # WebDriverWait(driver, 10).until(EC.url_contains("/EpaperDoc/DocQuery"))
        # print("點擊成功，已跳轉到【領件區】", flush=True)

        # # 選擇日期
        # print("請選取日期...", flush=True)
        # select_date(driver)

        # # 選擇起始頁數
        # start_page = 1
        # print("請輸入起始頁數（預設為1）：", flush=True)
        # start_page_input = input().strip()
        # if start_page_input.isdigit():
        #     start_page = int(start_page_input)
        # if start_page > 1:
        #     try:
        #         page_link = driver.find_element(
        #             By.XPATH, f"//li[@class='page-number' and text()='{start_page}']"
        #         )
        #         page_link.click()
        #         print(f"跳轉至第 {start_page} 頁", flush=True)
        #     except NoSuchElementException:
        #         print(f"無法跳轉至第 {start_page} 頁，可能不存在", flush=True)

        # # 處理每頁的「預覽列印」連結
        # print("是否要自動下載所有檔案？(Y/N，預設：Y)：", flush=True)
        # auto_download_input = input().strip().lower()
        # if not auto_download_input:
        #     auto_download_input = 'y'
        # auto_download = auto_download_input == 'y'

        # auto_confirm = False
        # if auto_download:
        #     print("是否只詢問一次存儲目錄？(Y/N，預設：Y)：", flush=True)
        #     auto_confirm_input = input().strip().lower()
        #     if not auto_confirm_input:
        #         auto_confirm_input = 'y'
        #     auto_confirm = auto_confirm_input == 'y'

        # 初始化選定的目錄變數
        # selected_output_dir = {'dir': None}

        # handle_preview_print(
        #     driver,
        #     auto_download=auto_download,
        #     auto_confirm=auto_confirm,
        #     custom_dir_options=custom_dir_options,
        #     selected_output_dir=selected_output_dir
        # )

    except Exception as e:
        print(f"在下載檔案過程中發生錯誤：{e}", flush=True)

# def download_pdf_with_cookies(
#     driver,
#     auto_confirm,
#     custom_dir_options=None,
#     selected_output_dir=None
# ):
#     try:
#         # 獲取嵌入的 PDF 連結
#         embed_elements = driver.find_elements(By.TAG_NAME, "embed") + \
#                          driver.find_elements(By.TAG_NAME, "iframe") + \
#                          driver.find_elements(By.TAG_NAME, "object")

#         if not embed_elements:
#             print("未找到任何嵌入的 PDF 連結。", flush=True)
#             return False  # 返回 False

#         for embed in embed_elements:
#             pdf_url = embed.get_attribute("src")
#             if pdf_url:
#                 print(f"找到 PDF 連結: {pdf_url}", flush=True)

#                 # 建立下載請求
#                 session = requests.Session()
#                 headers = {'User-Agent': 'Mozilla/5.0'}
#                 for cookie in driver.get_cookies():
#                     session.cookies.set(cookie['name'], cookie['value'])

#                 response = session.get(pdf_url, headers=headers, timeout=30)
#                 response.raise_for_status()  # 檢查HTTP請求是否成功

#                 file_name = f"downloaded_{int(time.time())}.pdf"
#                 with open(file_name, 'wb') as f:
#                     f.write(response.content)
#                     print(f"PDF下載並保存為: {file_name}", flush=True)
                
#                  # 提取檔案名稱資訊並重新命名
#                 print("提取檔案名稱資訊並重新命名...", flush=True)
#                 類別, 段名, 地號 = '未提取', '未提取', '未提取' #預設值先給定
#                 try:
#                     text = extract_text_from_first_page(file_name)
#                     cleaned_text = clean_text(text)
#                     類別, 段名, 地號 = extract_info_enhanced(cleaned_text)
#                     # print(f"提取到的資訊 - 類別: {類別}, 段名: {段名}, 地號: {地號}", flush=True) # 顯示提取到的資訊
#                 except Exception as e:
#                     print(f"提取檔案資訊時發生錯誤: {e}", flush=True)
                    

#                 output_dir = rename_pdf(file_name, output_dir=None, auto_confirm=auto_confirm, custom_dir_options=custom_dir_options, selected_output_dir=selected_output_dir)
#                 if not output_dir:
#                     print(f"\033[91m重新命名第 1 個檔案失敗。\033[0m", flush=True)
#                     return False

#                 return True # 返回 True
#             else:
#                 print("嵌入元素沒有 'src' 屬性。", flush=True)
#         return False # 迴圈結束沒有找到PDF也返回 False
#     except Exception as e:
#         print(f"下載 PDF 檔案時出錯: {e}", flush=True)
#         return False  # 確保函數在發生異常時正確結束
    

# def handle_preview_print(
#     driver,
#     auto_download=False,
#     auto_confirm=False,
#     retry_count=0,
#     custom_dir_options=None,
#     selected_output_dir=None
# ):
#     max_retries = 3  # 設定最大重試次數

#     try:
#         # 主頁面下載連結處理
#         print("檢查是否有『預覽列印』連結...", flush=True)
#         try:
#             print("嘗試尋找『預覽列印』連結元素...", flush=True)
#             preview_print_links = WebDriverWait(driver, 10).until(
#                 EC.presence_of_all_elements_located((By.XPATH, "//a[contains(text(), '預覽列印')]"))
#             )

#             print(f"找到 {len(preview_print_links)} 個『預覽列印』連結", flush=True)  # 新增日誌
#             for i, link in enumerate(preview_print_links):
#                 print(f"第 {i+1} 個連結文字: {link.text}", flush=True)
#                 print(f"第 {i+1} 個連結 href: {link.get_attribute('href')}", flush=True)


#         except TimeoutException:
#             print("找不到『預覽列印』連結，稍後重試...", flush=True)
#             if retry_count < max_retries:
#                 print("重新整理頁面以確保最新內容...", flush=True)
#                 driver.refresh()  # 重新載入頁面
#                 time.sleep(5)  # 等待頁面刷新
#                 return handle_preview_print(
#                     driver,
#                     auto_download=auto_download,
#                     auto_confirm=auto_confirm,
#                     retry_count=retry_count + 1,
#                     custom_dir_options=custom_dir_options,
#                     selected_output_dir=selected_output_dir
#                 )
#             else:
#                 print("已達到最大重試次數，無法找到『預覽列印』連結。", flush=True)
#                 return False

#         # 詢問目錄
#         if auto_confirm and selected_output_dir and selected_output_dir['dir']:
#             # 已經選擇過目錄，直接使用
#             output_dir = selected_output_dir['dir']
#         else:
#             # 第一次詢問或未自動確認
#             if custom_dir_options:
#                 print("請選擇保存檔案的目錄：", flush=True)
#                 print("1. 預設目錄（下載的謄本）", flush=True)
#                 print(f"2. 自訂目錄（{custom_dir_options}）", flush=True)
#                 print("（直接按下 Enter 選擇預設選項）", flush=True)
#                 while True:
#                     choice = input("請輸入選項編號（1 或 2，預設：1）：").strip()
#                     if not choice:
#                         choice = '1'
#                     if choice == '1':
#                         # 使用預設目錄
#                         output_dir = get_work_folder("下載的謄本")
#                         break
#                     elif choice == '2':
#                         # 使用自訂目錄
#                         output_dir = custom_dir_options
#                         break
#                     else:
#                         print("輸入無效，請重新輸入。", flush=True)
#             else:
#                 # 如果沒有提供自訂目錄，使用預設目錄
#                 output_dir = get_work_folder("下載的謄本")

#             if auto_confirm and selected_output_dir is not None:
#                 # 記錄選擇的目錄，供後續使用
#                 selected_output_dir['dir'] = output_dir

#         #將此段程式碼移動到 download_pdf_with_cookies
#         for index in range(len(preview_print_links)):
#             try:
#                 print(f"開始處理第 {index + 1} 個『預覽列印』連結...", flush=True)  # 新增日誌
#                 print("重新尋找『預覽列印』連結元素...", flush=True)
#                 preview_print_links = WebDriverWait(driver, 10).until(
#                     EC.presence_of_all_elements_located((By.XPATH, "//a[contains(text(), '預覽列印')]"))
#                 )
#                 link = preview_print_links[index]

#                 print(f"嘗試滾動到第 {index + 1} 個連結...", flush=True)
#                 driver.execute_script("arguments[0].scrollIntoView();", link)
#                 print(f"滾動到第 {index + 1} 個連結完成", flush=True)

#                 print(f"等待 uiLockId 元素消失...", flush=True)
#                 WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.ID, "uiLockId")))
#                 print(f"uiLockId 元素已消失", flush=True)

#                 print(f"準備點擊第 {index + 1} 個連結...", flush=True)
#                 time.sleep(1)
#                 link.click()
#                 print(f"第 {index + 1} 個連結點擊完成", flush=True)

#                 main_window_handle = driver.current_window_handle
#                 print(f"主視窗控制代碼: {main_window_handle}", flush=True)  # 新增日誌
#                 windows = driver.window_handles
#                 print(f"所有視窗控制代碼: {windows}", flush=True)  # 新增日誌

#                 if len(windows) > 1:
#                     new_window_handle = [handle for handle in windows if handle != main_window_handle][0]
#                     print(f"新視窗控制代碼: {new_window_handle}", flush=True)  # 新增日誌
#                     print("嘗試切換到新視窗...", flush=True)
#                     driver.switch_to.window(new_window_handle)
#                     print("已切換到新視窗", flush=True)  # 新增日誌

#                     try:
#                         print("嘗試尋找收費資訊元素...", flush=True)
#                         fee_element = WebDriverWait(driver, 5).until(
#                             EC.presence_of_element_located((By.XPATH, "//strong[contains(text(), '本筆案件將收費')]"))
#                         )
#                         fee_text = fee_element.text
#                         fee_amount = fee_text.split(" ")[1]
#                         print(f"收費資訊: \033[93m{fee_amount}元\033[0m", flush=True)

#                         if auto_download:
#                             # 如果選擇自動下載，直接下載
#                             print("選擇自動下載，嘗試尋找【yes】按鈕...", flush=True)
#                             pay_button = WebDriverWait(driver, 10).until(
#                                 EC.presence_of_element_located((By.NAME, "yes"))
#                             )
#                             print("【yes】按鈕找到，嘗試點擊...", flush=True)
#                             pay_button.click()
#                             print("開始下載...", flush=True)
                            
#                             print("等待新視窗開啟1...", flush=True)
#                             WebDriverWait(driver, 10).until(EC.new_window_is_opened(driver.window_handles))
#                             print("新視窗已開啟", flush=True)
#                             new_windows = [handle for handle in driver.window_handles if handle != main_window_handle and handle != new_window_handle]
#                             if new_windows:
#                                 download_window = new_windows[0]
#                                 print("嘗試切換到下載視窗...", flush=True)
#                                 driver.switch_to.window(download_window)
#                                 print(f"已切換到下載視窗: {download_window}", flush=True)  # 新增日誌

#                                 # 調用下載函數進行下載
#                                 print("調用 download_pdf_with_cookies 函數...", flush=True)
#                                 download_pdf_with_cookies(driver, auto_confirm=auto_confirm, custom_dir_options=custom_dir_options, selected_output_dir=selected_output_dir)
#                             else:
#                                 print("\033[91m未找到下載視窗\033[0m", flush=True)
#                                 driver.close()
#                                 driver.switch_to.window(main_window_handle)
#                                 continue

#                         else:
#                             # 如果未選擇自動下載，詢問使用者是否下載
#                             print("要下載此檔案請輸入『Y』，要取消請直接按【Enter】：", flush=True)
#                             user_choice = input().strip().upper()
#                             # user_choice = input("要下載此檔案請輸入『Y』，要取消請直接按【Enter】：").strip().upper()
#                             if not user_choice:
#                                 user_choice = 'N'
#                             if user_choice == 'Y':
#                                 print("使用者選擇下載，嘗試尋找【yes】按鈕...", flush=True)
#                                 pay_button = WebDriverWait(driver, 10).until(
#                                     EC.presence_of_element_located((By.NAME, "yes"))
#                                 )
#                                 print("【yes】按鈕找到，嘗試點擊...", flush=True)
#                                 pay_button.click()
#                                 print("開始下載...", flush=True)

#                                 print("等待新視窗開啟2...", flush=True)
#                                 WebDriverWait(driver, 10).until(EC.new_window_is_opened(driver.window_handles))
#                                 print("新視窗已開啟", flush=True)
#                                 new_windows = [handle for handle in driver.window_handles if handle != main_window_handle and handle != new_window_handle]
#                                 if new_windows:
#                                     download_window = new_windows[0]
#                                     print("嘗試切換到下載視窗...", flush=True)
#                                     driver.switch_to.window(download_window)
#                                     print(f"已切換到下載視窗: {download_window}", flush=True)  # 新增日誌

#                                     # 調用下載函數進行下載
#                                     print("調用 download_pdf_with_cookies 函數...", flush=True)
#                                     download_pdf_with_cookies(driver, auto_confirm=auto_confirm, custom_dir_options=custom_dir_options, selected_output_dir=selected_output_dir)
#                                 else:
#                                     print("\033[91m未找到下載視窗\033[0m", flush=True)
#                                     driver.close()
#                                     driver.switch_to.window(main_window_handle)
#                                     continue

#                             else:
#                                 print("您選擇不下載此檔案。", flush=True)

#                     except TimeoutException:
#                         # 如果沒有找到收費資訊，可能是已付費的檔案，直接下載
#                         print("沒有找到收費資訊，可能是已付費的檔案，直接下載...", flush=True)
#                         print("調用 download_pdf_with_cookies 函數...", flush=True)
#                         download_pdf_with_cookies(driver, auto_confirm=auto_confirm, custom_dir_options=custom_dir_options, selected_output_dir=selected_output_dir)

#                     finally:
#                         print("準備關閉新視窗並切換回主視窗...", flush=True)  # 新增日誌
#                         try:
#                             print("嘗試使用 JavaScript 關閉新視窗...", flush=True)
#                             driver.execute_script("window.close();")  # 使用 JavaScript 關閉視窗
#                             print("使用 JavaScript 關閉新視窗成功", flush=True)  # 新增日誌
#                         except Exception as e:
#                             print(f"使用 JavaScript 關閉視窗失敗: {e}", flush=True)  # 新增日誌
#                             try:
#                                 print("嘗試使用原始關閉視窗方法...", flush=True)
#                                 driver.close() # 原始的關閉視窗方法
#                                 print("原始關閉視窗方法", flush=True)
#                             except Exception as e:
#                                 print(f"原始關閉視窗方法失敗：{e}", flush=True)
#                         print("已關閉新視窗", flush=True)  # 新增日誌
#                         print("嘗試切換回主視窗...", flush=True)
#                         driver.switch_to.window(main_window_handle)
#                         print("已切換回主視窗", flush=True)  # 新增日誌
#                 else:
#                     print("沒有找到新視窗", flush=True)

#             except Exception as e:
#                 print(f"點擊連結時出錯: {e}", flush=True)

#         # 處理下一頁
#         try:
#             print("嘗試尋找下一頁按鈕...", flush=True)
#             next_page_button = WebDriverWait(driver, 3).until(
#                 EC.element_to_be_clickable((By.XPATH, "//li[@class='pgNext']/a"))
#             )
#             if next_page_button.is_displayed() and next_page_button.is_enabled():
#                 print("找到下一頁按鈕，嘗試點擊...", flush=True)
#                 next_page_button.click()
#                 time.sleep(2)
#                 handle_preview_print(
#                     driver,
#                     auto_download=auto_download,
#                     auto_confirm=auto_confirm,
#                     retry_count=retry_count,
#                     custom_dir_options=custom_dir_options,
#                     selected_output_dir=selected_output_dir
#                 )  # 遞迴呼叫
#             else:
#                 print("下一頁按鈕不可見或不可用", flush=True)

#         except (NoSuchElementException, TimeoutException):
#             print("已到最後一頁，無法找到下一頁按鈕", flush=True)

#     except Exception as e:
#         print(f"處理『預覽列印』連結時發生錯誤: {e}", flush=True)
#         if retry_count < max_retries:
#             print("重新載入選擇起始日", flush=True)
#             driver.get("https://ep.land.nat.gov.tw/EpaperDoc/DocQuery")
#             select_date(driver)
#             handle_preview_print(
#                 driver,
#                 auto_download=auto_download,
#                 auto_confirm=auto_confirm,
#                 retry_count=retry_count + 1,
#                 custom_dir_options=custom_dir_options,
#                 selected_output_dir=selected_output_dir
#             )
#         else:
#             print("處理『預覽列印』連結失敗，重試次數已達上限。", flush=True)
#     return True


def add_print_keyboard_shortcut(driver):
    """
    顯示增强版列印功能的說明
    """
    print("\n\033[93m=== 列印功能說明 ===\033[0m", flush=True)
    print("查詢登記謄本後，系統會自動嘗試點擊【列印】按鈕。", flush=True)
    print("系統會建議的PDF文件命名格式為:", flush=True)
    print("  [縣市][地區][地段]-[地號/建號]-[部分類型]", flush=True)
    print("  例如: 高雄市大樹區曹公段-64-28地號-標示部", flush=True)
    print("        高雄市大樹區曹公段-355建號-建物測量成果圖", flush=True)
    print("\n保存PDF操作步驟:", flush=True)
    print("1. 當列印對話框出現時，選擇「另存為PDF」選項", flush=True)
    print("2. 使用系統建議的文件名保存", flush=True)
    print("3. 選擇適當的保存位置", flush=True)
    print("4. 完成後返回程序繼續操作", flush=True)

    print("\n如需手動操作，可在網頁中點擊頂部菜單的【列印】按鈕。", flush=True)
    print("\033[93m==================\033[0m\n", flush=True)
    
    return True


def extract_and_query_land_registration(driver, county, district, section, custom_save_dir=None):
    """
    從建物所有權部抓取土地登記次序資訊，並詢問使用者是否要調閱

    參數:
        custom_save_dir: 自訂儲存目錄，如果為 None 則使用預設目錄
    """
    try:
        print("\n\033[93m檢查是否有建物坐落土地所有權登記次序...\033[0m", flush=True)

        # 尋找土地登記次序的 accordion 區塊
        try:
            accordion_header = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//h4[contains(text(), '建物坐落土地所有權登記次序')]"))
            )
            print("✓ 找到建物坐落土地所有權登記次序區塊", flush=True)
        except:
            print("未找到建物坐落土地所有權登記次序資訊", flush=True)
            return

        # 展開 accordion (如果還沒展開)
        try:
            if "ui-accordion-header-collapsed" in accordion_header.get_attribute("class"):
                accordion_header.click()
                time.sleep(1)
                print("已展開土地登記次序區塊", flush=True)
        except:
            pass

        # 抓取所有土地登記次序的連結
        land_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'linkLandQry')]")

        if not land_links:
            print("未找到任何土地登記次序連結", flush=True)
            return

        # 解析連結資訊
        land_data = []
        for link in land_links:
            onclick = link.get_attribute("onclick")
            text = link.text.strip()

            # 從 onclick 中解析參數: linkLandQry('3#EC;1259;1023;0000;0110')
            if onclick and "linkLandQry" in onclick:
                try:
                    params_start = onclick.find("'") + 1
                    params_end = onclick.rfind("'")
                    params = onclick[params_start:params_end]

                    # 參數格式: 3#EC;1259;1023;0000;0110
                    # 分解: [type];[district_code];[land_number];[sub_number];[registration_order]
                    parts = params.split(';')
                    if len(parts) >= 5:
                        registration_order = parts[4]  # 登記次序
                        land_number = parts[2]  # 地號
                        sub_number = parts[3]  # 小地號

                        # 組合完整地號
                        if sub_number and sub_number != '0000':
                            full_land_number = f"{land_number}-{sub_number}"
                        else:
                            full_land_number = land_number

                        land_data.append({
                            'land_number': full_land_number,
                            'registration_order': registration_order,
                            'onclick': onclick,
                            'params': params
                        })

                        print(f"  找到土地: {full_land_number}, 登記次序: {registration_order}", flush=True)
                except Exception as parse_e:
                    print(f"解析連結時出錯: {parse_e}", flush=True)

        if not land_data:
            print("未找到有效的土地登記次序", flush=True)
            return

        # 詢問使用者是否要調閱這些土地
        print(f"\n\033[93m========================================\033[0m", flush=True)
        print(f"\033[93m發現 {len(land_data)} 筆建物坐落土地\033[0m", flush=True)
        print(f"\033[93m========================================\033[0m", flush=True)

        for i, land in enumerate(land_data, 1):
            print(f"  {i}. 地號: {land['land_number']}, 登記次序: {land['registration_order']}", flush=True)

        print(f"\n\033[93m是否要調閱這些土地的所有權部？\033[0m", flush=True)
        print(f"\033[093m  [Y] = 是，調閱這些土地（預設）\033[0m", flush=True)
        print(f"\033[093m  [N] = 否，略過\033[0m", flush=True)
        print(f"\033[93m========================================\033[0m", flush=True)

        user_input = input("\033[93m請輸入選擇 [Y/N]: \033[0m").strip().upper()

        if user_input == 'N':
            print("\033[92m✓ 已略過土地調閱\033[0m", flush=True)
            return

        # 調閱每筆土地
        print("\033[92m✓ 開始調閱土地所有權部\033[0m", flush=True)

        for i, land in enumerate(land_data, 1):
            print(f"\n調閱第 {i}/{len(land_data)} 筆土地: {land['land_number']} (登記次序: {land['registration_order']})", flush=True)

            try:
                # 執行 onclick
                driver.execute_script(land['onclick'])

                # 🔥 等待頁面完全加載,確認不是列表頁面
                print("等待土地所有權部頁面加載...", flush=True)
                time.sleep(3)

                # 確認已經離開列表頁面
                try:
                    WebDriverWait(driver, 5).until_not(
                        EC.presence_of_element_located((By.XPATH, "//h4[contains(text(), '建物坐落土地所有權登記次序')]"))
                    )
                    print("✓ 已離開列表頁面,進入土地所有權部詳細頁面", flush=True)
                except:
                    print("確認頁面轉換...", flush=True)

                time.sleep(2)  # 額外等待確保內容完全載入

                # 列印土地所有權部
                click_print_and_save_pdf(driver, 'L', '03', land['land_number'], county, district, section, suffix=f"_登記次序{land['registration_order']}", custom_save_dir=custom_save_dir)

                # 返回建物頁面
                driver.back()
                time.sleep(2)

            except Exception as e:
                print(f"調閱土地 {land['land_number']} 時發生錯誤: {e}", flush=True)

        print("\033[92m✓ 完成土地調閱\033[0m", flush=True)

    except Exception as e:
        print(f"抓取土地登記次序時發生錯誤: {e}", flush=True)


def handle_multiple_rights(driver, item_type, item_value, number, county, district, section):
    """
    處理他項權利部有多個權利人的情況，逐個點擊並列印詳細信息
    """
    try:
        # 檢查是否顯示他項權利列表（更靈活的檢測方式）
        rights_count_element = None
        rights_count_text = ""

        try:
            # 方法1：嘗試查找"他項權利共X人"的文本
            rights_count_element = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//td[contains(text(), '他項權利共')]"))
            )
            rights_count_text = rights_count_element.text
            print(f"檢測到他項權利列表: {rights_count_text}", flush=True)
        except:
            # 方法2：嘗試通過檢查 div#other_owner_data 和連結來確認
            try:
                other_owner_data = driver.find_element(By.ID, "other_owner_data")
                test_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'treeQuery') and contains(@class, 'class_box')]")
                if test_links and len(test_links) > 0:
                    print(f"檢測到他項權利列表（通過連結數量確認: {len(test_links)} 個）", flush=True)
                else:
                    raise Exception("未找到權利連結")
            except:
                # 如果兩種方法都失敗，則認為不是列表頁面
                print("未檢測到他項權利列表，可能是單個權利或其他頁面", flush=True)
                return False
        
        # 查找所有的權利連結
        rights_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'treeQuery') and contains(@class, 'class_box')]")
        
        if not rights_links:
            print("未找到任何他項權利連結", flush=True)
            return False
        
        print(f"找到 {len(rights_links)} 個他項權利連結", flush=True)

        # 先列印他項權利列表頁面（總覽）
        print("列印他項權利列表總覽頁面...", flush=True)
        click_print_and_save_pdf(driver, item_type, item_value, number, county, district, section, suffix="_他項權利列表")
        time.sleep(3)  # 等待列印完成

        # 🚀 優化版：快速收集所有權利資料，減少不必要的輸出
        rights_data = []
        for link in rights_links:
            text = link.text.strip()
            onclick = link.get_attribute("onclick")
            # 從onclick中提取參數 (treeQuery函數的參數)
            params = None
            if onclick and "treeQuery" in onclick:
                try:
                    # 假設格式為 treeQuery('param1', 'param2', ...)
                    onclick = onclick.strip()
                    params_start = onclick.find("(") + 1
                    params_end = onclick.rfind(")")
                    if params_start > 0 and params_end > params_start:
                        params = onclick[params_start:params_end]
                except:
                    pass  # 靜默處理錯誤，加速執行

            rights_data.append({"text": text, "onclick": onclick, "params": params})

        print(f"✓ 快速收集完成，共 {len(rights_data)} 個他項權利", flush=True)

        # 若數量超過3個，詢問使用者是否全部調閱
        selected_indices = None
        skip_all_rights = False  # 🔥 新增：標記是否略過全部

        if len(rights_data) > 3:
            print(f"\n\033[93m========================================\033[0m", flush=True)
            print(f"\033[93m檢測到 {len(rights_data)} 個他項權利\033[0m", flush=True)
            print(f"\033[93m========================================\033[0m", flush=True)

            # 顯示他項權利清單供使用者參考
            for i, right in enumerate(rights_data, 1):
                print(f"  {i}. {right['text']}", flush=True)

            print(f"\n\033[93m請選擇操作：\033[0m", flush=True)
            print(f"\033[093m  [Enter] = 全部調閱（預設）\033[0m", flush=True)
            print(f"\033[093m  [N] = 略過全部（不調閱任何權利）\033[0m", flush=True)
            print(f"\033[093m  [序號] = 只調閱指定序號（例如：1,3,5 或 1-3,5 或 2-4）\033[0m", flush=True)
            print(f"\033[93m========================================\033[0m", flush=True)

            user_input = input("\033[93m請輸入選擇: \033[0m").strip()

            # 🔥 處理輸入
            if not user_input:
                # 直接按 Enter，調閱全部
                print("\033[92m✓ 將調閱全部他項權利\033[0m", flush=True)
            elif user_input.upper() == 'N':
                # 輸入 N，略過全部
                skip_all_rights = True
                print("\033[92m✓ 已選擇略過全部他項權利\033[0m", flush=True)
            elif user_input:  # 如果有輸入序號，解析序號
                selected_indices = set()
                parts = user_input.split(',')

                for part in parts:
                    part = part.strip()
                    if '-' in part:  # 範圍選擇，例如 1-3
                        try:
                            start, end = part.split('-')
                            start = int(start.strip())
                            end = int(end.strip())
                            for i in range(start, end + 1):
                                if 1 <= i <= len(rights_data):
                                    selected_indices.add(i - 1)  # 轉換為0-based索引
                        except ValueError:
                            print(f"\033[91m無效的範圍格式: {part}，將略過\033[0m", flush=True)
                    else:  # 單一序號
                        try:
                            idx = int(part)
                            if 1 <= idx <= len(rights_data):
                                selected_indices.add(idx - 1)  # 轉換為0-based索引
                            else:
                                print(f"\033[91m序號 {idx} 超出範圍，將略過\033[0m", flush=True)
                        except ValueError:
                            print(f"\033[91m無效的序號: {part}，將略過\033[0m", flush=True)

                selected_indices = sorted(list(selected_indices))
                print(f"\033[92m✓ 將調閱以下他項權利: {[i+1 for i in selected_indices]}\033[0m", flush=True)

        # 🔥 如果選擇略過全部，直接返回
        if skip_all_rights:
            print("\033[93m已略過所有他項權利的調閱\033[0m", flush=True)
            return True

        # 處理每個權利
        for idx, right_info in enumerate(rights_data):
            # 如果有選擇性調閱，檢查當前索引是否在選擇範圍內
            if selected_indices is not None and idx not in selected_indices:
                print(f"略過第 {idx+1} 個他項權利（未選擇）", flush=True)
                continue

            text = right_info["text"]
            print(f"\n處理第 {idx+1}/{len(rights_data)} 個他項權利: {text}", flush=True)
            
            # 先檢查我們是否仍在列表頁面（使用更靈活的檢測方式）
            list_detected = False
            try:
                # 方法1：尋找"他項權利共X人"文本
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, "//td[contains(text(), '他項權利共')]"))
                )
                list_detected = True
            except:
                # 方法2：檢查 div#other_owner_data 和連結
                try:
                    other_owner_data = driver.find_element(By.ID, "other_owner_data")
                    test_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'treeQuery') and contains(@class, 'class_box')]")
                    if test_links and len(test_links) >= len(rights_data):
                        list_detected = True
                except:
                    pass

            if list_detected:
                print("目前在他項權利列表頁面", flush=True)

                # 重新獲取所有連結元素
                current_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'treeQuery') and contains(@class, 'class_box')]")

                # 確保有足夠的連結
                if idx < len(current_links):
                    target_link = current_links[idx]
                    print(f"準備點擊連結: {target_link.text}", flush=True)

                    # 確保元素可見和可點擊
                    driver.execute_script("arguments[0].scrollIntoView(true);", target_link)
                    time.sleep(1)

                    # 點擊連結
                    target_link.click()
                    print(f"已點擊連結: {target_link.text}", flush=True)
                else:
                    print(f"連結索引 {idx} 超出範圍 (當前連結數: {len(current_links)})", flush=True)
                    continue
            else:
                print(f"未檢測到列表頁面，嘗試返回...", flush=True)
                # 嘗試使用瀏覽器的返回按鈕
                try:
                    driver.back()
                    time.sleep(3)  # 等待頁面加載
                    continue
                except Exception as back_e:
                    print(f"使用瀏覽器返回按鈕時出錯: {back_e}", flush=True)
                    print("嘗試使用瀏覽器返回按鈕回到列表頁面...", flush=True)

            try:
                pass  # 保持原有的 except 結構
            except Exception as e:
                print(f"在列表頁面操作時出錯: {e}", flush=True)
                print("嘗試使用瀏覽器返回按鈕回到列表頁面...", flush=True)
                
                # 嘗試使用瀏覽器的返回按鈕
                try:
                    driver.back()
                    time.sleep(3)  # 等待頁面加載
                    
                    # 確認是否返回到列表頁面
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//td[contains(text(), '他項權利共')]"))
                        )
                        print("已返回他項權利列表頁面", flush=True)
                        
                        # 繼續下一個循環，將在下一次迭代中嘗試點擊
                        continue
                    except:
                        print("返回後未能找到他項權利列表，跳過剩餘權利", flush=True)
                        return False
                except Exception as back_e:
                    print(f"使用瀏覽器返回按鈕時出錯: {back_e}", flush=True)
                    return False
            
            # 等待詳細頁面加載
            time.sleep(3)
            
            # 檢查是否成功加載詳細頁面
            try:
                # 檢查是否不再顯示列表頁面的元素
                WebDriverWait(driver, 3).until_not(
                    EC.presence_of_element_located((By.XPATH, "//td[contains(text(), '他項權利共')]"))
                )
                print("成功加載他項權利詳細頁面", flush=True)
            except:
                print("可能未能成功加載詳細頁面，但將繼續嘗試列印", flush=True)
            
            # 執行列印
            print(f"列印他項權利 {text} 的詳細信息", flush=True)
            # 格式化登記次序：0001000 -> 0001-000
            if len(text) == 7:
                formatted_text = f"{text[:4]}-{text[4:]}"
            else:
                formatted_text = text
            # 登記次序用括號包起來：地號-(登記次序)
            modified_number = f"{number}-({formatted_text})"
            click_print_and_save_pdf(driver, item_type, item_value, modified_number, county, district, section)
            
            # 等待列印完成
            time.sleep(3)

            # 不使用返回按鈕，而是點擊左側查詢結果樹的他項權利部節點來重新載入列表
            print("重新執行查詢以顯示他項權利列表...", flush=True)
            try:
                # 方法1：找到左側樹狀結構中的「他項權利部(X)」節點並點擊
                # 從 HTML 看，onclick 格式為: treeQuery('E','E14','1717','2337-0','05','','')
                rights_tree_node = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//span[contains(@class, 'other_folder') and contains(@onclick, \"'{item_value}'\") and contains(text(), '他項權利部')]"))
                )
                rights_tree_node.click()
                print("已點擊左側樹狀結構的「他項權利部」節點，重新載入列表", flush=True)
                time.sleep(3)  # 等待頁面加載
            except Exception as e:
                print(f"無法點擊左側樹狀結構的他項權利部節點: {e}", flush=True)
                print("嘗試使用其他方法返回列表...", flush=True)

                # 備選方案：嘗試使用瀏覽器返回
                try:
                    driver.back()
                    print("已使用瀏覽器返回功能", flush=True)
                    time.sleep(3)  # 等待返回操作完成
                except Exception as back_e:
                    print(f"使用瀏覽器返回時出錯: {back_e}", flush=True)
                    print("將跳過剩餘權利", flush=True)
                    return False

            # 確認是否已返回列表頁面
            try:
                # 嘗試找到列表頁面的特徵元素
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "other_owner_data"))
                )
                # 再確認有連結
                current_links = driver.find_elements(By.XPATH, "//a[contains(@onclick, 'treeQuery') and contains(@class, 'class_box')]")
                if len(current_links) >= len(rights_data):
                    print(f"已成功返回他項權利列表頁面，檢測到 {len(current_links)} 個連結", flush=True)
                    time.sleep(2)  # 確保頁面完全加載
                else:
                    print(f"返回列表頁面，但連結數量不符 (預期 {len(rights_data)}，實際 {len(current_links)})", flush=True)
                    print("將繼續嘗試處理下一個權利", flush=True)
            except Exception as e:
                print(f"未能確認返回列表頁面: {e}，將跳過剩餘權利", flush=True)
                return False
        
        return True
    
    except Exception as e:
        print(f"處理多個他項權利時出錯: {e}", flush=True)
        try:
            # 嘗試使用瀏覽器返回功能
            driver.back()
            print("已使用瀏覽器返回功能", flush=True)
            time.sleep(3)
        except:
            print("使用瀏覽器返回功能失敗", flush=True)
        return False




if __name__ == "__main__":
    main()
    print("子程式已完成執行", flush=True)  #  <--- 在這裡直接 print