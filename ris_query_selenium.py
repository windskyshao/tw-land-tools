"""
使用 Selenium 自動化查詢內政部戶政司村里街路門牌
支援使用者手動輸入 CAPTCHA 驗證碼
"""

import sys
import io
import os

# 🔥 在 PyInstaller 環境中，優先載入 python_embedded 的套件
# 這樣才能正確導入 selenium 和 rapidocr
try:
    from base_dir_helper import get_base_dir
    base_dir = get_base_dir()
    python_embedded_paths = [
        os.path.join(base_dir, 'python_embedded', 'Lib', 'site-packages'),
        os.path.join(base_dir, 'python_embedded', 'Lib'),
        os.path.join(base_dir, 'python_embedded')
    ]
    for path in python_embedded_paths:
        if os.path.exists(path) and path not in sys.path:
            sys.path.insert(0, path)
except:
    pass

# 設定 stdout 編碼為 UTF-8
# 🔥 在 PyInstaller 環境中，sys.stdout 可能是 None，需要檢查
if sys.stdout is not None and hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import re
from PIL import Image
from io import BytesIO
import warnings
import logging


def to_fullwidth_number(text):
    """將半形數字轉換為全形數字"""
    halfwidth = "0123456789"
    fullwidth = "０１２３４５６７８９"
    trans_table = str.maketrans(halfwidth, fullwidth)
    return text.translate(trans_table)


def chinese_number_to_arabic(chinese_num):
    """將中文數字轉換為阿拉伯數字（支援一到九十九）"""
    chinese_digits = {
        '零': 0, '一': 1, '二': 2, '三': 3, '四': 4,
        '五': 5, '六': 6, '七': 7, '八': 8, '九': 9,
        '十': 10
    }

    # 如果只有一個字
    if len(chinese_num) == 1:
        return chinese_digits.get(chinese_num, 0)

    # 處理十幾（如：十一、十二）
    if chinese_num.startswith('十'):
        if len(chinese_num) == 1:
            return 10
        else:
            return 10 + chinese_digits.get(chinese_num[1], 0)

    # 處理二十、三十等
    if '十' in chinese_num:
        parts = chinese_num.split('十')
        tens = chinese_digits.get(parts[0], 0) * 10
        ones = chinese_digits.get(parts[1], 0) if len(parts) > 1 and parts[1] else 0
        return tens + ones

    return 0


def recognize_captcha(image_element, debug=True, log_callback=None):
    """
    使用 RapidOCR 辨識驗證碼

    Args:
        image_element: Selenium WebElement of the captcha image
        debug: 是否顯示除錯訊息和儲存截圖
        log_callback: 訊息輸出回調函數

    Returns:
        str: 辨識出的驗證碼文字（只保留數字和字母）
    """
    log = log_callback or print

    try:
        # 抑制警告訊息
        import warnings
        import logging
        warnings.filterwarnings("ignore")
        logging.disable(logging.WARNING)

        # 🔥 初始化 RapidOCR
        from rapidocr_helper import create_rapidocr_reader, rapidocr_readtext
        import numpy as np
        import cv2

        if debug:
            log("  [除錯] 正在初始化 RapidOCR...")

        # 使用 RapidOCR（驗證碼優化參數，與 hinet.py 相同）
        ocr = create_rapidocr_reader(
            use_angle_cls=False,
            use_text_score=True,
            text_score=0.3,
            det_db_thresh=0.2,
            det_db_box_thresh=0.3,
            det_db_unclip_ratio=1.8,
            max_side_len=1280
        )

        if debug:
            log("  [除錯] RapidOCR 初始化完成")

        # 截圖驗證碼圖片
        image_screenshot = image_element.screenshot_as_png
        image = Image.open(BytesIO(image_screenshot))

        # 儲存原始截圖供除錯
        if debug:
            debug_path_original = "captcha_debug_original.png"
            image.save(debug_path_original)
            log(f"  [除錯] 原始驗證碼截圖已儲存：{debug_path_original}")
            log(f"  [除錯] 原始圖片大小：{image.size}")

        # 🔥 圖片預處理：僅放大，不做過度處理
        import numpy as np

        # 轉換為 numpy array
        img_array = np.array(image)

        # 🔥 放大圖片 2 倍（適度放大，避免過度處理）
        scale_factor = 2
        height, width = img_array.shape[:2]
        new_width = width * scale_factor
        new_height = height * scale_factor
        img_resized = cv2.resize(img_array, (new_width, new_height), interpolation=cv2.INTER_CUBIC)

        if debug:
            log(f"  [除錯] 圖片已放大：{width}x{height} → {new_width}x{new_height}")

        # 🔥 直接使用放大後的彩色圖片（RapidOCR 對彩色圖片效果可能更好）
        img_final = img_resized

        if debug:
            # 儲存處理後的圖片供除錯
            debug_path_processed = "captcha_debug_processed.png"
            cv2.imwrite(debug_path_processed, cv2.cvtColor(img_final, cv2.COLOR_RGB2BGR))
            log(f"  [除錯] 處理後圖片已儲存：{debug_path_processed}")

        # 🔥 使用 RapidOCR 辨識
        result = rapidocr_readtext(ocr, img_final, detail=1)

        if debug:
            log(f"  [除錯] OCR 原始結果：{result}")

        if result:
            # 🔥 根據 x 座標排序（確保從左到右的順序）
            sorted_results = sorted(result, key=lambda item: item[0][0][0])

            # 提取文字
            texts = []
            for detection in sorted_results:
                # RapidOCR 返回格式: (bbox, text, confidence)
                bbox, text, confidence = detection
                if debug:
                    log(f"  [除錯] 辨識到：'{text}' (信心度: {confidence:.2f})")
                # 🔥 降低信心度閾值，接受更多結果（驗證碼通常較難辨識）
                if confidence > 0.2:
                    texts.append(text)

            # 合併所有文字
            captcha_text = ''.join(texts)

            if debug:
                log(f"  [除錯] 合併文字：'{captcha_text}'")

            # 只保留數字和字母
            captcha_text = re.sub(r'[^a-zA-Z0-9]', '', captcha_text)

            if debug:
                log(f"  [除錯] 過濾後：'{captcha_text}'")

            return captcha_text

        return None

    except ImportError as e:
        log(f"⚠ 未安裝 RapidOCR：{e}")
        log("  請執行: pip install rapidocr-onnxruntime")
        return None
    except Exception as e:
        log(f"⚠ 驗證碼辨識失敗：{e}")
        if debug:
            import traceback
            error_detail = traceback.format_exc()
            log(f"  [除錯] 錯誤詳情：\n{error_detail}")
        return None


# 縣市名稱對照表（對應到 data-id）
CITY_ID_MAP = {
    '臺北市': '63000000',
    '台北市': '63000000',
    '新北市': '65000000',
    '桃園市': '68000000',
    '臺中市': '66000000',
    '台中市': '66000000',
    '臺南市': '67000000',
    '台南市': '67000000',
    '高雄市': '64000000',
    '基隆市': '10017000',
    '新竹市': '10018000',
    '嘉義市': '10020000',
    '新竹縣': '10004000',
    '苗栗縣': '10005000',
    '彰化縣': '10007000',
    '南投縣': '10008000',
    '雲林縣': '10009000',
    '嘉義縣': '10010000',
    '屏東縣': '10013000',
    '宜蘭縣': '10002000',
    '花蓮縣': '10015000',
    '臺東縣': '10014000',
    '台東縣': '10014000',
    '澎湖縣': '10016000',
    '金門縣': '09020000',
    '連江縣': '09007000'
}


class RISQueryBot:
    """內政部戶政司村里查詢機器人"""

    def __init__(self, headless=False, message_callback=None):
        """
        初始化 Selenium WebDriver

        Args:
            headless: 是否使用無頭模式（背景執行）
            message_callback: 訊息輸出回調函數
        """
        self.message_callback = message_callback or print

        self.log("=" * 60)
        self.log("內政部戶政司村里街路門牌查詢機器人")
        self.log("=" * 60)

        # 設定 Chrome 選項
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        # 🔥 強制使用淺色模式（避免暗黑模式影響截圖品質）
        options.add_argument('--disable-features=WebUIDarkMode,ForcedColors')
        options.add_argument('--force-color-profile=srgb')
        options.add_argument('--blink-settings=preferredColorScheme=1')  # 1=light

        # 設定 Chrome 偏好，強制淺色主題
        prefs = {
            'profile.default_content_setting_values.prefers_color_scheme': 1,  # 1=light, 2=dark
            'profile.managed_default_content_settings.images': 1,
            'browser.theme.color_scheme': 'light'
        }
        options.add_experimental_option('prefs', prefs)

        # 初始化 WebDriver
        try:
            self.log("\n[初始化] 啟動 Chrome 瀏覽器...")
            # 使用 webdriver-manager 自動管理 ChromeDriver
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 20)

            # 設定視窗位置和大小：左上角(0,0)，寬度1024，高度為螢幕高度扣掉工作列
            if not headless:
                import tkinter as tk
                root = tk.Tk()
                screen_height = root.winfo_screenheight()
                root.destroy()
                # 工作列通常約40-50像素，保守扣除60像素
                window_height = screen_height - 60
                self.driver.set_window_position(0, 0)
                self.driver.set_window_size(1024, window_height)

                # 🔥 將 Chrome 視窗置於最前方（解決被物件資料編輯器覆蓋的問題）
                try:
                    import ctypes
                    # 取得 Chrome 視窗的 HWND
                    hwnd = ctypes.windll.user32.GetForegroundWindow()
                    # 使用 driver 的 window_handles 確保取得正確視窗
                    # 首先讓 Chrome 視窗獲得焦點
                    self.driver.switch_to.window(self.driver.current_window_handle)
                    # 使用 JavaScript 讓視窗獲得焦點
                    self.driver.execute_script("window.focus();")
                    time.sleep(0.3)
                    # 再次取得前景視窗（應該是 Chrome）
                    hwnd = ctypes.windll.user32.GetForegroundWindow()
                    # 將視窗置於最前方
                    ctypes.windll.user32.SetForegroundWindow(hwnd)
                    self.log("✓ Chrome 視窗已置於最前方")
                except Exception as e:
                    self.log(f"⚠ 無法將 Chrome 視窗置前：{e}")

            self.log("✓ Chrome 瀏覽器已啟動")
        except Exception as e:
            self.log(f"✗ 啟動失敗：{e}")
            self.log("\n請確認：")
            self.log("1. 已安裝 Google Chrome 瀏覽器")
            self.log("2. 已安裝 Selenium 和 webdriver-manager")
            self.log("   執行：pip install selenium webdriver-manager")
            raise

    def log(self, message):
        """輸出訊息到 callback 或 print"""
        self.message_callback(message)

    def bring_to_front(self):
        """將 Chrome 視窗置於最前方（使用視窗標題查找）"""
        try:
            import ctypes
            from ctypes import wintypes

            # 定義 Windows API
            user32 = ctypes.windll.user32
            EnumWindows = user32.EnumWindows
            EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
            GetWindowText = user32.GetWindowTextW
            GetWindowTextLength = user32.GetWindowTextLengthW
            IsWindowVisible = user32.IsWindowVisible
            SetForegroundWindow = user32.SetForegroundWindow
            ShowWindow = user32.ShowWindow
            SW_RESTORE = 9

            chrome_hwnd = None

            def enum_callback(hwnd, lParam):
                nonlocal chrome_hwnd
                if IsWindowVisible(hwnd):
                    length = GetWindowTextLength(hwnd)
                    if length > 0:
                        buff = ctypes.create_unicode_buffer(length + 1)
                        GetWindowText(hwnd, buff, length + 1)
                        title = buff.value
                        # 查找包含 Chrome 或網頁標題的視窗
                        if 'Chrome' in title or '內政部' in title or 'ris.gov' in title:
                            chrome_hwnd = hwnd
                            return False  # 停止枚舉
                return True

            EnumWindows(EnumWindowsProc(enum_callback), 0)

            if chrome_hwnd:
                # 先恢復視窗（如果最小化）
                ShowWindow(chrome_hwnd, SW_RESTORE)
                time.sleep(0.1)
                # 將視窗置於最前方
                SetForegroundWindow(chrome_hwnd)
        except Exception:
            pass  # 靜默失敗，不影響主流程

    def query_by_address(self, city, district, road, number="", full_number="", save_pdf_path=None, pdf_keywords=None):
        """
        根據地址查詢村里鄰資訊

        Args:
            city: 縣市名稱（如：高雄市）
            district: 區鄉鎮名稱（如：左營區）
            road: 路街名稱（如：重惠街）
            number: 門牌號碼（可選，如：106號）- 不含樓層
            full_number: 完整門牌（可選，如：106號七樓）- 包含樓層
            save_pdf_path: PDF 儲存路徑（可選）
            pdf_keywords: 要在 PDF 上標記的關鍵字列表（可選）

        Returns:
            dict: 查詢結果 {"村里": "...", "鄰": "..."}
        """
        try:
            # 步驟 1: 訪問首頁
            self.log(f"\n[步驟 1] 訪問內政部戶政司查詢首頁...")
            self.driver.get("https://www.ris.gov.tw/app/portal/3053")
            time.sleep(2)

            # 切換到 iframe
            self.log("[步驟 1.1] 切換到查詢 iframe...")
            iframe = self.wait.until(
                EC.presence_of_element_located((By.ID, "content-frame"))
            )
            self.driver.switch_to.frame(iframe)
            self.log("✓ 已切換到 iframe")

            # 步驟 2: 點擊「以部分街路門牌查詢」
            self.log("\n[步驟 2] 點擊「以部分街路門牌查詢」...")
            # 使用 data-type 屬性定位按鈕（更精確）
            partial_road_btn = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-type='doorplate']"))
            )
            partial_road_btn.click()
            self.log("✓ 已點擊查詢按鈕（data-type='doorplate'）")

            # 等待頁面轉換完成，確保進入縣市選擇頁面
            self.log("⏳ 等待進入縣市選擇頁面...")
            time.sleep(3)

            # 步驟 3: 選擇縣市
            self.log(f"\n[步驟 3] 選擇縣市：{city}")

            # 查詢縣市對應的 data-id
            city_id = CITY_ID_MAP.get(city)
            if not city_id:
                self.log(f"✗ 找不到縣市「{city}」的對應 ID")
                self.log(f"支援的縣市：{list(CITY_ID_MAP.keys())}")
                return None

            self.log(f"✓ 找到縣市 ID：{city_id}")

            # 等待 image map 載入
            self.log("⏳ 等待縣市地圖載入...")
            try:
                # 等待圖片載入
                self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "img[usemap]"))
                )
                self.log("✓ 縣市地圖已載入")

                # 尋找對應的 area 元素並點擊
                area_selector = f"area[data-id='{city_id}']"
                self.log(f"⏳ 尋找縣市區域：{area_selector}")

                city_area = self.wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, area_selector))
                )
                self.log(f"✓ 找到縣市區域：{city}")

                # 點擊縣市區域
                city_area.click()
                self.log("✓ 已點擊縣市，等待進入表單頁面...")

            except TimeoutException:
                self.log(f"✗ 找不到縣市「{city}」的地圖區域")
                return None

            # 等待表單頁面完全載入
            time.sleep(3)
            self.log("✓ 表單頁面已載入")

            # 步驟 4: 填寫查詢表單
            self.log(f"\n[步驟 4] 填寫查詢表單...")

            # 填寫區鄉鎮（使用 areaCode 下拉選單）
            try:
                district_select = Select(self.wait.until(
                    EC.presence_of_element_located((By.ID, "areaCode"))
                ))
                # 使用 visible text 選擇（例如："左營區"）
                district_select.select_by_visible_text(district)
                self.log(f"✓ 選擇區鄉鎮：{district}")
            except Exception as e:
                self.log(f"⚠ 無法選擇區鄉鎮：{e}")
                # 嘗試使用部分匹配
                try:
                    for option in district_select.options:
                        if district in option.text:
                            option.click()
                            self.log(f"✓ 選擇區鄉鎮（部分匹配）：{option.text}")
                            break
                except:
                    self.log(f"✗ 無法選擇區鄉鎮：{district}")

            time.sleep(1)

            # 填寫路街名稱（使用 street 欄位，不含段、巷、弄、號）
            # 只填入基本路街名稱，例如："重惠街"
            street_input = self.wait.until(
                EC.presence_of_element_located((By.ID, "street"))
            )
            street_input.clear()
            street_input.send_keys(road)
            self.log(f"✓ 輸入路街名稱：{road}")

            time.sleep(0.5)

            # 解析門牌號碼並填入對應欄位
            if number:
                import re

                # 🔥 先進行全型轉半型轉換（用於 log 顯示）
                fullwidth = "０１２３４５６７８９"
                halfwidth = "0123456789"
                trans_table = str.maketrans(fullwidth, halfwidth)
                number_halfwidth = number.translate(trans_table)

                if number != number_halfwidth:
                    self.log(f"⏳ 解析門牌號碼：{number} → {number_halfwidth}")
                else:
                    self.log(f"⏳ 解析門牌號碼：{number}")

                # 解析門牌號碼的各個部分
                # 可能的格式：106號、1巷2號、5弄10號之2

                # 🔥 解析「號」前，會把「巷/弄」的部分從字串移除，
                #    避免把巷號（例：21巷）誤當成門牌「號」
                num_str = number

                # 巷
                lane_match = re.search(r'([０-９0-9]+)巷', number)
                if lane_match:
                    lane_value = to_fullwidth_number(lane_match.group(1))
                    try:
                        lane_input = self.driver.find_element(By.ID, "lane")
                        # 🔥 改用 JavaScript 設值：send_keys 在新版網站會被欄位事件清掉
                        #    （與「號」「樓」一致），設完再觸發 input 事件並驗證。
                        self.driver.execute_script("arguments[0].value = arguments[1];", lane_input, lane_value)
                        time.sleep(0.2)
                        lane_input.send_keys('')  # 觸發 input 事件
                        actual = lane_input.get_attribute('value')
                        if actual:
                            self.log(f"  ✓ 巷：{actual}")
                        else:
                            self.log(f"  ⚠ 巷欄位填入失敗（ID: lane）")
                    except Exception as e:
                        self.log(f"  ✗ 填入「巷」失敗：{e}")
                    num_str = num_str.replace(lane_match.group(0), '', 1)

                # 弄
                alley_match = re.search(r'([０-９0-9]+)弄', number)
                if alley_match:
                    alley_value = to_fullwidth_number(alley_match.group(1))
                    try:
                        alley_input = self.driver.find_element(By.ID, "alley")
                        # 🔥 同上，改用 JavaScript 設值
                        self.driver.execute_script("arguments[0].value = arguments[1];", alley_input, alley_value)
                        time.sleep(0.2)
                        alley_input.send_keys('')  # 觸發 input 事件
                        actual = alley_input.get_attribute('value')
                        if actual:
                            self.log(f"  ✓ 弄：{actual}")
                        else:
                            self.log(f"  ⚠ 弄欄位填入失敗（ID: alley）")
                    except Exception as e:
                        self.log(f"  ✗ 填入「弄」失敗：{e}")
                    num_str = num_str.replace(alley_match.group(0), '', 1)

                # 號（主要門牌號碼）
                # 🔥 在「已移除巷/弄」的字串上解析：先抓「…號」前的數字；
                #    若無「號」字（如 471之1），再退而取剩餘的第一個數字。
                number_match = re.search(r'([０-９0-9]+)號', num_str)
                if not number_match:
                    number_match = re.search(r'([０-９0-9]+)', num_str)
                if number_match:
                    # 取出數字並轉成半形（網站接受半形數字）
                    main_number_text = number_match.group(1)
                    # 將全形數字轉成半形
                    fullwidth = "０１２３４５６７８９"
                    halfwidth = "0123456789"
                    trans_table = str.maketrans(fullwidth, halfwidth)
                    main_number = main_number_text.translate(trans_table)

                    try:
                        number_input = self.driver.find_element(By.ID, "number")
                        # 先用 JavaScript 設定值
                        self.driver.execute_script(f"arguments[0].value = '{main_number}';", number_input)
                        time.sleep(0.2)
                        # 再用 send_keys 確保觸發事件
                        number_input.send_keys('')  # 觸發 input 事件
                        # 驗證是否填入成功
                        actual_value = number_input.get_attribute('value')
                        if actual_value:
                            self.log(f"  ✓ 已填入號：{actual_value}")
                        else:
                            self.log(f"  ⚠ 號碼欄位填入失敗（ID: number）")
                    except Exception as e:
                        self.log(f"  ✗ 找不到號碼欄位（ID: number）- {e}")

                # 之（門牌號碼後的附號）
                # 🔥 「號」字是可選的，因為有些門牌格式可能是「471之1」而不是「471號之1」
                sub_number_match = re.search(r'號?之([０-９0-9]+)', number)
                if sub_number_match:
                    sub_number_text = sub_number_match.group(1)
                    # 🔥 將半型數字轉換為全型（內政部網站要求）
                    # 注意：如果已經是全型，to_fullwidth_number 不會改變
                    sub_number = to_fullwidth_number(sub_number_text)
                    try:
                        number1_input = self.driver.find_element(By.ID, "number1")
                        # 🔥 改用 JavaScript 填入，跟「號」欄位相同的方式
                        self.driver.execute_script(f"arguments[0].value = '{sub_number}';", number1_input)
                        time.sleep(0.2)
                        # 再用 send_keys 確保觸發事件
                        number1_input.send_keys('')  # 觸發 input 事件
                        # 驗證是否填入成功
                        actual_value = number1_input.get_attribute('value')
                        if actual_value:
                            self.log(f"  ✓ 已填入之：{actual_value}")
                        else:
                            self.log(f"  ⚠ 之欄位填入失敗（ID: number1）")
                    except Exception as e:
                        self.log(f"  ✗ 填入「之」失敗：{e}")

            # 解析樓層（從完整門牌中解析，包含樓層資訊）
            # full_number 範例：106號七樓、1巷2號12樓之1
            address_with_floor = full_number if full_number else number

            # 樓層（支援中文數字和阿拉伯數字）
            # 範例：七樓、12樓、十一樓
            floor_match = re.search(r'([一二三四五六七八九十０-９0-9]+)樓', address_with_floor)
            if floor_match:
                floor_text = floor_match.group(1)
                self.log(f"⏳ 解析樓層：{floor_text}樓")

                # 判斷是中文數字還是阿拉伯數字
                if re.match(r'[一二三四五六七八九十]+', floor_text):
                    # 中文數字，轉換為阿拉伯數字（半形）
                    floor_arabic = chinese_number_to_arabic(floor_text)
                    floor_value = str(floor_arabic)  # 直接使用半形數字
                    self.log(f"  → 轉換：{floor_text} → {floor_arabic} → {floor_value}")
                else:
                    # 已經是數字，轉成半形
                    fullwidth = "０１２３４５６７８９"
                    halfwidth = "0123456789"
                    trans_table = str.maketrans(fullwidth, halfwidth)
                    floor_value = floor_text.translate(trans_table)

                try:
                    floor_input = self.driver.find_element(By.ID, "floor")
                    # 先用 JavaScript 設定值
                    self.driver.execute_script(f"arguments[0].value = '{floor_value}';", floor_input)
                    time.sleep(0.2)
                    # 再用 send_keys 確保觸發事件
                    floor_input.send_keys('')  # 觸發 input 事件
                    # 驗證是否填入成功
                    actual_floor_value = floor_input.get_attribute('value')
                    if actual_floor_value:
                        self.log(f"  ✓ 已填入樓層：{actual_floor_value}")
                    else:
                        self.log(f"  ⚠ 樓層欄位填入失敗（ID: floor）")
                except Exception as e:
                    self.log(f"  ✗ 找不到樓層欄位（ID: floor）- {e}")

            # 樓之（如：7樓之1、15樓之2）
            floor_ext_match = re.search(r'樓之([０-９0-9]+)', address_with_floor)
            if floor_ext_match:
                floor_ext_text = floor_ext_match.group(1)
                # 🔥 將半型數字轉換為全型（內政部網站要求）
                floor_ext = to_fullwidth_number(floor_ext_text)
                try:
                    ext_input = self.driver.find_element(By.ID, "ext")
                    # 🔥 改用 JavaScript 填入，跟「號」和「樓」欄位相同的方式
                    self.driver.execute_script(f"arguments[0].value = '{floor_ext}';", ext_input)
                    time.sleep(0.2)
                    # 再用 send_keys 確保觸發事件
                    ext_input.send_keys('')  # 觸發 input 事件
                    # 驗證是否填入成功
                    actual_value = ext_input.get_attribute('value')
                    if actual_value:
                        self.log(f"  ✓ 已填入樓之：{actual_value}")
                    else:
                        self.log(f"  ⚠ 樓之欄位填入失敗（ID: ext）")
                except Exception as e:
                    self.log(f"  ✗ 填入「樓之」失敗：{e}")

            # 步驟 5: 驗證碼處理（最多重試 3 次）
            self.log("\n[步驟 5] 處理驗證碼...")

            max_retries = 3  # 自動辨識 3 次（每次換新驗證碼且會驗證對錯），仍失敗就轉手動輸入
            captcha_success = False

            for attempt in range(1, max_retries + 1):
                try:
                    captcha_image = self.driver.find_element(By.ID, "captchaImage_captchaKey")
                    captcha_input = self.driver.find_element(By.ID, "captchaInput_captchaKey")
                    search_button = self.driver.find_element(By.ID, "main-query")

                    if attempt > 1:
                        self.log(f"\n⏳ 第 {attempt} 次嘗試辨識驗證碼...")
                        # 🔥 新版改用「產製新驗證碼」按鈕重新整理（不再是點圖片）
                        try:
                            self.driver.find_element(By.ID, "imageBlock_captchaKey").click()
                            time.sleep(1.5)  # 🔥 增加等待時間，確保驗證碼完全載入
                            # 重新取得驗證碼元素（避免 stale element）
                            captcha_image = self.driver.find_element(By.ID, "captchaImage_captchaKey")
                            self.log("  ✓ 已重新整理驗證碼")
                        except:
                            pass
                    else:
                        self.log("⏳ 正在使用 RapidOCR 辨識驗證碼...")
                        time.sleep(1)  # 🔥 首次也等待一下確保驗證碼載入

                    # 辨識驗證碼（開啟 debug 模式查看實際截圖）
                    captcha_text = recognize_captcha(captcha_image, debug=True, log_callback=self.log)

                    if captcha_text and len(captcha_text) >= 4:  # 驗證碼通常是 4-5 位
                        self.log(f"✓ 辨識到驗證碼：{captcha_text}（長度：{len(captcha_text)}）")

                        # 填入驗證碼
                        captcha_input.clear()
                        captcha_input.send_keys(captcha_text)
                        self.log("✓ 已自動填入驗證碼")

                        # 等待一下讓用戶確認
                        time.sleep(1)

                        # 自動點擊搜尋按鈕
                        search_button.click()
                        self.log("✓ 已自動點擊搜尋按鈕")

                        # 等待 2 秒，檢查是否出現驗證碼錯誤對話框
                        time.sleep(2)

                        # 檢查是否有 SweetAlert2 錯誤對話框（驗證碼失敗）
                        try:
                            error_dialog = self.driver.find_elements(By.CSS_SELECTOR, ".swal2-popup.swal2-show")
                            if error_dialog:
                                error_title = self.driver.find_element(By.ID, "swal2-title").text
                                if "驗證碼" in error_title or "驗證失敗" in error_title:
                                    self.log(f"⚠ 驗證碼錯誤：{error_title}")
                                    # 點擊確定按鈕
                                    confirm_btn = self.driver.find_element(By.CSS_SELECTOR, ".swal2-confirm")
                                    confirm_btn.click()
                                    self.log("  ✓ 已點擊確定按鈕")
                                    time.sleep(0.5)

                                    # 驗證碼錯誤，需要重試
                                    if attempt < max_retries:
                                        self.log(f"  → 驗證碼錯誤，將重試...")
                                        continue
                                    else:
                                        self.log(f"  → 已達最大重試次數")
                                        break
                        except:
                            # 沒有錯誤對話框，表示驗證成功
                            pass

                        # 驗證成功
                        captcha_success = True
                        break

                    else:
                        self.log(f"⚠ 辨識結果不完整：'{captcha_text}'（長度：{len(captcha_text) if captcha_text else 0}）")
                        if attempt < max_retries:
                            self.log(f"  → 將重新整理驗證碼並重試...")
                        continue

                except Exception as e:
                    self.log(f"⚠ 第 {attempt} 次辨識失敗：{e}")
                    if attempt < max_retries:
                        self.log(f"  → 將重試...")
                    continue

            # 🔥 如果所有重試都失敗，提示手動輸入 + 給較長等待時間 + 倒數計時
            # 此處邏輯與其他子程式不同：使用者需要先在 Chrome 視窗輸入驗證碼後，
            # **自行點擊「搜尋」按鈕**（不是按 Enter）才會送出查詢。
            if not captcha_success:
                self.log("\n" + "=" * 60)
                self.log("🔴 自動辨識三次都失敗 — 請手動處理")
                self.log("=" * 60)
                self.log("👉 請在 Chrome 視窗中：")
                self.log("   1. 看清楚驗證碼圖片")
                self.log("   2. 在「驗證碼」欄位輸入正確值")
                self.log("   3. **點擊「搜尋」按鈕** 送出（不是按 Enter）")
                self.log("=" * 60)
                # 手動模式給較長等待（90 秒）讓使用者有時間
                max_wait_time = 90
            else:
                # 自動成功時，預設 30 秒等網頁載入
                max_wait_time = 30

            # 等待查詢結果頁面載入（透過檢測特定元素）
            self.log(f"\n[步驟 6] 等待查詢結果（最多 {max_wait_time} 秒）...")
            result_found = False
            start_time = time.time()
            elapsed = 0
            last_countdown = -1

            while time.time() - start_time < max_wait_time:
                try:
                    elapsed = int(time.time() - start_time)
                    remaining = max_wait_time - elapsed

                    # 🔥 新版結果在 DataTables #view-datatable（舊版 jQGrid 已移除）
                    rows = self.driver.find_elements(By.CSS_SELECTOR, "#view-datatable tbody tr")
                    # 有資料的列會有多個 td；查無資料時是單一 td.dt-empty
                    data_rows = [r for r in rows
                                 if len(r.find_elements(By.TAG_NAME, "td")) >= 2]
                    if data_rows:
                        self.log(f"✓ 檢測到查詢結果（{elapsed} 秒）")
                        result_found = True
                        break

                    # 查無資料（td.dt-empty「目前沒有資料」）
                    # 🔥 只在「自動模式」（程式已自動送出查詢）時才據此判定查無資料；
                    #    手動模式下表格一開始就是「目前沒有資料」，使用者還在看/輸入驗證碼、尚未送出，
                    #    若據此就關閉會變成「還沒輸入就自動關閉」的 bug。手動模式只認「有資料列」或逾時。
                    if captcha_success:
                        empty_cell = self.driver.find_elements(By.CSS_SELECTOR, "#view-datatable td.dt-empty")
                        if empty_cell and elapsed >= 5:
                            self.log(f"✗ 查無資料：{empty_cell[0].text.strip()}")
                            return None

                    # 檢查是否有錯誤訊息
                    error_msg = self.driver.find_elements(By.CLASS_NAME, "alert-danger")
                    if len(error_msg) > 0:
                        self.log(f"✗ 查詢失敗：{error_msg[0].text}")
                        return None

                    # 🔥 手動模式：每 5 秒倒數提示一次，並提醒記得點搜尋
                    if not captcha_success:
                        if remaining > 0 and remaining % 5 == 0 and remaining != last_countdown:
                            self.log(f"⏳ 剩餘 {remaining} 秒... 別忘了輸入驗證碼後點「搜尋」按鈕")
                            last_countdown = remaining
                    else:
                        # 自動模式：原本的等待訊息
                        if elapsed % 5 == 0 and elapsed > 0:
                            self.log(f"⏳ 等待中... ({elapsed} 秒)")

                    time.sleep(1)
                except Exception as e:
                    self.log(f"⚠ 檢查結果時發生錯誤：{e}")
                    time.sleep(1)

            if not result_found:
                self.log(f"✗ 查詢逾時（{max_wait_time} 秒內未取得結果）")
                self.log("⚠ 可能原因：")
                self.log("  1. 驗證碼輸入錯誤")
                self.log("  2. 網路連線問題")
                self.log("  3. 查無符合條件的資料")
                return None

            # 步驟 7: 儲存 PDF（如果有指定）
            if save_pdf_path:
                self.save_result_as_pdf(save_pdf_path, pdf_keywords)

            # 步驟 8: 抓取查詢結果
            self.log("\n[步驟 8] 抓取並解析查詢結果...")
            result_data = self._extract_result()

            if result_data:
                self.log(f"\n✅ 查詢成功！")
                self.log(f"  村里：{result_data.get('村里', 'N/A')}")
                self.log(f"  鄰：第{result_data.get('鄰', 'N/A')}鄰")
                self.log("\n✓ 查詢完成，瀏覽器將在 3 秒後關閉...")
                time.sleep(3)  # 讓用戶看到結果
                return result_data
            else:
                self.log("✗ 無法解析查詢結果")
                self.log("⚠ 瀏覽器將在 5 秒後關閉...")
                time.sleep(5)
                return None

        except TimeoutException as e:
            self.log(f"✗ 操作逾時：{e}")
            self.log("⚠ 瀏覽器將在 5 秒後關閉...")
            time.sleep(5)
            return None
        except Exception as e:
            self.log(f"✗ 發生錯誤：{e}")
            self.log("\n錯誤詳情：")
            import traceback
            traceback.print_exc()
            self.log("\n⚠ 瀏覽器將在 5 秒後關閉...")
            time.sleep(5)
            return None

    def save_result_as_pdf(self, output_pdf_path, keywords_to_mark=None):
        """
        將當前查詢結果頁面儲存為 PDF，並用紅框標記關鍵字

        Args:
            output_pdf_path: 輸出 PDF 檔案路徑
            keywords_to_mark: 要標記的關鍵字列表（可選）

        Returns:
            bool: 是否成功儲存
        """
        try:
            import base64
            from PyPDF2 import PdfReader, PdfWriter
            from pdf_annotator import find_text_positions, add_red_boxes_to_pdf
            import tempfile
            import os

            self.log(f"\n[儲存PDF] 正在儲存查詢結果為 PDF...")

            # 1. 使用 Chrome DevTools Protocol 將網頁儲存為 PDF
            self.log("  ⏳ 正在將網頁轉換為 PDF...")

            # 計算列印設定（使用整個頁面）
            result = self.driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "landscape": False,
                "scale": 1.0,
                "paperWidth": 8.27,  # A4 寬度（英吋）
                "paperHeight": 11.69,  # A4 高度（英吋）
                "marginTop": 0.4,
                "marginBottom": 0.4,
                "marginLeft": 0.4,
                "marginRight": 0.4
            })

            # 2. 將 base64 編碼的 PDF 資料寫入臨時檔案
            pdf_data = base64.b64decode(result['data'])
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_pdf.write(pdf_data)
            temp_pdf.close()

            self.log(f"  ✓ 已將網頁儲存為臨時 PDF：{temp_pdf.name}")

            # 3. 如果有指定關鍵字，添加紅框標記
            if keywords_to_mark and len(keywords_to_mark) > 0:
                self.log(f"  ⏳ 正在標記關鍵字：{keywords_to_mark}")

                # 找到關鍵字位置
                positions = find_text_positions(temp_pdf.name, keywords_to_mark)
                self.log(f"  ✓ 找到 {len(positions)} 個匹配位置")

                if positions:
                    # 添加紅框
                    add_red_boxes_to_pdf(temp_pdf.name, output_pdf_path, positions)
                    self.log(f"  ✓ 已添加紅框標記")
                else:
                    # 沒有找到關鍵字，直接複製檔案
                    import shutil
                    shutil.copy2(temp_pdf.name, output_pdf_path)
                    self.log(f"  ⚠ 未找到關鍵字，保存原始 PDF")
            else:
                # 沒有關鍵字，直接複製檔案
                import shutil
                shutil.copy2(temp_pdf.name, output_pdf_path)

            # 4. 刪除臨時檔案
            try:
                os.unlink(temp_pdf.name)
            except:
                pass

            self.log(f"✓ PDF 已儲存：{output_pdf_path}")
            return True

        except Exception as e:
            self.log(f"✗ 儲存 PDF 失敗：{e}")
            import traceback
            self.log(traceback.format_exc())
            return False

    def _extract_result(self):
        """從結果頁面提取村里鄰資訊"""
        try:
            # 等待 jqGrid 表格載入
            time.sleep(2)

            # 🔥 新版 DataTables #view-datatable：取第一筆「資料列」的門牌資料欄（第 2 個 td）
            try:
                first_row_el = None
                for r in self.driver.find_elements(By.CSS_SELECTOR, "#view-datatable tbody tr"):
                    if len(r.find_elements(By.TAG_NAME, "td")) >= 2:  # 資料列有多個 td；查無資料是單一 td.dt-empty
                        first_row_el = r
                        break
                if first_row_el is None:
                    raise NoSuchElementException("無資料列")
                # 🔥 門牌資料欄會被 DataTables 響應式收合，但「整列文字」含完整門牌（含村里/鄰），例：
                #    "1 高雄市楠梓區清豐里025鄰立仁街２１巷１２號二樓 民國113年11月19日 門牌初編"
                full_address = (first_row_el.text or "").strip()
                self.log(f"✓ 找到查詢結果：{full_address}")

                # 解析地址：高雄市左營區福山里065鄰重惠街１０６號七樓
                # 格式：縣市 + 區 + 村里 + 鄰 + 路街門牌
                result_data = {}

                # 使用正則表達式提取村里和鄰
                import re

                # 提取村里（里或村）
                # 支援多種地址格式：
                # - 高雄市左營區福山里 -> 福山里
                # - 台南市安平區平通里 -> 平通里
                # - 屏東縣內埔鄉內田村 -> 內田村
                # 策略：找到最後一個行政區劃（區/鄉/鎮/市），然後捕獲之後的村里名稱
                village_match = re.search(r'[區鄉鎮市]([^區鄉鎮市縣]+[里村])', full_address)
                if village_match:
                    result_data['村里'] = village_match.group(1)
                    self.log(f"  → 村里：{result_data['村里']}")

                # 提取鄰（數字+鄰）
                neighbor_match = re.search(r'([０-９0-9]+)鄰', full_address)
                if neighbor_match:
                    neighbor_num = neighbor_match.group(1)
                    # 去除前導零
                    neighbor_num = neighbor_num.lstrip('０0')
                    result_data['鄰'] = neighbor_num
                    self.log(f"  → 鄰：第{result_data['鄰']}鄰")

                return result_data if result_data else None

            except NoSuchElementException:
                self.log("⚠ 未找到查詢結果")
                # 檢查是否有「查無資料」訊息
                try:
                    empty_msg = self.driver.find_element(By.CSS_SELECTOR, "#view-datatable td.dt-empty")
                    if empty_msg:
                        self.log(f"✗ {empty_msg.text}")
                except:
                    pass
                return None

            # 如果沒有從表格找到，嘗試從頁面文字中找
            if not result_data:
                page_text = self.driver.find_element(By.TAG_NAME, "body").text

                # 使用正則表達式尋找村里鄰資訊
                import re
                village_match = re.search(r'([^,，\s]+[村里])', page_text)
                neighbor_match = re.search(r'第?(\d+)鄰', page_text)

                if village_match:
                    result_data["村里"] = village_match.group(1)
                if neighbor_match:
                    result_data["鄰"] = neighbor_match.group(1)

            return result_data if result_data else None

        except Exception as e:
            self.log(f"提取結果失敗：{e}")
            return None

    def close(self):
        """關閉瀏覽器"""
        try:
            if hasattr(self, 'driver'):
                self.driver.quit()
                self.log("\n✓ 瀏覽器已關閉")
        except:
            pass


def test_query():
    """測試查詢功能"""
    bot = None
    try:
        # 建立查詢機器人（不使用無頭模式，讓使用者可以看到並輸入驗證碼）
        bot = RISQueryBot(headless=False)

        # 測試查詢高雄市左營區重惠街106號
        result = bot.query_by_address(
            city="高雄市",
            district="左營區",
            road="重惠街",
            number="106號"
        )

        if result:
            self.log("\n" + "=" * 60)
            self.log("查詢結果：")
            self.log(f"  村里：{result.get('村里', 'N/A')}")
            self.log(f"  鄰：{result.get('鄰', 'N/A')}")
            self.log("=" * 60)
        else:
            self.log("\n查詢失敗")

    except Exception as e:
        print(f"測試失敗：{e}")
        import traceback
        traceback.print_exc()
    finally:
        if bot:
            input("\n按 Enter 鍵關閉瀏覽器...")
            bot.close()


if __name__ == '__main__':
    test_query()
