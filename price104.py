# 導入所需的庫
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


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window

# 🔥 基準目錄設定（data.json 和工作資料夾的位置）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

import json
import os
import time
import sys
import traceback
import base64
import ctypes
import keyboard  # 用於系統級鍵盤事件

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
price104_window_width = 1024
price104_window_height = 1024
price104_window_x = 0
price104_window_y = 0

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

            price104_window_width = chrome_width_logical
            price104_window_height = chrome_height_logical
            price104_window_x = chrome_x_logical
            price104_window_y = chrome_y_logical
except Exception as e:
    pass

# 通知主程式子程式已完成執行的函數
def notify_main_program():
    print("104實價網已完成執行", flush=True)

# 確保在程式結束時能夠通知主程式
def exit_program():
    notify_main_program()
    sys.exit(0)

# 處理例外情況
def handle_exception(e):
    print(f"發生錯誤: {str(e)}", flush=True)
    print(traceback.format_exc(), flush=True)
    notify_main_program()
    sys.exit(1)

# 104實價網登入類
class Login104Woo:
    def __init__(self):
        self.url = "https://www.104woo.com.tw/price/login.asp?aver=104&sno=1"
        self.credentials_file = "104woo_credentials.json"
        self.driver = None
        
    def setup_driver(self):
        """設置Chrome瀏覽器驅動"""
        try:
            print("正在設置瀏覽器...", flush=True)
            chrome_options = Options()
            chrome_options.add_experimental_option("detach", True)  # 保持瀏覽器開啟
            
            # 添加其他選項以提高穩定性
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            # 🔥 不在啟動參數設定視窗大小，改為啟動後再調整（避免 DPI 改變時崩潰）
            # chrome_options.add_argument("--window-size=1024,1024")
            # chrome_options.add_argument("--window-position=0,0")

            # 禁用 "Chrome 正在被自動化軟體控制" 的提示
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option("useAutomationExtension", False)

            # 禁用 Chrome 背景服務,減少錯誤訊息
            chrome_options.add_argument("--disable-background-networking")
            chrome_options.add_argument("--disable-sync")
            chrome_options.add_argument("--disable-features=ChromeWhatsNewUI")
            chrome_options.add_argument("--log-level=3")  # 只顯示嚴重錯誤
            
            # 禁用密碼保存提示
            prefs = {
                "credentials_enable_service": False,
                "profile.password_manager_enabled": False,
                "autofill.profile_enabled": False
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            # 使用靜默模式安裝驅動
            os.environ['WDM_LOG_LEVEL'] = '0'  # 設置 webdriver-manager 的日誌級別為靜默
            print("正在安裝/更新 Chrome 驅動程式...", flush=True)
            self.driver = create_chrome_driver(options=chrome_options)

            # 🔥 啟動後立即設定視窗大小和位置
            try:
                self.driver.set_window_size(price104_window_width, price104_window_height)
                self.driver.set_window_position(price104_window_x, price104_window_y)
            except Exception as e:
                print(f"[WARNING] 視窗設定失敗: {e}，繼續執行")

            print("瀏覽器設置完成", flush=True)
            return self.driver
        except Exception as e:
            handle_exception(e)
    
    def save_credentials(self, username, password):
        """保存登入憑證到 Windows 認證管理員（透過 credential_helper）"""
        try:
            from credential_helper import save_credentials as _ch_save
            ok = _ch_save("104woo", username, password)
            if ok:
                print("登入資訊已加密儲存到 Windows 認證管理員！", flush=True)
            else:
                print("儲存登入資訊失敗", flush=True)
        except Exception as e:
            print(f"儲存登入資訊時發生錯誤: {str(e)}", flush=True)

    def load_credentials(self):
        """從 Windows 認證管理員讀取登入憑證；首次無資料時自動從舊 JSON 遷移"""
        try:
            from credential_helper import get_credentials as _ch_get
            user, pwd = _ch_get("104woo")
            if user and pwd:
                return {"username": user, "password": pwd}

            # 🔥 第一次使用 keyring 但有舊 JSON：自動遷移後刪除舊檔
            for old_path in [os.path.join("config", self.credentials_file),
                             self.credentials_file]:
                if os.path.exists(old_path):
                    try:
                        with open(old_path, 'r', encoding='utf-8') as f:
                            old = json.load(f)
                        if old.get("username") and old.get("password"):
                            print(f"偵測到舊版 {old_path}，自動遷移到 Windows 認證管理員...", flush=True)
                            self.save_credentials(old["username"], old["password"])
                            try:
                                os.remove(old_path)
                                print(f"已刪除舊憑證檔：{old_path}", flush=True)
                            except Exception as _de:
                                print(f"無法刪除舊憑證檔（可手動刪除）：{old_path}：{_de}", flush=True)
                            return old
                    except Exception as _le:
                        print(f"讀取舊憑證檔失敗：{old_path}：{_le}", flush=True)
            return None
        except Exception as e:
            print(f"讀取登入資訊時發生錯誤: {str(e)}", flush=True)
            return None
            
    def login(self, username, password):
        """登入104實價網"""
        try:
            if not self.driver:
                self.driver = self.setup_driver()
                
            self.driver.get(self.url)

            # 網頁載入後重新設定視窗大小（避免被網站重置）
            try:
                self.driver.set_window_size(price104_window_width, price104_window_height)
                self.driver.set_window_position(price104_window_x, price104_window_y)
            except:
                pass

            # 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
            verify_and_fix_chrome_window(self.driver)

            print("已開啟104實價網登入頁面", flush=True)
            
            try:
                # 等待登入表單加載
                print("等待頁面載入...", flush=True)
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.NAME, "id"))
                )
                
                # 處理登入表單
                username_input = self.driver.find_element(By.NAME, "id")
                password_input = self.driver.find_element(By.NAME, "code")
                
                # 確保元素可見且可交互
                WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.NAME, "id"))
                )
                
                # 填入帳號密碼
                print("正在填入帳號資訊...", flush=True)
                username_input.clear()
                username_input.send_keys(username)
                password_input.clear()
                password_input.send_keys(password)
                
                print("已填入帳號", flush=True)  # 移除顯示帳號名稱
                
                # 點擊登入按鈕
                print("正在點擊登入按鈕...", flush=True)
                login_button = self.driver.find_element(By.XPATH, "//input[@type='image']")
                login_button.click()
                
                # 等待登入成功（可以根據頁面特徵調整）
                time.sleep(2)  # 給予足夠時間進行登入

                print("已成功登入104實價網", flush=True)

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
                            zoom_count = 4  # 175%：按 4 次
                        elif current_dpi >= 1.5:
                            zoom_count = 2  # 150%：按 2 次
                        elif current_dpi >= 1.25 and logical_width < 1200:
                            zoom_count = 2  # 小螢幕 + 125%：縮放到 80%
                        else:
                            zoom_count = 0  # 100%、125%（大螢幕）：不縮放
                except Exception:
                    pass

                # 🔥 執行縮放（參考 V523.py 的實作）
                if zoom_count > 0:
                    try:
                        # 使用 JavaScript 來點擊 body，確保不會因為元素未找到而失敗
                        try:
                            self.driver.execute_script("document.body.click();")
                        except:
                            pass

                        time.sleep(0.5)

                        # 🔥 使用 keyboard 庫發送系統級的 Ctrl + - 按鍵
                        for _ in range(zoom_count):
                            keyboard.press_and_release('ctrl+-')
                            time.sleep(0.5)

                    except Exception:
                        pass  # 靜默處理縮放錯誤

                # 登入成功後，載入搜尋參數
                self.load_search_params()
                
                print("瀏覽器將保持開啟狀態，直到您選擇退出程式", flush=True)
                
                return self.driver
                
            except Exception as e:
                print(f"登入表單處理過程中發生問題: {str(e)}", flush=True)
                print("嘗試使用備用方法...", flush=True)
                
                # 提供使用者手動登入的選項
                print("已為您開啟104實價網登入頁面，請手動完成登入", flush=True)
                print("完成手動登入後，程式將繼續運行", flush=True)
                
                # 等待用戶確認
                print("按下Enter鍵繼續...", flush=True)
                sys.stdout.flush()
                input()
                
                # 登入成功後，載入搜尋參數
                self.load_search_params()
                
                return self.driver
                
        except Exception as e:
            print(f"登入過程中發生錯誤: {str(e)}", flush=True)
            handle_exception(e)
            
    def load_search_params(self):
        """從 data.json 讀取搜尋參數並填入表單"""
        try:
            print("正在載入搜尋參數...", flush=True)
            
            # 檢查 data.json 是否存在
            data_file = get_data_json_path()
            if not os.path.exists(data_file):
                # 檢查是否存在於 config 資料夾中
                data_file = os.path.join("config", "data.json")
                if not os.path.exists(data_file):
                    print("找不到 data.json 檔案，將使用預設搜尋參數", flush=True)
                    return
            
            # 讀取 data.json 檔案
            with open(data_file, 'r', encoding='utf-8') as f:
                data_list = json.load(f)
            
            # 確認是否為陣列格式
            if not isinstance(data_list, list) or len(data_list) == 0:
                print("data.json 格式不正確或沒有資料，將使用預設搜尋參數", flush=True)
                return
                
            # 使用第一筆資料作為預設搜尋條件
            data = data_list[0]
            
            print(f"成功載入搜尋參數，共有 {len(data_list)} 筆資料", flush=True)
            print(f"將使用第一筆資料進行搜尋: {data['city']} {data['area']} {data['section']} {data['lot_number']}", flush=True)
            
            # 等待頁面加載完成並確保我們在正確的頁面上
            try:
                # 首先確認是否已經在搜尋頁面，如果不是，需要先導航到搜尋頁面
                if "login.asp" in self.driver.current_url:
                    print("目前在登入頁面，等待登入完成後再載入搜尋參數", flush=True)
                    return
                    
                # 等待搜尋頁面加載
                try:
                    # 檢查是否已經在搜尋頁面
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.NAME, "acity1"))
                    )
                    print("已進入搜尋頁面，開始填入搜尋參數", flush=True)
                except Exception as e1:
                    # 如果找不到搜尋元素，可能需要導航到搜尋頁面
                    print("未找到搜尋元素，嘗試導航到搜尋頁面", flush=True)
                    # 這裡可以加入導航到搜尋頁面的代碼，如果需要的話
                    return
                
                # 1. 選擇縣市 (acity1)
                if 'city' in data:
                    try:
                        city_select = self.driver.find_element(By.NAME, "acity1")
                        # 使用 Select 類處理下拉選單
                        select_city = Select(city_select)
                        select_city.select_by_visible_text(data['city'])
                        print(f"已選擇縣市: {data['city']}", flush=True)
                        # 選擇縣市後等待頁面更新區域選單
                        time.sleep(1)
                    except Exception as e:
                        print(f"選擇縣市時發生錯誤: {str(e)}", flush=True)
                
                # 2. 選擇區域 (asblo)
                if 'area' in data:
                    try:
                        area_select = self.driver.find_element(By.NAME, "asblo")
                        select_area = Select(area_select)
                        select_area.select_by_visible_text(data['area'])
                        print(f"已選擇區域: {data['area']}", flush=True)
                    except Exception as e:
                        print(f"選擇區域時發生錯誤: {str(e)}", flush=True)
                
                # 3. 設定總價範圍 (預設值)
                try:
                    price_min = self.driver.find_element(By.NAME, "astp1")
                    price_min.clear()
                    price_min.send_keys("0")
                    print("已設定最低總價: 0", flush=True)
                    
                    price_max = self.driver.find_element(By.NAME, "astp2")
                    price_max.clear()
                    price_max.send_keys("999999")
                    print("已設定最高總價: 999999", flush=True)
                except Exception as e:
                    print(f"設定總價範圍時發生錯誤: {str(e)}", flush=True)
                
                # 4. 設定坪數 (預設不限)
                try:
                    size_select = self.driver.find_element(By.NAME, "asize1")
                    select_size = Select(size_select)
                    select_size.select_by_value("0")  # 0 表示不限範圍
                    print("已選擇坪數範圍: 不限範圍", flush=True)
                except Exception as e:
                    print(f"設定坪數時發生錯誤: {str(e)}", flush=True)
                
                # 5. 設定交易類型勾選框 (預設選擇買賣)
                try:
                    # 先將所有選項取消勾選
                    for i in range(1, 5):
                        try:
                            checkbox = self.driver.find_element(By.NAME, f"case{i}")
                            if checkbox.is_selected():
                                checkbox.click()  # 取消勾選
                        except Exception as e1:
                            pass
                    
                    # 勾選買賣 (case2)
                    checkbox = self.driver.find_element(By.NAME, "case2")
                    if not checkbox.is_selected():
                        checkbox.click()  # 勾選
                    print("已勾選交易類型: 買賣", flush=True)
                except Exception as e:
                    print(f"設定交易類型時發生錯誤: {str(e)}", flush=True)
                
                # 自動點擊搜尋按鈕
                try:
                    search_button = self.driver.find_element(By.XPATH, "//input[@type='button' and @onclick='SendSubmit();']")
                    search_button.click()
                    print("已自動點擊搜尋按鈕", flush=True)
                except Exception as e:
                    print(f"點擊搜尋按鈕時發生錯誤: {str(e)}", flush=True)
                
                # 等待搜尋結果頁面加載
                time.sleep(2)
                
                # 顯示可用的全部資料，讓使用者可以選擇查詢
                print("\n所有可用的地段資料:", flush=True)
                for i, item in enumerate(data_list):
                    print(f"[{i+1}] {item['city']} {item['area']} {item['section']} {item['lot_number']}", flush=True)
                
                print("\n您可以輸入數字來選擇要查詢的地段，或輸入 'Q' 退出程式", flush=True)
                print("選擇地段編號: ", flush=True)
                sys.stdout.flush()
                
            except Exception as e:
                print(f"設置搜尋參數時發生錯誤: {str(e)}", flush=True)
                print("請手動設置搜尋條件", flush=True)
        
        except Exception as e:
            print(f"載入搜尋參數時發生錯誤: {str(e)}", flush=True)
            
    def load_data_json(self):
        """載入 data.json 資料"""
        try:
            # 檢查 data.json 是否存在
            data_file = get_data_json_path()
            if not os.path.exists(data_file):
                # 檢查是否存在於 config 資料夾中
                data_file = os.path.join("config", "data.json")
                if not os.path.exists(data_file):
                    print("找不到 data.json 檔案", flush=True)
                    return None

            # 讀取 data.json 檔案
            with open(data_file, 'r', encoding='utf-8') as f:
                data_list = json.load(f)

            # 確認是否為陣列格式
            if not isinstance(data_list, list) or len(data_list) == 0:
                print("data.json 格式不正確或沒有資料", flush=True)
                return None

            print(f"成功載入data.json資料，共有 {len(data_list)} 筆資料", flush=True)
            return data_list

        except Exception as e:
            print(f"載入地段資料時發生錯誤: {str(e)}", flush=True)
            return None

    def get_unique_sections(self, data_list):
        """將地段資料去重，只保留唯一的地段組合"""
        if not data_list:
            return []

        unique_sections = []
        seen_sections = set()

        for data in data_list:
            # 建立唯一識別碼: 縣市-區域-地段
            section_key = f"{data.get('city', '')}-{data.get('area', '')}-{data.get('section', '')}"

            if section_key not in seen_sections:
                seen_sections.add(section_key)
                unique_sections.append(data)

        if len(unique_sections) < len(data_list):
            print(f"去重完成：原始 {len(data_list)} 筆 → 唯一地段 {len(unique_sections)} 筆", flush=True)

        return unique_sections
    
    def search_land_info(self, data):
        """根據地段資料查詢資訊"""
        try:
            # 檢查當前URL並保持在相同的搜尋頁面類型
            current_url = self.driver.current_url
            
            # 保存當前的URL類型（index6 或 index2_allok4）
            if "index6.asp" in current_url:
                base_url = "https://www.104woo.com.tw/price/index6.asp"
            else:
                base_url = "https://www.104woo.com.tw/price/index2_allok4.asp"
                
            # 重新導航到檢測到的基本頁面類型
            print(f"正在為 {data['city']} {data['area']} {data['section']} 開始搜尋...", flush=True)
            self.driver.get(base_url)
            time.sleep(3)  # 給予更長的加載時間
            
            # 等待頁面加載 - 嘗試等待常見元素
            try:
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.TAG_NAME, "select"))
                )
                print("搜尋頁面已加載完成", flush=True)
                
                # 簡化：不再顯示所有表單元素的調試信息
                selects = self.driver.find_elements(By.TAG_NAME, "select")
                inputs = self.driver.find_elements(By.TAG_NAME, "input")
                print("", flush=True)  # 空行
                
            except Exception as e:
                print(f"等待頁面加載時發生錯誤: {str(e)}", flush=True)
            
            # 1. 尋找和選擇縣市下拉選單
            try:
                # 通過名稱或其他屬性嘗試找到縣市選單
                city_select = None
                
                # 方法1: 通過名稱查找
                try:
                    city_select = self.driver.find_element(By.NAME, "acity1")
                except Exception as e1:
                    pass
                
                # 方法2: 通過 XPath 查找包含縣市選項的下拉選單
                if city_select is None:
                    try:
                        city_select = self.driver.find_element(By.XPATH, "//select[option[contains(text(), '高雄市')]]")
                    except Exception as e2:
                        pass
                
                # 方法3: 如果找到多個下拉選單，嘗試第一個
                if city_select is None and len(selects) > 0:
                    city_select = selects[0]
                    
                if city_select:
                    select_city = Select(city_select)
                    select_city.select_by_visible_text(data['city'])
                    print(f"已選擇縣市: {data['city']}", flush=True)
                    time.sleep(2)  # 等待區域選單更新
                else:
                    print("找不到縣市下拉選單", flush=True)
            except Exception as e:
                print(f"選擇縣市時發生錯誤: {str(e)}", flush=True)
            
            # 2. 尋找和選擇區域下拉選單
            try:
                # 通過不同方式尋找區域選單
                area_select = None
                
                # 方法1: 通過名稱查找
                try:
                    area_select = self.driver.find_element(By.NAME, "asblo")
                except Exception as e1:
                    pass
                
                # 方法2: 通過 XPath 查找包含區域選項的下拉選單
                if area_select is None:
                    try:
                        area_select = self.driver.find_element(By.XPATH, f"//select[option[contains(text(), '{data['area']}')]]")
                    except Exception as e2:
                        pass
                
                # 方法3: 如果找到多個下拉選單，嘗試第二個
                if area_select is None and len(selects) > 1:
                    area_select = selects[1]
                    
                if area_select:
                    select_area = Select(area_select)
                    select_area.select_by_visible_text(data['area'])
                    print(f"已選擇區域: {data['area']}", flush=True)
                else:
                    print("找不到區域下拉選單", flush=True)
            except Exception as e:
                print(f"選擇區域時發生錯誤: {str(e)}", flush=True)
            
            # 3. 尋找和選擇「類別」下拉選單，選擇「土地」
            try:
                # 通過不同方式尋找類別選單
                category_select = None
                
                # 方法1: 通過名稱查找
                try:
                    category_select = self.driver.find_element(By.NAME, "akind2")
                except Exception as e1:
                    pass
                
                # 方法2: 通過 XPath 查找包含「土地」選項的下拉選單
                if category_select is None:
                    try:
                        category_select = self.driver.find_element(By.XPATH, "//select[option[contains(text(), '土地')]]")
                    except Exception as e2:
                        pass
                
                # 方法3: 如果找到多個下拉選單，通過位置或其他屬性來猜測
                if category_select is None and len(selects) > 4:  # 通常「類別」是第5個下拉選單
                    try:
                        category_select = selects[4]
                    except Exception as e3:
                        pass
                
                if category_select:
                    select_category = Select(category_select)
                    # 嘗試選擇「土地」選項
                    try:
                        select_category.select_by_visible_text("土地")
                        print("已選擇類別: 土地", flush=True)
                    except Exception as e4:
                        # 如果找不到完全匹配的「土地」選項，嘗試找包含「土地」的選項
                        options = category_select.find_elements(By.TAG_NAME, "option")
                        for option in options:
                            try:
                                if "土地" in option.text:
                                    select_category.select_by_visible_text(option.text)
                                    print(f"已選擇類別: {option.text}", flush=True)
                                    break
                            except Exception as e5:
                                pass
                else:
                    print("找不到類別下拉選單", flush=True)
            except Exception as e:
                print(f"選擇類別時發生錯誤: {str(e)}", flush=True)
            
            # 4. 尋找和填寫關鍵字輸入框(地段)
            try:
                keyword_input = None
                
                # 收集所有文本輸入框
                text_inputs = []
                for inp in inputs:
                    try:
                        if inp.get_attribute("type") == "text":
                            text_inputs.append(inp)
                    except Exception as e1:
                        pass
                
                # 方法1: 通過名稱查找
                for inp in text_inputs:
                    try:
                        if inp.get_attribute("name") == "add1":
                            keyword_input = inp
                            break
                    except Exception as e2:
                        pass
                
                # 方法2: 尋找關鍵字輸入框的可能性 - 通常是比較寬的文本框
                if keyword_input is None:
                    for inp in text_inputs:
                        try:
                            style = inp.get_attribute("style")
                            if "width:160" in style or "width: 160" in style:
                                keyword_input = inp
                                break
                        except Exception as e3:
                            pass
                
                # 方法3: 如果找不到特定的輸入框，嘗試使用最後一個文本輸入框
                if keyword_input is None and len(text_inputs) > 0:
                    keyword_input = text_inputs[-1]
                    
                if keyword_input:
                    keyword_input.clear()
                    # 修正這行錯誤的程式碼
                    keyword_input.send_keys(data['section'])
                    print(f"已輸入關鍵字(地段): {data['section']}", flush=True)
                else:
                    print("找不到關鍵字輸入框", flush=True)
            except Exception as e:
                print(f"輸入關鍵字時發生錯誤: {str(e)}", flush=True)
            
            # 5. 嘗試設定交易類型為買賣
            try:
                checkbox_found = False
                
                # 收集所有的複選框
                checkboxes = []
                for inp in inputs:
                    try:
                        if inp.get_attribute("type") == "checkbox":
                            checkboxes.append(inp)
                    except Exception as e1:
                        pass
                
                # 方法1: 通過名稱或ID查找
                for checkbox in checkboxes:
                    try:
                        name = checkbox.get_attribute("name")
                        id_attr = checkbox.get_attribute("id")
                        if name == "case2" or id_attr == "case2":
                            if not checkbox.is_selected():
                                checkbox.click()
                            checkbox_found = True
                            break
                    except Exception as e2:
                        pass
                
                # 方法2: 通過附近的"買賣"文字來查找
                if not checkbox_found:
                    try:
                        checkbox = self.driver.find_element(By.XPATH, "//input[@type='checkbox']/following::font[contains(text(), '買賣')]/preceding::input[@type='checkbox'][1]")
                        if not checkbox.is_selected():
                            checkbox.click()
                        checkbox_found = True
                    except Exception as e3:
                        pass
                
                # 方法3: 如果找到了複選框但無法通過上述方法確定，嘗試第二個複選框(通常買賣選項是第二個)
                if not checkbox_found and len(checkboxes) >= 2:
                    try:
                        checkbox = checkboxes[1]
                        if not checkbox.is_selected():
                            checkbox.click()
                        checkbox_found = True
                    except Exception as e4:
                        pass
                        
                if checkbox_found:
                    print("已勾選交易類型: 買賣", flush=True)
                else:
                    print("找不到或無法選擇交易類型複選框", flush=True)
            except Exception as e:
                print(f"設定交易類型時發生錯誤: {str(e)}", flush=True)
            
            # 6. 尋找和點擊搜尋按鈕
            try:
                search_button = None
                
                # 收集所有的按鈕
                buttons = []
                for inp in inputs:
                    try:
                        if inp.get_attribute("type") == "button" or inp.get_attribute("type") == "submit" or inp.get_attribute("type") == "image":
                            buttons.append(inp)
                    except Exception as e1:
                        pass
                
                # 方法1: 通過 onclick 屬性查找
                for button in buttons:
                    try:
                        onclick = button.get_attribute("onclick")
                        if onclick and "SendSubmit" in onclick:
                            search_button = button
                            break
                    except Exception as e2:
                        pass
                
                # 方法2: 通過按鈕值查找
                if search_button is None:
                    for button in buttons:
                        try:
                            value = button.get_attribute("value")
                            if value and "搜尋" in value:
                                search_button = button
                                break
                        except Exception as e3:
                            pass
                
                # 方法3: 如果找不到特定按鈕，嘗試點擊最後一個按鈕
                if search_button is None and len(buttons) > 0:
                    search_button = buttons[-1]
                    
                if search_button:
                    search_button.click()
                    print("已點擊搜尋按鈕", flush=True)
                    time.sleep(3)  # 等待搜尋結果
                else:
                    print("找不到搜尋按鈕", flush=True)
            except Exception as e:
                print(f"點擊搜尋按鈕時發生錯誤: {str(e)}", flush=True)
            
            print(f"已完成 {data['city']} {data['area']} {data['section']} 的搜尋", flush=True)
            
            # 使用迴圈讓用戶可以反復檢查結果並保存PDF
            while True:
                # 顯示操作選項並等待使用者選擇
                print("\n1.請檢查結果，或是直接於網頁修改資料行查詢，完成後按ENTER繼續執行存檔功能", flush=True)
                print("2.退出程式請按\"Q\"", flush=True)
                sys.stdout.flush()
                
                # 等待使用者輸入
                user_choice = input().strip().lower()
                
                # 處理使用者選擇
                if user_choice == 'q':
                    print("正在關閉程式...", flush=True)
                    if self.driver:
                        self.driver.quit()
                    exit_program()
                    return False  # 返回 False 表示退出程式
                else:  # 無論是按 ENTER 還是輸入其他內容
                    # 將網頁儲存為PDF
                    self.save_page_as_pdf()
                    # 繼續循環，再次顯示選項，而不是返回選單
                    continue
            
        except Exception as e:
            print(f"搜尋地段資訊時發生錯誤: {str(e)}", flush=True)
            print("\n按 ENTER 返回選單或按 Q 退出程式...", flush=True)
            sys.stdout.flush()
            
            user_choice = input().strip().lower()
            if user_choice == 'q':
                if self.driver:
                    self.driver.quit()
                exit_program()
                return False
            return True  # 返回選單
    
    def save_page_as_pdf(self):
        """將當前頁面儲存為PDF，使用data.json中的第一筆資料建立目錄"""
        try:
            # 確保我們有第一筆資料
            if not hasattr(self, 'first_data'):
                # 嘗試載入資料
                data_list = self.load_data_json()
                if data_list and len(data_list) > 0:
                    self.first_data = data_list[0]
                else:
                    print("找不到資料，無法建立目錄", flush=True)
                    return
            
            # 獲取第一筆資料
            data = self.first_data
            
            # 建立儲存路徑 (始終使用第一筆資料)
            folder_name = f"{data['area']}{data['section']}-{data['lot_number']}"  # 例如: 鳳山區鳳甲段-3-35
            sub_folder = "3.行情"
            # 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
            save_path = os.path.join(get_work_folder(folder_name), sub_folder)
            
            # 確保目錄存在
            os.makedirs(save_path, exist_ok=True)
            
            # 準備檔案名稱的基本部分
            filename_base = f"104實價登錄-{data['area']}{data['section']}"  # 例如：104實價登錄-鳳山區鳳甲段
            
            # 檢查現有檔案，決定使用哪個字母
            existing_files = [f for f in os.listdir(save_path) if f.startswith(filename_base) and f.endswith('.pdf')]
            
            # 如果沒有現有檔案，使用 A
            if not existing_files:
                next_letter = 'A'
            else:
                # 找出已存在的最後一個字母
                letters = []
                for file in existing_files:
                    # 嘗試從檔名中提取字母
                    try:
                        letter = file.split('-')[-1].split('.')[0]  # 例如從 "104實價登錄-鳳山區鳳甲段-B.pdf" 提取 "B"
                        if len(letter) == 1 and letter.isalpha():
                            letters.append(letter)
                    except Exception as e1:
                        pass
                
                if not letters:
                    next_letter = 'A'
                else:
                    # 找出字母中最大的，並取得下一個字母
                    last_letter = max(letters)
                    if last_letter >= 'Z':
                        # 如果已經到 Z，則循環回 A 並加上數字
                        existing_numbered_files = [f for f in existing_files if f.endswith('Z.pdf') or f[-5].isdigit()]
                        if not existing_numbered_files:
                            next_letter = 'A1'
                        else:
                            # 找出最大的數字
                            max_num = 0
                            for file in existing_numbered_files:
                                try:
                                    if file[-5] == 'Z':  # 如果檔名以 Z.pdf 結尾
                                        num = 1
                                    else:
                                        num_str = ''.join(c for c in file.split('-')[-1].split('.')[0][1:] if c.isdigit())
                                        num = int(num_str) if num_str else 0
                                    max_num = max(max_num, num)
                                except Exception as e2:
                                    pass
                            next_letter = f'A{max_num + 1}'
                    else:
                        # 取得下一個字母
                        next_letter = chr(ord(last_letter) + 1)
            
            # 最終的檔案名稱和完整路徑
            pdf_filename = f"{filename_base}-{next_letter}.pdf"
            pdf_path = os.path.join(save_path, pdf_filename)
            
            # print(f"正在將網頁儲存為PDF: {pdf_path}", flush=True)
            
            # 使用Chrome的Print to PDF功能
            try:
                # 設定列印選項
                print_options = {
                    'landscape': False,
                    'displayHeaderFooter': False,
                    'printBackground': True,
                    'preferCSSPageSize': True,
                    'margin': {
                        'top': 0,
                        'bottom': 0,
                        'left': 0,
                        'right': 0
                    },
                    'scale': 1.0
                }
                
                # 執行列印指令 (使用CDP協議)
                pdf_data = self.driver.execute_cdp_cmd("Page.printToPDF", print_options)
                
                # 寫入檔案
                with open(pdf_path, 'wb') as f:
                    f.write(base64.b64decode(pdf_data['data']))
                    
                print(f"\033[93mPDF已成功儲存至: {pdf_path}\033[0m", flush=True)
                return
                
            except Exception as e:
                print(f"使用CDP指令儲存PDF時發生錯誤: {str(e)}", flush=True)
                print("嘗試使用備用方法...", flush=True)
            
            # 備用方法: 使用截圖 (不是真正的PDF，但至少可以保存頁面內容)
            try:
                # 使用截圖方法
                screenshot_path = os.path.join(save_path, f"{filename_base}-{next_letter}_截圖.png")
                self.driver.save_screenshot(screenshot_path)
                print(f"無法保存為PDF，已改為儲存截圖: {screenshot_path}", flush=True)
                return
                
            except Exception as backup_e:
                print(f"所有儲存方法都失敗: {str(backup_e)}", flush=True)
                print("請手動儲存網頁為PDF", flush=True)
                
        except Exception as e:
            print(f"儲存PDF過程中發生錯誤: {str(e)}", flush=True)
            print("請手動儲存網頁為PDF", flush=True)
    
    def auto_batch_mode(self):
        """自動批次查詢模式 - 自動對所有唯一地段執行查詢並存檔"""
        try:
            # 載入地段資料
            data_list = self.load_data_json()
            if not data_list:
                print("未找到有效的地段資料，進入一般瀏覽模式", flush=True)
                self.normal_mode()
                return

            # 保存第一筆資料，用於建立目錄
            self.first_data = data_list[0]

            # 去重，只保留唯一地段
            unique_sections = self.get_unique_sections(data_list)

            print(f"\n========== 開始自動批次查詢 ==========", flush=True)
            print(f"共有 {len(unique_sections)} 個唯一地段需要查詢", flush=True)

            # 自動逐一查詢每個唯一地段
            for i, data in enumerate(unique_sections, 1):
                print(f"\n>>> [{i}/{len(unique_sections)}] 正在處理: {data['city']} {data['area']} {data['section']}", flush=True)

                # 執行搜尋
                self.search_land_info_auto(data)

                # 自動存檔
                print(f"正在自動儲存 PDF...", flush=True)
                self.save_page_as_pdf()

                print(f"✓ [{i}/{len(unique_sections)}] 完成: {data['city']} {data['area']} {data['section']}", flush=True)

            # 所有自動查詢完成，進入手動調整模式
            print(f"\n========== 自動批次查詢已完成 ==========", flush=True)
            print(f"所有 {len(unique_sections)} 個地段已查詢並存檔", flush=True)
            print(f"\n現在您可以手動調整網頁查詢條件，並重複存檔", flush=True)

            # 進入手動調整與存檔循環
            while True:
                print("\n選擇操作：", flush=True)
                print("[ENTER] - 儲存目前網頁為 PDF", flush=True)
                print("[q] - 退出程式", flush=True)
                sys.stdout.flush()

                choice = input().strip().lower()

                if choice == 'q':
                    print("正在關閉程式...", flush=True)
                    if self.driver:
                        self.driver.quit()
                    exit_program()
                    break
                else:
                    # 按 ENTER 或其他鍵，儲存 PDF
                    self.save_page_as_pdf()

        except KeyboardInterrupt:
            print("\n批次模式已被中斷", flush=True)
            if self.driver:
                self.driver.quit()
            exit_program()
        except Exception as e:
            print(f"批次模式發生錯誤: {str(e)}", flush=True)
            self.normal_mode()

    def search_land_info_auto(self, data):
        """自動查詢地段資訊（無需使用者互動）"""
        try:
            # 檢查當前URL並保持在相同的搜尋頁面類型
            current_url = self.driver.current_url

            # 保存當前的URL類型（index6 或 index2_allok4）
            if "index6.asp" in current_url:
                base_url = "https://www.104woo.com.tw/price/index6.asp"
            else:
                base_url = "https://www.104woo.com.tw/price/index2_allok4.asp"

            # 重新導航到檢測到的基本頁面類型
            self.driver.get(base_url)
            time.sleep(3)

            # 等待頁面加載
            try:
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.TAG_NAME, "select"))
                )
                selects = self.driver.find_elements(By.TAG_NAME, "select")
                inputs = self.driver.find_elements(By.TAG_NAME, "input")
            except Exception as e:
                print(f"等待頁面加載時發生錯誤: {str(e)}", flush=True)
                return

            # 1. 選擇縣市
            try:
                city_select = None
                try:
                    city_select = self.driver.find_element(By.NAME, "acity1")
                except:
                    pass
                if city_select is None and len(selects) > 0:
                    city_select = selects[0]

                if city_select:
                    Select(city_select).select_by_visible_text(data['city'])
                    print(f"已選擇縣市: {data['city']}", flush=True)
                    time.sleep(2)
            except Exception as e:
                print(f"選擇縣市時發生錯誤: {str(e)}", flush=True)

            # 2. 選擇區域
            try:
                area_select = None
                try:
                    area_select = self.driver.find_element(By.NAME, "asblo")
                except:
                    pass
                if area_select is None and len(selects) > 1:
                    area_select = selects[1]

                if area_select:
                    Select(area_select).select_by_visible_text(data['area'])
                    print(f"已選擇區域: {data['area']}", flush=True)
            except Exception as e:
                print(f"選擇區域時發生錯誤: {str(e)}", flush=True)

            # 3. 選擇「類別」為「土地」
            try:
                category_select = None
                try:
                    category_select = self.driver.find_element(By.NAME, "akind2")
                except:
                    pass
                if category_select is None and len(selects) > 4:
                    category_select = selects[4]

                if category_select:
                    try:
                        Select(category_select).select_by_visible_text("土地")
                        print("已選擇類別: 土地", flush=True)
                    except:
                        pass
            except Exception as e:
                print(f"選擇類別時發生錯誤: {str(e)}", flush=True)

            # 4. 輸入地段關鍵字
            try:
                keyword_input = None
                text_inputs = [inp for inp in inputs if inp.get_attribute("type") == "text"]

                for inp in text_inputs:
                    if inp.get_attribute("name") == "add1":
                        keyword_input = inp
                        break
                if keyword_input is None and len(text_inputs) > 0:
                    keyword_input = text_inputs[-1]

                if keyword_input:
                    keyword_input.clear()
                    keyword_input.send_keys(data['section'])
                    print(f"已輸入關鍵字(地段): {data['section']}", flush=True)
            except Exception as e:
                print(f"輸入關鍵字時發生錯誤: {str(e)}", flush=True)

            # 5. 勾選「買賣」
            try:
                checkboxes = [inp for inp in inputs if inp.get_attribute("type") == "checkbox"]
                for checkbox in checkboxes:
                    if checkbox.get_attribute("name") == "case2":
                        if not checkbox.is_selected():
                            checkbox.click()
                        print("已勾選交易類型: 買賣", flush=True)
                        break
            except Exception as e:
                print(f"設定交易類型時發生錯誤: {str(e)}", flush=True)

            # 6. 點擊搜尋按鈕
            try:
                search_button = None
                buttons = [inp for inp in inputs if inp.get_attribute("type") in ["button", "submit", "image"]]

                for button in buttons:
                    onclick = button.get_attribute("onclick")
                    if onclick and "SendSubmit" in onclick:
                        search_button = button
                        break
                if search_button is None and len(buttons) > 0:
                    search_button = buttons[-1]

                if search_button:
                    search_button.click()
                    print("已點擊搜尋按鈕", flush=True)
                    time.sleep(3)
            except Exception as e:
                print(f"點擊搜尋按鈕時發生錯誤: {str(e)}", flush=True)

        except Exception as e:
            print(f"自動搜尋地段資訊時發生錯誤: {str(e)}", flush=True)
    
    def normal_mode(self):
        """一般瀏覽模式"""
        try:
            print("\n進入一般瀏覽模式", flush=True)
            print("您可以自由瀏覽104實價網，輸入 'exit' 退出程式", flush=True)
            
            while True:
                print("輸入 'exit' 以關閉程式，或按Enter繼續瀏覽: ", flush=True)
                sys.stdout.flush()
                choice = input().strip().lower()
                if choice == 'exit':
                    print("正在關閉程式...", flush=True)
                    if self.driver:
                        self.driver.quit()
                    exit_program()
                    break
                    
        except KeyboardInterrupt:
            print("\n一般模式已被中斷", flush=True)
            if self.driver:
                self.driver.quit()
            exit_program()
        except Exception as e:
            print(f"一般模式發生錯誤: {str(e)}", flush=True)
            if self.driver:
                self.driver.quit()
            exit_program()
    
    def run(self):
        """運行主要登入流程"""
        try:
            print("歡迎使用104實價網自動登入系統", flush=True)
            print("-" * 40, flush=True)
            
            # 🔥 從 Windows 認證管理員讀取或跳對話框輸入（含舊 JSON 自動遷移）
            saved_credentials = self.load_credentials()
            if saved_credentials:
                print("自動使用已儲存的帳號登入...", flush=True)
                self.login(saved_credentials['username'], saved_credentials['password'])
                self.auto_batch_mode()
                return

            # 沒有任何已儲存資料 → 跳 GUI 對話框讓使用者輸入並儲存
            try:
                from credential_helper import get_or_prompt_credentials
                username, password = get_or_prompt_credentials("104woo", "104實價網")
            except Exception as _e:
                print(f"❌ 載入 credential_helper 失敗：{_e}", flush=True)
                username, password = None, None

            if not username or not password:
                print("❌ 未取得帳密，程式中止", flush=True)
                return

            self.login(username, password)
            self.auto_batch_mode()
                
        except KeyboardInterrupt:
            print("\n程式已被用戶中斷", flush=True)
            if self.driver:
                self.driver.quit()
            exit_program()
        except Exception as e:
            handle_exception(e)

# 主程序
if __name__ == "__main__":
    try:
        # 清除螢幕並設定標題
        print("*" * 50, flush=True)
        print("        104實價網自動登入系統", flush=True)
        print("*" * 50, flush=True)
        print("啟動104實價網自動登入系統...", flush=True)
        print("", flush=True)  # 空行
        
        login_system = Login104Woo()
        login_system.run()
    except Exception as e:
        handle_exception(e)