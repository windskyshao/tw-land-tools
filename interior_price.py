# -*- coding: utf-8 -*-
"""
InteriorPriceSearch — Selenium 自動查詢（可手動延伸版）
2025-05-25‑update 15 - 修復截圖問題專門版
----------------------------------------------------------------
* 🖼 修復截圖只顯示單一項目的問題
  - 保持1024x1024視窗大小
  - 使用標準截圖方法而非CDP
  - 增加frame切換確保在正確的框架內
  - 等待表格完全載入後再截圖
----------------------------------------------------------------
"""

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


import base64
import datetime as dt
import json
import os
import re
import sys
import time
import threading
import queue
import ctypes
import keyboard  # 用於系統級鍵盤事件
from typing import Optional, Tuple, List, Dict

from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window

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
interior_window_width = 1024
interior_window_height = 1024
interior_window_x = 0
interior_window_y = 0

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

            interior_window_width = chrome_width_logical
            interior_window_height = chrome_height_logical
            interior_window_x = chrome_x_logical
            interior_window_y = chrome_y_logical
except Exception as e:
    pass

YELLOW = "\033[33m"
GREEN = "\033[32m"
CYAN = "\033[36m"
MAGENTA = "\033[35m"
BOLD = "\033[1m"
RESET = "\033[0m"

# 全局變量 - 輸入隊列和退出標誌
input_queue = queue.Queue()
should_exit = False  # 全局退出標誌，確保所有線程都能看到

# 格式化輸出函數
def yprint(msg: str) -> None:
    """打印黃色消息，並立即刷新輸出"""
    print(f"{YELLOW}{msg}{RESET}", flush=True)

def gprint(msg: str) -> None:
    """打印綠色消息，用於數據輸出"""
    print(f"{GREEN}{msg}{RESET}", flush=True)

def cprint(msg: str) -> None:
    """打印青色消息，用於重要提示"""
    print(f"{CYAN}{msg}{RESET}", flush=True)

def mprint(msg: str) -> None:
    """打印洋紅色消息，用於強調地號"""
    print(f"{MAGENTA}{msg}{RESET}", flush=True)

def bprint(msg: str) -> None:
    """打印粗體消息，用於標題"""
    print(f"{BOLD}{msg}{RESET}", flush=True)

def display_lot_numbers(items: List[Dict]):
    """顯示所有地號信息"""
    # 提取所有不同的城市、區域、地段組合
    locations = {}
    for item in items:
        city = item.get('city', 'N/A')
        area = item.get('area', 'N/A')
        section = item.get('section', 'N/A')
        lot_number = item.get('lot_number', 'N/A')
        
        key = (city, area, section)
        if key not in locations:
            locations[key] = []
        
        # 只添加不重複的地號
        if lot_number not in locations[key]:
            locations[key].append(lot_number)
    
    # 打印結果
    bprint("\n======= 所有地號資料 =======")
    counter = 1
    for (city, area, section), lot_numbers in locations.items():
        cprint(f"【位置】{city} {area} {section}")
        for lot_number in sorted(lot_numbers):
            mprint(f"[{counter}] {city} {area} {section} {lot_number}")
            counter += 1
    
    # 打印整合後的搜尋參數
    bprint("\n======= 整合後的搜尋參數 =======")
    counter = 1
    for (city, area, section), _ in locations.items():
        gprint(f"[{counter}] {city} {area} {section}")
        counter += 1
    
    print("")  # 空行


def _sanitize(name: str) -> str:
    """移除檔名中非法字元"""
    return re.sub(r"[\\/:*?\"<>|]", "", name)


def _unique_path(path: str) -> str:
    """若檔案已存在，於檔名前加 -A, -B, -C...等字母後綴"""
    if not os.path.exists(path):
        return path
        
    root, ext = os.path.splitext(path)
    # 使用ASCII碼從A開始
    for i in range(65, 91):  # A-Z的ASCII碼
        suffix = chr(i)  # 轉換ASCII碼為字母
        new_path = f"{root}-{suffix}{ext}"
        if not os.path.exists(new_path):
            return new_path
            
    # 如果A-Z都已使用，繼續使用AA, AB等（極少情況）
    for i in range(65, 91):
        for j in range(65, 91):
            suffix = chr(i) + chr(j)
            new_path = f"{root}-{suffix}{ext}"
            if not os.path.exists(new_path):
                return new_path
                
    # 最終返回一個帶隨機後綴的路徑
    import uuid
    return f"{root}-{uuid.uuid4().hex[:6]}{ext}"


class InteriorPriceSearch:
    """控制瀏覽器，自動／手動查詢實價登錄並截圖"""

    def __init__(self, base_dir: str) -> None:
        self.url = "https://lvr.land.moi.gov.tw/"
        self.base_dir = base_dir
        self.driver: Optional[webdriver.Chrome] = None
        self.wait: Optional[WebDriverWait] = None
        self.last_area = ""
        self.last_section = ""  # 新增: 保存上次成功讀取的地段
    
    # ---------------- Driver ----------------
    def setup_driver(self) -> None:
        yprint("準備配置Chrome瀏覽器選項...")
        
        # 改進瀏覽器設置以減少閃爍
        prefs = {
            "profile.default_content_setting_values.notifications": 1,
            "browser.enable_spellchecking": False,  # 減少UI元素
            "browser.enable_autospellcorrection": False
        }
        
        opts = Options()
        # 🔥 不在啟動參數設定視窗大小，改為啟動後再調整（避免 DPI 改變時崩潰）
        # opts.add_argument("--window-position=0,0")
        # opts.add_argument("--window-size=1024,1024")
        
        # 減少閃爍的其他選項
        opts.add_argument("--disable-animations")  # 禁用動畫
        opts.add_argument("--disable-infobars")    # 禁用信息欄
        opts.add_argument("--disable-extensions")  # 禁用擴展
        
        # 原有的設置
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_experimental_option("prefs", prefs)
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)
        
        yprint("正在初始化瀏覽器驅動...")
        
        # 使用安靜模式啟動瀏覽器
        self.driver = create_chrome_driver(options=opts)
        self.wait = WebDriverWait(self.driver, 20)

        # 🔥 啟動後立即設定視窗大小和位置
        try:
            self.driver.set_window_size(interior_window_width, interior_window_height)
            self.driver.set_window_position(interior_window_x, interior_window_y)
        except Exception as e:
            print(f"[WARNING] 視窗設定失敗: {e}，繼續執行")

        # 🔥 啟動 Chrome 死亡 watchdog：使用者按 X 關掉 Chrome 時自動觸發退出
        # 不管子程式正在做什麼（填表、截圖、等輸入...），watchdog 都能偵測
        self._start_chrome_watchdog()

        yprint("瀏覽器驅動已初始化，正在配置瀏覽器...")
        
        # 隱藏 webdriver 痕跡
        self.driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {
                "source": (
                    "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});"
                    "Object.defineProperty(navigator,'languages',{get:()=>['zh-TW','zh','en']});"
                    "Object.defineProperty(navigator,'plugins',{get:()=>[1,2,3]});"
                )
            },
        )
        
        yprint("瀏覽器配置完成")

    # ---------------- Helpers ----------------
    def switch_to_inner_frame(self) -> bool:
        try:
            self.driver.switch_to.default_content()
            frames = self.driver.find_elements(By.TAG_NAME, "frame")
            if frames:
                self.driver.switch_to.frame(frames[0])
                return True
            else:
                return False
        except Exception as e:
            yprint(f"切換框架時出錯: {e}")
            return False
    
    def _get_element_safely(self, by, value, wait_time=1, visible=False):
        """安全地嘗試獲取元素，如果失敗則返回None"""
        try:
            if visible:
                # 等待元素可見
                element = WebDriverWait(self.driver, wait_time).until(
                    EC.visibility_of_element_located((by, value))
                )
            else:
                # 等待元素存在，不管是否可見
                element = WebDriverWait(self.driver, wait_time).until(
                    EC.presence_of_element_located((by, value))
                )
            return element
        except (NoSuchElementException, TimeoutException):
            return None
        except Exception as e:
            yprint(f"獲取元素 {value} 時出錯: {e}")
            return None

    def _read_area_section(self, manual_mode=False) -> Tuple[str, str]:
        """從 DOM 讀取目前選取的鄉鎮市區 (area) 與輸入的地段 (section)
        
        Args:
            manual_mode: 如果為True，用於手動模式的檔名（區域+地段）
        """
        area = section = ""
        
        # 先嘗試切換到內部框架
        if not self.switch_to_inner_frame():
            yprint("無法切換到內部框架，嘗試在當前頁面讀取資訊...")
        
        # 嘗試在查詢頁面讀取區域和地段
        try:
            # 檢查是否在查詢頁面 - 查找地區下拉選單（l_town 或 p_town）
            town_select = None
            for town_id in ["l_town", "p_town", "l_town2"]:
                town_select = self._get_element_safely(By.ID, town_id, wait_time=0.5)
                if town_select:
                    break
            
            if town_select:
                try:
                    # 讀取地區
                    selected_option = Select(town_select).first_selected_option
                    area = selected_option.text.strip()
                    cprint(f"從查詢頁面讀取到區域: {area}")
                except:
                    yprint("無法讀取地區選項")
                
            # 嘗試讀取門牌/社區名稱（地段）- p_build 或 p_build2
            build_input = None
            for build_id in ["p_build", "p_build2"]:
                build_input = self._get_element_safely(By.ID, build_id, wait_time=0.5)
                if build_input:
                    break
                    
            if build_input:
                section = build_input.get_attribute("value").strip()
                if section:
                    cprint(f"從查詢頁面讀取到地段: {section}")
            
            # 如果在查詢頁面沒找到，可能在搜尋結果頁面
            if not area and not section:
                cprint("未檢測到查詢頁面表單，嘗試從結果頁面讀取...")
                
                # 嘗試從結果頁面的標題或其他元素讀取信息
                try:
                    # 查找可能包含地區資訊的元素
                    # 實價登錄結果頁面通常會顯示查詢條件
                    # 尋找包含"區"的元素
                    area_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(),'區') and not(contains(@id,'customCheck'))]")
                    for elem in area_elements:
                        text = elem.text
                        # 嘗試匹配常見的區域格式
                        area_match = re.search(r'([^縣市]+區)', text)
                        if area_match:
                            potential_area = area_match.group(1)
                            # 過濾掉一些非區域的文字
                            if len(potential_area) <= 6 and "地區" not in potential_area:
                                area = potential_area
                                cprint(f"從結果頁面解析出區域: {area}")
                                break
                    
                    # 尋找包含地段資訊的元素
                    # 通常會有類似 "門牌：xxx" 或 "社區：xxx" 的格式
                    section_patterns = [
                        r'門牌[：:]([^\s]+)',
                        r'社區[：:]([^\s]+)',
                        r'地址[：:].*?([^\s]+段)',
                        r'地段[：:]([^\s]+)'
                    ]
                    
                    all_text_elements = self.driver.find_elements(By.XPATH, "//*[text()]")
                    for elem in all_text_elements:
                        text = elem.text
                        for pattern in section_patterns:
                            match = re.search(pattern, text)
                            if match:
                                section = match.group(1)
                                cprint(f"從結果頁面解析出地段: {section}")
                                break
                        if section:
                            break
                            
                except Exception as e:
                    yprint(f"從結果頁面讀取時出錯: {e}")
        except Exception as e:
            yprint(f"嘗試讀取區域和地段時出錯: {e}")
        
        # 使用上次記錄的值作為備用
        if not area and self.last_area:
            area = self.last_area
            cprint(f"使用上次記錄的區域: {area}")
        
        if not section and self.last_section:
            section = self.last_section
            cprint(f"使用上次記錄的地段: {section}")
        
        # 更新記錄的值（如果讀取到了新值）
        if area:
            self.last_area = area
        if section:
            self.last_section = section

        return area, section

    def wait_for_results_load(self):
        """等待搜尋結果完全載入"""
        try:
            yprint("等待搜尋結果載入...")
            
            # 切換到正確的frame
            self.switch_to_inner_frame()
            
            # 等待基本頁面載入
            time.sleep(2)
            
            # 嘗試等待表格或列表元素
            try:
                # 等待任何一個可能的結果元素出現
                WebDriverWait(self.driver, 10).until(
                    lambda driver: 
                        driver.find_elements(By.CSS_SELECTOR, "table") or
                        driver.find_elements(By.CSS_SELECTOR, ".table") or
                        driver.find_elements(By.CSS_SELECTOR, "[class*='list']") or
                        driver.find_elements(By.CSS_SELECTOR, "tr") or
                        driver.find_elements(By.XPATH, "//table") or
                        driver.find_elements(By.XPATH, "//div[contains(@class,'result')]")
                )
                yprint("檢測到結果元素")
            except TimeoutException:
                yprint("未能檢測到明確的結果元素，繼續執行...")
            
            # 額外等待動態內容
            time.sleep(2)
            
            # 嘗試滾動以載入所有內容
            try:
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(0.5)
                self.driver.execute_script("window.scrollTo(0, 0);")
            except:
                pass
                
        except Exception as e:
            yprint(f"等待結果載入時出現警告: {e}")

    # ---------------- Capture ----------------
    def capture_page(self, area: str = "", section: str = "", postfix: Optional[str] = None, manual_mode: bool = False) -> None:
        # 等待頁面載入
        yprint("準備截圖...")
        self.wait_for_results_load()
        
        # 讀取區域和地段
        current_area, current_section = self._read_area_section()
        
        if manual_mode:
            # 手動模式：使用區域+地段的組合
            area_name = current_area or ""
            section_name = current_section or ""
            
            # 組合檔名：如果有區域和地段都有值，就組合使用
            if area_name and section_name:
                postfix = postfix or _sanitize(f"{area_name}{section_name}")
            elif area_name:
                postfix = postfix or _sanitize(area_name)
            elif section_name:
                postfix = postfix or _sanitize(section_name)
            else:
                postfix = postfix or "未知"
                
            cprint(f"手動模式檔名後綴: {postfix}")
        else:
            # 自動模式：保持原有邏輯
            area = current_area or area or "未知區"
            section = current_section or section or "未知段"
            postfix = postfix or _sanitize(f"{area}{section}")
        
        # 確保輸出目錄存在
        os.makedirs(self.base_dir, exist_ok=True)
        base_name = f"內政部實價登錄-{postfix}"
        png_path = _unique_path(os.path.join(self.base_dir, f"{base_name}.png"))
        pdf_path = png_path.replace(".png", ".pdf")

        # ----------- 截圖方法 -----------
        try:
            # 確保在正確的frame內
            self.switch_to_inner_frame()
            
            # 使用JavaScript獲取整個頁面的高度
            total_height = self.driver.execute_script("return document.body.scrollHeight")
            total_width = self.driver.execute_script("return document.body.scrollWidth")
            
            # 如果頁面很長，需要拼接截圖
            if total_height > 2000:  # 如果頁面高度超過2000像素
                yprint("檢測到長頁面，使用拼接截圖方法...")

                # 💡 不需要改變視窗大小，使用當前尺寸即可
                # self.driver.set_window_size(1024, 1024)
                
                # 創建完整圖片
                stitched_image = Image.new('RGB', (total_width, total_height))
                
                # 計算需要滾動的次數
                viewport_height = self.driver.execute_script("return window.innerHeight")
                scroll_count = (total_height // viewport_height) + 1
                
                for i in range(scroll_count):
                    # 滾動到指定位置
                    scroll_y = i * viewport_height
                    self.driver.execute_script(f"window.scrollTo(0, {scroll_y})")
                    time.sleep(0.5)  # 等待渲染
                    
                    # 截取當前視窗
                    screenshot = self.driver.get_screenshot_as_png()
                    img = Image.open(io.BytesIO(screenshot))
                    
                    # 計算粘貼位置
                    paste_y = scroll_y
                    if i == scroll_count - 1:  # 最後一張
                        # 調整最後一張圖片的位置
                        paste_y = total_height - img.height
                    
                    # 粘貼到完整圖片
                    stitched_image.paste(img, (0, paste_y))
                
                # 保存拼接後的圖片
                stitched_image.save(png_path)
                yprint("拼接截圖完成")
                
            else:
                # 頁面不太長，直接截圖
                # 💡 不需要改變視窗大小，使用當前尺寸即可
                # self.driver.set_window_size(1024, 1024)
                time.sleep(0.5)

                # 使用標準截圖方法
                self.driver.save_screenshot(png_path)
                yprint("標準截圖完成")
                
        except Exception as e:
            yprint(f"截圖時發生錯誤: {e}")
            # 備用方法：最基本的截圖
            try:
                self.driver.save_screenshot(png_path)
            except:
                yprint("所有截圖方法都失敗了")
                return

        # 轉換為PDF
        try:
            Image.open(png_path).convert("RGB").save(pdf_path, "PDF", resolution=110)
            yprint(f"📄 已輸出 {png_path} 和 PDF")
        except Exception as e:
            yprint(f"轉換PDF時發生錯誤: {e}")

    # ---------------- Core ----------------
    def fill_form(self, item: dict) -> bool:
        global should_exit
        if should_exit:
            return False
            
        cprint(f"\n【填寫表單】使用以下數據:")
        cprint(f"  城市: {item.get('city', 'N/A')}")
        cprint(f"  區域: {item.get('area', 'N/A')}")
        cprint(f"  地段: {item.get('section', 'N/A')}")
        
        if not self.switch_to_inner_frame():
            yprint("無法切換到內部框架")
            return False
        try:
            # 選擇城市
            city_element = self.wait.until(EC.element_to_be_clickable((By.ID, "p_city")))
            yprint(f"選擇城市: {item['city']}")
            Select(city_element).select_by_visible_text(item["city"])
            
            # 等待區域選項加載
            yprint("等待區域選項加載...")
            self.wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "#p_town option")) > 1)
            
            # 選擇區域
            yprint(f"選擇區域: {item['area']}")
            Select(self.driver.find_element(By.ID, "p_town")).select_by_visible_text(item["area"])
            self.last_area = item["area"]
            self.last_section = item["section"]  # 更新最後的地段
            
            # 勾選「地段」選項
            yprint("勾選「地段」選項")
            self.driver.execute_script("document.querySelector(\"label[for='customCheck2']\").click();")
            
            # 輸入地段
            yprint(f"輸入地段: {item['section']}")
            build_inp = self.driver.find_element(By.ID, "p_build")
            build_inp.clear()
            build_inp.send_keys(item["section"])
            
            # 點擊搜索按鈕
            yprint("點擊搜索按鈕")
            self.driver.execute_script("document.querySelector(\"button[go_type='list']\").click();")
            
            yprint("✅ 表單填寫完成，等待結果...")
            return True
        except Exception as e:
            yprint(f"❌ 自動填表錯誤: {e}")
            return False

    def wait_for_input(self, prompt_message):
        """等待用戶輸入，但能夠同時檢查退出標誌"""
        global input_queue, should_exit
        
        # 如果已經應該退出，直接返回
        if should_exit:
            return "q"
            
        yprint(prompt_message)
        
        # 清空隊列中的舊數據
        while not input_queue.empty():
            input_queue.get_nowait()
            
        # 等待新的輸入
        while not should_exit:
            try:
                cmd = input_queue.get(timeout=0.1)  # 短超時
                if cmd.lower() == "q":
                    should_exit = True
                    yprint("收到退出命令！")
                    return "q"
                else:
                    # 顯示用戶輸入的內容（如果是空字符串，表示按下了Enter）
                    if cmd == "":
                        gprint("您按下了 Enter")
                    else:
                        gprint(f"您輸入了: {cmd}")
                return cmd.lower()
            except queue.Empty:
                # 超時但沒有輸入，順便檢查瀏覽器是否還活著
                # （比照 nlma.py 的做法：使用者按 X 關掉 Chrome 時要能偵測到並退出）
                if self.driver:
                    try:
                        _ = self.driver.current_url
                    except Exception:
                        yprint("\n偵測到瀏覽器已關閉，程式將退出")
                        should_exit = True
                        return "q"

        return "q"  # 如果應該退出，返回q
            
    def search_one(self, item: dict, auto_capture=False) -> bool:
        global should_exit
        if should_exit:
            return False

        if not self.fill_form(item):
            cmd = self.wait_for_input("❗ 填表失敗，請手動完成後按 Enter 或 Q 結束…")
            if cmd == "q" or should_exit:
                return False
        else:
            if auto_capture:
                # 自動查詢模式：直接截圖不等待
                cprint("✅ 自動填表完成，自動截圖＋存檔")
                time.sleep(1)  # 稍等一下讓結果完全顯示
            else:
                # 手動模式：等待使用者確認
                cmd = self.wait_for_input("✅ 自動填表完成，請檢查結果後按 Enter 截圖＋存檔")
                if cmd == "q" or should_exit:
                    return False
        # 截圖前重新讀取當前的區域和地段
        self.capture_page()
        return True
    
    # 🔥 Chrome 死亡 watchdog（背景執行緒）
    def _start_chrome_watchdog(self):
        """每 1 秒檢查 self.driver 是否還連得上 Chrome，斷了就觸發退出。

        比照 nlma.py 用 driver.current_url 試讀的方式偵測。
        當 Chrome 被使用者按 X 關閉時，這個 thread 會：
          1. 設 should_exit = True 通知所有 while 迴圈退出
          2. 塞個 'q' 進 input_queue 把可能正在等輸入的 wait_for_input 喚醒
        """
        def watchdog():
            global should_exit, input_queue
            while not should_exit:
                time.sleep(1)
                if should_exit:
                    return
                try:
                    _ = self.driver.current_url
                except Exception:
                    yprint("\n[Watchdog] 偵測到 Chrome 視窗已被關閉，程式將退出")
                    should_exit = True
                    try:
                        input_queue.put('q')
                    except Exception:
                        pass
                    return
        threading.Thread(target=watchdog, daemon=True).start()

    # 關閉瀏覽器
    def close_browser(self):
        if self.driver:
            try:
                yprint("正在關閉瀏覽器...")
                self.driver.quit()
                yprint("瀏覽器已關閉")
            except Exception as e:
                yprint(f"關閉瀏覽器時出錯: {e}")


# 需要額外導入io模組
import io

# 監聽標準輸入的線程
def stdin_monitor():
    global input_queue, should_exit
    
    while not should_exit:  # 注意：現在檢查全局退出標誌
        try:
            line = sys.stdin.readline().strip()
            
            # 將所有輸入放入隊列
            input_queue.put(line)
            
            # 檢查是否為退出命令
            if line.lower() == 'q':
                yprint("stdin_monitor: 收到退出命令 'q'")
                should_exit = True  # 設置全局退出標誌，但不退出線程
                
        except Exception:
            # 讀取錯誤，短暫休息後繼續
            time.sleep(0.1)


# 創建一個專門監控退出標志的線程 - 強制關閉機制
def exit_monitor():
    global should_exit
    
    last_check = time.time()
    
    while not should_exit:
        time.sleep(0.1)  # 輕量檢查
        
        # 每隔2秒檢查一次
        current_time = time.time()
        if current_time - last_check >= 2:
            last_check = current_time
            
            # 嘗試讀取stdin中的'q'
            try:
                if not sys.stdin.isatty() and sys.stdin.readable():
                    # 讀取但不阻塞
                    import msvcrt
                    if msvcrt.kbhit():
                        ch = msvcrt.getch()
                        if ch == b'q':
                            yprint("exit_monitor: 收到退出命令 'q'")
                            should_exit = True
                            break
            except:
                pass


# ---------------- Main ----------------
if __name__ == "__main__":
    app = None
    
    # 啟動輸入監控線程
    input_thread = threading.Thread(target=stdin_monitor, daemon=True)
    input_thread.start()
    
    # 啟動退出監控線程 (備用機制)
    exit_thread = threading.Thread(target=exit_monitor, daemon=True)
    exit_thread.start()
    
    try:
        bprint("\n====== 實價登錄查詢程序啟動 ======\n")
        
        if not os.path.exists(get_data_json_path()):
            yprint("❌ 缺少 data.json")
            sys.exit(1)
        
        cprint(f"正在讀取 data.json...")
        raw_items = json.load(open(get_data_json_path(), encoding="utf-8"))
        cprint(f"data.json 包含 {len(raw_items)} 項資料")
        
        # 顯示所有地號數據
        display_lot_numbers(raw_items)
        
        # 去重處理 - 只保留不同的城市、區域、地段組合
        seen = set()
        items = []
        for it in raw_items:
            key = (it.get("city", ""), it.get("area", ""), it.get("section", ""))
            if key not in seen:
                seen.add(key)
                items.append(it)
        
        cprint(f"去重後剩餘 {len(items)} 組搜尋參數\n")
        
        # 獲取第一項數據，用於輸出目錄
        if not items:
            yprint("❌ 沒有有效的數據項")
            sys.exit(1)
            
        first = items[0]
        bprint("====== 設置輸出目錄 ======")
        lot_number = first.get("lot_number", "unknown")
        # 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
        out_dir = os.path.join(get_work_folder(f"{first['area']}{first['section']}-{lot_number}"), "3.行情")
        cprint(f"輸出目錄: {out_dir}\n")
        
        bprint("====== 啟動瀏覽器 ======")
        yprint("正在啟動瀏覽器，這可能需要幾秒鐘...")
        app = InteriorPriceSearch(out_dir)
        app.setup_driver()
        
        yprint("正在載入網頁，請稍候...")
        app.driver.get(app.url)

        # 🔥 網頁載入後重新設定視窗大小（避免被網站重置）
        try:
            app.driver.set_window_size(interior_window_width, interior_window_height)
            app.driver.set_window_position(interior_window_x, interior_window_y)
        except:
            pass

        # 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
        verify_and_fix_chrome_window(app.driver)

        # 等待網頁完全載入（內政部網站較慢，多等一些時間）
        time.sleep(5)

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
                elif current_dpi >= 1.25:
                    # 🔥 小螢幕 125% 需要更多縮放
                    if logical_width < 1200:
                        zoom_count = 2  # 小螢幕 125%：按 2 次
                    else:
                        zoom_count = 1  # 大螢幕 125%：按 1 次
                else:
                    zoom_count = 0  # 100%：不縮放
        except Exception:
            pass

        # 🔥 執行縮放（參考 V523.py 的實作）
        if zoom_count > 0:
            try:
                # 先確保頁面完全載入
                time.sleep(2)

                # 🔥 使用 keyboard 庫發送系統級的 Ctrl + - 按鍵（真正的瀏覽器縮放）
                # 使用 JavaScript 來點擊 body，確保不會因為元素未找到而失敗
                try:
                    app.driver.execute_script("document.body.click();")
                except:
                    pass

                time.sleep(0.5)

                # 🔥 根據 DPI 決定按幾次 Ctrl + -
                for i in range(zoom_count):
                    keyboard.press_and_release('ctrl+-')  # 系統級按鍵
                    time.sleep(0.5)

            except Exception:
                pass  # 靜默處理縮放錯誤

        # 🔥 收到退出訊號就直接跳過後續所有流程（避免 Chrome 被關後還印一堆訊息）
        if not should_exit:
            bprint("\n====== 開始處理數據 ======")
            for idx, itm in enumerate(items, 1):
                if should_exit:
                    yprint("主循環檢測到退出標志，準備退出...")
                    break

                cprint(f"\n[{idx}/{len(items)}] 處理: {itm.get('city', '')} {itm.get('area', '')} {itm.get('section', '')}")
                if not app.search_one(itm, auto_capture=True):
                    break
            else:
                if not should_exit:
                    yprint("⚠ 提醒：若存檔不如預期請手動更改檔名")

        # 手動延伸查詢循環
        if not should_exit:
            bprint("\n====== 進入手動查詢模式 ======")
            cprint("您現在可以手動在網頁上調整查詢條件")
            cprint("系統將自動偵測區域和地段或社區名稱作為檔名")
            cprint("例如：偵測到「鼓山區」和「鼎宇美術館」時，存檔為「內政部實價登錄-鼓山區鼎宇美術館.png」")

            while not should_exit:
                cmd = app.wait_for_input("\n若仍需查詢其它，完成後請按 Enter 截圖存檔，或輸入 Q 結束：")
                if cmd == "q" or should_exit:
                    break
                # 在手動模式下，capture_page會使用區域+地段名稱
                app.capture_page(manual_mode=True)
    
    except Exception as e:
        yprint(f"程式執行錯誤: {e}")
        import traceback
        yprint(traceback.format_exc())
    
    finally:
        # 設置退出標記，確保所有線程能夠退出
        should_exit = True
        
        # 確保關閉瀏覽器
        bprint("\n====== 程序結束 ======")
        if app:
            app.close_browser()
        
        yprint("程式結束")
        print("子程式已完成執行", flush=True)