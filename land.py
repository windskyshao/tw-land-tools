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

# 🔥 立即顯示歡迎訊息（在載入其他模組之前）
os.system('cls')
sys.stdout.reconfigure(line_buffering=True)
print("歡迎使用【地籍便民】自動化小程式，模組載入中...", flush=True)

# 🔥 基準目錄設定（優先載入，因為後面需要用到）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

# 🔥 快速載入的模組
import json
import ctypes
import time
import tempfile
import random

# 🔥 較慢的模組（selenium, PIL, fpdf, keyboard）
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from selenium.common.exceptions import JavascriptException, NoAlertPresentException, TimeoutException, UnexpectedAlertPresentException, NoSuchElementException, WebDriverException, InvalidSessionIdException
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window
from PIL import Image
from fpdf import FPDF
import keyboard

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
base_width = 1068
base_height = 1024
chrome_x = 0
chrome_y = 0

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

            base_width = chrome_width_logical
            base_height = chrome_height_logical
            chrome_x = chrome_x_logical
            chrome_y = chrome_y_logical
except Exception as e:
    pass

# 設置 WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36")
# 🔥 不在啟動參數設定視窗大小，改為啟動後再調整（避免 DPI 改變時崩潰）
# options.add_argument(f"--window-size={base_width},{base_height}")
# options.add_argument("--window-position=0,0")
# 🔥 注意：不使用 --force-device-scale-factor，讓 Chrome 自動處理縮放
# 隱藏 "Chrome 正在受到自動化測試軟體控制" 的訊息
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
# 每次執行時使用唯一的時間戳建立獨立的用戶資料目錄
user_data_dir = os.path.join(tempfile.gettempdir(), f"chrome_land_{int(time.time())}_{random.randint(1000, 9999)}")
options.add_argument(f"--user-data-dir={user_data_dir}")
# 使用 ChromeDriverManager 安裝 ChromeDriver 並應用選項
# driver = create_chrome_driver(options=options)

# 加入布林變數，記錄是否為第一次執行
# 🔥 若主程式帶入 --resume，表示要接續既有 data.json，不覆蓋
is_first_run = True
if '--resume' in sys.argv:
    is_first_run = False
    print("[接續模式] 將保留既有 data.json 內容，新資料會加在後面", flush=True)
# 初始化 driver 為 None，避免在 finally 區塊中出現 NameError
driver = None

def is_browser_open(driver):
    """檢查瀏覽器 session 是否仍然有效"""
    if driver is None:
        return False
    try:
        # 嘗試獲取當前 URL，如果 session 無效會拋出異常
        _ = driver.current_url
        return True
    except InvalidSessionIdException:
        # Session 已失效
        return False
    except WebDriverException as e:
        # 只有在錯誤訊息明確表示連線已斷開時才返回 False
        error_msg = str(e).lower()
        if "invalid session" in error_msg or "disconnected" in error_msg or "session deleted" in error_msg:
            return False
        # 其他 WebDriverException 可能是暫時性的，不應該視為瀏覽器已關閉
        return True
    except Exception:
        # 其他異常可能是暫時性的
        return True

def check_for_alert(driver):
    """在每次 Selenium 操作前檢查並處理可能出現的警告彈窗"""
    # 先檢查瀏覽器是否還開啟
    if not is_browser_open(driver):
        print("瀏覽器已關閉，程式即將退出。", flush=True)
        sys.exit()

    try:
        # 嘗試獲取警告彈窗
        alert = WebDriverWait(driver, 2).until(EC.alert_is_present())
        alert_text = alert.text
        print(f"檢測到警告視窗: {alert_text}", flush=True)

        if "請輸入驗証碼" in alert_text:
            alert.accept()  # 關閉驗證碼提示視窗
            handle_captcha(driver)  # 呼叫驗證碼處理函數
        else:
            alert.accept()  # 對非驗證碼的彈窗進行關閉處理
        return True
    except TimeoutException:
        return False  # 沒有檢測到警告彈窗
    except UnexpectedAlertPresentException as e:
        print(f"處理驗證碼時發生錯誤: {e}", flush=True)
        try:
            driver.switch_to.alert.accept()  # 接受錯誤後的彈出視窗
        except (InvalidSessionIdException, WebDriverException):
            print("瀏覽器已關閉，程式即將退出。", flush=True)
            sys.exit()
        return True
    except (InvalidSessionIdException, WebDriverException):
        print("瀏覽器已關閉，程式即將退出。", flush=True)
        sys.exit()
    except Exception as e:
        # 檢查是否為 session 相關錯誤
        if "invalid session" in str(e).lower() or "disconnected" in str(e).lower():
            print("瀏覽器已關閉，程式即將退出。", flush=True)
            sys.exit()
        print(f"處理警告時發生未知錯誤: {e}", flush=True)
        return True

def save_data_to_json(data, json_file=None, new_file=False):
    """保存數據到 JSON 文件，如果存在重複，則更新舊數據，保持順序"""
    # 🔥 使用 BASE_DIR 中的 data.json
    if json_file is None:
        json_file = get_data_json_path()

    if new_file or not os.path.exists(json_file):
        existing_data = [data]
        print("建立新 data.json 中...", flush=True)
    else:
        with open(json_file, 'r', encoding='utf-8') as f:
            existing_data = json.load(f)
            if isinstance(existing_data, dict):
                existing_data = [existing_data]
            elif not isinstance(existing_data, list):
                existing_data = []

        group_found = False
        for i, entry in enumerate(existing_data):
            if (entry.get('city') == data.get('city') and
                entry.get('area') == data.get('area') and
                entry.get('section') == data.get('section') and
                entry.get('lot_number') == data.get('lot_number')):
                existing_data[i] = data  # 覆蓋原來的數據
                group_found = True
                break

        if not group_found:
            existing_data.append(data)  # 如果沒有找到匹配的組，則添加新數據

        print("接續原始資料中...", flush=True)

    # 儲存到程式目錄
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(existing_data, f, ensure_ascii=False, indent=4)

    # 🔥 立即輸出第一筆資料給主程式更新標題
    if existing_data and len(existing_data) > 0:
        first_data = existing_data[0]
        city = first_data.get('city', '')
        area = first_data.get('area', '')
        section = first_data.get('section', '')
        lot_number = first_data.get('lot_number', '')
        coord_info = first_data.get('coord_info', '')

        # 輸出格式與主程式一致
        if coord_info:
            print(f"JSON: {city} | {area}{section} {lot_number} {coord_info}", flush=True)
        else:
            print(f"JSON: {city} | {area}{section} {lot_number}", flush=True)

    # 🔥 同時備份到建立的資料夾/4.其他相關/
    try:
        if existing_data and len(existing_data) > 0:
            first_data = existing_data[0]
            area = first_data.get('area', '')
            section = first_data.get('section', '')
            lot_number = first_data.get('lot_number', '')

            if area and section and lot_number:
                folder_name = f"{area}{section}-{lot_number}"
                # 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
                other_dir = os.path.join(get_work_folder(folder_name), "4.其他相關")

                # 確保資料夾存在
                if not os.path.exists(other_dir):
                    os.makedirs(other_dir)

                backup_path = os.path.join(other_dir, "data.json")
                with open(backup_path, 'w', encoding='utf-8') as f:
                    json.dump(existing_data, f, ensure_ascii=False, indent=4)
                print(f"[OK] 已備份 data.json 到 {backup_path}", flush=True)
    except Exception as e:
        print(f"[警告] 備份 data.json 失敗：{e}", flush=True)

def get_first_record_directory(json_file=None):
    """獲取 data.json 第一條記錄的目錄路徑"""
    # 🔥 使用 BASE_DIR 中的 data.json
    if json_file is None:
        json_file = get_data_json_path()

    if os.path.exists(json_file):
        with open(json_file, 'r', encoding='utf-8') as f:
            existing_data = json.load(f)
            if isinstance(existing_data, list) and existing_data:
                first_record = existing_data[0]
                area = first_record.get('area', '')
                section = first_record.get('section', '')
                lot_number = first_record.get('lot_number', '')
                # 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
                return get_work_folder(f"{area}{section}-{lot_number}")
    return None

def get_user_choice():
    """自動選擇資料儲存模式：第一次選擇新資料檔案，之後接續原始資料"""
    global is_first_run
    if is_first_run:
        is_first_run = False  # 之後的執行將不再是第一次
        return True  # 選擇新建 Data.json
    return False  # 接續原始資料

def create_directory_structure(base_dir):
    """創建目錄結構，如果目錄不存在則創建"""
    if not os.path.exists(base_dir):
        os.makedirs(base_dir)
    return base_dir

def generate_unique_filename(directory, base_name, extension=".pdf"):
    """直接返回文件名，允許覆蓋已存在的文件"""
    # 🔧 修改：直接返回文件路徑，不再添加序號
    # 如果文件已存在，將會被覆蓋
    return os.path.join(directory, f"{base_name}{extension}")



def handle_captcha(driver):
    """等待並處理驗證碼輸入"""
    try:
        # 等待驗證碼輸入框出現
        captcha_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "recaptcha_id"))
        )
        if captcha_element.is_displayed():
            print("\033[93m偵測到驗證碼：請在 Chrome 視窗的驗證碼框中輸入，打完程式會自動送出（不限時間；關閉瀏覽器即取消）。\033[0m", flush=True)

            # 等待用戶在 Chrome 中輸入驗證碼：不限時間，瀏覽器開著就一直等
            #（每 0.3 秒檢查一次，CPU 負擔極低；要中止就關掉瀏覽器即可）。
            # 判斷「打好了」的條件（避免慢慢打時被搶先送出）：
            #   ① 打滿 4 碼（input maxlength=4）→ 立即送出，或
            #   ② 有輸入但停止打字超過 1.5 秒（可能打了 <4 碼就停）→ 才送出。
            import time as _t
            CAPTCHA_LEN = 4
            SETTLE_SEC = 1.5
            POLL = 0.3
            _last_val = None
            _stable_for = 0.0
            while True:
                if not is_browser_open(driver):
                    print("瀏覽器已關閉，取消驗證碼等待。", flush=True)
                    return
                try:
                    val = (captcha_element.get_attribute("value") or "").strip()
                except Exception:
                    # 驗證碼框可能已隨頁面更新而消失，結束等待
                    return
                if len(val) >= CAPTCHA_LEN:        # 打滿 4 碼 → 立刻送
                    break
                if val:                            # 有值但停止輸入夠久 → 送
                    if val == _last_val:
                        _stable_for += POLL
                        if _stable_for >= SETTLE_SEC:
                            break
                    else:
                        _stable_for = 0.0
                _last_val = val
                _t.sleep(POLL)
            print("已偵測到驗證碼輸入，程式即將自動點擊【確認】按鈕...", flush=True)

            # 等待確認按鈕可點擊後自動點擊
            confirm_button = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='確認']"))
            )
            confirm_button.click()
            print("確認按鈕已自動點擊，繼續執行程序。", flush=True)
        else:
            print("未偵測到驗證碼欄位，繼續執行程序。", flush=True)
    except TimeoutException:
        print("驗證碼處理超時，未檢測到驗證碼或確認按鈕，繼續執行程序。", flush=True)
    except Exception as e:
        print("處理驗證碼時發生錯誤:", e , flush=True)

def handle_alert_immediately(driver):
    """立即處理按鈕被點擊後的警告彈窗"""
    try:
        alert = WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert_text = alert.text
        print(f"警示視窗出現: {alert_text}", flush=True)

        if "查無" in alert_text:
            print("查無資料，請重新輸入查詢條件...", flush=True)
            alert.accept()
            # 重置查詢按鈕狀態
            driver.execute_script("window.isQueryButtonClicked = false;")
            return "no_data"  # 返回特殊狀態表示查無資料
        elif "請輸入驗証碼" in alert_text:
            alert.accept()
            handle_captcha(driver)
        else:
            print("處理其他警告訊息...", flush=True)
            alert.accept()
        return "handled"
    except TimeoutException:
        return "no_alert"
    except UnexpectedAlertPresentException as e:
        print(f"處理警告時發生錯誤: {e}", flush=True)
        try:
            driver.switch_to.alert.accept()
        except:
            pass
        return "error"
    except Exception as e:
        print(f"處理警告時發生錯誤: {e}", flush=True)
        return "error"

def wait_for_button_click(driver):
    """等待查詢按鈕被點擊，或 ESC 被按下退出程式"""
    # 先檢查瀏覽器是否還開啟
    if not is_browser_open(driver):
        print("瀏覽器已關閉，程式即將退出。", flush=True)
        sys.exit()

    script = """
    document.getElementById('land_button').addEventListener('click', function() {
        window.isQueryButtonClicked = true;
    });
    window.isQueryButtonClicked = false;
    """
    driver.execute_script(script)

    while True:
        # 每次迴圈中先檢查瀏覽器是否還開啟
        if not is_browser_open(driver):
            print("瀏覽器已關閉，程式即將退出。", flush=True)
            sys.exit()

        # 檢查是否有彈窗
        check_for_alert(driver)  # 檢查並處理彈窗

        try:
            # 使用 JavaScript 檢查是否按鈕已被點擊
            is_clicked = driver.execute_script("return window.isQueryButtonClicked;")
            if is_clicked:
                print("查詢按鈕已被點擊!自動化即將開始", flush=True)

                print("檢查是否有彈窗，稍待一下...")
                check_for_alert(driver)
                break  # 結束等待，繼續程序
        except (InvalidSessionIdException, WebDriverException):
            print("瀏覽器已關閉，程式即將退出。", flush=True)
            sys.exit()
        except JavascriptException as e:
            if "invalid session" in str(e).lower() or "disconnected" in str(e).lower():
                print("瀏覽器已關閉，程式即將退出。", flush=True)
                sys.exit()
            print(f"檢查按鈕狀態時出錯: {e}", flush=True)
        except UnexpectedAlertPresentException as e:
            print(f"處理驗證碼時發生錯誤: {e}", flush=True)
            try:
                driver.switch_to.alert.accept()  # 接受錯誤後的彈出視窗
            except (InvalidSessionIdException, WebDriverException):
                print("瀏覽器已關閉，程式即將退出。", flush=True)
                sys.exit()

        check_exit()  # 檢查是否按下 ESC 鍵退出程式

def check_exit():
    """檢查是否按下 ESC 鍵"""
    if keyboard.is_pressed('esc'):
        print("檢測到 ESC 鍵按下，退出程式。", flush=True)
        sys.exit()

def handle_special_elements(driver):
    """處理查詢按鈕點擊後的特殊元素，例如進階訊息圖示、定位圖示等"""
    actions = ActionChains(driver)

    try:
        # # 滑鼠移動到進階訊息圖示並模擬懸停
        # info_icon = WebDriverWait(driver, 5).until(
        #     EC.element_to_be_clickable((By.ID, "cada_span_id"))
        # )
        # actions.move_to_element(info_icon).perform()
        # print("滑鼠已移到【進階訊息圖示】上，等待1秒...", flush=True)
        # time.sleep(1)  # 停留1秒
        # actions.move_by_offset(250, 250).perform()  # 滑鼠移開
        # # time.sleep(2)
        # print("滑鼠已從【進階訊息圖示】移開，提示訊息應該消失。", flush=True)

        # 關閉定位信息
        close_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "info_close"))
        )
        if close_button.is_displayed():
            print("關閉定位信息", flush=True)
            driver.execute_script("arguments[0].scrollIntoView(true);", close_button)
            driver.execute_script("arguments[0].click();", close_button)
            print("定位信息已關閉", flush=True)
        else:
            print("無定位信息需要關閉", flush=True)
    
    except TimeoutException:
        print("未能找到需要操作的特殊元素，略過此部分", flush=True)
    except Exception as e:
        print(f"處理特殊元素時發生錯誤: {e}", flush=True)

def wait_for_enter():
    """等待使用者按下 Enter 鍵，10秒後自動繼續"""
    print("\033[93m請調整地圖位置和顯示範圍，請於10秒內調整...\033[0m", flush=True)
    print("\033[93m(10秒後將自動繼續)\033[0m", flush=True)
    
    start_time = time.time()
    timeout = 10  # 10秒超時

    while True:
        # 檢查是否超時
        if time.time() - start_time > timeout:
            print("等待超時，自動繼續執行程序...", flush=True)
            break

        if keyboard.is_pressed('enter'):
            print("檢測到 Enter 鍵，繼續執行程序...", flush=True)
            time.sleep(0.5)  # 稍微延遲以確保按鍵被完全釋放
            break
        elif keyboard.is_pressed('esc'):
            print("檢測到 ESC 鍵，退出程序...", flush=True)
            sys.exit()
        
        # 顯示剩餘時間
        remaining_time = int(timeout - (time.time() - start_time))
        if remaining_time > 0:
            print(f"剩餘等待時間：{remaining_time}秒", flush=True)
            
        time.sleep(1)  # 每秒更新一次
    
    print("繼續執行後續操作...", flush=True)

def get_land_info():
    driver.get('https://easymap.land.moi.gov.tw/Index')
    # 🔥 網頁載入後重新設定視窗大小（避免被網站重置）
    try:
        driver.set_window_size(base_width, base_height)
        driver.set_window_position(chrome_x, chrome_y)
    except:
        pass

    # 🔥 DPI 自動修正：用共用 helper 驗證 + 修正視窗大小與頁面縮放
    # 只在偵測到 DPI 不同步時才動作，正常 PC 不受影響
    verify_and_fix_chrome_window(driver)

    os.system('cls')

    # 🔥 自動判斷是否需要縮放頁面（解決高 DPI 和小螢幕的排版問題）
    # 需要同時考慮 DPI 和螢幕解析度
    zoom_count = 0
    target_zoom = "100%"
    try:
        config_path = os.path.join(BASE_DIR, 'window_config.json')
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)

            current_dpi = config.get('dpi_scale', 1.0)
            screen_width = config.get('screen_width', 1920)  # 物理像素寬度

            # 🔥 根據 DPI 和螢幕解析度決定縮放
            # 計算邏輯像素寬度（實際可用空間）
            logical_width = screen_width / current_dpi

            # 🔥 縮放規則：
            # 1920x1080 @ 175% (logical=1097) → 67%
            # 1920x1080 @ 150% (logical=1280) → 80%
            # 1920x1080 @ 125% (logical=1536) → 100%
            # 1920x1080 @ 100% (logical=1920) → 100%
            # 1366x768 @ 125% (logical=1093) → 80%
            # 1366x768 @ 100% (logical=1366) → 100%

            if current_dpi >= 1.75:
                zoom_count = 4  # 175% 或更高：縮放到 67%
                target_zoom = "67%"
            elif current_dpi >= 1.5:
                zoom_count = 2  # 150%：縮放到 80%
                target_zoom = "80%"
            elif current_dpi >= 1.25 and logical_width < 1200:
                # 🔥 新增：小螢幕 + 125% DPI 時也需要縮放
                # 例如 1366x768@125% (logical=1093) 需要縮放到 80%
                zoom_count = 2  # 縮放到 80%
                target_zoom = "80%"
            else:
                zoom_count = 0  # 不縮放
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
            # 使用 JavaScript 讓視窗獲得焦點，避免點擊觸發頁面元素
            time.sleep(1)  # 等待頁面完全載入
            driver.execute_script("window.focus();")
            time.sleep(0.3)

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

    # 等待下拉選單出現
    select_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "mapTile_id"))
    )

    # 創建 Select 物件
    select = Select(select_element)

    # 選擇第一個選項 "NLSC 地圖"
    select.select_by_visible_text("NLSC 地圖")

    # 確認選擇
    selected_option = select.first_selected_option
    print(f"已選擇的地圖類型: {selected_option.text}", flush=True)

    # 🔥 等待額外圖層圖示出現（使用穩定的 class 和 title 組合，不依賴動態 ID）
    feature_icon = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.olControlButton2ItemActive[title='圖層']"))
    )

    # 創建 ActionChains 對象
    actions = ActionChains(driver)

    print("滑鼠移動到功能圖示，然後暫停一段時間", flush=True)
    actions.move_to_element(feature_icon).pause(0.5).perform()  # 暫停 0.5秒

    # 確保滑鼠在圖示上待 0.5秒
    time.sleep(0.5)  # 確保滑鼠保持在圖示上

    print("移開滑鼠，關閉額外圖層選單", flush=True)
    # 🔥 改用移動到 body 元素的中心，避免在小視窗/高DPI時超出邊界
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        actions.move_to_element(body).pause(1).perform()
    except Exception:
        # fallback: 嘗試較小的偏移量
        try:
            actions.move_by_offset(20, 20).pause(1).perform()
        except Exception:
            pass  # 忽略移動失敗，繼續執行

    print("本程式將自動轉存各筆資料至JSON、PNG、PDF", flush=True)
    print("長按【ESC】鍵，可退出程式", flush=True)
    print("\033[93m請選擇【縣市】【地區】【地段】及輸入【地號】後，點擊【查詢】\033[0m", flush=True)

    # 定義 JavaScript 代碼來監聽按鈕點擊事件
    script = """
    document.getElementById('land_button').addEventListener('click', function() {
        window.isQueryButtonClicked = true;
    });
    window.isQueryButtonClicked = false;
    """
    driver.execute_script(script)

    has_printed_timeout_message = False
    
    while True:
        try:
            # 等待並處理查詢按鈕點擊
            wait_for_button_click(driver)
            
            # 檢查並處理警告
            alert_result = handle_alert_immediately(driver)
            
            # 如果查無資料，重新開始等待新的查詢
            if alert_result == "no_data":
                print("請重新輸入查詢條件並點擊查詢按鈕...", flush=True)
                continue

            # 如果正常處理了警告或沒有警告，繼續執行後續操作
            if alert_result in ["handled", "no_alert"]:
                # 關閉信息窗口
                time.sleep(1)
                close_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "info_close"))
                )
                driver.execute_script("arguments[0].click();", close_button)
                print("信息視窗已關閉", flush=True)

                print("獲取選擇的縣市、地區、地段和地號")
                city_element = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.ID, 'select_city_id'))
                )
                city_select = Select(city_element)
                city_text = city_select.first_selected_option.text

                area_element = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.ID, 'select_town_id'))
                )
                area_select = Select(area_element)
                area_text = area_select.first_selected_option.text

                section_element = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.ID, 'select_sect_id'))
                )
                section_select = Select(section_element)
                section_text = section_select.first_selected_option.text.split(')')[-1].strip()

                lot_number_input = driver.find_element(By.ID, 'landno').get_attribute('value')
                
                time.sleep(2)
                # 等待地圖上的標記並模擬右鍵點擊
                map_marker = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'image[id^="OpenLayers\\.Geometry\\.Point_"]'))
                )
                
                actions = ActionChains(driver)
                time.sleep(1)
                actions.context_click(map_marker).perform()
                
                # 等待右鍵菜單出現並選擇「取得此位置坐標」
                coord_menu_item = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '//li[contains(@class, "icon-getCoordWGS84byMap")]//span'))
                )
                coord_menu_item.click()

                # 等待座標彈窗出現並獲取座標數據
                coord_text = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, 'coordDisplayLonLat'))
                ).text
                print(f"座標：{coord_text}", flush=True)

                # 使用 JavaScript 固定座標視窗的位置
                print("嘗試固定【座標視窗】的位置...", flush=True)
                try:
                    driver.execute_script("""
                        var coordWindow = document.querySelector('div[aria-describedby="leftClickCoordDisplayId"]');
                        if (coordWindow) {
                            coordWindow.style.position = 'fixed';
                            coordWindow.style.top = '130px';
                            coordWindow.style.left = '5px';
                            coordWindow.style.zIndex = '9999';
                        }
                    """)
                    print("【座標視窗】位置已成功固定。", flush=True)
                except Exception as e:
                    print(f"嘗試固定【座標視窗】位置時發生錯誤: {e}", flush=True)

                # 使用 JavaScript 固定查詢結果視窗的位置
                print("嘗試固定【查詢結果視窗】的位置...", flush=True)
                try:
                    driver.execute_script("""
                        var resultWindow = document.querySelector('div[aria-describedby="dlg_search_result"]');
                        if (resultWindow) {
                            resultWindow.style.position = 'fixed';
                            resultWindow.style.top = '300px';
                            resultWindow.style.left = '10px';
                            resultWindow.style.zIndex = '9999';
                        }
                    """)
                    print("【查詢結果視窗】位置已成功固定。", flush=True)
                except Exception as e:
                    print(f"嘗試固定【查詢結果視窗】位置時發生錯誤: {e}", flush=True)

                # 定位圖示處理
                print("開始查找【定位圖示】...", flush=True)
                try:
                    map_marker = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'image[id^="OpenLayers\\.Geometry\\.Point_"]'))
                    )

                    # 🔥 讀取縮放比例，調整錨點移動的偏移量
                    zoom_scale = 1.0  # 預設無縮放
                    try:
                        config_path = os.path.join(BASE_DIR, 'window_config.json')
                        if os.path.exists(config_path):
                            with open(config_path, 'r', encoding='utf-8') as f:
                                config = json.load(f)
                            current_dpi = config.get('dpi_scale', 1.0)
                            screen_width = config.get('screen_width', 1920)
                            logical_width = screen_width / current_dpi

                            # 🔥 根據 DPI 和螢幕解析度取得實際的縮放比例
                            if current_dpi >= 1.75:
                                zoom_scale = 0.67  # 175%: 縮放到 67%
                            elif current_dpi >= 1.5:
                                zoom_scale = 0.80  # 150%: 縮放到 80%
                            elif current_dpi >= 1.25 and logical_width < 1200:
                                zoom_scale = 0.80  # 小螢幕 + 125%: 縮放到 80%
                    except:
                        pass

                    location = map_marker.location
                    size = map_marker.size

                    # 🔥 使用 JavaScript 獲取實際的 viewport 尺寸（網頁可視區域，不含瀏覽器標題列等）
                    viewport_width = driver.execute_script("return window.innerWidth;")
                    viewport_height = driver.execute_script("return window.innerHeight;")

                    # 🔥 嘗試直接測量左側資訊欄的實際寬度
                    try:
                        # 查詢結果區域通常有固定的 class 或 id
                        info_panel = driver.find_element(By.CSS_SELECTOR, '.queryResult, #queryResult, [class*="query"]')
                        info_panel_width = info_panel.size['width']
                    except:
                        # 如果找不到資訊欄元素，使用固定值並根據縮放調整
                        if zoom_scale <= 0.67:
                            info_panel_width = 433  # 175% 模式
                        elif zoom_scale <= 0.80:
                            info_panel_width = 362  # 150% 模式
                        else:
                            info_panel_width = 290  # 100%, 125% 模式

                    # 🔥 計算右側剩餘空間的中心點
                    remaining_width = viewport_width - info_panel_width
                    center_x = info_panel_width + (remaining_width / 2)
                    center_y = viewport_height / 2

                    start_x = location['x'] + (size['width'] / 2)
                    start_y = location['y'] + (size['height'] / 2)
                    end_x = center_x
                    end_y = center_y

                    # 🔥 微調偏移量（因為已經計算右側空間中心，這裡只需小幅調整）
                    offset_x_correction = 0  # 不需要額外偏移
                    offset_y_correction = 0  # 不需要額外偏移

                    actions = ActionChains(driver)
                    actions.move_to_element(map_marker)
                    actions.move_by_offset(0, -10)
                    actions.click_and_hold()
                    actions.move_to_element_with_offset(map_marker, end_x - (start_x - offset_x_correction), end_y - (start_y + offset_y_correction))
                    actions.release().perform()

                    print(f"【定位圖示】已成功拖曳到視窗中心（縮放比例: {zoom_scale*100:.0f}%）。", flush=True)
                    print("\033[93m現在您可以手動調整地圖位置和顯示範圍...\033[0m", flush=True)
                    
                except Exception as e:
                    print(f"發生錯誤，未能找到【定位圖示】: {e}", flush=True)

                # 進階訊息圖示處理
                try:
                    print("再次查找【進階訊息圖示】並準備模擬滑鼠懸停...", flush=True)
                    info_icon = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.ID, "cada_span_id"))
                    )
                    actions.move_to_element(info_icon).perform()
                    print("滑鼠已移到【進階訊息圖示】上，等待1秒...", flush=True)
                    time.sleep(0.5)
                    # 移動到視窗左上角（安全位置）
                    actions.move_to_element_with_offset(driver.find_element(By.TAG_NAME, "body"), 10, 10).perform()
                    print("滑鼠已從【進階訊息圖示】移開，提示訊息應該消失。", flush=True)
                except Exception as e:
                    print(f"發生錯誤，未能找到【進階訊息圖示】: {e}", flush=True)

                # 保存數據
                data = {
                    'city': city_text,
                    'area': area_text,
                    'section': section_text,
                    'lot_number': lot_number_input,
                    'coordinates': coord_text
                }

                # 是否創建新文件
                new_file = get_user_choice()
                save_data_to_json(data, new_file=new_file)
                print("數據已保存到 data.json 文件。", flush=True)

                # 文件與目錄結構
                if new_file:
                    # 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
                    base_dir = get_work_folder(f"{area_text}{section_text}-{lot_number_input}")
                else:
                    base_dir = get_first_record_directory()

                pdf_dir = create_directory_structure(os.path.join(base_dir, "1.基本資料"))
                png_dir = create_directory_structure(os.path.join(base_dir, "1.基本資料", "png"))

                base_filename = f"05_地籍便民-{area_text}-{section_text}-{lot_number_input}"

                # 等待使用者調整完成
                print("\033[93m請調整地圖位置和顯示範圍，完成後請按 Enter 鍵繼續...\033[0m", flush=True)
                print("\033[93m(10秒後將自動繼續)\033[0m", flush=True)
                countdown(10)

                # 保存截圖
                screenshot_path = generate_unique_filename(png_dir, base_filename, ".png")
                driver.save_screenshot(screenshot_path)
                print(f"\033[93m網頁截圖已保存為 {screenshot_path}\033[0m", flush=True)

                # 截圖轉 PDF
                image = Image.open(screenshot_path)
                pdf_path = generate_unique_filename(pdf_dir, base_filename, ".pdf")
                pdf = FPDF(orientation='L', unit='mm', format='A4')
                pdf.add_page()
                pdf.image(screenshot_path, x=10, y=10, w=277)
                pdf.output(pdf_path)
                print(f"\033[93m截圖已轉換為 PDF 文件 {pdf_path}\033[0m", flush=True)

                # 關閉座標視窗
                try:
                    print("嘗試關閉座標視窗...", flush=True)
                    confirm_button = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.ID, "btnLeftClickCoordDisplayId"))
                    )
                    confirm_button.click()
                    print("座標視窗已成功關閉。", flush=True)
                except TimeoutException:
                    print("未能找到座標視窗的確定按鈕。", flush=True)
                except Exception as e:
                    print(f"關閉座標視窗時發生錯誤: {e}", flush=True)

                has_printed_timeout_message = False

                # 🔥 提示使用者本筆已完成，可繼續下一筆
                print("\n" + "="*39, flush=True)
                print("\033[92m✓ 本筆資料已完成處理！\033[0m", flush=True)
                print("\033[93m請繼續選擇【縣市】【地區】【地段】及輸入【地號】後，點擊【查詢】\033[0m", flush=True)
                print("\033[93m如要退出程式，請長按【ESC】鍵\033[0m", flush=True)
                print("="*39 + "\n", flush=True)

            # 重置查詢按鈕狀態，準備下一次查詢
            driver.execute_script("window.isQueryButtonClicked = false;")

        except (InvalidSessionIdException, WebDriverException) as e:
            # 瀏覽器已關閉
            print("瀏覽器已關閉，程式即將退出。", flush=True)
            break

        except TimeoutException:
            if not has_printed_timeout_message:
                print("請輸入縣市、行政區、地段、地號，\n如果要退出程式，請長按【ESC】鍵。", flush=True)
                has_printed_timeout_message = True
            check_exit()

        except UnexpectedAlertPresentException:
            alert_result = handle_alert_immediately(driver)
            if alert_result == "no_data":
                continue

        except Exception as e:
            # 檢查是否為 session 相關錯誤
            if "invalid session" in str(e).lower() or "disconnected" in str(e).lower():
                print("瀏覽器已關閉，程式即將退出。", flush=True)
                break
            print(f"發生錯誤: {e}", flush=True)
            if "no such alert" not in str(e):
                break

        # 檢查是否需要退出程序
        check_exit()
# handle_special_elements(driver)

def countdown(duration):
    for i in range(duration, 0, -1):
        # 使用 print 將倒數數字輸出到 stdout
        print(f"倒數：{i} 秒", flush=True)  # 將數字轉換為字串
        time.sleep(1)
    print("倒數結束", flush=True)

def notify_main_program():
    """通知主程式子程式執行完畢"""
    print("地籍便民系統已完成執行", flush=True)

if __name__ == "__main__":
    try:
        # 初始化 WebDriver
        driver = create_chrome_driver(options=options)

        # 🔥 啟動後立即設定視窗大小和位置（避免在啟動參數設定導致 DPI 改變時崩潰）
        try:
            driver.set_window_size(base_width, base_height)
            driver.set_window_position(chrome_x, chrome_y)
        except Exception as e:
            print(f"[WARNING] 視窗設定失敗: {e}，繼續執行", flush=True)

        # 子程式主邏輯
        # print("歡迎使用【地籍便民】自動化小程式，模組載入中...", flush=True)
        get_land_info()  # 呼叫主要的地籍查詢函式

    except Exception as e:
        print(f"執行過程中發生錯誤: {e}", flush=True)
        import traceback
        print(f"完整堆疊:\n{traceback.format_exc()}", flush=True)
    finally:
        # 通知主程式
        notify_main_program()
        # 關閉 WebDriver
        if driver:
            driver.quit()
        # 清理臨時的 user-data-dir
        try:
            import shutil
            if os.path.exists(user_data_dir):
                shutil.rmtree(user_data_dir, ignore_errors=True)
        except:
            pass  # 忽略清理錯誤

# Version Update: Modified Land V2_06