#重要設施查詢
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


import signal
import sys
import threading
import select
import os
import json
import time
import shutil
import ctypes
import keyboard  # 用於系統級鍵盤事件
from ctypes import wintypes
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window

# 🔥 基準目錄設定（data.json 和工作資料夾的位置）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

# 清除終端機並顯示歡迎訊息
os.system('cls')
print("歡迎使用【重要設施查詢系統】自動化小程式，模組載入中...", flush=True)

# 確保標準輸出即時更新
sys.stdout.reconfigure(line_buffering=True)

# 全局變數控制程序執行
should_exit = False
driver = None

def get_available_work_area():
    SPI_GETWORKAREA = 0x0030
    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", wintypes.LONG),
            ("top", wintypes.LONG),
            ("right", wintypes.LONG),
            ("bottom", wintypes.LONG)
        ]
    rect = RECT()
    ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
    return rect

def calculate_browser_window_position():
    """計算瀏覽器視窗的位置和大小，避開主程式介面和工作列"""
    # 獲取工作區域（已排除工作列）
    work_area = get_available_work_area()
    screen_width = ctypes.windll.user32.GetSystemMetrics(0)  # 完整螢幕寬度
    
    # 主程式佔用螢幕右側 1/3
    main_program_width = screen_width // 3
    
    # 瀏覽器使用剩餘的左側空間
    browser_width = work_area.right - work_area.left - main_program_width
    browser_height = work_area.bottom - work_area.top
    browser_x = work_area.left
    browser_y = work_area.top
    
    return browser_x, browser_y, browser_width, browser_height

def signal_handler(signum, frame):
    """處理終止信號"""
    global should_exit, driver
    print("接收到終止信號，正在清理資源...", flush=True)
    should_exit = True
    
    if driver:
        try:
            driver.quit()
            # print("瀏覽器已關閉", flush=True)
        except Exception as e:
            print(f"關閉瀏覽器時發生錯誤: {e}", flush=True)

    # print("重要設施已執行結束", flush=True)
    sys.exit(0)

def check_for_exit_input():
    """在背景執行緒中檢查退出輸入"""
    global should_exit
    try:
        while not should_exit:
            if sys.stdin in select.select([sys.stdin], [], [], 0.1)[0]:
                line = sys.stdin.readline().strip()
                if line.lower() in ['q', 'quit', 'exit']:
                    print("接收到退出指令", flush=True)
                    should_exit = True
                    break
            time.sleep(0.1)
    except Exception:
        pass

def display_and_select_data():
    """
    顯示data.json中的所有資料，並讓使用者選擇要執行的項目
    如果只有一組資料，則直接返回該資料
    返回：選中的資料列表
    """
    with open(get_data_json_path(), 'r', encoding='utf-8') as file:
        data_list = json.load(file)
    
    if not data_list:
        print("data.json中沒有資料。", flush=True)
        return []
    
    # 如果只有一組資料，直接返回不需詢問
    if len(data_list) == 1:
        print("data.json中只有一組資料，將直接執行：", flush=True)
        data = data_list[0]
        city = data.get('city', 'N/A')
        area = data.get('area', 'N/A')
        section = data.get('section', 'N/A')
        lot_number = data.get('lot_number', 'N/A')
        data_info = f"{city}{area} {section} {lot_number}"
        if 'coordinates' in data:
            data_info += f"地號， 座標: {data['coordinates']}"
        print(data_info, flush=True)
        return data_list
    
    print("\n===== 重要設施查詢 - 資料選擇 =====", flush=True)
    print(f"在data.json中共找到 {len(data_list)} 筆資料：", flush=True)
    
    # 顯示所有資料
    for idx, data in enumerate(data_list, 1):
        city = data.get('city', 'N/A')
        area = data.get('area', 'N/A')
        section = data.get('section', 'N/A')
        lot_number = data.get('lot_number', 'N/A')
        address = data.get('address', 'N/A')
        
        # 格式化顯示每一筆資料
        data_info = f"[{idx}] {city}{area} {section} {lot_number}"
        if 'coordinates' in data:
            data_info += f" 地號， 座標: {data['coordinates']}"
        print(data_info, flush=True)
    
    # 用戶選擇
    print("\n請選擇要執行的項目編號（多個項目請用逗號分隔，輸入'all'執行全部，輸入'q'退出）：", flush=True)
    user_input = input().strip()
    
    # 檢查是否為退出指令
    if user_input.lower() in ['q', 'quit', 'exit']:
        print("使用者選擇退出程序", flush=True)
        return []
    
    # 若用戶選擇全部執行
    if user_input.lower() == 'all':
        print("將執行所有項目", flush=True)
        return data_list
    
    # 否則，解析用戶選擇的項目
    selected_indices = []
    try:
        # 解析用戶輸入，支援逗號分隔和範圍選擇(如1-3)
        for part in user_input.split(','):
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                selected_indices.extend(range(start, end + 1))
            elif part:
                selected_indices.append(int(part))
    except ValueError:
        print("輸入格式錯誤，程序將退出。請重新執行並輸入正確的數字或數字範圍。", flush=True)
        return []  # 改為返回空列表而不是執行所有項目
    
    # 驗證選擇的索引是否有效
    selected_indices = [idx for idx in selected_indices if 1 <= idx <= len(data_list)]
    if not selected_indices:
        print("未選擇有效項目，程序將退出。", flush=True)
        return []
    
    # 返回選中的資料
    selected_data = [data_list[idx-1] for idx in selected_indices]
    print(f"將執行 {len(selected_data)} 個選定項目", flush=True)
    return selected_data

def main():
    global should_exit, driver

    # 註冊信號處理器
    signal.signal(signal.SIGTERM, signal_handler)
    signal.signal(signal.SIGINT, signal_handler)

    # 啟動輸入檢查執行緒（僅在Windows以外的系統使用）
    if os.name != 'nt':  # 非Windows系統
        input_thread = threading.Thread(target=check_for_exit_input, daemon=True)
        input_thread.start()

    # 初始化變數（確保 finally 區塊可以安全引用）
    download_directory = None

    try:
        # 設定 WebDriver
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        # 計算瀏覽器視窗位置和大小
        browser_x, browser_y, browser_width, browser_height = calculate_browser_window_position()

        # 設定視窗大小和位置
        options.add_argument(f"--window-size={browser_width},{browser_height}")
        options.add_argument(f"--window-position={browser_x},{browser_y}")

        print(f"瀏覽器視窗設定：位置({browser_x}, {browser_y})，大小({browser_width}x{browser_height})", flush=True)

        # 設定預設下載路徑
        # 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
        download_directory = get_work_folder("downloads")
        options.add_experimental_option("prefs", {
            "download.default_directory": download_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        # 確保下載目錄存在
        os.makedirs(download_directory, exist_ok=True)

        # 獲取完整的資料列表，用於獲取第一筆資料
        with open(get_data_json_path(), 'r', encoding='utf-8') as file:
            all_data_list = json.load(file)
        
        if not all_data_list:
            print("data.json中沒有資料，程式退出。", flush=True)
            return
        
        # 顯示資料並讓使用者選擇（如果只有一組資料，則自動選擇該資料）
        selected_data_list = display_and_select_data()
        
        if not selected_data_list:
            print("未選擇任何資料或使用者選擇退出，程式結束。", flush=True)
            return

        # 始終使用資料列表中的第一筆資料建立主目錄（無論使用者選擇的是哪些項目）
        first_data = all_data_list[0]
        first_area = first_data['area']
        first_section = first_data['section'] 
        first_lot_number = first_data['lot_number']
        
        # 建立存儲目錄路徑（使用第一筆資料）
        # 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
        save_directory = os.path.join(get_work_folder(f"{first_area}{first_section}-{first_lot_number}"), "1.基本資料")
        os.makedirs(save_directory, exist_ok=True)

        # 取得PDF檔的前綴名稱
        pdf_prefix = "AnalysisReport"

        # 初始化瀏覽器
        driver = create_chrome_driver(options=options)

        # 開啟目標網頁
        url = "https://service.map.com.tw/houseol/AnalysisObject.aspx"
        driver.get(url)

        # 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
        verify_and_fix_chrome_window(driver)

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

        # 執行查詢操作
        for index, data in enumerate(selected_data_list, 1):
            if should_exit:
                print("檢測到退出信號，停止執行", flush=True)
                break

            # 取得每組資料的座標資訊
            coordinates = data['coordinates']
            area = data['area']
            section = data['section']
            lot_number = data['lot_number']
            
            # 顯示進度
            if len(selected_data_list) > 1:
                print(f"\n處理項目 {index}/{len(selected_data_list)}: {area}{section} {lot_number}", flush=True)
            else:
                print(f"\n處理項目: {area}{section} {lot_number}", flush=True)

            # 檢查退出信號
            if should_exit:
                print("檢測到退出信號，停止執行", flush=True)
                break

            # 設定當前 PDF 檔名（使用當前項目資訊命名，但存放在共同目錄）
            pdf_file_name = f"03_重要設施-{area}{section}-{lot_number}.pdf"
            pdf_file_path = os.path.join(save_directory, pdf_file_name)

            # 在每個 WebDriverWait 操作前檢查
            if should_exit:
                break

            # 點擊「地號座標定位」並填入座標、定位
            locate_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//label[@for='CadastralLocate']"))
            )
            driver.execute_script("arguments[0].click();", locate_button)
            print("已點擊地號座標定位", flush=True)

            locate_input = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "LocateStr"))
            )
            locate_input.clear()
            locate_input.send_keys(coordinates)
            print(f"已填入座標：{coordinates}", flush=True)

            locate_confirm_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "Locate"))
            )
            driver.execute_script("arguments[0].click();", locate_confirm_button)
            print("已點擊定位按鈕", flush=True)
            time.sleep(3)

            # 在每個 WebDriverWait 操作前檢查
            if should_exit:
                break

            # 點擊「查詢重要設施」按鈕 - 使用重複點擊機制提升成功率
            search_facility_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "Search"))
            )

            # 第一次點擊
            driver.execute_script("arguments[0].click();", search_facility_button)
            print("已點擊查詢重要設施按鈕（第1次）", flush=True)
            time.sleep(3)  # 等待3秒讓頁面開始載入

            # 檢查退出信號
            if should_exit:
                break

            # 第二次點擊以確保載入完整
            try:
                # 重新定位按鈕（頁面可能已刷新）
                search_facility_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "Search"))
                )
                driver.execute_script("arguments[0].click();", search_facility_button)
                print("已點擊查詢重要設施按鈕（第2次）", flush=True)
                time.sleep(5)  # 較長的等待時間確保完整載入
            except TimeoutException:
                print("第二次點擊查詢按鈕時超時，可能頁面已載入完成", flush=True)
            except Exception as e:
                print(f"第二次點擊查詢按鈕時發生錯誤: {e}", flush=True)

            print("查詢重要設施載入完成，準備產生報表", flush=True)

            # 切換到包含 ReportButton 的 iframe
            # 🔥 優化：直接使用 iframe id（網站已改版，使用 id 最快最準確）
            iframe = None
            try:
                # 方法1：使用 id='DownloadFileFrame'（最準確，約2-3秒）
                iframe = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.ID, "DownloadFileFrame"))
                )
                # print("已找到 iframe (使用 id='DownloadFileFrame')", flush=True)
            except TimeoutException:
                print("方法1失敗，嘗試使用 src 屬性...", flush=True)
                # 方法2：使用包含 'Download' 的 src（向後相容）
                try:
                    iframe = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'Download')]"))
                    )
                    # print("已找到 iframe (使用 src 包含 Download)", flush=True)
                except TimeoutException:
                    print("方法2失敗，嘗試列舉所有 iframe...", flush=True)
                    # 方法3：列舉所有 iframe 並手動選擇
                    all_iframes = driver.find_elements(By.TAG_NAME, "iframe")
                    for frame in all_iframes:
                        frame_src = frame.get_attribute('src') or ''
                        frame_id = frame.get_attribute('id') or ''
                        if 'Download' in frame_src or 'Download' in frame_id:
                            iframe = frame
                            # print(f"已找到 iframe (列舉方式) - id='{frame_id}'", flush=True)
                            break

            if iframe:
                driver.switch_to.frame(iframe)
                # print("已切換到 iframe", flush=True)
            else:
                raise Exception("無法找到報表 iframe，請檢查網站是否改版")

            # 確認「選擇報表」的容器出現
            report_button_container = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ReportButton"))
            )
            driver.execute_script("arguments[0].scrollIntoView();", report_button_container)
            # print("已確認報表容器出現並滾動到可見範圍", flush=True)
            time.sleep(2)

            # 點擊「選擇報表」的下拉箭頭
            report_dropdown_arrow = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@id='ReportButton']//span[@class='dd-pointer dd-pointer-down']"))
            )
            driver.execute_script("arguments[0].click();", report_dropdown_arrow)
            print("已點擊選擇報表的下拉選單", flush=True)
            # 在每個 WebDriverWait 操作前檢查
            if should_exit:
                break
            # 選擇「PDF」選項
            pdf_option = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@id='ReportButton']//ul[@class='dd-options dd-click-off-close']//label[text()='PDF']"))
            )
            driver.execute_script("arguments[0].click();", pdf_option)
            print("已選擇PDF報表，處理PDF流程，請稍待...", flush=True)

            # 等待PDF檔案下載完成
            # print(f"[調試] 等待PDF下載，下載目錄：{download_directory}", flush=True)
            time.sleep(6)  # 🔥 延長等待時間到 10 秒

            # 搜尋下載的PDF檔案並重新命名
            # print(f"[調試] 開始搜尋下載檔案，檔名前綴：{pdf_prefix}", flush=True)
            downloaded_files = os.listdir(download_directory)
            # print(f"[調試] 下載目錄中的檔案：{downloaded_files}", flush=True)

            found = False
            for file in downloaded_files:
                # print(f"[調試] 檢查檔案：{file}", flush=True)
                if file.startswith(pdf_prefix) and file.endswith(".pdf"):
                    downloaded_pdf_path = os.path.join(download_directory, file)
                    shutil.move(downloaded_pdf_path, pdf_file_path)
                    print(f"\033[93m已成功下載PDF，檔案移動至：{pdf_file_path}\033[0m", flush=True)
                    found = True
                    break

            if not found:
                print("未找到下載的 PDF 檔案", flush=True)
                # print(f"[調試] 預期檔名格式：{pdf_prefix}*.pdf", flush=True)

            # 切換回主頁面
            driver.switch_to.default_content()
            time.sleep(2)  # 確保操作間隔

        print("重要設施查詢完成", flush=True)

    except (TimeoutException, NoSuchElementException) as e:
        print(f"載入網頁超時或元素未找到: {e}", flush=True)
    except Exception as e:
        print(f"執行過程中發生錯誤: {e}", flush=True)
    finally:
        # 確保瀏覽器在任何情況下都會關閉
        if driver:
            try:
                driver.quit()
                # print("瀏覽器已關閉", flush=True)
            except Exception as e:
                print(f"關閉瀏覽器時發生錯誤: {e}", flush=True)

        # 🔥 清理臨時下載資料夾
        try:
            if download_directory and os.path.exists(download_directory):
                shutil.rmtree(download_directory)
                # print("已清理臨時下載資料夾", flush=True)
        except Exception as e:
            print(f"清理下載資料夾時發生錯誤: {e}", flush=True)

def notify_main_program():
    # print("重要設施已執行結束", flush=True)
    pass

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"子程式執行過程中發生錯誤: {e}", flush=True)
    finally:
        # 在結尾通知主程式
        notify_main_program()