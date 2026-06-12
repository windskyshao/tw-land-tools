import sys
import os

# 打包環境修正：將 _internal 加入 sys.path
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
print("歡迎使用【全國土地使用分區】自動化小程式，模組載入中...", flush=True)
import json
import ctypes
import keyboard  # 用於系統級鍵盤事件
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import Select
from PIL import Image
from fpdf import FPDF
import time
from selenium.common.exceptions import UnexpectedAlertPresentException, TimeoutException, NoAlertPresentException
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window

# 基準目錄設定（data.json 和工作資料夾的位置）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

# 🔥 解析 --indices 參數：主程式傳遞使用者勾選的筆次（0-based，逗號分隔）
#    例如 --indices 0,2,5 表示只處理第 1、3、6 筆
#    若沒有 --indices 但有 --all → 全部（向後相容）
#    兩者都沒有 → 只處理第一筆（與舊版預設一致）
selected_indices = None
for _i, _arg in enumerate(sys.argv):
    if _arg == '--indices' and _i + 1 < len(sys.argv):
        try:
            selected_indices = [int(x.strip()) for x in sys.argv[_i + 1].split(',') if x.strip()]
        except ValueError:
            selected_indices = None
        break

run_all = '--all' in sys.argv  # 向後相容
if selected_indices is not None:
    print(f"[自選模式] 將處理使用者勾選的 {len(selected_indices)} 筆", flush=True)
elif run_all:
    print("[批次模式] 將逐筆查詢 data.json 中的所有地號", flush=True)

# 自動偵測 DPI 縮放比例
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

# 視窗大小設定：從 main.py 生成的設定檔讀取，自動填滿剩餘空間
luz_window_width = 1024
luz_window_height = 1024
luz_window_x = 0
luz_window_y = 0

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

            # Chrome 視窗高度需要加上標題欄高度（約 32px 邏輯像素）
            # 這樣整個視窗（包含標題欄）才會和主程式一樣高
            CHROME_TITLEBAR_HEIGHT = 32
            chrome_height_logical = int(chrome_height / dpi_scale) + CHROME_TITLEBAR_HEIGHT

            luz_window_width = chrome_width_logical
            luz_window_height = chrome_height_logical
            luz_window_x = chrome_x_logical
            luz_window_y = chrome_y_logical
except Exception as e:
    pass

# 讀取 data.json 文件
with open(get_data_json_path(), 'r', encoding='utf-8') as file:
    data_list = json.load(file)

# 設定 WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# 使用 ChromeDriverManager 安裝 ChromeDriver 並應用選項
driver = create_chrome_driver(options=options)

# 啟動後立即設定視窗大小和位置
try:
    driver.set_window_size(luz_window_width, luz_window_height)
    driver.set_window_position(luz_window_x, luz_window_y)
except Exception as e:
    print(f"[WARNING] 視窗設定失敗: {e}，繼續執行")

# 指定要打開的網址 - 全國土地使用分區查詢系統
url = "https://luz.tcd.gov.tw/WEB/"

# 使用瀏覽器打開網址
driver.get(url)

# 網頁載入後重新設定視窗大小
try:
    driver.set_window_size(luz_window_width, luz_window_height)
    driver.set_window_position(luz_window_x, luz_window_y)
except:
    pass

# 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
verify_and_fix_chrome_window(driver)

# 等待元素加載
driver.implicitly_wait(5)

print("【全國土地使用分區】網頁已開啟", flush=True)

# 自動判斷是否需要縮放頁面（解決高 DPI 或小螢幕下的排版問題）
zoom_count = 0
try:
    config_path = os.path.join(BASE_DIR, 'window_config.json')
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        current_dpi = config.get('dpi_scale', 1.0)
        screen_width = config.get('screen_width', 1920)
        logical_width = screen_width / current_dpi

        # 根據 DPI 和螢幕解析度決定縮放次數
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
        # 使用 keyboard 庫發送系統級的 Ctrl + - 按鍵（真正的瀏覽器縮放）
        # 先點擊頁面確保 Chrome 視窗有焦點
        body = driver.find_element(By.TAG_NAME, 'body')
        body.click()
        time.sleep(0.5)

        # 根據 DPI 決定按幾次 Ctrl + -
        print(f"[頁面縮放] 正在使用系統鍵盤縮放（按 {zoom_count} 次）...", flush=True)
        for i in range(zoom_count):
            keyboard.press_and_release('ctrl+-')  # 系統級按鍵
            time.sleep(0.5)

        print(f"[頁面縮放] ✓ 已使用系統鍵盤縮放至 {target_zoom}", flush=True)
    except Exception as e:
        print(f"[頁面縮放] 設定縮放時發生錯誤: {e}", flush=True)
else:
    print("[頁面縮放] 螢幕配置正常，不需要縮放", flush=True)

# 等待頁面 JavaScript 完全載入
time.sleep(3)

# ========== 步驟1：點擊「系統功能」按鈕 ==========
print("步驟1：點擊【系統功能】按鈕...", flush=True)
try:
    # 透過 ID 找到 menuLB 區塊內的圖片按鈕
    menu_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@id='menuLB']//img[@onclick='loadmenu();']"))
    )
    driver.execute_script("arguments[0].click();", menu_btn)
    print("  已點擊【系統功能】按鈕", flush=True)
    time.sleep(2)
except Exception as e:
    # 備援方案：直接執行 JavaScript 函數
    try:
        driver.execute_script("loadmenu();")
        print("  已透過 JavaScript 執行 loadmenu()", flush=True)
        time.sleep(2)
    except Exception as e2:
        print(f"  點擊【系統功能】失敗: {e2}", flush=True)

# 🔥 建立工作資料夾（以 data.json 第一筆為準，所有截圖/PDF 都存到此資料夾）
first_entry = data_list[0]
first_region = first_entry['area']
first_section = first_entry['section']
first_lot = first_entry['lot_number']
base_directory = get_work_folder(f"{first_region}{first_section}-{first_lot}")
os.makedirs(os.path.join(base_directory, "1.基本資料", "png"), exist_ok=True)

# 決定要處理的地號清單
# 決定要處理的地號清單：
#   1) --indices 優先：依使用者勾選的 index 順序挑選
#   2) --all 次之：全部
#   3) 都沒有：只第一筆（維持舊版預設）
if selected_indices is not None:
    entries_to_process = [data_list[i] for i in selected_indices if 0 <= i < len(data_list)]
elif run_all:
    entries_to_process = data_list
else:
    entries_to_process = [data_list[0]] if data_list else []
total_entries = len(entries_to_process)


def process_entry(driver, entry, idx, total):
    """查詢單筆地號 → 截圖 → PDF"""
    target_city = entry['city']
    target_area = entry['area']
    target_section = entry['section']
    target_lot = entry['lot_number']

    if total > 1:
        print("\n" + "=" * 50, flush=True)
        print(f"\033[96m【第 {idx}/{total} 筆】{target_city} {target_area}{target_section} {target_lot}\033[0m", flush=True)
        print("=" * 50, flush=True)

    # 🔥 多筆時：上一筆結束後選單是「收合」狀態（步驟7.10 收回的），
    #   表單收合時段名下拉點不到 → 這裡先確認表單已展開，避免第2筆起選段名失敗
    try:
        _expanded = driver.execute_script("""
            var m = document.getElementById('menuL');
            return !!(m && m.offsetParent !== null && m.getBoundingClientRect().width > 100);
        """)
        if not _expanded:
            print("  (偵測到表單收合，先展開系統功能選單)", flush=True)
            driver.execute_script("loadmenu();")
            time.sleep(1.2)
    except Exception:
        pass

    # ========== 步驟2：選擇「縣市」==========
    print(f"步驟2：選擇【縣市】: {target_city}...", flush=True)
    try:
        county_select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "COUNTY_0101"))
        )
        county_select = Select(county_select_element)
        county_select.select_by_visible_text(target_city)
        print(f"  已選擇縣市: {target_city}", flush=True)
        time.sleep(2)
    except Exception as e:
        print(f"  選擇縣市失敗: {e}", flush=True)

    # ========== 步驟3：選擇「鄉鎮市」==========
    print(f"步驟3：選擇【鄉鎮市】: {target_area}...", flush=True)
    try:
        town_select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "TOWN_0101"))
        )
        town_select = Select(town_select_element)
        town_select.select_by_visible_text(target_area)
        print(f"  已選擇鄉鎮市: {target_area}", flush=True)
        time.sleep(1)
    except Exception as e:
        print(f"  選擇鄉鎮市失敗: {e}", flush=True)

    # ========== 步驟4：確保「地籍」accordion 展開（而非盲目 toggle）==========
    # 🔥 關鍵：第二筆起 accordion 可能已經展開，再點會收合造成段名選單消失
    print("步驟4：確認【地籍】已展開...", flush=True)

    def _cadastral_is_active():
        """檢查地籍 accordion 是否已展開"""
        try:
            return bool(driver.execute_script("""
                var hs = document.querySelectorAll('h3.ui-accordion-header');
                for (var i = 0; i < hs.length; i++) {
                    var a = hs[i].querySelector('a');
                    if (a && a.textContent.trim() === '地籍') {
                        return hs[i].classList.contains('ui-accordion-header-active');
                    }
                }
                return false;
            """))
        except Exception:
            return False

    def _click_cadastral_header():
        """點擊地籍 accordion header"""
        btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//h3[contains(@class, 'ui-accordion-header')]//a[text()='地籍']"))
        )
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(0.8)

    try:
        if _cadastral_is_active():
            print("  【地籍】已展開，不重複點擊", flush=True)
        else:
            _click_cadastral_header()
            print("  已展開【地籍】", flush=True)
    except Exception as e:
        print(f"  【地籍】處理失敗: {e}", flush=True)

    # ========== 步驟5：選擇「段名」==========
    # 🔥 保留原本「點下拉 → 點 ui-menu-item」的路徑（原本跑得通的做法）；
    #    若失敗則用 Select options + JS 當備援；兩者都失敗才真正放棄
    def _try_select_section_once():
        """執行一次段名選取嘗試，成功回 True"""
        try:
            dropdown_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//select[@id='LANDSEC2_0101']/following-sibling::span//a[contains(@class, 'custom-combobox-toggle')]"))
            )
            driver.execute_script("arguments[0].click();", dropdown_btn)
            print(f"  已點擊段名下拉按鈕", flush=True)

            section_item = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, 'ui-menu-item')]//div[contains(text(), '{target_section}')]"))
            )
            section_item.click()
            print(f"  已選擇段名: {target_section}", flush=True)
            return True
        except Exception:
            pass

        # 備援：直接從 Select 找目標選項，用 JS 設值
        try:
            section_select = Select(driver.find_element(By.ID, "LANDSEC2_0101"))
            for option in section_select.options:
                if target_section in option.text:
                    val = option.get_attribute('value')
                    txt = option.text
                    driver.execute_script(f"""
                        $('#LANDSEC2_0101').val('{val}').trigger('change');
                        $('#LANDSEC2_0101').siblings('.custom-combobox').find('input.custom-combobox-input').val('{txt}');
                    """)
                    print(f"  已透過 JavaScript 選擇段名: {txt}", flush=True)
                    return True
        except Exception:
            pass
        return False

    print(f"步驟5：選擇【段名】: {target_section}...", flush=True)
    section_ok = _try_select_section_once()

    # 第一次失敗：重觸發鄉鎮市 change 強制重載段列表，再試第二次
    if not section_ok:
        print(f"  第 1 次失敗，重觸發鄉鎮市 change 以重載段列表...", flush=True)
        try:
            driver.execute_script("$('#TOWN_0101').trigger('change');")
            time.sleep(1)
        except Exception as e:
            print(f"  重觸發失敗: {e}", flush=True)
        section_ok = _try_select_section_once()

    if not section_ok:
        print(f"\033[91m✗ 第 {idx}/{total} 筆：段名選擇失敗，跳過此筆（不截圖避免誤存舊圖）\033[0m", flush=True)
        return False

    # ========== 步驟6：輸入「地號」（等待可互動）==========
    print(f"步驟6：輸入【地號】: {target_lot}...", flush=True)
    lot_ok = False
    try:
        lot_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "CADANO_0101"))
        )
        lot_input.clear()
        lot_input.send_keys(target_lot)
        print(f"  已輸入地號: {target_lot}", flush=True)
        lot_ok = True
        time.sleep(0.5)
    except Exception as e:
        print(f"  輸入地號失敗: {e}", flush=True)

    if not lot_ok:
        print(f"\033[91m✗ 第 {idx}/{total} 筆：地號輸入失敗，跳過此筆（不截圖避免誤存舊圖）\033[0m", flush=True)
        return False

    # ========== 步驟7：點擊「搜尋」按鈕 ==========
    print("步驟7：點擊【搜尋】按鈕...", flush=True)
    try:
        driver.execute_script("ZoomToData(8,'CADANO_0101','CADA');")
        print("  已執行搜尋", flush=True)
    except Exception as e:
        print(f"  搜尋失敗: {e}", flush=True)

    # 🔍 診斷：搜尋後 pin 狀態（判斷 pin 有沒有出現）
    time.sleep(2)
    _d = driver.execute_script("return [document.querySelectorAll('svg image').length, document.querySelectorAll('#mapDiv_graphics_layer image').length, document.querySelectorAll('#mapDiv_graphics_layer circle').length];")
    print(f"  [診斷-搜尋後] svgImgs={_d[0]}, layerImgs={_d[1]}, circles={_d[2]}", flush=True)

    # ========== 步驟7.5：縮回左側搜尋表框（點右上 back.png → loadmenu()），避免擋住地圖 ==========
    print("步驟7.5：縮回左側搜尋表框...", flush=True)
    try:
        time.sleep(1)  # 等搜尋結果穩定後再收合
        driver.execute_script("loadmenu();")
        print("  已縮回搜尋表框", flush=True)
    except Exception as e:
        print(f"  縮回表框失敗（不影響後續）: {e}", flush=True)
    _c = driver.execute_script("return document.querySelectorAll('svg image').length;")
    print(f"  [診斷-縮回後] svgImgs={_c}", flush=True)
    # 🔥 點「查詢」後 pin 會消失，所以先記住 pin 座標（＝地號所在），等下用座標去點
    #   小地號時「最下面尖端」可能落到下方鄰地/道路，改用「pin 中心」較準
    _pin_xy = driver.execute_script("""
        var pin = document.querySelector('#mapDiv_graphics_layer image');
        if (!pin) return null;
        var r = pin.getBoundingClientRect();
        return [Math.round(r.left + r.width/2), Math.round(r.top + r.height/2)];
    """)
    print(f"  [記住] 錨點中心座標 = {_pin_xy}", flush=True)

    # ========== 步驟7.6：點工具列「查詢」鈕（進入查詢/identify 模式）==========
    print("步驟7.6：點工具列【查詢】鈕...", flush=True)
    try:
        time.sleep(0.5)
        try:
            # w2ui 工具列按鈕：點該按鈕元素觸發 onclick（最接近真人點擊）
            _q_btn = driver.find_element(By.CSS_SELECTOR, "#tb_layout_main_toolbar_item_itemrdo3 table.w2ui-button")
            driver.execute_script("arguments[0].click();", _q_btn)
        except Exception:
            # 退路：直接呼叫 w2ui 工具列的 click
            driver.execute_script("w2ui['layout_main_toolbar'].click('itemrdo3');")
        print("  已點擊【查詢】鈕（進入查詢模式）", flush=True)
    except Exception as e:
        print(f"  點擊【查詢】鈕失敗: {e}", flush=True)
    _c = driver.execute_script("return document.querySelectorAll('svg image').length;")
    print(f"  [診斷-查詢後] svgImgs={_c}", flush=True)

    # ========== 步驟7.7：用記住的座標點地圖查詢，並「驗證地號」，不對就自動微調重點 ==========
    print("步驟7.7：在錨點座標點地圖（查詢該地號）...", flush=True)
    try:
        if not _pin_xy:
            print("  沒有記住錨點座標（搜尋後 pin 未出現，略過）", flush=True)
        else:
            bx, by = int(_pin_xy[0]), int(_pin_xy[1])
            target = str(target_lot).strip()
            _CLICK_JS = """
                var cx=arguments[0], cy=arguments[1];
                var tgt=document.elementFromPoint(cx,cy);
                if(!tgt) return false;
                var opts={bubbles:true,cancelable:true,view:window,clientX:cx,clientY:cy,button:0,buttons:1,pointerId:1,pointerType:'mouse',isPrimary:true};
                ['pointerover','pointermove','pointerdown','mousedown','pointerup','mouseup','click'].forEach(function(t){
                    var ev; try{ev=(t.indexOf('pointer')===0)?new PointerEvent(t,opts):new MouseEvent(t,opts);}catch(e){ev=new MouseEvent(t,opts);}
                    tgt.dispatchEvent(ev);
                });
                return true;
            """
            _READ_LOT_JS = """
                var el=document.querySelector('#identify_result'); if(!el) return null;
                var rows=el.querySelectorAll('tr');
                for(var i=0;i<rows.length;i++){var tds=rows[i].querySelectorAll('td');
                    if(tds.length>=2 && tds[0].innerText.trim()==='地號') return tds[1].innerText.trim();}
                return null;
            """
            # 先試 pin 中心，不對再上下左右微調
            candidates = [(0, 0), (0, -8), (0, 8), (-7, 0), (7, 0), (0, -15), (-7, -8), (7, -8)]
            matched = False
            for dx, dy in candidates:
                cx, cy = bx + dx, by + dy
                driver.execute_script(_CLICK_JS, cx, cy)
                lot = None
                for _ in range(6):  # 等查詢結果，最多約 3 秒
                    time.sleep(0.5)
                    lot = driver.execute_script(_READ_LOT_JS)
                    if lot:
                        break
                print(f"  試點 @({cx},{cy}) → 查到地號={lot}", flush=True)
                if lot and str(lot).strip() == target:
                    print(f"  ✓ 查詢結果正確：地號 {lot}", flush=True)
                    matched = True
                    break
            if not matched:
                print(f"  ⚠ 試了 {len(candidates)} 個位置仍未對到目標地號 {target}（地號可能太小，請人工確認）", flush=True)
    except Exception as e:
        print(f"  查詢點擊失敗: {e}", flush=True)

    # ========== 步驟7.8～7.10：重新顯示錨點(pin)，讓最終截圖更明確 ==========
    #   查詢結果確認後，再展開選單→重新搜尋(pin 重現)→收回選單
    try:
        print("步驟7.8：重新展開系統功能選單...", flush=True)
        driver.execute_script("loadmenu();")
        time.sleep(1)
        print("步驟7.9：再次搜尋（讓錨點重新出現）...", flush=True)
        driver.execute_script("ZoomToData(8,'CADANO_0101','CADA');")
        time.sleep(2)
        print("步驟7.10：再次縮回系統功能選單...", flush=True)
        driver.execute_script("loadmenu();")
        time.sleep(1)
        _imgs2 = driver.execute_script("return document.querySelectorAll('#mapDiv_graphics_layer image').length;")
        print(f"  錨點已重新顯示（image={_imgs2}）", flush=True)
    except Exception as e:
        print(f"  重新顯示錨點失敗（不影響截圖）: {e}", flush=True)

    # ========== 步驟8a：對準錨點放大一級（＝滑鼠對準錨點往前滾一格滾輪）==========
    #   先放大、再右移；放大以錨點為中心，之後再做右移，確保右移是最後一步不被洗掉
    print("步驟8a：對準錨點放大一級...", flush=True)
    try:
        zres = driver.execute_script("""
            if (typeof map === 'undefined') return 'map not found';
            try {
                // 取錨點(pin)的地理座標（找有圖形的圖層，取最後一個圖形）
                var pinGeo = null;
                var layers = [];
                if (map.graphics) layers.push(map.graphics);
                try { for (var k in map._layers) { var L = map.getLayer(k); if (L && L.graphics && L.graphics.length) layers.push(L); } } catch(e) {}
                for (var li = layers.length - 1; li >= 0 && !pinGeo; li--) {
                    var gs = layers[li].graphics;
                    if (gs && gs.length) {
                        var geom = gs[gs.length - 1].geometry;
                        if (geom) pinGeo = (geom.type === 'point') ? geom
                            : (geom.getCenter ? geom.getCenter()
                            : (geom.getExtent ? geom.getExtent().getCenter() : null));
                    }
                }
                var lvl = map.getLevel();
                var maxL = 19; try { if (map.getMaxZoom) maxL = map.getMaxZoom(); } catch(e) {}
                if (maxL >= 0 && lvl >= maxL) return 'already-max level=' + lvl;
                if (!pinGeo) { map.setLevel(lvl + 1); return 'zoom(no-pin) level=' + (lvl+1); }
                // 以錨點為中心放大一級（錨點移到畫面中央、變大）；右移留給步驟8b
                map.centerAndZoom(pinGeo, lvl + 1);
                return 'success level=' + lvl + '->' + (lvl+1);
            } catch(e) { return 'error: ' + e.message; }
        """)
        print(f"  放大結果: {zres}", flush=True)
        time.sleep(1.5)
    except Exception as e:
        print(f"  放大失敗（不影響截圖）: {e}", flush=True)

    # ========== 步驟8b：把地號往右移（置中於查詢結果面板右側剩餘空間）==========
    #   這是最後一個地圖動作，確保「右移」不會被放大覆蓋掉
    print("步驟8b：調整地圖位置（地號置中於查詢結果面板右側剩餘空間）...", flush=True)
    try:
        result = driver.execute_script("""
            if (typeof map === 'undefined') return 'map not found';
            try {
                // 地圖繪圖區(svg)的螢幕範圍
                var svg = document.querySelector('#mapDiv_gc');
                if (!svg) { var g = document.querySelector('#mapDiv_graphics_layer'); svg = g ? g.ownerSVGElement : null; }
                var mr = svg ? svg.getBoundingClientRect() : {left:0, right:map.width};
                // 查詢結果面板的右緣：用 #identify_result 往上找對話框（最穩），找不到退回左側選單
                var panelRight = mr.left; var src = 'none';
                var content = document.querySelector('#identify_result');
                var dlg = content ? content.closest('.ui-dialog') : null;
                if (dlg && dlg.offsetParent !== null) { panelRight = dlg.getBoundingClientRect().right; src = 'dialog'; }
                else {
                    var menu = document.getElementById('menuL') || document.getElementById('menuLB');
                    if (menu && menu.offsetParent !== null) { panelRight = menu.getBoundingClientRect().right; src = 'menu'; }
                }
                // 把地號往右移到「面板右緣 ~ 地圖右緣」的中間 → 平移量 =(面板右緣 - 地圖左緣)/2
                var panPixels = (panelRight - mr.left) / 2;
                if (panPixels < 0) panPixels = 0;
                var extent = map.extent;
                var pixelSize = (extent.xmax - extent.xmin) / map.width;
                var center = extent.getCenter();
                map.centerAt(new esri.geometry.Point(center.x - panPixels * pixelSize, center.y, map.spatialReference));
                return 'success src=' + src + ' panPixels=' + Math.round(panPixels) + ' panelRight=' + Math.round(panelRight) + ' mapLeft=' + Math.round(mr.left);
            } catch(e) { return 'error: ' + e.message; }
        """)
        print(f"  平移結果: {result}", flush=True)
    except Exception as e:
        print(f"  調整地圖位置失敗: {e}", flush=True)

    # 倒數等待使用者調整畫面
    print("\n" + "=" * 50, flush=True)
    print(f"\033[93m第 {idx}/{total} 筆查詢完成！\033[0m", flush=True)
    print("\033[93m開始倒數 10 秒，請調整網頁位置並確認查詢結果，倒數結束後將自動截圖\033[0m", flush=True)
    print("=" * 50, flush=True)

    for i in range(10, 0, -1):
        print(f"倒數：{i} 秒", end='\r', flush=True)
        time.sleep(1)
    print("倒數結束                ", flush=True)

    # 每筆地號存獨立的 PNG 與 PDF
    base_filename = f"07_全國土地使用分區-{target_area}{target_section}-{target_lot}"
    screenshot_path = os.path.join(base_directory, "1.基本資料", "png", base_filename + ".png")
    pdf_path = os.path.join(base_directory, "1.基本資料", base_filename + ".pdf")

    try:
        driver.save_screenshot(screenshot_path)
        print(f"\033[93m網頁截圖已保存為 {screenshot_path}\033[0m", flush=True)
    except Exception as e:
        print(f"截圖時發生錯誤: {e}", flush=True)
        return False

    try:
        img = Image.open(screenshot_path)
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.add_page()

        a4_width_mm, a4_height_mm = 297, 210
        dpi = 300
        scale_factor = dpi / 72
        img_width_px = int(a4_width_mm * scale_factor)
        img_height_px = int(a4_height_mm * scale_factor)

        img_resized = img.resize((img_width_px, img_height_px), Image.Resampling.LANCZOS)
        img_resized.save(screenshot_path, dpi=(dpi, dpi))

        pdf.image(screenshot_path, x=0, y=0, w=a4_width_mm, h=a4_height_mm)
        pdf.output(pdf_path)
        print(f"\033[93mPDF 已經儲存至 {pdf_path}\033[0m", flush=True)
    except Exception as e:
        print(f"將 PNG 轉成 PDF 時發生錯誤: {e}", flush=True)

    return True


# 主流程：依序處理每筆地號
if run_all and total_entries > 1:
    print(f"\n\033[96m【批次模式】共 {total_entries} 筆地號待查詢\033[0m\n", flush=True)

_success_count = 0
_skip_count = 0
_skipped_entries = []

for idx, entry in enumerate(entries_to_process, start=1):
    try:
        ok = process_entry(driver, entry, idx, total_entries)
        if ok:
            _success_count += 1
        else:
            _skip_count += 1
            _skipped_entries.append(f"{entry.get('area','')}{entry.get('section','')} {entry.get('lot_number','')}")
    except Exception as e:
        _skip_count += 1
        _skipped_entries.append(f"{entry.get('area','')}{entry.get('section','')} {entry.get('lot_number','')}")
        print(f"[錯誤] 處理第 {idx} 筆時發生未預期錯誤: {e}", flush=True)
        # 繼續處理下一筆，不中斷整體流程

if total_entries > 1:
    print(f"\n\033[92m✓ 成功：{_success_count} 筆\033[0m", flush=True)
    if _skip_count > 0:
        print(f"\033[91m✗ 跳過：{_skip_count} 筆（網頁狀態異常未截圖）\033[0m", flush=True)
        for s in _skipped_entries:
            print(f"    - {s}", flush=True)
    print("", flush=True)

# 結束 Selenium session
driver.quit()

def notify_main_program():
    print("全國土地使用分區查詢已完成執行", flush=True)

if __name__ == "__main__":
    try:
        # 子程式的主要邏輯
        print("執行全國土地使用分區查詢主要邏輯", flush=True)

    except Exception as e:
        print(f"全國土地使用分區查詢執行過程中發生錯誤: {e}", flush=True)
    finally:
        # 在結尾通知主程式
        notify_main_program()
