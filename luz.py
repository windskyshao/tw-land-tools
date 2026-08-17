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
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window, apply_page_zoom

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
# 🔥 不要再寫死 User-Agent！
#    本站(luz.nlma.gov.tw)已改掛 Cloudflare Turnstile 人機驗證。原本寫死 Chrome/127，
#    但實機是 Chrome 150，UA 與瀏覽器真實 Client Hints(sec-ch-ua) 版本對不上 →
#    Cloudflare 直接判定為機器人，導致「連手動點驗證方塊也一直失敗」。
#    移除覆寫後由 Chrome 送出自己真實且一致的 UA，通過率才正常。
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# ─────────────────────────────────────────────────────────────────────────
# 🔥 Cloudflare Turnstile 對策：附著模式（attach mode）
#    本站(luz.nlma.gov.tw)已掛 Cloudflare 人機驗證。實測「改 UA」「拿掉自動化旗標」
#    都仍被擋 —— 因為 Cloudflare 也會偵測 ChromeDriver 自身注入的痕跡(cdc_ 變數/CDP)，
#    導致連「手動點驗證方塊」都失敗。
#    解法：先用【一般方式】啟動 Chrome（完全不經 ChromeDriver，等同使用者自己開的），
#          由使用者通過驗證取得 cf_clearance cookie 後，程式再【附著】上去接手自動化。
#    設定檔固定於 chrome_profile_luz，cookie 會被記住，之後多半不必再驗。
# ─────────────────────────────────────────────────────────────────────────
LUZ_USE_ATTACH_MODE = True      # 若日後網站移除驗證，可改 False 回到原本模式
LUZ_MANUAL_ALL = True           # 🔥 全手動模式：程式只自動填欄位，之後【搜尋+查詢+點地號+過驗證+調整畫面】
                                #  全由使用者做，畫面確定後按 Enter → 程式才截圖並擷取分區文字，再走下一筆。
                                #  這是唯一能同時「有原生資訊框+有文字+不用跟 Cloudflare 纏」的方式。
LUZ_DO_IDENTIFY = False         # 是否要「查地號」帶出分區資訊框(左側查詢結果)。（LUZ_MANUAL_ALL=True 時由使用者手動查，此旗標不影響）
                                #  True：程式自動點【查詢】鈕，但「點圖釘查地號」由使用者親手點
                                #        (手動模式)——因為程式點 identify 會觸發隱形 Cloudflare 直接失敗，
                                #        只有真人點擊才通過。想要地號分區資訊框就用 True。
                                #  False：完全略過查地號，搜尋完直接截圖(只有圖釘+分區顏色，沒有資訊框)。
LUZ_ASK_MANUAL_PASS = False     # （已停用）舊做法：接手前先手動查一次取得通行證。
                                #  問題是按 Enter 後程式仍會從頭自動搜尋、又觸發驗證且必失敗，
                                #  已改由 LUZ_MANUAL_SEARCH 在「搜尋那一步」交給使用者。
LUZ_MANUAL_SEARCH = True        # ⚠ 實測結論：務必維持 True。
                                # True =【搜尋】由使用者親手按、自行過驗證後按 Enter 接續（唯一可行）
                                # False=程式用 execute_script 自動搜尋 → 2026-07-23 實測：
                                #        驗證彈窗立刻出現，使用者點勾選框一律失敗，且會無限重複要求驗證、
                                #        永遠不會通過。原因見下方 _luz_wait_if_challenge 的說明。
LUZ_DEBUG_PORT = 9222
LUZ_URL = "https://luz.tcd.gov.tw/WEB/"

_luz_profile = os.path.join(BASE_DIR, "chrome_profile_luz")


def _luz_find_chrome_exe():
    """找出實體 chrome.exe（附著模式要用一般方式啟動它）"""
    cands = [
        os.path.join(os.environ.get('PROGRAMFILES', r'C:\Program Files'),
                     r'Google\Chrome\Application\chrome.exe'),
        os.path.join(os.environ.get('PROGRAMFILES(X86)', r'C:\Program Files (x86)'),
                     r'Google\Chrome\Application\chrome.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''),
                     r'Google\Chrome\Application\chrome.exe'),
    ]
    for c in cands:
        if c and os.path.exists(c):
            return c
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe") as k:
            p = winreg.QueryValue(k, None)
            if p and os.path.exists(p):
                return p
    except Exception:
        pass
    return None


def _luz_port_open(port, host='127.0.0.1'):
    import socket
    try:
        with socket.socket() as s:
            s.settimeout(0.5)
            return s.connect_ex((host, port)) == 0
    except Exception:
        return False


def _luz_challenge_present(driver):
    """頁面上是否正顯示 Cloudflare 人機驗證。"""
    try:
        if driver.find_elements(By.CSS_SELECTOR, "iframe[src*='challenges.cloudflare.com']"):
            return True
        src = driver.page_source or ""
        for kw in ("正在進行驗證程序", "驗證您是人類", "Just a moment",
                   "cf-challenge", "challenge-platform"):
            if kw in src:
                return True
    except Exception:
        pass
    return False


def _luz_extract_zone(driver):
    """從『查詢結果』面板擷取分區文字（使用者手動查地號後，面板還在畫面上時呼叫）。
       回傳 dict：{'rows': {欄位:值...}, 'raw': 整段文字}；擷取不到回 {}。"""
    try:
        js = r"""
        var el = document.querySelector('#identify_result');
        var dlg = el ? el.closest('.ui-dialog') : null;
        if(!el && !dlg) return null;
        var out = {rows:{}, raw:''};
        var box = dlg || el;
        out.raw = ((box.innerText)||'').replace(/\s+\n/g,'\n').trim();
        var trs = (el||box).querySelectorAll('tr');
        for(var i=0;i<trs.length;i++){
            var tds = trs[i].querySelectorAll('td,th');
            if(tds.length>=2){
                var k=((tds[0].innerText)||'').trim();
                var v=((tds[1].innerText)||'').trim();
                if(k && k.length<=12) out.rows[k]=v;
            }
        }
        return JSON.stringify(out);
        """
        s = driver.execute_script(js)
        if not s:
            return {}
        d = json.loads(s)
        return d if (d.get('rows') or d.get('raw')) else {}
    except Exception as e:
        print(f"  擷取分區文字失敗: {e}", flush=True)
        return {}


def _luz_shot_and_pdf(driver, base_directory, base_filename):
    """截圖存 PNG + 轉 PDF 到 1.基本資料。回傳 True/False。"""
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


def _luz_save_zone(city, area, section, lot, zone):
    """把一筆分區文字加入累積清單並寫出 4.其他相關/全國土地使用分區資料-*.json（備用，高雄仍以高雄都市計畫優先）。"""
    if not zone:
        print("  [分區文字] 這次沒抓到查詢結果面板內容（可能沒查地號、或面板已關閉）。", flush=True)
        return
    try:
        rows = zone.get('rows', {}) or {}
        _use = rows.get('使用分區') or rows.get('分區') or rows.get('土地使用分區') or ''
        print(f"  [分區文字] 地號={lot} 使用分區={_use or '(詳見原始文字)'}", flush=True)
        _luz_zone_all.append({
            "縣市": city, "鄉鎮市區": area, "段名": section, "地號": lot,
            "查詢結果": rows, "原始文字": zone.get('raw', '')
        })
        os.makedirs(os.path.dirname(_luz_zone_json), exist_ok=True)
        with open(_luz_zone_json, 'w', encoding='utf-8') as f:
            json.dump(_luz_zone_all, f, ensure_ascii=False, indent=2)
        print(f"  [已存] 分區文字 → {_luz_zone_json}", flush=True)
    except Exception as e:
        print(f"  存分區 json 失敗: {e}", flush=True)


def _luz_wait_if_challenge(driver, max_wait=180):
    """偵測 Cloudflare 人機驗證畫面；一旦出現就【暫停自動化】直到它通過。

    ⚠ 這很重要：本站是在「點擊搜尋」時才觸發驗證。原本程式不會察覺，
       驗證還開著就一路往下跑（縮回表框、點查詢、點地圖、再次搜尋…），
       Cloudflare 看到「驗證期間仍有自動化互動」會直接判定失敗 → 永遠過不了。
       改成偵測到就停手等待，期間完全不碰頁面。
    """
    import time as _t

    def _present():
        return _luz_challenge_present(driver)

    if not _present():
        return True

    print("", flush=True)
    print("\033[93m⚠ 偵測到 Cloudflare 人機驗證 → 已【暫停自動化】\033[0m", flush=True)
    print("\033[93m   請到 Chrome 視窗點一下驗證方塊；期間程式不會再碰頁面。\033[0m", flush=True)
    print("\033[93m   ※ 若顯示「驗證失敗」先別關掉，Turnstile 會自己重試，\033[0m", flush=True)
    print("\033[93m     通常再等 7~10 秒就會自動通過、彈窗消失。\033[0m", flush=True)
    print("\033[93m   通過後程式會自動偵測到並接著跑（會自動幫你重新查詢）。\033[0m", flush=True)
    waited = 0
    while waited < max_wait:
        _t.sleep(2)
        waited += 2
        if not _present():
            print("\033[92m✓ 驗證已通過，恢復自動化\033[0m", flush=True)
            _t.sleep(2)   # 讓頁面把結果載完再繼續
            return True
        if waited % 20 == 0:
            print(f"   仍在等待驗證…（{waited} 秒 / 上限 {max_wait} 秒）", flush=True)
    print("\033[91m⚠ 等待驗證逾時，仍嘗試繼續\033[0m", flush=True)
    return False


def _luz_mark_profile_clean():
    """把專用設定檔標記為『上次正常關閉』。

    程式（或使用者按「關閉全國土地使用分區」）中途結束時，這顆 Chrome 是被強制結束的，
    會在設定檔留下 crash 記號 → 下次啟動就跳「你要還原網頁嗎？Chrome 未正確關閉」。
    啟動前先把記號改成正常，就不會再問。
    """
    try:
        pref = os.path.join(_luz_profile, "Default", "Preferences")
        if not os.path.exists(pref):
            return
        with open(pref, 'r', encoding='utf-8') as f:
            _d = json.load(f)
        _p = _d.setdefault("profile", {})
        _p["exit_type"] = "Normal"
        _p["exited_cleanly"] = True
        with open(pref, 'w', encoding='utf-8') as f:
            json.dump(_d, f)
    except Exception:
        pass


def _luz_create_driver_attached():
    """用一般方式開 Chrome（不經 ChromeDriver）再附著。失敗回傳 None 由呼叫端退回原模式。

    註：本站的人機驗證是「點擊搜尋」時才觸發，開頁階段不會出現，
        所以這裡不再要求使用者先按 Enter；真的跳驗證時，
        由 _luz_wait_if_challenge() 在搜尋步驟自動暫停並等待。
    """
    import subprocess
    import time as _t

    chrome_exe = _luz_find_chrome_exe()
    if not chrome_exe:
        print("[附著模式] 找不到 chrome.exe，改用一般模式", flush=True)
        return None

    if _luz_port_open(LUZ_DEBUG_PORT):
        print(f"[附著模式] 偵測到 {LUZ_DEBUG_PORT} 埠已有 Chrome，直接附著", flush=True)
    else:
        try:
            os.makedirs(_luz_profile, exist_ok=True)
            _luz_mark_profile_clean()   # 先清掉上次的 crash 記號
            subprocess.Popen(
                [chrome_exe,
                 f"--remote-debugging-port={LUZ_DEBUG_PORT}",
                 f"--user-data-dir={_luz_profile}",
                 "--no-first-run", "--no-default-browser-check",
                 # 🔥 不載入任何擴充功能：此設定檔只給自動查詢用，
                 #    可避免 Chrome 跳出「已新增○○擴充功能」詢問視窗，頁面也更乾淨
                 "--disable-extensions",
                 # 🔥 即使真的留下 crash 記號，也不要顯示「還原網頁」氣泡擋住畫面
                 "--hide-crash-restore-bubble",
                 LUZ_URL],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print(f"[附著模式] 已用一般方式開啟 Chrome（設定檔：{_luz_profile}）", flush=True)
        except Exception as e:
            print(f"[附著模式] 啟動 Chrome 失敗：{e}，改用一般模式", flush=True)
            return None

        # 等偵錯埠就緒（最多 25 秒）
        for _ in range(50):
            if _luz_port_open(LUZ_DEBUG_PORT):
                break
            _t.sleep(0.5)
        else:
            print("[附著模式] 等待偵錯埠逾時，改用一般模式", flush=True)
            return None

        # 🔥 先取得「通行證」再接手
        #    實測：只要 ChromeDriver 已附著(CDP 連線存在)，且搜尋是用 execute_script 觸發
        #    （非真人點擊），Cloudflare Turnstile 一律驗證失敗 —— 連手動點方塊也過不了。
        #    唯一可靠解法：在「程式尚未接手」的乾淨 Chrome 裡，由使用者自己按一次搜尋、
        #    完成驗證，取得 cf_clearance cookie（存在專用設定檔）。之後程式再附著操作，
        #    帶著 cookie 就不會再被要求驗證。cookie 有效期間內，下次可直接按 Enter 略過。
        if LUZ_ASK_MANUAL_PASS:
            print("", flush=True)
            print("\033[93m════════════ 首次請先取得「通行證」════════════\033[0m", flush=True)
            print("\033[93m 本站按下【搜尋】時會跳 Cloudflare 人機驗證，\033[0m", flush=True)
            print("\033[93m 而程式一旦接手就會被判定為機器人而驗證失敗。\033[0m", flush=True)
            print("\033[93m 請在剛開啟的 Chrome 手動做「一次」查詢：\033[0m", flush=True)
            print("\033[93m   ① 左側選 縣市／鄉鎮市（隨意）\033[0m", flush=True)
            print("\033[93m   ② 按【搜尋】\033[0m", flush=True)
            print("\033[93m   ③ 若跳出人機驗證，把它完成（這次會過）\033[0m", flush=True)
            print("\033[93m 完成後回到本程式，於下方輸入框按 Enter，程式即接手。\033[0m", flush=True)
            print("\033[93m（若先前已取得、通行證仍有效，可直接按 Enter 略過）\033[0m", flush=True)
            print("\033[93m═══════════════════════════════════════════════\033[0m", flush=True)
            try:
                input()
            except (EOFError, OSError):
                print("[附著模式] 讀不到輸入，改為等待 60 秒讓你完成…", flush=True)
                _t.sleep(60)

    try:
        attach_opts = webdriver.ChromeOptions()
        attach_opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{LUZ_DEBUG_PORT}")
        from webdriver_helper import get_chrome_driver_service
        d = webdriver.Chrome(service=get_chrome_driver_service(), options=attach_opts)
        print("[附著模式] ✓ 已成功附著到你的 Chrome，接手自動化", flush=True)
        return d
    except Exception as e:
        print(f"[附著模式] 附著失敗：{e}，改用一般模式", flush=True)
        return None


driver = None
if LUZ_USE_ATTACH_MODE:
    driver = _luz_create_driver_attached()

if driver is None:
    # 退路：原本的一般模式（stealth=True 仍可降低被偵測機率）
    try:
        os.makedirs(_luz_profile, exist_ok=True)
        options.add_argument(f"--user-data-dir={_luz_profile}")
    except Exception:
        pass
    driver = create_chrome_driver(options=options, stealth=True)
    try:
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
        })
    except Exception as _e:
        print(f"[提示] 反偵測腳本注入略過：{_e}", flush=True)

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

# 🔥 使用者在「Chrome縮放設定」面板調過的值優先（zoom_config.json）；沒調過才用上面的 DPI 預設
try:
    _zcp = os.path.join(BASE_DIR, 'zoom_config.json')
    if os.path.exists(_zcp):
        _zov = json.load(open(_zcp, encoding='utf-8')).get('overrides', {}).get('luz')
        if _zov is not None:
            zoom_count = int(_zov)
            print(f"[頁面縮放] 採用使用者設定：縮放 {zoom_count} 次", flush=True)
except Exception:
    pass

apply_page_zoom(driver, zoom_count)  # 共用：拉前景→Ctrl+0→縮放→驗證

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

# 🔥 累積各筆分區文字，存到 4.其他相關/全國土地使用分區資料-<案件>.json（作為備用；高雄仍以高雄都市計畫優先）
_luz_zone_all = []
_luz_zone_json = os.path.join(base_directory, "4.其他相關",
                              f"全國土地使用分區資料-{first_region}{first_section}-{first_lot}.json")

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
        # 🔥 用 JS 直接設值＋觸發事件，避免逐鍵 send_keys 時網站輸入遮罩把「-」吃掉
        #    （實測第 2 筆「2880-4」會被吃成「28804」→ 格式不正確）
        driver.execute_script("""
            var el=arguments[0], v=arguments[1];
            el.focus(); el.value=''; el.value=v;
            el.dispatchEvent(new Event('input',{bubbles:true}));
            el.dispatchEvent(new Event('change',{bubbles:true}));
            el.dispatchEvent(new KeyboardEvent('keyup',{bubbles:true}));
        """, lot_input, str(target_lot))
        time.sleep(0.3)
        _got = (lot_input.get_attribute('value') or '').strip()
        if _got != str(target_lot).strip():
            # 退路：clear + 逐鍵輸入
            try:
                lot_input.clear()
            except Exception:
                pass
            lot_input.send_keys(str(target_lot))
            time.sleep(0.3)
            _got = (lot_input.get_attribute('value') or '').strip()
        if _got == str(target_lot).strip():
            print(f"  已輸入地號: {_got}", flush=True)
        else:
            print(f"\033[91m  ⚠ 地號輸入異常：欄位顯示「{_got}」，目標是「{target_lot}」，請人工確認\033[0m", flush=True)
        lot_ok = True
        time.sleep(0.5)
    except Exception as e:
        print(f"  輸入地號失敗: {e}", flush=True)

    if not lot_ok:
        print(f"\033[91m✗ 第 {idx}/{total} 筆：地號輸入失敗，跳過此筆（不截圖避免誤存舊圖）\033[0m", flush=True)
        return False

    # ========== 步驟7：點擊「搜尋」按鈕 ==========
    # 🔥 本站在「按下搜尋」時觸發 Cloudflare Turnstile。
    #    ★ 2026-07-23 實測結論（兩種都試過）：
    #      ‧ 程式用 execute_script 觸發搜尋 → 驗證彈窗立刻出現，使用者點勾選框【一律失敗】，
    #        且會無限重複要求驗證，永遠不會通過。
    #      ‧ 由使用者【親手點搜尋鈕】觸發 → 第一次可能顯示失敗，但 Turnstile 會自動重試，
    #        約 7~10 秒後【自動通過】。
    #      → 關鍵不在「誰點驗證方塊」，而在「觸發這次請求的是不是真人手勢」。
    #        Turnstile 會把 challenge 綁定到觸發它的那個事件；腳本呼叫沒有真人手勢信任分，
    #        事後再怎麼點勾選框都補不回來。
    #    因此保持：前面欄位由程式填，【搜尋】必須由使用者親手按。
    #    因此改成：前面的縣市/鄉鎮市/段名/地號全部由程式填好，
    #    「按搜尋 + 過驗證」交給使用者親手做（真人點擊才通得過），
    #    完成後按 Enter，程式再接續後面的步驟（不會重跑搜尋）。
    if LUZ_MANUAL_ALL:
        # 🔥 全手動模式：搜尋/查詢/點地號/過驗證/調整畫面全部你自己做，好了按 Enter → 截圖+抓文字
        print("", flush=True)
        print("\033[93m════════ 請自行完成查詢並確認畫面 ════════\033[0m", flush=True)
        print(f"\033[93m 縣市／鄉鎮市／段名／地號({target_lot}) 已由程式填好。請你操作：\033[0m", flush=True)
        print("\033[93m ① 按左側的【定位】(不是搜尋) → 地圖出現紅色錨點；若跳驗證完成它\033[0m", flush=True)
        print("\033[93m    (第一次失敗別關，等幾秒會過；過了再按一次【定位】)\033[0m", flush=True)
        print("\033[93m ② 出現錨點後，按工具列的【查詢】鈕\033[0m", flush=True)
        print("\033[93m ③ 在地圖上『紅色錨點(地號)位置』點一下 → 左邊出現分區資訊；若跳驗證完成它\033[0m", flush=True)
        print("\033[93m ④ 縮回左側面板、把畫面／縮放調整到你要截的樣子\033[0m", flush=True)
        print("\033[93m ★ 全部完成、畫面確定後，回主程式輸入欄按 Enter → 我就截圖並抓分區文字，再做下一筆\033[0m", flush=True)
        print("\033[93m═══════════════════════════════════════════\033[0m", flush=True)
        try:
            input()
        except (EOFError, OSError):
            print("  讀不到輸入，改等待 90 秒讓你完成…", flush=True)
            time.sleep(90)
        print("步驟7：（已由使用者手動查詢並確認畫面）", flush=True)
    elif LUZ_MANUAL_SEARCH:
        print("", flush=True)
        print("\033[93m════════ 請由你手動按下【搜尋】════════\033[0m", flush=True)
        print(f"\033[93m 縣市／鄉鎮市／段名／地號({target_lot}) 已由程式填好。\033[0m", flush=True)
        print("\033[93m ① 到 Chrome 左側面板，親自點一下【搜尋】\033[0m", flush=True)
        print("\033[93m ② 若跳出人機驗證，請完成它\033[0m", flush=True)
        print("\033[93m    第一次會失敗，請等幾秒後再重新點選即可成功！\033[0m", flush=True)
        print("\033[93m    成功後，請自行再點擊【搜尋】一次後，\033[0m", flush=True)
        print("\033[93m    滑鼠移至主程式輸入欄點擊一下後，按 Enter 接續\033[0m", flush=True)
        print("\033[93m════════════════════════════════════\033[0m", flush=True)
        try:
            input()
        except (EOFError, OSError):
            print("  讀不到輸入，改等待 90 秒讓你完成…", flush=True)
            time.sleep(90)
        print("步驟7：（已由使用者手動搜尋並通過驗證）", flush=True)
    else:
        def _do_search(tag=""):
            try:
                driver.execute_script("ZoomToData(8,'CADANO_0101','CADA');")
                print(f"  已執行搜尋{tag}", flush=True)
            except Exception as e:
                print(f"  搜尋失敗: {e}", flush=True)

        print("步驟7：點擊【搜尋】按鈕...", flush=True)
        _do_search()
        time.sleep(2)
        # 🔥 程式觸發的搜尋一樣會跳驗證。偵測到就停手，等使用者點過驗證方塊
        #    （第一次可能顯示失敗，Turnstile 會自動重試，約 7~10 秒後通過）
        if _luz_challenge_present(driver):
            _luz_wait_if_challenge(driver)
            # 網站左側面板明示：「驗證完成後，請重新查詢。」
            # 第一次的搜尋請求已被驗證攔截掉，通過後一定要再送一次才會有結果。
            print("  驗證已通過 → 依網站規則自動重新查詢一次…", flush=True)
            time.sleep(1)
            _do_search("（驗證後重查）")
            time.sleep(2)
            _luz_wait_if_challenge(driver)   # 極少數情況會再驗一次

    # 🔥 全手動模式：使用者已完成查詢並確認畫面(按了 Enter) →
    #    直接「抓分區文字 + 截圖 + 存檔」，跳過所有會弄亂畫面的自動步驟(縮回/放大/平移/倒數)，本筆結束。
    if LUZ_MANUAL_ALL:
        zone = _luz_extract_zone(driver)
        _luz_save_zone(target_city, target_area, target_section, target_lot, zone)
        base_filename = f"07_全國土地使用分區-{target_area}{target_section}-{target_lot}"
        return _luz_shot_and_pdf(driver, base_directory, base_filename)

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
    if LUZ_DO_IDENTIFY:
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
    else:
        print("步驟7.6：手動模式略過（不進查詢模式，避免空的『查詢結果』框出現在截圖）", flush=True)

    # ========== 步驟7.7：用記住的座標點地圖查詢，並「驗證地號」，不對就自動微調重點 ==========
    print("步驟7.7：查地號（帶出分區資訊）...", flush=True)
    try:
        if not LUZ_DO_IDENTIFY:
            print("  設定為略過查地號：搜尋後圖釘已標在地號、分區顏色也都在，截圖已足夠（無資訊框）。", flush=True)
        elif LUZ_MANUAL_SEARCH:
            # 🔥 程式點地圖查地號(identify)會觸發隱形 Cloudflare 挑戰、直接失敗(沒有可點的方塊)；
            #    必須由「你的真人點擊」觸發才會通過(跟搜尋同理)。已幫你進入查詢模式，請親手點圖釘。
            print("", flush=True)
            print("\033[93m════════ 請點一下地圖上的『紅色圖釘』查地號 ════════\033[0m", flush=True)
            print("\033[93m 已幫你按好【查詢】鈕（進入查詢模式）。\033[0m", flush=True)
            print("\033[93m 請在地圖上『紅色圖釘』的位置點一下 → 左邊會出現該地號的分區資訊；\033[0m", flush=True)
            print("\033[93m 若跳出人機驗證，請完成它（第一次失敗別關，等幾秒會自動通過）。\033[0m", flush=True)
            print("\033[93m 資訊出現後，回主程式輸入欄按 Enter 截圖。\033[0m", flush=True)
            print("\033[93m（不需要地號資訊、只要地圖，也可直接按 Enter）\033[0m", flush=True)
            print("\033[93m══════════════════════════════════════════\033[0m", flush=True)
            try:
                input()
            except (EOFError, OSError):
                print("  讀不到輸入，改等待 40 秒讓你點圖釘…", flush=True)
                time.sleep(40)
            print("步驟7.7：（已由使用者手動點圖釘查地號）", flush=True)
        elif not _pin_xy:
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

            def _run_candidates():
                """逐點點地圖查地號；查到目標回 True。點地圖(identify)也會觸發 Cloudflare，
                   偵測到就暫停等使用者完成驗證，再重點一次。"""
                for dx, dy in candidates:
                    cx, cy = bx + dx, by + dy
                    driver.execute_script(_CLICK_JS, cx, cy)
                    lot = None
                    for _ in range(6):  # 等查詢結果，最多約 3 秒
                        time.sleep(0.5)
                        # 🔥 查地號這步同樣可能跳 Cloudflare → 偵測到就暫停等你過，再重點一次
                        if _luz_challenge_present(driver):
                            _luz_wait_if_challenge(driver)
                            driver.execute_script(_CLICK_JS, cx, cy)
                        lot = driver.execute_script(_READ_LOT_JS)
                        if lot:
                            break
                    print(f"  試點 @({cx},{cy}) → 查到地號={lot}", flush=True)
                    if lot and str(lot).strip() == target:
                        print(f"  ✓ 查詢結果正確：地號 {lot}", flush=True)
                        return True
                return False

            matched = _run_candidates()
            # 🔥 若整輪都 None、且畫面上有驗證框（第一次點就跳出、把 identify 全擋掉）→ 等你過驗證後再整輪重試一次
            if not matched and _luz_challenge_present(driver):
                _luz_wait_if_challenge(driver)
                matched = _run_candidates()
            if not matched:
                print(f"  ⚠ 試了 {len(candidates)} 個位置仍未對到目標地號 {target}（地號可能太小，請人工確認）", flush=True)
    except Exception as e:
        print(f"  查詢點擊失敗: {e}", flush=True)

    # ========== 步驟7.8～7.10：重新顯示錨點(pin)，讓最終截圖更明確 ==========
    #   查詢結果確認後，再展開選單→重新搜尋(pin 重現)→收回選單
    # 🔥 手動搜尋模式下略過：7.9 會再跑一次 ZoomToData（等於程式自動搜尋），
    #    那正是會觸發 Cloudflare 驗證、且程式觸發必失敗的動作，還可能讓驗證框擋住截圖。
    #    少了 pin 只是標示沒那麼醒目，查詢結果本身仍完整。
    try:
        if LUZ_MANUAL_SEARCH:
            print("步驟7.8～7.10：手動搜尋模式，略過『重新搜尋讓錨點重現』（避免再觸發驗證）", flush=True)
        else:
            print("步驟7.8：重新展開系統功能選單...", flush=True)
            driver.execute_script("loadmenu();")
            time.sleep(1)
            print("步驟7.9：再次搜尋（讓錨點重新出現）...", flush=True)
            driver.execute_script("ZoomToData(8,'CADANO_0101','CADA');")
            time.sleep(2)
            # 🔥 再次搜尋同樣可能觸發 Cloudflare 驗證 → 一樣先停手等它過
            _luz_wait_if_challenge(driver)
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
