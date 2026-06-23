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
os.system('cls')
print("歡迎使用【國土測繪】自動化小程式，模組載入中...", flush=True)
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
nlsc_window_width = 1024
nlsc_window_height = 1024
nlsc_window_x = 0
nlsc_window_y = 0

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

            nlsc_window_width = chrome_width_logical
            nlsc_window_height = chrome_height_logical
            nlsc_window_x = chrome_x_logical
            nlsc_window_y = chrome_y_logical
except Exception as e:
    pass

# 讀取 data.json 文件
with open(get_data_json_path(), 'r', encoding='utf-8') as file:
    data_list = json.load(file)

# 🔥 子批次模式：父程序透過 --batch-lots 指定本次 subprocess 只處理哪幾個地號
import argparse as _argparse
_arg_parser = _argparse.ArgumentParser(add_help=False)
_arg_parser.add_argument('--batch-lots', default=None,
                         help='comma-separated lot numbers (subprocess mode)')
_args, _ = _arg_parser.parse_known_args()

_is_child_batch = bool(_args.batch_lots)
if _is_child_batch:
    _wanted = {l.strip() for l in _args.batch_lots.split(',') if l.strip()}
    data_list = [d for d in data_list if d['lot_number'] in _wanted]
    print(f"[子批次] 限定處理 {len(data_list)} 筆: {sorted(_wanted)}", flush=True)


def _lot_main_int(lot_str):
    """從 '859' / '859-2' / '0859-0000' 取出母號 int(859)"""
    s = str(lot_str or '').strip().replace('地號', '')
    try:
        a = s.split('-', 1)[0]
        return int(a.lstrip('0') or '0')
    except (ValueError, IndexError):
        return 0


def _suggest_batching(data_list):
    """依原本順序，在相鄰兩筆母號差距最大處插入「/」。
    例：原順序 [859, 887, 446] → 相鄰差 (859,887)=28, (887,446)=441
        最大差在 887→446 之間 → '859,887/446'（保持原順序）
    """
    lots = [d['lot_number'] for d in data_list]
    if len(lots) < 2:
        return ','.join(lots)
    mains = [_lot_main_int(l) for l in lots]
    # 計算相鄰兩筆的母號絕對差（依原順序）
    gaps = [(abs(mains[i + 1] - mains[i]), i) for i in range(len(mains) - 1)]
    gaps.sort(reverse=True)
    cut_idx = gaps[0][1]  # 在 cut_idx 與 cut_idx+1 之間插入 '/'
    left = lots[:cut_idx + 1]
    right = lots[cut_idx + 1:]
    return ','.join(left) + '/' + ','.join(right)


def ask_user_for_batching(data_list):
    """≥3 筆地號時彈窗詢問是否分批。回傳 list of batch (each batch is sub-list of data_list)"""
    if len(data_list) < 3:
        return [data_list]

    import tkinter as tk

    all_lots = [d['lot_number'] for d in data_list]
    suggested = _suggest_batching(data_list)

    result = {"batches": [data_list]}  # 預設不分批

    win = tk.Tk()
    win.title("分批執行設定")
    # 🔥 加大視窗以容納加大的字體
    win.geometry("720x440")
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f"720x440+{(sw - 720) // 2}+{(sh - 440) // 2}")

    tk.Label(win, text=f"偵測到 {len(data_list)} 筆地號", font=("微軟正黑體", 17, "bold"),
             fg="#1976D2").pack(pady=(18, 6))
    tk.Label(win, text=', '.join(all_lots), font=("微軟正黑體", 16, "bold"),
             fg="#333").pack()

    tk.Label(win, text="\n如需分批執行（避免地號分散使截圖看不清），\n請輸入分組，用 / 分隔不同批次，用 , 分隔同批內的地號：",
             font=("微軟正黑體", 13), justify=tk.CENTER).pack()
    tk.Label(win, text=f"系統建議（依地號數字距離）：{suggested}",
             font=("微軟正黑體", 13), fg="#888").pack(pady=(8, 0))

    entry = tk.Entry(win, font=("微軟正黑體", 16, "bold"), width=40, justify='center')
    entry.insert(0, suggested)
    entry.pack(pady=14)

    def _parse_and_run_batches():
        text = entry.get().strip()
        batches = []
        seen = set()
        for group in text.split('/'):
            wanted = [l.strip() for l in group.split(',') if l.strip()]
            batch_data = [d for d in data_list if d['lot_number'] in wanted and d['lot_number'] not in seen]
            for d in batch_data:
                seen.add(d['lot_number'])
            if batch_data:
                batches.append(batch_data)
        result["batches"] = batches if batches else [data_list]
        win.destroy()

    def _run_all_together():
        result["batches"] = [data_list]
        win.destroy()

    btns = tk.Frame(win)
    btns.pack(pady=18)
    tk.Button(btns, text="分批執行", command=_parse_and_run_batches,
              font=("微軟正黑體", 14, "bold"), bg="#FF9800", fg="white",
              padx=30, pady=10, cursor="hand2").pack(side=tk.LEFT, padx=12)
    tk.Button(btns, text="不分批（全部一起截圖）", command=_run_all_together,
              font=("微軟正黑體", 14, "bold"), bg="#4CAF50", fg="white",
              padx=30, pady=10, cursor="hand2").pack(side=tk.LEFT, padx=12)

    win.protocol("WM_DELETE_WINDOW", _run_all_together)  # 關閉視窗 = 不分批
    win.mainloop()

    return result["batches"]


# 🔥 詢問分批執行（≥3 筆才會彈窗；子批次模式跳過對話框）
if _is_child_batch:
    _batches = [data_list]
    print(f"[子批次模式] 跳過分批對話框", flush=True)
else:
    _batches = ask_user_for_batching(data_list)
    if len(_batches) > 1:
        print(f"[分批] 共 {len(_batches)} 批：", flush=True)
        for _bi, _b in enumerate(_batches, 1):
            print(f"   批次 {_bi}: {', '.join(d['lot_number'] for d in _b)}", flush=True)
    else:
        print(f"[執行] 共 {len(data_list)} 筆地號，未分批", flush=True)

# 🔥 父程序 + 多批次：以 subprocess 逐批執行，最後合併 PDF
if not _is_child_batch and len(_batches) > 1:
    import subprocess as _subprocess

    # 預先計算工作目錄（與單批次邏輯一致）
    _first_entry = data_list[0]
    _region = _first_entry['area']
    _section = _first_entry['section']
    _all_lots_for_section = [
        entry['lot_number'] for entry in data_list
        if entry['area'] == _region and entry['section'] == _section
    ]
    _base_dir = get_work_folder(f"{_region}{_section}-{_all_lots_for_section[0]}")
    _png_dir = os.path.join(_base_dir, "1.基本資料", "png")
    os.makedirs(_png_dir, exist_ok=True)

    # 子程序 unbuffered 環境
    _child_env = os.environ.copy()
    _child_env['PYTHONUNBUFFERED'] = '1'
    _child_env['PYTHONIOENCODING'] = 'utf-8'

    # 逐批次啟動 subprocess（用 Popen + 主動 readline 即時轉發，避免 Windows pipe buffer 卡住）
    _batch_pngs = []
    for _bi, _b in enumerate(_batches, 1):
        _lots_str = ','.join(d['lot_number'] for d in _b)
        print(f"\n========== 啟動批次 {_bi}/{len(_batches)}: {_lots_str} ==========", flush=True)
        _cmd = [sys.executable, '-u', __file__, '--batch-lots', _lots_str]
        _proc = _subprocess.Popen(
            _cmd, env=_child_env,
            stdout=_subprocess.PIPE,
            stderr=_subprocess.STDOUT,
            bufsize=1,  # line-buffered
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        # 即時轉發子程序輸出（每讀到一行就立刻 flush 給 main.py）
        for _line in iter(_proc.stdout.readline, ''):
            print(_line, end='', flush=True)
        _proc.stdout.close()
        _proc.wait()
        _returncode = _proc.returncode
        if _returncode != 0:
            print(f"[警告] 批次 {_bi} 子程序回傳非零: {_returncode}", flush=True)

        # 子批次內部命名邏輯：06_國土測繪-{region}{section}-{lot1.lot2...}.png
        _batch_lot_list = [
            d['lot_number'] for d in _b
            if d['area'] == _region and d['section'] == _section
        ]
        _batch_png_name = f"06_國土測繪-{_region}{_section}-" + '.'.join(_batch_lot_list) + ".png"
        _batch_png_path = os.path.join(_png_dir, _batch_png_name)
        if os.path.exists(_batch_png_path):
            _batch_pngs.append(_batch_png_path)
            print(f"   ✓ 批次 {_bi} 截圖: {_batch_png_name}", flush=True)
        else:
            print(f"   ⚠ 批次 {_bi} 未找到預期截圖：{_batch_png_path}", flush=True)

    # 將所有批次 PNG 合併成單一 PDF
    if _batch_pngs:
        _combined_pdf_name = f"06_國土測繪-{_region}{_section}-" + '.'.join(_all_lots_for_section) + ".pdf"
        _combined_pdf_path = os.path.join(_base_dir, "1.基本資料", _combined_pdf_name)
        try:
            from fpdf import FPDF as _FPDF
            _pdf = _FPDF(orientation='L', unit='mm', format='A4')
            for _png in _batch_pngs:
                _pdf.add_page()
                _pdf.image(_png, x=0, y=0, w=297, h=210)
            _pdf.output(_combined_pdf_path)
            print(f"\n\033[93m✓ 合併 PDF 已儲存: {_combined_pdf_path}\033[0m", flush=True)
        except Exception as _e:
            print(f"[錯誤] 合併 PDF 失敗: {_e}", flush=True)
    else:
        print(f"[錯誤] 沒有任何批次截圖可合併", flush=True)

    print("\n[父程序] 所有批次處理完畢", flush=True)
    print("國土測繪圖資已完成執行", flush=True)  # 通知主程式（與 notify_main_program 一致）
    sys.exit(0)

# 設定 WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36")
# 🔥 不在啟動參數設定視窗大小，改為啟動後再調整（避免 DPI 改變時崩潰）
# options.add_argument("--window-size=1024,1024")
# options.add_argument("--window-position=0,0")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
# 🔥 預先「允許」地理位置（避免網站在被拒絕時呼叫 alert("使用者拒絕提供位置資訊") 卡住 Selenium）
# 通知改成拒絕，避免桌面通知干擾
# 1 = allow, 2 = block
options.add_experimental_option("prefs", {
    "profile.default_content_setting_values.geolocation": 1,    # 1 = allow
    "profile.default_content_setting_values.notifications": 2,
})
# 🔥 保險：仍設 unhandledPromptBehavior，萬一還有其他 alert 也不會卡住
options.set_capability("unhandledPromptBehavior", "dismiss")

# 使用 ChromeDriverManager 安裝 ChromeDriver 並應用選項
driver = create_chrome_driver(options=options)

# 🔥 啟動後立即設定視窗大小和位置（避免在啟動參數設定導致 DPI 改變時崩潰）
try:
    driver.set_window_size(nlsc_window_width, nlsc_window_height)
    driver.set_window_position(nlsc_window_x, nlsc_window_y)
except Exception as e:
    print(f"[WARNING] 視窗設定失敗: {e}，繼續執行")

# 🔥 用 CDP 直接授予地理位置權限（比 prefs 可靠；新版 Chrome 的 prefs 常失效）
try:
    driver.execute_cdp_cmd("Browser.grantPermissions", {"permissions": ["geolocation"]})
    print("[權限] 已自動授予地理位置權限", flush=True)
except Exception as _perm_e:
    print(f"[權限] 自動授予地理位置失敗（不影響流程）: {_perm_e}", flush=True)

# 指定要打開的網址
# 🔥 ?In_type=web 強制使用 PC 桌面版（沒帶這參數時會依視窗寬度退成窄版，缺 addIndexPage 等桌面版函式而報錯）
url = "https://maps.nlsc.gov.tw/T09/mapshow.action?In_type=web"

# 🔥 保險提醒：萬一仍跳出地理位置權限窗，告訴使用者怎麼點
print("\033[38;5;208m※ 若瀏覽器左上角跳出『存取您的位置資訊』，請點最上面的【造訪這個網站時允許】即可繼續。\033[0m", flush=True)
# 使用系統預設瀏覽器打開網址
driver.get(url)
# 🔥 重新整理一次才會套用 PC 桌面版：In_type=web 在「第一次載入」才寫進 session，
#    要「第二次載入」才讀得到（等同手動按 F5）。否則首次仍是窄版、缺 addIndexPage。
try:
    print("[版面] 首次載入完成，2 秒後自動重新整理（套用 PC 桌面版）...", flush=True)
    time.sleep(2)            # 等第一次載入把 In_type=web 設定寫進去
    driver.refresh()         # 等於按 F5 → 這次才是桌面版
    time.sleep(2)
    print("[版面] ✓ 已重新整理，現在是 PC 桌面版", flush=True)
except Exception as _e:
    print(f"[版面] 重新整理時發生問題（不影響流程）：{_e}", flush=True)
# 🔥 網頁載入後重新設定視窗大小（避免被網站重置）
try:
    driver.set_window_size(nlsc_window_width, nlsc_window_height)
    driver.set_window_position(nlsc_window_x, nlsc_window_y)
except:
    pass

# 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
verify_and_fix_chrome_window(driver)

# 等待元素加載 (可選)
driver.implicitly_wait(3)  # 等待 3 秒

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

# 🔥 使用者在「Chrome縮放設定」面板調過的值優先（zoom_config.json）；沒調過才用上面的 DPI 預設
try:
    _zcp = os.path.join(BASE_DIR, 'zoom_config.json')
    if os.path.exists(_zcp):
        _zov = json.load(open(_zcp, encoding='utf-8')).get('overrides', {}).get('nlscmaps')
        if _zov is not None:
            zoom_count = int(_zov)
            print(f"[頁面縮放] 採用使用者設定：縮放 {zoom_count} 次", flush=True)
except Exception:
    pass

def _force_chrome_foreground(driver):
    """用 AttachThreadInput 解除 Windows 前景鎖，把本支 Chrome 拉到前景，
    keyboard 的 Ctrl+- 才送得進去（背景程式直接 SetForegroundWindow 會被擋）。"""
    try:
        driver.switch_to.window(driver.current_window_handle)
    except Exception:
        pass
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        user32.GetWindowThreadProcessId.restype = wintypes.DWORD
        marker = "NLSCZOOMWIN8741"
        old = driver.execute_script("var t=document.title;document.title=arguments[0];return t;", marker)
        time.sleep(0.25)
        found = []
        WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def _cb(hwnd, lparam):
            n = user32.GetWindowTextLengthW(hwnd)
            if n:
                buf = ctypes.create_unicode_buffer(n + 1)
                user32.GetWindowTextW(hwnd, buf, n + 1)
                if marker in buf.value:
                    found.append(hwnd)
            return True
        user32.EnumWindows(WNDENUMPROC(_cb), 0)
        try:
            driver.execute_script("document.title=arguments[0];", old)
        except Exception:
            pass
        if not found:
            print("[頁面縮放] 找不到 Chrome 視窗（無法拉前景）", flush=True)
            return False
        hwnd = found[0]
        fg = user32.GetForegroundWindow()
        fg_thread = user32.GetWindowThreadProcessId(fg, None)
        cur_thread = kernel32.GetCurrentThreadId()
        user32.AttachThreadInput(cur_thread, fg_thread, True)
        user32.ShowWindow(hwnd, 9)        # SW_RESTORE
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
        user32.AttachThreadInput(cur_thread, fg_thread, False)
        time.sleep(0.3)
        return True
    except Exception as e:
        print(f"[頁面縮放] 拉前景失敗：{e}", flush=True)
        return False

apply_page_zoom(driver, zoom_count)  # 共用：拉前景→Ctrl+0→縮放→驗證→還原前景（避免打字跑到Chrome）

# 🔥 關閉「提醒您」彈窗（本網站內容僅供參考…）
# 這個彈窗是頁面載入後才彈出來的，加上 JS 主動 polling 等它出現
def _close_alert_dialog(max_wait_sec=10):
    """輪詢等待「提醒您」對話框出現，然後點「確定」關掉。
    用 JS 掃所有 div[role="dialog"]，相容於 jQuery UI dialog 結構。
    """
    import time as _t
    deadline = _t.time() + max_wait_sec
    attempt = 0
    while _t.time() < deadline:
        attempt += 1
        try:
            clicked = driver.execute_script("""
                var dialogs = document.querySelectorAll('div[role="dialog"]');
                for (var i = 0; i < dialogs.length; i++) {
                    var d = dialogs[i];
                    // 只處理「顯示中」的對話框
                    if (d.offsetParent === null) continue;
                    var btns = d.querySelectorAll('button');
                    for (var j = 0; j < btns.length; j++) {
                        var t = (btns[j].textContent || '').trim();
                        if (t === '確定' || t.indexOf('確定') !== -1) {
                            btns[j].click();
                            return true;
                        }
                    }
                }
                return false;
            """)
            if clicked:
                print(f"關閉提醒視窗（第 {attempt} 次輪詢成功）", flush=True)
                return True
        except Exception:
            # 例如 UnexpectedAlertPresentException —— 等 unhandledPromptBehavior 自動 dismiss
            pass
        _t.sleep(0.5)

    print("提醒視窗不存在或在等待期間未出現", flush=True)
    return False

_close_alert_dialog()

driver.implicitly_wait(3)

# 🔥 保險：如果有 native alert（例如「使用者拒絕提供位置資訊」），先 dismiss 掉
def _dismiss_alert_if_present(_driver):
    try:
        _alert = _driver.switch_to.alert
        _txt = _alert.text
        _alert.dismiss()
        print(f"已關閉瀏覽器 alert：{_txt}", flush=True)
    except NoAlertPresentException:
        pass
    except Exception:
        pass

_dismiss_alert_if_present(driver)

# 點擊「教學視窗關閉」按鈕
try:
    # 使用顯式等待按鈕出現，最多等待 3 秒
    close_button = WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@class='ui-button ui-corner-all ui-widget ui-button-icon-only ui-dialog-titlebar-close' and @title='Close']"))
    )
    close_button.click()
    print("關閉系統教學視窗", flush=True)
except TimeoutException:
    print("系統教學視窗沒有出現或無法點擊", flush=True)
except NoSuchElementException:
    print("找不到系統教學視窗", flush=True)
except UnexpectedAlertPresentException:
    _dismiss_alert_if_present(driver)
    print("已 dismiss alert 後跳過教學視窗等待", flush=True)

# 等待頁面 JavaScript 完全載入
# 連續模式下需要更長的等待時間
time.sleep(3)

try:
    # 等待 jQuery 和相關函式載入完成
    WebDriverWait(driver, 10).until(
        lambda d: d.execute_script("return typeof jQuery !== 'undefined' && typeof folder === 'function'")
    )

    # 先檢查 JavaScript 函數是否存在，然後執行
    driver.execute_script("""
        if (typeof folder === 'function' && typeof toggleControl === 'function') {
            folder('adg','CollapsiblePanel5');
            toggleControl('none');
            return true;
        } else {
            return false;
        }
    """)
    print("點擊【圖層設定】", flush=True)
except TimeoutException:
    print("【圖層設定】未出現或無法點擊", flush=True)
except Exception as e:
    error_msg = str(e)
    # 過濾掉網頁自身的 JavaScript 錯誤
    if "closeChatBox" in error_msg or "DEPRECATED_ENDPOINT" in error_msg:
        print("【圖層設定】偵測到網頁自身的 JavaScript 錯誤（可忽略）", flush=True)
    else:
        print(f"【圖層設定】JavaScript 執行失敗: {e}", flush=True)
    # print("嘗試使用其他方法點擊圖層設定...", flush=True)

    # # 備用方案：直接尋找並點擊圖層設定按鈕
    # try:
    #     layer_button = WebDriverWait(driver, 5).until(
    #         EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '圖層設定')]"))
    #     )
    #     driver.execute_script("arguments[0].click();", layer_button)
    #     print("使用備用方案點擊【圖層設定】成功", flush=True)
    # except Exception as backup_e:
    #     print("備用方案也失敗,繼續執行後續步驟", flush=True)

time.sleep(1)
try:
    # 使用 JavaScript 直接執行 onclick 事件
    driver.execute_script("addIndexPage('layerlist','/pro/layerlist.html');$('#layerlist_dialog').dialog('option', 'title', '圖層列表');$( '#layerlist_dialog' ).dialog('open');")
    print("點擊【圖層列表】", flush=True)
except TimeoutException:
    print("【圖層列表】未出現或無法點擊", flush=True)

time.sleep(1)
try:
    # 使用 JavaScript 直接執行 onclick 事件
    driver.execute_script("addMapsPage('002','layerlist','/pro/002_layer.html');")
    print("點擊【土地圖層】", flush=True)
except TimeoutException:
    print("【土地圖層】未出現或無法點擊", flush=True)

time.sleep(1)
try:
    # 使用 JavaScript 直接執行 onclick 事件
    driver.execute_script("chk_layertype('土地圖層',URBAN,'URBAN');")
    print("點擊【都市計畫】", flush=True)
except TimeoutException:
    print("【都市計畫】未出現或無法點擊", flush=True)

time.sleep(1)
try:
    # 使用 JavaScript 直接執行 onclick 事件
    driver.execute_script("chk_layertype('土地圖層',DMAPS,'DMAPS');")
    print("點擊【地籍圖】", flush=True)
except TimeoutException:
    print("【地籍圖】未出現或無法點擊", flush=True)

time.sleep(1)
# 處理地籍圖使用資訊對話框
try:
    # 通用的對話框定位，尋找標題包含"地籍圖使用資訊"的對話框
    dialog_titles = WebDriverWait(driver, 5).until(
        EC.presence_of_all_elements_located((By.XPATH, "//span[@class='ui-dialog-title']"))
    )
    
    found_dialog = False
    for title in dialog_titles:
        if "地籍圖使用資訊" in title.text:
            print(f"找到地籍圖使用資訊對話框: {title.text}", flush=True)
            found_dialog = True
            dialog = title.find_element(By.XPATH, "./ancestor::div[@role='dialog']")
            
            # 在對話框中尋找按鈕
            buttons = dialog.find_elements(By.XPATH, ".//button")
            print(f"在對話框中找到 {len(buttons)} 個按鈕", flush=True)
            
            # 尋找"確定"按鈕
            confirm_button = None
            for button in buttons:
                if button.text == "確定" or button.get_attribute("textContent") == "確定":
                    confirm_button = button
                    break
            
            # 如果找到確定按鈕，嘗試點擊
            if confirm_button:
                try:
                    driver.execute_script("arguments[0].scrollIntoView(true);", confirm_button)
                    driver.execute_script("arguments[0].click();", confirm_button)
                    print("成功點擊【確定】按鈕", flush=True)
                except Exception as e:
                    print(f"點擊【確定】按鈕時發生錯誤: {e}", flush=True)
            else:
                # 如果找不到確定按鈕，嘗試點擊第一個按鈕
                if buttons:
                    try:
                        driver.execute_script("arguments[0].scrollIntoView(true);", buttons[0])
                        driver.execute_script("arguments[0].click();", buttons[0])
                        print("點擊了對話框的第一個按鈕", flush=True)
                    except Exception as e:
                        print(f"點擊第一個按鈕時發生錯誤: {e}", flush=True)
                        
                        # 最後嘗試：直接使用 JavaScript 點擊任何可能的確定按鈕
                        try:
                            driver.execute_script("""
                                var dialogs = document.querySelectorAll('div[role="dialog"]');
                                for(var i=0; i<dialogs.length; i++) {
                                    if(dialogs[i].querySelector('.ui-dialog-title').textContent.includes('地籍圖使用資訊')) {
                                        var buttons = dialogs[i].querySelectorAll('button');
                                        for(var j=0; j<buttons.length; j++) {
                                            if(buttons[j].textContent.includes('確定')) {
                                                buttons[j].click();
                                                return;
                                            }
                                        }
                                        // 如果沒有找到確定按鈕，嘗試點擊第一個按鈕
                                        if(buttons.length > 0) {
                                            buttons[0].click();
                                            return;
                                        }
                                    }
                                }
                            """)
                            print("使用 JavaScript 嘗試點擊按鈕", flush=True)
                        except Exception as js_e:
                            print(f"JavaScript 點擊失敗: {js_e}", flush=True)
    
    if not found_dialog:
        print("沒有找到標題為【地籍圖使用資訊】的對話框", flush=True)
        
        # 嘗試直接通過 JavaScript 找到並點擊任何對話框中的按鈕
        try:
            result = driver.execute_script("""
                var result = { found: false, buttons: 0 };
                var dialogs = document.querySelectorAll('div[role="dialog"]');
                result.dialogs = dialogs.length;
                
                for(var i=0; i<dialogs.length; i++) {
                    var buttons = dialogs[i].querySelectorAll('button');
                    result.buttons += buttons.length;
                    
                    if(buttons.length > 0) {
                        result.found = true;
                        // 尋找確定按鈕
                        for(var j=0; j<buttons.length; j++) {
                            if(buttons[j].textContent.includes('確定')) {
                                buttons[j].click();
                                return result;
                            }
                        }
                        // 如果沒找到確定按鈕，點擊第一個按鈕
                        buttons[0].click();
                    }
                }
                return result;
            """)
            print(f"JavaScript 執行結果: 找到對話框: {result.get('found', False)}, 對話框數量: {result.get('dialogs', 0)}, 按鈕數量: {result.get('buttons', 0)}", flush=True)
        except Exception as js_e:
            print(f"JavaScript 操作失敗: {js_e}", flush=True)

except TimeoutException:
    print("未找到任何對話框，可能已經處理完畢", flush=True)
except NoSuchElementException:
    print("未找到所需元素", flush=True)
except Exception as e:
    print(f"處理地籍圖使用資訊對話框時發生錯誤: {e}", flush=True)
    
    # 最後嘗試：使用通用方法處理任何對話框
    try:
        driver.execute_script("""
            var dialogs = document.querySelectorAll('div[role="dialog"]');
            for(var i=0; i<dialogs.length; i++) {
                var buttons = dialogs[i].querySelectorAll('button');
                for(var j=0; j<buttons.length; j++) {
                    if(buttons[j].textContent.includes('確定')) {
                        buttons[j].click();
                        return;
                    }
                }
                if(buttons.length > 0) {
                    buttons[0].click();
                }
            }
        """)
        print("已嘗試使用 JavaScript 關閉任何對話框", flush=True)
    except Exception as js_error:
        print(f"JavaScript 關閉對話框失敗: {js_error}", flush=True)

# 確認對話框是否已關閉
time.sleep(1)
try:
    dialog_still_exists = driver.find_element(By.XPATH, "//span[contains(text(), '地籍圖使用資訊')]")
    print("警告：對話框似乎仍然存在，可能需要手動操作", flush=True)
except:
    print("確認：地籍圖使用資訊對話框已成功關閉", flush=True)

# 在點擊非都市使用地類別圖後，新增處理滿意度調查對話框的代碼
# 插入在以下代碼之後:
# driver.execute_script("chk_layertype('土地圖層',nURBAN2,'nURBAN2');")
# print("點擊【非都市使用地類別圖】", flush=True)
# time.sleep(1)

# 處理服務滿意度調查對話框
try:
    print("檢查是否出現服務滿意度調查對話框", flush=True)
    
    # 嘗試找到滿意度調查對話框 (通過標題文字識別)
    survey_dialog_title = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'ui-dialog-title') and contains(text(), '服務滿意度調查')]"))
    )
    
    if survey_dialog_title:
        print("找到服務滿意度調查對話框", flush=True)
        
        # 找到關閉按鈕 (通過標題欄中的關閉按鈕)
        close_button = survey_dialog_title.find_element(By.XPATH, "./parent::div/button[@title='Close']")
        
        # 使用JavaScript點擊關閉按鈕
        driver.execute_script("arguments[0].click();", close_button)
        print("已關閉服務滿意度調查對話框", flush=True)
        
except TimeoutException:
    print("未出現服務滿意度調查對話框或已自動關閉", flush=True)
except NoSuchElementException:
    print("找不到服務滿意度調查對話框的關閉按鈕", flush=True)
    
    # 退階方案：使用通用JavaScript方法關閉任何對話框
    try:
        driver.execute_script("""
            var dialogs = document.querySelectorAll('div[role="dialog"]');
            for (var i = 0; i < dialogs.length; i++) {
                var title = dialogs[i].querySelector('.ui-dialog-title');
                if (title && title.textContent.includes('服務滿意度調查')) {
                    var closeBtn = dialogs[i].querySelector('button[title="Close"]');
                    if (closeBtn) {
                        closeBtn.click();
                        return true;
                    }
                }
            }
            return false;
        """)
        print("嘗試使用JavaScript關閉服務滿意度調查對話框", flush=True)
    except Exception as js_e:
        print(f"JavaScript關閉對話框失敗: {js_e}", flush=True)
except Exception as e:
    print(f"處理服務滿意度調查對話框時發生錯誤: {e}", flush=True)

# 如果對話框沒有正確關閉，最後的解決方案是等待它自動關閉 (題目中提到 5 秒後關閉)
# time.sleep(5)
# try:
#     # 使用 JavaScript 直接執行 onclick 事件
#     driver.execute_script("chk_layertype('土地圖層',nURBAN,'nURBAN');")
#     print("點擊【非都市使用分區圖】", flush=True)
# except TimeoutException:
#     print("【非都市使用分區圖】未出現或無法點擊", flush=True)

time.sleep(1)
# 處理服務滿意度調查對話框
try:
    print("檢查是否出現服務滿意度調查對話框", flush=True)
    
    # 嘗試找到滿意度調查對話框 (通過標題文字識別)
    survey_dialog_title = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'ui-dialog-title') and contains(text(), '服務滿意度調查')]"))
    )
    
    if survey_dialog_title:
        print("找到服務滿意度調查對話框", flush=True)
        
        # 找到關閉按鈕 (通過標題欄中的關閉按鈕)
        close_button = survey_dialog_title.find_element(By.XPATH, "./parent::div/button[@title='Close']")
        
        # 使用JavaScript點擊關閉按鈕
        driver.execute_script("arguments[0].click();", close_button)
        print("已關閉服務滿意度調查對話框", flush=True)
        
except TimeoutException:
    print("未出現服務滿意度調查對話框或已自動關閉", flush=True)
except NoSuchElementException:
    print("找不到服務滿意度調查對話框的關閉按鈕", flush=True)
    
    # 退階方案：使用通用JavaScript方法關閉任何對話框
    try:
        driver.execute_script("""
            var dialogs = document.querySelectorAll('div[role="dialog"]');
            for (var i = 0; i < dialogs.length; i++) {
                var title = dialogs[i].querySelector('.ui-dialog-title');
                if (title && title.textContent.includes('服務滿意度調查')) {
                    var closeBtn = dialogs[i].querySelector('button[title="Close"]');
                    if (closeBtn) {
                        closeBtn.click();
                        return true;
                    }
                }
            }
            return false;
        """)
        print("嘗試使用JavaScript關閉服務滿意度調查對話框", flush=True)
    except Exception as js_e:
        print(f"JavaScript關閉對話框失敗: {js_e}", flush=True)
except Exception as e:
    print(f"處理服務滿意度調查對話框時發生錯誤: {e}", flush=True)

# 如果對話框沒有正確關閉，最後的解決方案是等待它自動關閉 (題目中提到 5 秒後關閉)
time.sleep(1)

# try:
#     # 使用 JavaScript 直接執行 onclick 事件
#     driver.execute_script("chk_layertype('土地圖層',nURBAN2,'nURBAN2');")
#     print("點擊【非都市使用地類別圖】", flush=True)
# except TimeoutException:
#     print("【非都市使用地類別圖】未出現或無法點擊", flush=True)

# time.sleep(1)

try:
    close_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[22]/div[1]/button/span[1]"))
    )

    driver.execute_script("arguments[0].scrollIntoView();", close_button)
    driver.execute_script("arguments[0].click();", close_button)
    print("點擊【土地圖層】關閉按鈕", flush=True)
except Exception as e:
    print(f"出現錯誤: {e}", flush=True)

# 等待頁面狀態穩定，連續模式下需要更多時間
time.sleep(3)

try:
    # 等待 jQuery 和相關函式載入完成
    WebDriverWait(driver, 10).until(
        lambda d: d.execute_script("return typeof jQuery !== 'undefined' && typeof folder === 'function'")
    )

    # 先檢查 JavaScript 函數是否存在，然後執行
    driver.execute_script("""
        if (typeof folder === 'function' && typeof toggleControl === 'function') {
            folder('adg','CollapsiblePanel1');
            toggleControl('none');
            return true;
        } else {
            return false;
        }
    """)
    print("點擊【定位查詢】", flush=True)

except TimeoutException:
    print("【定位查詢】未出現或無法點擊", flush=True)
except Exception as e:
    error_msg = str(e)
    # 過濾掉網頁自身的 JavaScript 錯誤
    if "closeChatBox" in error_msg or "DEPRECATED_ENDPOINT" in error_msg:
        print("【定位查詢】偵測到網頁自身的 JavaScript 錯誤(可忽略)", flush=True)
    else:
        print(f"【定位查詢】JavaScript 執行失敗: {e}", flush=True)
    print("嘗試使用其他方法點擊定位查詢...", flush=True)

    # 備用方案：直接尋找並點擊定位查詢按鈕
    try:
        locate_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '定位查詢')]"))
        )
        driver.execute_script("arguments[0].click();", locate_button)
        print("使用備用方案點擊【定位查詢】成功", flush=True)
    except Exception as backup_e:
        print("備用方案也失敗,繼續執行後續步驟", flush=True)


# ──────────────────────────────────────────────────────────────
# 🔥 方案A：當網站前端「定位查詢」下拉選單壞掉（選縣市後地段選不出來）時，
#    改打國土測繪「免申請開放API」拿代碼，再用 JS 把選項灌進下拉選單，繞過壞掉的前端。
#    API（已實測）：
#      ListCounty                         → 縣市碼
#      ListTown/{縣市碼}                  → 區碼
#      ListLandSection/{縣市碼}/{區碼}    → office, sectcode, sectstr(段名)
# ──────────────────────────────────────────────────────────────
import urllib.request as _urlreq
import xml.etree.ElementTree as _ET

def _norm_name(s):
    """名稱正規化：去空白、台↔臺 統一、巿→市，避免縣市/區/段名小差異害比對失敗。"""
    return (s or '').strip().replace('台', '臺').replace('巿', '市')

def _nlsc_api_xml(path):
    """打 api.nlsc.gov.tw/other/{path}，回傳 XML root；失敗回 None。"""
    try:
        url = "https://api.nlsc.gov.tw/other/" + path
        with _urlreq.urlopen(url, timeout=15) as _r:
            return _ET.fromstring(_r.read())
    except Exception as _e:
        print(f"[NLSC API] 取得 {path} 失敗：{_e}", flush=True)
        return None

def _nlsc_county_code(city_name):
    root = _nlsc_api_xml("ListCounty")
    if root is None:
        return None
    for it in root.iter('countyItem'):
        if _norm_name(it.findtext('countyname')) == _norm_name(city_name):
            return (it.findtext('countycode') or '').strip()
    return None

def _nlsc_town_code(county_code, area_name):
    root = _nlsc_api_xml(f"ListTown/{county_code}")
    if root is None:
        return None
    for it in list(root):
        # 容錯：不同節點名稱都掃 townname/towncode
        name = it.findtext('townname') or ''
        code = (it.findtext('towncode') or '').strip()
        if _norm_name(name) == _norm_name(area_name) and code:
            return code
    return None

def _nlsc_section(county_code, town_code, section_name):
    """回傳 (office, sectcode) 或 (None, None)"""
    root = _nlsc_api_xml(f"ListLandSection/{county_code}/{town_code}")
    if root is None:
        return None, None
    for it in root.iter('sectItem'):
        if _norm_name(it.findtext('sectstr')) == _norm_name(section_name):
            return (it.findtext('office') or '').strip(), (it.findtext('sectcode') or '').strip()
    return None, None

def _select_with_api_fallback(driver, select_id, want_text, kind, city_name=None, area_name=None):
    """先用可見文字選；失敗(前端沒填好選項)就打 API 拿代碼、JS 灌進選項再選。
    kind: 'area' or 'section'。回傳 True/False。"""
    from selenium.webdriver.support.ui import Select as _Select
    try:
        el = WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.ID, select_id)))
        # 🔥 診斷：印出目前選項的 value/text 格式（之後微調灌值格式用）
        try:
            opts = driver.execute_script(
                "return Array.from(arguments[0].options).slice(0,4).map(o=>o.value+'｜'+o.text);", el)
            print(f"[診斷] #{select_id} 現有選項(前4)：{opts}", flush=True)
        except Exception:
            pass
        _Select(el).select_by_visible_text(want_text)
        return True
    except Exception:
        print(f"[後備] #{select_id} 找不到「{want_text}」選項（前端可能壞了），改用 API 灌入…", flush=True)
        try:
            cc = _nlsc_county_code(city_name)
            if not cc:
                print("[後備] 取縣市碼失敗", flush=True); return False
            if kind == 'section':
                tc = _nlsc_town_code(cc, area_name)
                if not tc:
                    print("[後備] 取區碼失敗", flush=True); return False
                office, sectcode = _nlsc_section(cc, tc, want_text)
                if not sectcode:
                    print(f"[後備] API 找不到地段「{want_text}」", flush=True); return False
                # 🔥 網頁 #section 的 option value 格式為「office_sectcode」（例：EB_0901），實測確認
                inject_val = f"{office}_{sectcode}"
            else:
                inject_val = _nlsc_town_code(cc, want_text) or want_text
            # 用 JS 把選項灌進去並選取、觸發 change
            driver.execute_script("""
                var sel = document.getElementById(arguments[0]);
                if(!sel) return;
                var v = arguments[1], t = arguments[2];
                var found=false;
                for(var i=0;i<sel.options.length;i++){ if(sel.options[i].text===t){sel.selectedIndex=i;found=true;break;} }
                if(!found){ var o=document.createElement('option'); o.value=v; o.text=t; sel.add(o); sel.value=v; }
                sel.dispatchEvent(new Event('change',{bubbles:true}));
            """, select_id, inject_val, want_text)
            print(f"[後備] 已用 API 代碼灌入 #{select_id}=「{want_text}」(value={inject_val})", flush=True)
            return True
        except Exception as _e:
            print(f"[後備] 灌入 #{select_id} 失敗：{_e}", flush=True)
            return False


# 遍歷 data.json 中的每一組數據
for data in data_list:
    city = data['city']
    area = data['area']
    section = data['section']
    lot_number = data['lot_number']

    print(f"處理資料: 城市 = {city}, 區域 = {area}, 段 = {section}, 地號 = {lot_number}", flush=True)
    
    # 選擇縣市、區域、地段
    try:
        city_select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "city"))
        )
        city_select = Select(city_select_element)
        city_select.select_by_visible_text(city)
        print(f"已選擇縣市: {city}", flush=True)

        # 🔥 選區：正常用可見文字選；前端壞了就用 API 灌入（方案A）
        _select_with_api_fallback(driver, "area_office", area, 'area', city_name=city, area_name=area)
        print(f"已選擇鄉鎮市區: {area}", flush=True)
        time.sleep(0.6)  # 等「地段」選單依區載入

        # 🔥 選段：正常用可見文字選；前端壞了就用 API 灌入（方案A）
        _select_with_api_fallback(driver, "section", section, 'section', city_name=city, area_name=area)
        print(f"已選擇地段: {section}", flush=True)

        landcode_input_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "landcode"))
        )
        landcode_input_element.clear()
        landcode_input_element.send_keys(lot_number)
        print(f"已填入地號: {lot_number}", flush=True)
    except Exception as e:
        print(f"操作過程中發生錯誤: {e}", flush=True)

    # 點擊「定位」按鈕
    try:
        locate_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "div_cross_query"))
        )
        locate_button.click()
        print("已點擊【定位】按鈕", flush=True)
        
        # time.sleep(1)

        try:
            # 直接使用提供的XPath精确定位对话框和确定按钮
            confirm_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[27]/div[3]/div/button[1]"))
            )
            # 点击确定按钮
            driver.execute_script("arguments[0].scrollIntoView(true);", confirm_button)
            driver.execute_script("arguments[0].click();", confirm_button)
            print("已精確點擊【地籍圖使用資訊】對話框的【確定】按鈕", flush=True)
        except TimeoutException:
            # 如果特定XPath找不到，尝试通过标题精确定位对话框
            try:
                # 查找标题为"地籍圖使用資訊"的对话框
                dialog_title = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//span[@class='ui-dialog-title' and text()='地籍圖使用資訊']"))
                )
                # 找到对应的对话框容器
                dialog = dialog_title.find_element(By.XPATH, "./ancestor::div[@role='dialog']")
                
                # 在这个特定对话框中找确定按钮
                confirm_button = dialog.find_element(By.XPATH, ".//button[text()='確定']")
                driver.execute_script("arguments[0].scrollIntoView(true);", confirm_button)
                driver.execute_script("arguments[0].click();", confirm_button)
                print("通过标题找到【地籍圖使用資訊】对话框并点击【确定】按钮", flush=True)
            except Exception as e:
                print(f"通过标题查找【地籍圖使用資訊】对话框或点击确定按钮时出错：{e}", flush=True)
                
                # 最后尝试：通过JavaScript定向查找和关闭
                try:
                    result = driver.execute_script("""
                        // 查找标题为"地籍圖使用資訊"的对话框
                        var targetDialogs = [];
                        var allDialogs = document.querySelectorAll('div[role="dialog"]');
                        
                        for(var i=0; i<allDialogs.length; i++) {
                            var title = allDialogs[i].querySelector('.ui-dialog-title');
                            if(title && title.textContent === '地籍圖使用資訊') {
                                targetDialogs.push(allDialogs[i]);
                            }
                        }
                        
                        // 如果找到目标对话框，点击其中的确定按钮
                        if(targetDialogs.length > 0) {
                            var dialog = targetDialogs[0]; // 使用第一个匹配的对话框
                            var buttons = dialog.querySelectorAll('button');
                            var clicked = false;
                            
                            // 查找并点击确定按钮
                            for(var j=0; j<buttons.length; j++) {
                                if(buttons[j].textContent === '確定') {
                                    buttons[j].click();
                                    clicked = true;
                                    return {success: true, message: "成功点击确定按钮"};
                                }
                            }
                            
                            // 如果没有找到确定按钮，点击第一个按钮
                            if(!clicked && buttons.length > 0) {
                                buttons[0].click();
                                return {success: true, message: "找不到确定按钮，点击了第一个按钮"};
                            }
                            
                            return {success: false, message: "找到对话框但没有找到可点击的按钮"};
                        }
                        
                        return {success: false, message: "没有找到标题为地籍圖使用資訊的对话框"};
                    """)
                    
                    print(f"JavaScript执行结果: {result.get('success', False)}, {result.get('message', '')}", flush=True)
                except Exception as js_e:
                    print(f"JavaScript方法失败：{js_e}", flush=True)
        except Exception as e:
            print(f"精确定位【地籍圖使用資訊】对话框的确定按钮失败：{e}", flush=True)

    except TimeoutException:
        print("【定位】按鈕未出現或無法點擊", flush=True)
    except Exception as e:
        print(f"發生錯誤: {e}", flush=True)

# 所有「定位」完成後，開始點擊「著色」按鈕
try:
    # 獲取所有「著色」按鈕
    color_buttons = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, "//input[contains(@value, '著色')]"))
    )

    for color_button in color_buttons:
        try:
            driver.execute_script("arguments[0].scrollIntoView();", color_button)
            driver.execute_script("arguments[0].click();", color_button)
            print("已點擊「著色」按鈕", flush=True)
        except Exception as e:
            print(f"點擊「著色」按鈕時發生錯誤: {e}", flush=True)

except TimeoutException:
    print("「著色」按鈕未找到", flush=True)

time.sleep(1)  # 延時以便觀察效果，或視情況延遲時間

try:
    close_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[23]/div[1]/button/span[1]"))
    )
    driver.execute_script("arguments[0].scrollIntoView();", close_button)
    driver.execute_script("arguments[0].click();", close_button)
    print("點擊【定位查詢】關閉按鈕", flush=True)
except Exception as e:
    print(f"出現錯誤: {e}", flush=True)

# =======  倒數計時器  ========
print("\033[93m開始倒數 10 秒，請調整網頁截圖位置，\033[0m", flush=True)
print("\033[93m或點擊【手指】查詢，可增添相關資訊，如【土地資訊】\033[0m", flush=True)
# 🔥 改用 \n 換行而非 \r（主程式訊息區是 line-buffered，\r 不會即時顯示）
for i in range(10, 0, -1):
    print(f"倒數：{i} 秒", flush=True)
    time.sleep(1)
print("倒數結束", flush=True)
# =======  倒數計時器結束 ========

# 🔥 使用已過濾的 data_list（子批次模式時只有本批的地號）
# 避免重新讀取 data.json 而拿到全部地號，導致截圖檔名涵蓋非本批的地號
data = data_list

# 取得第一組資料（依本批決定 region/section/lot_numbers，用於檔名）
first_entry = data[0]
region = first_entry['area']
section = first_entry['section']
lot_numbers = [entry['lot_number'] for entry in data if entry['area'] == region and entry['section'] == section]

# 🔥 建立目錄：子批次需用「完整 data.json 第一筆」建 base_directory，
# 否則每批會跑到不同資料夾（父程序就找不到截圖了）
if _is_child_batch:
    try:
        with open(get_data_json_path(), 'r', encoding='utf-8') as _full_f:
            _full_data = json.load(_full_f)
        _full_region = _full_data[0]['area']
        _full_section = _full_data[0]['section']
        _full_first_lot = next(
            (e['lot_number'] for e in _full_data
             if e['area'] == _full_region and e['section'] == _full_section),
            lot_numbers[0]
        )
        base_directory = get_work_folder(f"{_full_region}{_full_section}-{_full_first_lot}")
        print(f"[子批次] base_directory 對齊到父程序：{base_directory}", flush=True)
    except Exception as _e:
        print(f"[子批次] 對齊 base_directory 失敗，退回本批第一筆：{_e}", flush=True)
        base_directory = get_work_folder(f"{region}{section}-{lot_numbers[0]}")
else:
    base_directory = get_work_folder(f"{region}{section}-{lot_numbers[0]}")
os.makedirs(os.path.join(base_directory, "1.基本資料", "png"), exist_ok=True)

# 定義 PDF 檔名
if len(set(entry['section'] for entry in data)) == 1:  # 地區地段相同
    pdf_filename = f"06_國土測繪-{region}{section}-" + '.'.join(lot_numbers) + ".pdf"
else:  # 地區相同但地段不同
    sections = set(entry['section'] for entry in data)
    pdf_filename = f"06_國土測繪-{region}-" + '+'.join([f"{sec}-{lot_numbers[0]}" for sec in sections]) + ".pdf"

# 嘗試確認並截取特定元素的螢幕
try:
    # 等待指定元素可見
    element_to_screenshot = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "normal_tools_img_select"))
    )

    # 滾動到該元素
    driver.execute_script("arguments[0].scrollIntoView();", element_to_screenshot)

    # 設定截圖路徑，命名要與 PDF 命名一致
    png_filename = f"06_國土測繪-{region}{section}-" + '.'.join(lot_numbers) + ".png"
    screenshot_path = os.path.join(base_directory, "1.基本資料", "png", png_filename)
    
    # 確保目錄存在
    os.makedirs(os.path.dirname(screenshot_path), exist_ok=True)
    
    # 保存截圖
    driver.save_screenshot(screenshot_path)
    print(f"\033[93m網頁截圖已保存為 {screenshot_path}\033[0m", flush=True)

except Exception as e:
    print(f"截圖時發生錯誤: {e}", flush=True)

# 將 PNG 轉成橫向的 PDF
pdf_path = os.path.join(base_directory, "1.基本資料", pdf_filename)
try:
    # 打開已經截取的 PNG 圖片
    img = Image.open(screenshot_path)
    pdf = FPDF(orientation='L', unit='mm', format='A4')  # 橫向 PDF
    pdf.add_page()
    
    # 設置更高解析度
    img_width, img_height = img.size
    a4_width_mm, a4_height_mm = 297, 210  # A4 尺寸的毫米大小 (橫向)
    
    # 設定 DPI (例如 300 DPI)
    dpi = 300
    scale_factor = dpi / 72  # 將 DPI 調整至 300
    
    # 計算縮放後的圖像尺寸
    img_width_px = int(a4_width_mm * scale_factor)
    img_height_px = int(a4_height_mm * scale_factor)

    # 重新調整圖像大小
    img_resized = img.resize((img_width_px, img_height_px), Image.Resampling.LANCZOS)
    
    # 保存縮放後的圖片，設定高 DPI
    img_resized.save(screenshot_path, dpi=(dpi, dpi))

    # 將縮放後的 PNG 轉成 PDF，確保圖片與 A4 尺寸對齊
    pdf.image(screenshot_path, x=0, y=0, w=a4_width_mm, h=a4_height_mm)
    pdf.output(pdf_path)
    print(f"\033[93mPDF 已經儲存至 {pdf_path}\033[0m", flush=True)
    
except Exception as e:
    print(f"將 PNG 轉成 PDF 時發生錯誤: {e}", flush=True)

# 結束 Selenium session
driver.quit()

def notify_main_program():
    # 🔥 子批次模式不能印此訊息：main.py 會以此字串判斷整個程序結束，
    # 一旦子批次印了，後續其他批次的輸出就會被 main.py 忽略
    if not _is_child_batch:
        print("國土測繪圖資已完成執行", flush=True)

if __name__ == "__main__":
    try:
        # 子程式的主要邏輯
        # 在這裡執行子程式的操作，例如表單填寫和資料處理
        if not _is_child_batch:
            print("執行國土測繪圖資主要邏輯", flush=True)

    except Exception as e:
        print(f"國土測繪圖資執行過程中發生錯誤: {e}", flush=True)
    finally:
        # 在結尾通知主程式
        notify_main_program()