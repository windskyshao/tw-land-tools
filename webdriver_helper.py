"""
WebDriver 輔助模組 - 統一處理 ChromeDriver 初始化

支援三種模式：
1. 本地 chromedriver.exe（打包環境）- 會自動檢查版本並更新
2. webdriver-manager 自動下載（開發環境）
3. 系統 PATH 中的 chromedriver
"""
import os
import sys
import re
import subprocess
import zipfile
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

def get_resource_path(relative_path):
    """
    獲取資源的絕對路徑，適用於開發環境和 PyInstaller 打包後的環境
    """
    try:
        # PyInstaller 會創建臨時資料夾，並將路徑存儲在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        # 如果不是打包環境，使用當前工作目錄
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_chrome_browser_version():
    """
    取得已安裝的 Chrome 瀏覽器版本

    Returns:
        版本字串（例如 "143.0.7499.41"）或 None
    """
    try:
        # Windows: 從登錄檔讀取 Chrome 版本
        import winreg
        key_path = r"SOFTWARE\Google\Chrome\BLBeacon"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
            version, _ = winreg.QueryValueEx(key, "version")
            winreg.CloseKey(key)
            return version
        except:
            pass

        # 備用方法：直接執行 Chrome 取得版本
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]
        for chrome_path in chrome_paths:
            if os.path.exists(chrome_path):
                result = subprocess.run(
                    [chrome_path, "--version"],
                    capture_output=True, text=True, timeout=10
                )
                if result.returncode == 0:
                    # 輸出格式: "Google Chrome 143.0.7499.41"
                    match = re.search(r'(\d+\.\d+\.\d+\.\d+)', result.stdout)
                    if match:
                        return match.group(1)
    except Exception as e:
        print(f"[WebDriver] 無法取得 Chrome 版本: {e}")

    return None

def get_chromedriver_version(driver_path):
    """
    取得 ChromeDriver 的版本

    Args:
        driver_path: chromedriver.exe 的路徑

    Returns:
        版本字串（例如 "141.0.7499.0"）或 None
    """
    try:
        result = subprocess.run(
            [driver_path, "--version"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            # 輸出格式: "ChromeDriver 141.0.7499.0 (...)"
            match = re.search(r'ChromeDriver\s+(\d+\.\d+\.\d+\.\d+)', result.stdout)
            if match:
                return match.group(1)
    except Exception as e:
        print(f"[WebDriver] 無法取得 ChromeDriver 版本: {e}")

    return None

def _curl_path():
    """回傳可用的 curl 路徑（Windows 10/11 內建 System32\\curl.exe，走 Schannel、不依賴 Python 的 OpenSSL）。"""
    import shutil as _sh
    c = _sh.which('curl')
    if c and os.path.exists(c):
        return c
    for cand in (r'C:\Windows\System32\curl.exe',):
        if os.path.exists(cand):
            return cand
    return None


def _http_get_bytes(url, timeout=30):
    """抓取 URL 內容（bytes）。優先用系統 curl，失敗才退回 urllib。
    修正：主程式以子程序啟動時，嵌入式 Python 的 _ssl 載入 OpenSSL DLL 失敗
    （_ssl.c:4192 DSO LOAD_FAILED）會讓 urllib 的 HTTPS 全掛 → ChromeDriver 更新下載失敗。
    curl 走 Windows Schannel，不碰 Python 的 OpenSSL，可繞過此問題。"""
    curl = _curl_path()
    if curl:
        try:
            r = subprocess.run([curl, '-fsSL', '--max-time', str(timeout), url],
                               capture_output=True, timeout=timeout + 15,
                               creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            if r.returncode == 0 and r.stdout:
                return r.stdout
            print(f"[WebDriver] curl 取得失敗（code {r.returncode}），改用 urllib...", flush=True)
        except Exception as _e:
            print(f"[WebDriver] curl 例外（{_e}），改用 urllib...", flush=True)
    import urllib.request
    with urllib.request.urlopen(url, timeout=timeout) as resp:
        return resp.read()


def _http_download_file(url, dest, timeout=180):
    """下載 URL 到 dest。優先 curl，失敗退回 urllib。"""
    curl = _curl_path()
    if curl:
        try:
            r = subprocess.run([curl, '-fsSL', '--max-time', str(timeout), '-o', dest, url],
                               capture_output=True, timeout=timeout + 15,
                               creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            if r.returncode == 0 and os.path.exists(dest) and os.path.getsize(dest) > 0:
                return True
            print(f"[WebDriver] curl 下載失敗（code {r.returncode}），改用 urllib...", flush=True)
        except Exception as _e:
            print(f"[WebDriver] curl 下載例外（{_e}），改用 urllib...", flush=True)
    import urllib.request
    urllib.request.urlretrieve(url, dest)
    return os.path.exists(dest) and os.path.getsize(dest) > 0


def download_chromedriver(chrome_version, target_path):
    """
    下載對應版本的 ChromeDriver

    Args:
        chrome_version: Chrome 瀏覽器版本（例如 "143.0.7499.41"）
        target_path: 下載目標路徑

    Returns:
        True 表示成功，False 表示失敗
    """
    try:
        import urllib.request
        import json

        # 取得主版本號
        major_version = chrome_version.split('.')[0]

        print(f"[WebDriver] 正在查詢 Chrome {major_version} 對應的 ChromeDriver...")

        # 使用 Chrome for Testing 的 JSON API 取得對應版本
        api_url = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"

        data = json.loads(_http_get_bytes(api_url, timeout=30).decode('utf-8'))

        # 尋找對應主版本號的最新版本
        matching_versions = []
        for version_info in data.get('versions', []):
            version = version_info.get('version', '')
            if version.startswith(f"{major_version}."):
                downloads = version_info.get('downloads', {})
                chromedriver_downloads = downloads.get('chromedriver', [])
                for download in chromedriver_downloads:
                    if download.get('platform') == 'win64':
                        matching_versions.append({
                            'version': version,
                            'url': download.get('url')
                        })

        if not matching_versions:
            print(f"[WebDriver] 找不到 Chrome {major_version} 對應的 ChromeDriver")
            return False

        # 取得最新版本
        latest = matching_versions[-1]
        download_url = latest['url']
        driver_version = latest['version']

        print(f"[WebDriver] 找到 ChromeDriver {driver_version}")
        print(f"[WebDriver] 正在下載...")

        # 下載 ZIP 檔案
        zip_path = os.path.join(os.path.dirname(target_path), "chromedriver_temp.zip")
        _http_download_file(download_url, zip_path)

        # 解壓縮
        print(f"[WebDriver] 正在解壓縮...")
        extract_dir = os.path.join(os.path.dirname(target_path), "chromedriver_temp")

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        # 尋找 chromedriver.exe
        for root, dirs, files in os.walk(extract_dir):
            if 'chromedriver.exe' in files:
                source_path = os.path.join(root, 'chromedriver.exe')

                # 刪除舊的 chromedriver.exe（如果存在）
                if os.path.exists(target_path):
                    try:
                        os.remove(target_path)
                    except:
                        # 如果無法刪除，嘗試重命名
                        backup_path = target_path + ".old"
                        if os.path.exists(backup_path):
                            os.remove(backup_path)
                        os.rename(target_path, backup_path)

                # 複製新的 chromedriver.exe
                shutil.copy2(source_path, target_path)
                print(f"[WebDriver] ChromeDriver {driver_version} 已更新完成！")
                break

        # 清理暫存檔案
        try:
            os.remove(zip_path)
            shutil.rmtree(extract_dir, ignore_errors=True)
        except:
            pass

        return True

    except Exception as e:
        print(f"[WebDriver] 下載 ChromeDriver 失敗: {e}")
        return False

def check_and_update_chromedriver(driver_path):
    """
    檢查 ChromeDriver 版本是否與 Chrome 相容，若不相容則自動更新

    Args:
        driver_path: chromedriver.exe 的路徑

    Returns:
        True 表示版本相容或已更新成功，False 表示失敗
    """
    chrome_version = get_chrome_browser_version()
    if not chrome_version:
        print("[WebDriver] 無法取得 Chrome 版本，跳過版本檢查")
        return True

    chrome_major = chrome_version.split('.')[0]

    # 如果 chromedriver.exe 不存在，直接下載
    if not os.path.exists(driver_path):
        print(f"[WebDriver] ChromeDriver 不存在，正在下載...")
        return download_chromedriver(chrome_version, driver_path)

    driver_version = get_chromedriver_version(driver_path)
    if not driver_version:
        print("[WebDriver] 無法取得 ChromeDriver 版本，嘗試重新下載...")
        return download_chromedriver(chrome_version, driver_path)

    driver_major = driver_version.split('.')[0]

    # 🔥 只有版本不匹配（需更新）時才印詳細版本資訊；相容時靜默通過
    if chrome_major != driver_major:
        print(f"[WebDriver] Chrome 版本: {chrome_version} (主版本: {chrome_major})")
        print(f"[WebDriver] ChromeDriver 版本: {driver_version} (主版本: {driver_major})")
        print(f"[WebDriver] 版本不匹配！正在自動更新 ChromeDriver...")
        return download_chromedriver(chrome_version, driver_path)

    return True

def get_chrome_driver_service():
    """
    取得 ChromeDriver Service

    優先順序：
    1. 本地 chromedriver.exe（打包環境）- 會自動檢查版本並更新
    2. webdriver-manager 自動下載（開發環境）
    3. 系統 PATH 中的 chromedriver
    """

    # 方法 1: 檢查是否有本地的 chromedriver.exe（打包環境）
    local_driver_path = get_resource_path('chromedriver.exe')

    # 🔥 自動檢查並更新 ChromeDriver 版本
    if os.path.exists(local_driver_path) or True:  # 即使不存在也嘗試下載
        check_and_update_chromedriver(local_driver_path)

    if os.path.exists(local_driver_path):
        return Service(local_driver_path)

    # 方法 2: 嘗試使用 webdriver-manager（開發環境）
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        print("[WebDriver] 使用 webdriver-manager 自動下載")
        return Service(ChromeDriverManager().install())
    except ImportError:
        print("[WebDriver] webdriver-manager 未安裝，嘗試使用系統 PATH")
    except Exception as e:
        print(f"[WebDriver] webdriver-manager 失敗: {e}")

    # 方法 3: 使用系統 PATH 中的 chromedriver
    print("[WebDriver] 使用系統 PATH 中的 ChromeDriver")
    return Service('chromedriver')

def create_chrome_driver(options=None, stealth=False):
    """
    建立 Chrome WebDriver

    Args:
        options: ChromeOptions 物件（可選）
        stealth: True 時略過「會暴露自動化身分」的命令列旗標（預設 False＝維持原行為）。
                 用於有 Cloudflare/Turnstile 人機驗證的網站（如 luz.nlma.gov.tw）。
                 ⚠ 下列旗標真實使用者絕不會有，Cloudflare 會直接判定為機器人：
                   --disable-web-security（最明顯）、--disable-gpu（WebGL 指紋異常）、
                   --blink-settings、--font-render-hinting 等。
                 這些原本是為了「Chrome printToPDF 中文字型」而加；只用 save_screenshot
                 的子程式（luz.py）不需要，關掉不影響輸出。
                 淺色模式改由 prefs 達成（prefs 不是命令列開關，不會被偵測）。

    Returns:
        WebDriver 實例
    """
    if options is None:
        options = webdriver.ChromeOptions()

    if not stealth:
        # 🔥 強制使用淺色模式（避免暗黑模式影響截圖品質）
        # 1. 停用暗黑模式相關功能
        options.add_argument('--disable-features=WebUIDarkMode,ForcedColors')
        options.add_argument('--force-color-profile=srgb')

        # 2. 強制淺色主題（多種方法確保生效）
        options.add_argument('--blink-settings=preferredColorScheme=1')  # 1=light

        # 🔥 PDF 中文字型修正：確保字型正確渲染
        options.add_argument('--font-render-hinting=none')
        options.add_argument('--disable-font-subpixel-positioning')
        options.add_argument('--disable-lcd-text')
        # 停用 GPU 加速渲染（避免字型渲染問題）
        options.add_argument('--disable-gpu')
        # 確保使用系統字型
        options.add_argument('--disable-web-security')  # 允許讀取本地字型

    # 3. 設定 Chrome 偏好，強制淺色主題
    # 🔥 先取得現有的 prefs（如果有的話），避免覆蓋其他設定（例如下載目錄）
    existing_prefs = {}
    if hasattr(options, 'experimental_options') and 'prefs' in options.experimental_options:
        existing_prefs = options.experimental_options['prefs'].copy()

    # 合併淺色模式偏好和現有偏好
    prefs = {
        **existing_prefs,  # 保留現有偏好（例如下載目錄設定）
        # 停用暗黑模式
        'profile.default_content_setting_values.prefers_color_scheme': 1,  # 1=light, 2=dark
        # 確保頁面使用淺色模式
        'profile.managed_default_content_settings.images': 1,
        # 強制淺色主題
        'browser.theme.color_scheme': 'light',
        # 🔥 PDF 中文字型修正：確保正確渲染和嵌入中文字型
        'printing.print_preview_sticky_settings.appState': '{"version":2,"isGcpPromoDismissed":true}',
        'printing.default_destination_selection_rules': '{"kind":"local","namePattern":".*"}',
        # 停用 PDF 字型替換（避免中文亂碼）
        'printing.rasterize_pdf': False,
        'printing.rasterize_pdf_dpi': 300
    }
    options.add_experimental_option('prefs', prefs)

    # 4. 透過 Chrome Flags 強制淺色
    options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
    options.add_experimental_option('useAutomationExtension', False)

    try:
        service = get_chrome_driver_service()
        driver = webdriver.Chrome(service=service, options=options)
        return driver
    except Exception as e:
        print(f"[錯誤] 無法啟動 Chrome WebDriver: {e}")
        print()
        print("請確認：")
        print("1. 已安裝 Google Chrome 瀏覽器")
        print("2. chromedriver.exe 存在於程式目錄中")
        print("3. Chrome 和 ChromeDriver 版本相容")
        raise

def verify_and_fix_chrome_window(driver, target_ratio=2/3):
    """
    驗證 Chrome 視窗是否為螢幕的指定比例（預設 2/3），若 DPI 不同步則自動修正。

    特例：只在偵測到 DPI 不同步時才動作。正常 PC（Windows DPI 設定一致）完全不會被影響。

    幂等保證：頁面縮放只在「同一個 driver 物件」的第一次呼叫時觸發；
              後續呼叫（例如重試載入網頁）只重設視窗大小，不會再按 Ctrl+-，
              避免累積縮放（80% → 64% → 51% → ... → 33%）。

    Args:
        driver: Chrome WebDriver（必須在 driver.get(URL) 之後且頁面已載入）
        target_ratio: 目標寬度佔螢幕的比例（預設 2/3）

    Returns:
        dict: {'dpi_mismatch': bool, 'chrome_screen_width': int, 'fixed_width': int}
        或 None（執行失敗時）
    """
    import time
    try:
        time.sleep(0.3)
        chrome_screen_width = driver.execute_script("return screen.width") or 0
        if chrome_screen_width <= 0:
            return None

        target_logical_width = int(chrome_screen_width * target_ratio)
        current_outer_width = driver.execute_script("return window.outerWidth") or 0
        current_outer_height = driver.execute_script("return window.outerHeight") or 800
        current_screen_x = driver.execute_script("return window.screenX") or 0
        current_screen_y = driver.execute_script("return window.screenY") or 0

        result = {
            'dpi_mismatch': False,
            'chrome_screen_width': chrome_screen_width,
            'fixed_width': current_outer_width
        }

        # 偏差超過 50 px 才動，避免抖動
        if abs(current_outer_width - target_logical_width) > 50:
            result['dpi_mismatch'] = True
            print(f"[視窗] 偵測到 DPI 不同步：Chrome screen.width={chrome_screen_width}, 當前寬度={current_outer_width}, 目標={target_logical_width}", flush=True)

            # 依目前視窗位置判斷放左還是放右
            center_x = chrome_screen_width // 2
            if (current_screen_x + current_outer_width / 2) <= center_x:
                new_x = 0
            else:
                new_x = chrome_screen_width - target_logical_width

            driver.set_window_size(target_logical_width, current_outer_height)
            driver.set_window_position(new_x, current_screen_y)
            result['fixed_width'] = target_logical_width
            print(f"[視窗] ✓ 已修正：寬度 {target_logical_width}, 位置 ({new_x}, {current_screen_y})", flush=True)

            # 特例：Chrome 邏輯寬度 < 1200 → 同步把頁面縮放到 80%
            # 🔥 用 driver 屬性當旗標，避免重試時累積縮放（80%→64%→51%→...→33%）
            already_zoomed = getattr(driver, '_dpi_zoom_applied', False)
            if chrome_screen_width < 1200 and not already_zoomed:
                try:
                    import keyboard
                    time.sleep(0.5)
                    driver.execute_script("window.focus();")
                    time.sleep(0.3)
                    print(f"[視窗] DPI 不同步特例：Chrome 邏輯寬度 {chrome_screen_width} < 1200，同步將頁面縮放至 80%", flush=True)
                    for _ in range(2):
                        keyboard.press_and_release('ctrl+-')
                        time.sleep(0.5)
                    print(f"[視窗] ✓ 頁面已縮放至 80%", flush=True)
                    driver._dpi_zoom_applied = True
                except Exception as _e:
                    print(f"[視窗] 頁面縮放修正失敗：{_e}", flush=True)
            elif already_zoomed:
                # 已縮放過（重試的場合）→ 只重設視窗大小，不再按 Ctrl+-，避免累積
                pass

        return result
    except Exception as _e:
        print(f"[視窗] 驗證寬度時發生錯誤（不影響執行）：{_e}", flush=True)
        return None


def bring_chrome_foreground(driver):
    """用 AttachThreadInput 解除 Windows 前景鎖，把本支 Chrome 拉到前景，
    keyboard 的 Ctrl+- 才送得進去（背景程式直接 SetForegroundWindow 會被擋）。"""
    import time
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
        marker = "CHROMEZOOMWIN8741"
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
            return None
        hwnd = found[0]
        fg = user32.GetForegroundWindow()   # 動作前的前景（通常是主程式）
        fg_thread = user32.GetWindowThreadProcessId(fg, None)
        cur_thread = kernel32.GetCurrentThreadId()
        user32.AttachThreadInput(cur_thread, fg_thread, True)
        user32.ShowWindow(hwnd, 9)        # SW_RESTORE
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
        user32.AttachThreadInput(cur_thread, fg_thread, False)
        time.sleep(0.3)
        return fg   # 🔥 回傳原本的前景視窗，縮放完用 restore_foreground 還原
    except Exception as e:
        print(f"[頁面縮放] 拉前景失敗：{e}", flush=True)
        return None


def restore_foreground(hwnd):
    """把 OS 前景還給指定視窗（通常是縮放前的主程式），
    否則縮放把 Chrome 拉到前景後，使用者打字會跑到 Chrome（要再點輸入框才正常）。"""
    if not hwnd:
        return
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        user32.GetWindowThreadProcessId.restype = wintypes.DWORD
        fg = user32.GetForegroundWindow()
        fg_thread = user32.GetWindowThreadProcessId(fg, None)
        cur_thread = kernel32.GetCurrentThreadId()
        user32.AttachThreadInput(cur_thread, fg_thread, True)
        user32.SetForegroundWindow(hwnd)
        user32.AttachThreadInput(cur_thread, fg_thread, False)
    except Exception:
        pass


def apply_page_zoom(driver, zoom_count):
    """把瀏覽器縮放到指定級數（按 zoom_count 次 Ctrl+-）。
    先拉前景→Ctrl+0 重設回100%→按 Ctrl+- 縮到目標→用前後 devicePixelRatio 比值驗證實際%。
    zoom_count<=0 不縮放。各子程式共用，確保 keyboard 縮放真的送進 Chrome。"""
    import time
    from selenium.webdriver.common.by import By
    if not zoom_count or zoom_count <= 0:
        print("[頁面縮放] 螢幕配置正常，不需要縮放", flush=True)
        return None
    try:
        import keyboard
    except Exception as e:
        print(f"[頁面縮放] 無 keyboard 模組：{e}", flush=True)
        return None
    STEP_PCT = {1: '90%', 2: '80%', 3: '75%', 4: '67%', 5: '50%'}
    target = STEP_PCT.get(zoom_count, f'{zoom_count}次')
    _prev_fg = None
    try:
        _prev_fg = bring_chrome_foreground(driver)
        try:
            driver.find_element(By.TAG_NAME, 'body').click()
        except Exception:
            pass
        time.sleep(0.4)
        keyboard.press_and_release('ctrl+0')   # 先重設回100%（絕對縮放、跨筆不累加）
        time.sleep(0.5)
        try:
            dpr0 = driver.execute_script("return window.devicePixelRatio;")
        except Exception:
            dpr0 = None
        print(f"[頁面縮放] 縮放中（按 {zoom_count} 次，目標 {target}）...", flush=True)
        for _ in range(zoom_count):
            keyboard.press_and_release('ctrl+-')
            time.sleep(0.4)
        actual = None
        try:
            dpr1 = driver.execute_script("return window.devicePixelRatio;")
            if dpr0:
                actual = round(dpr1 / dpr0 * 100)
        except Exception:
            pass
        if actual:
            try:
                ok = "✓ 成功" if abs(actual - int(str(target).rstrip('%'))) <= 3 else "⚠ 似乎沒生效"
            except Exception:
                ok = "已縮放"
            print(f"[頁面縮放] {ok}：實際約 {actual}%（目標 {target}）", flush=True)
        else:
            print(f"[頁面縮放] 已送出縮放至 {target}", flush=True)
        return actual
    except Exception as e:
        print(f"[頁面縮放] 設定縮放時發生錯誤: {e}", flush=True)
        return None
    finally:
        # 🔥 縮放完把 OS 前景還給原本的視窗（通常是主程式），
        #    否則 Chrome 還在前景，使用者打字會跑到 Chrome（要再點輸入框才正常）
        try:
            restore_foreground(_prev_fg)
        except Exception:
            pass


# 向後相容：提供舊的函數名稱
def get_webdriver(options=None):
    """向後相容的函數名稱"""
    return create_chrome_driver(options)
