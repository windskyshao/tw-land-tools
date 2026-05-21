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

        with urllib.request.urlopen(api_url, timeout=30) as response:
            data = json.loads(response.read().decode('utf-8'))

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
        urllib.request.urlretrieve(download_url, zip_path)

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

def create_chrome_driver(options=None):
    """
    建立 Chrome WebDriver

    Args:
        options: ChromeOptions 物件（可選）

    Returns:
        WebDriver 實例
    """
    if options is None:
        options = webdriver.ChromeOptions()

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

# 向後相容：提供舊的函數名稱
def get_webdriver(options=None):
    """向後相容的函數名稱"""
    return create_chrome_driver(options)
