"""
NBUPIC 全國建築管理資訊系統 - 自動化查詢
適用 9 個縣市（共用同一套 cloudbm.nlma.gov.tw/NBUPIC 系統）：
  苗栗縣、雲林縣、嘉義市、嘉義縣、臺南市、屏東縣、臺東縣、澎湖縣、連江縣

由 main.py 透過 --city 參數啟動指定縣市的查詢。
資料來源：data.json 的第 1 筆，取 city / area / section / lot_number。
"""
import sys
import os

# 打包環境路徑修正
if getattr(sys, 'frozen', False):
    internal_dir = os.path.dirname(os.path.abspath(__file__))
    if internal_dir not in sys.path:
        sys.path.insert(0, internal_dir)
elif '__file__' in globals():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)

import os
os.system('cls')
print("歡迎使用【全國建築執照存根查詢 - NBUPIC】~\n模組載入中...", flush=True)

import json
import time
import re
import argparse
import base64
import io
import atexit
import traceback

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    UnexpectedAlertPresentException, WebDriverException,
    InvalidSessionIdException, ElementClickInterceptedException,
)

from PIL import Image
from fpdf import FPDF
import numpy as np

from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window, apply_page_zoom
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder


def _nbupic_zoom_count():
    """依 window_config 的 DPI + zoom_config 覆寫（key=nlma_nbupic），算出要縮幾次。"""
    zc = 0
    try:
        _cfg = os.path.join(BASE_DIR, 'window_config.json')
        if os.path.exists(_cfg):
            c = json.load(open(_cfg, encoding='utf-8'))
            dpi = c.get('dpi_scale', 1.0) or 1.0
            lw = c.get('screen_width', 1920) / dpi
            if dpi >= 1.75:
                zc = 4
            elif dpi >= 1.5:
                zc = 2
            elif dpi >= 1.25 and lw < 1200:
                zc = 2
            else:
                zc = 0
    except Exception:
        pass
    try:
        _zp = os.path.join(BASE_DIR, 'zoom_config.json')
        if os.path.exists(_zp):
            ov = json.load(open(_zp, encoding='utf-8')).get('overrides', {}).get('nlma_nbupic')
            if ov is not None:
                zc = int(ov)
    except Exception:
        pass
    return zc

# ====== 縣市對應表（NBUPIC 9 個縣市）======
# LICENSING_UNIT 是「發照單位」select 的 value
# BMLAN_CITY 是「建築地號的縣市」select 的 value
# 兩個 select 雖然包含全國 22 縣市，但只有以下 9 個是 NBUPIC 系統實際支援的
LICENSING_UNIT_BY_CITY = {
    "苗栗縣": "I70",
    "雲林縣": "IC0",
    "嘉義市": "ID0",
    "嘉義縣": "IE0",
    "臺南市": "IF0",
    "台南市": "IF0",
    "屏東縣": "II0",
    "臺東縣": "IK0",
    "台東縣": "IK0",
    "澎湖縣": "IL0",
    "連江縣": "J10",
}

BMLAN_CITY_BY_CITY = {
    "苗栗縣": "10005",
    "雲林縣": "10009",
    "嘉義市": "10020",
    "嘉義縣": "10010",
    "臺南市": "67000",
    "台南市": "67000",
    "屏東縣": "10013",
    "臺東縣": "10014",
    "台東縣": "10014",
    "澎湖縣": "10016",
    "連江縣": "09007",
}

# ====== 命令列參數 ======
def parse_args():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--city', type=str, default=None,
                        help='指定縣市名稱（例 臺南市、屏東縣）；省略則用 data.json 第 1 筆的 city')
    parser.add_argument('--base-dir', type=str, default=None)
    args, unknown = parser.parse_known_args()
    return args


# ====== Selenium 工具 ======
driver = None  # 全域，方便 finally 區塊清理


def _get_dpi_scale():
    """偵測系統 DPI 縮放比例（fallback 1.0）"""
    try:
        import ctypes
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return dpi / 96.0
    except Exception:
        return 1.0


def _apply_chrome_window_layout():
    """讀 main.py 寫的 window_config.json，把 Chrome 擺到主程式對側、佔螢幕 2/3 寬。
    複製自 luz.py 的同名邏輯。讀不到設定檔就用 1024×1024@(0,0) fallback。
    """
    width, height, x, y = 1024, 1024, 0, 0  # fallback

    try:
        config_path = os.path.join(BASE_DIR, 'window_config.json')
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                window_config = json.load(f)

            work_area = window_config.get('work_area', {})
            main_window = window_config.get('main_window', {})

            if work_area and main_window:
                dpi_scale = window_config.get('dpi_scale') or _get_dpi_scale()
                main_x = main_window.get('x', 1280)
                main_w = main_window.get('width', 640)

                work_left = work_area.get('left', 0)
                work_top = work_area.get('top', 0)
                work_right = work_area.get('right', 1920)
                work_bottom = work_area.get('bottom', 1080)

                screen_width_physical = work_right - work_left
                chrome_width = (screen_width_physical * 2) // 3

                # 主程式中心點 vs 螢幕中心，決定 Chrome 放左還右
                main_center_x = main_x + main_w / 2
                screen_center_x = (work_left + work_right) / 2
                chrome_x = work_left if main_center_x > screen_center_x else (work_right - chrome_width)
                chrome_y = work_top
                chrome_height = work_bottom - work_top - 20

                if chrome_width < 800: chrome_width = 800
                if chrome_height < 600: chrome_height = 600

                # 轉邏輯像素
                CHROME_TITLEBAR = 32
                width = int(chrome_width / dpi_scale)
                height = int(chrome_height / dpi_scale) + CHROME_TITLEBAR
                x = int(chrome_x / dpi_scale)
                y = int(chrome_y / dpi_scale)
    except Exception as e:
        print(f"  [警告] 讀取 window_config.json 失敗：{e}", flush=True)

    try:
        driver.set_window_size(width, height)
        driver.set_window_position(x, y)
        # 一般情況靜默；設定失敗才印警告
    except Exception as e:
        print(f"  [警告] 設定 Chrome 視窗位置失敗：{e}", flush=True)


def safe_input(prompt=""):
    """input 的安全包裝，stdin 已關閉時不丟例外"""
    try:
        if prompt:
            print(prompt, end='', flush=True)
        return input()
    except (EOFError, OSError):
        return ""


def is_browser_open():
    if driver is None:
        return False
    try:
        _ = driver.current_url
        return True
    except (InvalidSessionIdException, WebDriverException):
        return False
    except Exception:
        return False


def accept_any_alert():
    """🔥 偵測並關閉任何待處理的 alert/confirm 對話框，避免阻塞 Selenium 呼叫。
    每次重要動作前都應該呼叫，因為 NBUPIC 的 JS 可能跳警告。
    """
    try:
        from selenium.webdriver.support.ui import WebDriverWait as _Wait
        _Wait(driver, 1).until(EC.alert_is_present())
        alert_text = driver.switch_to.alert.text
        driver.switch_to.alert.accept()
        print(f"  [處理 alert] {alert_text}", flush=True)
        time.sleep(0.3)
        return True
    except TimeoutException:
        return False
    except Exception:
        return False


# ====== NBUPIC 表單操作 ======
NBUPIC_URL_TEMPLATE = "http://cloudbm.nlma.gov.tw/NBUPIC/?organ={organ}"
# 🔥 直接帶對應縣市的 organ 進去，避免後續用 JS 改 LICENSING_UNIT 時 session
#    可能綁在初始 organ 造成查詢結果異常。例：
#      屏東縣 → organ=II0 → https://cloudbm.nlma.gov.tw/NBUPIC/?organ=II0
#      臺南市 → organ=IF0 → https://cloudbm.nlma.gov.tw/NBUPIC/?organ=IF0


def open_nbupic(city_name):
    """開啟對應縣市的 NBUPIC 首頁並等載入完成。
    🔥 即使 driver.get() 因 timeout 拋例外，仍繼續嘗試操作 —— 因為 'eager' 策略下
       DOMContentLoaded 通常已完成，只是某個慢資源拖累 page-load 訊號。
    """
    organ = LICENSING_UNIT_BY_CITY.get(city_name)
    if not organ:
        print(f"\033[91m[錯誤] 「{city_name}」沒有對應的 NBUPIC organ\033[0m", flush=True)
        return False
    url = NBUPIC_URL_TEMPLATE.format(organ=organ)
    print(f"開啟 Chrome 中（{city_name}）...", flush=True)

    try:
        driver.get(url)
    except TimeoutException:
        print(f"  [警告] 頁面載入超過 30 秒未完成，但 DOM 可能已就緒，繼續嘗試...", flush=True)
        # 嘗試 stop loading 讓 chromedriver 解除卡死
        try:
            driver.execute_script("window.stop();")
        except Exception:
            pass
    except Exception as e:
        print(f"\033[91m開啟頁面失敗：{e}\033[0m", flush=True)
        return False

    # 等真正的關鍵元素出現（不依賴 page-load 訊號）
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "Qry_LICENSING_UNIT"))
        )
    except TimeoutException:
        print("\033[91m[錯誤] 等待 Qry_LICENSING_UNIT 元素超時，頁面可能未正確載入\033[0m", flush=True)
        return False

    # 載入後可能被網站重設視窗大小，先套用原本的視窗布局
    try:
        _apply_chrome_window_layout()
    except Exception:
        pass

    # 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
    # 注意：要放在 _apply_chrome_window_layout() 之後，否則會被覆蓋掉
    verify_and_fix_chrome_window(driver)

    # 🔥 頁面縮放（可在主程式「Chrome縮放設定」面板調整，key=nlma_nbupic）
    try:
        apply_page_zoom(driver, _nbupic_zoom_count())
    except Exception:
        pass
    return True


def set_licensing_unit(city_name):
    """設定發照單位（Qry_LICENSING_UNIT）並觸發 onChange"""
    code = LICENSING_UNIT_BY_CITY.get(city_name)
    if not code:
        print(f"\033[91m[錯誤] 城市「{city_name}」不在 NBUPIC 支援清單\033[0m", flush=True)
        return False

    try:
        sel = Select(driver.find_element(By.ID, "Qry_LICENSING_UNIT"))
        sel.select_by_value(code)
        # 觸發 onChange（包含 get_dist + run_cb_addradr）
        driver.execute_script("""
            var el = document.getElementById('Qry_LICENSING_UNIT');
            if (el && typeof el.onchange === 'function') { el.onchange(); }
            try { get_dist('Qry', 'BMLAN_','DIST8'); run_cb_addradr(); } catch(e){}
        """)
        print(f"已設定發照單位：{city_name}", flush=True)
        time.sleep(1)
        return True
    except Exception as e:
        print(f"設定發照單位失敗：{e}", flush=True)
        return False


def _switch_query_type(value, expected_signal_id, mode_label):
    """通用：切換 Qry_QryType 並等對應表單元素出現。
    🔥 分步執行（避免一次塞太多動作觸發 alert 把 chromedriver 卡 120 秒）。
    🔥 每步都呼叫 accept_any_alert() 處理潛在彈窗。
    """
    print(f"\033[93m  正切換到【{mode_label}】中（網頁可能需要等候幾秒）...\033[0m", flush=True)
    try:
        accept_any_alert()
        try:
            driver.set_script_timeout(15)
        except Exception:
            pass

        # Step 1: Selenium Select 設值（比 jQuery trigger 安全，不會觸發莫名 confirm）
        sel = Select(driver.find_element(By.ID, "Qry_QryType"))
        sel.select_by_value(value)
        time.sleep(0.3)
        accept_any_alert()

        # Step 2: 確認 select.value 真的變了；若沒變，用 JS 直接設
        try:
            current_val = driver.execute_script(
                "return document.getElementById('Qry_QryType').value;"
            )
            if current_val != value:
                driver.execute_script(
                    f"document.getElementById('Qry_QryType').value = '{value}';"
                )
                time.sleep(0.2)
        except Exception:
            pass
        accept_any_alert()

        # Step 3: 直接呼叫 checkboxoc()（這就是 onchange handler 的內容）
        driver.execute_script("try { checkboxoc(); } catch(e){}")
        time.sleep(0.5)
        accept_any_alert()

        # Step 4: 等對應表單欄位出現
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, expected_signal_id))
        )
        print(f"已切換到【{mode_label}】查詢", flush=True)
        time.sleep(0.3)
        return True
    except Exception as e:
        print(f"切換到【{mode_label}】失敗：{e}", flush=True)
        return False


def set_query_type_to_lot():
    """設定查詢方式 = 4（建築地號）"""
    return _switch_query_type("4", "Qry_BMLAN_CITY", "建築地號")


def set_query_type_to_address():
    """設定查詢方式 = 3（建築地址）"""
    return _switch_query_type("3", "Qry_BMBADD_CITY", "建築地址")


def set_bmlan_city(city_name):
    """設定建築地號的縣市（Qry_BMLAN_CITY）並觸發載入鄉鎮市區"""
    code = BMLAN_CITY_BY_CITY.get(city_name)
    if not code:
        print(f"\033[91m[錯誤] 城市「{city_name}」不在 BMLAN 對照表\033[0m", flush=True)
        return False

    try:
        sel = Select(driver.find_element(By.ID, "Qry_BMLAN_CITY"))
        sel.select_by_value(code)
        # 觸發 onChange = run_floor_city() 載入 DIST8
        driver.execute_script("try { run_floor_city(); } catch(e){}")
        print(f"已設定地號縣市：{city_name}（{code}）", flush=True)
        time.sleep(1.5)  # 等 AJAX 載入鄉鎮區
        return True
    except Exception as e:
        print(f"設定地號縣市失敗：{e}", flush=True)
        return False


def set_district(area_name):
    """設定鄉鎮市區（Qry_BMLAN_DIST8），用文字比對找對應的 option"""
    try:
        # 等選單載入
        WebDriverWait(driver, 10).until(
            lambda d: any(
                area_name in opt.text
                for opt in Select(d.find_element(By.ID, "Qry_BMLAN_DIST8")).options
            )
        )

        sel = Select(driver.find_element(By.ID, "Qry_BMLAN_DIST8"))
        target_value = None
        for opt in sel.options:
            if area_name in opt.text:
                target_value = opt.get_attribute("value")
                break
        if not target_value:
            print(f"\033[91m找不到鄉鎮區：{area_name}\033[0m", flush=True)
            return False

        sel.select_by_value(target_value)
        # 觸發 onChange（setDist + get_dist 載入段名）
        driver.execute_script("""
            var el = document.getElementById('Qry_BMLAN_DIST8');
            try { setDist(el, '8'); get_dist('Qry', 'BMLAN_','DIST3'); } catch(e){}
        """)
        print(f"已設定鄉鎮區：{area_name}（{target_value}）", flush=True)
        time.sleep(1.5)  # 等段名 AJAX 載入
        return True
    except TimeoutException:
        print(f"\033[91m等鄉鎮區選單載入超時：{area_name}\033[0m", flush=True)
        return False
    except Exception as e:
        print(f"設定鄉鎮區失敗：{e}", flush=True)
        return False


def set_section(section_name):
    """設定段名（Qry_BMLAN_SECTION）"""
    try:
        # 等段選單真的有目標段名
        WebDriverWait(driver, 10).until(
            lambda d: any(
                section_name in opt.text
                for opt in Select(d.find_element(By.ID, "Qry_BMLAN_SECTION")).options
            )
        )

        sel = Select(driver.find_element(By.ID, "Qry_BMLAN_SECTION"))
        target_value = None
        target_text = None
        for opt in sel.options:
            if section_name in opt.text:
                target_value = opt.get_attribute("value")
                target_text = opt.text
                break
        if not target_value:
            print(f"\033[91m找不到段名：{section_name}\033[0m", flush=True)
            real_options = [o.text for o in sel.options if o.text.strip() and o.text.strip() != "全部"]
            print(f"  可用段名前 5 個：{real_options[:5]}", flush=True)
            return False

        sel.select_by_value(target_value)
        print(f"已設定段名：{target_text}（{target_value}）", flush=True)
        time.sleep(0.3)
        return True
    except TimeoutException:
        print(f"\033[91m等段名選單載入超時：{section_name}\033[0m", flush=True)
        return False
    except Exception as e:
        print(f"設定段名失敗：{e}", flush=True)
        return False


def set_bmbadd_city(city_name):
    """設定建築地址的縣市（Qry_BMBADD_CITY）並觸發載入鄉鎮市區。
    BMBADD_CITY 與 BMLAN_CITY 用同一套縣市代碼（10013=屏東縣 等）。
    """
    code = BMLAN_CITY_BY_CITY.get(city_name)  # 共用對照表
    if not code:
        print(f"\033[91m[錯誤] 城市「{city_name}」不在 BMBADD 對照表\033[0m", flush=True)
        return False

    try:
        sel = Select(driver.find_element(By.ID, "Qry_BMBADD_CITY"))
        sel.select_by_value(code)
        # 觸發 onChange = run_addradr() 載入 DIST8
        driver.execute_script("try { run_addradr(); } catch(e){}")
        print(f"已設定地址縣市：{city_name}（{code}）", flush=True)
        time.sleep(1.5)
        return True
    except Exception as e:
        print(f"設定地址縣市失敗：{e}", flush=True)
        return False


def set_bmbadd_district(area_name):
    """設定地址鄉鎮市區（Qry_BMBADD_DIST8），用文字比對找對應的 option"""
    try:
        WebDriverWait(driver, 10).until(
            lambda d: any(
                area_name in opt.text
                for opt in Select(d.find_element(By.ID, "Qry_BMBADD_DIST8")).options
            )
        )
        sel = Select(driver.find_element(By.ID, "Qry_BMBADD_DIST8"))
        target_value = None
        for opt in sel.options:
            if area_name in opt.text:
                target_value = opt.get_attribute("value")
                break
        if not target_value:
            print(f"\033[91m找不到鄉鎮區：{area_name}\033[0m", flush=True)
            return False
        sel.select_by_value(target_value)
        print(f"已設定地址鄉鎮區：{area_name}（{target_value}）", flush=True)
        time.sleep(0.5)
        return True
    except TimeoutException:
        print(f"\033[91m等地址鄉鎮區選單載入超時：{area_name}\033[0m", flush=True)
        return False
    except Exception as e:
        print(f"設定地址鄉鎮區失敗：{e}", flush=True)
        return False


def _parse_choice_input(user_input, max_count):
    """解析使用者下載編號輸入，支援單個、逗號分隔、範圍 N-M、混合。
    例：'1,3,5-9,11' → [0,2,4,5,6,7,8,10]（0-based、去重、排序）
    無效 token 拋 ValueError。
    """
    choices = set()
    for token in user_input.split(","):
        token = token.strip()
        if not token:
            continue
        if "-" in token:
            # 範圍
            parts = token.split("-", 1)
            try:
                start = int(parts[0].strip())
                end = int(parts[1].strip())
            except ValueError:
                raise ValueError(f"範圍格式錯誤：{token}")
            if start > end:
                start, end = end, start
            for i in range(start, end + 1):
                idx = i - 1
                if 0 <= idx < max_count:
                    choices.add(idx)
        else:
            try:
                idx = int(token) - 1
            except ValueError:
                raise ValueError(f"非數字：{token}")
            if 0 <= idx < max_count:
                choices.add(idx)
    return sorted(choices)


def _select_lots_dialog(data_list):
    """多筆地號時詢問使用者要查哪幾筆。

    回傳 list of dict（selected entries）。若使用者取消或輸入無效則回傳 None/空 list。
    若只有一筆則直接 return [data_list[0]] 不提示。
    """
    if not data_list:
        return []
    if len(data_list) == 1:
        return [data_list[0]]

    print("\n===== 建築地號查詢 - 選擇地段地號 =====", flush=True)
    print(f"在 data.json 中共有 {len(data_list)} 筆資料：", flush=True)
    for idx, d in enumerate(data_list, 1):
        city = d.get('city', '')
        area = d.get('area', '')
        section = d.get('section', '')
        lot = d.get('lot_number', '')
        print(f"  [{idx}] {city}{area} {section} {lot}", flush=True)

    print("\n請輸入要查詢的編號：", flush=True)
    print("  例：1,3,5     單個（多個用逗號分隔）", flush=True)
    print("  例：5-10      範圍（從 5 到 10 共 6 筆）", flush=True)
    print("  例：1,3,5-9   混合", flush=True)
    print("  輸入 a 或 all 全選；留空取消", flush=True)

    user_input = safe_input().strip().lower()
    if not user_input:
        print("已取消選擇", flush=True)
        return None

    if user_input in ("a", "all"):
        choices = list(range(len(data_list)))
    else:
        try:
            choices = _parse_choice_input(user_input, len(data_list))
        except ValueError as e:
            print(f"\033[91m輸入格式錯誤（{e}），取消\033[0m", flush=True)
            return None
        if not choices:
            print("\033[91m沒有有效的編號，取消\033[0m", flush=True)
            return None

    selected = [data_list[i] for i in choices]
    print(f"\n已選擇 {len(selected)} 筆：", flush=True)
    for i, d in enumerate(selected, 1):
        print(f"  [{i}] {d.get('area','')}{d.get('section','')} {d.get('lot_number','')}", flush=True)
    return selected


def parse_address_simple(address):
    """簡易解析地址字串，回傳 dict: {road, lane, alley, number}
    範例：「廣東路456巷59弄87號」 → road="廣東路", lane="456", alley="59", number="87"
         「廣東路87號」 → road="廣東路87", number="87"  ← 因為沒有「巷」「弄」
         其實第一個範例 road 會抓「廣東路」，看 regex 設計
    保守做法：只抓「路/街/段」「巷」「弄」「號」這四項，其他不填。
    """
    parts = {"road": "", "lane": "", "alley": "", "number": ""}
    if not address:
        return parts

    # 號（最後處理避免被路名吃掉）
    m_num = re.search(r'(\d+(?:[-之]\d+)?)號', address)
    if m_num:
        parts["number"] = m_num.group(1)

    # 路 / 街 / 段（吃到「路」或「街」止；如果後面有「段」連到段）
    m_road = re.search(r'(.+?(?:路|街)(?:\S*?段)?)', address)
    if m_road:
        parts["road"] = m_road.group(1)

    # 巷
    m_lane = re.search(r'(\d+)巷', address)
    if m_lane:
        parts["lane"] = m_lane.group(1)

    # 弄
    m_alley = re.search(r'(\d+)弄', address)
    if m_alley:
        parts["alley"] = m_alley.group(1)

    return parts


def fill_address_fields(address):
    """填入地址欄位（路、巷、弄、號）。其他欄位（村里鄰、樓室）不填。"""
    parsed = parse_address_simple(address)
    print(f"  解析地址：{parsed}", flush=True)

    field_map = [
        ("Qry_addrad2", parsed["road"]),    # 路街段
        ("Qry_addrad3", parsed["lane"]),    # 巷
        ("Qry_addrad4", parsed["alley"]),   # 弄
        ("Qry_addrad5", parsed["number"]),  # 號
    ]
    for elem_id, value in field_map:
        if not value:
            continue
        try:
            elem = driver.find_element(By.ID, elem_id)
            elem.clear()
            elem.send_keys(value)
        except Exception as e:
            print(f"  填入 {elem_id} 失敗：{e}", flush=True)
    return True


def fill_lot_number(lot_number):
    """填入地號（母號 / 子號）。lot_number 例：'2286' 或 '2286-1'。
    🔥 若 data.json 沒明確帶子號（純母號），子號欄位**留空**，這樣會搜出所有
       2286-X 系列分割地號，不會只搜到 2286-0000 那一筆。
       若 data.json 帶 '2286-3' 明確子號，才填子號。
    """
    try:
        if "-" in lot_number:
            parts = lot_number.split("-", 1)
            mother, child = parts[0].strip(), parts[1].strip()
        else:
            mother, child = lot_number.strip(), ""  # 子號留空 = 搜尋所有子號

        m_input = driver.find_element(By.ID, "Qry_road_no1")
        c_input = driver.find_element(By.ID, "Qry_road_no2")
        m_input.clear()
        m_input.send_keys(mother)
        c_input.clear()
        if child:
            c_input.send_keys(child)
            print(f"已填入地號：母號 {mother}，子號 {child}", flush=True)
        else:
            print(f"已填入地號：母號 {mother}，子號（留空，搜尋全部分割地號）", flush=True)
        return True
    except Exception as e:
        print(f"填入地號失敗：{e}", flush=True)
        return False


# ====== 驗證碼處理（OCR + 手動 fallback）======
def grab_captcha_image():
    """抓取驗證碼圖片並回傳 PIL Image"""
    try:
        img_elem = driver.find_element(By.ID, "vadimg")
        png_b64 = img_elem.screenshot_as_base64
        png_bytes = base64.b64decode(png_b64)
        return Image.open(io.BytesIO(png_bytes))
    except Exception as e:
        print(f"抓取驗證碼圖片失敗：{e}", flush=True)
        return None


def refresh_captcha():
    """換一張驗證碼"""
    try:
        driver.execute_script("try { refreshImg(); } catch(e){}")
        time.sleep(0.5)
    except Exception:
        pass


# 🔥 module-level 快取的 OCR reader：建立一次重複用，避免每次 OCR 都載入 ONNX 模型
_OCR_READER = None


def _get_ocr_reader():
    global _OCR_READER
    if _OCR_READER is not None:
        return _OCR_READER
    from rapidocr_helper import create_rapidocr_reader
    _OCR_READER = create_rapidocr_reader(
        use_angle_cls=False, use_text_score=True, text_score=0.3,
        det_db_thresh=0.2, det_db_box_thresh=0.3,
        det_db_unclip_ratio=1.8, max_side_len=1280,
        quiet=True,  # 不印 RapidOCR debug 訊息（模型目錄/參數）
    )
    return _OCR_READER


def ocr_captcha():
    """用 RapidOCR 辨識驗證碼，失敗回傳 None"""
    img = grab_captcha_image()
    if img is None:
        return None
    try:
        from rapidocr_helper import rapidocr_readtext
        reader = _get_ocr_reader()
        result = rapidocr_readtext(reader, np.array(img), detail=0)
        text = ''.join(result).replace(" ", "").strip()
        text = ''.join(filter(str.isalnum, text))
        return text if text else None
    except Exception as e:
        print(f"OCR 辨識錯誤：{e}", flush=True)
        return None


def fill_captcha(code):
    """填入驗證碼"""
    try:
        elem = driver.find_element(By.ID, "Qry_imageCodetxt")
        elem.clear()
        elem.send_keys(code)
        return True
    except Exception as e:
        print(f"填入驗證碼失敗：{e}", flush=True)
        return False


def handle_captcha_with_retry(max_attempts=3):
    """OCR 試 max_attempts 次，都失敗則改手動輸入"""
    for attempt in range(1, max_attempts + 1):
        code = ocr_captcha()
        if code and 4 <= len(code) <= 8:
            print(f"  OCR 驗證碼第 {attempt} 次：{code}", flush=True)
            fill_captcha(code)
            return code
        else:
            print(f"  OCR 第 {attempt} 次無結果，換一張重試...", flush=True)
            refresh_captcha()

    # OCR 都失敗，手動輸入
    print("\033[93m[手動輸入] OCR 多次失敗，請看 Chrome 視窗的驗證碼，手動輸入：\033[0m", flush=True)
    user_code = safe_input().strip()
    if user_code:
        fill_captcha(user_code)
        return user_code
    return None


# ====== 查詢與結果解析 ======
def click_query_button():
    """點擊「執行查詢」按鈕"""
    try:
        btn = driver.find_element(By.ID, "QueryParamButton_executeQuery")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", btn)
        print("已點擊【執行查詢】", flush=True)
        return True
    except Exception as e:
        print(f"點擊查詢按鈕失敗：{e}", flush=True)
        return False


def _read_total_count_from_pagination():
    """從 #body_content .pagination 區的「共 N 筆」抽出總筆數，找不到回 None。
    🔥 這是網站宣告的權威值，用來判斷我們抓到的 row 數是否完整。
    """
    try:
        text = driver.execute_script(
            "var el = document.querySelector('#body_content .pagination');"
            "return el ? el.textContent : '';"
        ) or ""
        m = re.search(r"共\s*(\d+)\s*筆", text)
        if m:
            return int(m.group(1))
    except Exception:
        pass
    return None


def wait_for_results():
    """等待結果載入到 #body_content。回傳 'success' / 'no_data' / 'captcha_error' / 'error'。
    🔥 不只等第一個 row 出現，還要等 row 數穩定（避免 parse 拿到中間狀態）。
    """
    try:
        # 第一階段：等 body_content 出現結果表格 OR 警告訊息
        WebDriverWait(driver, 15).until(
            lambda d: bool(d.find_elements(By.CSS_SELECTOR, "#body_content #body_table table tbody tr")) or
                      bool(d.execute_script(
                          "var el = document.getElementById('Qry_warning_txt');"
                          "return el && el.textContent && el.textContent.trim().length > 0;"
                      ))
        )
    except TimeoutException:
        print("\033[91m等待結果超時\033[0m", flush=True)
        return "error"

    # 檢查警告訊息（先處理錯誤狀況）
    try:
        warning = driver.execute_script(
            "var el = document.getElementById('Qry_warning_txt');"
            "return el ? el.textContent.trim() : '';"
        )
        if warning:
            print(f"\033[91m[警告訊息] {warning}\033[0m", flush=True)
            if "驗證碼" in warning or "認證" in warning:
                return "captcha_error"
            if "查無" in warning or "無資料" in warning:
                return "no_data"
            return "error"
    except Exception:
        pass

    # 🔥 第二階段：以分頁區「共 N 筆」當權威，等實際 row 數 ≥ N 才視為完成
    # 否則容易被「分批 ajax 載入」騙到（第一批穩定後突然又補一筆）
    prev_count = -1
    stable_since = None
    deadline = time.time() + 20  # 整體最多等 20 秒
    final_count = 0
    expected_total = None
    reached_target = False
    while time.time() < deadline:
        rows = driver.find_elements(By.CSS_SELECTOR,
            '#body_content #body_table table tbody tr td[data-name="執照號碼"] a')
        current = len([a for a in rows if a.text.strip()])
        expected_total = _read_total_count_from_pagination()

        if expected_total is not None and current >= expected_total:
            final_count = current
            reached_target = True
            break

        if current != prev_count:
            prev_count = current
            stable_since = time.time()
        elif stable_since and (time.time() - stable_since) > 3.0:
            # 連續穩定 3 秒沒變動，視為到頂（但可能比目標少）
            final_count = current
            break
        time.sleep(0.3)
    else:
        final_count = prev_count if prev_count > 0 else 0

    if final_count > 0:
        if reached_target:
            print(f"  結果列表已穩定，共 {final_count} 筆", flush=True)
        elif expected_total is not None and final_count < expected_total:
            print(f"  \033[93m結果穩定但筆數可能不足：抓到 {final_count} 筆，分頁顯示應有 {expected_total} 筆\033[0m", flush=True)
        else:
            print(f"  結果列表已穩定，共 {final_count} 筆", flush=True)
        # 🔥 把頁面捲到結果區頂端，讓使用者看得到
        try:
            driver.execute_script("""
                var el = document.querySelector('#body_content #body_table');
                if (el) el.scrollIntoView({behavior: 'smooth', block: 'start'});
                else window.scrollTo({top: 0, behavior: 'smooth'});
            """)
        except Exception:
            pass
        return "success"
    return "no_data"


def set_page_size_max():
    """把每頁筆數改成 200，避免分頁麻煩。
    🔥 改完要等 ajax 重新載入完成才 return，避免 parse 拿到中間狀態。
    """
    try:
        sel = driver.find_element(By.ID, "P_PAGE_SIZE")
        cur_value = sel.get_attribute("value")
        if cur_value == "200":
            return  # 已經是最大，不需要重載

        # 記下重載前的 row 數，等 ajax 後 row 數有變化（或穩定一段時間）才算完成
        before_rows = len(driver.find_elements(By.CSS_SELECTOR, "#body_content #body_table table tbody tr"))

        Select(sel).select_by_value("200")
        driver.execute_script("""
            var el = document.getElementById('P_PAGE_SIZE');
            if (el && typeof el.onchange === 'function') el.onchange();
        """)

        # 🔥 跟 wait_for_results 同邏輯：以分頁「共 N 筆」當權威，等實際 row 達標
        prev_count = -1
        stable_since = None
        deadline = time.time() + 15
        final_count = before_rows
        while time.time() < deadline:
            current = len(driver.find_elements(By.CSS_SELECTOR,
                '#body_content #body_table table tbody tr td[data-name="執照號碼"] a'))
            expected = _read_total_count_from_pagination()
            if expected is not None and current >= expected:
                final_count = current
                break
            if current != prev_count:
                prev_count = current
                stable_since = time.time()
            elif stable_since and (time.time() - stable_since) > 2.5:
                final_count = current
                break
            time.sleep(0.3)
        print(f"  每頁筆數已改為 200（最終列數 {final_count}）", flush=True)
        # 🔥 重載完把頁面捲到結果區頂端
        try:
            driver.execute_script("""
                var el = document.querySelector('#body_content #body_table');
                if (el) el.scrollIntoView({behavior: 'smooth', block: 'start'});
                else window.scrollTo({top: 0, behavior: 'smooth'});
            """)
        except Exception:
            pass
    except Exception as e:
        print(f"  [警告] 設定每頁 200 筆失敗（不影響繼續）：{e}", flush=True)


def parse_results():
    """解析結果列表，回傳 [info_dict, ...]
    🔥 用一次 JS 原子性抓所有結果資料，避免 Selenium 多次 round-trip 時遇到
       StaleElementReferenceException（ajax 重載 row 失效）造成漏抓。
    🔥 從 onclick="run_button('XXX','3',...)" 抽 IndexKey，用來組列印版 URL。
    🔥 「使」字執照（使用執照、變使、雜使...）標記 is_usage=True。
    """
    try:
        results = driver.execute_script(r"""
            var rows = document.querySelectorAll('#body_content #body_table table tbody tr');
            var out = [];
            rows.forEach(function(row) {
                var a = row.querySelector('td[data-name="執照號碼"] a');
                if (!a) return;
                var license_no = (a.textContent || '').trim();
                if (!license_no) return;

                var onclick = a.getAttribute('onclick') || '';
                var m = onclick.match(/run_button\(\s*['"]([^'"]+)['"]/);
                var index_key = m ? m[1] : null;

                function tdText(name) {
                    var td = row.querySelector('td[data-name="' + name + '"]');
                    return td ? (td.textContent || '').trim() : '';
                }

                out.push({
                    license_no: license_no,
                    owner: tdText('起造人'),
                    lot: tdText('建築地號'),
                    address: tdText('建築地址'),
                    date: tdText('發照日期'),
                    index_key: index_key,
                    is_usage: license_no.indexOf('使') !== -1
                });
            });
            return out;
        """)
        results = results or []
        print(f"  parse_results 抓到 {len(results)} 筆", flush=True)
        return results
    except Exception as e:
        print(f"解析結果時發生錯誤：{e}", flush=True)
        return []


def highlight_usage_license_rows():
    """🔥 在 Chrome 結果列表上，把含「使」字的執照那一列改成淺藍底+深藍粗體連結"""
    try:
        driver.execute_script("""
            var rows = document.querySelectorAll('#body_content #body_table table tbody tr');
            rows.forEach(function(row) {
                var a = row.querySelector('td[data-name="執照號碼"] a');
                if (a && a.textContent.indexOf('使') !== -1) {
                    row.style.backgroundColor = '#E3F2FD';   // 淺藍底
                    a.style.fontWeight = 'bold';
                    a.style.color = '#0D47A1';                // 深藍字
                }
            });
        """)
    except Exception:
        pass


def _td_text(row, name):
    try:
        return row.find_element(By.CSS_SELECTOR, f'td[data-name="{name}"]').text.strip()
    except Exception:
        return ""


def print_results(results):
    print(f"\n查詢結果共 {len(results)} 筆：", flush=True)
    print("-" * 60, flush=True)
    for i, r in enumerate(results, 1):
        line = f"({i}) {r['license_no']} | {r.get('date', '')} | {r.get('address', '')}"
        if r.get("is_usage"):
            # 「使」字執照用青色強調
            print(f"\033[96m★ {line}\033[0m", flush=True)
        else:
            print(f"  {line}", flush=True)
    usage_count = sum(1 for r in results if r.get("is_usage"))
    if usage_count > 0:
        print(f"\033[96m★ = 使用執照類（共 {usage_count} 筆，建議優先下載）\033[0m", flush=True)
    print("-" * 60, flush=True)


# ====== 詳細頁（modal popup）截圖存 PDF ======
def open_detail_modal(license_link_element):
    """點擊執照連結開啟 modal popup"""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", license_link_element)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", license_link_element)
        # 等 popup 出現
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'div[id$="_popWindow"]'))
        )
        time.sleep(0.5)  # 等內容完全載入
        return True
    except Exception as e:
        print(f"開啟詳細頁失敗：{e}", flush=True)
        return False


def close_detail_modal():
    """關閉 modal popup"""
    try:
        # 嘗試找 _close 按鈕
        close_btns = driver.find_elements(By.CSS_SELECTOR, 'a[id$="_close"]')
        for btn in close_btns:
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(0.3)
                return True
        # fallback：找 popup id 用 JS 關閉
        popups = driver.find_elements(By.CSS_SELECTOR, 'div[id$="_popWindow"]')
        for p in popups:
            full_id = p.get_attribute("id")
            hash_part = full_id.replace("_popWindow", "")
            try:
                driver.execute_script(f"try {{ closePopWindow('{hash_part}'); }} catch(e){{}}")
            except Exception:
                pass
        time.sleep(0.3)
        return True
    except Exception as e:
        print(f"關閉詳細頁失敗：{e}", flush=True)
        return False


PRINT_VIEW_URL_TEMPLATE = (
    "http://cloudbm.nlma.gov.tw/NBUPIC/maliportal/common/licInfo_rpt.do"
    "?IndexKey={index_key}&organ={organ}"
)


def save_via_print_view(info, organ, base_directory):
    """🔥 不靠 modal screenshot，直接開「列印版」獨立頁面 → Page.printToPDF。
    這個列印版頁面內容完整且乾淨（無捲軸、無摺疊），是最理想的 PDF 來源。
    流程：
      1. 開新分頁進 licInfo_rpt.do?IndexKey=...&organ=...
      2. printToPDF
      3. 關新分頁，切回主視窗（保留結果列表）
    """
    license_no = info.get("license_no", "unknown")
    index_key = info.get("index_key")
    if not index_key:
        print(f"  [警告] {license_no} 無 IndexKey，跳過", flush=True)
        return False

    pdf_dir = os.path.join(base_directory, "1.基本資料")
    os.makedirs(pdf_dir, exist_ok=True)
    safe_name = re.sub(r'[\\/:"*?<>|]+', '_', license_no)
    pdf_path = os.path.join(pdf_dir, f"02_使用執照-{safe_name}.pdf")

    main_handle = driver.current_window_handle
    print_url = PRINT_VIEW_URL_TEMPLATE.format(index_key=index_key, organ=organ)
    new_handle = None

    try:
        # 1. 開新分頁
        driver.execute_script(f"window.open('{print_url}', '_blank');")
        time.sleep(1)
        # 找剛開的新 handle
        for h in driver.window_handles:
            if h != main_handle:
                new_handle = h
                break
        if not new_handle:
            print("  [警告] 沒抓到新分頁 handle，跳過", flush=True)
            return False
        driver.switch_to.window(new_handle)

        # 2. 等列印版頁面 DOM 就緒
        try:
            WebDriverWait(driver, 15).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
        except TimeoutException:
            print("  [警告] 列印版頁面載入超時，仍嘗試列印", flush=True)
        time.sleep(0.6)

        # 3. printToPDF
        pdf_settings = {
            "landscape": False,
            "printBackground": True,
            "marginTop": 0.4,
            "marginBottom": 0.4,
            "marginLeft": 0.4,
            "marginRight": 0.4,
            "preferCSSPageSize": True,
        }
        result = driver.execute_cdp_cmd("Page.printToPDF", pdf_settings)
        with open(pdf_path, "wb") as f:
            f.write(base64.b64decode(result["data"]))
        print(f"  \033[93m已保存 PDF：{os.path.basename(pdf_path)}\033[0m", flush=True)
        return True

    except Exception as e:
        print(f"  保存 PDF 失敗：{e}", flush=True)
        return False
    finally:
        # 4. 關新分頁、切回主視窗（保留搜尋結果列表的狀態）
        try:
            if new_handle and driver.current_window_handle == new_handle:
                driver.close()
        except Exception:
            pass
        try:
            driver.switch_to.window(main_handle)
        except Exception:
            pass


def save_modal_as_pdf(license_no, base_directory):
    """[舊路徑] 直接截 modal 的 PDF 方法。保留以備不時之需，但建議用 save_via_print_view。"""
    pdf_dir = os.path.join(base_directory, "1.基本資料")
    os.makedirs(pdf_dir, exist_ok=True)

    safe_name = re.sub(r'[\\/:"*?<>|]+', '_', license_no)
    pdf_path = os.path.join(pdf_dir, f"02_使用執照-{safe_name}.pdf")

    try:
        # 🔥 強制展開：modal 自己 + 所有後代 + body/html，三層都解除高度與 overflow 限制。
        #    並把 modal 從 position:fixed 改回 static，讓它進入正常文件流，整份內容才會
        #    被 Page.printToPDF 完整捕捉到（不會被內部捲軸截斷）。
        #    順便展開 collapse-content 區塊，避免某些段落是「摺疊」狀態。
        driver.execute_script("""
            // 1) 解除 body / html 的高度與 overflow 限制
            document.documentElement.style.overflow = 'visible';
            document.documentElement.style.height = 'auto';
            document.body.style.overflow = 'visible';
            document.body.style.height = 'auto';

            // 2) 處理 popup 與其所有後代
            var popup = document.querySelector('div[id$="_popWindow"]');
            if (popup) {
                var elements = [popup].concat(Array.prototype.slice.call(popup.querySelectorAll('*')));
                elements.forEach(function(el) {
                    el.style.maxHeight = 'none';
                    el.style.minHeight = '0';
                    el.style.height = 'auto';
                    el.style.overflow = 'visible';
                    el.style.overflowY = 'visible';
                    el.style.overflowX = 'visible';
                });
                // 把 popup 從 fixed 改成 static，進入文件流
                popup.style.position = 'static';
                popup.style.transform = 'none';
                popup.style.top = 'auto';
                popup.style.left = 'auto';
                popup.style.right = 'auto';
                popup.style.bottom = 'auto';
                popup.style.width = '100%';
            }

            // 3) 強制展開所有 collapse-content（可能是 jQuery accordion 摺疊狀態）
            document.querySelectorAll('.collapse-content').forEach(function(el) {
                el.style.display = 'block';
                el.style.maxHeight = 'none';
                el.style.height = 'auto';
                el.style.overflow = 'visible';
            });

            // 4) 滾到最頂端
            window.scrollTo(0, 0);
        """)
        time.sleep(0.8)  # 等 layout 重排完成

        pdf_settings = {
            "landscape": False,
            "printBackground": True,
            "marginTop": 0.4,
            "marginBottom": 0.4,
            "marginLeft": 0.4,
            "marginRight": 0.4,
            "preferCSSPageSize": True,
        }
        result = driver.execute_cdp_cmd("Page.printToPDF", pdf_settings)
        with open(pdf_path, "wb") as f:
            f.write(base64.b64decode(result["data"]))
        print(f"  \033[93m已保存 PDF：{os.path.basename(pdf_path)}\033[0m", flush=True)
        return True
    except Exception as e:
        print(f"保存 PDF 失敗：{e}", flush=True)
        return False


# ====== 主流程 ======
def ask_query_mode():
    """詢問使用者要哪一種查詢模式。回傳 '1'/'2'/'3'，按 Enter 預設 1。"""
    LB = "\033[94m"   # 淺藍
    RST = "\033[0m"
    print("\n" + "=" * 50, flush=True)
    print(f"{LB}===== NBUPIC 建築執照查詢 - 請選擇查詢方式 ====={RST}", flush=True)
    print(f"{LB}[1] 建築地號查詢（從 data.json 自動帶入段+地號）{RST}", flush=True)
    print(f"{LB}[2] 建築地址查詢(從 data.json 取縣市區，提示輸入路名/號){RST}", flush=True)
    print(f"{LB}[3] 手動查詢（自己在網頁上填，腳本只負責驗證碼+送出）{RST}", flush=True)
    print(f"{LB}[Q] 退出{RST}", flush=True)
    print("=" * 50, flush=True)
    while True:
        choice = safe_input("請輸入選項（1-3，預設 1）：").strip().lower()
        if not choice:
            return "1"
        if choice in ("1", "2", "3"):
            return choice
        if choice == "q":
            return None
        print("輸入無效，請重新輸入", flush=True)


def setup_lot_query(city, area, section, lot_number):
    """模式 1：建築地號 — 設定全部查詢欄位"""
    if not set_query_type_to_lot():
        return False
    if not set_bmlan_city(city):
        return False
    if not set_district(area):
        return False
    if not set_section(section):
        return False
    if not fill_lot_number(lot_number):
        return False
    return True


def setup_address_query(city, area, address):
    """模式 2：建築地址 — 設定縣市+區+地址欄位"""
    if not set_query_type_to_address():
        return False
    if not set_bmbadd_city(city):
        return False
    if not set_bmbadd_district(area):
        return False
    if not fill_address_fields(address):
        return False
    return True


def setup_manual_query():
    """模式 3：手動 — 不預先設定任何 QType/欄位，讓使用者自己在網頁操作"""
    print("\033[93m[手動模式]\033[0m", flush=True)
    print("\033[93m網頁內容選填好後，請在主程式下方輸入欄內點擊後按 Enter，程式將自動幫你解驗證碼及查詢\033[0m", flush=True)
    user_input = safe_input("（按 Enter 繼續，輸入 q 取消）：").strip().lower()
    if user_input == "q":
        return False
    return True


def main():
    global driver

    args = parse_args()

    # ===== 1. 讀 data.json，取目標縣市/區/段/地號 =====
    data_json = get_data_json_path()
    if not os.path.exists(data_json):
        print(f"\033[91m找不到 data.json：{data_json}\033[0m", flush=True)
        return

    with open(data_json, 'r', encoding='utf-8') as f:
        data_list = json.load(f)
    if not data_list:
        print("\033[91mdata.json 是空的\033[0m", flush=True)
        return

    entry = data_list[0]
    data_json_city = entry.get("city", "")

    # 🔥 city 優先序：--city 參數 > data.json 第 1 筆。明確 print 來源以利除錯。
    if args.city:
        city = args.city
        city_source = "--city 參數（主程式按鈕對應）"
    else:
        city = data_json_city
        city_source = "data.json 第 1 筆（無 --city 參數）"
    area = entry.get("area", "")
    section = entry.get("section", "")
    lot_number = entry.get("lot_number", "")
    address = entry.get("address", "") or entry.get("addr", "")  # 部分 data.json 可能有 address 欄位

    print(f"\n查詢條件：city={city}, area={area}, section={section}, lot={lot_number}", flush=True)
    print(f"  ↑ city 來源：{city_source}", flush=True)

    # 驗證縣市是否在 NBUPIC 支援清單
    if city not in LICENSING_UNIT_BY_CITY:
        print(f"\033[91m[錯誤] 「{city}」不在 NBUPIC 支援的 9 個縣市內\033[0m", flush=True)
        print(f"  支援清單：{list(LICENSING_UNIT_BY_CITY.keys())}", flush=True)
        return

    # 🔥 偵測「按鈕城市 ≠ data.json 城市」—— 表示使用者要查無關 data.json 的案件
    #    此情境 mode 1/2 會用錯誤的段/地號/地址資料 → 強制走 mode 3 手動
    #    存檔也改到通用「使用執照查詢結果」資料夾，避免污染 data.json 案件目錄
    mismatched_city = bool(args.city and data_json_city and args.city != data_json_city)

    if mismatched_city:
        # 每行都要自帶 ANSI 顏色碼（print 之間不會延續）
        Y = "\033[93m"   # 亮黃
        R = "\033[91m"   # 亮紅（強調對比的兩個城市）
        RST = "\033[0m"
        print("", flush=True)
        print(f"{Y}{'=' * 50}{RST}", flush=True)
        print(f"{Y}⚠ 偵測到查詢城市與 data.json 不一致：{RST}", flush=True)
        print(f"{Y}  按鈕指定的縣市：{R}{city}{RST}", flush=True)
        print(f"{Y}  data.json 第 1 筆：{R}{data_json_city}{RST}", flush=True)
        print(f"{Y}  → 自動進入【手動查詢】模式（請在 Chrome 視窗自填查詢條件）{RST}", flush=True)
        print(f"{Y}  → PDF 會存到「使用執照查詢結果」通用資料夾，不會混進 data.json 案件目錄{RST}", flush=True)
        print(f"{Y}{'=' * 50}{RST}", flush=True)
        mode = "3"  # 強制手動
    else:
        # ===== 2. 詢問查詢模式 =====
        mode = ask_query_mode()
        if mode is None:
            print("使用者取消查詢", flush=True)
            return

    # 🔥 多筆地號時讓使用者勾選要查的（僅 mode 1 建築地號查詢有意義）
    if mode == "1" and not mismatched_city and len(data_list) > 1:
        selected_entries = _select_lots_dialog(data_list)
        if not selected_entries:
            return
    else:
        # 不一致城市 / mode 2 / mode 3 / 單筆：只用第一筆
        selected_entries = [entry]

    # ===== 3. 啟動瀏覽器 =====
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    # 🔥 'eager' = DOMContentLoaded 就視為載入完成，不等所有 image/iframe/廣告
    #    避免 driver.get() 卡在某個慢資源 300 秒不動
    options.page_load_strategy = 'eager'

    try:
        driver = create_chrome_driver(options=options)
    except Exception as e:
        print(f"\033[91m啟動 Chrome 失敗：{e}\033[0m", flush=True)
        return

    # 🔥 設 page load timeout = 30s，超時也不會 hang 住整個程式
    try:
        driver.set_page_load_timeout(30)
        driver.set_script_timeout(30)
    except Exception:
        pass

    # 🔥 套用 main.py 寫的 window_config.json，把 Chrome 擺到主程式對側、佔螢幕 2/3 寬
    try:
        _apply_chrome_window_layout()
    except Exception as e:
        print(f"  [警告] Chrome 視窗定位失敗（不影響功能）：{e}", flush=True)

    try:
        # 🔥 多筆地號逐筆查詢（mode 2/3 或單筆時迴圈只跑一次）
        for _idx, _entry in enumerate(selected_entries, 1):
            # 覆蓋外層 area/section/lot_number/address，讓後續邏輯沿用相同變數名
            area = _entry.get("area", "")
            section = _entry.get("section", "")
            lot_number = _entry.get("lot_number", "")
            address = _entry.get("address", "") or _entry.get("addr", "")

            if len(selected_entries) > 1:
                print(f"\n\033[96m===== 第 {_idx}/{len(selected_entries)} 筆：{area}{section} {lot_number} =====\033[0m", flush=True)

            # ===== 4. 開頁面（用對應縣市的 organ）+ 設定 LICENSING_UNIT =====
            if not open_nbupic(city):
                return
            if not set_licensing_unit(city):
                return

            # ===== 5. 依模式設定查詢條件 =====
            if mode == "1":
                if not setup_lot_query(city, area, section, lot_number):
                    return
            elif mode == "2":
                # 🔥 先把 Chrome 切到「建築地址」表單 + 設好縣市區，
                #    使用者就能在 Chrome 上看到對的欄位再輸入地址
                if not set_query_type_to_address():
                    return
                if not set_bmbadd_city(city):
                    return
                if not set_bmbadd_district(area):
                    return

                # 然後再從使用者拿地址（先把 data.json 既有資訊當預設）
                base_addr = f"{city}{area}"
                if address:
                    print(f"\033[93m  data.json 已有地址：{address}\033[0m", flush=True)
                    user_addr = safe_input(f"請輸入完整建築地址（直接 Enter 採用上述地址）：").strip() or address
                else:
                    print(f"\033[93m  請輸入完整建築地址\033[0m", flush=True)
                    print(f"\033[93m  範例：{base_addr}廣東路87號\033[0m", flush=True)
                    user_addr = safe_input(f"請輸入完整建築地址（含 {base_addr}）：").strip()
                if not user_addr:
                    print("\033[91m未輸入地址，取消\033[0m", flush=True)
                    return

                # 取出 path-after-area 部分當 address 填欄位
                address_only = user_addr.replace(city, "").replace(area, "").strip()
                print(f"  地址（路名+號等）：{address_only}", flush=True)
                if not fill_address_fields(address_only):
                    return
            elif mode == "3":
                if not setup_manual_query():
                    return

            # ===== 6. 驗證碼 + 送出查詢 =====
            # 🔥 Phase 1：OCR 自動嘗試 3 輪
            # 🔥 Phase 2：若都失敗轉手動，無限次重試直到成功或使用者按 q 放棄
            status = "error"
            # 🔥 多筆模式時用 flag 跳到下一筆（不是整批中止）
            skip_this_lot = False

            # ----- Phase 1: OCR 三輪 -----
            for round_n in range(1, 4):
                print(f"\n[OCR 第 {round_n}/3 輪]", flush=True)

                code = ocr_captcha()
                if code and 4 <= len(code) <= 8:
                    print(f"  OCR 辨識結果：{code}", flush=True)
                    fill_captcha(code)
                else:
                    print(f"  OCR 無結果，刷新換一張", flush=True)
                    refresh_captcha()
                    continue

                click_query_button()
                status = wait_for_results()

                if status == "success":
                    break
                elif status == "no_data":
                    print("\033[93m查無資料\033[0m", flush=True)
                    skip_this_lot = True
                    break  # 跳出 OCR 內層迴圈
                elif status == "captcha_error":
                    print("  OCR 驗證碼錯，刷新重試", flush=True)
                    refresh_captcha()
                    continue
                else:
                    print("  狀態不明，刷新重試", flush=True)
                    refresh_captcha()
                    continue

            if skip_this_lot:
                continue  # 跳到外層 for _entry 的下一筆

            # ----- Phase 2: 手動模式（無限次重試）-----
            if status != "success":
                Y = "\033[93m"
                RST = "\033[0m"
                print("", flush=True)
                print(f"{Y}{'=' * 50}{RST}", flush=True)
                print(f"{Y}OCR 三輪都失敗，切換到【手動輸入】模式{RST}", flush=True)
                print(f"{Y}請在主程式下方輸入欄內輸入驗證碼後，按 Enter{RST}", flush=True)
                print(f"{Y}輸入錯誤會自動換一張、繼續嘗試；亦可輸入 q 放棄查詢{RST}", flush=True)
                print(f"{Y}{'=' * 50}{RST}", flush=True)

                manual_attempt = 0
                while True:
                    manual_attempt += 1
                    user_code = safe_input(f"[手動 第 {manual_attempt} 次] 驗證碼 = ").strip()
                    if not user_code or user_code.lower() == "q":
                        print("使用者放棄查詢", flush=True)
                        return  # 使用者主動放棄 → 中止整批

                    fill_captcha(user_code)
                    click_query_button()
                    status = wait_for_results()

                    if status == "success":
                        break
                    elif status == "no_data":
                        print("\033[93m查無資料\033[0m", flush=True)
                        skip_this_lot = True
                        break  # 跳出手動 while
                    elif status == "captcha_error":
                        print("\033[91m驗證碼錯誤，已換一張，請重新輸入\033[0m", flush=True)
                        refresh_captcha()
                        continue
                    else:
                        print("\033[91m查詢狀態不明，已換一張驗證碼，請重新輸入\033[0m", flush=True)
                        refresh_captcha()
                        continue

                if skip_this_lot:
                    continue  # 跳到外層 for _entry 的下一筆

            if status != "success":
                print("\033[91m查詢未成功，跳過此筆\033[0m", flush=True)
                continue

            # ===== 5. 切換每頁 200 筆，避免分頁 =====
            set_page_size_max()

            # ===== 6. 解析結果 + 詢問下載 =====
            results = parse_results()
            if not results:
                print("\033[93m沒有解析到任何結果列\033[0m", flush=True)
                continue  # 多筆模式下這筆無結果，跳到下一筆

            # 🔥 在 Chrome 結果頁把「使」字執照那列高亮（淺藍底+深藍粗體連結）
            highlight_usage_license_rows()
            print_results(results)

            # 存檔目錄：
            #   1) 跨城市查詢（mismatched_city）→ 通用「使用執照查詢結果」資料夾，避免污染 data.json 案件目錄
            #   2) 一般情況 → data.json 第一筆建立的案件資料夾（跟 land.py/hinet.py/luz.py 一致）
            if mismatched_city:
                folder_name = "使用執照查詢結果"
            else:
                first_entry = data_list[0]
                f_area = first_entry.get("area", "")
                f_section = first_entry.get("section", "")
                f_lot = first_entry.get("lot_number", "")
                folder_name = f"{f_area}{f_section}-{f_lot}"
            base_directory = get_work_folder(folder_name)
            os.makedirs(base_directory, exist_ok=True)
            print(f"  存檔目錄：{os.path.join(base_directory, '1.基本資料')}", flush=True)

            # 詢問下載哪些
            if len(results) == 1:
                print("\n只有一筆，自動下載...", flush=True)
                choices = [0]
            else:
                print(f"\n請輸入要下載的編號：", flush=True)
                print(f"  例：1,3,5     單個（多個用逗號分隔）", flush=True)
                print(f"  例:5-10       範圍（從 5 到 10 共 6 筆）", flush=True)
                print(f"  例：1,3,5-9   混合", flush=True)
                print(f"  輸入 a 或 all 全選；留空跳過下載", flush=True)
                user_input = safe_input().strip().lower()
                if not user_input:
                    print("已跳過下載", flush=True)
                    continue  # 多筆模式下，跳過這筆，繼續下一筆
                if user_input in ("a", "all"):
                    choices = list(range(len(results)))
                else:
                    try:
                        choices = _parse_choice_input(user_input, len(results))
                    except ValueError as e:
                        print(f"輸入格式錯誤（{e}），跳過下載", flush=True)
                        continue
                    if not choices:
                        print("沒有有效的編號，跳過下載", flush=True)
                        continue

            # ===== 7. 逐筆開「列印版」新分頁 → printToPDF → 關分頁 =====
            # 🔥 改用 save_via_print_view（直接到列印版 URL），不再走 modal screenshot 路徑。
            organ = LICENSING_UNIT_BY_CITY.get(city)
            pdf_save_dir = os.path.join(base_directory, "1.基本資料")
            print(f"\n開始下載 {len(choices)} 筆執照詳細...", flush=True)
            print(f"\033[93m  PDF 將存至：{pdf_save_dir}\033[0m", flush=True)
            success_count = 0
            for i, idx in enumerate(choices, 1):
                r = results[idx]
                mark = "★ " if r.get("is_usage") else ""
                print(f"\n[{i}/{len(choices)}] {mark}{r['license_no']}", flush=True)

                if save_via_print_view(r, organ, base_directory):
                    success_count += 1
                time.sleep(0.5)

            print(f"\n\033[92m✓ 此筆完成：成功下載 {success_count}/{len(choices)} 筆\033[0m", flush=True)

        # 多筆地號處理完所有結束訊息
        if len(selected_entries) > 1:
            print(f"\n\033[92m✓ 多筆查詢全部完成（共 {len(selected_entries)} 筆）\033[0m", flush=True)

    except KeyboardInterrupt:
        print("\n程式被使用者中斷", flush=True)
    except Exception as e:
        print(f"\033[91m執行過程錯誤：{e}\033[0m", flush=True)
        traceback.print_exc()
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass


def notify_main_program():
    print("子程式已完成執行", flush=True)
    sys.stdout.flush()


atexit.register(notify_main_program)


if __name__ == '__main__':
    try:
        main()
    except SystemExit:
        pass
    except Exception as e:
        print(f"程式異常：{e}", flush=True)
