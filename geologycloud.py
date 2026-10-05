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
import json
import time
import ctypes
import keyboard  # 用於系統級鍵盤事件
from PIL import Image
from fpdf import FPDF
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from webdriver_helper import create_chrome_driver, verify_and_fix_chrome_window, apply_page_zoom

# 🔥 基準目錄設定（data.json 和工作資料夾的位置）
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

# 清除終端機並顯示歡迎訊息
os.system('cls')
print("歡迎使用【地質雲】自動化小程式，模組載入中...", flush=True)

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
geology_window_width = 1024
geology_window_height = 1024
geology_window_x = 0
geology_window_y = 0

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

            geology_window_width = chrome_width_logical
            geology_window_height = chrome_height_logical
            geology_window_x = chrome_x_logical
            geology_window_y = chrome_y_logical
except Exception as e:
    pass

# 讀取 data.json 文件
with open(get_data_json_path(), 'r', encoding='utf-8') as f:
    data_list = json.load(f)

if not data_list:
    print("data.json 是空的，請檢查資料。", flush=True)
    exit()

# 取得第一組資料，建立基礎目錄結構
first_entry = data_list[0]
base_region = first_entry['area']
base_section = first_entry['section']
base_lot_number = first_entry['lot_number']
# 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
base_directory = get_work_folder(f"{base_region}{base_section}-{base_lot_number}")
basic_data_dir = os.path.join(base_directory, "1.基本資料")
png_dir = os.path.join(basic_data_dir, "png")

# 建立目錄結構
os.makedirs(png_dir, exist_ok=True)
print(f"基礎目錄結構已建立於 {base_directory}", flush=True)

# 設定 WebDriver
options = webdriver.ChromeOptions()
# options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36")
# 🔥 不在啟動參數設定視窗大小，改為啟動後再調整（避免 DPI 改變時崩潰）
# options.add_argument("--window-size=1024,1024")
# options.add_argument("--window-position=0,0")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# 使用 ChromeDriverManager 安裝 ChromeDriver 並應用選項
driver = create_chrome_driver(options=options)

# 🔥 用 CDP 直接授予地理位置權限（比 prefs 可靠；新版 Chrome 的 prefs 常失效）
try:
    driver.execute_cdp_cmd("Browser.grantPermissions", {"permissions": ["geolocation"]})
    print("[權限] 已自動授予地理位置權限", flush=True)
except Exception as _perm_e:
    print(f"[權限] 自動授予地理位置失敗（不影響流程）: {_perm_e}", flush=True)

# 🔥 啟動後立即設定視窗大小和位置
try:
    driver.set_window_size(geology_window_width, geology_window_height)
    driver.set_window_position(geology_window_x, geology_window_y)
except Exception as e:
    print(f"[WARNING] 視窗設定失敗: {e}，繼續執行")

# 指定要打開的網址
# 🔥 2026-10-05 政府改版：舊的地質雲「土壤液化潛勢圖」**即日起停止服務**，
#    網站公告：「配合整體風險評估計畫推動進度，即日起停止土壤液化潛勢圖查詢服務，
#    接下來將為您重新導引到【環境地質雲】系統查詢」
#    新站的液化圖資是 2026-09-24 09:00 排程上線的（寫在它 panel1.js 裡）。
url = "https://envgeology.gsmma.gov.tw/landslide"

# 🔥 新站是 Leaflet 地圖，但地圖物件藏在模組內部、拿不到。
#    在網頁任何程式執行**之前**先攔截 L.Map 的初始化，把實例存起來，
#    之後就能直接命令地圖定位（比打搜尋框穩定很多）。
try:
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": """
      (function(){
        var tries = 0;
        var iv = setInterval(function(){
          tries++;
          try {
            if (window.L && L.Map && L.Map.prototype && !L.Map.prototype.__hooked) {
              var origInit = L.Map.prototype.initialize;
              L.Map.prototype.initialize = function(){
                var r = origInit.apply(this, arguments);
                try { window.__LMAPS = window.__LMAPS || []; window.__LMAPS.push(this); } catch(e){}
                return r;
              };
              L.Map.prototype.__hooked = true;
              clearInterval(iv);
            }
          } catch(e){}
          if (tries > 400) clearInterval(iv);
        }, 25);
      })();
    """})
except Exception as _e:
    print(f"[WARNING] 地圖攔截設定失敗（{_e}），將改用搜尋框退路", flush=True)

time.sleep(2)
driver.get(url)

# 🔥 一載入就先關掉「開場導覽」，別等到後面才關。
#    它是一層蓋住整個畫面的麵罩，不關的話要等很久才進得了系統。
#    官方按鈕：<div class="guidMask__closeBuidBtn" onclick="closeGuidMask(this)">關閉導覽(Close)</div>
_intro_deadline = time.time() + 25
while time.time() < _intro_deadline:
    try:
        _n = driver.execute_script("""
          var n = 0;
          var b = document.querySelector('.guidMask__closeBuidBtn');
          if (b) { try { b.click(); n++; } catch(e){} }
          if (!n && typeof closeGuidMask === 'function') {
            try { closeGuidMask(b || document.body); n++; } catch(e){}
          }
          document.querySelectorAll('.guidMask, [class*="guidMask"], .introjs-overlay, .introjs-helperLayer, .introjs-tooltipReferenceLayer, .introjs-tooltip')
            .forEach(function(e){ try { e.remove(); n++; } catch(e2){} });
          return n;
        """)
        if _n:
            print(f"已關閉開場導覽（{_n} 個元素）。", flush=True)
            break
    except Exception:
        pass
    time.sleep(0.5)
# 🔥 網頁載入後重新設定視窗大小（避免被網站重置）
try:
    driver.set_window_size(geology_window_width, geology_window_height)
    driver.set_window_position(geology_window_x, geology_window_y)
except:
    pass

# 🔥 DPI 自動修正（只在 DPI 不同步時動作，正常 PC 無影響）
verify_and_fix_chrome_window(driver)

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
            zoom_count = 2  # 175%：按 2 次
        elif current_dpi >= 1.5:
            zoom_count = 1  # 150%：按 1 次
        elif current_dpi >= 1.25 and logical_width < 1200:
            zoom_count = 1  # 小螢幕 + 125%：按 1 次
        else:
            zoom_count = 0  # 100%、125%（大螢幕）：不縮放
except Exception:
    pass

# 🔥 使用者在「Chrome縮放設定」面板調過的值優先（zoom_config.json）；沒調過才用上面的 DPI 預設
try:
    _zcp = os.path.join(BASE_DIR, 'zoom_config.json')
    if os.path.exists(_zcp):
        _zov = json.load(open(_zcp, encoding='utf-8')).get('overrides', {}).get('geologycloud')
        if _zov is not None:
            zoom_count = int(_zov)
except Exception:
    pass

apply_page_zoom(driver, zoom_count, key='geologycloud', base_dir=BASE_DIR)  # 共用：拉前景→Ctrl+0→縮放→驗證
if zoom_count > 0:
    time.sleep(1)  # 縮放後等待頁面穩定

driver.implicitly_wait(10)

def _dismiss_extra_popups(driver, rounds=3):
    """關掉「多出來的公告彈窗」(例：政府臨時維護公告「北部六縣市…更新」)，避免擋住後續操作。
    先按 ESC，再點掉任何可見的『知道了/確定/關閉/OK/×/data-dismiss』鈕；全程 try/except，沒有就跳過、不弄壞。"""
    from selenium.webdriver.common.keys import Keys
    import time as _t
    for _r in range(rounds):
        _closed = False
        try:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        except Exception:
            pass
        try:
            _xp = ("//button[normalize-space()='我知道了' or normalize-space()='知道了' or normalize-space()='確定'"
                   " or normalize-space()='關閉' or normalize-space()='OK' or contains(@class,'close')"
                   " or contains(@class,'btn-close') or @data-dismiss='modal' or @data-bs-dismiss='modal']"
                   " | //a[normalize-space()='我知道了' or normalize-space()='知道了' or normalize-space()='確定'"
                   " or normalize-space()='關閉' or contains(@class,'close')]"
                   " | //*[contains(@class,'modal') or contains(@class,'popup') or contains(@class,'dialog')]"
                   "//*[@aria-label='Close' or @aria-label='close' or normalize-space()='×' or normalize-space()='✕']")
            for _b in driver.find_elements(By.XPATH, _xp):
                try:
                    if _b.is_displayed():
                        driver.execute_script("arguments[0].click();", _b)
                        _closed = True
                        print("已關閉多餘公告彈窗。", flush=True)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            _n = driver.execute_script(
                "var n=0;"
                "document.querySelectorAll('.sweet-alert button.confirm, .swal2-confirm, .swal-button--confirm').forEach(function(b){try{b.click();n++;}catch(e){}});"
                "document.querySelectorAll('.sweet-overlay, .sweet-alert, .swal-overlay, .swal2-container, .modal-backdrop').forEach(function(e){try{e.style.display='none';if(e.parentNode)e.parentNode.removeChild(e);n++;}catch(e2){}});"
                "return n;")
            if _n:
                _closed = True
                print("已清除遮罩/公告彈窗(sweet-overlay 等)。", flush=True)
        except Exception:
            pass
        _t.sleep(0.4)
        if not _closed:
            break


# 舊站的「我知道了」(btn_Soillique_OK) 已隨舊站停止服務而移除，
# 新站改用 .guidMask__closeBuidBtn（已在載入時就關掉）。

# 🔥 關掉可能多出來的公告彈窗(例：政府臨時維護公告)，避免擋住後續搜尋操作
_dismiss_extra_popups(driver)


# ===== 新站（環境地質雲）初始化 =====
def _close_intro(driver):
    """關掉開場導覽（會蓋住整個畫面）。"""
    try:
        n = driver.execute_script("""
          var n = 0;
          var b = document.querySelector('.guidMask__closeBuidBtn');
          if (b) { try { b.click(); n++; } catch(e){} }
          document.querySelectorAll('.guidMask, [class*="guidMask"], .introjs-overlay, .introjs-helperLayer, .introjs-tooltipReferenceLayer, .introjs-tooltip')
            .forEach(function(e){ try { e.remove(); n++; } catch(e2){} });
          return n;""")
        if n:
            print(f"已關閉開場導覽（{n} 個元素）。", flush=True)
    except Exception:
        pass


def _enable_liquefaction_layer(driver, tries=3):
    """打開「土壤液化潛勢區域」圖層（新站預設是關的）。"""
    for _i in range(tries):
        try:
            r = driver.execute_script("""
              var cb = document.querySelector('#cgs_4c_liquefaction_0');
              if (cb) { if (!cb.checked) { cb.click(); } return cb.checked; }
              var bt = document.querySelector('#cgs_4c_liquefaction');
              if (bt) { bt.click(); return 'clicked'; }
              return null;""")
            if r:
                print(f"已開啟【土壤液化潛勢區域】圖層。", flush=True)
                time.sleep(3)
                return True
        except Exception:
            pass
        time.sleep(2)
    print("[警告] 找不到土壤液化圖層開關，截圖可能沒有液化色塊", flush=True)
    return False


def _parse_latlon(coord_text):
    """'22.57152,120.37395' -> (22.57152, 120.37395)。順序是緯度,經度。"""
    parts = [x.strip() for x in str(coord_text or '').replace('，', ',').split(',')]
    if len(parts) < 2:
        raise ValueError(f"座標格式不對：{coord_text!r}")
    a, b2 = float(parts[0]), float(parts[1])
    # 台灣：緯度 21~26、經度 118~123。若寫反了自動交換
    if a > 100 and b2 < 40:
        a, b2 = b2, a
    return a, b2


def _tiles_pending(driver):
    """還沒載完的圖磚數。
    直接看 <img> 的 complete / naturalWidth，比看 class 可靠
    （政府的 WMTS 圖層不一定會帶 .leaflet-tile-loaded）。"""
    try:
        return driver.execute_script("""
          var imgs = document.querySelectorAll('.leaflet-tile-pane img');
          var p = 0;
          for (var i = 0; i < imgs.length; i++) {
            var im = imgs[i];
            if (!im.complete || im.naturalWidth === 0) p++;
          }
          return p;
        """)
    except Exception:
        return 0


def _reload_failed_tiles(driver):
    """把載失敗的圖磚強制重新要一次。
    實測：灰色空白塊是 **NLSC 國土測繪的底圖**沒載起來，
    不是「該區沒有液化資料」（兩者很像但完全不同）：
    單獨去抓那些圖磚都是 HTTP 200 正常 JPEG，只是政府圖磚伺服器
    被同時要太多張時會掉包。重設 src 就會再要一次。"""
    try:
        return driver.execute_script("""
          var n = 0;
          document.querySelectorAll('.leaflet-tile-pane img').forEach(function(im){
            if (!im.complete || im.naturalWidth === 0) {
              var s = im.src;
              if (s) { im.src = ''; im.src = s; n++; }
            }
          });
          return n;
        """)
    except Exception:
        return 0


def _nudge_map(driver):
    """輕輕推一下地圖再推回來，迫使 Leaflet 重新要沒載到的圖磚。"""
    try:
        driver.execute_script("""
          var maps = window.__LMAPS || [];
          if (!maps.length) return;
          var m = maps[0];
          try { m.panBy([2, 2], {animate:false}); m.panBy([-2, -2], {animate:false}); } catch(e){}
          try { m.invalidateSize(); } catch(e){}
        """)
    except Exception:
        pass


def _wait_tiles(driver, timeout=12):
    """等地圖圖磚載完再截圖，但**卡住就立刻放棄**，不要讓使用者乾等。

    停滞偵測：每 0.5 秒看一次還沒載完的數量，
    若連續約 2.5 秒數量沒減少，就是「要不到了」（該區沒圖或政府伺服器掉包），
    繼續等只是浪費時間 → 馬上走人，讓外面把它藏成白底就好。"""
    last = None
    stall = 0
    end = time.time() + timeout
    while time.time() < end:
        n = _tiles_pending(driver)
        if n == 0:
            time.sleep(0.8)
            return True
        if last is not None and n >= last:
            stall += 1
            if stall >= 5:          # 約 2.5 秒沒進展 -> 不等了
                break
        else:
            stall = 0
        last = n
        time.sleep(0.5)

    # 只試一輪重要（真的沒圖的話再試也沒用，以前試 3 輪會多等幾十秒）
    _n = _reload_failed_tiles(driver)
    if _n:
        _end2 = time.time() + 5
        while time.time() < _end2:
            if _tiles_pending(driver) == 0:
                time.sleep(0.8)
                return True
            time.sleep(0.5)
    return False


def _map_center(driver):
    """回傳目前地圖中心 (lat, lon)，拿不到回傳 None。"""
    try:
        c = driver.execute_script("""
          var maps = window.__LMAPS || [];
          if (!maps.length) return null;
          var m = maps[0];
          var c = m.getCenter();
          return [c.lat, c.lng, m.getZoom()];
        """)
        return c
    except Exception:
        return None


def _search_coord(driver, coord_text, lat, lon):
    """用網站**官方搜尋框**輸入座標定位（跟手動操作完全一樣）。

    ⚠️ 關鍵：輸入完按 Enter **還不會定位**，它只是在搜尋框下方
       浮出一個「[坐標]22.57138,120.37363」的**可點擊按鈕**
       （`.search_result_button`），**點下去才會飛過去**，
       而且會在定位點標上官方的座標圖標與資訊框。
    回傳 True 表示地圖真的移到目標附近。"""
    try:
        from selenium.webdriver.common.keys import Keys
        inp = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='請輸入']")
        driver.execute_script("arguments[0].value = '';", inp)
        try:
            inp.clear()
        except Exception:
            pass
        inp.send_keys(str(coord_text))
        time.sleep(0.6)
        inp.send_keys(Keys.ENTER)

        # 等「座標候選按鈕」浮出來並點它（這步不做就不會定位）
        _clicked = None
        _end = time.time() + 8
        while time.time() < _end:
            _clicked = driver.execute_script("""
              var b = document.querySelector('.search_result_button');
              if (b) { b.click(); return (b.textContent || '').trim(); }
              return null;
            """)
            if _clicked:
                print(f"已點選座標候選：{_clicked}", flush=True)
                break
            time.sleep(0.4)

        # 等它真的把地圖移過去
        _end2 = time.time() + 10
        while time.time() < _end2:
            c = _map_center(driver)
            if c and abs(c[0] - lat) < 0.01 and abs(c[1] - lon) < 0.01:
                return True
            time.sleep(0.5)
    except Exception as _e:
        print(f"[搜尋框] 無法使用（{type(_e).__name__}）", flush=True)
    return False


def _goto_coord(driver, lat, lon, zoom=17, marker=True):
    """直接命令 Leaflet 地圖定位，並在中心畫一個紅圈標記。
    紅圈用 circleMarker（CSS 畫的），不用圖標圖片，免得破圖。"""
    return driver.execute_script("""
      var lat = arguments[0], lon = arguments[1], z = arguments[2], mk = arguments[3];
      var maps = window.__LMAPS || [];
      if (!maps.length) return {ok:false, msg:'沒攔到地圖物件'};
      var map = maps[0];
      for (var i = 0; i < maps.length; i++) {
        if (maps[i]._container && maps[i]._container.offsetParent) { map = maps[i]; break; }
      }
      // ⚠️ 一定要 animate:false：點完座標候選後網站會跑飛行動畫，
      //    帶動畫的 setView 會被它吃掉，結果停在 zoom 13（整個高雄）。
      map.setView([lat, lon], z, {animate: false});
      try {
        if (window.__LMK) { map.removeLayer(window.__LMK); window.__LMK = null; }
        if (mk && window.L && L.circleMarker) {
          window.__LMK = L.circleMarker([lat, lon], {
            radius: 10, color: '#d32f2f', weight: 3,
            fillColor: '#ff5252', fillOpacity: 0.45
          }).addTo(map);
        }
      } catch(e){}
      try { map.invalidateSize(); } catch(e){}
      return {ok:true, zoom:map.getZoom(),
              center:[map.getCenter().lat, map.getCenter().lng]};
    """, lat, lon, zoom, bool(marker))


_close_intro(driver)
time.sleep(1)
_liq_ok = _enable_liquefaction_layer(driver)
_map_hooked = bool(driver.execute_script("return (window.__LMAPS||[]).length;"))
print(f"[地圖] 攔截到 {int(_map_hooked)} 個地圖物件", flush=True)

# 列表來收集所有 PNG 路徑和標識符
all_png_files = []
all_identifiers = []
all_data_entries = []  # 儲存所有處理過的資料項目

# 遍歷 data.json 中的每一組數據
for index, data in enumerate(data_list):
    city = data['city']
    area = data['area']
    section = data['section']
    lot_number = data['lot_number']
    coordinates = data['coordinates']

    print(f"處理資料 {index+1}/{len(data_list)}: 城市 = {city}, 區域 = {area}, 段 = {section}, 地號 = {lot_number}, 座標 = {coordinates}", flush=True)

    try:
        # 註：新站沒有舊站那種 sweet-alert 公告，而且通用關閉會點到任何 .close，
        #    可能誤關掉**圖例**或**座標資訊框**，所以每筆查詢前不再跑。

        # 🔥 新站改用「直接命令地圖定位」，不再打搜尋框
        #    （舊站是輸入座標 -> 點放大鏡 -> 點結果清單；
        #      新站搜尋框對純座標不會定位，改攔 Leaflet 地圖物件直接 setView）
        _lat, _lon = _parse_latlon(coordinates)

        # 🔥 先清掉上一筆留下的座標圖標（資訊框裡的「移除」鈕），
        #    不清的話第二筆截圖會同時出現兩個圖標，分不出哪個是這筆。
        try:
            _rm = driver.execute_script("""
              var n = 0;
              document.querySelectorAll('button, a, div, span').forEach(function(e){
                var t = (e.innerText || '').trim();
                if (t === '移除' && e.offsetParent && e.children.length === 0) {
                  try { e.click(); n++; } catch(err){}
                }
              });
              return n;
            """)
            if _rm:
                time.sleep(0.8)
        except Exception:
            pass

        # (1) 先走官方搜尋框（會有官方座標圖標與資訊框，跟手動操作一樣）
        _searched = _search_coord(driver, coordinates, _lat, _lon)
        if _searched:
            print(f"已用搜尋框定位到 {coordinates}。", flush=True)
        else:
            print("[提示] 搜尋框沒把地圖移過去，改由程式定位", flush=True)

        # (2) 🔥 不論搜尋框成功與否，**最後都再校正一次比例尺**。
        #     因為點完座標候選後網站有自己的飛行動畫，會**覆蓋掉我們設的縮放**，
        #     曾出現同一批的兩張圖一張 1:4,168、另一張 1:66,692（整個高雄）的狀況。
        #     有官方圖標就不重複畫紅圈。
        time.sleep(1.5)
        _r = _goto_coord(driver, _lat, _lon, zoom=17, marker=not _searched)
        if not _r or not _r.get("ok"):
            print(f"[警告] 比例尺校正失敗：{_r}", flush=True)
        else:
            print(f"已校正到 zoom={_r.get('zoom')}。", flush=True)
        time.sleep(1.5)

        # 等圖磚載完再截（不等會截到一半空白）
        _wait_tiles(driver, timeout=12)
        # 實在要不到的底圖圖磚 -> 藏起來讓它變白底，別留難看的灰塊
        try:
            _h = driver.execute_script("""
              var n = 0;
              document.querySelectorAll('.leaflet-tile-pane img').forEach(function(im){
                if (!im.complete || im.naturalWidth === 0) { im.style.visibility = 'hidden'; n++; }
              });
              return n;
            """)
            if _h:
                print(f"[圖磚] {_h} 塊底圖要不到（該區沒圖或政府伺服器掉包），已藏成白底", flush=True)
        except Exception:
            pass

        filename_base = f"09_地質雲-{area}{section}-{lot_number}"
        screenshot_path = os.path.join(png_dir, f"{filename_base}.png")
        driver.save_screenshot(screenshot_path)
        print(f"[93m已截圖並保存為 {screenshot_path}[0m", flush=True)

        # 收集 PNG 路徑和標識符
        all_png_files.append(screenshot_path)
        all_identifiers.append(f"{area}{section}-{lot_number}")
        all_data_entries.append(data)

    except TimeoutException:
        print(f"處理座標 {coordinates} 時發生超時錯誤。", flush=True)

# 關閉瀏覽器
driver.quit()
print("查詢完成，關閉瀏覽器", flush=True)

# 合併所選 PNG 成為一個 PDF
def create_combined_pdf(png_files, identifiers, pdf_dir):
    if not png_files:
        print("沒有找到任何 PNG 檔案，無法生成 PDF。", flush=True)
        return

    # 使用新命名規則合併標識符
    area_section_dict = {}
    for identifier in identifiers:
        parts = identifier.split('-')
        region = parts[0]
        lot = '-'.join(parts[1:])  # 將後面的部分重新合併成地號

        if region in area_section_dict:
            area_section_dict[region].append(lot)
        else:
            area_section_dict[region] = [lot]

    # 建立 PDF 檔名，將每個區域的地號以「、」分隔，並用「+」連接不同地區地段
    pdf_filename = "09_地質雲-" + "+".join(
        [f"{region}-{','.join(lots)}" for region, lots in area_section_dict.items()]
    ) + ".pdf"

    pdf_path = os.path.join(pdf_dir, pdf_filename)

    # 🔥 檢查檔案是否已存在且被占用
    if os.path.exists(pdf_path):
        try:
            # 嘗試刪除舊檔案
            os.remove(pdf_path)
        except PermissionError:
            print(f"\033[91m警告：PDF 檔案正在被其他程式使用，請關閉該檔案後按 Enter 繼續...\033[0m", flush=True)
            print(f"檔案路徑：{pdf_path}", flush=True)
            input()
            # 再次嘗試刪除
            try:
                os.remove(pdf_path)
            except Exception as e:
                print(f"仍無法刪除舊檔案: {e}", flush=True)
                return None

    try:
        pdf = FPDF(orientation='L', unit='mm', format='A4')  # 橫向 PDF
        for png in png_files:
            pdf.add_page()
            pdf.image(png, x=0, y=0, w=297, h=210)  # 調整大小至 A4 橫向

        # 嘗試寫入 PDF，如果失敗則重試
        max_retries = 3
        for attempt in range(max_retries):
            try:
                pdf.output(pdf_path)
                print(f"\033[93mPDF 已經儲存至 {pdf_path}\033[0m", flush=True)
                return pdf_path
            except PermissionError as e:
                if attempt < max_retries - 1:
                    print(f"\033[91m寫入 PDF 失敗（嘗試 {attempt + 1}/{max_retries}），請關閉該檔案後按 Enter 重試...\033[0m", flush=True)
                    print(f"檔案路徑：{pdf_path}", flush=True)
                    input()
                else:
                    print(f"將 PNG 轉成 PDF 時發生錯誤: {e}", flush=True)
                    print(f"\033[91m提示：請檢查是否有 PDF 閱讀器正在開啟此檔案\033[0m", flush=True)
                    return None
    except Exception as e:
        print(f"將 PNG 轉成 PDF 時發生錯誤: {e}", flush=True)
        return None

# 詢問使用者要合併哪些地號並生成合併的 PDF
def ask_user_for_selection(all_data_entries, all_png_files, all_identifiers, basic_data_dir):
    # 如果只有一筆資料，直接生成 PDF
    if len(all_data_entries) == 1:
        print("\n只有一筆資料，直接生成 PDF...", flush=True)
        return create_combined_pdf(all_png_files, all_identifiers, basic_data_dir)

    # 多筆資料時才詢問使用者
    print("\n請選擇要合併為 PDF 的地號：", flush=True)
    print("0. 全部選取", flush=True)
    
    # 預設選擇最後一個
    default_selection = len(all_data_entries) - 1
    
    # 顯示所有可選項目
    for i, data in enumerate(all_data_entries):
        area = data['area']
        section = data['section']
        lot_number = data['lot_number']
        default_mark = " (預設)" if i == default_selection else ""
        print(f"{i+1}. {area}{section}-{lot_number}{default_mark}", flush=True)

    print("\n請輸入您要選擇的項目編號（多個選項以逗號分隔，例如：1,3,5），或直接按 Enter 使用預設選擇：", flush=True)
    selection = input().strip()
    
    selected_indices = []
    
    if not selection:  # 使用者直接按 Enter，選擇預設項目
        selected_indices = [default_selection]
        print(f"使用預設選擇：{all_data_entries[default_selection]['area']}{all_data_entries[default_selection]['section']}-{all_data_entries[default_selection]['lot_number']}", flush=True)
    elif selection == "0":  # 全部選擇
        selected_indices = list(range(len(all_data_entries)))
        print("已選擇所有項目", flush=True)
    else:
        try:
            # 解析用戶輸入，將編號轉換為索引（減 1）
            selected_indices = [int(idx.strip()) - 1 for idx in selection.split(',')]
            # 檢查索引是否有效
            for idx in selected_indices:
                if idx < 0 or idx >= len(all_data_entries):
                    print(f"警告：編號 {idx+1} 超出範圍，已忽略")
                    selected_indices.remove(idx)
            
            if not selected_indices:  # 如果沒有有效選擇，則使用預設
                selected_indices = [default_selection]
                print(f"沒有有效選擇，使用預設選擇：{all_data_entries[default_selection]['area']}{all_data_entries[default_selection]['section']}-{all_data_entries[default_selection]['lot_number']}")
            else:
                print(f"已選擇 {len(selected_indices)} 個項目")
        except ValueError:
            # 輸入格式錯誤，使用預設選擇
            selected_indices = [default_selection]
            print(f"輸入格式不正確，使用預設選擇：{all_data_entries[default_selection]['area']}{all_data_entries[default_selection]['section']}-{all_data_entries[default_selection]['lot_number']}")
    
    # 獲取選定項目的 PNG 和標識符
    selected_png_files = [all_png_files[i] for i in selected_indices]
    selected_identifiers = [all_identifiers[i] for i in selected_indices]
    
    # 生成 PDF
    return create_combined_pdf(selected_png_files, selected_identifiers, basic_data_dir)

# 呼叫詢問使用者並生成合併的 PDF
pdf_path = ask_user_for_selection(all_data_entries, all_png_files, all_identifiers, basic_data_dir)

def notify_main_program():
    print("地質雲已完成執行", flush=True)
    # if pdf_path:
        # print(f"生成的 PDF 文件路徑：{pdf_path}", flush=True)

if __name__ == "__main__":
    try:
        # 子程式的主要邏輯
        # 在這裡執行子程式的操作，例如表單填寫和資料處理
        # print("執行地質雲主要邏輯", flush=True)
        pass
    except Exception as e:
        print(f"地質雲執行過程中發生錯誤: {e}", flush=True)
    finally:
        # 在結尾通知主程式
        notify_main_program()