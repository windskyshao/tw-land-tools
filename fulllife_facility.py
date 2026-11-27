# -*- coding: utf-8 -*-
"""重要設施查詢（富住通版，2026-11-27）

台灣地圖網「重要設施分析」(gbmap.py 用的)對方伺服器壞了(HTTP 500)。公司富住通後台「週遭重要設施查詢」
(https://www.fulllife.com.tw/com/facility.aspx)裡嵌的是圖霸 map8「不動產說明書重要設施查詢系統」，一樣有
地號座標定位、300 公尺範圍、重要設施＋嫌惡設施清單、地圖。

流程：用使用者存在「修改帳密 → 富住通」的帳密登入富住通(驗證碼 ddddocr 自動辨識)→ 開重要設施頁 →
用 data.json 的座標定位 → 下載報表(PDF)→ 存成「1.基本資料/03_重要設施-<區段地號>.pdf」
(檔名跟 gbmap.py 一樣，行銷簡報、連續模式照樣接得到)。

問題順序跟 gbmap.py 一樣：data.json 只有一筆直接做；多筆時問一次「要執行的項目編號」(連續模式回答 1)。
帳密只從 Windows 認證管理員讀，不寫檔、不印出來。
"""
import sys
import os

if getattr(sys, 'frozen', False):
    _d = os.path.dirname(os.path.abspath(__file__))
else:
    _d = os.path.dirname(os.path.abspath(__file__))
if _d not in sys.path:
    sys.path.insert(0, _d)

import base64
import http.cookiejar
import io
import json
import re
import time
import urllib.parse
import urllib.request

try:
    sys.stdout.reconfigure(line_buffering=True)
except Exception:
    pass

print("歡迎使用【重要設施查詢（富住通）】自動化小程式，模組載入中...", flush=True)

FZT = "https://www.fulllife.com.tw"
LOGIN = FZT + "/com/index.aspx"
VCODE = FZT + "/func/vcode.aspx"
FACILITY = FZT + "/com/facility.aspx"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/128.0 Safari/537.36")


def pdf_name(entry):
    return f"03_重要設施-{entry.get('area', '')}{entry.get('section', '')}-{entry.get('lot_number', '')}.pdf"


def latlng(entry):
    """data.json 的 coordinates「緯度,經度」→ (lat, lng)；沒有或格式錯回 None"""
    try:
        a, b = [float(x) for x in str(entry.get("coordinates", "")).replace(" ", "").split(",")[:2]]
    except Exception:
        return None
    if not (21 <= a <= 27 and 118 <= b <= 123):        # 臺澎金馬
        return None
    return a, b


def hidden_fields(html):
    d = {}
    for m in re.finditer(r'<input[^>]+type="hidden"[^>]+>', html):
        n = re.search(r'name="([^"]+)"', m.group(0))
        v = re.search(r'value="([^"]*)"', m.group(0))
        if n:
            d[n.group(1)] = v.group(1) if v else ""
    return d


def pick_entries(data_list, ask=input):
    """跟 gbmap.py 一樣：一筆直接做；多筆問一次(1 / 1,3 / 1-3 / all / q)"""
    if not data_list:
        print("data.json 中沒有資料。", flush=True)
        return []
    if len(data_list) == 1:
        print("data.json 中只有一組資料，將直接執行。", flush=True)
        return data_list
    print("\n===== 重要設施查詢 - 資料選擇 =====", flush=True)
    for i, d in enumerate(data_list, 1):
        print(f"[{i}] {d.get('city', '')}{d.get('area', '')} {d.get('section', '')} {d.get('lot_number', '')}"
              + (f" 地號， 座標: {d['coordinates']}" if d.get('coordinates') else ""), flush=True)
    print("\n請選擇要執行的項目編號（多個項目請用逗號分隔，輸入'all'執行全部，輸入'q'退出）：", flush=True)
    s = (ask() or "").strip().lower()
    if s in ("q", "quit", "exit"):
        return []
    if s == "all":
        return list(data_list)
    idx = []
    try:
        for part in s.split(","):
            part = part.strip()
            if "-" in part:
                a, b = map(int, part.split("-"))
                idx.extend(range(a, b + 1))
            elif part:
                idx.append(int(part))
    except ValueError:
        print("輸入格式錯誤。", flush=True)
        return []
    return [data_list[i - 1] for i in idx if 1 <= i <= len(data_list)]


class _Session:
    def __init__(self):
        self.cj = http.cookiejar.CookieJar()
        self.op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cj))

    def get(self, url):
        return self.op.open(urllib.request.Request(url, headers={"User-Agent": UA}), timeout=30)

    def post(self, url, data):
        req = urllib.request.Request(url, data=urllib.parse.urlencode(data).encode("utf-8"), headers={
            "User-Agent": UA, "Content-Type": "application/x-www-form-urlencoded", "Referer": LOGIN, "Origin": FZT})
        return self.op.open(req, timeout=30)


def _captcha(sess, ocr):
    """富住通驗證碼＝透明底黑數字 → 白底 → ddddocr 4 碼"""
    from PIL import Image
    im = Image.open(io.BytesIO(sess.get(VCODE).read()))
    if im.mode == "RGBA":
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(bg, im)
    buf = io.BytesIO()
    im.convert("RGB").save(buf, "PNG")
    return re.sub(r"\D", "", ocr.classification(buf.getvalue()) or "")


def login(account, pwd, tries=5):
    """登入富住通，回傳已登入的 session；失敗回 (None, 原因)"""
    import ddddocr
    ocr = ddddocr.DdddOcr(show_ad=False)
    why = ""
    for _ in range(tries):
        sess = _Session()
        html = sess.get(LOGIN).read().decode("utf-8", "replace")
        code = _captcha(sess, ocr)
        if len(code) != 4:
            time.sleep(1)
            continue
        d = hidden_fields(html)
        d.update({"Login1$UserName": account, "Login1$Password": pwd, "Login1$cknum": code, "Login1$Button1": "登入"})
        r = sess.post(LOGIN, d)
        if "comindex.aspx" in r.geturl() or "landmain" in r.geturl():
            return sess, ""
        h2 = r.read().decode("utf-8", "replace")
        why = (re.search(r"alert\(['\"]([^'\"]{2,80})['\"]\)", h2) or [None, ""])[1] or why
        time.sleep(1)
    return None, why or "驗證碼一直辨識失敗"


class CaptchaReader:
    """富住通驗證碼(4 碼數字)。有 ddddocr 用它(準)；同事電腦的執行環境沒有 ddddocr(11-27 連續模式失敗)
    → 改用執行環境內建的 RapidOCR(跟 hinet.py 一樣的退路)，圖先放大三倍、加白邊比較好認"""

    def __init__(self, log=print):
        self.dddd = self.rapid = None
        try:
            import ddddocr
            self.dddd = ddddocr.DdddOcr(show_ad=False)
        except Exception:
            from rapidocr_helper import create_rapidocr_reader
            self.rapid = create_rapidocr_reader(use_angle_cls=False, use_text_score=True, text_score=0.3,
                                                det_db_thresh=0.2, det_db_box_thresh=0.3, det_db_unclip_ratio=1.8)
            log("（驗證碼改用 RapidOCR 辨識）")

    def read(self, png_bytes):
        from PIL import Image, ImageOps
        im = Image.open(io.BytesIO(png_bytes))
        if im.mode in ("RGBA", "LA"):
            bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
            im = Image.alpha_composite(bg, im.convert("RGBA"))
        im = im.convert("RGB")
        if self.dddd is not None:
            buf = io.BytesIO()
            im.save(buf, "PNG")
            return re.sub(r"\D", "", self.dddd.classification(buf.getvalue()) or "")
        import numpy as np
        from rapidocr_helper import rapidocr_readtext
        big = ImageOps.expand(im.resize((im.width * 3, im.height * 3), Image.LANCZOS), border=30, fill="white")
        txt = "".join(rapidocr_readtext(self.rapid, np.array(big), detail=0) or [])
        txt = txt.translate(str.maketrans("OoQDIlLiSZzBgq", "00001111522899"))
        return re.sub(r"\D", "", txt)


def browser_login(driver, account, pwd, tries=10, log=print):
    """在瀏覽器裡登入富住通(驗證碼圖截下來辨識)；成功回 True"""
    from selenium.webdriver.common.by import By
    ocr = CaptchaReader(log)
    for i in range(tries):
        driver.get(LOGIN)
        img = driver.find_element(By.ID, "Login1_Image1")
        time.sleep(0.5)
        code = ocr.read(img.screenshot_as_png)
        if len(code) != 4:
            continue
        for k, v in (("Login1_UserName", account), ("Login1_Password", pwd), ("Login1_cknum", code)):
            el = driver.find_element(By.ID, k)
            el.clear()
            el.send_keys(v)
        driver.find_element(By.ID, "Login1_Button1").click()
        time.sleep(1.5)
        try:
            al = driver.switch_to.alert
            msg = al.text
            al.accept()
            log(f"（第 {i + 1} 次登入沒過：{msg}）")
            continue
        except Exception:
            pass
        if "comindex.aspx" in driver.current_url or "landmain" in driver.current_url:
            return True
    return False


# 下載報表時網頁會開新視窗並叫出列印對話框 → 先換掉 window.open，讓新視窗不要跳列印(改由程式直接存 PDF)
_NO_PRINT_JS = """
if (!window.__lmNoPrint) {
  window.__lmNoPrint = true;
  var _o = window.open;
  window.open = function() { var w = _o.apply(window, arguments); try { w.print = function(){}; } catch (e) {} return w; };
}
"""


FEW_AMENITIES = 3          # 300 公尺內重要設施少於這個數 → 改查 500 公尺


def _wait_list(driver):
    """等結果區出現、讀取動畫消失、清單筆數連續幾次一樣 → 回 (重要設施筆數, 嫌惡設施筆數)"""
    from selenium.webdriver.support.ui import WebDriverWait
    WebDriverWait(driver, 60).until(lambda d: d.execute_script(
        "var r=document.getElementById('result');return r && !r.classList.contains('d-none');"))
    last, same = None, 0
    for _ in range(40):
        n = driver.execute_script(
            "var l=document.getElementById('loading');"
            "if (l && l.offsetParent!==null) return null;"
            "return [document.getElementById('amenityList').rows.length,"
            "        document.getElementById('nimbyList').rows.length];")
        if n is not None and n == last:
            same += 1
            if same >= 3:
                break
        else:
            same = 0
        last = n
        time.sleep(0.5)
    return tuple(last) if last else (0, 0)


def pick_radius(amen300, try_radius, log=print):
    """要用哪個範圍(網站選項 300/500/750；清單其實列到 500/750/1000 公尺)。
    11-27 使用者：比較偏的(300 重要設施少於 3 筆)自動放大；放大到最大都只有嫌惡設施(重要設施 0 筆)→ 維持 300。
    try_radius(r) 會把網頁切到 r 並回傳那個範圍的重要設施筆數"""
    if amen300 >= FEW_AMENITIES:
        return 300
    best = 300 if amen300 == 0 else None
    n = amen300
    for r in (500, 750):
        log(f"重要設施只有 {n} 筆，改查 {r} 公尺…")
        n = try_radius(r)
        if n >= FEW_AMENITIES:
            return r
        if n > 0:
            best = r
    if best is None:
        return 300
    if best == 300:
        log("放大到最大範圍也只有嫌惡設施，維持 300 公尺")
    return best


def save_facility_pdf(driver, lat, lng, out_path, log=print):
    """在已登入的瀏覽器裡：重要設施頁 → 座標定位 → 下載報表(PDF)→ 存檔"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    driver.get(FACILITY)
    if "logout" in driver.current_url or "index.aspx" in driver.current_url:
        raise RuntimeError("富住通登入狀態失效")
    WebDriverWait(driver, 30).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frame")))
    WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.ID, "nav-coordinate-tab"))).click()
    box = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "coordinate")))
    box.clear()
    box.send_keys(f"{lat:.6f},{lng:.6f}")
    driver.find_element(By.ID, "btnPositioning").click()
    log(f"已用座標 {lat:.6f},{lng:.6f} 定位，等設施清單…")
    try:
        amen, nimby = _wait_list(driver)
    except Exception:                            # 網站偶爾慢、結果沒出來 → 再按一次查詢
        log("設施清單沒出來，再查一次…")
        driver.find_element(By.ID, "btnPositioning").click()
        amen, nimby = _wait_list(driver)

    def set_radius(r):
        driver.execute_script(f"var s=document.getElementById('radius');s.value='{r}';"
                              "s.dispatchEvent(new Event('change'));")
        time.sleep(1)
        return _wait_list(driver)

    radius = pick_radius(amen, lambda r: set_radius(r)[0], log)
    if driver.execute_script("return document.getElementById('radius').value") != str(radius):
        amen, nimby = set_radius(radius)       # 例：放大到最大都沒有重要設施 → 改回 300
    else:
        amen, nimby = _wait_list(driver)
    log(f"週邊 {radius} 公尺：重要設施 {amen} 筆、嫌惡設施 {nimby} 筆，下載報表…")
    driver.execute_script(_NO_PRINT_JS)
    driver.execute_script("var s=document.getElementById('downloadFormat');s.selectedIndex=2;"
                          "s.dispatchEvent(new Event('change'));")
    before = set(driver.window_handles)
    driver.find_element(By.ID, "download").click()
    WebDriverWait(driver, 90).until(lambda d: len(set(d.window_handles) - before) > 0)
    main = driver.current_window_handle
    driver.switch_to.window((set(driver.window_handles) - before).pop())
    WebDriverWait(driver, 60).until(lambda d: d.execute_script(
        "return document.readyState==='complete' && Array.from(document.images).every(i=>i.complete);"))
    time.sleep(1)
    pdf = driver.execute_cdp_cmd("Page.printToPDF", {"printBackground": True, "paperWidth": 8.27, "paperHeight": 11.69,
                                                     "marginTop": 0.4, "marginBottom": 0.4, "marginLeft": 0.4,
                                                     "marginRight": 0.4, "preferCSSPageSize": False})
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "wb") as f:
        f.write(base64.b64decode(pdf["data"]))
    driver.close()
    driver.switch_to.window(main)
    return out_path


def chrome_options(headless=True):
    """不跳「要儲存密碼嗎」(11-27 使用者截圖)；預設無頭(背景跑，不開視窗、不搶滑鼠鍵盤)"""
    from selenium import webdriver
    o = webdriver.ChromeOptions()
    o.add_experimental_option("prefs", {"credentials_enable_service": False,
                                        "profile.password_manager_enabled": False,
                                        "profile.password_manager_leak_detection": False})
    o.add_argument("--disable-features=PasswordLeakDetection,PasswordManagerOnboarding")
    if headless:
        o.add_argument("--headless=new")
        o.add_argument("--window-size=1400,1000")
        o.add_argument("--force-device-scale-factor=1")
    return o


def main():
    from base_dir_helper import get_data_json_path, get_work_folder
    try:
        from credential_helper import get_credentials
        account, pwd = get_credentials("fulllife")
    except Exception:
        account = pwd = None
    if not account or not pwd:
        print("\033[91m還沒設定富住通帳密：請到主程式「修改帳密」→「富住通」填好再查。\033[0m", flush=True)
        return 2
    try:
        with open(get_data_json_path(), "r", encoding="utf-8") as f:
            data_list = json.load(f)
    except Exception as e:
        print(f"讀不到 data.json：{e}", flush=True)
        return 2
    picked = pick_entries(data_list)
    if not picked:
        return 0
    first = data_list[0]
    save_dir = os.path.join(get_work_folder(f"{first.get('area', '')}{first.get('section', '')}-{first.get('lot_number', '')}"),
                            "1.基本資料")
    from webdriver_helper import create_chrome_driver
    driver = create_chrome_driver(chrome_options(headless="--show" not in sys.argv))
    ok = 0
    try:
        print("登入富住通中（驗證碼自動辨識）…", flush=True)
        if not browser_login(driver, account, pwd, log=lambda m: print(m, flush=True)):
            print("\033[91m富住通登入失敗。請確認「修改帳密 → 富住通」的帳號密碼。\033[0m", flush=True)
            return 1
        print("✓ 富住通登入成功", flush=True)
        for i, e in enumerate(picked, 1):
            ll = latlng(e)
            label = f"{e.get('area', '')}{e.get('section', '')} {e.get('lot_number', '')}"
            if not ll:
                print(f"\033[93m[{label}] data.json 沒有座標，跳過（請先用地籍便民或萬用查詢建案件）\033[0m", flush=True)
                continue
            print(f"\n處理項目 {i}/{len(picked)}: {label}", flush=True)
            out = os.path.join(save_dir, pdf_name(e))
            try:
                save_facility_pdf(driver, ll[0], ll[1], out, log=lambda m: print(m, flush=True))
                print(f"\033[92m✓ 已存：{out}\033[0m", flush=True)
                ok += 1
            except Exception as ex:
                print(f"\033[91m[{label}] 重要設施查詢失敗：{ex}\033[0m", flush=True)
    finally:
        try:
            driver.quit()
        except Exception:
            pass
    print("重要設施查詢完成" if ok else "重要設施查詢沒有成功的項目", flush=True)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
