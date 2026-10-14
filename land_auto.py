"""地籍便民系統：自動截圖(讀 data.json 逐筆定位 → 截圖存 PNG + PDF，不改 data.json)

用法：land_auto.py --ver old|new [--base-dir ...]
  old = 官方舊版 easymap.moi.gov.tw/Index(跟手動版 land.py 同一個網頁)
  new = 官方最新版 easymap.moi.gov.tw/Z10Web/Normal(2026-08 改版)
Chrome 視窗大小/位置、頁面縮放、使用者的縮放設定都沿用 land.py(import 時就算好)。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))   # python_embedded 有 ._pth，要自己加才 import 得到 land

import argparse
import json
import re
import time

import land as L   # import 時會算好 Chrome 視窗大小/位置與 options(跟手動版一樣)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
from fpdf import FPDF

OLD_URL = "https://easymap.moi.gov.tw/Index"
NEW_URL = "https://easymap.moi.gov.tw/Z10Web/Normal"


def log(s):
    print(s, flush=True)


def _norm(s):
    return str(s or "").replace("台", "臺").replace(" ", "").strip()


def _pick(select_el, want, part=False, strip_code=False):
    """下拉選單依文字選：完全相同優先，其次去掉「(1050)」代碼後相同，再其次包含(行政區少「區」字也行)"""
    sel = Select(select_el)
    w = _norm(want)
    opts = [(o, _norm(o.text)) for o in sel.options]
    for o, t in opts:
        if t == w:
            sel.select_by_visible_text(o.text)
            return o.text
    if strip_code:
        for o, t in opts:
            if t.split(")")[-1] == w:
                sel.select_by_visible_text(o.text)
                return o.text
    if part:
        hits = [o for o, t in opts if t and (w in t or t in w) and len(t) >= 2]
        if len(hits) == 1:
            sel.select_by_visible_text(hits[0].text)
            return hits[0].text
    return None


def _wait_options(d, el_id, secs=15):
    WebDriverWait(d, secs).until(lambda x: len(Select(x.find_element(By.ID, el_id)).options) > 1)


def load_entries():
    try:
        with open(L.get_data_json_path(), encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        log(f"讀取 data.json 失敗：{e}")
        return []
    out, seen = [], set()
    for e in data if isinstance(data, list) else []:
        key = (e.get("city"), e.get("area"), e.get("section"), str(e.get("lot_number", "")))
        if all(key) and key not in seen:
            seen.add(key)
            out.append(e)
    return out


def save_shot(d, e, tag):
    """截圖存到案件資料夾 1.基本資料(PDF) 與 1.基本資料\\png(PNG)；檔名跟手動版同一套"""
    base_dir = L.get_first_record_directory()
    pdf_dir = L.create_directory_structure(os.path.join(base_dir, "1.基本資料"))
    png_dir = L.create_directory_structure(os.path.join(base_dir, "1.基本資料", "png"))
    base = f"05_地籍便民{tag}-{e['area']}-{e['section']}-{e['lot_number']}"
    png = L.generate_unique_filename(png_dir, base, ".png")
    d.save_screenshot(png)
    pdf_path = L.generate_unique_filename(pdf_dir, base, ".pdf")
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.image(png, x=10, y=10, w=277)
    pdf.output(pdf_path)
    log(f"\033[93m已存檔：{png}\n        {pdf_path}\033[0m")


def open_page(d, url):
    d.get(url)
    try:
        WebDriverWait(d, 1.5).until(EC.alert_is_present())
        d.switch_to.alert.accept()
    except Exception:
        pass
    try:
        d.set_window_size(L.base_width, L.base_height)
        d.set_window_position(L.chrome_x, L.chrome_y)
    except Exception:
        pass
    L.verify_and_fix_chrome_window(d)
    L.apply_land_zoom()


# ---------------------------------------------------------------- 舊版(/Index)
def close_layer_menu(d):
    """左下「≡」圖層圖示：滑鼠移上去停一下再移開，「額外圖層」黑色選單就會收起來(跟手動版 land.py 同一招)"""
    from selenium.webdriver import ActionChains
    try:
        icon = WebDriverWait(d, 10).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.olControlButton2ItemActive[title='圖層'], div[title='圖層']")))
        ActionChains(d).move_to_element(icon).pause(0.6).perform()
        time.sleep(0.4)
        L.move_mouse_to_map()             # 移到地圖中央(body 中央是分頁列，會碰到分頁)
    except Exception as ex:
        log(f"收起額外圖層選單時略過：{ex}")


def hide_tips(d):
    """截圖前收掉黃色「地籍查詢 - 進階」提示框"""
    d.execute_script("""[...document.querySelectorAll('div,span,p')].forEach(function(x){
        if (x.offsetParent && /^\\s*【地籍查詢\\s*-\\s*進階】/.test(x.innerText||'') && x.offsetWidth < 600) x.style.display='none';});""")


def run_old(d, entries):
    open_page(d, OLD_URL)
    Select(WebDriverWait(d, 15).until(EC.presence_of_element_located((By.ID, "mapTile_id")))).select_by_visible_text("NLSC 地圖")
    time.sleep(1)
    # 額外圖層：勾「地籍圖」(預設沒勾；用點 label 的方式讓網站自己的事件跑)，然後把選單收起來
    try:
        on = d.execute_script("var c=document.getElementById('ckb_OFFICE_LAND'); return c ? c.checked : null;")
        if on is False:
            d.execute_script("var c=document.getElementById('ckb_OFFICE_LAND'); c.click();")
            time.sleep(1)
        on = d.execute_script("return document.getElementById('ckb_OFFICE_LAND').checked")
        log("地籍圖圖層：" + ("已開啟" if on else "開啟失敗"))
    except Exception as ex:
        log(f"開啟地籍圖圖層時略過：{ex}")
    close_layer_menu(d)
    try:
        tab = WebDriverWait(d, 10).until(EC.element_to_be_clickable((By.ID, "button_cada")))
        if "TabbedPanelsTabSelected" not in (tab.get_attribute("class") or ""):
            d.execute_script("arguments[0].click();", tab)
            time.sleep(0.5)
    except Exception:
        pass
    ok = 0
    for i, e in enumerate(entries, 1):
        L.check_exit()
        label = f"{e['city']} {e['area']} {e['section']} {e['lot_number']}"
        log(f"\n>>> [{i}/{len(entries)}] {label}")
        try:
            if not _pick(d.find_element(By.ID, "select_city_id"), e["city"]):
                log(f"  找不到縣市「{e['city']}」，略過")
                continue
            _wait_options(d, "select_town_id")
            time.sleep(0.5)
            if not _pick(d.find_element(By.ID, "select_town_id"), e["area"], part=True):
                log(f"  找不到行政區「{e['area']}」，略過")
                continue
            _wait_options(d, "select_sect_id")
            time.sleep(0.5)
            if not _pick(d.find_element(By.ID, "select_sect_id"), e["section"], strip_code=True):
                log(f"  找不到地段「{e['section']}」，略過")
                continue
            # 選好地段後網站會自動查「本段最小號/最大號」，查完會把地號欄清空 → 等它查完再填，送出前再確認一次
            time.sleep(2.5)
            lot = str(e["lot_number"])
            for _ in range(4):
                box = d.find_element(By.ID, "landno")
                box.clear()
                box.send_keys(lot)
                time.sleep(1)
                if d.find_element(By.ID, "landno").get_attribute("value") == lot:
                    break
            d.execute_script("arguments[0].click();", d.find_element(By.ID, "land_button"))
            r = L.handle_alert_immediately(d)
            if r == "no_data":
                log("  查無資料，略過")
                continue
            time.sleep(1)
            try:
                btn = WebDriverWait(d, 8).until(EC.element_to_be_clickable((By.ID, "info_close")))
                d.execute_script("arguments[0].click();", btn)
            except Exception:
                pass
            time.sleep(1)
            # 跟手動版一樣：查詢結果視窗固定到左邊(不擋地號)，滑鼠移過「進階」提示再移開讓提示框消失
            d.execute_script("""var w=document.querySelector('div[aria-describedby="dlg_search_result"]');
                                if(w){w.style.position='fixed';w.style.top='300px';w.style.left='10px';w.style.zIndex='9999';}""")
            try:
                from selenium.webdriver import ActionChains
                ac = ActionChains(d)
                ac.move_to_element(d.find_element(By.ID, "cada_span_id")).pause(0.5).perform()
                L.move_mouse_to_map()
            except Exception:
                pass
            time.sleep(1)
            L.center_marker()
            close_layer_menu(d)           # 拖曳時滑鼠經過左下角可能又叫出選單 → 再收一次
            hide_tips(d)
            time.sleep(3)                 # 等地圖圖磚載完
            save_shot(d, e, "")
            ok += 1
        except Exception as ex:
            log(f"  這筆失敗：{ex}")
    return ok


# ---------------------------------------------------------------- 最新版(/Z10Web/Normal)
def _lot8(lot):
    m = re.match(r"^\s*(\d+)(?:-(\d+))?\s*$", str(lot))
    return f"{int(m.group(1)):04d}{int(m.group(2) or 0):04d}" if m else ""


def run_new(d, entries):
    open_page(d, NEW_URL)
    time.sleep(4)
    d.execute_script("[...document.querySelectorAll('button')].filter(b=>/我已瞭解|關閉導覽/.test(b.innerText)).forEach(b=>b.click())")
    time.sleep(1)
    ok = 0
    for i, e in enumerate(entries, 1):
        L.check_exit()
        label = f"{e['city']} {e['area']} {e['section']} {e['lot_number']}"
        log(f"\n>>> [{i}/{len(entries)}] {label}")
        try:
            WebDriverWait(d, 20).until(lambda x: len(Select(x.find_element(By.ID, "land_city_id")).options) > 1)
            if not _pick(d.find_element(By.ID, "land_city_id"), e["city"]):
                log(f"  找不到縣市「{e['city']}」，略過")
                continue
            _wait_options(d, "land_town_id")
            time.sleep(0.5)
            if not _pick(d.find_element(By.ID, "land_town_id"), e["area"], part=True):
                log(f"  找不到行政區「{e['area']}」，略過")
                continue
            _wait_options(d, "land_section_id")
            time.sleep(0.5)
            if not _pick(d.find_element(By.ID, "land_section_id"), e["section"], strip_code=True):
                log(f"  找不到地段「{e['section']}」，略過")
                continue
            box = d.find_element(By.ID, "land_landno")
            box.clear()
            box.send_keys(str(e["lot_number"]))
            d.execute_script("arguments[0].click();", d.find_element(By.ID, "btn_land_search_id"))
            want = _lot8(e["lot_number"])
            found = False
            end = time.time() + 20
            while time.time() < end:
                try:
                    a = d.switch_to.alert
                    log(f"  網站訊息：{a.text}")
                    a.accept()
                    break
                except Exception:
                    pass
                txt = d.execute_script("return document.body.innerText") or ""
                if want and re.search(r"地號\s*" + want, txt):
                    found = True
                    break
                time.sleep(0.5)
            if not found:
                log("  查無資料或網站沒有回應，略過")
                continue
            time.sleep(3)                 # 等地圖圖磚與地號框線載完
            save_shot(d, e, "新版")
            ok += 1
        except Exception as ex:
            log(f"  這筆失敗：{ex}")
    return ok


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ver", choices=["old", "new"], default="old")
    a, _ = ap.parse_known_args()
    name = "舊版" if a.ver == "old" else "最新版"
    entries = load_entries()
    if not entries:
        log("data.json 沒有地段地號資料。請先用萬用查詢或「手動建立」建立案件。")
        return
    log(f"地籍便民自動截圖（{name}）：data.json 共 {len(entries)} 筆，逐筆定位後截圖存檔")
    d = L.create_chrome_driver(options=L.options)
    L.driver = d                          # land.py 的共用函式讀它自己的全域 driver
    try:
        try:
            d.set_window_size(L.base_width, L.base_height)
            d.set_window_position(L.chrome_x, L.chrome_y)
        except Exception:
            pass
        try:
            d.execute_cdp_cmd("Browser.grantPermissions", {"permissions": ["geolocation"]})
        except Exception:
            pass
        ok = run_old(d, entries) if a.ver == "old" else run_new(d, entries)
        log(f"\n====== 地籍便民自動截圖（{name}）完成：成功 {ok} / {len(entries)} 筆 ======")
        if ok < len(entries):
            log("沒成功的筆數可改用「手動建立」查詢，或稍後再試（官方定位服務偶爾會暫停）")
    except Exception as ex:
        log(f"執行過程中發生錯誤：{ex}")
    finally:
        try:
            d.quit()
        except Exception:
            pass
        try:
            import shutil
            shutil.rmtree(L.user_data_dir, ignore_errors=True)
        except Exception:
            pass
        log("地籍便民自動截圖已完成執行")


if __name__ == "__main__":
    main()
