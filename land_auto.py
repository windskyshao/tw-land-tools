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


def wait_adjust(sec=10):
    """跟手動版一樣：截圖前留 10 秒讓使用者調整地圖位置與顯示範圍(地圖圖磚也趁這時載完)"""
    log("\033[93m請調整地圖位置和顯示範圍，10 秒後自動截圖...\033[0m")
    L.countdown(sec)


def save_shot(d, e, tag, base=None):
    """截圖存到案件資料夾 1.基本資料(PDF) 與 1.基本資料\\png(PNG)；檔名跟手動版同一套"""
    base_dir = L.get_first_record_directory()
    pdf_dir = L.create_directory_structure(os.path.join(base_dir, "1.基本資料"))
    png_dir = L.create_directory_structure(os.path.join(base_dir, "1.基本資料", "png"))
    base = base or f"05_地籍便民{tag}-{e['area']}-{e['section']}-{e['lot_number']}"
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


# ---------------------------------------------------------------- 座標(跟手動版一樣：截圖上看得到座標)
def _fill_coord(e, coord):
    """data.json 這一筆沒有座標時才補上(有就不動)"""
    if not e:
        return
    if not coord or str(e.get('coordinates', '')).strip():
        return
    try:
        path = L.get_data_json_path()
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        key = (e.get('city'), e.get('area'), e.get('section'), str(e.get('lot_number', '')))
        changed = False
        for x in data:
            if (x.get('city'), x.get('area'), x.get('section'), str(x.get('lot_number', ''))) == key                     and not str(x.get('coordinates', '')).strip():
                x['coordinates'] = coord
                changed = True
        if changed:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            e['coordinates'] = coord
            log(f"  data.json 原本沒有座標 → 已補上 {coord}")
    except Exception as ex:
        log(f"  補座標到 data.json 失敗：{ex}")


def grab_coord_old(d, e):
    """舊版：圖釘按右鍵 →「取得此位置坐標」→ 座標視窗固定在左上角(跟手動版同一套)"""
    from selenium.webdriver import ActionChains
    try:
        marker = WebDriverWait(d, 8).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'image[id^="OpenLayers\.Geometry\.Point_"]')))
        ActionChains(d).context_click(marker).perform()
        item = WebDriverWait(d, 8).until(EC.visibility_of_element_located(
            (By.XPATH, '//li[contains(@class, "icon-getCoordWGS84byMap")]//span')))
        try:
            item.click()
        except Exception:                 # 被查詢結果/門牌列表視窗擋住 → 直接觸發選單項目
            d.execute_script("arguments[0].click();", item)
        coord = WebDriverWait(d, 8).until(EC.visibility_of_element_located((By.ID, 'coordDisplayLonLat'))).text.strip()
        d.execute_script("""var w=document.querySelector('div[aria-describedby="leftClickCoordDisplayId"]');
                            if(w){w.style.position='fixed';w.style.top='130px';w.style.left='5px';w.style.zIndex='9999';}""")
        log(f"  座標：{coord}")
        _fill_coord(e, coord)
        L.move_mouse_to_map()
        return coord
    except Exception as ex:
        log(f"  抓座標略過：{ex.__class__.__name__}")
        return ''


def close_coord_old(d):
    try:
        btn = d.find_element(By.ID, 'btnLeftClickCoordDisplayId')
        if btn.is_displayed():
            d.execute_script("arguments[0].click();", btn)
    except Exception:
        pass


def grab_coord_new(d, e):
    """新版：結果下方「坐標」→ 讀出緯經度後關掉置中的大視窗(會擋住地圖)，改在地圖左上角放座標標籤"""
    try:
        btn = d.execute_script("return [...document.querySelectorAll('a.showcoord')].find(x=>x.offsetParent)")
        if not btn:
            log("  抓座標略過：找不到「坐標」按鈕")
            return ''
        d.execute_script("arguments[0].click();", btn)
        coord = dms = ''
        end = time.time() + 8
        while time.time() < end and not coord:
            time.sleep(0.4)
            t = d.execute_script("return document.body.innerText") or ''
            m = re.search(r'緯經度[:：]\s*([\d.]+\s*,\s*[\d.]+)', t)
            if m:
                coord = m.group(1).replace(' ', '')
                m2 = re.search(r'緯經度\(度分秒\)[:：]\s*(\S+)', t)
                dms = m2.group(1) if m2 else ''
        # 關掉「此位置」視窗
        d.execute_script("""[...document.querySelectorAll('button')].filter(b=>b.offsetParent && b.innerText.trim()=='確定'
                              && /此位置/.test((b.closest('.modal,.swal2-popup,[role=dialog]')||document.body).innerText||''))
                              .forEach(b=>b.click());""")
        time.sleep(0.8)
        if not coord:
            log("  抓座標略過：沒讀到緯經度")
            return ''
        log(f"  座標：{coord}")
        _fill_coord(e, coord)
        d.execute_script("""var old=document.getElementById('__lm_coord'); if(old) old.remove();
            var b=document.createElement('div'); b.id='__lm_coord';
            b.style.cssText='position:fixed;z-index:99999;background:#fff;border:2px solid #16A34A;border-radius:8px;'
              +'padding:8px 14px;font:600 16px "Microsoft JhengHei",sans-serif;color:#111;box-shadow:0 2px 8px rgba(0,0,0,.25);';
            // 地圖右上角(右邊留給那排綠色工具按鈕)，不擋左邊的查詢資料
            b.style.right='86px'; b.style.top='14px';
            b.innerHTML='此位置坐標（WGS84）<br>緯經度：'+arguments[0]+(arguments[1]?'<br>度分秒：'+arguments[1]:'');
            document.body.appendChild(b);""", coord, dms)
        return coord
    except Exception as ex:
        log(f"  抓座標略過：{ex.__class__.__name__}")
        return ''


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
            grab_coord_old(d, e)
            close_layer_menu(d)           # 拖曳時滑鼠經過左下角可能又叫出選單 → 再收一次
            hide_tips(d)
            wait_adjust()
            hide_tips(d)
            save_shot(d, e, "")
            close_coord_old(d)
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
            time.sleep(1)
            grab_coord_new(d, e)
            wait_adjust()
            save_shot(d, e, "新版")
            ok += 1
        except Exception as ex:
            log(f"  這筆失敗：{ex}")
    return ok


# ---------------------------------------------------------------- 門牌查詢(data.json 有門牌時，地號截完再用門牌查一次)
_FW = str.maketrans('０１２３４５６７８９－', '0123456789-')
_DOOR_RE = re.compile(r'^(?P<road>.+?(?:路|街|大道)(?:[一二三四五六七八九十]+段)?)?(?:(?P<lane>\d+)巷)?'
                      r'(?:(?P<alley>\d+)弄)?(?P<no>\d+(?:之\d+)?)號(?P<floor>.*)$')


def _hw(s):
    return str(s or '').translate(_FW).replace(' ', '').replace('　', '')


def load_doors(entries):
    """data.json 的門牌(door 與 builds[].door) → [(entry, 門牌)]，同門牌只查一次"""
    out, seen = [], set()
    for e in entries:
        for door in [e.get('door', '')] + [b.get('door', '') for b in (e.get('builds') or []) if isinstance(b, dict)]:
            door = _hw(door)
            for pre in (e.get('city', ''), e.get('area', '')):
                if pre and door.startswith(pre):
                    door = door[len(pre):]
            if door and '號' in door and (e.get('city'), e.get('area'), door) not in seen:
                seen.add((e.get('city'), e.get('area'), door))
                out.append((e, door))
    return out


def _door_name(e, door):
    safe = re.sub(r'[\\/:*?"<>|]', '', door)
    return f"05_地籍便民門牌{{tag}}-{e['area']}-{safe}"


def _match_door(candidates, door):
    """候選門牌(網站列出的，全形數字)裡挑跟 data.json 門牌一樣的；有樓層就比到樓層，沒有就取第一筆"""
    want = _hw(door)
    m = _DOOR_RE.match(want)
    base = want if not m else want[:len(want) - len(m.group('floor') or '')]
    for i, c in enumerate(candidates):
        if _hw(c).endswith(want):
            return i
    if not (m and m.group('floor')):
        for i, c in enumerate(candidates):
            if base in _hw(c):
                return i
    return None


def _loaded(d, el_id, secs):
    try:
        _wait_options(d, el_id, secs)
        return True
    except Exception:
        return False


def _switch_and_back(d, el_id, want, next_id):
    """下拉選單已經是同一個值時，網站不會再載入下一層 → 先切到別的選項、等它載完，再切回來"""
    sel = Select(d.find_element(By.ID, el_id))
    other = next((o.text for o in sel.options[1:] if _norm(o.text) != _norm(want) and want not in o.text), None)
    if other:
        sel.select_by_visible_text(other)
        _loaded(d, next_id, 8)
        time.sleep(0.5)


def _door_selects_ready(d, e):
    """門牌分頁：縣市 → 行政區 → 道路 三層都要載出來。
    查完地號後縣市/行政區常已是同一個值，再選一次網站不會載入下一層 →
    ① 直接選 ② 先切到別的縣市/行政區再切回來 ③ 重新載入網頁再來一次"""
    for attempt in range(3):
        if attempt == 2:
            log("  下拉清單沒載出來 → 重新載入網頁再試")
            open_page(d, OLD_URL)
            Select(WebDriverWait(d, 15).until(EC.presence_of_element_located((By.ID, 'mapTile_id')))).select_by_visible_text('NLSC 地圖')
            time.sleep(1)
            close_layer_menu(d)
        elif attempt == 1:
            log("  下拉清單沒載出來 → 先切到別的縣市/行政區再切回來")
        d.execute_script("document.getElementById('button_addr').click();")
        time.sleep(1)
        _pick(d.find_element(By.ID, 'doorPlateTypeId'), '地政門牌')
        # 縣市 → 行政區
        if attempt == 1:
            _switch_and_back(d, 'select_city_id1', e['city'], 'select_town_id1')
        if not _pick(d.find_element(By.ID, 'select_city_id1'), e['city']):
            return False
        if not _loaded(d, 'select_town_id1', 6 if attempt < 2 else 12):
            continue
        time.sleep(0.5)
        # 行政區 → 道路
        if attempt == 1:
            _switch_and_back(d, 'select_town_id1', e['area'], 'select_road_id')
        if _pick(d.find_element(By.ID, 'select_town_id1'), e['area'], part=True) is None:
            continue
        if not _loaded(d, 'select_road_id', 6 if attempt < 2 else 12):
            continue
        time.sleep(0.5)
        return True
    return False


def run_old_doors(d, doors):
    """舊版：門牌分頁(地政門牌)→縣市/鄉鎮/道路/巷/弄/號 → 門牌列表挑那一戶 → 定位、截圖"""
    ok = 0
    for i, (e, door) in enumerate(doors, 1):
        L.check_exit()
        log(f"\n>>> 門牌 [{i}/{len(doors)}] {e['city']}{e['area']}{door}")
        m = _DOOR_RE.match(door)
        if not m:
            log("  門牌格式看不懂，略過")
            continue
        try:
            if not _door_selects_ready(d, e):
                log(f"  縣市／行政區／道路清單載不出來（{e['city']}{e['area']}），略過")
                continue
            road = m.group('road') or ''
            free = d.find_element(By.ID, 'roadFreeKeyInId')
            free.clear()
            if road and not _pick(d.find_element(By.ID, 'select_road_id'), road):
                free.send_keys(road)                         # 下拉沒有就用「輸入道路名稱」
            for fid, val in (('doorLaneId', m.group('lane')), ('doorAlleyId', m.group('alley')), ('doorNoId', m.group('no'))):
                box = d.find_element(By.ID, fid)
                box.clear()
                if val:
                    box.send_keys(val)
            d.execute_script("document.getElementById('door_botton').click();")
            r = L.handle_alert_immediately(d)
            if r == 'no_data':
                log("  查無此門牌，略過")
                continue
            # 門牌列表：每一戶是一個 <a doorplate="和平二路２４０之１號十二樓">
            items = []
            for tries in range(2):                           # 網站慢的時候第一次常沒回應 → 再按一次查詢
                if tries:
                    log("  門牌列表沒出來 → 再查一次")
                    d.execute_script("document.getElementById('door_botton').click();")
                    L.handle_alert_immediately(d)
                end = time.time() + 15
                while time.time() < end and not items:
                    time.sleep(0.5)
                    items = d.find_elements(By.CSS_SELECTOR, "a[onclick*='searchDoorLocation']")
                    items = [x for x in items if x.is_displayed()]
                if items:
                    break
            if not items:
                try:
                    d.save_screenshot(os.path.join(os.environ.get("TEMP", "."), "land_auto_door_fail.png"))
                except Exception:
                    pass
                log("  門牌列表沒有結果，略過")
                continue
            k = _match_door([x.get_attribute('doorplate') or x.text for x in items], door)
            if k is None:
                log(f"  門牌列表裡沒有「{door}」，略過（列表第一筆：{items[0].text}）")
                continue
            log(f"  選定：{items[k].get_attribute('doorplate') or items[k].text}")
            d.execute_script("arguments[0].click();", items[k])
            # 查太多次網站會要驗證碼(「查詢次數過多，請輸入驗証碼」)→ 跟手動版一樣請使用者在 Chrome 輸入
            if L.handle_alert_immediately(d) == 'no_data':
                log("  查無此門牌的位置，略過")
                continue
            if not WebDriverWait(d, 20).until(lambda x: x.execute_script(
                    "return document.querySelectorAll('image[id^=\"OpenLayers.Geometry.Point_\"]').length")):
                log("  沒有定位成功，略過")
                continue
            time.sleep(1)
            # 收掉門牌列表、地號標籤；查詢結果移到左邊(跟地號截圖一樣)
            d.execute_script("""[...document.querySelectorAll('.ui-dialog')].forEach(function(w){
                                    if (/門牌列表/.test(w.innerText||'')) w.style.display='none';});
                                var w=document.querySelector('div[aria-describedby="dlg_search_result"]');
                                if(w){w.style.position='fixed';w.style.top='300px';w.style.left='10px';w.style.zIndex='9999';}""")
            try:
                btn = WebDriverWait(d, 4).until(EC.element_to_be_clickable((By.ID, 'info_close')))
                d.execute_script("arguments[0].click();", btn)
            except Exception:
                pass
            time.sleep(1)
            L.center_marker()
            grab_coord_old(d, {})          # 門牌座標只顯示在截圖上，不寫 data.json(那是地號的座標)
            close_layer_menu(d)
            hide_tips(d)
            wait_adjust()
            hide_tips(d)
            save_shot(d, e, '', _door_name(e, door).format(tag=''))
            close_coord_old(d)
            ok += 1
        except Exception as ex:
            log(f"  這筆門牌失敗：{ex}")
    return ok


def run_new_doors(d, doors):
    """最新版：門牌地址定位分頁 → 打完整門牌(含樓層)搜尋 → 結果不是那一戶就從建議清單挑 → 截圖"""
    from selenium.webdriver import ActionChains
    ok = 0

    def result_door():
        t = d.execute_script("return document.body.innerText") or ''
        k = t.find('查詢結果門牌')
        return t[k:].split('\n', 2)[1].strip() if k >= 0 and t[k:].count('\n') >= 1 else ''

    for i, (e, door) in enumerate(doors, 1):
        L.check_exit()
        full = f"{e['city']}{e['area']}{door}"
        log(f"\n>>> 門牌 [{i}/{len(doors)}] {full}")
        try:
            # 上一筆截圖時加的「藏建議清單」與座標標籤先拿掉(這一筆要從建議清單挑樓層)
            d.execute_script("['__lm_hide_sugg','__lm_coord'].forEach(function(i){var x=document.getElementById(i); if(x) x.remove();});"
                             "[...document.querySelectorAll('div[data-value]')].forEach(function(x){"
                             "var p=x.parentElement; if(p) p.style.removeProperty('display');});")
            d.execute_script("document.getElementById('doorplate-tab').click();")
            time.sleep(1)
            box = d.find_element(By.ID, 'roodplateText')
            box.clear()
            box.send_keys(full)
            before = result_door()
            d.execute_script("document.getElementById('doorplateSearch').click();")
            end = time.time() + 15
            got = ''
            while time.time() < end:
                time.sleep(0.5)
                try:
                    a = d.switch_to.alert
                    log(f"  網站訊息：{a.text}")
                    a.accept()
                    break
                except Exception:
                    pass
                got = result_door()
                if got and got != before:
                    break
            if _match_door([got], door) is None:
                # 結果不是那一戶 → 從下拉建議裡挑，挑完要再按一次搜尋
                sugg = [x for x in d.find_elements(By.CSS_SELECTOR, 'div[data-value]') if x.is_displayed()]
                k = _match_door([x.text for x in sugg], door)
                if k is None:
                    log(f"  找不到「{door}」{('（網站給的是：' + got + '）') if got else ''}，略過")
                    continue
                ActionChains(d).move_to_element(sugg[k]).click().perform()
                time.sleep(1)
                d.execute_script("document.getElementById('doorplateSearch').click();")
                time.sleep(4)
                got = result_door()
                if _match_door([got], door) is None:
                    log(f"  定位結果不是「{door}」（{got}），略過")
                    continue
            log(f"  定位：{got}")
            # 收掉搜尋框下方殘留的門牌建議清單，免得擋到截圖
            d.execute_script("document.activeElement && document.activeElement.blur();"
                             "var st=document.getElementById('__lm_hide_sugg')||document.createElement('style');"
                             "st.id='__lm_hide_sugg'; st.textContent='div[data-value]{display:none!important}';"
                             "document.head.appendChild(st);"
                             "[...document.querySelectorAll('div[data-value]')].forEach(function(x){"
                             "var p=x.parentElement; if(p) p.style.setProperty('display','none','important');});")
            grab_coord_new(d, {})
            wait_adjust()
            save_shot(d, e, '', _door_name(e, door).format(tag='新版'))
            ok += 1
        except Exception as ex:
            log(f"  這筆門牌失敗：{ex}")
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
        doors = load_doors(entries)
        dok = 0
        if doors:
            log(f"\n------ data.json 有 {len(doors)} 個門牌，接著用門牌查詢截圖 ------")
            dok = run_old_doors(d, doors) if a.ver == "old" else run_new_doors(d, doors)
        log(f"\n====== 地籍便民自動截圖（{name}）完成：地號 {ok} / {len(entries)} 筆"
            + (f"、門牌 {dok} / {len(doors)} 筆" if doors else "") + " ======")
        if ok < len(entries) or dok < len(doors):
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
