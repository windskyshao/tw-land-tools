# -*- coding: utf-8 -*-
"""使用分區證明（正式證明書，要繳費）—— 連續模式「需要你操作」最後一項。

高雄：開瀏覽器 → 自動登入(申請人名稱/Email/驗證碼 ddddocr) → 網路申請表帶入地段地號、取回使用分區
      → 自動送出申請、切到「案件繳費」→ 你繳費(繳完約 3 分鐘可下載) → 按「繳好了」→ 到「案件進度查詢及下載」抓證明書。
      之前送出過沒繳/沒下載 → 先查進度、直接到繳費頁，不重複申請。
台南：先用申請人＋電話查「跨區聯合發證查詢系統」(免驗證碼)；有可下載的就直接抓，不開瀏覽器。
      沒有 → 開「臺南市一站式整合服務系統」讓你申請繳費 → 記下申請日；台南約 4 個工作天核發，之後再跑一次就會自動下載。
★繳費一律你自己操作(信用卡/行動支付)，程式不碰錢。
證明書是有效力的文件 → 存 0.謄本\使用分區證明_<段地號>_<日期>.pdf；申請紀錄 4.其他相關\分區證明申請.json
申請人(名稱/電話/Email)存本機 config.json 的「分區證明申請人」(各人自己的，不會打包)。

輸出給連續模式：FAIL: 失敗原因／PENDING: 已申請等核發／WARN: 提醒
"""
import argparse
import datetime
import glob
import json
import os
import re
import shutil
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
# 打包版的 python_embedded 有 ._pth，不會自動把本資料夾加進 sys.path → 要自己加，才 import 得到 webdriver_helper
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

KH_URL = "https://landuse.kcg.gov.tw/eudb2012/web_page/LMA000100.jsp"
TN_QUERY = "https://uplanduse.tainan.gov.tw/landuse/LU0201.jsp"
TN_APPLY = "https://onestop.tainan.gov.tw/"
HDR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36",
       "Accept-Language": "zh-TW,zh;q=0.9"}
CFG_KEY = "分區證明申請人"
REC_NAME = "分區證明申請.json"


def say(t=""):
    print(t, flush=True)


def step(stage, btn="", color="busy"):
    """告訴連續模式的提示框「現在輪到誰做什麼」：大字說明＋按鈕文字/顏色(busy＝程式在做，按鈕變灰)"""
    say(f"STEP: {stage}|{btn}|{color}")


def ask(prompt, btn="下一步 ▶", stage=None, color="blue"):
    """等連續模式的按鈕：空白＝繼續，q＝這次不申請。按下去後提示框切回「程式處理中」"""
    step(stage or prompt, btn, color)
    say(prompt)
    try:
        a = (sys.stdin.readline() or "q").strip().lower()
    except Exception:
        a = "q"
    if a != "q":
        step("程式處理中，請稍候…")
    return a


# ---------------------------------------------------------------- 案件、設定
def load_case(base):
    with open(os.path.join(base, "data.json"), encoding="utf-8") as f:
        recs = json.load(f)
    d0 = recs[0]
    folder = os.path.join(base, f"{d0.get('area', '')}{d0.get('section', '')}-{d0.get('lot_number', '')}")
    # 網站一件申請只能填同一個行政區、同一個地段(最多 10 筆) → 依地段分組，一組送一件
    # (例：鼓山 2165、2165-1 在龍華段六小段，213 在龍華段二小段 → 兩件)
    groups = {}
    for r in recs:
        if r.get("city") != d0.get("city"):
            continue
        ln = str(r.get("lot_number", "")).strip()
        g = groups.setdefault((r.get("area", ""), r.get("section", "")), [])
        if ln and ln not in g:
            g.append(ln)
    gl = [dict(area=a, section=s, lots=ls[i:i + 10]) for (a, s), ls in groups.items() for i in range(0, len(ls), 10) if ls]
    first = gl[0] if gl else dict(area=d0.get("area", ""), section=d0.get("section", ""), lots=[])
    return dict(city=(d0.get("city") or "").replace("台", "臺"), area=first["area"], section=first["section"],
                lots=first["lots"], groups=gl, folder=folder)


def _lot_key(s):
    """地號比對用：0213-0000 / 02130000 / 213 → 213；2165-0001 / 2165-1 → 2165-1"""
    s = str(s).strip()
    if re.fullmatch(r"\d{8}", s):
        s = s[:4] + "-" + s[4:]
    m = re.match(r"0*(\d+)(?:-0*(\d+))?", s)
    if not m:
        return s
    sub = m.group(2)
    return m.group(1) if not sub or int(sub) == 0 else f"{m.group(1)}-{int(sub)}"


def has_cert(case, g):
    """這一組(地段)在 0.謄本 已經有證明書了嗎"""
    pre = f"使用分區證明_{g['section']}{g['lots'][0]}" if g["lots"] else f"使用分區證明_{g['section']}"
    return bool(glob.glob(os.path.join(glob.escape(os.path.join(case["folder"], "0.謄本")), glob.escape(pre) + "*")))


def which_group(case, path):
    """自己下載/手動申請的證明書是哪一組：看 PDF 上的地段全名、地號(例：龍華段六小段、2165-0000)"""
    txt = re.sub(r"\s+", "", _pdf_text(path))
    best, score = case["groups"][0], -1
    for g in case["groups"]:
        sc = (3 if g["section"] in txt else 0) + sum(
            1 for ln in g["lots"] if re.search(r"(?<!\d)0*" + re.escape(_lot_key(ln).split("-")[0]) + r"-0*"
                                               + (_lot_key(ln).split("-") + ["0"])[1] + r"(?!\d)", txt))
        if sc > score:
            best, score = g, sc
    return best


def load_applicant(base):
    try:
        with open(os.path.join(base, "config.json"), encoding="utf-8") as f:
            a = json.load(f).get(CFG_KEY) or {}
    except Exception:
        a = {}
    return {k: str(a.get(k, "")).strip() for k in ("name", "phone", "email")}


def rec_path(case):
    return os.path.join(case["folder"], "4.其他相關", REC_NAME)


def load_rec(case):
    try:
        with open(rec_path(case), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_rec(case, rec):
    os.makedirs(os.path.dirname(rec_path(case)), exist_ok=True)
    with open(rec_path(case), "w", encoding="utf-8") as f:
        json.dump(rec, f, ensure_ascii=False, indent=2)


def out_name(case, day=None):
    """檔名日期＝證明書上的核發日期(讀不到才用今天)；有效期八個月從這天算"""
    lots = "、".join(case["lots"][:3]) + ("等" if len(case["lots"]) > 3 else "")
    return f"使用分區證明_{case['section']}{lots}地號_{(day or datetime.date.today()):%Y%m%d}.pdf"


def issue_date(txt):
    """「中華民國 115 年 11 月 06 日」(高雄發文日期／台南核發日期) → date"""
    m = re.search(r"中華民國\s*(\d{2,3})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", txt or "")
    try:
        return datetime.date(int(m.group(1)) + 1911, int(m.group(2)), int(m.group(3))) if m else None
    except ValueError:
        return None


def _pdf_text(path):
    try:
        import fitz
        with fitz.open(path) as doc:
            return "".join(p.get_text() for p in doc)
    except Exception:
        return ""


def is_pdf(path):
    try:
        with open(path, "rb") as f:
            return f.read(5) == b"%PDF-"
    except Exception:
        return False


def keep_pdf(case, src):
    """收進 0.謄本；檢查證明書上有沒有這個地段(掃描檔讀不到字就不判斷)"""
    dst_dir = os.path.join(case["folder"], "0.謄本")
    os.makedirs(dst_dir, exist_ok=True)
    raw = _pdf_text(src)
    day = issue_date(raw)
    name = out_name(case, day)
    dst = os.path.join(dst_dir, name)
    n = 2
    while os.path.exists(dst):
        dst = os.path.join(dst_dir, name[:-4] + f"_{n}.pdf")
        n += 1
    shutil.move(src, dst)
    say(f"✓ 已存：0.謄本\\{os.path.basename(dst)}")
    if day:
        end = day.replace(year=day.year + (day.month + 8 - 1) // 12, month=(day.month + 8 - 1) % 12 + 1, day=min(day.day, 28))
        say(f"核發日期 {day.year - 1911}/{day.month:02d}/{day.day:02d}，有效期限約到 {end.year - 1911}/{end.month:02d}")
    txt = re.sub(r"\s+", "", raw)
    short = re.sub(r"段.*$", "", case["section"])            # 麻工段 → 麻工(證明書常寫簡稱)
    if txt and short and short not in txt:
        say(f"WARN: 證明書上找不到「{short}」，請打開確認是不是這個案件")
    return dst


# ---------------------------------------------------------------- 台南
def _roc(d):
    return f"{d.year - 1911}-{d.month:02d}-{d.day:02d}"


def _tags(h):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h or "")).strip()


def tn_query(app):
    """近 8 個月內這個申請人的案件 → [{text, dl(下載網址或 None)}]"""
    import requests
    import urllib3
    urllib3.disable_warnings()
    s = requests.Session()
    s.headers.update(HDR)
    s.verify = False
    end = datetime.date.today()
    start = end - datetime.timedelta(days=240)
    s.get(TN_QUERY, timeout=30)
    r = s.post(f"{TN_QUERY}?showTable=1&queryMode=1&startDate={_roc(start)}&endDate={_roc(end)}",
               data={"showHide": "0", "startDatepicker": _roc(start), "endDatepicker": _roc(end),
                     "USEM_PEPO": app["name"], "USEM_TEL": app["phone"]},
               headers={"Referer": TN_QUERY}, timeout=60)
    r.encoding = "utf-8"
    h = r.text
    i = h.find("處理狀態")
    if i < 0:
        raise RuntimeError("台南查詢系統回應格式變了")
    rows = []
    for tr in re.findall(r"(?is)<tr[^>]*>(.*?)</tr>", h[i:]):
        t = _tags(tr)
        if not re.search(r"\d{3}[/-]\d{1,2}[/-]\d{1,2}|\d{7}", t):
            continue
        m = re.search(r"""(dl\.jsp\?[0-9A-Za-z]+|https?://[^'"]+dl\.jsp\?[0-9A-Za-z]+)""", tr)
        dl = None
        if m:
            dl = m.group(1) if m.group(1).startswith("http") else "https://uplanduse.tainan.gov.tw/landuse/" + m.group(1)
        else:
            m = re.search(r"""(?:href|onclick)=["'][^"']*?([\w./?=&%-]*(?:download|Download|下載)[\w./?=&%-]*)""", tr)
            if m and "下載" in t:
                dl = m.group(1) if m.group(1).startswith("http") else "https://uplanduse.tainan.gov.tw/landuse/" + m.group(1).lstrip("./")
        rows.append(dict(text=t, dl=dl, html=tr))
    return rows, s


def _row_match(case, t):
    short = re.sub(r"段.*$", "", case["section"])
    t2 = t.replace(" ", "")
    return short in t2 and any(re.sub(r"-0*$", "", l).split("-")[0] in t2 for l in case["lots"])


def tainan(case, app, rec, log_dir):
    step("程式正在查台南有沒有已核發的證明，請稍候…")
    say(f"查台南「跨區聯合發證」進度：申請人 {app['name']}、電話 {app['phone']}")
    try:
        rows, s = tn_query(app)
    except Exception as e:
        say(f"FAIL: 台南查詢系統連不上（{str(e)[:40]}）")
        return
    mine = [r for r in rows if _row_match(case, r["text"])]
    say(f"近 8 個月有 {len(rows)} 件，這個地段 {len(mine)} 件")
    for r in mine:
        say("  " + r["text"][:120])
    ready = [r for r in mine if r["dl"]]
    if not ready and len(rows) and rec.get("tainan_applied"):        # 地段寫法對不上：申請日之後有可下載的也試(下載後再核對內容)
        ready = [r for r in rows if r["dl"]][:3]
    got = 0
    tmp = os.path.join(log_dir, "_dl.pdf")
    for r in ready:
        try:
            b = s.get(r["dl"], timeout=60, headers={"Referer": TN_QUERY}).content
        except Exception as e:
            say(f"（下載失敗：{e}）")
            continue
        if not b.startswith(b"%PDF-"):
            continue
        with open(tmp, "wb") as f:
            f.write(b)
        if not _row_match(case, r["text"]) and re.sub(r"段.*$", "", case["section"]) not in _pdf_text(tmp).replace(" ", ""):
            os.remove(tmp)
            continue
        keep_pdf(case, tmp)
        got += 1
    if got:
        rec["tainan_done"] = datetime.date.today().isoformat()
        save_rec(case, rec)
        return
    if mine:                                        # 已申請、還沒核發
        rec.setdefault("tainan_applied", datetime.date.today().isoformat())
        save_rec(case, rec)
        say("PENDING: 已申請，台南約 4 個工作天核發；之後再跑這一項就會自動下載")
        return
    if rec.get("tainan_applied"):
        say(f"（{rec['tainan_applied']} 已申請，查詢系統還沒出現；剛申請要等一下才查得到）")
        if ask("要再申請一次嗎？要就按「再申請一次」，不用就按「這次不申請」", "再申請一次 ▶",
               "之前申請的還查不到，要再申請一次嗎？", "orange") == "q":
            say("PENDING: 已申請，等台南核發（約 4 個工作天）")
            return
    # 開一站式讓使用者申請(登入、填寫、繳費都在網站上自己做)
    lots = "、".join(case["lots"])
    say("═" * 30)
    say("請在瀏覽器「臺南市一站式整合服務系統」申請使用分區證明：")
    say(f"  行政區：{case['area']}　地段：{case['section']}　地號：{lots}")
    say(f"  申請人填「{app['name']}」、電話「{app['phone']}」（之後靠這兩個自動下載）")
    say("繳完費按「申請好了」；不申請按「這次不申請」")
    drv = None
    try:
        drv = _driver(log_dir)
        drv.get(TN_APPLY)
        _fix_window(drv)
    except Exception as e:
        say(f"（瀏覽器開不起來：{e}；請自己開 {TN_APPLY}）")
    a = ask("等你申請…", "申請好了 ▶", "請在瀏覽器申請並繳費，好了按「申請好了」", "green")
    _quit(drv)
    if a == "q":
        say("SKIP: 這次不申請")
        return
    rec["tainan_applied"] = datetime.date.today().isoformat()
    save_rec(case, rec)
    say("PENDING: 已申請，台南約 4 個工作天核發；之後再跑這一項就會自動下載")


# ---------------------------------------------------------------- 瀏覽器
BASE = [os.path.dirname(os.path.abspath(sys.argv[0]))]


def _window_rect():
    """跟電子謄本一樣：主程式在右 → Chrome 佔左邊 2/3、整個工作區高(讀主程式寫的 window_config.json)"""
    try:
        with open(os.path.join(BASE[0], "window_config.json"), encoding="utf-8") as f:
            wc = json.load(f)
        wa, mw = wc.get("work_area") or {}, wc.get("main_window") or {}
        if not (wa and mw):
            return None
        sc = wc.get("dpi_scale")
        if not sc:
            import ctypes
            hdc = ctypes.windll.user32.GetDC(0)
            sc = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88) / 96.0
            ctypes.windll.user32.ReleaseDC(0, hdc)
        l, t, r, b = wa.get("left", 0), wa.get("top", 0), wa.get("right", 1920), wa.get("bottom", 1080)
        w = max(800, (r - l) * 2 // 3)
        x = l if mw.get("x", 1280) + mw.get("width", 640) / 2 > (l + r) / 2 else r - w
        h = max(600, b - t - 20)
        return int(x / sc), int(t / sc), int(w / sc), int(h / sc) + 32
    except Exception:
        return None


def _driver(dl_dir):
    from selenium import webdriver
    from webdriver_helper import create_chrome_driver
    os.environ.pop("LAND_HEADLESS", None)
    o = webdriver.ChromeOptions()
    o.add_experimental_option("prefs", {"download.default_directory": dl_dir, "download.prompt_for_download": False,
                                        "plugins.always_open_pdf_externally": True})
    drv = create_chrome_driver(o)
    rc = _window_rect()
    try:
        if rc:
            drv.set_window_size(rc[2], rc[3])
            drv.set_window_position(rc[0], rc[1])
        else:
            drv.set_window_size(1400, 1000)
    except Exception:
        pass
    return drv


def _fix_window(drv):
    try:
        from webdriver_helper import verify_and_fix_chrome_window
        verify_and_fix_chrome_window(drv)
    except Exception:
        pass


def _quit(drv):
    try:
        drv and drv.quit()
    except Exception:
        pass


_OCR = []


def _ocr(png):
    """ddddocr(打包版的 python_embedded 沒有)→ 退回 RapidOCR(有附)"""
    if not _OCR:
        try:
            import ddddocr
            o = ddddocr.DdddOcr(show_ad=False)
            _OCR.append(lambda b: o.classification(b) or "")
        except Exception:
            import io
            import numpy as np
            from PIL import Image, ImageOps
            from rapidocr_helper import create_rapidocr_reader
            rd = create_rapidocr_reader(quiet=True, use_angle_cls=False, text_score=0.1)

            def rapid(b):
                # 高雄驗證碼 74x25 四位數字：去淡色雜點→放大 3 倍→加白邊→只做「辨識」不做「找字框」(找框在小圖會漏字)
                # 實測 15 張全對；不去雜點會把 2518 認成 2218
                g = Image.open(io.BytesIO(b)).convert("L")
                out = ""
                for thr in (150, 0):
                    a = np.array(g)
                    if thr:
                        a = np.where(a < thr, 0, 255).astype("uint8")
                    im = Image.fromarray(a).resize((g.width * 3, g.height * 3), Image.LANCZOS)
                    im = np.array(ImageOps.expand(im, border=20, fill=255).convert("RGB"))
                    try:
                        r, _ = rd(im, use_det=False, use_cls=False, use_rec=True)
                        out = "".join(x[0] for x in (r or []))
                    except TypeError:                                # 舊版 RapidOCR 沒有 use_det
                        r, _ = rd(im)
                        out = "".join(x[1] for x in (r or []))
                    if len([c for c in out if c.isdigit()]) == 4:
                        break
                return out
            _OCR.append(rapid)
    try:
        return _OCR[0](png)
    except Exception:
        return ""


def _alerts(drv):
    """網站的提示框(歡迎使用、成功訊息…)按掉，內容印出來"""
    for _ in range(3):
        try:
            al = drv.switch_to.alert
            say(f"（網站訊息：{al.text[:60]}）")
            al.accept()
            time.sleep(0.5)
        except Exception:
            return


def _in_frame(drv):
    _alerts(drv)
    drv.switch_to.default_content()
    try:
        drv.switch_to.frame("content")
        return True
    except Exception:
        return False


def kh_login(drv, app):
    """自動登入；成功回 True(失敗就請使用者自己登入)"""
    from selenium.webdriver.common.by import By
    drv.get(KH_URL)
    _fix_window(drv)
    time.sleep(2)
    for n in range(6):
        if not _in_frame(drv):
            time.sleep(1)
            continue
        if drv.find_elements(By.NAME, "updateForm") or drv.find_elements(By.NAME, "USED_DIST"):
            return True
        btn = drv.find_elements(By.ID, "BTN01")
        if btn and not drv.find_elements(By.ID, "USERNAME"):
            btn[0].click()
            time.sleep(1.5)
            continue
        if not drv.find_elements(By.ID, "USERNAME"):
            time.sleep(1.5)
            continue
        img = drv.find_elements(By.CSS_SELECTOR, "img.CHKIMG")
        code = _ocr(img[0].screenshot_as_png) if img else ""
        code = re.sub(r"[^0-9A-Za-z]", "", code)[:4]
        drv.execute_script("""var f=document.FM01; f.USERNAME.value=arguments[0]; f.USEREMAIL.value=arguments[1];
            if(f.USERPHONE) f.USERPHONE.value=arguments[2];
            var c=document.querySelectorAll('[name=USERCD]'); for(var i=0;i<c.length;i++){c[i].value=arguments[3];}""",
                           app["name"], app["email"], app["phone"], code)
        if len(code) == 4:
            drv.execute_script("""window.__al=[]; window.alert=function(m){window.__al.push(String(m))};
                if(typeof PF01==='function') PF01('LMA010102','jsp','_self','','');""")
            time.sleep(2.5)
            _in_frame(drv)
            if not drv.find_elements(By.ID, "USERNAME"):
                say("✓ 已自動登入")
                return True
        say(f"（驗證碼辨識沒過，重試 {n + 1}）")
        try:
            drv.find_element(By.CSS_SELECTOR, ".A01").click()
        except Exception:
            drv.refresh()
        time.sleep(1.5)
    return False


def kh_open_form(drv):
    """登入後進「網路申請」(有同意頁就按同意)；回 True＝表單出來了"""
    from selenium.webdriver.common.by import By
    for _ in range(10):
        if _in_frame(drv) and drv.find_elements(By.NAME, "USED_DIST"):
            return True
        if _in_frame(drv):
            drv.execute_script("""window.alert=function(){};
                [].slice.call(document.querySelectorAll('input[type=checkbox]')).forEach(function(c){
                  if(!c.checked){ c.click(); } });
                var b=[].slice.call(document.querySelectorAll('input[type=submit],input[type=button],button,a')).filter(function(x){
                  return /同意/.test(x.value||x.innerText||'') && !/不同意/.test(x.value||x.innerText||'')})[0];
                if(b){ b.disabled=false; b.click(); }""")
            time.sleep(1.5)
            _in_frame(drv)
            time.sleep(1.5)
            if drv.find_elements(By.NAME, "USED_DIST"):
                return True
        drv.switch_to.default_content()
        drv.execute_script("""var a=[].slice.call(document.querySelectorAll('a')).filter(function(x){return /網路/.test(x.innerText)})[0];
            if(a) a.click();""")
        time.sleep(2)
    return False


def kh_fill(drv, case, app):
    """帶入申請人、行政區、地段、地號，按「取回使用分區」；回 (成功, 訊息)"""
    _in_frame(drv)
    r = drv.execute_script("""
      var f=document.updateForm, area=arguments[0], sect=arguments[1], app=arguments[2];
      window.__al=[]; window.alert=function(m){window.__al.push(String(m))};
      function set(n,v){ if(f[n] && v && !f[n].value){ f[n].value=v; } }
      set('USEM_PEPO',app.name); set('USEM_TEL1',app.phone); set('USEM_MAIL',app.email);
      var dc=''; for(var i=0;i<f.USED_DIST.options.length;i++){ if(f.USED_DIST.options[i].text==area) dc=f.USED_DIST.options[i].value; }
      if(!dc) return 'NODIST';
      f.USED_DIST.value=dc; f.USED_DIST.dispatchEvent(new Event('change'));
      return dc;""", case["area"], case["section"], app)
    if r == "NODIST":
        return False, f"高雄系統找不到行政區「{case['area']}」"
    time.sleep(2.5)
    r = drv.execute_script("""
      var f=document.updateForm, sect=arguments[0], dc=arguments[1], hit=-1;
      for(var i=0;i<f.USED_SECT.options.length;i++){
        if(f.USED_SECT.options[i].text==sect){ var ad=f.ALL_DIST && f.ALL_DIST.options[i]; if(hit<0 || (ad && ad.value==dc)) hit=i; } }
      if(hit<0) return 'NOSECT';
      f.USED_SECT.selectedIndex=hit; f.USED_SECT.dispatchEvent(new Event('change'));
      return f.USED_SECT.value;""", case["section"], r)
    if r == "NOSECT":
        return False, f"高雄系統找不到地段「{case['section']}」"
    time.sleep(1.5)
    for i, lot in enumerate(case["lots"]):
        drv.execute_script("""var el=document.updateForm.USED_LAND[arguments[0]]||document.updateForm.USED_LAND; el.focus();
            el.value=arguments[1]; el.dispatchEvent(new Event('input')); el.dispatchEvent(new Event('change'));
            el.blur(); el.dispatchEvent(new Event('blur'));""", i, lot)
        time.sleep(1.0)
    drv.execute_script("if(typeof getUSEDALL==='function') getUSEDALL();")
    time.sleep(5)
    got = drv.execute_script("""var f=document.updateForm,o=[];
      for(var i=0;i<10;i++){ var el=f.USED_LAND[i]; if(el && el.value){ var z=document.getElementById('ULN_'+i);
        o.push(el.value+' '+(z?z.value:'')); } }
      return o.join(' | ')+(window.__al.length?' ｜網站訊息：'+window.__al.join(' / '):'');""")
    return True, got or ""


def kh_submit(drv):
    """按「送出申請」(網站問確定就按確定)；回 (成功, 網站訊息)。表單還在＝沒送出(欄位沒填好之類)"""
    from selenium.webdriver.common.by import By
    _in_frame(drv)
    drv.execute_script("window.__al=[]; window.alert=function(m){window.__al.push(String(m))};")
    if drv.execute_script("return typeof checkForm!=='function'"):
        return False, "找不到「送出申請」按鈕"
    # 網站要 isDone(「取回使用分區」回來了)且算出金額才肯送；地號欄一失焦就會把 isDone、金額清掉
    # → 先讓地號欄失焦，捲到下面按一次「取回使用分區」(畫面看得到)，等金額出來(例：100)再直接呼叫 checkForm
    for k in range(2):
        drv.execute_script("var a=document.activeElement; if(a && a.blur) a.blur();")
        time.sleep(0.5)
        b = [x for x in drv.find_elements(By.CSS_SELECTOR, "input[type=button]")
             if "取回使用分區" in (x.get_attribute("value") or "")]
        if b:
            drv.execute_script("arguments[0].scrollIntoView({block:'center'});", b[0])
            time.sleep(0.5)
            b[0].click()
        else:
            drv.execute_script("getUSEDALL();")
        for _ in range(20):
            if drv.execute_script("return isDone===true"):
                break
            time.sleep(1)
        time.sleep(1)
        # 分區回來了但「應繳金額」還空白(11-13 鼓山實際遇到)→ 叫網站自己的算錢函式補上
        drv.execute_script("""if((document.updateForm.cost.value||'').replace(/\\s/g,'')==='' && typeof countCost==='function')
            countCost();""")
        say(f"應繳金額：{drv.execute_script('return document.updateForm.cost.value') or '（沒算出來）'} 元")
        drv.execute_script("""var s=[].slice.call(document.querySelectorAll('input[type=button]')).filter(function(x){
            return /送出申請/.test(x.value)})[0]; if(s) s.scrollIntoView({block:'center'});""")
        time.sleep(0.5)
        drv.execute_script("window.__al=[]; setTimeout(function(){checkForm()}, 50);")
        time.sleep(1.5)
        try:
            again = any("取回使用分區" in m for m in drv.execute_script("return window.__al||[]") or [])
        except Exception:
            again = False                                   # 跳出確認框/換頁了 → 交給下面處理
        if not again:
            break
    msgs = []
    for _ in range(20):
        time.sleep(1)
        try:
            al = drv.switch_to.alert                        # 確認框(confirm)或成功訊息
            msgs.append(al.text[:80])
            al.accept()
            continue
        except Exception:
            pass
        drv.switch_to.default_content()
        try:
            drv.switch_to.frame("content")
        except Exception:
            continue
        try:
            bad = drv.execute_script("var a=window.__al||[]; window.__al=[]; return a;") or []
        except Exception:
            bad = []
        msgs += bad
        if not drv.find_elements(By.NAME, "USED_DIST"):
            return True, " / ".join(msgs)
        if bad and drv.find_elements(By.NAME, "USED_DIST"):    # 表單檢查沒過(網站跳「XX不能為空白」之類)
            return False, " / ".join(msgs)
    return False, " / ".join(msgs) or "送出後畫面沒有變"


def kh_pay_page(drv):
    """切到「案件繳費」；回待繳費頁的文字(看得到這件就是送出成功)"""
    _alerts(drv)
    drv.switch_to.default_content()
    drv.execute_script("""var a=[].slice.call(document.querySelectorAll('a')).filter(function(x){return /案件繳費/.test(x.innerText)})[0];
        if(a) a.click();""")
    time.sleep(3)
    _in_frame(drv)
    try:
        return drv.execute_script("return document.body.innerText") or ""
    except Exception:
        return ""


def kh_rows(drv):
    """「案件進度查詢及下載」列表(預設列已繳費、8 個月內) → [{no, date, dist, sect, lots, status}]"""
    try:
        drv.switch_to.window(drv.window_handles[0])        # 繳費可能開了新分頁 → 回主視窗
    except Exception:
        pass
    drv.switch_to.default_content()
    drv.execute_script("""var a=[].slice.call(document.querySelectorAll('a')).filter(function(x){return /案件進度/.test(x.innerText)})[0];
        if(a) a.click();""")
    time.sleep(3)
    _in_frame(drv)
    return drv.execute_script("""var o=[];
        [].slice.call(document.querySelectorAll('tr')).forEach(function(tr){
          var td=tr.querySelectorAll('td'); if(td.length<8) return;
          var no=(td[3].innerText||'').trim(); if(!/^\\d{3}[A-Z]\\d+/.test(no)) return;
          o.push({no:no, date:(td[2].innerText||'').trim(), dist:(td[4].innerText||'').trim(), sect:(td[5].innerText||'').trim(),
                  lots:(td[6].innerText||'').trim(), status:(td[7].innerText||'').trim(),
                  dl:!!tr.querySelector('[onclick*="PF01"]')}); });
        return o;""") or []


def kh_download(drv, dl_dir, case, since=""):
    """找這個地段、這個地號、申請日 >= since 的最新一件，按它的「電子文件下載」；沒抓到回 None"""
    rows = kh_rows(drv)
    # 列表的地號欄只寫「最小的那個地號等N筆」(旗津填 859、887、446 → 顯示 446等3筆)→ 看它是不是這組的任一筆；
    # 這樣你自己在網站手動申請的(同地段、同地號)也抓得到
    keys = {_lot_key(x) for x in case["lots"]}
    mine = [r for r in rows if r["sect"] == case["section"] and r["dl"] and r["date"] >= since
            and _lot_key((re.match(r"[\d-]+", r["lots"].replace(" ", "")) or re.match("", "")).group(0)) in keys]
    say(f"案件進度：8 個月內 {len(rows)} 件，{case['section']} {'、'.join(case['lots'])} 可下載 {len(mine)} 件")
    if not mine:
        return None
    r = max(mine, key=lambda x: (x["date"], x["no"]))
    say(f"下載 {r['date']} 案件 {r['no']}（{r['sect']} {r['lots']}，{r['status']}）")
    before, wins = set(os.listdir(dl_dir)), set(drv.window_handles)
    drv.execute_script("$('#FDL01').val(''); PF01(arguments[0]);", r["no"])
    p = _wait_pdf(drv, dl_dir, before, wins, 120)          # 網站說第一次下載約要 1～2 分鐘
    if p:
        say("✓ 已下載")
    return p


def _roc7(d):
    return f"{d.year - 1911}{d.month:02d}{d.day:02d}"


def _pdf_from_zip(path):
    """高雄下載的是 ZIP(證明書 PDF＋電子簽章 .sig)→ 取出 PDF；不是 ZIP 回 None"""
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            pdfs = sorted((i for i in z.infolist() if i.filename.lower().endswith(".pdf")), key=lambda i: -i.file_size)
            if not pdfs:
                return None
            out = path + ".pdf"
            with z.open(pdfs[0]) as src, open(out, "wb") as dst:
                shutil.copyfileobj(src, dst)
        return out if is_pdf(out) else None
    except Exception:
        return None


def _wait_pdf(drv, dl_dir, before, wins, sec):
    """等下載資料夾出現新的 PDF；或新開視窗是 PDF 網址就用 cookies 抓"""
    t0 = time.time()
    while time.time() - t0 < sec:
        new = [f for f in os.listdir(dl_dir) if f not in before and not f.endswith((".crdownload", ".tmp"))]
        for f in new:
            p = os.path.join(dl_dir, f)
            if is_pdf(p):
                return p
            q = _pdf_from_zip(p)
            if q:
                return q
        for h in set(drv.window_handles) - wins:
            try:
                drv.switch_to.window(h)
                u = drv.current_url
                if u.startswith("http"):
                    import requests
                    ck = {c["name"]: c["value"] for c in drv.get_cookies()}
                    b = requests.get(u, cookies=ck, headers=HDR, timeout=60, verify=False).content
                    if b.startswith(b"%PDF-"):
                        p = os.path.join(dl_dir, "_kh.pdf")
                        with open(p, "wb") as fo:
                            fo.write(b)
                        return p
            except Exception:
                pass
        time.sleep(1)
    try:
        drv.switch_to.window(drv.window_handles[0])
    except Exception:
        pass
    return None


def _got_locals(log_dir):
    """使用者自己在瀏覽器按了下載(檔案落在本次的工作資料夾，每次開始前會清空)→ 取出的 PDF 路徑們；ZIP 取完就刪，不重複收"""
    out = []
    for f in os.listdir(log_dir):
        p = os.path.join(log_dir, f)
        if f.endswith((".crdownload", ".tmp", ".zip.pdf")) or not os.path.isfile(p):
            continue
        if f.lower().endswith(".zip"):
            q = _pdf_from_zip(p)
            if q:
                out.append(q)
                try:
                    os.remove(p)
                except OSError:
                    pass
        elif is_pdf(p):
            out.append(p)
    return out


def _gname(g):
    return f"{g['section']} {'、'.join(g['lots'])}"


def kaohsiung(case, app, rec, log_dir):
    drv = _driver(log_dir)
    groups = case["groups"]
    got = set()                                         # 這次已經存好的組(index)

    def store(i, p):
        keep_pdf(dict(case, **groups[i]), p)
        got.add(i)
        if p.endswith(".zip.pdf"):                      # 取完的 ZIP 刪掉，免得下一輪又被當成新下載
            try:
                os.remove(p[:-4])
            except OSError:
                pass

    def all_done():
        left = [g for i, g in enumerate(groups) if i not in got and not has_cert(case, g)]
        if not left:
            rec["kaohsiung_done"] = datetime.date.today().isoformat()
            rec.pop("missing", None)
        else:                                           # 連續模式清單顯示「還缺哪個地段」
            rec["missing"] = [_gname(g) for g in left]
        save_rec(case, rec)
        return not left

    def take_local():
        """你自己在瀏覽器下載的(包括你手動加申請的)：看證明書上的地段地號歸到對的組"""
        for p in _got_locals(log_dir):
            g = which_group(case, p)
            say(f"✓ 找到你在瀏覽器下載的證明書（{_gname(g)}）")
            store(groups.index(g), p)

    def fetch(i, since):
        p = kh_download(drv, log_dir, dict(case, **groups[i]), since=since)
        if p:
            store(i, p)
        return bool(p)
    try:
        step("程式正在自動登入、填表、送出申請，請稍候…")
        say("開啟高雄市土地使用分區證明核發系統…")
        if len(groups) > 1:
            say(f"這個案件的地號分在 {len(groups)} 個地段，網站一件只能填一個地段 → 分 {len(groups)} 件申請："
                + "；".join(_gname(g) for g in groups))
        if not kh_login(drv, app):
            if ask("請在瀏覽器自己登入（申請人名稱、Email、驗證碼），登入好了按「登入好了」", "登入好了 ▶",
                   "自動登入沒成功，請在瀏覽器自己登入") == "q":
                say("SKIP: 這次不申請")
                return
        applied = rec.get("kaohsiung_applied")
        missing = [i for i, g in enumerate(groups) if not has_cert(case, g)]
        to_apply = list(range(len(groups)))                # 全部都有了還來跑＝要重新申請(過期換新)
        if applied and missing:
            # 之前送出過還沒收齊 → 先查進度、繳費頁，別重複申請
            say(f"這個案件 {applied} 申請過，先查進度…")
            since = _roc7(datetime.date.fromisoformat(applied))
            for i in missing:
                fetch(i, since)
            if all_done():
                return
            pay = kh_pay_page(drv)
            waiting = [i for i in missing if i not in got and "目前無任何案件" not in pay and groups[i]["section"] in pay]
            to_apply = [i for i in missing if i not in got and i not in waiting]
            if waiting:
                say("還沒繳費（不重新申請）：" + "；".join(_gname(groups[i]) for i in waiting))
            if to_apply and ask("查不到申請（可能沒送出或已過期）：" + "；".join(_gname(groups[i]) for i in to_apply)
                                + "\n要重新申請就按「重新申請」", "重新申請 ▶",
                                f"{applied} 的申請有 {len(to_apply)} 件查不到，要重新申請嗎？", "orange") == "q":
                to_apply = []
                if not waiting:
                    say("SKIP: 這次不申請")
                    return
        elif missing:
            to_apply = missing
        for n, i in enumerate(to_apply, 1):                 # 一個地段送一件
            g = dict(case, **groups[i])
            tag = f"第 {n}/{len(to_apply)} 件 " if len(to_apply) > 1 else ""
            step(f"程式正在送出申請 {tag}（{_gname(g)}），請稍候…")
            if not kh_open_form(drv):
                if ask("請在瀏覽器進入「網路申請」頁(有同意就按同意)，表單出來後按「表單出來了」", "表單出來了 ▶",
                       "請在瀏覽器打開「網路申請」表單") == "q":
                    say("SKIP: 這次不申請")
                    return
            ok, msg = kh_fill(drv, g, app)
            if not ok:
                say(f"（{msg}，請在瀏覽器自己選）")
            else:
                say(f"{tag}已帶入：{g['area']} {_gname(g)}")
                if msg:
                    say(f"使用分區：{msg[:150]}")
            sent, msg = (kh_submit(drv) if ok else (False, ""))
            if sent:
                say(f"✓ {tag}已送出申請" + (f"（網站訊息：{msg[:80]}）" if msg else ""))
            else:
                if msg:
                    say(f"（送出沒成功：{msg[:100]}）")
                if ask(f"請在瀏覽器檢查表單（{_gname(g)}）、按「送出申請」，送出後按「送出好了」", "送出好了 ▶",
                       f"自動送出沒成功，請在瀏覽器按「送出申請」（{_gname(g)}）") == "q":
                    say("SKIP: 這次不申請")
                    return
            rec["kaohsiung_applied"] = datetime.date.today().isoformat()
            rec.pop("kaohsiung_done", None)
            save_rec(case, rec)
        take_local()                                        # 你已經自己繳完、下載了
        if all_done():
            return
        if to_apply and "目前無任何案件" in kh_pay_page(drv):
            say("（繳費頁沒有待繳的案件：如果你已經繳過了，直接按「繳好了，下載」；沒有的話請確認申請有送出）")
        say("═" * 30)
        say("💳 請在瀏覽器繳費：選「信用卡繳款」或「QRCode」行動支付（每件 100 元起；同一天申請的可以勾在一起一次繳）")
        say("繳完約 3 分鐘可以下載 → 按「繳好了，下載」；不繳按「這次不申請」")
        since = _roc7(datetime.date.fromisoformat(rec.get("kaohsiung_applied") or datetime.date.today().isoformat()))
        retry = False
        while True:
            a = ask("等你繳費…", "再下載一次 ▶" if retry else "繳好了，下載 ▶",
                    "證明書還沒好，等 1～2 分鐘再按「再下載一次」" if retry
                    else "💳 換你了：請在瀏覽器繳費（信用卡或行動支付）", "orange" if retry else "green")
            retry = True
            if a == "q":
                say("SKIP: 這次先不繳（申請已送出，沒繳費不會核發；之後單跑這一項會接著查、下載）")
                return
            take_local()
            for i, g in enumerate(groups):
                if i not in got and not has_cert(case, g):
                    fetch(i, since)
            if all_done():
                return
            left = [g for i, g in enumerate(groups) if i not in got and not has_cert(case, g)]
            say("還沒抓到：" + "；".join(_gname(g) for g in left)
                + "。剛繳費要等約 3 分鐘；也可以自己在瀏覽器「案件進度查詢及下載」按下載，程式會接走檔案。")
            before, wins = set(os.listdir(log_dir)), set(drv.window_handles)
            p = _wait_pdf(drv, log_dir, before, wins, 90)
            if p:
                g = which_group(case, p)
                store(groups.index(g), p)
                if all_done():
                    return
            say("過一下再按「再下載一次」；或按「這次不申請」結束（之後單跑這一項會接著查，不會重新申請）")
    finally:
        _quit(drv)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base-dir", default=os.path.dirname(os.path.abspath(sys.argv[0])))
    a = ap.parse_args()
    BASE[0] = a.base_dir
    try:
        case = load_case(a.base_dir)
    except Exception as e:
        say(f"FAIL: 讀不到目前案件（{e}）")
        return
    app = load_applicant(a.base_dir)
    if not (app["name"] and app["phone"] and app["email"]):
        say("FAIL: 還沒設定申請人（名稱、電話、Email）")
        return
    if not case["lots"]:
        say("FAIL: 案件沒有地號")
        return
    log_dir = os.path.join(case["folder"], "4.其他相關", "連續模式紀錄", "_zcert")
    shutil.rmtree(log_dir, ignore_errors=True)
    os.makedirs(log_dir, exist_ok=True)
    rec = load_rec(case)
    say(f"案件：{case['city']}{case['area']} {case['section']} 地號 {'、'.join(case['lots'])}")
    if case["city"] == "高雄市":
        kaohsiung(case, app, rec, log_dir)
    elif case["city"] == "臺南市":
        tainan(case, app, rec, log_dir)
    else:
        say("FAIL: 目前只支援高雄、台南線上申請；其他縣市請到公所申請")
    shutil.rmtree(log_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
