# -*- coding: utf-8 -*-
"""台南市建築執照補查（純 HTTP，不開瀏覽器）。

為什麼要補查：國土署 NBUPIC（organ=IF0）只有台南近年的執照；老建物的執照只在
台南工務局自己的「建築執照查詢系統」blis.tainan.gov.tw，而且要用「申請當時的地號」，
重測過的地段要先到台南地政局「重測前後新舊地建號」換成舊地號再查。
  例：孔廟段 318（富華大飯店）→ 重測前 中正段一小段 110 → 查到 (071)南工字第074571號

流程：
  1. old_lots()：land-query.tainan.gov.tw oldnew.jsp（免驗證碼）新地號 → 重測前舊地號
  2. blis_query()：blis.tainan.gov.tw（4 碼驗證碼，ddddocr 辨識；每次查詢前都要重過一次）
     ★段名要去掉最後一個「段」字（網頁 conArray 會 .replace(/段$/,"")）：中正段一小段 → 中正段一小
blis 只有執照號 / 地號 / 門牌 / 起造人（遮名），沒有存根細節。

單獨測試：python tainan_permit.py 中西區 孔廟段 318
"""
import re
import sys
import time
import json

LAND_URL = "https://land-query.tainan.gov.tw/query/rwd/"
BLIS_URL = "https://blis.tainan.gov.tw"
HDR = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9",
}


def _session():
    import requests
    import urllib3
    urllib3.disable_warnings()
    s = requests.Session()
    s.headers.update(HDR)
    s.verify = False          # 台南市府憑證鏈常缺中繼憑證
    return s


def _norm(x):
    return (x or "").replace(" ", "").replace("台", "臺").strip()


def split_lot(lot_number):
    """'318' / '318-2' / '0318-0002' → ('0318', '0002')；沒子號當 0000。"""
    s = str(lot_number or "").strip().replace("－", "-")
    m = re.match(r"^(\d+)(?:-(\d+))?$", s)
    if not m:
        return None
    return m.group(1).zfill(4)[-4:], (m.group(2) or "0").zfill(4)[-4:]


def _strip_tags(h):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h or "")).strip()


# ---------- 1. 重測前後新舊地號 ----------
def old_lots(area, section, lot_number, log=print):
    """回傳重測前的 [(區, 段名, '0110-0000'), ...]；沒重測過/查不到回 []。"""
    nums = split_lot(lot_number)
    if not nums:
        return []
    s = _session()
    page = s.get(LAND_URL + "oldnew.jsp?menu=false", timeout=30).text
    m = re.search(r"csrf\.param=([0-9A-F]+)", page)
    if not m:
        log("  [台南地政] 頁面格式變了，抓不到 csrf，略過舊地號查詢")
        return []
    csrf = m.group(1)
    site = None
    for code, name in re.findall(r'<OPTION\s+VALUE="(D[A-Z]-\d+)">([^<]+)</OPTION>', page, re.I):
        if _norm(name) == _norm(area):
            site = code
            break
    if not site:
        log(f"  [台南地政] 找不到行政區「{area}」")
        return []
    r = s.get(LAND_URL + "../ajaxAreaR48.jsp", params={
        "csrf.param": csrf, "SiteArea": site, "R48": "R48", "ViewActionName": "oldnew_va",
        "dt": str(int(time.time() * 1000))}, timeout=30)
    res = r.text.strip()
    secs = [p.split("#,#") for p in res[:-3].split("#;#")] if res else []
    sect_code = None
    want = _norm(section)
    for p in secs:
        if len(p) >= 2 and (_norm(p[0]) == want or _norm(p[0]) == want + "段"):
            sect_code = p[1]
            break
    if not sect_code:
        log(f"  [台南地政] {area}找不到地段「{section}」")
        return []
    data = {"action": "Query1", "r": "G", "SiteArea": site, "R48": sect_code, "R48check": sect_code,
            "TypeCF": "C", "NUM1": nums[0], "NUM2": nums[1], "TypeON": "O"}
    html = s.post(LAND_URL + "oldnew.jsp?menu=false", data=data, timeout=30).text
    tb = re.search(r'(?s)id="queryResult".*?<tbody>(.*?)</tbody>', html)
    out = []
    for tr in re.findall(r"(?s)<tr>(.*?)</tr>", tb.group(1) if tb else ""):
        tds = [_strip_tags(x) for x in re.findall(r"(?s)<td[^>]*>(.*?)</td>", tr)]
        if len(tds) >= 8 and re.match(r"^\d{4}-\d{4}$", tds[3]):
            item = (tds[0], tds[2], tds[3])
            if item not in out:
                out.append(item)
    return out


# ---------- 2. 台南工務局建築執照查詢 ----------
_OCR = None


def _ocr(img_bytes):
    global _OCR
    if _OCR is None:
        try:
            import ddddocr
            o = ddddocr.DdddOcr(show_ad=False)
            _OCR = lambda b: (o.classification(b) or "")
        except Exception:
            import io
            import numpy as np
            from PIL import Image
            from rapidocr_helper import create_rapidocr_reader, rapidocr_readtext
            rd = create_rapidocr_reader(quiet=True)
            _OCR = lambda b: "".join(rapidocr_readtext(rd, np.array(Image.open(io.BytesIO(b)).convert("RGB")), detail=0))
    return "".join(ch for ch in _OCR(img_bytes) if ch.isalnum()).upper()[:4]


def blis_query(area, section, lot_text, log=print):
    """查一個地號（lot_text 如 '0110-0000'）。回傳 [{license_no, land, address, applicant}]；
    驗證碼一直過不了回 None（跟「查無」區分）。"""
    s = _session()
    s.get(BLIS_URL + "/BuildingSearch.aspx", timeout=30)
    hj = {"Content-Type": "application/json; charset=utf-8", "X-Requested-With": "XMLHttpRequest",
          "Referer": BLIS_URL + "/BuildingSearch.aspx"}

    def call(method, payload):
        r = s.post(BLIS_URL + "/_asmx/db_BuildingSearch.asmx/" + method,
                   data=json.dumps(payload), headers=hj, timeout=30)
        return r.json().get("d")

    for _ in range(10):
        img = s.get(BLIS_URL + "/manage/CheckImageCode.aspx", timeout=30).content
        code = _ocr(img)
        if len(code) == 4 and call("checkCaptcha", {"code": code}) == "":
            break
    else:
        log("  [台南工務局] 驗證碼多次辨識失敗")
        return None
    sect = re.sub(r"段$", "", (section or "").strip())     # ★一定要去掉最後的「段」
    rows, page, total_pages = [], 1, 1
    while page <= total_pages and page <= 20:
        d = call("sqlGoPage", {"nPage": page, "dtS": [area, sect, lot_text]})
        if not d or len(d) < 3 or not d[1]:
            break
        for rec in d[1].split("Ä"):
            rec = rec.split("Ë", 1)[-1]
            f = rec.split("Ï")
            if len(f) >= 4:
                rows.append({"license_no": f[0].strip(), "land": f[1].strip(),
                             "address": f[2].strip(), "applicant": f[3].strip()})
        try:
            total_pages = int(str(d[2]).split(",")[1])
        except Exception:
            total_pages = 1
        page += 1
        if page <= total_pages:     # 換頁也要重過驗證碼
            for _ in range(10):
                img = s.get(BLIS_URL + "/manage/CheckImageCode.aspx", timeout=30).content
                code = _ocr(img)
                if len(code) == 4 and call("checkCaptcha", {"code": code}) == "":
                    break
    return rows


def lookup(area, section, lot_number, log=print):
    """新地號 + 重測前舊地號都查。回傳 (結果列表, 舊地號列表)。
    結果每筆多一個 via：'新地號 孔廟段0318-0000' 或 '舊地號 中正段一小段0110-0000'。"""
    nums = split_lot(lot_number)
    if not nums:
        log(f"  [台南工務局] 地號格式看不懂：{lot_number}")
        return [], []
    targets = [("新地號", area, section, f"{nums[0]}-{nums[1]}")]
    try:
        olds = old_lots(area, section, lot_number, log=log)
    except Exception as e:
        log(f"  [台南地政] 查舊地號失敗：{e}")
        olds = []
    for o_area, o_sect, o_lot in olds:
        if (_norm(o_sect), o_lot) != (_norm(section), targets[0][3]):
            targets.append(("舊地號", o_area, o_sect, o_lot))
    found, seen = [], set()
    for kind, a, sct, lt in targets:
        log(f"  查台南工務局：{kind} {a}{sct} {lt}")
        try:
            rows = blis_query(a, sct, lt, log=log)
        except Exception as e:
            log(f"  [台南工務局] 查詢失敗：{e}")
            rows = None
        if rows is None:
            continue
        log(f"    → {len(rows)} 筆")
        for r in rows:
            if r["license_no"] in seen:
                continue
            seen.add(r["license_no"])
            r["via"] = f"{kind} {sct}{lt}"
            found.append(r)
    return found, olds


def is_usage(license_no):
    """字號看起來是使用執照（使/更使/變使；民國 70 年代的「南工字」多為使照但無法確定，不算）。"""
    return "使" in (license_no or "")


def results_html(title, rows, olds, note=""):
    """把結果做成一頁可列印的 HTML（給 Chrome printToPDF 用）。"""
    import html as _h
    e = _h.escape
    old_txt = "、".join(f"{a}{s} {l}" for a, s, l in olds) or "沒有重測紀錄（或查不到）"
    trs = "".join(
        f"<tr{' class=u' if is_usage(r['license_no']) else ''}><td>{i}</td><td>{e(r['license_no'])}</td>"
        f"<td>{e(r['address'])}</td><td>{e(r['land'])}</td><td>{e(r['applicant'])}</td><td>{e(r['via'])}</td></tr>"
        for i, r in enumerate(rows, 1))
    return f"""<!doctype html><html><head><meta charset="utf-8"><title>{e(title)}</title>
<style>
body{{font-family:"Microsoft JhengHei","微軟正黑體",sans-serif;color:#111;margin:18px;font-size:14px}}
h1{{font-size:21px;margin:0 0 6px}} .m{{color:#333;margin:2px 0}}
table{{border-collapse:collapse;width:100%;margin-top:12px}}
th,td{{border:1px solid #888;padding:6px 7px;vertical-align:top;text-align:left}}
th{{background:#e8eef6}} td:nth-child(5),th:nth-child(5){{white-space:nowrap}} tr.u td{{background:#e3f0ff;font-weight:bold}}
.n{{margin-top:12px;color:#333;font-size:13px;line-height:1.6}}
</style></head><body>
<h1>{e(title)}</h1>
<div class="m">重測前舊地號：{e(old_txt)}</div>
<div class="m">查詢時間：{time.strftime('%Y-%m-%d %H:%M')}</div>
<table><tr><th>#</th><th>執照號碼</th><th>建築地址</th><th>地號</th><th>起造人</th><th>用哪個地號查到</th></tr>{trs}</table>
<div class="n">資料來源：臺南市政府工務局「建築執照查詢系統」blis.tainan.gov.tw；重測前後地號：臺南市政府地政局「重測前後新舊地建號」。<br>
藍底＝字號有「使」字的使用執照。此系統只列執照號、地號、門牌、起造人，存根內容需向工務局申請調閱。{e(note)}</div>
</body></html>"""


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    a = sys.argv[1:]
    if len(a) < 3:
        print(__doc__)
        sys.exit(1)
    rows, olds = lookup(a[0], a[1], a[2])
    print("舊地號：", olds)
    for r in rows:
        print(r)
