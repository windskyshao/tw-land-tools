# -*- coding: utf-8 -*-
"""
實價登錄申報書(租賃) 產生器
- 讀案件資料夾：4.其他相關\\data_final.json、trans_*.json(謄本結構化)、都市計畫資料彙整.json、
  6.成交\\...\\租賃契約(docx/pdf)
- 跳出確認視窗讓使用者檢查/修改 → 產生「每格獨立表單欄位」的可編輯 PDF 存到 6.成交
用法：rent_report.py --case-dir <案件資料夾> [--base-dir <主程式目錄>]
"""
import os
import re
import sys
import glob
import json
import argparse
from datetime import datetime

import fitz

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# 打包版 python_embedded 有 python39._pth → 不自動加腳本目錄，要自己加才 import 得到 deal_ui
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)
TEMPLATE_NAME = "實價登錄申報書_租_範本.pdf"
SETTINGS_NAME = "rent_report_settings.json"

DEFAULT_APPLICANT = {
    "類別": "不動產經紀業",
    "名稱": "富住通不動產仲介經紀股份有限公司鳳山分公司",
    "統一編號": "27217312",
    "電話": "07-382-7999",
    "地址": "高雄市三民區九如一路291號",
    "電子信箱": "",
}

EQUIPMENT = ["冷氣", "瓦斯或天然氣", "熱水器", "有線電視", "洗衣機", "網路", "電視機", "傢俱", "冰箱"]
BUILDING_TYPES = ["公寓(無電梯)", "華廈(10層含以下有電梯)", "透天厝", "店面(店舖)", "工廠",
                  "辦公商業大樓", "住宅大樓(11層含以上有電梯)", "廠辦", "農舍", "倉庫", "其他"]
USES = ["居住用", "非居住用", "混合用"]
LEASE_TYPES = ["整棟(戶)出租", "分層出租", "獨立套房", "分租套房", "分租雅房"]
ZONES = ["住", "商", "工", "農", "其他"]
FEES = ["管理費", "水費", "電費", "瓦斯費", "網路費"]
PARK_TYPES = ["①坡道平面", "②升降平面", "③坡道機械", "④升降機械", "⑤塔式車位", "⑥一樓平面", "⑦其他"]
PARK_XY = [[(101.9, 334.0), (198.3, 334.9), (294.5, 334.9), (390.9, 334.9), (101.9, 348.1), (198.3, 347.6), (294.5, 347.6)],
           [(101.9, 363.3), (198.3, 363.3), (294.5, 363.3), (390.9, 363.3), (101.9, 376.0), (198.3, 376.0), (294.5, 376.0)]]
SERVICES = ["無", "一般包租", "一般轉租", "一般代管", "社會住宅包租轉租", "社會住宅代管"]
PARK_PRICING = ["", "單獨計價", "未單獨計算(已含入租金總額)"]
# 24 備註(中/右欄)：(勾選框名稱, 視窗顯示文字)
REMARK_MID = [("多門牌", "有多個門牌且未個別計算租金"), ("增建", "含增建或未登記建物"),
              ("部分範圍", "租賃建物部分範圍，且未計入共有部分面積"), ("未租車位", "未租賃車位，建物登記含有車位面積"),
              ("不同租金", "租賃期間有約定不同租金")]
REMARK_RIGHT = [("親友", "親友、員工、共有人或其他特殊關係間之租賃"), ("急租", "急出租/急承租"),
                ("債務", "受債權債務關係影響之租賃"), ("民俗", "有民情風俗因素之租賃"),
                ("續租", "續租案件"), ("營業設備", "含營業設備"), ("外牆", "外牆作為廣告使用"),
                ("屋頂", "屋頂(突出物)作為基地台或廣告使用"), ("其他4", "其他")]
# 需要「請敘明」文字的項目 → 文字欄 key
REMARK_TEXT = {"費_其他": "其他費用", "多門牌": "全部門牌", "不同租金": "不同租金情形", "其他4": "其他說明"}


# ============================================================
# 小工具
# ============================================================
_FW = str.maketrans("０１２３４５６７８９－．，", "0123456789-.,")
_CN_DIGIT = {"零": 0, "〇": 0, "一": 1, "二": 2, "兩": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
             "壹": 1, "貳": 2, "參": 3, "叁": 3, "肆": 4, "伍": 5, "陸": 6, "柒": 7, "捌": 8, "玖": 9}
_CN_UNIT = {"十": 10, "拾": 10, "百": 100, "佰": 100, "千": 1000, "仟": 1000}


def fw2hw(s):
    return (s or "").translate(_FW)


def cn2int(s):
    """中文/阿拉伯數字混寫轉整數：十五→15、1萬7仟→17000、壹萬柒仟→17000；失敗回 None"""
    s = fw2hw(str(s)).replace(",", "").strip()
    if not s:
        return None
    if s.isdigit():
        return int(s)
    total, section, num = 0, 0, None
    i = 0
    while i < len(s):
        ch = s[i]
        if ch.isdigit():
            j = i
            while j < len(s) and s[j].isdigit():
                j += 1
            num = int(s[i:j])
            i = j
            continue
        if ch in _CN_DIGIT:
            num = _CN_DIGIT[ch]
        elif ch in _CN_UNIT:
            section += (num if num is not None else 1) * _CN_UNIT[ch]
            num = None
        elif ch in "萬万":
            section += num or 0
            total += section * 10000
            section, num = 0, None
        else:
            return None
        i += 1
    return total + section + (num or 0)


def frac(s):
    """'10000分之21' → 0.0021"""
    m = re.search(r"(\d+)\s*分之\s*(\d+)", fw2hw(s or "").replace(" ", ""))
    return int(m.group(2)) / int(m.group(1)) if m else None


def num(s):
    m = re.search(r"[\d,]+(?:\.\d+)?", fw2hw(s or ""))
    return float(m.group(0).replace(",", "")) if m else None


def lot_no(s):
    """'0013-0000' → '13-0'；'02589-000' → '2589-0'"""
    m = re.search(r"(\d+)-(\d+)", fw2hw(s or ""))
    return f"{int(m.group(1))}-{int(m.group(2))}" if m else ""


def short_county(s):
    return re.sub(r"[市縣]$", "", s or "")


def split_address(addr):
    """門牌拆欄位。別用貪婪/非貪婪亂切行政區：先找「區」，再退回 2 字+鄉鎮市"""
    a = fw2hw(addr or "").replace(" ", "")
    out = {k: "" for k in ("縣市", "區", "路街", "段", "巷", "弄", "號", "樓", "室")}
    m = re.match(r"^(..[市縣])", a)
    if m:
        out["縣市"] = m.group(1)
        a = a[len(m.group(1)):]
    for pat in (r"^(.{1,3}?區)", r"^(..[鄉鎮市])", r"^(.{1,3}?[鄉鎮市])"):
        m = re.match(pat, a)
        if m and not re.search(r"[路街巷弄號]", m.group(1)):
            out["區"] = m.group(1)
            a = a[len(m.group(1)):]
            break
    a = re.sub(r"^[^\d路街段巷弄號]{1,4}[里村](\d+鄰)?", "", a)       # 去掉 xx里xx鄰
    a = re.sub(r"^\d+鄰", "", a)
    m = re.match(r"^(.+?(?:大道|路|街))", a)
    if m:
        out["路街"] = m.group(1)
        a = a[len(m.group(1)):]
    for key, pat in (("段", r"^([\d一二三四五六七八九十]+)段"), ("巷", r"^(\d+)巷"), ("弄", r"^(\d+)弄"),
                     ("號", r"^([\d之\-]+)號"), ("樓", r"^([\d一二三四五六七八九十]+(?:之\d+)?)樓"),
                     ("室", r"^(?:之)?(\d+)室?")):
        m = re.match(pat, a)
        if m:
            v = m.group(1)
            if key == "樓" and not v[0].isdigit():
                v = str(cn2int(v) or v)
            a = a[m.end():]
            if key == "樓":                       # 7樓之1 → 樓欄填「7之1」
                m2 = re.match(r"^之(\d+)", a)
                if m2:
                    v += "之" + m2.group(1)
                    a = a[m2.end():]
            out[key] = v
    if not out["路街"] and a:  # 沒路街(例如只有村里鄰號)：剩下的放路街欄給人工看
        out["路街"] = a
    return out


def roc_date(y, m, d):
    return {"年": str(int(y)), "月": str(int(m)), "日": str(int(d))}


# ============================================================
# 讀案件資料
# ============================================================
def _load_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


_LEASE_EXCLUDE = re.compile(r"委託|封面|申報書|謄本|服務費|身分證|業績|說明書|水電|瓦斯|同意書|現況|協議書|點交|收據|切結|帳單")
_LEASE_KEY = re.compile(r"租|約|契")


def _pdf_text_chars(path, max_pages=3):
    try:
        d = fitz.open(path)
        return sum(len(d[i].get_text().strip()) for i in range(min(max_pages, len(d)))), len(d)
    except Exception:
        return 0, 0


def find_lease_candidates(case_dir):
    """6.成交 底下可能是租約的檔案。回傳 [{path, kind(文字檔/掃描檔), pages}]，排好順序(第一個=預設)
    排序：可直接讀文字的(docx/文字PDF)優先、同類取最新；掃描檔取頁數最多(合約通常最厚)"""
    deal = os.path.join(case_dir, "6.成交")
    out = []
    for root, _, files in os.walk(deal):
        for fn in files:
            low = fn.lower()
            if fn.startswith("~$") or not low.endswith((".docx", ".pdf")):
                continue
            if _LEASE_EXCLUDE.search(fn):
                continue
            stem = os.path.splitext(fn)[0]
            scanner_name = bool(re.fullmatch(r"[\d_\-\s]{8,}", stem))   # 掃描器自動命名(全數字)
            if not (_LEASE_KEY.search(fn) or scanner_name):
                continue
            p = os.path.join(root, fn)
            if low.endswith(".docx"):
                kind, pages = "文字檔", 0
            else:
                chars, pages = _pdf_text_chars(p)
                kind = "文字檔" if chars > 200 else "掃描檔"
            out.append({"path": p, "kind": kind, "pages": pages, "mtime": os.path.getmtime(p),
                        "unnamed": scanner_name})
    texts = sorted([c for c in out if c["kind"] == "文字檔"], key=lambda c: -c["mtime"])
    scans = sorted([c for c in out if c["kind"] == "掃描檔" and not c["unnamed"]], key=lambda c: (-c["pages"], -c["mtime"]))
    unnamed = sorted([c for c in out if c["unnamed"] and c["kind"] == "掃描檔"], key=lambda c: -c["mtime"])
    return texts + scans + unnamed


def find_lease_file(case_dir):
    """自動挑：檔名看得出是合約的才挑；掃描器數字檔名的看不出內容(可能是帳單/證件)，留給使用者選"""
    c = [x for x in find_lease_candidates(case_dir) if not x.get("unnamed")]
    return c[0]["path"] if c else None


# OCR 常把繁體認成簡體：只轉申報書會用到的關鍵字
_S2T = str.maketrans("号楼证赁约万电话书与为费设备气热视机网线间双门厅卫层区乡镇对应经时宝务权签预发台盖们贰参陆谢陈张刘吴黄缴纳满", "號樓證賃約萬電話書與為費設備氣熱視機網線間雙門廳衛層區鄉鎮對應經時寶務權簽預發臺蓋們貳參陸謝陳張劉吳黃繳納滿")


def _ocr_cache_path(path):
    """快取放程式自己的 ocr_cache 資料夾(別寫進共享的成交資料夾)"""
    import hashlib
    key = hashlib.md5(os.path.normcase(os.path.abspath(path)).encode("utf-8")).hexdigest()
    d = os.path.join(_settings_dir or SCRIPT_DIR, "ocr_cache")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, key + ".txt")


def ocr_pdf_lines(path, progress=None):
    """掃描檔 OCR(rapidocr)，結果快取在同資料夾 <檔名>.ocr.txt(檔案沒變就直接用)"""
    cache = _ocr_cache_path(path)
    if os.path.exists(cache) and os.path.getmtime(cache) >= os.path.getmtime(path):
        with open(cache, "r", encoding="utf-8") as f:
            return [l for l in f.read().splitlines() if l.strip()]
    import numpy as np
    from rapidocr_onnxruntime import RapidOCR
    engine = RapidOCR()
    doc = fitz.open(path)
    lines = []
    for i, page in enumerate(doc):
        if progress:
            progress(f"辨識掃描檔 第 {i + 1}/{len(doc)} 頁…")
        pix = page.get_pixmap(dpi=200)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)[:, :, :3]
        res, _ = engine(img)
        lines += [r[1] for r in (res or [])]
    lines = [l.translate(_S2T) for l in lines if l.strip()]
    try:
        with open(cache, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    except Exception:
        pass
    return lines


def read_lease_lines(path, progress=None):
    if path.lower().endswith(".docx"):
        import docx
        d = docx.Document(path)
        lines = [p.text for p in d.paragraphs]
        for t in d.tables:
            for r in t.rows:
                lines.append(" ".join(c.text for c in r.cells))
        return [l for l in lines if l.strip()], "文字檔"
    chars, _ = _pdf_text_chars(path)
    if chars > 200:
        doc = fitz.open(path)
        text = "\n".join(p.get_text() for p in doc)
        return [l for l in text.splitlines() if l.strip()], "文字檔"
    return ocr_pdf_lines(path, progress), "掃描檔"


_NUM = r"[\d一二三四五六七八九十壹貳參叁肆伍陸柒捌玖拾佰百零〇]+"
_MONEY = r"[\d,，一二三四五六七八九十壹貳參叁肆伍陸柒捌玖拾零萬仟千佰百]+"


def parse_lease(path, progress=None):
    """回傳 dict：承租人[]、起訖、訂約日、租金、管理費由誰付、附屬設備、來源類型
    兼容：自家類公證範本(承租人1(以下簡稱乙方)：xxx)、法院公證書(承租人：xxx女民國…)、OCR 文字"""
    lines, kind = read_lease_lines(path, progress)
    text = "\n".join(lines)
    flat = re.sub(r"\s+", "", text)
    res = {"承租人": [], "設備": set(), "冷氣台數": "", "類型": kind}

    name_re = re.compile(r"承租人\s*\d*\s*(?:[（(]以下簡稱乙方[)）])?\s*[：:]?\s*([一-鿿]{2,4}?)(?=\s|男|女|民國|先生|小姐|$|，|,)")
    id_re = re.compile(r"(?:統一編號|身分證字號|身分證統一編號|身分證號|國民身分證)\s*[：:]?\s*([A-Za-z][12]\d{8}|\d{8})")
    cur = None
    seen = set()
    for ln in lines:
        s = ln.strip().lstrip("：:、.． ")
        m = name_re.search(s) if s.startswith("承租人") else None
        # OCR 雜訊：跟已抓到的名字有兩個字重疊(例：謝東奇 vs 東奇奇)就當同一人
        if m and any(len(set(m.group(1)) & set(x)) >= 2 for x in seen):
            m = None
            cur = None
            continue
        if m and m.group(1) not in seen:
            seen.add(m.group(1))
            cur = {"姓名": m.group(1), "統編": "", "電話": "", "地址": ""}
            mid = id_re.search(s)
            if mid:
                cur["統編"] = mid.group(1).upper()
            res["承租人"].append(cur)
            continue
        if re.match(r"^(出租人|出租方|立契約|今甲|連帶保證人|保證人|二、|二．)", s):
            cur = None
            continue
        if cur is not None:
            if not cur["統編"]:
                mid = id_re.search(s)
                if mid:
                    cur["統編"] = mid.group(1).upper()
            if s.startswith("地址") or s.startswith("住址"):
                ma = re.match(r"^[地住]址\s*[：:]\s*(.+?)(?:\s*電話\s*[：:]\s*([\d\-\s+]+))?$", s)
                if ma:
                    cur["地址"] = ma.group(1).strip()
                    if ma.group(2):
                        cur["電話"] = re.sub(r"\s", "", ma.group(2))
    for t in res["承租人"]:
        if not t["電話"]:
            m = re.search(re.escape(t["姓名"]) + r".{0,200}?電話\s*[：:]\s*([\d\-]{9,})", flat)
            if m:
                t["電話"] = m.group(1)

    m = re.search(rf"自?(?:民國|中華民國)({_NUM})年({_NUM})月({_NUM})日起?[，,]?至(?:民國|中華民國)?({_NUM})年({_NUM})月({_NUM})日止", flat) or         re.search(rf"({_NUM})年({_NUM})月({_NUM})日起[，,]?至(?:民國|中華民國)?({_NUM})年({_NUM})月({_NUM})日止", flat)
    if m:
        g = [cn2int(x) for x in m.groups()]
        if all(x is not None for x in g):
            res["起"] = roc_date(*g[:3])
            res["迄"] = roc_date(*g[3:])
    dates = re.findall(r"中華民國(\d{2,3})年(\d{1,2})月(\d{1,2})日", flat)
    if dates:
        res["訂約"] = roc_date(*dates[-1])

    m = re.search(rf"租金(?:每[個]?月)?[^。；\n]{{0,8}}?新[台臺]幣\s*(?:[（(]以下同[)）])?\s*({_MONEY})\s*元", flat) or \
        re.search(rf"每[個]?月租金(?:新[台臺]幣)?(?:[（(]以下同[)）])?({_MONEY})元", flat) or         re.search(rf"月租金(?:新[台臺]幣)?[^。元\d壹貳參肆伍陸柒捌玖一二三四五六七八九]{{0,6}}({_MONEY})元", flat)
    if m:
        res["租金"] = cn2int(m.group(1).replace("，", ""))
    m = re.search(r"管理費[^。]{0,6}?由\s*([甲乙])\s*方負擔", flat) or re.search(r"管理費[^。]{0,6}?由(出租人|承租人)負擔", flat)
    if m:
        res["管理費由"] = "出租人" if m.group(1) in ("甲", "出租人") else "承租人"

    m = re.search(r"附屬?設備(?:計有|如下|包括|包含)[：:]?(.+?)。", flat) or         re.search(r"附屬?設備[：:](.{1,150}?)(?:。|第[一二三四五六七八九十]+條)", flat)
    eq_text = m.group(1) if m else ""
    if eq_text:
        rules = [("冷氣", r"冷氣"), ("熱水器", r"熱水器"), ("洗衣機", r"洗衣機"), ("電視機", r"電視(?!台)"),
                 ("冰箱", r"冰箱"), ("瓦斯或天然氣", r"瓦斯|天然氣"), ("有線電視", r"有線電視|第四台"),
                 ("網路", r"網路"), ("傢俱", r"床|沙發|衣櫥|衣櫃|書桌|餐桌|桌|椅|櫃")]
        for name, pat in rules:
            if re.search(pat, eq_text):
                res["設備"].add(name)
        m = re.search(r"冷氣[^\d、，；]*?(\d+)\s*[台組]", eq_text)
        if m:
            res["冷氣台數"] = m.group(1)
    return res


def gather(case_dir, lease_path=None, progress=None):
    """把能找到的資料整理成申報書欄位值；回傳 (values, notes)
    lease_path：使用者指定的租約檔(None=自動挑)；progress：OCR 進度回報 callback"""
    other = os.path.join(case_dir, "4.其他相關")
    df_path = os.path.join(other, "data_final.json")
    urban_path = os.path.join(other, "都市計畫資料彙整.json")
    df = _load_json(df_path) or {}
    urban = _load_json(urban_path) or {}
    trans_files = sorted(glob.glob(os.path.join(other, "trans_*.json")), key=os.path.getmtime)
    trans = _load_json(trans_files[-1]) if trans_files else []
    trans = trans if isinstance(trans, list) else []
    notes = []      # (等級, 訊息)  等級: "warn" 要人看 / "info" 來源說明
    rel = lambda p: os.path.relpath(p, case_dir) if p else ""
    sources = [  # (名稱, 檔案相對路徑, 狀態, 用途)
        ("物件資料", rel(df_path), "✔ 已讀" if df else "✖ 找不到", "門牌、租金、型態、電梯、車位、格局、管理費"),
        ("謄本結構化", rel(trans_files[-1]) if trans_files else "4.其他相關\\trans_*.json",
         "✔ 已讀" if trans else "✖ 找不到(先跑謄本結構化)", "土地/建物標的、面積、樓層"),
        ("都市計畫", rel(urban_path), "✔ 已讀" if urban else "— 沒有(用物件資料的分區)", "使用分區"),
    ]

    v = {"申報人": dict(DEFAULT_APPLICANT)}
    st = load_settings()
    if st.get("申報人"):
        v["申報人"].update(st["申報人"])

    # ---------- 門牌 ----------
    city_full = ""
    m = re.match(r"^(..[市縣])", df.get("土地標示", "") or "")
    if m:
        city_full = m.group(1)
    addr = fw2hw(df.get("建物門牌", ""))
    if addr and not re.match(r"^..[市縣]", addr) and city_full:
        addr = city_full + addr
    v["建物門牌"] = addr
    if not addr:
        notes.append(("warn", "找不到建物門牌(data_final.json)，請手動填"))

    # ---------- 謄本：土地/建物 ----------
    lands, bldgs = [], []
    floors_total, floor_no = "", ""
    for t in trans:
        kind = t.get("謄本類型", "")
        info = t.get("基本資訊", {}) or {}
        mark = t.get("標示部", {}) or {}
        ident = info.get("地號建號", "")
        mi = re.match(r"^(.{1,3}?區|..[鄉鎮市])?\s*(.+?段(?:.+?小段)?)\s*([\d\-]+)", ident)
        district = (mi.group(1) or "") if mi else ""
        section = mi.group(2) if mi else ""
        number = lot_no(mi.group(3)) if mi else ""
        if "土地" in kind:
            area = num(mark.get("面積", ""))
            share = sum((frac(o.get("權利範圍", "")) or 0) for o in t.get("所有權部", []) or [])
            lease_area = round(area * share, 2) if area and share else ""
            lands.append({"縣市": short_county(city_full), "區": district, "段": section, "地號": number,
                          "面積": f"{lease_area:.2f}" if lease_area != "" else "", "分區": ""})
        elif "建物" in kind:
            total = 0.0
            main_a = num(mark.get("總面積", "")) or 0
            total += main_a
            for k, val in mark.items():
                if k.startswith("附屬建物面積"):
                    total += num(val) or 0
            common_keys = sorted(k for k in mark if k.startswith("共有部分面積"))
            for ck in common_keys:
                suffix = ck[len("共有部分面積"):]
                ca, cr = num(mark.get(ck, "")), frac(mark.get("權利範圍" + suffix, ""))
                if ca and cr:
                    total += ca * cr
            bldgs.append({"縣市": short_county(city_full), "區": district, "段": section, "建號": number,
                          "面積": f"{total:.2f}" if total else ""})
            if not floors_total:
                floors_total = str(cn2int(re.sub(r"層", "", fw2hw(mark.get("層數", "")))) or "")
            if not floor_no:
                lv = re.sub(r"層$", "", mark.get("層次", ""))
                floor_no = str(cn2int(lv) or "") if lv else ""
    if not trans:
        notes.append(("warn", "找不到謄本結構化資料(trans_*.json)，土地/建物標的請手動填"))
    if len(lands) > 2 or len(bldgs) > 2:
        notes.append(("warn", f"土地 {len(lands)} 筆、建物 {len(bldgs)} 筆，申報書只有 2 列，多的請另附清冊"))

    zone_txt = df.get("都市土地使用分區", "") or ""
    if not zone_txt:
        for z in urban.get("都市計畫圖街廓資訊", []) or urban.get("分區查詢資料", []) or []:
            zone_txt = z.get("使用分區完整名稱") or z.get("使用分區") or ""
            if zone_txt:
                break
    zone = next((z for z, kw in (("住", "住"), ("商", "商"), ("工", "工"), ("農", "農")) if kw in zone_txt),
                "其他" if zone_txt else "")
    for ld in lands:
        ld["分區"] = zone
        ld["分區其他"] = zone_txt if zone == "其他" else ""
    if not zone:
        notes.append(("warn", "找不到都市土地使用分區(非都市土地免勾)，請確認"))
    v["土地"] = lands[:2]
    v["建物"] = bldgs[:2]

    # ---------- 物件資料 ----------
    floors_total = floors_total or str(df.get("地上層數", "") or "")
    if not floor_no:
        floor_no = split_address(addr).get("樓", "")
    v["總樓層"], v["租賃層次"] = floors_total, floor_no

    lay = df.get("格局", {}) or {}
    v["房"], v["廳"], v["衛"] = (str(lay.get(k, "") or "") for k in ("房", "廳", "衛"))
    for k in ("房", "廳", "衛"):
        n = cn2int(v[k]) if v[k] else None
        if n is not None and n > 9:
            notes.append(("warn", f"物件資料的格局「{k}」= {v[k]}，數字怪怪的(打錯？)，請確認"))
    v["無隔間"] = False

    typ, use = df.get("型式", ""), df.get("用途", "")
    fl = cn2int(floors_total) or 0
    elev = df.get("電梯", "") == "有"
    if use == "店面":
        btype = "店面(店舖)"
    elif use == "廠房":
        btype = "工廠"
    elif use == "辦公":
        btype = "辦公商業大樓"
    elif typ == "透天":
        btype = "透天厝"
    elif typ == "公寓" or (typ in ("華廈", "大樓") and not elev):
        btype = "公寓(無電梯)"
    elif typ == "華廈" or (typ == "大樓" and fl and fl <= 10):
        btype = "華廈(10層含以下有電梯)"
    elif typ == "大樓":
        btype = "住宅大樓(11層含以上有電梯)"
    else:
        btype = ""
        notes.append(("warn", "無法判斷建物型態，請選"))
    v["建物型態"] = btype
    v["用途"] = "居住用" if use in ("住家", "") else ("非居住用" if use in ("店面", "辦公", "廠房") else "")
    v["出租型態"] = "整棟(戶)出租" if v["用途"] == "居住用" else ""
    v["建物棟數"] = "1"

    v["電梯"] = elev
    v["管理組織"] = bool(num(df.get("管理費", "")))
    v["管理員"] = typ == "大樓"
    car = df.get("車位", {}) or {}
    v["車位"] = "有" if car.get("有無") == "有" else "無"
    v["車位個數"] = "1" if v["車位"] == "有" else ""
    v["車位型式"] = car.get("型式", "")
    v["車位計價"], v["車位租金總額"], v["車位清冊"] = "", "", []
    if v["車位"] == "有":
        v["車位計價"] = "未單獨計算(已含入租金總額)"
        kind = "①坡道平面" if "平面" in v["車位型式"] else ("③坡道機械" if "機械" in v["車位型式"] else "")
        v["車位清冊"] = [{"序號": "1", "類別": kind, "其他": "", "租金": "", "面積": "", "樓層": ""}]
        notes.append(("warn", "物件有車位：請確認 16 計價方式、23 車位類別(坡道/升降)、面積、樓層"))
    v["租賃住宅服務"] = "無"
    v["代理人"] = dict(st.get("代理人") or {})
    v["承租人信箱"] = ""
    v["土地筆數"] = v["房間間數"] = v["建物型態其他"] = ""
    rent0 = num(str(df.get("租金", "") or ""))
    if rent0 is not None and rent0 < 1000:          # 物件資料租金有時以「萬」為單位(1.62 = 16,200)
        rent0 = rent0 * 10000
    v["租金"] = str(int(round(rent0))) if rent0 else ""

    # ---------- 租約 ----------
    v["承租人"] = []
    v["設備"] = set()
    v["冷氣台數"] = ""
    v["起"] = v["迄"] = v["訂約"] = None
    v["含管理費"] = False
    cands = find_lease_candidates(case_dir)
    lease = lease_path or find_lease_file(case_dir)
    if not lease and cands:
        notes.append(("warn", f"6.成交 有 {len(cands)} 個掃描檔，但檔名看不出哪個是合約(可能是帳單/證件)，"
                              "請在下方「📂 讀取的資料來源」→「換一份租約」選擇"))
    if lease and all(os.path.normcase(c["path"]) != os.path.normcase(lease) for c in cands):
        cands.insert(0, {"path": lease, "kind": "", "pages": 0, "mtime": 0})
    v["租約候選"] = [c["path"] for c in cands]
    v["租約檔"] = lease or ""
    lease_status = "✖ 找不到(可按「選其他檔案」指定)"
    if lease:
        try:
            L = parse_lease(lease, progress)
            lease_status = "✔ 已讀(文字檔)" if L["類型"] == "文字檔" else "⚠ 掃描檔，已用 OCR 辨識"
            if L["類型"] == "掃描檔":
                notes.append(("warn", "租約是掃描檔，內容是 OCR 辨識的，承租人/身分證/電話/日期/金額請逐一核對"))
            v["承租人"] = L["承租人"]
            v["設備"] = L["設備"]
            v["冷氣台數"] = L["冷氣台數"]
            v["起"], v["迄"], v["訂約"] = L.get("起"), L.get("迄"), L.get("訂約")
            if L.get("租金"):
                if v["租金"] and str(L["租金"]) != v["租金"]:
                    if L["類型"] == "掃描檔":   # OCR 常漏字，數字以物件資料為準
                        notes.append(("warn", f"掃描租約辨識到租金 {L['租金']:,}，與物件資料 {int(v['租金']):,} 不同，"
                                              "先用物件資料，請對照合約"))
                    else:
                        notes.append(("warn", f"租約租金 {L['租金']:,} 與物件資料 {int(v['租金']):,} 不同，已採租約"))
                        v["租金"] = str(L["租金"])
                else:
                    v["租金"] = str(L["租金"])
            v["含管理費"] = L.get("管理費由") == "出租人"
            if not v["承租人"]:
                notes.append(("warn", "租約裡抓不到承租人，請手動填"))
            if not v["起"]:
                notes.append(("warn", "租約裡抓不到租賃期間，請手動填"))
        except Exception as e:
            lease_status = f"✖ 讀取失敗：{e}"
            notes.append(("warn", f"讀租約失敗：{e}"))
    else:
        notes.append(("warn", "6.成交 裡找不到租約，承租人/租期請手動填，或按「選其他檔案」指定合約"))
    sources.append(("租約", rel(lease), lease_status, "承租人、租期、租金、附屬設備、管理費由誰付"))
    sources.append(("申報人/代理人", SETTINGS_NAME, "✔ 上次填的" if load_settings() else "— 第一次用(預設值)", "01、02 欄"))
    v["來源"] = sources
    if not v["訂約"]:
        notes.append(("warn", "租約上的簽約日期是空白(多半手寫)，請填「訂約日期」"))
    if v["含管理費"]:
        notes.append(("info", "租約寫管理費由出租人負擔 → 備註勾「租金含管理費」"))
    return v, notes


# ============================================================
# 設定(申報人資料記憶)
# ============================================================
_settings_dir = None


def settings_path():
    return os.path.join(_settings_dir or SCRIPT_DIR, SETTINGS_NAME)


def load_settings():
    return _load_json(settings_path()) or {}


def save_settings(d):
    try:
        with open(settings_path(), "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ============================================================
# 產生 PDF
# ============================================================
B = lambda y: y + 10.0      # 12pt 字框上緣 → 基線

# 勾選框位置(範本上 □ 字元左上角)
BOX = {
    0: {"經紀業": (68.5, 151.4), "包租業": (68.5, 165.4),
        "居住用": (185.9, 360.9), "非居住用": (248.3, 360.5), "混合用": (316.7, 360.5),
        "整棟(戶)出租": (526.6, 353.0), "分層出租": (618.5, 352.6), "獨立套房": (686.3, 352.6),
        "分租套房": (526.6, 366.8), "分租雅房": (622.6, 366.8),
        "棟數_土地": (185.9, 386.0), "棟數_建物": (302.3, 386.9), "棟數_房間": (456.5, 386.0),
        "無隔間": (320.0, 433.8),
        "公寓(無電梯)": (185.8, 454.3), "華廈(10層含以下有電梯)": (185.8, 466.3), "透天厝": (311.9, 454.3),
        "店面(店舖)": (387.5, 454.3), "工廠": (387.5, 466.3), "辦公商業大樓": (490.0, 454.7),
        "住宅大樓(11層含以上有電梯)": (584.2, 455.1), "廠辦": (490.0, 467.3), "農舍": (533.8, 467.3),
        "倉庫": (582.4, 467.7), "其他": (625.6, 467.3),
        "冷氣": (185.9, 486.0), "瓦斯或天然氣": (185.9, 498.4), "熱水器": (299.3, 486.4),
        "有線電視": (299.3, 498.0), "洗衣機": (387.5, 486.4), "網路": (387.5, 498.0),
        "電視機": (463.1, 486.0), "傢俱": (463.1, 498.4), "冰箱": (535.4, 486.4), "無上述設備": (535.4, 498.4)},
    1: {"無車位": (425.0, 60.9), "有車位": (425.0, 75.2),
        "管理組織_有": (196.9, 111.2), "管理組織_無": (232.9, 110.8),
        "管理員_有": (439.0, 112.9), "管理員_無": (469.0, 112.5),
        "電梯_有": (644.4, 112.0), "電梯_無": (674.4, 111.6),
        "服務_無": (604.1, 137.4),
        "分區1_住": (589.8, 195.9), "分區1_商": (614.9, 195.6), "分區1_工": (640.1, 195.6),
        "分區1_農": (665.3, 195.6), "分區1_其他": (690.5, 194.6),
        "分區2_住": (590.2, 212.4), "分區2_商": (615.4, 212.4), "分區2_工": (640.6, 212.4),
        "分區2_農": (665.8, 212.4), "分區2_其他": (691.0, 211.5),
        "單獨計價": (527.0, 76.2), "未單獨計算": (527.0, 90.1),
        "服務_一般包租": (185.9, 137.1), "服務_一般轉租": (254.9, 137.1), "服務_一般代管": (323.3, 137.1),
        "服務_社會住宅包租轉租": (391.7, 137.1), "服務_社會住宅代管": (510.5, 137.1),
        **{f"車位{r + 1}_{t}": xy for r in range(2) for t, xy in zip(PARK_TYPES, PARK_XY[r])},
        "多人承租": (70.7, 418.5), "含衍生費用": (71.3, 455.4),
        "費_管理費": (84.5, 467.4), "費_水費": (140.9, 467.0), "費_電費": (184.7, 467.0),
        "費_瓦斯費": (84.5, 478.5), "費_網路費": (140.9, 478.5), "費_其他": (84.5, 490.5),
        "含稅": (71.3, 442.5),
        "多門牌": (272.1, 422.0), "增建": (272.1, 446.5), "部分範圍": (272.1, 458.5),
        "未租車位": (272.1, 482.0), "不同租金": (272.1, 494.0),
        "親友": (519.3, 418.5), "急租": (519.3, 430.5), "債務": (519.3, 442.5), "民俗": (519.3, 454.5),
        "續租": (519.3, 478.5), "營業設備": (588.3, 478.5), "外牆": (669.3, 478.5),
        "屋頂": (519.3, 490.5), "其他4": (519.3, 502.5)},
}


def _find_boxes(page):
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                if s["color"] == 0xFFFFFF:
                    continue
                for c in s["chars"]:
                    if c["c"] == "□":
                        out.append(fitz.Rect(c["bbox"]))
    return out


def _text_w(s, size):
    return sum(size * (0.5 if ord(ch) < 128 else 1.0) for ch in s)


def build_pdf(v, template_path, out_path):
    doc = fitz.open(template_path)
    fields = []     # (pno, rect, value, size, align)

    def tf(pno, x0, x1, base, value="", size=12, align=0):
        value = "" if value is None else str(value)
        if value and _text_w(value, size) > (x1 - x0) - 2:      # 太長自動縮字
            size = max(7, size * ((x1 - x0) - 2) / _text_w(value, size))
        fields.append((pno, fitz.Rect(x0, base - 11.5, x1, base + 3), value, size, align))

    def cell(pno, x0, x1, top, bot, value=""):
        base = (top + bot) / 2 + 12 * 0.36
        tf(pno, x0 + 1, x1 - 1, base, value, 12, 1)

    ap = v["申報人"]
    tf(0, 262, 566, B(122.4), ap.get("名稱"))
    tf(0, 634, 735, B(124.9), ap.get("統一編號"))
    tf(0, 634, 735, B(145.9), ap.get("電話"))
    tf(0, 262, 566, B(156.5), ap.get("地址"))
    tf(0, 634, 788, B(169.2), ap.get("電子信箱"))
    ag = v.get("代理人") or {}
    tf(0, 262, 566, B(192.4), ag.get("姓名"))
    tf(0, 634, 735, B(192.4), ag.get("統編"))
    tf(0, 634, 735, B(215.4), ag.get("電話"))
    tf(0, 634, 788, B(238.6), ag.get("電子信箱"))

    def addr_fields(addr, y1, y2):
        a = split_address(addr) if addr else {}
        tf(0, 252, 295, B(y1), short_county(a.get("縣市", "")))
        tf(0, 321, 350, B(y1), re.sub(r"區$", "", a.get("區", "")))
        tf(0, 399, 434, B(y1), re.sub(r"(路|街)$", "", a.get("路街", "")))
        tf(0, 471, 482, B(y1), a.get("段", "")); tf(0, 495, 506, B(y1), a.get("巷", ""))
        tf(0, 519, 548, B(y1), a.get("弄", ""))
        tf(0, 420, 458, B(y2), a.get("號", ""), align=2)
        tf(0, 471, 512, B(y2), a.get("樓", ""), align=1)
        tf(0, 524, 548, B(y2), a.get("室", ""))
    addr_fields(ag.get("地址", ""), 219.5, 239.4)

    tens = v.get("承租人") or []
    tf(0, 262, 566, B(262.8), "、".join(t["姓名"] for t in tens if t.get("姓名")))
    tf(0, 632, 788, B(261.3), "、".join(t["統編"] for t in tens if t.get("統編")), 11)
    tf(0, 632, 788, B(284.4), "、".join(t["電話"] for t in tens if t.get("電話")), 11)
    tf(0, 632, 788, B(307.8), v.get("承租人信箱"))
    addr_fields(tens[0].get("地址", "") if tens else "", 288.8, 308.8)

    ba = split_address(v.get("建物門牌", ""))
    tf(0, 182, 234, B(330.4), short_county(ba["縣市"]), align=1)
    tf(0, 265, 313, B(330.4), re.sub(r"區$", "", ba["區"]), align=1)
    tf(0, 373, 432, B(330.4), re.sub(r"(路|街)$", "", ba["路街"]), align=1)
    tf(0, 474, 498, B(330.4), ba["段"], align=1)
    tf(0, 523, 552, B(330.4), ba["巷"], align=1)
    tf(0, 571, 607, B(330.4), ba["弄"], align=1)
    tf(0, 625, 657, B(330.4), ba["號"], align=1)
    tf(0, 674, 710, B(330.4), ba["樓"], align=1)
    tf(0, 728, 760, B(330.4), ba["室"], align=1)

    tf(0, 224, 243, B(384.7), v.get("土地筆數"), align=1)
    tf(0, 340, 360, B(384.7), v.get("建物棟數"), align=1)
    tf(0, 494, 520, B(384.7), v.get("房間間數"), align=1)
    tf(0, 182, 415, B(410.8), v.get("總樓層"))
    tf(0, 522, 788, B(409.0), v.get("租賃層次"))
    tf(0, 182, 205, B(433.2), v.get("房"), align=1)
    tf(0, 220, 253, B(433.2), v.get("廳"), align=1)
    tf(0, 267, 295, B(433.2), v.get("衛"), align=1)
    tf(0, 222, 246, B(484.5), v.get("冷氣台數") if "冷氣" in v.get("設備", set()) else "", align=1)

    yb = B(519.9)
    s, e, sg = v.get("起") or {}, v.get("迄") or {}, v.get("訂約") or {}
    for x0, x1, val in [(188, 207, s.get("年")), (220, 235, s.get("月")), (246, 257, s.get("日")),
                        (282, 301, e.get("年")), (314, 329, e.get("月")), (340, 351, e.get("日")),
                        (530, 555, sg.get("年")), (570, 591, sg.get("月")), (602, 625, sg.get("日"))]:
        tf(0, x0, x1, yb, val, align=1)

    rent = fw2hw(v.get("租金", "")).replace(",", "")
    tf(1, 205, 284, B(76.0), f"{int(rent):,}" if rent.isdigit() else rent, 11, align=2)
    tf(1, 474, 490, B(74.2), v.get("車位個數") if v.get("車位") == "有" else "", align=1)
    tf(1, 680, 734, B(76.2), v.get("車位租金總額") if v.get("車位計價") == "單獨計價" else "", align=1)

    cols21 = [(56, 100), (100, 164.5), (164.5, 313), (313, 419), (419, 588)]
    lands = (v.get("土地") or []) + [{}] * 2
    for (top, bot), ld in zip([(192, 209), (209, 226)], lands[:2]):
        vals = [ld.get("縣市", ""), re.sub(r"區$", "", ld.get("區", "")), ld.get("段", ""), ld.get("地號", ""), ld.get("面積", "")]
        for (x0, x1), val in zip(cols21, vals):
            cell(1, x0, x1, top, bot, val)
    tf(1, 728, 788, B(195.6), lands[0].get("分區其他", ""), 10)
    tf(1, 728, 788, B(212.6), lands[1].get("分區其他", ""), 10)
    cols22 = [(56, 100), (100, 164.5), (164.5, 313), (313, 419), (419, 788)]
    blds = (v.get("建物") or []) + [{}] * 2
    for (top, bot), bd in zip([(262, 279), (279, 296)], blds[:2]):
        vals = [bd.get("縣市", ""), re.sub(r"區$", "", bd.get("區", "")), bd.get("段", ""), bd.get("建號", ""), bd.get("面積", "")]
        for (x0, x1), val in zip(cols22, vals):
            cell(1, x0, x1, top, bot, val)
    parks = (v.get("車位清冊") or []) + [{}] * 2
    for (top, bot), pk in zip([(331, 360), (360, 388)], parks[:2]):
        for (x0, x1), key in zip([(56, 100), (479, 588), (588, 702), (702, 788)], ["序號", "租金", "面積", "樓層"]):
            cell(1, x0, x1, top, bot, pk.get(key, ""))
    tf(1, 186, 204, B(417.0), str(len(tens)) if len(tens) > 1 else "", align=1)
    # 其餘「請敘明 / 其他」空白線：留空欄位，需要時直接在 PDF 點進去打
    tf(0, 663, 788, B(467.3), v.get("建物型態其他") if v.get("建物型態") == "其他" else "")   # 11 其他
    tf(1, 356, 478, B(347.6), parks[0].get("其他", "")); tf(1, 356, 478, B(376.0), parks[1].get("其他", ""))  # 23 ⑦其他
    rt = v.get("備註文字") or {}
    tf(1, 185, 253, B(490.5), rt.get("其他費用"), 11)        # 24② 其他(請敘明)
    tf(1, 406, 503, B(434.0), rt.get("全部門牌"), 11)        # 24③ 多個門牌(請敘明全部門牌)
    tf(1, 380, 484, B(506.0), rt.get("不同租金情形"), 11)    # 24③ 不同租金約定情形
    tf(1, 620, 719, B(502.5), rt.get("其他說明"), 11)        # 24④ 其他(請敘明)

    # ---- 勾選 ----
    on = {0: set(), 1: set()}
    on[0].add("經紀業" if ap.get("類別", "不動產經紀業") == "不動產經紀業" else "包租業")
    for k in (v.get("用途"), v.get("出租型態"), v.get("建物型態")):
        if k:
            on[0].add(k)
    for key, name in (("土地筆數", "棟數_土地"), ("建物棟數", "棟數_建物"), ("房間間數", "棟數_房間")):
        if v.get(key):
            on[0].add(name)
    if v.get("無隔間"):
        on[0].add("無隔間")
    eq = v.get("設備") or set()
    on[0].update(eq) if eq else on[0].add("無上述設備")
    on[1].add("有車位" if v.get("車位") == "有" else "無車位")
    if v.get("車位") == "有" and v.get("車位計價"):
        on[1].add("單獨計價" if v["車位計價"] == "單獨計價" else "未單獨計算")
    for i, pk in enumerate(parks[:2], 1):
        if pk.get("類別") in PARK_TYPES:
            on[1].add(f"車位{i}_{pk['類別']}")
    on[1].add("管理組織_有" if v.get("管理組織") else "管理組織_無")
    on[1].add("管理員_有" if v.get("管理員") else "管理員_無")
    on[1].add("電梯_有" if v.get("電梯") else "電梯_無")
    svc = v.get("租賃住宅服務") or ""
    if svc and v.get("用途") == "居住用":
        on[1].add("服務_" + svc)
    for i, ld in enumerate(lands[:2], 1):
        if ld.get("分區"):
            on[1].add(f"分區{i}_{ld['分區']}")
    if len(tens) > 1:
        on[1].add("多人承租")
    on[1].update(v.get("備註勾") or set())
    if v.get("含稅"):
        on[1].add("含稅")
    fees = v.get("衍生費用") or set()
    if fees:
        on[1].add("含衍生費用")
        on[1].update("費_" + f for f in fees)
    if (v.get("備註文字") or {}).get("其他費用"):
        on[1].update({"含衍生費用", "費_其他"})

    checks = []
    for pno, page in enumerate(doc):
        boxes = _find_boxes(page)
        for r in boxes:
            page.add_redact_annot(fitz.Rect(r.x0 + 1, r.y0 + 2, r.x1 - 1, r.y1 - 2))
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE,
                              text=fitz.PDF_REDACT_TEXT_REMOVE)
        names = {}
        for name, (x, y) in BOX[pno].items():
            hit = [r for r in boxes if abs(r.x0 - x) < 3 and abs(r.y0 - y) < 3]
            if not hit:
                raise RuntimeError(f"範本上找不到勾選框：{name}")
            names[id(hit[0])] = name
        for r in boxes:
            checks.append((pno, r, names.get(id(r)) in on[pno]))

    _add_form(doc, fields, checks)
    doc.save(out_path, garbage=3, deflate=True)
    return out_path


def _add_form(doc, fields, checks, box_inset=(1.4, 1.6, 1.4, 0.8), font_name="DFKaiShu-SB-Estd-BF",
              ap_font=None, bold=False):
    """欄位用不內嵌的標楷體(UCS2)，外觀自己畫 → Chrome/Edge/Acrobat 中文都正常，打字也正常
    ap_font=(頁面已 insert_font 的字型 xref, fitz.Font)：外觀改用這個「內嵌」字型(用 glyph id)，
    任何電腦都顯示該字型；編輯時仍用 font_name 指名的本機字型"""
    fd = doc.get_new_xref()
    doc.update_object(fd, f"<</Type/FontDescriptor/FontName/{font_name}/Flags 6"
                      "/FontBBox[0 -141 1000 859]/ItalicAngle 0/Ascent 859/Descent -141/CapHeight 859/StemV 80>>")
    cid = doc.get_new_xref()
    doc.update_object(cid, f"<</Type/Font/Subtype/CIDFontType2/BaseFont/{font_name}"
                      "/CIDSystemInfo<</Registry(Adobe)/Ordering(CNS1)/Supplement 0>>"
                      f"/FontDescriptor {fd} 0 R/DW 1000/W[1 95 500]>>")
    kai = doc.get_new_xref()
    doc.update_object(kai, f"<</Type/Font/Subtype/Type0/BaseFont/{font_name}"
                      f"/Encoding/UniCNS-UCS2-H/DescendantFonts[{cid} 0 R]>>")

    def xobj(w, h, content, res=""):
        x = doc.get_new_xref()
        doc.update_object(x, f"<</Type/XObject/Subtype/Form/BBox[0 0 {w:.2f} {h:.2f}]/Resources<<{res}>>>>")
        doc.update_stream(x, content.encode("latin-1"))
        return x

    for n, (pno, r, val, size, align) in enumerate(fields, 1):
        w = fitz.Widget()
        w.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        w.field_name = f"f{pno + 1}_{n:03d}"
        w.rect = r
        w.border_width = 0
        w.text_fontsize = size
        a = doc[pno].add_widget(w)
        wd, ht = r.width, r.height
        hexs = "".join(f"{ord(ch):04X}" for ch in val if ord(ch) <= 0xFFFF)
        if ap_font:
            fx, fobj = ap_font
            measure = lambda t: fobj.text_length(t, size)
            enc = lambda t: "".join(f"{fobj.has_glyph(ord(ch)):04X}" for ch in t)
            res, fres = f"/Font<</APF {fx} 0 R>>", "APF"
        else:
            measure = lambda t: _text_w(t, size)
            enc = lambda t: "".join(f"{ord(ch):04X}" for ch in t if ord(ch) <= 0xFFFF)
            res, fres = f"/Font<</Kai {kai} 0 R>>", "Kai"
        bold_ops = f"2 Tr {size * 0.04:.2f} w 0 0 1 RG " if bold else ""     # 描邊模擬粗體
        lines = val.split("\n") if val else []
        lead = size * 1.18
        ops = []
        for li, line in enumerate(lines):          # 多行：第一行從上方開始，最後一行落在基線 3
            tw = measure(line)
            tx = {0: 1.0, 1: (wd - tw) / 2, 2: wd - 1 - tw}[align]
            ty = 3 + lead * (len(lines) - 1 - li)
            ops.append(f"1 0 0 1 {tx:.2f} {ty:.2f} Tm <{enc(line)}> Tj")
        body = (f"/Tx BMC q BT /{fres} {size:.2f} Tf 0 0 1 rg {bold_ops}" + " ".join(ops) + " ET Q EMC"
                if lines else "/Tx BMC EMC")
        apx = xobj(wd, ht, body, res)
        if len(lines) > 1:
            doc.xref_set_key(a.xref, "Ff", "4096")    # 多行欄位
        doc.xref_set_key(a.xref, "AP", f"<</N {apx} 0 R>>")
        doc.xref_set_key(a.xref, "DA", f"(/Kai {size:.2f} Tf 0 0 1 rg)")
        doc.xref_set_key(a.xref, "Q", str(align))
        doc.xref_set_key(a.xref, "V", f"<FEFF{hexs}>" if val else "()")
        doc.xref_set_key(a.xref, "MK", "<<>>")

    for k, (pno, r, on) in enumerate(checks, 1):
        gr = fitz.Rect(r.x0 + box_inset[0], r.y0 + box_inset[1], r.x1 - box_inset[2], r.y1 - box_inset[3])
        w = fitz.Widget()
        w.field_type = fitz.PDF_WIDGET_TYPE_CHECKBOX
        w.field_name = f"c{pno + 1}_{k:03d}"
        w.rect = gr
        w.field_value = on
        a = doc[pno].add_widget(w)
        s = gr.width
        off = xobj(s, s, f"q 0 G 0.6 w 0.3 0.3 {s - 0.6:.2f} {s - 0.6:.2f} re S Q")
        yes = xobj(s, s, f"q 0 0 1 rg 0 0 {s:.2f} {s:.2f} re f Q")
        doc.xref_set_key(a.xref, "AP", f"<</N<</Yes {yes} 0 R/Off {off} 0 R>>/D<</Yes {yes} 0 R/Off {off} 0 R>>>>")
        doc.xref_set_key(a.xref, "AS", "/Yes" if on else "/Off")
        doc.xref_set_key(a.xref, "V", "/Yes" if on else "/Off")
        doc.xref_set_key(a.xref, "MK", "<</CA(n)/BC[0 0 0]>>")
        doc.xref_set_key(a.xref, "DA", "(/ZaDb 0 Tf 0 0 1 rg)")

    cat = doc.pdf_catalog()
    af = doc.xref_get_key(cat, "AcroForm")
    if af[0] == "xref":
        x = int(af[1].split()[0])
        doc.xref_set_key(x, "DR/Font/Kai", f"{kai} 0 R")
        doc.xref_set_key(x, "DA", "(/Kai 12 Tf 0 0 1 rg)")
        doc.xref_set_key(x, "NeedAppearances", "false")
    else:
        doc.xref_set_key(cat, "AcroForm/DR/Font/Kai", f"{kai} 0 R")
        doc.xref_set_key(cat, "AcroForm/DA", "(/Kai 12 Tf 0 0 1 rg)")
        doc.xref_set_key(cat, "AcroForm/NeedAppearances", "false")


def output_path(case_dir, v, prefix="實價登錄申報書(租)"):
    """存到租約所在資料夾(通常 6.成交\\日期)，沒有就 6.成交"""
    folder = os.path.dirname(v["租約檔"]) if v.get("租約檔") else os.path.join(case_dir, "6.成交")
    os.makedirs(folder, exist_ok=True)
    # 跟「物件個案調查表(租)_楠梓區芎蕉段-13_…」同一套命名
    lot = ""
    dj = _load_json(os.path.join(case_dir, "4.其他相關", "data.json"))
    if isinstance(dj, list) and dj and isinstance(dj[0], dict):
        d0 = dj[0]
        if d0.get("section") and d0.get("lot_number"):
            lot = f"{d0.get('area', '')}{d0['section']}-{d0['lot_number']}"
    ba = split_address(v.get("建物門牌", ""))
    floor = ""
    if ba["樓"]:
        fl, _, sub = ba["樓"].partition("之")
        floor = f"{fl}樓" + (f"之{sub}" if sub else "")
    addr = (ba["路街"] + (f"{ba['段']}段" if ba["段"] else "") + (f"{ba['巷']}巷" if ba["巷"] else "") +
            (f"{ba['弄']}弄" if ba["弄"] else "") + f"{ba['號']}號" + floor) if ba["路街"] and ba["號"] else ""
    tag = "_".join(x for x in (lot, addr) if x) or os.path.basename(case_dir)
    tag = re.sub(r'[\/:*?"<>|]', "", tag)
    base = os.path.join(folder, f"{prefix}_{tag}")
    path = base + ".pdf"
    i = 2
    while os.path.exists(path):         # 不覆蓋舊檔(可能有手動改過或正開著)
        path = f"{base}({i}).pdf"
        i += 1
    return path


# ============================================================
# 確認視窗
# ============================================================
def run_dialog(case_dir, v, notes, template_path):
    import tkinter as tk
    from tkinter import ttk, messagebox

    root = tk.Tk()
    root.title("實價登錄申報書(租賃) — 確認資料")
    import deal_ui as ui
    ui.apply(root)
    F, FB, FH = ui.F, ui.FB, ui.FH

    def big_dialog(title, msg, yesno=False):
        """大字訊息框(內建 messagebox 字太小)"""
        res = {"v": False}
        dlg = tk.Toplevel(root, bg=ui.CARD)
        dlg.title(title)
        dlg.transient(root)
        dlg.grab_set()
        tk.Frame(dlg, bg=ui.ACCENT if yesno else ui.PRIMARY, height=5).pack(fill=tk.X)
        tk.Label(dlg, text=msg, font=F, fg=ui.TEXT, bg=ui.CARD, justify="left", wraplength=560,
                 padx=28, pady=22).pack()
        bf = tk.Frame(dlg, bg=ui.CARD, pady=14)
        bf.pack()

        def done(v):
            res["v"] = v
            dlg.destroy()
        if yesno:
            ui.button(bf, "是", lambda: done(True)).pack(side=tk.LEFT, padx=8)
            ui.button(bf, "否", lambda: done(False), "ghost").pack(side=tk.LEFT, padx=8)
        else:
            ui.button(bf, "確定", lambda: done(True)).pack()
        dlg.update_idletasks()
        x = root.winfo_rootx() + (root.winfo_width() - dlg.winfo_width()) // 2
        y = root.winfo_rooty() + (root.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")
        root.wait_window(dlg)
        return res["v"]
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    W, H = min(1250, sw - 80), min(920, sh - 100)
    root.geometry(f"{W}x{H}+{(sw - W) // 2}+{max(10, (sh - H) // 2 - 30)}")

    ui.header(root, "實價登錄申報書（租賃）", os.path.basename(case_dir))
    foot = ui.footer(root)
    body, canvas = ui.scroll_body(root)

    warns = [m for lv, m in notes if lv == "warn"]
    infos = [m for lv, m in notes if lv == "info"]
    if warns:
        wc = ui.card(body, "請確認", "", "orange")
        for m in warns:
            tk.Label(wc, text="•  " + m, font=F, bg=wc.cget("bg"), fg=ui.WARN_FG, justify="left",
                     wraplength=W - 140).pack(anchor="w", pady=1)
    infos.insert(0, "讀的是物件編輯器「已儲存」的資料；剛改過還沒存的話，請先關掉這視窗、按「💾 儲存」再來")
    for m in infos:
        ui.note(body, m, "info", W - 120)

    # ---- 讀取的資料來源 ----
    srcf = ui.card(body, "📂 讀取的資料來源", "各欄位是從哪個檔案讀來的", "slate")
    for r_, (name, path, status, usage) in enumerate(v.get("來源") or []):
        col = ui.OK if status.startswith("✔") else (ui.BAD if status[:1] in "✖⚠" else ui.SUB)
        tk.Label(srcf, text=name, font=FB, fg="#1a1a1a", anchor="e").grid(row=r_ * 2, column=0, sticky="ne", padx=4)
        tk.Label(srcf, text=f"{status}　{path}", font=F, fg=col, anchor="w", justify="left",
                 wraplength=W - 260).grid(row=r_ * 2, column=1, sticky="w")
        tk.Label(srcf, text=f"用於：{usage}", font=ui.FS, fg=ui.SUB,
                 anchor="w").grid(row=r_ * 2 + 1, column=1, sticky="w", pady=(0, 4))
    lf2 = tk.Frame(srcf)
    lf2.grid(row=99, column=0, columnspan=2, sticky="w", pady=(4, 0))
    tk.Label(lf2, text="換一份租約：", font=F, fg="#1a1a1a").pack(side=tk.LEFT)
    cand_rel = [os.path.relpath(p_, case_dir) for p_ in v.get("租約候選") or []]
    cur_rel = os.path.relpath(v["租約檔"], case_dir) if v.get("租約檔") else ""
    lease_var = tk.StringVar(value=cur_rel)
    cb = ttk.Combobox(lf2, textvariable=lease_var, values=cand_rel, font=F, width=60, state="readonly")
    cb.pack(side=tk.LEFT, padx=4)

    def reload_with(path):
        if path and os.path.normcase(os.path.abspath(path)) != os.path.normcase(os.path.abspath(v.get("租約檔") or "")):
            result["reload"] = path
            root.destroy()
    cb.bind("<<ComboboxSelected>>", lambda e: reload_with(os.path.join(case_dir, lease_var.get())))

    def pick_file():
        from tkinter import filedialog
        p_ = filedialog.askopenfilename(parent=root, title="選擇租賃契約",
                                        initialdir=os.path.join(case_dir, "6.成交") if os.path.isdir(os.path.join(case_dir, "6.成交")) else case_dir,
                                        filetypes=[("合約(Word/PDF)", "*.docx *.pdf"), ("所有檔案", "*.*")])
        if p_:
            reload_with(p_)
    ui.button(lf2, "📁 選其他檔案…", pick_file, "chip", ui.THEMES["slate"][0]).pack(side=tk.LEFT, padx=6)
    tk.Label(srcf, text="(換租約會重新讀取，視窗裡改過的內容會還原成新租約的資料)", font=ui.FS,
             fg=ui.SUB).grid(row=100, column=0, columnspan=2, sticky="w")

    vars_ = {}

    themes = iter(["blue", "indigo", "teal", "green", "amber", "cyan", "purple", "brown", "slate", "rose"])

    def section(title):
        m = re.match(r"^(.*?)[(（](.*)[)）]$", title)
        t, hint = (m.group(1), m.group(2)) if m else (title, "")
        return ui.card(body, t, hint, next(themes, "blue"))

    def entry(parent, key, label, value, width=40, row=None, col=0):
        r = row if row is not None else parent.grid_size()[1]
        tk.Label(parent, text=label, font=F, fg="#1a1a1a", anchor="e").grid(row=r, column=col, sticky="e", padx=4, pady=2)
        var = tk.StringVar(value=value or "")
        tk.Entry(parent, textvariable=var, font=F, width=width, fg="#000000").grid(row=r, column=col + 1, sticky="w", padx=4, pady=2)
        vars_[key] = var
        return r

    def combo(parent, key, label, value, values, width=26, row=None, col=0):
        r = row if row is not None else parent.grid_size()[1]
        tk.Label(parent, text=label, font=F, fg="#1a1a1a", anchor="e").grid(row=r, column=col, sticky="e", padx=4, pady=2)
        var = tk.StringVar(value=value or "")
        ttk.Combobox(parent, textvariable=var, values=values, font=F, width=width).grid(row=r, column=col + 1, sticky="w", padx=4, pady=2)
        vars_[key] = var

    def check(parent, key, label, value, row, col):
        var = tk.BooleanVar(value=bool(value))
        tk.Checkbutton(parent, text=label, variable=var, font=F, fg="#1a1a1a").grid(row=row, column=col, sticky="w", padx=4)
        vars_[key] = var

    def date_row(parent, key, label, d):
        r = parent.grid_size()[1]
        tk.Label(parent, text=label, font=F, fg="#1a1a1a", anchor="e").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        d = d or {}
        for part in ("年", "月", "日"):
            vars_[f"{key}_{part}"] = tk.StringVar(value=d.get(part, ""))
        # 跟業績分配表同一套月曆；「迄」從「起」的月份開始翻
        start = (lambda: (vars_["起_年"].get(), vars_["起_月"].get())) if key == "迄" else None
        ui.date_field(parent, root, vars_[f"{key}_年"], vars_[f"{key}_月"], vars_[f"{key}_日"],
                      getattr(parent, "theme", (ui.PRIMARY,))[0], start).grid(row=r, column=1, sticky="w", pady=3)

    ap = v["申報人"]
    s1 = section("01 申報人(改了會記住，下次沿用)")
    combo(s1, "申報人_類別", "類別", ap.get("類別"), ["不動產經紀業", "租賃住宅包租業"])
    entry(s1, "申報人_名稱", "名稱", ap.get("名稱"), 50)
    entry(s1, "申報人_統一編號", "統一編號", ap.get("統一編號"), 20)
    entry(s1, "申報人_電話", "聯絡電話", ap.get("電話"), 20)
    entry(s1, "申報人_地址", "通訊地址", ap.get("地址"), 50)
    entry(s1, "申報人_電子信箱", "電子信箱", ap.get("電子信箱"), 50)

    ag = v.get("代理人") or {}
    s2 = section("02 申報代理人(受申報人委託才填；改了會記住)")
    entry(s2, "代理人_姓名", "姓名", ag.get("姓名"), 30)
    entry(s2, "代理人_統編", "統一編號", ag.get("統編"), 20)
    entry(s2, "代理人_電話", "聯絡電話", ag.get("電話"), 20)
    entry(s2, "代理人_地址", "通訊地址", ag.get("地址"), 50)
    entry(s2, "代理人_電子信箱", "電子信箱", ag.get("電子信箱"), 50)

    tens = v.get("承租人") or []
    s3 = section("03 承租人(多人時：姓名/統編/電話各自用「、」分隔；地址只填第一位)")
    entry(s3, "承租人_姓名", "姓名", "、".join(t.get("姓名", "") for t in tens), 50)
    entry(s3, "承租人_統編", "統一編號", "、".join(t.get("統編", "") for t in tens), 50)
    entry(s3, "承租人_電話", "聯絡電話", "、".join(t.get("電話", "") for t in tens), 50)
    entry(s3, "承租人_地址", "通訊地址", tens[0].get("地址", "") if tens else "", 50)
    entry(s3, "承租人信箱", "電子信箱", v.get("承租人信箱"), 50)

    s4 = section("04~14 建物與租約")
    entry(s4, "建物門牌", "04 建物門牌", v.get("建物門牌"), 50)
    combo(s4, "用途", "05 租賃用途", v.get("用途"), USES)
    combo(s4, "出租型態", "06 出租型態", v.get("出租型態"), LEASE_TYPES)
    r = s4.grid_size()[1]
    tk.Label(s4, text="07 租賃筆棟數", font=F, fg="#1a1a1a").grid(row=r, column=0, sticky="e", padx=4)
    cf = tk.Frame(s4)
    cf.grid(row=r, column=1, sticky="w")
    for key, lab in (("土地筆數", "土地(筆)"), ("建物棟數", "建物(棟/戶)"), ("房間間數", "房間(間)")):
        tk.Label(cf, text=lab, font=F, fg="#1a1a1a").pack(side=tk.LEFT, padx=(0, 2))
        var = tk.StringVar(value=v.get(key, ""))
        tk.Entry(cf, textvariable=var, font=F, width=4, fg="#000000").pack(side=tk.LEFT, padx=(0, 14))
        vars_[key] = var
    entry(s4, "總樓層", "08 總樓層數", v.get("總樓層"), 6)
    entry(s4, "租賃層次", "09 租賃層次", v.get("租賃層次"), 6)
    r = s4.grid_size()[1]
    tk.Label(s4, text="10 格局", font=F, fg="#1a1a1a").grid(row=r, column=0, sticky="e", padx=4)
    lf = tk.Frame(s4)
    lf.grid(row=r, column=1, sticky="w")
    for k in ("房", "廳", "衛"):
        var = tk.StringVar(value=v.get(k, ""))
        tk.Entry(lf, textvariable=var, font=F, width=3, fg="#000000").pack(side=tk.LEFT)
        tk.Label(lf, text=k, font=F, fg="#1a1a1a").pack(side=tk.LEFT, padx=(2, 10))
        vars_[k] = var
    var = tk.BooleanVar(value=v.get("無隔間"))
    tk.Checkbutton(lf, text="無隔間", variable=var, font=F).pack(side=tk.LEFT)
    vars_["無隔間"] = var
    combo(s4, "建物型態", "11 建物型態", v.get("建物型態"), BUILDING_TYPES, 30)
    entry(s4, "建物型態其他", "11 選「其他」時說明", v.get("建物型態其他"), 30)

    s12 = section("12 附屬設備")
    for i, name in enumerate(EQUIPMENT):
        check(s12, "設備_" + name, name, name in (v.get("設備") or set()), i // 5, i % 5)
    tk.Label(s12, text="冷氣台數", font=F).grid(row=2, column=0, sticky="e")
    var = tk.StringVar(value=v.get("冷氣台數", ""))
    tk.Entry(s12, textvariable=var, font=F, width=4).grid(row=2, column=1, sticky="w")
    vars_["冷氣台數"] = var

    s13 = section("13~14 期間與訂約日")
    date_row(s13, "起", "租賃期間 起", v.get("起"))
    date_row(s13, "迄", "租賃期間 迄", v.get("迄"))
    date_row(s13, "訂約", "14 訂約日期", v.get("訂約"))

    s15 = section("15~20")
    entry(s15, "租金", "15 租金(元/月)", v.get("租金"), 12)
    combo(s15, "車位", "16 車位", v.get("車位"), ["無", "有"], 6)
    entry(s15, "車位個數", "車位個數", v.get("車位個數"), 6)
    combo(s15, "車位計價", "車位計價方式", v.get("車位計價"), PARK_PRICING, 26)
    entry(s15, "車位租金總額", "車位租金總額(元/月)", v.get("車位租金總額"), 12)
    combo(s15, "租賃住宅服務", "20 租賃住宅服務", v.get("租賃住宅服務"), SERVICES, 20)
    r = s15.grid_size()[1]
    check(s15, "管理組織", "17 有管理組織", v.get("管理組織"), r, 1)
    r += 1
    check(s15, "管理員", "18 有管理員", v.get("管理員"), r, 1)
    r += 1
    check(s15, "電梯", "19 有電梯", v.get("電梯"), r, 1)

    s21 = section("21~22 土地/建物標的(謄本自動算：土地=面積×持分合計；建物=主建物+附屬+公設×持分)")
    hdr = ["", "縣市", "區", "段", "地號/建號", "面積(㎡)", "分區(住/商/工/農/其他)"]
    for c, h in enumerate(hdr):
        tk.Label(s21, text=h, font=FB, fg="#1a1a1a").grid(row=0, column=c, padx=3)
    rows = [("土地", i, (v.get("土地") or []) + [{}] * 2) for i in range(2)] + \
           [("建物", i, (v.get("建物") or []) + [{}] * 2) for i in range(2)]
    for rr, (kind, i, lst) in enumerate(rows, 1):
        d = lst[i]
        tk.Label(s21, text=f"{kind}{i + 1}", font=F, fg="#1a1a1a").grid(row=rr, column=0, sticky="e")
        keys = ["縣市", "區", "段", "地號" if kind == "土地" else "建號", "面積"]
        widths = [6, 6, 14, 10, 10]
        for c, (k, wdt) in enumerate(zip(keys, widths), 1):
            var = tk.StringVar(value=d.get(k, ""))
            tk.Entry(s21, textvariable=var, font=F, width=wdt, fg="#000000").grid(row=rr, column=c, padx=2, pady=1)
            vars_[f"{kind}{i}_{k}"] = var
        if kind == "土地":
            var = tk.StringVar(value=d.get("分區", ""))
            ttk.Combobox(s21, textvariable=var, values=ZONES, font=F, width=6).grid(row=rr, column=6, padx=2)
            vars_[f"{kind}{i}_分區"] = var
            var = tk.StringVar(value=d.get("分區其他", ""))
            tk.Entry(s21, textvariable=var, font=F, width=12, fg="#000000").grid(row=rr, column=7, padx=2)
            vars_[f"{kind}{i}_分區其他"] = var
    tk.Label(s21, text="分區「其他」說明", font=FB, fg="#1a1a1a").grid(row=0, column=7, padx=3)

    s23 = section("23 車位標的清冊(有租車位才填)")
    for c, h in enumerate(["", "序號", "車位類別", "⑦其他說明", "車位租金(元/月)", "面積(㎡)", "所在樓層"]):
        tk.Label(s23, text=h, font=FB, fg="#1a1a1a").grid(row=0, column=c, padx=3)
    parks = (v.get("車位清冊") or []) + [{}] * 2
    for i in range(2):
        pk = parks[i]
        tk.Label(s23, text=f"車位{i + 1}", font=F, fg="#1a1a1a").grid(row=i + 1, column=0, sticky="e")
        for c, (k, wdt) in enumerate((("序號", 4), ("類別", 0), ("其他", 12), ("租金", 10), ("面積", 8), ("樓層", 8)), 1):
            var = tk.StringVar(value=pk.get(k, ""))
            if k == "類別":
                ttk.Combobox(s23, textvariable=var, values=[""] + PARK_TYPES, font=F, width=11).grid(row=i + 1, column=c, padx=2)
            else:
                tk.Entry(s23, textvariable=var, font=F, width=wdt, fg="#000000").grid(row=i + 1, column=c, padx=2, pady=1)
            vars_[f"車位{i}_{k}"] = var

    s24 = section("24 備註(沒有的免勾；填了「請敘明」會自動打勾)")

    def sub(title, col):
        fr = tk.Frame(s24, bg="white", padx=12, pady=8, highlightthickness=1, highlightbackground=s24.theme[2])
        fr.grid(row=0, column=col, sticky="nsew", padx=4)
        tk.Label(fr, text=title, font=FB, fg=s24.theme[0], bg="white").pack(anchor="w", pady=(0, 4))
        return fr

    def rm_check(parent, key, label, value=False):
        var = tk.BooleanVar(value=bool(value))
        tk.Checkbutton(parent, text=label, variable=var, font=F, fg="#1a1a1a", anchor="w",
                       justify="left", wraplength=250).pack(anchor="w")
        vars_[key] = var

    def rm_text(parent, key, label):
        f = tk.Frame(parent)
        f.pack(anchor="w", padx=(24, 0), pady=(0, 4))
        tk.Label(f, text=label, font=F, fg="#444444").pack(side=tk.LEFT)
        var = tk.StringVar()
        tk.Entry(f, textvariable=var, font=F, width=14, fg="#000000").pack(side=tk.LEFT)
        vars_["文字_" + key] = var

    left = sub("①承租人 ②租金", 0)
    tk.Label(left, text="共同承租人數：依承租人姓名自動算", font=F, fg="#444444").pack(anchor="w")
    rm_check(left, "含稅", "含稅(限非居住用或混合用)")
    tk.Label(left, text="租金含下列費用：", font=F, fg="#1a1a1a").pack(anchor="w", pady=(4, 0))
    for fname in FEES:
        rm_check(left, "費_" + fname, "　" + fname, fname == "管理費" and v.get("含管理費"))
    rm_check(left, "費_其他", "　其他")
    rm_text(left, "費_其他", "請敘明")
    mid = sub("③租賃標的", 1)
    for key, label in REMARK_MID:
        rm_check(mid, key, label)
        if key in REMARK_TEXT:
            rm_text(mid, key, "請敘明全部門牌" if key == "多門牌" else "請敘明")
    right = sub("④特殊交易情況 / ④其他", 2)
    for key, label in REMARK_RIGHT:
        rm_check(right, key, label)
        if key in REMARK_TEXT:
            rm_text(right, key, "請敘明")
    for c in range(3):
        s24.grid_columnconfigure(c, weight=1)

    result = {"ok": False}

    def collect():
        g = lambda k: vars_[k].get().strip() if isinstance(vars_[k].get(), str) else vars_[k].get()
        out = dict(v)
        out["申報人"] = {"類別": g("申報人_類別"), "名稱": g("申報人_名稱"), "統一編號": g("申報人_統一編號"),
                      "電話": g("申報人_電話"), "地址": g("申報人_地址"), "電子信箱": g("申報人_電子信箱")}
        out["代理人"] = {k: g("代理人_" + k) for k in ("姓名", "統編", "電話", "地址", "電子信箱")}
        sp = lambda s: [x.strip() for x in re.split(r"[、,，]", s) if x.strip()]
        names, ids, tels = sp(g("承租人_姓名")), sp(g("承租人_統編")), sp(g("承租人_電話"))
        n = max(len(names), len(ids), len(tels))
        out["承租人"] = [{"姓名": names[i] if i < len(names) else "", "統編": ids[i] if i < len(ids) else "",
                       "電話": tels[i] if i < len(tels) else "", "地址": g("承租人_地址") if i == 0 else ""}
                      for i in range(n)]
        for k in ("建物門牌", "用途", "出租型態", "建物棟數", "總樓層", "租賃層次", "房", "廳", "衛",
                  "建物型態", "冷氣台數", "租金", "車位", "車位個數", "承租人信箱", "土地筆數", "房間間數",
                  "建物型態其他", "車位計價", "車位租金總額", "租賃住宅服務"):
            out[k] = g(k)
        for k in ("無隔間", "管理組織", "管理員", "電梯"):
            out[k] = bool(vars_[k].get())
        out["設備"] = {n for n in EQUIPMENT if vars_["設備_" + n].get()}
        for k in ("起", "迄", "訂約"):
            d = {p: g(f"{k}_{p}") for p in ("年", "月", "日")}
            out[k] = d if any(d.values()) else None
        lands, blds = [], []
        for i in range(2):
            ld = {k: g(f"土地{i}_{k}") for k in ("縣市", "區", "段", "地號", "面積", "分區")}
            ld["分區其他"] = g(f"土地{i}_分區其他") if ld["分區"] == "其他" else ""
            if any(ld[k] for k in ("段", "地號", "面積")):
                lands.append(ld)
            bd = {k: g(f"建物{i}_{k}") for k in ("縣市", "區", "段", "建號", "面積")}
            if any(bd[k] for k in ("段", "建號", "面積")):
                blds.append(bd)
        out["土地"], out["建物"] = lands, blds
        out["車位清冊"] = [{k: g(f"車位{i}_{k}") for k in ("序號", "類別", "其他", "租金", "面積", "樓層")} for i in range(2)]
        out["衍生費用"] = {f for f in FEES if vars_["費_" + f].get()}
        out["含稅"] = bool(vars_["含稅"].get())
        texts = {REMARK_TEXT[k]: g("文字_" + k) for k in REMARK_TEXT}
        out["備註文字"] = texts
        picked = {k for k, _ in REMARK_MID + REMARK_RIGHT if vars_[k].get()}
        picked |= {k for k in REMARK_TEXT if k != "費_其他" and texts[REMARK_TEXT[k]]}   # 有寫說明就自動勾
        out["備註勾"] = picked
        if vars_["費_其他"].get() or texts["其他費用"]:
            out["衍生費用"].add("其他")
        return out

    def on_ok():
        out = collect()
        missing = [k for k, ok in (("訂約日期", out.get("訂約")), ("租賃期間", out.get("起") and out.get("迄")),
                                   ("承租人", out.get("承租人"))) if not ok]
        if missing and not big_dialog("還有空白", "以下還沒填：" + "、".join(missing) +
                                               "\n\n要先產生 PDF(之後可在 PDF 直接點格子補)嗎？", yesno=True):
            return
        st = load_settings()
        st["申報人"] = out["申報人"]
        st["代理人"] = out["代理人"]
        save_settings(st)
        try:
            path = output_path(case_dir, out)
            build_pdf(out, template_path, path)
        except Exception as e:
            big_dialog("產生失敗", str(e))
            return
        result["ok"], result["path"] = True, path
        root.destroy()

    ui.repaint(body)
    ui.restyle_inputs(body)
    bf = foot
    tk.Label(bf, text="產生後會自動打開 PDF，每格都可再點進去改", font=ui.FS, fg=ui.SUB, bg=ui.CARD).pack(side=tk.LEFT)
    ui.button(bf, "📄  產生申報書 PDF", lambda: on_ok()).pack(side=tk.RIGHT)
    ui.button(bf, "取消", root.destroy, "ghost").pack(side=tk.RIGHT, padx=10)
    root.mainloop()
    return result


def gather_with_progress(case_dir, lease_path=None):
    """讀資料；遇到掃描檔要 OCR 時顯示進度小視窗(每頁約 10 秒)"""
    import threading
    import tkinter as tk
    out = {}
    root = tk.Tk()
    root.title("實價登錄申報書")
    root.attributes("-topmost", True)
    lab = tk.Label(root, text="讀取案件資料中…", font=("Microsoft JhengHei", 14), fg="#1a1a1a", padx=30, pady=24)
    lab.pack()
    root.update_idletasks()
    root.geometry(f"+{(root.winfo_screenwidth() - 420) // 2}+{root.winfo_screenheight() // 3}")
    msg = {"t": ""}

    def work():
        try:
            out["r"] = gather(case_dir, lease_path, progress=lambda t: msg.__setitem__("t", t))
        except Exception as e:
            out["e"] = e

    th = threading.Thread(target=work, daemon=True)
    th.start()

    def poll():
        if msg["t"]:
            lab.config(text=msg["t"] + "\n(掃描檔每頁約 10 秒，辨識完會記住，下次不用等)")
        if th.is_alive():
            root.after(300, poll)
        else:
            root.destroy()
    root.after(300, poll)
    root.mainloop()
    if "e" in out:
        raise out["e"]
    return out["r"]


def main():
    global _settings_dir
    ap = argparse.ArgumentParser()
    ap.add_argument("--case-dir", required=True)
    ap.add_argument("--base-dir", default="")
    ap.add_argument("--no-gui", action="store_true", help="測試用：直接用自動帶入值產生")
    ap.add_argument("--out", default="")
    a = ap.parse_args()
    _settings_dir = a.base_dir if a.base_dir and os.path.isdir(a.base_dir) else SCRIPT_DIR

    template = os.path.join(SCRIPT_DIR, TEMPLATE_NAME)
    if not os.path.exists(template) and a.base_dir:
        template = os.path.join(a.base_dir, TEMPLATE_NAME)
    if not os.path.exists(template):
        print(f"[錯誤] 找不到範本 {TEMPLATE_NAME}")
        return 1
    case_dir = os.path.abspath(a.case_dir)
    if a.no_gui:
        v, notes = gather(case_dir, progress=print)
        for lv, m in notes:
            print(f"[{'警告' if lv == 'warn' else '資訊'}] {m}")
        v["衍生費用"] = {"管理費"} if v.get("含管理費") else set()
        path = a.out or output_path(case_dir, v)
        build_pdf(v, template, path)
        print(f"[完成] {path}")
        return 0
    lease = None
    while True:
        v, notes = gather_with_progress(case_dir, lease)
        for lv, m in notes:
            print(f"[{'警告' if lv == 'warn' else '資訊'}] {m}")
        res = run_dialog(case_dir, v, notes, template)
        if res.get("reload"):
            lease = res["reload"]
            print(f"[資訊] 改用租約：{lease}")
            continue
        break
    if res.get("ok"):
        print(f"[完成] 已產生：{res['path']}")
        try:
            # 別讓 PDF 軟體繼承 _internal 工作目錄/Python 環境變數(會從 dist 載入 DLL 害打包時檔案被鎖)
            os.chdir(os.path.dirname(res["path"]))
            for k in ("PYTHONHOME", "PYTHONPATH", "PYTHONNOUSERSITE"):
                os.environ.pop(k, None)
            os.startfile(res["path"])
        except Exception:
            pass
    else:
        print("[取消] 未產生申報書")
    return 0


if __name__ == "__main__":
    sys.exit(main())
