"""行銷簡報產生器(大樓版)：讀案件資料夾 → 挑照片 → 產生 16:9 不動產行銷報告 PPTX

資料來源：4.其他相關/data_final.json(物件編輯器)、0.謄本 建物測量成果圖 PDF、2.照片(案件資料夾＋上次匯出的資料夾)、
3.行情 實價截圖、4.其他相關/gui_map_screenshot.png(位置圖)、通訊錄.xlsx(承辦人電話)。

照過去做行銷簡報踩過的雷內建：
- 圖片事先用 PIL 裁成框的比例再放(不用 PowerPoint 裁切 srcRect，iPad/iPhone 會顯示不出來)；從零建檔 → 關聯編號連號
- 配色明亮(天空藍＋金)，不用黑底；字大(長輩看得清楚)
- 照片自動轉正(EXIF)＋微提亮

用法：python marketing_deck.py --case-dir <案件資料夾> --base-dir <主程式資料夾>
"""
import os
import re
import math
import sys
import glob
import json
import argparse
import tempfile
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# 打包版 python_embedded 有 python39._pth → 不自動加腳本目錄，要自己加才 import 得到同資料夾模組
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from PIL import Image, ImageOps, ImageEnhance
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

COMPANY = "富住通不動產"
COMPANY_TEL = "07-382-7999"
FONT = "Microsoft JhengHei"

# 配色：天空藍＋金(高端、明亮，給長輩看不用黑底)
SKY = RGBColor(0x3E, 0x83, 0xB6)
SKY_DK = RGBColor(0x2B, 0x57, 0x79)
NUM = RGBColor(0x2C, 0x6C, 0x9E)
GOLD = RGBColor(0xB8, 0x8E, 0x3E)
GOLD_DK = RGBColor(0xA2, 0x7A, 0x2C)
INK = RGBColor(0x38, 0x46, 0x4F)
BG = RGBColor(0xF6, 0xFB, 0xFE)
PANEL = RGBColor(0xEA, 0xF5, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xC9, 0xDC, 0xEB)

# 版面用「設計座標」(13.333x7.5 吋)寫，輸出時依頁面大小換算(A4 橫式為預設，印 A4 滿版)
SW, SH = Inches(13.333), Inches(7.5)
PAGES = {"a4l": ("A4 橫式", 11.693, 8.268), "169": ("螢幕 16:9", 13.333, 7.5), "a4p": ("A4 直式", 8.268, 11.693)}
# 挑版面畫面用：(名稱, 說明, 寬, 高, 可用)
PAGE_CHOICES = [("a4l", "A4 橫式", "印成 A4 滿版（建議）", 297, 210, True),
                ("169", "螢幕 16:9", "電視、投影機播放", 16, 9, True),
                ("a4p", "A4 直式", "直式照片多、直式 DM", 210, 297, True)]
# 版型(依物件型式)：key → (圖示, 名稱, 適用, 包含頁面, 可用)
TEMPLATES = {
    "building": ("🏢", "大樓・華廈", "公寓、華廈、大樓、套房", "封面・基本資料・總覽・成果圖・室內・景觀・公設・行情", True),
    "villa": ("🏪", "透天・店面", "透天厝、透天店面、別墅", "滿版封面・物件亮點・樓層配置・周邊機能・街景（不用公司外框）", True),
    "office": ("🏬", "辦公・店面", "辦公室、商辦、店面", "樓層・面寬・管理・適合行業・商圈", False),
    "factory": ("🏭", "廠房・倉庫", "廠房、廠辦、倉庫", "坪數・挑高・電力・天車・卸貨・工廠登記", False),
    "land": ("🗺", "土地・建地", "建地、重劃區土地", "地籍圖・使用分區・面寬深度・臨路・開發試算", False),
    "farm": ("🌾", "農地", "農牧用地、田地", "地籍圖・農業用地・農舍資格・水源・現況", False),
}


def suggest_template(kind):
    """依物件編輯器的「型式」建議版型"""
    k = str(kind or "")
    for key, words in (("factory", ("廠", "倉庫")), ("farm", ("農", "田")), ("land", ("土地", "建地")),
                       ("villa", ("透天", "別墅", "墅")), ("office", ("辦公", "店面", "商辦", "店"))):
        if any(w in k for w in words):
            return key
    return "building"
SX = SY = 1.0                      # 設計座標 → 實際頁面 的縮放(由 set_page 設定)
REAL_W, REAL_H = SW, SH


def set_page(key):
    global SX, SY, REAL_W, REAL_H
    _, w, h = PAGES.get(key, PAGES["a4l"])
    REAL_W, REAL_H = Inches(w), Inches(h)
    SX, SY = REAL_W / SW, REAL_H / SH


# 色系：key → 名稱、說明、頁首漸層(左,右)、頁尾漸層(左,右)、主色、深色、數字色、頁底、卡片底、框線
# 富住通原色用公司外框圖(頁首藍→粉)；其它色系依同樣版面重畫漸層，LOGO/標語照貼
THEMES = {
    "fzt": dict(name="富住通原色", desc="藍粉漸層・公司公版", head=None, foot=None,
                sky="3E83B6", sky_dk="2B5779", num="2C6C9E", bg="F6FBFE", panel="EAF5FC", line="C9DCEB"),
    "navy": dict(name="經典深藍", desc="沉穩專業・公司套板色", head=("1E317E", "497DAD"), foot=("1E317E", "497DAD"),
                 sky="497DAD", sky_dk="1E317E", num="2A4A96", bg="F5F8FC", panel="E8EEF7", line="C5D2E6"),
    "warm": dict(name="香檳暖金", desc="明亮溫暖・豪宅長輩", head=("B8641E", "E2BC86"), foot=("8A5A2E", "C9A06A"),
                 sky="C07A35", sky_dk="6B4A2B", num="9A5B22", bg="FBF7F1", panel="F5ECDF", line="E5D3B8"),
    "green": dict(name="森林綠", desc="自然清新・土地農地", head=("2E6B47", "86B97C"), foot=("245C3B", "5E9A6C"),
                  sky="3E8E5E", sky_dk="245C3B", num="2F7A4E", bg="F5FAF6", panel="E6F2E9", line="C3DEC9"),
    "slate": dict(name="都會灰藍", desc="簡約俐落・廠辦店面", head=("37474F", "8FA4AF"), foot=("263238", "607D8B"),
                  sky="546E7A", sky_dk="263238", num="455A64", bg="F6F8F9", panel="ECEFF1", line="CFD8DC"),
}
THEMES["street"] = dict(name="炭灰琥珀", desc="店面街邊・透天專用(不用公司外框)", head=None, foot=None,
                        sky="B45309", sky_dk="1F2933", num="B45309", bg="FAF7F2", panel="F1ECE4", line="E2DACD")
THEME = "fzt"


def set_theme(key):
    """換色系：改掉全域配色(版面程式都讀這幾個常數)"""
    global THEME, SKY, SKY_DK, NUM, BG, PANEL, LINE
    THEME = key if key in THEMES else "fzt"
    t = THEMES[THEME]
    SKY, SKY_DK, NUM, BG, PANEL, LINE = (RGBColor.from_string(t[k]) for k in ("sky", "sky_dk", "num", "bg", "panel", "line"))


def _X(v):
    return int(v * SX)


def _Y(v):
    return int(v * SY)
IMG_EXT = (".jpg", ".jpeg", ".png", ".webp", ".bmp")

# 照片分類(挑照片視窗的下拉選單)；依檔名關鍵字猜
LAYOUT_CAT = "格局平面圖"      # 不是照片頁：每張自己一頁「格局平面圖」(左圖右重點)，放在照片頁前面
STREET_CAT = "周邊街景"
EXT_CAT = "標的外觀"          # 原「大樓外觀」：透天、店面也適用(舊設定讀進來自動換名)
SHOP_CAT = "店面空間"
CATS = ["封面", EXT_CAT, STREET_CAT, "騎樓", SHOP_CAT, "大廳公設",
        "客廳", "廳堂", "餐廳廚房", "臥室", "書房", "和室", "佛堂", "衛浴", "儲藏室", "樓梯間", "其他空間",
        "陽台景觀", "前後陽台", "頂樓露台", "庭院", "車位", "車庫",
        LAYOUT_CAT, "不使用"]
CAT_RENAME = {"大樓外觀": EXT_CAT}
CAT_SEP = "──────────"       # 下拉選單裡的分隔線(選到會自動退回)，把「格局平面圖」跟照片分類隔開
CAT_MENU = CATS[:-2] + [CAT_SEP, LAYOUT_CAT, CAT_SEP, "不使用"]
_GUESS = [
    (LAYOUT_CAT, ("平面圖", "格局", "配置圖", "floorplan", "floor_plan", "layout")),
    (STREET_CAT, ("街景", "周邊", "週邊", "街道", "路口", "巷口", "商圈", "street")),
    ("車庫", ("車庫", "garage")),
    ("車位", ("停車", "車位", "車道", "parking")),
    ("騎樓", ("騎樓",)),
    (SHOP_CAT, ("店面", "店內", "營業", "門市", "櫃台", "櫃檯", "shop", "store")),
    ("佛堂", ("佛堂", "神明", "神桌", "祖先", "公媽", "佛廳")),
    ("廳堂", ("廳堂", "正廳", "大廳堂")),
    ("書房", ("書房", "書桌", "study")),
    ("和室", ("和室", "榻榻米")),
    ("儲藏室", ("儲藏", "儲物", "置物", "收納")),
    ("樓梯間", ("樓梯", "梯間", "stair")),
    ("頂樓露台", ("頂樓", "屋頂", "露台", "rooftop")),
    ("庭院", ("庭院", "花園", "院子", "前院", "後院", "garden")),
    ("前後陽台", ("洗衣", "工作陽台", "前陽台", "後陽台", "生活陽台")),
    ("其他空間", ("更衣", "玄關", "走道", "夾層")),
    ("衛浴", ("浴", "衛", "廁", "bath")),
    ("餐廳廚房", ("廚", "餐", "kitchen")),
    ("客廳", ("客廳", "客餐", "起居", "living")),
    ("臥室", ("臥", "房間", "主臥", "bedroom")),
    ("陽台景觀", ("景觀", "視野", "夜景", "view")),
    ("前後陽台", ("陽台", "balcony")),             # 沒提到景觀的陽台 = 前/後陽台(不是每個陽台都有景觀)
    ("大廳公設", ("大廳", "公設", "健身", "ktv", "泳池", "中庭", "信箱", "交誼", "lobby", "gym")),
    (EXT_CAT, ("外觀", "大樓", "建築", "門面", "facade")),
]


def log(msg):
    print(msg, flush=True)


# ---------------------------------------------------------------- 資料
def _fw2hw(s):
    """全形數字/字母 → 半形"""
    return "".join(chr(ord(c) - 0xFEE0) if "０" <= c <= "９" or "Ａ" <= c <= "Ｚ" or "ａ" <= c <= "ｚ" else c for c in str(s or ""))


def _num(v, default=None):
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return default


def _fmt_money_wan(v):
    n = _num(v)
    if n is None:
        return str(v or "")
    return f"{n:,.0f}" if n == int(n) else f"{n:,.2f}".rstrip("0").rstrip(".")      # 6.25 萬不要變 6.2


_CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def _cn_int(s):
    """十二 → 12、二十三 → 23、7 → 7"""
    s = _fw2hw(s)
    m = re.search(r"\d+", s)
    if m:
        return int(m.group())
    s = re.sub(r"[^一二三四五六七八九十]", "", s)
    if not s:
        return None
    if "十" in s:
        a, _, b = s.partition("十")
        return (_CN.get(a, 1) if a else 1) * 10 + (_CN.get(b, 0) if b else 0)
    return _CN.get(s[0])


def find_export_dir(case_dir, base_dir):
    """物件編輯器「匯出」過的資料夾(照片常放那邊)：config.json export_settings 對得上這個案件才用"""
    for root in (base_dir, os.path.dirname(case_dir)):
        try:
            with open(os.path.join(root, "config.json"), "r", encoding="utf-8") as f:
                es = json.load(f).get("export_settings") or {}
        except Exception:
            continue
        cur = os.path.normcase(os.path.abspath(es.get("current_folder") or ""))
        d = os.path.join(es.get("last_export_path") or "", es.get("last_archive_name") or "")
        if es.get("last_archive_name") and cur == os.path.normcase(os.path.abspath(case_dir)) and os.path.isdir(d):
            return d
    # 對不上 → 用案件資料夾名稱(地段-地號)在上次匯出的上層找同名開頭的資料夾
    return None


def load_case(case_dir, base_dir):
    exp = find_export_dir(case_dir, base_dir)
    cands = [os.path.join(case_dir, "4.其他相關", "data_final.json")]
    if exp:
        cands.append(os.path.join(exp, "4.其他相關", "data_final.json"))
    cands = [p for p in cands if os.path.exists(p)]
    if not cands:
        raise SystemExit("找不到 data_final.json，請先在物件編輯器按「💾 儲存」")
    path = max(cands, key=os.path.getmtime)          # 兩邊都有 → 用較新的
    with open(path, "r", encoding="utf-8") as f:
        d = json.load(f)
    log(f"物件資料：{path}")
    try:
        import parking_info
        if parking_info.fill_missing(d, case_dir) or (exp and parking_info.fill_missing(d, exp)):
            pk = d.get("車位", {})
            log(f"車位：從謄本結構化補上 {pk.get('編號', '')}" + (f"，約 {pk.get('坪數')} 坪" if pk.get("坪數") else ""))
    except Exception as e:
        log(f"（補車位資訊略過：{e}）")
    return d, exp


def find_photos(case_dir, exp_dir):
    """兩個資料夾的 2.照片(含子資料夾)；「廣告」子資料夾(整理過的)排前面。同檔名只留一份"""
    found = {}
    for root in (exp_dir, case_dir):
        if not root:
            continue
        top = os.path.join(root, "2.照片")
        for r, _, fs in os.walk(top):
            for f in sorted(fs):
                if not f.lower().endswith(IMG_EXT):
                    continue
                ad = "廣告" in os.path.relpath(r, top)
                old = found.get(f.lower())
                if old and (old["ad"] or not ad):          # 同檔名：保留「廣告」那份(整理過的)
                    continue
                found[f.lower()] = {"path": os.path.join(r, f), "name": f, "ad": ad, "cat": guess_cat(f)}
    out = list(found.values())
    out.sort(key=lambda x: (not x["ad"], x["name"]))
    return out


SETTINGS_NAME = "行銷簡報設定.json"


def settings_path(case_dir):
    return os.path.join(case_dir, "4.其他相關", SETTINGS_NAME)


def load_settings(case_dir, photos):
    """上次挑的照片分類/說明/版面/版型 → 套回 photos；回傳上次的設定 dict(沒有存過回 {})"""
    try:
        with open(settings_path(case_dir), "r", encoding="utf-8") as f:
            st = json.load(f)
    except Exception:
        return {}
    saved = st.get("photos") or {}
    have = {p["name"].lower() for p in photos}
    for path in st.get("extra") or []:               # 上次在「整理與確認」新增的照片
        if os.path.exists(path) and os.path.basename(path).lower() not in have:
            photos.append({"path": path, "name": os.path.basename(path), "ad": False, "cat": guess_cat(path), "added": True})
            have.add(os.path.basename(path).lower())
    n = 0
    old = (st.get("version") or 1) < 2          # 舊設定還沒有「格局平面圖」分類 → 當時存的「不使用」不算數
    for p in photos:
        v = saved.get(p["name"].lower())
        if v:
            if not (old and v.get("cat") == "不使用" and p["cat"] == LAYOUT_CAT):
                p["cat"] = CAT_RENAME.get(v.get("cat"), v.get("cat", p["cat"]))
                if p["cat"] not in CATS:
                    p["cat"] = "其他空間"
            if "cap" in v:
                p["cap"] = v["cap"]
            for k in ("order", "prev"):
                if k in v:
                    p[k] = v[k]
            n += 1
    log(f"已載入上次的設定（{n} 張照片的分類/說明）")
    return st


def save_settings(case_dir, photos, page, template="building", show_agents=True, back_cover=False, theme="fzt", sections=None):
    """記住這次挑的照片分類/說明/版面(存在案件資料夾 4.其他相關)"""
    st = {"version": 2, "saved": datetime.now().strftime("%Y-%m-%d %H:%M"), "page": page, "template": template,
          "photos": {p["name"].lower(): dict({"cat": p["cat"]}, **({"cap": p["cap"]} if "cap" in p else {}),
                                             **({"order": p["order"]} if "order" in p else {}),
                                             **({"prev": p["prev"]} if "prev" in p else {}))
                     for p in photos},      # 說明只記「使用者改過的」，沒改過的下次依分類/檔名自動帶
          "extra": [p["path"] for p in photos if p.get("added")],
          "show_agents": show_agents, "back_cover": back_cover, "theme": theme,
          "sections": sections or {}}
    try:
        os.makedirs(os.path.dirname(settings_path(case_dir)), exist_ok=True)
        with open(settings_path(case_dir), "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=1)
    except Exception as e:
        log(f"（設定沒存成功：{e}）")


# 照片頁：(頁名, 包含的分類, 每頁最多幾張)；順序＝簡報裡的頁面順序
PHOTO_PAGES = [("室內精選｜客餐廳", ("客廳", "廳堂", "餐廳廚房"), 4),
               ("室內精選｜臥室・衛浴", ("臥室", "書房", "和室", "佛堂", "衛浴", "儲藏室", "樓梯間", "其他空間"), 4),
               ("景觀・陽台", ("陽台景觀", "前後陽台", "頂樓露台", "庭院"), 3),
               ("大樓外觀・公設", (EXT_CAT, "大廳公設", "車位", "車庫", "騎樓", SHOP_CAT), 4),
               ("周邊環境・街景", (STREET_CAT,), 4)]
# 透天・店面版的照片頁：店面先看外觀、街景
STREET_PHOTO_PAGES = [("外觀・周邊街景", (EXT_CAT, STREET_CAT, "騎樓", "車位", "車庫"), 4),
                      ("店面・營業空間", (SHOP_CAT,), 4),
                      ("室內空間", ("客廳", "廳堂", "餐廳廚房", "臥室", "書房", "和室", "佛堂", "衛浴", "儲藏室", "樓梯間",
                                    "其他空間"), 4),
                      ("景觀・陽台・頂樓", ("陽台景觀", "前後陽台", "頂樓露台", "庭院", "大廳公設"), 3)]


def split_pages(n, cap):
    """n 張照片、每頁最多 cap 張 → 每頁幾張(平均分配)。例 5 張每頁 4 → [3, 2]；9 張 → [3, 3, 3]"""
    if n <= 0:
        return []
    pages = (n + cap - 1) // cap
    base, extra = divmod(n, pages)
    return [base + (1 if i < extra else 0) for i in range(pages)]


def ordered(items):
    """依使用者在「整理與確認」排的順序(order)，沒排過的照原本順序"""
    return sorted(items, key=lambda p: (p.get("order", 10 ** 6), p["name"]))


def guess_cat(name):
    n = name.lower()
    if "成果圖" in n:                      # 建物測量成果圖另有專頁，不當照片
        return "不使用"
    for cat, keys in _GUESS:
        if any(k in n for k in keys):
            return cat
    return "不使用"


def caption_of(p):
    """照片說明：檔名有描述就用(去副檔名、去英文)；IMG_xxxx 就用分類名"""
    base = os.path.splitext(p["name"])[0]
    if re.match(r"^(img|dsc|dji|pxl|s__|line_|photo|image)[\s_\-]*\d", base, re.I) or re.fullmatch(r"[A-Za-z0-9_\s\-]+", base) \
            or re.search(r"\d{4}[-_./]\d{1,2}[-_./]\d{1,2}|[上下]午\s*\d|\d{1,2}[:：]\d{2}|^\s*(相片|照片|圖片|截圖)", base):
        # 手機自動檔名(相片 2026-10-15 下午3:05:51、IMG_1234…)不是說明 → 用分類名
        return p.get("cat") if p.get("cat") not in (None, "不使用", "封面") else ""
    return re.sub(r"\s*[A-Za-z][A-Za-z\s]*$", "", base).strip()


def lookup_phones(names, base_dir):
    """通訊錄.xlsx(第 11 列是標題：店別/職稱/姓名/電話/編號) 找承辦人電話"""
    paths = [os.path.join(base_dir, "_internal", "通訊錄.xlsx"), os.path.join(base_dir, "通訊錄.xlsx"),
             os.path.join(SCRIPT_DIR, "通訊錄.xlsx")]
    path = next((p for p in paths if os.path.exists(p)), None)
    out = {}
    if not path:
        return out
    try:
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        head = None
        for row in ws.iter_rows(values_only=True):
            cells = ["" if c is None else str(c).strip() for c in row]
            if head is None:
                if "姓名" in cells and "電話" in cells:
                    head = {k: cells.index(k) for k in ("姓名", "電話")}
                continue
            nm, ph = cells[head["姓名"]], cells[head["電話"]]
            if nm in names and ph and nm not in out:
                ph = ph.split(".")[0]
                if ph.isdigit() and not ph.startswith("0"):
                    ph = "0" + ph
                if len(ph) == 10 and ph.startswith("09"):
                    ph = f"{ph[:4]}-{ph[4:7]}-{ph[7:]}"
                out[nm] = ph
    except Exception as e:
        log(f"（讀通訊錄失敗：{e}）")
    return out


def facts(d):
    """把 data_final.json 整理成簡報要用的欄位"""
    door = _fw2hw(d.get("建物門牌", ""))
    door_short = re.sub(r"^(臺|台)?(北|中|南|高雄|新北|桃園|基隆|新竹|嘉義|苗栗|彰化|南投|雲林|屏東|宜蘭|花蓮|臺東|台東|澎湖|金門|連江)[市縣]", "", door)
    floor = None
    for k in (d.get("主建物") or {}):
        floor = _cn_int(k)
        if floor:
            break
    total_fl = _cn_int(d.get("地上層數", "")) if d.get("地上層數") else None
    age = ""
    m = re.search(r"民國\s*(\d+)\s*年", str(d.get("竣工日期", "")))
    if m:
        age = f"約 {datetime.now().year - 1911 - int(m.group(1))} 年"
    lay = d.get("格局") or {}
    layout = "".join(f"{lay.get(k)}{k}" for k in ("房", "廳", "衛") if str(lay.get(k, "")).strip())
    park = d.get("車位") or {}
    park_txt = ""
    if isinstance(park, dict) and "有" in str(park.get("有無", "")):
        park_txt = " ".join(x for x in (park.get("型式", ""), park.get("編號", "")) if x) or "有"
    points = [re.sub(r"^\s*[一二三四五六七八九十\d]+\s*[、.．]\s*", "", ln).strip()
              for ln in str(d.get("訴求重點", "")).splitlines()]
    points = [p for p in points if p]
    elev = ""
    if str(d.get("電梯", "")).startswith("有"):
        elev = f"{d.get('電梯數量')}部" if d.get("電梯數量") else "有"
        if d.get("一層幾戶"):
            elev += f"・一層 {d.get('一層幾戶')} 戶"
    attach = d.get("附屬建物") or {}
    attach_txt = "、".join(f"{k}{v}" for k, v in attach.items()) if isinstance(attach, dict) else ""
    out = _facts_dict(d, door, door_short, floor, total_fl, age, layout, park_txt, points, elev, attach_txt)
    # 透天/店面：各層面積、樓層範圍、面臨道路、適合行業
    fl = []
    for k, v in (d.get("主建物") or {}).items():
        n = _num(v)
        if n:
            fl.append((k, n))
    fl.sort(key=lambda x: _floor_order(x[0]))
    out["floors"] = fl
    nums = [_cn_int(k) for k, _ in fl if "地下" not in k and _cn_int(k)]
    if len(nums) >= 2 and total_fl:
        out["floor"] = f"{min(nums)}~{max(nums)} 樓 / 共 {total_fl} 樓"
    out["road_w"] = f"{_fmt_money_wan(d.get('面臨道路'))} 米" if _num(d.get("面臨道路")) else ""
    tr = d.get("適合行業別") or {}
    trades = [k for k, v in tr.items() if v is True and not k.startswith("其他")] if isinstance(tr, dict) else []
    if isinstance(tr, dict) and tr.get("其他_checked") and str(tr.get("其他_text", "")).strip():
        trades.append(str(tr["其他_text"]).strip())
    out["trades"] = trades
    out["above"] = total_fl
    out["below"] = _cn_int(d.get("地下層數", "")) if d.get("地下層數") else None
    out["use"] = d.get("用途", "")
    if not str(out["price"]).strip():            # 沒填價格 → 顯示「洽詢」(不要只剩一個「萬」字)
        out["price"], out["price_unit"] = "洽詢", ""
    return out


def _floor_order(name):
    """樓層排序：地下 → 一層、二層… → 夾層等 → 屋頂突出物"""
    n = _cn_int(name) or 0
    if "地下" in name:
        return (0, -n)
    if "騎樓" in name:                    # 騎樓在地面(一樓前面)：排在一層旁邊，不要跑到頂樓上面
        return (0.5, 0)
    if "屋頂" in name or "突出" in name:
        return (3, 0)
    return (1, n) if n else (2, 0)


def _zero(v):
    """「0」「0.00」「0.00 坪」這種沒有數值的欄位(透天沒有公設)→ 不顯示"""
    n = _num(re.sub(r"[^\d.,-]", "", str(v or "")) or None)
    return n is not None and n == 0


def _facts_dict(d, door, door_short, floor, total_fl, age, layout, park_txt, points, elev, attach_txt):
    return {
        "title": d.get("案名") or door_short,
        "short": (re.sub(r"[（(].*?[）)]", "", d.get("案名") or "").strip() or door_short)[:26],
        "door": door_short or door,
        "price": _fmt_money_wan(d.get("租金", "") if "租" in str(d.get("交易類型", "")) else d.get("總價", "")),
        "deal": "租" if "租" in str(d.get("交易類型", "")) else "售",
        "price_label": "月租" if "租" in str(d.get("交易類型", "")) else "總價",
        "price_unit": "萬/月" if "租" in str(d.get("交易類型", "")) else "萬",
        "dims": (f"{d.get('面寬')} 米 ／ {d.get('深度')} 米" if str(d.get("面寬") or "").strip() and str(d.get("深度") or "").strip()
                 else (f"面寬 {d.get('面寬')} 米" if str(d.get("面寬") or "").strip() else "")),
        "unit": d.get("建物單價", ""),
        "area": d.get("總建坪") or d.get("產權建坪", ""),
        "main": d.get("主建物坪數", ""), "attach": d.get("附屬建物坪數", ""), "attach_txt": attach_txt,
        "public": d.get("公共設施坪數", ""), "land": d.get("總地坪", ""),
        "floor": f"{floor} 樓" + (f" / 共 {total_fl} 樓" if total_fl else "") if floor else (f"共 {total_fl} 樓" if total_fl else ""),
        "age": age, "built": d.get("竣工日期", ""), "layout": layout, "park": park_txt,
        "park_ping": ((d.get("車位") or {}).get("坪數", "") if isinstance(d.get("車位"), dict) and park_txt else ""),
        "struct": d.get("建築結構", ""), "facing": d.get("座向", ""),
        "fee": f"{_fmt_money_wan(d.get('管理費'))} 元/月" if d.get("管理費") else "",
        "zone": d.get("都市土地使用分區", "") or d.get("非都市土地使用分區", ""),
        "elev": elev, "school": d.get("鄰近學校", ""), "district": d.get("所在商圈", ""),
        "points": points, "type": d.get("型式", ""), "agents": d.get("承辦人列表") or [],
        "lots": re.sub(r"^(臺|台)?..[市縣]", "", _fw2hw(d.get("土地標示", ""))).strip(),
        "bcr_far": (f"{d.get('建蔽率')}% ／ {d.get('容積率')}%" if d.get("建蔽率") and d.get("容積率") else ""),
    }


# ---------------------------------------------------------------- 圖片
IMG_DPI = 200                 # 照片放進簡報的解析度(依框大小換算像素；A4 印刷清楚、檔案不肥)
IMG_MAX = 2600                # 單張長邊上限(封面滿版)
IMG_MIN = 900                 # 小格也至少這麼大(放大看不糊)


LINE_DPI = 300                # 地籍圖/成果圖這類線條小字：較高解析度＋PNG 無損(不糊)
LINE_MAX = 5000               # ★線條圖不重新取樣:原圖長邊 ≤ 此值就保留原像素(縮一點點反而讓細線、小字變糊)
PDF_LINE_ZOOM = 6.0           # PDF 線條圖轉圖倍率(6 倍 ≈ 432 dpi;成果圖/地籍圖印 A4 約 490 dpi)


CAPTIONS = {}                 # 照片路徑 → 燒在照片邊角的文字標籤範圍(0～1)；裁切時別切到(photo_captions.py 用 OCR 找)


def load_captions(paths):
    """找出照片邊角的文字標籤(例：「附近街景圖」)。結果存快取(路徑＋修改時間＋大小)，同一張不重算；
    OCR 在 python_embedded 跑(主程式沒打包 OCR)，最多等 3 分鐘，失敗就照舊置中裁"""
    import subprocess
    frozen = getattr(sys, "frozen", False)
    exe_dir = os.path.dirname(sys.executable) if frozen else os.path.dirname(os.path.abspath(__file__))
    internal = getattr(sys, "_MEIPASS", exe_dir) if frozen else exe_dir
    cache_p = os.path.join(exe_dir, "photo_captions_cache.json")
    try:
        with open(cache_p, encoding="utf-8") as f:
            cache = json.load(f)
    except Exception:
        cache = {}

    def key(p):
        st = os.stat(p)
        return f"{os.path.abspath(p)}|{int(st.st_mtime)}|{st.st_size}"
    todo = []
    for p in dict.fromkeys(paths):
        try:
            k = key(p)
        except OSError:
            continue
        if k in cache:
            CAPTIONS[p] = cache[k]
        elif p.lower().endswith((".jpg", ".jpeg", ".png", ".webp", ".bmp")):
            todo.append(p)
    if not todo:
        return
    log(f"檢查 {len(todo)} 張照片邊角有沒有文字標籤（裁切時避開）…")
    script = os.path.join(internal, "photo_captions.py")
    if frozen:
        py_dir = os.path.join(exe_dir, "python_embedded")
        py = os.path.join(py_dir, "python.exe")
        env = dict(os.environ, PYTHONHOME=py_dir, PYTHONNOUSERSITE="1",
                   PYTHONPATH=internal + os.pathsep + os.path.join(py_dir, "Lib", "site-packages"))
    else:
        py, env = sys.executable, dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    tmpd = tempfile.mkdtemp(prefix="cap_")
    src, dst = os.path.join(tmpd, "in.json"), os.path.join(tmpd, "out.json")
    with open(src, "w", encoding="utf-8") as f:
        json.dump(todo, f, ensure_ascii=False)
    try:
        subprocess.run([py, script, src, dst], env=env, cwd=internal, capture_output=True, timeout=180,
                       creationflags=0x08000000)
    except Exception as e:
        log(f"（照片文字標籤檢查沒完成：{e}）")
    try:
        with open(dst, encoding="utf-8") as f:
            got = json.load(f)
    except Exception:
        got = {}
    n = 0
    for p, box in got.items():
        CAPTIONS[p] = box
        n += bool(box)
        try:
            cache[key(p)] = box
        except OSError:
            pass
    if n:
        log(f"  有 {n} 張照片邊角有文字標籤，裁切時會保留")
    try:
        with open(cache_p, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception:
        pass


def _keep_span(lo_default, size, full, a, b):
    """裁切窗(長 size)在 0～full 之間滑動：盡量包住 [a, b](標籤)，包不住就貼齊標籤那一側"""
    start = lo_default
    if b - a <= size:
        start = min(max(start, b - size), a)
    else:
        start = 0 if a < full - b else full - size
    return int(max(0, min(start, full - size)))


def prep_image(path, box_w, box_h, mode="cover", tmpdir=None, max_side=None, enhance=True, lossless=False):
    """事先把圖裁成框的比例(cover)或整張放進(contain)，縮小、轉正、微提亮 → 存成 JPEG/PNG 暫存檔。
    回傳 (暫存檔, 實際寬, 實際高)；contain 時寬高可能比框小"""
    im = Image.open(path)
    im = ImageOps.exif_transpose(im)
    has_alpha = im.mode in ("RGBA", "LA", "P")
    im = im.convert("RGB")
    ratio = box_w / box_h
    w, h = im.size
    if mode == "cover":
        cap = CAPTIONS.get(path)                          # 照片邊角燒進去的字(例：「附近街景圖」)：裁切窗挪過去包住它
        if w / h > ratio:
            nw = int(h * ratio)
            x0 = (w - nw) // 2
            if cap:
                x0 = _keep_span(x0, nw, w, cap[0] * w, cap[2] * w)
            im = im.crop((x0, 0, x0 + nw, h))
        else:
            nh = int(w / ratio)
            top = max(0, int((h - nh) * 0.45))           # 稍微偏上(天花板少、主體多)
            if cap:
                top = _keep_span(top, nh, h, cap[1] * h, cap[3] * h)
            im = im.crop((0, top, w, top + nh))
        out_w, out_h = box_w, box_h
    else:
        if w / h > ratio:
            out_w, out_h = box_w, int(box_w * h / w)
        else:
            out_w, out_h = int(box_h * w / h), box_h
    if max_side is None:          # 框寬(EMU)→吋→像素
        max_side = int(max(box_w, box_h) / 914400 * (LINE_DPI if lossless else IMG_DPI))
        max_side = max(IMG_MIN, min(3600 if lossless else IMG_MAX, max_side))
    if lossless:                  # 地籍圖/成果圖:不縮(除非大到離譜),保持原本每一個像素的清晰度
        max_side = max(max_side, LINE_MAX)
    im.thumbnail((max_side, max_side), Image.LANCZOS)
    if enhance and mode == "cover" and not lossless:
        im = ImageEnhance.Brightness(im).enhance(1.05)
        im = ImageEnhance.Color(im).enhance(1.08)
    fd, out = tempfile.mkstemp(suffix=".png" if (lossless or (has_alpha and mode != "cover")) else ".jpg", dir=tmpdir)
    os.close(fd)
    if out.endswith(".png"):
        im.save(out, "PNG", optimize=True)
    else:
        im.save(out, "JPEG", quality=85, optimize=True, progressive=True, subsampling=0)
    return out, out_w, out_h


def _trim(im, thr=235, pad=12):
    """去掉四周白邊；貼邊的整條框線也去掉(去完再去一次白邊)"""
    g = im.convert("L")
    for _ in range(2):
        box = g.point(lambda v: 255 if v < thr else 0).getbbox()
        if not box:
            break
        g, im = g.crop(box), im.crop(box)
        w, h = g.size
        px = g.load()

        def dark_row(y):
            return sum(1 for x in range(0, w, 3) if px[x, y] < thr) > (w / 3) * 0.6

        def dark_col(x):
            return sum(1 for y in range(0, h, 3) if px[x, y] < thr) > (h / 3) * 0.6
        t, b, l, r = 0, h, 0, w
        while t < h // 4 and dark_row(t):
            t += 1
        while b > h * 3 // 4 and dark_row(b - 1):
            b -= 1
        while l < w // 4 and dark_col(l):
            l += 1
        while r > w * 3 // 4 and dark_col(r - 1):
            r -= 1
        if (t, b, l, r) == (0, h, 0, w):
            break
        g, im = g.crop((l, t, r, b)), im.crop((l, t, r, b))
    out = Image.new("RGB", (im.width + pad * 2, im.height + pad * 2), "white")
    out.paste(im, (pad, pad))
    return out


def pdf_region_png(pdf, tmpdir, clip_fn, name, zoom=PDF_LINE_ZOOM):
    """PDF 第一頁：clip_fn(頁) 回傳要裁的範圍(None=整頁)，高倍率轉圖(向量線條小字清楚)、去白邊，存 PNG"""
    import fitz
    pg = fitz.open(pdf)[0]
    clip = None
    try:
        clip = clip_fn(pg)
    except Exception as e:
        log(f"（{name} 範圍判斷失敗，用整頁：{e}）")
    im = _pdf_native_scan(pg, clip)   # ★掃描型(成果圖):直接取 PDF 內原始掃描圖、在原像素上裁,不經任何重新取樣
    out = os.path.join(tmpdir, f"{name}_{os.getpid()}.png")
    if im is not None:
        if clip is None and clip_fn is plan_clip:   # 掃描圖上的字是圖片、找不到「位置圖」→ 用影像找中間圖框
            im = _scan_plan_crop(im)
        im.save(out, optimize=True)   # 掃描底色是淡灰,_trim 會誤判把圖切成一條 → 掃描型不用 _trim
        return out
    # 向量型(地籍圖):高倍率轉圖,線條小字越高倍越清楚
    pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    _trim(im).save(out, optimize=True)
    return out


def _scan_plan_crop(im):
    """掃描型建物測量成果圖:找左邊面積表右框(寬 25~55%)與右邊簽章欄左框(寬 75~93%)兩條長直線,
    以兩線的上下端為圖框上下緣 → 只留中間「位置圖＋平面圖」。找不到就整張不裁(寧可多,不可切掉)。"""
    try:
        import numpy as np
        a = np.asarray(im.convert("L"), dtype=np.int16)
        H, W = a.shape
        thr = min(170, int(np.percentile(a, 50)) - 45)   # 依底色自適應:淡掃描的框線只有灰色(固定 150 會抓不到)
        dark = a < thr
        soft = a < int(np.percentile(a, 50)) - 25
        col = dark.sum(0) / H

        def best_in(f0, f1):   # 該區最強的一條直線,且要明顯強過方格紙格線(≥ 周圍中位數 5 倍)
            seg = col[int(W * f0):int(W * f1)]
            x = int(W * f0) + int(seg.argmax())
            return x if col[x] > max(0.18, 5 * float(np.median(seg))) else None
        xl, xr = best_in(0.25, 0.55), best_in(0.75, 0.93)
        if xl is None or xr is None or xr - xl < W * 0.3:
            return im

        def extent(x):
            ys = np.where(soft[:, max(0, x - 3):x + 4].any(1))[0]   # 掃描線會微歪 → 左右各看 3px;沿線用寬鬆門檻(淡掃描線會斷)
            best, s, prev = (0, 0), ys[0], ys[0]
            for y in ys[1:]:
                if y - prev > 25:   # 印章/筆跡/掃描淡處造成的小斷點不算斷
                    if prev - s > best[1] - best[0]:
                        best = (s, prev)
                    s = y
                prev = y
            return best if prev - s <= best[1] - best[0] else (s, prev)
        (t1, b1), (t2, b2) = extent(xl), extent(xr)
        top, bot = min(t1, t2), max(b1, b2)
        # ★最後把關:成果圖中間圖框一定很高(約原圖 75~85%)、寬約 45~65%;不像就整張不裁(淡掃描易誤判,寧可多不可切掉)
        if bot - top < H * 0.7 or not (W * 0.4 <= xr - xl <= W * 0.7):
            log("（成果圖圖框判斷不確定，整張放入）")
            return im
        return im.crop((xl + 3, int(top) + 3, xr - 2, int(bot) - 2))
    except Exception as e:
        log(f"（成果圖範圍判斷失敗，用整張：{e}）")
        return im


def _pdf_native_scan(pg, clip):
    """PDF 頁面若是「一張大掃描圖(佔頁面三成以上)、沒有向量線條」(地政建物測量成果圖就是)，取出原始掃描像素並依 clip 裁切。
    清晰度上限就是原始掃描(約 140 dpi)；轉圖再放大不會多細節、只會多一次插補 → 這裡保留原像素。不是掃描型回 None。"""
    import fitz
    import math
    try:
        if pg.get_drawings():
            return None
        best = None
        for x in pg.get_images(full=True):
            for r, m in pg.get_image_rects(x[0], transform=True):
                if r.width * r.height >= pg.rect.width * pg.rect.height * 0.3 and (best is None or r.width * r.height > best[1].width * best[1].height):
                    best = (x[0], r, m)
        if not best:
            return None
        xref, r, mat = best
        pix = fitz.Pixmap(pg.parent, xref)
        if pix.n - pix.alpha >= 4:      # CMYK → RGB
            pix = fitz.Pixmap(fitz.csRGB, pix)
        im = Image.frombytes("RGB" if pix.n - pix.alpha == 3 else "L", (pix.width, pix.height), pix.samples).convert("RGB")
        # mat:圖片單位方塊(0~1) → 頁面(未旋轉)座標;把「顯示用」頁面座標換回去 → 圖片像素。翻轉/旋轉都一併處理
        to_unrot = pg.derotation_matrix
        inv = ~mat
        clip = clip or pg.rect

        def to_px(pt):
            q = fitz.Point(pt) * to_unrot * inv
            return q.x * im.width, q.y * im.height
        pts = [to_px(p) for p in (clip.tl, clip.tr, clip.bl, clip.br)]
        xs, ys = [p[0] for p in pts], [p[1] for p in pts]
        box = (max(0, int(min(xs))), max(0, int(min(ys))), min(im.width, int(max(xs))), min(im.height, int(max(ys))))
        if box[2] - box[0] > 50 and box[3] - box[1] > 50:
            im = im.crop(box)
        # 顯示方向:若圖片在頁面上是旋轉 90/270 度擺放,轉正
        ang = round(math.degrees(math.atan2(mat.b, mat.a))) % 360 if (mat.a or mat.b) else 0
        if ang in (90, 180, 270):
            im = im.rotate(-ang, expand=True)
        if mat.a < 0 and abs(mat.b) < 1e-6 or mat.d < 0 and abs(mat.c) < 1e-6:
            pass   # (掃描檔極少鏡像;遇到再處理)
        return im
    except Exception as e:
        log(f"（取原始掃描圖失敗，改用轉圖：{e}）")
        return None


def plan_clip(pg):
    """建物測量成果圖：只留中間圖框(位置圖＋平面圖)；左邊面積表、下方「一、二、…」說明、右邊 QR 與說明都不要"""
    import fitz
    head = pg.search_for("位置圖") or pg.search_for("平面圖")
    if not head:
        return None
    x0, y0 = head[0].x0 - 6, head[0].y0 - 4
    notes = [r for r in pg.search_for("一、") if r.y0 > y0 + 50]
    y1 = notes[0].y0 - 3 if notes else pg.rect.y1 * 0.75
    side = pg.search_for("本建物測量成果圖之比例尺") or pg.search_for("QR")
    x1 = side[0].x0 - 12 if side else pg.rect.x1 * 0.8
    return fitz.Rect(x0, y0, x1, y1)


def cadastral_clip(pg):
    """地籍圖謄本：只留中間地籍圖框；上面表頭、下面比例尺與說明不要"""
    import fitz
    H = pg.rect.y1
    # 表頭固定有「資料管轄機關／核發機關／主任」，緊接著「中華民國…」日期那行；以這幾行定位表頭下緣
    # (不能用「上方 3 成的文字」判斷：地圖上的路名常落在那裡，會被截掉)
    anchors = [r.y1 for t in ("資料管轄機關", "核發機關", "主任") for r in pg.search_for(t) if r.y1 < H * 0.35]
    if anchors:
        a_y = max(anchors)
        near = [b[3] for b in pg.get_text("blocks") if a_y - 2 <= b[3] <= a_y + 26]
        top = max(near + [a_y]) + 12
    else:
        top = H * 0.18
    scale = [r for r in pg.search_for("比例尺") if r.y0 > H * 0.5]
    bottom = scale[0].y0 - 3 if scale else H * 0.93
    return fitz.Rect(pg.rect.x0 + 20, top, pg.rect.x1 - 20, bottom)


def parcel_dims(pdf):
    """地籍圖量面寬/深度(建議值)：標的外框邊長(同方向線段合併) + 標註版的路名判斷臨路邊。
    回傳 dict(frontage, depth, road, roads=[(路名, 邊長)], edges=[...]) 或 None"""
    import fitz
    pg = fitz.open(pdf)[0]
    t = pg.get_text()
    m = re.search(r"比例尺[：:\s]*1\s*/\s*([\d,]+)", t)
    if not m:
        return None
    k = 0.0254 / 72 * int(m.group(1).replace(",", ""))          # 1 pt → 公尺
    best = None
    for dr in pg.get_drawings():
        col, fill = dr.get("color"), dr.get("fill")
        red = col and col[0] > 0.8 and col[1] < 0.3 and col[2] < 0.3
        gray = fill and abs(fill[0] - fill[1]) < 0.03 and abs(fill[1] - fill[2]) < 0.03 and 0.8 < fill[0] < 0.97
        if not (red or gray):
            continue
        segs = [((it[1].x, it[1].y), (it[2].x, it[2].y)) for it in dr["items"] if it[0] == "l"]
        if len(segs) < 3:
            continue
        area = dr["rect"].width * dr["rect"].height
        score = (1 if red else 0, area)                          # 紅框(標註版的標的)優先
        if best is None or score > best[0]:
            best = (score, segs, dr["items"])
    if not best:
        return None
    segs = best[1]
    # 同方向且相連的線段合併成一條邊
    edges = []
    for a, b in segs:
        ang = math.atan2(b[1] - a[1], b[0] - a[0])
        if edges:
            pa, pb, pang = edges[-1]
            d = abs((ang - pang + math.pi) % (2 * math.pi) - math.pi)
            if d < math.radians(8) and math.dist(pb, a) < 2:
                edges[-1] = (pa, b, math.atan2(b[1] - pa[1], b[0] - pa[0]))
                continue
        edges.append((a, b, ang))
    pts = [p for s in segs for p in s]

    def seg_dist(p, a, b):
        ax, ay, bx, by = *a, *b
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy or 1e-9
        u = max(0, min(1, ((p[0] - ax) * dx + (p[1] - ay) * dy) / L2))
        return math.dist(p, (ax + u * dx, ay + u * dy))

    def line_dist(p, a, b):
        dx, dy = b[0] - a[0], b[1] - a[1]
        return abs(dy * (p[0] - a[0]) - dx * (p[1] - a[1])) / (math.hypot(dx, dy) or 1e-9)
    # 路名(標註版)：文字以 路/街/大道 結尾(直書會被拆成一字一行 → 合併)
    roads = []
    import unicodedata as _ud
    texts = [(b[0], b[1], b[2], b[3], _ud.normalize("NFKC", b[4])) for b in pg.get_text("blocks")]
    # 直書路名：PDF 會一字一個 → 同一直行、上下相鄰的單字接起來
    import unicodedata
    words = [w[:4] + (unicodedata.normalize("NFKC", w[4]),) for w in pg.get_text("words")]   # 康熙部首「⼆」→「二」
    chars = sorted([w for w in words if len(w[4]) == 1 and "一" <= w[4] <= "鿿"],
                   key=lambda w: w[1])
    cols = []
    for w in chars:
        cx = (w[0] + w[2]) / 2
        for c in cols:
            last = c[-1]
            if abs((last[0] + last[2]) / 2 - cx) < 4 and 0 <= w[1] - last[3] < (last[3] - last[1]) * 1.6:
                c.append(w)
                break
        else:
            cols.append([w])
    for c in cols:
        if len(c) >= 2:
            texts.append((min(w[0] for w in c), c[0][1], max(w[2] for w in c), c[-1][3], "".join(w[4] for w in c)))
    for b in texts:
        name = re.sub(r"\s+", "", b[4])
        if re.fullmatch(r"[一-鿿\d一二三四五六七八九十]{1,8}(路|街|大道)(\d+段|[一二三四五六七八九十]段)?", name):
            c = ((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
            e = min(edges, key=lambda e: seg_dist(c, e[0], e[1]))
            if seg_dist(c, e[0], e[1]) < 120:                   # 太遠的(圖上其他路)不算
                roads.append((name, e))
    if roads:
        name, e = max(roads, key=lambda r: math.dist(r[1][0], r[1][1]))
    else:
        name, e = None, max(edges, key=lambda e: math.dist(e[0], e[1]))
    frontage = math.dist(e[0], e[1]) * k
    depth = max(line_dist(p, e[0], e[1]) for p in pts) * k
    return {"frontage": round(frontage, 1), "depth": round(depth, 1), "road": name,
            "roads": sorted({(n, round(math.dist(x[0], x[1]) * k, 1)) for n, x in roads}, key=lambda r: -r[1]),
            "edges": sorted([round(math.dist(a, b) * k, 1) for a, b, _ in edges], reverse=True)}


def plan_dims(pdf, floor_m2=None):
    """建物測量成果圖(向量檔)→ 平面圖外框 (短邊, 長邊) 公尺；掃描檔、量不出來、或面積跟登記差太多 → None。
    做法：平面圖標籤旁的比例尺 → 平面圖區內的線段依端點相連分群(相碰的再併)→ 最大那群的最小外接矩形。
    floor_m2(該層登記面積 ㎡)有給就驗證：量出來的面積要在 0.7～1.4 倍之間才算數(寧可不給，不給錯的)"""
    import fitz
    from shapely.geometry import MultiPoint
    doc = fitz.open(pdf)
    pg = doc[0]
    words = pg.get_text("words")
    labs = [w for w in words if w[4].strip() in ("平面圖", "平面圖比例尺")]
    if not labs:
        return None
    lab = min(labs, key=lambda w: w[1])
    scs = [(abs(w[0] - lab[2]), w) for w in words
           if re.fullmatch(r"1[/:]\d{2,4}", w[4]) and abs(w[1] - lab[1]) < 6 and w[0] > lab[0]]
    if not scs:
        return None
    scale = int(re.split(r"[/:]", min(scs)[1][4])[1])
    pos = next((w for w in words if "位置圖" in w[4] and abs(w[1] - lab[1]) < 6), None)
    x0 = (pos[2] + lab[0]) / 2 + 40 if pos and pos[0] < lab[0] else pg.rect.x0
    note = next((w for w in words if "本建物係" in w[4] or w[4].startswith("二、") or w[4].startswith("一、本建物")), None)
    region = fitz.Rect(x0, lab[3] + 2, pg.rect.x1 - 150, note[1] - 2 if note else pg.rect.y1)
    segs = []
    for dr in pg.get_drawings():
        for it in dr["items"]:
            if it[0] == "l" and region.contains(it[1]) and region.contains(it[2]) and it[1].distance_to(it[2]) > 0.5:
                segs.append(((it[1].x, it[1].y), (it[2].x, it[2].y)))
    if len(segs) < 4:
        return None
    par = {}

    def find(x):
        while par.setdefault(x, x) != x:
            par[x] = par[par[x]]
            x = par[x]
        return x

    def key(p):
        return round(p[0] * 2), round(p[1] * 2)
    for a, b in segs:
        par[find(key(a))] = find(key(b))
    groups = {}
    for a, b in segs:
        groups.setdefault(find(key(a)), []).extend([a, b])
    boxes = []
    for g in groups.values():
        xs, ys = [q[0] for q in g], [q[1] for q in g]
        boxes.append([min(xs) - 3, min(ys) - 3, max(xs) + 3, max(ys) + 3, list(g)])
    merged = True
    while merged:
        merged = False
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                A, B = boxes[i], boxes[j]
                if A[0] <= B[2] and B[0] <= A[2] and A[1] <= B[3] and B[1] <= A[3]:
                    A[:4] = [min(A[0], B[0]), min(A[1], B[1]), max(A[2], B[2]), max(A[3], B[3])]
                    A[4] += B[4]
                    del boxes[j]
                    merged = True
                    break
            if merged:
                break
    pts = max((b[4] for b in boxes), key=lambda g: MultiPoint(g).convex_hull.area)
    if len(pts) < 3:
        return None
    c = list(MultiPoint(pts).minimum_rotated_rectangle.exterior.coords)
    k = 25.4 / 72 * scale / 1000
    a, b = sorted((math.dist(c[0], c[1]) * k, math.dist(c[1], c[2]) * k))
    if a < 2 or b > 80:
        return None
    if floor_m2 and not (0.7 <= a * b / floor_m2 <= 1.4):
        return None
    return round(a, 2), round(b, 2)


def cadastral_scale(pdf):
    """地籍圖比例尺文字(例：1/500)"""
    try:
        import fitz
        t = fitz.open(pdf)[0].get_text()
        m = re.search(r"比例尺[：:\s]*(1\s*/\s*[\d,]+)", t)
        return m.group(1).replace(" ", "") if m else ""
    except Exception:
        return ""


def pdf_to_png(pdf, tmpdir, zoom=2.5):
    import fitz
    doc = fitz.open(pdf)
    outs = []
    for i, pg in enumerate(doc):
        pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        p = os.path.join(tmpdir, f"pdf_{os.getpid()}_{i}.png")
        pix.save(p)
        outs.append(p)
        if i >= 1:
            break
    return outs


# ---------------------------------------------------------------- 版面元件
# 富住通公版素材(從「富住通不動產行銷報告公版」取出)：內容頁外框(漸層頁首＋標誌＋Full Life 頁尾，中間透明)、白字標誌
ASSET_FRAME = "fzt_frame.png"
ASSET_LOGO = "fzt_logo_hd.png"          # 公司正式白字 LOGO(2659px，印 A4 也清楚)
ASSET_LOGO_V = "fzt_logo_vertical.png"   # 直式 LOGO(封底)
ASSET_FULLLIFE = "fzt_fulllife_blue.png" # Full Life 草寫字(封底)
STAMP_RED = RGBColor(0xC4, 0x27, 0x22)   # 公司套板的售/租章紅
HEADER_H = Inches(1.0)       # 公版頁首高
FOOTER_Y = Inches(7.23)      # 公版頁尾起點
GOLD_HI = RGBColor(0xE6, 0xBE, 0x55)   # 金色漸層(亮)
GOLD_LO = RGBColor(0xB8, 0x88, 0x2B)   # 金色漸層(深)


def asset(name):
    for d in (SCRIPT_DIR, os.path.join(SCRIPT_DIR, "_internal")):
        p = os.path.join(d, name)
        if os.path.exists(p):
            return p
    return None


def _font(run, size, bold=False, color=INK):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", FONT)


def text(slide, x, y, w, h, s, size=16, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         spacing=1.15):
    x, y, w, h = _X(x), _Y(y), _X(w), _Y(h)
    if "\n" not in str(s) and h < Pt(size) * 2.7:      # 單行字：放不下就縮小字級(A4 比螢幕窄)
        while size > 10 and _text_w(s, size) > w - Inches(0.12):
            size -= 1
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    for i, line in enumerate(str(s).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
        _font(r, size, bold, color)
    return tb


def box(slide, x, y, w, h, fill=WHITE, line=None, radius=True, alpha=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE, _X(x), _Y(y), _X(w), _Y(h))
    if radius:
        shp.adjustments[0] = 0.08
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if alpha is not None:                 # 半透明(0~1，1=不透明)
        clr = shp.fill._xPr.find(qn("a:solidFill"))[0]
        a = clr.makeelement(qn("a:alpha"), {"val": str(int(alpha * 100000))})
        clr.append(a)
    if line:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def gold(slide, x, y, w, h, radius=True, oval=False):
    """金色漸層牌(亮金→深金，比單色土黃好看)"""
    if oval:
        shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, _X(x), _Y(y), _X(w), _Y(h))
        shp.line.fill.background()
        shp.shadow.inherit = False
    else:
        shp = box(slide, x, y, w, h, GOLD, radius=radius)
    shp.fill.gradient()
    shp.fill.gradient_angle = 90
    st = shp.fill.gradient_stops
    st[0].color.rgb, st[0].position = GOLD_HI, 0.0
    st[1].color.rgb, st[1].position = GOLD_LO, 1.0
    return shp


def tone(img_path, frac=(0.0, 0.6, 0.6, 1.0), dark=0.45):
    """照片某區塊(左,上,右,下 比例)的平均色，壓暗一點 → 當半透明底色(同色系、白字看得清楚)"""
    try:
        im = Image.open(img_path).convert("RGB")
        w, h = im.size
        im = im.crop((int(w * frac[0]), int(h * frac[1]), int(w * frac[2]), int(h * frac[3]))).resize((32, 32))
        px = list(im.getdata())
        r, g, b = (sum(c[i] for c in px) / len(px) for i in range(3))
    except Exception:
        r, g, b = 0x3E, 0x83, 0xB6
    k = 1 - dark
    return RGBColor(int(r * k), int(g * k), int(b * k))


def picture(slide, path, x, y, w, h, tmpdir, mode="cover", lossless=False):
    """放一張圖到框裡(事先裁好、不用 srcRect)。contain 時置中。回傳(圖形, 處理後的圖檔)"""
    x, y, w, h = _X(x), _Y(y), _X(w), _Y(h)
    img, iw, ih = prep_image(path, int(w), int(h), mode, tmpdir, lossless=lossless)
    if mode == "contain":
        x, y = x + (w - iw) // 2, y + (h - ih) // 2
        w, h = iw, ih
    return slide.shapes.add_picture(img, x, y, w, h), img


def make_frame(tmpdir):
    """富住通外框依目前頁面大小重組：頁首/頁尾漸層拉寬，標誌、Full Life 標語保持原比例貼回(不壓扁)"""
    src = asset(ASSET_FRAME)
    if not src:
        return None
    a = Image.open(src).convert("RGBA")              # 1920x1080：頁首 0~144、頁尾 1041~1080
    k = 144.0                                        # 每設計吋的像素
    W, H = int(REAL_W / 914400 * k), int(REAL_H / 914400 * k)
    hh, fh = int(round(144 * SY)), int(round(39 * SY))
    out = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    t = THEMES.get(THEME, THEMES["fzt"])

    def strip(y0, y1, sample_rows):
        row = a.crop((0, sample_rows[0], a.width, sample_rows[1])).resize((a.width, 1), Image.BOX)
        return row.resize((W, y1 - y0), Image.BILINEAR)

    def grad(h, c1, c2):
        """左→右漸層(其它色系用)"""
        c1, c2 = tuple(int(c1[i:i + 2], 16) for i in (0, 2, 4)), tuple(int(c2[i:i + 2], 16) for i in (0, 2, 4))
        row = Image.new("RGBA", (256, 1))
        for x in range(256):
            k = x / 255
            row.putpixel((x, 0), tuple(int(c1[i] + (c2[i] - c1[i]) * k) for i in range(3)) + (255,))
        return row.resize((W, h), Image.BILINEAR)
    if t["head"]:
        out.paste(grad(hh, *t["head"]), (0, 0))
        out.paste(grad(fh, *t["foot"]), (0, H - fh))
        lg = asset(ASSET_LOGO)                       # 白字透明 LOGO(公版外框裡的 LOGO 連著藍粉底，不能用)
        if lg:
            logo = Image.open(lg).convert("RGBA")
            lh = int(round(96 * SY))
            logo = logo.resize((int(logo.width * lh / logo.height), lh), Image.LANCZOS)
            out.alpha_composite(logo, (W - int(41 * SY) - logo.width, int(24 * SY)))
        # 頁尾標語「關懷・家庭・文化・國際 Full Life」：從公版取白字(亮度當透明度)貼到新漸層上
        sl = a.crop((770, 1041, 1150, 1080)).convert("L")
        sl = sl.resize((int(sl.width * SY), fh), Image.LANCZOS)
        mask = sl.point(lambda v: max(0, min(255, (v - 150) * 3)))
        white = Image.new("RGBA", sl.size, (255, 255, 255, 255))
        white.putalpha(mask)
        out.alpha_composite(white, ((W - sl.width) // 2, H - fh))
    else:
        out.paste(strip(0, hh, (0, 28)), (0, 0))
        out.paste(strip(H - fh, H, (1041, 1043)), (0, H - fh))
        logo = a.crop((1376, 22, 1880, 124))
        logo = logo.resize((int(logo.width * SY), int(logo.height * SY)), Image.LANCZOS)
        out.alpha_composite(logo, (W - int(51 * SY) - logo.width + int(10 * SY), int(22 * SY)))
        slog = a.crop((770, 1041, 1150, 1080))
        slog = slog.resize((int(slog.width * SY), fh), Image.LANCZOS)
        out.alpha_composite(slog, ((W - slog.width) // 2, H - fh))
    p = os.path.join(tmpdir, "frame.png")
    out.save(p)
    return p


def _text_w(s, size):
    """估文字寬度(中文≈1 字寬、英數≈0.55)"""
    em = sum(1.0 if ord(c) > 0x2E80 else 0.55 for c in str(s))
    return int(Pt(size) * em)


class Deck:
    def __init__(self, f, tmpdir):
        self.f, self.tmp = f, tmpdir
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = REAL_W, REAL_H
        self.blank = self.prs.slide_layouts[6]
        self.n = 0
        self.frame = make_frame(tmpdir)

    def new(self, title=None):
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        bg = s.background.fill
        bg.solid()
        bg.fore_color.rgb = BG
        if title:
            if self.frame:              # 富住通公版外框(頁首漸層＋標誌、頁尾 Full Life)
                s.shapes.add_picture(self.frame, 0, 0, REAL_W, REAL_H)
            else:
                box(s, 0, 0, SW, HEADER_H, SKY, radius=False)
            text(s, Inches(0.45), Inches(0.1), Inches(8.6), Inches(0.8), title, 30, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
            text(s, Inches(11.9), FOOTER_Y, Inches(1.2), SH - FOOTER_Y, str(self.n), 11, True, WHITE, PP_ALIGN.RIGHT,
                 MSO_ANCHOR.MIDDLE)
        return s

    # 照片格：1 張大圖；2 張左右；3 張左大右二；4 張田字
    def photos(self, s, items, x, y, w, h, gap=Inches(0.18)):
        items = items[:4]
        n = len(items)
        cells = []
        if n == 1:
            cells = [(x, y, w, h)]
        elif n == 2:
            cw = (w - gap) // 2
            cells = [(x, y, cw, h), (x + cw + gap, y, cw, h)]
        elif n == 3:
            lw = int(w * 0.6)
            rw = w - lw - gap
            rh = (h - gap) // 2
            cells = [(x, y, lw, h), (x + lw + gap, y, rw, rh), (x + lw + gap, y + rh + gap, rw, rh)]
        elif n == 4:
            cw, ch = (w - gap) // 2, (h - gap) // 2
            cells = [(x, y, cw, ch), (x + cw + gap, y, cw, ch), (x, y + ch + gap, cw, ch), (x + cw + gap, y + ch + gap, cw, ch)]
        caps = [p["cap"] if "cap" in p else caption_of(p) for p in items]     # 挑照片視窗改過的說明優先
        # 說明條寬度：以這頁最長的說明為準(全頁一致)，不再整條橫跨蓋住照片下緣
        cap_w = int(max([_text_w(c, 14) for c in caps if c] or [0]) / SX) + Inches(0.4)
        for p, cap, (cx, cy, cw, ch) in zip(items, caps, cells):
            _, img = picture(s, p["path"], cx, cy, cw, ch, self.tmp)
            if cap:
                bh = Inches(0.42)
                bw = min(cap_w, cw - Inches(0.3))
                bx, by = cx + Inches(0.15), cy + ch - bh - Inches(0.15)
                box(s, bx, by, bw, bh, tone(img, (0.0, 0.75, 0.5, 1.0)), alpha=0.68)
                text(s, bx + Inches(0.12), by, bw - Inches(0.2), bh, cap, 14, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------- 各頁
def slide_cover(deck, cover, phones, show_agents=True):
    f = deck.f
    s = deck.new()
    img = None
    if cover:
        _, img = picture(s, cover["path"], 0, 0, SW, SH, deck.tmp)
    # 左下資訊卡：跟照片同色系的半透明底，透得到背景照片
    cx, cy, cw, ch = Inches(0.6), Inches(3.55), Inches(7.8), Inches(3.3)
    card = tone(img, (0.0, 0.45, 0.6, 0.95)) if img else SKY_DK
    box(s, cx, cy, cw, ch, card, alpha=0.62)
    gold(s, cx, cy, Inches(0.14), ch, radius=False)
    text(s, cx + Inches(0.4), cy + Inches(0.25), cw - Inches(0.6), Inches(1.3), f["title"], 28, True, WHITE)
    text(s, cx + Inches(0.4), cy + Inches(1.55), cw - Inches(0.6), Inches(0.5), f["door"], 18, True, WHITE)
    gold(s, cx + Inches(0.4), cy + Inches(2.25), Inches(3.8), Inches(0.78))
    text(s, cx + Inches(0.4), cy + Inches(2.25), Inches(3.8), Inches(0.78),
         f"{'月租' if f['deal'] == '租' else '開價'} {f['price']} {f['price_unit']}", 28, True, WHITE,
         PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    # 左上：紅色售／租章(公司套板規格：紅底圓角方塊＋白色粗體字)
    st = box(s, Inches(0.6), Inches(0.45), Inches(1.25), Inches(1.25), STAMP_RED)
    st.adjustments[0] = 0.12
    text(s, Inches(0.6), Inches(0.45), Inches(1.25), Inches(1.25), f["deal"], 66, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    if f["park"]:
        text(s, cx + Inches(4.35), cy + Inches(2.25), Inches(3.2), Inches(0.78), "含車位", 20, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    # 右上：富住通標誌(白字)放在同色系半透明底上
    logo = asset(ASSET_LOGO)
    if logo:
        lw, lh = Inches(3.9), Inches(0.72)
        box(s, SW - lw - Inches(0.75), Inches(0.4), lw + Inches(0.3), lh + Inches(0.24),
            tone(img, (0.6, 0.0, 1.0, 0.25)) if img else SKY_DK, alpha=0.55)
        rw = _X(lw)
        iw, ih = Image.open(logo).size
        s.shapes.add_picture(logo, _X(SW - lw - Inches(0.6)), _Y(Inches(0.52)), rw, int(rw * ih / iw))
    # 右下：專案人員(原本最後一頁的聯絡資訊併到封面)
    agents = [f"{nm} {phones.get(nm, '')}".strip() for nm in f["agents"][:3]] if show_agents else []
    if agents:
        line = "專案人員：" + "　".join(agents)
        bw = int(_text_w(line, 16) / SX) + Inches(0.5)
        bx = SW - bw - Inches(0.45)
        box(s, bx, Inches(6.55), bw, Inches(0.6), tone(img, (0.55, 0.85, 1.0, 1.0)) if img else SKY_DK, alpha=0.62)
        text(s, bx + Inches(0.2), Inches(6.55), bw - Inches(0.3), Inches(0.6), line, 16, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)


def slide_basic(deck):
    f = deck.f
    s = deck.new("基本資料")
    cards = [(f["price_label"], f"{f['price']}", f["price_unit"]), ("權狀坪數", f"{f['area']}", "坪"), ("格局", f["layout"] or "—", ""),
             ("樓層", f["floor"] or "—", ""), ("屋齡", f["age"] or "—", ""), ("車位", f["park"] or "無", "")]
    # 上半部卡片縮小一點，空間讓給下面的明細(也是重點)
    x0, y0, cw, chh, gx, gy = Inches(0.5), Inches(1.2), Inches(3.95), Inches(1.12), Inches(0.24), Inches(0.16)
    for i, (k, v, unit) in enumerate(cards):
        x, y = x0 + (cw + gx) * (i % 3), y0 + (chh + gy) * (i // 3)
        box(s, x, y, cw, chh, WHITE, LINE)
        if i == 0:
            gold(s, x, y, Inches(0.12), chh, radius=False)
        else:
            box(s, x, y, Inches(0.12), chh, SKY, radius=False)
        text(s, x + Inches(0.3), y + Inches(0.08), cw - Inches(0.4), Inches(0.36), k, 15, True, INK)
        big = 32 if len(v) <= 6 else (26 if len(v) <= 10 else 20)
        t = text(s, x + Inches(0.3), y + Inches(0.42), cw - Inches(0.4), Inches(0.62), v, big, True, GOLD_LO if i == 0 else NUM,
                 anchor=MSO_ANCHOR.MIDDLE)
        if unit:
            r = t.text_frame.paragraphs[0].add_run()
            r.text = f" {unit}"
            _font(r, 16, True, INK)
    rows = [("主建物", f"{f['main']} 坪"), ("附屬建物", f"{f['attach']} 坪" + (f"（{f['attach_txt']}）" if f["attach_txt"] else "")),
            ("公共設施", f"{f['public']} 坪"),
            ("車位坪數", f"約 {f['park_ping']} 坪（含在公設內）" if f.get("park_ping") else ""),
            ("土地持分", f"{f['land']} 坪"), ("面寬／深度", f.get("dims_show", "")),
            ("建築結構", f["struct"]),
            ("座向", f["facing"]), ("管理費", f["fee"]), ("電梯", f["elev"]), ("使用分區", f["zone"]),
            ("竣工日期", f["built"]), ("鄰近學校", f["school"]), ("所在商圈", f["district"])]
    rows = [(k, v) for k, v in rows if str(v).strip() and not str(v).strip().startswith(" 坪") and str(v).strip() != "坪"
            and not _zero(v)]
    extra = []                                   # 各層面積在成果圖頁已逐層列出，這裡不重複(字太多會被縮得很小)
    if f.get("road_w"):
        extra.append(("面臨道路", f["road_w"]))
    if f.get("trades"):
        extra.append(("適合行業", "、".join(f["trades"])))
    rows = rows[:1] + extra + rows[1:]
    rows = rows[:14]
    ty = y0 + (chh + gy) * 2 + Inches(0.02)
    th = FOOTER_Y - Inches(0.12) - ty
    box(s, x0, ty, SW - Inches(1.0), th, PANEL)
    half = (len(rows) + 1) // 2
    colw = (SW - Inches(1.0)) // 2
    rh = min(Inches(0.5), (th - Inches(0.2)) // max(1, half))
    for i, (k, v) in enumerate(rows):
        cx = x0 + Inches(0.35) + colw * (i // half)
        cy = ty + Inches(0.1) + rh * (i % half)
        text(s, cx, cy, Inches(1.7), rh, k, 19, True, SKY_DK, anchor=MSO_ANCHOR.MIDDLE)
        text(s, cx + Inches(1.7), cy, colw - Inches(2.1), rh, v, 19 if len(str(v)) <= 16 else 16, True, INK, anchor=MSO_ANCHOR.MIDDLE)


def slide_overview(deck, map_png):
    f = deck.f
    s = deck.new("物件總覽")
    pts = f["points"][:5] or ["（請在物件編輯器填寫「訴求重點」）"]
    lw = Inches(6.9) if map_png else SW - Inches(1.0)
    y = Inches(1.2)
    avail = FOOTER_Y - Inches(0.1) - y
    each = min(Inches(1.3), avail // len(pts))
    for i, p in enumerate(pts):
        box(s, Inches(0.5), y, lw, each - Inches(0.14), PANEL, LINE)      # 淡色框底
        oy = y + (each - Inches(0.14) - Inches(0.55)) // 2
        gold(s, Inches(0.7), oy, Inches(0.55), Inches(0.55), oval=True)
        text(s, Inches(0.7), oy, Inches(0.55), Inches(0.55), str(i + 1), 18, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        text(s, Inches(1.45), y, lw - Inches(1.1), each - Inches(0.14), p, 17 if len(p) < 60 else 15, False, INK,
             anchor=MSO_ANCHOR.MIDDLE)
        y += each
    if map_png:
        mx, my, mw, mh = Inches(7.65), Inches(1.2), Inches(5.2), Inches(4.75)
        box(s, mx - Inches(0.08), my - Inches(0.08), mw + Inches(0.16), mh + Inches(0.16), WHITE, LINE)
        picture(s, map_png, mx, my, mw, mh, deck.tmp)
        info = "\n".join(x for x in (f"商圈：{f['district']}" if f["district"] else "", f"學區：{f['school']}" if f["school"] else "") if x)
        text(s, mx, my + mh + Inches(0.15), mw, Inches(0.9), info or f["door"], 15, True, SKY_DK)


NLSC_TILE = "https://wmts.nlsc.gov.tw/wmts/{layer}/default/GoogleMapsCompatible/{z}/{y}/{x}"


def _fetch_tile(layer, z, x, y):
    """國土測繪中心圖磚；偶爾回 0 位元組(不是被擋)→重抓 3 次，還是空就回 None(該格留白)"""
    import urllib.request
    import io as _io
    for _ in range(3):
        try:
            req = urllib.request.Request(NLSC_TILE.format(layer=layer, z=z, x=x, y=y), headers={"User-Agent": "Mozilla/5.0"})
            b = urllib.request.urlopen(req, timeout=15).read()
            if len(b) > 200:
                return Image.open(_io.BytesIO(b)).convert("RGB")
        except Exception:
            pass
    return None


def tile_map(lat, lng, z, w_px, h_px, layer="EMAP", rings=()):
    """以座標為中心拼一張 w_px×h_px 的地圖(圖磚原始像素、不縮放 → 地圖上的字清楚)，畫紅色圖釘與距離圈。
    rings=[(公尺, 標籤)]。失敗回 None"""
    from concurrent.futures import ThreadPoolExecutor
    from PIL import ImageDraw, ImageFont
    n = 2 ** z
    fx = (lng + 180) / 360 * n * 256
    fy = (1 - math.log(math.tan(math.radians(lat)) + 1 / math.cos(math.radians(lat))) / math.pi) / 2 * n * 256
    x0, y0 = int(fx - w_px / 2), int(fy - h_px / 2)
    tx0, ty0, tx1, ty1 = x0 // 256, y0 // 256, (x0 + w_px) // 256, (y0 + h_px) // 256
    jobs = [(tx, ty) for ty in range(ty0, ty1 + 1) for tx in range(tx0, tx1 + 1)]
    with ThreadPoolExecutor(8) as ex:
        tiles = list(ex.map(lambda t: _fetch_tile(layer, z, t[0], t[1]), jobs))
    if sum(t is not None for t in tiles) < len(tiles) * 0.6:
        return None
    big = Image.new("RGB", ((tx1 - tx0 + 1) * 256, (ty1 - ty0 + 1) * 256), (236, 236, 236))
    for (tx, ty), t in zip(jobs, tiles):
        if t is not None:
            big.paste(t, ((tx - tx0) * 256, (ty - ty0) * 256))
    im = big.crop((x0 - tx0 * 256, y0 - ty0 * 256, x0 - tx0 * 256 + w_px, y0 - ty0 * 256 + h_px))
    dr = ImageDraw.Draw(im, "RGBA")
    cx, cy = w_px / 2, h_px / 2
    mpp = 156543.03392 * math.cos(math.radians(lat)) / n          # 每像素幾公尺
    try:
        font = ImageFont.truetype("msjhbd.ttc", max(18, int(h_px / 42)))
    except Exception:
        font = ImageFont.load_default()
    for meters, label in rings:
        r = meters / mpp
        dr.ellipse((cx - r, cy - r, cx + r, cy + r), outline=(196, 39, 34, 210), width=max(3, int(h_px / 300)))
        tw = dr.textlength(label, font=font)
        lx, ly = cx + r * 0.707 - tw / 2, cy - r * 0.707 - font.size * 0.7
        dr.rounded_rectangle((lx - 8, ly - 4, lx + tw + 8, ly + font.size * 1.3), 8, fill=(196, 39, 34, 225))
        dr.text((lx, ly), label, font=font, fill="white")
    # 紅色水滴圖釘(尖端在座標上)
    pr = max(16, int(h_px / 34))
    dr.polygon([(cx, cy), (cx - pr * 0.82, cy - pr * 1.6), (cx + pr * 0.82, cy - pr * 1.6)], fill=(196, 39, 34))
    dr.ellipse((cx - pr, cy - pr * 2.6, cx + pr, cy - pr * 0.6), fill=(196, 39, 34), outline="white", width=max(3, pr // 6))
    dr.ellipse((cx - pr * 0.38, cy - pr * 1.98, cx + pr * 0.38, cy - pr * 1.22), fill="white")
    return im


def case_coords(case_dir, base_dir, d):
    """案件座標：案件資料夾 4.其他相關/data.json → 主程式 data.json(同地號才算) → 回 (lat, lng) 或 None"""
    lots = []
    for p in (os.path.join(case_dir, "4.其他相關", "data.json"), os.path.join(base_dir or "", "data.json")):
        try:
            with open(p, encoding="utf-8") as fh:
                lots = json.load(fh)
            for e in lots or []:
                c = str(e.get("coordinates", "")).split(",")
                if len(c) >= 2 and e.get("lot_number") and str(e["lot_number"]) in os.path.basename(case_dir):
                    return float(c[0]), float(c[1])
        except Exception:
            continue
    return None


def slide_location(deck, lat, lng):
    """地理位置：左邊大地圖(約 3 公里範圍、500 公尺/1 公里距離圈)，右邊航照近照＋商圈/學區/交通"""
    f = deck.f
    mx, my, mw, mh = Inches(0.5), Inches(1.2), Inches(8.3), FOOTER_Y - Inches(0.12) - Inches(1.2)
    rw_px = int(_X(mw) / 914400 * 200)                     # 200 dpi
    rh_px = int(rw_px * _Y(mh) / _X(mw))
    big = tile_map(lat, lng, 16, rw_px, rh_px, "EMAP", rings=((500, "500 公尺・步行約 6 分"), (1000, "1 公里")))
    if big is None:
        log("（地理位置：國土測繪圖磚抓不到，略過這頁）")
        return False
    s = deck.new("地理位置")
    p_big = os.path.join(deck.tmp, "loc_map.png")
    big.save(p_big, optimize=True)
    box(s, mx - Inches(0.06), my - Inches(0.06), mw + Inches(0.12), mh + Inches(0.12), WHITE, LINE)
    picture(s, p_big, mx, my, mw, mh, deck.tmp, lossless=True)
    rx, rw = Inches(9.1), Inches(3.75)
    ah = Inches(2.55)
    aw_px = int(_X(rw) / 914400 * 200)
    ah_px = int(aw_px * _Y(ah) / _X(rw))
    aerial = tile_map(lat, lng, 18, aw_px, ah_px, "PHOTO2")
    y = Inches(1.2)
    if aerial is not None:
        p_a = os.path.join(deck.tmp, "loc_aerial.jpg")
        aerial.save(p_a, quality=90)
        box(s, rx - Inches(0.06), y - Inches(0.06), rw + Inches(0.12), ah + Inches(0.12), WHITE, LINE)
        picture(s, p_a, rx, y, rw, ah, deck.tmp)
        text(s, rx + Inches(0.1), y + ah - Inches(0.42), Inches(1.6), Inches(0.36), "航照近照", 13, True, WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
        y += ah + Inches(0.2)
    traffic = ""
    for p in f.get("points") or []:                         # 從訴求重點挑出講交通的那一句
        m = re.search(r"[^，。；,]*(捷運|輕軌|車站|火車|高鐵|交流道|公車|快速道路|國道)[^，。；,]*", p)
        if m:
            traffic = m.group(0).strip()
            break
    rows = [(k, v) for k, v in (("所在商圈", f.get("district")), ("鄰近學校", f.get("school")), ("交通", traffic)) if v]
    if not rows:
        rows = [("地址", f.get("door", ""))]
    avail = FOOTER_Y - Inches(0.55) - y
    each = min(Inches(1.05), avail // len(rows))
    for k, v in rows:
        box(s, rx, y, rw, each - Inches(0.1), PANEL, LINE)
        text(s, rx + Inches(0.18), y + Inches(0.04), rw - Inches(0.3), Inches(0.38), k, 14, True, SKY_DK)
        text(s, rx + Inches(0.18), y + Inches(0.38), rw - Inches(0.3), each - Inches(0.5), v, 15 if len(v) <= 22 else 13, True, INK)
        y += each
    text(s, rx, FOOTER_Y - Inches(0.48), rw, Inches(0.36), "圖資：內政部國土測繪中心（電子地圖、正射影像）", 11, False, INK)
    return True


def find_layouts(roots, limit=2):
    """使用者自己放的格局平面圖(0.謄本、4.其他相關、案件資料夾裡，檔名含「平面圖/格局/配置圖」，不含建物測量成果圖)，最多 2 張"""
    out = []
    for r in roots:
        for sub in ("0.謄本", "4.其他相關", ""):
            for h in sorted(glob.glob(os.path.join(r, sub, "*"))):
                b = os.path.basename(h)
                if (h.lower().endswith((".png", ".jpg", ".jpeg", ".webp", ".pdf")) and re.search(r"平面圖|格局|配置圖", b)
                        and "成果圖" not in b and b.lower() not in {os.path.basename(x).lower() for x in out}):
                    out.append(h)
    return out[:limit] if limit else out


def layout_features(f):
    """格局平面圖右邊的重點：住起來怎樣(採光、陽台、衛浴、格局特色)優先，從「訴求重點」抓；
    坪數上一頁成果圖已經列過，只有重點不夠 6 項時才補回來"""
    t = "　".join(f.get("points") or [])
    light = re.search(r"([四三雙兩二]面採光|邊間)", t)
    light = light.group(1).replace("兩面", "雙面").replace("二面", "雙面") if light else ""
    balc = re.search(r"(前後大?陽台|雙陽台|景觀陽台|大陽台|前陽台|後陽台)", t)
    balc = balc.group(1).replace("大陽台", "陽台") if balc else ""
    if balc == "陽台":
        balc = "大陽台"
    if not balc and "陽台" in (f.get("attach_txt") or ""):
        balc = "有陽台"
    bath = ""
    if re.search(r"(雙|兩|二)(套)?衛", t):
        bath = "雙衛"
    if bath and re.search(r"衛[^，。；]{0,6}開窗|開窗[^，。；]{0,4}衛", t):
        bath += "開窗"
    if not bath:
        m = re.search(r"(\d+)衛", str(f.get("layout") or ""))
        if m and int(m.group(1)) >= 2:
            bath = f"{m.group(1)} 套衛浴"
    t2 = re.sub(r"挑高(大廳|門廳|大堂)", "", t)          # 「挑高大廳」是公設，不是這戶挑高
    feats = [k for k in ("格局方正", "無暗房", "挑高", "雙主臥", "大面寬") if k in t2]
    if "方正" in t and "格局方正" not in feats:
        feats.insert(0, "格局方正")
    rows = [(k, v) for k, v in (("格局", f.get("layout")), ("採光", light), ("陽台", balc), ("衛浴", bath),
                                ("格局特色", "、".join(feats[:3])), ("座向", f.get("facing")), ("樓層", f.get("floor"))) if v]
    for k, v in (("權狀坪數", f"{f['area']} 坪" if f.get("area") else ""),
                 ("主建物", f"{f['main']} 坪" if f.get("main") else "")):
        if len(rows) < 6 and v:
            rows.append((k, v))
    return rows[:7]


def slide_layout(deck, src, cont=False):
    """格局平面圖：左邊整張圖(不裁切、原像素不糊)，右邊格局與坪數重點"""
    f = deck.f
    if src.lower().endswith(".pdf"):
        import fitz
        doc = fitz.open(src)
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(4, 4), alpha=False)
        png = os.path.join(deck.tmp, f"layout_{abs(hash(src)) % 10 ** 6}.png")
        pix.save(png)
        doc.close()
    else:
        png = src
    s = deck.new("格局平面圖" + ("（續）" if cont else ""))
    x, y, w, h = Inches(0.5), Inches(1.2), Inches(8.3), FOOTER_Y - Inches(0.12) - Inches(1.2)
    box(s, x, y, w, h, WHITE, LINE)
    picture(s, png, x + Inches(0.12), y + Inches(0.12), w - Inches(0.24), h - Inches(0.24), deck.tmp, mode="contain", lossless=True)
    rx, rw = Inches(9.1), Inches(3.75)
    rows = layout_features(f)
    yy = Inches(1.2)
    each = min(Inches(0.95), (FOOTER_Y - Inches(0.7) - yy) // max(1, len(rows)))
    for k, v in rows:
        box(s, rx, yy, rw, each - Inches(0.1), PANEL, LINE)
        text(s, rx + Inches(0.2), yy, Inches(1.3), each - Inches(0.1), k, 15, True, SKY_DK, anchor=MSO_ANCHOR.MIDDLE)
        text(s, rx + Inches(1.45), yy, rw - Inches(1.6), each - Inches(0.1), str(v), 18 if len(str(v)) <= 12 else 14, True, INK,
             anchor=MSO_ANCHOR.MIDDLE)
        yy += each
    text(s, rx, FOOTER_Y - Inches(0.55), rw, Inches(0.4), "平面圖為示意，實際以現況及權狀為準", 12, False, INK)


def image_plan_png(path, tmpdir):
    """電傳謄本的成果圖是圖片檔：取原始像素(不放大縮小)、找中間圖框裁切(找不到就整張)，存 PNG 無損"""
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    im = _scan_plan_crop(im)
    out = os.path.join(tmpdir, f"planimg_{os.getpid()}_{abs(hash(path)) % 10 ** 6}.png")
    im.save(out, optimize=True)
    return out


def slide_floorplan(deck, src, cont=False):
    f = deck.f
    s = deck.new("建物測量成果圖" + ("（續）" if cont else ""))
    if src.lower().endswith(".pdf"):
        png = pdf_region_png(src, deck.tmp, plan_clip, "plan")
    else:
        png = image_plan_png(src, deck.tmp)
    x, y, w, h = Inches(0.5), Inches(1.2), Inches(8.3), FOOTER_Y - Inches(0.12) - Inches(1.2)
    box(s, x, y, w, h, WHITE, LINE)
    picture(s, png, x + Inches(0.1), y + Inches(0.1), w - Inches(0.2), h - Inches(0.2), deck.tmp, mode="contain", lossless=True)
    rx, rw = Inches(9.1), Inches(3.75)
    # 總價再秀一次(提醒)
    gold(s, rx, Inches(1.2), rw, Inches(1.05))
    text(s, rx + Inches(0.25), Inches(1.2), Inches(1.4), Inches(1.05), f["price_label"], 18, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    pt = f"{f['price']} {f['price_unit']}"
    text(s, rx + Inches(1.2), Inches(1.2), rw - Inches(1.4), Inches(1.05), pt, 30 if len(pt) <= 7 else 24, True, WHITE,
         PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
    text(s, rx, Inches(2.4), rw, Inches(0.45), "登記面積（坪）", 17, True, SKY_DK)
    rows = [("主建物", f["main"]), ("附屬建物", f["attach"]), ("公共設施", f["public"])]
    rows = [(k, v) for k, v in rows if str(v).strip() and not _zero(v)]
    if f.get("park_ping"):
        rows.append(("　其中車位", f["park_ping"]))      # 車位坪數含在公設內
    if f.get("floors") and len(f["floors"]) > 1 and len(rows) <= 2:     # 透天：沒有公設 → 改列各層面積(較有用)
        rows += [("　" + k, f"{v:g}") for k, v in f["floors"][:4]]
    rows.append(("權狀合計", f["area"]))
    rh, step = (Inches(0.72), Inches(0.84)) if len(rows) <= 4 else (Inches(0.62), Inches(0.72))
    yy = Inches(2.9)
    for i, (k, v) in enumerate(rows):
        last = i == len(rows) - 1
        if last:
            gold(s, rx, yy, rw, rh)
        else:
            box(s, rx, yy, rw, rh, WHITE, LINE)
        sub = k.startswith("　")
        text(s, rx + Inches(0.25), yy, Inches(1.8), rh, k, 15 if sub else 17, True, WHITE if last else INK, anchor=MSO_ANCHOR.MIDDLE)
        text(s, rx + Inches(1.9), yy, Inches(1.65), rh, str(v), 18 if sub else 22, True, WHITE if last else NUM, PP_ALIGN.RIGHT,
             MSO_ANCHOR.MIDDLE)
        yy += step
    text(s, rx, yy + Inches(0.05), rw, Inches(0.6), "資料來源：地政事務所建物測量成果圖", 12, False, INK)


def slide_cadastral(deck, pdf):
    f = deck.f
    s = deck.new("地籍圖")
    png = pdf_region_png(pdf, deck.tmp, cadastral_clip, "cadastral")
    x, y, w, h = Inches(0.5), Inches(1.2), Inches(8.3), FOOTER_Y - Inches(0.12) - Inches(1.2)
    box(s, x, y, w, h, WHITE, LINE)
    picture(s, png, x + Inches(0.1), y + Inches(0.1), w - Inches(0.2), h - Inches(0.2), deck.tmp, mode="contain", lossless=True)
    rx, rw = Inches(9.1), Inches(3.75)
    text(s, rx, Inches(1.2), rw, Inches(0.45), "土地資料", 17, True, SKY_DK)
    rows = [("地段地號", f.get("lots") or ""), ("使用分區", f["zone"]),
            ("建蔽／容積", f.get("bcr_far") or ""), ("土地持分", f"{f['land']} 坪" if f["land"] else ""),
            ("面寬／深度" + ("" if f.get("dims") else "（量測參考）"), f.get("dims") or f.get("dims_cad", "")),
            ("臨路", f.get("roads_show", "")),
            ("比例尺", cadastral_scale(pdf))]
    rows = [(k, v) for k, v in rows if str(v).strip()]
    yy = Inches(1.72)
    pitch = min(Inches(1.04), (FOOTER_Y - Inches(0.65) - yy) // max(1, len(rows)))   # 項目多就縮格高，不被頁尾擋住
    for k, v in rows:
        box(s, rx, yy, rw, pitch - Inches(0.1), WHITE, LINE)
        text(s, rx + Inches(0.22), yy + Inches(0.04), rw - Inches(0.3), Inches(0.34), k, 13, True, INK)
        text(s, rx + Inches(0.22), yy + Inches(0.31), rw - Inches(0.3), pitch - Inches(0.42), str(v), 15, True, NUM)
        yy += pitch
    text(s, rx, yy + Inches(0.05), rw, Inches(0.6), "資料來源：地政事務所地籍圖謄本", 12, False, INK)


def slide_photos(deck, title, items, note=None):
    if not items:
        return
    s = deck.new(title)
    top = Inches(1.2)
    bottom = FOOTER_Y - Inches(0.12) if not note else FOOTER_Y - Inches(0.7)
    deck.photos(s, items, Inches(0.5), top, SW - Inches(1.0), bottom - top)
    if note:
        text(s, Inches(0.5), bottom + Inches(0.06), SW - Inches(1.0), Inches(0.55), note, 17, True, SKY_DK, anchor=MSO_ANCHOR.MIDDLE)


def crop_104(pdf, tmpdir, rows=8, max_ratio=None):
    """104 實價網 PDF：從表格標題列往下截前幾筆成交(字大、好讀)。
    max_ratio(高/寬)有給 → 在放得下的範圍內盡量多截幾筆(最多 15 筆)"""
    import fitz
    doc = fitz.open(pdf)
    pg = doc[0]
    head = pg.search_for("交易標的")
    if not head:
        return None
    top = head[0].y0 - 6
    # 每筆左邊有編號「1.」「2.」…(在第一欄) → 用編號定位每一列，截到第 rows 筆的下緣
    want = 15 if max_ratio else rows
    ys = []
    for k in range(1, want + 2):
        hits = [r for r in pg.search_for(f"{k}.") if r.x0 < head[0].x1 + 30 and r.y0 > head[0].y1]
        if not hits:
            break
        ys.append(min(h.y0 for h in hits))
    words_all = pg.get_text("words")
    xs0 = [w[0] for w in words_all if w[1] >= top] + [w[2] for w in words_all if w[1] >= top]
    tab_w = (max(xs0) - min(xs0) + 12) if xs0 else pg.rect.width
    if len(ys) >= 2:
        pitch = (ys[-1] - ys[0]) / (len(ys) - 1)
        n = min(want, len(ys))
        if max_ratio:                       # 依版面能放的高度決定筆數
            while n > 3 and (ys[n - 1] + pitch * 0.62 - top) / tab_w > max_ratio * 1.15:   # 容許稍窄、把高度填滿
                n -= 1
        bottom = ys[n - 1] + pitch * 0.62
    else:
        bottom = min(pg.rect.y1 - 20, top + 360)
    words = pg.get_text("words")
    xs = [w[0] for w in words if top <= w[1] <= bottom] + [w[2] for w in words if top <= w[1] <= bottom]
    left, right = (min(xs) - 6, max(xs) + 6) if xs else (pg.rect.x0, pg.rect.x1)
    clip = fitz.Rect(max(0, left), top, min(pg.rect.x1, right), bottom)
    pix = pg.get_pixmap(matrix=fitz.Matrix(3, 3), clip=clip, alpha=False)
    out = os.path.join(tmpdir, f"p104_{os.getpid()}.png")
    pix.save(out)
    # 截到的這幾筆的單價(「@34.5萬」)→ 右邊顯示範圍與平均(照抄來源數字，不推算)
    units = []
    for w in pg.get_text("words", clip=clip):
        m = re.match(r"^(?:地)?[@＠]([\d.]+)萬", w[4])    # 104 用全形「＠」
        if m:
            units.append(float(m.group(1)))
    return out, [u for u in units if u > 0]


NOTE_H = Inches(0.8)          # 行情頁下方「行情說明」只留一兩行
MARKET_LW = Inches(9.75)      # 行情頁左邊表格區寬(右欄收窄，表格放大)


def market_ratio():
    """行情頁左邊表格區能放的 高/寬(實際頁面)：給 crop_104 決定截幾筆"""
    lw = MARKET_LW - Inches(0.2)
    h = FOOTER_Y - Inches(0.12) - Inches(1.2) - Inches(0.2) - Inches(0.15) - NOTE_H
    return _Y(h) / _X(lw)


def slide_market(deck, market_img, source, units=None):
    f = deck.f
    s = deck.new("市場行情")
    lx, lw = Inches(0.5), MARKET_LW
    bottom = FOOTER_Y - Inches(0.12)
    note_top = Inches(1.2)
    if market_img:
        # 表格框貼合圖片高度(A4 比螢幕高，下面空間給「行情說明」)
        iw, ih = Image.open(market_img).size
        real_w = _X(lw - Inches(0.2))
        dh = int(real_w * ih / iw / SY)
        dh = min(dh, bottom - Inches(1.2) - Inches(0.2) - Inches(0.15) - NOTE_H)
        box(s, lx, Inches(1.2), lw, dh + Inches(0.2), WHITE, LINE)
        picture(s, market_img, lx + Inches(0.1), Inches(1.3), lw - Inches(0.2), dh, deck.tmp, mode="contain")
        note_top = Inches(1.2) + dh + Inches(0.35)
    if bottom - note_top > Inches(0.4):
        box(s, lx, note_top, lw, bottom - note_top, PANEL, LINE)
        text(s, lx + Inches(0.2), note_top + Inches(0.06), lw - Inches(0.4), bottom - note_top - Inches(0.12),
             "行情說明：（請依實價登錄補充周邊近期租金行情與比較）" if f["deal"] == "租" else
             "行情說明：（請依實價登錄補充同社區、周邊近期成交與比較）", 15, False, INK, anchor=MSO_ANCHOR.MIDDLE)
    rx, rw = Inches(10.45), Inches(2.4)
    gold(s, rx, Inches(1.2), rw, Inches(1.9))
    text(s, rx + Inches(0.22), Inches(1.3), rw, Inches(0.45), "本案租金" if f["deal"] == "租" else "本案開價", 16, True, WHITE)
    text(s, rx + Inches(0.22), Inches(1.72), rw - Inches(0.35), Inches(0.75), f"{f['price']} {f['price_unit']}", 32, True, WHITE)
    if f["unit"] and f["deal"] != "租":
        text(s, rx + Inches(0.22), Inches(2.38), rw - Inches(0.35), Inches(0.68), f"建物單價約\n{f['unit']} 萬/坪", 15, True, WHITE,
             spacing=1.0)
    y = Inches(3.3)
    if units:
        box(s, rx, y, rw, Inches(2.2), WHITE, LINE)
        text(s, rx + Inches(0.2), y + Inches(0.1), rw - Inches(0.3), Inches(0.45), f"表列 {len(units)} 筆成交單價", 15, True, SKY_DK)
        text(s, rx + Inches(0.2), y + Inches(0.6), rw - Inches(0.3), Inches(0.6),
             f"{min(units):.1f}～{max(units):.1f}", 26, True, NUM)
        text(s, rx + Inches(0.2), y + Inches(1.15), rw - Inches(0.3), Inches(0.4), "萬/坪", 14, True, INK)
        text(s, rx + Inches(0.2), y + Inches(1.55), rw - Inches(0.3), Inches(0.5),
             f"平均約 {sum(units) / len(units):.1f} 萬/坪", 17, True, GOLD_LO)
        y += Inches(2.35)
    text(s, rx, y, rw, Inches(0.9), f"資料來源：{source}（提供客戶前請再查證）", 11, False, INK)


def slide_back(deck, phones):
    """封底：直式 LOGO＋Full Life＋公司電話(可選)"""
    f = deck.f
    s = deck.new()
    box(s, 0, 0, SW, SH, WHITE, radius=False)
    gold(s, 0, SH - Inches(0.35), SW, Inches(0.35), radius=False)
    lv = asset(ASSET_LOGO_V)
    if lv:
        iw, ih = Image.open(lv).size
        h = _Y(Inches(3.3))
        w = int(h * iw / ih)
        s.shapes.add_picture(lv, (REAL_W - w) // 2, _Y(Inches(0.9)), w, h)
    fl = asset(ASSET_FULLLIFE)
    if fl:
        iw, ih = Image.open(fl).size
        w = _X(Inches(3.6))
        s.shapes.add_picture(fl, (REAL_W - w) // 2, _Y(Inches(4.45)), w, int(w * ih / iw))
    text(s, Inches(0.5), Inches(5.35), SW - Inches(1.0), Inches(0.7), f"{COMPANY}　{COMPANY_TEL}", 26, True, SKY_DK,
         PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    agents = [f"{nm} {phones.get(nm, '')}".strip() for nm in f["agents"][:3]]
    if agents:
        text(s, Inches(0.5), Inches(6.05), SW - Inches(1.0), Inches(0.6), "專案人員：" + "　".join(agents), 18, True, INK,
             PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------- 主流程
def build(case_dir, base_dir, d, exp, chosen, out_path, show_agents=True, back_cover=False, template="building", sel=None):
    """sel＝「頁面內容」步驟的選擇 {單元: {on, pick}}；None＝照預設自動帶"""
    f = facts(d)
    try:
        load_captions([p["path"] for p in chosen if p.get("path")])
    except Exception as e:
        log(f"（照片文字標籤檢查失敗：{e}）")
    if REAL_H > REAL_W:                     # A4 直式：另一套直式排版(保留富住通外框、LOGO、售/租章、封底)
        import marketing_deck_portrait
        return marketing_deck_portrait.build(sys.modules[__name__], case_dir, base_dir, d, exp, chosen, out_path, f,
                                             show_agents=show_agents, back_cover=True, sel=sel, template=template)
    if template == "villa":
        import marketing_deck_street
        return marketing_deck_street.build(sys.modules[__name__], case_dir, base_dir, d, exp, chosen, out_path, f,
                                           show_agents=show_agents, sel=sel)
    tmp = tempfile.mkdtemp(prefix="deck_")
    deck = Deck(f, tmp)
    by = {c: ordered([p for p in chosen if p["cat"] == c]) for c in CATS}
    cover = (by["封面"] or by["陽台景觀"] or by[EXT_CAT] or by["客廳"] or [None])[0]
    roots = [r for r in (case_dir, exp) if r]

    def first(pattern):
        for r in roots:
            hits = sorted(glob.glob(os.path.join(r, pattern)))
            if hits:
                return hits[0]
        return None
    map_png = first(os.path.join("4.其他相關", "gui_map_screenshot.png"))
    # 建物測量成果圖：電子謄本 PDF 優先；沒有就用電傳謄本的圖片檔(檔名含「建物測量成果圖」，可能分 _1、_2 多張)
    plans = []
    for r in roots:
        hits = sorted(glob.glob(os.path.join(r, "0.謄本", "*建物測量成果圖*.pdf")))
        if hits:
            plans = hits[:1]
            break
    if not plans:
        for r in roots:
            hits = sorted(h for h in glob.glob(os.path.join(r, "0.謄本", "*建物測量成果圖*"))
                          if h.lower().endswith((".png", ".jpg", ".jpeg")))
            if hits:
                plans = hits[:3]
                break
    plans = sel_paths(sel, "plan", plans)
    cad_pdf = first(os.path.join("0.謄本", "*地籍圖*標註*.pdf")) or first(os.path.join("0.謄本", "*地籍圖*.pdf"))
    cad_pages = sel_paths(sel, "cadastral", [cad_pdf] if cad_pdf else [])
    if cad_pages:
        cad_pdf = cad_pages[0]
    # 行情：使用者選的(可多份/不放)；沒選過 → 104(社區優先)、沒有才內政部；租件只用「租賃」那份
    rent = f["deal"] == "租"
    market_picks = sel_paths(sel, "market", market_default(roots, rent))

    # 面寬／深度：物件編輯器有填就用；沒填 → 有建物先量建物測量成果圖(向量檔、面積對得上才用)，
    #   再不行才依地籍圖比例尺量土地(標註版有路名可判斷臨路邊)，都標明「參考」
    f["dims_show"], f["roads_show"] = f["dims"], ""
    if not f["dims"] and plans and plans[0].lower().endswith(".pdf"):
        try:
            big = max([v for _, v in f.get("floors") or []] or [_num(f.get("main")) or 0])
            pd_ = plan_dims(plans[0], big * 3.30579 if big else None)
        except Exception as e:
            pd_ = None
            log(f"（成果圖量測失敗：{e}）")
        if pd_:
            f["dims"] = f["dims_show"] = f"約 {pd_[0]:g}／{pd_[1]:g} 米（依建物測量成果圖，參考）"
            log(f"面寬／深度：物件編輯器沒填，依建物測量成果圖量測 {f['dims_show']}")
    if cad_pdf:
        try:
            dm = parcel_dims(cad_pdf)
        except Exception as e:
            dm = None
            log(f"（地籍圖量測失敗：{e}）")
        if dm:
            f["roads_show"] = "、".join(f"{n} {l:g}米" for n, l in dm["roads"])
            if not f["dims"]:
                f["dims_show"] = f"約 {dm['frontage']:g}／{dm['depth']:g} 米（參考）"
                f["dims_cad"] = f"約 {dm['frontage']:g}／{dm['depth']:g} 米" + ("" if dm["road"] else "（未判斷臨路）")
                log(f"面寬／深度：物件編輯器沒填，依地籍圖量測建議 {f['dims_show']}")
    phones = lookup_phones(f["agents"], base_dir)
    log("製作：封面"); slide_cover(deck, cover, phones, show_agents)
    log("製作：基本資料"); slide_basic(deck)
    if sel_on(sel, "overview"):
        log("製作：物件總覽"); slide_overview(deck, map_png)
    ll = case_coords(case_dir, base_dir, d) if sel_on(sel, "location") else None
    if ll:
        log("製作：地理位置（下載國土測繪地圖）")
        slide_location(deck, *ll)
    else:
        log("（data.json 沒有這個案件的座標，略過地理位置頁）")
    for src in cad_pages:
        log("製作：地籍圖"); slide_cadastral(deck, src)
    for i, src in enumerate(plans):
        log("製作：建物測量成果圖" + (f"（{i + 1}）" if len(plans) > 1 else "")); slide_floorplan(deck, src, cont=i > 0)
    # 格局平面圖：照片裡選了「格局平面圖」的(2.照片)優先，再加 0.謄本/4.其他相關 裡檔名有平面圖/格局的；同檔名只放一次
    layouts = [p["path"] for p in by[LAYOUT_CAT]]
    names = {os.path.basename(x).lower() for x in layouts}
    layouts += [x for x in sel_paths(sel, "layout", find_layouts(roots)) if os.path.basename(x).lower() not in names]
    for i, src in enumerate(layouts[:6]):
        log(f"製作：格局平面圖（{os.path.basename(src)}）"); slide_layout(deck, src, cont=i > 0)
    log("製作：照片頁")
    for title, cats, cap in PHOTO_PAGES:
        items = ordered([p for p in chosen if p["cat"] in cats])
        note = None
        if "陽台景觀" in cats:
            # 頁名/「X 樓景觀視野」要跟照片相符：分類是「陽台景觀」且說明提到景觀/視野才算景觀；
            # 前後陽台(或說明沒提景觀的，例 陽台洗衣空間)不加景觀字樣
            def _cap(p):
                return (p.get("cap") if p.get("cap") is not None else caption_of(p)) or ""
            view = any(p["cat"] == "陽台景觀" and ("景" in _cap(p) or "視野" in _cap(p)) for p in items)
            plain = any(not (p["cat"] == "陽台景觀" and ("景" in _cap(p) or "視野" in _cap(p))) for p in items)
            title = "景觀・陽台" if (view and plain) else ("景觀陽台" if view else "前後陽台")
            if view and f["floor"]:
                note = f"{f['floor'].split('/')[0].strip()}景觀視野"
        start = 0
        for k, size in enumerate(split_pages(len(items), cap)):   # 一頁放不下 → 平均分成幾頁(5 張=3+2，不會剩 1 張孤零零)
            slide_photos(deck, title + ("（續）" if k else ""), items[start:start + size], note)
            start += size
    market_slides(deck, f, market_picks, tmp)
    if back_cover:
        log("製作：封底"); slide_back(deck, phones)
    unique_shape_ids(deck.prs)
    deck.prs.save(out_path)
    try:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
    except Exception:
        pass
    return deck.n


# ---------------------------------------------------------------- 頁面內容(使用者選要不要、用哪份資料)
def _globs(roots, *pats):
    out, seen = [], set()
    for r in roots:
        for pat in pats:
            for h in sorted(glob.glob(os.path.join(r, pat))):
                k = os.path.basename(h).lower()
                if k not in seen:
                    seen.add(k)
                    out.append(h)
    return out


def _no_data(path):
    """104 查詢結果 PDF 寫「無成交資料」→ True(這份放進簡報只會是空表)"""
    if not path.lower().endswith(".pdf"):
        return False
    try:
        import fitz
        doc = fitz.open(path)
        t = doc[0].get_text()
        doc.close()
        return "無成交資料" in t
    except Exception:
        return False


def market_default(roots, rent):
    """沒選過時的行情預設：104(社區那份優先於地段那份，查無資料的不算)；沒有才用內政部截圖。租件只用檔名有「租賃」的"""
    p104 = [p for p in _globs(roots, os.path.join("3.行情", "104實價登錄-*.pdf"))
            if ("租賃" in os.path.basename(p)) == rent and not _no_data(p)]
    p104.sort(key=lambda p: ("段" in os.path.basename(p), os.path.basename(p)))
    if p104:
        return [p104[0]]
    pngs = [p for p in _globs(roots, os.path.join("3.行情", "*.png")) if ("租賃" in os.path.basename(p)) == rent]
    pick = [p for p in pngs if "內政部實價登錄" in os.path.basename(p)] or pngs
    return pick[:1]


def collect_sections(case_dir, exp, template, d):
    """「頁面內容」步驟要列的單元(依簡報順序)＋每個單元找得到的資料。
    回傳 [{key, title, desc, multi, cands:[(路徑, 顯示名)], pick:[預設路徑], on, empty}]
    empty＝沒資料時的說明(這種單元不能勾)"""
    f = facts(d)
    roots = [r for r in (case_dir, exp) if r]
    rent = f["deal"] == "租"
    villa = template == "villa"
    out = []

    def add(key, title, desc, cands=None, pick=None, multi=False, on=True, empty=""):
        cands = cands if cands is not None else []
        out.append({"key": key, "title": title, "desc": desc, "multi": multi,
                    "cands": [(p, _cand_label(p)) for p in cands], "pick": list(pick or []),
                    "on": bool(on) and (not empty or bool(cands)) and (not cands or bool(pick)),   # 有資料但預設一份都沒選(例：租件沒有租賃行情)→ 先不勾
                    "empty": empty if not cands else "", "needs": cands is not None})

    if villa:
        add("overview", "物件亮點", "面寬、臨路、坪數、樓層等數字＋物件特色", None)
        if f.get("floors"):
            add("floors", "樓層配置", "各樓層面積長條圖(透天專用)", None)
    else:
        add("overview", "物件總覽", "訴求重點＋地圖截圖", None)
    add("location", "地理位置", "國土測繪中心電子地圖＋航照(需要案件座標)", None)
    if villa:
        fac = _globs(roots, os.path.join("1.基本資料", "*重要設施*.pdf"))
        add("facility", "周邊機能", "重要設施地圖＋距離表", fac, fac[:1], empty="1.基本資料 沒有「重要設施」查詢結果")
        zon = _globs(roots, os.path.join("1.基本資料", "png", "*都市計畫圖*.png"),
                     os.path.join("1.基本資料", "png", "*都市計畫使用分區*.png"))
        add("zoning", "使用分區", "都市計畫圖(只取地圖區)＋分區、建蔽容積", zon, zon[:1], empty="1.基本資料 沒有都市計畫截圖")
    cads = [c for c in _globs(roots, os.path.join("0.謄本", "*地籍圖*標註*.pdf"), os.path.join("0.謄本", "*地籍圖*.pdf"))
            if "合併" not in os.path.basename(c) and "_temp" not in os.path.basename(c)]      # 檔名自己改過的(例：425地號畫線地籍圖)也算
    add("cadastral", "地籍圖", "地籍圖謄本(標註版優先)", cads, cads[:1], empty="0.謄本 沒有地籍圖謄本")
    plans = _globs(roots, os.path.join("0.謄本", "*建物測量成果圖*.pdf"))
    imgs = [h for h in _globs(roots, os.path.join("0.謄本", "*建物測量成果圖*")) if h.lower().endswith((".png", ".jpg", ".jpeg"))]
    add("plan", "建物測量成果圖", "可選多張，每張一頁", plans + imgs, plans[:1] or imgs[:3], multi=True,
        empty="0.謄本 沒有建物測量成果圖")
    lays = find_layouts(roots, limit=None)
    add("layout", "格局平面圖", "0.謄本／4.其他相關 裡檔名有平面圖、格局的圖(2.照片裡的在「挑照片」選分類)", lays, lays[:4],
        multi=True, on=bool(lays), empty="沒有找到平面圖檔(可在挑照片時把照片分類成「格局平面圖」)")
    mk = _globs(roots, os.path.join("3.行情", "104實價登錄-*.pdf"), os.path.join("3.行情", "*.png"))
    mk.sort(key=lambda p: (("租賃" in os.path.basename(p)) != rent, not os.path.basename(p).startswith("104"),
                           os.path.basename(p)))
    add("market", "市場行情", "可選 104 實價網、內政部實價，或都不放；選幾份就幾頁" +
        ("（租件：請選檔名有「租賃」的）" if rent else ""), mk, market_default(roots, rent), multi=True,
        empty="3.行情 沒有實價資料，請先跑「實價登錄」「104 實價網」")
    return out


def _cand_label(path):
    b = os.path.splitext(os.path.basename(path))[0]
    for pre, tag in (("104實價登錄-", "104｜"), ("內政部實價登錄-", "內政部｜"), ("地籍圖謄本-", ""), ("建物測量成果圖-", "")):
        if b.startswith(pre):
            b = tag + b[len(pre):]
            break
    return b + ("（無成交資料）" if b.startswith("104｜") and _no_data(path) else "")


def apply_saved_sections(sections, saved):
    """上次存的選擇(on、pick)套回去；存的檔案已不存在就用預設"""
    for sec in sections:
        sv = (saved or {}).get(sec["key"])
        if not sv:
            continue
        if not sec["empty"]:
            sec["on"] = bool(sv.get("on", True))
        if sec["cands"] and "pick" in sv:
            have = {p for p, _ in sec["cands"]}
            pk = [p for p in sv["pick"] if p in have]
            if pk or not sv["pick"]:
                sec["pick"] = pk
    return sections


def sections_to_sel(sections):
    return {s["key"]: {"on": s["on"], "pick": list(s["pick"])} for s in sections}


def sel_on(sel, key):
    return not sel or key not in sel or bool(sel[key].get("on", True))


def sel_paths(sel, key, default):
    """這個單元要用的檔案：沒選過 → default；關掉 → []；選過 → 使用者選的(檔案還在的)"""
    if not sel or key not in sel:
        return list(default)
    if not sel[key].get("on", True):
        return []
    pk = [p for p in sel[key].get("pick") or [] if os.path.exists(p)]
    return pk if pk or not sel[key].get("pick") else list(default)


def market_slides(deck, f, picks, tmp):
    """依選的行情檔各做一頁：104 PDF 截表格(字大)、內政部 PNG 直接放"""
    rent = f["deal"] == "租"
    n = 0
    for p in picks:
        if p.lower().endswith(".pdf"):
            try:
                r = crop_104(p, tmp, max_ratio=market_ratio())
            except Exception as e:
                r = None
                log(f"（104 實價截取失敗：{e}）")
            if not r:
                log(f"（{os.path.basename(p)} 截不到成交表格，略過）")
                continue
            img, units, src = r[0], r[1], "104 行情通實價網"
        else:
            img, units, src = p, None, "內政部實價登錄"
        log(f"製作：市場行情（{_cand_label(p)}）")
        slide_market(deck, img, src, None if rent else units)
        n += 1
    if not picks and rent:
        log("（行情頁沒有選資料；租件請先重跑「實價登錄」「104 實價網」，會自動查租賃）")
    return n


def unique_shape_ids(prs):
    """整份簡報的圖形編號(cNvPr id)改成不重複。
    python-pptx 每頁都從 1 重新編號，iPhone 預覽(LINE 開檔就是用它)會把前一頁同編號的圖形殘影疊到下一頁
    (實例：成果圖頁冒出上一頁物件總覽的 5 個條列框、臥室頁冒出編號金圓點)"""
    n = 1
    for slide in prs.slides:
        for el in slide._element.iter(qn("p:cNvPr")):
            el.set("id", str(n))
            n += 1


def _pdf_restore_images(pptx_path, pdf_path):
    """PowerPoint 轉 PDF 一定會把圖壓到約 220ppi(地籍圖 3050px→1015px，小字變糊；改列印品質也一樣)
    → 把 PDF 裡每張被縮小的圖換回簡報裡的原圖(同一頁、長寬比相同的那張)。文字仍是向量。"""
    import io
    import fitz
    prs = Presentation(pptx_path)
    doc = fitz.open(pdf_path)
    done = set()
    for slide, page in zip(prs.slides, doc):
        pics = []
        for shp in slide.shapes:
            if shp.shape_type == 13:                      # PICTURE
                try:
                    im = Image.open(io.BytesIO(shp.image.blob))
                    if im.mode in ("RGBA", "LA", "P"):    # 有透明的(LOGO)換了會失去透明，保留 PowerPoint 版
                        continue
                    pics.append([shp.image.blob, im.size[0], im.size[1], False,
                                 fitz.Rect(shp.left / 12700, shp.top / 12700, (shp.left + shp.width) / 12700,
                                           (shp.top + shp.height) / 12700)])
                except Exception:
                    pass
        for info in page.get_images(full=True):
            xref, smask, w, h = info[0], info[1], info[2], info[3]
            if xref in done or not h or smask:            # 有透明遮罩的(外框、LOGO)不換
                continue
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            best = None
            for p in pics:
                if p[3] or p[1] <= w * 1.05 or abs(p[1] / p[2] - w / h) >= 0.01 * (w / h):
                    continue
                # ★還要「位置對得上」：只比長寬比會配錯(A4 平面圖跟整頁外框比例幾乎一樣 → 平面圖被貼成整頁背景)
                if rects and not any(abs(r.x0 - p[4].x0) < 6 and abs(r.y0 - p[4].y0) < 6 and
                                     abs(r.width - p[4].width) < 8 for r in rects):
                    continue
                best = p
                break
            if best:
                page.replace_image(xref, stream=best[0])
                best[3] = True
                done.add(xref)
    tmp = pdf_path + ".tmp.pdf"
    doc.save(tmp, garbage=3, deflate=True)
    doc.close()
    os.replace(tmp, pdf_path)
    return len(done)


def export_pdf(pptx_path):
    """用 PowerPoint 另存一份 PDF(傳 LINE/手機看最穩)。沒裝 PowerPoint 或失敗回 None。
    COM 存中文路徑會失敗 → 複製到英文暫存資料夾轉檔再搬回；PowerPoint 原本就開著時絕不 Quit(會關掉使用者的簡報)"""
    import shutil
    import subprocess
    pdf_path = os.path.splitext(pptx_path)[0] + ".pdf"
    tmpd = tempfile.mkdtemp(prefix="mdpdf_")
    src, dst = os.path.join(tmpd, "deck.pptx"), os.path.join(tmpd, "deck.pdf")
    try:
        shutil.copy2(pptx_path, src)
        ps = ("$ErrorActionPreference='Stop';"
              "$was = $null -ne (Get-Process POWERPNT -ErrorAction SilentlyContinue);"
              "$pp = New-Object -ComObject PowerPoint.Application;"
              "$p = $pp.Presentations.Open($env:MD_SRC, $true, $false, $false);"      # 🔒 路徑用環境變數傳
              "try { $p.SaveAs($env:MD_DST, 32) } finally { $p.Close() };"
              "if (-not $was) { $pp.Quit() }")
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], timeout=180,
                       capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                       env=dict(os.environ, MD_SRC=src, MD_DST=dst))
        if not os.path.exists(dst):
            return None
        _pdf_restore_images(src, dst)
        shutil.copy2(dst, pdf_path)
        return pdf_path
    except Exception as e:
        log(f"（另存 PDF 失敗：{e}）")
        return None
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)


def default_out_path(case_dir, exp, d):
    f = facts(d)
    folder = exp or case_dir
    name = re.sub(r'[\\/:*?"<>|]', "", f["door"] or os.path.basename(case_dir))
    p = os.path.join(folder, f"{name}_不動產行銷報告.pptx")
    if os.path.exists(p):
        p = os.path.join(folder, f"{name}_不動產行銷報告_{datetime.now():%m%d_%H%M}.pptx")
    return p


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case-dir", required=True)
    ap.add_argument("--base-dir", default=os.path.dirname(SCRIPT_DIR))
    ap.add_argument("--auto", action="store_true", help="不開挑照片視窗，直接用自動分類(測試用)")
    ap.add_argument("--out", default=None, help="輸出檔(測試用)")
    ap.add_argument("--page", default="a4l", choices=list(PAGES), help="a4l=A4 橫式(預設)、169=螢幕 16:9")
    ap.add_argument("--no-open", action="store_true", help="完成後不自動開啟(測試用)")
    ap.add_argument("--template", default="building", help="--auto 時用的版型(測試用)")
    a = ap.parse_args()
    set_page(a.page)
    d, exp = load_case(a.case_dir, a.base_dir)
    photos = find_photos(a.case_dir, exp)
    log(f"照片：{len(photos)} 張（案件資料夾＋{'匯出資料夾 ' + exp if exp else '無匯出資料夾'}）")
    if a.auto:
        chosen = [p for p in photos if p["cat"] != "不使用"]
        out = a.out or default_out_path(a.case_dir, exp, d)
    else:
        import marketing_deck_ui
        last = load_settings(a.case_dir, photos)
        sug = suggest_template(d.get("型式", ""))
        res = marketing_deck_ui.pick(photos, facts(d), default_out_path(a.case_dir, exp, d), last.get("page") or a.page,
                                     on_save=lambda pg, tp, ag, bc, th, sec: save_settings(a.case_dir, photos, pg, tp, ag, bc, th, sec),
                                     get_sections=lambda tp: apply_saved_sections(collect_sections(a.case_dir, exp, tp, d),
                                                                                  last.get("sections")),
                                     template=last.get("template") or "building", suggest=sug,
                                     show_agents=last.get("show_agents", True), back_cover=last.get("back_cover", False),
                                     theme=last.get("theme") or "fzt")
        if not res:
            log("已取消（版型、照片分類與說明已記住，下次打開會帶回）")
            return
        chosen, out, page, template, show_agents, back_cover, theme, sel = res
        set_page(page)
        set_theme(theme)
        log(f"版型：{TEMPLATES.get(template, TEMPLATES['building'])[1]}　版面：{PAGES[page][0]}　色系：{THEMES[THEME]['name']}")
    n = build(a.case_dir, a.base_dir, d, exp, chosen, out, show_agents if not a.auto else True,
              back_cover if not a.auto else False, template if not a.auto else a.template, None if a.auto else sel)
    log(f"✓ 行銷簡報完成（{n} 頁）：{out}")
    log("正在另存一份 PDF（傳 LINE、手機看用 PDF 最不會跑版）…")
    pdf = export_pdf(out)
    if pdf:
        log(f"✓ PDF：{pdf}")
    else:
        log("（這台電腦沒有 PowerPoint 或轉檔失敗，未產生 PDF；可用 PowerPoint 開啟後「另存為 PDF」）")
    if not a.no_open:
        try:
            os.startfile(out)
        except Exception:
            pass


if __name__ == "__main__":
    main()
