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
PAGES = {"a4l": ("A4 橫式", 11.693, 8.268), "169": ("螢幕 16:9", 13.333, 7.5)}
SX = SY = 1.0                      # 設計座標 → 實際頁面 的縮放(由 set_page 設定)
REAL_W, REAL_H = SW, SH


def set_page(key):
    global SX, SY, REAL_W, REAL_H
    _, w, h = PAGES.get(key, PAGES["a4l"])
    REAL_W, REAL_H = Inches(w), Inches(h)
    SX, SY = REAL_W / SW, REAL_H / SH


def _X(v):
    return int(v * SX)


def _Y(v):
    return int(v * SY)
IMG_EXT = (".jpg", ".jpeg", ".png", ".webp", ".bmp")

# 照片分類(挑照片視窗的下拉選單)；依檔名關鍵字猜
CATS = ["封面", "大樓外觀", "大廳公設", "客廳", "餐廳廚房", "臥室", "衛浴", "其他空間", "陽台景觀", "車位", "不使用"]
_GUESS = [
    ("車位", ("停車", "車位", "車道", "parking")),
    ("其他空間", ("洗衣", "工作陽台", "儲藏", "更衣", "玄關", "走道")),
    ("衛浴", ("浴", "衛", "廁", "bath")),
    ("餐廳廚房", ("廚", "餐", "kitchen")),
    ("客廳", ("客廳", "客餐", "起居", "living")),
    ("臥室", ("臥", "房間", "主臥", "bedroom", "study")),
    ("陽台景觀", ("陽台", "景觀", "視野", "夜景", "view", "露台")),
    ("大廳公設", ("大廳", "公設", "健身", "ktv", "泳池", "中庭", "信箱", "交誼", "lobby", "gym")),
    ("大樓外觀", ("外觀", "大樓", "建築", "門面", "facade")),
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
    return f"{n:,.0f}" if n == int(n) else f"{n:,.1f}"


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
    """上次挑的照片分類/說明/版面 → 套回 photos；回傳上次的版面(沒有存過回 None)"""
    try:
        with open(settings_path(case_dir), "r", encoding="utf-8") as f:
            st = json.load(f)
    except Exception:
        return None
    saved = st.get("photos") or {}
    n = 0
    for p in photos:
        v = saved.get(p["name"].lower())
        if v:
            p["cat"] = v.get("cat", p["cat"])
            if "cap" in v:
                p["cap"] = v["cap"]
            n += 1
    log(f"已載入上次的設定（{n} 張照片的分類/說明）")
    return st.get("page")


def save_settings(case_dir, photos, page):
    """記住這次挑的照片分類/說明/版面(存在案件資料夾 4.其他相關)"""
    st = {"version": 1, "saved": datetime.now().strftime("%Y-%m-%d %H:%M"), "page": page,
          "photos": {p["name"].lower(): dict({"cat": p["cat"]}, **({"cap": p["cap"]} if "cap" in p else {}))
                     for p in photos}}      # 說明只記「使用者改過的」，沒改過的下次依分類/檔名自動帶
    try:
        os.makedirs(os.path.dirname(settings_path(case_dir)), exist_ok=True)
        with open(settings_path(case_dir), "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=1)
    except Exception as e:
        log(f"（設定沒存成功：{e}）")


def guess_cat(name):
    n = name.lower()
    for cat, keys in _GUESS:
        if any(k in n for k in keys):
            return cat
    return "不使用"


def caption_of(p):
    """照片說明：檔名有描述就用(去副檔名、去英文)；IMG_xxxx 就用分類名"""
    base = os.path.splitext(p["name"])[0]
    if re.match(r"^(img|dsc|dji|pxl|s__|line_|photo|image)[\s_\-]*\d", base, re.I) or re.fullmatch(r"[A-Za-z0-9_\s\-]+", base):
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
    return {
        "title": d.get("案名") or door_short,
        "short": (re.sub(r"[（(].*?[）)]", "", d.get("案名") or "").strip() or door_short)[:26],
        "door": door_short or door,
        "price": _fmt_money_wan(d.get("總價", "")),
        "unit": d.get("建物單價", ""),
        "area": d.get("總建坪") or d.get("產權建坪", ""),
        "main": d.get("主建物坪數", ""), "attach": d.get("附屬建物坪數", ""), "attach_txt": attach_txt,
        "public": d.get("公共設施坪數", ""), "land": d.get("總地坪", ""),
        "floor": f"{floor} 樓" + (f" / 共 {total_fl} 樓" if total_fl else "") if floor else (f"共 {total_fl} 樓" if total_fl else ""),
        "age": age, "built": d.get("竣工日期", ""), "layout": layout, "park": park_txt,
        "struct": d.get("建築結構", ""), "facing": d.get("座向", ""),
        "fee": f"{_fmt_money_wan(d.get('管理費'))} 元/月" if d.get("管理費") else "",
        "zone": d.get("都市土地使用分區", "") or d.get("非都市土地使用分區", ""),
        "elev": elev, "school": d.get("鄰近學校", ""), "district": d.get("所在商圈", ""),
        "points": points, "type": d.get("型式", ""), "agents": d.get("承辦人列表") or [],
    }


# ---------------------------------------------------------------- 圖片
def prep_image(path, box_w, box_h, mode="cover", tmpdir=None, max_side=1600, enhance=True):
    """事先把圖裁成框的比例(cover)或整張放進(contain)，縮小、轉正、微提亮 → 存成 JPEG/PNG 暫存檔。
    回傳 (暫存檔, 實際寬, 實際高)；contain 時寬高可能比框小"""
    im = Image.open(path)
    im = ImageOps.exif_transpose(im)
    has_alpha = im.mode in ("RGBA", "LA", "P")
    im = im.convert("RGB")
    ratio = box_w / box_h
    w, h = im.size
    if mode == "cover":
        if w / h > ratio:
            nw = int(h * ratio)
            im = im.crop(((w - nw) // 2, 0, (w - nw) // 2 + nw, h))
        else:
            nh = int(w / ratio)
            top = max(0, int((h - nh) * 0.45))           # 稍微偏上(天花板少、主體多)
            im = im.crop((0, top, w, top + nh))
        out_w, out_h = box_w, box_h
    else:
        if w / h > ratio:
            out_w, out_h = box_w, int(box_w * h / w)
        else:
            out_w, out_h = int(box_h * w / h), box_h
    im.thumbnail((max_side, max_side), Image.LANCZOS)
    if enhance and mode == "cover":
        im = ImageEnhance.Brightness(im).enhance(1.05)
        im = ImageEnhance.Color(im).enhance(1.08)
    fd, out = tempfile.mkstemp(suffix=".png" if (has_alpha and mode != "cover") else ".jpg", dir=tmpdir)
    os.close(fd)
    if out.endswith(".png"):
        im.save(out, "PNG", optimize=True)
    else:
        im.save(out, "JPEG", quality=86, optimize=True)
    return out, out_w, out_h


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
ASSET_LOGO = "fzt_logo.png"
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


def picture(slide, path, x, y, w, h, tmpdir, mode="cover"):
    """放一張圖到框裡(事先裁好、不用 srcRect)。contain 時置中。回傳(圖形, 處理後的圖檔)"""
    x, y, w, h = _X(x), _Y(y), _X(w), _Y(h)
    img, iw, ih = prep_image(path, int(w), int(h), mode, tmpdir)
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

    def strip(y0, y1, sample_rows):
        row = a.crop((0, sample_rows[0], a.width, sample_rows[1])).resize((a.width, 1), Image.BOX)
        return row.resize((W, y1 - y0), Image.BILINEAR)
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
def slide_cover(deck, cover, phones):
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
    text(s, cx + Inches(0.4), cy + Inches(2.25), Inches(3.8), Inches(0.78), f"開價 {f['price']} 萬", 28, True, WHITE,
         PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    if f["park"]:
        text(s, cx + Inches(4.35), cy + Inches(2.25), Inches(3.2), Inches(0.78), "含車位", 20, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    # 右上：富住通標誌(白字)放在同色系半透明底上
    logo = asset(ASSET_LOGO)
    if logo:
        lw, lh = Inches(3.9), Inches(0.72)
        box(s, SW - lw - Inches(0.75), Inches(0.4), lw + Inches(0.3), lh + Inches(0.24),
            tone(img, (0.6, 0.0, 1.0, 0.25)) if img else SKY_DK, alpha=0.55)
        rw = _X(lw)
        s.shapes.add_picture(logo, _X(SW - lw - Inches(0.6)), _Y(Inches(0.52)), rw, int(rw * 196 / 1064))
    # 右下：專案人員(原本最後一頁的聯絡資訊併到封面)
    agents = [f"{nm} {phones.get(nm, '')}".strip() for nm in f["agents"][:3]]
    if agents:
        line = "專案人員：" + "　".join(agents)
        bw = int(_text_w(line, 16) / SX) + Inches(0.5)
        bx = SW - bw - Inches(0.45)
        box(s, bx, Inches(6.55), bw, Inches(0.6), tone(img, (0.55, 0.85, 1.0, 1.0)) if img else SKY_DK, alpha=0.62)
        text(s, bx + Inches(0.2), Inches(6.55), bw - Inches(0.3), Inches(0.6), line, 16, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)


def slide_basic(deck):
    f = deck.f
    s = deck.new("基本資料")
    cards = [("總價", f"{f['price']}", "萬"), ("權狀坪數", f"{f['area']}", "坪"), ("格局", f["layout"] or "—", ""),
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
            ("公共設施", f"{f['public']} 坪"), ("土地持分", f"{f['land']} 坪"), ("建築結構", f["struct"]),
            ("座向", f["facing"]), ("管理費", f["fee"]), ("電梯", f["elev"]), ("使用分區", f["zone"]),
            ("竣工日期", f["built"]), ("鄰近學校", f["school"]), ("所在商圈", f["district"])]
    rows = [(k, v) for k, v in rows if str(v).strip() and not str(v).strip().startswith(" 坪") and str(v).strip() != "坪"]
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


def slide_floorplan(deck, pdf):
    f = deck.f
    s = deck.new("建物測量成果圖")
    pngs = pdf_to_png(pdf, deck.tmp)
    if not pngs:
        return
    x, y, w, h = Inches(0.5), Inches(1.2), Inches(8.3), FOOTER_Y - Inches(0.12) - Inches(1.2)
    box(s, x, y, w, h, WHITE, LINE)
    picture(s, pngs[0], x + Inches(0.1), y + Inches(0.1), w - Inches(0.2), h - Inches(0.2), deck.tmp, mode="contain")
    rx, rw = Inches(9.1), Inches(3.75)
    # 總價再秀一次(提醒)
    gold(s, rx, Inches(1.2), rw, Inches(1.05))
    text(s, rx + Inches(0.25), Inches(1.2), Inches(1.4), Inches(1.05), "總價", 18, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    text(s, rx + Inches(1.2), Inches(1.2), rw - Inches(1.4), Inches(1.05), f"{f['price']} 萬", 30, True, WHITE, PP_ALIGN.RIGHT,
         MSO_ANCHOR.MIDDLE)
    text(s, rx, Inches(2.4), rw, Inches(0.45), "登記面積（坪）", 17, True, SKY_DK)
    rows = [("主建物", f["main"]), ("附屬建物", f["attach"]), ("公共設施", f["public"]), ("權狀合計", f["area"])]
    yy = Inches(2.9)
    for i, (k, v) in enumerate(rows):
        last = i == len(rows) - 1
        if last:
            gold(s, rx, yy, rw, Inches(0.72))
        else:
            box(s, rx, yy, rw, Inches(0.72), WHITE, LINE)
        text(s, rx + Inches(0.25), yy, Inches(1.8), Inches(0.72), k, 17, True, WHITE if last else INK, anchor=MSO_ANCHOR.MIDDLE)
        text(s, rx + Inches(1.9), yy, Inches(1.65), Inches(0.72), str(v), 22, True, WHITE if last else NUM, PP_ALIGN.RIGHT,
             MSO_ANCHOR.MIDDLE)
        yy += Inches(0.84)
    text(s, rx, yy + Inches(0.05), rw, Inches(0.6), "資料來源：地政事務所建物測量成果圖", 12, False, INK)


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
             "行情說明：（請依實價登錄補充同社區、周邊近期成交與比較）", 15, False, INK, anchor=MSO_ANCHOR.MIDDLE)
    rx, rw = Inches(10.45), Inches(2.4)
    gold(s, rx, Inches(1.2), rw, Inches(1.9))
    text(s, rx + Inches(0.22), Inches(1.3), rw, Inches(0.45), "本案開價", 16, True, WHITE)
    text(s, rx + Inches(0.22), Inches(1.72), rw - Inches(0.35), Inches(0.75), f"{f['price']} 萬", 32, True, WHITE)
    if f["unit"]:
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


# ---------------------------------------------------------------- 主流程
def build(case_dir, base_dir, d, exp, chosen, out_path):
    f = facts(d)
    tmp = tempfile.mkdtemp(prefix="deck_")
    deck = Deck(f, tmp)
    by = {c: [p for p in chosen if p["cat"] == c] for c in CATS}
    cover = (by["封面"] or by["陽台景觀"] or by["大樓外觀"] or by["客廳"] or [None])[0]
    roots = [r for r in (case_dir, exp) if r]

    def first(pattern):
        for r in roots:
            hits = sorted(glob.glob(os.path.join(r, pattern)))
            if hits:
                return hits[0]
        return None
    map_png = first(os.path.join("4.其他相關", "gui_map_screenshot.png"))
    plan_pdf = first(os.path.join("0.謄本", "建物測量成果圖*.pdf"))
    # 行情：104 實價網(字大)優先，社區那份優先於地段那份；沒有才用內政部截圖
    market_img, source = None, "內政部實價登錄"
    p104 = [p for r in roots for p in sorted(glob.glob(os.path.join(r, "3.行情", "104實價登錄-*.pdf")))]
    p104.sort(key=lambda p: ("段" in os.path.basename(p), os.path.basename(p)))
    units = None
    for p in p104:
        try:
            r = crop_104(p, tmp, max_ratio=market_ratio())
            if r:
                market_img, units = r
        except Exception as e:
            log(f"（104 實價截取失敗：{e}）")
        if market_img:
            source = "104 行情通實價網"
            break
    if not market_img:
        market_img = first(os.path.join("3.行情", "內政部實價登錄*.png")) or first(os.path.join("3.行情", "*.png"))

    log("製作：封面"); slide_cover(deck, cover, lookup_phones(f["agents"], base_dir))
    log("製作：基本資料"); slide_basic(deck)
    log("製作：物件總覽"); slide_overview(deck, map_png)
    if plan_pdf:
        log("製作：建物測量成果圖"); slide_floorplan(deck, plan_pdf)
    log("製作：照片頁")
    slide_photos(deck, "室內精選｜客餐廳", by["客廳"] + by["餐廳廚房"])
    slide_photos(deck, "室內精選｜臥室・衛浴", by["臥室"] + by["衛浴"] + by["其他空間"])
    views = by["陽台景觀"][:3]
    slide_photos(deck, "景觀陽台", views, f"{f['floor'].split('/')[0].strip()}景觀視野" if f["floor"] else None)
    slide_photos(deck, "大樓外觀・公設", by["大樓外觀"] + by["大廳公設"] + by["車位"])
    log("製作：市場行情"); slide_market(deck, market_img, source, units)
    deck.prs.save(out_path)
    try:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
    except Exception:
        pass
    return deck.n


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
        last_page = load_settings(a.case_dir, photos)
        res = marketing_deck_ui.pick(photos, facts(d), default_out_path(a.case_dir, exp, d), last_page or a.page,
                                     on_save=lambda pg: save_settings(a.case_dir, photos, pg))
        if not res:
            log("已取消（照片分類與說明已記住，下次打開會帶回）")
            return
        chosen, out, page = res
        set_page(page)
    n = build(a.case_dir, a.base_dir, d, exp, chosen, out)
    log(f"✓ 行銷簡報完成（{n} 頁）：{out}")
    if not a.no_open:
        try:
            os.startfile(out)
        except Exception:
            pass


if __name__ == "__main__":
    main()
