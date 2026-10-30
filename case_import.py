"""整理既有資料夾：使用者自己建的資料夾(可能有好幾層子資料夾)→ 每個檔案判斷該放系統的哪個資料夾。
判斷依據(由強到弱)：檔案內容(PDF/Word/Excel/PowerPoint 文字、圖片是文件時用 OCR)＞檔名＞使用者自己的資料夾名稱。
只負責「判斷」與「照判斷結果複製」；原檔永遠不動。要 fitz / rapidocr → 主程式用 python_embedded 跑。
用法：python case_import.py <來源資料夾> <輸出.json> [--no-ocr]
  輸出 {"files": [{"src", "rel", "target", "kind", "reason", "conf", "lot"}...], "lots": {地號: 次數}}"""
import os
import re
import sys
import json
import zipfile

TARGETS = ["0.謄本", "1.基本資料", "2.照片", "3.行情", "4.其他相關", "5.委託", "6.成交", "根目錄"]
IMG = (".jpg", ".jpeg", ".png", ".heic", ".webp", ".bmp", ".gif", ".tif", ".tiff")
SKIP_NAMES = {"thumbs.db", "desktop.ini", ".ds_store"}

# (目標, 種類, 關鍵字) — 由上往下先中先贏；同一份文件同時有好幾種字時，越上面越優先
RULES = [
    # 自家產生的文件先認(物調表裡也會出現「履約保證」「契約編號」，不能被成交搶走)
    ("6.成交", "實價登錄申報書", ["實價登錄申報書", "不動產成交案件實際資訊申報書", "成交案件實際資訊申報"]),
    ("6.成交", "業績分配表", ["業績分配表"]),
    ("根目錄", "物件調查表", ["物件個案調查表", "物調表", "不動產說明書"]),
    ("根目錄", "行銷簡報", ["不動產行銷報告", "行銷簡報", "行銷報告"]),
    ("5.委託", "委託契約書", ["委託銷售契約", "委託租賃契約", "專任委託", "一般委託", "委託契約書", "委託書", "委任契約"]),
    ("5.委託", "斡旋／要約", ["斡旋", "要約書", "買方出價"]),
    ("6.成交", "買賣／租賃契約", ["不動產買賣契約", "買賣契約書", "房屋租賃契約", "租賃契約書", "租約", "類公證",
                            "價金信託", "點交證明", "成交確認"]),
    ("0.謄本", "謄本", ["土地登記第一類謄本", "土地登記第二類謄本", "建物登記第一類謄本", "建物登記第二類謄本",
                      "登記第一類謄本", "登記第二類謄本", "土地登記謄本", "建物登記謄本", "地籍異動索引", "異動索引",
                      "土地參考資訊", "地政電傳", "全方位地籍資料查詢", "謄本"]),
    ("0.謄本", "地籍圖／成果圖", ["地籍圖謄本", "地籍圖", "建物測量成果圖", "成果圖", "地籍測繪", "畫線-", "平面圖", "位置圖"]),
    ("1.基本資料", "土地增值稅", ["土地增值稅", "土增稅", "增值稅"]),
    ("1.基本資料", "使用執照／建照", ["使用執照", "使照", "建造執照", "建照", "建築執照"]),
    ("1.基本資料", "分區／都計", ["使用分區", "都市計畫", "國土分區", "國土功能分區", "分區證明", "土地使用分區", "分區"]),
    ("1.基本資料", "環境／地圖資料", ["重要設施", "環境設施", "週邊", "周邊", "地質", "地籍便民", "國土測繪", "V523",
                              "建築線", "地籍圖資", "淹水", "土壤液化"]),
    ("3.行情", "實價行情", ["實價登錄", "不動產交易實價", "實價查詢", "成交月份", "漲跌率", "訂約日期", "實價", "行情",
                          "104實價", "樂居"]),
]
# 使用者自己的資料夾名稱 → 弱提示(內容、檔名都判斷不出來時才用)
DIR_HINTS = [
    ("0.謄本", ["謄本", "地籍"]), ("2.照片", ["照片", "相片", "現場", "外觀", "室內", "photo", "pic"]),
    ("3.行情", ["行情", "實價", "比價"]), ("5.委託", ["委託", "委任", "專任"]),
    ("6.成交", ["成交", "簽約", "合約", "契約", "租約"]), ("1.基本資料", ["基本資料", "使照", "分區", "稅"]),
]
SYS_FILES = {"gui_map_screenshot.png", "photo_captions_cache.json"}
SYS_JSON = re.compile(r"^(data|data_final|trans_.*|.*_tax_result|都市計畫資料彙整|行銷簡報設定|合併列印.*|V523資料.*|"
                      r"全國土地使用分區資料.*|業績分配表設定|.*設定)\.json$")
LOT_RE = re.compile(r"([一-鿿]{1,6}段(?:[一-鿿]{1,3}小段)?)\s*0*(\d{1,4})-?0*(\d{0,4})\s*地號")


# ───────────── 讀內容 ─────────────
def _zip_text(path, members):
    try:
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if any(re.fullmatch(m, n) for m in members)]
            raw = " ".join(z.read(n).decode("utf-8", "ignore") for n in names[:6])
        import html
        return html.unescape(re.sub(r"<[^>]+>", "", raw))          # Excel 會把中文存成 &#36001; 這種
    except Exception:
        return ""


def _pdf_text(path, ocr):
    try:
        import fitz
        with fitz.open(path) as doc:
            txt = "".join(doc[i].get_text() for i in range(min(2, len(doc))))
            if len(txt.strip()) < 20 and ocr and len(doc):        # 掃描的 PDF：第一頁轉圖 OCR
                pix = doc[0].get_pixmap(dpi=110)
                from PIL import Image
                im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                txt += _ocr_img(im, ocr)
        return txt
    except Exception:
        return ""


def _ocr_img(im, ocr, top=0.55):
    """OCR 只看上半部(契約、謄本的標題都在最上面)，縮到 960 寬：一張約 0.5 秒"""
    import numpy as np
    if top:
        im = im.crop((0, 0, im.width, int(im.height * top)))
    k = min(1.0, 960 / max(im.size))
    if k < 1:
        im = im.resize((int(im.width * k), int(im.height * k)))
    try:
        res, _ = ocr(np.asarray(im.convert("RGB"))[:, :, ::-1], use_cls=False)
        return " ".join(r[1] for r in res or [] if r[2] >= 0.5)
    except Exception:
        return ""


def _doc_like(im):
    """圖片像不像『拍文件／掃描』：白底多、顏色少、深色細字。回傳 0~1"""
    from PIL import ImageStat
    s = im.convert("RGB").resize((160, 160))
    hsv = s.convert("HSV")
    px = list(zip(s.getdata(), hsv.getdata()))
    white = sum(1 for (r, g, b), (h, sa, v) in px if v > 175 and sa < 45) / len(px)
    dark = sum(1 for (r, g, b), (h, sa, v) in px if v < 90) / len(px)
    sat = ImageStat.Stat(hsv).mean[1] / 255
    score = white * 1.3 - sat * 1.2 + (0.15 if 0.02 < dark < 0.35 else 0)
    return max(0.0, min(1.0, score))


def read_content(path, ocr=None):
    """回傳 (文字, 附加資訊 dict)"""
    low = path.lower()
    info = {}
    if low.endswith(".pdf"):
        return _pdf_text(path, ocr), info
    if low.endswith(".docx"):
        return _zip_text(path, [r"word/document\.xml"]), info
    if low.endswith((".xlsx", ".xlsm")):
        return _zip_text(path, [r"xl/sharedStrings\.xml", r"xl/worksheets/sheet\d+\.xml"]), info
    if low.endswith(".pptx"):
        return _zip_text(path, [r"ppt/slides/slide[1-3]\.xml"]), info
    if low.endswith(IMG):
        try:
            from PIL import Image, ImageOps
            im = Image.open(path)
            im.draft("RGB", (1400, 1400))                          # JPEG 直接縮小解碼(手機原圖 4032 很慢)
            im = ImageOps.exif_transpose(im)
            ex = im.getexif() if hasattr(im, "getexif") else {}
            info["camera"] = bool(ex.get(271) or ex.get(272))      # Make / Model
            info["doc"] = round(_doc_like(im), 2)
            if ocr and info["doc"] >= 0.45:
                t = _ocr_img(im, ocr)
                info["doc_text"] = len(t.replace(" ", ""))
                return t, info
        except Exception:
            pass
        return "", info
    if low.endswith((".txt", ".csv")):
        try:
            return open(path, encoding="utf-8", errors="ignore").read(4000), info
        except Exception:
            return "", info
    return "", info


# ───────────── 判斷 ─────────────
GENERIC = {"謄本", "分區", "周邊", "週邊", "實價", "行情", "成果圖", "平面圖", "位置圖", "地籍圖", "使照", "建照",
           "增值稅", "土增稅", "地質", "畫線-", "物調表", "行銷報告", "委託書", "租約", "斡旋"}
OWN_DOCS = {"實價登錄申報書", "業績分配表", "物件調查表", "行銷簡報", "委託契約書"}   # 檔名寫明就以檔名為準


def _match(text, generic=True):
    t = text.replace(" ", "").replace("\u3000", "")
    for target, kind, keys in RULES:
        for k in keys:
            if (generic or k not in GENERIC) and k in t:
                return target, kind, k
    return None


def classify(path, rel_dirs=(), ocr=None):
    """回傳 dict(target, kind, reason, conf=高/中/低, lot)。
    順序：系統檔 → 檔名寫明的自家文件/系統附件編號 → 內容(只用明確的詞) → 檔名 → 圖片(看是不是拍文件＋資料夾名稱) → 資料夾名稱"""
    name = os.path.basename(path)
    low = name.lower()
    stem = os.path.splitext(name)[0]
    dirs = "/".join(rel_dirs)
    if SYS_JSON.match(name) or name in SYS_FILES or low.endswith((".txt", ".log")) and "紀錄" in dirs:
        return dict(target="4.其他相關", kind="系統資料", reason="系統產生的資料檔", conf="高", lot="")
    name_hit = _match(stem)
    sys_no = re.match(r"^\d{2}(\.\d|-\d)*_", stem)
    if name_hit and (name_hit[1] in OWN_DOCS or name_hit[0] == "0.謄本"):   # 檔名寫明謄本/地籍圖/成果圖也以檔名為準
        return dict(target=name_hit[0], kind=name_hit[1], reason=f"檔名有「{name_hit[2]}」", conf="高", lot="")
    if sys_no:
        return dict(target="1.基本資料", kind=(name_hit[1] if name_hit and name_hit[0] == "1.基本資料" else "附件"),
                    reason="系統附件編號開頭", conf="高", lot="")
    if low.endswith((".txt", ".log", ".ini", ".json")):
        return dict(target="4.其他相關", kind="文字/資料檔", reason="文字或資料檔", conf="低", lot="")
    text, info = read_content(path, None if (name_hit and low.endswith(".pdf")) else ocr)   # 檔名就認得的 PDF 不再 OCR(慢)
    lot = ""
    m = LOT_RE.search(text.replace(" ", "")) or LOT_RE.search(stem)
    if m:
        sect = re.sub(r"^.*[區鄉鎮市]", "", m.group(1)) or m.group(1)        # 「市前鎮區盛興段」→「盛興段」
        lot = f"{sect}{m.group(2)}" + (f"-{m.group(3)}" if m.group(3) and m.group(3) != "0" else "")
    # ① 內容(只用明確的詞)
    hit = _match(text[:3000], generic=False) if text else None
    if hit:
        return dict(target=hit[0], kind=hit[1], reason=f"內容有「{hit[2]}」", conf="高", lot=lot)
    # ② 檔名
    if name_hit:
        return dict(target=name_hit[0], kind=name_hit[1], reason=f"檔名有「{name_hit[2]}」", conf="中", lot=lot)
    # ③ 圖片
    if low.endswith(IMG):
        hint = next((t for t, keys in DIR_HINTS if any(k in dirs for k in keys)), None)
        if info.get("doc", 0) >= 0.45 and info.get("doc_text", 0) >= 60:   # 白白的、而且讀得出一大段字＝拍文件(招牌只有幾個字)
            if hint in ("5.委託", "6.成交", "0.謄本"):
                return dict(target=hint, kind="文件照片", reason=f"像拍文件、放在「{dirs}」", conf="中", lot=lot)
            return dict(target="4.其他相關", kind="文件照片(看不出是什麼)", reason="像拍文件但讀不出關鍵字", conf="低", lot=lot)
        if hint in ("5.委託", "6.成交"):                 # 例：成交\出租時現況\IMG_7056.JPG＝點交時的現況照
            return dict(target=hint, kind="照片", reason=f"放在「{dirs}」的照片", conf="中", lot=lot)
        line = name.startswith("S__") and not info.get("camera")       # LINE 傳來的圖：常是拍契約，請人看一下
        return dict(target="2.照片", kind="照片", reason="一般照片" + ("（相機拍的）" if info.get("camera") else
                                                                  "（LINE 傳的，請確認）" if line else ""),
                    conf="高" if info.get("camera") else ("低" if line else "中"), lot=lot)
    # ④ 使用者資料夾名稱
    for target, keys in DIR_HINTS:
        if any(k in dirs for k in keys):
            return dict(target=target, kind="依資料夾名稱", reason=f"放在「{dirs}」", conf="低", lot=lot)
    if low.endswith(".pptx"):
        return dict(target="根目錄", kind="簡報", reason="PowerPoint 檔", conf="低", lot=lot)
    return dict(target="4.其他相關", kind="其他", reason="看不出類別", conf="低", lot=lot)


def scan(src, ocr=None, progress=None):
    out, lots = [], {}
    all_files = []
    for dp, dns, fns in os.walk(src):
        dns[:] = [d for d in dns if not d.startswith((".", "~"))]
        for fn in fns:
            if fn.lower() in SKIP_NAMES or fn.startswith("~$") or fn.lower().endswith((".lnk", ".tmp", ".ini")):
                continue
            all_files.append(os.path.join(dp, fn))
    for i, p in enumerate(all_files):
        rel = os.path.relpath(p, src)
        r = classify(p, tuple(os.path.dirname(rel).replace("\\", "/").split("/")) if os.path.dirname(rel) else (), ocr)
        r.update(src=p, rel=rel)
        out.append(r)
        if r["lot"]:
            lots[r["lot"]] = lots.get(r["lot"], 0) + 1
        if progress:
            progress(i + 1, len(all_files), rel)
    return dict(files=out, lots=lots)


def main():
    src, dst = sys.argv[1], sys.argv[2]
    ocr = None
    if "--no-ocr" not in sys.argv:
        try:
            from rapidocr_onnxruntime import RapidOCR
            ocr = RapidOCR()
        except Exception:
            ocr = None
    res = scan(src, ocr, lambda i, n, rel: print(f"__PROGRESS__ {i}/{n} {rel}", flush=True))
    with open(dst, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)
    print("OK", flush=True)


if __name__ == "__main__":
    main()
