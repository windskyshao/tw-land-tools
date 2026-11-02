"""行銷簡報「工業・廠房」版型：工業地、廠房共用一個版型，依有沒有建物自動換基本資料欄位。
參考公司工業地產團隊報告書的手法：滿版實景＋半透明圖層(字有底襯著、照片還看得到)。
★半透明一律先用 PIL 合成進圖片再放進簡報，iPad/手機預覽不會漏圖；字仍是簡報文字，可以直接改。
★分割方案、應注意事項存在案件 4.其他相關/工業版補充.json(「頁面內容」步驟按「填寫」)。
共用 marketing_deck 的畫圖工具(M=marketing_deck 模組本身)。"""
import json
import math
import os
import re
import urllib.request

EXTRA_FILE = "工業版補充.json"
SPLIT_COLORS = ("B8862B", "B5322E", "2B7A9B", "4F7A3A")      # A/B/C/D 方案標籤色
COUNTY_CODE = {"臺北": "A", "台北": "A", "臺中": "B", "台中": "B", "基隆": "C", "臺南": "D", "台南": "D", "高雄": "E", "新北": "F",
               "宜蘭": "G", "桃園": "H", "嘉義市": "I", "新竹縣": "J", "苗栗": "K", "南投": "M", "彰化": "N", "新竹市": "O", "雲林": "P",
               "嘉義縣": "Q", "屏東": "T", "花蓮": "U", "臺東": "V", "台東": "V", "澎湖": "X", "金門": "W", "連江": "Z"}
FACTORY_URL = "https://map.windsky-sky.com/data/factories/{}.geojson"
SERVICES = ("土地變更相關諮詢", "後續建廠作業諮詢", "工廠毗連擴廠諮詢", "土壤檢測改良諮詢", "大型產業用地整體開發諮詢")


def load_extra(case_dir):
    try:
        with open(os.path.join(case_dir, "4.其他相關", EXTRA_FILE), encoding="utf-8") as fh:
            return json.load(fh) or {}
    except Exception:
        return {}


def save_extra(case_dir, data):
    p = os.path.join(case_dir, "4.其他相關", EXTRA_FILE)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=1)


def _val(v):
    s = str(v if v is not None else "").strip()
    try:
        return "" if float(s.replace(",", "")) == 0 else s
    except ValueError:
        return s


def is_plant(f):
    """有建物或型式寫廠房 → 廠房內容；否則工業地"""
    return bool(_val(f.get("area"))) or "廠" in str(f.get("type", "")) + str(f.get("use", ""))


def _kind(f):
    return "廠房" if is_plant(f) else "工業地"


def _rgb(hex_):
    return tuple(int(hex_[i:i + 2], 16) for i in (0, 2, 4))


def _veil(im, shapes):
    """在圖上蓋半透明形狀(合成進圖片)。shapes=[(種類, 座標, (r,g,b), 不透明度0~255)]"""
    from PIL import Image, ImageDraw
    w, h = im.size
    ss = 2
    lay = Image.new("RGBA", (w * ss, h * ss), (0, 0, 0, 0))
    dr = ImageDraw.Draw(lay)
    for kind, xy, rgb, a in shapes:
        pts = [v * ss for v in xy]
        if kind == "ellipse":
            dr.ellipse(pts, fill=rgb + (a,))
        else:
            dr.rounded_rectangle(pts, radius=int(h * ss * 0.012), fill=rgb + (a,))
    lay = lay.resize((w, h), Image.LANCZOS)
    return Image.alpha_composite(im.convert("RGBA"), lay).convert("RGB")


def _aerial(M, ll, w_px, h_px, z=18, pin=True):
    return M.tile_map(ll[0], ll[1], z, w_px, h_px, "PHOTO2", pin=pin)


def _fit(path, w_px, h_px):
    """照片 cover 裁成指定像素(EXIF 轉正)"""
    from PIL import Image, ImageOps
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    return ImageOps.fit(im, (w_px, h_px), Image.LANCZOS)


def _base_image(M, photo, ll, w_px, h_px):
    """底圖：有照片用照片；沒有照片用國測航照；都沒有回 None"""
    if photo:
        return _fit(photo["path"], w_px, h_px)
    if ll:
        return _aerial(M, ll, w_px, h_px, 18, pin=False)     # 圖釘會透出標題，用在底圖時不畫
    return None


def _coords(M, case_dir, base_dir, d):
    ll = M.case_coords(case_dir, base_dir, d)
    if ll:
        return ll
    lat, lng = M._num(d.get("緯度")), M._num(d.get("經度"))
    return (lat, lng) if lat and lng else None


def build(M, case_dir, base_dir, d, exp, chosen, out_path, f, show_agents=True, sel=None):
    from pptx.dml.color import RGBColor
    t = M.THEMES[M.THEME]
    C = {"main": RGBColor.from_string(t["sky"]), "dark": RGBColor.from_string(t["sky_dk"]),
         "dark_rgb": _rgb(t["sky_dk"]), "white": M.WHITE, "panel": M.PANEL, "line": M.LINE,
         "ink": RGBColor.from_string("1F2933"), "ink2": RGBColor.from_string("3F4A56"), "red": M.STAMP_RED}
    M.INK = C["ink"]
    tmp = M.tempfile.mkdtemp(prefix="deck_")
    deck = M.Deck(f, tmp)
    deck.frame = None
    _chrome(M, deck, C, f)
    extra = load_extra(case_dir)
    by = {c: M.ordered([p for p in chosen if p["cat"] == c]) for c in M.CATS}
    cover = (by["封面"] or by["空拍"] or by[M.EXT_CAT] or by[M.STREET_CAT] or [None])[0]
    side = (by[M.EXT_CAT] or by[M.STREET_CAT] or by["空拍"] or [cover])[0]
    roots = [r for r in (case_dir, exp) if r]
    ll = _coords(M, case_dir, base_dir, d)
    cads = M.sel_paths(sel, "cadastral", [c for c in M._globs(roots, os.path.join("0.謄本", "*地籍圖*標註*.pdf"),
                                                              os.path.join("0.謄本", "*地籍圖*.pdf"))][:1])
    cards = M.sel_paths(sel, "card", M._globs(roots, os.path.join("0.謄本", "*_行銷*.png"))[:1])
    # 面寬深度：物件編輯器優先，沒填用地籍圖量
    f["dims_show"] = f["dims"]
    if not f["dims_show"] and cads:
        try:
            dm = M.parcel_dims(cads[0])
        except Exception:
            dm = None
        if dm:
            f["dims_show"] = f"約 {dm['frontage']:g}／{dm['depth']:g} 米"
    phones = M.lookup_phones(f["agents"], base_dir)
    M.log(f"製作：封面（工業版・{_kind(f)}）"); cover_slide(M, deck, C, f, cover, ll, phones, show_agents)
    M.log("製作：基本資訊"); basics(M, deck, C, f, d, side, ll)
    if cards:
        M.log("製作：地籍與現況"); card_slide(M, deck, C, f, cards[0], ll)
    else:
        for src in cads:
            M.log("製作：地籍圖"); M.slide_cadastral(deck, src)
    if M.sel_on(sel, "split") and extra.get("split"):
        M.log("製作：分割銷售方案")
        split_slide(M, deck, C, f, d, extra, ll, (M.sel_paths(sel, "split", []) or [None])[0])
    if ll and M.sel_on(sel, "location"):
        M.log("製作：交通位置"); location(M, deck, C, f, ll)
    if M.sel_on(sel, "law") and (extra.get("notes") or _law_rows(f, d)):
        M.log("製作：法規與應注意事項"); law_slide(M, deck, C, f, d, extra)
    if ll and M.sel_on(sel, "neighbors"):
        try:
            nb = nearby(d, ll)
            if nb:
                M.log("製作：周邊廠商"); neighbors_slide(M, deck, C, f, nb)
            else:
                M.log("（周邊廠商：2 公里內查無登記工廠，略過）")
        except Exception as e:
            M.log(f"（周邊廠商頁略過：{str(e)[:80]}）")
    plans = M.sel_paths(sel, "plan", M._globs(roots, os.path.join("0.謄本", "*建物測量成果圖*.pdf"))[:1])
    plans = M.plan_pages(plans)
    for i, src in enumerate(plans):
        M.log("製作：建物測量成果圖"); M.slide_floorplan(deck, src, cont=i > 0)
    M.log("製作：照片頁")
    for title, cats, cap in M.FACTORY_PHOTO_PAGES:
        items = M.ordered([p for p in chosen if p["cat"] in cats])
        if items and cover in items and len(items) > 1:
            items = [p for p in items if p is not cover] + [cover]
        start = 0
        for k, n in enumerate(M.split_pages(len(items), cap)):
            M.slide_photos(deck, title + ("（續）" if k else ""), items[start:start + n])
            start += n
    M.market_slides(deck, f, M.sel_paths(sel, "market", M.market_default(roots, f["deal"] == "租")), tmp)
    M.log("製作：專案團隊"); team(M, deck, C, f, phones, show_agents)
    M.unique_shape_ids(deck.prs)
    deck.prs.save(out_path)
    return deck.n


# ---------------------------------------------------------------- 頁框：主色短條＋標題；頁尾深色帶(公司名、頁碼)
def _chrome(M, deck, C, f):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    H = M.SH - M.FOOTER_Y

    def footer(s):
        M.box(s, 0, M.FOOTER_Y, M.SW, H, C["dark"], radius=False)
        M.text(s, Inches(0.5), M.FOOTER_Y, Inches(9), H, "富住通商用不動產　大型工業地產團隊", 12, True, C["white"],
               anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, M.SW - Inches(1.7), M.FOOTER_Y, Inches(1.2), H, str(deck.n), 12, True, C["white"], PP_ALIGN.RIGHT,
               MSO_ANCHOR.MIDDLE)

    def new(title=None):
        s = deck.prs.slides.add_slide(deck.blank)
        deck.n += 1
        bg = s.background.fill
        bg.solid()
        bg.fore_color.rgb = M.BG
        if title:
            M.box(s, Inches(0.5), Inches(0.36), Inches(0.13), Inches(0.58), C["main"], radius=False)
            M.text(s, Inches(0.78), Inches(0.25), Inches(7.4), Inches(0.8), title, 28, True, C["ink"], anchor=MSO_ANCHOR.MIDDLE)
            M.text(s, Inches(8.3), Inches(0.3), M.SW - Inches(8.8), Inches(0.7), f["short"], 13, True, C["ink2"], PP_ALIGN.RIGHT,
                   MSO_ANCHOR.MIDDLE)
            footer(s)
        return s
    deck.new = new
    deck.footer = footer


def _price_text(M, f, d):
    """開價主字串、單價字串"""
    main = f"{f['price']} {f['price_unit']}".strip()
    unit = M._num(d.get("建物單價")) if is_plant(f) else M._num(d.get("土地單價"))
    return main, (f"{unit:g} 萬/坪" if unit else "")


# ---------------------------------------------------------------- 封面：滿版實景＋半透明大圓弧襯標題
def cover_slide(M, deck, C, f, cover, ll, phones, show_agents):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new()
    SW = M.SW
    wp, hp = 2400, int(2400 * M.SH / SW)
    im = _base_image(M, cover, ll, wp, hp)
    if im is None:
        from PIL import Image
        im = Image.new("RGB", (wp, hp), C["dark_rgb"])
    r = wp * 0.31
    im = _veil(im, [("ellipse", (wp / 2 - r, hp * 0.27, wp / 2 + r, hp * 0.27 + r * 2), (28, 28, 30), 205)])
    p = os.path.join(deck.tmp, "ind_cover.jpg")
    im.save(p, quality=90)
    s.shapes.add_picture(p, 0, 0, M.REAL_W, M.REAL_H)
    st = M.box(s, Inches(0.45), Inches(0.4), Inches(1.05), Inches(1.05), C["red"])
    st.adjustments[0] = 0.12
    M.text(s, Inches(0.45), Inches(0.4), Inches(1.05), Inches(1.05), f["deal"], 54, True, C["white"], PP_ALIGN.CENTER,
           MSO_ANCHOR.MIDDLE)
    x, w = Inches(3.1), SW - Inches(6.2)
    title = f["title"] if len(f["title"]) <= 16 else f["short"]
    M.text(s, x, Inches(2.55), w, Inches(0.95), title, 36 if len(title) <= 12 else 28, True, C["white"], PP_ALIGN.CENTER,
           MSO_ANCHOR.MIDDLE)
    M.text(s, x, Inches(3.5), w, Inches(0.8), f"{_kind(f)}{'出租' if f['deal'] == '租' else '銷售'}報告書", 34, True, C["white"],
           PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    agents = [f"{nm}　{phones.get(nm, '')}".strip() for nm in f["agents"][:2]] if show_agents else []
    if agents:
        M.text(s, x, Inches(4.5), w, Inches(0.9), "\n".join(agents), 17, True, C["white"], PP_ALIGN.CENTER, spacing=1.05)
    deck.footer(s)


def _zone(f, d):
    return str(d.get("非都市土地使用地類別") or f.get("zone") or "").replace("(空白) / (空白)", "").strip()


def _heights(d):
    vals = [_val(d.get(k)) for k in ("樑下", "滴水", "中脊")]
    return "／".join(v or "－" for v in vals) + " 米" if any(vals) else ""


def _stats(f, d):
    """基本資料卡：(標題, 值)；沒資料的不列"""
    land = _val(f.get("land"))
    rows = [("土地面積", f"{land} 坪" if land else ""), ("使用分區", _zone(f, d)), ("建蔽／容積", f.get("bcr_far", "")),
            ("臨路寬度", f.get("road_w", "")), ("面寬／深度", f.get("dims_show") or "")]
    if is_plant(f):
        rows += [("建物面積", f"{_val(f.get('area'))} 坪" if _val(f.get("area")) else ""), ("建物結構", f.get("struct", "")),
                 ("屋齡", f.get("age", "")), ("樑下／滴水／中脊", _heights(d)), ("天車", _val(d.get("天車"))),
                 ("工廠登記", _val(d.get("工廠登記"))), ("廠房保存登記", _val(d.get("廠房保存登記")))]
    return [(k, str(v)) for k, v in rows if str(v or "").strip()]


# ---------------------------------------------------------------- 基本資訊：左邊實景＋半透明開價帶，右邊資料卡
def basics(M, deck, C, f, d, photo, ll):
    from pptx.util import Inches
    from pptx.enum.text import MSO_ANCHOR
    s = deck.new("基本資訊")
    x0, y0, w0 = Inches(0.5), Inches(1.2), Inches(6.5)
    h0 = M.FOOTER_Y - y0 - Inches(0.2)
    wp, hp = 1500, int(1500 * h0 / w0)
    im = _base_image(M, photo, ll, wp, hp)
    main, unit = _price_text(M, f, d)
    if im is not None:
        im = _veil(im, [("rect", (-20, hp * 0.8, wp + 20, hp + 20), (24, 24, 26), 195)])
        p = os.path.join(deck.tmp, "ind_basic.jpg")
        im.save(p, quality=90)
        s.shapes.add_picture(p, M._X(x0), M._Y(y0), M._X(w0), M._Y(h0))
    else:
        M.box(s, x0, y0, w0, h0, C["dark"], radius=False)
    by, bh = y0 + int(h0 * 0.8), h0 - int(h0 * 0.8)
    M.text(s, x0 + Inches(0.2), by, Inches(1.1), bh, "月租" if f["deal"] == "租" else "開價", 18, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
    M.text(s, x0 + Inches(1.2), by, w0 - Inches(1.3), bh, main + (f"（{unit}）" if unit else ""), 26, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
    rows = _stats(f, d)[:10]
    x1 = x0 + w0 + Inches(0.25)
    w1 = M.SW - x1 - Inches(0.5)
    n = max(1, math.ceil(len(rows) / 2))
    pts = (f.get("points") or [])[:3]
    box_h = Inches(1.35) if pts else 0
    gap = Inches(0.1)
    ch = min(Inches(1.0), (h0 - box_h - gap * n) // n)
    cw = (w1 - gap) // 2
    for i, (k, v) in enumerate(rows):
        cx = x1 + (cw + gap) * (i % 2)
        cy = y0 + (ch + gap) * (i // 2)
        M.box(s, cx, cy, cw, ch, C["white"], C["line"], radius=False)
        M.box(s, cx, cy, Inches(0.08), ch, C["main"], radius=False)
        M.text(s, cx + Inches(0.15), cy + Inches(0.03), cw - Inches(0.2), Inches(0.34), k, 13, True, C["ink2"])
        M.text(s, cx + Inches(0.15), cy + Inches(0.35), cw - Inches(0.2), ch - Inches(0.38), v, 17 if len(v) <= 11 else 13, True,
               C["ink"], anchor=MSO_ANCHOR.MIDDLE)
    if pts:
        y = y0 + h0 - box_h
        M.box(s, x1, y, w1, box_h, C["panel"], C["line"])
        M.text(s, x1 + Inches(0.12), y + Inches(0.06), w1 - Inches(0.24), box_h - Inches(0.1), "\n".join("• " + p for p in pts),
               14, True, C["ink"])


# ---------------------------------------------------------------- 地籍與現況：行銷地籍圖＋航照
def card_slide(M, deck, C, f, png, ll):
    from PIL import Image
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new("地籍與現況")
    y, h = Inches(1.15), M.FOOTER_Y - Inches(1.25)
    iw, ih = Image.open(png).size
    w = int(h * iw / ih * M.SY / M.SX)         # 設計座標 → 實際頁面比例不同，寬度要換算才不變形
    s.shapes.add_picture(png, M._X(Inches(0.5)), M._Y(y), M._X(w), M._Y(h))
    x = Inches(0.5) + w + Inches(0.3)
    rw = M.SW - x - Inches(0.5)
    if ll:
        wp, hp = 1400, int(1400 * h / rw)
        a = _aerial(M, ll, wp, hp, 18)
        if a is not None:
            a = _veil(a, [("rect", (wp * 0.03, hp * 0.03, wp * 0.36, hp * 0.12), (24, 24, 26), 195)])
            p = os.path.join(deck.tmp, "ind_aerial.jpg")
            a.save(p, quality=90)
            s.shapes.add_picture(p, M._X(x), M._Y(y), M._X(rw), M._Y(h))
            M.text(s, x + int(rw * 0.03), y + int(h * 0.03), int(rw * 0.33), int(h * 0.09), "航照現況", 16, True, C["white"],
                   PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            M.text(s, x, M.FOOTER_Y - Inches(0.1) - Inches(0.34), rw, Inches(0.32), "圖資：內政部國土測繪中心正射影像", 11, True, C["white"])


# ---------------------------------------------------------------- 分割銷售方案：滿版圖＋半透明同色標籤＋自動算差額
def split_calc(M, f, d, rows):
    """回傳 [(區塊, 坪, 單價, 總價, 說明)]、整售單價、整售總價、分售合計(萬)"""
    out = []
    for r in rows:
        a, u = M._num(r.get("area")), M._num(r.get("unit"))
        out.append((str(r.get("name") or ""), a, u, a * u if a and u else None, str(r.get("note") or "")))
    wu, wa = M._num(d.get("土地單價")), M._num(f.get("land"))
    if not wu and wa and M._num(d.get("總價")):
        wu = M._num(d.get("總價")) / wa
    whole = wu * wa if wu and wa else None
    tot = sum(x[3] for x in out) if out and all(x[3] for x in out) else None
    return out, wu, whole, tot


def _wan(v):
    if v >= 10000:
        y, w = int(v // 10000), v % 10000
        return f"{y} 億" + (f" {w:,.0f} 萬" if round(w) else "")
    return f"{v:,.0f} 萬"


def split_slide(M, deck, C, f, d, extra, ll, img):
    from PIL import Image
    from pptx.util import Inches
    from pptx.enum.text import MSO_ANCHOR
    s = deck.new()
    SW, HA = M.SW, M.FOOTER_Y
    wp, hp = 2400, int(2400 * HA / SW)
    if img and os.path.exists(img):
        im = _fit(img, wp, hp)
    else:
        im = (_aerial(M, ll, wp, hp, 18) if ll else None) or Image.new("RGB", (wp, hp), C["dark_rgb"])
    rows, wu, whole, tot = split_calc(M, f, d, extra.get("split") or [])
    px, py = (lambda v: v / SW * wp), (lambda v: v / HA * hp)
    ch = Inches(1.12)
    gap = Inches(0.12)
    cx, cw = Inches(0.4), Inches(4.4)
    ys = [Inches(1.12) + (ch + gap) * i for i in range(len(rows))]
    shp = [("rect", (px(cx), py(Inches(0.25)), px(cx + Inches(3.6)), py(Inches(0.98))), (24, 24, 26), 205)]
    for i, y in enumerate(ys):
        shp.append(("rect", (px(cx), py(y), px(cx + cw), py(y + ch)), _rgb(SPLIT_COLORS[i % 4]), 215))
    sum_y = (ys[-1] + ch + gap) if ys else Inches(1.12)
    show_sum = bool(whole and tot)
    if show_sum:
        shp.append(("rect", (px(cx), py(sum_y), px(cx + cw), py(sum_y + Inches(1.15))), (24, 24, 26), 215))
    im = _veil(im, shp)
    p = os.path.join(deck.tmp, "ind_split.jpg")
    im.save(p, quality=90)
    s.shapes.add_picture(p, 0, 0, M.REAL_W, M._Y(HA))
    M.text(s, cx + Inches(0.15), Inches(0.25), Inches(3.3), Inches(0.73), "分割銷售方案", 28, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
    tx, tw = cx + Inches(0.15), cw - Inches(0.3)
    for (name, a, u, tt, note), y in zip(rows, ys):
        M.text(s, tx, y + Inches(0.05), tw, Inches(0.46), f"{name}　{a:g} 坪" if a else name, 20, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
        line = (f"{u:g} 萬/坪" if u else "") + (f"　約 {_wan(tt)}" if tt else "")
        M.text(s, tx, y + Inches(0.5), tw, Inches(0.34), line, 17, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
        if note:
            M.text(s, tx, y + Inches(0.83), tw, Inches(0.27), note, 12, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
    if show_sum:
        diff = tot - whole
        M.text(s, tx, sum_y + Inches(0.05), tw, Inches(0.36), f"整售　{round(wu, 2):g} 萬/坪　約 {_wan(whole)}", 15, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, tx, sum_y + Inches(0.41), tw, Inches(0.34), f"分割合計　約 {_wan(tot)}", 15, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, tx, sum_y + Inches(0.75), tw, Inches(0.36),
               f"差額　{'+' if diff >= 0 else '-'}{_wan(abs(diff))}（{abs(diff) / whole * 100:.1f}%）", 16, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
    deck.footer(s)


# ---------------------------------------------------------------- 交通位置：大範圍地圖(距離圈 1、3 公里看聯外)＋航照
def location(M, deck, C, f, ll):
    from pptx.util import Inches
    y, h = Inches(1.2), M.FOOTER_Y - Inches(1.32)
    mw = Inches(8.3)
    wp, hp = int(M._X(mw) / 914400 * 200), int(M._Y(h) / 914400 * 200)
    big = M.tile_map(ll[0], ll[1], 14, wp, hp, "EMAP", rings=((1000, "1 公里"), (3000, "3 公里")))
    if big is None:
        M.log("（交通位置：國土測繪圖磚抓不到，略過這頁）")
        return
    s = deck.new("交通位置")
    p = os.path.join(deck.tmp, "ind_map.png")
    big.save(p, optimize=True)
    s.shapes.add_picture(p, M._X(Inches(0.5)), M._Y(y), M._X(mw), M._Y(h))
    rx, rw = Inches(9.0), M.SW - Inches(9.5)
    ah = h - Inches(1.75)
    a = _aerial(M, ll, int(M._X(rw) / 914400 * 200), int(M._Y(ah) / 914400 * 200), 17)
    yy = y
    if a is not None:
        pa = os.path.join(deck.tmp, "ind_loc_a.jpg")
        a.save(pa, quality=90)
        s.shapes.add_picture(pa, M._X(rx), M._Y(yy), M._X(rw), M._Y(ah))
        yy += ah + Inches(0.15)
    tr = []
    for pt in f.get("points") or []:
        m = re.search(r"[^，。；,]*(交流道|國道|快速道路|工業區|園區|港|機場|高鐵|車站)[^，。；,]*", pt)
        if m:
            tr.append(m.group(0).strip())
    rows = [("臨路", f.get("road_w", ""))] + [("交通", x) for x in tr[:3]]
    rows = [(k, v) for k, v in rows if v] or [("位置", f.get("door") or f.get("lots", ""))]
    M.box(s, rx, yy, rw, y + h - yy - Inches(0.4), C["panel"], C["line"])
    M.text(s, rx + Inches(0.1), yy + Inches(0.08), rw - Inches(0.2), y + h - yy - Inches(0.55),
           "\n".join(f"{k}：{v}" for k, v in rows), 15, True, C["ink"])
    M.text(s, rx, y + h - Inches(0.36), rw, Inches(0.34), "圖資：內政部國土測繪中心", 11, False, C["ink2"])


# ---------------------------------------------------------------- 法規與應注意事項
def _law_rows(f, d):
    rows = [("使用分區", _zone(f, d)), ("建蔽／容積", f.get("bcr_far", "")), ("工廠登記", _val(d.get("工廠登記"))),
            ("廠房保存登記", _val(d.get("廠房保存登記")))]
    return [(k, str(v)) for k, v in rows if str(v or "").strip()]


def law_slide(M, deck, C, f, d, extra):
    from pptx.util import Inches
    s = deck.new("法規與應注意事項")
    y = Inches(1.25)
    lw = Inches(5.2)
    each = Inches(1.0)
    for k, v in _law_rows(f, d):
        M.box(s, Inches(0.5), y, lw, each - Inches(0.12), C["white"], C["line"], radius=False)
        M.box(s, Inches(0.5), y, Inches(0.08), each - Inches(0.12), C["main"], radius=False)
        M.text(s, Inches(0.68), y + Inches(0.04), lw - Inches(0.3), Inches(0.32), k, 13, True, C["ink2"])
        M.text(s, Inches(0.68), y + Inches(0.36), lw - Inches(0.3), Inches(0.46), v, 18 if len(v) <= 18 else 14, True, C["ink"])
        y += each
    x = Inches(6.0)
    w = M.SW - x - Inches(0.5)
    h = max(each * len(_law_rows(f, d)) - Inches(0.12), Inches(3.2))
    M.box(s, x, Inches(1.25), w, h, C["panel"], C["line"], radius=False)
    M.text(s, x + Inches(0.2), Inches(1.33), w - Inches(0.4), Inches(0.45), "應注意事項", 20, True, C["dark"])
    notes = [ln.strip() for ln in str(extra.get("notes") or "").splitlines() if ln.strip()]
    M.text(s, x + Inches(0.2), Inches(1.85), w - Inches(0.4), h - Inches(0.75),
           "\n".join("• " + n for n in notes) or "（請洽專案人員確認可進駐行業別、環保與水保規定）", 17, True, C["ink"])


# ---------------------------------------------------------------- 周邊廠商：經濟部生產中工廠清冊
def county_code(d):
    s = str(d.get("土地標示") or d.get("建物門牌") or "")[:4]
    for k, v in COUNTY_CODE.items():
        if s.startswith(k):
            return v
    return None


def nearby(d, ll, r=2000):
    code = county_code(d)
    if not code:
        return None
    req = urllib.request.Request(FACTORY_URL.format(code), headers={"User-Agent": "Mozilla/5.0", "Referer": "https://map.windsky-sky.com/"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        geo = json.loads(resp.read().decode("utf-8"))
    lat, lng = ll
    cos = math.cos(math.radians(lat))
    out = []
    for ft in geo.get("features") or []:
        try:
            x, y = ft["geometry"]["coordinates"][:2]
        except Exception:
            continue
        dist = math.hypot((x - lng) * 111320 * cos, (y - lat) * 110540)
        if dist <= r:
            pr = ft.get("properties") or {}
            cat = re.split(r"[、,，]", str(pr.get("c", "")))[0]          # 清冊可能列多個產業，取第一個
            out.append((round(dist), pr.get("n", ""), re.sub(r"^\d+", "", cat).strip()))
    out.sort()
    return out


def neighbors_slide(M, deck, C, f, items):
    from collections import Counter
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new("周邊廠商")
    M.text(s, Inches(0.5), Inches(1.13), Inches(12), Inches(0.45), f"半徑 2 公里內共 {len(items)} 家登記工廠", 18, True, C["dark"])
    top = Counter(c for _, _, c in items if c).most_common(6)
    y0, x0 = Inches(1.7), Inches(0.5)
    M.text(s, x0, y0, Inches(5.5), Inches(0.4), "主要產業", 16, True, C["ink"])
    y = y0 + Inches(0.45)
    mx = top[0][1] if top else 1
    bw = Inches(2.1)
    for name, n in top:
        M.text(s, x0, y, Inches(3.0), Inches(0.5), name.replace("製造業", ""), 14, True, C["ink"], anchor=MSO_ANCHOR.MIDDLE)
        M.box(s, x0 + Inches(3.0), y + Inches(0.1), max(Inches(0.1), bw * n // mx), Inches(0.3), C["main"], radius=False)
        M.text(s, x0 + Inches(3.05) + bw * n // mx, y, Inches(0.7), Inches(0.5), str(n), 14, True, C["ink"], anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(0.58)
    x1 = Inches(6.7)
    w1 = M.SW - x1 - Inches(0.5)
    M.text(s, x1, y0, w1, Inches(0.4), "最近的工廠", 16, True, C["ink"])
    y = y0 + Inches(0.45)
    for dist, name, cat in items[:7]:
        M.box(s, x1, y, w1, Inches(0.5), C["white"], C["line"], radius=False)
        M.text(s, x1 + Inches(0.08), y, Inches(3.5), Inches(0.5), name, 13, True, C["ink"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, x1 + Inches(3.6), y, w1 - Inches(4.55), Inches(0.5), cat.replace("製造業", ""), 11, False, C["ink2"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, x1 + w1 - Inches(0.95), y, Inches(0.9), Inches(0.5), f"{dist} 米", 12, True, C["dark"], PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
        y += Inches(0.56)
    M.text(s, Inches(0.5), M.FOOTER_Y - Inches(0.36), Inches(12), Inches(0.32), "資料來源：經濟部生產中工廠清冊（距離依登記地址定位推估）", 11, False, C["ink2"])


# ---------------------------------------------------------------- 專案團隊(公司通用介紹＋本案承辦人)
def team(M, deck, C, f, phones, show_agents):
    from PIL import Image
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new("專案團隊")
    lv = M.asset(M.ASSET_LOGO_V)
    if lv:
        iw, ih = Image.open(lv).size
        h = Inches(2.6)
        w = int(h * iw / ih * M.SY / M.SX)
        s.shapes.add_picture(lv, M._X(Inches(2.7) - w // 2), M._Y(Inches(1.3)), M._X(w), M._Y(h))
    M.text(s, Inches(0.4), Inches(4.05), Inches(4.6), Inches(0.5), M.COMPANY, 18, True, C["ink"], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    M.text(s, Inches(0.4), Inches(4.55), Inches(4.6), Inches(0.5), M.COMPANY_TEL, 20, True, C["dark"], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    x = Inches(5.5)
    w = M.SW - x - Inches(0.5)
    M.box(s, x, Inches(1.25), w, Inches(2.65), C["panel"], C["line"])
    M.text(s, x + Inches(0.2), Inches(1.32), w - Inches(0.4), Inches(0.42), "免費諮詢服務", 18, True, C["dark"])
    M.text(s, x + Inches(0.2), Inches(1.76), w - Inches(0.4), Inches(2.1), "\n".join("• " + v for v in SERVICES), 16, True, C["ink"])
    agents = [(nm, phones.get(nm, "")) for nm in f["agents"][:3]] if show_agents else []
    y = Inches(4.05)
    if agents:
        M.text(s, x, y, w, Inches(0.42), "本案承辦", 18, True, C["dark"])
    for i, (nm, ph) in enumerate(agents):
        yy = y + Inches(0.48) + Inches(0.6) * i
        M.text(s, x, yy, Inches(2.5), Inches(0.55), nm, 22, True, C["ink"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, x + Inches(2.5), yy, w - Inches(2.5), Inches(0.55), ph, 22, True, C["dark"], anchor=MSO_ANCHOR.MIDDLE)
