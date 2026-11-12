"""行銷簡報「土地・建地」「農地」版型：沒有建物，重點是地籍、使用分區、面寬臨路、能蓋多少(建地)或農舍條件(農地)。
滿版空拍/航照＋半透明圖層的手法沿用工業版(★半透明一律 PIL 合成進圖片，iPad/手機預覽不漏圖)；
使用分區、周邊機能頁沿用透天版。利多消息/現況說明、分割方案存在案件 4.其他相關/土地版補充.json(「頁面內容」步驟按「填寫」)。
共用 marketing_deck 的畫圖工具(M=marketing_deck 模組本身)。"""
import json
import os

import marketing_deck_factory as mf
import marketing_deck_street as ms

EXTRA_FILE = "土地版補充.json"
PING_M2 = 3.30579
FARM_HOUSE = ("申請人須為農民，且無自用住宅", "土地取得滿 2 年（民國 89 年後取得的農地）", "農業用地面積至少 0.25 公頃",
              "農舍建築面積不超過土地面積 10%", "總樓地板面積不超過 495 平方公尺", "高度不超過 3 層樓、10.5 公尺")


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


def notes_label(farm):
    return "現況說明（一行一點，例：有灌溉水圳、現種植芒果、可通行農路）" if farm else "利多消息（一行一點，例：重劃區、捷運站 500 公尺、大型建案林立）"


def _kind(f, farm):
    return "農地" if farm else ("建地" if f.get("zone") else "土地")


def _is_farmland(f, d):
    return "農牧" in str(d.get("非都市土地使用地類別") or "")


def build(M, case_dir, base_dir, d, exp, chosen, out_path, f, show_agents=True, sel=None, farm=False):
    from pptx.dml.color import RGBColor
    t = M.THEMES[M.THEME]
    C = {"main": RGBColor.from_string(t["sky"]), "dark": RGBColor.from_string(t["sky_dk"]),
         "dark_rgb": mf._rgb(t["sky_dk"]), "white": M.WHITE, "panel": M.PANEL, "line": M.LINE,
         "ink": RGBColor.from_string("1F2933"), "ink2": RGBColor.from_string("3F4A56"), "red": M.STAMP_RED}
    # 透天版的頁面(使用分區、周邊機能)用的顏色名稱
    C.update(char=C["ink"], char2=C["ink2"], amber=C["main"], amber_dk=C["dark"], grey=RGBColor.from_string("718096"))
    M.INK = C["ink"]
    tmp = M.tempfile.mkdtemp(prefix="deck_")
    deck = M.Deck(f, tmp)
    deck.frame = None
    _chrome(M, deck, C, f)
    extra = load_extra(case_dir)
    by = {c: M.ordered([p for p in chosen if p["cat"] == c]) for c in M.CATS}
    cover = (by["封面"] or by["空拍"] or by["基地現況"] or by[M.EXT_CAT] or by[M.STREET_CAT] or [None])[0]
    side = (by["空拍"] or by["基地現況"] or by[M.EXT_CAT] or by[M.STREET_CAT] or [cover])[0]
    roots = [r for r in (case_dir, exp) if r]
    ll = mf._coords(M, case_dir, base_dir, d)
    cads = M.sel_paths(sel, "cadastral", [c for c in M._globs(roots, os.path.join("0.謄本", "*地籍圖*標註*.pdf"),
                                                              os.path.join("0.謄本", "*地籍圖*.pdf"))][:1])
    cards = M.sel_paths(sel, "card", M._globs(roots, os.path.join("0.謄本", "*_行銷*.png"))[:1])
    f["dims_show"] = f["dims"]
    if not f["dims_show"] and cads:
        try:
            dm = M.parcel_dims(cads[0])
        except Exception:
            dm = None
        if dm:
            f["dims_show"] = f"約 {dm['frontage']:g}／{dm['depth']:g} 米"
    kind = _kind(f, farm)
    phones = M.lookup_phones(f["agents"], base_dir)
    M.log(f"製作：封面（{kind}）"); cover_slide(M, deck, C, f, cover, ll, phones, show_agents, kind)
    M.log("製作：基本資訊"); basics(M, deck, C, f, d, side, ll, kind)
    if cards:
        M.log("製作：地籍與現況"); mf.card_slide(M, deck, C, f, cards[0], ll)
    else:
        for src in cads:
            M.log("製作：地籍圖"); M.slide_cadastral(deck, src)
    zon = M.sel_paths(sel, "zoning", M._globs(roots, os.path.join("1.基本資料", "png", "*都市計畫圖*.png"),
                                              os.path.join("1.基本資料", "*全國土地使用分區*.pdf"))[:1])
    for src in zon:
        M.log("製作：使用分區"); ms.zoning(M, deck, C, f, src)
    if M.sel_on(sel, "develop"):
        if farm:
            M.log("製作：農地資訊"); farm_slide(M, deck, C, f, d)
        elif _num_land(M, f) and M._num(d.get("容積率")):
            M.log("製作：開發試算"); develop_slide(M, deck, C, f, d)
    if M.sel_on(sel, "split") and extra.get("split"):
        M.log("製作：分割銷售方案")
        mf.split_slide(M, deck, C, f, d, extra, ll, (M.sel_paths(sel, "split", []) or [None])[0])
    if ll and M.sel_on(sel, "location"):
        M.log("製作：地理位置"); M.slide_location(deck, *ll)
    for src in M.sel_paths(sel, "facility", M._globs(roots, os.path.join("1.基本資料", "*重要設施*.pdf"))[:1]):
        try:
            M.log("製作：周邊機能（重要設施）"); ms.facilities(M, deck, C, f, src)
        except Exception as e:
            M.log(f"（周邊機能頁略過：{e}）")
    notes = [ln.strip() for ln in str(extra.get("notes") or "").splitlines() if ln.strip()]
    if notes and M.sel_on(sel, "notes"):
        M.log("製作：" + ("現況說明" if farm else "利多消息")); notes_slide(M, deck, C, f, notes, farm)
    M.log("製作：照片頁")
    for title, cats, cap in M.LAND_PHOTO_PAGES:
        items = M.ordered([p for p in chosen if p["cat"] in cats])
        if items and cover in items and len(items) > 1:
            items = [p for p in items if p is not cover] + [cover]
        start = 0
        for k, n in enumerate(M.split_pages(len(items), cap)):
            M.slide_photos(deck, title + ("（續）" if k else ""), items[start:start + n])
            start += n
    M.market_slides(deck, f, M.sel_paths(sel, "market", M.market_default(roots, f["deal"] == "租")), tmp)
    M.log("製作：聯絡資訊"); contact(M, deck, C, f, phones, show_agents, kind)
    M.unique_shape_ids(deck.prs)
    deck.prs.save(out_path)
    return deck.n


def _num_land(M, f):
    return M._num(f.get("land"))


# ---------------------------------------------------------------- 頁框：主色短條＋標題；頁尾深色帶(公司名、頁碼)
def _chrome(M, deck, C, f):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    H = M.SH - M.FOOTER_Y

    def footer(s):
        M.box(s, 0, M.FOOTER_Y, M.SW, H, C["dark"], radius=False)
        M.text(s, Inches(0.5), M.FOOTER_Y, Inches(9), H, f"{M.COMPANY}　{M.COMPANY_TEL}", 12, True, C["white"],
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
    main = f"{f['price']} {f['price_unit']}".strip()
    unit = M._num(d.get("土地單價"))
    return main, (f"{unit:g} 萬/坪" if unit else "")


# ---------------------------------------------------------------- 封面：四種樣式給使用者選(步驟① 選封面樣式)
COVER_STYLES = {"A": ("底部漸層", "照片幾乎全露，最下面漸暗放字"),
                "B": ("左側直欄", "左邊色塊寫資料，字最清楚"),
                "C": ("上照片下白卡", "照片不遮，資料放下方白底"),
                "D": ("左下小卡", "照片完整，資料放左下白卡")}
PHOTO_H = 0.68                 # C 樣式照片佔上面 68%


def cover_background(im, style, w, h, accent):
    """封面底圖(PIL)：照片＋該樣式的遮罩/色塊，半透明都合成進圖片(iPad/手機不漏圖)。im=照片(已裁成 w×h)或 None"""
    from PIL import Image, ImageDraw, ImageOps
    if im is None:
        im = Image.new("RGB", (w, h), accent)
    if style == "C":
        out = Image.new("RGB", (w, h), "white")
        ph = int(h * PHOTO_H)
        out.paste(ImageOps.fit(im, (w, ph), Image.LANCZOS), (0, 0))
        ImageDraw.Draw(out).rectangle((0, ph, w, ph + max(4, h // 110)), fill=accent)
        return out
    lay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(lay)
    if style == "B":
        d.rectangle((0, 0, int(w * 0.38), h), fill=tuple(accent) + (225,))
    elif style == "D":
        x0, y0, x1, y1 = w * 0.0375, h * 0.545, w * 0.49, h * 0.915
        d.rounded_rectangle((x0, y0, x1, y1), radius=int(h * 0.025), fill=(255, 255, 255, 235))
        d.rectangle((x0, y0 + h * 0.035, x0 + max(4, w * 0.006), y1 - h * 0.035), fill=tuple(accent) + (255,))
    else:                                       # A：下半部由透明漸漸變暗
        top = int(h * 0.5)
        for y in range(top, h):
            d.line((0, y, w, y), fill=(10, 18, 12, int(215 * ((y - top) / (h - top)) ** 0.8)))
    return Image.alpha_composite(im.convert("RGBA"), lay).convert("RGB")


def cover_lines(f, kind, phones, show_agents):
    try:
        lv = float(str(f.get("land") or "").replace(",", "").replace("坪", "").strip() or 0)
    except ValueError:
        lv = 0
    land = f"土地 {lv:g} 坪" if lv else ""
    title = f["title"] if len(f["title"]) <= 16 else f["short"]
    return dict(title=title, kind=f"{kind}{'出租' if f['deal'] == '租' else '銷售'}報告書", land=land,
                zone=f.get("zone") or "", agents=[f"{nm}　{phones.get(nm, '')}".strip() for nm in f["agents"][:2]] if show_agents else [])


# 各樣式文字位置(佔頁面寬高的比例)：(x, y, w, h, 字級, 粗體, 顏色鍵, 對齊)
_LAYOUT = {
    "A": [("title", .045, .63, .62, .11, 40, "w", "l"), ("sub", .047, .745, .62, .06, 18, "w", "l"),
          ("agents", .58, .70, .375, .12, 16, "w", "r")],
    "B": [("title", .04, .22, .32, .12, 32, "w", "l"), ("kind", .04, .37, .32, .07, 22, "w", "l"),
          ("land", .04, .45, .32, .055, 17, "w", "l"), ("zone", .04, .505, .32, .055, 17, "w", "l"),
          ("agents", .04, .64, .32, .12, 16, "w", "l")],
    "C": [("title", .045, .705, .58, .1, 34, "ink", "l"), ("sub", .047, .81, .58, .06, 16, "ink2", "l"),
          ("agents", .6, .71, .355, .12, 16, "accent", "r")],
    "D": [("title", .065, .565, .41, .09, 28, "ink", "l"), ("kind", .066, .655, .41, .06, 18, "accent", "l"),
          ("sub2", .066, .715, .41, .05, 14, "ink2", "l"), ("agents", .066, .775, .41, .11, 15, "ink", "l")],
}


def _texts(L, key):
    sub = "　｜　".join(x for x in (L["kind"], "　".join(y for y in (L["land"], L["zone"]) if y)) if x)
    return {"title": L["title"], "kind": L["kind"], "land": L["land"], "zone": L["zone"], "sub": sub,
            "sub2": "　".join(y for y in (L["land"], L["zone"]) if y), "agents": "\n".join(L["agents"])}[key]


def cover_preview(photo_path, style, f, kind, accent_hex, w=320, h=180):
    """步驟①封面樣式的縮圖(PIL)：用這個案件的封面照片，字的位置跟正式簡報一樣"""
    from PIL import Image, ImageDraw, ImageFont, ImageOps
    acc = tuple(int(accent_hex[i:i + 2], 16) for i in (0, 2, 4))
    im = None
    if photo_path:
        try:
            src = Image.open(photo_path)
            src.draft("RGB", (w * 2, h * 2))                    # JPEG 先縮小解碼，縮圖秒出
            im = ImageOps.fit(ImageOps.exif_transpose(src).convert("RGB"), (w * 2, h * 2), Image.LANCZOS)
        except Exception:
            im = None
    W2, H2 = w * 2, h * 2
    bg = cover_background(im, style, W2, H2, acc)
    dr = ImageDraw.Draw(bg)
    L = cover_lines(f, kind, {}, True)
    L["agents"] = [a.split("　")[0] + "　09xx-xxx-xxx" for a in L["agents"]] or ["經紀人　09xx-xxx-xxx"]
    col = {"w": (255, 255, 255), "ink": (31, 41, 51), "ink2": (63, 74, 86), "accent": acc}
    try:
        font = lambda sz: ImageFont.truetype("C:/Windows/Fonts/msjhbd.ttc", sz)
        font(10)
    except Exception:
        font = lambda sz: ImageFont.load_default()
    for key, x, y, bw, bh, sz, ck, al in _LAYOUT.get(style, _LAYOUT["A"]):
        t = _texts(L, key)
        if not t:
            continue
        px = max(9, int(sz * H2 / 540 * 1.0))
        X = W2 * (x + bw) if al == "r" else W2 * x
        dr.multiline_text((X, H2 * y), t, font=font(px), fill=col[ck], anchor="ra" if al == "r" else "la",
                          align="right" if al == "r" else "left", spacing=4)
    dr.rounded_rectangle((W2 * .02, H2 * .035, W2 * .075, H2 * .13), radius=6, fill=(200, 40, 40))     # 售
    dr.rectangle((0, H2 - max(6, H2 // 26), W2, H2), fill=acc)                                          # 頁尾
    return bg.resize((w, h), Image.LANCZOS)


def cover_slide(M, deck, C, f, cover, ll, phones, show_agents, kind):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    style = getattr(M, "COVER_STYLE", "A")
    style = style if style in COVER_STYLES else "A"
    s = deck.new()
    wp, hp = 2400, int(2400 * M.SH / M.SW)
    acc = C["dark_rgb"]
    im = mf._base_image(M, cover, ll, wp, hp)
    im = cover_background(im, style, wp, hp, acc)
    p = os.path.join(deck.tmp, "land_cover.jpg")
    im.save(p, quality=90)
    s.shapes.add_picture(p, 0, 0, M.REAL_W, M.REAL_H)
    st = M.box(s, Inches(0.45), Inches(0.4), Inches(1.05), Inches(1.05), C["red"])
    st.adjustments[0] = 0.12
    M.text(s, Inches(0.45), Inches(0.4), Inches(1.05), Inches(1.05), f["deal"], 54, True, C["white"], PP_ALIGN.CENTER,
           MSO_ANCHOR.MIDDLE)
    L = cover_lines(f, kind, phones, show_agents)
    col = {"w": C["white"], "ink": C["ink"], "ink2": C["ink2"], "accent": C["dark"]}
    for key, x, y, bw, bh, sz, ck, al in _LAYOUT[style]:
        t = _texts(L, key)
        if not t:
            continue
        if key == "title" and len(t) > 12:
            sz = int(sz * 0.8)
        M.text(s, int(M.SW * x), int(M.SH * y), int(M.SW * bw), int(M.SH * bh), t, sz, True, col[ck],
               PP_ALIGN.RIGHT if al == "r" else PP_ALIGN.LEFT, MSO_ANCHOR.TOP, spacing=1.05)
    if style == "B":                            # 標題下的金色短線
        M.box(s, int(M.SW * .04), int(M.SH * .345), int(M.SW * .075), int(M.SH * .008), M.RGBColor.from_string("FFC850"))
    deck.footer(s)


def _stats(M, f, d, farm):
    """基本資料卡：(標題, 值)；沒資料的不列"""
    land = M._num(f.get("land"))
    area = f"{land:g} 坪" + (f"（約 {land * PING_M2 / 10000:.3f} 公頃）" if farm else f"（{land * PING_M2:,.2f} ㎡）") if land else ""
    cur = M._num(d.get("公告現值"))
    rows = [("土地面積", area), ("使用分區", mf._zone(f, d)), ("建蔽／容積", f.get("bcr_far", "")),
            ("臨路寬度", f.get("road_w", "")), ("面寬／深度", f.get("dims_show") or ""),
            ("公告現值", f"{cur:,.0f} 元/㎡（約 {cur * PING_M2 / 10000:.2f} 萬/坪）" if cur else ""),
            ("地段地號", f.get("lots", ""))]
    return [(k, str(v)) for k, v in rows if str(v or "").strip()]


# ---------------------------------------------------------------- 基本資訊：左邊實景＋半透明開價帶，右邊資料卡
def basics(M, deck, C, f, d, photo, ll, kind):
    import math
    from pptx.util import Inches
    from pptx.enum.text import MSO_ANCHOR
    s = deck.new("基本資訊")
    x0, y0, w0 = Inches(0.5), Inches(1.2), Inches(6.5)
    h0 = M.FOOTER_Y - y0 - Inches(0.2)
    wp, hp = 1500, int(1500 * h0 / w0)
    im = mf._base_image(M, photo, ll, wp, hp)
    main, unit = _price_text(M, f, d)
    if im is not None:
        im = mf._veil(im, [("rect", (-20, hp * 0.8, wp + 20, hp + 20), (22, 26, 24), 195)])
        p = os.path.join(deck.tmp, "land_basic.jpg")
        im.save(p, quality=90)
        s.shapes.add_picture(p, M._X(x0), M._Y(y0), M._X(w0), M._Y(h0))
    else:
        M.box(s, x0, y0, w0, h0, C["dark"], radius=False)
    by, bh = y0 + int(h0 * 0.8), h0 - int(h0 * 0.8)
    M.text(s, x0 + Inches(0.2), by, Inches(1.1), bh, "月租" if f["deal"] == "租" else "開價", 18, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
    M.text(s, x0 + Inches(1.2), by, w0 - Inches(1.3), bh, main + (f"（{unit}）" if unit else ""), 26, True, C["white"], anchor=MSO_ANCHOR.MIDDLE)
    rows = _stats(M, f, d, kind == "農地")[:8]
    x1 = x0 + w0 + Inches(0.25)
    w1 = M.SW - x1 - Inches(0.5)
    pts = (f.get("points") or [])[:3]
    box_h = Inches(1.35) if pts else 0
    gap = Inches(0.1)
    # 長的(地段地號、公告現值、面積)整列，短的兩欄
    wide = [r for r in rows if len(r[1]) > 12]
    narrow = [r for r in rows if len(r[1]) <= 12]
    n_lines = len(wide) + math.ceil(len(narrow) / 2)
    ch = min(Inches(0.95), (h0 - box_h - gap * max(1, n_lines)) // max(1, n_lines))
    cw = (w1 - gap) // 2
    y = y0
    cells = [(r, True) for r in wide] + [(r, False) for r in narrow]
    col = 0
    for (k, v), full in cells:
        if full:
            cx, ww = x1, w1
        else:
            cx, ww = x1 + (cw + gap) * col, cw
        M.box(s, cx, y, ww, ch, C["white"], C["line"], radius=False)
        M.box(s, cx, y, Inches(0.08), ch, C["main"], radius=False)
        M.text(s, cx + Inches(0.15), y + Inches(0.03), ww - Inches(0.2), Inches(0.32), k, 13, True, C["ink2"])
        M.text(s, cx + Inches(0.15), y + Inches(0.33), ww - Inches(0.2), ch - Inches(0.36), v, 16 if len(v) <= 22 else 13, True,
               C["ink"], anchor=MSO_ANCHOR.MIDDLE)
        if full or col == 1:
            y += ch + gap
            col = 0
        else:
            col = 1
    if pts:
        y = y0 + h0 - box_h
        M.box(s, x1, y, w1, box_h, C["panel"], C["line"])
        M.text(s, x1 + Inches(0.12), y + Inches(0.06), w1 - Inches(0.24), box_h - Inches(0.1), "\n".join("• " + p for p in pts),
               14, True, C["ink"])


# ---------------------------------------------------------------- 開發試算(建地)：能蓋多少、每坪容積攤到多少地價
def develop_calc(M, f, d):
    land = M._num(f.get("land"))
    bcr, far = M._num(d.get("建蔽率")), M._num(d.get("容積率"))
    unit = M._num(d.get("土地單價"))
    total = M._num(d.get("總價"))
    if not unit and total and land:
        unit = total / land
    rows = [("土地面積", f"{land:g} 坪")]
    if bcr:
        rows.append(("建築面積上限", f"約 {land * bcr / 100:,.1f} 坪（建蔽率 {bcr:g}%）"))
    rows.append(("基準容積樓地板", f"約 {land * far / 100:,.1f} 坪（容積率 {far:g}%）"))
    big = []
    if unit:
        big.append(("土地單價", f"{unit:,.2f}".rstrip("0").rstrip("."), "萬/坪"))
        big.append(("每坪容積地價", f"{unit / (far / 100):,.2f}".rstrip("0").rstrip("."), "萬/坪"))
    big.insert(0, ("可建樓地板(基準容積)", f"{land * far / 100:,.0f}", "坪"))
    return rows, big


def develop_slide(M, deck, C, f, d):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new("開發試算")
    rows, big = develop_calc(M, f, d)
    x0, y0 = Inches(0.5), Inches(1.3)
    gap = Inches(0.25)
    bw = (M.SW - Inches(1.0) - gap * (len(big) - 1)) // len(big)
    bh = Inches(1.9)
    for i, (k, v, u) in enumerate(big):
        x = x0 + (bw + gap) * i
        M.box(s, x, y0, bw, bh, C["panel"] if i else C["dark"], C["line"])
        M.text(s, x, y0 + Inches(0.15), bw, Inches(0.45), k, 16, True, C["white"] if not i else C["ink2"], PP_ALIGN.CENTER)
        tb = M.text(s, x, y0 + Inches(0.65), bw, Inches(1.0), v, 40 if len(v) <= 7 else 30, True,
                    C["white"] if not i else C["dark"], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        r = tb.text_frame.paragraphs[0].add_run()
        r.text = f" {u}"
        M._font(r, 18, True, C["white"] if not i else C["ink2"])
    y = y0 + bh + Inches(0.35)
    for k, v in rows:
        M.box(s, x0, y, M.SW - Inches(1.0), Inches(0.62), C["white"], C["line"], radius=False)
        M.box(s, x0, y, Inches(0.08), Inches(0.62), C["main"], radius=False)
        M.text(s, x0 + Inches(0.25), y, Inches(3.0), Inches(0.62), k, 16, True, C["ink2"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, x0 + Inches(3.3), y, M.SW - Inches(4.3), Inches(0.62), v, 18, True, C["ink"], anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(0.72)
    M.text(s, x0, M.FOOTER_Y - Inches(0.75), M.SW - Inches(1.0), Inches(0.6),
           "※ 以基準容積概算，未含容積獎勵、移轉及退縮、法定空地等規定；實際可建量以建築師規劃、主管機關核定為準。",
           13, True, C["ink2"], anchor=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------- 農地資訊：使用地類別、面積(公頃)、農舍興建條件
def farm_slide(M, deck, C, f, d):
    from pptx.util import Inches
    from pptx.enum.text import MSO_ANCHOR
    s = deck.new("農地資訊")
    land = M._num(f.get("land"))
    ha = land * PING_M2 / 10000 if land else 0
    x0, y0 = Inches(0.5), Inches(1.3)
    lw = Inches(5.4)
    rows = [("使用地類別", str(d.get("非都市土地使用地類別") or "").replace("(空白) / (空白)", "").strip() or mf._zone(f, d)),
            ("土地面積", f"{land:g} 坪（約 {ha:.3f} 公頃、{land * PING_M2:,.0f} ㎡）" if land else ""),
            ("臨路寬度", f.get("road_w", "")), ("面寬／深度", f.get("dims_show") or "")]
    y = y0
    for k, v in [r for r in rows if r[1]]:
        M.box(s, x0, y, lw, Inches(0.95), C["white"], C["line"], radius=False)
        M.box(s, x0, y, Inches(0.08), Inches(0.95), C["main"], radius=False)
        M.text(s, x0 + Inches(0.2), y + Inches(0.05), lw - Inches(0.3), Inches(0.32), k, 13, True, C["ink2"])
        M.text(s, x0 + Inches(0.2), y + Inches(0.38), lw - Inches(0.3), Inches(0.5), v, 17 if len(v) <= 20 else 13, True, C["ink"],
               anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(1.07)
    x = x0 + lw + Inches(0.3)
    w = M.SW - x - Inches(0.5)
    h = M.FOOTER_Y - y0 - Inches(0.25)
    M.box(s, x, y0, w, h, C["panel"], C["line"], radius=False)
    if _is_farmland(f, d):
        M.text(s, x + Inches(0.25), y0 + Inches(0.12), w - Inches(0.5), Inches(0.45), "農舍興建條件（一般規定）", 20, True, C["dark"])
        ok = ha >= 0.25
        M.text(s, x + Inches(0.25), y0 + Inches(0.62), w - Inches(0.5), Inches(0.45),
               ("✓ 本筆面積已達 0.25 公頃" if ok else "✗ 本筆面積未達 0.25 公頃，不能單獨申請農舍") if land else "",
               16, True, C["main"] if ok else C["red"])
        M.text(s, x + Inches(0.25), y0 + Inches(1.15), w - Inches(0.5), h - Inches(1.9), "\n".join("• " + v for v in FARM_HOUSE),
               16, True, C["ink"], spacing=1.15)
        M.text(s, x + Inches(0.25), y0 + h - Inches(0.7), w - Inches(0.5), Inches(0.6),
               "依農業用地興建農舍辦法，實際以申請人資格與主管機關審查為準。", 12, True, C["ink2"], anchor=MSO_ANCHOR.MIDDLE)
    else:
        M.text(s, x + Inches(0.25), y0 + Inches(0.12), w - Inches(0.5), Inches(0.45), "使用說明", 20, True, C["dark"])
        M.text(s, x + Inches(0.25), y0 + Inches(0.7), w - Inches(0.5), h - Inches(1.0),
               "此筆不是非都市「農牧用地」，可作的用途與興建條件請洽專案人員確認。", 16, True, C["ink"])


# ---------------------------------------------------------------- 利多消息／現況說明
def notes_slide(M, deck, C, f, notes, farm):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new("現況說明" if farm else "利多消息")
    notes = notes[:8]
    cols = 2 if len(notes) > 4 else 1
    rows = (len(notes) + cols - 1) // cols
    x0, y0 = Inches(0.5), Inches(1.3)
    gap = Inches(0.2)
    cw = (M.SW - Inches(1.0) - gap * (cols - 1)) // cols
    ch = min(Inches(1.15), (M.FOOTER_Y - y0 - Inches(0.3) - gap * (rows - 1)) // rows)
    for i, n in enumerate(notes):
        x = x0 + (cw + gap) * (i // rows)
        y = y0 + (ch + gap) * (i % rows)
        M.box(s, x, y, cw, ch, C["white"], C["line"])
        M.box(s, x + Inches(0.2), y + (ch - Inches(0.55)) // 2, Inches(0.55), Inches(0.55), C["main"])
        M.text(s, x + Inches(0.2), y + (ch - Inches(0.55)) // 2, Inches(0.55), Inches(0.55), str(i + 1), 18, True, C["white"],
               PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        M.text(s, x + Inches(0.95), y, cw - Inches(1.15), ch, n, 18 if len(n) <= 24 else 14, True, C["ink"], anchor=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------- 聯絡頁
def contact(M, deck, C, f, phones, show_agents, kind):
    from PIL import Image
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new("專案人員")
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
    M.text(s, x, Inches(1.3), w, Inches(0.45), f"{kind}{'出租' if f['deal'] == '租' else '出售'}", 16, True, C["main"])
    M.text(s, x, Inches(1.75), w, Inches(1.0), "歡迎預約看地", 40, True, C["ink"])
    M.text(s, x, Inches(2.75), w, Inches(0.5), f.get("lots") or f["title"], 15, True, C["ink2"])
    agents = [(nm, phones.get(nm, "")) for nm in f["agents"][:3]] if show_agents else []
    for i, (nm, ph) in enumerate(agents):
        y = Inches(3.5) + Inches(0.75) * i
        M.text(s, x, y, Inches(2.4), Inches(0.65), nm, 24, True, C["ink"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, x + Inches(2.4), y, w - Inches(2.4), Inches(0.65), ph, 24, True, C["dark"], anchor=MSO_ANCHOR.MIDDLE)
    tb = M.text(s, x, M.FOOTER_Y - Inches(1.0), w, Inches(0.8), f"{f['price_label']}　", 18, True, C["ink2"], anchor=MSO_ANCHOR.MIDDLE)
    r = tb.text_frame.paragraphs[0].add_run()
    r.text = f"{f['price']} {f['price_unit']}"
    M._font(r, 30, True, C["dark"])
