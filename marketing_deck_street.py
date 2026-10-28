"""行銷簡報「透天・店面」版型：跟大樓版明顯不同(不用藍粉外框)，但保留富住通 LOGO、售/租章、封底。
★淺色省碳粉：大面積都是白/淡底，顏色只用在細線、小標籤與數字(使用者 10-21：深色底列印很傷碳粉)。
半頁照片封面＋大字租金/開價、物件亮點(面寬/臨路/坪數/樓層)、樓層配置、周邊機能(1.基本資料 重要設施)、
使用分區(都市計畫截圖)、地籍圖、成果圖、格局圖、照片(外觀街景排前面)、行情、聯絡頁。
共用 marketing_deck 的畫圖工具(M=marketing_deck 模組本身，避免 __main__ 重複載入)。"""
import os
import re
import glob


def build(M, case_dir, base_dir, d, exp, chosen, out_path, f, show_agents=True, sel=None):
    from pptx.dml.color import RGBColor
    t = M.THEMES[M.THEME]
    # amber=主色(細線/標籤/小方塊)、amber_dk=深主色(大數字)；名稱沿用，顏色跟著使用者選的色系
    C = {
        "char": RGBColor.from_string("2D3748"), "char2": RGBColor.from_string("4A5568"),
        "amber": RGBColor.from_string(t["sky"]), "amber_dk": RGBColor.from_string(t["sky_dk"]),
        "grey": RGBColor.from_string("718096"), "light": RGBColor.from_string("A0AEC0"),
        "white": M.WHITE, "panel": M.PANEL, "line": M.LINE, "red": M.STAMP_RED,
    }
    if M.THEME == "street":
        C["amber"] = RGBColor.from_string("D97706")
        M.GOLD_HI, M.GOLD_LO = RGBColor.from_string("FBBF24"), RGBColor.from_string("D97706")   # 金牌改琥珀
    M.INK = C["char"]
    tmp = M.tempfile.mkdtemp(prefix="deck_")
    C["logo"] = _dark_logo(M, tmp)
    deck = M.Deck(f, tmp)
    deck.frame = None
    _street_chrome(M, deck, C, f)
    by = {c: M.ordered([p for p in chosen if p["cat"] == c]) for c in M.CATS}
    cover = (by["封面"] or by[M.EXT_CAT] or by[M.STREET_CAT] or by["陽台景觀"] or by["客廳"] or [None])[0]
    roots = [r for r in (case_dir, exp) if r]

    def first(pattern):
        for r in roots:
            hits = sorted(glob.glob(os.path.join(r, pattern)))
            if hits:
                return hits[0]
        return None

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
    plans = M.sel_paths(sel, "plan", plans)
    cad_pdf = first(os.path.join("0.謄本", "*地籍圖*標註*.pdf")) or first(os.path.join("0.謄本", "*地籍圖*.pdf"))
    cad_pages = M.sel_paths(sel, "cadastral", [cad_pdf] if cad_pdf else [])
    if cad_pages:
        cad_pdf = cad_pages[0]
    fac = first(os.path.join("1.基本資料", "*重要設施*.pdf"))
    fac_pages = M.sel_paths(sel, "facility", [fac] if fac else [])
    zon = first(os.path.join("1.基本資料", "png", "*都市計畫圖*.png"))
    zone_pages = M.sel_paths(sel, "zoning", [zon] if zon else [])

    # 面寬/深度：物件編輯器 → 成果圖 → 地籍圖(土地)
    f["dims_show"] = f["dims"]
    if not f["dims"] and plans and plans[0].lower().endswith(".pdf"):
        try:
            big = max([v for _, v in f.get("floors") or []] or [M._num(f.get("main")) or 0])
            pd_ = M.plan_dims(plans[0], big * 3.30579 if big else None)
        except Exception:
            pd_ = None
        if pd_:
            f["dims_show"] = f"約 {pd_[0]:g}／{pd_[1]:g} 米（成果圖量測）"
    if not f["dims_show"] and cad_pdf:
        try:
            dm = M.parcel_dims(cad_pdf)
        except Exception:
            dm = None
        if dm:
            f["dims_show"] = f"約 {dm['frontage']:g}／{dm['depth']:g} 米（地籍圖量測）"

    phones = M.lookup_phones(f["agents"], base_dir)
    M.log("製作：封面（透天・店面版）"); cover_slide(M, deck, C, f, cover, phones, show_agents)
    if M.sel_on(sel, "overview"):
        M.log("製作：物件亮點"); highlights(M, deck, C, f)
    if f.get("floors") and M.sel_on(sel, "floors"):
        M.log("製作：樓層配置"); floors_slide(M, deck, C, f)
    ll = M.case_coords(case_dir, base_dir, d) if M.sel_on(sel, "location") else None
    if ll:
        M.log("製作：地理位置"); M.slide_location(deck, *ll)
    for src in fac_pages:
        try:
            M.log("製作：周邊機能（重要設施）"); facilities(M, deck, C, f, src)
        except Exception as e:
            M.log(f"（周邊機能頁略過：{e}）")
    for src in zone_pages:
        M.log("製作：使用分區"); zoning(M, deck, C, f, src)
    for src in cad_pages:
        M.log("製作：地籍圖"); M.slide_cadastral(deck, src)
    for i, src in enumerate(plans):
        M.log("製作：建物測量成果圖"); M.slide_floorplan(deck, src, cont=i > 0)
    layouts = [p["path"] for p in by[M.LAYOUT_CAT]]
    names = {os.path.basename(x).lower() for x in layouts}
    layouts += [x for x in M.sel_paths(sel, "layout", M.find_layouts(roots)) if os.path.basename(x).lower() not in names]
    for i, src in enumerate(layouts[:6]):
        M.log(f"製作：格局平面圖（{os.path.basename(src)}）"); M.slide_layout(deck, src, cont=i > 0)
    M.log("製作：照片頁")
    for title, cats, cap in M.STREET_PHOTO_PAGES:
        items = M.ordered([p for p in chosen if p["cat"] in cats])
        if items and cover in items and len(items) > 1:
            items = [p for p in items if p is not cover] + [cover]      # 封面那張放最後，避免一翻開又看到
        start = 0
        for k, n in enumerate(M.split_pages(len(items), cap)):
            M.slide_photos(deck, title + ("（續）" if k else ""), items[start:start + n])
            start += n
    M.market_slides(deck, f, M.sel_paths(sel, "market", M.market_default(roots, f["deal"] == "租")), tmp)
    M.log("製作：聯絡資訊"); contact(M, deck, C, f, phones, show_agents)
    M.unique_shape_ids(deck.prs)
    deck.prs.save(out_path)
    return deck.n


# ---------------------------------------------------------------- 頁框：白底、左上琥珀短條＋深灰標題、細線、頁碼
def _street_chrome(M, deck, C, f):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    def new(title=None):
        s = deck.prs.slides.add_slide(deck.blank)
        deck.n += 1
        bg = s.background.fill
        bg.solid()
        bg.fore_color.rgb = M.BG
        if title:
            M.box(s, Inches(0.5), Inches(0.38), Inches(0.13), Inches(0.58), C["amber"], radius=False)
            M.text(s, Inches(0.78), Inches(0.26), Inches(8.5), Inches(0.8), title, 28, True, C["char"], anchor=MSO_ANCHOR.MIDDLE)
            if C.get("logo"):
                _logo(M, s, C["logo"], SW_ - Inches(0.5), Inches(0.42), Inches(0.48), right=True)
            else:
                M.text(s, Inches(8.0), Inches(0.3), SW_ - Inches(8.5), Inches(0.7), f["short"], 13, True, C["grey"], PP_ALIGN.RIGHT,
                       MSO_ANCHOR.MIDDLE)
            M.box(s, Inches(0.5), Inches(1.06), SW_ - Inches(1.0), Inches(0.02), C["line"], radius=False)
            M.box(s, Inches(0.5), M.FOOTER_Y + Inches(0.02), SW_ - Inches(1.0), Inches(0.015), C["line"], radius=False)
            M.text(s, Inches(0.5), M.FOOTER_Y + Inches(0.03), Inches(8), M.SH - M.FOOTER_Y - Inches(0.03), f["door"], 10, False,
                   C["grey"], anchor=MSO_ANCHOR.MIDDLE)
            M.text(s, SW_ - Inches(1.7), M.FOOTER_Y + Inches(0.03), Inches(1.2), M.SH - M.FOOTER_Y - Inches(0.03),
                   f"{deck.n:02d}", 11, True, C["amber_dk"], PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
        return s
    SW_ = M.SW
    deck.new = new


def _dark_logo(M, tmp):
    """公司橫式 LOGO 是白字(給深色底用)→ 淺色底要把白字改成公司藍；左邊紅藍標誌不動。回傳暫存 PNG(沒有素材回 None)"""
    from PIL import Image
    src = M.asset(M.ASSET_LOGO)
    if not src:
        return None
    im = Image.open(src).convert("RGBA")
    w, h = im.size
    px = im.load()
    # 左邊紅藍方形標誌：從左往右找有彩色的欄，彩色斷掉處＋一點白邊＝標誌結束(之後才是白字)
    mark = 0
    for x in range(w):
        if any(px[x, y][3] > 128 and max(px[x, y][:3]) - min(px[x, y][:3]) > 60 for y in range(0, h, 2)):
            mark = x
        elif mark and x - mark > 5:
            break
    mark += max(4, h // 40)
    for y in range(h):
        for x in range(mark, w):
            r, g, b, a = px[x, y]
            if a and r > 200 and g > 200 and b > 200:
                px[x, y] = (0x1F, 0x3F, 0x99, a)
    out = os.path.join(tmp, "logo_dark.png")
    im.save(out)
    return out


def _logo(M, s, path, x, y, h, right=False):
    """放 LOGO：給高度(設計座標)，寬度照比例；right=True 時 x 是右緣"""
    from PIL import Image
    iw, ih = Image.open(path).size
    w = int(h * iw / ih)
    s.shapes.add_picture(path, M._X(x - w if right else x), M._Y(y), M._X(w), M._Y(h))


def _label(f):
    kind = "店面" if "店" in str(f.get("use", "")) + str(f.get("type", "")) else "透天"
    return kind + ("出租" if f["deal"] == "租" else "出售")


# ---------------------------------------------------------------- 封面：左邊滿版照片、右邊深灰資訊欄
def cover_slide(M, deck, C, f, cover, phones, show_agents):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new()
    SW, SH = M.SW, M.SH
    pw = Inches(8.1)
    if cover:
        M.picture(s, cover["path"], 0, 0, pw, SH, deck.tmp)
    else:
        M.box(s, 0, 0, pw, SH, C["char2"], radius=False)
    M.box(s, pw, 0, SW - pw, SH, C["white"], radius=False)
    M.box(s, pw, 0, Inches(0.07), SH, C["amber"], radius=False)
    # 左上：紅色售／租章(公司套板)
    st = M.box(s, Inches(0.45), Inches(0.4), Inches(1.05), Inches(1.05), C["red"])
    st.adjustments[0] = 0.12
    M.text(s, Inches(0.45), Inches(0.4), Inches(1.05), Inches(1.05), f["deal"], 54, True, C["white"], PP_ALIGN.CENTER,
           MSO_ANCHOR.MIDDLE)
    x, w = pw + Inches(0.5), SW - pw - Inches(0.85)
    if C.get("logo"):
        _logo(M, s, C["logo"], x, Inches(0.42), Inches(0.5))
    tag = _label(f)
    M.box(s, x, Inches(1.2), Inches(1.6), Inches(0.44), C["amber"])
    M.text(s, x, Inches(1.2), Inches(1.6), Inches(0.44), tag, 15, True, C["white"], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    M.text(s, x, Inches(1.75), w, Inches(1.05), f["title"], 24 if len(f["title"]) <= 14 else 20, True, C["char"])
    M.text(s, x, Inches(2.8), w, Inches(0.4), f["door"], 14, False, C["grey"])
    M.box(s, x, Inches(3.28), Inches(0.8), Inches(0.04), C["amber"], radius=False)
    M.text(s, x, Inches(3.36), w, Inches(0.36), "月租" if f["deal"] == "租" else "開價", 14, True, C["amber"])
    tb = M.text(s, x, Inches(3.62), w, Inches(0.9), f["price"], 46 if len(str(f["price"])) <= 6 else 36, True, C["amber_dk"],
                anchor=MSO_ANCHOR.MIDDLE)
    if f["price_unit"]:
        r = tb.text_frame.paragraphs[0].add_run()
        r.text = f"  {f['price_unit']}"
        M._font(r, 18, True, C["grey"])
    stats = [(v, k) for k, v in (("面寬／深度", _short_dims(f.get("dims_show") or f.get("dims"))), ("面臨道路", f.get("road_w")),
                                 ("權狀坪數", f"{f['area']} 坪" if str(f.get("area", "")).strip() else ""),
                                 ("樓層", f.get("floor")), ("土地", f"{f['land']} 坪" if str(f.get("land", "")).strip() else ""),
                                 ("屋齡", f.get("age"))) if v][:4]
    for i, (v, k) in enumerate(stats):
        cx = x + (w // 2) * (i % 2)
        cy = Inches(4.65) + Inches(0.85) * (i // 2)
        M.text(s, cx, cy, w // 2, Inches(0.45), v, 18 if len(v) <= 9 else 14, True, C["char"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, cx, cy + Inches(0.42), w // 2, Inches(0.3), k, 11, False, C["grey"])
    agents = [f"{nm}　{phones.get(nm, '')}".strip() for nm in f["agents"][:3]] if show_agents else []
    if agents:
        M.box(s, x, Inches(6.38), w, Inches(0.015), C["line"], radius=False)
        M.text(s, x, Inches(6.45), w, Inches(0.8), "\n".join(agents), 13, True, C["char"], spacing=1.0)


def _short_dims(s):
    s = re.sub(r"[（(].*?[）)]", "", str(s or "")).replace(" ", "")
    m = re.search(r"([\d.]+)\D+([\d.]+)", s)
    return f"{m.group(1)}／{m.group(2)} 米" if m else s


# ---------------------------------------------------------------- 物件亮點：數字方塊＋訴求重點
def highlights(M, deck, C, f):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new("物件亮點")
    SW = M.SW
    tiles = [(f["price_label"], f"{f['price']}", f["price_unit"]),
             ("面寬／深度", _short_dims(f.get("dims_show") or f.get("dims")) or "—", ""),
             ("面臨道路", f.get("road_w") or "—", ""),
             ("權狀坪數", f"{f['area']}" if str(f.get("area", "")).strip() else "—", "坪" if str(f.get("area", "")).strip() else ""),
             ("樓層", f.get("floor") or "—", ""),
             ("土地坪數", f"{f['land']}" if str(f.get("land", "")).strip() else "—", "坪" if str(f.get("land", "")).strip() else ""),
             ("屋齡", f.get("age") or "—", ""),
             ("使用分區", f.get("zone") or "—", "")]
    x0, y0, gx, gy = Inches(0.5), Inches(1.3), Inches(0.2), Inches(0.2)
    cw, ch = (SW - Inches(1.0) - gx * 3) // 4, Inches(1.2)
    for i, (k, v, unit) in enumerate(tiles):
        x, y = x0 + (cw + gx) * (i % 4), y0 + (ch + gy) * (i // 4)
        dark = i == 0
        M.box(s, x, y, cw, ch, C["panel"] if dark else C["white"], C["amber"] if dark else C["line"])
        if dark:
            M.box(s, x, y, Inches(0.1), ch, C["amber"], radius=False)
        M.text(s, x + Inches(0.22), y + Inches(0.1), cw - Inches(0.3), Inches(0.35), k, 13, True, C["amber"] if dark else C["grey"])
        big = 28 if len(v) <= 6 else (22 if len(v) <= 9 else 17)
        tb = M.text(s, x + Inches(0.22), y + Inches(0.45), cw - Inches(0.3), Inches(0.65), v, big, True,
                    C["amber_dk"] if dark else C["char"], anchor=MSO_ANCHOR.MIDDLE)
        if unit:
            r = tb.text_frame.paragraphs[0].add_run()
            r.text = f" {unit}"
            M._font(r, 14, True, C["grey"])
    # 下半：訴求重點(左)＋適合行業/補充(右)
    ty = y0 + (ch + gy) * 2 + Inches(0.1)
    th = M.FOOTER_Y - Inches(0.15) - ty
    lw = Inches(8.1)
    M.box(s, x0, ty, lw, th, C["panel"])
    M.text(s, x0 + Inches(0.3), ty + Inches(0.12), Inches(4), Inches(0.4), "物件特色", 15, True, C["amber_dk"])
    pts = f["points"][:5] or ["（請在物件編輯器填寫「訴求重點」）"]
    each = min(Inches(0.56), (th - Inches(0.6)) // max(1, len(pts)))
    yy = ty + Inches(0.55)
    for i, p in enumerate(pts):
        M.box(s, x0 + Inches(0.3), yy + (each - Inches(0.36)) // 2, Inches(0.36), Inches(0.36), C["amber"])
        M.text(s, x0 + Inches(0.3), yy + (each - Inches(0.36)) // 2, Inches(0.36), Inches(0.36), str(i + 1), 13, True, C["white"],
               PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        M.text(s, x0 + Inches(0.85), yy, lw - Inches(1.1), each, p, 16 if len(p) <= 30 else 13, True, C["char"],
               anchor=MSO_ANCHOR.MIDDLE)
        yy += each
    rx = x0 + lw + Inches(0.25)
    rw = SW - Inches(0.5) - rx
    M.box(s, rx, ty, rw, th, C["white"], C["line"])
    rows = [(k, v) for k, v in (("適合行業", "、".join(f.get("trades") or [])), ("格局", f.get("layout")),
                                ("建築結構", f.get("struct")), ("座向", f.get("facing")), ("建蔽／容積", f.get("bcr_far")),
                                ("所在商圈", f.get("district")), ("鄰近學校", f.get("school"))) if str(v or "").strip()][:6]
    rh = min(Inches(0.5), (th - Inches(0.3)) // max(1, len(rows)))
    for i, (k, v) in enumerate(rows):
        cy = ty + Inches(0.15) + rh * i
        M.text(s, rx + Inches(0.25), cy, Inches(1.3), rh, k, 13, True, C["grey"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, rx + Inches(1.55), cy, rw - Inches(1.7), rh, str(v), 14 if len(str(v)) <= 12 else 12, True, C["char"],
               anchor=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------- 樓層配置：每層一列＋面積長條
def floors_slide(M, deck, C, f):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new("樓層配置")
    SW = M.SW
    fl = list(f["floors"])
    have = {M._cn_int(k) for k, _ in fl if "地下" not in k}
    missing = []
    if f.get("above"):
        missing = [n for n in range(1, int(f["above"]) + 1) if n not in have]
    rows = [(k, v, True) for k, v in fl]
    # 不在本建號的樓層(例：一樓、地下室另有登記)也畫出來(灰色)，整棟看起來才完整
    cn = "一二三四五六七八九十"
    for n in missing:
        rows.append((f"{cn[n - 1] if n <= 10 else n}層", 0, False))
    if f.get("below") and not any("地下" in k for k, _ in fl):
        rows.append(("地下層", 0, False))
    rows.sort(key=lambda r: M._floor_order(r[0]), reverse=True)        # 高樓層在上，像一棟樓
    lx, lw = Inches(0.5), Inches(8.3)
    top = Inches(1.35)
    avail = M.FOOTER_Y - Inches(0.9) - top
    if missing or f.get("below"):
        avail -= Inches(0.7)                                             # 留給下方「不在本建號」的說明
    # 樓層少(透天兩層)時每層拉高、整欄疊滿，像一棟樓；長條本身最粗 0.9 吋，字跟著放大
    rh = min(Inches(2.6), avail // max(1, len(rows)))
    big = rh / Inches(0.85)
    fs = int(min(26, max(17, 17 * big ** 0.5)))
    bh = min(Inches(0.9), rh - Inches(0.32))
    vmax = max([v for _, v, _ in rows] + [0])
    for i, (k, v, ok) in enumerate(rows):
        y = top + rh * i
        M.box(s, lx, y + Inches(0.06), Inches(1.6), rh - Inches(0.12), C["panel"] if ok else C["white"], C["amber"] if ok else C["line"])
        M.text(s, lx, y + Inches(0.06), Inches(1.6), rh - Inches(0.12), k, fs, True, C["amber_dk"] if ok else C["light"],
               PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        if not ok:
            M.text(s, lx + Inches(1.75), y, lw - Inches(1.9), rh, "不在本建號登記範圍", 14, True, C["grey"], anchor=MSO_ANCHOR.MIDDLE)
            continue
        bw = int((lw - Inches(3.4)) * (v / vmax)) if vmax else 0
        M.box(s, lx + Inches(1.75), y + (rh - bh) // 2, max(bw, Inches(0.2)), bh, C["amber"])
        M.text(s, lx + Inches(1.75) + bw + Inches(0.15), y, Inches(1.6), rh, f"{v:g} 坪", fs, True, C["char"],
               anchor=MSO_ANCHOR.MIDDLE)
    note = []
    if missing:
        note.append("、".join(f"{n} 樓" for n in missing) + " 不在本建號登記範圍")
    if f.get("below"):
        if not any("地下" in k for k, _ in fl):
            note.append(f"地下 {f['below']} 層不在本建號登記範圍")
    if note:
        M.text(s, lx, M.FOOTER_Y - Inches(0.75), lw, Inches(0.6), "※ " + "；".join(note) + "（以謄本、現況為準）", 13, True,
               C["amber_dk"], anchor=MSO_ANCHOR.MIDDLE)
    # 右欄：合計
    rx, rw = Inches(9.2), SW - Inches(9.7)
    M.box(s, rx, top, rw, Inches(1.35), C["panel"], C["amber"])
    M.box(s, rx, top, Inches(0.1), Inches(1.35), C["amber"], radius=False)
    M.text(s, rx + Inches(0.3), top + Inches(0.12), rw, Inches(0.4), "權狀合計", 14, True, C["amber"])
    tb = M.text(s, rx + Inches(0.3), top + Inches(0.5), rw - Inches(0.5), Inches(0.75), f"{f['area']}", 34, True, C["amber_dk"],
                anchor=MSO_ANCHOR.MIDDLE)
    r = tb.text_frame.paragraphs[0].add_run()
    r.text = " 坪"
    M._font(r, 16, True, C["grey"])
    info = [(k, v) for k, v in (("主建物", f"{f['main']} 坪" if not M._zero(f.get("main")) and str(f.get("main", "")).strip() else ""),
                                ("附屬建物", f"{f['attach']} 坪" if not M._zero(f.get("attach")) and str(f.get("attach", "")).strip() else ""),
                                ("總樓層", f"地上 {f['above']} 層" + (f"・地下 {f['below']} 層" if f.get("below") else "")
                                 if f.get("above") else ""),
                                ("建築結構", f.get("struct")), ("竣工", re.sub(r"(?<=[國年月])0+(?=\d)", "", str(f.get("built") or ""))),
                                ("格局", f.get("layout"))) if str(v or "").strip()]
    y = top + Inches(1.55)
    for k, v in info:
        M.box(s, rx, y, rw, Inches(0.62), C["white"], C["line"])
        M.text(s, rx + Inches(0.25), y, Inches(1.3), Inches(0.62), k, 13, True, C["grey"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, rx + Inches(1.5), y, rw - Inches(1.65), Inches(0.62), str(v), 14 if len(str(v)) <= 12 else 12, True, C["char"],
               anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(0.72)


# ---------------------------------------------------------------- 周邊機能：重要設施查詢結果(地圖＋距離表)
def facilities(M, deck, C, f, pdf):
    import fitz
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    doc = fitz.open(pdf)
    pg = doc[0]
    imgs = pg.get_image_info()
    if not imgs:
        raise ValueError("找不到地圖")
    r = fitz.Rect(max(imgs, key=lambda i: fitz.Rect(i["bbox"]).width * fitz.Rect(i["bbox"]).height)["bbox"])   # 舊版 PyMuPDF 沒有 get_area
    png = os.path.join(deck.tmp, "facility_map.png")
    pg.get_pixmap(matrix=fitz.Matrix(4, 4), clip=r, alpha=False).save(png)
    # 清單是一般文字(不是表格物件)：依第一頁表頭「名稱」「地址」「距離」的位置切欄，同一列(y 相近)的字歸一起
    hw = {w[4]: w for w in pg.get_text("words") if w[4] in ("名稱", "地址") or w[4].startswith("距離")}
    x_name = hw["名稱"][0] - 45 if "名稱" in hw else 100
    x_addr = hw["地址"][0] - 100 if "地址" in hw else 215
    x_dist = next((w[0] for k, w in hw.items() if k.startswith("距離")), 440) - 5
    rows = []
    for p in doc:
        lines = {}
        for w in p.get_text("words"):
            lines.setdefault(round(w[1] / 4), []).append(w)
        for ws in (lines[k] for k in sorted(lines)):
            ws.sort(key=lambda w: w[0])
            no = next((w[4] for w in ws if w[0] < x_name and w[4].isdigit()), "")
            name = "".join(w[4] for w in ws if x_name <= w[0] < x_addr)
            dist = next((w[4] for w in ws if w[0] >= x_dist and re.fullmatch(r"[\d.]+", w[4])), "")
            if no and name:
                rows.append((no, name, dist))
    title = "周邊機能"
    m = re.search(r"週邊\s*(\d+)\s*公尺", pg.get_text())
    s = deck.new(title + (f"（{m.group(1)} 公尺內）" if m else ""))
    SW = M.SW
    mx, my, mw, mh = Inches(0.5), Inches(1.3), Inches(6.6), M.FOOTER_Y - Inches(1.5)
    M.box(s, mx, my, mw, mh, C["white"], C["line"])
    M.picture(s, png, mx + Inches(0.1), my + Inches(0.1), mw - Inches(0.2), mh - Inches(0.2), deck.tmp, mode="contain")
    rx, rw = Inches(7.4), SW - Inches(7.9)
    rows = rows[:10]
    if not rows:
        M.text(s, rx, my, rw, Inches(0.6), "（重要設施清單讀取不到，請看左圖）", 15, True, C["grey"])
        return
    rh = min(Inches(0.52), (mh - Inches(0.5)) // len(rows))
    M.text(s, rx, my, rw, Inches(0.4), "名稱", 12, True, C["grey"])
    M.text(s, rx, my, rw, Inches(0.4), "距離", 12, True, C["grey"], PP_ALIGN.RIGHT)
    y = my + Inches(0.45)
    for no, name, dist in rows:
        M.box(s, rx, y + (rh - Inches(0.34)) // 2, Inches(0.34), Inches(0.34), C["amber"])
        M.text(s, rx, y + (rh - Inches(0.34)) // 2, Inches(0.34), Inches(0.34), no, 11, True, C["white"], PP_ALIGN.CENTER,
               MSO_ANCHOR.MIDDLE)
        M.text(s, rx + Inches(0.48), y, rw - Inches(1.6), rh, name, 14 if len(name) <= 14 else 12, True, C["char"],
               anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, rx + rw - Inches(1.1), y, Inches(1.1), rh, f"{dist} m" if dist else "", 14, True, C["amber_dk"], PP_ALIGN.RIGHT,
               MSO_ANCHOR.MIDDLE)
        M.box(s, rx, y + rh - Inches(0.01), rw, Inches(0.01), C["line"], radius=False)
        y += rh
    M.text(s, rx, M.FOOTER_Y - Inches(0.5), rw, Inches(0.4), "資料來源：重要環境設施查詢（僅供參考）", 11, False, C["grey"])


# ---------------------------------------------------------------- 使用分區：都市計畫截圖(只取地圖區)＋分區/建蔽容積
def zoning(M, deck, C, f, png):
    from PIL import Image
    from pptx.util import Inches
    from pptx.enum.text import MSO_ANCHOR
    im = Image.open(png).convert("RGB")
    w, h = im.size
    # 高雄都計系統畫面：左邊選單約 21%、上方標題列約 8%、下方比例尺列約 7% → 只留地圖
    crop = im.crop((int(w * 0.215), int(h * 0.18), w, int(h * 0.92)))      # 也切掉地圖上方那排工具按鈕
    out = os.path.join(deck.tmp, "zone_map.png")
    crop.save(out)
    s = deck.new("使用分區")
    SW = M.SW
    mx, my, mw, mh = Inches(0.5), Inches(1.3), Inches(8.6), M.FOOTER_Y - Inches(1.5)
    M.box(s, mx, my, mw, mh, C["white"], C["line"])
    M.picture(s, out, mx + Inches(0.1), my + Inches(0.1), mw - Inches(0.2), mh - Inches(0.2), deck.tmp, mode="contain")
    rx, rw = Inches(9.4), SW - Inches(9.9)
    M.box(s, rx, my, rw, Inches(1.5), C["panel"], C["amber"])
    M.box(s, rx, my, Inches(0.1), Inches(1.5), C["amber"], radius=False)
    M.text(s, rx + Inches(0.25), my + Inches(0.12), rw, Inches(0.4), "使用分區", 13, True, C["amber"])
    M.text(s, rx + Inches(0.25), my + Inches(0.5), rw - Inches(0.4), Inches(0.85), f.get("zone") or "—",
           22 if len(f.get("zone") or "") <= 7 else 17, True, C["amber_dk"], anchor=MSO_ANCHOR.MIDDLE)
    y = my + Inches(1.7)
    for k, v in (("建蔽率／容積率", f.get("bcr_far")), ("土地坪數", f"{f['land']} 坪" if str(f.get("land", "")).strip() else ""),
                 ("地段地號", f.get("lots"))):
        if not str(v or "").strip():
            continue
        M.box(s, rx, y, rw, Inches(0.95), C["white"], C["line"])
        M.text(s, rx + Inches(0.25), y + Inches(0.08), rw, Inches(0.35), k, 12, True, C["grey"])
        M.text(s, rx + Inches(0.25), y + Inches(0.4), rw - Inches(0.4), Inches(0.5), str(v), 16 if len(str(v)) <= 14 else 12, True,
               C["char"], anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(1.08)
    M.text(s, rx, M.FOOTER_Y - Inches(0.5), rw, Inches(0.4), "資料來源：都市計畫地理資訊系統", 11, False, C["grey"])


# ---------------------------------------------------------------- 聯絡頁(封底)：白底，左邊公司直式 LOGO＋Full Life＋電話，右邊預約看屋
def contact(M, deck, C, f, phones, show_agents):
    from PIL import Image
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    s = deck.new()
    SW, SH = M.SW, M.SH
    M.box(s, 0, 0, SW, SH, C["white"], radius=False)
    M.box(s, 0, SH - Inches(0.14), SW, Inches(0.14), C["amber"], radius=False)
    M.box(s, Inches(5.2), Inches(0.9), Inches(0.02), SH - Inches(1.8), C["line"], radius=False)
    lv = M.asset(M.ASSET_LOGO_V)
    if lv:
        iw, ih = Image.open(lv).size
        h = Inches(2.9)
        w = int(h * iw / ih)
        s.shapes.add_picture(lv, M._X(Inches(2.6) - w // 2), M._Y(Inches(1.0)), M._X(w), M._Y(h))
    fl = M.asset(M.ASSET_FULLLIFE)
    if fl:
        iw, ih = Image.open(fl).size
        w = Inches(3.0)
        s.shapes.add_picture(fl, M._X(Inches(2.6) - w // 2), M._Y(Inches(4.15)), M._X(w), M._Y(int(w * ih / iw)))
    M.text(s, Inches(0.4), Inches(5.0), Inches(4.4), Inches(0.5), M.COMPANY, 18, True, C["char"], PP_ALIGN.CENTER,
           MSO_ANCHOR.MIDDLE)
    M.text(s, Inches(0.4), Inches(5.5), Inches(4.4), Inches(0.5), M.COMPANY_TEL, 18, True, C["amber_dk"], PP_ALIGN.CENTER,
           MSO_ANCHOR.MIDDLE)
    x = Inches(5.8)
    w = SW - x - Inches(0.7)
    M.text(s, x, Inches(1.0), w, Inches(0.45), _label(f), 16, True, C["amber"])
    M.text(s, x, Inches(1.45), w, Inches(1.0), "歡迎預約看屋", 40, True, C["char"])
    M.text(s, x, Inches(2.5), w, Inches(0.5), f"{f['title']}　｜　{f['door']}", 15, True, C["grey"])
    M.box(s, x, Inches(3.15), Inches(1.0), Inches(0.05), C["amber"], radius=False)
    agents = [(nm, phones.get(nm, "")) for nm in f["agents"][:3]] if show_agents else []
    for i, (nm, ph) in enumerate(agents):
        y = Inches(3.45) + Inches(0.8) * i
        M.text(s, x, y, Inches(2.4), Inches(0.7), nm, 24, True, C["char"], anchor=MSO_ANCHOR.MIDDLE)
        M.text(s, x + Inches(2.4), y, Inches(4.0), Inches(0.7), ph, 24, True, C["amber_dk"], anchor=MSO_ANCHOR.MIDDLE)
    tb = M.text(s, x, Inches(6.0), w, Inches(0.8), f"{f['price_label']}　", 18, True, C["grey"], anchor=MSO_ANCHOR.MIDDLE)
    r = tb.text_frame.paragraphs[0].add_run()
    r.text = f"{f['price']} {f['price_unit']}"
    M._font(r, 30, True, C["amber_dk"])
