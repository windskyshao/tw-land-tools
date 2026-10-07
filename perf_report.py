# -*- coding: utf-8 -*-
"""
業績分配表 產生器(買賣、租賃都用)
- 自動帶入：契約編號、物件編號、物件名稱、物件地址、租件/售件(契約起迄＝委託書期間，使用者自選)、
  賣方(出租方)經紀人 = 物件編輯器承辦人(通訊錄已選人員)
- 確認視窗調整比例(賣方/買方合計 100%；賣方各經紀人加總 = 賣方%)→ 產生每格可編輯的 PDF
用法：perf_report.py --case-dir <案件資料夾> [--base-dir <主程式目錄>]
"""
import os
import re
import sys
import argparse
from datetime import date

import fitz

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# 打包版 python_embedded 有 python39._pth → 不吃 PYTHONPATH、也不自動加腳本目錄，要自己加才 import 得到 rent_report
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)
import rent_report as rr  # noqa: E402
TEMPLATE_NAME = "業績分配表_範本.pdf"
MAX_AGENTS = 6                               # 每方 6 人：大格內左三右三

# 範本上 □ 字元左上角
BOX = {"租件": (484.1, 184.7), "售件": (540.9, 184.7), "上網": (484.7, 223.1), "不上網": (541.7, 223.1),
       "代書": (379.0, 466.0), "其他": (415.0, 466.0), "帆布": (487.0, 466.0)}
SELLER_BOX = fitz.Rect(58.2, 276.5, 352.8, 379.4)    # 賣方(出租方)經紀人大格
BUYER_BOX = fitz.Rect(355.7, 276.5, 556.8, 379.4)    # 買方(承租方)經紀人大格
NAME_FONT_FILES = ["秀風體W3.TTC", "華康秀風體W3.TTC", "DFHsiuW3.TTC"]
_KAIU = r"C:/Windows/Fonts/kaiu.ttf"


def find_name_font():
    """姓名用華康秀風體(商用字型，不隨系統散布；用本機已安裝的)，沒有就退回標楷體"""
    dirs = [os.path.join(os.environ.get("WINDIR", r"C:/Windows"), "Fonts"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")]
    for d in dirs:
        for n in NAME_FONT_FILES:
            p = os.path.join(d, n)
            if os.path.exists(p):
                return p, True
    for d in dirs:                                   # 檔名不同時：用字型內部名稱找
        try:
            for fn in os.listdir(d):
                if fn.lower().endswith((".ttc", ".ttf")) and ("秀風" in fn or "hsiu" in fn.lower()):
                    return os.path.join(d, fn), True
        except OSError:
            pass
    return _KAIU, False


def gather(case_dir):
    other = os.path.join(case_dir, "4.其他相關")
    df_path = os.path.join(other, "data_final.json")
    df = rr._load_json(df_path) or {}
    notes = []
    v = {"data_final": df_path if df else ""}
    if not df:
        notes.append(("warn", "找不到物件資料(4.其他相關\\data_final.json)，請先在物件編輯器按「💾 儲存」"))
    v["契約編號"] = str(df.get("契約編號", "") or "")
    v["物件編號"] = str(df.get("物件編號", "") or "")
    v["物件名稱"] = str(df.get("案名", "") or "")
    addr = rr.fw2hw(df.get("建物門牌", "") or "")
    land = df.get("土地標示", "") or ""
    m = re.match(r"^(..[市縣])", land)
    if addr and m and not re.match(r"^..[市縣]", addr):
        addr = m.group(1) + addr
    v["物件地址"] = addr or land
    v["建物門牌"] = addr
    v["類型"] = "租件" if "租" in (df.get("交易類型", "") or "") else "售件"
    v["上網"] = True
    v["賣方%"] = "50"
    agents = [a for a in (df.get("承辦人列表") or []) if a][:MAX_AGENTS]
    v["經紀人"] = [[a, ""] for a in agents]
    v["買方經紀人"] = []          # 成交時才知道，預設空白
    _split_even(v)
    if not agents:
        notes.append(("warn", "物件編輯器的承辦人是空的，賣方(出租方)經紀人請手動填"))

    # 契約起迄＝買賣/租賃委託書上的期間(手寫在委託書，系統沒有資料)，由使用者自己選
    v["起"] = v["迄"] = None
    v["租約檔"] = rr.find_lease_file(case_dir) or ""      # 只用來決定存檔位置(跟租約同資料夾)
    return v, notes


def _split_even(v):
    """賣方% 平均分給各經紀人(整數，餘數給第一位)"""
    ags = v["經紀人"]
    if not ags:
        return
    try:
        total = int(float(v["賣方%"]))
    except ValueError:
        return
    each, rest = divmod(total, len(ags))
    for i, a in enumerate(ags):
        a[1] = str(each + (rest if i == 0 else 0))


def _agent_layout(box, n, texts, fnt):
    """像手寫一樣排：人少字大、人多自然縮小；左右置中、偏上置中。
    回傳 [(x 起點, 基線, 字級, 簽名線長)]；1~3 人一欄，4~6 人兩欄(左邊放多的一半)"""
    cols = 1 if n <= 3 else 2
    rows = -(-n // cols)
    col_w = box.width / cols
    pad_top = 6
    avail_h = box.height - pad_top * 2
    for size in [x / 2 for x in range(44, 15, -1)]:          # 22pt → 8pt
        row_h = size * 2.05
        line = size * 3.6                                      # 簽名線長
        gap = size * 0.5
        if row_h * rows > avail_h:
            continue
        if all(fnt.text_length(t, size) + gap + line <= col_w * 0.9 for t in texts):
            break
    block_h = row_h * rows
    top = box.y0 + pad_top + (avail_h - block_h) * 0.35          # 偏上置中，不留一大片白
    out = []
    for i in range(n):
        c, r = (0, i) if cols == 1 else ((0, i) if i < rows else (1, i - rows))
        unit = fnt.text_length(texts[i], size) + gap + line
        cx = box.x0 + col_w * (c + 0.5)
        base = top + row_h * r + row_h * 0.62
        out.append((cx - unit / 2, base, size, line, gap))
    return out


_KFNT = []


def _kai_font():
    if not _KFNT:
        _KFNT.append(fitz.Font(fontfile=_KAIU))
    return _KFNT[0]


def _draw_agents(page, box, agents, font_path, fnt):
    if not agents:
        return
    texts = [f"{n} {p}%" if p else n for n, p in agents]
    for (x, base, size, line, gap), t in zip(_agent_layout(box, len(agents), texts, fnt), texts):
        # 秀風體沒有粗體 → 填色+描邊模擬加粗
        page.insert_text((x, base), t, fontname="nm", fontfile=font_path, fontsize=size,
                         color=(0, 0, 1), fill=(0, 0, 1), render_mode=2, border_width=0.04)
        lx0 = x + fnt.text_length(t, size) + gap
        lx1 = lx0 + line
        page.draw_line((lx0, base + 2), (lx1, base + 2), color=(0, 0, 0), width=0.7)
        hint = "(請簽名)"
        hs = max(6.5, size * 0.42)
        hw = _kai_font().text_length(hint, hs)
        page.insert_text(((lx0 + lx1) / 2 - hw / 2, base - size * 0.55), hint, fontname="kai", fontfile=_KAIU,
                         fontsize=hs, color=(0.55, 0.55, 0.55))


def _date_text(d, tail=""):
    if not d or not all(d.get(k) for k in ("年", "月", "日")):
        return ""
    return f"{int(d['年'])}年{int(d['月']):02d}月{int(d['日']):02d}日{tail}"


def build_pdf(v, template_path, out_path):
    doc = fitz.open(template_path)
    page = doc[0]
    fields = []
    font_path, is_hsiu = find_name_font()
    fnt = fitz.Font(fontfile=font_path)
    tw = (lambda t, sz: fnt.text_length(t, sz)) if is_hsiu else (lambda t, sz: rr._text_w(t, sz))

    def wrap(text, sz, width):
        """逐字排，滿了就換行；不在數字/英文中間斷(往回退到「、，」或非數字處)"""
        lines, cur = [], ""
        for ch in text:
            if cur and tw(cur + ch, sz) > width:
                cut = len(cur)
                if ch.isalnum() and cur[-1].isalnum():
                    k = len(cur)
                    while k > 0 and cur[k - 1].isalnum():
                        k -= 1
                    if k > len(cur) // 2:
                        cut = k
                lines.append(cur[:cut])
                cur = cur[cut:] + ch
            else:
                cur += ch
        return lines + ([cur] if cur else [])

    def tf(x0, x1, base, value="", size=12, align=0, min_size=9, top=None):
        """字放大為原則；太長先縮字(不小於 min_size)保持一行；給了 top(可用上緣)還放不下就分多行"""
        value = "" if value is None else str(value)
        width = (x1 - x0) - 3
        rect_top = base - max(11.5, size * 0.95)
        if value and tw(value, size) > width:
            fit = size * width / tw(value, size)
            if fit >= min_size or top is None:
                size = max(7, fit)
            else:                                    # 多行：找能放進高度的最大字級
                avail_h = base + 3 - top
                sz = min_size
                while sz > 7:
                    ls = wrap(value, sz, width)
                    if len(ls) * sz * 1.18 <= avail_h:
                        break
                    sz -= 0.5
                size = sz
                value = "\n".join(wrap(value, sz, width))
                rect_top = top
        fields.append((0, fitz.Rect(x0, rect_top, x1, base + 3), value, size, align))

    tf(507, 557, 95, v.get("物件編號"), 14)
    tf(60, 153, 145, v.get("契約編號"), 16, 1)
    tf(443, 555, 131, _date_text(v.get("起")), 14, 1)
    tf(443, 555, 167, _date_text(v.get("迄")), 14, 1)
    tf(137, 437, 205, v.get("物件名稱"), 18, 0, 13, top=174)
    tf(137, 437, 243.5, v.get("物件地址"), 18, 0, 13, top=211)
    seller = str(v.get("賣方%", "") or "")
    buyer = str(100 - int(float(seller))) if re.fullmatch(r"\d+(\.\d+)?", seller) else ""
    tf(236, 300, 270, seller, 18, 1)
    tf(503, 531, 270, buyer, 18, 1)

    ags = [a for a in (v.get("經紀人") or []) if a[0].strip()][:MAX_AGENTS]
    bags = [a for a in (v.get("買方經紀人") or []) if a[0].strip()][:MAX_AGENTS]
    # 認證欄、應收/實收、合計、檢核：成交後手填，留空欄位
    tf(157, 181, 476, "", 13, 1); tf(193, 217, 476, "", 13, 1); tf(229, 253, 476, "", 13, 1)
    tf(290, 354, 476, "", 13, 2); tf(428, 486, 476, "", 12)
    for base in (673, 708.5):
        tf(150, 165, base, "", 12, 1); tf(205, 290, base, "", 13)
        tf(402, 417, base, "", 12, 1); tf(451, 554, base, "", 13)
    tf(282, 395, 750, "", 14); tf(282, 395, 802, "", 14)

    on = {v.get("類型", "售件"), "上網" if v.get("上網", True) else "不上網"}
    boxes = rr._find_boxes(page)
    for r in boxes:
        page.add_redact_annot(fitz.Rect(r.x0 + 1, r.y0 + 2, r.x1 - 1, r.y1 - 2))
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE,
                          text=fitz.PDF_REDACT_TEXT_REMOVE)
    checks = []
    for r in boxes:
        name = next((n for n, (x, y) in BOX.items() if abs(r.x0 - x) < 3 and abs(r.y0 - y) < 3), None)
        checks.append((0, r, name in on))
    # 經紀人姓名(秀風體) + 簽名底線
    nm_xref = page.insert_font(fontname="nm", fontfile=font_path)
    # 契約起迄：「至」放兩排中間
    page.insert_text((499 - fnt.text_length("至", 14) / 2, 149.5), "至", fontname="nm", fontfile=font_path,
                     fontsize=14, color=(0, 0, 1), fill=(0, 0, 1), render_mode=2, border_width=0.04)
    for box, lst in ((SELLER_BOX, ags), (BUYER_BOX, bags)):
        _draw_agents(page, box, lst, font_path, fnt)
    v["_字型"] = "華康秀風體" if is_hsiu else "標楷體(本機沒有華康秀風體)"
    print(f"[資訊] 經紀人姓名字型：{v['_字型']}")

    # 輸入欄位也用秀風體：只「指名」不嵌入(授權不允許嵌入後拿來編輯)，本機有裝就顯示秀風體
    # 欄位外觀用內嵌秀風體(只內嵌用到的字)：subset 只看頁面文字，所以把欄位文字以「隱形字」也寫進頁面，字才會被保留
    if is_hsiu:
        for _, r, val, size, _a in fields:
            if val:
                page.insert_text((r.x0, r.y1 - 3), val, fontname="nm", fontfile=font_path, fontsize=size, render_mode=3)
    rr._add_form(doc, fields, checks, box_inset=(0.7, 0.8, 0.8, 0.9),
                 font_name="DFHsiuW3-B5" if is_hsiu else "DFKaiShu-SB-Estd-BF",
                 ap_font=(nm_xref, fnt) if is_hsiu else None, bold=True)
    try:
        doc.subset_fonts()       # 只嵌入用到的字(秀風體整套 7MB)
    except Exception:
        pass
    doc.save(out_path, garbage=3, deflate=True)
    return out_path


# ============================================================
# 確認視窗
# ============================================================
def run_dialog(case_dir, v, notes, template_path):
    import tkinter as tk
    from tkinter import ttk
    import deal_ui as ui

    root = tk.Tk()
    root.title("業績分配表")
    ui.apply(root)
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    W, H = min(1040, sw - 80), min(900, sh - 100)
    root.geometry(f"{W}x{H}+{(sw - W) // 2}+{max(10, (sh - H) // 2 - 30)}")
    root.minsize(860, 600)

    def big_dialog(title, msg, yesno=False):
        res = {"v": False}
        dlg = tk.Toplevel(root, bg=ui.CARD)
        dlg.title(title)
        dlg.transient(root)
        dlg.grab_set()
        tk.Frame(dlg, bg=ui.ACCENT if yesno else ui.PRIMARY, height=5).pack(fill=tk.X)
        tk.Label(dlg, text=msg, font=ui.F, fg=ui.TEXT, bg=ui.CARD, justify="left", wraplength=560,
                 padx=28, pady=22).pack()
        bf = tk.Frame(dlg, bg=ui.CARD, pady=14)
        bf.pack()

        def done(val):
            res["v"] = val
            dlg.destroy()
        if yesno:
            ui.button(bf, "仍要產生", lambda: done(True)).pack(side=tk.LEFT, padx=8)
            ui.button(bf, "回去修改", lambda: done(False), "ghost").pack(side=tk.LEFT, padx=8)
        else:
            ui.button(bf, "確定", lambda: done(True)).pack()
        dlg.update_idletasks()
        dlg.geometry(f"+{root.winfo_rootx() + (root.winfo_width() - dlg.winfo_width()) // 2}+{root.winfo_rooty() + 160}")
        root.wait_window(dlg)
        return res["v"]

    ui.header(root, "業績分配表", os.path.basename(case_dir))
    foot = ui.footer(root)
    body, canvas = ui.scroll_body(root)

    for lv, m in notes:
        ui.note(body, m, lv, W - 120)
    ui.note(body, "讀的是物件編輯器「已儲存」的資料；賣方名單＝物件編輯器的承辦人(通訊錄已選人員)", "info", W - 120)

    vars_ = {}

    # ---- 基本資料 ----
    c1 = ui.card(body, "基本資料", "自動帶入，可直接修改", "blue")
    for r_, (key, lab, wdt) in enumerate((("契約編號", "契約編號", 22), ("物件編號", "物件編號", 22),
                                          ("物件名稱", "物件名稱", 52), ("物件地址", "物件地址", 52))):
        ui.label(c1, lab, color=ui.SUB).grid(row=r_, column=0, sticky="e", padx=(0, 12), pady=4)
        vars_[key] = tk.StringVar(value=v.get(key) or "")
        ui.entry(c1, vars_[key], wdt, row=r_, column=1, sticky="w", pady=4)

    # ---- 契約起迄：共用月曆(deal_ui.date_field) ----
    c2 = ui.card(body, "契約起迄", "委託書上的期間；點日期框叫出月曆", "teal")
    T2 = ui.THEMES["teal"][0]
    for r_, key in enumerate(("起", "迄")):
        d = v.get(key) or {}
        for p in ("年", "月", "日"):
            vars_[f"{key}_{p}"] = tk.StringVar(value=str(int(d[p])) if d.get(p) else "")
        ui.label(c2, key, bold=True).grid(row=r_, column=0, padx=(0, 12), pady=5)
        start = (lambda: (vars_["起_年"].get(), vars_["起_月"].get())) if key == "迄" else None
        ui.date_field(c2, root, vars_[f"{key}_年"], vars_[f"{key}_月"], vars_[f"{key}_日"], T2, start).grid(
            row=r_, column=1, sticky="w", pady=4)

    # ---- 類別 ----
    c3 = ui.card(body, "類別", "租件/售件依物件編輯器自動選；上網預設", "purple")
    T3 = ui.THEMES["purple"][0]
    type_var = tk.StringVar(value=v.get("類型", "售件"))
    web_var = tk.StringVar(value="上網" if v.get("上網", True) else "不上網")
    ui.segmented(c3, type_var, ("租件", "售件"), T3).grid(row=0, column=0, sticky="w")
    ui.segmented(c3, web_var, ("上網", "不上網"), T3).grid(row=0, column=1, sticky="w", padx=(40, 0))

    # ---- 比例 ----
    c4 = ui.card(body, "經紀人與比例", "賣方＋買方＝100%；各方名單加總＝該方%", "orange")
    T4, TINT4, EDGE4 = ui.THEMES["orange"]
    top_row = tk.Frame(c4, bg=TINT4)
    top_row.pack(fill=tk.X, pady=(0, 10))
    seller_var = tk.StringVar(value=v.get("賣方%", "50"))
    ui.label(top_row, "賣方(出租方)", bold=True).pack(side=tk.LEFT)
    ttk.Combobox(top_row, textvariable=seller_var, values=["60", "50", "40"], font=ui.FB, width=4).pack(side=tk.LEFT, padx=6)
    ui.label(top_row, "%").pack(side=tk.LEFT)
    buyer_lab = ui.label(top_row, "", bold=True, color=T4)
    buyer_lab.pack(side=tk.LEFT, padx=(24, 0))
    qf = tk.Frame(top_row, bg=TINT4)
    qf.pack(side=tk.RIGHT)

    pct_vals = [str(i) for i in range(5, 101, 5)]

    def agent_list(title):
        box = tk.Frame(c4, bg="white", padx=12, pady=10, highlightthickness=1, highlightbackground=EDGE4)
        box.pack(fill=tk.X, pady=5)
        hd = tk.Frame(box, bg="white")
        hd.pack(fill=tk.X, pady=(0, 6))
        tk.Label(hd, text=title, font=ui.FB, fg=T4, bg="white").pack(side=tk.LEFT)
        tk.Label(hd, text="左三右三，對應表格位置", font=ui.FS, fg=ui.SUB, bg="white").pack(side=tk.LEFT, padx=10)
        lab = tk.Label(hd, text="", font=ui.FB, bg="white")
        lab.pack(side=tk.RIGHT)
        g = tk.Frame(box, bg="white")
        g.pack(anchor="w")
        rs = []
        return box, g, lab, rs

    def fill_list(g, rs, init):
        init = (init or []) + [["", ""]] * MAX_AGENTS
        for i in range(MAX_AGENTS):
            r_, c0 = i % 3, (i // 3) * 5
            nv = tk.StringVar(value=init[i][0])
            pv = tk.StringVar(value=init[i][1])
            tk.Label(g, text=f"{i + 1}", font=ui.FB, fg="white", bg=T4 if i < 3 else "#EBA25A",
                     width=2).grid(row=r_, column=c0, padx=(30 if c0 else 0, 6), pady=3)
            ui.entry(g, nv, 9, row=r_, column=c0 + 1, pady=3)
            ttk.Combobox(g, textvariable=pv, values=pct_vals, font=ui.F, width=4).grid(row=r_, column=c0 + 2, padx=(6, 2))
            tk.Label(g, text="%", font=ui.F, fg=ui.TEXT, bg="white").grid(row=r_, column=c0 + 3, sticky="w")
            rs.append((nv, pv))

    _, g1, sum_lab, rows = agent_list("賣方(出租方)經紀人")
    fill_list(g1, rows, v.get("經紀人"))
    _, g2, bsum_lab, brows = agent_list("買方(承租方)經紀人")
    fill_list(g2, brows, v.get("買方經紀人"))
    ui.note(c4, "PDF 上每位後面會留簽名底線；買方成交時再填也可以，空著就不畫線", "info", W - 160)

    def to_num(s):
        try:
            return float(s)
        except (TypeError, ValueError):
            return None

    def refresh(*_):
        sv = to_num(seller_var.get())
        buyer_lab.config(text=f"買方(承租方) {100 - sv:g} %" if sv is not None else "買方 ? %")
        for rs, lab, target, who in ((rows, sum_lab, sv, "賣方"),
                                     (brows, bsum_lab, (100 - sv) if sv is not None else None, "買方")):
            named = [p for n, p in rs if n.get().strip()]
            total = sum(to_num(p.get()) or 0 for p in named)
            if not named:
                lab.config(text="尚未填", fg=ui.SUB)
                continue
            ok = target is not None and abs(total - target) < 1e-6
            lab.config(text=f"合計 {total:g}%  " + (f"✔ 符合{who}%" if ok else f"✖ 要等於 {target:g}%" if target is not None else ""),
                       fg=ui.OK if ok else ui.BAD)
    for nv, pv in rows + brows:
        nv.trace_add("write", refresh)
        pv.trace_add("write", refresh)
    seller_var.trace_add("write", refresh)

    def set_ratio(s):
        seller_var.set(s)
        even()

    def even():
        sv = to_num(seller_var.get())
        if sv is None:
            return
        for rs, target in ((rows, int(sv)), (brows, 100 - int(sv))):
            names = [(n, p) for n, p in rs if n.get().strip()]
            if not names:
                continue
            each, rest = divmod(target, len(names))
            for i, (n, p) in enumerate(names):
                p.set(str(each + (rest if i == 0 else 0)))
    for t, s_ in (("60 / 40", "60"), ("50 / 50", "50")):
        ui.button(qf, t, lambda s_=s_: set_ratio(s_), "chip", T4).pack(side=tk.LEFT, padx=4)
    ui.button(qf, "平均分配", even, "chip", T4).pack(side=tk.LEFT, padx=4)
    refresh()

    result = {"ok": False}

    def collect():
        out = dict(v)
        for k in ("契約編號", "物件編號", "物件名稱", "物件地址"):
            out[k] = vars_[k].get().strip()
        for k in ("起", "迄"):
            d = {p: vars_[f"{k}_{p}"].get() for p in ("年", "月", "日")}
            out[k] = d if all(d.values()) else None
        out["類型"] = type_var.get()
        out["上網"] = web_var.get() == "上網"
        out["賣方%"] = seller_var.get().strip()
        out["經紀人"] = [[n.get().strip(), p.get().strip()] for n, p in rows if n.get().strip()]
        out["買方經紀人"] = [[n.get().strip(), p.get().strip()] for n, p in brows if n.get().strip()]
        return out

    def on_ok():
        out = collect()
        sv = to_num(out["賣方%"])
        total = sum(to_num(p) or 0 for _, p in out["經紀人"])
        probs = []
        if sv is None or not 0 <= sv <= 100:
            probs.append("賣方% 不是 0~100 的數字")
        else:
            if out["經紀人"] and abs(total - sv) > 1e-6:
                probs.append(f"賣方經紀人加總 {total:g}% 不等於賣方 {sv:g}%")
            btotal = sum(to_num(p) or 0 for _, p in out["買方經紀人"])
            if out["買方經紀人"] and abs(btotal - (100 - sv)) > 1e-6:
                probs.append(f"買方經紀人加總 {btotal:g}% 不等於買方 {100 - sv:g}%")
        if probs and not big_dialog("請確認", "\n".join("• " + p for p in probs) + "\n\n仍要產生 PDF 嗎？(之後可在 PDF 直接改)", True):
            return
        try:
            path = rr.output_path(case_dir, {"租約檔": out.get("租約檔", ""), "建物門牌": out.get("建物門牌", "")},
                                  prefix="業績分配表")
            build_pdf(out, template_path, path)
        except Exception as e:
            big_dialog("產生失敗", str(e))
            return
        result["ok"], result["path"] = True, path
        root.destroy()

    ui.button(foot, "📄  產生業績分配表 PDF", on_ok).pack(side=tk.RIGHT)
    ui.button(foot, "取消", root.destroy, "ghost").pack(side=tk.RIGHT, padx=10)
    tk.Label(foot, text="產生後會自動打開 PDF，每格都可再點進去改", font=ui.FS, fg=ui.SUB, bg=ui.CARD).pack(side=tk.LEFT)
    root.mainloop()
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case-dir", required=True)
    ap.add_argument("--base-dir", default="")
    ap.add_argument("--no-gui", action="store_true", help="測試用：直接用自動帶入值產生")
    ap.add_argument("--out", default="")
    a = ap.parse_args()
    rr._settings_dir = a.base_dir if a.base_dir and os.path.isdir(a.base_dir) else SCRIPT_DIR
    template = os.path.join(SCRIPT_DIR, TEMPLATE_NAME)
    if not os.path.exists(template) and a.base_dir:
        template = os.path.join(a.base_dir, TEMPLATE_NAME)
    if not os.path.exists(template):
        print(f"[錯誤] 找不到範本 {TEMPLATE_NAME}")
        return 1
    case_dir = os.path.abspath(a.case_dir)
    v, notes = gather(case_dir)
    for lv, m in notes:
        print(f"[{'警告' if lv == 'warn' else '資訊'}] {m}")
    if a.no_gui:
        path = a.out or rr.output_path(case_dir, v, prefix="業績分配表")
        build_pdf(v, template, path)
        print(f"[完成] {path}")
        return 0
    res = run_dialog(case_dir, v, notes, template)
    if res.get("ok"):
        print(f"[完成] 已產生：{res['path']}")
        try:
            os.chdir(os.path.dirname(res["path"]))
            for k in ("PYTHONHOME", "PYTHONPATH", "PYTHONNOUSERSITE"):
                os.environ.pop(k, None)
            os.startfile(res["path"])
        except Exception:
            pass
    else:
        print("[取消] 未產生業績分配表")
    return 0


if __name__ == "__main__":
    sys.exit(main())
