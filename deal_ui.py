# -*- coding: utf-8 -*-
"""成交文件確認視窗共用樣式(業績分配表 / 實價登錄申報書)
圓角卡片(各區塊同色系色調)、圓角按鈕、漸層標題列；字 14 起跳、說明文字不用淡灰"""
import tkinter as tk
from tkinter import ttk, font as tkfont

BG = "#F1F4F8"          # 視窗底色
CARD = "#FFFFFF"
BORDER = "#D5DCE6"
PRIMARY = "#1565C0"     # 與 PDF 藍字一致
PRIMARY_DK = "#0D47A1"
ACCENT = "#F28C28"      # 富住通橘
TEXT = "#1F2937"
SUB = "#374151"         # 說明文字(看得清楚)
OK = "#2E7D32"
BAD = "#C62828"
WARN_BG = "#FFF4E5"
WARN_FG = "#8A3B00"

FAMILY = "Microsoft JhengHei"
F = (FAMILY, 14)
FB = (FAMILY, 14, "bold")
FS = (FAMILY, 13)
FH = (FAMILY, 18, "bold")
FT = (FAMILY, 15, "bold")

# 區塊色系：(主色, 淡底, 邊框)
THEMES = {
    "blue":   ("#1565C0", "#F0F6FE", "#C9DCF6"),
    "teal":   ("#00796B", "#ECF7F5", "#BFE3DD"),
    "purple": ("#5E35B1", "#F4F0FC", "#D9CDF3"),
    "orange": ("#D9700E", "#FFF6EC", "#F6D7B5"),
    "green":  ("#2E7D32", "#F0F8F0", "#C8E3C9"),
    "rose":   ("#C2185B", "#FDF0F5", "#F3C8D9"),
    "indigo": ("#3949AB", "#F1F3FC", "#CDD3F1"),
    "amber":  ("#B26A00", "#FFF8E8", "#F2DDAE"),
    "cyan":   ("#00838F", "#EBF8FA", "#BDE5EA"),
    "slate":  ("#455A64", "#F2F5F7", "#CFD8DC"),
    "brown":  ("#795548", "#F8F3F0", "#E0D2CB"),
}


def apply(root):
    root.configure(bg=BG)
    root.option_add("*Font", F)
    root.option_add("*TCombobox*Listbox.font", F)
    root.option_add("*TCombobox*Listbox.selectBackground", PRIMARY)
    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except tk.TclError:
        pass
    st.configure("TCombobox", padding=4, arrowsize=16, fieldbackground="white", background="white",
                 bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
    st.map("TCombobox", bordercolor=[("focus", PRIMARY)], fieldbackground=[("readonly", "white")])
    st.configure("Vertical.TScrollbar", gripcount=0, background="#C9D2DE", troughcolor=BG, bordercolor=BG,
                 arrowcolor=SUB, relief="flat")


def _mix(c1, c2, t):
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{int(a[k] + (b[k] - a[k]) * t):02x}" for k in range(3))


def _round_rect(cv, x0, y0, x1, y1, r, **kw):
    r = max(1, min(r, (x1 - x0) / 2, (y1 - y0) / 2))
    pts = [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1, x1 - r, y1,
           x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]
    return cv.create_polygon(pts, smooth=True, **kw)


def header(root, title, subtitle=""):
    """漸層標題列(深藍→青綠)＋橘色標記"""
    h = 88 if subtitle else 60
    cv = tk.Canvas(root, height=h, highlightthickness=0, bg=PRIMARY_DK)
    cv.pack(fill=tk.X)

    def draw(e=None):
        cv.delete("all")
        w = max(cv.winfo_width(), 400)
        steps = 60
        for i in range(steps):
            cv.create_rectangle(w * i / steps, 0, w * (i + 1) / steps + 1, h, width=0,
                                fill=_mix("#0D47A1", "#00796B", i / (steps - 1)))
        _round_rect(cv, 20, 16, 26, h - 16, 3, fill=ACCENT, outline="")
        cv.create_text(40, 31 if subtitle else h / 2, text=title, anchor="w", font=FH, fill="white")
        if subtitle:
            cv.create_text(40, 64, text=subtitle, anchor="w", font=FS, fill="#E3F2FD")
    cv.bind("<Configure>", draw)
    return cv


def footer(root):
    bar = tk.Frame(root, bg=CARD, padx=16, pady=10, highlightthickness=1, highlightbackground=BORDER)
    bar.pack(fill=tk.X, side=tk.BOTTOM)
    return bar


def scroll_body(root):
    """可捲動內容區(滑鼠滾輪)，回傳 (body, canvas)"""
    outer = tk.Frame(root, bg=BG)
    outer.pack(fill=tk.BOTH, expand=True)
    canvas = tk.Canvas(outer, highlightthickness=0, bg=BG)
    sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    body = tk.Frame(canvas, bg=BG, padx=16, pady=10)
    body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    win = canvas.create_window((0, 0), window=body, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
    canvas.configure(yscrollcommand=sb.set)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
    return body, canvas


def card(parent, title, hint="", theme="blue", radius=14):
    """圓角卡片(同色系淡底＋邊框，標題用主色)；回傳內容 frame。content.theme = (主色, 淡底, 邊框)"""
    main, tint, edge = THEMES.get(theme, THEMES["blue"])
    pbg = parent.cget("bg")
    cv = tk.Canvas(parent, bg=pbg, highlightthickness=0, height=40)
    cv.pack(fill=tk.X, pady=7)
    pad = radius
    inner = tk.Frame(cv, bg=tint)
    win = cv.create_window(pad + 4, pad - 2, window=inner, anchor="nw")
    head = tk.Frame(inner, bg=tint)
    head.pack(fill=tk.X, pady=(0, 8))
    bar = tk.Canvas(head, width=6, height=24, bg=tint, highlightthickness=0)
    _round_rect(bar, 0, 0, 6, 24, 3, fill=main, outline="")
    bar.pack(side=tk.LEFT, padx=(0, 10))
    tk.Label(head, text=title, font=FT, fg=main, bg=tint).pack(side=tk.LEFT)
    if hint:
        tk.Label(head, text=hint, font=FS, fg=SUB, bg=tint).pack(side=tk.LEFT, padx=(12, 0))
    content = tk.Frame(inner, bg=tint)
    content.pack(fill=tk.X)
    content.theme = (main, tint, edge)

    def fit(e=None):
        w = cv.winfo_width()
        if w < 50:
            return
        cv.itemconfigure(win, width=w - 2 * pad - 8)
        hgt = inner.winfo_reqheight() + 2 * pad - 4
        if int(float(cv.cget("height"))) != hgt:
            cv.configure(height=hgt)
        cv.delete("bg")
        _round_rect(cv, 1, 1, w - 2, hgt - 2, radius, fill=tint, outline=edge, width=1.5, tags="bg")
        cv.tag_lower("bg")
    inner.bind("<Configure>", fit)
    cv.bind("<Configure>", fit)
    return content


def note(parent, text, kind="info", wrap=900):
    if kind == "warn":
        f = tk.Frame(parent, bg=WARN_BG, padx=14, pady=8, highlightthickness=1, highlightbackground="#F5C27A")
        f.pack(fill=tk.X, pady=4)
        tk.Label(f, text="⚠  " + text, font=F, fg=WARN_FG, bg=WARN_BG, justify="left", wraplength=wrap).pack(anchor="w")
    else:
        tk.Label(parent, text="ⓘ  " + text, font=FS, fg=SUB, bg=parent.cget("bg"), justify="left",
                 wraplength=wrap).pack(anchor="w", pady=1)


def entry(parent, var, width=20, **grid):
    e = tk.Entry(parent, textvariable=var, font=F, width=width, fg=TEXT, bg="white", relief="flat",
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=PRIMARY, insertbackground=PRIMARY)
    if grid:
        e.grid(ipady=5, **grid)
    return e


def label(parent, text, bold=False, color=TEXT, **kw):
    return tk.Label(parent, text=text, font=FB if bold else F, fg=color, bg=parent.cget("bg"), **kw)


class RoundButton(tk.Canvas):
    """圓角按鈕(畫在 Canvas 上)。kind: primary/accent/ghost/chip；color 可指定主色(chip 用同色系)"""

    def __init__(self, parent, text, command, kind="primary", color=None, font=None, radius=None, padx=None, pady=None):
        pbg = parent.cget("bg")
        main = color or PRIMARY
        if kind == "primary":
            fill, hover, fg = main, _mix(main, "#000000", 0.18), "white"
        elif kind == "accent":
            fill, hover, fg = ACCENT, "#D9731A", "white"
        elif kind == "ghost":
            fill, hover, fg = "#E3E8EF", "#D3DAE4", TEXT
        else:   # chip：同色系淡底＋主色字
            fill, hover, fg = _mix(main, "#FFFFFF", 0.86), _mix(main, "#FFFFFF", 0.74), main
        self.font = font or (FB if kind in ("primary", "accent") else F)
        f = tkfont.Font(font=self.font)
        px = padx if padx is not None else (24 if kind in ("primary", "accent", "ghost") else 14)
        py = pady if pady is not None else (10 if kind in ("primary", "accent", "ghost") else 5)
        w, h = f.measure(text) + px * 2, f.metrics("linespace") + py * 2
        super().__init__(parent, width=w, height=h, bg=pbg, highlightthickness=0, cursor="hand2")
        r = radius if radius is not None else h / 2
        self._shape = _round_rect(self, 1, 1, w - 1, h - 1, r, fill=fill, outline="")
        self._text = self.create_text(w / 2, h / 2, text=text, font=self.font, fill=fg)
        self.fill, self.hover, self.command = fill, hover, command
        self.bind("<Enter>", lambda e: self.itemconfigure(self._shape, fill=self.hover))
        self.bind("<Leave>", lambda e: self.itemconfigure(self._shape, fill=self.fill))
        self.bind("<Button-1>", lambda e: self.command())

    def invoke(self):
        self.command()

    def cget(self, key):
        if key == "text":
            return self.itemcget(self._text, "text")
        return super().cget(key)


def button(parent, text, command, kind="primary", color=None, **kw):
    return RoundButton(parent, text, command, kind, color, **kw)


def segmented(parent, var, options, color=None):
    """圓角分段切換鈕：選中＝主色實心白字，其餘白底主色框"""
    main = color or PRIMARY
    pbg = parent.cget("bg")
    wrap = tk.Frame(parent, bg=pbg)
    btns = {}
    f_on, f_off = tkfont.Font(font=FB), tkfont.Font(font=F)
    for i, opt in enumerate(options):
        w = max(f_on.measure(opt), f_off.measure(opt)) + 40
        h = f_on.metrics("linespace") + 12
        cv = tk.Canvas(wrap, width=w, height=h, bg=pbg, highlightthickness=0, cursor="hand2")
        cv.grid(row=0, column=i, padx=(0 if i == 0 else 6, 0))
        cv.bind("<Button-1>", lambda e, o=opt: var.set(o))
        btns[opt] = (cv, w, h)

    def paint(*_):
        for opt, (cv, w, h) in btns.items():
            on = var.get() == opt
            cv.delete("all")
            _round_rect(cv, 1, 1, w - 1, h - 1, h / 2, fill=main if on else "white", outline=main, width=1.5)
            cv.create_text(w / 2, h / 2, text=opt, font=FB if on else F, fill="white" if on else main)
    var.trace_add("write", paint)
    paint()
    return wrap


def radio(parent, text, var, value):
    return tk.Radiobutton(parent, text=text, variable=var, value=value, font=F, fg=TEXT, bg=parent.cget("bg"),
                          activebackground=parent.cget("bg"), selectcolor="white", highlightthickness=0,
                          cursor="hand2", padx=4)


def check(parent, text, var):
    return tk.Checkbutton(parent, text=text, variable=var, font=F, fg=TEXT, bg=parent.cget("bg"),
                          activebackground=parent.cget("bg"), selectcolor="white", highlightthickness=0,
                          cursor="hand2", anchor="w", justify="left")


def repaint(widget):
    """把預設灰底的 Label/Frame/Check/Radio/LabelFrame 改成跟父層同底色(套舊介面用)"""
    default = {"SystemButtonFace", "#f0f0f0", "#F0F0F0", "#d9d9d9"}
    for w in widget.winfo_children():
        cls = w.winfo_class()
        try:
            pbg = w.master.cget("bg")
            if cls in ("Label", "Frame", "Labelframe", "Checkbutton", "Radiobutton") and w.cget("bg") in default:
                w.configure(bg=pbg)
                if cls in ("Checkbutton", "Radiobutton"):
                    w.configure(activebackground=pbg, selectcolor="white", highlightthickness=0)
                if cls == "Labelframe":
                    w.configure(relief="flat", bd=0)
        except tk.TclError:
            pass
        repaint(w)


def restyle_inputs(widget):
    """舊介面的 Entry 改成扁平白底細框、聚焦變藍；勾選框白底"""
    for w in widget.winfo_children():
        cls = w.winfo_class()
        try:
            if cls == "Entry":
                w.configure(relief="flat", bd=0, highlightthickness=1, highlightbackground=BORDER,
                            highlightcolor=PRIMARY, bg="white", fg=TEXT, insertbackground=PRIMARY)
                mgr = w.winfo_manager()
                if mgr == "grid":
                    w.grid_configure(ipady=4)
                elif mgr == "pack":
                    w.pack_configure(ipady=4)
            elif cls in ("Checkbutton", "Radiobutton"):
                w.configure(selectcolor="white", cursor="hand2", highlightthickness=0)
        except tk.TclError:
            pass
        restyle_inputs(w)


def date_field(parent, root, yvar, mvar, dvar, color=None, start_from=None, width=18):
    """日期欄：顯示框＋📅選日期＋清除；點了跳出月曆(民國年)。
    yvar/mvar/dvar：民國年/月/日 StringVar；start_from()：空白時月曆從哪個 (民國年, 月) 開始翻"""
    import calendar
    from datetime import date
    main = color or PRIMARY
    dark = _mix(main, "#000000", 0.25)
    pbg = parent.cget("bg")
    fr = tk.Frame(parent, bg=pbg)
    lb = tk.Label(fr, font=FB, width=width, anchor="w", bg="white", padx=10, pady=5, cursor="hand2",
                  highlightthickness=1, highlightbackground=BORDER)
    lb.pack(side=tk.LEFT)

    def show(*_):
        y, m, d = yvar.get(), mvar.get(), dvar.get()
        ok = bool(y and m and d)
        try:
            txt = f"{int(y)} 年 {int(m):02d} 月 {int(d):02d} 日" if ok else "點這裡選日期"
        except ValueError:
            txt, ok = f"{y} 年 {m} 月 {d} 日", True
        lb.config(text=txt, fg=dark if ok else SUB)

    def open_cal():
        cur = [yvar.get(), mvar.get(), dvar.get()]
        today = date.today()
        if not cur[0] and start_from:
            sf = start_from()
            if sf and sf[0]:
                cur = [str(sf[0]), str(sf[1]), ""]
        try:
            state = {"y": int(cur[0]) + 1911 if cur[0] else today.year, "m": int(cur[1]) if cur[1] else today.month}
            sel = (int(cur[0]) + 1911, int(cur[1]), int(cur[2])) if all(cur) else None
        except ValueError:
            state, sel = {"y": today.year, "m": today.month}, None
        top = tk.Toplevel(root, bg=CARD)
        top.title("選擇日期")
        top.transient(root)
        top.grab_set()
        top.resizable(False, False)
        head = tk.Frame(top, bg=dark, padx=10, pady=10)
        head.pack(fill=tk.X)
        title = tk.Label(head, font=FT, fg="white", bg=dark)
        grid = tk.Frame(top, bg=CARD, padx=14, pady=10)
        grid.pack()

        def nav(text, dm):
            b = tk.Label(head, text=text, font=FB, fg="white", bg=dark, padx=8, cursor="hand2")
            b.bind("<Button-1>", lambda e: step(dm))
            b.bind("<Enter>", lambda e: b.config(bg=main))
            b.bind("<Leave>", lambda e: b.config(bg=dark))
            return b

        def draw():
            for w in grid.winfo_children():
                w.destroy()
            y, m = state["y"], state["m"]
            title.config(text=f"民國 {y - 1911} 年  {m} 月")
            for c, wd in enumerate("日一二三四五六"):
                tk.Label(grid, text=wd, font=FB, bg=CARD, width=3,
                         fg=BAD if c in (0, 6) else SUB).grid(row=0, column=c, pady=(0, 6))
            for r, week in enumerate(calendar.Calendar(firstweekday=6).monthdayscalendar(y, m), 1):
                for c, dd in enumerate(week):
                    if not dd:
                        continue
                    is_sel = sel == (y, m, dd)
                    is_today = (y, m, dd) == (today.year, today.month, today.day)
                    bg = main if is_sel else CARD
                    fg = "white" if is_sel else (BAD if c in (0, 6) else TEXT)
                    cell = tk.Label(grid, text=str(dd), font=FB if (is_sel or is_today) else F, width=3, pady=6,
                                    bg=bg, fg=fg, cursor="hand2", highlightthickness=2,
                                    highlightbackground=ACCENT if is_today and not is_sel else bg)
                    cell.grid(row=r, column=c, padx=2, pady=2)
                    cell.bind("<Button-1>", lambda e, dd=dd: pick(dd))
                    if not is_sel:
                        hov = _mix(main, "#FFFFFF", 0.85)
                        cell.bind("<Enter>", lambda e, w=cell: w.config(bg=hov))
                        cell.bind("<Leave>", lambda e, w=cell, b=bg: w.config(bg=b))

        def step(dm):
            m = state["m"] + dm
            state["y"] += (m - 1) // 12
            state["m"] = (m - 1) % 12 + 1
            draw()

        def pick(dd):
            yvar.set(str(state["y"] - 1911))
            mvar.set(str(state["m"]))
            dvar.set(str(dd))
            top.destroy()

        nav("«", -12).pack(side=tk.LEFT)
        nav("‹", -1).pack(side=tk.LEFT)
        title.pack(side=tk.LEFT, expand=True, padx=16)
        nav("»", 12).pack(side=tk.RIGHT)
        nav("›", 1).pack(side=tk.RIGHT)
        ft = tk.Frame(top, bg=CARD, pady=10)
        ft.pack()
        button(ft, "回到今天", lambda: (state.update(y=today.year, m=today.month), draw()), "chip", main).pack(side=tk.LEFT, padx=6)
        button(ft, "取消", top.destroy, "chip", main).pack(side=tk.LEFT, padx=6)
        draw()
        top.update_idletasks()
        # 開在該日期項目(整列：日期框＋按鈕)的右側、上下置中；右邊放不下改左側，上下不超出螢幕
        tw, th = top.winfo_reqwidth(), top.winfo_reqheight()
        sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
        x = fr.winfo_rootx() + fr.winfo_width() + 12
        if x + tw > sw - 8:
            x = max(8, lb.winfo_rootx() - tw - 12)
        y = fr.winfo_rooty() + fr.winfo_height() // 2 - th // 2
        y = max(8, min(y, sh - th - 60))          # 60：留給工作列
        top.geometry(f"+{x}+{y}")
        top.update_idletasks()
        dy = top.winfo_rooty() - y                 # 視窗標題列高度：內容區要置中，外框得再往上移
        if dy > 0:
            top.geometry(f"+{x}+{max(8, y - dy)}")

    lb.bind("<Button-1>", lambda e: open_cal())
    lb.bind("<Enter>", lambda e: lb.config(highlightbackground=main))
    lb.bind("<Leave>", lambda e: lb.config(highlightbackground=BORDER))
    button(fr, "📅 選日期", open_cal, "chip", main).pack(side=tk.LEFT, padx=(10, 4))
    button(fr, "清除", lambda: [v.set("") for v in (yvar, mvar, dvar)], "chip", main).pack(side=tk.LEFT, padx=4)
    for v in (yvar, mvar, dvar):
        v.trace_add("write", show)
    show()
    fr.open = open_cal
    return fr
