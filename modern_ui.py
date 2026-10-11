# -*- coding: utf-8 -*-
"""主介面現代化外觀(macOS 風)：色票、圓角按鈕、卡片、深色訊息面板、膠囊輸入框。

高解析度/縮放：主程式以「物理像素」運作(SetProcessDpiAwareness(2))，所有像素尺寸都乘上 S(=DPI 倍率，
主程式啟動時 set_scale)；按鈕/輸入框高度一律由字型實際行高推算，所以字級怎麼變都不會擠壞。

ModernButton 相容 tk.Button 常用介面(config/configure/cget/invoke、text/bg/fg/state/command/font…)，
主程式原本 button.config(text=..., state=...)、button.config(**舊樣式dict) 等寫法都不用改；
舊樣式的 bg 顏色會自動換成新外觀(實心 / 淡彩 / 白底)。"""
import colorsys
import tkinter as tk
import tkinter.font as tkfont

# ---------- 色票(Apple 系統色＋淺灰底) ----------
BG = "#E3E7ED"            # 視窗底色(冷灰，讓白卡與彩色按鈕浮出來)
CARD = "#FFFFFF"
LINE = "#D1D1D6"          # 細框線
INK = "#1C1C1E"           # 主要文字
INK2 = "#3A3A3C"          # 次要文字(不用淡灰)
BLUE = "#0A6CDB"
CONSOLE_BG = "#1E1F24"    # 訊息面板(深色，像 Terminal)
CONSOLE_HEAD = "#2A2B31"
CONSOLE_FG = "#F2F2F7"

S = 1.0                   # DPI 倍率(主程式啟動時設定)
ON_PRESS = None           # 任一功能鈕按下前先呼叫(主程式用來自動關掉萬用查詢面板)


def set_scale(s):
    global S
    S = max(1.0, float(s or 1.0))


def px(v):
    return int(round(v * S))


# 主頁按鈕的舊 bg → 新外觀(樣式, 主色)；沒列到的依顏色自動換算
_LEGACY = {
    "#FF8C00": ("solid", "#FF8A00"),   # 地籍便民(橘，最優先)
    "#9C27B0": ("solid", "#A54BD6"),   # 謄本與土增稅(紫)
    "#FFCDD2": ("tint", "#E5484D"),    # 土地使用分區(紅)
    "#FFF3CD": ("tint", "#C27C00"),    # 實價登錄(琥珀)
    "#CE93D8": ("tint", "#9B4DCA"),    # 編輯物件資料(紫)
    "#87CEEB": ("tint", "#0A84FF"),    # 連續模式(藍)
    "#81C784": ("tint", "#248A3D"),    # 功能分頁(綠)
    "#FF6B6B": ("run", "#E5484D"),     # 停止連續模式
    "#E0E0E0": ("plain", INK),
}
_BACK_TEXTS = ("回至主頁", "回主選單", "回主頁")
BACK_COLOR = "#5B6472"       # 回主頁：灰藍色系

# 功能色系(給原本白底的按鈕)
INDIGO, TEAL, CYAN, BROWN, SLATE, PINK, AMBER, RED, GREEN, PURPLE = (
    "#5856D6", "#0F8F84", "#0A7EA4", "#9A6B3F", "#4E6A8A", "#C2417A", "#C27C00", "#E5484D", "#248A3D", "#9B4DCA")


def mix(c1, c2, t):
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02X%02X%02X" % tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def _hex(c, widget=None):
    """'white'/'black' 這類色名轉 #RRGGBB(混色要用)"""
    if not c:
        return INK
    if c.startswith("#") and len(c) == 7:
        return c.upper()
    try:
        w = widget or tk._default_root
        r, g, b = w.winfo_rgb(c)
        return "#%02X%02X%02X" % (r // 256, g // 256, b // 256)
    except Exception:
        return INK


def _deepen(c):
    """淡色(舊的粉彩鈕) → 同色相的飽和主色"""
    r, g, b = (int(c[i:i + 2], 16) / 255 for i in (1, 3, 5))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    r, g, b = colorsys.hsv_to_rgb(h, max(0.7, s), 0.62)
    return "#%02X%02X%02X" % (int(r * 255), int(g * 255), int(b * 255))


def _luma(c):
    r, g, b = (int(c[i:i + 2], 16) for i in (1, 3, 5))
    return 0.299 * r + 0.587 * g + 0.114 * b


def round_rect(cv, x0, y0, x1, y1, r, **kw):
    r = max(0, min(r, (x1 - x0) / 2, (y1 - y0) / 2))
    pts = [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1, x1 - r, y1,
           x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]
    return cv.create_polygon(pts, smooth=True, **kw)


class ModernButton(tk.Canvas):
    """圓角功能鈕。page=True 會在右側畫「›」(代表會打開分頁)。
    同一頁的按鈕可用 uniform_group() 綁成一組：字太長要縮時整組一起縮，字級一致"""

    def __init__(self, parent, text="", command=None, font=None, bg="#E0E0E0", fg="black", state=tk.NORMAL,
                 page=False, **legacy):
        super().__init__(parent, bg=parent.cget("bg"), highlightthickness=0, bd=0, cursor="hand2", width=px(60))
        self._o = {"text": text, "command": command, "font": font or ("Microsoft JhengHei", 14),
                   "bg": bg, "fg": fg, "state": state}
        self.page = page
        self.color = None            # 指定色系(白底按鈕改成這個顏色的淡彩)
        self.filled = False          # True＝用 color 畫實心白字(同組裡的主要按鈕)
        self.group = None
        self._hover = self._press = False
        self._fonts = {}
        self._apply_height()
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

    # ---- tk.Button 相容介面 ----
    def configure(self, cnf=None, **kw):
        if cnf:
            kw.update(cnf)
        changed = False
        for k in ("text", "command", "font", "bg", "fg", "state", "background", "foreground"):
            if k in kw:
                key = {"background": "bg", "foreground": "fg"}.get(k, k)
                self._o[key] = kw.pop(k)
                changed = True
        for k in ("width", "height", "relief", "borderwidth", "bd", "activebackground", "activeforeground",
                  "padx", "pady", "anchor", "justify", "wraplength", "overrelief", "disabledforeground"):
            kw.pop(k, None)                         # 舊樣式欄位：新外觀自己處理，忽略
        if kw:
            super().configure(**kw)
        if changed:
            self._fonts.clear()
            self._apply_height()
            super().configure(cursor="hand2" if self._o["state"] != tk.DISABLED else "arrow")
            if self.group:
                self.group.refresh()
            else:
                self._draw()

    config = configure

    def cget(self, key):
        key = {"background": "bg", "foreground": "fg"}.get(key, key)
        if key in self._o:
            return self._o[key]
        return super().cget(key)

    def __getitem__(self, key):
        return self.cget(key)

    def __setitem__(self, key, value):
        self.configure(**{key: value})

    def invoke(self):
        if self._o["state"] != tk.DISABLED and self._o["command"]:
            if ON_PRESS:
                try:
                    ON_PRESS()
                except Exception:
                    pass
            return self._o["command"]()

    # ---- 尺寸 ----
    def _base_font(self):
        return tkfont.Font(font=self._o["font"])

    def _apply_height(self):
        """高度＝字型行高 × 1.9(字級/DPI/系統字型倍率怎麼變都跟著變)"""
        try:
            h = int(self._base_font().metrics("linespace") * 1.9) + px(3)
            if int(float(super().cget("height"))) != h:
                super().configure(height=h)
        except tk.TclError:
            pass

    def _display_text(self):
        kind, _ = self._style()
        t = self._o["text"] or ""
        if kind == "run":
            return "■ " + t
        if t in _BACK_TEXTS:
            return "‹ " + t
        return t

    def _room(self):
        w = self.winfo_width()
        return w - px(16) - (px(18) if self.page else 0)

    def fit_size(self):
        """不被切掉的最大字級(最少 9)"""
        base = self._base_font()
        size = abs(base.actual("size"))
        text = self._display_text()
        room = self._room()
        fam, weight = base.actual("family"), "bold"
        while size > 9:
            if tkfont.Font(family=fam, size=size, weight=weight).measure(text) <= room:
                break
            size -= 1
        return size

    def _font_for(self, size):
        if size not in self._fonts:                 # 要留住參照，否則字型被回收會變回預設字
            base = self._base_font()
            self._fonts[size] = tkfont.Font(family=base.actual("family"), size=size, weight="bold")
        return self._fonts[size]

    # ---- 外觀 ----
    def _style(self):
        bg = _hex(self._o["bg"] or "#E0E0E0", self)
        text = self._o["text"] or ""
        if text.startswith("關閉") or text.startswith("停止"):
            return "run", "#E5484D"
        if text in _BACK_TEXTS:
            return "tint", BACK_COLOR
        if self.color:
            return ("solid" if self.filled else "tint"), self.color
        if bg in _LEGACY:
            return _LEGACY[bg]
        if _luma(bg) < 150:                          # 舊的深色鈕(白字)→ 實心
            return "solid", bg
        if abs(int(bg[1:3], 16) - int(bg[3:5], 16)) < 12 and abs(int(bg[3:5], 16) - int(bg[5:7], 16)) < 12:
            return "plain", INK                      # 灰色系 → 白底
        return "tint", _deepen(bg)                   # 粉彩 → 淡彩

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10:
            return
        pbg = _hex(self.master.cget("bg"), self)
        if _hex(super().cget("bg"), self) != pbg:
            super().configure(bg=pbg)
        kind, main = self._style()
        disabled = self._o["state"] == tk.DISABLED
        if kind in ("solid", "run"):
            fill, fg, edge = main, "#FFFFFF", ""
            if self._hover:
                fill = mix(main, "#000000", 0.12)
        elif kind == "tint":
            fill, fg, edge = mix(main, "#FFFFFF", 0.72), mix(main, "#000000", 0.52), mix(main, "#FFFFFF", 0.38)
            if self._hover:
                fill = mix(main, "#FFFFFF", 0.58)
        else:
            fill, fg, edge = CARD, INK, LINE
            if self._hover:
                fill = "#F5F5F7"
        if self._press and not disabled:
            fill = mix(fill, "#000000", 0.10)
        if disabled:                                  # 停用：整顆淡掉
            fill, fg = mix(fill, pbg, 0.55), mix(fg, pbg, 0.55)
            edge = mix(edge, pbg, 0.55) if edge else ""
        r = px(10)
        sh = max(1, px(2))
        if not disabled:                              # 柔和陰影
            round_rect(self, 2, sh + 1, w - 2, h - 1, r, fill=mix(pbg, "#000000", 0.13), outline="")
        round_rect(self, 1, 1, w - 2, h - sh - 1, r, fill=fill, outline=edge, width=1 if edge else 0)
        size = self.group.size if self.group and self.group.size and kind != "run" else self.fit_size()
        text = self._display_text()
        f = self._font_for(size)
        if f.measure(text) > self._room():           # 組內字級仍放不下(極窄)時自己再縮
            f = self._font_for(self.fit_size())
        show_chev = self.page and kind != "run" and f.measure(text) + px(30) < w
        cx = w / 2 - (px(7) if show_chev else 0)
        cy = (h - sh) / 2
        self.create_text(cx, cy, text=text, font=f, fill=fg)
        if show_chev:
            self.create_text(w - px(14), cy, text="›", font=self._font_for(int(size * 1.15)),
                             fill=mix(fg, fill, 0.25))

    def _set_hover(self, on):
        self._hover = on and self._o["state"] != tk.DISABLED
        self._draw()

    def _on_press(self, e):
        if self._o["state"] != tk.DISABLED:
            self._press = True
            self._draw()

    def _on_release(self, e):
        was = self._press
        self._press = False
        self._draw()
        if was and 0 <= e.x <= self.winfo_width() and 0 <= e.y <= self.winfo_height():
            self.invoke()


class _Group:
    def __init__(self, buttons):
        self.buttons = [b for b in buttons if isinstance(b, ModernButton)]
        self.size = None
        for b in self.buttons:
            b.group = self
            b.bind("<Configure>", lambda e: self.refresh(), add="+")

    def refresh(self):
        # 執行中那顆(■ 關閉…)字比較長，讓它自己縮，不拖累整頁字級
        live = [b for b in self.buttons if b.winfo_ismapped() and b.winfo_width() > 20 and b._style()[0] != "run"]
        if live:
            self.size = min(b.fit_size() for b in live)
        for b in self.buttons:
            b._draw()


def uniform_group(buttons):
    """同一頁的按鈕綁一組：字級取「大家都放得下」的最大值，整頁一致"""
    return _Group(buttons)


class IconButton(tk.Canvas):
    """標題卡右側的小圓鈕(說明/回饋/更新)。相容主程式原本的 .config(bg=...)：
    紅色系＝有新版(右上紅點)、較深的顏色＝滑鼠移上去"""
    _HOVER = {"#1976D2", "#D32F2F", "#F57C00", "#388E3C"}
    _RED = {"#D32F2F", "#F44336"}

    def __init__(self, parent, text="", command=None, color=BLUE, size=None, **legacy):
        size = size or px(30)
        super().__init__(parent, width=size, height=size, bg=parent.cget("bg"), highlightthickness=0, cursor="hand2")
        self.glyph, self.command, self.color, self.size = text, command, color, size
        self.dot = False
        self._hover = False
        self.bind("<Button-1>", self._click)
        self._draw()

    def _click(self, e=None):
        if ON_PRESS:
            try:
                ON_PRESS()
            except Exception:
                pass
        if self.command:
            self.command()

    def configure(self, cnf=None, **kw):
        if cnf:
            kw.update(cnf)
        bg = kw.pop("bg", kw.pop("background", None))
        if bg:
            b = _hex(bg, self)
            self._hover = b in self._HOVER
            self.dot = b in self._RED
        if "command" in kw:
            self.command = kw.pop("command")
        if "text" in kw:
            self.glyph = kw.pop("text")
        for k in ("fg", "foreground", "font", "relief", "bd", "borderwidth", "padx", "pady", "activebackground"):
            kw.pop(k, None)
        if kw:
            super().configure(**kw)
        self._draw()

    config = configure

    def cget(self, key):
        if key == "text":
            return self.glyph
        return super().cget(key)

    def _draw(self):
        self.delete("all")
        s = self.size
        self.create_oval(1, 1, s - 1, s - 1, fill=mix(self.color, "#FFFFFF", 0.74 if self._hover else 0.87), outline="")
        self.create_text(s / 2, s / 2, text=self.glyph, font=("Segoe UI Emoji", max(8, int(s * 0.30 / S))), fill=self.color)
        if self.dot:
            d = px(9)
            self.create_oval(s - d - 1, 1, s - 1, d + 1, fill="#FF3B30", outline="white", width=max(1, px(1.5)))


class Card(tk.Canvas):
    """圓角白卡；內容放 .inner"""

    def __init__(self, parent, radius=None, fill=CARD, pad=None, edge=LINE, accent=None):
        super().__init__(parent, bg=parent.cget("bg"), highlightthickness=0, height=px(30))
        self.radius, self.fill, self.pad, self.edge = radius or px(12), fill, pad or px(9), edge
        self.accent = accent                       # 左側色條(醒目用)
        self._lx = self.pad + (px(8) if accent else 0)
        self.inner = tk.Frame(self, bg=fill)
        self._win = self.create_window(self._lx, self.pad, window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._fit)
        self.bind("<Configure>", self._fit)

    def _fit(self, e=None):
        w = self.winfo_width()
        if w < 30:
            return
        self.itemconfigure(self._win, width=w - self._lx - self.pad)
        h = self.inner.winfo_reqheight() + self.pad * 2
        if int(float(self.cget("height"))) != h:
            self.configure(height=h)
        self.delete("bg")
        round_rect(self, 1, 2, w - 1, h, self.radius, fill=mix(_hex(self.cget("bg"), self), "#000000", 0.05),
                   outline="", tags="bg")
        round_rect(self, 1, 1, w - 2, h - 2, self.radius, fill=self.fill, outline=self.edge, tags="bg")
        if self.accent:
            bw = px(6)
            round_rect(self, px(6), px(8), px(6) + bw, h - px(9), bw / 2, fill=self.accent, outline="", tags="bg")
        self.tag_lower("bg")


class ConsolePanel(tk.Frame):
    """深色訊息面板：上方一條 macOS 風標題列(紅黃綠三點＋標題)，內容放 .body"""

    def __init__(self, parent, title="訊息"):
        super().__init__(parent, bg=parent.cget("bg"))
        self.cv = tk.Canvas(self, bg=parent.cget("bg"), highlightthickness=0)
        self.cv.pack(fill=tk.BOTH, expand=True)
        self.title = title
        self.body = tk.Frame(self.cv, bg=CONSOLE_BG)
        self.head_h = max(px(24), tkfont.Font(family="Microsoft JhengHei", size=10, weight="bold").metrics("linespace") + px(8))
        self._win = self.cv.create_window(px(8), self.head_h + px(4), window=self.body, anchor="nw")
        self.cv.bind("<Configure>", self._draw)

    def _draw(self, e=None):
        cv = self.cv
        w, h = cv.winfo_width(), cv.winfo_height()
        hh, r = self.head_h, px(11)
        cv.delete("bg")
        round_rect(cv, 0, 0, w - 1, h - 1, r, fill=CONSOLE_BG, outline="", tags="bg")
        round_rect(cv, 0, 0, w - 1, hh, r, fill=CONSOLE_HEAD, outline="", tags="bg")
        cv.create_rectangle(0, hh - r, w - 1, hh, fill=CONSOLE_HEAD, outline="", tags="bg")
        d, gap = px(10), px(16)
        for i, c in enumerate(("#FF5F57", "#FEBC2E", "#28C840")):
            x = px(12) + i * gap
            cv.create_oval(x, hh / 2 - d / 2, x + d, hh / 2 + d / 2, fill=c, outline="", tags="bg")
        cv.create_text(w / 2, hh / 2, text=self.title, font=("Microsoft JhengHei", 10, "bold"), fill="#C7C7CC", tags="bg")
        cv.tag_lower("bg")
        cv.coords(self._win, px(8), hh + px(4))
        cv.itemconfigure(self._win, width=max(10, w - px(16)), height=max(10, h - hh - px(12)))


class PillEntry(tk.Frame):
    """膠囊輸入框外殼：把原本的 tk.Entry 放進來(去邊框)，左邊放大鏡、聚焦時藍色外框。
    高度依輸入框字型行高推算"""

    def __init__(self, parent, on_menu=None):
        super().__init__(parent, bg=parent.cget("bg"))
        self.cv = tk.Canvas(self, bg=parent.cget("bg"), highlightthickness=0, height=px(40))
        self.cv.pack(fill=tk.X)
        self.focused = False
        self.on_menu = on_menu               # 右側「☰ 選單」按鈕(打開地籍/地址選單面板)
        self._menu_hover = False
        self._menu_w = 0
        if on_menu:
            self.cv.tag_bind("menu", "<Button-1>", lambda e: self.on_menu())
            self.cv.tag_bind("menu", "<Enter>", lambda e: self._mh(True))
            self.cv.tag_bind("menu", "<Leave>", lambda e: self._mh(False))
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.entry = None

    def attach(self, entry):
        self.entry = entry
        entry.configure(relief="flat", bd=0, highlightthickness=0, bg=CARD)
        h = int(tkfont.Font(font=entry.cget("font")).metrics("linespace") * 1.75)
        self.cv.configure(height=h)
        self._win = self.cv.create_window(h * 0.85, h / 2, window=entry, anchor="w")
        entry.lift()        # 輸入框比外框早建立，堆疊在下面會被外框蓋住(字打得進去但看不到)→ 提到最上層
        entry.bind("<FocusIn>", lambda e: self._focus(True), add="+")
        entry.bind("<FocusOut>", lambda e: self._focus(False), add="+")
        self._draw()

    def _focus(self, on):
        self.focused = on
        self._draw()

    def _mh(self, on):
        self._menu_hover = on
        self.cv.configure(cursor="hand2" if on else "")
        self._draw()

    def _draw(self):
        cv = self.cv
        w, h = cv.winfo_width(), cv.winfo_height()
        if w < 30:
            return
        cv.delete("bg")
        ring = max(2, px(3))
        if self.focused:
            round_rect(cv, 0, 0, w - 1, h - 1, h / 2, fill=mix(BLUE, "#FFFFFF", 0.7), outline="", tags="bg")
        round_rect(cv, ring, ring, w - ring - 1, h - ring - 1, (h - ring * 2) / 2, fill=CARD,
                   outline=BLUE if self.focused else LINE, width=2 if self.focused else 1, tags="bg")
        cv.create_text(h * 0.48, h / 2, text="🔎", font=("Segoe UI Emoji", max(9, int(h * 0.26 / S))), fill=INK2, tags="bg")
        cv.delete("menu")
        self._menu_w = 0
        if self.on_menu:                     # 右側膠囊鈕「☰ 選單」
            f = tkfont.Font(family="Microsoft JhengHei", size=max(9, int(h * 0.20 / S)), weight="bold")
            label = "☰ 選單"
            bw = f.measure(label) + h * 0.6
            x1 = w - ring - h * 0.18
            x0 = x1 - bw
            y0, y1 = h * 0.2, h * 0.8
            fill = mix(BLUE, "#FFFFFF", 0.72 if self._menu_hover else 0.84)
            round_rect(cv, x0, y0, x1, y1, (y1 - y0) / 2, fill=fill, outline="", tags="menu")
            self._menu_font = f
            cv.create_text((x0 + x1) / 2, h / 2, text=label, font=f, fill=mix(BLUE, "#000000", 0.3), tags="menu")
            self._menu_w = bw + h * 0.3
        cv.tag_lower("bg")
        if self.entry is not None:
            cv.coords(self._win, h * 0.85, h / 2)
            cv.itemconfigure(self._win, width=max(10, w - h * 0.85 - h * 0.45 - self._menu_w))


class StatusChip(tk.Label):
    """謄本結構化狀態小標籤。相容主程式原本 json_status_label.config(text=..., fg=...)：
    綠字(#006400)＝已做結構化 → 綠底「✓ 已做謄本結構化：地號」；橘字＝尚未 → 橘底"""

    def __init__(self, parent, font=None, **legacy):
        super().__init__(parent, font=font, padx=px(8), pady=max(1, px(1)), anchor="w", justify="left")
        self._raw = ""
        self.configure(text="尚未謄本結構化", fg="#FF8C00")

    def configure(self, cnf=None, **kw):
        if cnf:
            kw.update(cnf)
        text = kw.pop("text", None)
        fg = kw.pop("fg", kw.pop("foreground", None))
        for k in ("bg", "background", "padx", "pady", "anchor"):
            kw.pop(k, None)
        if text is not None or fg is not None:
            if text is not None:
                self._raw = text
            done = "已做謄本結構化" in self._raw
            if "不符" in self._raw:                   # 謄本與目前案件不符：紅底白字，最醒目
                super().configure(text=f"⚠ {self._raw}", fg="#FFFFFF", bg="#D93025")
            elif done:
                lot = self._raw.replace("←已做謄本結構化", "").strip()
                shown = f"✓ 已做謄本結構化：{lot}" if lot else "✓ 已做謄本結構化"
                super().configure(text=shown, fg="#1B7F3B", bg="#E3F5E8")
            else:
                super().configure(text=f"○ {self._raw}", fg="#B35400", bg="#FFEBD6")
        if kw:
            super().configure(**kw)

    config = configure

    def cget(self, key):
        if key == "text":
            return self._raw
        return super().cget(key)


def paint_bg(widget, color=BG):
    """把一整棵 Frame 的預設灰底換成新底色(只換系統預設灰的 Frame，自訂顏色的不動)"""
    try:
        cls = widget.winfo_class()
    except tk.TclError:
        return
    if cls == "Frame":
        try:
            cur = _hex(widget.cget("bg"), widget)
            if cur in ("#F0F0F0", "#ECECEC", "#D9D9D9") or cur == _hex("SystemButtonFace", widget):
                widget.configure(bg=color)
        except tk.TclError:
            pass
    if cls == "Toplevel":
        return                                       # 對話框不動(第三批再處理)
    for ch in widget.winfo_children():
        paint_bg(ch, color)
