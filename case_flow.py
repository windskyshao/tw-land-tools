"""案件進度表(主頁新版)：照辦一件案子的順序列出步驟，看案件資料夾判斷每一步做到哪，「下一步」最醒目；
阿生(橘貓)在上方用動作＋對話框報目前狀態。
情境：委託＝全部做完呈報公司；開發＝先免費查詢，有希望再調謄本，物件資料等委託後再做(想先做也可以)。
本檔只負責「看進度、畫畫面」；按下按鈕要做的事由主程式傳 actions={key: 函式} 進來。"""
import os
import glob
import json
import math
import tkinter as tk
import tkinter.font as tkfont

import modern_ui as M

try:
    import case_attach as CA
except Exception:                       # 單獨測試畫面時沒有也能畫
    CA = None

ORANGE = "#E07A1F"
ORANGE_DK = "#B35A10"
ORANGE_BG = "#FFF1E3"
GREEN = "#23864B"
GREEN_BG = "#E6F4EA"
RED = "#C5221F"
INK = "#1C1C1E"
INK2 = "#3A3A3C"                         # 說明文字(不用淡灰，要看得清楚)
DIM = "#8A8A8E"                          # 開發情境暫不用做的步驟
FAM = "Microsoft JhengHei"

PHASES = [("查", "免費查詢", "先看看有沒有希望"), ("調", "謄本與計算", "要付費，有希望再調"),
          ("做", "委託資料", "委託後製作"), ("交", "交件", "")]


# ───────────────────────── 看資料夾判斷進度 ─────────────────────────
def _glob(folder, *pats):
    out = []
    for p in pats:
        out += glob.glob(os.path.join(folder, p), recursive=True)
    return [x for x in out if not os.path.basename(x).startswith("~$")]


def _attach_items(case):
    """這個縣市適用的附件項目 [(名稱, 有沒有檔)]（跟連續模式同一份清單）"""
    if not CA:
        return []
    out = []
    for it in CA.ITEMS + [CA.LUZ]:
        if CA.applicable(it, case.get("city", "")):
            out.append((it["name"].split(" ", 1)[-1], bool(CA.found_files(it, case))))
    return out


def avm_ok(account):
    return bool(account) and account.get("tier") in ("supporter", "member")


def scan(case, stage="委託", running=None, account=None):
    """case＝dict(city, area, section, lot, folder, door, deal)；stage＝委託/開發；
    running＝背景正在跑的資訊 dict(key, done, total, wait) 或 None；
    account＝阿生地圖登入者 dict(name, tier=user/supporter/member) 或 None(沒登入)。回傳步驟清單(每步一個 dict)"""
    f = case["folder"]
    dev = stage == "開發"
    steps = []

    def add(key, phase, title, state, detail, actions=(), tip=""):
        steps.append(dict(key=key, phase=phase, title=title, state=state, detail=detail,
                          actions=list(actions), tip=tip))

    # ① 建立案件
    door = case.get("door") or ""
    add("case", "查", "建立案件", "done",
        f"{case.get('section', '')} {case.get('lot', '')} 地號" + (f"・{door}" if door else ""),
        [("📥 整理資料夾", "import", False,
          "自己已經有放檔案的資料夾（裡面有子資料夾也可以）？\n選那個資料夾，阿生會看每個檔案的內容，\n"
          "自動分到謄本、基本資料、照片、行情、委託、成交…（原來的檔案不動）。")],
        tip="換案件：在最下面的輸入框打地址或地段地號。")

    # 阿生地圖(不算進度)：要 Google 登入；行情推估限抖內支持者／會員
    LOGIN = "要先用 Google 登入阿生地圖（右上「🔑 Google 登入」，登入一次就會記住）。"
    AVM_TIP = ("【行情推估】用周邊實價登錄（比較法＋迴歸）推估這間的合理價格，"
               "可以放進給屋主的行情資料。\n（抖內支持者或會員才能使用）")
    if not account:
        lm_acts = [("地圖", "lm_view", False, LOGIN), ("周邊行情", "lm_near", False, LOGIN),
                   ("推估🔒", "lm_avm", False, LOGIN + "\n\n" + AVM_TIP)]
    else:
        lm_acts = [("地圖", "lm_view", False, "在阿生地圖打開這個地號：地籍線、分區、周邊實價、在售物件。"),
                   ("周邊行情", "lm_near", False, "在阿生地圖看這附近的實價登錄成交、在售、出租。"),
                   ("推估" if avm_ok(account) else "推估🔒", "lm_avm", False,
                    AVM_TIP if not avm_ok(account) else "【行情推估】在阿生地圖用周邊實價登錄推估這間的合理價格。")]
    add("lmap", "查", "阿生地圖", "extra", "" if account else "（要登入）", lm_acts,
        tip="阿生地圖的資料直接在這裡查，不用另外開網頁。\n" + ("" if account else LOGIN))

    # ② 附件查詢(免費、自動)
    items = _attach_items(case)
    have = [n for n, ok in items if ok]
    miss = [n for n, ok in items if not ok]
    if running and running.get("key") == "query":
        st = "wait" if running.get("wait") else "run"
        det = f"阿生查詢中 {running.get('done', len(have))}/{running.get('total', len(items))}"
        if running.get("wait"):
            det += f"・{running['wait']}"
        acts = [("停止", "query_stop", False), ("📋 清單", "query_list", False)]
    else:
        st = "done" if items and not miss else "todo"
        det = f"已有 {len(have)}/{len(items)}" + (f"　缺：{'、'.join(miss[:3])}{'…' if len(miss) > 3 else ''}" if miss else "")
        acts = [("▶ 查缺的", "query_run", True), ("📋 清單", "query_list", False)] if miss else [("📋 清單", "query_list", False)]
    add("query", "查", "附件查詢", st, det, acts,
        tip="使用執照、重要設施、V523、地籍便民、國土測繪、都市計畫、國土分區、地質雲、"
            "內政部／104 實價，免費、不用謄本。\n按「▶ 查缺的」會打開連續模式，一項一項自動查，已經有的不重查。\n"
            "「清單」可以看每一項、單獨開某一項手動操作。")

    # ③ 調閱謄本(付費)
    tr = _glob(f, "0.謄本/*.pdf")
    land = any("地號" in os.path.basename(p) and "謄本" in os.path.basename(p) for p in tr)
    bldg = any("建號" in os.path.basename(p) and "謄本" in os.path.basename(p) and "成果圖" not in os.path.basename(p) for p in tr)
    cad = [p for p in tr if any(k in os.path.basename(p) for k in ("地籍圖", "畫線"))]
    got = land or bldg or cad
    det = "　".join(f"{n} {'✓' if ok else '—'}" for n, ok in (("土地", land), ("建物", bldg), ("地籍圖", bool(cad))))
    if not got:
        det = "還沒調閱" + ("（有希望再調）" if dev else "")
    add("trans", "調", "調閱謄本", "done" if got else ("opt" if dev else "todo"), det,
        [("電子謄本", "epaper", not got), ("電傳", "telereport", False)],
        tip="要付費：電子謄本（全國地政）或電傳資訊系統。\n調好後會問你要不要接著做謄本結構化。")

    # ④ 謄本結構化
    js = _glob(f, "4.其他相關/trans_*.json")
    if "struct_ok" in case:                 # 主程式判斷過(目前選用的結構化資料是不是這個案件、有沒有更新的謄本)
        js = case["struct_ok"]
    why = case.get("struct_why") or ""
    if js:
        add("struct", "調", "謄本結構化", "done", "已完成（本案）", [("重做", "struct", False)],
            tip="把謄本 PDF 轉成資料：面積、持分、所有權人、前次移轉現值…\n物件編輯器和土增稅都會用到。")
    else:
        add("struct", "調", "謄本結構化", "lock" if not got else "todo",
            "調閱謄本後再做" if not got else (why if why and "還沒做" not in why else "謄本有了，還沒結構化"),
            [("結構化", "struct", bool(got))], tip="把謄本 PDF 轉成資料：面積、持分、所有權人、前次移轉現值…")

    # ⑤ 土地增值稅
    ft = _glob(f, "1.基本資料/01_土地增值稅_財政部版*")
    kt = _glob(f, "1.基本資料/01-1_土地增值稅_高雄市版*")
    if ft or kt:
        add("tax", "調", "土地增值稅", "done",
            "　".join(x for x in ("財政部 ✓" if ft else "", "高雄市 ✓" if kt else "") if x),
            [("財政部", "tax_fin", False), ("高雄版", "tax_kh", False)], tip="用結構化資料自動試算，不用再手打。")
    else:
        add("tax", "調", "土地增值稅", "lock" if not js else "todo",
            "結構化後就能算" if not js else "可以試算了",
            [("試算", "tax_fin", bool(js))], tip="用結構化資料自動試算（預設財政部版，高雄的案子可以加算高雄市版）。")

    # ⑥ 地籍圖標註
    marked = [p for p in cad if "_標註" in os.path.basename(p)]
    if marked:
        add("map", "調", "地籍圖標註", "done", "已標註（邊長、路名、路寬）", [("開啟工具", "cadastral", False)],
            tip="地籍圖處理工具：標出標的邊長、臨路路名路寬、附近地標。")
    else:
        add("map", "調", "地籍圖標註", "lock" if not cad else "todo",
            "調到地籍圖後才能標" if not cad else "地籍圖有了，還沒標註",
            [("開啟工具", "cadastral", bool(cad))],
            tip="地籍圖處理工具：標出標的邊長、臨路路名路寬、附近地標。")

    # ⑦⑧⑨ 委託資料(開發情境：委託後再做，想先做也可以按)
    later = "委託後再做（想先做也可以）"
    pics = [p for p in _glob(f, "2.照片/*") if p.lower().endswith((".jpg", ".jpeg", ".png", ".heic", ".webp"))]
    add("photo", "做", "現場照片", "done" if pics else ("skip" if dev else "todo"),
        f"{len(pics)} 張" if pics else (later if dev else "拍完直接拖進主程式"),
        [("📂 打開照片資料夾", "folder_photo", True)],
        tip="把照片拖進主程式視窗，選「2.照片」就會放進案件資料夾；行銷簡報會用到。")
    fin = _glob(f, "4.其他相關/data_final.json")
    add("editor", "做", "物件資料", "done" if fin else ("skip" if dev else "todo"),
        "已填寫" if fin else (later if dev else "開物件編輯器填寫"),
        [("物件編輯器", "editor", not fin and not dev)],
        tip="物件編輯器：謄本、分區、使照的資料會自動帶進來，補上售價、格局、車位等。")
    rent = "租" in str(case.get("deal", ""))
    docs = [("物調表", bool(_glob(f, "物件個案調查表*.docx", "物件個案調查表*.pdf")), "doc_survey"),
            ("簡報", bool(_glob(f, "*不動產行銷報告*.pptx", "*不動產行銷報告*.pdf")), "doc_deck"),
            ("業績分配表", bool(_glob(f, "**/業績分配表_*.pdf")), "doc_perf")]
    if rent:
        docs.append(("申報書", bool(_glob(f, "**/實價登錄申報書*.pdf")), "doc_rent"))
    done_n = sum(ok for _, ok, _ in docs)
    add("docs", "做", "物調表・簡報・業績", "done" if done_n == len(docs) else ("skip" if dev and not done_n else "todo"),
        "　".join(f"{n} {'✓' if ok else '○'}" for n, ok, _ in docs) if (done_n or not dev) else later,
        [(n, k, False) for n, ok, k in docs if not ok][:3],
        tip="都從物件編輯器的資料產生；做好的檔案放在案件資料夾。")

    # ⑩ 交件
    merged = _glob(f, "1.基本資料/00_合併列印_*.pdf")
    if dev:
        add("deliver", "交", "給屋主的行情資料", "todo",
            "實價行情為主、基本資料為輔" + ("＋阿生行情推估" if avm_ok(account) else "（加行情推估 🔒）"),
            [("📑 產生", "owner_pack", True), ("📂", "folder", False)],
            tip="內政部／104 實價（3.行情）放前面，分區、地籍圖等基本資料放後面，合成一份給屋主看。\n（目前先用合併列印，自己勾要給屋主看的檔案）\n"
                + ("阿生地圖的行情推估會放在最前面。" if avm_ok(account)
                   else "🔒 抖內支持者或會員可以再加上阿生地圖的「行情推估」頁，更有說服力。"))
    else:
        add("deliver", "交", "呈報公司", "done" if merged else "todo",
            "已合併列印" if merged else "所有做好的檔案合成一份 PDF 上繳",
            [("📑 合併列印", "merge", not merged), ("📂", "folder", False)],
            tip="把謄本、附件、行情、物調表等全部依順序合併成一份 PDF（會先讓你勾選、預覽）。")

    # 下一步＝第一個「該做還沒做」的步驟
    for s in steps:
        if s["state"] == "todo":
            s["state"] = "next"
            break
    return steps


def cat_line(steps, stage):
    """阿生要說的話(短短一句)"""
    by = {s["key"]: s for s in steps}
    q = by.get("query")
    if q and q["state"] == "wait":
        return "wait", "全國分區要你點一下驗證喔！"
    if q and q["state"] == "run":
        return "work", "我在幫你查附件，你可以先去調謄本～"
    nxt = next((s for s in steps if s["state"] == "next"), None)
    if nxt:
        return "idle", {"query": "先免費查查看有沒有希望", "trans": "有希望就調謄本吧", "struct": "謄本到了，我來轉成資料",
                        "tax": "可以算土增稅了", "map": "地籍圖到了，來標邊長和路寬", "photo": "拍完照片拖進來就好",
                        "editor": "打開物件編輯器補一補", "docs": "物調表、簡報、業績分配表一鍵產生",
                        "deliver": "都好了，合成一份交給公司" if stage == "委託" else "產生一份行情資料給屋主看"
                        }.get(nxt["key"], f"下一步：{nxt['title']}")
    return "done", "這件都齊了！辛苦了 🎉"


# ───────────────────────── 小元件 ─────────────────────────
class Tip:
    """滑鼠停 0.45 秒顯示說明(大字、深色字)"""

    def __init__(self, widget, text_fn):
        self.w, self.text_fn, self.win, self.job = widget, text_fn, None, None
        widget.bind("<Enter>", self._enter, add="+")
        widget.bind("<Leave>", self._leave, add="+")
        widget.bind("<ButtonPress>", self._leave, add="+")

    def _enter(self, e=None):
        self.job = self.w.after(450, self.show)

    def _leave(self, e=None):
        if self.job:
            self.w.after_cancel(self.job)
            self.job = None
        if self.win:
            self.win.destroy()
            self.win = None

    def show(self, x=None, y=None):
        text = self.text_fn() if callable(self.text_fn) else self.text_fn
        if not text:
            return
        self.win = tw = tk.Toplevel(self.w)
        tw.overrideredirect(True)
        tw.attributes("-topmost", True)
        fr = tk.Frame(tw, bg="#5C4A32", padx=1, pady=1)
        fr.pack()
        tk.Label(fr, text=text, font=(FAM, 12), fg=INK, bg="#FFF8E7", justify="left", wraplength=M.px(330),
                 padx=M.px(10), pady=M.px(7)).pack()
        tw.update_idletasks()
        if x is None:
            x = self.w.winfo_rootx() + M.px(10)
            y = self.w.winfo_rooty() + self.w.winfo_height() + M.px(4)
        tw.geometry(f"+{x}+{y}")
        return tw


class Pill(tk.Canvas):
    """小圓角按鈕：primary＝橘底白字(下一步)，否則淺底深字"""

    def __init__(self, parent, text, command=None, primary=False, dim=False, size=11):
        self.font = tkfont.Font(family=FAM, size=size, weight="bold")
        w = self.font.measure(text) + M.px(20)
        h = self.font.metrics("linespace") + M.px(8)
        super().__init__(parent, width=w, height=h, bg=parent.cget("bg"), highlightthickness=0, cursor="hand2")
        self.text, self.command, self.primary, self.dim, self.hover = text, command, primary, dim, False
        self.bind("<Enter>", lambda e: self._h(True))
        self.bind("<Leave>", lambda e: self._h(False))
        self.bind("<ButtonRelease-1>", lambda e: self.command and self.command())
        self._draw()

    def _h(self, v):
        self.hover = v
        self._draw()

    def _draw(self):
        self.delete("all")
        w, h = int(self.cget("width")), int(self.cget("height"))
        if self.primary:
            fill, fg, edge = (ORANGE_DK if self.hover else ORANGE), "white", ""
        else:
            fill = "#E6E1DA" if self.hover else "#F2EFEA"
            fg, edge = (DIM if self.dim else INK), "#D6D0C8"
        M.round_rect(self, 1, 1, w - 1, h - 1, h / 2 - 1, fill=fill, outline=edge)
        self.create_text(w / 2, h / 2, text=self.text, font=self.font, fill=fg)


class Asheng(tk.Canvas):
    """阿生：圓形貓照片＋動作。idle 慢慢呼吸、work 晃頭忙碌、wait 跳起來加「!」、done 跳躍加愛心"""

    def __init__(self, parent, img_path, size=None):
        self.size = size or M.px(58)
        super().__init__(parent, width=self.size + M.px(8), height=self.size + M.px(12), bg=parent.cget("bg"),
                         highlightthickness=0)
        self.mode, self.t = "idle", 0
        self.frames = self._load(img_path)
        self.item = self.create_image(self.size / 2 + M.px(4), self.size / 2 + M.px(8), image=self.frames[0])
        self.badge = None
        self._tick()

    def _load(self, path):
        from PIL import Image, ImageDraw, ImageTk
        s = self.size
        im = Image.open(path).convert("RGBA").resize((s * 3, s * 3), Image.LANCZOS)
        mask = Image.new("L", im.size, 0)
        ImageDraw.Draw(mask).ellipse((0, 0, s * 3 - 1, s * 3 - 1), fill=255)
        im.putalpha(mask)
        ring = Image.new("RGBA", im.size, (0, 0, 0, 0))
        ImageDraw.Draw(ring).ellipse((2, 2, s * 3 - 3, s * 3 - 3), outline=(255, 255, 255, 255), width=9)
        im = Image.alpha_composite(im, ring).resize((s, s), Image.LANCZOS)
        self._pil = im
        return [ImageTk.PhotoImage(im.rotate(a, resample=Image.BICUBIC)) for a in (0, 7, -7)]

    def set_mode(self, mode):
        self.mode = mode
        if self.badge:
            self.delete(self.badge)
            self.badge = None
        s = self.size
        if mode == "wait":
            self.badge = self.create_text(s + M.px(2), M.px(10), text="!", font=(FAM, 15, "bold"), fill=RED)
        elif mode == "done":
            self.badge = self.create_text(s + M.px(1), M.px(10), text="♥", font=(FAM, 13, "bold"), fill="#E0398A")

    def _tick(self):
        try:
            self.t += 1
            s, t = self.size, self.t
            cx, cy = s / 2 + M.px(4), s / 2 + M.px(8)
            if self.mode == "work":
                self.itemconfigure(self.item, image=self.frames[1 if (t // 4) % 2 else 2])
                dy = -abs(math.sin(t / 2.0)) * M.px(3)
            elif self.mode == "wait":
                self.itemconfigure(self.item, image=self.frames[0])
                dy = -abs(math.sin(t / 2.2)) * M.px(7)
            elif self.mode == "done":
                self.itemconfigure(self.item, image=self.frames[0])
                dy = -abs(math.sin(t / 3.0)) * M.px(6)
            else:
                self.itemconfigure(self.item, image=self.frames[0])
                dy = math.sin(t / 8.0) * M.px(1.5)
            self.coords(self.item, cx, cy + dy)
            self.after(70, self._tick)
        except tk.TclError:
            pass


class Bubble(tk.Canvas):
    """阿生的對話框(左側有尖角)"""

    def __init__(self, parent):
        super().__init__(parent, bg=parent.cget("bg"), highlightthickness=0, height=M.px(50))
        self.text = ""
        self.font = tkfont.Font(family=FAM, size=12, weight="bold")
        self.bind("<Configure>", lambda e: self._draw())

    def set(self, text):
        self.text = text
        self._draw()

    def _draw(self):
        self.delete("all")
        w = self.winfo_width()
        if w < 40:
            return
        tail = M.px(10)
        tid = self.create_text(tail + M.px(12), M.px(9), text=self.text, font=self.font, fill=INK, anchor="nw",
                               width=w - tail - M.px(26))
        x0, y0, x1, y1 = self.bbox(tid)
        h = y1 + M.px(9)
        self.configure(height=h + 2)
        M.round_rect(self, tail, 1, w - 2, h, M.px(12), fill="#FFFFFF", outline="#E3C9A8", width=1.5, tags="bg")
        my = min(h / 2, M.px(24))
        self.create_polygon(tail + 1, my - M.px(6), 1, my, tail + 1, my + M.px(6), fill="#FFFFFF", outline="#E3C9A8",
                            tags="bg")
        self.create_line(tail + 1, my - M.px(5), tail + 1, my + M.px(5), fill="#FFFFFF", width=2, tags="bg")
        self.tag_lower("bg")


class StatusDot(tk.Canvas):
    """步驟左邊的狀態圓：✓完成 ▶下一步 轉圈＝阿生在跑 !＝等你 ○未做 鎖＝前一步還沒好"""

    def __init__(self, parent, state, num):
        self.d = M.px(26)
        super().__init__(parent, width=self.d + 2, height=self.d + 2, bg=parent.cget("bg"), highlightthickness=0)
        self.state, self.num, self.a = state, num, 0
        self._draw()
        if state == "run":
            self._spin()

    def _draw(self):
        self.delete("all")
        d, st = self.d, self.state
        f = tkfont.Font(family=FAM, size=11, weight="bold")
        if st == "done":
            self.create_oval(1, 1, d, d, fill=GREEN, outline="")
            self.create_text(d / 2 + 1, d / 2 + 1, text="✓", font=f, fill="white")
        elif st == "next":
            self.create_oval(1, 1, d, d, fill=ORANGE, outline="")
            self.create_text(d / 2 + 1, d / 2 + 1, text=self.num, font=f, fill="white")
        elif st in ("run", "wait"):
            self.create_oval(1, 1, d, d, fill=ORANGE_BG, outline=ORANGE, width=1.5)
            if st == "run":
                self.create_arc(3, 3, d - 2, d - 2, start=self.a, extent=100, style="arc", outline=ORANGE, width=3)
                self.create_text(d / 2 + 1, d / 2 + 1, text="🐾", font=("Segoe UI Emoji", 8), fill=ORANGE_DK)
            else:
                self.create_text(d / 2 + 1, d / 2 + 1, text="!", font=f, fill=RED)
        else:
            col = DIM if st in ("skip", "lock") else INK2
            self.create_oval(2, 2, d - 1, d - 1, fill="white", outline="#C9C4BC", width=1.5)
            self.create_text(d / 2 + 1, d / 2 + 1, text=self.num, font=f, fill=col)

    def _spin(self):
        try:
            self.a = (self.a - 24) % 360
            self._draw()
            self.after(60, self._spin)
        except tk.TclError:
            pass


NUMS = "①②③④⑤⑥⑦⑧⑨⑩"


# ───────────────────────── 進度表 ─────────────────────────
TIER_NAME = {"member": "會員", "supporter": "支持者", "user": "一般"}
PHASE_ICON = {"查": "🔍", "調": "📄", "做": "📝", "交": "📑"}


class Tile(tk.Canvas):
    """主頁的大方塊：階段名稱(大字)＋一個狀態(✓ / 2/4 / 轉圈 / !)"""

    def __init__(self, parent, title, sub, state, command, dim=False, hot=False):
        super().__init__(parent, height=M.px(84), bg=parent.cget("bg"), highlightthickness=0, cursor="hand2")
        self.title, self.sub, self.state, self.command, self.dim, self.hot = title, sub, state, command, dim, hot
        self.hover = False
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", lambda e: self._h(True))
        self.bind("<Leave>", lambda e: self._h(False))
        self.bind("<ButtonRelease-1>", lambda e: self.command and self.command())

    def _h(self, v):
        self.hover = v
        self._draw()

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), int(self.cget("height"))
        if w < 20:
            return
        fill = ORANGE_BG if self.hot else ("#F7F5F2" if self.dim else "#FFFFFF")
        if self.hover:
            fill = M.mix(fill, ORANGE, 0.12)
        edge = ORANGE if self.hot else ("#E3DDD5" if self.dim else "#D9CDBE")
        M.round_rect(self, 2, 2, w - 2, h - 2, M.px(14), fill=fill, outline=edge, width=2 if self.hot else 1.5)
        fg = DIM if self.dim else INK
        self.create_text(M.px(16), h * 0.36, text=self.title, anchor="w", font=(FAM, 17, "bold"), fill=fg)
        col = {"done": GREEN, "wait": RED, "run": ORANGE_DK}.get(self.state, DIM if self.dim else INK2)
        self.create_text(M.px(16), h * 0.72, text=self.sub, anchor="w", font=(FAM, 12, "bold"), fill=col)
        cx, cy, r = w - M.px(30), h / 2, M.px(15)
        if self.state == "done":
            self.create_oval(cx - r, cy - r, cx + r, cy + r, fill=GREEN, outline="")
            self.create_text(cx, cy + 1, text="✓", font=(FAM, 14, "bold"), fill="white")
        elif self.state == "wait":
            self.create_oval(cx - r, cy - r, cx + r, cy + r, fill=RED, outline="")
            self.create_text(cx, cy + 1, text="!", font=(FAM, 14, "bold"), fill="white")
        elif self.state == "run":
            self.create_text(cx, cy, text="🐾", font=("Segoe UI Emoji", 15), fill=ORANGE_DK)
        else:
            self.create_text(cx, cy, text="›", font=(FAM, 22, "bold"), fill=DIM)


class CaseFlow(tk.Frame):
    """主頁的案件進度：主頁只有四個大方塊＋「下一步」大按鈕；點方塊進該階段分頁(重跑、單獨開、細節都在分頁)。
    on_action(key) 由主程式處理；on_stage(stage) 切換委託/開發時通知主程式存檔"""

    def __init__(self, parent, img_path, on_action=None, on_stage=None):
        super().__init__(parent, bg=M.BG)
        self.on_action, self.on_stage = on_action or (lambda k: None), on_stage or (lambda s: None)
        self.card = M.Card(self, fill="#FFFFFF", edge="#E3C9A8")
        self.card.pack(fill=tk.X)
        inner = self.card.inner
        top = tk.Frame(inner, bg="#FFFFFF")
        top.pack(fill=tk.X)
        self.cat = Asheng(top, img_path, size=M.px(52))
        self.cat.pack(side=tk.LEFT, anchor="n")
        right = tk.Frame(top, bg="#FFFFFF")
        right.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.bubble = Bubble(right)
        self.bubble.pack(fill=tk.X, padx=(0, M.px(2)), pady=(M.px(2), 0))
        bar = tk.Frame(right, bg="#FFFFFF")
        bar.pack(fill=tk.X, pady=(M.px(4), 0), padx=(M.px(10), 0))
        self.seg = tk.Frame(bar, bg="#FFFFFF")
        self.seg.pack(side=tk.LEFT)
        self.acct = tk.Frame(bar, bg="#FFFFFF")
        self.acct.pack(side=tk.RIGHT)
        self.body = tk.Frame(inner, bg="#FFFFFF")
        self.body.pack(fill=tk.X, pady=(M.px(6), 0))
        self.steps, self.stage, self.account, self.page = [], "委託", None, None
        self.asking = None                 # (程式名稱, 最後幾行, on_answer)：子程式在等回答
        self.note = None                   # (阿生要說的話, 按鈕文字, 按鈕動作)：剛做完一步的提示
        self.busy = None                   # 主程式有子程式在跑：那個功能的名稱(阿生改說「X 執行中」)
        self.note_box = tk.Frame(bar, bg="#FFFFFF")
        self.note_box.pack(side=tk.LEFT, padx=(M.px(8), 0))

    def _draw_seg(self):
        for w in self.seg.winfo_children():
            w.destroy()
        for s, tip in (("開發", "還在開發：先免費查詢，有希望再調謄本；物件資料委託後再做。"),
                       ("委託", "已經委託：全部做完，合併列印呈報公司。")):
            on = s == self.stage
            b = tk.Label(self.seg, text=s, font=(FAM, 13, "bold"), padx=M.px(12), pady=M.px(2), cursor="hand2",
                         fg="white" if on else INK2, bg=ORANGE if on else "#F2EFEA")
            b.pack(side=tk.LEFT, padx=(0, 1))
            b.bind("<Button-1>", lambda e, s=s: self.set_stage(s))
            Tip(b, tip)

    def _draw_acct(self):
        for w in self.acct.winfo_children():
            w.destroy()
        a = self.account
        if not a:
            b = Pill(self.acct, "🔑 Google 登入", lambda: self.on_action("login"))
            Tip(b, "登入阿生地圖（Google 帳號），就能在這裡查周邊行情；\n抖內支持者或會員還能用「行情推估」。")
        else:
            b = Pill(self.acct, f"👤 {a.get('name', '')}", lambda: self.on_action("account"))
            Tip(b, f"已登入阿生地圖（{TIER_NAME.get(a.get('tier'), '一般')}）。點一下可以登出。")
        b.pack(side=tk.LEFT)
        t = Pill(self.acct, "🧰", lambda: self.on_action("toolbox"))
        t.pack(side=tk.LEFT, padx=(M.px(4), 0))
        Tip(t, "工具箱：所有個別功能（想自己一項一項操作時用）")

    def set_stage(self, s):
        if s != self.stage:
            self.stage = s
            self.on_stage(s)
            self.on_action("refresh")

    def show(self, steps, stage, account=None):
        self.steps, self.stage, self.account = steps, stage, account
        self._draw_seg()
        self._draw_acct()
        for w in self.body.winfo_children():
            w.destroy()
        if self.asking:                   # 子程式在問：問答卡放最上面
            prog, lines, cb = self.asking
            AskCard(self.body, prog, lines, cb).pack(fill=tk.X, padx=M.px(3), pady=(0, M.px(8)))
        if self.page:
            self._phase_page(self.page)
        else:
            self._home()
        self.card._fit()
        mode, line = cat_line(steps, stage)
        for w in self.note_box.winfo_children():
            w.destroy()
        if self.asking:
            mode, line = "wait", f"{self.asking[0]}在問你問題，看下面就好～"
        elif self.busy:
            mode, line = "work", f"「{self.busy}」執行中，要回答的話在最下面輸入框打～"
        elif self.note:
            mode, line = "done", self.note[0]
            if self.note[1]:
                Pill(self.note_box, self.note[1], self.note[2]).pack(side=tk.LEFT)
        self.cat.set_mode(mode)
        self.bubble.set(line)

    def show_empty(self, text, btn_text, command, account=None):
        """還沒有案件：阿生說怎麼建案＋一顆大按鈕"""
        self.steps, self.account, self.page = [], account, None
        self._draw_seg()
        self._draw_acct()
        for w in self.body.winfo_children():
            w.destroy()
        for w in self.note_box.winfo_children():
            w.destroy()
        BigButton(self.body, btn_text, command).pack(fill=tk.X, padx=M.px(3), pady=(M.px(4), M.px(2)))
        self.card._fit()
        self.cat.set_mode("idle")
        self.bubble.set(text)

    def ask(self, program, lines, on_answer):
        """子程式在等回答 → 顯示問答卡；回答後自動收起"""
        def answer(v):
            self.asking = None
            self.show(self.steps, self.stage, self.account)
            on_answer(v)
        self.asking = (program, lines, answer)
        self.show(self.steps, self.stage, self.account)

    def notify(self, text, btn=None, action=None, auto=None):
        """剛做完一步：阿生說一句＋(可選)一顆按鈕；auto＝秒數，倒數完自動做 action(按按鈕＝先不要)"""
        self.note = (text, btn, action)
        self.show(self.steps, self.stage, self.account)

    def open_page(self, ph):
        self.page = ph
        self.show(self.steps, self.stage, self.account)

    # 主頁：四個大方塊＋下一步
    def _home(self):
        grid = tk.Frame(self.body, bg="#FFFFFF")
        grid.pack(fill=tk.X)
        grid.columnconfigure(0, weight=1, uniform="t")
        grid.columnconfigure(1, weight=1, uniform="t")
        for n, ph in enumerate(PHASES):
            sts = [s for s in self.steps if s["phase"] == ph[0] and s["state"] != "extra"]
            done = sum(s["state"] == "done" for s in sts)
            dev_later = self.stage == "開發" and ph[0] == "做"
            if any(s["state"] == "wait" for s in sts):
                state, sub = "wait", "等你處理"
            elif any(s["state"] == "run" for s in sts):
                state, sub = "run", "阿生處理中…"
            elif done == len(sts):
                state, sub = "done", "完成"
            elif dev_later:
                state, sub = "skip", "委託後再做"
            else:
                state, sub = "todo", f"{done} / {len(sts)}"
            hot = any(s["state"] == "next" for s in sts)
            title = {"查": "查詢", "調": "謄本", "做": "物件資料", "交": "給屋主" if self.stage == "開發" else "交件"}[ph[0]]
            t = Tile(grid, f"{PHASE_ICON[ph[0]]} {title}", sub, state, (lambda k=ph[0]: self.open_page(k)),
                     dim=dev_later, hot=hot)
            t.grid(row=n // 2, column=n % 2, sticky="ew", padx=M.px(3), pady=M.px(3))
            Tip(t, {"查": "建立案件、阿生地圖、免費附件查詢（使照、分區、實價…）",
                    "調": "調閱謄本、結構化、土地增值稅、地籍圖標註",
                    "做": "現場照片、物件編輯器、物調表、行銷簡報、業績分配表",
                    "交": "合併列印呈報公司" if self.stage == "委託" else "實價行情為主的資料，給屋主看"}[ph[0]])
        nxt = next((s for s in self.steps if s["state"] == "next"), None)
        if nxt:
            act = next((a for a in nxt["actions"] if a[2]), nxt["actions"][0] if nxt["actions"] else None)
            big = BigButton(self.body, f"▶ 下一步：{nxt['title']}",
                            (lambda k=act[1]: self.on_action(k)) if act else (lambda p=nxt["phase"]: self.open_page(p)))
            big.pack(fill=tk.X, padx=M.px(3), pady=(M.px(8), M.px(2)))
            Tip(big, nxt["tip"])

    # 分頁：該階段每一步(重跑、單獨開都在這)
    def _phase_page(self, key):
        ph = next(p for p in PHASES if p[0] == key)
        hdr = tk.Frame(self.body, bg="#FFFFFF")
        hdr.pack(fill=tk.X, pady=(0, M.px(4)))
        Pill(hdr, "← 回主頁", lambda: self.open_page(None)).pack(side=tk.LEFT)
        tk.Label(hdr, text=f"  {PHASE_ICON[key]} {ph[1]}", font=(FAM, 15, "bold"), bg="#FFFFFF", fg=INK).pack(side=tk.LEFT)
        for i, s in enumerate(self.steps):
            if s["phase"] == key:
                self._row(i, s)

    def _row(self, i, s):
        st = s["state"]
        hi = st in ("next", "wait")
        bg = ORANGE_BG if hi else "#FFFFFF"
        row = tk.Frame(self.body, bg=bg)
        row.pack(fill=tk.X, pady=M.px(2))
        num = "🗺" if st == "extra" else str(sum(1 for x in self.steps[:i + 1] if x["state"] != "extra"))
        StatusDot(row, st, num).pack(side=tk.LEFT, padx=(M.px(6), M.px(8)), pady=M.px(6))
        dim = st in ("skip", "lock")
        txt = tk.Frame(row, bg=bg)
        txt.pack(side=tk.LEFT, fill=tk.X, expand=True)
        t = tk.Label(txt, text=s["title"], font=(FAM, 14, "bold"), bg=bg, fg=DIM if dim else INK, anchor="w")
        t.pack(anchor="w")
        if s["detail"]:
            dfg = DIM if dim else (RED if st == "wait" else (GREEN if st == "done" else INK2))
            tk.Label(txt, text=s["detail"], font=(FAM, 11), bg=bg, fg=dfg, anchor="w").pack(anchor="w")
        acts = tk.Frame(row, bg=bg)
        acts.pack(side=tk.RIGHT, padx=M.px(6))
        for a in s["actions"]:
            label, key, primary = a[:3]
            p = Pill(acts, label, (lambda k=key: self.on_action(k)), primary=primary and st == "next", dim=dim)
            p.pack(side=tk.LEFT, padx=(M.px(4), 0))
            if len(a) > 3 and a[3]:
                Tip(p, a[3])
        for w in (row, txt, t):
            Tip(w, s["tip"])
        return row


class BigButton(tk.Canvas):
    """主頁最大的那顆橘色按鈕(下一步)；問答卡的按鈕也用它"""

    def __init__(self, parent, text, command, height=None, width=None, color=None, size=17, fg="white"):
        super().__init__(parent, height=height or M.px(56), bg=parent.cget("bg"), highlightthickness=0, cursor="hand2")
        if width:
            self.configure(width=width)
        self.size, self.fg = size, fg
        self.color = color or ORANGE
        self.text, self.command, self.hover = text, command, False
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", lambda e: self._h(True))
        self.bind("<Leave>", lambda e: self._h(False))
        self.bind("<ButtonRelease-1>", lambda e: self.command and self.command())

    def _h(self, v):
        self.hover = v
        self._draw()

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), int(self.cget("height"))
        if w < 20:
            return
        M.round_rect(self, 2, 2, w - 2, h - 2, M.px(14), fill=M.mix(self.color, "#000000", 0.18) if self.hover else self.color,
                     outline="")
        self.create_text(w / 2, h / 2, text=self.text, font=(FAM, self.size, "bold"), fill=self.fg, width=w - M.px(16),
                         justify="center")


TOOLBOX = [   # 只放「步驟」裡沒有的：各項查詢在第 2 步「清單」(連續模式)可以單跑，謄本/土增稅/物件編輯器在各步驟
    ("建立案件", [("萬用查詢（打地址/地號）", "quick"), ("地籍便民系統-手動", "land_manual"), ("地籍便民-舊版截圖", "land_old")]),
    ("案件檔案", [("載入案件 data.json", "load_case"), ("讀取結構化JSON", "read_json"), ("匯入封存", "import_arc"),
              ("匯出封存", "export_arc"), ("開啟工作資料夾", "folder")]),
    ("設定", [("修改帳密", "cred"), ("Chrome 縮放設定", "zoom")]),
    ("阿生地圖（Google 登入）", [("在地圖看", "lm_view"), ("周邊行情・在售", "lm_near"), ("行情推估🔒", "lm_avm")]),
]


def open_toolbox(parent, on_pick, x=None, y=None, account=None):
    """工具箱：所有個別功能一層攤開(不用再進分頁、按回主頁)；account 能用推估就拿掉🔒"""
    tw = tk.Toplevel(parent)
    tw.title("工具箱")
    tw.configure(bg="#FFFFFF")
    tk.Label(tw, text="🧰 其他功能", font=(FAM, 15, "bold"), bg="#FFFFFF", fg=INK,
             anchor="w").pack(fill=tk.X, padx=M.px(14), pady=(M.px(12), 0))
    tk.Label(tw, text="各項查詢（使照、分區、實價…）在第 2 步按「清單」可以一項一項單獨跑。", font=(FAM, 12),
             bg="#FFFFFF", fg=INK2, anchor="w", justify="left").pack(fill=tk.X, padx=M.px(14), pady=(0, M.px(6)))
    acc = tk.Frame(tw, bg="#FFFFFF")
    acc.pack(fill=tk.X, padx=M.px(14), pady=(0, M.px(6)))
    Pill(acc, (f"👤 {account.get('name', '')}（點一下登出）" if account else "🔑 Google 登入阿生地圖"),
         (lambda: (tw.destroy(), on_pick("account" if account else "login")))).pack(anchor="w")
    grid = tk.Frame(tw, bg="#FFFFFF")
    grid.pack(fill=tk.BOTH, padx=M.px(14), pady=(0, M.px(12)))
    for gi, (gname, tools) in enumerate(TOOLBOX):
        col = tk.Frame(grid, bg="#FFFFFF")
        col.grid(row=gi // 2, column=gi % 2, sticky="nw", padx=(0, M.px(16)), pady=M.px(4))
        tk.Label(col, text=gname, font=(FAM, 13, "bold"), fg=ORANGE_DK, bg="#FFFFFF").pack(anchor="w")
        wrap = tk.Frame(col, bg="#FFFFFF")
        wrap.pack(anchor="w")
        for name, key in tools:
            if key == "lm_avm" and avm_ok(account):
                name = name.replace("🔒", "")
            Pill(wrap, name, (lambda k=key: (tw.destroy(), on_pick(k))), size=13).pack(anchor="w", pady=2)
    back = tk.Label(tw, text="↩ 切回舊版介面（原本的按鈕畫面）", font=(FAM, 11, "underline"), fg="#1F5FAF",
                    bg="#FFFFFF", cursor="hand2")
    back.pack(anchor="w", padx=M.px(14), pady=(0, M.px(12)))
    back.bind("<Button-1>", lambda e: (tw.destroy(), on_pick("classic_ui")))
    Tip(back, "換回原本的按鈕畫面。\n舊版畫面右上角的 ✨ 可以再切回新版。")
    if x is not None:
        tw.geometry(f"+{x}+{y}")
    return tw


# ───────────────────────── 問答卡：子程式問問題時，用大字＋按鈕回答(取代在訊息框看、在輸入框打) ─────────────────────────
import re as _re

_OPT = _re.compile(r"^\s*[\[(（]?\s*(\d{1,2})\s*[\])）.、:：]\s*(.{1,30}?)\s*$")


def parse_prompt(lines):
    """子程式最後印的幾行 → (問題, 按鈕[(文字, 要送的值)], 要不要輸入框, 要不要遮住)"""
    lines = [_re.sub(r"\x1b\[[0-9;]*m", "", ln).strip() for ln in lines if ln and ln.strip()]
    if not lines:
        return "", [("繼續", "")], False, False
    q = lines[-1]
    secret = bool(_re.search(r"密碼|PIN|身分證", q, _re.I))
    if _re.search(r"[(（]\s*Y\s*/\s*N", q, _re.I):
        dflt = "Y" if _re.search(r"預設\s*Y", q, _re.I) else ("N" if _re.search(r"預設\s*N", q, _re.I) else "")
        q2 = _re.sub(r"[(（][^)）]*[)）]\s*[:：]?\s*$", "", q).strip()
        return q2, [("是" + ("（預設）" if dflt == "Y" else ""), "Y"), ("否" + ("（預設）" if dflt == "N" else ""), "N")], False, False
    opts = []
    for ln in reversed(lines[:-1][-8:]):
        m = _OPT.match(ln)
        if m:
            opts.insert(0, (m.group(2), m.group(1)))
        elif opts:
            break
    if opts and _re.search(r"選項|選擇|編號|\d\s*/\s*\d", q):
        return q.rstrip("：: "), opts, False, False
    if _re.search(r"Enter|按.{0,4}繼續|完成後", q, _re.I) or q.endswith(("...", "…")):
        return q, [("繼續", "")], False, False
    return q.rstrip("：: "), [], True, secret


class AskCard(tk.Frame):
    """主畫面上的問答卡：哪個程式在問、問題(大字)、按鈕或輸入框。on_answer(值) 送回子程式"""

    def __init__(self, parent, program, lines, on_answer):
        super().__init__(parent, bg=ORANGE_BG, highlightbackground=ORANGE, highlightthickness=2)
        q, btns, need_entry, secret = parse_prompt(lines)
        tk.Label(self, text=f"🐱 {program} 在問你：", font=(FAM, 12, "bold"), bg=ORANGE_BG, fg=ORANGE_DK,
                 anchor="w").pack(fill=tk.X, padx=M.px(12), pady=(M.px(10), 0))
        tk.Label(self, text=q, font=(FAM, 16, "bold"), bg=ORANGE_BG, fg=INK, anchor="w", justify="left",
                 wraplength=M.px(520)).pack(fill=tk.X, padx=M.px(12), pady=(M.px(4), M.px(8)))
        row = tk.Frame(self, bg=ORANGE_BG)
        row.pack(fill=tk.X, padx=M.px(12), pady=(0, M.px(12)))
        if need_entry:
            var = tk.StringVar()
            e = tk.Entry(row, textvariable=var, font=(FAM, 16), show="●" if secret else "", relief="solid", bd=1)
            e.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=M.px(4))
            e.bind("<Return>", lambda ev: on_answer(var.get()))
            e.focus_set()
            BigButton(row, "送出", lambda: on_answer(var.get()), height=M.px(44), width=M.px(110)).pack(side=tk.LEFT, padx=(M.px(8), 0))
            if secret:
                tk.Label(self, text="（輸入的內容不會顯示在畫面和紀錄裡）", font=(FAM, 11), bg=ORANGE_BG, fg=INK2).pack(
                    anchor="w", padx=M.px(12), pady=(0, M.px(8)))
        elif len(btns) <= 2:                 # 是／否、繼續：一排
            for i, (label, val) in enumerate(btns):
                b = BigButton(row, label, (lambda v=val: on_answer(v)), height=M.px(48),
                              color=ORANGE if i == 0 else "#7A7A80")
                b.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0 if i == 0 else M.px(8), 0))
        else:                                # 好幾個選項：一個一排，大小一樣
            for label, val in btns:
                BigButton(row, label, (lambda v=val: on_answer(v)), height=M.px(44), color="#D9822B").pack(
                    fill=tk.X, pady=M.px(3))


class LogStrip(tk.Frame):
    """收起來的過程紀錄：平常一行(最新一句)，點開才看全部；有錯誤時變紅"""

    def __init__(self, parent, on_open=None):
        super().__init__(parent, bg=M.BG)
        self.lbl = tk.Label(self, text="📜 過程紀錄", font=(FAM, 12), bg="#F2EFEA", fg=INK2, anchor="w", padx=M.px(10),
                            pady=M.px(6), cursor="hand2")
        self.lbl.pack(fill=tk.X)
        self.lbl.bind("<Button-1>", lambda e: on_open and on_open())
        Tip(self.lbl, "每個程式做了什麼的完整紀錄（出問題要回報時用）。點一下打開。")

    def set(self, last, error=False):
        last = last if len(last) <= 26 else last[:25] + "…"
        self.lbl.configure(text=f"📜 {last}　　▸ 看全部", fg="white" if error else INK2,
                           bg=RED if error else "#F2EFEA")


# ───────────────────────── 帶路版主頁：一次只做一件事(使用者 10-30：四方塊版看不懂、字太小) ─────────────────────────
GUIDE_HELP = {
    "case": "在最下面的輸入框打地址、門牌或「地段 地號」，按 Enter 就會建好案件資料夾。",
    "query": "免費：使照、分區、地籍便民、實價…自動一項一項查。",
    "trans": "要付費：調土地、建物謄本和地籍圖。",
    "struct": "把謄本 PDF 轉成資料，土增稅和物件資料會用到。",
    "tax": "用結構化資料自動算土地增值稅。",
    "map": "在地籍圖標出邊長、臨路路名和路寬。",
    "photo": "照片放進「2.照片」：直接拖進主程式視窗也可以。",
    "editor": "物件編輯器：補上售價、格局、車位（謄本資料已自動帶入）。",
    "docs": "一鍵產生物調表、行銷簡報、業績分配表。",
    "deliver": "把做好的檔案合成一份 PDF。",
}


def _autowrap(lbl):
    lbl.bind("<Configure>", lambda e: lbl.configure(wraplength=max(100, e.width - 6)), add="+")
    return lbl


class Guide(tk.Frame):
    """主頁(帶路版)：上面一張大卡片＝現在這一步(說明＋大按鈕)；下面大字清單＝全部步驟，點一行可以跳過去。
    on_action(key)；on_stage(stage)；on_skip(key, True/False)＝這步不用做/取消不用做"""

    def __init__(self, parent, img_path, on_action=None, on_stage=None, on_skip=None):
        super().__init__(parent, bg=M.BG)
        self.on_action = on_action or (lambda k: None)
        self.on_stage, self.on_skip = on_stage or (lambda s: None), on_skip or (lambda k, v: None)
        self.card = M.Card(self, fill="#FFFFFF", edge="#E3C9A8")
        self.card.pack(fill=tk.X)
        inner = self.card.inner
        self.img_path = img_path
        self.cat = None                    # 阿生放在大卡片標題旁(每次重畫重建)
        self.cur = tk.Frame(inner, bg="#FFFFFF")
        self.cur.pack(fill=tk.X)
        self.lst = tk.Frame(inner, bg="#FFFFFF")
        self.lst.pack(fill=tk.X, pady=(M.px(8), 0))
        self.foot = tk.Frame(inner, bg="#FFFFFF")
        self.foot.pack(fill=tk.X, pady=(M.px(8), 0))
        self.steps, self.stage, self.account, self.skipped = [], "委託", None, set()
        self.pick, self.busy, self.note, self._prev_done, self._last = None, None, None, None, None
        self.page = None                   # 相容 CaseFlow(主程式拿來判斷要不要重畫)
        self.list_open = True              # 全部步驟清單展開/收起(主程式記住)

    def show(self, steps, stage, account=None, skipped=()):
        self._last = (steps, stage, account, set(skipped))
        self.steps, self.stage, self.account, self.skipped = [dict(s) for s in steps], stage, account, set(skipped)
        main = [s for s in self.steps if s["state"] != "extra"]
        for s in main:
            if s["state"] == "next":
                s["state"] = "todo"
            if s["state"] != "done" and s["key"] in self.skipped:
                s["state"] = "skipped"
        todo = [s for s in main if s["state"] in ("todo", "run", "wait")]
        auto = todo[0] if todo else None
        cur = next((s for s in main if s["key"] == self.pick), None) or auto
        done = {s["key"] for s in main if s["state"] == "done"}
        if self._prev_done is not None and auto:
            new = [s for s in main if s["key"] in done - self._prev_done]
            if new:                         # 剛做好一步 → 打勾、自動換到下一步
                self.note = f"✓ {new[-1]['title']} 做好了！接下來：第 {main.index(auto) + 1} 步 {auto['title']}"
                self.pick = None
                cur = auto
        self._prev_done = done
        for fr in (self.cur, self.lst, self.foot):
            for w in fr.winfo_children():
                w.destroy()
        self.lst.configure(height=1)
        self._draw_cur(main, cur, auto)
        if self.busy:                       # 程式執行中：收起清單，讓訊息框大一點(子程式的提問在那裡)
            tk.Label(self.lst, text="（程式執行中，步驟清單先收起來）",
                     font=(FAM, 12), bg="#FFFFFF", fg=INK2, anchor="w").pack(fill=tk.X, padx=M.px(4))
        else:
            self._draw_list(main, cur)
        self._draw_foot()
        self.card._fit()
        if self.cat:
            self.cat.set_mode("work" if self.busy else ("done" if self.note or not auto else "idle"))

    def _redraw(self):
        if self._last:
            self.note = None
            self.show(*self._last)

    def _goto(self, key):
        self.pick = key
        self._redraw()

    def _head(self, box, title, color=ORANGE_DK):
        """大卡片標題列：阿生＋「第 N 步 XXX」；下面一行阿生說的話(剛做好／執行中)"""
        hd = tk.Frame(box, bg=ORANGE_BG)
        hd.pack(fill=tk.X, padx=M.px(10), pady=(M.px(6), 0))
        self.cat = Asheng(hd, self.img_path, size=M.px(40))
        self.cat.pack(side=tk.LEFT)
        tk.Label(hd, text=title, font=(FAM, 21, "bold"), bg=ORANGE_BG, fg=color, anchor="w").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(M.px(6), 0))
        say = (f"「{self.busy}」正在做，做完我會自動打勾、接下一步" if self.busy else self.note)
        if say:
            _autowrap(tk.Label(box, text=say, font=(FAM, 14, "bold"), bg=ORANGE_BG, anchor="w", justify="left",
                               fg=ORANGE_DK if self.busy else GREEN)).pack(fill=tk.X, padx=M.px(14), pady=(M.px(2), 0))

    def _draw_cur(self, main, cur, auto):
        box = tk.Frame(self.cur, bg=ORANGE_BG, highlightbackground=ORANGE, highlightthickness=2)
        box.pack(fill=tk.X, padx=M.px(2))
        pad = dict(padx=M.px(14))
        if not cur:
            self._head(box, "🎉 全部做好了", GREEN)
            tk.Label(box, text="想重做哪一步，點下面清單那一行就好。", font=(FAM, 14), bg=ORANGE_BG, fg=INK2).pack(
                anchor="w", pady=(M.px(4), M.px(10)), **pad)
            BigButton(box, "📂 打開案件資料夾", lambda: self.on_action("folder"), height=M.px(54), size=17).pack(
                fill=tk.X, pady=(0, M.px(12)), **pad)
            return
        self._head(box, f"第 {main.index(cur) + 1} 步　{cur['title']}")
        st = cur["state"]
        if st == "done":
            msg, col = "✓ 這步已經做好了。要重做可以再按下面的按鈕。", GREEN
        elif st in ("skipped", "skip"):
            msg, col = "這步目前不用做。要做的話按下面的按鈕。", INK2
        elif st == "lock":
            msg, col = f"還不能做：{cur['detail']}", RED
        else:
            msg, col = GUIDE_HELP.get(cur["key"], cur.get("tip", "")), INK
        _autowrap(tk.Label(box, text=msg, font=(FAM, 14), bg=ORANGE_BG, fg=col, anchor="w", justify="left")).pack(
            fill=tk.X, pady=(M.px(2), 0), **pad)
        if cur["detail"] and st != "lock":
            _autowrap(tk.Label(box, text=f"目前：{cur['detail']}", font=(FAM, 13), bg=ORANGE_BG, fg=INK2, anchor="w",
                               justify="left")).pack(fill=tk.X, **pad)
        acts = [a for a in cur["actions"] if a[1] != "query_stop"]
        big, small = acts[:2], acts[2:]
        if big:
            row = tk.Frame(box, bg=ORANGE_BG)
            row.pack(fill=tk.X, pady=(M.px(8), 0), **pad)
            for i, a in enumerate(big):
                label = a[0] if i or a[0].startswith(("▶", "📂", "📑", "📥")) else f"▶ {a[0]}"
                b = BigButton(row, label, (lambda k=a[1]: self.on_action(k)), height=M.px(50), size=18 if i == 0 else 16,
                              color=ORANGE if i == 0 else "#2F7D63")
                row.columnconfigure(i, weight=2 if i == 0 else 1, uniform="b")
                b.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else M.px(8), 0))
                if len(a) > 3 and a[3]:
                    Tip(b, a[3])
        if small:
            row2 = tk.Frame(box, bg=ORANGE_BG)
            row2.pack(fill=tk.X, pady=(M.px(6), 0), **pad)
            for j, a in enumerate(small):
                row2.columnconfigure(j, weight=1, uniform="s")
                BigButton(row2, a[0], (lambda k=a[1]: self.on_action(k)), height=M.px(42), size=14,
                          color="#E9E2D8", fg=INK).grid(row=0, column=j, sticky="ew", padx=(0, M.px(6)))
        low = tk.Frame(box, bg=ORANGE_BG)
        low.pack(fill=tk.X, pady=(M.px(6), M.px(8)), **pad)

        def link(text, cmd):
            lb = tk.Label(low, text=text, font=(FAM, 13, "underline"), bg=ORANGE_BG, fg="#1F5FAF", cursor="hand2")
            lb.pack(side=tk.LEFT, padx=(0, M.px(18)))
            lb.bind("<Button-1>", lambda e: cmd())
        if st in ("todo", "lock"):
            link("這步不用做，跳下一步", lambda k=cur["key"]: (self.on_skip(k, True), self._goto(None)))
        elif st == "skipped":
            link("改回要做", lambda k=cur["key"]: (self.on_skip(k, False), self._redraw()))
        if auto and cur is not auto:
            link(f"回到現在該做的（第 {main.index(auto) + 1} 步）", lambda: self._goto(None))

    def _draw_list(self, main, cur):
        n_done = sum(s["state"] == "done" for s in main)
        hd = tk.Label(self.lst, text=(f"▾ 全部步驟　已做好 {n_done}/{len(main)}（點這裡收起來）"
                                      if self.list_open else f"▸ 全部步驟　已做好 {n_done}/{len(main)}（點這裡展開）"),
                      font=(FAM, 13, "bold"), bg="#FFFFFF", fg="#1F5FAF", anchor="w", cursor="hand2")
        hd.pack(fill=tk.X, padx=M.px(4))
        hd.bind("<Button-1>", lambda e: self.on_action("list_toggle"))
        if not self.list_open:
            return
        grid = tk.Frame(self.lst, bg="#FFFFFF")
        grid.pack(fill=tk.X)
        half = (len(main) + 1) // 2
        for c in (0, 1):
            grid.columnconfigure(c, weight=1, uniform="l")
        for i, s in enumerate(main):
            st = s["state"]
            on = s is cur
            bg = ORANGE_BG if on else "#FFFFFF"
            row = tk.Frame(grid, bg=bg, cursor="hand2")
            row.grid(row=i % half, column=i // half, sticky="ew", padx=M.px(2), pady=1)
            icon, icol = {"done": ("✓", GREEN), "skipped": ("—", DIM), "skip": ("—", DIM), "lock": ("○", DIM),
                          "run": ("…", ORANGE_DK), "wait": ("!", RED)}.get(st, ("○", INK2))
            if on:
                icon, icol = "▶", ORANGE
            ws = [tk.Label(row, text=icon, font=(FAM, 14, "bold"), bg=bg, fg=icol, width=2),
                  tk.Label(row, text=f"{i + 1}. {s['title']}", font=(FAM, 14, "bold" if on else "normal"), bg=bg,
                           fg=DIM if st in ("skipped", "skip", "lock") else INK, anchor="w")]
            ws[0].pack(side=tk.LEFT)
            ws[1].pack(side=tk.LEFT, fill=tk.X, expand=True)
            for w in [row] + ws:
                w.bind("<Button-1>", lambda e, k=s["key"]: self._goto(k))
            Tip(row, {"done": "做好了", "skipped": "設成不用做", "skip": "委託後再做", "lock": "還不能做：" + s["detail"]}.get(st, "還沒做"))

    def _draw_foot(self, stage=True):
        if stage:
            self._draw_stage()
        self._draw_tools()

    def _draw_stage(self):
        tk.Label(self.foot, text="這件是：", font=(FAM, 14), bg="#FFFFFF", fg=INK2).pack(side=tk.LEFT, padx=(M.px(4), 0))
        for s, tip in (("委託", "已經委託：全部做完，合併列印交給公司。"),
                       ("開發", "還在開發：先免費查詢，有希望再調謄本；物件資料委託後再做。")):
            on = s == self.stage
            b = tk.Label(self.foot, text=s, font=(FAM, 14, "bold"), padx=M.px(14), pady=M.px(3), cursor="hand2",
                         fg="white" if on else INK2, bg=ORANGE if on else "#F2EFEA")
            b.pack(side=tk.LEFT, padx=(0, 2))
            b.bind("<Button-1>", lambda e, s=s: self._set_stage(s))
            Tip(b, tip)

    def _draw_tools(self):
        q = tk.Label(self.foot, text="🔍 查詢新案件", font=(FAM, 14, "bold"), padx=M.px(12), pady=M.px(3), cursor="hand2",
                     fg="white", bg="#2F7D63")
        q.pack(side=tk.RIGHT, padx=(M.px(6), 0))
        q.bind("<Button-1>", lambda e: self.on_action("quick"))
        Tip(q, "打開左邊的萬用查詢面板：打地址、門牌或地段地號，查好就建立新案件。")
        tb = tk.Label(self.foot, text="🧰 其他功能", font=(FAM, 14, "bold"), padx=M.px(12), pady=M.px(3), cursor="hand2",
                      fg=INK, bg="#F2EFEA")
        tb.pack(side=tk.RIGHT)
        tb.bind("<Button-1>", lambda e: self.on_action("toolbox"))
        Tip(tb, "所有個別功能（想自己一項一項操作時用），最下面可以切回舊版介面。")
        if not self.account:                # 登入後就不佔位置(登出在「其他功能」裡)
            lg = tk.Label(self.foot, text="🔑 登入", font=(FAM, 13), padx=M.px(10), pady=M.px(3), cursor="hand2",
                          fg=INK, bg="#F2EFEA")
            lg.pack(side=tk.RIGHT, padx=(0, M.px(6)))
            lg.bind("<Button-1>", lambda e: self.on_action("login"))
            Tip(lg, "用 Google 帳號登入阿生地圖（周邊行情、行情推估要登入）。")

    def _set_stage(self, s):
        if s != self.stage:
            self.stage = s
            self.on_stage(s)
            self.on_action("refresh")

    def show_empty(self, text, btn_text, command, account=None):
        """還沒有案件：第 1 步 建立案件"""
        self.account, self._prev_done, self.pick, self._last = account, None, None, None
        for fr in (self.cur, self.lst, self.foot):
            for w in fr.winfo_children():
                w.destroy()
        self.lst.configure(height=1)        # 清空的 Frame 不會自己縮回去
        box = tk.Frame(self.cur, bg=ORANGE_BG, highlightbackground=ORANGE, highlightthickness=2)
        box.pack(fill=tk.X, padx=M.px(2))
        self.note = None
        self._head(box, "第 1 步　建立案件")
        _autowrap(tk.Label(box, text=GUIDE_HELP["case"], font=(FAM, 14), bg=ORANGE_BG, fg=INK, anchor="w",
                           justify="left")).pack(fill=tk.X, padx=M.px(14), pady=(M.px(4), 0))
        BigButton(box, btn_text, command, height=M.px(58), size=18).pack(fill=tk.X, padx=M.px(14),
                                                                          pady=(M.px(10), M.px(12)))
        self._draw_foot(stage=False)
        self.card._fit()
