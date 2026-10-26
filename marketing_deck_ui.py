"""行銷簡報視窗：① 版型與版面 → ② 頁面內容(要哪些頁、用哪份資料，可預覽) → ③ 挑照片 → ④ 整理與確認。
pick() 回傳 (選用的照片清單, 輸出路徑, 版面, 版型, 專案人員, 封底, 色系, 頁面內容選擇) 或 None(取消)"""
import os
import sys
import queue
import threading
import tkinter as tk
try:
    import os as _os, sys as _sys
    _sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
    import tk_textmenu as _tkm
    _tkm.auto()          # 輸入框/文字區右鍵：剪下、複製、貼上、全選
except Exception:
    pass
from tkinter import ttk, filedialog, messagebox

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)
import deal_ui as ui
from PIL import Image, ImageOps, ImageTk

INK = "#0F172A"
INK2 = "#1E293B"
MUTED = "#475569"      # 製作中的字(不用淡灰，仍看得清楚)
BLUE = "#0A6CDB"
BLUE_BG = "#EAF3FF"
USED = "#3E83B6"       # 有選用的框色(天空藍)
COVER = "#B88E3E"      # 封面(金)
UNUSED = "#CBD5E1"
LAYOUT = "#7C3AED"      # 格局平面圖(紫)，跟照片分類區隔


def _guard_sep(cb, v, p):
    """下拉選到分隔線 → 退回原本的分類"""
    from marketing_deck import CAT_SEP
    last = {"v": v.get()}

    def on_sel(e=None):
        if v.get() == CAT_SEP:
            v.set(last["v"])
        else:
            last["v"] = v.get()
    cb.bind("<<ComboboxSelected>>", on_sel, add="+")
GREEN = "#1B7F3B"
GOLD_FG = "#8A6416"
ORANGE = "#B35400"
FAM = "Microsoft JhengHei"
EMOJI = "Segoe UI Emoji"


def pick(photos, facts, default_out, page="a4l", on_save=None, template="building", suggest="building", show_agents=True,
         back_cover=False, theme="fzt", get_sections=None):
    from marketing_deck import TEMPLATES
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    ui.apply(root)
    root.title("行銷簡報")
    sc = max(1.0, root.winfo_fpixels("1i") / 96)
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    W, H = min(int(1180 * sc), sw - 60), min(int(900 * sc), sh - 90)
    root.geometry(f"{W}x{H}+{(sw - W) // 2}+{max(0, (sh - H) // 2 - 20)}")
    if not TEMPLATES.get(template, (0, 0, 0, 0, False))[4]:
        template = "building"
    state = {"page": page, "template": template, "out": default_out, "show_agents": show_agents, "back_cover": back_cover,
             "theme": theme, "sections": None, "sections_tp": None}

    def sel():
        from marketing_deck import sections_to_sel
        return sections_to_sel(state["sections"]) if state["sections"] else None

    def save():
        if on_save:
            on_save(state["page"], state["template"], state["show_agents"], state["back_cover"], state["theme"], sel() or {})

    step = 1
    while True:
        if step == 1:
            r = _step_template(root, facts, state, suggest, sc, W)
            step = 2 if r == "next" else 0
        elif step == 2:
            if get_sections and state["sections_tp"] != state["template"]:     # 換了版型 → 單元不同，重新收集
                try:
                    state["sections"] = get_sections(state["template"])
                except Exception as e:
                    print(f"（收集頁面內容失敗：{e}）", flush=True)
                    state["sections"] = []
                state["sections_tp"] = state["template"]
            if not state["sections"]:
                step = 3
                continue
            r = _step_sections(root, state, facts, sc, W)
            step = {"back": 1, "next": 3}.get(r, 0)
        elif step == 3:
            r = _step_photos(root, photos, facts, state, sc, W)
            step = {"back": 2 if state["sections"] else 1, "next": 4}.get(r, 0)
        elif step == 4:
            r = _step_arrange(root, photos, facts, state, sc, W)
            if r == "back":
                step = 3
                continue
            save()
            _close(root)
            if r != "go":
                return None
            chosen = [p for p in photos if p["cat"] != "不使用"]
            return (chosen, state["out"], state["page"], state["template"], state["show_agents"], state["back_cover"],
                    state["theme"], sel())
        if step == 0:
            save()
            _close(root)
            return None


def _close(root):
    try:
        root.destroy()
    except tk.TclError:
        pass


def _steps_bar(parent, active):
    """步驟條：① 版型與版面 → ② 頁面內容 → ③ 挑照片 → ④ 整理與確認"""
    bar = tk.Frame(parent, bg=ui.BG, padx=18, pady=10)
    bar.pack(fill=tk.X)
    names = ("版型與版面", "頁面內容", "挑照片", "整理與確認")
    for i, name in enumerate(names):
        on = i + 1 == active
        done = i + 1 < active
        chip = tk.Frame(bar, bg=BLUE if on else ("#D7E7FB" if done else "#E2E8F0"), padx=14, pady=6)
        chip.pack(side=tk.LEFT)
        tk.Label(chip, text=f"{'✓' if done else i + 1}  {name}", font=(FAM, 15, "bold"),
                 fg="white" if on else (BLUE if done else MUTED), bg=chip.cget("bg")).pack()
        if i < len(names) - 1:
            tk.Label(bar, text="  ›  ", font=(FAM, 18, "bold"), fg=MUTED, bg=ui.BG).pack(side=tk.LEFT)


# ---------------------------------------------------------------- ① 版型與版面
def _step_template(root, facts, state, suggest, sc, W):
    from marketing_deck import TEMPLATES, PAGE_CHOICES, THEMES
    stage = tk.Frame(root, bg=ui.BG)
    stage.pack(fill=tk.BOTH, expand=True)
    done = tk.StringVar(master=root)
    root.protocol("WM_DELETE_WINDOW", lambda: done.set("close"))
    ui.header(stage, "🎞 行銷簡報", f"{facts.get('short') or facts.get('title', '')}　先選版型與版面，下一步選要放哪些頁")
    foot = ui.footer(stage)
    _steps_bar(stage, 1)
    body, _cv = ui.scroll_body(stage)

    # ---- 版型 ----
    hd = tk.Frame(body, bg=ui.BG)
    hd.pack(fill=tk.X, pady=(4, 8))
    tk.Label(hd, text="選擇版型", font=(FAM, 20, "bold"), fg=INK, bg=ui.BG).pack(side=tk.LEFT)
    sname = TEMPLATES.get(suggest, TEMPLATES["building"])[1]
    tip = f"依物件型式「{facts.get('type') or '—'}」建議：{sname}"
    if not TEMPLATES.get(suggest, (0, 0, 0, 0, True))[4]:
        tip += "（製作中，先用大樓版）"
    tk.Label(hd, text=tip, font=(FAM, 14), fg=INK2, bg=ui.BG).pack(side=tk.LEFT, padx=16)

    grid = tk.Frame(body, bg=ui.BG)
    grid.pack(fill=tk.X)
    cols = 3
    for c in range(cols):
        grid.grid_columnconfigure(c, weight=1, uniform="tpl")
    cards = {}

    def paint_cards():
        for key, (fr, parts, ok) in cards.items():
            sel = key == state["template"]
            bg = BLUE_BG if sel else ("white" if ok else "#F1F5F9")
            fr.configure(bg=bg, highlightbackground=BLUE if sel else ("#CBD5E1" if ok else "#E2E8F0"),
                         highlightthickness=3 if sel else 1)
            for w in parts:
                try:
                    w.configure(bg=bg)
                except tk.TclError:
                    pass
            fr.check.configure(text=" ✓ 已選 " if sel else "", bg=BLUE if sel else bg)

    def choose_tpl(key, ok):
        if not ok:
            ui.message(root, "製作中", f"「{TEMPLATES[key][1]}」版型還在製作中，做好會在這裡開放。\n目前先用「大樓・華廈」版。")
            return
        state["template"] = key
        paint_cards()

    for i, (key, (icon, name, fits, pages, ok)) in enumerate(TEMPLATES.items()):
        fr = tk.Frame(grid, bg="white", highlightthickness=1, highlightbackground="#CBD5E1", padx=16, pady=14,
                      cursor="hand2" if ok else "arrow")
        fr.grid(row=i // cols, column=i % cols, padx=8, pady=8, sticky="nsew")
        top = tk.Frame(fr, bg="white")
        top.pack(fill=tk.X)
        ic = tk.Label(top, text=icon, font=(EMOJI, 24), fg=INK if ok else MUTED, bg="white")
        ic.pack(side=tk.LEFT)
        t1 = tk.Label(top, text=name, font=(FAM, 19, "bold"), fg=INK if ok else MUTED, bg="white")
        t1.pack(side=tk.LEFT, padx=(8, 0))
        badges = tk.Frame(top, bg="white")
        badges.pack(side=tk.RIGHT, anchor="n")
        fr.check = tk.Label(badges, text="", font=(FAM, 12, "bold"), fg="white", bg=BLUE)
        fr.check.pack(side=tk.RIGHT, padx=(6, 0))
        if key == suggest:
            tk.Label(badges, text=" 建議 ", font=(FAM, 12, "bold"), fg="white", bg="#5B6472").pack(side=tk.RIGHT, padx=(6, 0))
        if not ok:                                    # 能用的不必標，沒標「製作中」就是可以用
            tk.Label(badges, text=" 製作中 ", font=(FAM, 12, "bold"), fg="white", bg=ORANGE).pack(side=tk.RIGHT)
        t2 = tk.Label(fr, text=f"適用：{fits}", font=(FAM, 13), fg=INK2 if ok else MUTED, bg="white", anchor="w")
        t2.pack(fill=tk.X, pady=(6, 0))
        t3 = tk.Label(fr, text=pages, font=(FAM, 12), fg=INK2 if ok else MUTED, bg="white", anchor="w",
                      justify="left", wraplength=int((W - 120) / cols) - 40)
        t3.pack(fill=tk.X, pady=(4, 0))
        parts = [top, ic, t1, badges, t2, t3]
        cards[key] = (fr, parts, ok)
        for w in [fr] + parts + [fr.check] + list(badges.winfo_children()):
            w.bind("<Button-1>", lambda e, k=key, o=ok: choose_tpl(k, o))

    # ---- 版面 ----
    hd2 = tk.Frame(body, bg=ui.BG)
    hd2.pack(fill=tk.X, pady=(16, 6))
    tk.Label(hd2, text="選擇版面", font=(FAM, 20, "bold"), fg=INK, bg=ui.BG).pack(side=tk.LEFT)
    tk.Label(hd2, text="要列印選 A4（直式照片多選 A4 直式）；在電視或投影機播放選 16:9", font=(FAM, 14), fg=INK2,
             bg=ui.BG).pack(side=tk.LEFT, padx=16)
    prow = tk.Frame(body, bg=ui.BG)
    prow.pack(fill=tk.X)
    for c in range(len(PAGE_CHOICES)):
        prow.grid_columnconfigure(c, weight=1, uniform="pg")
    ptiles = {}
    PH = int(78 * sc)                                 # 紙張圖示的框高

    def paint_pages():
        for key, (fr, cv, parts, ok, wh) in ptiles.items():
            sel = key == state["page"]
            bg = BLUE_BG if sel else ("white" if ok else "#F1F5F9")
            fr.configure(bg=bg, highlightbackground=BLUE if sel else ("#CBD5E1" if ok else "#E2E8F0"),
                         highlightthickness=3 if sel else 1)
            for w in parts:
                w.configure(bg=bg)
            cv.configure(bg=bg)
            cv.delete("all")
            pw, ph = wh
            k = min((PH - 8) / ph, (PH * 1.6) / pw)
            w_, h_ = pw * k, ph * k
            cx, cy = int(PH * 1.6) // 2 + 4, PH // 2
            x0, y0 = cx - w_ / 2, cy - h_ / 2
            cv.create_rectangle(x0 + 4, y0 + 4, x0 + w_ + 4, y0 + h_ + 4, fill="#CBD5E1", outline="")
            cv.create_rectangle(x0, y0, x0 + w_, y0 + h_, fill="white", outline=BLUE if sel else "#94A3B8", width=2)
            cv.create_rectangle(x0, y0, x0 + w_, y0 + max(6, h_ * 0.13), fill=BLUE if ok else "#94A3B8", outline="")
            for j in range(3):
                yy = y0 + h_ * (0.3 + j * 0.2)
                cv.create_rectangle(x0 + w_ * 0.1, yy, x0 + w_ * (0.9 - j * 0.15), yy + max(3, h_ * 0.07),
                                    fill="#D7E7FB" if ok else "#E2E8F0", outline="")

    def choose_page(key, ok):
        if not ok:
            ui.message(root, "製作中", "A4 直式還在製作中，等各版型橫式都做好後開放。")
            return
        state["page"] = key
        paint_pages()

    for i, (key, name, desc, pw, ph, ok) in enumerate(PAGE_CHOICES):
        fr = tk.Frame(prow, bg="white", highlightthickness=1, highlightbackground="#CBD5E1", padx=14, pady=12,
                      cursor="hand2" if ok else "arrow")
        fr.grid(row=0, column=i, padx=8, pady=6, sticky="nsew")
        cv = tk.Canvas(fr, width=int(PH * 1.6) + 10, height=PH, bg="white", highlightthickness=0)
        cv.pack(side=tk.LEFT)
        txt = tk.Frame(fr, bg="white")
        txt.pack(side=tk.LEFT, padx=(10, 0))
        a = tk.Label(txt, text=name, font=(FAM, 18, "bold"), fg=INK if ok else MUTED, bg="white", anchor="w")
        a.pack(anchor="w")
        b = tk.Label(txt, text=desc, font=(FAM, 13), fg=INK2 if ok else MUTED, bg="white", anchor="w")
        b.pack(anchor="w")
        c = tk.Label(txt, text="" if ok else "製作中", font=(FAM, 12, "bold"), fg=ORANGE, bg="white", anchor="w")
        c.pack(anchor="w")
        parts = [txt, a, b, c]
        ptiles[key] = (fr, cv, parts, ok, (pw, ph))
        for w in [fr, cv] + parts:
            w.bind("<Button-1>", lambda e, k=key, o=ok: choose_page(k, o))
    # ---- 色系 ----
    hd3 = tk.Frame(body, bg=ui.BG)
    hd3.pack(fill=tk.X, pady=(16, 6))
    tk.Label(hd3, text="選擇色系", font=(FAM, 20, "bold"), fg=INK, bg=ui.BG).pack(side=tk.LEFT)
    tk.Label(hd3, text="換頁首頁尾、卡片底色與重點字的顏色；照片、地籍圖不受影響", font=(FAM, 14), fg=INK2,
             bg=ui.BG).pack(side=tk.LEFT, padx=16)
    trow = tk.Frame(body, bg=ui.BG)
    trow.pack(fill=tk.X, pady=(0, 10))
    for c in range(len(THEMES)):
        trow.grid_columnconfigure(c, weight=1, uniform="th")
    ttiles = {}
    TW_, TH_ = int(150 * sc), int(84 * sc)            # 小色卡(一頁簡報的縮圖)

    def _mix(c1, c2, k):
        a = [int(c1[i:i + 2], 16) for i in (0, 2, 4)]
        b = [int(c2[i:i + 2], 16) for i in (0, 2, 4)]
        return "#%02X%02X%02X" % tuple(int(a[i] + (b[i] - a[i]) * k) for i in range(3))

    def draw_swatch(cv, t):
        cv.delete("all")
        head = t["head"] or ("0276AB", "E89AB3")       # 富住通原色：公版的藍→粉
        foot = t["foot"] or ("8E1D4F", "19408F")
        hh, fh = int(TH_ * 0.2), max(4, int(TH_ * 0.07))
        cv.create_rectangle(0, 0, TW_, TH_, fill="#" + t["bg"], outline="#CBD5E1")
        for x in range(0, TW_, 3):
            k = x / TW_
            cv.create_rectangle(x, 0, x + 3, hh, fill=_mix(*head, k), outline="")
            cv.create_rectangle(x, TH_ - fh, x + 3, TH_, fill=_mix(*foot, k), outline="")
        cv.create_rectangle(8, hh * 0.3, TW_ * 0.42, hh * 0.7, fill="white", outline="")     # 標題字
        for j in range(3):                                                            # 三張卡片
            x0 = 8 + j * (TW_ - 16) / 3
            cv.create_rectangle(x0 + 2, hh + 8, x0 + (TW_ - 16) / 3 - 4, hh + 8 + TH_ * 0.22,
                                fill="white", outline="#" + t["line"])
            cv.create_rectangle(x0 + 2, hh + 8, x0 + 6, hh + 8 + TH_ * 0.22,
                                fill="#B88E3E" if j == 0 else "#" + t["sky"], outline="")
        cv.create_rectangle(8, hh + 14 + TH_ * 0.24, TW_ - 8, TH_ - fh - 6, fill="#" + t["panel"], outline="")
        for j in range(2):
            yy = hh + 20 + TH_ * 0.24 + j * TH_ * 0.12
            cv.create_rectangle(16, yy, TW_ * 0.35, yy + 4, fill="#" + t["sky_dk"], outline="")
            cv.create_rectangle(TW_ * 0.4, yy, TW_ * 0.8, yy + 4, fill="#94A3B8", outline="")

    def paint_themes():
        for key, (fr, cv, parts) in ttiles.items():
            sel = key == state.get("theme", "fzt")
            bg = BLUE_BG if sel else "white"
            fr.configure(bg=bg, highlightbackground=BLUE if sel else "#CBD5E1", highlightthickness=3 if sel else 1)
            for w in parts:
                w.configure(bg=bg)
            parts[0].configure(text=("✓ " if sel else "") + THEMES[key]["name"], fg=BLUE if sel else INK)

    def choose_theme(key):
        state["theme"] = key
        paint_themes()

    for i, (key, t) in enumerate(THEMES.items()):
        fr = tk.Frame(trow, bg="white", highlightthickness=1, highlightbackground="#CBD5E1", padx=10, pady=10,
                      cursor="hand2")
        fr.grid(row=0, column=i, padx=6, pady=6, sticky="nsew")
        cv = tk.Canvas(fr, width=TW_, height=TH_, highlightthickness=0, bg="white")
        cv.pack()
        draw_swatch(cv, t)
        a = tk.Label(fr, text=t["name"], font=(FAM, 15, "bold"), fg=INK, bg="white")
        a.pack(pady=(6, 0))
        b = tk.Label(fr, text=t["desc"], font=(FAM, 12), fg=INK2, bg="white")
        b.pack()
        ttiles[key] = (fr, cv, [a, b])
        for w in (fr, cv, a, b):
            w.bind("<Button-1>", lambda e, k=key: choose_theme(k))
    paint_cards()
    paint_pages()
    paint_themes()

    # ---- 頁尾 ----
    tk.Label(foot, text="選好按「下一步」選頁面內容；之後也可以按「‹ 上一步」回來改", font=(FAM, 13), fg=INK2,
             bg=foot.cget("bg")).pack(side=tk.LEFT)
    ui.button(foot, "下一步：頁面內容  ›", lambda: done.set("next"), "primary").pack(side=tk.RIGHT)
    ui.button(foot, "取消", lambda: done.set("close"), "ghost").pack(side=tk.RIGHT, padx=10)
    root.bind("<Return>", lambda e: done.set("next"))
    root.wait_variable(done)
    root.unbind("<Return>")
    r = done.get()
    try:
        stage.destroy()
    except tk.TclError:
        pass
    return r


# ---------------------------------------------------------------- ② 挑照片
def _step_photos(root, photos, facts, state, sc, W):
    from marketing_deck import CATS, CAT_MENU, CAT_SEP, LAYOUT_CAT, PAGES, TEMPLATES, caption_of
    stage = tk.Frame(root, bg=ui.BG)
    stage.pack(fill=tk.BOTH, expand=True)
    done = tk.StringVar(master=root)
    TW, TH = int(210 * sc), int(158 * sc)
    tname = TEMPLATES.get(state["template"], TEMPLATES["building"])[1]
    ui.header(stage, f"🎞 行銷簡報（{tname}・{PAGES.get(state['page'], PAGES['a4l'])[0]}）",
              f"{facts.get('short') or facts.get('title', '')}　替照片選分類、照片下方可改說明（會記住，下次自動帶回）")
    foot = ui.footer(stage)
    _steps_bar(stage, 3)

    # 上方：分類統計
    top = tk.Frame(stage, bg=ui.BG, padx=16, pady=2)
    top.pack(fill=tk.X)
    stat = tk.Label(top, text="", font=(FAM, 14, "bold"), fg=INK, bg=ui.BG, justify="left", anchor="w")
    stat.pack(side=tk.LEFT, fill=tk.X, expand=True)
    tk.Label(top, text="⭐＝「廣告」資料夾（整理過的照片）", font=(FAM, 13), fg=INK2, bg=ui.BG).pack(side=tk.RIGHT)

    body, canvas = ui.scroll_body(stage)
    cols = max(3, (W - 60) // (TW + 24))
    tiles = []
    imgs = []

    def refresh(*_):
        cnt = {}
        for p, v, fr in tiles:
            c = v.get()
            p["cat"] = c
            cnt[c] = cnt.get(c, 0) + 1
            fr.configure(highlightbackground=COVER if c == "封面" else LAYOUT if c == LAYOUT_CAT else (USED if c != "不使用" else UNUSED),
                         highlightthickness=3 if c != "不使用" else 1)
        used = [f"{c} {cnt[c]}" for c in CATS if c != "不使用" and cnt.get(c)]
        stat.configure(text=("已選：" + "・".join(used)) if used else "還沒選任何照片（把分類從「不使用」改成要放的頁面）")

    blank = tk.PhotoImage(master=root, width=TW, height=TH)
    quiet = {"on": False}
    for i, p in enumerate(photos):
        fr = tk.Frame(body, bg="white", highlightthickness=1, highlightbackground=UNUSED, padx=6, pady=6)
        fr.grid(row=i // cols, column=i % cols, padx=8, pady=8, sticky="n")
        img_lbl = tk.Label(fr, bg="#E2E8F0", image=blank, bd=0)
        img_lbl.pack()
        name = ("⭐ " if p["ad"] else "") + p["name"]
        tk.Label(fr, text=name, font=(FAM, 12), fg=INK, bg="white", wraplength=TW, justify="left").pack(anchor="w", pady=(4, 2))
        v = tk.StringVar(master=root, value=p["cat"])
        cb = ttk.Combobox(fr, textvariable=v, values=CAT_MENU, state="readonly", height=26, font=(FAM, 13), width=10)
        cb.pack(anchor="w")
        _guard_sep(cb, v, p)
        v.trace_add("write", refresh)
        # 說明(照片上的小字條)：預設用檔名描述，可以改；清空＝不顯示說明
        cap = tk.StringVar(master=root, value=p["cap"] if "cap" in p else caption_of(p))
        tk.Entry(fr, textvariable=cap, font=(FAM, 12), width=max(12, TW // 14), fg=INK, relief="flat",
                 highlightthickness=1, highlightbackground=UNUSED, highlightcolor=USED).pack(anchor="w", fill=tk.X, pady=(4, 0), ipady=2)

        def on_cap(*a, p=p, cap=cap):
            if not quiet["on"]:                       # 使用者自己改的才記
                p["cap"] = cap.get()
        cap.trace_add("write", on_cap)

        def on_cat(*a, p=p, cap=cap, v=v):
            if "cap" not in p:                        # 沒改過說明 → 換分類時跟著換預設說明
                quiet["on"] = True
                p["cat"] = v.get()
                cap.set(caption_of(p))
                quiet["on"] = False
        v.trace_add("write", on_cat)
        tiles.append((p, v, fr))
        imgs.append(img_lbl)
    refresh()

    # 縮圖在背景載入(JPEG 用 draft 快速縮小)，載好再回主執行緒貼上
    q = queue.Queue()
    alive = {"on": True}

    def loader():
        for idx, p in enumerate(photos):
            if not alive["on"]:
                return
            try:
                im = Image.open(p["path"])
                im.draft("RGB", (TW * 2, TH * 2))
                im = ImageOps.exif_transpose(im).convert("RGB")
                im.thumbnail((TW, TH))
                q.put((idx, im))
            except Exception:
                q.put((idx, None))

    keep = {}

    def pump():
        if not alive["on"]:
            return
        try:
            while True:
                idx, im = q.get_nowait()
                if im is not None:
                    ph = ImageTk.PhotoImage(im)
                    keep[idx] = ph
                    imgs[idx].configure(image=ph)
        except (queue.Empty, tk.TclError):
            pass
        if len(keep) < len(photos):
            root.after(60, pump)
    threading.Thread(target=loader, daemon=True).start()
    root.after(60, pump)

    # 頁尾：按鈕先放(靠右)，路徑放剩下的空間(太長中間用…)
    out_var = tk.StringVar(master=root, value=state["out"])
    shown = tk.StringVar(master=root)

    def _short(p, n=52):
        return p if len(p) <= n else p[:24] + " … " + p[-(n - 27):]
    out_var.trace_add("write", lambda *a: shown.set(_short(out_var.get())))
    shown.set(_short(state["out"]))
    btns = tk.Frame(foot, bg=foot.cget("bg"))
    btns.pack(side=tk.RIGHT)
    left = tk.Frame(foot, bg=foot.cget("bg"))
    left.pack(side=tk.LEFT, fill=tk.X, expand=True)
    tk.Label(left, text="存到：", font=(FAM, 13, "bold"), fg=INK, bg=foot.cget("bg")).pack(side=tk.LEFT)
    tk.Label(left, textvariable=shown, font=(FAM, 12), fg=INK2, bg=foot.cget("bg"), anchor="w").pack(side=tk.LEFT)

    def change_out():
        p = filedialog.asksaveasfilename(parent=root, title="簡報存檔位置", defaultextension=".pptx",
                                         initialdir=os.path.dirname(out_var.get()), initialfile=os.path.basename(out_var.get()),
                                         filetypes=[("PowerPoint 簡報", "*.pptx")])
        if p:
            out_var.set(os.path.normpath(p))

    result = {}

    def go():
        done.set("next")

    root.protocol("WM_DELETE_WINDOW", lambda: done.set("close"))
    ui.button(btns, "‹ 上一步", lambda: done.set("back"), "ghost").pack(side=tk.LEFT, padx=(0, 8))
    ui.button(btns, "變更位置…", change_out, "ghost").pack(side=tk.LEFT, padx=(0, 8))
    ui.button(btns, "取消", lambda: done.set("close"), "ghost").pack(side=tk.LEFT, padx=(0, 8))
    ui.button(btns, "下一步：整理與確認  ›", go, "primary").pack(side=tk.LEFT)
    root.wait_variable(done)
    alive["on"] = False
    for p, v, _ in tiles:
        p["cat"] = v.get()
    state["out"] = out_var.get()
    r = done.get()
    try:
        stage.destroy()
    except tk.TclError:
        pass
    return r if r in ("back", "next") else None


# ---------------------------------------------------------------- ② 頁面內容
def _thumb(path, w, h, zoom_w=None):
    """縮圖(PDF 取第一頁)→ PIL Image；讀不到回 None"""
    try:
        if path.lower().endswith(".pdf"):
            import fitz
            doc = fitz.open(path)
            pg = doc[0]
            z = (zoom_w or w * 2) / max(pg.rect.width, 1)
            pix = pg.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
            im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            doc.close()
        else:
            im = Image.open(path)
            im.draft("RGB", (w * 2, h * 2))
            im = ImageOps.exif_transpose(im).convert("RGB")
        im.thumbnail((w, h))
        return im
    except Exception:
        return None


def _preview(root, path, title):
    """點縮圖 → 大圖預覽：開啟就最大化；滑鼠滾輪縮放(以游標為中心)、按住拖曳移動；
    下方「－／＋／適合視窗／原尺寸」；雙擊＝在適合視窗與原尺寸之間切換；可用預設程式開原檔"""
    top = tk.Toplevel(root)
    top.title(f"預覽：{title}")
    top.configure(bg="#1E293B")
    try:
        top.state("zoomed")                           # 最大化
    except tk.TclError:
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        top.geometry(f"{sw}x{sh - 60}+0+0")
    top.transient(root)
    bar = tk.Frame(top, bg="#1E293B", padx=12, pady=8)
    bar.pack(fill=tk.X, side=tk.BOTTOM)
    tk.Label(bar, text=os.path.basename(path), font=(FAM, 13, "bold"), fg="white", bg="#1E293B").pack(side=tk.LEFT)
    zl = tk.Label(bar, text="", font=(FAM, 13, "bold"), fg="#93C5FD", bg="#1E293B")
    zl.pack(side=tk.LEFT, padx=(16, 0))
    tk.Label(bar, text="滾輪縮放・按住拖曳移動・雙擊切換大小", font=(FAM, 12), fg="#CBD5E1", bg="#1E293B").pack(side=tk.LEFT,
                                                                                                  padx=(16, 0))
    cv = tk.Canvas(top, bg="#334155", highlightthickness=0, cursor="fleur")
    cv.pack(fill=tk.BOTH, expand=True)
    msg = cv.create_text(40, 40, text="載入中…", fill="white", font=(FAM, 16), anchor="nw")
    top.update_idletasks()
    # 原圖：PDF 用高倍率轉一次(放大也清楚)，之後縮放都從這張縮
    src = None
    try:
        if path.lower().endswith(".pdf"):
            import fitz
            doc = fitz.open(path)
            pg = doc[0]
            z = min(4.0, 3600 / max(pg.rect.width, pg.rect.height, 1))
            pix = pg.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
            src = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            doc.close()
        else:
            src = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    except Exception:
        src = None
    if src is None:
        cv.itemconfigure(msg, text="這個檔案無法預覽，請用「用預設程式開啟原檔」")
        ui.button(bar, "關閉", top.destroy, "ghost").pack(side=tk.RIGHT)
        return
    cv.delete(msg)
    st = {"z": 1.0, "fit": 1.0, "x": 0.0, "y": 0.0, "img": None, "job": None, "drag": None}

    def fit_scale():
        cw, ch = max(cv.winfo_width(), 200), max(cv.winfo_height(), 200)
        return min((cw - 30) / src.width, (ch - 30) / src.height)

    def render(fast=False):
        cw, ch = max(cv.winfo_width(), 200), max(cv.winfo_height(), 200)
        z = st["z"]
        # 只取畫面看得到的那一塊來縮放(放很大也不吃記憶體)
        x0 = max(0, int(-st["x"] / z)); y0 = max(0, int(-st["y"] / z))
        x1 = min(src.width, int((cw - st["x"]) / z) + 1); y1 = min(src.height, int((ch - st["y"]) / z) + 1)
        cv.delete("img")
        if x1 <= x0 or y1 <= y0:
            return
        crop = src.crop((x0, y0, x1, y1))
        w_, h_ = max(1, int((x1 - x0) * z)), max(1, int((y1 - y0) * z))
        crop = crop.resize((w_, h_), Image.BILINEAR if fast else Image.LANCZOS)
        st["img"] = ImageTk.PhotoImage(crop)
        cv.create_image(st["x"] + x0 * z, st["y"] + y0 * z, image=st["img"], anchor="nw", tags="img")
        zl.configure(text=f"{z / st['fit'] * 100:.0f}%（適合視窗＝100%）")

    def later():
        if st["job"]:
            top.after_cancel(st["job"])
        st["job"] = top.after(180, render)            # 停下來再用高品質重畫

    def set_zoom(nz, cx=None, cy=None):
        nz = max(st["fit"] * 0.5, min(nz, st["fit"] * 12, 8.0))
        cw, ch = cv.winfo_width(), cv.winfo_height()
        cx = cw / 2 if cx is None else cx
        cy = ch / 2 if cy is None else cy
        k = nz / st["z"]
        st["x"] = cx - (cx - st["x"]) * k
        st["y"] = cy - (cy - st["y"]) * k
        st["z"] = nz
        render(fast=True)
        later()

    def fit():
        st["fit"] = st["z"] = fit_scale()
        cw, ch = cv.winfo_width(), cv.winfo_height()
        st["x"] = (cw - src.width * st["z"]) / 2
        st["y"] = (ch - src.height * st["z"]) / 2
        render()

    def actual():
        set_zoom(1.0)

    def wheel(e):
        set_zoom(st["z"] * (1.2 if e.delta > 0 else 1 / 1.2), e.x, e.y)

    def press(e):
        st["drag"] = (e.x, e.y, st["x"], st["y"])

    def move(e):
        if st["drag"]:
            sx, sy, ox, oy = st["drag"]
            st["x"], st["y"] = ox + e.x - sx, oy + e.y - sy
            render(fast=True)
            later()

    def dbl(e):
        if abs(st["z"] - st["fit"]) < 1e-3:
            set_zoom(min(1.0, st["fit"] * 3) if st["fit"] < 1 else st["fit"] * 2.5, e.x, e.y)
        else:
            fit()
    cv.bind("<MouseWheel>", wheel)
    cv.bind("<ButtonPress-1>", press)
    cv.bind("<B1-Motion>", move)
    cv.bind("<ButtonRelease-1>", lambda e: st.update(drag=None))
    cv.bind("<Double-Button-1>", dbl)
    cv.bind("<Configure>", lambda e: (fit() if not st["img"] else later()))
    top.bind("<Escape>", lambda e: top.destroy())
    top.bind("<plus>", lambda e: set_zoom(st["z"] * 1.25))
    top.bind("<equal>", lambda e: set_zoom(st["z"] * 1.25))
    top.bind("<minus>", lambda e: set_zoom(st["z"] / 1.25))
    ui.button(bar, "關閉", top.destroy, "ghost").pack(side=tk.RIGHT)
    ui.button(bar, "用預設程式開啟原檔", lambda: os.startfile(path), "ghost").pack(side=tk.RIGHT, padx=(0, 8))
    ui.button(bar, "原尺寸", actual, "ghost").pack(side=tk.RIGHT, padx=(0, 8))
    ui.button(bar, "適合視窗", fit, "ghost").pack(side=tk.RIGHT, padx=(0, 8))
    ui.button(bar, "＋", lambda: set_zoom(st["z"] * 1.25), "ghost", padx=14).pack(side=tk.RIGHT, padx=(0, 8))
    ui.button(bar, "－", lambda: set_zoom(st["z"] / 1.25), "ghost", padx=14).pack(side=tk.RIGHT, padx=(0, 8))
    top.focus_set()
    top.after(80, fit)


def _step_sections(root, state, facts, sc, W):
    """每個單元一張卡：大勾選＝要不要這頁；有好幾份資料的列出縮圖(點縮圖放大預覽)，勾選要用哪幾份"""
    from marketing_deck import PAGES, TEMPLATES
    stage = tk.Frame(root, bg=ui.BG)
    stage.pack(fill=tk.BOTH, expand=True)
    done = tk.StringVar(master=root)
    tname = TEMPLATES.get(state["template"], TEMPLATES["building"])[1]
    ui.header(stage, f"🎞 行銷簡報（{tname}・{PAGES.get(state['page'], PAGES['a4l'])[0]}）",
              f"{facts.get('short') or facts.get('title', '')}　勾選要放哪些頁、用哪份資料（點縮圖可放大預覽；會記住）")
    foot = ui.footer(stage)
    _steps_bar(stage, 2)
    tk.Label(stage, text="封面、基本資料、照片頁一定會有；照片在下一步挑。", font=(FAM, 13), fg=MUTED, bg=ui.BG,
             anchor="w").pack(fill=tk.X, padx=20)
    body, canvas = ui.scroll_body(stage)
    TW, TH = int(150 * sc), int(112 * sc)
    cols = max(3, (W - 120) // (TW + 30))
    jobs = []                 # (縮圖 Label, 路徑) → 背景載入
    keep = []
    q = queue.Queue()
    alive = {"on": True}
    blank = tk.PhotoImage(master=root, width=TW, height=TH)      # 縮圖還沒載好前的空白(要留著參照，否則被回收)
    keep.append(blank)

    for sec in state["sections"]:
        card = tk.Frame(body, bg="white", highlightthickness=1, highlightbackground=UNUSED, padx=14, pady=10)
        card.pack(fill=tk.X, pady=6)
        head = tk.Frame(card, bg="white")
        head.pack(fill=tk.X)
        on = tk.BooleanVar(master=root, value=sec["on"])
        ui.big_check(head, on, enabled=not sec["empty"]).pack(side=tk.LEFT)
        tl = tk.Label(head, text=sec["title"], font=(FAM, 17, "bold"), fg=INK if not sec["empty"] else MUTED, bg="white",
                      cursor="hand2")
        tl.pack(side=tk.LEFT, padx=(10, 0))
        tk.Label(head, text=sec["empty"] or sec["desc"], font=(FAM, 13), fg=ORANGE if sec["empty"] else INK2, bg="white",
                 wraplength=W - 420, justify="left").pack(side=tk.LEFT, padx=(14, 0))

        def paint(*_, card=card, on=on):
            card.configure(highlightbackground=USED if on.get() else UNUSED, highlightthickness=2 if on.get() else 1)
        if not sec["empty"]:
            tl.bind("<Button-1>", lambda e, on=on: on.set(not on.get()))

        def on_toggle(*_, sec=sec, on=on):
            sec["on"] = on.get()
            paint()
        on.trace_add("write", on_toggle)
        paint()
        if not sec["cands"]:
            continue
        grid = tk.Frame(card, bg="white")
        grid.pack(fill=tk.X, pady=(8, 0), padx=(40, 0))
        for i, (path, label) in enumerate(sec["cands"]):
            cell = tk.Frame(grid, bg="white", padx=4, pady=4)
            cell.grid(row=i // cols, column=i % cols, sticky="n", padx=4, pady=4)
            th = tk.Label(cell, bg="#E2E8F0", cursor="hand2", bd=0, image=blank)
            th.pack()
            th.bind("<Button-1>", lambda e, p=path, t=label: _preview(root, p, t))
            jobs.append((th, path))
            v = tk.BooleanVar(master=root, value=path in sec["pick"])
            row = tk.Frame(cell, bg="white")
            row.pack(fill=tk.X, pady=(4, 0))
            ui.big_check(row, v).pack(side=tk.LEFT)
            tk.Label(row, text=label, font=(FAM, 11, "bold"), fg=INK, bg="white", wraplength=TW - 30, justify="left",
                     anchor="w").pack(side=tk.LEFT, padx=(4, 0))

            def on_pick(*_, sec=sec, path=path, v=v, on=on):
                if v.get():
                    if path not in sec["pick"]:
                        sec["pick"].append(path)
                    if not sec["multi"]:              # 單選：勾這個就取消別的
                        for o_path, o_var in sec["_vars"]:
                            if o_path != path and o_var.get():
                                o_var.set(False)
                    if not on.get():
                        on.set(True)
                else:
                    if path in sec["pick"]:
                        sec["pick"].remove(path)
            v.trace_add("write", on_pick)
            sec.setdefault("_vars", []).append((path, v))
        tk.Label(card, text="（點縮圖放大預覽）" + ("　可勾選多份，每份一頁" if sec["multi"] else "　只能選一份"),
                 font=(FAM, 11), fg=MUTED, bg="white").pack(anchor="w", padx=(44, 0))

    def loader():
        for th, path in jobs:
            if not alive["on"]:
                return
            q.put((th, _thumb(path, TW, TH)))

    def pump():
        if not alive["on"]:
            return
        try:
            while True:
                th, im = q.get_nowait()
                if im is not None:
                    ph = ImageTk.PhotoImage(im)
                    keep.append(ph)
                    try:
                        th.configure(image=ph)
                    except tk.TclError:
                        pass
        except queue.Empty:
            pass
        root.after(80, pump)
    threading.Thread(target=loader, daemon=True).start()
    root.after(80, pump)

    root.protocol("WM_DELETE_WINDOW", lambda: done.set("close"))
    btns = tk.Frame(foot, bg=foot.cget("bg"))
    btns.pack(side=tk.RIGHT)
    ui.button(btns, "‹ 上一步", lambda: done.set("back"), "ghost").pack(side=tk.LEFT, padx=(0, 8))
    ui.button(btns, "取消", lambda: done.set("close"), "ghost").pack(side=tk.LEFT, padx=(0, 8))
    ui.button(btns, "下一步：挑照片  ›", lambda: done.set("next"), "primary").pack(side=tk.LEFT)
    root.wait_variable(done)
    alive["on"] = False
    for sec in state["sections"]:
        sec.pop("_vars", None)
        if sec["cands"] and sec["on"] and not sec["pick"]:      # 勾了單元卻一份都沒選 → 當作不要這頁
            sec["on"] = False
    r = done.get()
    try:
        stage.destroy()
    except tk.TclError:
        pass
    return r if r in ("back", "next") else None


# ---------------------------------------------------------------- ③ 整理與確認
SEL_BG = "#EAF3FF"      # 已選用區(淡藍)
OFF_BG = "#EEF1F5"      # 不使用區(淡灰)


def _step_arrange(root, photos, facts, state, sc, W):
    from marketing_deck import CATS, CAT_MENU, CAT_SEP, LAYOUT_CAT, PAGES, TEMPLATES, PHOTO_PAGES, STREET_PHOTO_PAGES, caption_of, guess_cat, ordered, split_pages
    if state.get("template") == "villa":                # 透天・店面版的照片頁分法不同(外觀街景在前)
        PHOTO_PAGES = STREET_PHOTO_PAGES
    stage = tk.Frame(root, bg=ui.BG)
    stage.pack(fill=tk.BOTH, expand=True)
    done = tk.StringVar(master=root)
    TW, TH = int(168 * sc), int(118 * sc)            # 已選用的縮圖
    SW_, SH_ = int(120 * sc), int(84 * sc)           # 不使用的縮圖(小一點)
    tname = TEMPLATES.get(state["template"], TEMPLATES["building"])[1]
    ui.header(stage, f"🎞 行銷簡報（{tname}・{PAGES.get(state['page'], PAGES['a4l'])[0]}）",
              f"{facts.get('short') or facts.get('title', '')}　確認每頁放哪些照片、順序、封面與說明，好了就產生")
    foot = ui.footer(stage)
    _steps_bar(stage, 4)

    # 上方工具列：封面顯示專案人員
    tool = tk.Frame(stage, bg=ui.BG, padx=16, pady=4)
    tool.pack(fill=tk.X)
    ag_var = tk.BooleanVar(master=root, value=state.get("show_agents", True))
    ui.big_check(tool, ag_var).pack(side=tk.LEFT)
    names = "、".join(facts.get("agents") or []) or "（物件編輯器沒有承辦人）"
    lab = tk.Label(tool, text=f"封面顯示專案人員：{names}", font=(FAM, 14, "bold"), fg=INK, bg=ui.BG, cursor="hand2")
    lab.pack(side=tk.LEFT, padx=(8, 0))
    lab.bind("<Button-1>", lambda e: ag_var.set(not ag_var.get()))
    # 封底(直式 LOGO＋Full Life＋公司電話)：使用者自己選要不要
    bc_var = tk.BooleanVar(master=root, value=state.get("back_cover", False))
    ui.big_check(tool, bc_var).pack(side=tk.LEFT, padx=(28, 0))
    lab2 = tk.Label(tool, text="加封底（公司 LOGO、Full Life、電話）", font=(FAM, 14, "bold"), fg=INK, bg=ui.BG, cursor="hand2")
    lab2.pack(side=tk.LEFT, padx=(8, 0))
    lab2.bind("<Button-1>", lambda e: bc_var.set(not bc_var.get()))

    body, canvas = ui.scroll_body(stage)
    thumbs = {}               # (路徑, 小圖?) → PhotoImage
    labels = {}               # (路徑, 小圖?) → 目前畫面上的縮圖 Label
    pending = set()
    q = queue.Queue()
    alive = {"on": True}

    def _load(path, small):
        try:
            im = Image.open(path)
            w, h = (SW_, SH_) if small else (TW, TH)
            im.draft("RGB", (w * 2, h * 2))
            im = ImageOps.exif_transpose(im).convert("RGB")
            im.thumbnail((w, h))
            return im
        except Exception:
            return None

    def want_thumb(path, small):
        key = (path, small)
        if key in thumbs:
            return thumbs[key]
        if key not in pending:
            pending.add(key)
            threading.Thread(target=lambda: q.put((key, _load(path, small))), daemon=True).start()
        return None

    def pump():
        if not alive["on"]:
            return
        try:
            while True:
                key, im = q.get_nowait()
                if im is not None:
                    thumbs[key] = ImageTk.PhotoImage(im)
                    lb = labels.get(key)
                    if lb is not None:
                        try:
                            lb.configure(image=thumbs[key])
                        except tk.TclError:
                            pass
        except queue.Empty:
            pass
        root.after(80, pump)
    root.after(80, pump)
    blank_big = tk.PhotoImage(master=root, width=TW, height=TH)
    blank_small = tk.PhotoImage(master=root, width=SW_, height=SH_)

    def next_order():
        return max([p.get("order", 0) for p in photos] + [0]) + 1

    def set_cat(p, cat):
        if cat == p["cat"]:
            return
        if cat == "封面":                             # 封面只有一張：原本的封面回到它之前的分類
            for o in photos:
                if o is not p and o["cat"] == "封面":
                    o["cat"] = o.pop("prev", None) or guess_cat(o["name"])
                    if o["cat"] == "封面":
                        o["cat"] = "不使用"
            p["prev"] = p["cat"] if p["cat"] != "不使用" else "陽台景觀"
        p["cat"] = cat
        p["order"] = next_order()
        render()

    def move(p, items, d):
        """在同一頁裡往前/往後移一格"""
        for i, it in enumerate(items):
            it["order"] = i + 1
        i = items.index(p)
        j = i + d
        if 0 <= j < len(items):
            items[i]["order"], items[j]["order"] = items[j]["order"], items[i]["order"]
        render()

    # ===== 滑鼠拖曳排序 =====
    # 拖照片縮圖：放到別張照片上 → 插到那張的位置(跨頁就換到那一頁)；放到某頁標題/空白處 → 排到那頁最後；
    # 放到「不使用」區 → 不使用。按住不動直接放開 = 沒事(不會誤觸)
    tile_of = {}              # 照片框 Frame → 照片
    zone_of = {}              # 頁面區塊 Frame → 那頁的分類(tuple)；不使用區 → ("不使用",)
    drag = {"p": None, "x": 0, "y": 0, "on": False, "ghost": None, "hl": None}

    def _find(widget, table):
        w = widget
        while w is not None:
            if w in table:
                return w
            w = getattr(w, "master", None)
        return None

    def _highlight(fr):
        old = drag["hl"]
        if old is fr:
            return
        if old is not None:
            try:
                old.configure(highlightbackground=old._orig_hl, highlightthickness=old._orig_th)
            except tk.TclError:
                pass
        drag["hl"] = fr
        if fr is not None:
            try:
                fr._orig_hl = fr.cget("highlightbackground")
                fr._orig_th = fr.cget("highlightthickness")
                fr.configure(highlightbackground="#E8590C", highlightthickness=max(3, fr._orig_th))
            except tk.TclError:
                drag["hl"] = None

    def drag_start(e, p):
        drag.update(p=p, x=e.x_root, y=e.y_root, on=False)

    def drag_motion(e):
        p = drag["p"]
        if p is None:
            return
        if not drag["on"]:
            if abs(e.x_root - drag["x"]) + abs(e.y_root - drag["y"]) < 8:
                return
            drag["on"] = True
            g = tk.Toplevel(root)
            g.overrideredirect(True)
            g.attributes("-topmost", True)
            try:
                g.attributes("-alpha", 0.85)
            except tk.TclError:
                pass
            img = thumbs.get((p["path"], True)) or thumbs.get((p["path"], False))
            tk.Label(g, image=img, text="" if img else p["name"], font=(FAM, 12, "bold"), bg="#E8590C", fg="white",
                     bd=3, relief="solid").pack()
            drag["ghost"] = g
            root.configure(cursor="fleur")
        drag["ghost"].geometry(f"+{e.x_root + 14}+{e.y_root + 10}")
        w = root.winfo_containing(e.x_root, e.y_root)
        hit = _find(w, tile_of) if w else None
        if hit is None and w is not None:
            hit = _find(w, zone_of)
        _highlight(hit if (hit is None or tile_of.get(hit) is not p) else None)
        # 拖到捲動區的上下邊緣 → 自動捲動
        cy = canvas.winfo_rooty()
        ch = canvas.winfo_height()
        if e.y_root < cy + 50:
            canvas.yview_scroll(-1, "units")
        elif e.y_root > cy + ch - 50:
            canvas.yview_scroll(1, "units")

    def drag_end(e):
        p, was_on = drag["p"], drag["on"]
        drag.update(p=None, on=False)
        if drag["ghost"] is not None:
            drag["ghost"].destroy()
            drag["ghost"] = None
        root.configure(cursor="")
        _highlight(None)
        if p is None or not was_on:
            return
        w = root.winfo_containing(e.x_root, e.y_root)
        if w is None:
            return
        t = _find(w, tile_of)
        if t is not None:
            drop_on_photo(p, tile_of[t])
            return
        z = _find(w, zone_of)
        if z is not None:
            drop_on_zone(p, zone_of[z])

    def _group_cats(cat):
        if cat == "封面":
            return ("封面",)
        for _t, cats, _c in PHOTO_PAGES:
            if cat in cats:
                return tuple(cats)
        return (cat,)

    def drop_on_photo(p, q):
        if q is p:
            return
        if q["cat"] == "不使用":
            set_cat(p, "不使用")
            return
        if q["cat"] == "封面":
            set_cat(p, "封面")
            return
        cats = _group_cats(q["cat"])
        items = ordered([o for o in photos if o["cat"] in cats])
        same = p in items
        before = same and items.index(p) < items.index(q)
        if p["cat"] not in cats:                      # 跨頁：換成目標那頁(用目標照片的分類)
            if p["cat"] == "封面":
                p.pop("prev", None)
            p["cat"] = q["cat"]
        rest = [o for o in items if o is not p]
        k = rest.index(q) + (1 if before else 0)       # 往後拖放在目標後面、往前拖放在目標前面
        rest.insert(k, p)
        for i, o in enumerate(rest):
            o["order"] = i + 1
        render()

    def drop_on_zone(p, cats):
        if cats == ("不使用",):
            set_cat(p, "不使用")
            return
        if cats == ("封面",):
            set_cat(p, "封面")
            return
        if p["cat"] in cats:                          # 同一頁 → 排到最後
            p["order"] = next_order()
            render()
            return
        if p["cat"] == "封面":
            p.pop("prev", None)
        p["cat"] = cats[0]
        p["order"] = next_order()
        render()

    def draggable(widget, p):
        widget.configure(cursor="fleur")
        widget.bind("<ButtonPress-1>", lambda e, p=p: drag_start(e, p))
        widget.bind("<B1-Motion>", drag_motion)
        widget.bind("<ButtonRelease-1>", drag_end)

    def add_photos():
        paths = filedialog.askopenfilenames(parent=root, title="新增照片",
                                            filetypes=[("照片", "*.jpg *.jpeg *.png *.webp *.bmp"), ("全部", "*.*")])
        have = {p["name"].lower() for p in photos}
        n = 0
        for path in paths:
            nm = os.path.basename(path)
            if nm.lower() in have:
                continue
            cat = guess_cat(nm)
            photos.append({"path": path, "name": nm, "ad": False, "added": True,
                           "cat": cat if cat != "不使用" else "其他空間", "order": next_order()})
            have.add(nm.lower())
            n += 1
        if n:
            render()

    def remove_added(p):
        if ui.message(root, "移除照片", f"把「{p['name']}」從清單移除？\n（只是不放進簡報，原本的照片檔不會刪掉）", kind="ask", ok="移除"):
            photos.remove(p)
            render()

    def caption_entry(parent, p):
        cap = tk.StringVar(master=root, value=p["cap"] if "cap" in p else caption_of(p))
        e = tk.Entry(parent, textvariable=cap, font=(FAM, 12), fg=INK, relief="flat", highlightthickness=1,
                     highlightbackground=UNUSED, highlightcolor=USED, width=max(10, TW // 13))
        e.pack(fill=tk.X, pady=(4, 0), ipady=2)
        cap.trace_add("write", lambda *a: p.__setitem__("cap", cap.get()))

    def small_btn(parent, text, cmd, fg=INK, bg="#E2E8F0"):
        b = tk.Button(parent, text=text, command=cmd, font=(FAM, 12, "bold"), fg=fg, bg=bg, activebackground="#CBD5E1",
                      relief="flat", bd=0, padx=8, pady=2, cursor="hand2")
        b.pack(side=tk.LEFT, padx=(0, 4))
        return b

    def sel_tile(parent, p, items, idx, over):
        color = COVER if p["cat"] == "封面" else LAYOUT if p["cat"] == LAYOUT_CAT else (ORANGE if over else USED)
        fr = tk.Frame(parent, bg="white", highlightthickness=2, highlightbackground=color, padx=6, pady=6)
        img = want_thumb(p["path"], False)
        lb = tk.Label(fr, image=img or blank_big, bg="#E2E8F0", bd=0)
        lb.pack()
        labels[(p["path"], False)] = lb
        tile_of[fr] = p
        draggable(lb, p)
        top = tk.Frame(fr, bg="white")
        top.pack(fill=tk.X, pady=(4, 0))
        tag = "封面" if p["cat"] == "封面" else f"{idx + 1}"
        tk.Label(top, text=f" {tag} ", font=(FAM, 12, "bold"), fg="white", bg=color).pack(side=tk.LEFT)
        nm = ("⭐" if p["ad"] else "") + p["name"]
        tk.Label(top, text=nm if len(nm) <= 13 else nm[:12] + "…", font=(FAM, 11), fg=INK2, bg="white").pack(side=tk.LEFT, padx=4)
        caption_entry(fr, p)
        v = tk.StringVar(master=root, value=p["cat"])
        cb = ttk.Combobox(fr, textvariable=v, values=CAT_MENU, state="readonly", height=26, font=(FAM, 12), width=10)
        cb.pack(fill=tk.X, pady=(4, 0))
        cb.bind("<<ComboboxSelected>>", lambda e: set_cat(p, v.get()) if v.get() != CAT_SEP else v.set(p["cat"]))
        row = tk.Frame(fr, bg="white")
        row.pack(fill=tk.X, pady=(4, 0))
        if items is not None:
            small_btn(row, "◀", lambda: move(p, items, -1))
            small_btn(row, "▶", lambda: move(p, items, +1))
            small_btn(row, "★封面", lambda: set_cat(p, "封面"), fg=GOLD_FG, bg="#FBF1D9")
        small_btn(row, "✕", (lambda: remove_added(p)) if p.get("added") else (lambda: set_cat(p, "不使用")),
                  fg="#B42318", bg="#FDE2E1")
        return fr

    def render():
        labels.clear()
        tile_of.clear()
        zone_of.clear()
        for w in body.winfo_children():
            w.destroy()
        cols = max(3, (W - 70) // (TW + 30))
        # ===== 已選用 =====
        sel = tk.Frame(body, bg=SEL_BG, padx=12, pady=10, highlightthickness=1, highlightbackground="#B9D4F5")
        sel.pack(fill=tk.X, pady=(0, 12))
        used = [p for p in photos if p["cat"] != "不使用"]
        tk.Label(sel, text=f"✔ 已選用（{len(used)} 張）— 依簡報頁面順序；可用滑鼠拖曳照片排序或換頁", font=(FAM, 17, "bold"), fg=BLUE, bg=SEL_BG).pack(anchor="w")
        groups = [("封面", ("封面",), 1), (LAYOUT_CAT, (LAYOUT_CAT,), 99)] + list(PHOTO_PAGES)
        for title, cats, cap in groups:
            items = ordered([p for p in photos if p["cat"] in cats])
            gh = tk.Frame(sel, bg=SEL_BG)
            gh.pack(fill=tk.X, pady=(10, 4))
            tk.Label(gh, text=title, font=(FAM, 15, "bold"), fg=INK, bg=SEL_BG).pack(side=tk.LEFT)
            hint = f"{len(items)} 張"
            if title == "封面" and not items:
                hint = "還沒指定 → 會自動用第一張陽台景觀或外觀照（按照片上的 ★封面 指定）"
            elif title == LAYOUT_CAT:
                hint = (f"{len(items)} 張，每張自己一頁（左圖、右邊格局重點），放在照片頁前面" if items else
                        "沒有選 → 會自動找 0.謄本／4.其他相關 裡檔名有「平面圖、格局」的圖")
            elif title != "封面" and len(items) > cap:
                hint += f"（每頁最多 {cap} 張，會平均分成 {len(split_pages(len(items), cap))} 頁：{'＋'.join(map(str, split_pages(len(items), cap)))} 張）"
            tk.Label(gh, text=hint, font=(FAM, 13), fg=ORANGE if (title != "封面" and len(items) > cap) else INK2,
                     bg=SEL_BG).pack(side=tk.LEFT, padx=10)
            gh.configure(highlightthickness=3, highlightbackground=SEL_BG)
            zone_of[gh] = tuple(cats)
            if not items:
                continue
            g = tk.Frame(sel, bg=SEL_BG, highlightthickness=3, highlightbackground=SEL_BG)
            g.pack(fill=tk.X)
            zone_of[g] = tuple(cats)
            for c in range(cols):
                g.grid_columnconfigure(c, uniform="sel", minsize=TW + 16)
            page_of = [k for k, n in enumerate(split_pages(len(items), cap) if title != "封面" else [len(items)]) for _ in range(n)]
            for i, p in enumerate(items):
                over = title != "封面" and page_of[i] % 2 == 1          # 接續頁的照片用橘框標出
                sel_tile(g, p, items if title != "封面" else None, i, over).grid(
                    row=i // cols, column=i % cols, padx=6, pady=6, sticky="nsew")
        # ===== 不使用 =====
        off = tk.Frame(body, bg=OFF_BG, padx=12, pady=10, highlightthickness=1, highlightbackground="#CBD5E1")
        off.pack(fill=tk.X)
        zone_of[off] = ("不使用",)
        unused = [p for p in photos if p["cat"] == "不使用"]
        tk.Label(off, text=f"不使用（{len(unused)} 張）— 選一頁、或直接拖到上面那一頁", font=(FAM, 17, "bold"), fg=MUTED, bg=OFF_BG).pack(anchor="w")
        g2 = tk.Frame(off, bg=OFF_BG)
        g2.pack(fill=tk.X, pady=(6, 0))
        c2 = max(4, (W - 70) // (SW_ + 22))
        for c in range(c2):
            g2.grid_columnconfigure(c, uniform="off", minsize=SW_ + 10)
        for i, p in enumerate(unused):
            fr = tk.Frame(g2, bg="white", highlightthickness=1, highlightbackground="#CBD5E1", padx=4, pady=4)
            fr.grid(row=i // c2, column=i % c2, padx=4, pady=4, sticky="nsew")
            img = want_thumb(p["path"], True)
            lb = tk.Label(fr, image=img or blank_small, bg="#E2E8F0", bd=0)
            lb.pack()
            labels[(p["path"], True)] = lb
            tile_of[fr] = p
            draggable(lb, p)
            tk.Label(fr, text=(("⭐" if p["ad"] else "") + p["name"])[:14], font=(FAM, 10), fg=INK2, bg="white").pack(anchor="w")
            v = tk.StringVar(master=root, value="放到…")
            cb = ttk.Combobox(fr, textvariable=v, values=CAT_MENU[:-2], state="readonly", height=26,
                              font=(FAM, 11), width=8)
            cb.pack(fill=tk.X)
            cb.bind("<<ComboboxSelected>>", lambda e, p=p, v=v: set_cat(p, v.get()) if v.get() != CAT_SEP else v.set("放到…"))
        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
    render()

    root.protocol("WM_DELETE_WINDOW", lambda: done.set("close"))
    btns = tk.Frame(foot, bg=foot.cget("bg"))
    btns.pack(side=tk.RIGHT)
    tk.Label(foot, text="照片會依頁面大小自動壓縮（清晰、檔案不肥）", font=(FAM, 13), fg=INK2, bg=foot.cget("bg")).pack(side=tk.LEFT)
    ui.button(btns, "‹ 上一步", lambda: done.set("back"), "ghost").pack(side=tk.LEFT, padx=(0, 8))
    ui.button(btns, "＋ 新增照片…", add_photos, "ghost").pack(side=tk.LEFT, padx=(0, 8))
    ui.button(btns, "取消", lambda: done.set("close"), "ghost").pack(side=tk.LEFT, padx=(0, 8))
    ui.button(btns, "🎞 產生簡報", lambda: done.set("go"), "primary").pack(side=tk.LEFT)
    root.wait_variable(done)
    alive["on"] = False
    state["show_agents"] = ag_var.get()
    state["back_cover"] = bc_var.get()
    r = done.get()
    try:
        stage.destroy()
    except tk.TclError:
        pass
    return r
