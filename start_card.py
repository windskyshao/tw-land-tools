"""開始卡：主程式一開先問「今天要做哪一件？」——接著做上次的案件／切到最近的案件／建立新案件。
勾「啟動時不用問我」就直接接著上次的案件(記在 ui_tips.json)，工具箱「說明」裡可以再叫出來。
本檔只負責畫卡片、列最近案件；真正切換案件、建立新案件由主程式傳函式進來。"""
import os
import json
import time
import tkinter as tk

import modern_ui as M
from case_flow import BigButton, Tip, ORANGE, ORANGE_DK, ORANGE_BG, GREEN, INK, INK2, FAM

OFF_KEY = "start_card_off"
_STATE = {"win": None, "queue": []}


# ───────────────────────── 設定(和萬用查詢面板共用 ui_tips.json) ─────────────────────────
def _tips_path(flag_dir):
    return os.path.join(flag_dir, "ui_tips.json")


def disabled(flag_dir):
    try:
        with open(_tips_path(flag_dir), "r", encoding="utf-8") as f:
            return bool(json.load(f).get(OFF_KEY))
    except Exception:
        return False


def set_disabled(flag_dir, off):
    p = _tips_path(flag_dir)
    try:
        with open(p, "r", encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        d = {}
    d[OFF_KEY] = bool(off)
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ───────────────────────── 開著時，其他彈窗(更新通知、功能提示)先排隊 ─────────────────────────
def is_open():
    w = _STATE["win"]
    try:
        return bool(w is not None and w.winfo_exists())
    except tk.TclError:
        return False


def when_closed(fn):
    """開始卡開著 → 等它關掉再跑；沒開 → 馬上跑"""
    if is_open():
        _STATE["queue"].append(fn)
    else:
        fn()


def _flush():
    q, _STATE["queue"] = _STATE["queue"], []
    for fn in q:
        try:
            fn()
        except Exception as e:
            print(f"[開始卡] 關閉後的工作失敗：{e}")


# ───────────────────────── 最近案件 ─────────────────────────
def _touched(folder):
    """案件最後動過的時間：資料夾本身、0~4 子資料夾、data.json、進度檔取最新"""
    ts = []
    for p in (folder, os.path.join(folder, "4.其他相關", "data.json"),
              os.path.join(folder, "4.其他相關", "case_stage.json")):
        try:
            ts.append(os.path.getmtime(p))
        except OSError:
            pass
    try:
        for e in os.scandir(folder):
            if e.is_dir() and e.name[:1].isdigit():
                ts.append(e.stat().st_mtime)
    except OSError:
        pass
    return max(ts) if ts else 0


def recent_cases(base_dir, n=5, exclude=None):
    """程式資料夾裡有 4.其他相關/data.json 的案件，最近動過的排前面"""
    out = []
    ex = os.path.normcase(os.path.abspath(exclude)) if exclude else None
    try:
        entries = list(os.scandir(base_dir))
    except OSError:
        return []
    for e in entries:
        if not e.is_dir():
            continue
        dj = os.path.join(e.path, "4.其他相關", "data.json")
        if not os.path.isfile(dj):
            continue
        if ex and os.path.normcase(os.path.abspath(e.path)) == ex:
            continue
        out.append(dict(folder=e.path, name=e.name, data_json=dj, mtime=_touched(e.path)))
    out.sort(key=lambda c: c["mtime"], reverse=True)
    return out[:n]


def when_text(ts):
    if not ts:
        return ""
    now = time.localtime()
    t = time.localtime(ts)
    days = (time.mktime(now[:3] + (0, 0, 0, 0, 0, -1)) - time.mktime(t[:3] + (0, 0, 0, 0, 0, -1))) // 86400
    if days <= 0:
        return time.strftime("今天 %H:%M", t)
    if days == 1:
        return time.strftime("昨天 %H:%M", t)
    if t.tm_year == now.tm_year:
        return f"{t.tm_mon}/{t.tm_mday}"
    return f"{t.tm_year - 1911}/{t.tm_mon}/{t.tm_mday}"


# ───────────────────────── 卡片 ─────────────────────────
BLUE = "#0A6CDB"            # 跟主程式上方「目前案件」同色
BLUE_BG = "#EAF2FD"
LINE = "#C9D1DC"
WEEK = "一二三四五六日"


def _dim(root):
    """主程式蓋一層半透明深色：看得到、點不到；跟著主程式移動"""
    ov = tk.Toplevel(root)
    ov.overrideredirect(True)
    ov.configure(bg="#1C1C1E", cursor="arrow")
    ov.transient(root)
    ov.attributes("-topmost", True)     # ★實測主程式還是會壓在它上面(transient 擋不住)→ 跟開始卡一樣設最上層，卡片關掉就一起消失
    try:
        ov.attributes("-alpha", 0.35)   # 淡一點：進度還看得清楚
    except tk.TclError:
        pass
    # 提示條另開一個不透明的小視窗(放在半透明層上字會變淡)
    bn = tk.Toplevel(ov)
    bn.overrideredirect(True)
    bn.transient(ov)
    bn.attributes("-topmost", True)
    bn.configure(bg="#FFF4E5", highlightthickness=2, highlightbackground="#E58135")
    ov.msg = tk.Label(bn, text="請先選擇要從哪裡開始", font=(FAM, 15, "bold"), bg="#FFF4E5", fg="#7C2D12",
                      justify="center", padx=M.px(14), pady=M.px(10))
    ov.msg.pack()
    tk.Label(bn, text="這裡是目前案件的進度（先看就好），選好之後就能操作", font=(FAM, 12), bg="#FFF4E5", fg="#1C1917",
             padx=M.px(14)).pack(pady=(0, M.px(10)))

    last = {}
    subs = [w_ for w_ in bn.winfo_children() if isinstance(w_, tk.Label)]

    def fit(e=None, raise_=False):
        # (主程式 Configure 觸發時不能 lift/update：會再觸發 Configure → 無限迴圈；只有計時呼叫(raise_)才做)
        # ★實測 lift 這種無邊框最上層視窗，Tk 會把它的位置重設到左上角 → lift 完一定要把位置再設一次
        try:
            if e is not None and e.widget is not root:
                return
            c = getattr(ov, "_card", None)
            if raise_:
                ov.lift()
                bn.lift()
                if c is not None:
                    c.attributes("-topmost", True)       # 開始卡放回最上面
                bn.update_idletasks()
            w, h, x, y = root.winfo_width(), root.winfo_height(), root.winfo_rootx(), root.winfo_rooty()
            bw = max(M.px(200), w - M.px(40))            # 提示條固定比主程式窄一點，字太長就換行(不會超出螢幕)
            g = (w, h, x, y, bn.winfo_reqheight())
            if raise_ or g != last.get("g"):
                last["g"] = g
                for l_ in subs:
                    l_.configure(wraplength=bw - M.px(30))
                ov.geometry(f"{w}x{h}+{x}+{y}")
                bh = bn.winfo_reqheight()
                bn.geometry(f"{bw}x{bh}+{x + (w - bw) // 2}+{y + h - bh - M.px(140)}")   # 放下面(蓋訊息區)，上面的進度看得到
                if raise_ and c is not None and getattr(c, "_pos", None):
                    c.geometry("%dx%d+%d+%d" % c._pos)
        except tk.TclError:
            pass
    ov.fit = fit
    fit(raise_=True)
    root.after(300, lambda: fit(raise_=True))     # 開機時主程式可能還沒畫好(寬高 1)：等一下再對一次
    root.after(1500, lambda: fit(raise_=True))
    bid = root.bind("<Configure>", fit, add="+")
    bn.bind("<Button-1>", lambda e: ov.event_generate("<Button-1>"))
    return ov, bid


def _badge(parent, n, bg):
    c = tk.Canvas(parent, width=M.px(30), height=M.px(30), bg=bg, highlightthickness=0)
    c.create_oval(1, 1, M.px(29), M.px(29), fill=BLUE, outline=BLUE)
    c.create_text(M.px(15), M.px(15), text=str(n), fill="white", font=(FAM, 13, "bold"))
    return c


def _section(parent, n, title, bg):
    row = tk.Frame(parent, bg=bg)
    row.pack(fill=tk.X, pady=(M.px(12), M.px(5)))
    _badge(row, n, bg).pack(side=tk.LEFT)
    tk.Label(row, text=title, font=(FAM, 16, "bold"), bg=bg, fg=INK).pack(side=tk.LEFT, padx=(M.px(8), 0))
    return row


def show(root, base_dir, current=None, progress_fn=None, on_continue=None, on_switch=None, on_new=None,
         on_browse=None, on_cancel=None, n_recent=4):
    """current={'name','folder'} 或 None；progress_fn(folder, data_json)→「做到第幾步」文字(可回空)；
    on_switch(data_json) 切到那個案件；on_new() 建立新案件；on_browse() 從資料夾挑別的案件；
    on_cancel() 按右上 ✕ 或 Esc＝先不選(留在上次的案件，跟「接著做」一樣不切換，只是不急著開始)。
    主程式蓋半透明(看得到、點不到)，卡片放在主程式旁邊不擋住它"""
    if is_open():
        _STATE["win"].lift()
        return _STATE["win"]
    bg = "#FFFFFF"
    ov, bid = _dim(root)
    tw = tk.Toplevel(root)
    ov._card = tw
    _STATE["win"] = tw
    tw.title("請問目前我們要從哪裡開始？")
    tw.overrideredirect(True)                 # 自己畫標題列(正式一點，也不會被當成一般小視窗關掉)
    tw.configure(bg=LINE)
    tw.attributes("-topmost", True)
    frame = tk.Frame(tw, bg=bg)
    frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
    pad = M.px(20)

    def close(then=None):
        try:
            tw.grab_release()
        except tk.TclError:
            pass
        try:
            # ★不能用 root.unbind("<Configure>", bid)：tkinter 會把主程式 Configure 的「所有」綁定一起清掉
            #   (主視窗拖走自動放回就是被它清掉的)→ 只拿掉自己這一條
            script = root.bind("<Configure>") or ""
            root.bind("<Configure>", "\n".join(l_ for l_ in script.split("\n") if bid not in l_))
            root.deletecommand(bid)
        except tk.TclError:
            pass
        for w in (tw, ov):
            try:
                w.destroy()
            except tk.TclError:
                pass
        _STATE["win"] = None
        try:
            root.focus_force()
        except tk.TclError:
            pass
        if then:
            try:
                then()
            except Exception as e:
                print(f"[開始卡] {e}")
        root.after(200, _flush)

    # 標題列：阿生＋問句＋今天日期(拖這裡可以移動)
    head = tk.Frame(frame, bg=BLUE)
    head.pack(fill=tk.X)
    try:
        from quick_case import _cat_photo
        tw._cat = _cat_photo(tw, M.px(56), ring=3)
    except Exception:
        tw._cat = None
    if tw._cat:
        tk.Label(head, image=tw._cat, bg=BLUE).pack(side=tk.LEFT, padx=(pad, M.px(10)), pady=M.px(14))
    ht = tk.Frame(head, bg=BLUE)
    ht.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0 if tw._cat else pad, pad), pady=M.px(14))
    t = time.localtime()
    tk.Label(ht, text="請問目前我們要從哪裡開始？", font=(FAM, 19, "bold"), bg=BLUE, fg="white", anchor="w").pack(fill=tk.X)
    tk.Label(ht, text=f"{t.tm_mon} 月 {t.tm_mday} 日（星期{WEEK[t.tm_wday]}）・請選一項", font=(FAM, 13),
             bg=BLUE, fg="#DCE9FB", anchor="w").pack(fill=tk.X)
    # 右上 ✕：先不選(2026-11-06 使用者：沒有取消)
    xb = tk.Label(head, text="✕", font=(FAM, 20, "bold"), bg=BLUE, fg="white", cursor="hand2", padx=M.px(12))
    xb.pack(side=tk.RIGHT, anchor="n", pady=M.px(8), padx=(0, M.px(6)), before=ht)
    xb.bind("<Enter>", lambda e: xb.configure(bg="#E5484D"))
    xb.bind("<Leave>", lambda e: xb.configure(bg=BLUE))
    xb.bind("<Button-1>", lambda e: close(on_cancel))
    Tip(xb, "先不選，關掉這張卡片（Esc 也可以）。\n" + ("會留在上次的案件：" + current["name"] if current else "")
        + "\n之後要換案件：按主程式「🔍 新案件」，或「🧰 其他功能」→ 開始卡。")

    drag = {}

    def press(e):
        drag.update(x=e.x_root - tw.winfo_x(), y=e.y_root - tw.winfo_y())

    def move(e):
        nx, ny = e.x_root - drag['x'], e.y_root - drag['y']
        tw.geometry(f"+{nx}+{ny}")
        if getattr(tw, "_pos", None):
            tw._pos = tw._pos[:2] + (nx, ny)
    for wdg in [head, ht] + [w_ for w_ in head.winfo_children() if w_ is not xb] + list(ht.winfo_children()):
        wdg.bind("<ButtonPress-1>", press)
        wdg.bind("<B1-Motion>", move)

    body = tk.Frame(frame, bg=bg)
    body.pack(fill=tk.BOTH, expand=True, padx=pad)

    # ① 接著做上次的案件
    n = 1
    if current:
        _section(body, n, "接著做上次的案件", bg)
        n += 1
        box = tk.Frame(body, bg=ORANGE_BG, highlightbackground=ORANGE, highlightthickness=2)
        box.pack(fill=tk.X)
        tk.Label(box, text=current["name"], font=(FAM, 19, "bold"), bg=ORANGE_BG, fg=INK, anchor="w").pack(
            fill=tk.X, padx=M.px(14), pady=(M.px(10), 0))
        sub = []
        if progress_fn:
            try:
                p = progress_fn(current["folder"], os.path.join(current["folder"], "4.其他相關", "data.json"))
                if p:
                    sub.append(p)
            except Exception:
                pass
        w = when_text(_touched(current["folder"])) if current.get("folder") else ""
        if w:
            sub.append(f"上次動：{w}")
        if sub:
            tk.Label(box, text="\n".join(sub), font=(FAM, 14), bg=ORANGE_BG, fg=INK2, anchor="w",
                     justify="left").pack(fill=tk.X, padx=M.px(14))
        BigButton(box, "▶ 接著做這件", lambda: close(on_continue), height=M.px(46), size=17,
                  color="#55B092", fg=INK).pack(fill=tk.X, padx=M.px(14), pady=(M.px(8), M.px(12)))

    # ② 建立新案件
    _section(body, n, "建立新案件", bg)
    n += 1
    tk.Label(body, text="打地址、門牌或「地段 地號」，查好就建立案件資料夾。", font=(FAM, 14), bg=bg, fg=INK2,
             anchor="w").pack(fill=tk.X)
    BigButton(body, "＋ 建立新案件", lambda: close(on_new), height=M.px(46), size=17,
              color="#E58135", fg=INK).pack(fill=tk.X, pady=(M.px(6), 0))

    # ③ 切到其他案件
    recent = recent_cases(base_dir, n_recent, exclude=current.get("folder") if current else None)
    if recent or on_browse:
        _section(body, n, "切到其他案件", bg)
    for c in recent:
        row = tk.Frame(body, bg=BLUE_BG, cursor="hand2", highlightbackground=LINE, highlightthickness=1)
        row.pack(fill=tk.X, pady=(0, M.px(5)))
        a = tk.Label(row, text=c["name"], font=(FAM, 15, "bold"), bg=BLUE_BG, fg=INK, anchor="w", cursor="hand2")
        a.pack(side=tk.LEFT, padx=(M.px(10), 0), pady=M.px(6))
        b = tk.Label(row, text=when_text(c["mtime"]) + "　›", font=(FAM, 13), bg=BLUE_BG, fg=INK2, cursor="hand2")
        b.pack(side=tk.RIGHT, padx=M.px(10))
        for wdg in (row, a, b):
            wdg.bind("<Button-1>", lambda e, c=c: close(lambda: on_switch and on_switch(c["data_json"])))
        if progress_fn:
            def tip(c=c):
                try:
                    return progress_fn(c["folder"], c["data_json"]) or "點一下切到這個案件"
                except Exception:
                    return "點一下切到這個案件"
            Tip(row, tip)
    if on_browse:
        lk = tk.Label(body, text="📂 其他案件（從資料夾挑）", font=(FAM, 14, "underline"), bg=bg, fg="#1F5FAF",
                      cursor="hand2", anchor="w")
        lk.pack(fill=tk.X, pady=(M.px(2), 0))
        lk.bind("<Button-1>", lambda e: close(on_browse))

    # 頁尾：不要再問
    foot = tk.Frame(frame, bg="#F4F6F9", highlightbackground=LINE, highlightthickness=1)
    foot.pack(fill=tk.X, pady=(M.px(16), 0))
    off = tk.BooleanVar(master=tw, value=disabled(base_dir))
    off.trace_add("write", lambda *a: set_disabled(base_dir, off.get()))
    fr = tk.Frame(foot, bg="#F4F6F9")
    fr.pack(fill=tk.X, padx=pad, pady=M.px(10))
    try:
        import deal_ui
        deal_ui.big_check(fr, off).pack(side=tk.LEFT, padx=(0, M.px(8)))
    except Exception:
        tk.Checkbutton(fr, variable=off, bg="#F4F6F9", activebackground="#F4F6F9").pack(side=tk.LEFT)
    cb = tk.Label(fr, text="啟動時不用問我（直接接著上次的案件）\n想再打開：🧰 其他功能 →「說明」→ 開始卡",
                  font=(FAM, 13), bg="#F4F6F9", fg=INK, cursor="hand2", justify="left", anchor="w")
    cb.pack(side=tk.LEFT)
    cb.bind("<Button-1>", lambda e: off.set(not off.get()))

    # 位置：主程式左邊(不擋住它)；左邊放不下才放右邊，再不行就蓋在上面
    tw.update_idletasks()
    W, H = max(tw.winfo_reqwidth(), M.px(500)), tw.winfo_reqheight()
    sw, sh = tw.winfo_screenwidth(), tw.winfo_screenheight()
    rx, ry, rw = root.winfo_rootx(), root.winfo_rooty(), root.winfo_width()
    gap = M.px(14)
    if rx - W - gap >= 0:
        x = rx - W - gap
        ov.msg.configure(text="← 請先在左邊選擇\n要從哪裡開始")
    elif rx + rw + gap + W <= sw:
        x = rx + rw + gap
        ov.msg.configure(text="請先在右邊選擇 →\n要從哪裡開始")
    else:
        x = max(0, rx + (rw - W) // 2)
    y = max(0, min(ry, sh - H - M.px(48)))
    tw.geometry(f"{W}x{H}+{x}+{y}")
    tw._pos = (W, H, x, y)
    tw.after(50, lambda: ov.fit(raise_=True))      # 提示字換了(←左邊/右邊→)：重新置中
    tw.bind("<Escape>", lambda e: close(on_cancel))   # Esc＝右上 ✕：先不選
    ov.bind("<Button-1>", lambda e: (tw.lift(), tw.focus_force()))
    tw.lift()
    tw.focus_force()
    try:
        tw.grab_set()
    except tk.TclError:
        pass
    return tw
