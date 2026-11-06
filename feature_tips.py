"""功能提示(像阿生地圖的💡)：亮框指著按鈕＋大字說明，一次一個。
更新後只介紹「這一版新增的」；看過的不再跳。勾「不要再提示」就不自動跳，工具箱「說明 › 💡 功能提示」可以重看全部。
新功能要介紹 → 在 TIPS 加一筆(ver 填那一版的版號，target 填主程式 resolve 認得的名字；None＝開始卡選完就跳)。說明卡放在主程式旁邊，不擋住它。"""
import os
import re
import json
import tkinter as tk

import modern_ui as M
from case_flow import Tip, ORANGE, ORANGE_DK, INK, INK2, FAM

# target=None：啟動時開始卡選完就跳；target 有值：那個按鈕／區塊「第一次出現在畫面上」才跳(用到才講)
# after：這幾則有任一則看過了才跳(例如「提示可以重看」要先看過別的提示才有意義)
TIPS = [
    dict(id="start_card", ver="1.1.9n", target=None, title="📋 開始卡：每次打開先選要做哪一件",
         text="剛剛那張「請問目前我們要從哪裡開始？」每次打開都會出現。\n"
              "不想每次問：卡片最下面勾「啟動時不用問我」。\n"
              "想再叫出來：🧰 其他功能 →「說明」→ 開始卡。"),
    dict(id="stage_ask", ver="1.1.9n", target="stage_ask", title="新案件先問：委託還是開發？",
         text="每個案件第一次打開，會先問「已委託」還是「開發中」，只問一次。\n"
              "已委託＝屋主已經簽委託，資料全部備齊交公司。\n"
              "開發中＝還在爭取屋主簽，先免費查，最後做行情資料給屋主看。\n"
              "談成委託後，按最上面「已委託」切過去就好。"),
    dict(id="doc_items", ver="1.1.9n", target="doc_items", title="第 8 步：一項一格",
         text="物調表、簡報、業績分配表一項一格：\n綠底「✓」＝做好了，白底橘框「□」＝還沒做，\n灰框＝成交時再做（不影響打勾）。點格子就直接去做那一項。"),
    dict(id="toolbox_help", ver="1.1.9n", target="toolbox", after=("start_card", "stage_ask", "doc_items"),
         title="💡 提示看過可以重看",
         text="這些提示看過就不會再跳。\n想再看一次：🧰 其他功能 →「說明」→ 💡 功能提示。"),
]

SEEN_KEY, OFF_KEY = "ft_seen", "ft_off"
_STATE = {"wins": [], "running": False, "queue": []}


def _path(flag_dir):
    return os.path.join(flag_dir, "ui_tips.json")


def _load(flag_dir):
    try:
        with open(_path(flag_dir), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save(flag_dir, **kw):
    d = _load(flag_dir)
    d.update(kw)
    try:
        with open(_path(flag_dir), "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _mark_seen(flag_dir, ids):
    d = _load(flag_dir)
    seen = set(d.get(SEEN_KEY, [])) | set(ids)
    _save(flag_dir, **{SEEN_KEY: sorted(seen)})


def _vt(v):
    """'1.1.9m' → (1, 1, 9, 'm')：數字照數字比、尾巴字母照字母比"""
    m = re.match(r"([\d.]+)([a-z]*)", str(v or ""))
    if not m:
        return (0,)
    return tuple(int(x) for x in m.group(1).split(".") if x) + (m.group(2),)


def pending(flag_dir, version):
    """還沒看過的提示；第一次用(沒有紀錄)只介紹這一版新增的，舊版的當作看過"""
    d = _load(flag_dir)
    if d.get(OFF_KEY):
        return []
    if SEEN_KEY not in d:
        _mark_seen(flag_dir, [t["id"] for t in TIPS if _vt(t["ver"]) < _vt(version)])
        d = _load(flag_dir)
    seen = set(d.get(SEEN_KEY, []))
    return [t for t in TIPS if t["id"] not in seen]


def is_open():
    return _STATE["running"]


def when_closed(fn):
    """提示還在跳 → 等看完再跑(例如更新通知)；沒在跳 → 馬上跑"""
    if _STATE["running"]:
        _STATE["queue"].append(fn)
    else:
        fn()


def _finish():
    _clear()
    _STATE["running"] = False
    q, _STATE["queue"] = _STATE["queue"], []
    for fn in q:
        try:
            fn()
        except Exception as e:
            print(f"[功能提示] 看完後的工作失敗：{e}")


# ───────────────────────── 畫提示 ─────────────────────────
def _clear():
    for w in _STATE["wins"]:
        try:
            w.destroy()
        except tk.TclError:
            pass
    _STATE["wins"] = []


def _visible(w):
    try:
        return bool(w is not None and w.winfo_exists() and w.winfo_ismapped() and w.winfo_width() > 1)
    except tk.TclError:
        return False


def _frame_around(root, w):
    """目標按鈕外面畫一圈橘色亮框(四條無邊框小窗)"""
    x, y, ww, hh = w.winfo_rootx(), w.winfo_rooty(), w.winfo_width(), w.winfo_height()
    t, g = M.px(4), M.px(4)
    for gx, gy, gw, gh in ((x - g - t, y - g - t, ww + 2 * (g + t), t), (x - g - t, y + hh + g, ww + 2 * (g + t), t),
                           (x - g - t, y - g, t, hh + 2 * g), (x + ww + g, y - g, t, hh + 2 * g)):
        s = tk.Toplevel(root)
        s.overrideredirect(True)
        s.attributes("-topmost", True)
        s.configure(bg=ORANGE)
        s.geometry(f"{max(1, gw)}x{max(1, gh)}+{gx}+{gy}")
        _STATE["wins"].append(s)
    return x, y, ww, hh


def _pointer(root, x, y, left):
    """卡片和亮框之間的橘色小三角(指向主程式那一邊)"""
    key = "#FF00FE"
    w, h = M.px(18), M.px(30)
    p = tk.Toplevel(root)
    p.overrideredirect(True)
    p.attributes("-topmost", True)
    p.configure(bg=key)
    try:
        p.attributes("-transparentcolor", key)
    except tk.TclError:
        pass
    c = tk.Canvas(p, width=w, height=h, bg=key, highlightthickness=0)
    c.pack()
    c.create_polygon(*((0, 0, w, h // 2, 0, h) if left else (w, 0, 0, h // 2, w, h)), fill=ORANGE, outline=ORANGE)
    p.geometry(f"{w}x{h}+{x if left else x - w}+{y - h // 2}")
    _STATE["wins"].append(p)


def _card(root, tip, idx, total, on_next, on_prev, on_off, target):
    """說明卡放在主程式旁邊(不擋住它)，小三角指著亮框"""
    bg = "#FFFDF7"
    cw = tk.Toplevel(root)
    cw.overrideredirect(True)
    cw.attributes("-topmost", True)
    _STATE["wins"].append(cw)
    box = tk.Frame(cw, bg=bg, highlightbackground=ORANGE, highlightthickness=3)
    box.pack(fill=tk.BOTH, expand=True)
    hd = tk.Frame(box, bg=bg)
    hd.pack(fill=tk.X, padx=M.px(14), pady=(M.px(10), 0))
    tk.Label(hd, text="💡 功能提示", font=(FAM, 12, "bold"), bg=bg, fg=INK2, anchor="w").pack(side=tk.LEFT)
    if total > 1:
        tk.Label(hd, text=f"{idx + 1}/{total}", font=(FAM, 13), bg=bg, fg=INK2).pack(side=tk.RIGHT)
    W = M.px(440)
    tk.Label(box, text=tip["title"], font=(FAM, 16, "bold"), bg=bg, fg=ORANGE_DK, anchor="w", justify="left",
             wraplength=W - M.px(30)).pack(fill=tk.X, padx=M.px(14), pady=(M.px(2), 0))
    tk.Label(box, text=tip["text"], font=(FAM, 14), bg=bg, fg=INK, anchor="w", justify="left",
             wraplength=W - M.px(30)).pack(fill=tk.X, padx=M.px(14), pady=(M.px(6), M.px(8)))
    row = tk.Frame(box, bg=bg)
    row.pack(fill=tk.X, padx=M.px(14), pady=(0, M.px(12)))
    nxt = tk.Label(row, text="下一個 ▶" if idx + 1 < total else "知道了", font=(FAM, 14, "bold"), bg="#55B092", fg=INK,
                   padx=M.px(16), pady=M.px(4), cursor="hand2")
    nxt.pack(side=tk.RIGHT)
    nxt.bind("<Button-1>", lambda e: on_next())
    if idx > 0:
        pv = tk.Label(row, text="◀ 上一個", font=(FAM, 14, "bold"), bg="#E6E3DD", fg=INK, padx=M.px(14),
                      pady=M.px(4), cursor="hand2")
        pv.pack(side=tk.RIGHT, padx=(0, M.px(8)))
        pv.bind("<Button-1>", lambda e: on_prev())
    if on_off:
        off = tk.Label(row, text="不要再提示", font=(FAM, 13, "underline"), bg=bg, fg="#1F5FAF", cursor="hand2")
        off.pack(side=tk.LEFT)
        off.bind("<Button-1>", lambda e: on_off())
        Tip(off, "之後有新功能也不自動跳。\n想重看：🧰 其他功能 →「說明」→ 💡 功能提示。")
    cw.update_idletasks()
    H = cw.winfo_reqheight()
    sw, sh = cw.winfo_screenwidth(), cw.winfo_screenheight()
    rx, ry, rw = root.winfo_rootx(), root.winfo_rooty(), root.winfo_width()
    gap = M.px(22)
    ty = (target[1] + target[3] // 2) if target else ry + M.px(200)      # 要指的高度
    if rx - W - gap >= 0:                    # 主程式左邊
        cx, side = rx - W - gap, "left"
    elif rx + rw + gap + W <= sw:            # 右邊
        cx, side = rx + rw + gap, "right"
    else:                                    # 兩邊都放不下才蓋在主程式上(放亮框下面)
        cx, side = rx + (rw - W) // 2, None
        if target:
            ty = target[1] + target[3] + M.px(14) + H // 2
    cy = max(M.px(4), min(ty - H // 2, sh - H - M.px(48)))
    cw.geometry(f"{W}x{H}+{cx}+{cy}")
    if side and target:
        py = max(cy + M.px(20), min(ty, cy + H - M.px(20)))
        _pointer(root, cx + W if side == "left" else cx, py, side == "left")


def run(root, flag_dir, tips, resolve=None, mark=True, on_done=None):
    """一個一個顯示(可以回上一個)；resolve(target 名)→ widget(找不到或沒顯示就不畫框)"""
    _clear()
    tips = list(tips)
    _STATE["running"] = True

    def step(i):
        _clear()
        if i >= len(tips):
            _finish()
            if on_done:
                on_done()
            return
        t = tips[i]
        w = None
        if t.get("target") and resolve:
            try:
                w = resolve(t["target"])
            except Exception:
                w = None
        box = _frame_around(root, w) if _visible(w) else None

        def nxt():
            if mark:
                _mark_seen(flag_dir, [t["id"]])
            step(i + 1)

        def off():
            _save(flag_dir, **{OFF_KEY: True})
            if mark:
                _mark_seen(flag_dir, [x["id"] for x in tips])
            _finish()
        _card(root, t, i, len(tips), nxt, lambda: step(i - 1), off if mark else None, box)

    step(0)


def show_ready(root, flag_dir, version, resolve=None, startup=False):
    """用到才講：沒看過、而且「現在畫面上有那個東西」的提示才跳。
    startup=True(開始卡剛選完)才跳不指按鈕的那幾則。回傳有沒有跳"""
    if _STATE["running"]:
        return False
    p = pending(flag_dir, version)
    if not p:
        return False
    seen = set(_load(flag_dir).get(SEEN_KEY, []))
    ready = []
    for t in p:
        if t.get("after") and not (seen & set(t["after"])):
            continue
        if t.get("target") is None:
            if startup:
                ready.append(t)
            continue
        try:
            w = resolve(t["target"]) if resolve else None
        except Exception:
            w = None
        if _visible(w):
            ready.append(t)
    if ready:
        run(root, flag_dir, ready, resolve)
    return bool(ready)


def show_new(root, flag_dir, version, resolve=None):
    """(舊名，保留相容)啟動時"""
    return show_ready(root, flag_dir, version, resolve, startup=True)


def show_all(root, flag_dir, resolve=None):
    """工具箱「💡 功能提示」：列出全部，可以一個一個重看，也可以打開/關掉自動提示"""
    bg = "#FFFFFF"
    tw = tk.Toplevel(root)
    tw.title("💡 功能提示")
    tw.configure(bg=bg)
    tw.transient(root)
    tk.Label(tw, text="💡 功能提示", font=(FAM, 18, "bold"), bg=bg, fg=INK, anchor="w").pack(
        fill=tk.X, padx=M.px(16), pady=(M.px(14), M.px(6)))
    tk.Label(tw, text="點一項再看一次說明：", font=(FAM, 13), bg=bg, fg=INK2, anchor="w").pack(fill=tk.X, padx=M.px(16))
    for t in reversed(TIPS):                         # 新的在上面
        r = tk.Label(tw, text=f"{t['title']}　（v{t['ver']}）", font=(FAM, 14), bg="#F6F4F0", fg=INK, anchor="w",
                     padx=M.px(10), pady=M.px(6), cursor="hand2")
        r.pack(fill=tk.X, padx=M.px(16), pady=(M.px(4), 0))
        r.bind("<Button-1>", lambda e, t=t: (tw.destroy(), run(root, flag_dir, [t], resolve, mark=False)))
    allb = tk.Label(tw, text="▶ 從頭全部看一遍", font=(FAM, 14, "bold"), bg="#55B092", fg=INK, padx=M.px(14),
                    pady=M.px(5), cursor="hand2")
    allb.pack(anchor="w", padx=M.px(16), pady=(M.px(12), 0))
    allb.bind("<Button-1>", lambda e: (tw.destroy(), run(root, flag_dir, TIPS, resolve, mark=False)))
    on = tk.BooleanVar(master=tw, value=not _load(flag_dir).get(OFF_KEY))
    tk.Checkbutton(tw, text="有新功能時自動提示我", variable=on, font=(FAM, 13), bg=bg, fg=INK, activebackground=bg,
                   command=lambda: _save(flag_dir, **{OFF_KEY: not on.get()})).pack(
        anchor="w", padx=M.px(16), pady=(M.px(12), M.px(14)))
    tw.update_idletasks()
    W, H = max(tw.winfo_reqwidth(), M.px(420)), tw.winfo_reqheight()
    tw.geometry(f"{W}x{H}+{root.winfo_rootx() + max(0, (root.winfo_width() - W) // 2)}+{root.winfo_rooty() + M.px(80)}")
    return tw
