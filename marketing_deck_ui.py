"""行銷簡報：挑照片視窗(縮圖＋分類下拉)。pick() 回傳 (選用的照片清單, 輸出路徑) 或 None(取消)"""
import os
import sys
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)
import deal_ui as ui
from PIL import Image, ImageOps, ImageTk

INK = "#0F172A"
INK2 = "#1E293B"
USED = "#3E83B6"       # 有選用的框色(天空藍)
COVER = "#B88E3E"      # 封面(金)
UNUSED = "#CBD5E1"


def pick(photos, facts, default_out, page="a4l", on_save=None):
    from marketing_deck import CATS, PAGES, caption_of
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    ui.apply(root)
    root.title("行銷簡報（大樓版）")
    sc = max(1.0, root.winfo_fpixels("1i") / 96)
    TW, TH = int(210 * sc), int(158 * sc)
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    W, H = min(int(1180 * sc), sw - 60), min(int(900 * sc), sh - 90)
    root.geometry(f"{W}x{H}+{(sw - W) // 2}+{max(0, (sh - H) // 2 - 20)}")
    ui.header(root, "🎞 行銷簡報（大樓版）",
              f"{facts.get('short') or facts.get('title', '')}　替照片選分類、照片下方可改說明（會記住，下次自動帶回）")
    foot = ui.footer(root)

    # 上方：分類統計
    top = tk.Frame(root, bg=ui.BG, padx=16, pady=6)
    top.pack(fill=tk.X)
    stat = tk.Label(top, text="", font=("Microsoft JhengHei", 14, "bold"), fg=INK, bg=ui.BG, justify="left", anchor="w")
    stat.pack(side=tk.LEFT, fill=tk.X, expand=True)
    page_names = [v[0] for v in PAGES.values()]
    page_var = tk.StringVar(master=root, value=PAGES.get(page, PAGES["a4l"])[0])
    ttk.Combobox(top, textvariable=page_var, values=page_names, state="readonly", width=10,
                 font=("Microsoft JhengHei", 13)).pack(side=tk.RIGHT)
    tk.Label(top, text="版面：", font=("Microsoft JhengHei", 13, "bold"), fg=INK, bg=ui.BG).pack(side=tk.RIGHT, padx=(16, 0))
    tk.Label(top, text="⭐＝「廣告」資料夾（整理過的照片）", font=("Microsoft JhengHei", 13), fg=INK2, bg=ui.BG).pack(side=tk.RIGHT)

    body, canvas = ui.scroll_body(root)
    cols = max(3, (W - 60) // (TW + 24))
    tiles = []
    vars_ = []

    def refresh(*_):
        cnt = {}
        for p, v, fr in tiles:
            c = v.get()
            p["cat"] = c
            cnt[c] = cnt.get(c, 0) + 1
            fr.configure(highlightbackground=COVER if c == "封面" else (USED if c != "不使用" else UNUSED),
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
        tk.Label(fr, text=name, font=("Microsoft JhengHei", 12), fg=INK, bg="white", wraplength=TW, justify="left").pack(anchor="w", pady=(4, 2))
        v = tk.StringVar(master=root, value=p["cat"])
        cb = ttk.Combobox(fr, textvariable=v, values=CATS, state="readonly", font=("Microsoft JhengHei", 13), width=10)
        cb.pack(anchor="w")
        v.trace_add("write", refresh)
        # 說明(照片上的小字條)：預設用檔名描述，可以改；清空＝不顯示說明
        cap = tk.StringVar(master=root, value=p["cap"] if "cap" in p else caption_of(p))
        ce = tk.Entry(fr, textvariable=cap, font=("Microsoft JhengHei", 12), width=max(12, TW // 14), fg=INK,
                      relief="flat", highlightthickness=1, highlightbackground=UNUSED, highlightcolor=USED)
        ce.pack(anchor="w", fill=tk.X, pady=(4, 0), ipady=2)
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
        vars_.append(img_lbl)
    refresh()

    # 縮圖在背景載入(JPEG 用 draft 快速縮小)，載好再回主執行緒貼上
    q = queue.Queue()

    def loader():
        for idx, p in enumerate(photos):
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
        try:
            while True:
                idx, im = q.get_nowait()
                if im is not None:
                    ph = ImageTk.PhotoImage(im)
                    keep[idx] = ph
                    vars_[idx].configure(image=ph)
        except queue.Empty:
            pass
        if len(keep) < len(photos) and root.winfo_exists():
            root.after(60, pump)
    threading.Thread(target=loader, daemon=True).start()
    root.after(60, pump)

    # 頁尾：按鈕先放(靠右)，路徑放剩下的空間(太長中間用…)
    out_var = tk.StringVar(master=root, value=default_out)
    shown = tk.StringVar(master=root)

    def _short(p, n=60):
        return p if len(p) <= n else p[:28] + " … " + p[-(n - 31):]
    out_var.trace_add("write", lambda *a: shown.set(_short(out_var.get())))
    shown.set(_short(default_out))
    btns = tk.Frame(foot, bg=foot.cget("bg"))
    btns.pack(side=tk.RIGHT)
    left = tk.Frame(foot, bg=foot.cget("bg"))
    left.pack(side=tk.LEFT, fill=tk.X, expand=True)
    tk.Label(left, text="存到：", font=("Microsoft JhengHei", 13, "bold"), fg=INK, bg=foot.cget("bg")).pack(side=tk.LEFT)
    tk.Label(left, textvariable=shown, font=("Microsoft JhengHei", 12), fg=INK2, bg=foot.cget("bg"), anchor="w").pack(side=tk.LEFT)

    def change_out():
        p = filedialog.asksaveasfilename(parent=root, title="簡報存檔位置", defaultextension=".pptx",
                                         initialdir=os.path.dirname(out_var.get()), initialfile=os.path.basename(out_var.get()),
                                         filetypes=[("PowerPoint 簡報", "*.pptx")])
        if p:
            out_var.set(os.path.normpath(p))

    result = {}

    def cur_page():
        return next((k for k, v in PAGES.items() if v[0] == page_var.get()), "a4l")

    def remember():
        for p, v, _ in tiles:
            p["cat"] = v.get()
        if on_save:
            on_save(cur_page())

    def cancel():
        remember()
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", cancel)

    def go():
        chosen = [p for p, v, _ in tiles if v.get() != "不使用"]
        if not chosen and not messagebox.askyesno("沒有照片", "一張照片都沒選，仍要產生（只有文字頁）？", parent=root):
            return
        remember()
        result["v"] = (chosen, out_var.get(), cur_page())
        root.destroy()

    ui.button(btns, "🎞 產生簡報", go, "primary").pack(side=tk.RIGHT, padx=(8, 0))
    ui.button(btns, "取消", cancel, "ghost").pack(side=tk.RIGHT, padx=(8, 0))
    ui.button(btns, "變更位置…", change_out, "ghost").pack(side=tk.RIGHT)
    root.mainloop()
    return result.get("v")
