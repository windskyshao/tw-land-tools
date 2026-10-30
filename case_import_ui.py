"""整理既有資料夾的預覽視窗：列出每個檔案判斷要放哪、為什麼、有多確定；使用者可以改，確認後才複製(原檔不動)"""
import os
import re
import shutil
import hashlib
import tkinter as tk
from tkinter import ttk

import modern_ui as M
import case_flow as CF

FAM = CF.FAM
CONF_COL = {"高": "#23864B", "中": "#C98A00", "低": "#C5221F"}
SHOW = {"根目錄": "案件資料夾（物調表、簡報）"}
LABELS = [SHOW.get(t, t) for t in ("0.謄本", "1.基本資料", "2.照片", "3.行情", "4.其他相關", "5.委託", "6.成交", "根目錄")]
TKEY = dict(zip(LABELS, ("0.謄本", "1.基本資料", "2.照片", "3.行情", "4.其他相關", "5.委託", "6.成交", "根目錄")))


def _md5(p):
    h = hashlib.md5()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def find_dups(files):
    """內容完全一樣的檔案：只留第一個，其他標 dup"""
    by_size = {}
    for r in files:
        try:
            by_size.setdefault(os.path.getsize(r["src"]), []).append(r)
        except OSError:
            pass
    for grp in by_size.values():
        if len(grp) < 2:
            continue
        seen = {}
        for r in grp:
            h = _md5(r["src"])
            if h in seen:
                r["dup"] = seen[h]["rel"]
            else:
                seen[h] = r


def copy_into(files, case_dir, move=False):
    """照判斷結果複製到案件資料夾；同名不覆蓋(加序號)。回傳 (成功數, 錯誤清單)"""
    done, errs = 0, []
    for r in files:
        if not r.get("use", True) or r.get("dup"):
            continue
        sub = "" if r["target"] == "根目錄" else r["target"]
        d = os.path.join(case_dir, sub)
        os.makedirs(d, exist_ok=True)
        base, ext = os.path.splitext(r.get("newname") or os.path.basename(r["src"]))
        dest = os.path.join(d, base + ext)
        n = 1
        while os.path.exists(dest):
            dest = os.path.join(d, f"{base} ({n}){ext}")
            n += 1
        try:
            (shutil.move if move else shutil.copy2)(r["src"], dest)
            done += 1
        except Exception as e:
            errs.append(f"{r['rel']}：{e}")
    return done, errs


def open_preview(parent, src, result, case_name, on_go=None, on_preview=None):
    """預覽視窗。on_go(files, move) 由主程式接著複製並刷新進度表"""
    files = result["files"]
    tw = tk.Toplevel(parent)
    tw.title("整理資料夾")
    tw.configure(bg="#FFFFFF")
    W = "#FFFFFF"
    head = tk.Frame(tw, bg=W)
    head.pack(fill=tk.X, padx=M.px(14), pady=(M.px(12), M.px(4)))
    tk.Label(head, text="📥 整理資料夾", font=(FAM, 15, "bold"), bg=W, fg=CF.INK).pack(anchor="w")
    ndirs = len({os.path.dirname(r["rel"]) for r in files if os.path.dirname(r["rel"])})
    tk.Label(head, text=f"來源：{src}（{len(files)} 個檔案" + (f"，{ndirs} 個子資料夾）" if ndirs else "）"),
             font=(FAM, 11), bg=W, fg=CF.INK2, anchor="w").pack(anchor="w")
    lots = sorted(result.get("lots", {}).items(), key=lambda x: -x[1])
    who = tk.Frame(head, bg=CF.ORANGE_BG)
    who.pack(fill=tk.X, pady=(M.px(6), 0))
    msg = f"放進案件：{case_name}"
    if lots:
        msg += f"　（從謄本認出：{'、'.join(l for l, _ in lots[:3])} 地號）"
    tk.Label(who, text=msg, font=(FAM, 12, "bold"), bg=CF.ORANGE_BG, fg=CF.INK, anchor="w", padx=M.px(8),
             pady=M.px(4)).pack(side=tk.LEFT)
    CF.Pill(who, "改", None).pack(side=tk.RIGHT, padx=M.px(6))

    # 各資料夾數量
    chips = tk.Frame(tw, bg=W)
    chips.pack(fill=tk.X, padx=M.px(14), pady=M.px(6))
    cnt = {}
    for r in files:
        if not r.get("dup"):
            cnt[r["target"]] = cnt.get(r["target"], 0) + 1
    for t in ("0.謄本", "1.基本資料", "2.照片", "3.行情", "5.委託", "6.成交", "根目錄", "4.其他相關"):
        if cnt.get(t):
            tk.Label(chips, text=f"{SHOW.get(t, t).split('（')[0]} {cnt[t]}", font=(FAM, 11, "bold"), bg="#F2EFEA",
                     fg=CF.INK, padx=M.px(8), pady=M.px(2)).pack(side=tk.LEFT, padx=(0, M.px(4)))
    low = sum(r["conf"] == "低" and not r.get("dup") for r in files)
    if low:
        tk.Label(chips, text=f"● 要你看一下 {low}", font=(FAM, 11, "bold"), bg="#FDECEA", fg=CF.RED, padx=M.px(8),
                 pady=M.px(2)).pack(side=tk.LEFT)

    # 清單：預設只列「要你看一下」的(紅/黃)，有把握的收起來；按「顯示全部」才攤開
    outer = tk.Frame(tw, bg=W)
    outer.pack(fill=tk.BOTH, expand=True, padx=M.px(14))
    cv = tk.Canvas(outer, bg=W, highlightthickness=0, height=M.px(150))
    sb = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
    body = tk.Frame(cv, bg=W)
    body.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
    win_id = cv.create_window(0, 0, window=body, anchor="nw")
    cv.bind("<Configure>", lambda e: cv.itemconfigure(win_id, width=e.width))      # 清單撐滿寬度
    cv.configure(yscrollcommand=sb.set)
    cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    order = {"低": 0, "中": 1, "高": 2}
    for r in files:
        r["use_var"] = tk.BooleanVar(value=not r.get("dup"))
        r["target_var"] = tk.StringVar(value=SHOW.get(r["target"], r["target"]))
    state = {"all": False}

    def render():
        for w in body.winfo_children():
            w.destroy()
        rows = sorted(files, key=lambda r: (bool(r.get("dup")), order[r["conf"]], r["target"], r["rel"]))
        if not state["all"]:
            rows = [r for r in rows if r["conf"] != "高" and not r.get("dup")]
        if not rows:
            tk.Label(body, text="✓ 每個檔案都認得出來，可以直接開始整理", font=(FAM, 14, "bold"), bg=W,
                     fg=CF.GREEN).pack(anchor="w", pady=M.px(10))
        for r in rows:
            row = tk.Frame(body, bg=W)
            row.pack(fill=tk.X, pady=M.px(2))
            tk.Checkbutton(row, variable=r["use_var"], bg=W, activebackground=W).pack(side=tk.LEFT)
            tk.Label(row, text="●", font=(FAM, 12), bg=W,
                     fg=CF.DIM if r.get("dup") else CONF_COL[r["conf"]]).pack(side=tk.LEFT)
            name = tk.Label(row, text=os.path.basename(r["rel"]), font=(FAM, 13, "bold"), bg=W,
                            fg=CF.DIM if r.get("dup") else CF.INK, anchor="w")
            name.pack(side=tk.LEFT, padx=(M.px(4), M.px(8)))
            CF.Tip(name, (f"位置：{r['rel']}\n" + (f"跟「{r['dup']}」一模一樣，只放一份" if r.get("dup")
                                                   else f"判斷：{r['kind']}（{r['reason']}）")))
            CF.Pill(row, "預覽", (lambda p=r["src"]: on_preview and on_preview(p))).pack(side=tk.RIGHT, padx=(M.px(4), 0))
            ttk.Combobox(row, textvariable=r["target_var"], values=LABELS, state="readonly", width=12,
                         font=(FAM, 12)).pack(side=tk.RIGHT)
        tog.configure(text="只看要確認的" if state["all"] else f"顯示全部 {len(files)} 個檔案")

    tog = tk.Label(tw, text="", font=(FAM, 11, "underline"), fg="#1F5FAF", bg=W, cursor="hand2")
    tog.pack(anchor="w", padx=M.px(14), pady=(M.px(4), 0))
    tog.bind("<Button-1>", lambda e: (state.update(all=not state["all"]), render()))
    render()

    move_var = tk.BooleanVar(value=False)
    mv = tk.Checkbutton(tw, text="整理完把原來的檔案搬走（不勾＝複製，原資料夾不動）", variable=move_var, font=(FAM, 11),
                        bg=W, activebackground=W)
    btns = tk.Frame(tw, bg=W)
    btns.pack(side=tk.BOTTOM, fill=tk.X, padx=M.px(14), pady=M.px(12))
    mv.pack(side=tk.BOTTOM, anchor="w", padx=M.px(14), pady=(M.px(6), 0))
    tog.pack_forget()
    tog.pack(side=tk.BOTTOM, anchor="w", padx=M.px(14), pady=(M.px(4), 0))
    n_go = sum(not r.get("dup") for r in files)

    def go():
        for r in files:
            r["use"] = r["use_var"].get()
            r["target"] = TKEY.get(r["target_var"].get(), r["target"])
        tw.destroy()
        if on_go:
            on_go(files, move_var.get())

    b = CF.Pill(btns, f"▶ 開始整理 {n_go} 個檔案", go, primary=True)
    b.primary = True
    b._draw()
    b.pack(side=tk.RIGHT)
    CF.Pill(btns, "取消", tw.destroy).pack(side=tk.RIGHT, padx=M.px(6))
    return tw
