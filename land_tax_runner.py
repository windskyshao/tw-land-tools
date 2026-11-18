"""謄本結構化程式(電子／電傳)共用：轉好的謄本資料 → 算土地增值稅

2026-11-17 同事只想「轉謄本順便算土增稅」，不想先回主程式建案件。
- 稅法只放在 finance_tax.py 一個地方：這裡只負責找到它、用 --calc(程式照稅法算，不開財政部網站)跑它、把結果用大字顯示
- 試算表 PDF／Excel 存在 <主程式資料夾>\\<行政區段名-地號>\\1.基本資料(跟主程式按「財政部土增稅」同一個地方)
- 結構化 exe 跟主程式 exe 放在同一個資料夾：用主程式的 python_embedded 跑 _internal\\finance_tax.py
"""
import glob
import json
import os
import re
import subprocess
import sys
import threading
import time

PREF_FILE = "transcript_tax_pref.json"
_ANSI = re.compile(r"\x1b\[[0-9;]*m")
FAM = "微軟正黑體"


# ───────────────────────── 不碰畫面的部分(測試用得到) ─────────────────────────
def load_pref(base_dir):
    """「轉好後順便算土增稅」上次有沒有勾"""
    try:
        with open(os.path.join(base_dir, PREF_FILE), encoding="utf-8") as f:
            return bool(json.load(f).get("with_tax"))
    except Exception:
        return False


def save_pref(base_dir, on):
    try:
        with open(os.path.join(base_dir, PREF_FILE), "w", encoding="utf-8") as f:
            json.dump({"with_tax": bool(on)}, f)
    except Exception:
        pass


def is_transcript_json(path):
    """謄本結構化的資料：清單、裡面是一份份謄本(有「謄本類型」或「基本資訊」)，而且至少一份是土地"""
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return False
    if not isinstance(data, list):
        return False
    recs = [d for d in data if isinstance(d, dict) and ("謄本類型" in d or "基本資訊" in d)]
    return any("土地" in str(d.get("謄本類型", "")) for d in recs)


def newest_transcript_json(dirs, since):
    """這次處理產生的謄本 JSON：幾個資料夾裡、since(時間戳)之後寫的、最新的一份"""
    cands = []
    for d in dirs:
        if not d or not os.path.isdir(d):
            continue
        for p in glob.glob(os.path.join(d, "*.json")):
            try:
                mt = os.path.getmtime(p)
            except OSError:
                continue
            if mt >= since:
                cands.append((mt, p))
    for _, p in sorted(cands, reverse=True):
        if is_transcript_json(p):
            return p
    return None


def finance_cmd(json_path, base_dir, frozen=None, exe_dir=None, here=None):
    """組出跑 finance_tax.py --calc 的指令。回傳 (cmd, env, cwd, 錯誤說明)；找不到就 cmd=None"""
    frozen = getattr(sys, "frozen", False) if frozen is None else frozen
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    if frozen:
        app = exe_dir or os.path.dirname(sys.executable)
        py = os.path.join(app, "python_embedded", "python.exe")
        internal = os.path.join(app, "_internal")
        script = os.path.join(internal, "finance_tax.py")
        if not (os.path.isfile(py) and os.path.isfile(script)):
            return None, None, None, ("找不到土增稅試算程式。\n\n這個功能要跟「地籍資料查詢系統」放在同一個資料夾才能用"
                                      "（它會用主程式裡的試算程式）。")
        env["PYTHONHOME"] = os.path.join(app, "python_embedded")
        env["PYTHONPATH"] = internal + os.pathsep + os.path.join(app, "python_embedded", "Lib", "site-packages")
        env["PYTHONNOUSERSITE"] = "1"
        cwd = internal
    else:
        root = here or os.path.dirname(os.path.abspath(__file__))
        py, script, cwd = sys.executable, os.path.join(root, "finance_tax.py"), root
        if not os.path.isfile(script):
            return None, None, None, f"找不到 {script}"
    return [py, script, "--json", json_path, "--calc", "--base-dir", base_dir], env, cwd, ""


def parse_output(lines):
    """從試算程式印出來的字抓重點：合計、試算表 PDF 路徑、成功沒有"""
    out = {"total": "", "pdf": "", "ok": False, "zero": False, "rows": []}
    for raw in lines:
        s = _ANSI.sub("", raw).strip()
        if "合計（" in s:
            out["total"] = s
        elif s.startswith("✓ 試算表："):
            out["pdf"] = s.split("：", 1)[1].strip()
        elif re.match(r"^\d+\.\s", s) and "元" in s or "✗" in s and "沒算進合計" in s:
            out["rows"].append(s)
        if "土地增值稅程式試算完成" in s:
            out["ok"] = True
        if "稅額為 0" in s:
            out["zero"] = True
    return out


def clean(line):
    s = _ANSI.sub("", line).rstrip()
    if s.startswith("__PROGRESS__") or s.startswith("[DEBUG]"):
        return None
    return s


# ───────────────────────── 視窗 ─────────────────────────
def run_tax_window(master, json_path, base_dir, wait=True):
    """開一個大字視窗：跑試算 → 顯示每位所有權人的稅額和合計，可以打開試算表 PDF／資料夾"""
    import tkinter as tk
    from tkinter import scrolledtext

    cmd, env, cwd, err = finance_cmd(json_path, base_dir)
    win = tk.Toplevel(master)
    win.title("土地增值稅試算（財政部稅率，程式直接算）")
    win.configure(bg="#FFFFFF")
    try:
        win.transient(master)
    except Exception:
        pass
    k = max(1.0, win.winfo_fpixels("1i") / 96)
    W, H = int(760 * k), int(600 * k)
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    W, H = min(W, sw - 40), min(H, sh - 80)
    win.geometry(f"{W}x{H}+{(sw - W) // 2}+{max(0, (sh - H) // 3)}")

    head = tk.Label(win, text="💰 土地增值稅試算中…", font=(FAM, 18, "bold"), fg="#7C2D12", bg="#FFFFFF", anchor="w")
    head.pack(fill="x", padx=20, pady=(16, 4))
    sub = tk.Label(win, text="用財政部稅率，照謄本上的前次移轉現值、公告現值、持分和物價指數算（要連網抓物價指數）",
                   font=(FAM, 13), fg="#1C1917", bg="#FFFFFF", anchor="w", justify="left", wraplength=W - 40)
    sub.pack(fill="x", padx=20)
    total_lbl = tk.Label(win, text="", font=(FAM, 17, "bold"), fg="#1B5E20", bg="#FFFFFF", anchor="w", justify="left",
                         wraplength=W - 40)
    total_lbl.pack(fill="x", padx=20, pady=(8, 4))
    txt = scrolledtext.ScrolledText(win, font=(FAM, 13), fg="#111111", bg="#FAFAF7", wrap="word", height=12)
    txt.pack(fill="both", expand=True, padx=20, pady=(4, 8))
    bf = tk.Frame(win, bg="#FFFFFF")
    bf.pack(fill="x", padx=20, pady=(0, 16))
    state = {"pdf": "", "lines": [], "proc": None}

    def btn(text, cmd_, bg="#E6E3DD"):
        b = tk.Button(bf, text=text, font=(FAM, 14, "bold"), bg=bg, fg="#111111", relief="flat", padx=16, pady=8,
                      cursor="hand2", command=cmd_)
        return b

    def open_pdf():
        if state["pdf"] and os.path.exists(state["pdf"]):
            os.startfile(state["pdf"])

    def open_dir():
        p = state["pdf"]
        d = os.path.dirname(p) if p else ""
        if d and os.path.isdir(d):
            os.startfile(d)

    def close():
        pr = state["proc"]
        if pr is not None and pr.poll() is None:
            try:
                pr.kill()
            except Exception:
                pass
        win.destroy()

    close_b = btn("關閉", close)
    close_b.pack(side="right")
    pdf_b = btn("📄 打開試算表", open_pdf, "#55B092")
    dir_b = btn("📂 打開資料夾", open_dir)
    win.protocol("WM_DELETE_WINDOW", close)
    win.bind("<Escape>", lambda e: close())

    def add(line):
        try:
            txt.insert("end", line + "\n")
            txt.see("end")
        except tk.TclError:
            pass

    def finished(rc):
        try:
            if not win.winfo_exists():
                return
        except tk.TclError:
            return
        r = parse_output(state["lines"])
        state["pdf"] = r["pdf"]
        if r["ok"] and r["total"]:
            head.config(text="✅ 土地增值稅算好了", fg="#1B5E20")
            t = r["total"].replace("合計", "合計 ", 1)
            if r["zero"]:
                t += "\n（稅額 0：公告現值沒有超過物價調整後的前次移轉現值）"
            total_lbl.config(text=t)
            where = os.path.dirname(r["pdf"]) if r["pdf"] else ""
            sub.config(text=f"試算表存在：{where}" if where else "試算完成")
            if r["pdf"]:
                pdf_b.pack(side="right", padx=(0, 10))
                dir_b.pack(side="right", padx=(0, 10))
        else:
            head.config(text="⚠️ 土地增值稅沒有算出來", fg="#B71C1C")
            sub.config(text="下面是試算程式的訊息（常見原因：沒有網路抓不到物價指數、謄本缺前次移轉現值）。"
                            "也可以回主程式按「財政部土增稅」再試。")
            total_lbl.config(text="")

    if not cmd:
        head.config(text="⚠️ 不能算土增稅", fg="#B71C1C")
        sub.config(text=err)
    else:
        def work():
            rc = -1
            try:
                pr = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                                      text=True, encoding="utf-8", errors="replace", env=env, cwd=cwd,
                                      creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                state["proc"] = pr
                for line in pr.stdout:
                    s = clean(line)
                    if s is None:
                        continue
                    state["lines"].append(s)
                    win.after(0, add, s)
                rc = pr.wait()
            except Exception as e:
                state["lines"].append(f"啟動試算程式失敗：{e}")
                try:
                    win.after(0, add, f"啟動試算程式失敗：{e}")
                except Exception:
                    pass
            try:
                win.after(0, finished, rc)
            except Exception:
                pass
        threading.Thread(target=work, daemon=True).start()

    try:
        win.lift()
        win.focus_force()
        win.grab_set()
    except Exception:
        pass
    if wait:
        master.wait_window(win)
    return win


def ask_and_run(master, json_path, base_dir, show_msg=None):
    """完成卡的「💰 算土增稅」按鈕用：找不到資料就說明，有就開試算視窗"""
    if not json_path or not os.path.exists(json_path):
        msg = "找不到這次轉好的謄本資料（JSON），沒辦法算土增稅。\n請確認輸出格式有勾 JSON，再處理一次。"
        if show_msg:
            show_msg("算土增稅", msg)
        return False
    if not is_transcript_json(json_path):
        if show_msg:
            show_msg("算土增稅", "這次轉的謄本裡沒有土地謄本（只有建物），土增稅只算土地。")
        return False
    run_tax_window(master, json_path, base_dir)
    return True

