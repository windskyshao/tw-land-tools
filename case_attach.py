"""連續模式・案件附件：依阿生地圖「案件附件📁」的做法，在使用者自己的電腦背景逐項跑子程式。
- 項目依存檔編號排(02 使照→03 重要設施→…→09 地質雲，實價放後面)；狀態直接看案件資料夾有沒有那個檔
- 預設不顯示瀏覽器(LAND_HEADLESS=1)，可勾「顯示瀏覽器視窗」；不顯示時有明顯的進度(第幾項、秒數、子程式最新訊息)
- 子程式停下等輸入時，依序送預排答案(與阿生地圖 case_runner 同一套)；每項有逾時，跑完關掉它自己開的 Chrome
- 全國土地使用分區(Cloudflare 真人驗證)：勾選時提醒；其他項全部跑完後才單獨開視窗跑，過程中有「繼續」按鈕
被 main.py import(打包進 exe)；子程式照主程式的方式啟動(打包版用 python_embedded)。"""
import os
import re
import sys
import json
import glob
import time
import queue
import threading
import subprocess
import tkinter as tk
from tkinter import ttk

import deal_ui as ui

NO_WINDOW = 0x08000000
FAM = "Microsoft JhengHei"
INK, INK2, MUTED = "#0F172A", "#334155", "#64748B"
GREEN, ORANGE, RED, BLUE = "#15803D", "#B45309", "#B91C1C", "#0A6CDB"

# key, 名稱, 子程式, 參數方式, 預排答案, 逾時秒, 縣市限制, 檔案(相對案件資料夾), 說明
ITEMS = [
    dict(key="nlma", name="02 使用執照（高雄市）", script="case_nlma.py", wrap="base", timeout=600, only="高雄市",
         files=["1.基本資料/02_使用執照*"], note="同地號有好幾張使照會全部存"),
    dict(key="nbupic", name="02 使用執照（屏東、台南…）", script="nlma_nbupic.py", answers=["1", "a", "1"], timeout=420,
         skip="高雄市", city_arg=True, files=["1.基本資料/02_使用執照*"]),
    dict(key="gbmap", name="03 重要設施", script="gbmap.py", answers=[""], timeout=180, files=["1.基本資料/03_重要設施*"]),
    dict(key="v523", name="04 V523 地籍資料", script="V523.py", answers=["", "n", ""], timeout=300, idle=20,
         only="高雄市|屏東縣", files=["1.基本資料/04_V523*"], note="自動登入（帳密在這台電腦）"),
    dict(key="easymap", name="05 地籍便民（新版自動截圖）", script="land_auto.py", extra=["--ver", "new"], answers=[""],
         timeout=420, files=["1.基本資料/05_地籍便民*"]),
    dict(key="nlsc", name="06 國土測繪圖", script="nlscmaps.py", answers=[""], timeout=300, files=["1.基本資料/06_國土測繪*"]),
    dict(key="urbangis", name="07 高雄都市計畫", script="urbangis.py", answers=["", ""], timeout=360, only="高雄市",
         files=["1.基本資料/07_都市計畫*"]),
    dict(key="tcdmap", name="08 國土功能分區", script="case_tcdmap.py", wrap="info", timeout=180,
         files=["1.基本資料/08_國土分區*"]),
    dict(key="geology", name="09 地質雲", script="geologycloud.py", answers=[""], timeout=300, files=["1.基本資料/09_地質雲*"]),
    dict(key="price", name="內政部實價登錄", script="interior_price.py", answers=["q"], timeout=420, idle=45, deal=True,
         files=["3.行情/內政部實價登錄*"]),
    dict(key="price104", name="104 實價網行情", script="price104.py", answers=["q"], timeout=420, idle=30, deal=True,
         files=["3.行情/104實價登錄*"], note="自動登入（帳密在這台電腦）"),
]
LUZ = dict(key="luz", name="07 全國土地使用分區", script="luz.py", timeout=900, files=["1.基本資料/07_全國土地使用分區*"],
           note="網站有真人驗證，要你在瀏覽器點一下")


def _paths():
    frozen = getattr(sys, "frozen", False)
    exe_dir = os.path.dirname(sys.executable) if frozen else os.path.dirname(os.path.abspath(__file__))
    internal = getattr(sys, "_MEIPASS", os.path.join(exe_dir, "_internal")) if frozen else exe_dir
    if frozen:
        py_dir = os.path.join(exe_dir, "python_embedded")
        py = os.path.join(py_dir, "python.exe")
        env = dict(os.environ, PYTHONHOME=py_dir,
                   PYTHONPATH=internal + os.pathsep + os.path.join(py_dir, "Lib", "site-packages"), PYTHONNOUSERSITE="1")
    else:
        py, env = sys.executable, dict(os.environ)
    env.update(PYTHONIOENCODING="utf-8", PYTHONUTF8="1", PYTHONUNBUFFERED="1")
    return exe_dir, internal, py, env


def _case():
    """目前案件(主程式資料夾 data.json 第一筆)→ dict(city, area, section, lot, folder)；沒有回 None"""
    exe_dir = _paths()[0]
    try:
        with open(os.path.join(exe_dir, "data.json"), encoding="utf-8") as f:
            d0 = json.load(f)[0]
    except Exception:
        return None
    lot = str(d0.get("lot_number", ""))
    folder = os.path.join(exe_dir, f"{d0.get('area', '')}{d0.get('section', '')}-{lot}")
    return dict(city=d0.get("city", ""), area=d0.get("area", ""), section=d0.get("section", ""), lot=lot,
                folder=folder, title=f"{d0.get('area', '')}{d0.get('section', '')}-{lot}")


def _deal_default(case):
    try:
        with open(os.path.join(case["folder"], "4.其他相關", "data_final.json"), encoding="utf-8") as f:
            df = json.load(f)
        df = df[0] if isinstance(df, list) else df
        return "租賃" if "租" in str(df.get("交易類型", "")) else "買賣"
    except Exception:
        return "買賣"


def applicable(item, city):
    c = (city or "").replace("臺", "台")
    if item.get("only") and c not in item["only"].replace("臺", "台").split("|"):
        return False
    if item.get("skip") and item["skip"].replace("臺", "台") == c:
        return False
    return True


def found_files(item, case):
    out = []
    for pat in item["files"]:
        out += [p for p in glob.glob(os.path.join(case["folder"], pat)) if p.lower().endswith((".pdf", ".png", ".xlsx"))]
    return out


def _kill_child_chrome(pid):
    ps = ("Get-CimInstance Win32_Process -Filter \"Name='chromedriver.exe' AND ParentProcessId=%d\" | "
          "ForEach-Object { $_.ProcessId }" % int(pid))
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, timeout=30,
                             creationflags=NO_WINDOW).stdout.decode("ascii", "ignore")
        for x in out.split():
            if x.isdigit():
                subprocess.run(["taskkill", "/PID", x, "/T", "/F"], capture_output=True, creationflags=NO_WINDOW)
    except Exception:
        pass


_ANSI = re.compile(r"\x1b\[[0-9;]*m")


class Runner:
    """在背景執行緒逐項跑；訊息丟進 q 給畫面"""

    def __init__(self, case, items, show_browser, deal, q):
        self.case, self.items, self.show, self.deal, self.q = case, items, show_browser, deal, q
        self.stop = False
        self.proc = None

    def send(self, text):
        try:
            self.proc.stdin.write((text + "\n").encode("utf-8"))
            self.proc.stdin.flush()
        except Exception:
            pass

    def run_one(self, item, interactive=False):
        exe_dir, internal, py, env = _paths()
        if not self.show and not interactive:
            env["LAND_HEADLESS"] = "1"
        c = self.case
        script = os.path.join(internal, item["script"])
        if item.get("wrap") == "base":                       # case_nlma：輸出到案件資料夾(它自己再加 1.基本資料)
            args = [py, "-u", script, c["city"], c["area"], c["section"], c["lot"], c["folder"]]
        elif item.get("wrap") == "info":                     # case_tcdmap：輸出到 1.基本資料
            os.makedirs(os.path.join(c["folder"], "1.基本資料"), exist_ok=True)
            args = [py, "-u", script, c["city"], c["area"], c["section"], c["lot"], os.path.join(c["folder"], "1.基本資料")]
        else:
            args = [py, "-u", script, "--base-dir", exe_dir] + list(item.get("extra") or [])
            if item.get("city_arg"):
                args += ["--city", c["city"]]
            if item.get("deal"):
                args += ["--deal", self.deal]
        t0 = time.time()
        before = {p: os.path.getmtime(p) for p in found_files(item, c)}
        lines, last = [], [time.time()]
        self.proc = p = subprocess.Popen(args, cwd=internal, env=env, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                         stderr=subprocess.STDOUT, creationflags=NO_WINDOW)

        def rd():
            for b in iter(p.stdout.readline, b""):
                t = _ANSI.sub("", b.decode("utf-8", "replace")).strip()
                if t:
                    lines.append(t)
                    last[0] = time.time()
                    if not t.startswith("[DEBUG]"):
                        self.q.put(("line", item["key"], t))
        threading.Thread(target=rd, daemon=True).start()
        answers, sent = list(item.get("answers") or []), 0
        while p.poll() is None and time.time() - t0 < item["timeout"] and not self.stop:
            time.sleep(1)
            self.q.put(("tick", item["key"], int(time.time() - t0)))
            if interactive:                                  # 全國土地使用分區：使用者在畫面按「繼續」才送 Enter
                continue
            if not item.get("wrap") and time.time() - last[0] > item.get("idle", 12) and sent < 8:
                ans = answers.pop(0) if answers else "q"
                self.send(ans)
                last[0] = time.time()
                sent += 1
        timed_out = p.poll() is None
        if timed_out:
            subprocess.run(["taskkill", "/PID", str(p.pid), "/T", "/F"], capture_output=True, creationflags=NO_WINDOW)
        time.sleep(1)
        _kill_child_chrome(p.pid)
        txt = "\n".join(lines)
        try:
            os.makedirs(os.path.join(c["folder"], "4.其他相關", "連續模式紀錄"), exist_ok=True)
            with open(os.path.join(c["folder"], "4.其他相關", "連續模式紀錄", item["key"] + ".txt"), "w",
                      encoding="utf-8") as f:
                f.write(txt)
        except Exception:
            pass
        new = [x for x in found_files(item, c) if x not in before or os.path.getmtime(x) > before[x] + 1]
        sec = int(time.time() - t0)
        if new:
            msg = ""
            if "WARN:" in txt:
                msg = "⚠ " + txt[txt.rfind("WARN:") + 5:].split("\n")[0].strip()
            return "done", msg, sec
        if self.stop:
            return "stopped", "已停止", sec
        if "NODATA" in txt or "查無" in txt[-800:]:
            return "nodata", "官方查無這筆資料", sec
        if re.search(r"網站.{0,6}故障|HTTP 500|伺服器錯誤", txt):
            return "failed", "對方網站故障，請稍後重跑", sec
        if "FAIL:" in txt:
            return "failed", txt[txt.rfind("FAIL:") + 5:].split("\n")[0].strip()[:60], sec
        if timed_out:
            return "failed", f"超過 {item['timeout']} 秒沒完成，已中止", sec
        return "failed", "沒有產生檔案（紀錄在 4.其他相關\\連續模式紀錄）", sec

    def run(self):
        normal = [i for i in self.items if i["key"] != "luz"]
        luz = [i for i in self.items if i["key"] == "luz"]
        total = len(self.items)
        for n, item in enumerate(normal + luz, 1):
            if self.stop:
                break
            inter = item["key"] == "luz"
            self.q.put(("start", item["key"], n, total, inter))
            try:
                st, msg, sec = self.run_one(item, interactive=inter)
            except Exception as e:
                st, msg, sec = "failed", str(e)[:60], 0
            self.q.put(("end", item["key"], st, msg, sec))
        self.q.put(("alldone",))


def open_window(root, on_finish=None, old_mode=None):
    case = _case()
    if not case:
        ui.message(root, "連續模式", "目前沒有案件（data.json 是空的），請先用地籍便民或快速建立案件建立案件。")
        return
    win = tk.Toplevel(root)
    win.title("連續模式・案件附件")
    ui.apply(win)
    sc = max(1.0, win.winfo_fpixels("1i") / 96)
    W, H = int(820 * sc), int(860 * sc)
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    W, H = min(W, sw - 40), min(H, sh - 80)
    win.geometry(f"{W}x{H}+{(sw - W) // 2}+{max(0, (sh - H) // 2 - 20)}")
    win.transient(root)
    ui.header(win, "🧩 連續模式・案件附件", f"{case['title']}　勾選要抓的附件，按「開始」逐項自動完成")
    foot = ui.footer(win)
    body, canvas = ui.scroll_body(win)
    bg = ui.BG

    # ---- 選項 ----
    opt = tk.Frame(body, bg=bg)
    opt.pack(fill=tk.X, pady=(0, 6))
    show_var = tk.BooleanVar(master=win, value=False)
    ui.big_check(opt, show_var).pack(side=tk.LEFT)
    tk.Label(opt, text="顯示瀏覽器視窗", font=(FAM, 14, "bold"), fg=INK, bg=bg).pack(side=tk.LEFT, padx=(6, 4))
    tk.Label(opt, text="（不勾＝背景執行，比較快、不干擾你做別的事）", font=(FAM, 12), fg=INK2, bg=bg).pack(side=tk.LEFT)
    opt2 = tk.Frame(body, bg=bg)
    opt2.pack(fill=tk.X, pady=(0, 8))
    tk.Label(opt2, text="實價登錄查：", font=(FAM, 14, "bold"), fg=INK, bg=bg).pack(side=tk.LEFT)
    deal_var = tk.StringVar(master=win, value=_deal_default(case))
    ttk.Combobox(opt2, textvariable=deal_var, values=["買賣", "租賃", "兩者"], state="readonly", width=6,
                 font=(FAM, 13)).pack(side=tk.LEFT, padx=6)
    tk.Label(opt2, text="（依物件編輯器的交易類型預選）", font=(FAM, 12), fg=INK2, bg=bg).pack(side=tk.LEFT)

    rows = {}

    def status_text(item):
        if not applicable(item, case["city"]):
            return "— 這個縣市不適用", MUTED
        fs = found_files(item, case)
        if fs:
            t = max(os.path.getmtime(x) for x in fs)
            return f"✓ 已有（{time.strftime('%m/%d %H:%M', time.localtime(t))}）", GREEN
        return "○ 還沒有", ORANGE

    def add_row(parent, item, warn=False):
        ok = applicable(item, case["city"])
        fr = tk.Frame(parent, bg="white", highlightthickness=1, highlightbackground="#CBD5E1", padx=10, pady=6)
        fr.pack(fill=tk.X, pady=3)
        st, col = status_text(item)
        v = tk.BooleanVar(master=win, value=ok and not found_files(item, case) and not warn)
        cb = ui.big_check(fr, v, enabled=ok)
        cb.pack(side=tk.LEFT)
        mid = tk.Frame(fr, bg="white")
        mid.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
        tk.Label(mid, text=item["name"], font=(FAM, 15, "bold"), fg=INK if ok else MUTED, bg="white", anchor="w").pack(anchor="w")
        if item.get("note"):
            tk.Label(mid, text=item["note"], font=(FAM, 11), fg=ORANGE if warn else INK2, bg="white", anchor="w").pack(anchor="w")
        sl = tk.Label(fr, text=st, font=(FAM, 13, "bold"), fg=col, bg="white", width=22, anchor="e")
        sl.pack(side=tk.LEFT, padx=(6, 6))
        b = ui.button(fr, "單跑", lambda it=item: start([it]), "ghost", padx=10, pady=2)
        if ok:
            b.pack(side=tk.LEFT)
        rows[item["key"]] = dict(var=v, status=sl, frame=fr, item=item)
        if warn:
            def on_check(*_):
                if v.get():
                    ui.message(win, "全國土地使用分區",
                               "這個網站有「真人驗證」（Cloudflare），一定要你自己在瀏覽器上點驗證。\n\n"
                               "勾選後，它會等其他項目全部跑完，最後才單獨開瀏覽器視窗，"
                               "請在電腦前完成驗證，再按這裡的「繼續」。")
            v.trace_add("write", on_check)

    tk.Label(body, text="自動完成（背景執行）", font=(FAM, 16, "bold"), fg=BLUE, bg=bg, anchor="w").pack(fill=tk.X, pady=(6, 2))
    lst = tk.Frame(body, bg=bg)
    lst.pack(fill=tk.X)
    for it in ITEMS:
        add_row(lst, it)
    tk.Label(body, text="需要你操作（最後單獨跑）", font=(FAM, 16, "bold"), fg=ORANGE, bg=bg, anchor="w").pack(fill=tk.X, pady=(14, 2))
    lst2 = tk.Frame(body, bg=bg)
    lst2.pack(fill=tk.X)
    add_row(lst2, LUZ, warn=True)
    tk.Label(body, text="已經有檔案的預設不勾；要重抓就勾起來或按「單跑」。檔案存到案件資料夾的 1.基本資料、3.行情。",
             font=(FAM, 12), fg=INK2, bg=bg, wraplength=W - 80, justify="left").pack(anchor="w", pady=(10, 0))

    # ---- 進度區(固定在下方，背景執行時看得到在做什麼) ----
    prog = tk.Frame(win, bg="#EEF4FF", highlightthickness=1, highlightbackground="#B9D4F5", padx=14, pady=10)
    prog_lbl = tk.Label(prog, text="", font=(FAM, 15, "bold"), fg=INK, bg="#EEF4FF", anchor="w", justify="left")
    prog_lbl.pack(fill=tk.X)
    bar_item = ttk.Progressbar(prog, mode="indeterminate")
    bar_item.pack(fill=tk.X, pady=(6, 2))
    bar_all = ttk.Progressbar(prog, mode="determinate")
    bar_all.pack(fill=tk.X, pady=(2, 4))
    line_lbl = tk.Label(prog, text="", font=(FAM, 11), fg=INK2, bg="#EEF4FF", anchor="w", justify="left", wraplength=W - 80)
    line_lbl.pack(fill=tk.X)
    cont_btn = ui.button(prog, "我完成驗證了，繼續 ▶", lambda: runner_box[0] and runner_box[0].send(""), "primary")

    q = queue.Queue()
    runner_box = [None]
    state = {"done": 0, "total": 0, "cur": None, "results": []}
    btns = tk.Frame(foot, bg=foot.cget("bg"))
    btns.pack(side=tk.RIGHT)

    def start(sel=None):
        if runner_box[0]:
            return
        items = sel or [r["item"] for r in rows.values() if r["var"].get() and applicable(r["item"], case["city"])]
        if not items:
            ui.message(win, "連續模式", "沒有勾選任何項目。")
            return
        order = [i["key"] for i in ITEMS] + ["luz"]
        items.sort(key=lambda i: order.index(i["key"]))
        state.update(done=0, total=len(items), results=[])
        prog.pack(fill=tk.X, side=tk.BOTTOM, before=foot)
        bar_all.configure(maximum=len(items), value=0)
        r = Runner(case, items, show_var.get(), deal_var.get(), q)
        runner_box[0] = r
        threading.Thread(target=r.run, daemon=True).start()

    def stop():
        if runner_box[0]:
            runner_box[0].stop = True
            prog_lbl.configure(text="正在停止…（等目前這一項結束）")

    def pump():
        try:
            while True:
                m = q.get_nowait()
                kind = m[0]
                if kind == "start":
                    _, key, n, total, inter = m
                    state["cur"] = key
                    rows[key]["status"].configure(text="⏳ 執行中…", fg=BLUE)
                    rows[key]["frame"].configure(highlightbackground=BLUE, highlightthickness=2)
                    prog_lbl.configure(text=f"第 {n}/{total} 項：{rows[key]['item']['name']}")
                    bar_item.start(12)
                    if inter:
                        cont_btn.pack(anchor="w", pady=(6, 0))
                        line_lbl.configure(text="已開啟瀏覽器：請完成真人驗證，網頁出現查詢結果後按「繼續」")
                elif kind == "tick":
                    _, key, sec = m
                    if key == state["cur"]:
                        t = prog_lbl.cget("text").split("　")[0]
                        prog_lbl.configure(text=f"{t}　已 {sec} 秒")
                elif kind == "line":
                    line_lbl.configure(text="→ " + m[2][:120])
                elif kind == "end":
                    _, key, st, msg, sec = m
                    bar_item.stop()
                    cont_btn.pack_forget()
                    state["done"] += 1
                    bar_all.configure(value=state["done"])
                    rows[key]["frame"].configure(highlightbackground="#CBD5E1", highlightthickness=1)
                    if st == "done":
                        rows[key]["status"].configure(text=f"✓ 完成（{sec} 秒）" + (" ⚠" if msg else ""), fg=GREEN)
                        rows[key]["var"].set(False)
                    elif st == "nodata":
                        rows[key]["status"].configure(text="— 官方查無資料", fg=MUTED)
                    else:
                        rows[key]["status"].configure(text=f"✗ {msg[:18]}", fg=RED)
                    state["results"].append((rows[key]["item"]["name"], st, msg))
                elif kind == "alldone":
                    runner_box[0] = None
                    bar_item.stop()
                    ok = sum(1 for _, s, _m in state["results"] if s == "done")
                    bad = [(n_, m_) for n_, s, m_ in state["results"] if s not in ("done", "nodata")]
                    prog_lbl.configure(text=f"完成：{ok}/{state['total']} 項成功" + (f"，{len(bad)} 項失敗" if bad else ""))
                    line_lbl.configure(text="；".join(f"{n_}：{m_}" for n_, m_ in bad)[:300] if bad else "全部完成，檔案在案件資料夾")
                    if on_finish:
                        on_finish(state["results"])
        except queue.Empty:
            pass
        try:
            win.after(300, pump)
        except tk.TclError:
            pass

    def on_close():
        if runner_box[0]:
            if not ui.message(win, "連續模式", "還在執行中，要停止並關閉嗎？", kind="ask", ok="停止並關閉", cancel="繼續執行"):
                return
            stop()
        win.destroy()
    win.protocol("WM_DELETE_WINDOW", on_close)
    if old_mode:                                     # 謄本、結構化、土增稅、編輯物件資料這些要人操作的，照舊版連續模式跑
        ui.button(foot, "舊版連續模式（謄本、土增稅…）", lambda: (win.destroy(), old_mode()), "ghost").pack(side=tk.LEFT)
    ui.button(btns, "關閉", on_close, "ghost").pack(side=tk.LEFT, padx=(0, 8))
    ui.button(btns, "■ 停止", stop, "ghost").pack(side=tk.LEFT, padx=(0, 8))
    run_btn = ui.button(btns, "▶ 開始", lambda: start(), "primary")
    run_btn.pack(side=tk.LEFT)
    win.after(300, pump)
    return win
