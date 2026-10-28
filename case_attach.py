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
# 土地增值稅：用目前選用的謄本結構化 JSON(--json)；跑之前先檢查結構化資料是不是這個案件(跟主程式按鈕一樣)
TAX = [
    dict(key="ftax", name="01 財政部土增稅", script="finance_tax.py", json=True, timeout=3600, idle=15, default="",
         max_send=80, hide_flag="--headless", files=["1.基本資料/01_土地增值稅_財政部版*", "1.基本資料/01.1_土地增值稅_財政部版*"],
         note="產生試算 PDF＋Excel 列表"),
    dict(key="ktax", name="01-1 高雄市土增稅估算", script="kaohsiung_tax.py", json=True, timeout=3600, idle=15, default="",
         max_send=80, show_flag="--show", only="高雄市",
         files=["1.基本資料/01-1_土地增值稅_高雄市版*", "1.基本資料/01-1.1_土地增值稅_高雄市版*"],
         note="高雄市稅捐處網站，驗證碼自動辨識"),
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

    def __init__(self, case, items, show_browser, deal, q, json_path=None):
        self.case, self.items, self.show, self.deal, self.q = case, items, show_browser, deal, q
        self.json_path = json_path
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
            if item.get("json") and self.json_path:
                args += ["--json", self.json_path]
            if item.get("show_flag") and (self.show or interactive):
                args.append(item["show_flag"])
            if item.get("hide_flag") and not (self.show or interactive):
                args.append(item["hide_flag"])
        t0 = time.time()
        before = {p: os.path.getmtime(p) for p in found_files(item, c)}
        lines, last = [], [time.time()]
        self.proc = p = subprocess.Popen(args, cwd=internal, env=env, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                         stderr=subprocess.STDOUT, creationflags=NO_WINDOW)

        def rd():
            for b in iter(p.stdout.readline, b""):
                t = _ANSI.sub("", b.decode("utf-8", "replace")).strip()
                if t and not t.startswith(("__PROGRESS__", "__GUI_PREFILL__")):
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
            if not item.get("wrap") and time.time() - last[0] > item.get("idle", 12) and sent < item.get("max_send", 8):
                ans = answers.pop(0) if answers else item.get("default", "q")
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


def open_folder(path):
    try:
        os.startfile(path)
    except Exception:
        subprocess.Popen(["explorer", path])


# ---- 合併列印：已合併過的檔案(開著沒、改過沒、附件有沒有變) ----
MERGE_PREFIX = "00_合併列印_"
MERGE_LOG = os.path.join("4.其他相關", "合併列印紀錄.json")
_cm_mod = [None]


def _case_merge():
    """載入 case_merge.py(只用它挑檔/排序；fitz 在它裡面是用到才 import)"""
    if _cm_mod[0] is None:
        import importlib.util
        spec = importlib.util.spec_from_file_location("case_merge", os.path.join(_paths()[1], "case_merge.py"))
        m = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(m)
        _cm_mod[0] = m
    return _cm_mod[0]


def file_locked(path):
    """被別的程式(Acrobat 等)開著鎖住 → True"""
    import ctypes
    k = ctypes.windll.kernel32
    k.CreateFileW.restype = ctypes.c_void_p
    h = k.CreateFileW(str(path), 0x80000000, 0, None, 3, 0x80, None)      # GENERIC_READ、不共用、OPEN_EXISTING
    if h is None or h == ctypes.c_void_p(-1).value:
        return k.GetLastError() in (32, 33)                                 # 共用違規／鎖定違規
    k.CloseHandle(ctypes.c_void_p(h))
    return False


def _find_window(name):
    """標題含檔名的視窗(Acrobat「xxx.pdf - Adobe Acrobat」、瀏覽器分頁等)"""
    import ctypes
    from ctypes import wintypes
    u = ctypes.windll.user32
    stem = os.path.splitext(name)[0]
    found = []

    def cb(hwnd, _):
        if u.IsWindowVisible(hwnd):
            n = u.GetWindowTextLengthW(hwnd)
            if n:
                buf = ctypes.create_unicode_buffer(n + 1)
                u.GetWindowTextW(hwnd, buf, n + 1)
                if stem in buf.value:
                    found.append(hwnd)
                    return False
        return True
    u.EnumWindows(ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(cb), 0)
    return found[0] if found else None


def open_or_focus(path):
    """已經開著 → 切換過去；沒開 → 用預設 PDF 程式打開。回 focused / opened / failed"""
    try:
        h = _find_window(os.path.basename(path))
        if h:
            import ctypes
            u = ctypes.windll.user32
            if u.IsIconic(h):
                u.ShowWindow(h, 9)                   # 縮到工作列的先還原
            u.SetForegroundWindow(h)
            return "focused"
    except Exception:
        pass
    try:
        os.startfile(path)                           # Acrobat 開著但在別的分頁：它會自己切到這份
        return "opened"
    except Exception:
        return "failed"


def merge_record(folder):
    try:
        with open(os.path.join(folder, MERGE_LOG), encoding="utf-8") as f:
            rec = json.load(f)
    except Exception:
        return {}
    src = rec.get("sources")
    if src is not None and any("/" not in k for k in src):      # 舊紀錄只存檔名(都是 1.基本資料)
        rec["sources"] = {(k if "/" in k else "1.基本資料/" + k): v for k, v in src.items()}
    return rec


def merge_status(folder):
    """已合併過的檔案狀態；沒有回 None。
    changes：之後附件的變動(新增/更新/拿掉)；edited：合併後在 PDF 程式裡改過存檔；locked：目前開著"""
    info = os.path.join(folder, "1.基本資料")
    olds = sorted(glob.glob(os.path.join(info, MERGE_PREFIX + "*.pdf")), key=os.path.getmtime)
    if not olds:
        return None
    path = olds[-1]
    mt = os.path.getmtime(path)
    rec = merge_record(folder)
    if os.path.basename(rec.get("file", "")) != os.path.basename(path):
        rec = {}
    try:
        cm = _case_merge()
        cur = {cm.rel(folder, x): os.path.getmtime(x) for _sec, x in cm.candidates(folder)}
    except Exception:
        cur = {}
    changes = []
    src = rec.get("sources")
    if src is not None:
        skip = set(rec.get("unpicked") or [])
        old_fmt = "unpicked" not in rec              # 加入謄本挑選之前的合併：0.謄本 本來就沒合併，不算新增
        for n in sorted(cur):
            if old_fmt and n.startswith(("0.謄本/", "3.行情/")):
                continue
            if n not in src and n not in skip:
                changes.append("新增 " + n)
            elif n in src and cur[n] > src[n] + 2:
                changes.append("更新 " + n)
        changes += ["拿掉 " + n for n in sorted(src) if n not in cur]
    else:                                            # 舊版合併沒有紀錄：比檔案時間
        changes = ["更新 " + n for n, t in sorted(cur.items()) if t > mt + 2]
    if src is None:
        changes = [c for c in changes if c.split(" ", 1)[1].startswith("1.基本資料")]
    no_trans = 0                                     # 舊版合併(還沒有謄本挑選)：0.謄本 一份都沒放進去
    if "unpicked" not in rec:
        no_trans = sum(1 for n in cur if n.startswith(("0.謄本/", "3.行情/")))
    return dict(path=path, mtime=mt, changes=changes, locked=file_locked(path), no_trans=no_trans,
                edited=bool(rec.get("mtime")) and mt > rec["mtime"] + 2)


def pick_merge_files(parent, folder):
    """合併前挑檔案：列出 0.謄本＋1.基本資料，每份可勾、可「預覽」(用 PDF 程式打開)。
    預設：上次有勾的照勾、上次沒勾的不勾；新檔依規則(附件全勾；謄本同一張圖只勾標註版)。
    回傳 (要合併的路徑[已排序], 沒勾的相對路徑[]) 或 None(取消)"""
    cm = _case_merge()
    cands = cm.candidates(folder)
    if not cands:
        ui.message(parent, "合併列印", "0.謄本 和 1.基本資料 裡都沒有可以合併的檔案。")
        return None
    rec = merge_record(folder)
    picked_before, unpicked_before = set(rec.get("sources") or {}), set(rec.get("unpicked") or [])
    tdef = cm.default_picks([p for sec, p in cands if sec == "0.謄本"])

    top = tk.Toplevel(parent)
    top.withdraw()
    top.title("合併列印・選擇要合併的檔案")
    ui.apply(top)
    top.transient(parent)
    ui.header(top, "📑 合併列印", "勾選要合併的檔案，可以先按「預覽」看內容；由上到下就是合併後的順序")
    foot = ui.footer(top)
    body, _cv = ui.scroll_body(top)
    bg = ui.BG
    res = {"v": None}
    rows = []
    allv = tk.BooleanVar(master=top, value=False)
    sync = [False]

    def upd(*_):
        if sync[0]:
            return
        sync[0] = True
        n = sum(1 for r in rows if r[1].get())
        allv.set(bool(rows) and n == len(rows))
        go.itemconfigure(go._text, text=f"▶ 合併 {n} 份")
        sync[0] = False

    def on_all(*_):
        if sync[0]:
            return
        sync[0] = True
        for r in rows:
            r[1].set(allv.get())
        sync[0] = False
        upd()
    ar = tk.Frame(body, bg=bg)
    ar.pack(fill=tk.X, pady=(0, 4))
    ui.big_check(ar, allv).pack(side=tk.LEFT)
    tk.Label(ar, text="全選／全不選", font=(FAM, 14, "bold"), fg=INK, bg=bg).pack(side=tk.LEFT, padx=(6, 0))
    allv.trace_add("write", on_all)
    cur_sec = None
    for sec, path in cands:
        if sec != cur_sec:
            cur_sec = sec
            tk.Label(body, text={"0.謄本": "0.謄本（放在最前面）", "3.行情": "3.行情（實價登錄，放在最後）"}.get(
                sec, "1.基本資料（附件，依編號）"),
                     font=(FAM, 16, "bold"), fg=BLUE, bg=bg, anchor="w").pack(fill=tk.X, pady=(10, 2))
        key = cm.rel(folder, path)
        if key in picked_before:
            on = True
        elif key in unpicked_before:
            on = False
        else:
            on = path in tdef if sec == "0.謄本" else True
        fr = tk.Frame(body, bg="white", highlightthickness=1, highlightbackground="#CBD5E1", padx=10, pady=5)
        fr.pack(fill=tk.X, pady=2)
        v = tk.BooleanVar(master=top, value=on)
        ui.big_check(fr, v).pack(side=tk.LEFT)
        name = os.path.basename(path)
        lb = tk.Label(fr, text=name, font=(FAM, 13, "bold"), fg=INK, bg="white", anchor="w", justify="left")
        lb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 6))
        lb.bind("<Button-1>", lambda e, v=v: v.set(not v.get()))
        ui.button(fr, "預覽", lambda p=path: open_or_focus(p), "ghost", padx=12, pady=2).pack(side=tk.RIGHT)
        tk.Label(fr, text=time.strftime("%m/%d %H:%M", time.localtime(os.path.getmtime(path))), font=(FAM, 12),
                 fg=INK2, bg="white").pack(side=tk.RIGHT, padx=(6, 10))
        v.trace_add("write", upd)
        rows.append((key, v, path))

    def done(ok):
        if ok:
            files = [p for _k, v, p in rows if v.get()]
            if not files:
                ui.message(top, "合併列印", "沒有勾選任何檔案。")
                return
            res["v"] = (files, [k for k, v, _p in rows if not v.get()])
        top.destroy()
    btns = tk.Frame(foot, bg=foot.cget("bg"))
    btns.pack(side=tk.RIGHT)
    ui.button(btns, "取消", lambda: done(False), "ghost").pack(side=tk.LEFT, padx=(0, 8))
    go = ui.button(btns, "▶ 合併 999 份", lambda: done(True), "primary")
    go.pack(side=tk.LEFT)
    tk.Label(foot, text="「預覽」會用你的 PDF 程式打開該檔", font=(FAM, 12), fg=INK2, bg=foot.cget("bg")).pack(side=tk.LEFT)
    upd()
    top.protocol("WM_DELETE_WINDOW", lambda: done(False))
    top.bind("<Escape>", lambda e: done(False))
    top.update_idletasks()
    try:
        sc = max(1.0, top.winfo_fpixels("1i") / 96)
        W = min(int(900 * sc), top.winfo_screenwidth() - 40)
        H = min(int(860 * sc), top.winfo_screenheight() - 80)
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - W) // 2)
        top.geometry(f"{W}x{H}+{max(0, x)}+{max(0, parent.winfo_rooty())}")
    except Exception:
        pass
    top.deiconify()
    top.grab_set()
    top.focus_set()
    parent.wait_window(top)
    return res["v"]


def open_window(root, on_finish=None, old_mode=None, on_line=None, tax_check=None, json_path=None):
    """on_line(text)：每一項的過程訊息也送一份到主程式的訊息欄(跟單獨執行時一樣看得到提示)
    tax_check()：謄本結構化有沒有問題(""=沒問題，否則是原因)；json_path()：目前選用的結構化 JSON"""
    def tax_problem():
        if not tax_check:
            return ""
        try:
            return tax_check() or ""
        except Exception:
            return ""

    def cur_json():
        try:
            p = json_path() if callable(json_path) else json_path
            return p if p and os.path.exists(p) else None
        except Exception:
            return None
    def say(t):
        if on_line:
            try:
                on_line(t)
            except Exception:
                pass

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
    try:                                       # 跟其他子程式一樣：佔主程式左邊的空間(主程式固定在右側 1/3)，不蓋住主程式
        rx = root.winfo_x()
        if rx > 600:
            W, H = rx - 10, sh - 80
            win.geometry(f"{W}x{H}+0+0")
        else:
            raise ValueError
    except (tk.TclError, ValueError):
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

    opt3 = tk.Frame(body, bg=bg)
    opt3.pack(fill=tk.X, pady=(0, 8))
    duplex_var = tk.BooleanVar(master=win, value=False)
    ui.big_check(opt3, duplex_var).pack(side=tk.LEFT)
    tk.Label(opt3, text="合併列印時每份文件從新的一張紙開始", font=(FAM, 14, "bold"), fg=INK, bg=bg).pack(side=tk.LEFT, padx=(6, 4))
    tk.Label(opt3, text="（單數頁的文件後面會補空白頁；不勾＝直接接著印，不浪費紙）", font=(FAM, 12), fg=INK2, bg=bg).pack(side=tk.LEFT)

    rows = {}

    def status_text(item):
        if not applicable(item, case["city"]):
            return "— 這個縣市不適用", MUTED
        fs = found_files(item, case)
        if fs:
            t = max(os.path.getmtime(x) for x in fs)
            return f"✓ 已有（{time.strftime('%m/%d %H:%M', time.localtime(t))}）", GREEN
        return "○ 還沒有", ORANGE

    def add_row(parent, item, warn=False, default=None):
        ok = applicable(item, case["city"])
        fr = tk.Frame(parent, bg="white", highlightthickness=1, highlightbackground="#CBD5E1", padx=10, pady=6)
        fr.pack(fill=tk.X, pady=3)
        st, col = status_text(item)
        dv = ok and not found_files(item, case) and not warn
        if default is not None:
            dv = dv and default
        v = tk.BooleanVar(master=win, value=dv)
        cb = ui.big_check(fr, v, enabled=ok)
        cb.pack(side=tk.LEFT)
        mid = tk.Frame(fr, bg="white")
        mid.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
        tk.Label(mid, text=item["name"], font=(FAM, 15, "bold"), fg=INK if ok else MUTED, bg="white", anchor="w").pack(anchor="w")
        if item.get("note"):
            tk.Label(mid, text=item["note"], font=(FAM, 11), fg=ORANGE if warn else INK2, bg="white", anchor="w").pack(anchor="w")
        if item.get("json") and ok and tax_reason:
            tk.Label(mid, text="⚠ " + tax_reason, font=(FAM, 12, "bold"), fg=RED, bg="white", anchor="w", justify="left",
                     wraplength=max(300, W - 380)).pack(anchor="w")
        sl = tk.Label(fr, text=st, font=(FAM, 13, "bold"), fg=col, bg="white", width=22, anchor="e")
        sl.pack(side=tk.LEFT, padx=(6, 6))
        b = ui.button(fr, "單跑", lambda it=item: start([it]), "ghost", padx=10, pady=2)
        if ok:
            b.pack(side=tk.LEFT)
        rows[item["key"]] = dict(var=v, status=sl, frame=fr, item=item)
        if ok and item["key"] != "luz":
            v.trace_add("write", lambda *_: sync_all())
        if warn:
            def on_check(*_):
                if v.get():
                    ui.message(win, "全國土地使用分區",
                               "這個網站有「真人驗證」（Cloudflare），一定要你自己在瀏覽器上點驗證。\n\n"
                               "勾選後，它會等其他項目全部跑完，最後才單獨開瀏覽器視窗，"
                               "請在電腦前完成驗證，再按這裡的「繼續」。")
            v.trace_add("write", on_check)

    # ---- 全選／全不選(全國土地使用分區要真人驗證，不跟著全選) ----
    all_row = tk.Frame(body, bg=bg)
    all_row.pack(fill=tk.X, pady=(2, 0))
    all_var = tk.BooleanVar(master=win, value=False)
    _sync = [False]

    def _pickable():
        return [r for k, r in rows.items() if k != "luz" and applicable(r["item"], case["city"])]

    def on_all(*_):
        if _sync[0]:
            return
        _sync[0] = True
        for r in _pickable():
            r["var"].set(all_var.get())
        _sync[0] = False

    def sync_all():
        if _sync[0]:
            return
        _sync[0] = True
        ps = _pickable()
        all_var.set(bool(ps) and all(r["var"].get() for r in ps))
        _sync[0] = False
    ui.big_check(all_row, all_var).pack(side=tk.LEFT)
    tk.Label(all_row, text="全選／全不選", font=(FAM, 14, "bold"), fg=INK, bg=bg).pack(side=tk.LEFT, padx=(6, 4))
    tk.Label(all_row, text="（不含最下面的全國土地使用分區）", font=(FAM, 12), fg=INK2, bg=bg).pack(side=tk.LEFT)
    all_var.trace_add("write", on_all)

    tax_reason = tax_problem()
    tk.Label(body, text="土地增值稅（用謄本結構化資料計算）", font=(FAM, 16, "bold"), fg=BLUE, bg=bg, anchor="w").pack(
        fill=tk.X, pady=(6, 2))
    lst0 = tk.Frame(body, bg=bg)
    lst0.pack(fill=tk.X)
    for it in TAX:                                   # 預設只勾財政部；結構化資料不是這個案件就都不勾
        add_row(lst0, it, default=(it["key"] == "ftax" and not tax_reason))
    tk.Label(body, text="自動完成（背景執行）", font=(FAM, 16, "bold"), fg=BLUE, bg=bg, anchor="w").pack(fill=tk.X, pady=(14, 2))
    lst = tk.Frame(body, bg=bg)
    lst.pack(fill=tk.X)
    for it in ITEMS:
        add_row(lst, it)
    tk.Label(body, text="需要你操作（最後單獨跑）", font=(FAM, 16, "bold"), fg=ORANGE, bg=bg, anchor="w").pack(fill=tk.X, pady=(14, 2))
    lst2 = tk.Frame(body, bg=bg)
    lst2.pack(fill=tk.X)
    add_row(lst2, LUZ, warn=True)
    sync_all()
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
    float_box = [None]

    def show_float(name):
        """要人操作的項目(全國土地使用分區)：右下角永遠在最上層的小提示框，不會被瀏覽器蓋住。
        連續模式視窗先藏起來：它掛在主程式底下，點主程式時會跟著跳出來擋住瀏覽器"""
        hide_float(restore=False)
        try:
            win.withdraw()
        except tk.TclError:
            pass
        f = tk.Toplevel(root)
        f.title("連續模式：需要你操作")
        f.attributes("-topmost", True)
        f.configure(bg="#FFF4E5", highlightthickness=2, highlightbackground="#E58135")
        tk.Label(f, text=f"🖐 {name}", font=(FAM, 16, "bold"), fg="#7C2D12", bg="#FFF4E5").pack(anchor="w", padx=16, pady=(12, 4))
        tk.Label(f, text="1. 在瀏覽器上完成「真人驗證」（點一下勾選框）\n"
                         "2. 等網頁出現查詢結果（地號定位、分區圖）\n"
                         "3. 再按下面的「我完成驗證了，繼續」",
                 font=(FAM, 13), fg="#1C1917", bg="#FFF4E5", justify="left").pack(anchor="w", padx=16)
        f._line = tk.Label(f, text="", font=(FAM, 11), fg="#57534E", bg="#FFF4E5", wraplength=380, justify="left",
                           height=2, anchor="nw")      # 固定兩行高：訊息再長也不會把按鈕擠出框外
        f._line.pack(anchor="w", padx=16, pady=(6, 0))
        ui.button(f, "我完成驗證了，繼續 ▶", lambda: runner_box[0] and runner_box[0].send(""), "primary").pack(
            anchor="e", padx=16, pady=12)
        f.update_idletasks()
        w, h = max(f.winfo_reqwidth(), 420), f.winfo_reqheight()
        f.geometry(f"{w}x{h}+{f.winfo_screenwidth() - w - 24}+{f.winfo_screenheight() - h - 70}")
        f.protocol("WM_DELETE_WINDOW", lambda: None)            # 不讓它被誤關(做完會自己關)
        float_box[0] = f

    def hide_float(restore=True):
        if float_box[0] is not None:
            try:
                float_box[0].destroy()
            except Exception:
                pass
            float_box[0] = None
        if restore:                                      # 這一項做完：連續模式視窗再叫回來
            try:
                if win.state() == "withdrawn":
                    win.deiconify()
            except tk.TclError:
                pass

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
        order = [i["key"] for i in TAX + ITEMS] + ["luz"]
        items.sort(key=lambda i: order.index(i["key"]))
        jp = None
        if any(i.get("json") for i in items):        # 土增稅：再檢查一次謄本結構化(可能剛在主程式重做過)
            reason = tax_problem() or ("" if cur_json() else "找不到這個案件的謄本結構化資料")
            others = [i for i in items if not i.get("json")]
            if reason:
                if not others:
                    ui.message(win, "土增稅", f"⚠ {reason}\n\n土地增值稅要用這個案件的「謄本結構化」資料才算得出來，"
                                              "請先到【謄本與土增稅】做謄本結構化，再回來執行。", kind="warn")
                    return
                if not ui.message(win, "土增稅", f"⚠ {reason}\n\n土地增值稅要用這個案件的「謄本結構化」資料才算得出來。\n"
                                                 "要先略過土增稅，其他項目照跑嗎？", kind="ask",
                                  ok="略過土增稅，其他照跑", cancel="取消"):
                    return
                for i in items:
                    if i.get("json"):
                        rows[i["key"]]["var"].set(False)
                        rows[i["key"]]["status"].configure(text="✗ 結構化不是這個案件", fg=RED)
                items = others
            else:
                jp = cur_json()
                say(f"[連續模式] 土增稅使用結構化資料：{os.path.basename(jp)}")
        state.update(done=0, total=len(items), results=[])
        prog.pack(fill=tk.X, side=tk.BOTTOM, before=foot)
        bar_all.configure(maximum=len(items), value=0)
        r = Runner(case, items, show_var.get(), deal_var.get(), q, json_path=jp)
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
                    say(f"[連續模式] 第 {n}/{total} 項：{rows[key]['item']['name']} 開始")
                    if inter:
                        cont_btn.pack(anchor="w", pady=(6, 0))
                        line_lbl.configure(text="已開啟瀏覽器：請完成真人驗證，網頁出現查詢結果後按「繼續」")
                        say("  🖐 需要你操作：在瀏覽器完成真人驗證，網頁出現查詢結果後，按右下角提示框的「我完成驗證了，繼續」")
                        show_float(rows[key]['item']['name'])
                elif kind == "tick":
                    _, key, sec = m
                    if key == state["cur"]:
                        t = prog_lbl.cget("text").split("　")[0]
                        prog_lbl.configure(text=f"{t}　已 {sec} 秒")
                elif kind == "line":
                    line_lbl.configure(text="→ " + m[2][:120])
                    say("  " + m[2])
                    if float_box[0] is not None:
                        try:
                            _t = m[2].strip().strip("═=─-—").strip()      # 分隔線不顯示(會撐高提示框、擠掉按鈕)
                            if _t:
                                float_box[0]._line.configure(text="→ " + (_t[:44] + "…" if len(_t) > 45 else _t))
                        except Exception:
                            pass
                elif kind == "end":
                    _, key, st, msg, sec = m
                    bar_item.stop()
                    cont_btn.pack_forget()
                    hide_float()
                    _txt = {"done": "✓ 完成", "nodata": "— 官方查無資料", "stopped": "■ 已停止"}.get(st, "✗ 失敗")
                    say(f"[連續模式] {rows[key]['item']['name']}：{_txt}（{sec} 秒）" + (f"　{msg}" if msg else ""))
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
                elif kind == "mline":
                    line_lbl.configure(text=m[1][:120])
                elif kind == "merged":
                    merged(*m[1:])
                elif kind == "alldone":
                    runner_box[0] = None
                    bar_item.stop()
                    ok = sum(1 for _, s, _m in state["results"] if s == "done")
                    bad = [(n_, m_) for n_, s, m_ in state["results"] if s not in ("done", "nodata")]
                    prog_lbl.configure(text=f"完成：{ok}/{state['total']} 項成功" + (f"，{len(bad)} 項失敗" if bad else ""))
                    line_lbl.configure(text="；".join(f"{n_}：{m_}" for n_, m_ in bad)[:300] if bad else
                                       "全部完成，檔案在案件資料夾；可按下方「合併列印」把附件接成一份")
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
        hide_float(restore=False)
        win.destroy()
    win.protocol("WM_DELETE_WINDOW", on_close)
    if old_mode:                                     # 謄本、結構化、土增稅、編輯物件資料這些要人操作的，照舊版連續模式跑
        ui.button(foot, "舊版連續模式（謄本、土增稅…）", lambda: (win.destroy(), old_mode()), "ghost").pack(side=tk.LEFT)
    def do_open_folder():
        if os.path.isdir(case["folder"]):
            open_folder(case["folder"])
        else:
            ui.message(win, "案件資料夾", f"還沒有案件資料夾：\n{case['folder']}")

    merging = [False]

    def do_merge():
        if merging[0]:
            return
        if runner_box[0]:
            ui.message(win, "合併列印", "附件還在抓，請等全部跑完再合併。")
            return
        if not os.path.isdir(case["folder"]):
            ui.message(win, "合併列印", "還沒有案件資料夾，沒有檔案可以合併。")
            return
        try:
            st = merge_status(case["folder"])
        except Exception:
            st = None
        if st:                                       # 合併過了：打開舊的還是重新合併
            name = os.path.basename(st["path"])
            lines = [f"這個案件已經合併過：\n{name}\n（{time.strftime('%m/%d %H:%M', time.localtime(st['mtime']))} 存檔"
                     + ("，目前開在 PDF 程式裡）" if st["locked"] else "）")]
            if st["changes"]:
                more = f"…等 {len(st['changes'])} 項" if len(st["changes"]) > 5 else ""
                lines.append("⚠ 合併之後附件有變動，建議重新合併：\n" +
                             "\n".join("• " + c for c in st["changes"][:5]) + more)
            elif not st["no_trans"]:
                lines.append("✓ 合併之後附件沒有變動")
            if st["no_trans"]:
                lines.append(f"⚠ 這份是舊版合併的，沒有包含 0.謄本／3.行情 的 {st['no_trans']} 份檔案。\n"
                             "要一起印，請按「重新合併」，再勾選要哪些。")
            redo = bool(st["changes"] or st["no_trans"])
            if st["edited"]:
                lines.append("✏ 這份合併檔之後在 PDF 程式裡改過、存過檔。\n"
                             "重新合併前會先把它備份成「合併列印_手動修改備份_…pdf」，你的修改不會不見。")
            ans = ui.choice(win, "合併列印", "\n\n".join(lines),
                            [("打開已合併的檔案", "ghost" if redo else "primary"),
                             ("重新合併", "primary" if redo else "ghost"), ("取消", "ghost")],
                            kind="warn" if redo else "ask", default=1 if redo else 0)
            if ans is None or ans == 2:
                return
            if ans == 0:
                r = open_or_focus(st["path"])
                say(f"[合併列印] {'切換到' if r == 'focused' else '打開'}已合併的檔案：{name}")
                if r == "failed":
                    ui.message(win, "合併列印", f"打不開：\n{st['path']}\n\n請確認電腦有裝 PDF 程式。", kind="warn")
                return
        try:
            picked = pick_merge_files(win, case["folder"])
        except Exception as e:
            ui.message(win, "合併列印", f"列出檔案失敗：{e}", kind="warn")
            return
        if not picked:
            return
        if st:
            name = os.path.basename(st["path"])
            while file_locked(st["path"]):           # 重新合併：舊檔開著會蓋不掉
                open_or_focus(st["path"])
                a2 = ui.choice(win, "合併列印",
                               f"舊的合併檔還開在 PDF 程式裡：\n{name}\n\n"
                               "請先在 PDF 程式把它關掉（已經幫你切換過去了），再按「重試」。\n"
                               "若要保留舊的那份開著，可以選「存成新檔名」。",
                               [("重試", "primary"), ("存成新檔名", "ghost"), ("取消", "ghost")], kind="warn")
                if a2 is None or a2 == 2:
                    return
                if a2 == 1:
                    break
            if st["edited"] and not file_locked(st["path"]):
                bak = os.path.join(os.path.dirname(st["path"]),
                                   f"合併列印_手動修改備份_{time.strftime('%Y%m%d_%H%M%S', time.localtime(st['mtime']))}.pdf")
                try:
                    os.replace(st["path"], bak)
                    say(f"[合併列印] 你改過的合併檔已備份：{os.path.basename(bak)}")
                except OSError as e:
                    ui.message(win, "合併列印", f"備份你改過的合併檔失敗，先不重新合併：\n{e}", kind="warn")
                    return
        merging[0] = True
        merge_btn.itemconfigure(merge_btn._text, text="合併中…")
        prog.pack(fill=tk.X, side=tk.BOTTOM, before=foot)
        prog_lbl.configure(text="📑 合併列印：依編號把附件接成一份 PDF…")
        say("[合併列印] 開始：" + case["title"])
        exe_dir, internal, py, env = _paths()
        lst = os.path.join(case["folder"], "4.其他相關", "合併列印清單.json")
        try:
            os.makedirs(os.path.dirname(lst), exist_ok=True)
            with open(lst, "w", encoding="utf-8") as f:
                json.dump({"files": picked[0], "unpicked": picked[1]}, f, ensure_ascii=False, indent=1)
        except Exception as e:
            merged(None, [], f"寫入合併清單失敗：{e}")
            return
        args = [py, "-u", os.path.join(internal, "case_merge.py"), case["folder"], "--title", case["title"],
                "--list", lst]
        if duplex_var.get():
            args.append("--pad-blank")

        def work():
            out, warns, err = None, [], ""
            try:
                p = subprocess.Popen(args, cwd=internal, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                     creationflags=NO_WINDOW)
                for b in iter(p.stdout.readline, b""):
                    t = b.decode("utf-8", "replace").strip()
                    if t.startswith("OK:"):
                        out = t[3:]
                    elif t.startswith("FAIL:"):
                        err = t[5:]
                    elif t.startswith("WARN:"):
                        warns.append(t[5:])
                    elif t:
                        q.put(("mline", t))
                p.wait(timeout=600)
            except Exception as e:
                err = str(e)
            q.put(("merged", out, warns, err))
        threading.Thread(target=work, daemon=True).start()

    def merged(out, warns, err):
        merging[0] = False
        merge_btn.itemconfigure(merge_btn._text, text="📑 合併列印")
        if not out:
            prog_lbl.configure(text="✗ 合併失敗：" + (err or "原因不明"))
            say("[合併列印] 失敗：" + (err or "原因不明"))
            ui.message(win, "合併列印", "合併失敗：" + (err or "原因不明"), kind="warn")
            return
        name = os.path.basename(out)
        prog_lbl.configure(text=f"✓ 已合併：1.基本資料\\{name}")
        line_lbl.configure(text="已用你的 PDF 程式打開，看過沒問題就可以列印")
        say(f"[合併列印] 完成：{out}")
        for w_ in warns:
            say("  ⚠ " + w_)
        open_or_focus(out)                           # 用預設的 PDF 程式(Acrobat 等)開來預覽
        msg = (f"已合併並存成：\n1.基本資料\\{name}\n\n"
               "已經用你的 PDF 程式打開，請先看一下內容：\n"
               "• 沒問題 → 直接在 PDF 程式按列印（可選「雙面列印」，每頁已統一成 A4）\n"
               "• 要調整 → 可在 PDF 程式裡刪頁、調順序後存檔；或補好附件再按一次「合併列印」")
        if re.search(r"_\d{8}_\d{6}\.pdf$", name):
            msg += ("\n\n（上一份合併檔還開在 PDF 程式裡，所以這次存成新檔名；"
                    "請把舊的那個視窗關掉，下次合併時會自動清掉舊檔）")
        if warns:
            msg += "\n\n⚠ 以下沒有合併進去：\n" + "\n".join("• " + w_ for w_ in warns)
        ui.message(win, "合併列印", msg, kind="warn" if warns else "info")

    def tint_btn(b, main):                          # 跟主程式右邊彩色按鈕同一套(modern_ui tint)：主色淡底＋深字＋邊框
        b.fill, b.hover = ui._mix(main, "#FFFFFF", 0.72), ui._mix(main, "#FFFFFF", 0.58)
        b.itemconfigure(b._shape, fill=b.fill, outline=ui._mix(main, "#FFFFFF", 0.38), width=1)
        b.itemconfigure(b._text, fill=ui._mix(main, "#000000", 0.52))
        return b
    BF = (FAM, 14, "bold")
    tint_btn(ui.button(foot, "📂 開啟案件資料夾", do_open_folder, "ghost", font=BF), "#9A6B3F").pack(   # 棕(資料夾)
        side=tk.LEFT, padx=(0, 8))
    merge_btn = tint_btn(ui.button(foot, "📑 合併列印", do_merge, "ghost", font=BF), "#0F8F84")         # 藍綠
    merge_btn.pack(side=tk.LEFT)
    ui.button(btns, "關閉", on_close, "ghost").pack(side=tk.LEFT, padx=(0, 8))
    ui.button(btns, "■ 停止", stop, "ghost").pack(side=tk.LEFT, padx=(0, 8))
    run_btn = ui.button(btns, "▶ 開始", lambda: start(), "primary")
    win._start = start                              # 給測試/其他程式直接開始指定項目
    run_btn.pack(side=tk.LEFT)
    win.after(300, pump)
    return win
