"""地籍資料查詢系統 ↔ 阿生地圖 Google 登入(裝置配對)。
按「Google 登入」→ 顯示確認碼、開瀏覽器 → 使用者在網頁按「允許」→ 這裡收到登入憑證。
憑證用 Windows DPAPI 加密存在 %APPDATA%\\地籍資料查詢系統\\landmap_login.bin(只有這台電腦的這個 Windows 帳號解得開)，
不放 config.json、不會被打包。"""
import os
import json
import ctypes
import ctypes.wintypes as wt
import threading
import webbrowser
import urllib.request
import urllib.error
import tkinter as tk

import landmap_api as LA

_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "地籍資料查詢系統")
_FILE = os.path.join(_DIR, "landmap_login.bin")
FAM = "Microsoft JhengHei"


# ───────────── DPAPI 加密存檔 ─────────────
class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(data, protect):
    src = _Blob(len(data), ctypes.cast(ctypes.create_string_buffer(data, len(data)), ctypes.POINTER(ctypes.c_char)))
    out = _Blob()
    fn = ctypes.windll.crypt32.CryptProtectData if protect else ctypes.windll.crypt32.CryptUnprotectData
    if not fn(ctypes.byref(src), None, None, None, None, 0, ctypes.byref(out)):
        raise OSError("DPAPI 失敗")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out.pbData)


def _save(d):
    os.makedirs(_DIR, exist_ok=True)
    with open(_FILE, "wb") as f:
        f.write(_dpapi(json.dumps(d, ensure_ascii=False).encode("utf-8"), True))


def _load():
    try:
        with open(_FILE, "rb") as f:
            return json.loads(_dpapi(f.read(), False).decode("utf-8"))
    except Exception:
        return None


# ───────────── 呼叫阿生地圖 ─────────────
def _req(path, token=None, method="GET", timeout=20, body=None):
    h = {"User-Agent": LA._UA, "Accept": "application/json"}
    if token:
        h["Cookie"] = f"session={token}"
    data = None
    if method == "POST":
        data = json.dumps(body or {}).encode("utf-8")
        h["Content-Type"] = "application/json"
    req = urllib.request.Request(LA.BASE + path, headers=h, method=method, data=data)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            return json.loads(e.read().decode("utf-8"))
        except Exception:
            return {"status": "ERROR", "msg": f"HTTP {e.code}"}


def token():
    d = _load()
    return (d or {}).get("token")


def call(path):
    """帶登入憑證呼叫阿生地圖 API(給之後的周邊行情、行情推估用)"""
    return _req(path, token())


def account(refresh=False):
    """目前登入者 dict(name, email, tier, avm) 或 None。refresh=True 會向主機確認(憑證被登出就清掉)"""
    d = _load()
    if not d or not d.get("token"):
        return None
    if refresh:
        try:
            j = _req("/api/auth/me", d["token"])
        except Exception:
            return d.get("user")                  # 連不上：先用上次的資料
        u = j.get("user")
        if not u:
            logout(local_only=True)
            return None
        tier = "member" if u.get("is_member") else ("supporter" if u.get("is_supporter") else "user")
        d["user"] = dict(name=u.get("name") or "", email=u.get("email") or "", tier=tier, avm=bool(u.get("avm")),
                         raw_tier=u.get("tier") or "")
        _save(d)
    return d.get("user")


def logout(local_only=False):
    t = token()
    if t and not local_only:
        try:
            _req("/api/auth/logout", t)
        except Exception:
            pass
    try:
        os.remove(_FILE)
    except OSError:
        pass


# ───────────── 登入視窗 ─────────────
_active = {"tw": None}


def _safe_after(w, fn):
    """背景執行緒回主畫面；主視窗已關就算了"""
    try:
        w.after(0, fn)
    except Exception:
        pass


def login(parent, on_done=None):
    """開確認碼小視窗＋瀏覽器；背景輪詢。成功 on_done(account)，取消/逾時 on_done(None)"""
    try:   # 已經有一個登入視窗開著：拉到前面，不重開(重開會產生兩組確認碼)
        if _active["tw"] is not None and _active["tw"].winfo_exists():
            _active["tw"].deiconify()
            _active["tw"].lift()
            return _active["tw"]
    except tk.TclError:
        pass
    try:
        j = _req("/api/auth/desktop/start")
    except Exception as e:
        j = {"status": "NET", "msg": str(e)}
    if j.get("status") != "OK":
        from tkinter import messagebox
        if j.get("status") == "NET":
            m = ("連不上阿生地圖。\n\n請確認網路正常；如果是公司網路，可能擋了 map.windsky-sky.com，"
                 "可以先換手機熱點試試。\n\n（" + str(j.get("msg", ""))[:120] + "）")
        else:
            m = "阿生地圖暫時無法登入，請過幾分鐘再試。\n" + str(j.get("msg", ""))
        messagebox.showerror("Google 登入", m, parent=parent)
        return
    code, sec = j["code"], j["secret"]
    url = LA.BASE + j["url"]
    webbrowser.open(url)

    tw = tk.Toplevel(parent)
    tw.title("Google 登入阿生地圖")
    tw.configure(bg="#FFFFFF")
    tw.attributes("-topmost", True)
    tw.resizable(False, False)
    tk.Label(tw, text="🐱", font=("Segoe UI Emoji", 30), bg="#FFFFFF").pack(pady=(14, 0))
    tk.Label(tw, text="在剛打開的網頁輸入這組確認碼，再按「允許登入」", font=(FAM, 15, "bold"), bg="#FFFFFF",
             fg="#1C1C1E").pack(padx=28)
    tk.Label(tw, text="確認碼：", font=(FAM, 12), bg="#FFFFFF", fg="#3A3A3C").pack(pady=(10, 0))
    tk.Label(tw, text=code, font=(FAM, 30, "bold"), bg="#FFFFFF", fg="#B35A10").pack()
    st = tk.Label(tw, text="等你按允許…", font=(FAM, 12), bg="#FFFFFF", fg="#3A3A3C")
    st.pack(pady=(4, 6))
    row = tk.Frame(tw, bg="#FFFFFF")
    row.pack(pady=(0, 16))
    state = {"stop": False}

    def cancel():
        state["stop"] = True
        _active["tw"] = None
        tw.destroy()
        on_done and on_done(None)

    tk.Button(row, text="網頁沒打開？再開一次", font=(FAM, 11), command=lambda: webbrowser.open(url)).pack(side=tk.LEFT, padx=4)
    tk.Button(row, text="取消", font=(FAM, 11), command=cancel).pack(side=tk.LEFT, padx=4)
    tw.protocol("WM_DELETE_WINDOW", cancel)
    tw.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - tw.winfo_reqwidth()) // 2
    y = parent.winfo_rooty() + 120
    tw.geometry(f"+{max(0, x)}+{max(0, y)}")
    _active["tw"] = tw

    def status(t, color="#3A3A3C"):
        def _f():
            try:
                st.config(text=t, fg=color)
            except tk.TclError:
                pass
        _safe_after(tw, _f)

    def worker():
        import time
        t0 = time.time()
        t_end = t0 + j.get("ttl", 600)
        bad = 0
        while not state["stop"] and time.time() < t_end:
            time.sleep(2)
            try:
                r = _req("/api/auth/desktop/poll", method="POST", body={"secret": sec})   # 密鑰放 body，不進網址
            except Exception:
                bad += 1
                if bad >= 3:
                    status("網路不穩，持續重試中…（不用關掉，連上就會自動完成）", "#B71C1C")
                continue
            if bad >= 3:
                status("等你按允許…")
            bad = 0
            if r.get("status") == "OK":
                try:
                    _save({"token": r["token"], "user": r.get("user") or {}})
                except Exception as e:
                    _safe_after(parent, lambda: finish(None, f"登入成功，但這台電腦存不了登入資料：\n{e}\n\n"
                                                             "請確認防毒軟體沒有擋，再按一次「Google 登入」"))
                    return
                acct = account(refresh=True) or r.get("user") or {"name": "", "email": ""}
                _safe_after(parent, lambda: finish(acct))
                return
            if r.get("status") == "EXPIRED":
                break
            if time.time() - t0 > 120:   # 等很久了：提醒常見卡住的原因
                status("還沒收到允許。網頁上要先用 Google 登入，再輸入上面的確認碼按「允許登入」", "#7A3E00")
        if not state["stop"]:
            _safe_after(parent, lambda: finish(None, "確認碼過期或失效了（10 分鐘內要按允許），"
                                                     "請再按一次「Google 登入」，會給你新的確認碼"))

    def finish(acct, err=None):
        if state["stop"]:
            return
        state["stop"] = True
        _active["tw"] = None
        try:
            tw.destroy()
        except tk.TclError:
            pass
        if err:
            from tkinter import messagebox
            messagebox.showinfo("Google 登入", err, parent=parent)
        on_done and on_done(acct)

    threading.Thread(target=worker, daemon=True).start()
    return tw
