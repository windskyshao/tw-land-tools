"""案件「匯出」的紀錄與兩邊對齊(使用者 11-11)。

工作資料夾＝程式資料夾\\<區段-地號>；匯出的資料夾＝物件編輯器「匯出」複製出去的那份
(例：0成交區\\115年\\115.11.11鼓山區龍華段-2165,門牌)。匯出不等於歸檔：案件可能還在進行，
同事也可能在匯出的資料夾加檔案 → 兩邊都當「還在用」，比對誰比較新，只補缺的、留新的，絕不刪檔。

紀錄存在每件自己的 4.其他相關\\匯出紀錄.json(資料夾搬家、換電腦都跟著走)；
舊版只在 config.json 記「最後匯出的那一件」，找不到紀錄時才拿來用。"""
import json
import os
import re
import shutil
import time
from datetime import datetime

REC_NAME = "匯出紀錄.json"
_SKIP = {"thumbs.db", "desktop.ini", ".dropbox", REC_NAME.lower()}
_cache = {}


def _rec_path(folder):
    return os.path.join(folder, "4.其他相關", REC_NAME)


def load(case_dir):
    try:
        with open(_rec_path(case_dir), encoding="utf-8") as f:
            r = json.load(f)
        return r if isinstance(r, dict) else {}
    except Exception:
        return {}


def _save(folder, rec):
    p = _rec_path(folder)
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(rec, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def record(case_dir, dest, fmt):
    """匯出完成後記下來；資料夾形式的匯出那邊也寫一份(從那邊開也認得是哪一件)"""
    rec = load(case_dir)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    rec["name"] = os.path.basename(os.path.normpath(case_dir))
    rec["latest"] = {"path": dest, "format": fmt, "date": now}
    hist = rec.get("history") or []
    hist.append({"path": dest, "format": fmt, "date": now})
    rec["history"] = hist[-30:]
    _save(case_dir, rec)
    if fmt in ("folder", "both") and os.path.isdir(dest):
        _save(dest, rec)
    _cache.clear()


def _dated(name, item):
    """item 是「民國日期＋區段-地號」開頭的匯出名稱 → 回可比較的日期數字，不是回 0"""
    m = re.match(r"^(\d{2,3})\.(\d{1,2})\.(\d{1,2})" + re.escape(name) + r"(?:,|_|\.zip$|$)", item)
    return int(m.group(1)) * 10000 + int(m.group(2)) * 100 + int(m.group(3)) if m else 0


def _newest_in(parent, name):
    """在 parent 找這一件最新日期的匯出(資料夾優先於 zip)"""
    best = None
    try:
        for it in os.listdir(parent):
            d = _dated(name, it)
            p = os.path.join(parent, it)
            if d and (os.path.isdir(p) or it.lower().endswith(".zip")):
                key = (d, os.path.isdir(p), os.path.getmtime(p))
                if not best or key > best[0]:
                    best = (key, p)
    except OSError:
        return None
    return best[1] if best else None


def find(case_dir):
    """這一件匯出到哪裡 → dict(path, format, date, moved) 或 None。
    紀錄的路徑不見了(改了新日期、被刪)→ 在同一個上層資料夾找同一件最新的那份"""
    name = os.path.basename(os.path.normpath(case_dir))
    rec = load(case_dir)
    la = rec.get("latest") or {}
    if not la:                                       # 舊版：config.json 只記最後一件
        try:
            with open(os.path.join(os.path.dirname(case_dir), "config.json"), encoding="utf-8") as f:
                es = json.load(f).get("export_settings") or {}
            if os.path.normcase(os.path.abspath(es.get("current_folder") or "")) == \
                    os.path.normcase(os.path.abspath(case_dir)) and es.get("last_archive_name"):
                fmt = es.get("last_export_format") or "zip"
                p = os.path.join(es.get("last_export_path") or "", es["last_archive_name"])
                la = {"path": p if fmt in ("folder", "both") else p + ".zip", "format": fmt,
                      "date": es.get("last_export_date") or ""}
        except Exception:
            pass
    if not la.get("path"):
        return None
    p = la["path"]
    if os.path.exists(p):
        return dict(la, moved=False)
    alt = _newest_in(os.path.dirname(p), name)
    if alt:
        return dict(la, path=alt, format="folder" if os.path.isdir(alt) else "zip", moved=True)
    return dict(la, missing=True)


def _tree(root):
    out = {}
    for d, dirs, files in os.walk(root):
        dirs[:] = [x for x in dirs if x.lower() not in _SKIP]
        for fn in files:
            if fn.startswith("~$") or fn.lower() in _SKIP:
                continue
            p = os.path.join(d, fn)
            try:
                out[os.path.relpath(p, root)] = os.path.getmtime(p)
            except OSError:
                pass
    return out


def compare(work, exp, cache_sec=0):
    """回 (push, pull)：push＝工作資料夾比較新或匯出那邊沒有的檔案；pull＝反過來(例：同事加的照片)"""
    key = (os.path.normcase(work), os.path.normcase(exp))
    hit = _cache.get(key)
    if cache_sec and hit and time.time() - hit[0] < cache_sec:
        return hit[1]
    W, E = _tree(work), _tree(exp)
    push = sorted(r for r, t in W.items() if r not in E or t > E[r] + 2)
    pull = sorted(r for r, t in E.items() if r not in W or t > W[r] + 2)
    _cache[key] = (time.time(), (push, pull))
    return push, pull


def copy_files(src, dst, rels):
    """把 rels(相對路徑)從 src 複製到 dst(保留修改時間)；回 (成功數, 錯誤清單)"""
    n, errs = 0, []
    for r in rels:
        try:
            t = os.path.join(dst, r)
            os.makedirs(os.path.dirname(t), exist_ok=True)
            shutil.copy2(os.path.join(src, r), t)
            n += 1
        except Exception as e:
            errs.append(f"{r}：{e}")
    _cache.clear()
    return n, errs


def align(work, exp):
    """兩邊對齊：各自補缺的、留比較新的，不刪檔。回 (拿回來數, 送過去數, 錯誤清單)"""
    push, pull = compare(work, exp)
    a, e1 = copy_files(exp, work, pull)
    b, e2 = copy_files(work, exp, push)
    return a, b, e1 + e2


def roc_date(s):
    """'2026-11-11 15:00' → '115.11.11'"""
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(s or ""))
    return f"{int(m.group(1)) - 1911}.{m.group(2)}.{m.group(3)}" if m else ""
