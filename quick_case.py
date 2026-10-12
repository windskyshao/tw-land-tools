# -*- coding: utf-8 -*-
"""萬用查詢欄的結果視窗：地址/地號 → 地號、建號、社區名 → 寫入 data.json(新案件 / 加入目前案件)
由主程式在「沒有子程式執行時」按 Enter 觸發。查詢走 landmap_api(阿生地圖)，在背景執行緒跑。
樣式沿用 deal_ui，但不呼叫 deal_ui.apply()(那會改到主視窗的全域樣式)。"""
import json
import math
import re
import os
import sys
import shutil
import threading
import tkinter as tk
import webbrowser
from tkinter import ttk

import deal_ui as ui
import landmap_api as L

INK = "#0F172A"        # 內容文字(加深)
INK2 = "#1E293B"       # 說明文字(加深，不用淡灰)

def _asset(name):
    """打包後在 _internal(sys._MEIPASS)，開發時在本檔旁邊"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


CAT_ORANGE = "#E8730C"


def _cat_photo(master, size, ring=6):
    """阿生圓形頭像(外圈富住通橘)；PIL 不可用時回 None"""
    im = _cat_pil(size, ring)
    if im is None:
        return None
    from PIL import ImageTk
    return ImageTk.PhotoImage(im, master=master)


def _cat_pil(size, ring=6):
    try:
        from PIL import Image, ImageDraw
        inner = size - ring * 2
        im = Image.open(_asset("landmap_cat.png")).convert("RGBA").resize((inner, inner), Image.LANCZOS)
        mask = Image.new("L", (inner * 4, inner * 4), 0)
        ImageDraw.Draw(mask).ellipse((0, 0, inner * 4 - 1, inner * 4 - 1), fill=255)
        im.putalpha(mask.resize((inner, inner), Image.LANCZOS))
        big = Image.new("RGBA", (size * 4, size * 4), (0, 0, 0, 0))
        ImageDraw.Draw(big).ellipse((0, 0, size * 4 - 1, size * 4 - 1), fill=(242, 140, 40, 255))
        out = big.resize((size, size), Image.LANCZOS)
        out.alpha_composite(im, (ring, ring))
        return out
    except Exception:
        return None


def _wobble_frames(master, size):
    """仿阿生地圖 .catwork：以 (50%, 88%) 為支點左右擺 ±6°；另備 .catdone 開心彈一下的縮放影格"""
    base = _cat_pil(size)
    if base is None:
        return None, None
    from PIL import Image, ImageTk
    pad = int(size * 0.2)
    canvas_sz = size + pad * 2
    piv = (canvas_sz / 2, pad + size * 0.88)
    work = {}
    for k in range(-12, 13):                          # 每 0.5° 一張，擺動更順
        deg = k / 2
        c = Image.new("RGBA", (canvas_sz, canvas_sz), (0, 0, 0, 0))
        c.alpha_composite(base, (pad, pad))
        work[k] = ImageTk.PhotoImage(c.rotate(-deg, resample=Image.BICUBIC, center=piv), master=master)
    done = []
    for sc, deg in ((0.55, -8), (0.75, -4), (0.95, 0), (1.12, 3), (1.2, 4), (1.1, 2), (1.03, 1), (1.0, 0)):
        n = max(8, int(size * sc))
        im = base.resize((n, n), Image.LANCZOS)
        c = Image.new("RGBA", (canvas_sz, canvas_sz), (0, 0, 0, 0))
        c.alpha_composite(im, ((canvas_sz - n) // 2, (canvas_sz - n) // 2))
        done.append(ImageTk.PhotoImage(c.rotate(-deg, resample=Image.BICUBIC), master=master))
    return work, done


def _paw(cv, x, y, ang, s, color):
    """貓腳印：微三角的大肉球＋四顆弧形排列的趾頭，整個朝 ang(弧度)方向旋轉"""
    import math as _m
    ca, sa = _m.cos(ang), _m.sin(ang)

    def pt(fw, side):           # fw=往前距離, side=往右距離 → 畫面座標
        return x + ca * fw - sa * side, y + sa * fw + ca * side
    # 肉球：前窄後寬的圓潤形(平滑多邊形)，跟著方向轉
    pad = []
    for k in range(24):
        th = 2 * _m.pi * k / 24
        fw = _m.cos(th) * s * 0.30
        side = _m.sin(th) * s * (0.40 - 0.08 * _m.cos(th))   # 前方(cos>0)較窄
        pad += pt(fw - s * 0.18, side)
    cv.create_polygon(pad, smooth=True, fill=color, outline="")
    # 四顆趾頭：弧形排在肉球前方
    for deg, dist, r in ((-62, 0.50, 0.15), (-22, 0.60, 0.165), (22, 0.60, 0.165), (62, 0.50, 0.15)):
        th = _m.radians(deg)
        tx, ty = pt(_m.cos(th) * s * dist + s * 0.05, _m.sin(th) * s * dist)
        cv.create_oval(tx - s * r, ty - s * r * 1.15, tx + s * r, ty + s * r * 1.15, fill=color, outline="")


STEPS = {"address": ["找地址位置", "確認地號", "比對建號", "查社區"],
         "lot": ["確認地段", "找出位置", "列出建物", "查社區"]}
QUIPS = {0: "阿生出發囉，正在翻地圖找位置…", 1: "阿生找到地方了，正在確認地號…",
         2: "阿生在一戶一戶比對門牌…", 3: "阿生去問實價登錄，這裡是哪個社區…", 4: "阿生查好了！整理結果中…"}

TIPS = [
    ("🏢", "地址加上樓層（例：236之2號12樓），就能直接查到那一戶的建號"),
    ("🗺️", "用地號查（例：前鎮區竹北段一小段2096），可以看到這塊地上所有門牌"),
    ("🏘️", "查到社區名稱後，內政部實價會自動改用社區名查詢，不用再手動找大樓名"),
    ("📁", "建立案件後，地籍便民、使用分區等各查詢功能都會自動帶入地號"),
    ("➕", "同一個案子有好幾筆地號？選「加入目前案件」就能一筆一筆加進去"),
    ("📋", "也可以一次輸入多筆，用「、」分開：竹北段一小段2096、2097、2100-1"),
    ("🐾", "縣市可以省略，沒寫的話預設是高雄市"),
    ("🔢", "「236號之2」「236之2號」「240-1號」寫法都可以，全形數字也行"),
]

_RETURN = {}          # 從萬用查詢面板送出時記下(分頁＋已選的值)，給結果視窗「回上一步重新查」用

CASE_SUBDIRS = ["0.謄本", "1.基本資料", "2.照片", "3.行情", "4.其他相關", "5.委託", "6.成交"]


def _case_folder_name(e):
    return f"{e['area']}{e['section']}-{e['lot_number']}"


def _load_data(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, list) else []
    except Exception:
        return []


def _merge_builds(a, b):
    """兩份 builds 清單合併(依建號去重，保留先後順序)"""
    out, seen = [], set()
    for x in (a or []) + (b or []):
        if x.get("build_no") and x["build_no"] not in seen:
            seen.add(x["build_no"])
            out.append(x)
    return out


def _upsert(items, entry):
    """同地號(縣市+行政區+段+地號)更新該列，否則加到最後。
    build_no/door 維持舊程式讀的「第一戶」；builds 清單記這個地號上要用的全部戶"""
    key = (entry["city"], entry["area"], entry["section"], entry["lot_number"])
    for it in items:
        if (it.get("city"), it.get("area"), it.get("section"), it.get("lot_number")) == key:
            builds = _merge_builds(it.get("builds"), entry.get("builds"))
            keep_first = it.get("build_no")
            it.update({k: v for k, v in entry.items() if v and k != "builds"})
            if keep_first:                                   # 已指定的第一戶不被後來加的蓋掉
                it["build_no"] = keep_first
                it["door"] = next((b["door"] for b in builds if b["build_no"] == keep_first), it.get("door", ""))
            if builds:
                it["builds"] = builds
            return
    items.append(entry)


def save_entries(entries, mode, data_json_path, base_dir):
    """mode='new'：data.json 只放這幾筆；'append'：加到現有清單(同地號則更新)。
    同時建好案件資料夾並備份 data.json 到 4.其他相關(與地籍便民的做法一致)。回傳案件資料夾路徑"""
    items = [] if mode == "new" else _load_data(data_json_path)
    for entry in entries:
        _upsert(items, entry)
    with open(data_json_path, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=4)
    folder = os.path.join(base_dir, _case_folder_name(items[0]))
    for sd in CASE_SUBDIRS:
        os.makedirs(os.path.join(folder, sd), exist_ok=True)
    try:
        shutil.copy2(data_json_path, os.path.join(folder, "4.其他相關", "data.json"))
    except Exception:
        pass
    return folder


def save_entry(entry, mode, data_json_path, base_dir):
    return save_entries([entry], mode, data_json_path, base_dir)


def entry_from_result(r, build=None):
    """查詢結果 → data.json 一列。build=選定的那一戶(可為 None)"""
    e = {"city": r["city"], "area": r["area"], "section": r["section"], "lot_number": r["lot_number"],
         "coordinates": f"{r['lat']:.5f},{r['lng']:.5f}"}
    if build:
        e["build_no"], e["door"] = build["build_no"], build["door"]
        e["builds"] = [{"build_no": build["build_no"], "door": build["door"]}]
    elif r.get("door"):
        e["door"] = r["door"]
    if r.get("community"):
        e["community"] = r["community"]
    return e


class QuickCaseWindow:
    def __init__(self, root, text, data_json_path, base_dir, on_saved=None, log=None, input_widget=None):
        self.root, self.text = root, text
        self.data_json_path, self.base_dir = data_json_path, base_dir
        self.on_saved, self.log = on_saved, (log or (lambda m: None))
        self.input_widget = input_widget
        self.back = dict(_RETURN) if _RETURN else None     # 從面板來的 → 回面板；否則 → 文字放回輸入欄
        _RETURN.clear()
        self.res, self.err = None, None
        self.queries = L.split_batch(text) or [text]
        self.batch = len(self.queries) > 1          # 一次輸入多筆(用、，；分開)
        self.cur_i = 0
        self.sel_build = tk.IntVar(master=root, value=0)
        self.extra_builds = []

        self.top = top = tk.Toplevel(root, bg=ui.BG)
        top.title("快速建立案件")
        top.transient(root)
        geo = self._side_geometry()
        if geo:
            W, H, X, Y = geo
        else:
            sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
            W, H = min(860, sw - 80), min(820, sh - 100)
            X, Y = (sw - W) // 2, max(10, (sh - H) // 2 - 30)
        top.geometry(f"{W}x{H}+{X}+{Y}")
        top.minsize(600, 480)
        self.W = min(W, 1100)                     # 文字換行寬度上限(太寬一行讀起來累)
        ui.header(top, "快速建立案件", f"查詢：{text}")
        self.foot = ui.footer(top)
        self.body, self.canvas = ui.scroll_body(top)
        self._loading()
        threading.Thread(target=self._work, daemon=True).start()
        self._poll()

    def _side_geometry(self):
        """放在主介面旁邊剩餘的空間(跟子程式開 Chrome 的位置一樣)。
        工作區讀 window_config.json(主程式依每台電腦螢幕算好的，實體像素)；主介面以「現在」的實際位置為準。
        兩側空間都不足 600px 或讀不到設定 → 回 None(改用置中)"""
        try:
            with open(os.path.join(self.base_dir, "window_config.json"), "r", encoding="utf-8") as f:
                wa = json.load(f)["work_area"]
            self.root.update_idletasks()
            mx, my = self.root.winfo_rootx(), self.root.winfo_rooty()
            mw = self.root.winfo_width()
            fx = self.root.winfo_x()                       # 含外框的左上角
            frame_l = mx - fx                              # 左右外框寬
            gap = 10
            left = fx - wa["left"] - gap                   # 主介面左邊可用寬(含外框)
            right = wa["right"] - (fx + mw + frame_l * 2) - gap
            space, on_left = (left, True) if left >= right else (right, False)
            if space < 600:
                return None
            W = space - frame_l * 2                        # geometry 寬度不含外框
            fy = self.root.winfo_y()
            title_h = my - fy                              # 標題列高度
            # 上下緣對齊主介面(同樣的標題列與外框 → 內容高度相同)；超出工作區才往內縮
            Y = max(fy, wa["top"])
            H = self.root.winfo_height() - (Y - fy)
            H = min(H, wa["bottom"] - Y - title_h)         # 下外框多半是透明陰影，不必扣
            if H < 480:
                H = (wa["bottom"] - wa["top"]) - title_h - frame_l
                Y = wa["top"]
            X = wa["left"] if on_left else fx + mw + frame_l * 2 + gap
            return int(W), int(H), int(X), int(Y)
        except Exception:
            return None

    # ---------- 查詢 ----------
    def _work(self):
        if self.batch:
            self._work_batch()
            return
        try:
            self.res = L.resolve(self.text, progress=lambda n: setattr(self, "step", n))
        except L.LandmapError as e:
            self.err = str(e)
        except Exception as e:
            self.err = f"查詢時發生錯誤：{e}"

    def _work_batch(self):
        """一筆一筆查(不並行：照順序沿用前一筆的行政區，也不會一次打太多請求)"""
        out, prev = [], None
        for i, q in enumerate(self.queries):
            self.cur_i, self.step = i, 0
            try:
                r = L.resolve_with_context(q, prev, progress=lambda n: setattr(self, "step", n))
                out.append({"q": q, "res": r, "err": ""})
                prev = r
            except L.LandmapError as e:
                out.append({"q": q, "res": None, "err": str(e)})
            except Exception as e:
                out.append({"q": q, "res": None, "err": f"查詢時發生錯誤：{e}"})
        self.step = 4
        self.res = out

    def _poll(self):
        if self.res is None and self.err is None:
            try:
                self.top.after(200, self._poll)
            except tk.TclError:
                pass
            return
        if self.res is not None and not getattr(self, "_finish_shown", False):
            self._finish_shown = True               # 先讓四步全打勾、阿生說「查好了」再切換結果
            self.step = 4
            self.top.after(600, self._poll)
            return
        for w in self.body.winfo_children():
            w.destroy()
        if self.err:
            self._error()
        elif self.batch:
            self._show_batch()
        else:
            self._show()

    def _loading(self):
        """等待畫面：阿生(橘貓)浮動＋地圖圖釘繞圈＋四步驟進度條(每完成一步打勾)"""
        self.step = 0
        self.tick = 0
        kind = "lot" if L.looks_like_lot(L.norm_city(self.queries[0]).translate(L._FW).replace(" ", "")) else "address"
        self.step_names = STEPS[kind]
        self.top_spacer = tk.Frame(self.body, bg=ui.BG, height=1)
        self.top_spacer.pack(fill=tk.X)
        sub = (f"共 {len(self.queries)} 筆，一筆一筆查，每筆約 10 秒" if self.batch
               else "地址 → 地號 → 建號 → 社區，約 10 秒")
        c = ui.card(self.body, "阿生查詢中", sub, "blue_strong")
        bg = c.cget("bg")
        self.anim = tk.Canvas(c, height=230, bg=bg, highlightthickness=0)
        self.anim.pack(fill=tk.X)
        self.cat_img = _cat_photo(self.top, 140)
        self.cat_work, self.cat_done = _wobble_frames(self.top, 140)
        self.done_i = None
        self.caption = tk.Label(c, text=QUIPS[0], font=ui.FT, fg=ui.THEMES["blue_strong"][0], bg=bg)
        self.caption.pack(pady=(4, 10))
        # 步驟條
        self.stepper = tk.Canvas(c, height=70, bg=bg, highlightthickness=0)
        self.stepper.pack(fill=tk.X, pady=(0, 6))
        # 阿生小提醒(輪播)＋目前案件
        tc = ui.card(self.body, "💡 阿生小提醒", "", "orange_strong")
        self.tip_bg = tc.cget("bg")
        cur = _load_data(self.data_json_path)
        if cur and all(k in cur[0] for k in ("area", "section", "lot_number")):
            more = f"（共 {len(cur)} 筆）" if len(cur) > 1 else ""
            tk.Label(tc, text=f"目前案件：{_case_folder_name(cur[0])}{more}", font=ui.FB,
                     fg=ui.THEMES["orange_strong"][0], bg=self.tip_bg).pack(anchor="w", pady=(0, 6))
        row = tk.Frame(tc, bg=self.tip_bg)
        row.pack(fill=tk.X)
        self.tip_icon = tk.Canvas(row, width=38, height=38, bg=self.tip_bg, highlightthickness=0)
        self.tip_icon.pack(side=tk.LEFT, anchor="n")
        self.tip_lab = tk.Label(row, text="", font=(ui.FAMILY, 15), fg=INK, bg=self.tip_bg, justify="left",
                                anchor="nw", height=2)
        self.tip_lab.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
        # 依實際寬度換行(不然長句會被切掉)
        self.tip_lab.bind("<Configure>", lambda e: self.tip_lab.config(wraplength=max(200, e.width - 10)))
        import time as _time
        self.tip_t0 = _time.monotonic()
        self.tip_dots = tk.Canvas(tc, height=14, bg=self.tip_bg, highlightthickness=0)
        self.tip_dots.pack(anchor="w", pady=(6, 0))
        ui.button(self.foot, "取消", self.top.destroy, "ghost").pack(side=tk.RIGHT)
        self._add_back()
        self.top.after(80, self._center_loading)
        self._animate()

    def _center_loading(self):
        """把等待畫面上下置中：空白平均分到上下"""
        try:
            self.top.update_idletasks()
            content = sum(w.winfo_reqheight() for w in self.body.winfo_children() if w is not self.top_spacer)
            spare = self.canvas.winfo_height() - content - 30
            self.top_spacer.configure(height=max(1, spare // 2))
        except tk.TclError:
            pass

    def _draw_tip(self):
        import time as _time
        n = len(TIPS)
        el = _time.monotonic() - self.tip_t0                 # 用實際時間：每則 4 秒，前後各 0.4 秒淡入淡出
        idx = int(el // 4.0) % n
        pos = el % 4.0
        a = min(1.0, pos / 0.4, (4.0 - pos) / 0.4)
        icon, text = TIPS[idx]
        self.tip_icon.delete("all")                           # 每則前面一個橘色貓腳印(跟著淡入淡出)
        _paw(self.tip_icon, 19, 21, -math.pi / 2, 28, ui._mix(self.tip_bg, CAT_ORANGE, a))
        self.tip_lab.config(text=text, fg=ui._mix(self.tip_bg, INK, a))
        cv = self.tip_dots
        cv.delete("all")
        for i in range(n):                                    # 第幾則的小圓點
            x = 6 + i * 16
            on = i == idx
            cv.create_oval(x - (5 if on else 3), 7 - (5 if on else 3), x + (5 if on else 3), 7 + (5 if on else 3),
                           fill=ui.THEMES["orange_strong"][0] if on else "#E9B98A", outline="")

    def _animate(self):
        try:
            if not self.anim.winfo_exists():
                return
        except tk.TclError:
            return
        self.tick += 1
        t = self.tick
        cv = self.anim
        w = max(cv.winfo_width(), 600)
        cx, cy = w / 2, 110
        ph = (t * 0.05) % 1.0                         # 每 tick 50ms → 1 秒一個來回
        wob = -6 * math.cos(2 * math.pi * ph)          # -6° → +6° → -6°
        bob = -2 * (1 - math.cos(2 * math.pi * ph))    # 擺到中間時上抬 4px
        cv.delete("all")
        # 影子(貓浮高時變小變淡)
        sw = 46 - bob * 1.2
        cv.create_oval(cx - sw, 196 - 5, cx + sw, 196 + 5, fill="#B9CBE6", outline="")
        # 外圈虛線軌道
        R = 92
        cv.create_oval(cx - R, cy - R + bob, cx + R, cy + R + bob, outline="#9DBBE8", dash=(3, 6), width=2)
        # 小光點
        for k in range(3):
            ang = t / 9 + k * 2.1
            sx, sy = cx + math.cos(ang) * (R + 26), cy + math.sin(ang) * (R - 10) + bob
            r = 3 + (math.sin(t / 3 + k) + 1) * 1.5
            cv.create_oval(sx - r, sy - r, sx + r, sy + r, fill="#F6B26B", outline="")
        # 阿生：工作中左右搖晃；查好時開心彈一下
        if self.step >= 4 and self.cat_done:
            self.done_i = 0 if self.done_i is None else min(self.done_i + 1, len(self.cat_done) - 1)
            cv.create_image(cx, cy, image=self.cat_done[self.done_i])
        elif self.cat_work:
            cv.create_image(cx, cy + bob, image=self.cat_work[int(round(wob * 2))])
        elif self.cat_img:
            cv.create_image(cx, cy + bob, image=self.cat_img)
        else:
            cv.create_oval(cx - 64, cy - 64 + bob, cx + 64, cy + 64 + bob, fill="#F28C28", outline="")
            cv.create_text(cx, cy + bob, text="阿生", font=(ui.FAMILY, 22, "bold"), fill="white")
        # 阿生的腳印沿著軌道走：每 5 格踏出一步，左右腳交替，越早的腳印越淡
        step_i = t // 6
        trail = 6
        for k in range(trail):
            n = step_i - k
            ang = n * 0.55
            side = 11 if n % 2 else -11               # 左右腳(往軌道內外錯開)
            heading = ang + math.pi / 2                # 順時針前進方向
            px = cx + math.cos(ang) * (R + side)
            py = cy + math.sin(ang) * (R + side) + bob
            fade = k / trail
            col = ui._mix(CAT_ORANGE, "#DCEAFE", min(0.9, fade * 1.1))
            _paw(cv, px, py, heading, 22 - k * 1.2, col)
        self._draw_steps()
        dots = "．" * (t // 4 % 4)
        head = f"第 {self.cur_i + 1}/{len(self.queries)} 筆　" if self.batch and self.step < 4 else ""
        self.caption.config(text=head + QUIPS.get(self.step, QUIPS[4]).rstrip("…") + "…" + dots)
        self._draw_tip()
        self.top.after(50, self._animate)

    def _draw_steps(self):
        cv = self.stepper
        w = max(cv.winfo_width(), 600)
        n = len(self.step_names)
        x0, x1, y = 70, w - 70, 22
        gap = (x1 - x0) / (n - 1)
        done_c, act_c, pend_c = "#1B6B3A", ui.THEMES["blue_strong"][0], "#94A3B8"
        cv.delete("all")
        for i in range(n - 1):
            col = done_c if self.step > i else "#C7D2E0"
            cv.create_line(x0 + gap * i + 16, y, x0 + gap * (i + 1) - 16, y, fill=col, width=4, capstyle="round")
        pulse = (math.sin(self.tick / 3) + 1) / 2
        for i, name in enumerate(self.step_names):
            x = x0 + gap * i
            if self.step > i:                                     # 完成
                cv.create_oval(x - 15, y - 15, x + 15, y + 15, fill=done_c, outline="")
                cv.create_text(x, y, text="✓", font=(ui.FAMILY, 14, "bold"), fill="white")
                fg = done_c
            elif self.step == i:                                  # 進行中(呼吸光圈)
                r = 15 + pulse * 5
                cv.create_oval(x - r, y - r, x + r, y + r, outline=act_c, width=2)
                cv.create_oval(x - 13, y - 13, x + 13, y + 13, fill=act_c, outline="")
                cv.create_text(x, y, text=str(i + 1), font=(ui.FAMILY, 12, "bold"), fill="white")
                fg = act_c
            else:                                                 # 尚未
                cv.create_oval(x - 13, y - 13, x + 13, y + 13, fill="white", outline=pend_c, width=2)
                cv.create_text(x, y, text=str(i + 1), font=(ui.FAMILY, 12, "bold"), fill=pend_c)
                fg = "#475569"
            cv.create_text(x, y + 32, text=name, font=(ui.FAMILY, 13, "bold"), fill=fg)

    def _error(self):
        c = ui.card(self.body, "查不到", "", "orange_strong")
        tk.Label(c, text=self.err, font=ui.F, fg=ui.WARN_FG, bg=c.cget("bg"), justify="left",
                 wraplength=self.W - 120).pack(anchor="w", pady=4)
        tk.Label(c, text="可以這樣輸入：\n  ‧ 地址：高雄市前鎮區和平二路236之2號12樓（縣市可省略，預設高雄市）\n"
                         "  ‧ 地號：前鎮區竹北段一小段2096、楠梓區芎蕉段13-1\n"
                         "若阿生地圖暫時連不上，請照原本方式用「地籍便民系統」查詢。",
                 font=ui.FS, fg=INK2, bg=c.cget("bg"), justify="left").pack(anchor="w", pady=(8, 0))
        self._footer(can_save=False)

    # ---------- 結果 ----------
    def _kv(self, parent, r, k, v, bold=False):
        tk.Label(parent, text=k, font=ui.F, fg=INK2, bg=parent.cget("bg")).grid(row=r, column=0, sticky="e", padx=(0, 14), pady=3)
        tk.Label(parent, text=v or "—", font=ui.FB if bold else ui.F, fg=INK, bg=parent.cget("bg"),
                 justify="left", wraplength=self.W - 260).grid(row=r, column=1, sticky="w", pady=3)

    def _show(self):
        r = self.res
        main, tint, _edge = ui.THEMES["blue_strong"]
        c1 = ui.card(self.body, "📍 位置", "地號、座標已確認", "blue_strong")
        g = tk.Frame(c1, bg=c1.cget("bg"))
        g.pack(fill=tk.X)
        self._kv(g, 0, "縣市 / 行政區", f"{r['city']} {r['area']}")
        self._kv(g, 1, "地段", r["section"], True)
        self._kv(g, 2, "地號", r["lot_number"], True)
        self._kv(g, 3, "座標", f"{r['lat']:.5f}, {r['lng']:.5f}")
        btns = tk.Frame(c1, bg=c1.cget("bg"))
        btns.pack(anchor="w", pady=(8, 0))
        self.cat_small = _cat_photo(self.top, 28, ring=2)
        ui.button(btns, "在阿生地圖上看", lambda: webbrowser.open(r["map_url"]), "primary", CAT_ORANGE,
                  padx=14, pady=5, image=self.cat_small).pack(side=tk.LEFT)

        # 建物
        c2 = ui.card(self.body, "🏢 建物 / 建號", "電子、電傳謄本用", "green_strong")
        self.build_box = tk.Frame(c2, bg=c2.cget("bg"))
        self.build_box.pack(fill=tk.X)
        self._render_builds(r["builds"], r["build_note"])
        if r["kind"] == "lot" and r["door_groups"]:
            tk.Label(c2, text="此地號上的門牌：", font=ui.FB, fg=INK, bg=c2.cget("bg")).pack(anchor="w", pady=(8, 2))
            chips = tk.Frame(c2, bg=c2.cget("bg"))
            chips.pack(anchor="w")
            for i, grp in enumerate(r["door_groups"][:12]):
                tk.Label(chips, text=L.door_group_text(grp), font=ui.FS, fg=ui.THEMES["green_strong"][0], bg="white",
                         padx=10, pady=3, highlightthickness=1, highlightbackground=ui.THEMES["green_strong"][2]).grid(
                    row=i // 3, column=i % 3, sticky="w", padx=3, pady=3)
        # 補打門牌再查
        f = tk.Frame(c2, bg=c2.cget("bg"))
        f.pack(anchor="w", pady=(10, 0))
        tk.Label(f, text="要查某一戶：", font=ui.F, fg=INK, bg=c2.cget("bg")).pack(side=tk.LEFT)
        self.door_var = tk.StringVar(master=self.top, value="")
        e = ui.entry(f, self.door_var, 26)
        e.pack(side=tk.LEFT, ipady=4, padx=6)
        e.bind("<Return>", lambda ev: self._lookup_door())
        ui.button(f, "查建號", self._lookup_door, "primary", ui.THEMES["green_strong"][0], padx=16, pady=6).pack(side=tk.LEFT)
        tk.Label(c2, text="例：和平二路236之2號12樓（打門牌＋樓層）", font=ui.FS, fg=INK2, bg=c2.cget("bg")).pack(anchor="w", pady=(4, 0))

        # 社區
        c3 = ui.card(self.body, "🏘️ 社區", "內政部實價會用來查詢", "purple_strong")
        if r["community"]:
            tk.Label(c3, text=r["community"], font=(ui.FAMILY, 20, "bold"), fg=ui.THEMES["purple_strong"][0],
                     bg=c3.cget("bg")).pack(anchor="w")
            tk.Label(c3, text="依歷年實價登錄在這塊地上的成交紀錄比對。", font=ui.FS, fg=INK2, bg=c3.cget("bg")).pack(anchor="w")
        else:
            tk.Label(c3, text="查無社區名稱", font=ui.FB, fg=INK, bg=c3.cget("bg")).pack(anchor="w")
            tk.Label(c3, text="土地、透天，或這棟大樓還沒有成交紀錄；實價會照原本用地段查詢。", font=ui.FS,
                     fg=INK2, bg=c3.cget("bg")).pack(anchor="w")
        self._footer(can_save=True)

    def _render_builds(self, builds, note):
        for w in self.build_box.winfo_children():
            w.destroy()
        self.cur_builds = builds
        bg = self.build_box.cget("bg")
        teal, _t, edge = ui.THEMES["green_strong"]
        if not builds:
            tk.Label(self.build_box, text=note or "沒有建號資料", font=ui.F, fg=INK, bg=bg, justify="left",
                     wraplength=self.W - 120).pack(anchor="w")
            return
        self.sel_build.set(0)
        none_row = tk.Frame(self.build_box, bg=bg)
        none_row.pack(anchor="w")
        ui.big_radio(none_row, self.sel_build, -1, teal).pack(side=tk.LEFT, padx=(0, 8))
        nl = tk.Label(none_row, text="不指定建號（純土地或之後再查）", font=ui.FS, fg=INK2, bg=bg, cursor="hand2")
        nl.pack(side=tk.LEFT)
        nl.bind("<Button-1>", lambda e: self.sel_build.set(-1))
        rows = tk.Frame(self.build_box, bg=bg)
        rows.pack(fill=tk.X, pady=(2, 0))
        for i, b in enumerate(builds[:40]):
            att = "、".join(f"{a.get('use', '')}{a.get('m2', '')}㎡" for a in b.get("attached") or [])
            row = tk.Frame(rows, bg="white", highlightthickness=1, highlightbackground=edge, padx=10, pady=6)
            row.pack(fill=tk.X, pady=3)
            ui.big_radio(row, self.sel_build, i, teal).grid(row=0, column=0, rowspan=2, sticky="n", padx=(0, 10), pady=(2, 0))
            tk.Label(row, text=f"建號 {b['build_no']}", font=ui.FB, fg=teal, bg="white", anchor="w").grid(row=0, column=1, sticky="w")
            tk.Label(row, text=b["door"], font=ui.F, fg=INK, bg="white", anchor="w").grid(row=0, column=2, sticky="w", padx=(14, 0))
            info = "　·　".join(x for x in (b.get("floor", ""), f"主建物 {b['area_m2']}㎡" if b.get("area_m2") else "",
                                            f"附屬：{att}" if att else "") if x)
            tk.Label(row, text=info, font=ui.FS, fg=INK2, bg="white", anchor="w", justify="left",
                     wraplength=self.W - 200).grid(row=1, column=1, columnspan=2, sticky="w", pady=(2, 0))
            for w in (row,) + tuple(row.winfo_children()):
                w.bind("<Button-1>", lambda e, i=i: self.sel_build.set(i))
        if note:
            tk.Label(self.build_box, text="ⓘ " + note, font=ui.FS, fg=INK2, bg=bg, justify="left",
                     wraplength=self.W - 120).pack(anchor="w", pady=(4, 0))

    def _lookup_door(self):
        door = self.door_var.get().strip()
        if not door:
            return
        door = L._official_door(door.translate(L._FW))
        for w in self.build_box.winfo_children():
            w.destroy()
        tk.Label(self.build_box, text=f"查詢「{door}」中…", font=ui.F, fg=INK2, bg=self.build_box.cget("bg")).pack(anchor="w")
        out = {}

        def work():
            try:
                out["r"] = L.lookup_door_in_lot(self.res, door)
            except Exception as e:
                out["e"] = str(e)

        def poll():
            if not out:
                self.top.after(300, poll)
                return
            if "e" in out:
                self._render_builds([], f"查詢失敗：{out['e']}")
            else:
                builds, note = out["r"]
                if builds:
                    self.res["door"] = door
                self._render_builds(builds, note if builds else f"這個地號上找不到「{door}」，請確認門牌與樓層寫法")
        threading.Thread(target=work, daemon=True).start()
        poll()

    # ---------- 存檔 ----------
    def _add_back(self):
        """頁尾左側「‹ 回上一步重新查」：面板送出的回面板(保留已選的值)，輸入欄打的把文字放回輸入欄"""
        if not self.back and self.input_widget is None:
            return
        ui.button(self.foot, "‹  回上一步重新查", self._go_back, "chip", ui.THEMES["blue_strong"][0]).pack(side=tk.LEFT)

    def _go_back(self):
        try:
            self.top.destroy()
        except tk.TclError:
            pass
        b = self.back
        if b:
            open_panel(b["root"], b["inp"], b["flag_dir"], b["on_query"], b.get("tab"), restore=b.get("values"),
                       case_info=b.get("case_info"))
            return
        w = self.input_widget
        try:
            w.focus_set()
            w.delete(0, "end")
            w.insert(0, self.text)
            w.config(fg="black")
            w.icursor("end")
        except (tk.TclError, AttributeError):
            pass

    def _footer(self, can_save):
        for w in self.foot.winfo_children():
            w.destroy()
        cur = _load_data(self.data_json_path)
        if can_save:
            save = self._save_batch if self.batch else self._save
            ui.button(self.foot, "＋  建立新案件", lambda: save("new")).pack(side=tk.RIGHT)
            if cur:
                name = _case_folder_name(cur[0]) if all(k in cur[0] for k in ("area", "section", "lot_number")) else "目前案件"
                ui.button(self.foot, f"加入目前案件（{name}）", lambda: save("append"), "chip").pack(side=tk.RIGHT, padx=10)
        ui.button(self.foot, "取消" if can_save else "關閉", self.top.destroy, "ghost").pack(side=tk.RIGHT, padx=(0, 10))
        self._add_back()
        tk.Label(self.body, text="資料來源：阿生地圖（國土測繪中心、內政部實價登錄）", font=ui.FS, fg=INK2,
                 bg=ui.BG).pack(anchor="e", pady=(4, 8))

    def _save(self, mode):
        r = self.res
        i = self.sel_build.get()
        builds = getattr(self, "cur_builds", []) or []
        entry = entry_from_result(r, builds[i] if 0 <= i < len(builds) else None)
        msg = self._city_conflict([entry], mode)
        if msg:
            tk.Label(self.body, text=msg, font=ui.FB, fg=ui.BAD, bg=ui.BG, wraplength=self.W - 80,
                     justify="left").pack(anchor="w")
            return
        try:
            folder = save_entry(entry, mode, self.data_json_path, self.base_dir)
        except Exception as e:
            tk.Label(self.body, text=f"儲存失敗：{e}", font=ui.FB, fg=ui.BAD, bg=ui.BG).pack(anchor="w")
            return
        self.top.destroy()
        if self.on_saved:
            self.on_saved(entry, mode, folder)


    # ---------- 批次 ----------
    @staticmethod
    def _batch_build(r):
        """地址查到剛好一戶 → 自動帶那一戶；多戶或地號查詢 → 不指定(之後可單筆查)"""
        if r["kind"] == "address" and len(r["builds"]) == 1:
            return r["builds"][0]
        return None

    def _show_batch(self):
        items = self.res
        ok = [it for it in items if it["res"]]
        first_city = L.norm_city(ok[0]["res"]["city"]) if ok else ""
        c = ui.card(self.body, "📋 批次查詢結果", f"共 {len(items)} 筆，查到 {len(ok)} 筆", "blue_strong")
        bg = c.cget("bg")
        teal = ui.THEMES["green_strong"][0]
        # 提醒：跨地段/行政區、同地號會合併
        keys = [(it["res"]["area"], it["res"]["section"]) for it in ok]
        lots = [(it["res"]["area"], it["res"]["section"], it["res"]["lot_number"]) for it in ok]
        notes = []
        if len(set(keys)) > 1:
            notes.append("這批地號分屬不同行政區或地段，請確認是同一個案子。")
        if len(set(lots)) < len(lots):
            notes.append("有幾筆落在同一個地號上，存檔時會合併成一筆，建號都會記下來。")
        for n in notes:
            tk.Label(c, text="⚠ " + n, font=ui.FB, fg=ui.THEMES["orange_strong"][0], bg=bg, justify="left",
                     wraplength=self.W - 120).pack(anchor="w", pady=(0, 4))
        tk.Label(c, text="勾選要加入案件的項目（照輸入順序存，第一筆當案件名稱）：", font=ui.F, fg=INK,
                 bg=bg).pack(anchor="w", pady=(2, 6))
        self.batch_vars = []
        for i, it in enumerate(items, 1):
            r = it["res"]
            bad_city = bool(r) and L.norm_city(r["city"]) != first_city
            usable = bool(r) and not bad_city
            edge = ui.THEMES["green_strong"][2] if usable else "#E8A0A0"
            row = tk.Frame(c, bg="white", highlightthickness=1, highlightbackground=edge, padx=10, pady=6)
            row.pack(fill=tk.X, pady=3)
            var = tk.BooleanVar(master=self.top, value=usable)
            self.batch_vars.append((var, it))
            cb = ui.big_check(row, var, ui.THEMES["blue_strong"][0], enabled=usable)
            cb.grid(row=0, column=0, rowspan=3, sticky="n", padx=(0, 12), pady=(2, 0))
            tk.Label(row, text=f"{i}. {it['q']}", font=ui.FB, fg=INK, bg="white", anchor="w").grid(row=0, column=1, sticky="w")
            if not r:
                tk.Label(row, text="✗ " + it["err"], font=ui.F, fg=ui.BAD, bg="white", anchor="w", justify="left",
                         wraplength=self.W - 200).grid(row=1, column=1, sticky="w")
                continue
            tk.Label(row, text=f"{r['area']} {r['section']} {r['lot_number']} 地號", font=ui.F,
                     fg=ui.THEMES["blue_strong"][0], bg="white", anchor="w").grid(row=1, column=1, sticky="w")
            b = self._batch_build(r)
            if b:
                info = f"建號 {b['build_no']}（{b['door']}）"
            elif r["kind"] == "address" and r["builds"]:
                info = f"符合 {len(r['builds'])} 戶，先不指定建號（要哪一戶可之後單筆查）"
            else:
                info = r["build_note"] or "沒有建號資料"
            if r.get("community"):
                info += f"　·　社區：{r['community']}"
            if bad_city:
                info = f"不同縣市（{r['city']}），請另外建立案件"
            fg = ui.BAD if bad_city else (teal if b else INK2)
            tk.Label(row, text=info, font=ui.FS, fg=fg, bg="white", anchor="w", justify="left",
                     wraplength=self.W - 200).grid(row=2, column=1, sticky="w", pady=(2, 0))
            if usable:
                for w in (row,) + tuple(x for x in row.winfo_children() if x is not cb):
                    w.bind("<Button-1>", lambda e, v=var: v.set(not v.get()))
        self._footer(can_save=bool(ok))

    def _save_batch(self, mode):
        entries = [entry_from_result(it["res"], self._batch_build(it["res"]))
                   for var, it in self.batch_vars if var.get() and it["res"]]
        if not entries:
            tk.Label(self.body, text="請至少勾選一筆", font=ui.FB, fg=ui.BAD, bg=ui.BG).pack(anchor="w")
            return
        msg = self._city_conflict(entries, mode)
        if msg:
            tk.Label(self.body, text=msg, font=ui.FB, fg=ui.BAD, bg=ui.BG, wraplength=self.W - 80,
                     justify="left").pack(anchor="w")
            return
        try:
            folder = save_entries(entries, mode, self.data_json_path, self.base_dir)
        except Exception as e:
            tk.Label(self.body, text=f"儲存失敗：{e}", font=ui.FB, fg=ui.BAD, bg=ui.BG).pack(anchor="w")
            return
        self.top.destroy()
        if self.on_saved:
            self.on_saved(entries, mode, folder)

    def _city_conflict(self, entries, mode):
        """同一案件限同一縣市(子程式切換縣市的邏輯沒逐一驗過)；有衝突回提示文字"""
        cities = {L.norm_city(e["city"]) for e in entries}
        if mode == "append":
            cur = _load_data(self.data_json_path)
            if cur and cur[0].get("city"):
                cities.add(L.norm_city(cur[0]["city"]))
        if len(cities) > 1:
            return f"同一個案件只能放同一縣市的地號（目前有：{'、'.join(sorted(cities))}），不同縣市請另外建立案件。"
        return ""


INTRO_FLAG = "quick_lookup_intro_hide"
INTRO_EXAMPLES = [
    ("地址", "前鎮區和平二路236之2號12樓", "查到這一戶的地號、建號、社區"),
    ("地號", "前鎮區竹北段一小段2096", "看這塊地上有哪些門牌、建物"),
    ("多筆", "前鎮區竹北段一小段2096、2095", "用「、」分開，後面可以只寫號碼"),
]


def _tips_path(flag_dir):
    return os.path.join(flag_dir, "ui_tips.json")


def intro_hidden(flag_dir):
    try:
        with open(_tips_path(flag_dir), "r", encoding="utf-8") as f:
            return bool(json.load(f).get(INTRO_FLAG))
    except Exception:
        return False


def _set_intro_hidden(flag_dir, hide):
    p = _tips_path(flag_dir)
    try:
        with open(p, "r", encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        d = {}
    d[INTRO_FLAG] = bool(hide)
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ======================================================================
# 萬用查詢面板(主介面左側)：① 地籍選單 ② 地址選單 ③ 怎麼用
# ======================================================================
_ROADS = None


def load_roads():
    """全國路名(內政部戶政司 全國路名資料 115 年，data.gov.tw/dataset/35321)：{縣市: {鄉鎮市區: [街路]}}"""
    global _ROADS
    if _ROADS is None:
        try:
            import gzip
            with gzip.open(_asset("tw_roads.json.gz"), "rb") as f:
                _ROADS = json.loads(f.read().decode("utf-8")).get("roads", {})
        except Exception:
            _ROADS = {}
    return _ROADS


def _tips_get(flag_dir, key, default=None):
    try:
        with open(_tips_path(flag_dir), "r", encoding="utf-8") as f:
            return json.load(f).get(key, default)
    except Exception:
        return default


def _tips_set(flag_dir, **kw):
    p = _tips_path(flag_dir)
    try:
        with open(p, "r", encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        d = {}
    d.update(kw)
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


_CN = "零一二三四五六七八九十"


def _seg_cn(s):
    """段：阿拉伯數字 → 國字(官方門牌寫法「一段」)；已是國字就照用"""
    s = (s or "").strip().translate(L._FW).replace("段", "")
    if not s:
        return ""
    if s.isdigit():
        n = int(s)
        if n <= 10:
            s = _CN[n]
        elif n < 20:
            s = "十" + (_CN[n % 10] if n % 10 else "")
        else:
            s = _CN[n // 10] + "十" + (_CN[n % 10] if n % 10 else "")
    return s + "段"


def _num(s):
    """巷/弄/號/之/樓：全形轉半形、去掉使用者自己打的單位字"""
    return (s or "").strip().translate(L._FW).replace(" ", "").rstrip("巷弄衖號樓之")


_LOT_TOKEN = re.compile(r"\d+(?:\s*[-－–]\s*\d+)?")


def normalize_lots(text):
    """地號欄 → 「2096、2096-1、2100」。
    分隔可用 、，, ；; 空白 / 換行；多打或沒接數字的分隔號、雜字都丟掉；同一個地號只留一次。
    子號寫法 2096-1 / 2096－1 都可以(子號 0 視同沒有)"""
    t = (text or "").translate(L._FW)
    out = []
    for tok in _LOT_TOKEN.findall(t):
        tok = re.sub(r"\s*[-－–]\s*", "-", tok)
        mother, _, child = tok.partition("-")
        lot = str(int(mother)) + (f"-{int(child)}" if child and int(child) else "")
        if lot not in out:
            out.append(lot)
    return "、".join(out)


class FilterCombo(ttk.Combobox):
    """可打字過濾的下拉選單(例：打「中山」只剩含中山的路名)。
    打字時在下方列出候選(一般清單，不會像原生下拉那樣鎖住整個程式的滑鼠鍵盤)；
    滑鼠點選，或按 ↓ 移到清單、Enter 確定、Esc 收起"""

    def __init__(self, parent, var, font, width=18):
        super().__init__(parent, textvariable=var, font=font, width=width, height=14)
        self.all = []
        self._font = font
        self.sug = None
        self.bind("<KeyRelease>", self._filter)
        self.bind("<Down>", self._to_list)
        self.bind("<Escape>", lambda e: self._hide())
        self.bind("<FocusOut>", lambda e: self.after(150, self._maybe_hide))
        self.bind("<<ComboboxSelected>>", lambda e: self._hide())
        self.bind("<Destroy>", lambda e: self._hide(), add="+")

    def set_values(self, values):
        self.all = list(values)
        self["values"] = self.all
        self._hide()

    def _filter(self, e):
        if e.keysym in ("Up", "Down", "Return", "Escape", "Tab", "Left", "Right", "Shift_L", "Shift_R",
                        "Control_L", "Control_R"):
            return
        t = self.get().strip()
        vals = [v for v in self.all if t in v] if t else []
        self["values"] = vals or self.all
        if t and vals and not (len(vals) == 1 and vals[0] == t):
            self._show(vals[:200])
        else:
            self._hide()

    def _show(self, vals):
        top = self.winfo_toplevel()
        if self.sug is None or not self.sug.winfo_exists():
            self.sug = tk.Listbox(top, font=self._font, activestyle="dotbox", exportselection=False,
                                  bg="white", fg=INK, selectbackground=ui.PRIMARY, selectforeground="white",
                                  highlightthickness=1, highlightbackground=ui.BORDER, relief="flat")
            self.sug.bind("<ButtonRelease-1>", lambda e: self._choose())
            self.sug.bind("<Return>", lambda e: self._choose())
            self.sug.bind("<Escape>", lambda e: (self._hide(), self.focus_set()))
            self.sug.bind("<FocusOut>", lambda e: self.after(150, self._maybe_hide))
        self.sug.delete(0, tk.END)
        for v in vals:
            self.sug.insert(tk.END, v)
        self.sug.configure(height=min(8, len(vals)))
        x = self.winfo_rootx() - top.winfo_rootx()
        y = self.winfo_rooty() - top.winfo_rooty() + self.winfo_height()
        self.sug.place(x=x, y=y, width=max(self.winfo_width(), 200))
        self.sug.lift()

    def _to_list(self, e=None):
        if self.sug is not None and self.sug.winfo_ismapped() and self.sug.size():
            self.sug.focus_set()
            self.sug.selection_clear(0, tk.END)
            self.sug.selection_set(0)
            self.sug.activate(0)
            return "break"

    def _choose(self):
        try:
            sel = self.sug.curselection()
            if sel:
                self.set(self.sug.get(sel[0]))
        except tk.TclError:
            pass
        self._hide()
        self.focus_set()
        self.icursor(tk.END)

    def _maybe_hide(self):
        try:
            f = self.focus_get()
        except (tk.TclError, KeyError):
            f = None
        if f is not self and f is not self.sug:
            self._hide()

    def _hide(self):
        try:
            if self.sug is not None and self.sug.winfo_exists():
                self.sug.place_forget()
        except tk.TclError:
            pass

    def pick(self):
        """回傳確定的選項：完全符合，或過濾後只剩一個；否則 None"""
        t = self.get().strip()
        if t in self.all:
            return t
        cand = [v for v in self.all if t and t in v]
        return cand[0] if len(cand) == 1 else None


def _num_key(s):
    """地號/建號比對鍵：1375、1375-0、1375-0000 → '1375-0'；26539-000 → '26539-0'"""
    m = re.match(r"\s*0*(\d+)(?:\s*-\s*0*(\d+))?\s*$", str(s or ""))
    if not m:
        return None
    return f"{int(m.group(1))}-{int(m.group(2) or 0)}"


def _short_no(key):
    """'1375-0' → '1375'、'13-2' → '13-2'(顯示用)"""
    a, _, b = (key or "").partition("-")
    return a if b in ("", "0") else f"{a}-{b}"


def compare_case(data_path, trans_path):
    """目前案件(data.json)逐筆列出，並和謄本結構化 JSON 比對。
    不切縣市/行政區字串(會切錯)：謄本的「區段」前綴只用「結尾是段名、裡面含行政區」判斷同一段。
    回傳 {status: ok/partial/bad/none/nocase, case, trans_file, entries, extras, rows}
      entries：data.json 每一筆 {label, lot_state, builds:[(文字, 狀態)], community, coord}
      extras：謄本有、案件沒有 [(種類, 文字)]
      狀態：ok=謄本有、missing=謄本沒有、None=還沒做結構化"""
    fix = lambda s: str(s or "").replace("巿", "市").replace("\n", "").replace(" ", "").replace("\u3000", "")
    res = {"status": "nocase", "case": "", "trans_file": "", "entries": [], "extras": [], "rows": []}
    try:
        with open(data_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = []
    if isinstance(data, dict):
        data = [data]
    trans = []
    if trans_path and os.path.exists(trans_path):
        res["trans_file"] = os.path.basename(trans_path)
        try:
            with open(trans_path, "r", encoding="utf-8") as f:
                tj = json.load(f)
            for it in tj if isinstance(tj, list) else []:
                s = str((it.get("基本資訊") or {}).get("地號建號", ""))
                m = re.search(r"(\S+?)\s*(\d+)\s*-\s*(\d+)\s*(地號|建號)", s)
                if m:
                    trans.append([m.group(4), fix(m.group(1)), _num_key(f"{m.group(2)}-{m.group(3)}"), False])
        except Exception:
            pass

    def find(kind, area, sect, k):
        for t in trans:
            if t[0] == kind and t[2] == k and sect and t[1].endswith(sect) and (not area or area in t[1] or t[1] == sect):
                t[3] = True
                return "ok"
        return "missing" if trans else None
    n_lot = n_bld = 0
    for e in data if isinstance(data, list) else []:
        if not isinstance(e, dict):
            continue
        area, sect = fix(e.get("area")), fix(e.get("section"))
        k = _num_key(e.get("lot_number"))
        label = f"{fix(e.get('city'))}{area}{sect} {_short_no(k) if k else (e.get('lot_number') or '？')} 地號"
        ent = {"label": label, "lot_state": find("地號", area, sect, k) if k else None, "builds": [],
               "community": str(e.get("community") or "").strip(), "coord": bool(str(e.get("coordinates") or "").strip())}
        if k:
            n_lot += 1
            res["rows"].append(("地號", label, ent["lot_state"]))
        blds = [(b.get("build_no"), b.get("door", "")) for b in (e.get("builds") or []) if isinstance(b, dict)]
        if not blds and e.get("build_no"):
            blds = [(e.get("build_no"), e.get("door", ""))]
        seen = set()
        for bno, door in blds:
            bk = _num_key(bno)
            if not bk or bk in seen:
                continue
            seen.add(bk)
            n_bld += 1
            stt = find("建號", area, sect, bk)
            door = str(door or "").strip()
            ent["builds"].append((f"建號 {_short_no(bk)}" + (f"（{door}）" if door else ""), stt))
            res["rows"].append(("建號", bk, stt))
        res["entries"].append(ent)
    if not res["entries"]:
        return res
    res["case"] = f"共 {n_lot} 筆地號" + (f"、{n_bld} 戶建號" if n_bld else "")
    for t in trans:
        if not t[3]:
            res["extras"].append((t[0], f"{t[1]} {_short_no(t[2])} {t[0]}"))
    if not trans:
        res["status"] = "none"
    else:
        ok = sum(1 for r in res["rows"] if r[2] == "ok")
        bad = sum(1 for r in res["rows"] if r[2] == "missing") + sum(1 for x in res["extras"] if x[0] == "地號")
        res["status"] = "ok" if not bad else ("bad" if ok == 0 else "partial")
    return res


def _short_trans_name(name):
    """trans_岡山區石螺段1375_土地謄本_第二類_20251114_1850.json → 岡山區石螺段1375 土地謄本 第二類（2025-11-14 18:50）"""
    m = re.match(r"(?:trans|transcript_data)_(.*?)_(\d{8})_(\d{4})\.json$", name or "")
    if not m:
        return name
    d, t = m.group(2), m.group(3)
    return f"{m.group(1).replace('_土地建物謄本', '').replace('_土地謄本', '').replace('_', ' ')}　{d[:4]}-{d[4:6]}-{d[6:]} {t[:2]}:{t[2:]}"


class QuickIntro:
    """萬用查詢面板(擺在主介面旁邊)：分頁 ① 地籍選單(縣市→行政區→段小段→地號)
    ② 地址選單(國土測繪中心門牌格式：縣市→鄉鎮市區→街路→段→巷→弄→號(之)→樓) ③ 怎麼用。
    選好按「查詢」→ 組成文字交給 on_query(走萬用查詢流程)。
    guide=True(啟動自動跳出)時，箭頭指著輸入欄、輸入欄閃橘框。勾「下次不要再顯示」寫入 ui_tips.json"""

    TABS = [("lot", "🗺️ 地籍選單"), ("addr", "🏠 地址選單"), ("help", "💡 怎麼用")]

    def __init__(self, root, input_widget, flag_dir, placeholder="", on_query=None, tab=None, guide=True, restore=None,
                 case_info=None):
        self.case_info = case_info              # 回傳 (data.json 路徑, 目前選用的謄本結構化 JSON 路徑)
        self.root, self.inp, self.flag_dir, self.placeholder = root, input_widget, flag_dir, placeholder
        self.rv = dict(restore or {})            # 回上一步：要帶回的值(段小段/街路等清單載入後才填)
        self.on_query, self.guide = on_query, guide
        self.top = top = tk.Toplevel(root, bg=ui.BG)
        top.title("萬用查詢")
        top.transient(root)
        try:
            top.attributes("-alpha", 0.0)      # 排版、擺放好才現身(不然會看到畫一半的黑塊)
        except tk.TclError:
            pass
        top.resizable(False, False)
        self.W = self._fit_width()
        self.FONT = (ui.FAMILY, 18)                  # 選單/輸入框
        self.LF = (ui.FAMILY, 17, "bold")            # 欄位名稱
        self.HF = (ui.FAMILY, 15)                    # 說明文字
        self.BF = (ui.FAMILY, 17, "bold")            # 查詢按鈕
        top.option_add("*TCombobox*Listbox.font", self.FONT)
        ui.header(top, "🔎 萬用查詢", "用選單挑地號或地址，阿生幫你查地號、建號、社區，並建立案件")
        foot = ui.footer(top)
        self.body = body = tk.Frame(top, bg=ui.BG, padx=16, pady=8)
        body.pack(fill=tk.BOTH, expand=True)
        self.wrap = self.W - 110

        # 分頁切換
        self.tab_var = tk.StringVar(master=top)
        names = [n for _, n in self.TABS]
        seg = ui.segmented(body, self.tab_var, names, ui.THEMES["blue_strong"][0])
        seg.pack(anchor="w", pady=(4, 6))
        self.pages = {}
        self._build_lot()
        self._build_addr()
        self._build_help()
        self._build_status()
        self._anchor = tk.Frame(body, bg=ui.BG, height=1)      # 分頁/比對卡都排在它前面
        self._anchor.pack(fill=tk.X)
        # 錯誤提示：有訊息才浮在面板底部(不佔版面，空間留給比對清單)
        self.msg = tk.Label(body, text="", font=ui.FB, fg=ui.BAD, bg="#FFF4E5", justify="left", wraplength=self.wrap,
                            padx=10, pady=4, highlightthickness=1, highlightbackground="#F5C27A")
        self.tab_var.trace_add("write", lambda *a: self._show_tab())
        start = tab or ("help" if guide and not _tips_get(flag_dir, "qp_used") else _tips_get(flag_dir, "qp_tab", "lot"))
        self.tab_var.set(dict(self.TABS).get(start, names[0]))

        # 頁尾
        self.hide_var = tk.BooleanVar(master=top, value=intro_hidden(flag_dir))
        hide_box = tk.Frame(foot, bg=foot.cget("bg"))
        hide_box.pack(side=tk.LEFT)
        ui.big_check(hide_box, self.hide_var).pack(side=tk.LEFT, padx=(0, 8))
        hl = tk.Label(hide_box, text="啟動時不要自動跳出", font=ui.F, fg=INK, bg=foot.cget("bg"), cursor="hand2")
        hl.pack(side=tk.LEFT)
        hl.bind("<Button-1>", lambda e: self.hide_var.set(not self.hide_var.get()))
        ui.button(foot, "關閉", self._close, "ghost").pack(side=tk.RIGHT)
        top.protocol("WM_DELETE_WINDOW", self._close)

        self.arrow = None
        self._orig = None
        self._maxH = 0
        self._measuring = True
        # 三個分頁都先量一次高度(視窗還透明)，取最高的當固定大小 → 切換分頁時位置、大小都不跳
        start_name = self.tab_var.get()
        t = 40
        for k, n in self.TABS:
            top.after(t, lambda n=n: self.tab_var.set(n))
            top.after(t + 160, self._place)
            t += 320
        top.after(t, lambda: self.tab_var.set(start_name))
        top.after(t + 60, self._end_measure)
        top.after(t + 200, self._reveal)
        if guide:
            self._blink(0)

    def _end_measure(self):
        self._measuring = False
        self._place()

    # ---------- 共用 ----------
    def _row(self, parent, label, width_px=None):
        r = tk.Frame(parent, bg=parent.cget("bg"))
        r.pack(fill=tk.X, pady=4)
        tk.Label(r, text=label, font=self.LF, fg=INK, bg=parent.cget("bg"), width=6, anchor="e").pack(side=tk.LEFT, padx=(0, 10))
        return r

    def _combo(self, parent, var, width=16, filt=False):
        if filt:
            c = FilterCombo(parent, var, self.FONT, width)
        else:
            c = ttk.Combobox(parent, textvariable=var, font=self.FONT, width=width, state="readonly", height=14)
        return c

    def _entry(self, parent, var, width=6):
        e = tk.Entry(parent, textvariable=var, font=self.FONT, width=width, relief="flat", bg="white", fg=INK,
                     highlightthickness=1, highlightbackground=ui.BORDER, highlightcolor=ui.PRIMARY, justify="center")
        return e

    def _bg(self, work, done):
        """背景抓清單，抓完回到介面執行 done(結果, 錯誤)"""
        box = {}

        def run():
            try:
                box["r"] = work()
            except Exception as e:
                box["e"] = str(e)

        def poll():
            if not box:
                self.top.after(120, poll)
                return
            try:
                done(box.get("r"), box.get("e"))
            except tk.TclError:
                pass
        threading.Thread(target=run, daemon=True).start()
        poll()

    def _city_town(self, parent, city_var, town_var, on_town=None):
        """縣市＋行政區兩個下拉；行政區清單跟著縣市背景載入"""
        r = self._row(parent, "縣市")
        cb_city = self._combo(r, city_var, 10)
        cb_city["values"] = list(L.CITY_CODES)
        cb_city.pack(side=tk.LEFT)
        tk.Label(r, text="行政區", font=self.LF, fg=INK, bg=parent.cget("bg")).pack(side=tk.LEFT, padx=(16, 10))
        cb_town = self._combo(r, town_var, 10)
        cb_town.pack(side=tk.LEFT)
        state = {"towns": []}

        def load(*_):
            city = city_var.get()
            cb_town["values"] = ["載入中…"]
            town_var.set("")

            def done(res, err):
                if err:
                    self._say(err)
                    return
                state["towns"] = res or []
                names = [n for n, _ in state["towns"]]
                cb_town["values"] = names
                want = _tips_get(self.flag_dir, "qp_town", "")
                if want in names and city == _tips_get(self.flag_dir, "qp_city", ""):
                    town_var.set(want)
            self._bg(lambda: L.list_towns(city), done)
        city_var.trace_add("write", load)
        if on_town:
            town_var.trace_add("write", lambda *a: on_town(dict(state["towns"]).get(town_var.get(), "")))
        return state

    def _say(self, text, ok=False):
        try:
            self.msg.config(text=text, fg=ui.THEMES["green_strong"][0] if ok else ui.BAD)
            if text:
                self.msg.place(x=0, rely=1.0, y=-2, anchor="sw")
                self.msg.lift()
            else:
                self.msg.place_forget()
        except tk.TclError:
            pass

    # ---------- ① 地籍選單 ----------
    def _build_lot(self):
        c = ui.card(self.body, "地籍：縣市 → 行政區 → 段小段 → 地號", "", "blue_strong")
        self.pages["lot"] = c
        bg = c.cget("bg")
        self.lot_city = tk.StringVar(master=self.top)
        self.lot_town = tk.StringVar(master=self.top)
        self.lot_sect = tk.StringVar(master=self.top)
        self.lot_no = tk.StringVar(master=self.top)

        def on_town(tid):
            self.lot_sect.set("")
            cb_sect.set_values([])
            if not tid:
                return
            city = self.lot_city.get()
            cb_sect["values"] = ["載入中…"]
            self._bg(lambda: L.list_sections(city, tid),
                     lambda res, err: self._say(err) if err else self._sect_loaded(res or []))
        self._city_town(c, self.lot_city, self.lot_town, on_town)
        r = self._row(c, "段小段")
        cb_sect = self._combo(r, self.lot_sect, 24, filt=True)
        cb_sect.pack(side=tk.LEFT)
        tk.Label(r, text="可打字過濾", font=self.HF, fg=INK2, bg=bg).pack(side=tk.LEFT, padx=10)
        self.cb_sect = cb_sect
        r = self._row(c, "地號")
        e = self._entry(r, self.lot_no, 22)
        e.configure(justify="left")
        e.pack(side=tk.LEFT, ipady=3)
        e.bind("<Return>", lambda ev: self._query_lot())
        e.bind("<FocusOut>", lambda ev: self._tidy_lots(), add="+")     # 離開欄位就整理一次
        self.lot_entry = e
        blue = ui.THEMES["blue_strong"][0]
        # 「、」鍵：數字和中文頓號不用切輸入法；按一下在游標處插入
        ui.button(r, "、下一筆", self._insert_sep, "chip", blue, font=self.LF, padx=14, pady=4).pack(side=tk.LEFT, padx=(8, 0))
        ui.button(r, "清除", lambda: (self.lot_no.set(""), e.focus_set()), "ghost", font=self.HF, padx=12, pady=4).pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(c, text="例：2096、2096-1；多筆用「、」或「,」分開（可按「、下一筆」）",
                 font=self.HF, fg=INK2, bg=bg).pack(anchor="w", padx=(8, 0), pady=(2, 0))
        tk.Label(c, text="多打的符號、空白或不是地號的字，查詢前會自動整理",
                 font=self.HF, fg=INK2, bg=bg).pack(anchor="w", padx=(8, 0), pady=(0, 6))
        ui.button(c, "🔍  查詢地號", self._query_lot, "primary", ui.THEMES["blue_strong"][0], font=self.BF, padx=28, pady=12).pack(anchor="e", pady=(4, 2))
        self.lot_city.set(_tips_get(self.flag_dir, "qp_city", "") or "高雄市")
        if self.rv.get("lots"):
            self.lot_no.set(self.rv["lots"])

    def _sect_loaded(self, names):
        self.cb_sect.set_values(names)
        want = self.rv.pop("sect", "")
        if want in names:
            self.lot_sect.set(want)

    def _insert_sep(self):
        e = self.lot_entry
        e.focus_set()
        txt = e.get()
        pos = e.index(tk.INSERT)
        if pos == 0 or txt[pos - 1] in "、,，":       # 開頭或前面已經是分隔號就不再加
            return
        e.insert(pos, "、")
        e.icursor(pos + 1)

    def _tidy_lots(self):
        """把地號欄整理成「2096、2096-1」：分隔用「、」，多餘的符號、空白、非地號的字都拿掉"""
        lots = normalize_lots(self.lot_no.get())
        if lots != self.lot_no.get():
            self.lot_no.set(lots)
        return lots

    def _query_lot(self):
        city, town = self.lot_city.get(), self.lot_town.get()
        sect = self.cb_sect.pick()
        lots = self._tidy_lots()
        if not (city and town and town != "載入中…"):
            return self._say("請先選縣市與行政區")
        if not sect:
            return self._say("請從清單選擇段小段（可打幾個字過濾）")
        if not re.search(r"\d", lots):
            return self._say("請輸入地號，例如 2096 或 2096-1")
        self._remember(city, town, "lot")
        self._go(f"{city}{town}{sect}{lots}")

    # ---------- ② 地址選單(官方門牌格式) ----------
    def _build_addr(self):
        c = ui.card(self.body, "地址：依國土測繪中心門牌格式", "街路→段→巷→弄→號(之)→樓", "green_strong")
        self.pages["addr"] = c
        bg = c.cget("bg")
        self.ad_city = tk.StringVar(master=self.top)
        self.ad_town = tk.StringVar(master=self.top)
        self.ad_road = tk.StringVar(master=self.top)
        self.ad = {k: tk.StringVar(master=self.top) for k in ("seg", "lane", "alley", "no", "no_zhi", "fl", "fl_zhi")}

        def on_town(_tid):
            self.ad_road.set("")
            roads = load_roads().get(L.norm_city(self.ad_city.get()), {}).get(self.ad_town.get(), [])
            cb_road.set_values(roads)
            want = self.rv.get("road", "")
            if want and want in roads:            # 清單真的載入、而且有這條路才帶回(行政區清空時不要用掉)
                self.ad_road.set(want)
                self.rv.pop("road", None)
        self._city_town(c, self.ad_city, self.ad_town, on_town)
        r = self._row(c, "街路")
        cb_road = self._combo(r, self.ad_road, 20, filt=True)
        cb_road.pack(side=tk.LEFT)
        tk.Label(r, text="可打字過濾", font=self.HF, fg=INK2, bg=bg).pack(side=tk.LEFT, padx=10)
        self.cb_road = cb_road

        def units(parent, items, label):
            """一排門牌欄位：(欄位, 單位字, 寬度, 單位字在前？)。「之」照門牌讀法放在格子前面(236號 之2)"""
            r = self._row(parent, label)
            for key, unit, w, before in items:
                if key is None:                       # 分組間距
                    tk.Frame(r, bg=bg, width=26).pack(side=tk.LEFT)
                    continue
                if before:
                    tk.Label(r, text=unit, font=self.LF, fg=INK, bg=bg).pack(side=tk.LEFT, padx=(0, 4))
                e = self._entry(r, self.ad[key], w)
                e.pack(side=tk.LEFT, ipady=3)
                e.bind("<Return>", lambda ev: self._query_addr())
                if not before:
                    tk.Label(r, text=unit, font=self.LF, fg=INK, bg=bg).pack(side=tk.LEFT, padx=(4, 12))
        units(c, [("seg", "段", 4, False), ("lane", "巷", 5, False), ("alley", "弄", 5, False)], "段巷弄")
        # 號：國土測繪門牌寫法是「236之2號」(之在號前面)
        units(c, [("no", "之", 6, False), ("no_zhi", "號", 4, False)], "門牌號")
        units(c, [("fl", "樓", 4, False), ("fl_zhi", "之", 4, True)], "樓層")
        for line in ("・沒有的欄位空著；段可打「1」或「一」；沒有「之」只填前面那格",
                     "・例：中山北路一段12巷3弄 236之2號 12樓之1（樓層用來比對建號）"):
            tk.Label(c, text=line, font=self.HF, fg=INK2, bg=bg, anchor="w").pack(anchor="w", padx=(8, 0))
        br = tk.Frame(c, bg=bg)
        br.pack(fill=tk.X, pady=(4, 2))
        ui.button(br, "🔍  查詢地址", self._query_addr, "primary", ui.THEMES["green_strong"][0], font=self.BF, padx=28, pady=12).pack(side=tk.RIGHT)
        tk.Label(br, text="路名資料：內政部戶政司\n「全國路名資料」115 年", font=self.HF, fg=INK2, bg=bg, justify="left"
                 ).pack(side=tk.LEFT, padx=(8, 0))
        self.ad_city.set(_tips_get(self.flag_dir, "qp_city", "") or "高雄市")
        for k, v in (self.rv.get("ad") or {}).items():
            if k in self.ad:
                self.ad[k].set(v)

    def compose_address(self, city, town, road, v):
        """國土測繪中心門牌格式：街路＋段(國字)＋巷＋弄＋號(「N之M號」)＋樓(之K)"""
        s = f"{city}{town}{road}{_seg_cn(v['seg'])}"
        for key, unit in (("lane", "巷"), ("alley", "弄")):
            n = _num(v[key])
            if n:
                s += n + unit
        no, zhi = _num(v["no"]), _num(v["no_zhi"])
        if no:
            s += f"{no}之{zhi}號" if zhi else f"{no}號"
        fl, fzhi = _num(v["fl"]), _num(v["fl_zhi"])
        if fl:
            s += f"{fl}樓" + (f"之{fzhi}" if fzhi else "")
        return s

    def _query_addr(self):
        city, town = self.ad_city.get(), self.ad_town.get()
        road = self.cb_road.pick()
        v = {k: var.get() for k, var in self.ad.items()}
        if not (city and town and town != "載入中…"):
            return self._say("請先選縣市與鄉鎮市區")
        if not road:
            return self._say("請從清單選擇街路（可打幾個字過濾）")
        if not _num(v["no"]):
            return self._say("請輸入門牌號碼（號）")
        self._remember(city, town, "addr")
        self._go(self.compose_address(city, town, road, v))

    # ---------- ③ 怎麼用 ----------
    def _build_help(self):
        page = tk.Frame(self.body, bg=ui.BG)
        self.pages["help"] = page
        wrap = self.wrap
        c1 = ui.card(page, "建立案件的兩種方式", "存的是同一份 data.json，可以混用", "orange_strong")
        bg1 = c1.cget("bg")
        row = tk.Frame(c1, bg=bg1)
        row.pack(fill=tk.X)
        self.cat = _cat_photo(self.top, 64, ring=3)
        if self.cat:
            tk.Label(row, image=self.cat, bg=bg1).pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(row, text="①【地籍便民系統】（主畫面橙色按鈕）：在地籍圖網頁上查地號。\n"
                           "②【萬用查詢】：用上面的「地籍選單」「地址選單」挑好按查詢，"
                           "或在主畫面最下方的輸入欄打地址／地號按 Enter（還會查建號、社區）。",
                 font=ui.F, fg=INK, bg=bg1, justify="left", wraplength=wrap - 80).pack(side=tk.LEFT, anchor="w")
        tk.Label(c1, text="※執行地籍便民等功能的過程中，輸入欄是給該功能用的，請等它結束再查。\n"
                          "※建立「新案件」後，若原本的謄本結構化是別的案件，會自動解除，請到【謄本與土增稅】重新做結構化。",
                 font=ui.FS, fg=ui.WARN_FG, bg=bg1, justify="left", wraplength=wrap).pack(anchor="w", pady=(8, 0))
        c2 = ui.card(page, "輸入欄可以打什麼", "按「試試看」填進輸入欄，再按 Enter", "blue_strong")
        blue = ui.THEMES["blue_strong"][0]
        for kind, ex, desc in INTRO_EXAMPLES:
            r = tk.Frame(c2, bg="white", highlightthickness=1, highlightbackground=ui.THEMES["blue_strong"][2],
                         padx=10, pady=6)
            r.pack(fill=tk.X, pady=3)
            tk.Label(r, text=kind, font=ui.FB, fg="white", bg=blue, padx=8).grid(row=0, column=0, rowspan=2, sticky="nw")
            tk.Label(r, text=ex, font=ui.FB, fg=INK, bg="white", anchor="w").grid(row=0, column=1, sticky="w", padx=(10, 0))
            tk.Label(r, text=desc, font=ui.FS, fg=INK2, bg="white", anchor="w").grid(row=1, column=1, sticky="w", padx=(10, 0))
            r.grid_columnconfigure(1, weight=1)
            ui.button(r, "試試看", lambda ex=ex: self._try(ex), "chip", blue).grid(row=0, column=2, rowspan=2, sticky="e")
        c3 = ui.card(page, "查到之後", "", "green_strong")
        tk.Label(c3, text="確認位置、選好建號 → 按「建立新案件」或「加入目前案件」。"
                          "會自動寫入 data.json 並建好案件資料夾，各查詢功能就能直接接著用。",
                 font=ui.F, fg=INK, bg=c3.cget("bg"), justify="left", wraplength=wrap).pack(anchor="w")
        tk.Label(page, text="之後要再打開這個面板：按輸入欄右邊的「☰ 選單」，或輸入欄空白時按 Enter。",
                 font=ui.FS, fg=INK2, bg=ui.BG).pack(anchor="w", pady=(2, 0))

    # ---------- 地籍、地址頁下方：目前案件 ↔ 謄本結構化 比對 ----------
    # 第二部分(謄本結構化)：一條色帶簡單講結果，顏色跟主介面狀態標籤同一套
    _ST_BANNER = {
        "ok": ("✓ 謄本結構化：與目前案件相符", "#1B7F3B", "#E3F5E8"),
        "partial": ("⚠ 謄本結構化部分對不上 → 請到【謄本與土增稅】重做", "#FFFFFF", "#D93025"),
        "bad": ("⚠ 謄本結構化是別的案件 → 請到【謄本與土增稅】重做", "#FFFFFF", "#D93025"),
        "none": ("○ 謄本結構化：目前案件還沒做（調好謄本後到【謄本與土增稅】做）", "#B35400", "#FFEBD6"),
        "nocase": ("○ 謄本結構化：還沒有案件", "#B35400", "#FFEBD6"),
    }

    def _build_status(self):
        """地籍、地址兩頁各放一張比對卡(各自依剩餘高度收筆數)"""
        self.st = {}
        if not self.case_info:
            return
        try:
            data_path, trans_path = self.case_info()
            r = compare_case(data_path, trans_path)
        except Exception as e:
            print(f"[萬用查詢] 案件比對失敗：{e}")
            return
        for tab in ("lot", "addr"):
            self.st[tab] = self._make_status(r)

    def _make_status(self, r):
        """第一部分(重點)：目前案件 data.json 的內容——幾筆地段地號、幾戶建號、逐筆列出，
        讓使用者判斷還要加什麼，或重新建立新案件。第二部分：謄本結構化一條色帶簡單比對"""
        blue, tint, edge = ui.THEMES["blue_strong"]
        box = tk.Frame(self.body, bg=ui.BG)
        st = {"box": box, "rows": [], "more": None, "tip": None, "total": 0, "bad": 0, "cv": None}
        n_lot = len(r["entries"])
        n_bld = sum(len(e["builds"]) for e in r["entries"])

        # ===== 第一部分：data.json =====
        part1 = tk.Frame(box, bg="white", highlightthickness=2, highlightbackground=blue)
        part1.pack(fill=tk.X)
        head = tk.Frame(part1, bg=blue)
        head.pack(fill=tk.X)
        tk.Label(head, text="📋 目前案件（data.json）", font=(ui.FAMILY, 16, "bold"), fg="white", bg=blue,
                 padx=12).pack(side=tk.LEFT, pady=4)
        cnt = (f"地段地號 {n_lot} 筆　建號 {n_bld} 戶" if n_lot else "還沒有案件")
        tk.Label(head, text=cnt, font=(ui.FAMILY, 16, "bold"), fg="white", bg=blue, padx=12).pack(side=tk.RIGHT, pady=4)
        c = tk.Frame(part1, bg="white", padx=12, pady=6)
        c.pack(fill=tk.X)
        if not n_lot:
            tk.Label(c, text="用上面的選單查地號或地址，按「建立新案件」就會寫進 data.json", font=self.HF,
                     fg=INK, bg="white", anchor="w").pack(anchor="w")
        else:
            wrapf = tk.Frame(c, bg="white", highlightthickness=1, highlightbackground=edge)
            wrapf.pack(fill=tk.X)
            cv = tk.Canvas(wrapf, bg="white", highlightthickness=0, height=40)
            sb = ttk.Scrollbar(wrapf, orient="vertical", command=cv.yview)
            cv.configure(yscrollcommand=sb.set)
            cv.pack(side=tk.LEFT, fill=tk.X, expand=True)
            tb = tk.Frame(cv, bg="white", padx=8, pady=2)
            win = cv.create_window(0, 0, window=tb, anchor="nw")
            st.update(cv=cv, sb=sb, tb=tb, full=0)

            def fit(_e=None):
                cv.itemconfigure(win, width=cv.winfo_width())
                h = tb.winfo_reqheight()
                cv.configure(scrollregion=(0, 0, 0, h))
                if st["full"] != h:                    # 內容高度變了 → 先給全高，_place 再依螢幕縮
                    st["full"] = h
                    cv.configure(height=h)
                sb.pack(side=tk.RIGHT, fill=tk.Y) if int(float(cv.cget("height"))) < h else sb.pack_forget()
            tb.bind("<Configure>", fit)
            cv.bind("<Configure>", fit)
            st["fit"] = fit
            row = 0
            for i, e in enumerate(r["entries"], 1):
                lines = [("", f"{i}. {e['label']}", INK, (ui.FAMILY, 15, "bold"))]
                for txt, _stt in e["builds"]:
                    lines.append(("", "　　" + txt, INK, (ui.FAMILY, 15)))
                if not e["builds"]:
                    lines.append(("", "　　建號：（沒有）", INK2, (ui.FAMILY, 14)))
                if e["community"]:
                    lines.append(("", f"　　社區：{e['community']}", INK2, (ui.FAMILY, 14)))
                if not e["coord"]:
                    lines.append(("⚠", "　　沒有座標（地圖類查詢會抓不到位置，建議重新建立）", "#C62828", (ui.FAMILY, 14, "bold")))
                for icon, txt, col, f in lines:
                    tk.Label(tb, text=icon, font=(ui.FAMILY, 14, "bold"), fg=col, bg="white").grid(row=row, column=0)
                    tk.Label(tb, text=txt, font=f, fg=col, bg="white", anchor="w").grid(row=row, column=1, sticky="w")
                    row += 1
            tb.grid_columnconfigure(1, weight=1)

            def wheel(ev):
                if int(float(cv.cget("height"))) < st["full"]:
                    cv.yview_scroll(-1 if ev.delta > 0 else 1, "units")
                return "break"

            def bind_wheel(w):
                w.bind("<MouseWheel>", wheel)
                for ch in w.winfo_children():
                    bind_wheel(ch)
            bind_wheel(cv)
            st["tip"] = tk.Label(c, text="→加地號／建號：查詢後按「加入目前案件」；換案件：按「建立新案件」",
                                 font=self.HF, fg=blue, bg="white", anchor="w", justify="left")
            st["tip"].pack(anchor="w", pady=(6, 0))

        # ===== 第二部分：謄本結構化(簡單比對) =====
        text, fg, bbg = self._ST_BANNER[r["status"]]
        part2 = tk.Frame(box, bg=bbg, highlightthickness=1,
                         highlightbackground=bbg if r["status"] in ("partial", "bad") else edge)
        part2.pack(fill=tk.X, pady=(6, 0))
        tk.Label(part2, text=text, font=(ui.FAMILY, 15, "bold"), fg=fg, bg=bbg, anchor="w", padx=12
                 ).pack(anchor="w", pady=(4, 0))
        name = _short_trans_name(r["trans_file"])
        sub = {"ok": f"使用：{name}", "partial": f"使用：{name}", "bad": f"謄本是：{name}"}.get(r["status"], "")
        if sub:
            tk.Label(part2, text=sub, font=self.HF, fg=fg, bg=bbg, anchor="w", padx=12).pack(anchor="w", pady=(0, 4))
        return st

    def _st_more_update(self, st):
        if st["more"] is None:
            return
        hidden = st["total"] - len(st["rows"])
        if hidden <= 0:
            st["more"].grid_remove()
            return
        if st["rows"]:
            st["more"].config(text=f"…另有 {hidden} 筆")
        else:
            st["more"].config(text=f"共 {st['total']} 筆" + (f"，其中 {st['bad']} 筆對不上" if st["bad"] else ""))
        st["more"].grid()

    def _st_shrink(self, st, avail):
        """螢幕不夠高：先把清單區縮到約 3 行高(可捲動)，再收提示行，最後清單區縮到 1 行半"""
        top = self.top
        cv = st.get("cv")
        if cv is None:
            if st["tip"] is not None and top.winfo_reqheight() > avail:
                st["tip"].destroy()
                st["tip"] = None
            return
        line = max(24, st["tb"].winfo_reqheight() // max(1, len(st["tb"].grid_slaves(column=1))))

        def squeeze(floor):
            over = top.winfo_reqheight() - avail
            h = int(float(cv.cget("height")))
            if over > 0 and h > floor:
                cv.configure(height=max(floor, h - over))
                top.update_idletasks()
                st["fit"]()
        squeeze(min(st["full"], int(line * 2.5)))
        if top.winfo_reqheight() > avail and st["tip"] is not None:
            st["tip"].destroy()
            st["tip"] = None
            top.update_idletasks()
            h = int(float(cv.cget("height")))
            room = avail - top.winfo_reqheight()
            if room > 0:                               # 收掉提示行多出來的空間還給清單
                cv.configure(height=min(st["full"], h + room))
                top.update_idletasks()
        squeeze(int(line * 1.6))
        st["fit"]()

    def _show_tab(self):
        key = next((k for k, n in self.TABS if n == self.tab_var.get()), "lot")
        for k, w in self.pages.items():
            target = w.master if k != "help" else w      # ui.card 回傳內容框，外層 canvas 才是要切換的
            if k != "help":
                while target.master is not self.body:
                    target = target.master
            if k == key:
                target.pack(fill=tk.X, before=self._anchor)
            else:
                target.pack_forget()
        for k, st in getattr(self, "st", {}).items():   # 比對卡只放在地籍、地址兩頁下方
            if k == key:
                st["box"].pack(fill=tk.X, pady=(2, 4), before=self._anchor)
            else:
                st["box"].pack_forget()
        self._say("")
        if not getattr(self, "_measuring", False):
            _tips_set(self.flag_dir, qp_tab=key)

    def _remember(self, city, town, tab):
        _tips_set(self.flag_dir, qp_city=city, qp_town=town, qp_tab=tab, qp_used=True)

    def _go(self, text):
        key = next((k for k, n in self.TABS if n == self.tab_var.get()), "lot")
        _RETURN.clear()
        _RETURN.update(root=self.root, inp=self.inp, flag_dir=self.flag_dir, on_query=self.on_query, tab=key,
                       case_info=self.case_info,
                       values={"sect": self.lot_sect.get(), "lots": self.lot_no.get(), "road": self.ad_road.get(),
                               "ad": {k: v.get() for k, v in self.ad.items()}})
        self._close()
        if self.on_query:
            self.on_query(text)

    def _fit_width(self):
        """依主介面旁的空間決定寬度(越寬越矮)：620~920"""
        try:
            self.root.update_idletasks()
            fx = self.root.winfo_x()
            fw = self.root.winfo_width() + (self.root.winfo_rootx() - fx) * 2
            space = max(fx, self.root.winfo_screenwidth() - (fx + fw)) - 96 - 16
            return int(max(620, min(920, space)))
        except Exception:
            return 700

    def _reveal(self):
        try:
            self.top.update_idletasks()
            self.top.attributes("-alpha", 1.0)
        except tk.TclError:
            pass

    # ---------- 擺放與箭頭 ----------
    def _place(self):
        top, root = self.top, self.root
        try:
            top.update_idletasks()
        except tk.TclError:
            return
        avail = self._avail_height()
        for st in getattr(self, "st", {}).values():
            if st["box"].winfo_ismapped():
                self._st_shrink(st, avail)
        H = max(top.winfo_reqheight(), getattr(self, "_maxH", 0))   # 固定用三頁中最高的高度
        self._maxH = H
        W = self.W
        try:
            root.update_idletasks()
            iy = self.inp.winfo_rooty() + self.inp.winfo_height() // 2      # 輸入欄中線(螢幕座標)
            fx, fw = root.winfo_x(), root.winfo_width() + (root.winfo_rootx() - root.winfo_x()) * 2
            sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
            gap = 96                                                          # 留給箭頭
            on_left = fx >= W + gap or fx >= sw - (fx + fw)
            X = fx - gap - W if on_left else fx + fw + gap
            X = max(0, min(X, sw - W - 16))
            wa_top, wa_bot = 0, sh
            try:                                                               # 工作區(扣掉工作列)
                import ctypes
                from ctypes import wintypes
                rc = wintypes.RECT()
                if ctypes.windll.user32.SystemParametersInfoW(0x30, 0, ctypes.byref(rc), 0):
                    wa_top, wa_bot = rc.top, rc.bottom
            except Exception:
                pass
            title_h = 45                                                       # 標題列＋下外框
            H = min(H, wa_bot - wa_top - title_h)                              # 螢幕太矮：頁尾按鈕仍留在畫面內
            Y = max(wa_top, min(iy + 60 - H, wa_bot - H - title_h))            # 底部略低於輸入欄
            top.geometry(f"{W}x{H}+{int(X)}+{int(Y)}")
            if self.arrow is None and self.guide:
                self._make_arrow(on_left, (fx - gap + 4) if on_left else (fx + fw + 4), iy)
        except Exception:
            top.geometry(f"{W}x{H}")

    def _avail_height(self):
        """工作區(扣掉工作列)高度 − 標題列"""
        h = self.top.winfo_screenheight()
        try:
            import ctypes
            from ctypes import wintypes
            rc = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(0x30, 0, ctypes.byref(rc), 0):
                h = rc.bottom - rc.top
        except Exception:
            pass
        return h - 45

    def _make_arrow(self, on_left, x, iy):
        """主介面旁的透明小視窗，畫一支會前後動的橘色箭頭指著輸入欄"""
        try:
            a = tk.Toplevel(self.top)
            a.overrideredirect(True)
            key = "#FF00FE"
            a.configure(bg=key)
            a.attributes("-transparentcolor", key)
            a.attributes("-topmost", True)
            w, h = 92, 64
            a.geometry(f"{w}x{h}+{int(x)}+{int(iy - h / 2)}")
            cv = tk.Canvas(a, width=w, height=h, bg=key, highlightthickness=0)
            cv.pack()
            self.arrow, self.arrow_cv, self.arrow_dir = a, cv, (1 if on_left else -1)
        except Exception:
            self.arrow = None

    def _draw_arrow(self, t):
        if not self.arrow:
            return
        cv, d = self.arrow_cv, self.arrow_dir
        cv.delete("all")
        off = 8 * math.sin(t / 3)                    # 前後輕推
        w, h = 92, 64
        tip = (w - 6 + off - 8) if d > 0 else (6 - off + 8)
        tail = tip - 70 * d
        cy = h / 2
        pts = [tail, cy - 9, tip - 30 * d, cy - 9, tip - 30 * d, cy - 24, tip, cy,
               tip - 30 * d, cy + 24, tip - 30 * d, cy + 9, tail, cy + 9]
        cv.create_polygon(pts, fill=CAT_ORANGE, outline="#8A3B00", width=2)

    def _blink(self, t):
        """輸入欄閃橘框＋箭頭動畫，約 15 秒後停在亮框(視窗關掉才恢復)"""
        try:
            if not self.top.winfo_exists():
                return
        except tk.TclError:
            return
        if self._orig is None:
            try:
                self._orig = {k: self.inp.cget(k) for k in ("bg", "highlightthickness", "highlightbackground", "highlightcolor")}
            except tk.TclError:
                self._orig = {}
        on = (t // 6) % 2 == 0 or t > 300
        try:
            self.inp.config(highlightthickness=4, highlightbackground=CAT_ORANGE if on else "#FFD2A6",
                            highlightcolor=CAT_ORANGE, bg="#FFF4E8" if on else "white")
        except tk.TclError:
            pass
        self._draw_arrow(t)
        self.top.after(50, self._blink, t + 1)

    # ---------- 動作 ----------
    def _restore(self):
        if self._orig:
            try:
                self.inp.config(**self._orig)
            except tk.TclError:
                pass
        if self.arrow:
            try:
                self.arrow.destroy()
            except tk.TclError:
                pass
            self.arrow = None

    def _close(self, focus=False):
        _set_intro_hidden(self.flag_dir, self.hide_var.get())
        self._restore()
        for cb in (getattr(self, "cb_sect", None), getattr(self, "cb_road", None)):
            try:
                cb.tk.call("ttk::combobox::Unpost", cb)      # 原生下拉若開著先收起(放掉滑鼠鎖)
            except Exception:
                pass
        try:
            self.top.grab_release()
        except tk.TclError:
            pass
        try:
            self.top.destroy()
        except tk.TclError:
            pass
        if focus:
            try:
                self.inp.focus_set()
            except tk.TclError:
                pass

    def _try(self, text):
        """把範例填進輸入欄(去掉灰色提示字)，關掉介紹，游標停在輸入欄等使用者按 Enter"""
        self._close()
        try:
            self.inp.focus_set()
            self.inp.delete(0, "end")
            self.inp.insert(0, text)
            self.inp.config(fg="black")
            self.inp.icursor("end")
        except tk.TclError:
            pass


_PANEL = {"w": None}


def show_intro(root, input_widget, flag_dir, force=False, placeholder="", on_query=None, tab=None, guide=True,
               restore=None, case_info=None):
    """啟動時呼叫：沒勾過「啟動時不要自動跳出」才跳；force=True(打「說明」/按選單)一定跳。
    同時只開一個面板(已開著就換分頁並帶到最前面)"""
    if not force and intro_hidden(flag_dir):
        return None
    w = _PANEL["w"]
    try:
        if w is not None and w.top.winfo_exists():
            if tab:
                w.tab_var.set(dict(QuickIntro.TABS).get(tab, w.tab_var.get()))
            w.top.lift()
            return w
    except tk.TclError:
        pass
    _PANEL["w"] = QuickIntro(root, input_widget, flag_dir, placeholder, on_query, tab, guide, restore, case_info)
    return _PANEL["w"]


def close_panel():
    """主介面按了其他功能鈕 → 自動關掉萬用查詢面板(連同指向輸入欄的箭頭、閃爍外框)"""
    w = _PANEL["w"]
    try:
        if w is not None and w.top.winfo_exists():
            w._close()
    except tk.TclError:
        pass
    _PANEL["w"] = None


def open_panel(root, input_widget, flag_dir, on_query, tab=None, restore=None, case_info=None):
    """輸入欄「☰ 選單」或空白按 Enter：打開查詢面板(不閃箭頭)"""
    return show_intro(root, input_widget, flag_dir, True, "", on_query, tab, guide=False, restore=restore,
                      case_info=case_info)


def open_quick_case(root, text, data_json_path, base_dir, on_saved=None, log=None, input_widget=None):
    return QuickCaseWindow(root, text, data_json_path, base_dir, on_saved, log, input_widget)
