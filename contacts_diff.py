"""通訊錄 ↔ 公司名冊 對照(管理者用)。
名冊＝阿生地圖每晚從公司後台抓的「現職」員工(店別/姓名/電話)；通訊錄.xlsx 要管理者自己維護(多了員工編號、職稱、離職標註)。
這裡把兩邊的差異列出來：新到職(要補編號)、已離職(要標離職)、換電話、改名、換店。只讀，不改 Excel。"""
import os
import re
import tkinter as tk

FAM = "Microsoft JhengHei"
_BAD_CODE = ("離職", "成仁", "離退")
_NOT_PERSON = re.compile(r"試用|總管理|專用|測試")


def _phone(v):
    d = re.sub(r"\D", "", str(v or "").split(".")[0])
    if len(d) == 9 and d.startswith("9"):      # Excel 把開頭的 0 吃掉了
        d = "0" + d
    return d


def _store(s):
    return (s or "").strip().replace("臺", "台")


def load_contacts(path):
    """讀通訊錄.xlsx(第 11 列是表頭：店別/職稱/姓名/電話/編號)→ [dict(name, store, phone, code, active)]"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    out, store, cols = [], "", None
    for row in ws.iter_rows(values_only=True):
        cells = ["" if v is None else str(v).strip() for v in row]
        if cols is None:
            if "姓名" in cells and "編號" in cells:
                cols = {k: cells.index(k) for k in ("店別", "職稱", "姓名", "電話", "編號") if k in cells}
            continue

        def g(k):
            i = cols.get(k)
            return cells[i] if i is not None and i < len(cells) else ""
        if g("店別"):
            store = g("店別")
        name = g("姓名")
        if not name or "＜" in name or "＞" in name or "姓名" in name or "：" in name:
            continue
        code = g("編號")
        out.append(dict(name=name, store=_store(store), phone=_phone(g("電話")), code=code, title=g("職稱"),
                        active=bool(code) and code not in _BAD_CODE))
    wb.close()
    return out


def fetch_roster():
    """阿生地圖的公司名冊 {姓名: [{store, phone, line}]}；沒登入/不是公司成員回 None"""
    import landmap_login
    if not landmap_login.token():
        return None
    j = landmap_login.call("/api/fuzutong/staff")
    if (j or {}).get("status") != "OK":
        return None
    return j.get("staff") or {}


def compare(contacts, roster):
    """回 dict(new, gone, renamed, phone, store, rehired)，每項是 list[dict]"""
    ros = {}
    for nm, lst in (roster or {}).items():
        if not nm or _NOT_PERSON.search(nm) or not lst:
            continue
        r = lst[0]
        ros[nm] = dict(name=nm, store=_store(r.get("store")), phone=_phone(r.get("phone")))
    by_name = {c["name"]: c for c in contacts}
    by_phone = {}
    for c in contacts:
        if c["phone"] and c["name"] not in ros:
            by_phone.setdefault(c["phone"], c)
    d = dict(new=[], gone=[], renamed=[], phone=[], store=[], rehired=[])
    renamed_from = set()
    for nm, r in ros.items():
        c = by_name.get(nm)
        if c is None:
            old = by_phone.get(r["phone"]) if r["phone"] else None
            if old and old["name"] not in renamed_from:      # 電話一樣、名字不同＝改名(或錯字)
                d["renamed"].append(dict(old=old["name"], new=nm, store=r["store"], phone=r["phone"]))
                renamed_from.add(old["name"])
            else:
                d["new"].append(r)
            continue
        if c["code"] in _BAD_CODE:                           # 編號空白的是行政人員，不算
            d["rehired"].append(dict(r, code=c["code"]))
            continue
        if c["phone"] and r["phone"] and c["phone"] != r["phone"]:
            d["phone"].append(dict(name=nm, old=c["phone"], new=r["phone"]))
        if c["store"] and r["store"] and c["store"] != r["store"]:
            d["store"].append(dict(name=nm, old=c["store"], new=r["store"]))
    for c in contacts:
        if c["active"] and c["store"] != "總公司管理部" and c["name"] not in ros and c["name"] not in renamed_from:
            d["gone"].append(dict(name=c["name"], store=c["store"], code=c["code"], phone=c["phone"]))
    for r in list(d["new"]):                                 # 同店、名字只差一個字＝打錯字(電話可能也換了)
        for g in d["gone"]:
            if g["store"] == r["store"] and len(g["name"]) == len(r["name"]) and \
                    sum(a != b for a, b in zip(g["name"], r["name"])) == 1:
                d["renamed"].append(dict(old=g["name"], new=r["name"], store=r["store"], phone=r["phone"],
                                         old_phone=g["phone"], code=g["code"]))
                d["new"].remove(r)
                d["gone"].remove(g)
                break
    return d


def count(d):
    return sum(len(v) for v in (d or {}).values())


def summary(d):
    parts = [("新到職", "new"), ("離職", "gone"), ("改名", "renamed"), ("換電話", "phone"), ("換店", "store"),
             ("標離職但還在職", "rehired")]
    return "、".join(f"{t} {len(d[k])}" for t, k in parts if d.get(k))


def show_window(parent, d, xlsx_path=None):
    """大字清單視窗：照著改 Excel 就好"""
    tw = tk.Toplevel(parent)
    tw.title("通訊錄 ↔ 公司名冊 對照")
    tw.configure(bg="#FFFFFF")
    tw.geometry("760x640")
    n = count(d)
    tk.Label(tw, text=("通訊錄跟公司名冊一致 ✓" if n == 0 else f"有 {n} 處不同，請照著改通訊錄 Excel，再按「上傳通訊錄」"),
             font=(FAM, 15, "bold"), bg="#FFFFFF", fg="#1C1C1E", wraplength=720, justify="left").pack(anchor="w", padx=18, pady=(14, 2))
    tk.Label(tw, text="名冊＝阿生地圖每晚從公司後台抓的現職員工；公司後台最先更新，兩邊不同時一律以名冊為準", font=(FAM, 11), bg="#FFFFFF",
             fg="#3A3A3C").pack(anchor="w", padx=18)
    fr = tk.Frame(tw, bg="#FFFFFF")
    fr.pack(fill=tk.BOTH, expand=True, padx=14, pady=8)
    sb = tk.Scrollbar(fr)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    tx = tk.Text(fr, font=(FAM, 13), wrap="word", yscrollcommand=sb.set, relief="flat", bg="#FAFAFA",
                 fg="#1C1C1E", padx=10, pady=8, spacing1=3)
    tx.pack(fill=tk.BOTH, expand=True)
    sb.config(command=tx.yview)
    tx.tag_configure("h", font=(FAM, 14, "bold"), foreground="#7A3E00", spacing1=10)

    def sec(title, items, fmt):
        if not items:
            return
        tx.insert("end", f"{title}（{len(items)}）\n", "h")
        for x in items:
            tx.insert("end", "　• " + fmt(x) + "\n")
    sec("🆕 新到職：通訊錄沒有 → 新增一列並補上員工編號", d["new"],
        lambda x: f"{x['name']}　{x['store']}　{x['phone']}")
    sec("👋 已離職：名冊沒有了 → 編號欄改成「離職」", d["gone"],
        lambda x: f"{x['name']}　{x['store']}　{x['code']}")
    sec("✏️ 改名／錯字：同一人名字不同 → 一律照名冊改", d["renamed"],
        lambda x: f"{x['old']} → {x['new']}　{x['store']}" +
        (f"　電話也改：{x['old_phone']} → {x['phone']}" if x.get("old_phone") and x["old_phone"] != x["phone"]
         else f"　{x['phone']}"))
    sec("📞 換電話", d["phone"], lambda x: f"{x['name']}　{x['old']} → {x['new']}")
    sec("🏪 換店", d["store"], lambda x: f"{x['name']}　{x['old']} → {x['new']}")
    sec("↩️ 通訊錄標離職，但名冊還在職 → 補回員工編號", d["rehired"],
        lambda x: f"{x['name']}　{x['store']}　（編號欄現在：{x['code'] or '空白'}）")
    tx.config(state="disabled")
    row = tk.Frame(tw, bg="#FFFFFF")
    row.pack(pady=(0, 14))
    if xlsx_path and os.path.exists(xlsx_path):
        tk.Button(row, text="📂 打開通訊錄 Excel", font=(FAM, 12, "bold"), bg="#55b092", fg="#10261d",
                  command=lambda: os.startfile(xlsx_path)).pack(side=tk.LEFT, padx=6)
    tk.Button(row, text="關閉", font=(FAM, 12), command=tw.destroy).pack(side=tk.LEFT, padx=6)
    return tw


if __name__ == "__main__":     # 命令列測試：python contacts_diff.py 通訊錄.xlsx
    import sys
    import json
    cs = load_contacts(sys.argv[1] if len(sys.argv) > 1 else "通訊錄.xlsx")
    ro = fetch_roster()
    print("通訊錄", len(cs), "在職", sum(c["active"] for c in cs), "名冊", None if ro is None else len(ro))
    if ro is not None:
        dd = compare(cs, ro)
        print(summary(dd))
        print(json.dumps(dd, ensure_ascii=False, indent=1))
