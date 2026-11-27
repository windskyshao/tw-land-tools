# -*- coding: utf-8 -*-
"""物件周邊「有出處的事實」：給 ✨ AI 行銷文(訴求重點)用，也給行銷簡報讀重要設施 PDF。

2026-11-27 使用者：AI 寫的訴求重點沒有新資料、還會自己編(「位於綠園新村商圈」「交通動線便捷」)。
→ 把下面這些查得到、有出處的資料交給 AI，並規定只能用這些寫環境/建設/交通：
  ① 重要設施：案件「1.基本資料/03_重要設施*.pdf」(富住通/圖霸 map8 或舊台灣地圖網)的清單＋距離(不含嫌惡設施)
  ② 重大建設新聞：阿生地圖新聞(/data/news.json，約近兩週)，同縣市、離物件幾公里內或寫到同一區的
     重大建設／產業設廠／都更
"""
import glob
import json
import math
import os
import re
import urllib.request


# ---------------------------------------------------------------- ① 重要設施 PDF
def facility_rows(doc):
    """重要設施 PDF(fitz 文件) → [(編號, 名稱, 距離)]。
    清單是一般文字(不是表格物件)：依第一頁表頭「名稱/重要設施」「地址」「距離」的位置切欄。
    表頭：台灣地圖網＝「名稱」；富住通(圖霸 map8，2026-11-27 起)＝「重要設施」「嫌惡設施」。
    嫌惡設施表(高壓電塔、宮廟…)行銷用不列：那張表以下(含之後的頁)都略過。
    以編號欄的數字當每一列的中心；名稱太長會折成上下兩行 → 中心上下 12pt 內的名稱字都算這一列"""
    pg = doc[0]
    hw = {w[4]: w for w in pg.get_text("words")
          if w[4] in ("名稱", "地址", "重要設施", "嫌惡設施") or w[4].startswith("距離")}
    nh = next((hw[k] for k in ("重要設施", "嫌惡設施") if k in hw), None)
    x_name = hw["名稱"][0] - 45 if "名稱" in hw else (nh[0] - 20 if nh else 100)
    x_addr = hw["地址"][0] - 100 if "地址" in hw else 215
    x_dist = next((w[0] for k, w in hw.items() if k.startswith("距離")), 440) - 5
    rows, stop_y = [], {}
    for pi, p in enumerate(doc):
        if stop_y:                                   # 前一頁已經進到嫌惡設施表 → 後面都是它
            break
        words = p.get_text("words")
        for w in words:
            if w[4] == "嫌惡設施" and pi not in stop_y:
                stop_y[pi] = w[1] - 2
        nos = sorted((w for w in words if w[0] < x_name and w[4].isdigit()), key=lambda w: w[1])
        for i, nw in enumerate(nos):
            y0 = nw[1]
            if pi in stop_y and y0 >= stop_y[pi]:
                continue
            lo = (nos[i - 1][1] + y0) / 2 if i else y0 - 12
            hi = (nos[i + 1][1] + y0) / 2 if i + 1 < len(nos) else y0 + 12
            lo, hi = max(lo, y0 - 12), min(hi, y0 + 12)
            near = sorted((w for w in words if lo <= w[1] < hi), key=lambda w: (round(w[1]), w[0]))
            name = "".join(w[4] for w in near if x_name <= w[0] < x_addr)
            dist = next((w[4] for w in near if w[0] >= x_dist and re.fullmatch(r"[\d.]+", w[4])), "")
            if name:
                rows.append((nw[4], name, dist))
    return rows


def facility_facts(case_dir, limit=10):
    """案件的重要設施 PDF → (句子清單, 出處)；沒有檔案回 ([], "")"""
    if not case_dir:
        return [], ""
    pdfs = sorted(glob.glob(os.path.join(case_dir, "1.基本資料", "*重要設施*.pdf")), key=os.path.getmtime)
    if not pdfs:
        return [], ""
    import fitz
    doc = fitz.open(pdfs[-1])
    try:
        rows = facility_rows(doc)
        m = re.search(r"週邊\s*(\d+)\s*公尺", doc[0].get_text())
    finally:
        doc.close()
    if not rows:
        return [], ""
    seen, out = set(), []
    for _no, name, dist in sorted(rows, key=lambda r: float(r[2]) if r[2] else 9e9):
        if name in seen:
            continue
        seen.add(name)
        out.append(f"{name}（約 {dist} 公尺）" if dist else name)
        if len(out) >= limit:
            break
    src = f"重要設施查詢（{os.path.basename(pdfs[-1])}" + (f"，查詢範圍週邊 {m.group(1)} 公尺" if m else "") + "）"
    return out, src


# ---------------------------------------------------------------- ② 阿生地圖重大建設新聞
NEWS_CATS = ("重大建設", "產業設廠", "都更")


def _km(a, b, c, d):
    r = 6371.0
    p1, p2 = math.radians(a), math.radians(c)
    dp, dl = p2 - p1, math.radians(d - b)
    h = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(h))


def pick_news(items, lat=None, lng=None, city="", area="", max_km=6.0, limit=4):
    """新聞清單 → 跟這個物件有關的重大建設新聞(新到舊)。
    條件：分類是重大建設/產業設廠/都更；同縣市；而且(有座標的離物件 max_km 公里內，或標題/地點寫到同一個區)"""
    norm = lambda s: (s or "").replace("台", "臺")
    city, area = norm(city), norm(area)
    out = []
    for it in items or []:
        if it.get("cat") not in NEWS_CATS:
            continue
        if city and norm(it.get("county")) and norm(it.get("county")) != city:
            continue
        text = norm(it.get("place", "") + it.get("title", "") + it.get("sum", ""))
        near = False
        if lat is not None and lng is not None:
            for loc in it.get("locs") or []:
                try:
                    if _km(float(lat), float(lng), float(loc["lat"]), float(loc["lng"])) <= max_km:
                        near = True
                        break
                except Exception:
                    pass
        if area and area in text:
            near = True
        if area and len(area) >= 3 and area[:-1] in text:        # 「仁武區」→ 標題常只寫「仁武」
            near = True
        if near:
            out.append(it)
    out.sort(key=lambda x: x.get("dt", ""), reverse=True)
    return out[:limit]


def news_facts(lat, lng, city, area, base="https://map.windsky-sky.com", timeout=20):
    """→ (句子清單, 出處)；連不上回 ([], "")"""
    try:
        req = urllib.request.Request(base + "/data/news.json", headers={"User-Agent": "Mozilla/5.0 LandRegDesktop"})
        items = json.loads(urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8")).get("items") or []
    except Exception:
        return [], ""
    got = pick_news(items, lat, lng, city, area)
    lines = [f"{(it.get('dt') or '')[:10]}【{it.get('src', '')}】{it.get('title', '')}"
             + (f"：{it['sum']}" if it.get("sum") else "") for it in got]
    return lines, ("阿生地圖房產新聞（近期重大建設／產業設廠／都更，同縣市且在附近或寫到同一區）" if lines else "")


def facts_block(sections):
    """[(標題, 句子清單, 出處)] → 給 AI 的「周邊資料」文字；全部沒有回 "" """
    parts = []
    for title, lines, src in sections:
        if lines:
            parts.append(f"【{title}】（出處：{src}）\n" + "\n".join(f"- {x}" for x in lines))
    return "\n\n".join(parts)
