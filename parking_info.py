"""車位資訊：從謄本結構化的建物標示部抓出車位編號、持分、坪數(主程式轉物調表、物件編輯器、行銷簡報共用)

標示部的車位欄位(pdf_parser 產生)：
  第 1 組共有部分：「含停車位編號」「含停車位編號_1_2」…，面積在「共有部分面積」
  第 N 組共有部分：「含停車位編號N」「含停車位編號N_2」…，面積在「共有部分面積N」
  內容例：「地下三層3A ,權利範圍：100000分之280」「００２號」「B1-23」
車位坪數 = 該組共有部分面積 × 車位權利範圍 × 0.3025
"""
import glob
import json
import os
import re

_FW = str.maketrans("０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ－，：",
                    "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-,:")
_CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def _cn2int(s):
    s = str(s)
    if s.isdigit():
        return int(s)
    if s == "十":
        return 10
    if s.startswith("十"):
        return 10 + _CN.get(s[1:], 0)
    if "十" in s:
        a, _, b = s.partition("十")
        return _CN.get(a, 1) * 10 + (_CN.get(b, 0) if b else 0)
    return _CN.get(s, 0)


def parking_no(text):
    """「地下三層3A ,權利範圍…」→「B3-3A」；「地上一層第5號」→「1F-5」；「００２號」→「002號」"""
    s = str(text or "").translate(_FW)
    s = re.split(r"[,，]?\s*權利範圍", s)[0].strip(" ,、")
    m = re.search(r"地下\s*(?:第)?\s*([一二三四五六七八九十]+|\d+)\s*層\s*(?:第)?\s*([A-Za-z0-9\-]+)", s)
    if m:
        if re.match(r"^B\d", m.group(2), re.I):          # 「地下二層B2-15」編號本身已帶樓層
            return m.group(2).upper()
        return f"B{_cn2int(m.group(1))}-{m.group(2)}"
    m = re.search(r"(?:地上\s*)?(?:第)?\s*([一二三四五六七八九十]+|\d+)\s*層\s*(?:第)?\s*([A-Za-z0-9\-]+)", s)
    if m:
        return f"{_cn2int(m.group(1))}F-{m.group(2)}"
    m = re.match(r"\s*([A-Za-z0-9\-]+(?:號)?)", s)
    if m:
        return m.group(1).rstrip(",").strip()
    return s[:20]


def _ratio(text):
    m = re.search(r"權利範圍[:：]?\s*(\d+)\s*分之\s*(\d+)", str(text or "").translate(_FW))
    if m and int(m.group(1)) > 0:
        return f"{m.group(1)}分之{m.group(2)}", int(m.group(2)) / int(m.group(1))
    return "", None


def _area_m2(v):
    try:
        return float(str(v).replace(",", "").replace("平方公尺", "").strip())
    except (TypeError, ValueError):
        return None


def parse_bld(標示部):
    """一筆建物標示部 → [{"編號","持分","坪數"(float 或 None)}]"""
    out = []
    for key, val in (標示部 or {}).items():
        m = re.match(r"^含停車位編號(\d*)(?:_(\d+)_\d+)?$", key)
        if not m or not val:
            continue
        g = int(m.group(2) or m.group(1) or 1)
        area = _area_m2(標示部.get("共有部分面積" if g == 1 else f"共有部分面積{g}")) \
            or _area_m2(標示部.get("共有部分面積"))
        share, r = _ratio(val)
        ping = round(area * r * 0.3025, 2) if (area and r) else None
        no = parking_no(val)
        if no and not any(x["編號"] == no for x in out):
            out.append({"編號": no, "持分": share, "坪數": ping})
    return out


def parse_transcripts(trans):
    """結構化 JSON(list，每筆一份謄本) → 全部建物的車位清單"""
    out = []
    for t in trans if isinstance(trans, list) else []:
        if not isinstance(t, dict):
            continue
        for p in parse_bld(t.get("標示部") or {}):
            if not any(x["編號"] == p["編號"] for x in out):
                out.append(p)
    return out


def summary(items):
    """車位清單 → (編號字串, 持分字串, 總坪數字串)"""
    nos = "、".join(p["編號"] for p in items)
    shares = "、".join(p["持分"] for p in items if p["持分"])
    pings = [p["坪數"] for p in items if p["坪數"]]
    return nos, shares, (f"{sum(pings):.2f}" if pings else "")


def from_case(case_dir):
    """案件資料夾 4.其他相關/trans_*.json(最新一份)→ 車位清單；沒有就回 []"""
    try:
        files = sorted(glob.glob(os.path.join(case_dir, "4.其他相關", "trans_*.json")), key=os.path.getmtime)
        if not files:
            return []
        with open(files[-1], encoding="utf-8") as f:
            return parse_transcripts(json.load(f))
    except Exception:
        return []


def fill_missing(data, case_dir):
    """物調表 data 的「車位」沒有坪數/持分時，從案件的結構化 JSON 補上(不蓋掉使用者填的)。回傳是否有補"""
    park = data.get("車位")
    if not isinstance(park, dict):
        park = {}
    if park.get("坪數") and park.get("持分"):
        return False
    items = from_case(case_dir)
    if not items:
        return False
    nos, shares, ping = summary(items)
    changed = False
    for k, v in (("坪數", ping), ("持分", shares), ("編號", nos)):
        if v and not str(park.get(k, "")).strip():
            park[k] = v
            changed = True
    if nos and not str(park.get("有無", "")).strip():
        park["有無"] = "有，公設內"
        changed = True
    data["車位"] = park
    return changed
