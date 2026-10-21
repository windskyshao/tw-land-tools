# -*- coding: utf-8 -*-
"""阿生地圖(map.windsky-sky.com)查詢：地址/地號 → 縣市、行政區、段、地號、座標、建號、社區名
主程式「萬用查詢欄」用。全部免登入、免費(資料源：國土測繪中心、內政部實價登錄)。
連不上時拋 LandmapError，由呼叫端提示使用者改用原本方式。"""
import json
import re
import ssl
import urllib.parse
import urllib.request
import urllib.error

BASE = "https://map.windsky-sky.com"
_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0 Safari/537.36 LandRegDesktop"

# 地政縣市代碼
CITY_CODES = {
    "臺北市": "A", "臺中市": "B", "基隆市": "C", "臺南市": "D", "高雄市": "E", "新北市": "F",
    "宜蘭縣": "G", "桃園市": "H", "嘉義市": "I", "新竹縣": "J", "苗栗縣": "K", "南投縣": "M",
    "彰化縣": "N", "新竹市": "O", "雲林縣": "P", "嘉義縣": "Q", "屏東縣": "T", "花蓮縣": "U",
    "臺東縣": "V", "金門縣": "W", "澎湖縣": "X", "連江縣": "Z",
}
_FW = str.maketrans("０１２３４５６７８９－", "0123456789-")


class LandmapError(Exception):
    pass


def _get(path, timeout=25, **params):
    url = BASE + path + ("?" + urllib.parse.urlencode(params) if params else "")
    req = urllib.request.Request(url, headers={"User-Agent": _UA, "Accept": "application/json"})
    ctxs = [None]
    try:
        import certifi
        ctxs.append(ssl.create_default_context(cafile=certifi.where()))
    except Exception:
        pass
    last = None
    for ctx in ctxs:
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            raise LandmapError(f"阿生地圖回應錯誤（HTTP {e.code}）")
        except urllib.error.URLError as e:
            last = e
            if not isinstance(getattr(e, "reason", None), ssl.SSLError):
                break
        except Exception as e:
            last = e
            break
    raise LandmapError(f"連不上阿生地圖：{last}")


def norm_city(s):
    s = (s or "").replace("台", "臺")
    return s


_TOWNS_CACHE = {}


def _towns(code):
    if code not in _TOWNS_CACHE:                   # 批次查詢時同縣市只抓一次
        j = _get("/api/towns", city=code)
        _TOWNS_CACHE[code] = j.get("towns") or []
    return _TOWNS_CACHE[code]


def _nlsc_xml(path):
    """國土測繪中心免申請 API(回 XML)；阿生地圖連不上時的備援"""
    import xml.etree.ElementTree as ET
    req = urllib.request.Request("https://api.nlsc.gov.tw/other/" + path, headers={"User-Agent": _UA})
    with urllib.request.urlopen(req, timeout=20) as r:
        return ET.fromstring(r.read())


def list_towns(city):
    """縣市 → [(行政區名, 代碼)]。先問阿生地圖，失敗改問國土測繪中心"""
    code = CITY_CODES.get(norm_city(city), "")
    if not code:
        return []
    try:
        return [(t.get("name", ""), t.get("id", "")) for t in _towns(code) if t.get("name")]
    except LandmapError:
        pass
    try:
        root = _nlsc_xml(f"ListTown1/{code}")
        out = [(n.findtext("townname", ""), n.findtext("towncode01", "")) for n in root.iter("townItem")]
        _TOWNS_CACHE[code] = [{"name": a, "id": b} for a, b in out]
        return out
    except Exception as e:
        raise LandmapError(f"抓不到行政區清單：{e}")


_SECT_CACHE = {}


def list_sections(city, town_id):
    """縣市＋行政區代碼 → [段小段名]。先問阿生地圖，失敗改問國土測繪中心"""
    code = CITY_CODES.get(norm_city(city), "")
    key = (code, town_id)
    if key in _SECT_CACHE:
        return _SECT_CACHE[key]
    names = []
    try:
        names = [s.get("name", "") for s in (_get("/api/sections", city=code, area=town_id).get("sections") or [])]
    except LandmapError:
        try:
            root = _nlsc_xml(f"ListLandSection/{code}/{town_id}")
            names = [n.findtext("sectstr", "") for n in root if n.findtext("sectstr")]
        except Exception as e:
            raise LandmapError(f"抓不到地段清單：{e}")
    names = [n for n in dict.fromkeys(names) if n]
    _SECT_CACHE[key] = names
    return names


_BATCH_SEP = re.compile(r"[、，,；;\n]+")


def split_batch(text):
    """批次輸入拆成多筆完整查詢字串。只寫號碼的那筆沿用前一筆的前段：
    「竹北段一小段2096、2097、2100-1」→ 2097 補成「竹北段一小段2097」
    「和平二路240之1號12樓、240之2號12樓」→ 補成「和平二路240之2號12樓」"""
    out = []
    for piece in _BATCH_SEP.split(text or ""):
        p = piece.strip().translate(_FW).replace(" ", "")
        if not p:
            continue
        if out and re.match(r"^\d", p):
            m = re.match(r"^(.*[段路街道巷弄])(?=\d)", out[-1])     # 前一筆最後一個「段/路/巷…」之前
            if m:
                p = m.group(1) + p
        out.append(p)
    return out


def resolve_with_context(text, prev=None, progress=None):
    """批次用：這筆沒寫行政區(地號會查不到)時，沿用前一筆查到的縣市＋行政區再試一次"""
    try:
        return resolve(text, progress=progress)
    except LandmapError as e:
        t = norm_city(text)
        if prev and prev.get("area") and prev["area"] not in t and "行政區" in str(e):
            base = t[len(prev["city"]):] if t.startswith(prev["city"]) else t
            return resolve(prev["city"] + prev["area"] + base, progress=progress)
        raise


def _split_city_town(text):
    """從地址/地號字串切出 縣市、行政區(用阿生地圖的正式行政區清單比對，不用正則猜)。回傳 (city, code, town_name, town_id, rest)"""
    t = norm_city(text.strip())
    city = next((c for c in CITY_CODES if t.startswith(c)), "")
    if not city:
        return "", "", "", "", t
    code = CITY_CODES[city]
    rest = t[len(city):]
    best = None
    for tw in _towns(code):
        name = tw.get("name", "")
        if rest.startswith(name) and (best is None or len(name) > len(best["name"])):
            best = tw
    if best:
        return city, code, best["name"], best["id"], rest[len(best["name"]):]
    return city, code, "", "", rest


def looks_like_lot(text):
    """「…段…號」且沒有路/街/巷/弄 → 視為地號(例：前鎮區竹北段一小段2096、高雄市楠梓區芎蕉段13-1)"""
    t = text.translate(_FW).replace(" ", "")
    return bool(re.search(r"段[^路街巷弄]*?\d+(-\d+)?(地號)?$", t)) and not re.search(r"[路街巷弄]", t)


def map_url(lat, lng, door=""):
    q = {"lat": f"{lat:.6f}", "lng": f"{lng:.6f}"}
    if door:
        q["door"] = door
    return BASE + "/?" + urllib.parse.urlencode(q)


def _community(code, lat, lng, sect, no, door=""):
    try:
        j = _get("/api/lvr_communityat", city=code, lat=lat, lng=lng, sect=sect, no=no, door=door)
        return j.get("com") or ""
    except LandmapError:
        return ""


def community_for(entry):
    """data.json 的一筆 → 社區名(查不到或連不上回 "")。
    已有 community 欄位(快速建立案件存的)就直接用；否則用座標＋地段地號問阿生地圖(依歷年實價登錄比對)"""
    if entry.get("community"):
        return entry["community"]
    try:
        lat, lng = (float(x) for x in str(entry.get("coordinates", "")).split(",")[:2])
        code = CITY_CODES.get(norm_city(entry.get("city", "")), "")
        if not code:
            return ""
        return _community(code, lat, lng, entry.get("section", ""), entry.get("lot_number", ""), entry.get("door", ""))
    except Exception:
        return ""


def roads_for(entries, entry=None):
    """data.json 的門牌(door 與 builds[].door) → 路名清單(例 廟前路、建國路三段)，去重照順序。
    entry 有給就只看同縣市+行政區+地段的那幾筆。沒有社區名的透天/店面用路名查周邊成交"""
    out = []
    for e in entries or []:
        if not isinstance(e, dict):
            continue
        if entry and (e.get("city"), e.get("area"), e.get("section")) != (entry.get("city"), entry.get("area"), entry.get("section")):
            continue
        doors = [e.get("door", "")] + [b.get("door", "") for b in (e.get("builds") or []) if isinstance(b, dict)]
        for d in doors:
            d = str(d or "").replace(" ", "").translate(_FW)
            for pre in (e.get("city", ""), e.get("area", "")):
                if pre and d.startswith(pre):
                    d = d[len(pre):]
            d = re.sub(r"^[^\d路街段巷弄號]{1,4}[里村](\d+鄰)?", "", d)
            m = re.match(r"(.{1,8}?(?:路|街|大道)(?:[一二三四五六七八九十]段)?)", d)
            if m and m.group(1) not in out:
                out.append(m.group(1))
    return out


def _parcel_at(lat, lng):
    j = _get("/api/parcel", lat=lat, lng=lng)
    if j.get("status") != "OK":
        raise LandmapError("這個位置查不到地號")
    return j


_progress = None


def _step(n):
    if _progress:
        try:
            _progress(n)
        except Exception:
            pass


def resolve(text, default_city="高雄市", progress=None):
    """主入口：自動判斷地址或地號。回傳 dict(見下方 result)"""
    raw = (text or "").strip()
    if not raw:
        raise LandmapError("請輸入地址或地號")
    t = raw.translate(_FW).replace(" ", "")
    if not any(t.startswith(c) or t.startswith(c.replace("臺", "台")) for c in CITY_CODES):
        t = default_city + t
    global _progress
    _progress = progress
    try:
        return _resolve_lot(t, raw) if looks_like_lot(t) else _resolve_address(t, raw)
    finally:
        _progress = None


def _base_result(kind, raw):
    return {"kind": kind, "input": raw, "city": "", "city_code": "", "area": "", "section": "",
            "sect_code": "", "office": "", "lot_number": "", "lat": None, "lng": None,
            "door": "", "builds": [], "build_note": "", "door_groups": [], "community": "", "map_url": ""}


def _builds_from(m):
    out = []
    for b in m.get("matches") or []:
        out.append({"build_no": b.get("disp", ""), "door": (b.get("door") or "").translate(_FW),
                    "floor": b.get("floor", ""), "area_m2": b.get("area", ""),
                    "use": b.get("mainuse", ""), "attached": b.get("attached") or [],
                    "main_floors": b.get("main_floors_free") or []})     # 便民系統分層明細[{floor,m2}]
    return out


def _door_variants(t):
    """定位用門牌：完整 → 去樓層 → 去「之N」(帶之N/樓層的門牌，定位常落到隔壁地號)"""
    v = [t]
    no_floor = re.sub(r"(\d+(?:之\d+)?號)(之\d+)?[\d一二三四五六七八九十]+樓.*$", r"\1\2", t)
    v.append(no_floor)
    v.append(re.sub(r"(\d+號)之\d+$", r"\1", no_floor))          # 236號之2 → 236號
    v.append(re.sub(r"(\d+)之\d+號$", r"\1號", no_floor))        # 236之2號 → 236號
    out = []
    for x in v:
        if x and x not in out:
            out.append(x)
    return out


def _official_door(door):
    """門牌轉成地政登記寫法：236號之2 → 236之2號(樓層保留)"""
    return re.sub(r"(\d+)號之(\d+)", r"\1之\2號", door)


def _resolve_address(t, raw):
    from concurrent.futures import ThreadPoolExecutor
    res = _base_result("address", raw)
    t = _official_door(t)
    _c2, _code2, _t2, _id2, door = _split_city_town(t)
    door = re.sub(r"^[^\d路街段巷弄號]{1,4}[里村](\d+鄰)?", "", door)
    res["door"] = door
    variants = _door_variants(t)

    def locate(q):                      # 門牌 → (候選, 地號資訊)
        try:
            j = _get("/api/address", q=q)
            cands = [c for c in (j.get("results") or []) if c.get("type") == "門牌"]
            if not cands:
                return None
            c = cands[0]
            lat, lng = float(c["lat"]), float(c["lng"])
            return c, _parcel_at(lat, lng), lat, lng
        except LandmapError:
            return None

    def match(p):
        try:
            return _get("/api/matchdoor", timeout=40, city=p.get("cityCode"), sect=p.get("sectNo"),
                        landno=p.get("landNo"), door=door, office=p.get("office"))
        except LandmapError:
            return {}

    with ThreadPoolExecutor(4) as ex:
        located = [x for x in ex.map(locate, variants) if x]          # 依寫法順序
        _step(1)
        _step(2)
        if not located:
            raise LandmapError("找不到這個地址，請確認縣市、行政區、路名與門牌號")
        uniq, order = {}, []
        for x in located:
            key = (x[1].get("sectNo"), x[1].get("landNo"))
            if key not in uniq:
                uniq[key] = x
                order.append(key)
        matches = dict(zip(order, ex.map(lambda k: match(uniq[k][1]), order)))
        _step(3)
    chosen = next((k for k in order if (matches.get(k) or {}).get("matches")), None)
    c, p, lat, lng = uniq[chosen or order[0]]
    city, code, town, _tid, _rest = _split_city_town(c.get("name", "") or t)
    if not town:
        city, code, town, _tid, _rest = _split_city_town(t)
    res.update(lat=lat, lng=lng, city=norm_city(p.get("cityName") or city), city_code=p.get("cityCode") or code,
               area=town, section=p.get("sectName", ""), sect_code=p.get("sectNo", ""), office=p.get("office", ""),
               lot_number=p.get("landNo", ""))
    if chosen:
        res["builds"] = _builds_from(matches[chosen])
        res["build_note"] = matches[chosen].get("note", "")
    else:
        res["build_note"] = "這個門牌在此地號上找不到對應建號（可能是透天新建未登記、或門牌與地號不一致）"
    res["community"] = _community(res["city_code"], res["lat"], res["lng"], res["section"], res["lot_number"], door)
    _step(4)
    res["map_url"] = map_url(res["lat"], res["lng"], door)
    return res


def _resolve_lot(t, raw):
    res = _base_result("lot", raw)
    city, code, town, tid, rest = _split_city_town(t)
    if not (city and town):
        raise LandmapError("地號請包含行政區，例如：前鎮區竹北段一小段2096")
    m = re.match(r"(.+?段(?:.+?小段)?)(\d+)(?:-(\d+))?(?:地號)?$", rest)
    if not m:
        raise LandmapError("看不懂地號格式，例如：前鎮區竹北段一小段2096 或 芎蕉段13-1")
    sect_name, mother, child = m.group(1), m.group(2), m.group(3)
    secs = (_get("/api/sections", city=code, area=tid).get("sections") or [])
    _step(1)
    sec = next((s for s in secs if s.get("name") == sect_name), None)
    if not sec:
        raise LandmapError(f"{town}找不到「{sect_name}」，請確認段名（含小段）")
    lot = mother + (f"-{child}" if child and int(child) else "")
    res.update(city=city, city_code=code, area=town, section=sect_name, sect_code=sec["id"],
               office=sec.get("office", ""), lot_number=lot)
    landno = f"{int(mother):04d}{int(child or 0):04d}"
    pb = _get("/api/parcelbatch", city=code, office=res["office"], sect=res["sect_code"], landno=landno)
    _step(2)
    if pb.get("status") != "OK" or pb.get("lat") is None:
        raise LandmapError("查不到這個地號的位置（可能地號不存在或已合併/分割）")
    res["lat"], res["lng"] = float(pb["lat"]), float(pb["lng"])
    from concurrent.futures import ThreadPoolExecutor
    with ThreadPoolExecutor(2) as ex:
        f_com = ex.submit(_community, code, res["lat"], res["lng"], sect_name, lot)
        try:
            g = _get("/api/parcelbuildings", timeout=40, city=code, sect=res["sect_code"], landno=landno)
            res["door_groups"] = g.get("groups") or []
            res["build_note"] = f"此地號上共有 {g.get('total', 0)} 個建號" if g.get("total") else "此地號上沒有登記建物（純土地）"
        except LandmapError:
            res["build_note"] = "建物資料暫時查不到"
        _step(3)
        res["community"] = f_com.result()
        _step(4)
    res["map_url"] = map_url(res["lat"], res["lng"])
    return res


def door_group_text(g):
    """parcelbuildings 的門牌群組 → 「和平二路232~240號」「和平二路236號之1~之2」"""
    pre, lo, hi = g.get("pre", ""), g.get("lo"), g.get("hi")
    s = f"{pre}{lo}號" if lo == hi else f"{pre}{lo}~{hi}號"
    if g.get("sub"):
        zlo, zhi = g.get("zlo"), g.get("zhi")
        s = f"{pre}{lo}號之{zlo}" + (f"~之{zhi}" if zhi != zlo else "")
    return f"{s}（{g.get('n', 0)} 戶）"


def lookup_door_in_lot(res, door):
    """地號查詢後，使用者補打門牌(含樓層)→ 在這個地號找建號"""
    m = _get("/api/matchdoor", timeout=40, city=res["city_code"], sect=res["sect_code"],
             landno=f"{int(res['lot_number'].split('-')[0]):04d}{int((res['lot_number'].split('-') + ['0'])[1]):04d}",
             door=door, office=res["office"])
    return _builds_from(m), m.get("note", "")
