"""高雄國小學區 PDF(115年國小學區.pdf；build_school_district_data.build_kh_elementary_layout 用)(無框線表格)依座標切欄：編號 | 學校 | 行政區 | 學區文字。
學區文字以「。」分段(每個行政區一段)；行政區標籤垂直置中在它那段；
學校名稱垂直置中在整列(連續幾段)→ 用 DP 把段落分給學校(列中心離校名中心最近)。"""
import re
import sys
import pdfplumber

X_SCHOOL, X_DIST, X_TEXT = 130, 230, 300


def lines_of(page):
    words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
    rows = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        for r in rows:
            if abs(r["top"] - w["top"]) < 3:
                r["words"].append(w)
                break
        else:
            rows.append({"top": w["top"], "bottom": w["bottom"], "words": [w]})
    return sorted(rows, key=lambda r: r["top"])


def parse(pdf_path):
    blocks, schools = [], []     # blocks: {page, top, bottom, text, dist}; schools: {page, mid, name}
    cur = None
    with pdfplumber.open(pdf_path) as pdf:
        for pi, page in enumerate(pdf.pages):
            dists, names = [], []
            for r in lines_of(page):
                ws = r["words"]
                if r["top"] > 765 and re.fullmatch(r"\d+", "".join(w["text"] for w in ws)):
                    continue
                if any(w["text"] in ("學校名稱", "行政區") for w in ws) or any("劃分一覽表" in w["text"] for w in ws):
                    continue
                # 備註(1 .雄峰里自106學年度…)：編號欄出現非數字、或左邊欄位有長串文字 → 不是表格列
                if any(w["x0"] < X_SCHOOL and not w["text"].isdigit() for w in ws) or                         any(w["x0"] < X_TEXT and w["x1"] - w["x0"] > 110 for w in ws):
                    continue
                txt = "".join(w["text"] for w in ws if w["x0"] >= X_TEXT)
                sch = "".join(w["text"] for w in ws if X_SCHOOL <= w["x0"] < X_DIST)
                dst = "".join(w["text"] for w in ws if X_DIST <= w["x0"] < X_TEXT)
                mid = (r["top"] + r["bottom"]) / 2
                if sch:
                    names.append((mid, sch))
                if dst:
                    dists.append((mid, dst))
                if txt:
                    if cur is None or cur["page"] != pi or cur["closed"]:
                        if cur is not None and not cur["closed"]:
                            cur["cont"] = True          # 跨頁接續
                        cur = {"page": pi, "top": r["top"], "bottom": r["bottom"], "text": txt, "dist": [], "lines": [],
                               "closed": False, "cont_from_prev": bool(blocks) and not blocks[-1]["closed"]}
                        blocks.append(cur)
                    else:
                        cur["text"] += txt
                        cur["bottom"] = r["bottom"]
                    cur["lines"].append((r["top"], r["bottom"], txt))
                    if txt.rstrip().endswith("。"):
                        cur["closed"] = True
            # 校名可能分兩行(紅毛港 / 國小)→ 相鄰合併
            merged = []
            for mid, nm in names:
                if merged and mid - merged[-1][0] < 22 and (not merged[-1][1].endswith(("國小", "分校")) or nm.endswith("分校")):
                    merged[-1] = ((merged[-1][0] + mid) / 2, merged[-1][1] + nm)
                else:
                    merged.append((mid, nm))
            for mid, nm in merged:
                schools.append({"page": pi, "mid": mid, "name": nm})
            pb = [b for b in blocks if b["page"] == pi]
            for mid, d in dists:
                if not d.endswith("區"):
                    print("!! 怪行政區", pi + 1, d, file=sys.stderr)
                best = min(pb, key=lambda b: 0 if b["top"] - 2 <= mid <= b["bottom"] + 2 else
                           min(abs(mid - b["top"]), abs(mid - b["bottom"]))) if pb else None
                if best is not None:
                    best["dist"].append((mid, d))
    return blocks, schools


def split_multi(blocks):
    out = []
    for b in blocks:
        ds = sorted(b["dist"])
        if len(ds) < 2:
            out.append(b)
            continue
        L = b["lines"]
        n, m = len(L), len(ds)
        INF = float("inf")
        dp = [[INF] * (m + 1) for _ in range(n + 1)]
        bk = [[None] * (m + 1) for _ in range(n + 1)]
        dp[0][0] = 0
        for j in range(1, m + 1):
            for i in range(j, n + 1):
                for k in range(j - 1, i):
                    if dp[k][j - 1] == INF:
                        continue
                    c = abs((L[k][0] + L[i - 1][1]) / 2 - ds[j - 1][0])
                    if dp[k][j - 1] + c < dp[i][j]:
                        dp[i][j], bk[i][j] = dp[k][j - 1] + c, k
        if dp[n][m] == INF:
            print("!! 拆不開", ds, b["text"][:30], file=sys.stderr)
            out.append(b)
            continue
        parts, i = [], n
        for j in range(m, 0, -1):
            k = bk[i][j]
            parts.append((ds[j - 1], L[k:i]))
            i = k
        for idx, (d, ls) in enumerate(parts[::-1]):
            nb = dict(b, top=ls[0][0], bottom=ls[-1][1], text="".join(t for _, _, t in ls), dist=[d], lines=ls,
                      cont_from_prev=b.get("cont_from_prev") if idx == 0 else False)
            out.append(nb)
        print("   拆段", [d for d, _ in parts[::-1]], "|", " / ".join("".join(t for _, _, t in ls)[:15] for _, ls in parts[::-1]), file=sys.stderr)
    return out


def assign(blocks, schools):
    """逐頁 DP：本頁段落 → 本頁學校(可有開頭段屬上一頁最後一校)"""
    out = []      # (school, dist, text, block)
    pages = sorted({b["page"] for b in blocks} | {s["page"] for s in schools})
    last_school = None
    for pg in pages:
        B = [b for b in blocks if b["page"] == pg]
        S = [s for s in schools if s["page"] == pg]
        if not S:
            for b in B:
                out.append((last_school, b))
            continue
        n, m = len(B), len(S)
        INF = float("inf")
        # dp[i][j]: 前 i 段分給 前 j 校(第 0 校前可有接續段 c 段)的最小成本
        best = (INF, None)
        for c in range(0, n - m + 1):
            if c and last_school is None:
                break
            dp = [[INF] * (m + 1) for _ in range(n + 1)]
            back = [[None] * (m + 1) for _ in range(n + 1)]
            dp[c][0] = 0
            for j in range(1, m + 1):
                for i in range(c + j, n - (m - j) + 1):
                    for k in range(c + j - 1, i):
                        if dp[k][j - 1] == INF:
                            continue
                        center = (B[k]["top"] + B[i - 1]["bottom"]) / 2
                        cost = abs(center - S[j - 1]["mid"])
                        if j == m:                 # 本頁最後一校可能延到下一頁：不能低於校名
                            cost = 0 if B[k]["top"] - 3 <= S[j - 1]["mid"] <= B[i - 1]["bottom"] + 3 else 1000
                        if dp[k][j - 1] + cost < dp[i][j]:
                            dp[i][j] = dp[k][j - 1] + cost
                            back[i][j] = k
            if dp[n][m] < best[0]:
                parts, i = [], n
                for j in range(m, 0, -1):
                    k = back[i][j]
                    parts.append((S[j - 1]["name"], B[k:i]))
                    i = k
                best = (dp[n][m], (B[:c], parts[::-1]))
        if best[1] is None:
            print("!! 第", pg + 1, "頁分不出來", file=sys.stderr)
            continue
        cont, parts = best[1]
        for b in cont:
            out.append((last_school, b))
        for nm, bs in parts:
            for b in bs:
                out.append((nm, b))
        last_school = parts[-1][0]
    return out


def triples(pdf_path):
    blocks, schools = parse(pdf_path)
    blocks = split_multi(blocks)
    res, last_dist = [], None
    for sch, b in assign(blocks, schools):
        if b["dist"]:
            if len(b["dist"]) > 1:
                print("!! 一段有兩個行政區", sch, b["dist"], b["text"][:40], file=sys.stderr)
            last_dist = b["dist"][0][1]
        elif not b.get("cont_from_prev"):
            print("!! 段落沒有行政區(沿用上一個)", sch, last_dist, b["text"][:40], file=sys.stderr)
        if b.get("cont_from_prev") and res and res[-1][0] == sch and res[-1][1] == last_dist:
            res[-1] = (sch, last_dist, res[-1][2] + b["text"])      # 跨頁接續(「瑞」+「和里」)接回同一段
        else:
            res.append((sch, last_dist, b["text"]))
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    for t in triples(sys.argv[1]):
        print(t[0], "|", t[1], "|", t[2])
