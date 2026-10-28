"""案件附件合併列印：把案件資料夾 1.基本資料 裡的附件依編號接成一份 PDF(可雙面列印)。
- 順序跟使用者手動用 Acrobat「合併檔案」一樣：01.1 財政部 Excel → 01 財政部 PDF → 01-1.1 高雄 Excel → 01-1 高雄 PDF → 02 … 09
- 同一份附件有好幾個日期的版本(…_20260519、…_20260520)只取最新的
- Excel 用這台電腦的 Excel 轉 PDF(PowerShell COM，不需 pywin32)；沒有 Excel 改試 LibreOffice；都沒有就略過並說明
- 每頁統一成 A4(直/橫依內容)，等比縮放置中：Letter、國土分區 759×434 這類非 A4 的頁印出來大小才一致
- 預設不補空白頁(直接雙面列印)；--pad-blank：單數頁的文件後面補空白頁，每份從新的一張紙開始
- 存成 1.基本資料/00_合併列印_{案件}_{日期}.pdf(同名先刪舊的)
- 0.謄本 的謄本也能一起合併(放最前面)；3.行情 的實價登錄(內政部、104)放最後；主程式的挑選視窗決定要哪些 → --list 清單.json {"files":[…], "unpicked":[…]}
用法：python case_merge.py <案件資料夾> [--pad-blank] [--title 標題] [--list 清單.json]
最後一行印 OK:<合併檔路徑> 或 FAIL:<原因>"""
import os
import re
import sys
import glob
import shutil
import argparse
import tempfile
import json
import subprocess
from datetime import datetime

NO_WINDOW = 0x08000000
OUT_PREFIX = "00_合併列印_"
EXTS = (".pdf", ".xlsx", ".xls", ".png", ".jpg", ".jpeg")
_KEY = re.compile(r"^(\d+)(?:-(\d+))?(?:\.(\d+))?_")
_DATE = re.compile(r"_(\d{8})(?=\.[^.]+$)")


def sort_key(name):
    """01.1 → (1,0,0)；01 → (1,0,1)；01-1.1 → (1,1,0)；01-1 → (1,1,1)：同一號 Excel 摘要在 PDF 明細前面"""
    m = _KEY.match(name)
    if not m:
        return None
    major, sub, dot = int(m.group(1)), int(m.group(2) or 0), m.group(3)
    return (major, sub, 0 if dot else 1, name)


def pick_files(info_dir):
    """回傳要合併的檔案(已排序)。沒編號的(例如 組合 1.pdf)和合併檔本身不算"""
    best = {}
    for p in glob.glob(os.path.join(info_dir, "*")):
        n = os.path.basename(p)
        if not os.path.isfile(p) or not n.lower().endswith(EXTS) or n.startswith(OUT_PREFIX) or n.startswith("~$"):
            continue
        if sort_key(n) is None:
            continue
        m = _DATE.search(n)
        base = _DATE.sub("", n) if m else n                 # 去掉日期 → 同一份附件
        d = m.group(1) if m else ""
        if base not in best or (d, os.path.getmtime(p)) > best[base][0]:
            best[base] = ((d, os.path.getmtime(p)), p)
    return sorted((v[1] for v in best.values()), key=lambda p: sort_key(os.path.basename(p)))


TRANS_EXTS = (".pdf", ".png", ".jpg", ".jpeg")
_VARIANT = re.compile(r"(_標註|_重新處理|_unlocked|_\d)+$")


def _trans_cat(name):
    for k, c in (("土地登記", 0), ("建物登記", 1), ("地籍圖", 2), ("畫線", 2), ("成果圖", 3)):
        if k in name:
            return c
    return 4


def transcript_files(case_dir):
    """0.謄本 裡的謄本(只看第一層)：土地登記→建物登記→地籍圖/畫線→建物測量成果圖→其他"""
    d = os.path.join(case_dir, "0.謄本")
    out = [p for p in glob.glob(os.path.join(d, "*")) if os.path.isfile(p)
           and p.lower().endswith(TRANS_EXTS) and not os.path.basename(p).startswith("~$")]
    return sorted(out, key=lambda p: (_trans_cat(os.path.basename(p)), os.path.basename(p)))


def default_picks(paths):
    """謄本預設勾哪些：同一張圖有好幾版(原檔/_標註/_重新處理/_unlocked)只勾一份——有標註版勾最新的標註版，否則勾原檔"""
    groups = {}
    for p in paths:
        stem = os.path.splitext(os.path.basename(p))[0]
        groups.setdefault(_VARIANT.sub("", stem), []).append(p)
    pick = set()
    for g in groups.values():
        marked = [p for p in g if "_標註" in os.path.basename(p)]
        if marked:
            pick.add(max(marked, key=os.path.getmtime))
        else:
            plain = [p for p in g if not _VARIANT.search(os.path.splitext(os.path.basename(p))[0])]
            pick.add(plain[0] if plain else max(g, key=os.path.getmtime))
    return pick


def rel(case_dir, p):
    return os.path.relpath(p, case_dir).replace("\\", "/")


def candidates(case_dir):
    """挑選視窗要列的全部檔案：[(分區, 路徑)]，0.謄本 在前、1.基本資料 在後(已排好合併順序)"""
    out = [("0.謄本", p) for p in transcript_files(case_dir)]
    info = os.path.join(case_dir, "1.基本資料")
    if os.path.isdir(info):
        out += [("1.基本資料", p) for p in pick_files(info)]
    out += [("3.行情", p) for p in market_files(case_dir)]
    return out


def market_files(case_dir):
    """3.行情：內政部實價登錄 → 104 實價網 → 其他；同名的 PDF 跟 PNG 只取 PDF"""
    d = os.path.join(case_dir, "3.行情")
    fs = [p for p in glob.glob(os.path.join(d, "*")) if os.path.isfile(p)
          and p.lower().endswith(TRANS_EXTS) and not os.path.basename(p).startswith("~$")]
    stems = {os.path.splitext(p)[0].lower() for p in fs if p.lower().endswith(".pdf")}
    fs = [p for p in fs if p.lower().endswith(".pdf") or os.path.splitext(p)[0].lower() not in stems]
    cat = lambda n: 0 if n.startswith("內政部") else 1 if n.startswith("104") else 2
    return sorted(fs, key=lambda p: (cat(os.path.basename(p)), os.path.basename(p)))


def _ps_quote(s):
    return "'" + s.replace("'", "''") + "'"


def excel_to_pdf(src, dst):
    """Excel COM 轉 PDF；只關自己開的活頁簿，Excel 裡沒有別的活頁簿才結束 Excel(不動使用者開著的檔)"""
    ps = (
        "$ErrorActionPreference='Stop';"
        "$x=New-Object -ComObject Excel.Application;$x.DisplayAlerts=$false;"
        f"$wb=$x.Workbooks.Open({_ps_quote(src)},0,$true);"
        f"try{{$wb.ExportAsFixedFormat(0,{_ps_quote(dst)})}}finally{{$wb.Close($false)}};"
        "if($x.Workbooks.Count -eq 0){$x.Quit()};"
        "[void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)"
    )
    try:
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], capture_output=True,
                       timeout=120, creationflags=NO_WINDOW)
    except Exception:
        pass
    if os.path.exists(dst) and os.path.getsize(dst) > 0:
        return "Excel"
    for soffice in (shutil.which("soffice"), r"C:\Program Files\LibreOffice\program\soffice.exe",
                    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if soffice and os.path.exists(soffice):
            outdir = os.path.dirname(dst)
            try:
                subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, src],
                               capture_output=True, timeout=180, creationflags=NO_WINDOW)
            except Exception:
                continue
            made = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
            if os.path.exists(made):
                if made != dst:
                    shutil.move(made, dst)
                return "LibreOffice"
    return None


A4 = (595.28, 841.89)


def _append_a4(doc, src):
    """src 每一頁放進 doc 的 A4 頁；本來就是 A4 的直接插入(保留連結/註解)"""
    baked = False
    for pg in src:
        r = pg.rect                                   # 已含旋轉
        land = r.width > r.height
        w, h = (A4[1], A4[0]) if land else A4
        if abs(r.width - w) < 3 and abs(r.height - h) < 3:
            doc.insert_pdf(src, from_page=pg.number, to_page=pg.number)
            continue
        if not baked:                                 # 表單欄位/註解先燒進內容，否則縮放貼上會不見
            try:
                src.bake()
            except Exception:
                pass
            baked = True
            pg = src[pg.number]
        new = doc.new_page(width=w, height=h)
        new.show_pdf_page(new.rect, src, pg.number, keep_proportion=True)


def merge(case_dir, duplex=False, title=None, log=print, files=None, unpicked=None):
    info = os.path.join(case_dir, "1.基本資料")
    os.makedirs(info, exist_ok=True)
    if files is None:
        files = pick_files(info)
    if not files:
        raise RuntimeError("沒有勾選任何要合併的檔案")
    title = title or os.path.basename(os.path.normpath(case_dir))
    out = os.path.join(info, f"{OUT_PREFIX}{title}_{datetime.now():%Y%m%d}.pdf")
    import fitz                                       # 用到才載入(主程式只借 pick_files 判斷有沒有變動)
    tmp = tempfile.mkdtemp(prefix="case_merge_")
    doc = fitz.open()
    skipped, blanks, toc = [], 0, []
    try:
        for i, p in enumerate(files, 1):
            n = os.path.basename(p)
            ext = os.path.splitext(n)[1].lower()
            src = None
            try:
                if ext in (".xlsx", ".xls"):
                    conv = os.path.join(tmp, f"{i:02d}.pdf")
                    how = excel_to_pdf(p, conv)
                    if not how:
                        skipped.append(f"{n}（這台電腦沒有 Excel，無法轉成 PDF）")
                        log(f"  ✗ {n}：沒有 Excel，略過")
                        continue
                    src = fitz.open(conv)
                elif ext == ".pdf":
                    src = fitz.open(p)
                else:                                       # 圖片 → 一頁 A4，等比放進去
                    img = fitz.open(p)
                    r = img[0].rect
                    w, h = (842, 595) if r.width > r.height else (595, 842)
                    src = fitz.open()
                    pg = src.new_page(width=w, height=h)
                    pg.insert_image(pg.rect + (20, 20, -20, -20), filename=p, keep_proportion=True)
                if src.needs_pass:
                    skipped.append(f"{n}（有密碼保護）")
                    continue
                start = len(doc)
                _append_a4(doc, src)
                toc.append([1, os.path.splitext(n)[0], start + 1])          # 書籤：Acrobat 左邊可跳到各份
                log(f"  ✓ {n}（{len(src)} 頁）")
                if duplex and len(src) % 2 == 1 and i < len(files):
                    doc.new_page(width=A4[0], height=A4[1])
                    blanks += 1
            except Exception as e:
                skipped.append(f"{n}（{str(e)[:40]}）")
                log(f"  ✗ {n}：{e}")
            finally:
                if src is not None:
                    src.close()
        if len(doc) == 0:
            raise RuntimeError("沒有任何附件合併成功")
        try:
            doc.set_toc(toc)
        except Exception:
            pass
        locked = False
        for old in glob.glob(os.path.join(info, OUT_PREFIX + "*.pdf")):
            try:
                os.remove(old)
            except OSError:                               # 上一份還開在 Acrobat 裡(被鎖住)
                if os.path.normcase(old) == os.path.normcase(out):
                    locked = True
        if locked:                                        # 同名被鎖 → 加時間另存，舊的下次合併再清
            out = os.path.join(info, f"{OUT_PREFIX}{title}_{datetime.now():%Y%m%d_%H%M%S}.pdf")
            log("  ⚠ 上一份合併檔還開著，這次存成新檔名")
        doc.save(out, garbage=3, deflate=True)
        try:                                              # 紀錄：之後判斷「附件有沒有變」「合併檔有沒有被手動改過」
            os.makedirs(os.path.join(case_dir, "4.其他相關"), exist_ok=True)
            with open(os.path.join(case_dir, "4.其他相關", "合併列印紀錄.json"), "w", encoding="utf-8") as f:
                json.dump({"file": os.path.basename(out), "mtime": os.path.getmtime(out),
                           "sources": {rel(case_dir, x): os.path.getmtime(x) for x in files},
                           "unpicked": list(unpicked or [])},
                          f, ensure_ascii=False, indent=1)
        except Exception:
            pass
        return out, len(doc), blanks, skipped, files
    finally:
        doc.close()
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("case_dir")
    ap.add_argument("--pad-blank", action="store_true")
    ap.add_argument("--no-duplex", action="store_true")              # 舊參數(現在預設就不補)
    ap.add_argument("--title", default=None)
    ap.add_argument("--list", default=None)
    a = ap.parse_args()
    files = unpicked = None
    if a.list:
        with open(a.list, encoding="utf-8") as f:
            lst = json.load(f)
        files, unpicked = [x for x in lst["files"] if os.path.exists(x)], lst.get("unpicked", [])
    try:
        out, pages, blanks, skipped, files = merge(a.case_dir, duplex=a.pad_blank and not a.no_duplex, title=a.title,
                                                   files=files, unpicked=unpicked,
                                                   log=lambda t: print(t, flush=True))
        print(f"共 {len(files) - len(skipped)} 份、{pages} 頁" + (f"（含 {blanks} 張補的空白頁）" if blanks else ""),
              flush=True)
        for s in skipped:
            print("WARN:" + s, flush=True)
        print("OK:" + out, flush=True)
    except Exception as e:
        print("FAIL:" + str(e), flush=True)


if __name__ == "__main__":
    main()
