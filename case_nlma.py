"""案件附件:高雄使用執照(建築執照存根) 無螢幕查詢+驗證碼+存 PDF。
沿用 nlma.py 的 select_city / perform_search / click_and_save_pdf,只把瀏覽器換成無螢幕。
用法: python case_nlma.py 縣市 行政區 段名 地號 輸出資料夾(案件\\1.基本資料)"""
import os, sys, time, re
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ["LAND_HEADLESS"] = "1"
city, area, section, lot, out = sys.argv[1:6]
# ★2026-10-20 使用者指出:同地號有多張使照時只存第一張是錯的。--lic 字號1|字號2 = 網頁勾選的使照;沒帶→全部使照都存(各自一個 PDF)。
want = [x for x in sys.argv[sys.argv.index("--lic") + 1].split("|") if x] if "--lic" in sys.argv else []
_norm = lambda t: re.sub(r"\s+", "", (t or "")).translate(str.maketrans("０１２３４５６７８９（）－", "0123456789()-"))
os.makedirs(out, exist_ok=True)
import nlma                                      # 匯入:只會清畫面、建暫存資料夾
sys.stderr = sys.__stderr__
from selenium import webdriver
from webdriver_helper import create_chrome_driver

t0 = time.time()
log = lambda m: print(f"[{time.time() - t0:5.1f}s] {m}", flush=True)
opt = webdriver.ChromeOptions()
opt.add_experimental_option("prefs", {"savefile.default_directory": nlma.TMP_PDF_PATH,
                                      "download.prompt_for_download": False})
nlma.driver = create_chrome_driver(options=opt)
found = False
try:
    nlma.driver.get("https://cloudbm.nlma.gov.tw/CPTL/cpt0407m.do?")
    nlma.select_city(city)
    status, results = nlma.perform_search("4", city, area, section, lot)
    log(f"查詢結果:{status},{len(results or [])} 筆")
    if status == "success" and results:
        uses = [r for r in results if "使" in (r.get("license_number") or "")]
        if want:
            ws = {_norm(w) for w in want}
            targets = [r for r in results if _norm(r.get("license_number")) in ws]
            if not targets:
                log("勾選的使照在查詢結果找不到→改存全部使照")
                targets = uses or results[:1]
        else:
            targets = uses or results[:1]
        log(f"要存 {len(targets)} 張:" + "、".join((t.get("license_number") or "") for t in targets))
        for i, t in enumerate(targets):
            if i > 0:   # 存過一張後頁面已離開清單→回查詢頁重查一次,再點下一張
                nlma.driver.get("https://cloudbm.nlma.gov.tw/CPTL/cpt0407m.do?")
                nlma.select_city(city)
                st2, res2 = nlma.perform_search("4", city, area, section, lot)
                t2 = next((r for r in (res2 or []) if _norm(r.get("license_number")) == _norm(t.get("license_number"))), None)
                if not t2:
                    log(f"重查後找不到 {t.get('license_number')},略過"); continue
                t = t2
            log(f"存:{t.get('license_number')}")
            try:
                nlma.click_and_save_pdf(t, out)
                found = True
            except Exception as e:
                log(f"存 {t.get('license_number')} 失敗:{e}")
finally:
    try:
        nlma.driver.quit()
    except Exception:
        pass
if not found:
    print("NODATA: 這筆地號查無建照/使照資料", flush=True)
