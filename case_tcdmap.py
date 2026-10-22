"""案件附件:國土功能分區(nluz) 無螢幕查詢+截圖存 PDF。
用法: python case_tcdmap.py 縣市 行政區 段名 地號 輸出資料夾(案件\\1.基本資料)"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ["LAND_HEADLESS"] = "1"
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_helper import create_chrome_driver

city, area, section, lot, out = sys.argv[1:6]
os.makedirs(out, exist_ok=True)
t0 = time.time()
log = lambda m: print(f"[{time.time() - t0:5.1f}s] {m}", flush=True)
driver = create_chrome_driver()
try:
    try:
        driver.execute_cdp_cmd("Browser.grantPermissions", {"permissions": ["geolocation"]})
    except Exception:
        pass
    driver.get(f"https://nluz.nlma.gov.tw/ex3smap/default.aspx?CN={city}&version=報部版")
    log("開啟國土功能分區網頁")
    try:
        WebDriverWait(driver, 8).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.btn.btn-primary[data-bs-dismiss='modal']"))).click()
    except Exception:
        pass
    WebDriverWait(driver, 25).until(lambda d: len(d.find_element(By.ID, "cada_townname").find_elements(By.TAG_NAME, "option")) > 1)
    for o in driver.find_element(By.ID, "cada_townname").find_elements(By.TAG_NAME, "option"):
        if o.text.strip() == area:
            o.click(); break
    WebDriverWait(driver, 15).until(lambda d: len(d.find_element(By.ID, "cada_secname").find_elements(By.TAG_NAME, "option")) > 1)
    time.sleep(1)
    picked = None
    for o in driver.find_element(By.ID, "cada_secname").find_elements(By.TAG_NAME, "option"):
        if o.text.strip() == section or o.text.strip().endswith(section):
            o.click(); picked = o.text; break
    if not picked:
        print("FAIL: 國土分區網站找不到這個段名", flush=True); sys.exit(2)
    li = driver.find_element(By.ID, "cada_num"); li.clear(); li.send_keys(lot)
    driver.find_element(By.ID, "cada_search").click()
    log(f"查詢 {area}{picked}{lot}")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".esri-popup__main-container, .esri-popup")))
    except Exception:
        pass
    time.sleep(5)
    base = os.path.join(out, f"08_國土分區-{area}{section}-{lot}")
    driver.save_screenshot(base + ".png")
    from PIL import Image
    Image.open(base + ".png").convert("RGB").save(base + ".pdf", "PDF", resolution=150)
    os.remove(base + ".png")
    log("存成 PDF")
finally:
    driver.quit()
