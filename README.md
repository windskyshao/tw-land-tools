# 地籍資料查詢系統（tw-land-tools）

不動產營業員作業便捷工具集 — 整合台灣地籍便民系統、土地使用分區、實價登錄、電子謄本、土地增值稅、學區查詢等政府網站的自動化操作。

> **僅供台灣地區使用**。所有功能依賴台灣官方政府網站（內政部、財政部、各縣市政府等）。

## 主要功能

| 模組 | 功能 |
|---|---|
| 地籍便民系統 | 解析地號 → 自動建立案件資料夾結構 |
| 電子謄本下載 | `hinet.py` / `qpt_hinet.py` 自動下載第一/二類謄本（自然人憑證）|
| 電子謄本結構化 | 解析謄本 PDF → 結構化 JSON（土地/建物分區排序、多筆合併）|
| 土地使用分區 | 國土測繪 (`nlscmaps.py`)、高雄都市計畫 (`urbangis.py`)、全國土地使用分區 (`luz.py`) |
| 土地增值稅試算 | 財政部 (`finance_tax.py`)、高雄市 (`kaohsiung_tax.py`) 兩版 |
| 實價登錄 | 內政部 (`interior_price.py`)、104 實價網 (`price104.py`) |
| 使用執照查詢 | 高雄 (`nlma.py`)、屏東/台南 (`nlma_nbupic.py`) |
| 其他 | V523、地質雲、重要設施、學區查詢（高雄/屏東/台南，國中/國小）|
| 物件資料編輯 | `data_editor.py` 編輯案件物件 → 匯出 Word 物件調查表 |
| 自動更新檢查 | 啟動時背景比對 GitHub Release，有新版時右上 🔄 按鈕轉紅 |

## 環境需求

- **作業系統**：Windows 10 / 11
- **Python**：3.9+（執行原始碼）/ 不需 Python（執行打包後的 .exe）
- **瀏覽器**：Google Chrome（建議最新版）
- **ChromeDriver**：對應 Chrome 版本，請至 <https://googlechromelabs.github.io/chrome-for-testing/> 下載

## 從原始碼執行

```bash
git clone https://github.com/windskyshao/tw-land-tools.git
cd tw-land-tools
pip install -r requirements.txt
python main.py
```

> `requirements.txt` 共約 294 個套件，是用 `pip freeze` 整理出來的完整環境快照。如果只想跑特定模組，實際只需要 selenium、tkinter、customtkinter、pandas、openpyxl 等核心套件。

## 打包成 .exe

### 一、必備檔案

下載/準備這些放在專案根目錄：

| 檔案/資料夾 | 用途 | 怎麼取得 |
|---|---|---|
| `chromedriver.exe` | Selenium 控制 Chrome | <https://googlechromelabs.github.io/chrome-for-testing/>（版本須對應 Chrome）|
| `python_embedded/` | 子腳本執行環境（~1.5 GB）| 見下方「python_embedded 準備」 |
| `通訊錄.xlsx` | 個人聯絡人清單（選用）| 自行建立，沒有也能執行（對應功能會無資料）|

`通訊錄.xlsx` 欄位範例：

| 姓名 | 電話 | 備註 |
|---|---|---|
| 王小明 | 0912-345-678 | 同事 |

### 二、python_embedded 準備

主程式 (`地籍資料查詢系統.exe`) 啟動子腳本（`hinet.py`、`finance_tax.py` 等）時，需要一個包含完整 selenium / pandas / easyocr 等套件的 Python 環境，這個資料夾就是它。

1. 從 <https://www.python.org/downloads/windows/> 下載「Windows embeddable package (64-bit)」zip（建議 **Python 3.9.x**）
2. 解壓到 `python_embedded/`，與 `main.py` 同層
3. 編輯 `python_embedded/python39._pth`，把 `#import site` 那行的 `#` 拿掉（啟用 site-packages）
4. 下載 [get-pip.py](https://bootstrap.pypa.io/get-pip.py) 到 `python_embedded/`
5. 安裝套件：
   ```cmd
   cd python_embedded
   python.exe get-pip.py
   python.exe -m pip install -r ..\requirements.txt
   ```

這個資料夾不附在 repo（太大），請每位開發者自行準備。

### 三、打包

直接執行：

```cmd
pyinstaller "地籍資料查詢系統.spec" -y
```

或執行 **`0.bat`**（會同步到 `dist-0/`，加速增量更新）：

```cmd
.\0.bat
```

打包完成後，`dist\地籍資料查詢系統\` 內含：
- `_internal/`（PyInstaller 自動產生）
- `地籍資料查詢系統.exe`
- `系統使用說明文檔.html`（已內建在 .spec 的 COLLECT 步驟）

### 四、執行時 dist 資料夾應有的內容

| 項目 | 說明 |
|---|---|
| `_internal/` | PyInstaller 產生 |
| `地籍資料查詢系統.exe` | 主程式 |
| `系統使用說明文檔.html` | 系統說明（打包自動建立）|
| `python_embedded/` | **必備**，子腳本執行環境 |
| `chromedriver.exe` | **必備**，需與 Chrome 同版本 |
| `通訊錄.xlsx` | 選用，個人聯絡人 |
| `地籍圖處理工具*.exe` | **獨立程式**，本 repo **不包含其原始碼**，但主程式提供「地籍圖處理工具」按鈕會啟動它 |
| `電子謄本結構化*.exe` | 由 `GUI_transcript_pdf 1141021-01.py` **另外打包**，主程式按「電子謄本結構化」會啟動。檔名支援 `電子謄本結構化v1.5b.exe` 之類，會自動找最新版本 |

> 缺少 `地籍圖處理工具*.exe` 或 `電子謄本結構化*.exe` 不影響主程式啟動，只有對應按鈕點下去會顯示找不到。

## 自動更新機制

### 使用者端

- 啟動 2 秒後在背景檢查 GitHub Releases
- 本地版號 ≥ 遠端最新 release：右上 🔄 按鈕維持藍色
- 遠端有新版：🔄 按鈕轉紅色 + 滑鼠移上會提示「有新版可下載」
- 點 🔄 跳出對話框：
  - **若該 release 附有 `*.zip` 檔**：詢問是否自動下載 + 替換 → 啟動 updater.bat → 退出主程式 → updater 等 3 秒後覆寫 `_internal/` 和 `地籍資料查詢系統.exe` → 重新啟動
  - **若無 zip**：開瀏覽器到 release 頁面讓使用者手動下載

### 自動更新會保留的使用者個人檔

`_internal/` 採 `robocopy /E` 增量覆寫（不是 `/MIR`），所以 zip 內**沒有**的檔案不會被刪除。因此：

- `_internal/通訊錄.xlsx`：保留（zip 製作時主動排除）
- `_internal/chromedriver.exe`：保留（zip 製作時主動排除）
- exe 同層的 `python_embedded/`、`地籍圖處理工具*.exe`、`電子謄本結構化*.exe`、`data.json`、案件資料夾：完全不動

### 維護者端：發新版的標準流程

1. 改 `main.py` 的 `VERSION`
2. 跑 `0.bat` 打包到 `dist-0/`
3. 跑本機 `2.bat` 同步原始碼到 A_github
4. 跑本機 `3.bat` 推上 GitHub + 自動建立 tag + 開瀏覽器到 release 頁面
5. 跑本機 `4.bat` 製作 release zip（自動排除 `通訊錄.xlsx` 和 `chromedriver*.exe`）
6. 在 release 頁面把 `release_packages/地籍資料查詢系統_v{X.Y.Z}.zip` 拖進 Attach binaries → Publish

完成後，所有舊版的使用者點 🔄 都會自動下載並更新到新版。

> **限制**：自動更新只能更新 `地籍資料查詢系統.exe` 和 `_internal/`。如果版本動到 `python_embedded/`（例如新增套件）或外部 `地籍圖處理工具`、`電子謄本結構化` 程式，使用者需手動更新那些檔案，自動更新無法處理。

## 注意事項

- 本工具**僅自動化操作公開政府網站**，不會繞過驗證碼、不會盜取他人資料
- **自然人憑證帳密**儲存於 Windows 認證管理員（透過 `keyring` 套件），請至「帳號設定」設定
- 查詢結果僅供參考，**法律依據以官方文件為準**
- 各政府網站介面常更新，本工具可能需要對應修改

## 授權

尚未設定 License。

## 問題回報

請至 [Issues](https://github.com/windskyshao/tw-land-tools/issues) 開 issue。
