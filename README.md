# 地籍資料查詢系統（tw-land-tools）

不動產營業員作業便捷工具集 — 整合地籍便民系統、土地使用分區查詢、實價登錄、電子謄本、土地增值稅試算等政府網站的自動化操作。

> **僅供台灣地區使用**。所有功能依賴台灣官方政府網站（內政部、財政部、各縣市政府等）。

## 主要功能

| 模組 | 功能 |
|---|---|
| 地籍便民系統 | 自動建立案件資料夾結構 |
| 電子謄本結構化 | 解析謄本 PDF 轉成結構化 JSON |
| 全國地政電子謄本 | hinet.py / qpt_hinet.py 自動下載第一/二類謄本 |
| 土地使用分區 | 國土測繪、高雄都市計畫、全國土地使用分區 |
| 土地增值稅試算 | 財政部、高雄市版兩種試算工具 |
| 實價登錄 | 內政部、104 實價網 |
| 使用執照查詢 | 高雄市、屏東縣、台南市 |
| 其他 | V523、地質雲、重要設施、學區查詢 |

## 環境需求

- **作業系統**：Windows 10 / 11
- **Python**：3.9+ （或使用打包後的 .exe，不需自行安裝 Python）
- **瀏覽器**：Google Chrome（最新版）
- **ChromeDriver**：對應 Chrome 版本，請至 <https://chromedriver.chromium.org/> 下載

## 安裝（從原始碼跑）

```bash
git clone https://github.com/windskyshao/tw-land-tools.git
cd tw-land-tools
pip install -r requirements.txt   # 請依需要套件自行整理
python main.py
```

## 打包成 .exe（PyInstaller）

需要先準備兩個檔案：

1. `chromedriver.exe` — 放在程式根目錄
2. `通訊錄.xlsx` — 你的聯絡人清單，格式：

   | 欄位 | 範例 |
   |---|---|
   | 姓名 | 王小明 |
   | 電話 | 0912-345-678 |
   | 備註 | 同事 |

   （若不需此功能可建立空白 xlsx，但執行時對應功能會無資料）

接著執行：

```bash
pyinstaller "地籍資料查詢系統.spec" -y
```

或直接執行 `0.bat`（會同步到 `dist-0/`）。

## 注意事項

- 本工具**僅自動化操作公開政府網站**，不會繞過驗證碼、不會盜取他人資料
- **自然人憑證帳密**儲存於 Windows 認證管理員，請至「帳號設定」設定
- 查詢結果僅供參考，**法律依據以官方文件為準**
- 各政府網站介面常更新，本工具可能需要對應修改

## 授權

尚未設定 License。

## 問題回報

請至 [Issues](https://github.com/windskyshao/tw-land-tools/issues) 開 issue。
