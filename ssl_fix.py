# -*- coding: utf-8 -*-
"""HTTPS 連線保險(2026-11-27 同事筆電：Google 登入「連不上阿生地圖」
[DSO: LOAD_FAILED] could not load the shared library (_ssl.c:4192))。

Python 在 Windows 建 HTTPS 連線時會把「Windows 憑證存放區」裡所有憑證讀進來；
有些電腦裝了特殊憑證(銀行、自然人憑證、公司軟體…)，OpenSSL 讀它時要另外載入元件、載不到就整個失敗。
→ 建預設連線設定失敗時，改用程式內建的憑證清單(certifi，跟瀏覽器用的公開憑證一樣)，一樣有驗證、不是不檢查。
主程式一開始 install() 一次，之後 urllib(阿生地圖、Google 登入、學區…)都會走這個。"""
import ssl

_orig = ssl.create_default_context
_installed = False


def _certifi_file():
    try:
        import certifi
        return certifi.where()
    except Exception:
        return None


def safe_create_default_context(purpose=ssl.Purpose.SERVER_AUTH, *, cafile=None, capath=None, cadata=None):
    try:
        return _orig(purpose, cafile=cafile, capath=capath, cadata=cadata)
    except ssl.SSLError:
        if cafile or capath or cadata:
            raise
        ca = _certifi_file()
        if not ca:
            raise
        return _orig(purpose, cafile=ca)          # 給了 cafile 就不會去讀 Windows 憑證存放區


def install():
    global _installed
    if _installed:
        return
    ssl.create_default_context = safe_create_default_context
    ssl._create_default_https_context = safe_create_default_context
    _installed = True
