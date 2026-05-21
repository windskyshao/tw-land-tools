# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

# 🔥 datas 中的檔案會被打包到 _internal 目錄
# 注意：原本還包含 ('通訊錄.xlsx', '.') 與 ('chromedriver.exe', '.')，
# 公開版本已移除（前者為使用者個人聯絡人，後者需自行下載對應 Chrome 版本）
datas = [('base_dir_helper.py', '.'), ('credential_helper.py', '.'), ('config_manager.py', '.'), ('data_editor.py', '.'), ('export_archive_new.py', '.'), ('webdriver_helper.py', '.'), ('easyocr_helper.py', '.'), ('rapidocr_helper.py', '.'), ('land.py', '.'), ('luz.py', '.'), ('nlma.py', '.'), ('nlma_nbupic.py', '.'), ('gbmap.py', '.'), ('V523.py', '.'), ('nlscmaps.py', '.'), ('urbangis.py', '.'), ('tcdmap.py', '.'), ('geologycloud.py', '.'), ('interior_price.py', '.'), ('price104.py', '.'), ('hinet.py', '.'), ('qpt_hinet.py', '.'), ('GUI_transcript_pdf 1141021-01.py', '.'), ('finance_tax.py', '.'), ('kaohsiung_tax.py', '.'), ('captcha_utils.py', '.'), ('build_school_district_data.py', '.'), ('build_tainan_school_district_data.py', '.'), ('build_pingtung_school_district_data.py', '.'), ('school_district_query.py', '.'), ('school_district_lookup.py', '.'), ('ris_query_selenium.py', '.'), ('ris_query_wrapper.py', '.'), ('高雄市_國中學區資料.json', '.'), ('高雄市_國小學區資料.json', '.'), ('屏東縣國中學區資料.json', '.'), ('屏東縣國小學區資料.json', '.'), ('台南市國中學區資料.json', '.'), ('台南市國小學區資料.json', '.'), ('settings.json', '.'), ('google_popup_coords.json', '.')]
# 🔥 系統使用說明文檔.html 移到 COLLECT 階段，放在與 exe 同層目錄
binaries = []
# 🔥 只保留主程式需要的套件（GUI相關）
# 子程式的套件（selenium, torch, easyocr等）由 python_embedded 提供
# 🔥 添加 pandas 和 Excel 相關套件（通訊錄載入功能需要）
hiddenimports = ['tkintermapview', 'geocoder', 'customtkinter', 'timeit', 'importlib', 'importlib.util', 'pandas', 'openpyxl', 'xlrd']


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 🔥 排除子程式專用的大型套件（由 python_embedded 提供）
    # 注意：psutil 和 docx 主程式需要，不能排除！
    # 🔥 pandas 和 openpyxl 主程式需要（通訊錄載入功能），不能排除！
    excludes=[
        'torch', 'torchvision', 'torchaudio', 'torch.utils',
        'selenium', 'webdriver_manager',
        'easyocr', 'cv2', 'opencv',
        'scipy', 'scipy.stats', 'scipy.special',
        'sklearn', 'scikit-image', 'skimage',
        # 'PIL', 'Pillow',  # 🔥 不排除 Pillow，因為 tkintermapview 需要
        # 'numpy', 'numpy.core', 'numpy.lib',  # 🔥 不排除 numpy，因為 pandas 需要
        'fitz', 'PyMuPDF', 'pikepdf',
        'fpdf', 'fpdf2',
        'keyboard', 'pyautogui', 'pyscreeze',
        # 'openpyxl',  # 🔥 不排除 openpyxl，通訊錄載入需要
        'matplotlib',  # 🔥 matplotlib 可以排除
        # 'pandas',  # 🔥 不排除 pandas，通訊錄載入需要
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='地籍資料查詢系統',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    # 🔥 將系統使用說明文檔.html 放在與 exe 同層目錄（不是 _internal 目錄）
    [('系統使用說明文檔.html', '系統使用說明文檔.html', 'DATA')],
    strip=False,
    upx=True,
    upx_exclude=[],
    name='地籍資料查詢系統',
)
