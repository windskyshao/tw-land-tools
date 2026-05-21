"""
EasyOCR 輔助函數
統一管理 EasyOCR 模型目錄設置
"""
import os
import sys


def get_easyocr_model_dir():
    """
    取得 EasyOCR 模型目錄路徑
    優先使用 python_embedded 下的目錄

    返回: 模型目錄的完整路徑
    """
    # 尋找 python_embedded 目錄
    current_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(current_dir)

    # 打包環境：在 _internal 或主程式目錄的 python_embedded 下
    python_embedded_dir = None
    for search_dir in [current_dir, parent_dir, os.path.dirname(parent_dir)]:
        candidate = os.path.join(search_dir, 'python_embedded')
        if os.path.exists(candidate):
            python_embedded_dir = candidate
            break

    # 如果找到 python_embedded，使用其下的 easyocr_models；否則使用當前目錄
    if python_embedded_dir:
        model_dir = os.path.join(python_embedded_dir, 'easyocr_models')
    else:
        model_dir = os.path.join(current_dir, 'easyocr_models')

    os.makedirs(model_dir, exist_ok=True)
    return model_dir


def create_easyocr_reader(languages=['en'], gpu=False, verbose=False):
    """
    創建 EasyOCR Reader 實例
    自動設置模型目錄到 python_embedded

    參數:
        languages: 語言列表，預設 ['en']
        gpu: 是否使用 GPU，預設 False
        verbose: 是否顯示詳細輸出，預設 False

    返回: easyocr.Reader 實例
    """
    import easyocr

    model_dir = get_easyocr_model_dir()
    print(f"EasyOCR 模型目錄: {model_dir}", flush=True)

    # 檢查模型檔案是否存在
    import os
    craft_model = os.path.join(model_dir, 'craft_mlt_25k.pth')
    english_model = os.path.join(model_dir, 'english_g2.pth')

    if os.path.exists(craft_model) and os.path.exists(english_model):
        print(f"✓ 模型檔案已存在，停用自動下載", flush=True)
        download_enabled = False
    else:
        print(f"⚠ 模型檔案不存在，啟用自動下載", flush=True)
        download_enabled = True

    reader = easyocr.Reader(
        languages,
        gpu=gpu,
        model_storage_directory=model_dir,
        download_enabled=download_enabled,
        verbose=verbose
    )

    return reader
