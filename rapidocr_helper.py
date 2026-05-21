"""
RapidOCR 輔助函數
統一管理 RapidOCR 模型目錄設置
"""
import os
import sys


def get_rapidocr_model_dir():
    """
    取得 RapidOCR 模型目錄路徑
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

    # 如果找到 python_embedded，使用其下的 rapidocr_models；否則使用當前目錄
    if python_embedded_dir:
        model_dir = os.path.join(python_embedded_dir, 'rapidocr_models')
    else:
        model_dir = os.path.join(current_dir, 'rapidocr_models')

    os.makedirs(model_dir, exist_ok=True)
    return model_dir


def create_rapidocr_reader(
    use_angle_cls=True,
    use_text_score=True,
    text_score=0.5,
    # 檢測參數（影響文字區域偵測）
    det_db_thresh=0.3,           # 二值化閾值，越小越容易檢測到文字（預設 0.3）
    det_db_box_thresh=0.5,       # 文字框置信度閾值（預設 0.5）
    det_db_unclip_ratio=1.6,     # 文字框擴展比例，越大框越大（預設 1.6）
    # 識別參數
    rec_batch_num=6,             # 批次處理數量（預設 6）
    max_side_len=960,            # 最大邊長限制（預設 960）
    # 🔥 quiet=True 時不印「模型目錄」與「參數」debug 訊息
    quiet=False,
):
    """
    創建 RapidOCR Reader 實例
    自動設置模型目錄到 python_embedded

    🔧 針對驗證碼優化的參數說明：

    檢測參數（影響能否找到文字）：
        det_db_thresh: 二值化閾值
            - 預設 0.3，範圍 0.0~1.0
            - 驗證碼建議 0.2~0.3（更容易檢測到文字）
            - 值越小越敏感，但可能產生噪點

        det_db_box_thresh: 文字框置信度閾值
            - 預設 0.5，範圍 0.0~1.0
            - 驗證碼建議 0.3~0.5（降低可提高檢出率）

        det_db_unclip_ratio: 文字框擴展比例
            - 預設 1.6
            - 驗證碼建議 1.5~2.0（擴大文字框範圍）

        max_side_len: 圖片最大邊長
            - 預設 960
            - 驗證碼建議 1280~2048（提高解析度）

    識別參數：
        text_score: 文字置信度閾值
            - 預設 0.5，範圍 0.0~1.0
            - 驗證碼建議 0.3~0.5（降低可接受更多結果）

        use_angle_cls: 是否使用方向分類
            - 驗證碼通常不需要，可設為 False 加速

        use_text_score: 是否使用文字置信度
            - 建議 True，可過濾低品質結果

    返回: RapidOCR 實例
    """
    from rapidocr_onnxruntime import RapidOCR

    model_dir = get_rapidocr_model_dir()
    if not quiet:
        print(f"RapidOCR 模型目錄: {model_dir}", flush=True)

    # 🔧 針對驗證碼優化的配置
    reader = RapidOCR(
        # 文字方向分類
        use_angle_cls=use_angle_cls,
        use_text_score=use_text_score,
        text_score=text_score,

        # 檢測模型參數（影響文字區域偵測）
        det_db_thresh=det_db_thresh,
        det_db_box_thresh=det_db_box_thresh,
        det_db_unclip_ratio=det_db_unclip_ratio,

        # 識別模型參數
        rec_batch_num=rec_batch_num,

        # 圖片處理參數
        max_side_len=max_side_len,
    )

    if not quiet:
        print(f"✓ RapidOCR 參數: det_thresh={det_db_thresh}, box_thresh={det_db_box_thresh}, "
              f"unclip_ratio={det_db_unclip_ratio}, max_side={max_side_len}, text_score={text_score}", flush=True)

    return reader


def rapidocr_readtext(reader, image, detail=0):
    """
    使用 RapidOCR 辨識圖片文字
    提供與 EasyOCR 相似的介面

    參數:
        reader: RapidOCR 實例
        image: numpy array 格式的圖片
        detail: 0=僅返回文字列表, 1=返回完整結果（座標、文字、置信度）

    返回:
        detail=0: ['text1', 'text2', ...]
        detail=1: [[[x1,y1], [x2,y2], [x3,y3], [x4,y4]], 'text', confidence]
    """
    import numpy as np

    # 確保圖片是 numpy array
    if not isinstance(image, np.ndarray):
        import PIL.Image
        if isinstance(image, PIL.Image.Image):
            image = np.array(image)
        else:
            raise ValueError("圖片必須是 numpy array 或 PIL Image")

    # RapidOCR 的返回格式: result, elapse
    # result 是 list，每個元素是 [[[x1,y1],[x2,y2],[x3,y3],[x4,y4]], text, confidence]
    result, elapse = reader(image)

    if result is None:
        return [] if detail == 0 else []

    if detail == 0:
        # 僅返回文字列表
        return [item[1] for item in result]
    else:
        # 返回完整結果，格式與 EasyOCR 相容
        return result
