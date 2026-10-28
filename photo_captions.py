"""照片邊角燒進去的文字標籤(例：591 照片右上角「附近街景圖」)：用 OCR 找位置，給行銷簡報裁照片時避開(不切到字)。
用法：python photo_captions.py <輸入清單.json> <輸出.json>
  輸入：["照片路徑", …]；輸出：{"照片路徑": [x0, y0, x1, y1](0～1 比例，已轉正) 或 null}
只算貼近照片邊緣的大字(招牌、路牌在照片中間不算)；被主程式用 python_embedded 叫(主程式本身沒有 OCR)。"""
import os
import sys
import json

EDGE = 0.05          # 離邊緣 5% 內才算「貼邊」
MIN_H = 0.035        # 字高至少照片高的 3.5%(太小的字不是標籤)


def caption_box(reader, path):
    import numpy as np
    from PIL import Image, ImageOps
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    W, H = im.size
    k = min(1.0, 1200 / max(W, H))                       # 縮到長邊 1200 再辨識(手機原圖 4032 太慢)
    if k < 1:
        im = im.resize((int(W * k), int(H * k)))
    w, h = im.size
    res, _ = reader(np.asarray(im)[:, :, ::-1])
    boxes = []
    for box, txt, sc in res or []:
        if sc < 0.6 or len(str(txt).strip()) < 2:
            continue
        xs, ys = [p[0] for p in box], [p[1] for p in box]
        x0, y0, x1, y1 = min(xs) / w, min(ys) / h, max(xs) / w, max(ys) / h
        if y1 - y0 < MIN_H:
            continue
        near_tb = y0 < EDGE or y1 > 1 - EDGE
        near_lr = x0 < EDGE or x1 > 1 - EDGE
        if near_tb and (near_lr or x1 - x0 > 0.3):       # 角落的標籤，或上/下緣橫跨的長標題
            boxes.append((x0, y0, x1, y1))
    if not boxes:
        return None
    pad = 0.012                                          # 標籤底色比字大一點
    return [max(0.0, min(b[0] for b in boxes) - pad), max(0.0, min(b[1] for b in boxes) - pad),
            min(1.0, max(b[2] for b in boxes) + pad), min(1.0, max(b[3] for b in boxes) + pad)]


def main():
    src, dst = sys.argv[1], sys.argv[2]
    with open(src, encoding="utf-8") as f:
        paths = json.load(f)
    from rapidocr_onnxruntime import RapidOCR          # 預設參數就好(rapidocr_helper 那套是調給驗證碼的)
    reader = RapidOCR()
    out = {}
    for p in paths:
        try:
            out[p] = caption_box(reader, p)
        except Exception as e:
            print(f"WARN:{os.path.basename(p)} {e}", flush=True)
            out[p] = None
        with open(dst, "w", encoding="utf-8") as f:      # 每張寫一次：中途逾時也留得住已算好的
            json.dump(out, f, ensure_ascii=False)
    print("OK", flush=True)


if __name__ == "__main__":
    main()
