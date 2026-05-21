# captcha_utils.py

import cv2
import numpy as np

def get_captcha_coordinates(image_path):
    """
    从给定的图像中自动获取验证码区域的坐标。
    
    :param image_path: 图像文件的路径。
    :return: 验证码区域的坐标 (left, top, right, bottom)。
    """
    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edged = cv2.Canny(blurred, 50, 150)
    
    contours, _ = cv2.findContours(edged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = sorted(contours, key=cv2.contourArea, reverse=True)

    if contours:
        x, y, w, h = cv2.boundingRect(contours[0])
        left, top = x, y
        right, bottom = x + w, y + h
        return left, top, right, bottom
    else:
        return None
