# enhanced_transcript_gui.py - 美化版不動產登記謄本處理器
# =============================================================================
# 完全重新設計的現代化GUI介面
# =============================================================================

import time
import shutil
import os
import sys
import io

# 🔧 修復 Windows 打包環境的 UTF-8 編碼問題
# 確保 stdout/stderr 使用 UTF-8 以支援 emoji 和中文輸出
# 🔥 加入 line_buffering=True 以確保訊息即時顯示（不會等到程式結束才一次顯示）
try:
    if sys.stdout is not None and hasattr(sys.stdout, 'buffer'):
        if sys.stdout.encoding != 'utf-8':
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)
    if sys.stderr is not None and hasattr(sys.stderr, 'buffer'):
        if sys.stderr.encoding != 'utf-8':
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', line_buffering=True)
except (AttributeError, ValueError):
    # 在某些環境下（如 GUI 模式）可能沒有 buffer，忽略錯誤
    pass

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import pandas as pd
import json
import fitz  # PyMuPDF
import re
import threading
import queue
from datetime import datetime
from typing import List, Dict, Any
from collections import OrderedDict
from multiprocessing import freeze_support
# 🔧 加入這個條件導入
try:
    import send2trash
    HAS_SEND2TRASH = True
except ImportError:
    HAS_SEND2TRASH = False

# ==================== 螢幕適應性工具函數 ====================
class ScreenUtils:
    """螢幕適應性工具類"""
    
    @staticmethod
    def get_screen_geometry():
        """獲取可用螢幕區域（扣除工作列）"""
        root = tk.Tk()
        root.withdraw()  # 隱藏測試視窗
        
        # 獲取完整螢幕尺寸
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        
        # 嘗試獲取實際可用區域
        try:
            # Windows 特定方法
            if os.name == 'nt':
                import ctypes
                from ctypes import wintypes
                
                # 獲取工作區域
                class RECT(ctypes.Structure):
                    _fields_ = [("left", ctypes.c_long),
                               ("top", ctypes.c_long),
                               ("right", ctypes.c_long),
                               ("bottom", ctypes.c_long)]
                
                rect = RECT()
                ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0)
                
                work_width = rect.right - rect.left
                work_height = rect.bottom - rect.top
                work_x = rect.left
                work_y = rect.top
            else:
                # 其他系統的備用方法
                work_width = screen_width
                work_height = screen_height - 60  # 假設工作列高度
                work_x = 0
                work_y = 0
                
        except:
            # 備用方案：保守估計
            work_width = screen_width
            work_height = screen_height - 80  # 保守估計工作列高度
            work_x = 0
            work_y = 0
        
        root.destroy()
        
        return {
            'screen_width': screen_width,
            'screen_height': screen_height,
            'work_width': work_width,
            'work_height': work_height,
            'work_x': work_x,
            'work_y': work_y
        }
    
    @staticmethod
    def calculate_optimal_window_size(target_width=1400, target_height=800, margin=50):
        """計算最佳視窗大小和位置"""
        screen_info = ScreenUtils.get_screen_geometry()
        
        # 計算可用空間（留邊距）
        available_width = screen_info['work_width'] - (margin * 2)
        available_height = screen_info['work_height'] - (margin * 2)
        
        # 調整視窗大小以適應可用空間
        window_width = min(target_width, available_width)
        window_height = min(target_height, available_height)
        
        # 計算居中位置
        pos_x = screen_info['work_x'] + (screen_info['work_width'] - window_width) // 2
        pos_y = screen_info['work_y'] + (screen_info['work_height'] - window_height) // 2
        
        # 確保視窗不會超出邊界
        pos_x = max(screen_info['work_x'] + margin, pos_x)
        pos_y = max(screen_info['work_y'] + margin, pos_y)
        
        return {
            'width': int(window_width),
            'height': int(window_height),
            'x': int(pos_x),
            'y': int(pos_y)
        }
    
    @staticmethod
    def calculate_dialog_position(parent, dialog_width, dialog_height):
        """計算對話框最佳位置"""
        # 獲取父視窗位置和大小
        parent.update_idletasks()
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        # 計算居中位置
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        # 獲取螢幕資訊以確保對話框不會超出邊界
        screen_info = ScreenUtils.get_screen_geometry()
        
        # 調整位置確保對話框完全可見
        x = max(screen_info['work_x'], min(x, screen_info['work_x'] + screen_info['work_width'] - dialog_width))
        y = max(screen_info['work_y'], min(y, screen_info['work_y'] + screen_info['work_height'] - dialog_height))
        
        return int(x), int(y)

# ==================== 現代化樣式設定 ====================
class ModernStyle:
    """現代化UI樣式"""
    
    # 🔧 保留原有的顏色主題
    PRIMARY = "#2E86AB"      # 主要藍色
    SECONDARY = "#A23B72"    # 次要紫色
    SUCCESS = "#28A745"      # 成功綠色
    WARNING = "#FFC107"      # 警告黃色
    DANGER = "#DC3545"       # 危險紅色
    INFO = "#17A2B8"         # 資訊藍色
    
    # 🔧 新增：主題模式切換
    current_mode = "light"  # 預設為淺色模式
    
    # 🔧 新增：淺色模式定義（使用你原有的顏色）
    LIGHT_THEME = {
        "BG_PRIMARY": "#F8F9FA",   # 主背景
        "BG_SECONDARY": "#E9ECEF", # 次背景
        "BG_DARK": "#343A40",      # 深色背景
        "BG_LIGHT": "#FFFFFF",     # 淺色背景
        "TEXT_PRIMARY": "#212529",
        "TEXT_SECONDARY": "#6C757D",
        "TEXT_LIGHT": "#FFFFFF"
    }
    
    # 🔧 新增：暗黑模式定義
    DARK_THEME = {
        "BG_PRIMARY": "#2b2b2b",   # 主背景
        "BG_SECONDARY": "#3c3c3c", # 次背景
        "BG_DARK": "#1a1a1a",      # 深色背景
        "BG_LIGHT": "#404040",     # 淺色背景
        "TEXT_PRIMARY": "#ffffff",
        "TEXT_SECONDARY": "#b0b0b0",
        "TEXT_LIGHT": "#ffffff"
    }
    
    # 🔧 保留原有的靜態顏色（向後兼容）
    BG_PRIMARY = "#F8F9FA"
    BG_SECONDARY = "#E9ECEF"
    BG_DARK = "#343A40"
    BG_LIGHT = "#FFFFFF"
    TEXT_PRIMARY = "#212529"
    TEXT_SECONDARY = "#6C757D"
    TEXT_LIGHT = "#FFFFFF"
    
    # 🔧 保留原有字體設定
    FONT_MAIN = ("微軟正黑體", 12)
    FONT_HEADER = ("微軟正黑體", 14, "bold")
    FONT_TITLE = ("微軟正黑體", 16, "bold")
    FONT_CODE = ("微軟正黑體", 12)
    
    # 在 ModernStyle 類別中加入
    @classmethod
    def get_current_theme(cls):
        """取得當前主題的顏色"""
        if cls.current_mode == "dark":
            return {
                "BG_PRIMARY": "#1e1e1e",
                "BG_LIGHT": "#2d2d2d", 
                "BG_SECONDARY": "#3a3a3a",
                "TEXT_PRIMARY": "#ffffff",
                "TEXT_SECONDARY": "#cccccc"
            }
        else:
            return {
                "BG_PRIMARY": "#F8F9FA",
                "BG_LIGHT": "#FFFFFF",
                "BG_SECONDARY": "#E9ECEF", 
                "TEXT_PRIMARY": "#212529",
                "TEXT_SECONDARY": "#6C757D"
            }
    
    @classmethod
    def get_color(cls, color_name):
        """取得當前主題的特定顏色"""
        theme = cls.get_current_theme()
        return theme.get(color_name, getattr(cls, color_name, "#000000"))
    
    @classmethod
    def switch_theme(cls):
        """切換主題"""
        cls.current_mode = "dark" if cls.current_mode == "light" else "light"
        # 更新靜態屬性以保持向後兼容
        theme = cls.get_current_theme()
        cls.BG_PRIMARY = theme["BG_PRIMARY"]
        cls.BG_SECONDARY = theme["BG_SECONDARY"]
        cls.BG_DARK = theme["BG_DARK"]
        cls.BG_LIGHT = theme["BG_LIGHT"]
        cls.TEXT_PRIMARY = theme["TEXT_PRIMARY"]
        cls.TEXT_SECONDARY = theme["TEXT_SECONDARY"]
        cls.TEXT_LIGHT = theme["TEXT_LIGHT"]
    
    @classmethod
    def set_theme(cls, mode):
        """設定主題模式"""
        if mode in ["light", "dark"]:
            cls.current_mode = mode
            cls.switch_theme()
    
    # 🔧 保留原有的 ttk 配置方法
    @staticmethod
    def configure_ttk_style():
        """配置ttk樣式"""
        style = ttk.Style()
        
        # 配置Notebook樣式
        style.configure('Modern.TNotebook', background=ModernStyle.BG_PRIMARY)
        style.configure('Modern.TNotebook.Tab', 
                       padding=[20, 10],
                       font=ModernStyle.FONT_MAIN)
        
        # 配置Frame樣式
        style.configure('Modern.TFrame', background=ModernStyle.BG_LIGHT)
        style.configure('Card.TFrame', 
                       background=ModernStyle.BG_LIGHT,
                       relief='solid',
                       borderwidth=1)

# ==================== 基本處理函數 ====================
# 🔥 基準目錄設定
from base_dir_helper import BASE_DIR, get_data_json_path, get_work_folder

# 🔥 雙重保險：確保 BASE_DIR 不是 _internal 目錄
if '_internal' in BASE_DIR or BASE_DIR.endswith('_internal'):
    BASE_DIR = os.path.dirname(BASE_DIR)
    print(f"[警告] 修正 BASE_DIR（移除 _internal）: {BASE_DIR}", flush=True)

PDF_DIR = BASE_DIR  # 使用 BASE_DIR 而非 os.getcwd()，避免受工作目錄影響
OUT_DIR = get_work_folder("output")
os.makedirs(OUT_DIR, exist_ok=True)

def extract_pdf_text_only(pdf_path: str) -> str:
    """提取PDF文字"""
    try:
        doc = fitz.open(pdf_path)
        all_text = []
        
        for page in doc:
            text_dict = page.get_text("dict")
            for block in text_dict["blocks"]:
                if block.get("type") == 0:
                    for line in block.get("lines", []):
                        line_text = ""
                        for span in line.get("spans", []):
                            if span.get("text"):
                                line_text += span["text"]
                        if line_text.strip():
                            all_text.append(line_text.strip())
        
        doc.close()
        return "\n".join(all_text)
    except Exception as e:
        print(f"[錯誤] PDF 文字提取失敗: {e}", flush=True)
        return ""

def detect_transcript_info(text: str) -> Dict[str, str]:
    """檢測謄本基本資訊"""
    info = {
        "謄本類別": "未知",
        "謄本類型": "未知", 
        "地號建號": "未找到",
        "列印時間": "未找到"
    }
    
    if re.search(r'登記第二類謄本', text):
        info["謄本類別"] = "第二類"
    elif re.search(r'登記第一類謄本', text):
        info["謄本類別"] = "第一類"
    
    if re.search(r'建物登記第[一二]類謄本', text):
        info["謄本類型"] = "建物謄本"
    elif re.search(r'土地登記第[一二]類謄本', text):
        info["謄本類型"] = "土地謄本"
    
    location_match = re.search(r'([\u4e00-\u9faf]+(?:鄉|鎮|市|巿|區)[\u4e00-\u9faf]+段(?:[\u4e00-\u9faf]*小段)?\s+[\d\-]+(?:地號|建號))', text)
    if location_match:
        info["地號建號"] = location_match.group(1)
    
    time_match = re.search(r'列印時間[:：]\s*([^\n]+?)(?:\s+頁次[:：].*?)?(?:\n|$)', text)
    if time_match:
        info["列印時間"] = time_match.group(1)
    
    return info

def clean_pagination_simple(text: str) -> str:
    """簡化版跨頁清理"""
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        stripped = line.strip()
        
        if any([
            '續次頁' in stripped,
            re.search(r'^[0-9A-F]{1,4}\s+[0-9A-F]{1,4}\s+', stripped),
            re.search(r'列印時間：.*?頁次：\d+', stripped),
            stripped.endswith('地政事務所') and len(stripped) < 20,
            re.search(r'^\s*頁次[:：]\s*\d+\s*$', stripped),
        ]):
            continue
            
        cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)

def split_documents_by_title(text: str) -> List[str]:
    """分割多份謄本"""
    patterns = [
        r"土地登記第[一二]類謄本.*?(?=土地登記第[一二]類謄本|建物登記第[一二]類謄本|$)",
        r"建物登記第[一二]類謄本.*?(?=土地登記第[一二]類謄本|建物登記第[一二]類謄本|$)"
    ]
    
    documents = []
    for pattern in patterns:
        matches = re.findall(pattern, text, re.DOTALL)
        documents.extend(matches)
    
    return [doc.strip() for doc in documents if doc.strip()]

def extract_basic_fields(text: str) -> Dict[str, str]:
    """提取基本欄位"""
    data = {}
    
    patterns = {
        "謄本類型": r"([土地建物]+登記第[一二]類謄本.*?)(?:\n|（)",
        "地號建號": r"([\u4e00-\u9faf]+(?:鄉|鎮|市|巿|區)[\u4e00-\u9faf]+段(?:[\u4e00-\u9faf]*小段)?\s+[\d\-]+(?:地號|建號))",
        "列印時間": r"列印時間[:：]\s*([^\n]+?)(?:\s+頁次[:：].*?)?(?:\n|$)",
        "登記日期": r"登記日期[:：]\s*([^\n]+)",
        "登記原因": r"登記原因[:：]\s*([^\n]+)",
        "建物門牌": r"建物門牌[:：]\s*([^\n]+)",
        "主要用途": r"主要用途[:：]\s*([^\n]+)",
        "總面積": r"總面積[:：]\s*([^\n]+)",
        "所有權人": r"所有權人[:：]\s*(.+?)(?=\s*統一編號)",  # 🔥 修正：支援跨行
        "統一編號": r"統一編號[:：]\s*([^\n]+)",
        "住址": r"住址[:：]\s*([^\n]+)",
        "權利範圍": r"權利範圍[:：]\s*([^\n]+)",
        "權利種類": r"權利種類[:：]\s*([^\n]+)",
        "擔保債權總金額": r"擔保債權總金額[:：]\s*([^\n]+)",
    }

    for field, pattern in patterns.items():
        flags = re.MULTILINE | re.DOTALL if field == "所有權人" else re.MULTILINE
        match = re.search(pattern, text, flags)
        if match:
            value = match.group(1).strip()
            # 🔥 清理多餘空白和換行
            value = re.sub(r'\s+', ' ', value).strip()

            data[field] = value

    return data

# ==================== 原模組導入 ====================
def import_original_modules():
    """直接導入 pdf_parser.py 模組"""
    try:
        print("🔍 嘗試導入 pdf_parser.py 模組...", flush=True)

        # 🔥 直接指定 pdf_parser.py 的可能位置
        search_dirs = [
            BASE_DIR,  # 程式基礎目錄
            os.getcwd(),  # 當前目錄
        ]

        # 如果是打包環境，也搜尋 sys._MEIPASS
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            search_dirs.insert(0, sys._MEIPASS)

        module_filename = 'pdf_parser.py'

        for search_dir in search_dirs:
            if not os.path.exists(search_dir):
                continue

            full_path = os.path.join(search_dir, module_filename)

            if os.path.exists(full_path):
                try:
                    print(f"  📂 找到模組: {full_path}", flush=True)

                    import importlib.util
                    spec = importlib.util.spec_from_file_location("pdf_parser_module", full_path)
                    original_module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(original_module)

                    print(f"  ✅ 成功導入模組: {module_filename}", flush=True)

                    # 嘗試找到處理器和匯出器類別
                    processor_classes = [
                        'UnifiedTranscriptProcessor',
                        'SecondClassProcessor',
                        'TranscriptProcessor'
                    ]

                    exporter_classes = [
                        'UnifiedTranscriptExporter',
                        'SecondClassExporter',
                        'TranscriptExporter'
                    ]

                    processor_class = None
                    exporter_class = None

                    for class_name in processor_classes:
                        processor_class = getattr(original_module, class_name, None)
                        if processor_class:
                            print(f"    ✅ 找到處理器類別: {class_name}", flush=True)
                            break

                    for class_name in exporter_classes:
                        exporter_class = getattr(original_module, class_name, None)
                        if exporter_class:
                            print(f"    ✅ 找到匯出器類別: {class_name}", flush=True)
                            break

                    if processor_class and exporter_class:
                        print(f"  🎯 pdf_parser.py 包含所需的類別", flush=True)
                        return processor_class, exporter_class
                    else:
                        print(f"  ❌ pdf_parser.py 缺少必要的類別", flush=True)

                except Exception as e:
                    print(f"    ❌ 導入 pdf_parser.py 失敗: {e}", flush=True)
                    continue

        print("  ⚠️ 無法找到或載入 pdf_parser.py", flush=True)
        return None, None

    except Exception as e:
        print(f"❌ 導入過程發生錯誤: {e}", flush=True)
        return None, None

# 導入原有模組
# 🔧 延後到 main() 顯示啟動畫面後再載入（pdf_parser 會帶入 OCR/pandas/fitz，是啟動最耗時的部分），
#    這樣使用者一開始就看得到「載入中」提示，而不是對著黑畫面乾等。
ProcessorClass, ExporterClass = None, None

# ==================== 增強處理器 ====================
class EnhancedTranscriptProcessor:
    """修正版增強處理器 - 完整支援他項權利部"""
    
    def __init__(self):
        # 引入原程式的模式和邏輯
        self.patterns = self._build_enhanced_patterns()
    
    def _build_enhanced_patterns(self):
        """建立完整的謄本解析模式（從原程式複製）"""
        return {
            # === 基本資訊 ===
            "謄本類型": r"([土地建物]+登記第[一二]類謄本.*?)(?:\n|（)",
            "地號建號": r"([\u4e00-\u9faf]+(?:鄉|鎮|市|巿|區)[\u4e00-\u9faf]+段(?:[\u4e00-\u9faf]*小段)?\s+[\d\-]+(?:地號|建號))",
            "列印時間": r"列印時間[:：]\s*([^\n]+?)(?:\s+頁次[:：].*?)?(?:\n|$)",
            
            # === 他項權利部完整模式 ===
            "權利種類": r"權利種類[:：]\s*([^\s\n]+)",
            "收件年期": r"收件年期[:：]\s*([^\s]+)",
            "字號": r"字號[:：]\s*([^\s\n]+)",
            "他項_登記日期": r"登記日[期日][:：]\s*([^\s]+(?:\s+[^\s]+)*?)(?:\s+登記原因|$)",
            "他項_登記原因": r"登記原因[:：]\s*([^\s\n]+)",
            "權利人": r"權\s*利\s*人[:：]\s*([^\n]+)",
            "統一編號_他項": r"統一編號[:：]\s*([^\s\n]+)",
            "住址_他項": r"(?:住|居)\s*址[:：]\s*([^\n]+(?:\n(?!債權額比例|擔保債權總金額)[^\n]+)*)",
            "債權額比例": r"債權額比例[:：]\s*([^\n]+)",
            "擔保債權總金額": r'擔保債權總金額[:：]\s*(新台幣[^0-9]*[\d,]+\.?\d*\s*元正)',
            "存續期間": r"存續期間[:：]\s*([^\n]+)",
            "擔保債權種類及範圍": r'擔保債權種類及範圍[:：]\s*(.*?)(?=\s*擔保債權確定期[日期][:：])',
            "擔保債權確定期日": r"擔保債權確定期[日期][:：]\s*([^\n。]*?)(?=\s*[。\n]|清償日[期日][:：])",
            "清償日期": r"清償日[期日][:：]\s*(.*?)(?=\s*利息\s*[（(]?率[）)]?[:：])",
            "利息率": r"利息\s*[（(]?率[）)]?[:：]\s*(.*?)(?=\s*遲延利息\s*[（(]?率[）)]?[:：])",
            "遲延利息率": r"遲延利息\s*[（(]?率[）)]?[:：]\s*(.*?)(?=\s*違\s*約\s*金[:：])",
            "違約金": r"違\s*約\s*金[:：]\s*(.*?)(?=\s*其他擔保範圍約定[:：])",
            "其他擔保範圍約定": r"其他擔保範圍約定[:：]\s*(.*?)(?=\s*債務人及債務額比例[:：]|\s*權利標的[:：]|\s*標的登記次序[:：])",
            "債務人及債務額比例": r"債務人及債務額比例[:：]\s*(.*?)(?=\s*權利標的[:：]|$)",
            "權利標的": r"權利標的[:：]\s*([^\s\n]+)",
            "標的登記次序": r"標的登記次序[:：]\s*([^\n]+)",
            "設定權利範圍": r"設定權利範圍[:：]\s*([^\n]+)",
            "證明書字號": r"證明書字號[:：]\s*([^\s\n]+)",
            "設定義務人": r"設定義務人[:：]\s*([^\n]+)",
            "共同擔保地號": r"共同擔保地號[:：]\s*(.*?)(?=\s*共同擔保建號|其他登記事項|$)",
            "共同擔保建號": r"共同擔保建號[:：]\s*(.*?)(?=\s*其他登記事項|$)",
        }

    def process_multiple_pdfs(self, pdf_paths: List[str]) -> List[Dict[str, Any]]:
        """修正版處理方法 - 確保他項權利部完整處理"""
        all_results = []
        
        for i, pdf_path in enumerate(pdf_paths, 1):
            print(f"📄 修正版處理第 {i} 個PDF: {os.path.basename(pdf_path)}", flush=True)
            
            if not os.path.exists(pdf_path):
                continue
            
            try:
                # 提取文字
                full_text = extract_pdf_text_only(pdf_path)
                if not full_text:
                    continue
                
                # 清理跨頁
                cleaned_text = clean_pagination_simple(full_text)
                
                # 分割文件
                documents = split_documents_by_title(cleaned_text)
                if not documents:
                    documents = [cleaned_text]
                
                for doc_i, doc_text in enumerate(documents):
                    try:
                        # 🔧 關鍵修正：使用完整的解析邏輯
                        structured = self._extract_complete_transcript_data(doc_text)
                        
                        structured["源檔案"] = os.path.basename(pdf_path)
                        structured["檔案序號"] = i
                        structured["謄本序號"] = doc_i + 1
                        
                        all_results.append(structured)
                        
                    except Exception as e:
                        print(f"謄本 {doc_i+1} 解析失敗: {e}", flush=True)
                        continue
                        
            except Exception as e:
                print(f"PDF {i} 處理失敗: {e}", flush=True)
                continue
        
        return all_results
    
    def _extract_complete_transcript_data(self, text: str) -> Dict[str, Any]:
        """完整的謄本資料提取 - 確保他項權利部不遺失"""
        
        print("🔄 開始完整謄本解析...", flush=True)
        
        # 1. 基本資訊
        basic_info = self._extract_basic_info_enhanced(text)
        transcript_info = detect_transcript_info(text)
        
        # 2. 章節分割（使用更精確的方法）
        sections = self._split_sections_precise(text)
        
        # 3. 檢查他項權利部是否存在
        rights_section = sections.get("他項權利部", "")
        print(f"📊 他項權利部檢查:", flush=True)
        print(f"   章節存在: {'✅' if rights_section else '❌'}", flush=True)
        print(f"   內容長度: {len(rights_section)} 字元", flush=True)
        
        if rights_section and len(rights_section) > 50:  # 有實質內容
            print("   內容預覽: " + rights_section[:100] + "...", flush=True)
        
        # 4. 完整解析各部分
        structured = OrderedDict([
            ("謄本類型", f"{transcript_info['謄本類型']} - {transcript_info['謄本類別']}"),
            ("基本資訊", basic_info),
            ("標示部", self._extract_indicator_enhanced(sections.get("標示部", ""))),
            ("所有權部", self._extract_ownership_enhanced(sections.get("所有權部", ""))),
            ("他項權利部", self._extract_rights_complete(sections.get("他項權利部", "")))  # 🔧 關鍵修正
        ])
        
        # 5. 驗證他項權利部資料
        rights_data = structured["他項權利部"]
        if isinstance(rights_data, list) and len(rights_data) > 0:
            print(f"✅ 他項權利部解析成功: {len(rights_data)} 筆記錄", flush=True)
            
            # 驗證關鍵欄位
            for i, record in enumerate(rights_data):
                key_fields = ["權利種類", "權利人", "擔保債權總金額"]
                found_fields = [field for field in key_fields if record.get(field)]
                print(f"   記錄{i+1}: 包含 {len(found_fields)} 個關鍵欄位 ({found_fields})", flush=True)
        else:
            print(f"⚠️ 他項權利部可能為空或解析失敗", flush=True)
        
        return structured
    
    def _split_sections_precise(self, text: str) -> Dict[str, str]:
        """精確的章節分割 - 確保他項權利部完整"""
        sections = {}

        print("🔍 精確章節分割...", flush=True)

        # 使用星號標記進行精確分割
        # 🔧 修正：他項權利部使用貪婪匹配 (.*) 而非非貪婪 (.*?)，確保匹配到「本謄本列印完畢」之前的所有內容
        patterns = {
            "標示部": r"(\*{10,}\s*(?:[土地建物]*)?標示部\s*\*{10,})(.*?)(?=\*{10,}.*?所有權部|\*{10,}.*?他項權利部|$)",
            "所有權部": r"(\*{10,}\s*(?:[土地建物]*)?所有權部\s*\*{10,})(.*?)(?=\*{10,}.*?他項權利部|$)",
            "他項權利部": r"(\*{10,}\s*(?:[土地建物]*)?他項權利部\s*\*{10,})(.*?)(?=〈\s*本謄本列印完畢\s*〉|$)"
        }
        
        for section_name, pattern in patterns.items():
            match = re.search(pattern, text, re.DOTALL)
            if match:
                content = match.group(2).strip()
                sections[section_name] = content
                print(f"   ✅ {section_name}: {len(content)} 字元", flush=True)
                
                # 特別檢查他項權利部
                if section_name == "他項權利部":
                    key_terms = ["權利種類", "權利人", "擔保債權總金額", "共同擔保"]
                    found_terms = [term for term in key_terms if term in content]
                    print(f"      包含關鍵詞: {found_terms}", flush=True)
            else:
                print(f"   ❌ {section_name}: 未找到", flush=True)
                sections[section_name] = ""
        
        return sections
    
    def _extract_basic_info_enhanced(self, text: str) -> Dict[str, str]:
        """增強版基本資訊提取"""
        data = {}
        
        basic_patterns = {
            "謄本類型": r"([土地建物]+登記第[一二]類謄本.*?)(?:\n|（)",
            "地號建號": r"([\u4e00-\u9faf]+(?:鄉|鎮|市|巿|區)[\u4e00-\u9faf]+段(?:[\u4e00-\u9faf]*小段)?\s+[\d\-]+(?:地號|建號))",
            "列印時間": r"列印時間[:：]\s*([^\n]+?)(?:\s+頁次[:：].*?)?(?:\n|$)"
        }
        
        for field, pattern in basic_patterns.items():
            match = re.search(pattern, text, re.MULTILINE)
            if match:
                data[field] = match.group(1).strip()
        
        return data
    
    def _extract_indicator_enhanced(self, section_text: str) -> Dict[str, str]:
        """增強版標示部提取"""
        if not section_text:
            return {}
        
        # 使用基本欄位提取，但可以擴展
        return extract_basic_fields(section_text)
    
    def _extract_ownership_enhanced(self, section_text: str) -> List[Dict[str, str]]:
        """增強版所有權部提取 - 支援多筆記錄和統一編號"""
        if not section_text:
            return []

        print(f"🔄 解析所有權部（長度: {len(section_text)} 字元）...", flush=True)

        ownership_list = []

        # 1. 檢查是否有多筆記錄（例如：（0001）登記次序：0002）
        record_pattern = r'（(\d+)）登記次序[:：]\s*([\d\-]+)'
        records = re.findall(record_pattern, section_text)

        if not records:
            print("   使用整個章節作為單一記錄", flush=True)
            records = [('0001', '0001')]
            record_contents = [section_text]
        else:
            print(f"   找到 {len(records)} 個所有權記錄", flush=True)
            record_contents = []

            # 分割每個記錄的內容
            for i, (num, seq) in enumerate(records):
                start_pattern = f'（{num}）登記次序[:：]\\s*{re.escape(seq)}'
                start_match = re.search(start_pattern, section_text)

                if start_match:
                    start_pos = start_match.start()

                    if i + 1 < len(records):
                        next_num, next_seq = records[i + 1]
                        end_pattern = f'（{next_num}）登記次序[:：]\\s*{re.escape(next_seq)}'
                        end_match = re.search(end_pattern, section_text)
                        end_pos = end_match.start() if end_match else len(section_text)
                    else:
                        end_pos = len(section_text)

                    record_content = section_text[start_pos:end_pos].strip()
                    record_contents.append(record_content)

        # 2. 處理每個記錄，提取所有權人、統一編號等欄位
        for i, ((num, seq), content) in enumerate(zip(records, record_contents)):
            print(f"   🔄 處理所有權記錄 {i+1}: 編號={num}, 次序={seq}", flush=True)

            record_data = {}
            record_data["記錄編號"] = num
            record_data["登記次序"] = seq

            # 3. 提取各個欄位（包括統一編號）
            field_patterns = {
                "登記日期": r"登記日期[:：]\s*([^\n]+)",
                "登記原因": r"登記原因[:：]\s*([^\n]+)",
                "原因發生日期": r"原因發生日期[:：]\s*([^\n]+)",
                "所有權人": r"所有權人[:：]\s*(.+?)(?=\s*統一編號)",  # 🔥 修正：匹配到「統一編號」之前，支援跨行
                "統一編號": r"統一編號[:：]\s*([^\n]+)",
                "出生日期": r"出生日期[:：]\s*([^\n]+)",
                "住址": r"住址[:：]\s*([^\n]+)",
                "權利範圍": r"權利範圍[:：]\s*([^\n]+)",
                "權狀字號": r"權狀字號[:：]\s*([^\n]+)",
                "當期申報地價": r"當期申報地價[:：]\s*([^\n]+)",
                "前次移轉現值或原規定": r"前次移轉現值或原規定(?:地價)?[:：]\s*([^\n]+)",
            }

            for field_name, pattern in field_patterns.items():
                flags = re.MULTILINE | re.DOTALL if field_name == "所有權人" else re.MULTILINE
                match = re.search(pattern, content, flags)
                if match:
                    value = match.group(1).strip()
                    # 🔥 清理多餘空白和換行
                    value = re.sub(r'\s+', ' ', value).strip()
                    record_data[field_name] = value
                    print(f"      ✅ {field_name}: {value[:50]}...", flush=True)
                else:
                    record_data[field_name] = ""
                    if field_name in ["所有權人", "統一編號"]:  # 只對重要欄位顯示警告
                        print(f"      ⚠️ 未找到 {field_name}", flush=True)

            # 4. 提取歷次取得權利範圍（可能有多筆）
            history_pattern = r'序號(\d+)[:：]\s*([^\n]*)\n.*?年月\1[:：]\s*([^\n]*)\n.*?(?:地價|元/平方公尺)\1[:：]\s*([^\n]*)\n.*?歷次取得權利範圍\1[:：]\s*([^\n]*)'
            history_matches = re.findall(history_pattern, content, re.MULTILINE | re.DOTALL)
            for idx, (序號, 年月, 地價, 範圍) in enumerate(history_matches, 1):
                record_data[f"序號{idx}"] = 序號.strip()
                record_data[f"年月{idx}"] = 年月.strip()
                record_data[f"地價{idx}"] = 地價.strip()
                record_data[f"歷次取得權利範圍{idx}"] = 範圍.strip()

            # 5. 提取其他登記事項
            other_match = re.search(r'其他登記事項[:：]\s*([^\n]+(?:\n(?!（\d+）)[^\n]+)*)', content, re.MULTILINE)
            if other_match:
                record_data["其他登記事項"] = other_match.group(1).strip()
            else:
                record_data["其他登記事項"] = ""

            ownership_list.append(record_data)
            print(f"   📊 所有權記錄 {i+1} 完成，共 {len(record_data)} 個欄位", flush=True)

        print(f"✅ 所有權部解析完成，共 {len(ownership_list)} 筆記錄", flush=True)
        return ownership_list
    
    def _extract_rights_complete(self, section_text: str) -> List[Dict[str, str]]:
        """🔧 關鍵修正：完整的他項權利部提取"""
        
        if not section_text or len(section_text.strip()) < 20:
            print("❌ 他項權利部內容為空或太短", flush=True)
            return []
        
        print(f"🔄 完整解析他項權利部（長度: {len(section_text)} 字元）...", flush=True)
        
        rights_list = []
        
        # 1. 檢查是否有多筆記錄
        record_pattern = r'（(\d+)）登記次序[:：]\s*([\d\-]+)'
        records = re.findall(record_pattern, section_text)
        
        if not records:
            print("   使用整個章節作為單一記錄", flush=True)
            records = [('0001', '0001')]
            record_contents = [section_text]
        else:
            print(f"   找到 {len(records)} 個他項權利記錄", flush=True)
            record_contents = []
            
            # 分割每個記錄的內容
            for i, (num, seq) in enumerate(records):
                start_pattern = f'（{num}）登記次序[:：]\\s*{re.escape(seq)}'
                start_match = re.search(start_pattern, section_text)
                
                if start_match:
                    start_pos = start_match.start()
                    
                    if i + 1 < len(records):
                        next_num, next_seq = records[i + 1]
                        end_pattern = f'（{next_num}）登記次序[:：]\\s*{re.escape(next_seq)}'
                        end_match = re.search(end_pattern, section_text)
                        end_pos = end_match.start() if end_match else len(section_text)
                    else:
                        end_pos = len(section_text)
                    
                    record_content = section_text[start_pos:end_pos].strip()
                    record_contents.append(record_content)
        
        # 2. 處理每個記錄
        for i, ((num, seq), content) in enumerate(zip(records, record_contents)):
            print(f"   🔄 處理記錄 {i+1}: 編號={num}, 次序={seq}", flush=True)
            
            record_data = {}
            record_data["記錄編號"] = num
            record_data["登記次序"] = seq
            
            # 3. 提取各個欄位
            for field_name, pattern in self.patterns.items():
                if field_name.startswith("他項_") or field_name in [
                    "權利種類", "收件年期", "字號", "權利人", "統一編號_他項", "住址_他項",
                    "債權額比例", "擔保債權總金額", "存續期間", "擔保債權種類及範圍",
                    "擔保債權確定期日", "清償日期", "利息率", "遲延利息率", "違約金",
                    "其他擔保範圍約定", "債務人及債務額比例", "權利標的", "標的登記次序",
                    "設定權利範圍", "證明書字號", "設定義務人", "共同擔保地號", "共同擔保建號"
                ]:
                    match = re.search(pattern, content, re.MULTILINE | re.DOTALL)
                    if match:
                        # 清理欄位名稱
                        clean_name = field_name.replace("他項_", "").replace("_他項", "")
                        if clean_name == "統一編號_他項":
                            clean_name = "統一編號"
                        elif clean_name == "住址_他項":
                            clean_name = "住址"
                        
                        value = self._clean_value(match.group(1))
                        record_data[clean_name] = value
                        print(f"      ✅ {clean_name}: {value[:50]}...", flush=True)
                    else:
                        clean_name = field_name.replace("他項_", "").replace("_他項", "")
                        if clean_name == "統一編號_他項":
                            clean_name = "統一編號"
                        elif clean_name == "住址_他項":
                            clean_name = "住址"
                        
                        record_data[clean_name] = ""
                        print(f"      ❌ 未找到 {clean_name}", flush=True)
            
            # 4. 提取其他登記事項
            other_items = self._extract_other_items_for_rights(content)
            if other_items:
                record_data["其他登記事項"] = other_items
            
            rights_list.append(record_data)
            print(f"   📊 記錄 {i+1} 完成，共 {len(record_data)} 個欄位", flush=True)
        
        print(f"✅ 他項權利部解析完成，共 {len(rights_list)} 筆記錄", flush=True)
        
        # 5. 驗證結果
        if rights_list:
            total_fields = sum(len(record) for record in rights_list)
            print(f"📊 他項權利部統計:", flush=True)
            print(f"   記錄數: {len(rights_list)}", flush=True)
            print(f"   總欄位數: {total_fields}", flush=True)
            
            # 檢查關鍵欄位
            key_fields = ["權利種類", "權利人", "擔保債權總金額"]
            for i, record in enumerate(rights_list):
                found_keys = [field for field in key_fields if record.get(field)]
                print(f"   記錄{i+1}關鍵欄位: {found_keys}", flush=True)
        
        return rights_list
    
    def _extract_other_items_for_rights(self, content: str) -> str:
        """提取他項權利部的其他登記事項"""
        
        # 尋找其他登記事項
        pattern = r"其他登記事項[:：]\s*(.*?)(?=\s*共同擔保|（\d+）登記次序|$)"
        match = re.search(pattern, content, re.DOTALL)
        
        if match:
            other_content = match.group(1).strip()
            # 清理內容
            other_content = re.sub(r'\*+', '', other_content)
            other_content = re.sub(r'\s+', ' ', other_content)
            
            if other_content and other_content not in ["（空白）", "(空白)"]:
                return other_content
        
        return ""
    
    def _clean_value(self, value: str) -> str:
        """清理值"""
        if not value:
            return ""
        
        # 移除星號（保留統一編號的*號）
        if re.match(r'^[A-Z]\d\*+\d+$', value.strip()):
            # 統一編號格式，保留*號
            cleaned = value
        else:
            cleaned = re.sub(r'\*+', '', value)
        
        # 清理跨頁標記
        cleaned = re.sub(r'[（(]續次頁[）)]\s*', '', cleaned)
        
        # 處理多行
        lines = cleaned.split('\n')
        cleaned_lines = [line.strip() for line in lines if line.strip()]
        
        return ' '.join(cleaned_lines) if cleaned_lines else ""

class EnhancedTranscriptExporter:
    """增強版匯出器"""
    def __init__(self, gui_instance=None):
        # 🔧 修正：正確設定 GUI 實例
        self.gui_instance = gui_instance
    
    def get_output_directory(self):
        """取得當前設定的輸出目錄"""
        if self.gui_instance and hasattr(self.gui_instance, 'output_directory'):
            output_dir = self.gui_instance.output_directory
        else:
            # 回傳預設目錄
            # 🔥 使用 get_work_folder 確保在主程式目錄下建立資料夾
            output_dir = get_work_folder("output")

        # 確保目錄存在
        os.makedirs(output_dir, exist_ok=True)
        return output_dir
    
    @property
    def output_dir(self):
        """動態取得輸出目錄"""
        return self.get_output_directory()
    
    def export_to_txt_report(self, data: List[Dict], filename: str = None):
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"transcript_report_{timestamp}.txt"
        
        # 🔧 使用動態輸出目錄
        filepath = os.path.join(self.output_dir, filename)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write("不動產登記謄本處理報告\n")
            f.write("="*60 + "\n")
            f.write(f"生成時間: {datetime.now()}\n")
            f.write(f"總謄本數: {len(data)}\n\n")
            
            for i, transcript in enumerate(data, 1):
                f.write(f"謄本 {i}: {transcript.get('源檔案', '未知檔案')}\n")
                f.write("-"*40 + "\n")
                self._write_dict_to_file(f, transcript, indent=0)
                f.write("\n\n")
        
        return filepath
    
    def _write_dict_to_file(self, f, data, indent=0):
        prefix = "  " * indent
        
        if isinstance(data, dict):
            for key, value in data.items():
                if isinstance(value, (dict, list)):
                    f.write(f"{prefix}{key}:\n")
                    self._write_dict_to_file(f, value, indent + 1)
                else:
                    f.write(f"{prefix}{key}: {value}\n")
        elif isinstance(data, list):
            for i, item in enumerate(data):
                f.write(f"{prefix}[{i+1}]:\n")
                self._write_dict_to_file(f, item, indent + 1)
        else:
            f.write(f"{prefix}{data}\n")
    
    def export_to_excel(self, data: List[Dict], filename: str = None):
        """匯出到 Excel，按土地/建物和標示部/所有權部/他項權利部分工作表"""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"transcript_data_{timestamp}.xlsx"
        
        # 使用動態輸出目錄
        filepath = os.path.join(self.output_dir, filename)
        
        # 分類資料：土地/建物 x 標示部/所有權部/他項權利部
        categorized_data = {
            '土地-標示部': [],
            '土地-所有權部': [],
            '土地-他項權利部': [],
            '建物-標示部': [],
            '建物-所有權部': [],
            '建物-他項權利部': []
        }
        
        for transcript in data:
            # 判斷是土地還是建物
            transcript_type = str(transcript.get('謄本類型', ''))
            is_land = '土地' in transcript_type
            is_building = '建物' in transcript_type
            
            # 扁平化謄本資料
            flattened = self._flatten_dict(transcript)
            
            # 根據謄本類型分類到對應的工作表
            if is_land:
                # 土地謄本：提取各部資料
                if '標示部' in transcript or any('標示部' in str(k) for k in flattened.keys()):
                    categorized_data['土地-標示部'].append(flattened)
                if '所有權部' in transcript or any('所有權部' in str(k) for k in flattened.keys()):
                    categorized_data['土地-所有權部'].append(flattened)
                if '他項權利部' in transcript or any('他項權利部' in str(k) for k in flattened.keys()):
                    categorized_data['土地-他項權利部'].append(flattened)
                # 如果沒有明確分部，就都加入
                if not any('部' in str(k) for k in flattened.keys() if '部' in str(k)):
                    categorized_data['土地-標示部'].append(flattened)
            elif is_building:
                # 建物謄本：提取各部資料
                if '標示部' in transcript or any('標示部' in str(k) for k in flattened.keys()):
                    categorized_data['建物-標示部'].append(flattened)
                if '所有權部' in transcript or any('所有權部' in str(k) for k in flattened.keys()):
                    categorized_data['建物-所有權部'].append(flattened)
                if '他項權利部' in transcript or any('他項權利部' in str(k) for k in flattened.keys()):
                    categorized_data['建物-他項權利部'].append(flattened)
                # 如果沒有明確分部，就都加入
                if not any('部' in str(k) for k in flattened.keys() if '部' in str(k)):
                    categorized_data['建物-標示部'].append(flattened)
        
        # 使用 openpyxl engine 並創建 ExcelWriter
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            for sheet_name, records in categorized_data.items():
                if not records:  # 跳過空的工作表
                    continue
                
                # 創建 DataFrame
                df = pd.DataFrame(records)
                
                # 寫入工作表
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 🔧 取得工作表物件並自動調整欄寬（改進版）
                worksheet = writer.sheets[sheet_name]

                # 🔧 自動調整所有欄位寬度
                for column_cells in worksheet.columns:
                    max_length = 0
                    column_letter = column_cells[0].column_letter

                    for cell in column_cells:
                        try:
                            # 計算儲存格內容的長度（考慮中文字元）
                            if cell.value:
                                cell_value = str(cell.value)
                                # 中文字元計為2個單位，英文計為1個單位
                                cell_length = sum(2 if ord(c) > 127 else 1 for c in cell_value)
                                max_length = max(max_length, cell_length)
                        except:
                            pass

                    # 設定欄寬，增加一些邊距讓內容不會太擠
                    adjusted_width = max_length + 2

                    # 設定最小寬度為10，最大寬度為100
                    adjusted_width = max(10, min(adjusted_width, 100))

                    worksheet.column_dimensions[column_letter].width = adjusted_width
        
        return filepath


    def _flatten_dict(self, d, parent_key='', sep='_'):
        items = []
        if isinstance(d, dict):
            for k, v in d.items():
                new_key = f"{parent_key}{sep}{k}" if parent_key else k
                if isinstance(v, dict):
                    items.extend(self._flatten_dict(v, new_key, sep=sep).items())
                elif isinstance(v, list):
                    for i, item in enumerate(v):
                        items.extend(self._flatten_dict(item, f"{new_key}_{i+1}", sep=sep).items())
                else:
                    items.append((new_key, v))
        return dict(items)
    
    def export_to_json(self, data: List[Dict], filename: str = None):
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"transcript_data_{timestamp}.json"
        
        # 🔧 使用動態輸出目錄
        filepath = os.path.join(self.output_dir, filename)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return filepath

    def _create_flexible_filename(self, data: List[Dict], format_suffix: str) -> str:
        """靈活檔案命名系統 - 支援一類+二類,友善的中文檔名"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        
        # 分析謄本類型
        first_class_count = sum(1 for t in data if "第一類" in str(t.get("謄本類型", "")))
        second_class_count = sum(1 for t in data if "第二類" in str(t.get("謄本類型", "")))
        
        # 類別描述
        if first_class_count > 0 and second_class_count > 0:
            class_part = f"一類{first_class_count}份_二類{second_class_count}份"
        elif first_class_count > 0:
            class_part = f"一類{first_class_count}份"
        elif second_class_count > 0:
            class_part = f"二類{second_class_count}份"
        else:
            class_part = f"{len(data)}份"
        
        # 嘗試從第一筆提取地段或地址資訊
        location_part = ""
        if data:
            first_item = data[0]
            if "地段地號" in first_item and first_item["地段地號"]:
                location_part = str(first_item["地段地號"]).split()[0][:15]  # 取地段部分
            elif "建號" in first_item and first_item["建號"]:
                location_part = f"建號{first_item['建號']}"[:15]
            elif "地址" in first_item and first_item["地址"]:
                addr = str(first_item["地址"])
                # 提取鄉鎮市區
                import re
                match_area = re.search(r'([^縣市]+[縣市])([^鄉鎮市區]+[鄉鎮市區])', addr)
                if match_area:
                    location_part = match_area.group(2)[:10]
        
        # 謄本類型描述
        type_part = "謄本"
        if "土地" in str(data[0].get("謄本類型", "")) if data else False:
            type_part = "土地"
        elif "建物" in str(data[0].get("謄本類型", "")) if data else False:
            type_part = "建物"
        
        # 組合檔名(加上 trans_ 前綴表示結構化謄本)
        if location_part:
            if format_suffix == "base":
                return f"trans_{location_part}_{type_part}謄本_{class_part}_{timestamp}"
            else:
                return f"trans_{location_part}_{type_part}謄本_{class_part}_{timestamp}.{format_suffix}"
        else:
            if format_suffix == "base":
                return f"trans_{type_part}謄本_{class_part}_{timestamp}"
            else:
                return f"trans_{type_part}謄本_{class_part}_{timestamp}.{format_suffix}"

# ==================== 自定義美化元件 ====================

    def _create_flexible_filename(self, data: List[Dict], format_suffix: str) -> str:
        """靈活檔案命名系統 - 支援一類+二類,友善的中文檔名"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")

        # 分析謄本類型
        first_class_count = sum(1 for t in data if "第一類" in str(t.get("謄本類型", "")))
        second_class_count = sum(1 for t in data if "第二類" in str(t.get("謄本類型", "")))
        building_count = sum(1 for t in data if "建物謄本" in str(t.get("謄本類型", "")))
        land_count = sum(1 for t in data if "土地謄本" in str(t.get("謄本類型", "")))

        # 提取地段地號資訊
        location_part = ""
        if data and len(data) > 0:
            # 從第一份謄本的基本資訊中提取地號建號
            first_transcript = data[0]
            basic_info = first_transcript.get("基本資訊", {})
            lot_info = basic_info.get("地號建號", "")

            if lot_info:
                # 解析地號建號,例如:"大寮區仁德段 0193-0000地號"
                # 提取區段和地號/建號
                match = re.match(r'([\u4e00-\u9fff]+區)?([\u4e00-\u9fff]+段)\s*(\d+[\-\d]*)(地號|建號)', lot_info)
                if match:
                    area = match.group(1) or ""  # 可能沒有區
                    section = match.group(2)
                    number = match.group(3).replace('-0000', '').replace('-000', '').lstrip('0') or '0'
                    
                    # 移除多餘的0,例如 0193 -> 193
                    if '-' in number:
                        parts = number.split('-')
                        number = parts[0].lstrip('0') or '0'
                        if parts[1] != '0000' and parts[1] != '000':
                            number += '-' + parts[1].lstrip('0')

                    location_part = f"{area}{section}{number}"

        # 智慧命名邏輯(中文化)
        if first_class_count > 0 and second_class_count > 0:
            class_part = "混合類"
        elif second_class_count > 0:
            class_part = "第二類"
        else:
            class_part = "第一類"

        if building_count > 0 and land_count > 0:
            type_part = "土地建物"
        elif building_count > 0:
            type_part = "建物"
        else:
            type_part = "土地"

        # 組合檔名(加上 trans_ 前綴表示結構化謄本)
        if location_part:
            # 有地段地號資訊
            return f"trans_{location_part}_{type_part}謄本_{class_part}_{timestamp}.{format_suffix}"
        else:
            # 沒有地段地號資訊,使用舊格式
            if format_suffix == "txt":
                return f"trans_{type_part}謄本_{class_part}_報告_{timestamp}.{format_suffix}"
            else:
                return f"trans_{type_part}謄本_{class_part}_{timestamp}.{format_suffix}"

class ModernButton(tk.Button):
    """現代化按鈕"""
    def __init__(self, parent, text="", command=None, style="primary", **kwargs):
        # 記錄樣式類型
        self.button_style = style
        
        # 根據樣式設定顏色
        if style == "primary":
            bg = ModernStyle.PRIMARY
            fg = ModernStyle.TEXT_LIGHT
            active_bg = "#1F5F85"
        elif style == "success":
            bg = ModernStyle.SUCCESS
            fg = ModernStyle.TEXT_LIGHT
            active_bg = "#1E7E34"
        elif style == "warning":
            bg = ModernStyle.WARNING
            fg = ModernStyle.TEXT_PRIMARY
            active_bg = "#E0A800"
        elif style == "danger":
            bg = ModernStyle.DANGER
            fg = ModernStyle.TEXT_LIGHT
            active_bg = "#C82333"
        elif style == "info":
            bg = ModernStyle.INFO
            fg = ModernStyle.TEXT_LIGHT
            active_bg = "#138496"
        else:  # secondary
            bg = ModernStyle.BG_SECONDARY
            fg = ModernStyle.TEXT_PRIMARY
            active_bg = "#495057"  # 🔧 改成深灰而不是黑色
        
        default_config = {
            "bg": bg,
            "fg": fg,
            "activebackground": active_bg,
            "activeforeground": fg,
            "font": ("微軟正黑體", 12),  # 原本是12，改成10
            "relief": "flat",
            "borderwidth": 0,
            "cursor": "hand2",
            "padx": 12,
            "pady": 8
        }
        
        # 合併使用者參數
        final_config = {**default_config, **kwargs}
        
        super().__init__(parent, text=text, command=command, **final_config)
        
        # 添加hover效果
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        
        self.default_bg = bg
        self.hover_bg = active_bg
    
    def _on_enter(self, event):
        """滑鼠進入事件"""
        # 🔧 修正：不同樣式使用不同的hover顏色
        if hasattr(self, 'button_style'):
            if self.button_style == "secondary":
                hover_color = "#495057"  # 深灰色，但不是純黑
            elif self.button_style == "danger":
                hover_color = "#C82333"
            elif self.button_style == "primary":
                hover_color = "#1F5F85"
            elif self.button_style == "success":
                hover_color = "#1E7E34"
            elif self.button_style == "warning":
                hover_color = "#E0A800"
            elif self.button_style == "info":
                hover_color = "#138496"
            else:
                hover_color = "#495057"
        else:
            hover_color = self.hover_bg
        
        self.config(bg=hover_color)

    def _on_leave(self, event):
        """滑鼠離開事件"""
        self.config(bg=self.default_bg)
    
    def _on_leave(self, event):
        self.config(bg=self.default_bg)

class ModernFrame(tk.Frame):
    """現代化框架"""
    def __init__(self, parent, title="", **kwargs):
        default_config = {
            "bg": ModernStyle.BG_LIGHT,
            "relief": "solid",
            "borderwidth": 1,
            "padx": 15,
            "pady": 15
        }
        
        final_config = {**default_config, **kwargs}
        super().__init__(parent, **final_config)
        
        if title:
            title_label = tk.Label(self, text=title, 
                                 font=ModernStyle.FONT_HEADER,
                                 bg=ModernStyle.BG_LIGHT,
                                 fg=ModernStyle.PRIMARY)
            title_label.pack(anchor="w", pady=(0, 10))

class ProgressIndicator(tk.Frame):
    """進度指示器"""
    def __init__(self, parent):
        super().__init__(parent, bg=ModernStyle.BG_LIGHT)
        
        self.progress_var = tk.StringVar()
        self.progress_var.set("就緒")
        
        # 狀態指示燈
        self.status_canvas = tk.Canvas(self, width=20, height=20, 
                                     bg=ModernStyle.BG_LIGHT, 
                                     highlightthickness=0)
        self.status_canvas.pack(side="left", padx=(0, 10))
        
        self.status_circle = self.status_canvas.create_oval(5, 5, 15, 15, 
                                                          fill=ModernStyle.SUCCESS, 
                                                          outline="")
        
        # 狀態文字
        self.status_label = tk.Label(self, textvariable=self.progress_var,
                                   font=ModernStyle.FONT_MAIN,
                                   bg=ModernStyle.BG_LIGHT,
                                   fg=ModernStyle.TEXT_SECONDARY)
        self.status_label.pack(side="left")
    
    def set_status(self, text, status="ready"):
        self.progress_var.set(text)
        
        if status == "processing":
            color = ModernStyle.WARNING
        elif status == "success":
            color = ModernStyle.SUCCESS
        elif status == "error":
            color = ModernStyle.DANGER
        else:  # ready
            color = ModernStyle.INFO
        
        self.status_canvas.itemconfig(self.status_circle, fill=color)

# 🔧 可選：添加更進階的狀態指示器
class StatusIndicator(tk.Frame):
    """狀態指示器組件"""
    def __init__(self, parent):
        super().__init__(parent, bg=ModernStyle.BG_LIGHT)
        
        self.status_canvas = tk.Canvas(self, width=12, height=12, 
                                     bg=ModernStyle.BG_LIGHT, 
                                     highlightthickness=0)
        self.status_canvas.pack(side="left", padx=(0, 5))
        
        self.status_dot = self.status_canvas.create_oval(2, 2, 10, 10, 
                                                        fill="#6C757D", 
                                                        outline="")
    
    def set_status(self, status="none"):
        """設置狀態顏色"""
        colors = {
            "none": "#6C757D",      # 灰色
            "preview": "#17A2B8",   # 藍色  
            "processing": "#FD7E14", # 橘色
            "processed": "#28A745",  # 綠色
            "brief": "#FFC107"       # 黃色
        }
        
        color = colors.get(status, colors["none"])
        self.status_canvas.itemconfig(self.status_dot, fill=color)

# ==================== 美化版主GUI ====================
class ModernTranscriptGUI:
    def __init__(self, master):
        self.master = master
        # 🔧 記錄程式啟動時間
        self.start_time = time.time()
        
        # 🔧 添加缺少的屬性 - 要放在最前面
        self.progress_indicator = self  # 暫時指向自己，避免錯誤
        
        self.set_custom_icon()
        self.pdf_listbox = None
        self.current_folder = None
        self.file_tree = None

        # 🔧 導航歷史（瀏覽器式上一頁/下一頁）
        self.back_stack = []     # 返回前頁用：記錄離開過的目錄
        self.forward_stack = []  # 前進下一頁用

        # 🔧 新增：慢速雙擊重新命名的變數
        self.last_click_time = 0
        # 🔧 設定自定義圖示
        #🔧 在初始化早期載入設定
        self.load_settings()

        self.last_click_item = None
        self.rename_delay = 1000  # 慢速雙擊間隔（毫秒）
        
        self.setup_window()
        self.init_processors()
                
        self.setup_modern_ui()
        self.load_pdf_list()
        self.is_processed = False
        self.processing_state = "none"
        # 🔧 progress_indicator 已經在上面設定了，這裡不用重複
        # self.progress_indicator = None  # ← 刪除這行或註解掉
        self.processed_files = set()  # 記錄已處理的檔案
        self.current_selected_files = set()  # 當前選中的檔案
        
        # 🔧 確認這兩個設定有沒有加上：
        self.export_templates = {
            "完整": "all",        # 全部匯出
            "核心": "core",       # 基本資訊 + 所有權部（最重要的）
            "房產": "property",   # 基本資訊 + 標示部 + 所有權部（房仲用）
            "權利": "rights",     # 基本資訊 + 所有權部 + 他項權利部（法務用）
        }
        self.current_template = "完整"  # 預設模板
        # 🔧 綁定關閉事件，自動儲存設定
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    # 🔧 加入這些新方法
    def comprehensive_debug_theme_switch(self):
        """全面調試主題切換的每一步"""
        print("\n" + "="*50, flush=True)
        print("🔍 開始全面調試主題切換", flush=True)
        print("="*50, flush=True)
        
        if hasattr(self, 'file_tree'):
            tree_parent = self.file_tree.master
            
            # 🔍 步驟1：檢查初始狀態
            print("\n📋 步驟1：初始狀態檢查", flush=True)
            print(f"  檔案樹父容器：{tree_parent}", flush=True)
            print(f"  父容器類型：{tree_parent.winfo_class()}", flush=True)
            
            initial_children = []
            for i, child in enumerate(tree_parent.winfo_children()):
                widget_info = f"[{i}] {child.winfo_class()} - {child}"
                initial_children.append(widget_info)
                print(f"    {widget_info}", flush=True)
                
                if isinstance(child, ttk.Scrollbar):
                    try:
                        orient = child.cget('orient')
                        geometry = child.winfo_geometry()
                        manager = child.winfo_manager()
                        visible = child.winfo_viewable()
                        print(f"      ↳ 方向:{orient}, 幾何:{geometry}, 管理器:{manager}, 可見:{visible}", flush=True)
                    except Exception as e:
                        print(f"      ↳ 檢查失敗: {e}", flush=True)
            
            # 🔍 步驟2：檢查是否在其他地方有捲軸
            print("\n📋 步驟2：檢查整個widget樹中的所有捲軸", flush=True)
            self.find_all_scrollbars_in_app(self.master, "", 0)
            
            # 🔍 步驟3：檢查檔案樹的配置
            print("\n📋 步驟3：檔案樹配置檢查", flush=True)
            try:
                x_cmd = self.file_tree.cget('xscrollcommand')
                y_cmd = self.file_tree.cget('yscrollcommand')
                print(f"  X捲軸命令：{x_cmd}", flush=True)
                print(f"  Y捲軸命令：{y_cmd}", flush=True)
                
                # 檢查命令是否指向已銷毀的對象
                if x_cmd:
                    try:
                        test_result = str(x_cmd)
                        print(f"  X命令測試：{test_result}", flush=True)
                    except:
                        print(f"  X命令可能指向已銷毀對象", flush=True)
                        
                if y_cmd:
                    try:
                        test_result = str(y_cmd)
                        print(f"  Y命令測試：{test_result}", flush=True)
                    except:
                        print(f"  Y命令可能指向已銷毀對象", flush=True)
                        
            except Exception as e:
                print(f"  檢查捲軸命令失敗：{e}", flush=True)

    def find_all_scrollbars_in_app(self, widget, path, level):
        """遞迴查找整個應用程式中的所有捲軸"""
        try:
            indent = "  " * level
            if isinstance(widget, ttk.Scrollbar):
                orient = widget.cget('orient') if hasattr(widget, 'cget') else "未知"
                print(f"{indent}🔍 發現捲軸：{path} - {orient} - {widget}", flush=True)
            
            if hasattr(widget, 'winfo_children'):
                for i, child in enumerate(widget.winfo_children()):
                    child_path = f"{path}.{i}" if path else f"root.{i}"
                    self.find_all_scrollbars_in_app(child, child_path, level + 1)
                    
        except Exception as e:
            print(f"  檢查 {path} 時發生錯誤：{e}", flush=True)

    def collect_all_scrollbars(self, widget, scrollbar_list):
        """收集整個widget樹中的所有捲軸"""
        try:
            if isinstance(widget, ttk.Scrollbar):
                scrollbar_list.append(widget)
            
            if hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    self.collect_all_scrollbars(child, scrollbar_list)
                    
        except:
            pass

    def apply_scrollbar_theme_simple(self, v_scrollbar, h_scrollbar):
        """簡化的捲軸主題套用"""
        try:
            style = ttk.Style()
            
            if ModernStyle.current_mode == "dark":
                style.configure("Vertical.TScrollbar", background="#3a3a3a", troughcolor="#2d2d2d")
                style.configure("Horizontal.TScrollbar", background="#3a3a3a", troughcolor="#2d2d2d")
            else:
                style.configure("Vertical.TScrollbar", background="#C0C0C0", troughcolor="#F0F0F0")
                style.configure("Horizontal.TScrollbar", background="#C0C0C0", troughcolor="#F0F0F0")
                
            print("✅ 捲軸主題已套用", flush=True)
        except Exception as e:
            print(f"❌ 套用捲軸主題失敗：{e}", flush=True)

    def nuclear_fix_scrollbars_v2(self):
        """改進的核彈級捲軸清理"""
        try:
            print("💥 執行改進版核彈級清理...", flush=True)
            
            if hasattr(self, 'file_tree'):
                tree_parent = self.file_tree.master
                
                # 🔍 先進行全面調試
                self.comprehensive_debug_theme_switch()
                
                # 🔧 1. 暴力清除：直接遍歷整個widget樹找捲軸
                all_scrollbars_found = []
                self.collect_all_scrollbars(self.master, all_scrollbars_found)
                
                print(f"\n💥 在整個應用程式中發現 {len(all_scrollbars_found)} 個捲軸", flush=True)
                for scrollbar in all_scrollbars_found:
                    try:
                        orient = scrollbar.cget('orient')
                        parent = scrollbar.master
                        print(f"  準備銷毀：{orient}捲軸 - 父容器:{parent}", flush=True)
                        scrollbar.destroy()
                        print(f"  ✅ 已銷毀", flush=True)
                    except Exception as e:
                        print(f"  ❌ 銷毀失敗：{e}", flush=True)
                
                # 🔧 2. 清除檔案樹的捲軸命令
                try:
                    self.file_tree.configure(xscrollcommand="", yscrollcommand="")
                    print("✅ 已清除檔案樹的捲軸命令", flush=True)
                except:
                    pass
                
                # 🔧 3. 強制更新，確保銷毀生效
                self.master.update_idletasks()
                self.master.update()
                
                # 🔧 4. 重新佈局檔案樹
                self.file_tree.pack_forget()
                
                # 🔧 5. 重新創建捲軸
                print("🔧 重新創建捲軸系統...", flush=True)
                v_scrollbar = ttk.Scrollbar(tree_parent, orient='vertical')
                h_scrollbar = ttk.Scrollbar(tree_parent, orient='horizontal')
                
                # 🔧 6. 佈局
                h_scrollbar.pack(side='bottom', fill='x')
                v_scrollbar.pack(side='right', fill='y')  
                self.file_tree.pack(side='left', fill='both', expand=True)
                
                # 🔧 7. 連接
                self.file_tree.configure(
                    yscrollcommand=v_scrollbar.set,
                    xscrollcommand=h_scrollbar.set
                )
                v_scrollbar.configure(command=self.file_tree.yview)
                h_scrollbar.configure(command=self.file_tree.xview)
                
                # 🔧 8. 主題
                self.apply_scrollbar_theme_simple(v_scrollbar, h_scrollbar)
                
                # 🔧 9. 最終檢查
                print("\n📋 最終狀態檢查:", flush=True)
                final_children = []
                for child in tree_parent.winfo_children():
                    final_children.append(child.winfo_class())
                print(f"  父容器子元件：{final_children}", flush=True)
                
                # 再次搜尋整個應用程式
                remaining_scrollbars = []
                self.collect_all_scrollbars(self.master, remaining_scrollbars)
                print(f"  整個應用程式剩餘捲軸數：{len(remaining_scrollbars)}", flush=True)
                
                # 🔧 10. 新增：修正字體大小
                self.fix_treeview_font_ultimate_v2()
                
        except Exception as e:
            print(f"💥 改進版核彈級清理失敗：{e}", flush=True)

    def fix_treeview_font_ultimate_v2(self):
        """終極字體修正 v2 - 更強力的方法"""
        try:
            print("🔧 開始終極字體修正 v2...", flush=True)
            
            if hasattr(self, 'file_tree'):
                # 🔧 方法1：強制重新設定 TTK 主題
                style = ttk.Style()
                
                # 先重置主題
                current_theme = style.theme_use()
                print(f"  當前主題：{current_theme}", flush=True)
                
                # 嘗試切換到其他主題再切回來
                try:
                    available_themes = style.theme_names()
                    print(f"  可用主題：{available_themes}", flush=True)
                    
                    if 'clam' in available_themes:
                        style.theme_use('clam')
                        print("  切換到 clam 主題", flush=True)
                        
                except Exception as e:
                    print(f"  切換主題失敗：{e}", flush=True)
                
                # 🔧 方法2：使用不同的字體大小設定方式
                font_configs = [
                    ("微軟正黑體", 12),
                    ("Microsoft YaHei", 12), 
                    ("Arial", 11),
                    ("Default", 11)
                ]
                
                for font_family, font_size in font_configs:
                    try:
                        style.configure("Treeview",
                                    font=(font_family, font_size),
                                    rowheight=int(font_size * 1.8))
                        
                        style.configure("Treeview.Heading",
                                    font=(font_family, font_size, "bold"))
                        
                        print(f"  ✅ 嘗試字體：{font_family} {font_size}", flush=True)
                        break
                    except Exception as e:
                        print(f"  ❌ 字體 {font_family} 失敗：{e}", flush=True)
                        continue
                
                # 🔧 方法3：強制刷新檔案樹
                try:
                    # 儲存當前展開狀態
                    expanded_items = []
                    for item in self.file_tree.get_children():
                        if self.file_tree.item(item, 'open'):
                            expanded_items.append(self.file_tree.item(item, 'text'))
                    
                    # 儲存選中狀態
                    current_selection = self.file_tree.selection()
                    
                    # 清空並重新載入
                    for item in self.file_tree.get_children():
                        self.file_tree.delete(item)
                    
                    # 重新載入內容
                    if hasattr(self, 'load_folder_contents'):
                        self.load_folder_contents()
                    
                    # 嘗試恢復展開狀態
                    for item in self.file_tree.get_children():
                        item_text = self.file_tree.item(item, 'text')
                        if item_text in expanded_items:
                            self.file_tree.item(item, open=True)
                    
                    print("  ✅ 檔案樹已重新載入", flush=True)
                    
                except Exception as e:
                    print(f"  ❌ 重新載入檔案樹失敗：{e}", flush=True)
                
                # 🔧 方法4：檢查最終字體設定
                try:
                    current_font = style.lookup("Treeview", "font")
                    current_rowheight = style.lookup("Treeview", "rowheight")
                    print(f"  最終字體設定：{current_font}", flush=True)
                    print(f"  最終行高設定：{current_rowheight}", flush=True)
                except Exception as e:
                    print(f"  無法檢查字體設定：{e}", flush=True)
                
                print("✅ 終極字體修正 v2 完成", flush=True)
                
        except Exception as e:
            print(f"❌ 終極字體修正 v2 失敗：{e}", flush=True)

    # 🔧 同時加入這個方法
    def set_status(self, message, progress=None):
        """修正 set_status 方法 - 支援進度參數"""
        try:
            if progress is not None:
                print(f"📢 狀態：{message} ({progress}%)", flush=True)
            else:
                print(f"📢 狀態：{message}", flush=True)
            # 如果你有狀態列元件，可以在這裡更新
            # 例如：self.status_label.config(text=message)
            # 如果有進度條：self.progress_bar.config(value=progress)
        except Exception as e:
            print(f"❌ 更新狀態失敗：{e}", flush=True)

    # def load_settings(self):
    #     """載入程式設定"""
    #     import json
        
    #     # 設定檔路徑（與程式同目錄）
    #     script_dir = os.path.dirname(os.path.abspath(__file__))
    #     settings_file = os.path.join(script_dir, "settings.json")
        
    #     # 預設設定
    #     default_settings = {
    #         "source_directory": script_dir,  # 預設來源目錄為程式所在目錄
    #         "output_directory": os.path.join(script_dir, "output")  # 預設輸出目錄
    #     }
        
    #     try:
    #         if os.path.exists(settings_file):
    #             with open(settings_file, 'r', encoding='utf-8') as f:
    #                 settings = json.load(f)
                
    #             # 載入設定，如果不存在就使用預設值
    #             self.current_folder = settings.get("source_directory", default_settings["source_directory"])
    #             self.output_directory = settings.get("output_directory", default_settings["output_directory"])
                
    #             print(f"📁 已載入設定：來源={self.current_folder}, 輸出={self.output_directory}")
    #         else:
    #             # 第一次執行，使用預設設定
    #             self.current_folder = default_settings["source_directory"]
    #             self.output_directory = default_settings["output_directory"]
    #             self.save_settings()  # 儲存預設設定
                
    #     except Exception as e:
    #         print(f"載入設定失敗，使用預設設定: {e}")
    #         self.current_folder = default_settings["source_directory"]
    #         self.output_directory = default_settings["output_directory"]

    def save_settings(self):
        """儲存程式設定"""
        import json

        # 🔥 修正：使用 BASE_DIR 而不是 __file__，避免在打包後指向 _internal
        script_dir = BASE_DIR
        settings_file = os.path.join(script_dir, "settings.json")
        
        settings = {
            "source_directory": getattr(self, 'current_folder', script_dir),
            "output_directory": getattr(self, 'output_directory', os.path.join(script_dir, "output")),
            "dark_mode": ModernStyle.current_mode == "dark"  # 👈 新增
        }
        
        try:
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"儲存設定失敗: {e}", flush=True)

    def load_settings(self):
        """載入程式設定"""
        import json

        # 🔥 修正：使用 BASE_DIR 而不是 __file__，避免在打包後指向 _internal
        script_dir = BASE_DIR
        settings_file = os.path.join(script_dir, "settings.json")

        default_settings = {
            "source_directory": script_dir,
            "output_directory": os.path.join(script_dir, "output"),
            "dark_mode": False  # 👈 新增
        }

        try:
            if os.path.exists(settings_file):
                with open(settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)

                # 🔥 處理來源目錄：如果儲存的路徑不存在（例如跨電腦同步），使用預設路徑
                saved_source = settings.get("source_directory", default_settings["source_directory"])
                if os.path.exists(saved_source):
                    self.current_folder = saved_source
                else:
                    self.current_folder = default_settings["source_directory"]
                    print(f"[提示] 儲存的來源目錄不存在，使用預設目錄：{self.current_folder}", flush=True)

                # 🔥 處理輸出目錄：如果儲存的路徑不存在（例如跨電腦同步），使用預設路徑
                saved_output = settings.get("output_directory", default_settings["output_directory"])
                if os.path.exists(saved_output) or os.path.dirname(saved_output) == script_dir:
                    # 如果路徑存在，或者是腳本目錄下的子目錄（相對路徑），就使用它
                    self.output_directory = saved_output
                else:
                    # 否則使用預設的相對路徑
                    self.output_directory = default_settings["output_directory"]
                    print(f"[提示] 儲存的輸出目錄不存在，使用預設目錄：{self.output_directory}", flush=True)
                
                # 🔧 載入暗黑模式設定
                dark_mode = settings.get("dark_mode", False)
                if dark_mode:
                    ModernStyle.current_mode = "dark"
                else:
                    ModernStyle.current_mode = "light"
                    
            else:
                self.current_folder = default_settings["source_directory"]
                self.output_directory = default_settings["output_directory"]
                self.save_settings()
                
        except Exception as e:
            print(f"載入設定失敗: {e}", flush=True)
            self.current_folder = default_settings["source_directory"]
            self.output_directory = default_settings["output_directory"]


    def on_closing(self):
        """程式關閉時的處理"""
        self.save_settings()  # 儲存設定
        self.master.destroy()  # 關閉程式

    def create_status_bar(self):
        """創建底部狀態列"""
        status_frame = tk.Frame(self.master, bg="#34495e", height=22, relief="flat")
        status_frame.pack(side="bottom", fill="x")
        status_frame.pack_propagate(False)
        
        # 左側狀態訊息
        self.status_label = tk.Label(status_frame, 
                                    text="🟢 就緒", 
                                    bg="#34495e", 
                                    fg="#ecf0f1",
                                    font=("微軟正黑體", 9),
                                    anchor="w")
        self.status_label.pack(side="left", padx=8, pady=1)
        
        # 右側版本訊息
        version_label = tk.Label(status_frame,
                               text="v1.7 (2026-07-23) | 電子謄本功能測試版",
                               bg="#34495e",
                               fg="#95a5a6",
                               font=("微軟正黑體", 9))
        version_label.pack(side="right", padx=8, pady=1)

    # def update_status(self, message):
    #     """更新狀態列訊息"""
    #     if hasattr(self, 'status_label'):
    #         self.status_label.config(text=f"🟢 {message}")
        
    def set_custom_icon(self):
        """設定自定義圖示 - PDF文件圖標版本"""
        try:
            # 創建48x48的圖示（更大更清楚）
            icon_img = tk.PhotoImage(width=48, height=48)

            # 🔧 繪製PDF文件圖示
            # 背景透明/白色
            for x in range(48):
                for y in range(48):
                    icon_img.put("#FFFFFF", (x, y))

            # 文件外框（紅色 - PDF特色）
            # 左邊框
            for y in range(8, 44):
                for x in range(10, 12):
                    icon_img.put("#DC143C", (x, y))
            # 右邊框
            for y in range(8, 44):
                for x in range(36, 38):
                    icon_img.put("#DC143C", (x, y))
            # 上邊框
            for x in range(10, 38):
                for y in range(8, 10):
                    icon_img.put("#DC143C", (x, y))
            # 下邊框
            for x in range(10, 38):
                for y in range(42, 44):
                    icon_img.put("#DC143C", (x, y))

            # 文件摺角（右上角）
            for x in range(30, 38):
                for y in range(8, 16):
                    if (x - 30) + (y - 8) >= 6:
                        icon_img.put("#FFB6C1", (x, y))  # 淡紅色

            # PDF 文字（紅色）
            # 字母 P
            for y in range(16, 24):
                icon_img.put("#DC143C", (14, y))
            for x in range(14, 18):
                icon_img.put("#DC143C", (x, 16))
            for x in range(14, 18):
                icon_img.put("#DC143C", (x, 20))
            icon_img.put("#DC143C", (18, 17))
            icon_img.put("#DC143C", (18, 18))
            icon_img.put("#DC143C", (18, 19))

            # 字母 D
            for y in range(16, 24):
                icon_img.put("#DC143C", (20, y))
            for x in range(20, 24):
                icon_img.put("#DC143C", (x, 16))
            for x in range(20, 24):
                icon_img.put("#DC143C", (x, 23))
            icon_img.put("#DC143C", (24, 17))
            icon_img.put("#DC143C", (24, 18))
            icon_img.put("#DC143C", (24, 19))
            icon_img.put("#DC143C", (24, 20))
            icon_img.put("#DC143C", (24, 21))
            icon_img.put("#DC143C", (24, 22))

            # 字母 F
            for y in range(16, 24):
                icon_img.put("#DC143C", (26, y))
            for x in range(26, 30):
                icon_img.put("#DC143C", (x, 16))
            for x in range(26, 29):
                icon_img.put("#DC143C", (x, 20))

            # 文件線條（表示內容）
            for x in range(14, 34):
                icon_img.put("#B0B0B0", (x, 28))
            for x in range(14, 34):
                icon_img.put("#B0B0B0", (x, 31))
            for x in range(14, 28):
                icon_img.put("#B0B0B0", (x, 34))
            for x in range(14, 30):
                icon_img.put("#B0B0B0", (x, 37))

            # 🔧 設定視窗圖示
            self.master.iconphoto(True, icon_img)

            # 🔧 Windows特定：設定工作列圖示
            if os.name == 'nt':
                try:
                    # 設定App User Model ID，使其在工作列中獨立顯示
                    import ctypes
                    myappid = 'TaiwanRealEstate.TranscriptProcessor.GUI.v1'
                    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
                except:
                    pass

            print("✅ PDF圖示已設定", flush=True)

        except Exception as e:
            print(f"❌ 設定圖示失敗: {e}", flush=True)

    #====================新增】3個新方法======================
    def show_export_selection_dialog(self, title="處理完成", message="", processed_count=0, pdf_files_count=0, transcripts_count=0):
        """顯示匯出選擇對話框 - 使用現有匯出方法"""
        dialog = tk.Toplevel(self.master)
        dialog.title("選擇匯出格式")
        dialog.geometry("500x400")
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.resizable(False, False)

        # 居中顯示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)  # 🔧 修正：使用實際寬度的一半
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)  # 🔧 修正：使用實際高度的一半
        dialog.geometry(f"500x400+{x}+{y}")

        # 主容器
        container = tk.Frame(dialog, bg=ModernStyle.BG_LIGHT)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        # 標題
        title_label = tk.Label(container,
                            text="🎉 " + title,
                            font=("微軟正黑體", 16, "bold"),
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.SUCCESS)
        title_label.pack(pady=(0, 15))

        # 處理結果訊息
        if message:
            msg_label = tk.Label(container,
                                text=message,
                                font=("微軟正黑體", 12),
                                bg=ModernStyle.BG_LIGHT,
                                fg=ModernStyle.TEXT_PRIMARY)
            msg_label.pack(pady=(0, 10))

        # 處理統計（根據參數顯示不同的訊息）
        if pdf_files_count > 0 and transcripts_count > 0:
            # 顯示 PDF 檔案數量和謄本數量
            stats_text = f"已處理 {pdf_files_count} 個PDF檔案，成功提取 {transcripts_count} 份謄本"
        elif processed_count > 0:
            # 向後相容：使用舊的參數
            stats_text = f"已成功處理 {processed_count} 份謄本"
        else:
            stats_text = "處理完成"

        stats_label = tk.Label(container,
                            text=stats_text,
                            font=("微軟正黑體", 12, "bold"),
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.INFO)
        stats_label.pack(pady=(0, 20))
        
        # 匯出選項標題
        export_title = tk.Label(container,
                            text="📤 選擇要匯出的格式：",
                            font=("微軟正黑體", 14, "bold"),
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.TEXT_PRIMARY)
        export_title.pack(pady=(0, 15))
        
        # 匯出按鈕區域
        buttons_frame = tk.Frame(container, bg=ModernStyle.BG_LIGHT)
        buttons_frame.pack(expand=True, fill="both", pady=(0, 20))
        
        def export_and_close(format_type):
            """匯出並關閉對話框"""
            try:
                dialog.destroy()  # 先關閉對話框
                
                # 使用現有的匯出方法
                if format_type == "txt":
                    self.export_format("txt", silent_mode=False)
                elif format_type == "excel":
                    self.export_format("excel", silent_mode=False)
                elif format_type == "json":
                    self.export_format("json", silent_mode=False)
                elif format_type == "all":
                    self.export_all()
                
            except Exception as e:
                print(f"❌ 匯出失敗：{e}", flush=True)
                self.show_custom_messagebox("匯出錯誤", f"匯出失敗：{e}")
        
        # 匯出按鈕配置
        export_options = [
            ("📄 匯出 TXT", "txt", "#6c757d"),
            ("📈 匯出 Excel", "excel", "#28a745"),
            ("🔗 匯出 JSON", "json", "#ffc107"),
            ("📦 匯出全部", "all", "#17a2b8")
        ]
        
        # 創建匯出按鈕（2x2 網格）
        for i, (text, format_type, color) in enumerate(export_options):
            row = i // 2
            col = i % 2
            
            btn_frame = tk.Frame(buttons_frame, bg=ModernStyle.BG_LIGHT)
            btn_frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
            
            btn = tk.Button(btn_frame,
                        text=text,
                        command=lambda fmt=format_type: export_and_close(fmt),
                        font=("微軟正黑體", 12, "bold"),
                        bg=color,
                        fg="white" if color != "#ffc107" else "black",
                        relief="flat",
                        borderwidth=0,
                        cursor="hand2",
                        height=3,
                        width=15)
            btn.pack(fill="both", expand=True)
        
        # 配置網格權重
        buttons_frame.grid_columnconfigure(0, weight=1)
        buttons_frame.grid_columnconfigure(1, weight=1)
        buttons_frame.grid_rowconfigure(0, weight=1)
        buttons_frame.grid_rowconfigure(1, weight=1)
        
        # 底部跳過按鈕
        bottom_frame = tk.Frame(container, bg=ModernStyle.BG_LIGHT)
        bottom_frame.pack(side="bottom", fill="x")
        
        skip_btn = tk.Button(bottom_frame,
                            text="⏭️ 稍後再匯出",
                            command=dialog.destroy,
                            font=("微軟正黑體", 11),
                            bg="#95a5a6",
                            fg="white",
                            relief="flat",
                            cursor="hand2")
        skip_btn.pack(pady=10)

    # def show_auto_close_success_message(self, title, processed_count):
    #     """顯示自動關閉的成功訊息 - 加入動態執行時間"""
    #     msg_window = tk.Toplevel(self.master)
    #     msg_window.title(title)
    #     msg_window.geometry("450x250")  # 稍微加高以容納執行時間
    #     msg_window.transient(self.master)
    #     msg_window.grab_set()
    #     msg_window.resizable(False, False)
        
    #     # 居中顯示
    #     msg_window.update_idletasks()
    #     x = (msg_window.winfo_screenwidth() // 2) - (225)
    #     y = (msg_window.winfo_screenheight() // 2) - (125)
    #     msg_window.geometry(f"450x250+{x}+{y}")
        
    #     container = tk.Frame(msg_window, bg=ModernStyle.BG_LIGHT)
    #     container.pack(fill="both", expand=True, padx=20, pady=20)
        
    #     # 成功圖示和訊息
    #     success_label = tk.Label(container,
    #                         text="🎉 " + title,
    #                         font=("微軟正黑體", 16, "bold"),
    #                         bg=ModernStyle.BG_LIGHT,
    #                         fg=ModernStyle.SUCCESS)
    #     success_label.pack(pady=(20, 10))
        
    #     detail_label = tk.Label(container,
    #                         text=f"已成功處理 {processed_count} 個檔案",
    #                         font=("微軟正黑體", 12),
    #                         bg=ModernStyle.BG_LIGHT,
    #                         fg=ModernStyle.TEXT_PRIMARY)
    #     detail_label.pack(pady=(0, 15))
        
    #     # 🔧 新增：執行時間顯示
    #     runtime_label = tk.Label(container,
    #                         text="⏱️ 執行時間：計算中...",
    #                         font=("微軟正黑體", 11),
    #                         bg=ModernStyle.BG_LIGHT,
    #                         fg=ModernStyle.INFO)
    #     runtime_label.pack(pady=(0, 10))
        
    #     # 倒數計時
    #     countdown_label = tk.Label(container,
    #                             text="4 秒後自動關閉...",
    #                             font=("微軟正黑體", 10),
    #                             bg=ModernStyle.BG_LIGHT,
    #                             fg=ModernStyle.TEXT_SECONDARY)
    #     countdown_label.pack()
        
    #     # 🔧 動態更新執行時間和倒數
    #     import time
    #     start_time = time.time()
        
    #     def update_display(remaining):
    #         if remaining > 0:
    #             # 更新執行時間
    #             current_time = time.time()
    #             elapsed = current_time - start_time
                
    #             # 格式化執行時間
    #             if elapsed >= 3600:  # 超過1小時
    #                 hours = int(elapsed // 3600)
    #                 minutes = int((elapsed % 3600) // 60)
    #                 seconds = int(elapsed % 60)
    #                 time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    #             elif elapsed >= 60:  # 超過1分鐘
    #                 minutes = int(elapsed // 60)
    #                 seconds = int(elapsed % 60)
    #                 time_str = f"{minutes:02d}:{seconds:02d}"
    #             else:  # 小於1分鐘
    #                 time_str = f"{elapsed:.1f} 秒"
                
    #             runtime_label.config(text=f"⏱️ 執行時間：{time_str}")
    #             countdown_label.config(text=f"{remaining:.2f} 秒後自動關閉...")
                
    #             msg_window.after(100, lambda: update_display(remaining - 0.1))
    #         else:
    #             msg_window.destroy()
        
    #     # 🔧 延遲4秒關閉（批次處理需要更多時間）
    #     close_delay = 4.0 if "批次處理" in title else 3.0
    #     update_display(close_delay)

    def on_batch_process_complete(self, pdf_files_count, transcripts_count):
        """批次處理完成處理 - 直接顯示匯出選擇"""
        # 🔧 移除中間成功訊息，直接顯示匯出選擇對話框
        self.show_export_selection_dialog(
            title="批次處理完成",
            message="批次處理已完成！",
            pdf_files_count=pdf_files_count,
            transcripts_count=transcripts_count
        )

    # ==================== 修改2: 添加更新標籤的簡單函數 ====================
    def update_tab_title(self, is_processed=False):
        """更新結構化資料標籤的標題"""
        if is_processed:
            # 🔥 用火焰表示已處理（橘紅色感覺，非常醒目）
            new_title = "🔥 已結構化處理"
            self.is_processed = True
        else:
            # 未處理：顯示原始標題
            new_title = "📊 結構化資料"
            self.is_processed = False
        
        # 🔧 更新按鈕文字而不是notebook標籤
        if hasattr(self, 'tab_buttons') and 'struct' in self.tab_buttons:
            self.tab_buttons['struct'].config(text=new_title)
        
    def setup_window(self):
        """設置主視窗"""
        self.master.title("📄 電子謄本PDF處理工具")
        # 🔧 使用螢幕適應性工具計算最佳視窗大小
        window_config = ScreenUtils.calculate_optimal_window_size(
            target_width=1400, 
            target_height=800, 
            margin=50
        )

        # 設置視窗大小和位置
        geometry_string = f"{window_config['width']}x{window_config['height']}+{window_config['x']}+{window_config['y']}"
        self.master.geometry(geometry_string)

        # 🔧 設置最小視窗大小以防過度縮小
        self.master.minsize(1000, 600)
        self.master.configure(bg=ModernStyle.BG_PRIMARY)
        
        try:
            # 備用方案：使用更亮的背景色
            import ctypes
            import tkinter as tk
            from PIL import Image, ImageDraw, ImageTk
            
            self.master.iconbitmap("")
            
            # 使用更亮的背景色
            img = Image.new('RGBA', (32, 32), (100, 150, 200, 255))  # 更亮的藍色
            draw = ImageDraw.Draw(img)
            
            # 超寬版本
            draw.polygon([(16, 5), (5, 15), (27, 15)], fill='white')  # 更寬的屋頂
            draw.rectangle([7, 15, 25, 27], fill='white')              # 更寬的房身
            draw.rectangle([14, 20, 18, 27], fill='#8B4513')           # 門
            draw.rectangle([8, 17, 11, 19], fill='#87CEEB')            # 左窗
            draw.rectangle([21, 17, 24, 19], fill='#87CEEB')           # 右窗
            
            # 加個邊框讓圖示更明顯
            draw.rectangle([0, 0, 31, 31], outline='white', width=1)
            
            icon_photo = ImageTk.PhotoImage(img)
            self.master.iconphoto(True, icon_photo)
            
            myappid = 'RealEstateTool.Application.1.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            
        except Exception as e:
            print(f"備用圖示設置失敗: {e}", flush=True)
            pass
        
        # 配置ttk樣式
        ModernStyle.configure_ttk_style()


    def init_processors(self):
        """初始化處理器"""
        global ProcessorClass, ExporterClass
        
        print("\n🔧 處理器初始化...", flush=True)
        
        if ProcessorClass and ExporterClass:
            try:
                print(f"  🎯 使用原程式的處理器類別", flush=True)
                self.processor = ProcessorClass()
                self.exporter = ExporterClass()
                print("  ✅ 原有處理器初始化成功", flush=True)
                self.using_original = True
                
                if hasattr(self.processor, 'process_multiple_pdfs'):
                    print("    ✅ process_multiple_pdfs 方法存在", flush=True)
                else:
                    print("    ❌ process_multiple_pdfs 方法不存在", flush=True)
                    raise Exception("缺少必要的處理方法")
                    
            except Exception as e:
                print(f"  ❌ 原有處理器初始化失敗: {e}", flush=True)
                print(f"  🔄 切換到增強版處理器...", flush=True)
                self.init_enhanced_processors()
        else:
            print("  ⚠️ 無法找到原有處理模組，使用增強版本", flush=True)
            self.init_enhanced_processors()
        
        print(f"  📊 最終使用: {'原程式處理器' if self.using_original else '增強版處理器'}", flush=True)
    
    def init_enhanced_processors(self):
        """初始化增強處理器"""
        print("  🔄 初始化增強版處理器...", flush=True)
        self.processor = EnhancedTranscriptProcessor()
        self.exporter = EnhancedTranscriptExporter(gui_instance=self)  # 🔥 修正：加上 self.
        self.using_original = False
        print("  ✅ 增強處理器初始化完成", flush=True)
    
    def setup_modern_ui(self):
        """建立現代化使用者介面"""

        # 頂部標題欄
        self.create_header()

        # 主要內容區域
        main_container = tk.Frame(self.master, bg=ModernStyle.BG_LIGHT)
        main_container.pack(fill="both", expand=True, padx=1, pady=(1, 0))

        # 🔧 使用 PanedWindow 來創建可調整大小的分隔線
        # 第一層 PanedWindow：分隔左側+中間 vs 右側
        self.main_paned = tk.PanedWindow(main_container, orient=tk.HORIZONTAL,
                                          sashwidth=5, sashrelief=tk.RAISED,
                                          bg=ModernStyle.BG_SECONDARY)
        self.main_paned.pack(fill="both", expand=True)

        # 第二層 PanedWindow：分隔左側 vs 中間
        self.left_center_paned = tk.PanedWindow(self.main_paned, orient=tk.HORIZONTAL,
                                                 sashwidth=5, sashrelief=tk.RAISED,
                                                 bg=ModernStyle.BG_SECONDARY)
        self.main_paned.add(self.left_center_paned, stretch="always")

        # 左側面板 - 檔案管理（添加到 left_center_paned）
        self.create_left_panel(self.left_center_paned)

        # 中間面板 - 預覽區域（添加到 left_center_paned）
        self.create_center_panel(self.left_center_paned)

        # 右側面板 - 操作控制（添加到 main_paned）
        self.create_right_panel(self.main_paned)

        # 底部狀態列
        self.create_footer()

        # 🔧 視窗排版完成後，把左側與中間調成等寬
        self.master.after(400, self._balance_left_center)

    def _balance_left_center(self):
        """把左側檔案管理與中間預覽欄調成一樣寬（分隔線置中）。"""
        try:
            self.left_center_paned.update_idletasks()
            total = self.left_center_paned.winfo_width()
            if total > 50:
                self.left_center_paned.sash_place(0, total // 2, 0)
        except Exception as e:
            print(f"左右等寬調整失敗: {e}", flush=True)


    def create_header(self):
        """創建頂部標題欄 - 調細版本"""
        header = tk.Frame(self.master, bg=ModernStyle.PRIMARY, height=50)  # 從80改為50
        header.pack(fill="x")
        header.pack_propagate(False)
        
        # 主標題
        title_label = tk.Label(header, 
                            text="🏛️ 不動產登記謄本智慧處理工具",
                            font=("Segoe UI", 18, "bold"),
                            bg=ModernStyle.PRIMARY,
                            fg=ModernStyle.TEXT_LIGHT)
        title_label.pack(side="left", padx=30, pady=12)  # 從pady=20改為12
        
        # 版本與狀態資訊
        status_text = "✅電子謄本功能測試版" if self.using_original else "⚠️ 增強功能版本"
        version_label = tk.Label(header,
                            text=f"v1.7 (2026-07-23) | {status_text}",
                            font=ModernStyle.FONT_MAIN,
                            bg=ModernStyle.PRIMARY,
                            fg=ModernStyle.TEXT_LIGHT)
        version_label.pack(side="right", padx=30, pady=12)  # 從pady=20改為12

    def apply_theme(self):
        """套用主題到整個界面 - 修正變數作用域"""
        print("🔍 調試：apply_theme 被呼叫", flush=True)

        # 🔧 確保 colors 變數在整個方法中可用
        if ModernStyle.current_mode == "dark":
            colors = {
                "BG_PRIMARY": "#1e1e1e",
                "BG_LIGHT": "#2d2d2d",
                "BG_SECONDARY": "#3a3a3a",
                "TEXT_PRIMARY": "#ffffff",
                "TEXT_SECONDARY": "#cccccc",
                "HEADER_BG": "#1a365d"
            }
        else:
            colors = {
                "BG_PRIMARY": "#F8F9FA",
                "BG_LIGHT": "#FFFFFF",
                "BG_SECONDARY": "#E9ECEF",
                "TEXT_PRIMARY": "#212529",
                "TEXT_SECONDARY": "#6C757D",
                "HEADER_BG": "#2E86AB"
            }

        print(f"🔍 調試：使用顏色方案: {colors}", flush=True)

        try:
            # 1. 更新主視窗
            self.master.configure(bg=colors["BG_PRIMARY"])
            print("✅ 主視窗背景已更新", flush=True)

            # 2. 智能更新所有元件
            self.smart_update_widgets(colors)

            # 3. 特別處理按鈕
            self.update_buttons_theme(colors)

            # 4. 🔧 修正：強制更新TTK樣式 (確保 colors 變數可用)
            self.force_update_ttk_styles(colors)

            # 5. 特別強制更新檔案樹標題
            self.force_treeview_heading_update()

            # 🔧 6. 核彈級捲軸修復（替換原來的方法）
            self.nuclear_fix_scrollbars_v2()

            # 🔧 7. 更新 PanedWindow 的分隔線顏色
            self.update_panedwindow_theme(colors)

            # 8. 強制刷新
            self.master.update_idletasks()
            self.master.update()

            print("✅ 主題套用完成", flush=True)

        except Exception as e:
            print(f"❌ 套用主題時發生錯誤: {e}", flush=True)

    def update_panedwindow_theme(self, colors):
        """更新 PanedWindow 分隔線的主題顏色"""
        try:
            # 更新主 PanedWindow（左側+中間 vs 右側）
            if hasattr(self, 'main_paned'):
                self.main_paned.configure(bg=colors["BG_SECONDARY"])
                print("✅ 主 PanedWindow 顏色已更新", flush=True)

            # 更新左側-中間 PanedWindow
            if hasattr(self, 'left_center_paned'):
                self.left_center_paned.configure(bg=colors["BG_SECONDARY"])
                print("✅ 左側-中間 PanedWindow 顏色已更新", flush=True)

        except Exception as e:
            print(f"❌ 更新 PanedWindow 主題失敗: {e}", flush=True)

    def force_treeview_heading_update(self):
        """強制更新檔案樹標題樣式 - 恢復成功版本"""
        try:
            if hasattr(self, 'file_tree'):
                style = ttk.Style()
                
                # 🔧 保持成功的設定：強制使用clam主題
                try:
                    style.theme_use('clam')
                except:
                    pass
                
                if ModernStyle.current_mode == "dark":
                    # 🔧 保持成功的暗黑模式設定
                    style.configure("Treeview",
                                background="#2d2d2d",
                                foreground="#ffffff",
                                selectbackground="#0078d4",
                                selectforeground="white",
                                fieldbackground="#2d2d2d")
                    
                    style.configure("Treeview.Heading",
                                background="#505050",      # 成功的深灰背景
                                foreground="#ffffff",      # 成功的白色文字
                                borderwidth=2,
                                relief="raised",
                                font=("微軟正黑體", 10, "bold"))
                    
                    print("🔍 暗黑模式標題：背景=#505050, 文字=#ffffff", flush=True)
                    
                else:
                    # 🔧 保持成功的淺色模式設定
                    style.configure("Treeview",
                                background="#ffffff",
                                foreground="#212529",
                                selectbackground="#0078d4",
                                selectforeground="white",
                                fieldbackground="#ffffff")
                    
                    style.configure("Treeview.Heading",
                                background="#D0D0D0",      # 成功的深灰背景
                                foreground="#000000",      # 成功的黑色文字
                                borderwidth=2,
                                relief="raised",
                                font=("微軟正黑體", 10, "bold"))
                    
                    print("🔍 淺色模式標題：背景=#D0D0D0, 文字=#000000", flush=True)
                
                # 🔧 保持成功的懸停效果
                style.map("Treeview.Heading",
                        background=[('active', '#606060' if ModernStyle.current_mode == "dark" else '#C0C0C0')])
                
                # 🔧 保持成功的檔案樹設定
                self.file_tree.configure(style="Treeview")
                
                # 🔧 保持成功的多重刷新
                self.file_tree.update_idletasks()
                self.file_tree.update()
                
                if self.file_tree.master:
                    self.file_tree.master.update_idletasks()
                    self.file_tree.master.update()
                
                print("✅ 檔案樹標題更新完成", flush=True)
                
        except Exception as e:
            print(f"❌ 更新檔案樹標題失敗: {e}", flush=True)


    def force_update_ttk_styles(self, colors):
        """強制更新所有TTK樣式"""
        try:
            style = ttk.Style()
            
            # 🔧 直接修改預設 Treeview 樣式
            style.configure("Treeview",
                        background=colors["BG_LIGHT"],
                        foreground=colors["TEXT_PRIMARY"],
                        selectbackground="#0078d4",
                        selectforeground="white",
                        fieldbackground=colors["BG_LIGHT"])
            
            # 🔧 重點：確保標題列正確
            style.configure("Treeview.Heading",
                        background=colors["BG_SECONDARY"],
                        foreground=colors["TEXT_PRIMARY"],
                        font=("微軟正黑體", 10, "bold"))
            
            # 🔧 強制所有 Treeview 重新載入樣式
            for widget in self.get_all_widgets():
                if isinstance(widget, ttk.Treeview):
                    # 暫時切換樣式強制刷新
                    widget.configure(style="Treeview")
                    widget.update_idletasks()
            
            print(f"✅ 強制更新TTK樣式完成", flush=True)
            
        except Exception as e:
            print(f"❌ 強制更新TTK樣式失敗: {e}", flush=True)

    def get_all_widgets(self):
        """取得所有widget的列表"""
        widgets = []
        
        def collect_widgets(parent):
            widgets.append(parent)
            for child in parent.winfo_children():
                collect_widgets(child)
        
        collect_widgets(self.master)
        return widgets

    def smart_update_widgets(self, colors):
        """智能更新widget - 修正左側面板版本"""
        print("🔍 開始智能更新widget...", flush=True)
        
        def update_widget_smart(widget, path="root"):
            try:
                widget_class = widget.winfo_class()
                
                if widget_class == "Frame":
                    current_bg = widget.cget('bg')
                    
                    # 🔧 直接指定：第一個Frame就是標題
                    if path == "root.0":
                        new_bg = colors["HEADER_BG"]
                        widget.configure(bg=new_bg)
                        print(f"  🔵 標題Frame更新: {path} {current_bg} -> {new_bg}", flush=True)
                    else:
                        widget.configure(bg=colors["BG_PRIMARY"])
                        print(f"  Frame更新: {path} {current_bg} -> {colors['BG_PRIMARY']}", flush=True)
                    
                elif widget_class == "Label":
                    # 🔧 直接指定：在標題Frame下的標籤
                    if path.startswith("root.0."):
                        # 標題下的標籤用白字
                        parent_bg = colors["HEADER_BG"]
                        widget.configure(bg=parent_bg, fg="white")
                        print(f"  🔵 藍色標籤更新: {path}", flush=True)
                    else:
                        # 一般標籤
                        widget.configure(bg=colors["BG_PRIMARY"], fg=colors["TEXT_PRIMARY"])
                    
                elif widget_class in ["Text", "Entry"]:
                    widget.configure(bg=colors["BG_LIGHT"], 
                                fg=colors["TEXT_PRIMARY"],
                                insertbackground=colors["TEXT_PRIMARY"])
                    
                elif widget_class == "Canvas":
                    widget.configure(bg=colors["BG_LIGHT"])
                    print(f"  Canvas更新: {path}", flush=True)
                    
                # 🔧 新增：處理 Listbox
                elif widget_class == "Listbox":
                    widget.configure(bg=colors["BG_LIGHT"], 
                                fg=colors["TEXT_PRIMARY"],
                                selectbackground=ModernStyle.PRIMARY,
                                selectforeground="white")
                    print(f"  Listbox更新: {path}", flush=True)
                    
                # 🔧 新增：強制處理所有 TTK 元件
                elif widget_class in ["Treeview", "TScrollbar"]:
                    self.update_ttk_widget(widget, colors)
                    print(f"  TTK元件更新: {path} ({widget_class})", flush=True)
                
                # 遞迴處理子元件
                for i, child in enumerate(widget.winfo_children()):
                    update_widget_smart(child, f"{path}.{i}")
                    
            except Exception as e:
                print(f"  更新widget失敗: {path} - {e}", flush=True)
        
        update_widget_smart(self.master)

    def update_buttons_theme(self, colors):
        """更新所有按鈕主題 - 修正跳過問題"""
        print("🔍 開始更新按鈕主題...", flush=True)
        
        # 🔧 明確定義需要切換的按鈕顏色
        if ModernStyle.current_mode == "dark":
            # 切換到暗黑模式的對應
            button_color_map = {
                "#E9ECEF": "#404040",    # 一般灰色按鈕 -> 暗色
                "#FFFFFF": "#2d2d2d",    # 白色按鈕 -> 深暗色
                "#F8F9FA": "#404040",    # 淺灰按鈕 -> 暗色
            }
            default_fg = "#ffffff"
        else:
            # 切換到淺色模式的對應
            button_color_map = {
                "#404040": "#E9ECEF",    # 暗色按鈕 -> 灰色
                "#2d2d2d": "#FFFFFF",    # 深暗色 -> 白色
                "#3a3a3a": "#E9ECEF",    # 中暗色 -> 灰色
                "#1e1e1e": "#E9ECEF",    # 極暗色 -> 灰色
            }
            default_fg = "#212529"
        
        # 🔧 功能性顏色按鈕 - 保持不變
        keep_original = [
            '#2E86AB', '#DC3545', '#28A745', '#FFC107', '#17A2B8', 
            '#95a5a6', '#3498db', '#27ae60', '#2ecc71', '#A23B72'
        ]
        
        def find_and_update_buttons(widget, path="root"):
            try:
                if isinstance(widget, tk.Button):
                    current_bg = widget.cget('bg')
                    button_text = widget.cget('text')
                    
                    print(f"    檢查按鈕: '{button_text}' BG:{current_bg}", flush=True)
                    
                    # 🔧 1. 檢查是否為功能性顏色（完全匹配）
                    if current_bg in keep_original:
                        print(f"      🎨 保持功能色: {button_text}", flush=True)
                        
                    # 🔧 2. 檢查對應表（完全匹配）
                    elif current_bg in button_color_map:
                        new_bg = button_color_map[current_bg]
                        widget.configure(bg=new_bg, fg=default_fg)
                        print(f"      ✅ 按鈕更新: {button_text} {current_bg} -> {new_bg}", flush=True)
                        
                    # 🔧 3. 特殊處理：淺色模式下所有暗色按鈕
                    elif ModernStyle.current_mode == "light" and current_bg.startswith('#'):
                        # 轉換為RGB值判斷亮度
                        try:
                            hex_color = current_bg.lstrip('#')
                            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
                            brightness = (r * 299 + g * 587 + b * 114) / 1000
                            
                            if brightness < 128:  # 暗色
                                widget.configure(bg="#E9ECEF", fg="#212529")
                                print(f"      🔄 暗色->淺色: {button_text} {current_bg} -> #E9ECEF", flush=True)
                            else:
                                print(f"      ⚪ 保持淺色: {button_text}", flush=True)
                        except:
                            print(f"      ⚪ 顏色解析失敗: {button_text}", flush=True)
                            
                    # 🔧 4. 特殊處理：暗黑模式下所有淺色按鈕  
                    elif ModernStyle.current_mode == "dark" and current_bg.startswith('#'):
                        try:
                            hex_color = current_bg.lstrip('#')
                            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
                            brightness = (r * 299 + g * 587 + b * 114) / 1000
                            
                            if brightness > 128:  # 淺色
                                widget.configure(bg="#404040", fg="#ffffff")
                                print(f"      🔄 淺色->暗色: {button_text} {current_bg} -> #404040", flush=True)
                            else:
                                print(f"      ⚪ 保持暗色: {button_text}", flush=True)
                        except:
                            print(f"      ⚪ 顏色解析失敗: {button_text}", flush=True)
                            
                    else:
                        print(f"      ⚪ 未處理: {button_text} (顏色:{current_bg})", flush=True)
                
                # 遞迴處理子元件
                for i, child in enumerate(widget.winfo_children()):
                    find_and_update_buttons(child, f"{path}.{i}")
                    
            except Exception as e:
                print(f"      ❌ 按鈕更新失敗: {path} - {e}", flush=True)
        
        find_and_update_buttons(self.master)

    def is_dark_color(self, color):
        """判斷是否為暗色"""
        dark_indicators = ['#1', '#2', '#3', '#4', '#5', 'black', 'dark']
        return any(indicator in color.lower() for indicator in dark_indicators)

    def is_light_color(self, color):
        """判斷是否為淺色"""
        light_indicators = ['#F', '#E', '#D', '#C', '#B', 'white', 'light']
        return any(indicator in color.lower() for indicator in light_indicators) or color == '#FFFFFF'

    def update_ttk_widget(self, widget, colors):
        """更新TTK元件主題 - 修正標題顏色對比"""
        try:
            style = ttk.Style()
            
            if isinstance(widget, ttk.Treeview):
                # 設定主體樣式
                style_name = "Custom.Treeview"
                
                style.configure(style_name,
                            background=colors["BG_LIGHT"],           # 檔案列表背景
                            foreground=colors["TEXT_PRIMARY"],       # 檔案列表文字
                            selectbackground=ModernStyle.PRIMARY,    # 選中背景
                            selectforeground="white",                # 選中文字
                            fieldbackground=colors["BG_LIGHT"],      # 欄位背景
                            borderwidth=0,
                            relief="flat")
                
                # 🔧 重點修正：標題列要有明顯對比
                print(f"🔍 設定標題列樣式，當前模式: {ModernStyle.current_mode}", flush=True)
                
                if ModernStyle.current_mode == "dark":
                    # 暗黑模式：深灰背景 + 白色文字
                    heading_bg = "#404040"    # 較深的灰色背景
                    heading_fg = "#ffffff"    # 白色文字
                    print(f"🔍 暗黑模式標題：背景={heading_bg}, 文字={heading_fg}", flush=True)
                else:
                    # 淺色模式：淺灰背景 + 深色文字
                    heading_bg = "#E9ECEF"    # 淺灰背景
                    heading_fg = "#212529"    # 深色文字
                    print(f"🔍 淺色模式標題：背景={heading_bg}, 文字={heading_fg}", flush=True)
                
                style.configure(f"{style_name}.Heading",
                            background=heading_bg,
                            foreground=heading_fg,
                            borderwidth=1,
                            relief="raised",  # 🔧 加一點立體感
                            font=("微軟正黑體", 10, "bold"))
                
                # 🔧 應用自定義樣式到 widget
                widget.configure(style=style_name)
                
                # 強制更新父容器
                if widget.master:
                    try:
                        widget.master.configure(bg=colors["BG_PRIMARY"])
                    except:
                        pass
                
                print(f"✅ Treeview主題已更新: 主體={colors['BG_LIGHT']}, 標題={heading_bg}", flush=True)
                
            elif isinstance(widget, ttk.Scrollbar):
                # 更新捲軸樣式
                style_name = "Custom.Vertical.TScrollbar"
                
                style.configure(style_name,
                            background=colors["BG_SECONDARY"],
                            troughcolor=colors["BG_PRIMARY"],
                            borderwidth=0,
                            arrowcolor=colors["TEXT_PRIMARY"])
                
                widget.configure(style=style_name)
                print(f"✅ Scrollbar主題已更新", flush=True)
                
            # ... 其他元件保持不變 ...
                
        except Exception as e:
            print(f"❌ 更新TTK元件失敗: {e}", flush=True)

    # def update_widget_theme(self, widget, colors):
    #     """遞迴更新widget主題"""
    #     try:
    #         # 更新當前widget
    #         if isinstance(widget, (tk.Frame, tk.LabelFrame)):
    #             widget.configure(bg=colors["BG_PRIMARY"])
    #         elif isinstance(widget, tk.Label):
    #             widget.configure(bg=colors["BG_PRIMARY"], fg=colors["TEXT_PRIMARY"])
    #         elif isinstance(widget, tk.Text):
    #             widget.configure(bg=colors["BG_LIGHT"], fg=colors["TEXT_PRIMARY"],
    #                         insertbackground=colors["TEXT_PRIMARY"])
    #         elif isinstance(widget, tk.Entry):
    #             widget.configure(bg=colors["BG_LIGHT"], fg=colors["TEXT_PRIMARY"],
    #                         insertbackground=colors["TEXT_PRIMARY"])
            
    #         # 遞迴處理子元件
    #         for child in widget.winfo_children():
    #             self.update_widget_theme(child, colors)
                
    #     except Exception as e:
    #         print(f"更新widget主題時發生錯誤: {e}")

    def show_settings_dialog(self):
        """顯示設定對話框 - 修正按鈕版本"""
        dialog = tk.Toplevel(self.master)
        dialog.title("系統設定")
        dialog.geometry("500x500")  # 👈 再增加一點高度
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # 居中顯示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (500 // 2)
        dialog.geometry(f"500x500+{x}+{y}")
        
        # 🔧 整體容器 - 使用 pack 而不是 grid
        container = tk.Frame(dialog)
        container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 標題
        title_label = tk.Label(container, 
                            text="⚙️ 系統設定", 
                            font=("微軟正黑體", 16, "bold"),
                            fg="#2c3e50")
        title_label.pack(pady=(0, 20))
        
        # 🔧 內容區域
        content_frame = tk.Frame(container)
        content_frame.pack(fill="both", expand=True)
        
        # 來源目錄設定
        source_frame = tk.LabelFrame(content_frame, text="📂 來源目錄（PDF檔案所在位置）", 
                                    font=("微軟正黑體", 12, "bold"),
                                    fg="#34495e", padx=10, pady=10)
        source_frame.pack(fill="x", pady=(0, 15))
        
        current_source_var = tk.StringVar(value=self.current_folder)
        source_entry = tk.Entry(source_frame, textvariable=current_source_var,
                            font=("微軟正黑體", 10),
                            state="readonly", bg="#ecf0f1")
        source_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        source_browse_btn = tk.Button(source_frame, text="📁 瀏覽",
                                    command=lambda: self.browse_folder(current_source_var, "選擇PDF檔案目錄"),
                                    font=("微軟正黑體", 10),
                                    bg="#3498db", fg="white")
        source_browse_btn.pack(side="right")
        
        # 輸出目錄設定  
        output_frame = tk.LabelFrame(content_frame, text="📁 輸出目錄（處理結果存放位置）", 
                                    font=("微軟正黑體", 12, "bold"),
                                    fg="#34495e", padx=10, pady=10)
        output_frame.pack(fill="x", pady=(0, 15))
        
        if not hasattr(self, 'output_directory'):
            # 🔥 修正：使用 BASE_DIR 而不是 __file__
            self.output_directory = os.path.join(BASE_DIR, "output")
        
        current_output_var = tk.StringVar(value=self.output_directory)
        output_entry = tk.Entry(output_frame, textvariable=current_output_var,
                            font=("微軟正黑體", 10),
                            state="readonly", bg="#ecf0f1")
        output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        output_browse_btn = tk.Button(output_frame, text="📁 瀏覽",
                                    command=lambda: self.browse_folder(current_output_var, "選擇輸出目錄"),
                                    font=("微軟正黑體", 10),
                                    bg="#27ae60", fg="white")
        output_browse_btn.pack(side="right")
        
        # 外觀設定
        appearance_frame = tk.LabelFrame(content_frame, text="🎨 外觀設定", 
                                        font=("微軟正黑體", 12, "bold"),
                                        fg="#34495e", padx=15, pady=15)
        appearance_frame.pack(fill="x", pady=(0, 20))
        
        # 暗黑模式選項
        dark_mode_var = tk.BooleanVar(value=(ModernStyle.current_mode == "dark"))
        dark_mode_check = tk.Checkbutton(appearance_frame, 
                                        text="🌙 啟用暗黑模式",
                                        variable=dark_mode_var,
                                        font=("微軟正黑體", 12))
        dark_mode_check.pack(anchor="w", pady=5)
        
        # 說明文字
        desc_label = tk.Label(appearance_frame,
                            text="切換為深色主題，適合低光環境使用",
                            font=("微軟正黑體", 10),
                            fg="#7f8c8d")
        desc_label.pack(anchor="w", padx=20, pady=(0, 10))
        
        # 🔧 按鈕區域 - 固定在最底部
        button_frame = tk.Frame(container)
        button_frame.pack(side="bottom", fill="x", pady=(20, 0))
        
        # 確定按鈕
        def apply_settings():
            new_source = current_source_var.get()
            new_output = current_output_var.get()
            new_dark_mode = dark_mode_var.get()
            
            print(f"🔍 調試：套用設定 - 暗黑模式: {new_dark_mode}", flush=True)
            
            # 更新目錄
            if new_source != self.current_folder:
                self.current_folder = new_source
                self.load_folder_contents()
            
            self.output_directory = new_output
            
            if not os.path.exists(self.output_directory):
                os.makedirs(self.output_directory, exist_ok=True)
            
            # 🔧 新增：統一更新所有匯出器的輸出目錄
            self.update_exporter_output_directory()
            
            # 強制更新暗黑模式
            ModernStyle.current_mode = "dark" if new_dark_mode else "light"
            print(f"🔍 調試：ModernStyle 模式設為: {ModernStyle.current_mode}", flush=True)
            
            # 手動更新顏色
            if new_dark_mode:
                ModernStyle.BG_PRIMARY = "#2b2b2b"
                ModernStyle.BG_LIGHT = "#404040"
                ModernStyle.BG_SECONDARY = "#3c3c3c"
                ModernStyle.TEXT_PRIMARY = "#ffffff"
                ModernStyle.TEXT_SECONDARY = "#b0b0b0"
            else:
                ModernStyle.BG_PRIMARY = "#F8F9FA"
                ModernStyle.BG_LIGHT = "#FFFFFF"
                ModernStyle.BG_SECONDARY = "#E9ECEF"
                ModernStyle.TEXT_PRIMARY = "#212529"
                ModernStyle.TEXT_SECONDARY = "#6C757D"
            
            print(f"🔍 調試：顏色已更新 - BG_PRIMARY: {ModernStyle.BG_PRIMARY}", flush=True)
            
            # 套用主題並儲存
            self.apply_theme()
            self.save_settings()
            
            self.show_custom_messagebox("設定完成", "設定已更新並儲存")
            dialog.destroy()
        
        # 🔧 按鈕使用明確的位置設定
        cancel_btn = tk.Button(button_frame, text="❌ 取消",
                            command=dialog.destroy,
                            font=("微軟正黑體", 12),
                            bg="#95a5a6", fg="white",
                            width=10, height=1)
        cancel_btn.pack(side="left")
        
        # 間隔
        spacer = tk.Frame(button_frame)
        spacer.pack(side="left", fill="x", expand=True)
        
        apply_btn = tk.Button(button_frame, text="✅ 套用設定",
                            command=apply_settings,
                            font=("微軟正黑體", 12, "bold"),
                            bg="#2ecc71", fg="white",
                            width=12, height=1)
        apply_btn.pack(side="right")

    # 🔧 在這裡加入新方法
    def get_current_output_directory(self):
        """取得當前輸出目錄設定"""
        if hasattr(self, 'output_directory') and self.output_directory:
            output_dir = self.output_directory
        else:
            # 🔥 修正：預設輸出目錄使用 BASE_DIR
            output_dir = os.path.join(BASE_DIR, "output")
            self.output_directory = output_dir
        
        # 確保目錄存在
        os.makedirs(output_dir, exist_ok=True)
        return output_dir

    def update_exporter_output_directory(self):
        """更新匯出器的輸出目錄（確保目錄存在）"""
        try:
            current_output = self.get_current_output_directory()
            # exporter.output_dir 是動態 property，會自動從 GUI 讀取
            # 只需確保目錄存在即可
            os.makedirs(current_output, exist_ok=True)
            print(f"📁 輸出目錄已確認：{current_output}", flush=True)
        except Exception as e:
            print(f"❌ 確認輸出目錄失敗：{e}", flush=True)

    def open_output_directory(self):
        """打開輸出目錄 - 使用統一設定"""
        try:
            output_dir = self.get_current_output_directory()
            
            if os.path.exists(output_dir):
                os.startfile(output_dir)  # Windows
                print(f"📂 已打開輸出目錄：{output_dir}", flush=True)
            else:
                print(f"❌ 輸出目錄不存在：{output_dir}", flush=True)
                
        except Exception as e:
            print(f"❌ 打開輸出目錄失敗：{e}", flush=True)

    def preview_dark_mode(self, enable_dark):
        """預覽暗黑模式"""
        # 🔧 修正：使用 set_theme 而不是 switch_mode
        if enable_dark:
            ModernStyle.set_theme("dark")
        else:
            ModernStyle.set_theme("light")

    # 🔧 在 apply_settings 函數中使用更高對比的顏色
    # def apply_settings():
    #     new_source = current_source_var.get()
    #     new_output = current_output_var.get()
    #     new_dark_mode = dark_mode_var.get()
        
    #     print(f"🔍 調試：套用設定 - 暗黑模式: {new_dark_mode}")
        
    #     # 更新目錄
    #     if new_source != self.current_folder:
    #         self.current_folder = new_source
    #         self.load_folder_contents()
        
    #     self.output_directory = new_output
        
    #     if not os.path.exists(self.output_directory):
    #         os.makedirs(self.output_directory, exist_ok=True)
        
    #     # 🔧 記錄原本的模式
    #     old_mode = ModernStyle.current_mode
        
    #     # 設定新模式
    #     ModernStyle.current_mode = "dark" if new_dark_mode else "light"
        
    #     if new_dark_mode:
    #         ModernStyle.BG_PRIMARY = "#1e1e1e"
    #         ModernStyle.BG_LIGHT = "#2d2d2d"
    #         ModernStyle.BG_SECONDARY = "#3a3a3a"
    #         ModernStyle.TEXT_PRIMARY = "#ffffff"
    #         ModernStyle.TEXT_SECONDARY = "#cccccc"
    #     else:
    #         ModernStyle.BG_PRIMARY = "#F8F9FA"
    #         ModernStyle.BG_LIGHT = "#FFFFFF"
    #         ModernStyle.BG_SECONDARY = "#E9ECEF"
    #         ModernStyle.TEXT_PRIMARY = "#212529"
    #         ModernStyle.TEXT_SECONDARY = "#6C757D"
        
    #     # 🔧 如果模式有變化，強制重新創建UI
    #     if old_mode != ModernStyle.current_mode:
    #         self.refresh_ui_theme()
        
    #     # 儲存設定
    #     self.save_settings()
        
    #     self.show_custom_messagebox("設定完成", "主題已切換並儲存設定")
    #     dialog.destroy()

    def refresh_ui_theme(self):
        """刷新整個UI主題"""
        print("🔍 調試：開始刷新UI主題", flush=True)
        
        try:
            # 🔧 強制更新主視窗背景
            self.master.configure(bg=ModernStyle.BG_PRIMARY)
            
            # 🔧 更新頂部標題列
            self.update_header_theme()
            
            # 🔧 更新檔案樹和標題
            if hasattr(self, 'file_tree'):
                self.update_file_tree_theme()
                self.update_left_panel_labels()  # 新增這個方法
            
            # 🔧 更新狀態列
            self.update_status_bar_theme()
            
            # 🔧 更新所有框架
            self.update_all_frames()
            
            # 🔧 強制刷新顯示
            self.master.update()
            
            print("✅ UI主題刷新完成", flush=True)
            
        except Exception as e:
            print(f"❌ 刷新UI主題時發生錯誤: {e}", flush=True)


    def browse_folder(self, var, title):
        """瀏覽資料夾"""
        from tkinter import filedialog
        
        folder = filedialog.askdirectory(title=title, initialdir=var.get())
        if folder:
            var.set(folder)
    
    def create_left_panel(self, parent):
        """創建左側檔案管理面板 - 檔案總管風格"""
        left_panel = ModernFrame(parent, title="📁 檔案管理")

        # 🔧 使用 PanedWindow 的 add 方法而不是 pack
        # 設定初始寬度為 450，但可以調整
        parent.add(left_panel, width=450, minsize=250)

        # 不需要 pack_propagate，因為使用 PanedWindow
        
        # 🔧 工具列（檔案總管風格）
        toolbar = tk.Frame(left_panel, bg=ModernStyle.BG_LIGHT)
        toolbar.pack(fill="x", padx=1, pady=(1, 0))
        
        # 🔧 統一按鈕尺寸和樣式
        button_config = {
            "font": ("微軟正黑體", 16),  # 統一字體大小
            "bg": ModernStyle.BG_SECONDARY, 
            "relief": "flat",
            "width": 3,  # 👈 統一寬度
            "height": 1,  # 👈 統一高度
            "cursor": "hand2"
        }
        
        # 🔧 本機（所有磁碟機）— 放最左
        mycomp_btn = tk.Button(toolbar, text="🖥", command=self.go_to_my_computer, **button_config)
        mycomp_btn.pack(side="left", padx=(0, 5))

        # 🔧 網路（區網電腦）— 放最左
        net_btn = tk.Button(toolbar, text="🌐", command=self.go_to_network, **button_config)
        net_btn.pack(side="left", padx=(0, 5))

        # ⬅️ 返回前頁（瀏覽歷史的上一頁）
        self.back_btn = tk.Button(toolbar, text="⬅️", command=self.go_back, **button_config, state=tk.DISABLED)
        self.back_btn.pack(side="left", padx=(0, 5))

        # ⬆️ 返回上層（回上一層資料夾）
        up_btn = tk.Button(toolbar, text="⬆️", command=self.go_parent_folder, **button_config)
        up_btn.pack(side="left", padx=(0, 5))

        # ➡️ 前進到下一頁（瀏覽歷史的下一頁）
        self.forward_btn = tk.Button(toolbar, text="➡️", command=self.go_forward, **button_config, state=tk.DISABLED)
        self.forward_btn.pack(side="left", padx=(0, 5))

        # 刷新按鈕
        refresh_btn = tk.Button(toolbar, text="🔄", command=self.refresh_folder, **button_config)
        refresh_btn.pack(side="left", padx=(0, 5))

        # 檢視模式切換
        view_btn = tk.Button(toolbar, text="📋", command=self.toggle_view_mode, **button_config)
        view_btn.pack(side="left", padx=(0, 5))

        # 🔧 新增：設定按鈕
        settings_btn = tk.Button(toolbar, text="⚙️", command=self.show_settings_dialog, **button_config)
        settings_btn.pack(side="left", padx=(0, 5))

        # 🔧 修正：改善工具提示位置
        self.create_tooltip(mycomp_btn, "本機（所有磁碟機）")
        self.create_tooltip(net_btn, "網路（區網電腦）")
        self.create_tooltip(self.back_btn, "返回前頁")
        self.create_tooltip(up_btn, "返回上層資料夾")
        self.create_tooltip(self.forward_btn, "前進到下一頁")
        self.create_tooltip(refresh_btn, "重新整理檔案列表")
        self.create_tooltip(view_btn, "切換檢視模式")
        self.create_tooltip(settings_btn, "目錄設定")
        
        # 🔧 路徑列
        path_frame = tk.Frame(left_panel, bg=ModernStyle.BG_LIGHT)
        path_frame.pack(fill="x", padx=1, pady=(1, 0))
        
        tk.Label(path_frame, text="位置:", font=("微軟正黑體", 10),
                bg=ModernStyle.BG_LIGHT, fg=ModernStyle.TEXT_SECONDARY).pack(side="left")

        # 🔧 新增：「前往」按鈕（放在路徑列右側）
        go_btn = tk.Button(path_frame, text="前往", font=("微軟正黑體", 10),
                           bg=ModernStyle.PRIMARY, fg=ModernStyle.TEXT_LIGHT,
                           relief="flat", cursor="hand2",
                           command=self.navigate_to_path)
        go_btn.pack(side="right", padx=(4, 0))

        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(path_frame, textvariable=self.path_var,
                                font=("微軟正黑體", 10), bg="white",
                                relief="solid", borderwidth=1)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(2, 0))
        self.path_entry.bind('<Return>', self.navigate_to_path)
        
        # 🔧 檔案清單區域
        list_container = tk.Frame(left_panel, bg=ModernStyle.BG_LIGHT)
        list_container.pack(fill="both", expand=True, padx=1, pady=(1, 0))
        
        # 使用Treeview創建檔案總管風格列表
        style = ttk.Style()
        style.configure("FileExplorer.Treeview",
                    font=("微軟正黑體", 12),
                    background=ModernStyle.BG_LIGHT,
                    foreground=ModernStyle.TEXT_PRIMARY,
                    selectbackground=ModernStyle.PRIMARY,
                    selectforeground=ModernStyle.TEXT_LIGHT)
        
        # 🔧 設定標題字體大小
        style.configure("FileExplorer.Treeview.Heading",
                    font=("微軟正黑體", 11, "bold"))  # 👈 稍微小一點
                    
        # 🔧 修正：欄位順序改為：名稱、修改日期、類型、大小
        columns = ("modified", "type", "size")
        self.file_tree = ttk.Treeview(list_container,
                                    style="FileExplorer.Treeview",
                                    columns=columns,
                                    show="tree headings",
                                    height=15)

        # 🔧 修正：按新順序設定欄位標題
        self.file_tree.heading("#0", text="📝 名稱", anchor="w",
                            command=lambda: self.sort_files("name"))
        self.file_tree.heading("modified", text="📅 修改日期", anchor="w",
                            command=lambda: self.sort_files("modified"))
        self.file_tree.heading("type", text="📁 類型", anchor="w",
                            command=lambda: self.sort_files("type"))
        self.file_tree.heading("size", text="📏 大小", anchor="e",
                            command=lambda: self.sort_files("size"))

        # 🔧 欄位寬度：名稱欄加大（長檔名一眼看清、會隨面板伸縮）
        self.file_tree.column("#0", width=420, minwidth=260, stretch=True)        # 名稱欄位（加大、可伸縮）
        self.file_tree.column("modified", width=150, minwidth=120, stretch=False) # 修改日期欄位
        self.file_tree.column("type", width=80, minwidth=60, stretch=False)       # 類型欄位
        self.file_tree.column("size", width=90, minwidth=60, stretch=False)       # 大小欄位

        # 🔧 預設只顯示「修改日期」欄；類型/大小可在欄位標題上按右鍵勾選顯示
        self.all_value_columns = ("modified", "type", "size")
        self.visible_columns = ["modified"]
        self._col_vars = {c: tk.BooleanVar(value=(c in self.visible_columns))
                          for c in self.all_value_columns}
        self.file_tree.config(displaycolumns=tuple(self.visible_columns))
        
        # 垂直捲軸
        v_scrollbar = ttk.Scrollbar(list_container, orient="vertical",
                                command=self.file_tree.yview)
        v_scrollbar.pack(side="right", fill="y")

        # 🔧 確保檔案樹正確配置
        self.file_tree.pack(side="left", fill="both", expand=True, padx=0, pady=0)
        self.file_tree.config(yscrollcommand=v_scrollbar.set)
        
        #橫向捲軸
        h_scrollbar = ttk.Scrollbar(left_panel, orient="horizontal",
                                command=self.file_tree.xview)
        h_scrollbar.pack(side="bottom", fill="x", padx=1, pady=(0, 1))
        self.file_tree.config(xscrollcommand=h_scrollbar.set)
        
        # 🔧 事件綁定
        self.file_tree.bind('<Double-Button-1>', self.on_file_double_click)
        self.file_tree.bind('<Button-1>', self.on_file_single_click)
        self.file_tree.bind('<Button-3>', self.show_context_menu)  # 右鍵選單

        # 🔧 鍵盤操作（方便不用滑鼠也能瀏覽）
        #    Enter：開啟選取的資料夾／預覽檔案
        #    Backspace：回上一層　　Alt+← 返回前頁　Alt+→ 前進下一頁
        #    上下方向鍵：Treeview 內建，自動移動選取
        self.file_tree.bind('<Return>', self.on_tree_enter)
        self.file_tree.bind('<KP_Enter>', self.on_tree_enter)
        self.file_tree.bind('<BackSpace>', self.on_tree_backspace)
        self.file_tree.bind('<Alt-Left>', lambda e: self._kb(self.go_back))
        self.file_tree.bind('<Alt-Right>', lambda e: self._kb(self.go_forward))

        # 🔧 一般檔案操作快捷鍵
        #    Del 刪除　Ctrl+C 複製　Ctrl+X 剪下　Ctrl+V 貼上　Ctrl+Z 復原　F2 重新命名
        self.file_tree.bind('<Delete>', lambda e: self._kb(self.delete_selected_files))
        self.file_tree.bind('<Control-c>', lambda e: self._kb(self.copy_selected_files))
        self.file_tree.bind('<Control-C>', lambda e: self._kb(self.copy_selected_files))
        self.file_tree.bind('<Control-x>', lambda e: self._kb(self.cut_selected_files))
        self.file_tree.bind('<Control-X>', lambda e: self._kb(self.cut_selected_files))
        self.file_tree.bind('<Control-v>', lambda e: self._kb(self.paste_files))
        self.file_tree.bind('<Control-V>', lambda e: self._kb(self.paste_files))
        self.file_tree.bind('<Control-z>', lambda e: self._kb(self.undo_last_action))
        self.file_tree.bind('<Control-Z>', lambda e: self._kb(self.undo_last_action))
        self.file_tree.bind('<F2>', lambda e: self._kb(self.rename_selected_file))

        # 🔧 底部資訊列
        info_frame = tk.Frame(left_panel, bg=ModernStyle.BG_LIGHT)
        info_frame.pack(fill="x", padx=1, pady=1)
        
        self.file_count_label = tk.Label(info_frame,
                                        text="",
                                        font=("微軟正黑體", 10),
                                        bg=ModernStyle.BG_LIGHT,
                                        fg=ModernStyle.TEXT_SECONDARY)
        self.file_count_label.pack(side="left")
        
        self.selected_info_label = tk.Label(info_frame,
                                        text="",
                                        font=("微軟正黑體", 10),
                                        bg=ModernStyle.BG_LIGHT,
                                        fg=ModernStyle.TEXT_SECONDARY)
        self.selected_info_label.pack(side="right")
        
        # 初始化
        self.current_folder = PDF_DIR
        self.view_mode = "details"  # details 或 list
        # 🔧 設定預設排序：修改日期降序
        self.sort_column = "modified"
        self.sort_reverse = True  # True = 降序（最新的在上面）
        self.load_folder_contents()

    def on_click_preview(self, event):
        """點擊檔案時預覽"""
        selection = self.pdf_listbox.curselection()
        if selection:
            self.smart_preview()
    
    def sort_files(self, column):
        """排序檔案列表"""
        if not hasattr(self, 'sort_column'):
            self.sort_column = None
            self.sort_reverse = False
        
        # 如果點擊同一欄位，就反向排序
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        
        # 更新欄位標題顯示排序方向
        self.update_sort_indicators()
        
        # 重新載入並排序
        self.load_folder_contents()

    def update_sort_indicators(self):
        """更新排序指示器"""
        # 清除所有排序指示器
        self.file_tree.heading("#0", text="📝 名稱")
        self.file_tree.heading("type", text="📁 類型")
        self.file_tree.heading("size", text="📏 大小")
        self.file_tree.heading("modified", text="📅 修改日期")
        
        # 添加當前排序的指示器
        if hasattr(self, 'sort_column') and self.sort_column:
            arrow = " ▲" if not self.sort_reverse else " ▼"
            
            if self.sort_column == "name":
                self.file_tree.heading("#0", text=f"📝 名稱{arrow}")
            elif self.sort_column == "type":
                self.file_tree.heading("type", text=f"📁 類型{arrow}")
            elif self.sort_column == "size":
                self.file_tree.heading("size", text=f"📏 大小{arrow}")
            elif self.sort_column == "modified":
                self.file_tree.heading("modified", text=f"📅 修改日期{arrow}")

    def create_tooltip(self, widget, text):
        """創建工具提示 - 修正位置"""
        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.configure(bg="lightyellow")
            
            label = tk.Label(tooltip, text=text, 
                            bg="lightyellow", 
                            fg="black",
                            font=("微軟正黑體", 9),
                            padx=5, pady=3)
            label.pack()
            
            # 🔧 修正位置計算 - 避免與滑鼠重疊
            x = event.widget.winfo_rootx() - 50  # 👈 往左偏移
            y = event.widget.winfo_rooty() - 30  # 👈 往上偏移
            
            # 確保不會超出螢幕邊界
            screen_width = tooltip.winfo_screenwidth()
            screen_height = tooltip.winfo_screenheight()
            
            if x < 0:
                x = 10
            if y < 0:
                y = event.widget.winfo_rooty() + event.widget.winfo_height() + 5
            
            tooltip.geometry(f"+{x}+{y}")
            widget.tooltip = tooltip
        
        def on_leave(event):
            if hasattr(event.widget, 'tooltip'):
                event.widget.tooltip.destroy()
                del event.widget.tooltip
        
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def on_file_clicked(self, event):
        """檔案點擊事件處理 - 添加狀態更新"""
        try:
            # 🔧 原有的處理邏輯
            self.master.after(50, self._handle_file_click_with_status)
            
            # 🔧 新增：更新按鈕狀態
            self.master.after(150, self.on_file_selection_changed)  # 稍微延遲確保選擇狀態已更新
            
        except Exception as e:
            print(f"檔案點擊處理錯誤: {e}", flush=True)

        except Exception as e:
            print(f"檔案點擊處理錯誤: {e}", flush=True)

    def _handle_file_click_with_status(self):
        """處理檔案點擊並更新狀態"""
        try:
            selection = self.pdf_listbox.curselection()
            if not selection:
                self.update_struct_tab_status("none")
                return
            
            current_selection = set(selection)
            selected_files = [self.pdf_listbox.get(idx) for idx in selection]
            
            # 🔧 檢查是否有已處理的檔案
            processed_in_selection = any(f in self.processed_files for f in selected_files)
            
            if processed_in_selection and len(selected_files) == 1:
                # 單選且已處理過
                self.update_struct_tab_status("processed", selected_files)
            elif processed_in_selection:
                # 多選包含已處理檔案
                self.update_struct_tab_status("brief", selected_files)
            else:
                # 未處理檔案，進入預覽模式
                self.update_struct_tab_status("preview", selected_files)
            
            # 檢查是否選擇了新的檔案
            if hasattr(self, '_last_preview_selection'):
                if self._last_preview_selection == current_selection:
                    print("🔍 點擊了相同的檔案，跳過重複預覽", flush=True)
                    return
            
            print(f"🖱️ 用戶點擊檔案，直接觸發預覽: {current_selection}", flush=True)
            
            # 直接執行預覽
            self._execute_click_preview(current_selection)
            
        except Exception as e:
            print(f"處理檔案點擊錯誤: {e}", flush=True)

    def _handle_file_click(self):
        """處理檔案點擊的實際邏輯 - 加上重置狀態"""
        try:
            selection = self.pdf_listbox.curselection()
            if not selection:
                return
            
            # 🔧 新增：點擊檔案時重置為未處理狀態
            self.update_tab_title(is_processed=False)
            
            current_selection = set(selection)
            
            # 檢查是否選擇了新的檔案
            if hasattr(self, '_last_preview_selection'):
                if self._last_preview_selection == current_selection:
                    print("🔍 點擊了相同的檔案，跳過重複預覽", flush=True)
                    return
            
            print(f"🖱️ 用戶點擊檔案，直接觸發預覽: {current_selection}", flush=True)
            
            # 直接執行預覽
            self._execute_click_preview(current_selection)
            
        except Exception as e:
            print(f"處理檔案點擊錯誤: {e}", flush=True)



    def _restore_listbox_selection(self, selection_set):
        """恢復列表框的選擇狀態"""
        try:
            self.pdf_listbox.selection_clear(0, tk.END)
            for index in selection_set:
                self.pdf_listbox.selection_set(index)
        except Exception as e:
            print(f"恢復選擇狀態錯誤: {e}", flush=True)

    def _execute_click_preview(self, selected_files):
        """執行點擊觸發的預覽 - 新檔案系統版本"""
        try:
            # 🔧 修改：selected_files 現在是檔案路徑列表，不是索引
            pdf_paths = []
            
            if isinstance(selected_files, (list, set)):
                # 如果是檔案路徑列表
                for file_path in selected_files:
                    if isinstance(file_path, str) and os.path.exists(file_path):
                        pdf_paths.append(file_path)
            
            if not pdf_paths:
                self.show_custom_messagebox("錯誤", "找不到選中的PDF檔案")
                return
            
            # 記錄當前預覽的選擇
            self._last_preview_selection = pdf_paths.copy()
            
            # 根據選擇數量提供不同的狀態提示
            if len(pdf_paths) == 1:
                filename = os.path.basename(pdf_paths[0])
                self.progress_indicator.set_status(f"正在預覽: {filename}", "processing")
            else:
                self.progress_indicator.set_status(f"正在預覽 {len(pdf_paths)} 個檔案...", "processing")
            
            # 只有選擇過多檔案時才提醒（閾值提高到8個）
            if len(pdf_paths) > 8:
                result = self.show_custom_messagebox(
                    "大量檔案預覽", 
                    f"您選擇了 {len(pdf_paths)} 個檔案。\n\n"
                    f"預覽大量檔案可能需要較長時間，\n"
                    f"是否要繼續？", 
                    "yesno"
                )
                if not result:
                    return
            elif len(pdf_paths) > 5:
                # 5-8個檔案只顯示提示，不阻塞
                self.progress_indicator.set_status(f"預覽 {len(pdf_paths)} 個檔案，請稍候...", "warning")
            
            # 立即執行預覽
            self._perform_click_preview(pdf_paths)
            
        except Exception as e:
            print(f"執行點擊預覽錯誤: {e}", flush=True)
            self.progress_indicator.set_status("預覽失敗", "error")

    def _perform_click_preview(self, pdf_paths):
        """執行點擊預覽的核心邏輯 - 優化版本"""
        try:
            # 記錄是否有文字選擇（僅用於日誌）
            had_text_selection = False
            try:
                for text_widget in [self.raw_text, self.struct_text]:
                    if text_widget.tag_ranges("sel"):
                        had_text_selection = True
                        selected_text = text_widget.selection_get()
                        print(f"🔍 覆蓋了用戶的文字選擇 ({len(selected_text)} 字元)", flush=True)
                        break
            except:
                pass
            
            # 清空所有文本框（直接覆蓋，不詢問）
            for text_widget in [self.raw_text, self.struct_text]:
                text_widget.config(state="normal")
                text_widget.delete(1.0, tk.END)
            
            # 清空資訊卡片
            self.clear_info_cards()
            
            # 自動切換到原始文本標籤頁
            self.switch_tab("raw")
            
            # 處理檔案
            for i, pdf_path in enumerate(pdf_paths):
                pdf_name = os.path.basename(pdf_path)
                
                # 更好的進度提示
                if len(pdf_paths) > 1:
                    self.progress_indicator.set_status(
                        f"預覽進度 {i+1}/{len(pdf_paths)}: {pdf_name[:20]}{'...' if len(pdf_name) > 20 else ''}", 
                        "processing"
                    )
                
                # 提取原始文本
                raw_text = extract_pdf_text_only(pdf_path)
                if not raw_text:
                    self.add_preview_content(self.raw_text, f"❌ 檔案: {pdf_name}", "無法提取PDF文字內容")
                    continue
                
                # 根據檔案數量動態調整顯示行數
                if len(pdf_paths) == 1:
                    max_lines = None  # 單檔案顯示完整內容
                elif len(pdf_paths) <= 3:
                    max_lines = 300   # 2-3個檔案顯示較多內容
                else:
                    max_lines = 150   # 多檔案顯示精簡內容
                
                # 顯示原始文本
                self.add_preview_content(self.raw_text, f"📄 檔案: {pdf_name}", raw_text, max_lines)
                
                # 清理文本
                cleaned_text = clean_pagination_simple(raw_text)
                
                # 基本結構化解析
                basic_fields = extract_basic_fields(cleaned_text)
                struct_content = self.format_structured_data(basic_fields)
                self.add_preview_content(self.struct_text, f"📊 基本欄位: {pdf_name}", struct_content)
                
                # 謄本資訊卡片
                transcript_info = detect_transcript_info(raw_text)
                self.add_info_card(pdf_name, transcript_info, basic_fields, raw_text, cleaned_text)
                
                # 每處理幾個檔案更新一次界面（避免阻塞）
                if (i + 1) % 2 == 0:  # 每2個檔案更新一次
                    self.master.update_idletasks()
            
            # 設為只讀
            for text_widget in [self.raw_text, self.struct_text]:
                text_widget.config(state="disabled")
            
            # 滾動到文本開頭
            self.raw_text.see("1.0")
            
            # 更好的完成狀態提示
            if len(pdf_paths) == 1:
                filename = os.path.basename(pdf_paths[0])
                self.progress_indicator.set_status(f"✅ 已預覽: {filename}", "success")
            else:
                self.progress_indicator.set_status(f"✅ 已預覽 {len(pdf_paths)} 個檔案", "success")
                
            # 如果之前有文字選擇，給個提示
            if had_text_selection:
                # 1秒後顯示溫和提示
                self.master.after(1000, lambda: self.progress_indicator.set_status(
                    f"已預覽 {len(pdf_paths)} 個檔案（已清除文字選擇）", "info"
                ))
            
        except Exception as e:
            print(f"點擊預覽執行錯誤: {e}", flush=True)
            self.progress_indicator.set_status("預覽失敗", "error")
            
    def _force_preview_update(self):
        """強制更新預覽（檔案選擇變化時）"""
        try:
            # 取消任何待執行的預覽
            if hasattr(self, '_preview_timer'):
                self.master.after_cancel(self._preview_timer)
            
            # 立即執行預覽更新
            self._execute_auto_preview()
            
        except Exception as e:
            print(f"強制預覽更新錯誤: {e}", flush=True)


    def _execute_auto_preview(self):
        """執行自動預覽 - 最終檢查版本"""
        try:
            selection = self.pdf_listbox.curselection()
            if not selection:
                self.clear_preview()
                return
            
            current_selection = set(selection)
            
            # 🔧 最終檢查：如果用戶仍在選擇文字，再次延遲
            user_still_selecting = False
            try:
                focused_widget = self.master.focus_get()
                if focused_widget in [self.raw_text, self.struct_text]:
                    for text_widget in [self.raw_text, self.struct_text]:
                        if text_widget.tag_ranges("sel"):
                            user_still_selecting = True
                            break
            except:
                pass
            
            if user_still_selecting:
                print("🔍 用戶仍在選擇文字，再次延遲預覽", flush=True)
                # 再延遲1秒
                self._preview_timer = self.master.after(1000, self._execute_auto_preview)
                return
            
            # 檢查選擇是否真的變化了
            if hasattr(self, '_last_selection') and self._last_selection == current_selection:
                print("🔍 檔案選擇未變化，跳過預覽更新", flush=True)
                return
            
            # 記錄當前選擇
            self._last_selection = current_selection.copy()
            
            print(f"🔍 執行預覽更新: {current_selection}", flush=True)
            
            # 更新進度指示器
            self.progress_indicator.set_status(f"自動預覽 {len(selection)} 個檔案...", "processing")
            
            pdf_paths = []
            for idx in selection:
                pdf_name = self.pdf_listbox.get(idx)
                pdf_path = os.path.join(PDF_DIR, pdf_name)
                if os.path.exists(pdf_path):
                    pdf_paths.append(pdf_path)
            
            if not pdf_paths:
                self.clear_preview()
                return
            
            if len(pdf_paths) > 5:
                self.progress_indicator.set_status(f"選擇了 {len(pdf_paths)} 個檔案，自動預覽可能較慢", "warning")
            
            # 執行預覽
            self._perform_auto_preview(pdf_paths)
            
        except Exception as e:
            print(f"自動預覽錯誤: {e}", flush=True)
            self.progress_indicator.set_status("自動預覽失敗", "error")

    def _perform_auto_preview(self, pdf_paths):
        """執行自動預覽的核心邏輯 - 保護文字選擇版本"""
        try:
            # 🔧 執行前的最後檢查
            user_has_selection = False
            selected_text = ""
            selected_widget = None
            selection_start = None
            selection_end = None
            
            # 保存當前的文字選擇狀態
            try:
                for text_widget in [self.raw_text, self.struct_text]:
                    if text_widget.tag_ranges("sel"):
                        user_has_selection = True
                        selected_widget = text_widget
                        selected_text = text_widget.selection_get()
                        selection_start = text_widget.index("sel.first")
                        selection_end = text_widget.index("sel.last")
                        print(f"🔍 保存文字選擇: {len(selected_text)} 字元", flush=True)
                        break
            except:
                pass
            
            if user_has_selection:
                # 如果用戶有文字選擇，詢問是否要繼續
                print("🔍 檢測到文字選擇，跳過此次預覽更新", flush=True)
                self.progress_indicator.set_status("跳過預覽更新（保護文字選擇）", "warning")
                return
            
            # 清空所有文本框
            for text_widget in [self.raw_text, self.struct_text]:
                text_widget.config(state="normal")
                text_widget.delete(1.0, tk.END)
            
            # 清空資訊卡片
            self.clear_info_cards()
            
            # 自動切換到原始文本標籤頁
            self.switch_tab("raw")
            
            # 處理檔案
            for i, pdf_path in enumerate(pdf_paths):
                pdf_name = os.path.basename(pdf_path)
                
                if len(pdf_paths) > 1:
                    self.progress_indicator.set_status(f"預覽中 {i+1}/{len(pdf_paths)}: {pdf_name}", "processing")
                
                # 提取原始文本
                raw_text = extract_pdf_text_only(pdf_path)
                if not raw_text:
                    self.add_preview_content(self.raw_text, f"❌ 檔案: {pdf_name}", "無法提取PDF文字內容")
                    continue
                
                max_lines = 200 if len(pdf_paths) > 1 else 500
                
                # 顯示原始文本
                self.add_preview_content(self.raw_text, f"📄 檔案: {pdf_name}", raw_text, max_lines)
                
                # 清理文本
                cleaned_text = clean_pagination_simple(raw_text)
                
                # 基本結構化解析
                basic_fields = extract_basic_fields(cleaned_text)
                struct_content = self.format_structured_data(basic_fields)
                self.add_preview_content(self.struct_text, f"📊 基本欄位: {pdf_name}", struct_content)
                
                # 謄本資訊卡片
                transcript_info = detect_transcript_info(raw_text)
                self.add_info_card(pdf_name, transcript_info, basic_fields, raw_text, cleaned_text)
                
                self.master.update_idletasks()
            
            # 設為只讀
            for text_widget in [self.raw_text, self.struct_text]:
                text_widget.config(state="disabled")
            
            # 滾動到文本開頭
            self.raw_text.see("1.0")
            
            # 更新狀態
            if len(pdf_paths) == 1:
                self.progress_indicator.set_status(f"已預覽: {os.path.basename(pdf_paths[0])}", "success")
            else:
                self.progress_indicator.set_status(f"已預覽 {len(pdf_paths)} 個檔案", "success")
            
        except Exception as e:
            print(f"自動預覽執行錯誤: {e}", flush=True)
            self.progress_indicator.set_status("自動預覽失敗", "error")

    # 🔧 3. 修改後的 create_preview_tabs 方法
    def create_preview_tabs(self):
        """創建預覽標籤頁 - 簡化版本"""
        
        # 📄 原始文本標籤頁
        self.raw_frame = ttk.Frame(self.notebook, style='Modern.TFrame')
        self.notebook.add(self.raw_frame, text="📄 原始文本")
        
        raw_container = tk.Frame(self.raw_frame, bg=ModernStyle.BG_LIGHT)
        raw_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 添加提示標籤
        raw_tip = tk.Label(raw_container, 
                        text="💡 顯示PDF提取的原始文字內容",
                        font=("微軟正黑體", 10),
                        bg=ModernStyle.BG_LIGHT,
                        fg=ModernStyle.TEXT_SECONDARY)
        raw_tip.pack(anchor="w", pady=(0, 5))
        
        # 在創建文字框後，添加事件綁定
        self.raw_text = scrolledtext.ScrolledText(raw_container, 
                                                font=("微軟正黑體", 14),
                                                bg=ModernStyle.BG_LIGHT,
                                                fg=ModernStyle.TEXT_PRIMARY,
                                                relief="solid",
                                                borderwidth=1,
                                                wrap=tk.WORD,
                                                selectbackground="#3390FF",
                                                selectforeground="white")
        self.raw_text.pack(fill="both", expand=True)
        self._add_text_context_menu(self.raw_text)

        # 🔧 添加滑鼠事件，改善文字選擇體驗
        def on_text_click(event):
            """文字框點擊事件"""
            event.widget.focus_set()
            
        def on_text_drag(event):
            """文字框拖拽事件"""
            event.widget.focus_set()
        
        self.raw_text.bind("<Button-1>", on_text_click)
        self.raw_text.bind("<B1-Motion>", on_text_drag)
        
        # 📊 結構化資料標籤頁
        self.struct_frame = ttk.Frame(self.notebook, style='Modern.TFrame')
        self.notebook.add(self.struct_frame, text="📊 結構化資料")
        
        struct_container = tk.Frame(self.struct_frame, bg=ModernStyle.BG_LIGHT)
        struct_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 添加提示標籤
        struct_tip = tk.Label(struct_container, 
                            text="💡 提取的結構化欄位資料和處理結果",
                            font=("微軟正黑體", 9),
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.TEXT_SECONDARY)
        struct_tip.pack(anchor="w", pady=(0, 5))
        
        self.struct_text = scrolledtext.ScrolledText(struct_container, 
                                                font=("微軟正黑體", 14),
                                                bg=ModernStyle.BG_LIGHT,
                                                fg=ModernStyle.TEXT_PRIMARY,
                                                relief="solid",
                                                borderwidth=1,
                                                wrap=tk.WORD,
                                                selectbackground="#3390FF",
                                                selectforeground="white")
        self.struct_text.pack(fill="both", expand=True)
        self._add_text_context_menu(self.struct_text)

        # 對結構化文字框做相同處理
        self.struct_text.bind("<Button-1>", on_text_click)
        self.struct_text.bind("<B1-Motion>", on_text_drag)
        
        # ℹ️ 謄本資訊標籤頁
        self.info_frame = ttk.Frame(self.notebook, style='Modern.TFrame')
        self.notebook.add(self.info_frame, text="ℹ️ 謄本資訊")
        
        info_container = tk.Frame(self.info_frame, bg=ModernStyle.BG_LIGHT)
        info_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 添加提示標籤
        info_tip = tk.Label(info_container, 
                        text="💡 謄本基本資訊和統計數據的卡片式展示",
                        font=("微軟正黑體", 9),
                        bg=ModernStyle.BG_LIGHT,
                        fg=ModernStyle.TEXT_SECONDARY)
        info_tip.pack(anchor="w", pady=(0, 5))
        
        # 謄本資訊使用更友善的卡片式布局
        self.create_info_cards(info_container)

    def _add_text_context_menu(self, text_widget):
        """為文字框添加右鍵選單"""
        context_menu = tk.Menu(text_widget, tearoff=0)
        context_menu.add_command(label="📄 複製", command=lambda: self._copy_selected_text(text_widget))
        context_menu.add_command(label="🔍 全選", command=lambda: self._select_all_text(text_widget))
        context_menu.add_separator()
        context_menu.add_command(label="🔍 尋找...", command=lambda: self._show_find_dialog(text_widget))
        
        def show_context_menu(event):
            print("🔍 預覽框右鍵選單被觸發")  # 調試用
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()
        
        text_widget.bind("<Button-3>", show_context_menu)  # 右鍵

    def _copy_selected_text(self, text_widget):
        """複製選中的文字"""
        try:
            selected_text = text_widget.selection_get()
            self.master.clipboard_clear()
            self.master.clipboard_append(selected_text)
            print("✅ 文字已複製到剪貼簿", flush=True)
        except tk.TclError:
            print("❌ 沒有選擇文字", flush=True)

    def _select_all_text(self, text_widget):
        """全選文字"""
        text_widget.tag_add(tk.SEL, "1.0", tk.END)
        text_widget.mark_set(tk.INSERT, "1.0")
        text_widget.see(tk.INSERT)

    def _show_find_dialog(self, text_widget):
        """顯示尋找對話框"""
        # 創建尋找對話框
        find_dialog = tk.Toplevel(self.master)
        find_dialog.title("尋找")
        find_dialog.geometry("400x200")
        find_dialog.transient(self.master)
        find_dialog.grab_set()
        
        # 居中顯示
        find_dialog.update_idletasks()
        x = (find_dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (find_dialog.winfo_screenheight() // 2) - (200 // 2)
        find_dialog.geometry(f"400x200+{x}+{y}")
        
        # 尋找關鍵字輸入
        tk.Label(find_dialog, text="尋找:", font=("微軟正黑體", 12)).pack(pady=(10, 5))
        
        search_var = tk.StringVar()
        search_entry = tk.Entry(find_dialog, textvariable=search_var, 
                            font=("微軟正黑體", 12), width=35)
        search_entry.pack(pady=5)
        search_entry.focus()
        
        # 選項框架
        options_frame = tk.Frame(find_dialog)
        options_frame.pack(pady=10)
        
        # 大小寫敏感選項
        case_sensitive_var = tk.BooleanVar()
        case_check = tk.Checkbutton(options_frame, text="區分大小寫", 
                                variable=case_sensitive_var,
                                font=("微軟正黑體", 10))
        case_check.pack(side="left", padx=10)
        
        # 結果顯示
        result_label = tk.Label(find_dialog, text="", 
                            font=("微軟正黑體", 10),
                            fg="blue")
        result_label.pack(pady=5)
        
        # 搜尋狀態變數
        current_index = tk.StringVar(value="1.0")
        matches_found = []
        current_match = 0
        
        def do_search():
            """執行搜尋 - 修正版本"""
            nonlocal matches_found, current_match
            
            search_text = search_var.get().strip()
            if not search_text:
                result_label.config(text="請輸入搜尋關鍵字", fg="red")
                return
            
            # 清除之前的高亮
            text_widget.tag_remove("search_highlight", "1.0", tk.END)
            text_widget.tag_remove("current_highlight", "1.0", tk.END)
            
            # 🔧 修正：使用tkinter的search方法而不是Python的字串搜尋
            matches_found = []
            
            # 準備搜尋選項
            if case_sensitive_var.get():
                search_options = []
            else:
                search_options = ["-nocase"]  # 不區分大小寫
            
            # 從文字開頭開始搜尋
            start_pos = "1.0"
            
            while True:
                # 🔧 使用tkinter的內建搜尋方法
                if case_sensitive_var.get():
                    pos = text_widget.search(search_text, start_pos, tk.END)
                else:
                    pos = text_widget.search(search_text, start_pos, tk.END, nocase=True)
                
                if not pos:
                    break
                
                # 計算結束位置
                end_pos = f"{pos}+{len(search_text)}c"
                matches_found.append((pos, end_pos))
                
                # 移動到下一個搜尋位置
                start_pos = f"{pos}+1c"
            
            if matches_found:
                current_match = 0
                highlight_current_match()
                result_label.config(text=f"找到 {len(matches_found)} 個結果", fg="green")
            else:
                result_label.config(text="找不到符合的結果", fg="red") 


        def highlight_current_match():
            """高亮當前匹配項"""
            if not matches_found:
                return
            
            # 清除所有高亮
            text_widget.tag_remove("search_highlight", "1.0", tk.END)
            text_widget.tag_remove("current_highlight", "1.0", tk.END)
            
            # 高亮所有匹配項
            for start, end in matches_found:
                text_widget.tag_add("search_highlight", start, end)
            
            # 特別高亮當前項
            if 0 <= current_match < len(matches_found):
                start, end = matches_found[current_match]
                text_widget.tag_add("current_highlight", start, end)
                text_widget.see(start)  # 滾動到當前匹配項
                
                result_label.config(
                    text=f"第 {current_match + 1} / {len(matches_found)} 個結果", 
                    fg="blue"
                )
        
        def find_next():
            """下一個匹配項"""
            nonlocal current_match
            if matches_found and current_match < len(matches_found) - 1:
                current_match += 1
                highlight_current_match()
        
        def find_previous():
            """上一個匹配項"""
            nonlocal current_match
            if matches_found and current_match > 0:
                current_match -= 1
                highlight_current_match()
        
        # 按鈕框架
        button_frame = tk.Frame(find_dialog)
        button_frame.pack(pady=10)
        
        # 搜尋按鈕
        search_btn = tk.Button(button_frame, text="搜尋", command=do_search,
                            font=("微軟正黑體", 11), width=8)
        search_btn.pack(side="left", padx=5)
        
        # 下一個按鈕
        next_btn = tk.Button(button_frame, text="下一個", command=find_next,
                            font=("微軟正黑體", 11), width=8)
        next_btn.pack(side="left", padx=5)
        
        # 上一個按鈕
        prev_btn = tk.Button(button_frame, text="上一個", command=find_previous,
                            font=("微軟正黑體", 11), width=8)
        prev_btn.pack(side="left", padx=5)
        
        # 關閉按鈕
        close_btn = tk.Button(button_frame, text="關閉", 
                            command=lambda: close_dialog(),
                            font=("微軟正黑體", 11), width=8)
        close_btn.pack(side="left", padx=5)
        
        def close_dialog():
            """關閉對話框並清除高亮"""
            text_widget.tag_remove("search_highlight", "1.0", tk.END)
            text_widget.tag_remove("current_highlight", "1.0", tk.END)
            find_dialog.destroy()
        
        # 設定高亮樣式
        text_widget.tag_configure("search_highlight", background="yellow")
        text_widget.tag_configure("current_highlight", background="orange", foreground="black")
        
        # 綁定Enter鍵執行搜尋
        search_entry.bind('<Return>', lambda e: do_search())
        
        # 處理對話框關閉
        find_dialog.protocol("WM_DELETE_WINDOW", close_dialog)

    def create_center_panel(self, parent):
        """創建中間預覽面板 - 統一邊框包含標題版本"""
        # 🔧 整個中間面板（統一大邊框，包含標題）
        center_panel = tk.Frame(parent, bg=ModernStyle.BG_LIGHT,
                            relief="solid", borderwidth=1)

        # 🔧 使用 PanedWindow 的 add 方法而不是 pack
        # 設定為可伸縮，占用剩餘空間
        parent.add(center_panel, stretch="always")
        
        # 🔧 標題和提示區域（在邊框內的頂部）
        header_frame = tk.Frame(center_panel, bg=ModernStyle.BG_LIGHT)
        header_frame.pack(fill="x", padx=15, pady=(5, 4))
        
        # 左側標題
        title_label = tk.Label(header_frame, 
                            text="👁️ 智慧預覽與分析",
                            font=ModernStyle.FONT_HEADER,
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.PRIMARY)
        title_label.pack(side="left")
        
        # 右側提示文字
        tip_label = tk.Label(header_frame, 
                            text="💡 點擊左側檔案即可預覽，可自由選擇和複製文字內容",
                            font=("微軟正黑體", 9),
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.TEXT_SECONDARY)
        tip_label.pack(side="right")
        
        # 🔧 標籤頁標題列
        tab_header = tk.Frame(center_panel, bg=ModernStyle.BG_LIGHT)
        tab_header.pack(fill="x", padx=10, pady=(0, 5))
        
        # 左側：手動建立標籤頁按鈕
        tab_buttons_frame = tk.Frame(tab_header, bg=ModernStyle.BG_LIGHT)
        tab_buttons_frame.pack(side="left")
        
        # 標籤頁按鈕
        self.tab_buttons = {}
        tab_configs = [
            ("📄 原始文本", "raw"),
            ("📊 結構化資料", "struct"), 
            ("ℹ️ 謄本資訊", "info")
        ]
        
        for i, (text, key) in enumerate(tab_configs):
            btn = tk.Button(tab_buttons_frame,
                        text=text,
                        command=lambda k=key: self.switch_tab(k),
                        font=("微軟正黑體", 12), #"原始文本"、"結構化處理"、"謄本資訊"字體大小
                        bg=ModernStyle.BG_SECONDARY if i == 0 else ModernStyle.BG_LIGHT,
                        fg=ModernStyle.TEXT_PRIMARY,
                        relief="solid",
                        borderwidth=1,
                        padx=15, pady=8,
                        cursor="hand2")
            btn.pack(side="left", padx=(0, 2))
            self.tab_buttons[key] = btn
        
        # 右側：控制按鈕（與標籤頁同行）
        control_buttons = tk.Frame(tab_header, bg=ModernStyle.BG_LIGHT)
        control_buttons.pack(side="right")
        
        ModernButton(control_buttons, "🔄 重新預覽", 
                    command=self.smart_preview, 
                    style="primary").pack(side="left", padx=(0, 5))
        
        ModernButton(control_buttons, "🧹 清空預覽", 
                    command=self.clear_preview, 
                    style="secondary").pack(side="left")
        
        # 🔧 內容區域（在同一個邊框內）
        self.content_area = tk.Frame(center_panel, bg=ModernStyle.BG_LIGHT)
        self.content_area.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # 創建各個標籤頁內容
        self.create_tab_contents()
        
        # 預設顯示第一個標籤頁
        self.current_tab = "raw"
        self.switch_tab("raw")

    def create_tab_contents(self):
        """創建標籤頁內容"""
        
        # 📄 原始文本內容
        self.raw_frame = tk.Frame(self.content_area, bg=ModernStyle.BG_LIGHT)
        
        raw_tip = tk.Label(self.raw_frame, 
                        text="💡 顯示PDF提取的原始文字內容",
                        font=("微軟正黑體", 10),
                        bg=ModernStyle.BG_LIGHT,
                        fg=ModernStyle.TEXT_SECONDARY)
        raw_tip.pack(anchor="w", pady=(0, 5))
        
        self.raw_text = scrolledtext.ScrolledText(self.raw_frame, 
                                                font=("微軟正黑體", 14),
                                                bg=ModernStyle.BG_LIGHT,
                                                fg=ModernStyle.TEXT_PRIMARY,
                                                relief="solid",
                                                borderwidth=1,
                                                wrap=tk.WORD,
                                                selectbackground="#3390FF",
                                                selectforeground="white")
        self.raw_text.pack(fill="both", expand=True)
        self._add_text_context_menu(self.raw_text)  # 👈 加入這一行

        # 📊 結構化資料內容
        self.struct_frame = tk.Frame(self.content_area, bg=ModernStyle.BG_LIGHT)
        
        struct_tip = tk.Label(self.struct_frame, 
                            text="💡 提取的結構化欄位資料和處理結果",
                            font=("微軟正黑體", 10),
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.TEXT_SECONDARY)
        struct_tip.pack(anchor="w", pady=(0, 5))
        
        self.struct_text = scrolledtext.ScrolledText(self.struct_frame, 
                                                font=("微軟正黑體", 14),
                                                bg=ModernStyle.BG_LIGHT,
                                                fg=ModernStyle.TEXT_PRIMARY,
                                                relief="solid",
                                                borderwidth=1,
                                                wrap=tk.WORD,
                                                selectbackground="#3390FF",
                                                selectforeground="white")
        self.struct_text.pack(fill="both", expand=True)
        self._add_text_context_menu(self.struct_text)  # 👈 加入這一行

        # ℹ️ 謄本資訊內容
        self.info_frame = tk.Frame(self.content_area, bg=ModernStyle.BG_LIGHT)
        
        info_tip = tk.Label(self.info_frame, 
                        text="💡 謄本基本資訊和統計數據的卡片式展示",
                        font=("微軟正黑體", 10),
                        bg=ModernStyle.BG_LIGHT,
                        fg=ModernStyle.TEXT_SECONDARY)
        info_tip.pack(anchor="w", pady=(0, 5))
        
        # 謄本資訊使用更友善的卡片式布局
        self.create_info_cards(self.info_frame)
        
        # 🔧 添加文字選擇事件
        def on_text_click(event):
            event.widget.focus_set()
            
        def on_text_drag(event):
            event.widget.focus_set()
        
        self.raw_text.bind("<Button-1>", on_text_click)
        self.raw_text.bind("<B1-Motion>", on_text_drag)
        self.struct_text.bind("<Button-1>", on_text_click)
        self.struct_text.bind("<B1-Motion>", on_text_drag)

    def switch_tab(self, tab_key):
        """切換標籤頁"""
        # 隱藏所有標籤頁
        for frame in [self.raw_frame, self.struct_frame, self.info_frame]:
            frame.pack_forget()
        
        # 重設所有按鈕樣式
        for key, btn in self.tab_buttons.items():
            if key == tab_key:
                btn.config(bg=ModernStyle.PRIMARY, fg=ModernStyle.TEXT_LIGHT)
            else:
                btn.config(bg=ModernStyle.BG_SECONDARY, fg=ModernStyle.TEXT_PRIMARY)
        
        # 顯示對應標籤頁
        if tab_key == "raw":
            self.raw_frame.pack(fill="both", expand=True)
        elif tab_key == "struct":
            self.struct_frame.pack(fill="both", expand=True)
        elif tab_key == "info":
            self.info_frame.pack(fill="both", expand=True)
        
        self.current_tab = tab_key

    def create_info_cards(self, parent):
        """創建謄本資訊卡片"""
        
        # 捲動容器
        canvas = tk.Canvas(parent, bg=ModernStyle.BG_LIGHT, highlightthickness=0)
        scrollbar = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self.info_scrollable_frame = tk.Frame(canvas, bg=ModernStyle.BG_LIGHT)
        
        self.info_scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.info_scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 初始提示
        self.info_placeholder = tk.Label(self.info_scrollable_frame,
                                        text="📋 請選擇PDF檔案並點擊預覽以查看謄本資訊",
                                        font=ModernStyle.FONT_MAIN,
                                        bg=ModernStyle.BG_LIGHT,
                                        fg=ModernStyle.TEXT_SECONDARY)
        self.info_placeholder.pack(pady=50)

    def create_right_panel(self, parent):
        """創建右側操作面板"""
        right_panel = ModernFrame(parent, title="⚙️ 智慧操作控制")

        # 🔧 使用 PanedWindow 的 add 方法而不是 pack
        # 設定初始寬度為 200，最小寬度為 180
        parent.add(right_panel, width=200, minsize=180)

        # 不需要 pack_propagate，因為使用 PanedWindow
        
        # 處理功能區
        self.create_processing_section(right_panel)
        
        # 匯出功能區  
        self.create_export_section(right_panel)
        
        # 工具功能區
        self.create_tools_section(right_panel)
    
    def create_processing_section(self, parent):
        """創建處理功能區 - 簡化版本（移除簡略查詢）"""
        processing_section = tk.Frame(parent, bg=ModernStyle.BG_LIGHT)
        processing_section.pack(fill="x", pady=(0, 3)) #調間距

        # 處理按鈕
        ModernButton(processing_section, "📋 批次處理",
                    command=self.batch_process,
                    style="primary").pack(fill="x", pady=3)

        # 🔥 跨資料夾掃描批次處理（自動遞迴掃描 + 篩選謄本 + 按案件分組）
        ModernButton(processing_section, "🔍 跨資料夾掃描",
                    command=self.open_batch_scan_dialog,
                    style="warning").pack(fill="x", pady=3)


    def create_export_section(self, parent):
        # 🔧 臨時修復：如果沒有模板設定就先建立
        if not hasattr(self, 'export_templates'):
            self.export_templates = {
                "完整": "all",
                "核心": "core", 
                "房產": "property",
                "權利": "rights",
            }
            self.current_template = "完整"
        """創建匯出功能區 - 添加快速模板選擇"""
        export_section = tk.Frame(parent, bg=ModernStyle.BG_LIGHT)
        export_section.pack(fill="x", pady=(0, 3)) #調間距
        
        # 區域標題
        title = tk.Label(export_section, text="📤 資料匯出",
                        font=ModernStyle.FONT_HEADER,
                        bg=ModernStyle.BG_LIGHT,
                        fg=ModernStyle.SECONDARY)
        title.pack(anchor="w", pady=(0, 3))
        
        # 🔧 新增：快速模板選擇區
        template_frame = tk.Frame(export_section, 
                                bg=ModernStyle.BG_SECONDARY,
                                relief="solid", 
                                borderwidth=1,
                                padx=10, pady=8)
        template_frame.pack(fill="x", pady=(0, 10))
        
        # 快速模板選擇標題
        template_label = tk.Label(template_frame, 
                                # text="💡",
                                font=("微軟正黑體", 10),
                                bg=ModernStyle.BG_SECONDARY,
                                fg=ModernStyle.TEXT_PRIMARY)
        template_label.pack(side="left")
        
        # 🔧 模板按鈕（支援面板縮放調整）
        template_buttons_container = tk.Frame(template_frame, bg=ModernStyle.BG_SECONDARY)
        template_buttons_container.pack(fill="both", expand=True, pady=(5, 0))

        # 創建兩行模板按鈕
        self.template_buttons = {}
        template_names = list(self.export_templates.keys())

        # 第一行：完整、核心
        row1_frame = tk.Frame(template_buttons_container, bg=ModernStyle.BG_SECONDARY)
        row1_frame.pack(fill="both", expand=True, pady=(0, 3))

        # 配置第一行的欄位權重
        row1_frame.grid_columnconfigure(0, weight=1)
        row1_frame.grid_columnconfigure(1, weight=1)

        for i, template_name in enumerate(template_names[:2]):  # 前兩個
            btn = tk.Button(row1_frame,
                        text=template_name,
                        command=lambda t=template_name: self.select_template(t),
                        font=("微軟正黑體", 11),
                        relief="solid",
                        borderwidth=1,
                        padx=4, pady=2,
                        cursor="hand2")
            btn.grid(row=0, column=i, sticky="ew", padx=(0, 2) if i == 0 else (0, 0))
            self.template_buttons[template_name] = btn

        # 第二行：房產、權利
        row2_frame = tk.Frame(template_buttons_container, bg=ModernStyle.BG_SECONDARY)
        row2_frame.pack(fill="both", expand=True)

        # 配置第二行的欄位權重
        row2_frame.grid_columnconfigure(0, weight=1)
        row2_frame.grid_columnconfigure(1, weight=1)

        for i, template_name in enumerate(template_names[2:]):  # 後兩個
            btn = tk.Button(row2_frame,
                        text=template_name,
                        command=lambda t=template_name: self.select_template(t),
                        font=("微軟正黑體", 11),
                        relief="solid",
                        borderwidth=1,
                        padx=4, pady=2,
                        cursor="hand2")
            btn.grid(row=0, column=i, sticky="ew", padx=(0, 2) if i == 0 else (0, 0))
            self.template_buttons[template_name] = btn
        
        # 設定初始選中狀態
        self.update_template_buttons()
        
        # 🔧 當前選擇顯示
        self.current_selection_label = tk.Label(template_frame,
                                            text="(完整資料)",
                                            font=("微軟正黑體", 11),
                                            bg=ModernStyle.BG_SECONDARY,
                                            fg=ModernStyle.INFO)
        self.current_selection_label.pack(side="right")
        
        # 🔧 修改：使用模板的匯出按鈕（支援面板縮放調整）
        buttons_container = tk.Frame(export_section, bg=ModernStyle.BG_LIGHT)
        buttons_container.pack(fill="both", expand=True)

        ModernButton(buttons_container, "📄 匯出 TXT",
                    command=lambda: self.export_with_template("txt"),
                    style="secondary").pack(fill="both", expand=True, pady=2)

        ModernButton(buttons_container, "📈 匯出 Excel",
                    command=lambda: self.export_with_template("excel"),
                    style="success").pack(fill="both", expand=True, pady=2)

        ModernButton(buttons_container, "🔗 匯出 JSON",
                    command=lambda: self.export_with_template("json"),
                    style="warning").pack(fill="both", expand=True, pady=2)

        ModernButton(buttons_container, "📦 匯出全部",
                    command=self.export_all_with_template,
                    style="info").pack(fill="both", expand=True, pady=2)


    def create_tools_section(self, parent):
        """創建工具功能區"""
        tools_section = tk.Frame(parent, bg=ModernStyle.BG_LIGHT)
        tools_section.pack(fill="x", pady=(0, 5)) #調間距
        
        # 區域標題
        title = tk.Label(tools_section, text="🛠️實用工具",
                        font=ModernStyle.FONT_HEADER,
                        bg=ModernStyle.BG_LIGHT,
                        fg=ModernStyle.INFO)
        title.pack(anchor="w", pady=(0, 10))
        
        # 工具按鈕
        # 第一行工具按鈕
        row1_frame = tk.Frame(tools_section, bg=ModernStyle.BG_LIGHT)
        row1_frame.pack(fill="x", pady=2)

        ModernButton(row1_frame, "❓ 說明", 
                    command=self.show_help, 
                    style="secondary", width=5).pack(side="left", fill="x", expand=True, padx=(0, 2))

        ModernButton(row1_frame, "📂 輸出", 
            command=self.open_output_directory,  # ← 改用正確的方法
            style="secondary", width=5).pack(side="right", fill="x", expand=True, padx=(2, 0))

        # 第二行工具按鈕
        row2_frame = tk.Frame(tools_section, bg=ModernStyle.BG_LIGHT)
        row2_frame.pack(fill="x", pady=2)

        # 🔧 修改設定按鈕為程式狀態
        ModernButton(row2_frame, "📊 狀態", 
            command=self.show_program_status,
            style="secondary", width=5).pack(side="left", fill="x", expand=True, padx=(0, 2))

        ModernButton(row2_frame, "❌ 關閉", 
                    command=self.master.quit, 
                    style="danger", width=5).pack(side="right", fill="x", expand=True, padx=(2, 0))

    def create_footer(self):
        """創建底部狀態列 - 配合UI色調"""
        footer = tk.Frame(self.master, bg=ModernStyle.BG_LIGHT, height=30)  # 👈 使用UI統一背景色
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        
        # 簡單的狀態訊息
        self.status_label = tk.Label(footer, 
                                    text="🟢 就緒", 
                                    bg=ModernStyle.BG_LIGHT, 
                                    fg=ModernStyle.TEXT_PRIMARY,
                                    font=("微軟正黑體", 9))
        self.status_label.pack(side="left", padx=10, pady=5)
        
        # 版本資訊
        version_label = tk.Label(footer,
                            text="v1.7 (2026-07-23) | 電子謄本功能測試版",
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.TEXT_SECONDARY,
                            font=("微軟正黑體", 9))
        version_label.pack(side="right", padx=10, pady=5)

    def update_status(self, message):
        """更新狀態列訊息"""
        if hasattr(self, 'status_label'):
            self.status_label.config(text=f"🟢 {message}")

    def show_custom_messagebox(self, title, message, msg_type="info"):
        """自定義訊息對話框 - 修正版本"""
        dialog = tk.Toplevel(self.master)
        dialog.title(title)
        
        # 🔧 關鍵修正：根據訊息長度動態調整對話框大小
        message_lines = message.count('\n') + 1
        base_height = 200
        line_height = 25
        button_height = 80
        padding = 60
        
        # 計算所需高度
        calculated_height = base_height + (message_lines * line_height) + button_height + padding
        
        # 設定最小和最大高度
        min_height = 280
        max_height = 600
        final_height = max(min_height, min(calculated_height, max_height))
        
        dialog_width = 500
        dialog.geometry(f"{dialog_width}x{final_height}")
        dialog.configure(bg=ModernStyle.BG_LIGHT)
        dialog.transient(self.master)
        dialog.grab_set()
        
        # 🔧 修正定位：手動計算中央位置
        # 等待主視窗更新
        self.master.update_idletasks()
        dialog.update_idletasks()
        
        # 獲取主視窗位置和大小
        main_x = self.master.winfo_rootx()
        main_y = self.master.winfo_rooty()
        main_width = self.master.winfo_width()
        main_height = self.master.winfo_height()
        
        # 計算對話框應該出現的中央位置
        dialog_x = main_x + (main_width - dialog_width) // 2
        dialog_y = main_y + (main_height - final_height) // 2
        
        # 確保對話框不會超出螢幕邊界
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        
        dialog_x = max(0, min(dialog_x, screen_width - dialog_width))
        dialog_y = max(0, min(dialog_y, screen_height - final_height))
        
        # 設定位置
        dialog.geometry(f"{dialog_width}x{final_height}+{dialog_x}+{dialog_y}")
        
        
        # 🔧 修正：使用 pack 而不是固定位置，確保所有元件都能顯示
        # 主要內容容器
        main_frame = tk.Frame(dialog, bg=ModernStyle.BG_LIGHT)
        main_frame.pack(fill="both", expand=True, padx=25, pady=25)
        
        # 圖示區域
        icon_frame = tk.Frame(main_frame, bg=ModernStyle.BG_LIGHT)
        icon_frame.pack(side="top", pady=(0, 15))
        
        # 根據類型顯示不同圖示
        if msg_type == "info" or msg_type == "success":
            icon_text = "✅"
            icon_color = ModernStyle.SUCCESS
        elif msg_type == "warning":
            icon_text = "⚠️"
            icon_color = ModernStyle.WARNING
        elif msg_type == "error":
            icon_text = "❌"
            icon_color = ModernStyle.DANGER
        else:
            icon_text = "❓"
            icon_color = ModernStyle.INFO
        
        icon_label = tk.Label(icon_frame, 
                            text=icon_text,
                            font=("Segoe UI", 24),
                            bg=ModernStyle.BG_LIGHT,
                            fg=icon_color)
        icon_label.pack()
        
        # 🔧 修正：訊息文字區域 - 使用可滾動的文字框處理長訊息
        if message_lines > 8:  # 如果訊息超過8行，使用滾動文字框
            msg_frame = tk.Frame(main_frame, bg=ModernStyle.BG_LIGHT)
            msg_frame.pack(fill="both", expand=True, pady=15)
            
            # 創建滾動文字框
            msg_text = scrolledtext.ScrolledText(msg_frame,
                                            height=8,
                                            wrap=tk.WORD,
                                            font=("微軟正黑體", 12),
                                            bg=ModernStyle.BG_LIGHT,
                                            fg=ModernStyle.TEXT_PRIMARY,
                                            relief="flat",
                                            borderwidth=0,
                                            state="normal")
            msg_text.pack(fill="both", expand=True)
            msg_text.insert("1.0", message)
            msg_text.config(state="disabled")
        else:
            # 普通標籤（短訊息）
            msg_label = tk.Label(main_frame, 
                                text=message,
                                font=("微軟正黑體", 12),
                                bg=ModernStyle.BG_LIGHT,
                                fg=ModernStyle.TEXT_PRIMARY,
                                justify="center",
                                wraplength=400)
            msg_label.pack(fill="x", expand=True, pady=15)
        
        # 🔧 關鍵修正：按鈕區域固定在底部
        button_frame = tk.Frame(main_frame, bg=ModernStyle.BG_LIGHT)
        button_frame.pack(side="bottom", fill="x", pady=(15, 0))
        
        result = [False]  # 用列表來存儲結果
        
        def on_yes():
            result[0] = True
            dialog.destroy()
        
        def on_no():
            result[0] = False
            dialog.destroy()
        
        # 🔧 修正：確保按鈕有足夠大小和間距
        if msg_type == "yesno":
            no_btn = ModernButton(button_frame, "否", command=on_no, style="secondary")
            no_btn.pack(side="right", padx=(5, 0))
            
            yes_btn = ModernButton(button_frame, "是", command=on_yes, style="success")
            yes_btn.pack(side="right", padx=(0, 5))
        else:
            ok_btn = ModernButton(button_frame, "確定", command=on_no, style="primary")
            ok_btn.pack(side="right")

        # 🔧 鍵盤操作：
        #    是/否對話框 → Y 或 Enter = 是，N 或 Esc = 否
        #    一般對話框   → Enter 或 Esc = 確定（關閉）
        if msg_type == "yesno":
            for k in ('<Return>', '<KP_Enter>', '<y>', '<Y>'):
                dialog.bind(k, lambda e: on_yes())
            for k in ('<n>', '<N>', '<Escape>'):
                dialog.bind(k, lambda e: on_no())
            try:
                yes_btn.focus_set()  # 預設焦點在「是」
            except Exception:
                pass
        else:
            for k in ('<Return>', '<KP_Enter>', '<Escape>'):
                dialog.bind(k, lambda e: on_no())
            try:
                ok_btn.focus_set()
            except Exception:
                pass

        # 🔧 新增：確保對話框大小適合內容
        dialog.update_idletasks()
        dialog.minsize(500, min_height)

        dialog.wait_window()
        return result[0]

    # ==================== 功能實現方法 ====================
    
    def load_pdf_list(self):
        """載入PDF檔案清單 - 新檔案系統版本"""
        print("🔍 調試：開始載入PDF清單", flush=True)
        
        # 🔧 如果還沒有建立檔案系統，就先跳過
        if not hasattr(self, 'file_tree') or self.file_tree is None:
            print("⚠️ 調試：file_tree 還沒建立，跳過載入", flush=True)
            return
        
        # 🔧 如果還沒有設定 current_folder，就使用預設值
        if not hasattr(self, 'current_folder') or self.current_folder is None:
            self.current_folder = PDF_DIR
            print(f"🔍 調試：設定預設資料夾: {self.current_folder}", flush=True)
        
        # 載入資料夾內容
        try:
            self.load_folder_contents()
            print("✅ 調試：PDF清單載入完成", flush=True)
        except Exception as e:
            print(f"❌ 調試：載入PDF清單失敗: {e}", flush=True)

    def select_folder(self):
        """選擇PDF資料夾"""
        global PDF_DIR
        folder = filedialog.askdirectory(title="選擇包含PDF檔案的資料夾")
        if folder:
            PDF_DIR = folder
            self.load_pdf_list()
            self.progress_indicator.set_status(f"切換到: {os.path.basename(PDF_DIR)}", "ready")
    
    def get_selected_pdfs(self):
        """取得選中的PDF檔案"""
        selected_files = []
        
        # 🔧 檢查 file_tree 選擇
        if hasattr(self, 'file_tree'):
            try:
                selection = self.file_tree.selection()
                for item in selection:
                    item_text = self.file_tree.item(item, 'text')
                    if item_text.startswith("📄"):
                        filename = item_text[2:]  # 移除📄前綴
                        file_path = os.path.join(self.current_folder, filename)
                        if os.path.exists(file_path):
                            selected_files.append(file_path)
            except Exception:
                pass
        
        # 🔧 檢查 current_selection 屬性
        if not selected_files and hasattr(self, 'current_selection'):
            if isinstance(self.current_selection, list):
                selected_files = self.current_selection
            elif isinstance(self.current_selection, str):
                selected_files = [self.current_selection]
        
        # 🔧 備用：檢查舊的 pdf_listbox
        if not selected_files and hasattr(self, 'pdf_listbox') and self.pdf_listbox:
            try:
                selected_indices = self.pdf_listbox.curselection()
                for index in selected_indices:
                    filename = self.pdf_listbox.get(index)
                    file_path = os.path.join(PDF_DIR, filename)
                    if os.path.exists(file_path):
                        selected_files.append(file_path)
            except Exception:
                pass
        
        return selected_files

    def smart_preview(self):
        """智慧預覽功能 - 新檔案系統版本"""
        try:
            # 🔧 使用新的檔案選擇方式
            pdf_paths = self.get_selected_pdfs()
            if not pdf_paths:
                self.show_custom_messagebox("提示", "請先選擇要預覽的PDF檔案")
                return
            
            # print("🔄 手動觸發重新預覽")
            
            # 🔧 直接呼叫原有的預覽方法
            self._perform_click_preview(pdf_paths)
            
        except Exception as e:
            # print(f"手動預覽錯誤: {e}")
            pass

    # 🔧 在這個區域加入新的按鈕狀態管理方法
    def on_file_selection_changed(self):
        """檔案選擇改變時的處理 - 更新按鈕狀態"""
        try:
            print("🔍 [測試] on_file_selection_changed 被呼叫了!")  # 新增這行測試
            # 🔧 檢查當前選中的檔案
            selected_files = self.get_currently_selected_files()
            
            if selected_files:
                # 🔧 檢查選中檔案的處理狀態
                self.update_processing_button_state(selected_files)
                print(f"🔍 檔案選擇改變，已更新按鈕狀態：{len(selected_files)}個檔案", flush=True)
            else:
                # 🔧 沒有選中檔案時，恢復預設狀態
                self.reset_processing_button_state()
                print("🔍 沒有選中檔案，按鈕已恢復預設狀態", flush=True)
                
        except Exception as e:
            print(f"❌ 更新檔案選擇狀態失敗：{e}", flush=True)

    def update_processing_button_state(self, selected_files):
        """根據選中檔案更新處理按鈕狀態"""
        try:
            # 🔧 修正：檢查正確的按鈕
            print(f"🔍 [調試] 檢查按鈕是否存在: {hasattr(self, 'tab_buttons') and 'struct' in self.tab_buttons}", flush=True)
            
            # 🔧 檢查所選檔案是否都已處理
            all_processed = True
            any_processed = False
            
            for file_path in selected_files:
                if file_path in self.processed_files:
                    any_processed = True
                else:
                    all_processed = False
            
            # 🔧 更新按鈕狀態和文字
            if hasattr(self, 'tab_buttons') and 'struct' in self.tab_buttons:
                if all_processed and len(selected_files) > 0:
                    # 所有選中檔案都已處理
                    self.tab_buttons['struct'].configure(text="📊 已結構化處理")
                    print("✅ 按鈕更新：已結構化處理", flush=True)
                elif any_processed:
                    # 部分檔案已處理
                    processed_count = sum(1 for f in selected_files if f in self.processed_files)
                    self.tab_buttons['struct'].configure(text=f"📊 部分已處理 ({processed_count}/{len(selected_files)})")
                    print(f"✅ 按鈕更新：部分已處理 {processed_count}/{len(selected_files)}", flush=True)
                else:
                    # 沒有檔案被處理過
                    self.tab_buttons['struct'].configure(text="📊 結構化資料")
                    print("✅ 按鈕更新：結構化資料", flush=True)
            else:
                print("❌ 找不到 tab_buttons['struct'] 按鈕", flush=True)
                    
        except Exception as e:
            print(f"❌ 更新處理按鈕狀態失敗：{e}", flush=True)

    def reset_processing_button_state(self):
        """重置處理按鈕到預設狀態"""
        try:
            if hasattr(self, 'tab_buttons') and 'struct' in self.tab_buttons:
                self.tab_buttons['struct'].configure(text="📊 結構化資料")
                print("✅ 按鈕已重置為預設狀態", flush=True)
            else:
                print("❌ 找不到 tab_buttons['struct'] 按鈕", flush=True)
        except Exception as e:
            print(f"❌ 重置按鈕狀態失敗：{e}", flush=True)

    def get_currently_selected_files(self):
        """取得當前選中的檔案列表"""
        try:
            selected_files = []
            
            if hasattr(self, 'file_tree') and self.file_tree:
                # 取得檔案樹中選中的項目
                selection = self.file_tree.selection()
                
                for item in selection:
                    # 取得檔案路徑
                    item_text = self.file_tree.item(item, 'text')
                    if item_text and not item_text.startswith('📁'):  # 排除資料夾
                        # 構建完整路徑
                        if hasattr(self, 'current_folder') and self.current_folder:
                            full_path = os.path.join(self.current_folder, item_text.replace('📄 ', ''))
                            selected_files.append(full_path)
            
            return selected_files
            
        except Exception as e:
            print(f"❌ 取得選中檔案失敗：{e}", flush=True)
            return []


    def show_preview_completion(self, message):
        """顯示預覽完成訊息（非阻塞式）"""
        # 創建一個簡短的提示而不是彈出對話框
        # 這樣用戶可以繼續查看預覽內容
        original_status = self.progress_indicator.progress_var.get()
        self.progress_indicator.set_status("✅ 預覽完成！點擊標籤頁切換內容", "success")
        
        # 3秒後恢復原狀態
        self.master.after(3000, lambda: self.progress_indicator.set_status(original_status, "success"))

    def add_preview_content(self, text_widget, title, content, max_lines=None):
        """添加預覽內容到文本框 - 改進版本"""
        text_widget.insert(tk.END, f"{'='*60}\n")
        text_widget.insert(tk.END, f"{title}\n")
        text_widget.insert(tk.END, f"{'='*60}\n")
        
        if max_lines and content:
            lines = content.split('\n')
            total_lines = len(lines)
            
            if total_lines > max_lines:
                # 顯示前面的行
                for line_num, line in enumerate(lines[:max_lines], 1):
                    text_widget.insert(tk.END, f"{line_num:04d} {line}\n")
                
                remaining_lines = total_lines - max_lines
                text_widget.insert(tk.END, f"\n🔽 省略了 {remaining_lines} 行內容\n")
                text_widget.insert(tk.END, f"💡 如需查看完整內容，請使用「📋 批次處理」功能\n")
            else:
                # 顯示所有行
                for line_num, line in enumerate(lines, 1):
                    text_widget.insert(tk.END, f"{line_num:04d} {line}\n")
        else:
            # 無行數限制或內容為空
            if content:
                text_widget.insert(tk.END, content)
            else:
                text_widget.insert(tk.END, "📭 內容為空")
        
        text_widget.insert(tk.END, "\n\n")

    
    def format_structured_data(self, data):
        """格式化結構化資料顯示"""
        if not data:
            return "未找到結構化資料"
        
        formatted = []
        for key, value in data.items():
            display_value = str(value)[:100] + "..." if len(str(value)) > 100 else str(value)
            formatted.append(f"{key:<20}: {display_value}")
        
        return "\n".join(formatted)
    
    def clear_info_cards(self):
        """清空資訊卡片"""
        for widget in self.info_scrollable_frame.winfo_children():
            widget.destroy()
        
        self.info_placeholder = tk.Label(self.info_scrollable_frame,
                                        text="📋 請選擇PDF檔案並點擊預覽以查看謄本資訊",
                                        font=ModernStyle.FONT_MAIN,
                                        bg=ModernStyle.BG_LIGHT,
                                        fg=ModernStyle.TEXT_SECONDARY)
        self.info_placeholder.pack(pady=50)
    
    def add_info_card(self, filename, transcript_info, basic_fields, raw_text, cleaned_text):
        """添加謄本資訊卡片 - 簡化版本"""
        # 移除佔位符
        if hasattr(self, 'info_placeholder'):
            self.info_placeholder.destroy()
        
        # 創建資訊卡片
        card = tk.Frame(self.info_scrollable_frame, 
                    bg=ModernStyle.BG_LIGHT,
                    relief="solid",
                    borderwidth=1,
                    padx=15,
                    pady=15)
        card.pack(fill="x", padx=10, pady=10)
        
        # 檔案標題
        title_frame = tk.Frame(card, bg=ModernStyle.BG_LIGHT)
        title_frame.pack(fill="x", pady=(0, 10))
        
        file_icon = tk.Label(title_frame, text="📄", 
                        font=("Segoe UI", 16),
                        bg=ModernStyle.BG_LIGHT)
        file_icon.pack(side="left")
        
        file_title = tk.Label(title_frame, text=filename,
                            font=ModernStyle.FONT_HEADER,
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.PRIMARY)
        file_title.pack(side="left", padx=(10, 0))
        
        # 謄本類型標籤
        transcript_type = f"{transcript_info['謄本類型']} - {transcript_info['謄本類別']}"
        type_color = ModernStyle.SUCCESS if "第一類" in transcript_type else ModernStyle.WARNING
        
        type_label = tk.Label(title_frame, text=transcript_type,
                            font=ModernStyle.FONT_MAIN,
                            bg=type_color,
                            fg=ModernStyle.TEXT_LIGHT,
                            padx=10, pady=3)
        type_label.pack(side="right")
        
        # 基本資訊網格
        info_grid = tk.Frame(card, bg=ModernStyle.BG_LIGHT)
        info_grid.pack(fill="x", pady=10)
        
        # 創建兩列佈局
        left_col = tk.Frame(info_grid, bg=ModernStyle.BG_LIGHT)
        left_col.pack(side="left", fill="both", expand=True)
        
        right_col = tk.Frame(info_grid, bg=ModernStyle.BG_LIGHT)
        right_col.pack(side="right", fill="both", expand=True)
        
        # 左列資訊
        info_items = [
            ("🏠 地號建號", transcript_info.get("地號建號", "未查到")),
            ("⏰ 列印時間", transcript_info.get("列印時間", "未查到")),
            ("👤 所有權人", basic_fields.get("所有權人", "未查到")),
            ("🏛️ 主要用途", basic_fields.get("主要用途", "未查到"))
        ]
        
        for i, (label, value) in enumerate(info_items):
            if i < 2:  # 前兩項放左列
                self.add_info_item(left_col, label, value)
            else:      # 後兩項放右列  
                self.add_info_item(right_col, label, value)
        
        # 統計資訊（簡化）
        stats_frame = tk.Frame(card, bg=ModernStyle.BG_SECONDARY, padx=10, pady=8)
        stats_frame.pack(fill="x", pady=(10, 0))
        
        documents = split_documents_by_title(raw_text)
        # 簡化統計資訊顯示
        stats_text = f"📊 統計: 文字 {len(raw_text.split())} 字 | 謄本數 {len(documents)} | 欄位數 {len(basic_fields)}"
        
        stats_label = tk.Label(stats_frame, text=stats_text,
                            font=("Consolas", 8),
                            bg=ModernStyle.BG_SECONDARY,
                            fg=ModernStyle.TEXT_SECONDARY)
        stats_label.pack()

    def add_info_item(self, parent, label, value):
        """添加資訊項目"""
        item_frame = tk.Frame(parent, bg=ModernStyle.BG_LIGHT)
        item_frame.pack(fill="x", pady=2)
        
        label_widget = tk.Label(item_frame, text=label,
                              font=ModernStyle.FONT_MAIN,
                              bg=ModernStyle.BG_LIGHT,
                              fg=ModernStyle.TEXT_SECONDARY,
                              width=12, anchor="w")
        label_widget.pack(side="left")
        
        # 限制值的顯示長度
        display_value = str(value)[:30] + "..." if len(str(value)) > 30 else str(value)
        
        value_widget = tk.Label(item_frame, text=display_value,
                              font=ModernStyle.FONT_MAIN,
                              bg=ModernStyle.BG_LIGHT,
                              fg=ModernStyle.TEXT_PRIMARY,
                              anchor="w")
        value_widget.pack(side="left", fill="x", expand=True)
    
    def detailed_analysis(self):
        """詳細分析功能"""
        pdf_paths = self.get_selected_pdfs()
        if not pdf_paths:
            return
        
        self.show_custom_messagebox("開發中", "詳細分析功能正在開發中，敬請期待！")
    
    def clear_preview(self):
        """清空預覽 - 簡化版本"""
        # 直接清空，不詢問
        for text_widget in [self.raw_text, self.struct_text]:
            text_widget.config(state="normal")
            text_widget.delete(1.0, tk.END)
            text_widget.config(state="disabled")
        
        self.clear_info_cards()

        # 🔧 新增：清空時重置標籤狀態
        self.update_tab_title(is_processed=False)

        # 清空記錄的選擇
        if hasattr(self, '_last_preview_selection'):
            delattr(self, '_last_preview_selection')
        
        self.progress_indicator.set_status("預覽已清空", "ready")
    
    def brief_query_single(self, pdf_path):
        """單檔簡略查詢"""
        try:
            raw_text = extract_pdf_text_only(pdf_path)
            if not raw_text:
                return None
            
            cleaned_text = clean_pagination_simple(raw_text)
            basic_fields = extract_basic_fields(cleaned_text)
            transcript_info = detect_transcript_info(raw_text)
            
            # 簡化結果結構
            result = {
                "謄本類型": f"{transcript_info['謄本類型']} - {transcript_info['謄本類別']}",
                "基本資訊": basic_fields,
                "源檔案": os.path.basename(pdf_path),
                "處理方式": "簡略查詢"
            }
            
            return result
            
        except Exception as e:
            print(f"簡略查詢 {pdf_path} 失敗: {e}", flush=True)
            return None
    


    def open_batch_scan_dialog(self):
        """跨資料夾掃描：遞迴掃描多個來源資料夾 → 篩選謄本 → 按案件資料夾分組批次處理。

        設計重點：
          - 篩選：抽 PDF 第一頁文字，找「土地/建物登記第一/二類謄本」關鍵字（毫秒級，準確）
          - 分組：按 PDF 所在資料夾（同資料夾 = 同案件）
          - 處理：每組各呼叫一次 self.processor.process_multiple_pdfs（引擎免改）
          - 輸出：每組各自 export 到 output 目錄（檔名自動依地號區分）
          - 安全：背景執行緒只做「處理 + 寫檔」，不碰 tkinter UI；進度用 queue + after 回主執行緒
        """
        import os, re, threading, queue
        from tkinter import filedialog
        from collections import defaultdict

        dlg = tk.Toplevel(self.master)
        dlg.title("🔍 跨資料夾批次掃描謄本")
        dlg.geometry("920x720")
        try:
            dlg.configure(bg=ModernStyle.BG_LIGHT)
        except Exception:
            pass
        bg = getattr(ModernStyle, 'BG_LIGHT', '#F5F5F5')
        font_n = ("Microsoft JhengHei", 11)
        font_b = ("Microsoft JhengHei", 11, "bold")

        state = {'sources': [], 'found': [], 'stop': False, 'scanning': False}

        # ===== 篩選函式（只抽第一頁，快又準）=====
        def check_transcript(pdf_path):
            """回傳 (是否合格謄本, 類型標籤如 '土地·一類')"""
            try:
                import fitz
                doc = fitz.open(pdf_path)
                txt = doc[0].get_text() if doc.page_count > 0 else ""
                doc.close()
                if not re.search(r'[土地建物]+登記第[一二]類謄本', txt):
                    return (False, "")
                kind = "建物" if "建物登記" in txt else "土地"
                cls = "二類" if "第二類" in txt else "一類"
                return (True, f"{kind}·{cls}")
            except Exception:
                return (False, "")

        def mk_btn(parent, text, cmd, color='#E67E22'):
            return tk.Button(parent, text=text, command=cmd, font=font_n,
                             bg=color, fg='white', relief="flat", padx=10, pady=4,
                             cursor="hand2", activebackground=color)

        # ===== 上方：來源資料夾清單 =====
        top = tk.Frame(dlg, bg=bg); top.pack(fill="x", padx=12, pady=(12, 6))
        tk.Label(top, text="掃描來源（可加多個資料夾，會自動遞迴往下掃所有層）：",
                 font=font_b, bg=bg, anchor="w").pack(fill="x")
        src_listbox = tk.Listbox(top, height=4, font=font_n)
        src_listbox.pack(fill="x", pady=4)

        def pick_multiple_folders():
            """自訂多選資料夾對話框（Ctrl / Shift 複選 + 完整磁碟機導覽）。"""
            import string
            picker = tk.Toplevel(dlg)
            picker.title("選擇資料夾（可 Ctrl / Shift 多選，雙擊進入下一層）")
            # 置中於螢幕
            _pw, _ph = 760, 580
            _sw = picker.winfo_screenwidth(); _sh = picker.winfo_screenheight()
            picker.geometry(f"{_pw}x{_ph}+{(_sw - _pw) // 2}+{(_sh - _ph) // 2}")
            try:
                picker.configure(bg=bg)
            except Exception:
                pass
            picker.transient(dlg)
            picker.grab_set()

            # 預設開在程式所在目錄（BASE_DIR）；若已加過來源則延續上一個
            start = state['sources'][-1] if state['sources'] else BASE_DIR
            cur = {'path': start}
            result = {'folders': []}
            folder_map = {}         # listbox index -> 完整路徑
            DRIVES = "::DRIVES::"   # 「本機」層（列磁碟機）
            NETWORK = "::NETWORK::" # 「網路」層（列網路電腦）

            def list_network_computers():
                """列出區網內的電腦。優先用 Windows Shell COM（跟檔案總管同一套探索，
                最可靠）；失敗再退而用 net view。"""
                comps = []
                # 方法1：Shell COM 列舉「網路」資料夾（ssfNETWORK = 18）
                try:
                    import win32com.client
                    shell = win32com.client.Dispatch("Shell.Application")
                    net = shell.NameSpace(18)
                    if net:
                        for item in net.Items():
                            try:
                                if item.IsFolder:  # 電腦是資料夾，印表機等不是
                                    nm = str(item.Name).strip()
                                    if nm:
                                        comps.append("\\\\" + nm)
                            except Exception:
                                continue
                except Exception:
                    pass
                # 方法2：net view（後備；新版 Windows 常因服務停用而列不到）
                if not comps:
                    import subprocess
                    try:
                        r = subprocess.run(
                            ["net", "view"],
                            capture_output=True, timeout=20,
                            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                        raw = r.stdout or b""
                        text = ""
                        for enc in ('cp950', 'utf-8', 'big5', 'cp437', 'mbcs'):
                            try:
                                text = raw.decode(enc); break
                            except Exception:
                                continue
                        for line in text.splitlines():
                            line = line.strip()
                            if line.startswith("\\\\"):
                                comps.append(line.split()[0])
                    except Exception as e:
                        print(f"net view 失敗: {e}", flush=True)
                # 去重（保序）
                seen2, uniq = set(), []
                for c in comps:
                    if c not in seen2:
                        seen2.add(c); uniq.append(c)
                return uniq

            def _is_unc_computer_root(p):
                """判斷是不是「電腦根」UNC（如 \\TABLE1，只有電腦名沒有共用名）。
                這種路徑 os.listdir 會 WinError 67，必須改用 Shell COM 列共用。"""
                if not p.startswith("\\\\"):
                    return False
                parts = [x for x in p[2:].split("\\") if x]
                return len(parts) == 1

            def list_shell_folders(p):
                """用 Shell COM 列出某路徑下的子資料夾（共用），回傳 [(名稱, 完整路徑)]。
                專用於 UNC 電腦根——os.listdir 對它會失敗。"""
                items = []
                try:
                    import win32com.client
                    shell = win32com.client.Dispatch("Shell.Application")
                    ns = shell.NameSpace(p)
                    if ns:
                        for it in ns.Items():
                            try:
                                if it.IsFolder:
                                    nm = str(it.Name).strip()
                                    if nm:
                                        items.append((nm, p.rstrip("\\") + "\\" + nm))
                            except Exception:
                                continue
                except Exception:
                    pass
                return items

            def list_net_view_shares(p):
                """後備：用 net view \\電腦 列共用資料夾名。"""
                items = []
                import subprocess
                try:
                    r = subprocess.run(
                        ["net", "view", p],
                        capture_output=True, timeout=20,
                        creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                    raw = r.stdout or b""
                    text = ""
                    for enc in ('cp950', 'utf-8', 'big5', 'cp437', 'mbcs'):
                        try:
                            text = raw.decode(enc); break
                        except Exception:
                            continue
                    for line in text.splitlines():
                        # 含「Disk / 磁碟」類型的行，第一欄即共用名
                        if ('Disk' in line) or ('磁碟' in line):
                            parts = line.split()
                            if parts:
                                nm = parts[0]
                                items.append((nm, p.rstrip("\\") + "\\" + nm))
                except Exception:
                    pass
                return items

            # 路徑列
            pf = tk.Frame(picker, bg=bg); pf.pack(fill="x", padx=10, pady=8)
            path_var = tk.StringVar(value=cur['path'])

            def go_up():
                p = cur['path']
                if p in (DRIVES, NETWORK):
                    return
                stripped = p.rstrip("\\/")
                # UNC 網路路徑 \\電腦\共用\...
                if stripped.startswith("\\\\"):
                    parts = [x for x in stripped[2:].split("\\") if x]
                    if len(parts) <= 1:
                        cur['path'] = NETWORK  # \\電腦 → 回網路層
                    else:
                        cur['path'] = "\\\\" + "\\".join(parts[:-1])
                    refresh()
                    return
                parent = os.path.dirname(stripped)
                # 已在磁碟機根（如 C:\）→ 再上一層到「本機」
                if parent == stripped or parent == "" or (len(stripped) == 2 and stripped[1] == ":"):
                    cur['path'] = DRIVES
                else:
                    cur['path'] = parent
                refresh()

            def go_to(p):
                if p in (DRIVES, NETWORK):
                    cur['path'] = p
                    refresh()
                    return
                # UNC 路徑
                if p.startswith("\\\\"):
                    # 電腦根（\\TABLE1）：os.listdir 會 WinError 67，直接交給
                    # refresh 用 Shell COM 列共用資料夾
                    if _is_unc_computer_root(p):
                        cur['path'] = p
                        refresh()
                        return
                    # 更深層（\\電腦\共用\...）：用 listdir 驗證可存取
                    try:
                        os.listdir(p)
                        cur['path'] = p
                        refresh()
                    except Exception:
                        messagebox.showinfo("提示",
                                            f"無法存取：\n{p}\n\n"
                                            "請確認：電腦名正確、網路相通、有存取權限。",
                                            parent=picker)
                    return
                if os.path.isdir(p):
                    cur['path'] = p
                    refresh()
                else:
                    messagebox.showinfo("提示", "路徑不存在", parent=picker)

            def add_current():
                if cur['path'] in (DRIVES, NETWORK):
                    messagebox.showinfo("提示", "請先進入一個磁碟機、資料夾或網路共用", parent=picker)
                    return
                result['folders'] = [cur['path']]
                picker.destroy()

            mk_btn(pf, "⬆ 上一層", go_up, '#95A5A6').pack(side="left")
            path_entry = tk.Entry(pf, textvariable=path_var, font=font_n)
            path_entry.pack(side="left", fill="x", expand=True, padx=6)
            mk_btn(pf, "前往", lambda: go_to(path_var.get()), '#3498DB').pack(side="left")
            # 🔥 「加入此資料夾」移到路徑列、前往右邊，改橘色
            mk_btn(pf, "✔ 加入此資料夾", add_current, '#E67E22').pack(side="left", padx=4)

            # ===== 常用位置快捷列（含雲端硬碟、網路）=====
            def detect_quick_locations():
                import json as _json
                home = os.path.expanduser("~")
                locs = []
                for nm, sub in [("🖥 桌面", "Desktop"), ("📄 文件", "Documents"), ("📥 下載", "Downloads")]:
                    p = os.path.join(home, sub)
                    if os.path.isdir(p):
                        locs.append((nm, p))
                # === Dropbox（支援多帳號）===
                seen = set()
                # 1. 讀 Dropbox info.json（最準，含所有登入帳號的路徑）
                for base in [os.environ.get("APPDATA"), os.environ.get("LOCALAPPDATA")]:
                    if not base:
                        continue
                    info = os.path.join(base, "Dropbox", "info.json")
                    if os.path.isfile(info):
                        try:
                            with open(info, encoding='utf-8') as f:
                                data = _json.load(f)
                            for acc_type, acc in data.items():  # personal / business
                                p = acc.get("path")
                                if p and os.path.isdir(p) and p not in seen:
                                    seen.add(p)
                                    locs.append((f"☁ Dropbox-{acc_type}", p))
                        except Exception:
                            pass
                # 2. 後備：掃描家目錄下所有 Dropbox* 資料夾
                try:
                    for name in sorted(os.listdir(home)):
                        full = os.path.join(home, name)
                        if name.lower().startswith("dropbox") and os.path.isdir(full) and full not in seen:
                            seen.add(full)
                            locs.append((f"☁ {name}", full))
                except Exception:
                    pass
                # 3. 後備：掃描各磁碟機根的 Dropbox（例如 D:\Dropbox，另一個帳號常放這）
                for d in string.ascii_uppercase:
                    cand = f"{d}:\\Dropbox"
                    if os.path.isdir(cand) and cand not in seen:
                        seen.add(cand)
                        locs.append((f"☁ Dropbox({d}:)", cand))
                # === OneDrive（支援個人 + 公司）===
                try:
                    for name in sorted(os.listdir(home)):
                        full = os.path.join(home, name)
                        if name.lower().startswith("onedrive") and os.path.isdir(full):
                            locs.append((f"☁ {name}", full))
                except Exception:
                    pass
                # === Google Drive（虛擬碟 X:\My Drive 或同步資料夾）===
                gd = None
                for d in string.ascii_uppercase:
                    for sub in ["My Drive", "我的雲端硬碟"]:
                        c = f"{d}:\\{sub}"
                        if os.path.isdir(c):
                            gd = c; break
                    if gd:
                        break
                if not gd:
                    for cand in [os.path.join(home, "Google Drive"), os.path.join(home, "My Drive")]:
                        if os.path.isdir(cand):
                            gd = cand; break
                if gd:
                    locs.append(("☁ Google雲端", gd))
                return locs

            # 常用列：放在最上面（路徑列之前），grid 自動換行，按鈕之間不留空白
            qf = tk.Frame(picker, bg=bg); qf.pack(fill="x", padx=10, pady=(8, 4), before=pf)
            _quick = [("💻 本機", DRIVES)] + detect_quick_locations() + [("🌐 網路", NETWORK)]
            _col, _row = 0, 0
            for _nm, _p in _quick:
                mk_btn(qf, _nm, (lambda pp=_p: go_to(pp)), '#7F8C8D').grid(
                    row=_row, column=_col, padx=2, pady=2, sticky="w")
                _col += 1
                if _col > 6:   # 每行最多 7 顆快捷鈕，超過換行
                    _col, _row = 0, _row + 1

            # 資料夾清單（多選）
            lf = tk.Frame(picker, bg=bg); lf.pack(fill="both", expand=True, padx=10)
            sb = tk.Scrollbar(lf); sb.pack(side="right", fill="y")
            folder_lb = tk.Listbox(lf, selectmode="extended", font=font_n, yscrollcommand=sb.set)
            folder_lb.pack(side="left", fill="both", expand=True)
            sb.config(command=folder_lb.yview)

            def refresh():
                folder_lb.delete(0, "end")
                folder_map.clear()
                # 「本機」層：列出所有磁碟機
                if cur['path'] == DRIVES:
                    path_var.set("💻 本機（所有磁碟機）")
                    i = 0
                    for d in string.ascii_uppercase:
                        root = f"{d}:\\"
                        if os.path.exists(root):
                            folder_lb.insert("end", f"💽 {root}")
                            folder_map[i] = root
                            i += 1
                    return
                # 「網路」層：列出區網電腦
                if cur['path'] == NETWORK:
                    path_var.set("🌐 網路（雙擊電腦進入看共用資料夾）")
                    folder_lb.insert("end", "（正在搜尋網路電腦，請稍候 5~20 秒...）")
                    picker.update_idletasks()
                    comps = list_network_computers()
                    folder_lb.delete(0, "end")
                    folder_map.clear()
                    for i2, c in enumerate(comps):
                        folder_lb.insert("end", f"💻 {c}")
                        folder_map[i2] = c
                    if not comps:
                        folder_lb.insert("end", "（找不到網路電腦；可在路徑列直接輸入 \\\\電腦名 後按前往）")
                    return
                path_var.set(cur['path'])
                # UNC 電腦根（\\TABLE1）：os.listdir 會 WinError 67，改用 Shell COM 列共用
                if _is_unc_computer_root(cur['path']):
                    folder_lb.insert("end", "（正在讀取共用資料夾...）")
                    picker.update_idletasks()
                    shares = list_shell_folders(cur['path'])
                    if not shares:
                        shares = list_net_view_shares(cur['path'])  # 後備
                    folder_lb.delete(0, "end")
                    folder_map.clear()
                    i = 0
                    for nm, full in shares:
                        folder_lb.insert("end", f"📁 {nm}")
                        folder_map[i] = full
                        i += 1
                    if i == 0:
                        folder_lb.insert("end",
                                         "（無法列出共用資料夾，可在路徑列輸入 \\\\電腦\\共用名 後按前往）")
                    return
                try:
                    names = sorted(os.listdir(cur['path']))
                    i = 0
                    for name in names:
                        full = os.path.join(cur['path'], name)
                        if os.path.isdir(full):
                            folder_lb.insert("end", f"📁 {name}")
                            folder_map[i] = full
                            i += 1
                    if i == 0:
                        folder_lb.insert("end", "（此資料夾沒有子資料夾，可直接按「加入此資料夾」）")
                except Exception as e:
                    messagebox.showerror("錯誤", f"無法讀取：{e}", parent=picker)

            def enter_folder(_e):
                sel = folder_lb.curselection()
                if len(sel) == 1 and sel[0] in folder_map:
                    cur['path'] = folder_map[sel[0]]
                    refresh()
            folder_lb.bind("<Double-Button-1>", enter_folder)

            def confirm_selected():
                sel = folder_lb.curselection()
                picked = [folder_map[i] for i in sel if i in folder_map]
                if not picked:
                    messagebox.showinfo("提示", "請先選取一或多個資料夾（可 Ctrl / Shift 複選）", parent=picker)
                    return
                result['folders'] = picked
                picker.destroy()

            bf = tk.Frame(picker, bg=bg); bf.pack(fill="x", padx=10, pady=10)
            mk_btn(bf, "✔ 加入選取的資料夾", confirm_selected, '#27AE60').pack(side="left")
            mk_btn(bf, "取消", picker.destroy, '#95A5A6').pack(side="right")

            refresh()
            picker.wait_window()
            return result['folders']

        def add_folder():
            for d in pick_multiple_folders():
                if d and d not in state['sources']:
                    state['sources'].append(d)
                    src_listbox.insert("end", d)

        def remove_folder():
            for i in reversed(list(src_listbox.curselection())):
                src_listbox.delete(i)
                del state['sources'][i]

        def clear_folders():
            src_listbox.delete(0, "end")
            state['sources'].clear()

        brow = tk.Frame(top, bg=bg); brow.pack(fill="x")
        mk_btn(brow, "+ 新增資料夾", add_folder).pack(side="left", padx=2)
        mk_btn(brow, "移除選取", remove_folder, '#95A5A6').pack(side="left", padx=2)
        mk_btn(brow, "清空", clear_folders, '#95A5A6').pack(side="left", padx=2)

        # ===== 掃描按鈕 + 進度/狀態 =====
        srow = tk.Frame(dlg, bg=bg); srow.pack(fill="x", padx=12, pady=6)
        scan_btn = mk_btn(srow, "🔍 開始掃描", lambda: do_scan(), '#27AE60')
        scan_btn.pack(side="left")
        stop_btn = mk_btn(srow, "⏹ 停止", lambda: state.update(stop=True), '#E74C3C')
        stop_btn.pack(side="left", padx=4)
        stop_btn.config(state="disabled")  # 初始禁用，掃描時才啟用
        scan_status = tk.Label(srow, text="尚未掃描", font=font_n, bg=bg, fg='#555')
        scan_status.pack(side="left", padx=10)

        # ===== 結果勾選清單 =====
        mid = tk.Frame(dlg, bg=bg); mid.pack(fill="both", expand=True, padx=12, pady=6)
        tk.Label(mid, text="掃描結果（合格謄本，可用 Ctrl / Shift 複選）：",
                 font=font_b, bg=bg, anchor="w").pack(fill="x")
        lframe = tk.Frame(mid); lframe.pack(fill="both", expand=True)
        scroll = tk.Scrollbar(lframe); scroll.pack(side="right", fill="y")
        result_listbox = tk.Listbox(lframe, selectmode="extended",
                                    font=("Microsoft JhengHei", 10),
                                    yscrollcommand=scroll.set)
        result_listbox.pack(side="left", fill="both", expand=True)
        scroll.config(command=result_listbox.yview)

        selrow = tk.Frame(mid, bg=bg); selrow.pack(fill="x", pady=3)
        mk_btn(selrow, "全選", lambda: result_listbox.select_set(0, "end"), '#3498DB').pack(side="left", padx=2)
        mk_btn(selrow, "全不選", lambda: result_listbox.select_clear(0, "end"), '#3498DB').pack(side="left", padx=2)

        def sel_invert():
            cur = set(result_listbox.curselection())
            for i in range(result_listbox.size()):
                (result_listbox.select_clear if i in cur else result_listbox.select_set)(i)
        mk_btn(selrow, "反選", sel_invert, '#3498DB').pack(side="left", padx=2)

        # ===== 輸出選項 =====
        orow = tk.Frame(dlg, bg=bg); orow.pack(fill="x", padx=12, pady=4)
        output_mode = tk.StringVar(value="merge")
        tk.Radiobutton(orow, text="全部合併成一份總表（放 output 目錄）",
                       variable=output_mode, value="merge", font=font_n, bg=bg).pack(anchor="w")
        tk.Radiobutton(orow, text="按案件分組：各案件一份，集中放 output 目錄",
                       variable=output_mode, value="group_central", font=font_n, bg=bg).pack(anchor="w")
        tk.Radiobutton(orow, text="按案件分組：各案件一份，放回各案件PDF所在的資料夾",
                       variable=output_mode, value="group_inplace", font=font_n, bg=bg).pack(anchor="w")
        # 輸出格式複選（對「分組」和「合併」都適用）
        fmt_row = tk.Frame(orow, bg=bg); fmt_row.pack(fill="x", pady=(4, 0))
        tk.Label(fmt_row, text="輸出格式：", font=font_n, bg=bg).pack(side="left")
        fmt_txt = tk.BooleanVar(value=False)
        fmt_json = tk.BooleanVar(value=False)
        fmt_xlsx = tk.BooleanVar(value=True)
        tk.Checkbutton(fmt_row, text="TXT", variable=fmt_txt, font=font_n, bg=bg).pack(side="left", padx=6)
        tk.Checkbutton(fmt_row, text="JSON", variable=fmt_json, font=font_n, bg=bg).pack(side="left", padx=6)
        tk.Checkbutton(fmt_row, text="XLSX", variable=fmt_xlsx, font=font_n, bg=bg).pack(side="left", padx=6)

        # ===== 開始處理 =====
        proc_btn = mk_btn(dlg, "▶ 開始處理選取的謄本", lambda: do_process(), '#E74C3C')
        proc_btn.pack(fill="x", padx=12, pady=(6, 12))

        # ===== 掃描邏輯（背景執行緒，即時列出 + 可停止）=====
        def _disp_path(full):
            for s in state['sources']:
                if full.startswith(s):
                    return os.path.relpath(full, os.path.dirname(s))
            return full

        def do_scan():
            if state['scanning']:
                return
            if not state['sources']:
                messagebox.showinfo("提示", "請先新增至少一個掃描來源資料夾", parent=dlg)
                return
            state['stop'] = False
            state['scanning'] = True
            state['found'] = []
            result_listbox.delete(0, "end")
            scan_btn.config(state="disabled")
            stop_btn.config(state="normal")
            sq = queue.Queue()

            def worker():
                total = 0
                for root_folder in list(state['sources']):
                    if state['stop']:
                        break
                    for dirpath, _dirs, files in os.walk(root_folder):
                        if state['stop']:
                            break
                        for fn in files:
                            if state['stop']:
                                break
                            if fn.lower().endswith('.pdf'):
                                total += 1
                                full = os.path.join(dirpath, fn)
                                ok, label = check_transcript(full)
                                if ok:
                                    # 🔥 找到就即時回報（邊掃邊列）
                                    sq.put(('found', full, label, total))
                                elif total % 5 == 0:
                                    sq.put(('scanned', total))
                sq.put(('done', total, state['stop']))

            threading.Thread(target=worker, daemon=True).start()

            def poll():
                try:
                    while True:
                        msg = sq.get_nowait()
                        if msg[0] == 'found':
                            full, label, total = msg[1], msg[2], msg[3]
                            state['found'].append((full, label))
                            result_listbox.insert("end", f"[{label}] {_disp_path(full)}")
                            result_listbox.select_set("end")  # 即時加入並預選
                            scan_status.config(text=f"掃描中... 已掃 {total}，找到 {len(state['found'])} 個謄本")
                        elif msg[0] == 'scanned':
                            scan_status.config(text=f"掃描中... 已掃 {msg[1]}，找到 {len(state['found'])} 個謄本")
                        elif msg[0] == 'done':
                            total, stopped = msg[1], msg[2]
                            state['scanning'] = False
                            skipped = total - len(state['found'])
                            word = "已停止" if stopped else "完成"
                            scan_status.config(text=f"{word}：掃描 {total} 個 PDF，找到 {len(state['found'])} 個謄本（已篩掉 {skipped} 個）")
                            scan_btn.config(state="normal")
                            stop_btn.config(state="disabled")
                            return
                except queue.Empty:
                    pass
                dlg.after(100, poll)
            dlg.after(100, poll)

        # ===== 處理邏輯（背景執行緒只處理+寫檔）=====
        def do_process():
            sel = result_listbox.curselection()
            if not sel:
                messagebox.showinfo("提示", "請先勾選要處理的謄本", parent=dlg)
                return
            if not (hasattr(self, 'processor') and hasattr(self, 'exporter')):
                messagebox.showerror("錯誤", "處理引擎未就緒", parent=dlg)
                return
            # 依勾選決定輸出格式
            want_fmts = []
            if fmt_txt.get():
                want_fmts.append('txt')
            if fmt_json.get():
                want_fmts.append('json')
            if fmt_xlsx.get():
                want_fmts.append('xlsx')
            if not want_fmts:
                messagebox.showinfo("提示", "請至少勾選一種輸出格式（TXT / JSON / XLSX）", parent=dlg)
                return
            chosen = [state['found'][i] for i in sel]
            groups = defaultdict(list)
            for full, _label in chosen:
                groups[os.path.dirname(full)].append(full)
            merge = (output_mode.get() == "merge")
            _mode_desc = {
                "merge": "合併成一份總表（放 output 目錄）",
                "group_central": "按案件分組，集中放 output 目錄",
                "group_inplace": "按案件分組，放回各案件原資料夾",
            }.get(output_mode.get(), "按案件分組")
            if not self.show_custom_messagebox(
                "確認處理",
                f"將處理 {len(chosen)} 個謄本，分屬 {len(groups)} 個案件資料夾\n"
                f"輸出方式：{_mode_desc}\n"
                f"輸出格式：{'、'.join(f.upper() for f in want_fmts)}\n\n確定開始？",
                "yesno"):
                return
            proc_btn.config(state="disabled")
            scan_btn.config(state="disabled")

            # 🔥 重用原本程式的完整進度視窗（執行時間計時器、6 階段、整體/單檔進度、取消、暫停、處理記錄）
            all_pdfs = [p for pdfs in groups.values() for p in pdfs]
            try:
                total_size = sum(os.path.getsize(p) / (1024 * 1024) for p in all_pdfs)
            except Exception:
                total_size = 0
            progress_window = self.create_enhanced_batch_progress_window(all_pdfs, total_size)
            rq = queue.Queue()

            def export_results(results, out_dir=None):
                # out_dir 不為 None 時，暫時把輸出目錄指到該案件資料夾（放回原處）
                paths = []
                old_dir = self.exporter.output_dir
                if out_dir:
                    try:
                        os.makedirs(out_dir, exist_ok=True)
                        self.exporter.output_dir = out_dir
                    except Exception as e:
                        print(f"[批次掃描] 切換輸出目錄失敗: {e}", flush=True)
                try:
                    for ext in want_fmts:  # 只輸出使用者勾選的格式
                        try:
                            fn = self.exporter._create_flexible_filename(results, ext)
                            if ext == 'txt':
                                p = self.exporter.export_to_txt_report(results, fn)
                            elif ext == 'xlsx':
                                p = self.exporter.export_to_excel(results, fn)
                            else:
                                p = self.exporter.export_to_json(results, fn)
                            if p:
                                paths.append(p)
                        except Exception as e:
                            print(f"[批次掃描] 匯出 {ext} 失敗: {e}", flush=True)
                finally:
                    self.exporter.output_dir = old_dir  # 還原
                return paths

            def worker():
                import time as _t
                try:
                    out_files, idx = [], 0

                    def process_group(pdfs):
                        nonlocal idx
                        gres = []
                        for pdf in pdfs:
                            if progress_window.cancelled:
                                return None
                            while progress_window.paused and not progress_window.cancelled:
                                _t.sleep(0.3)
                            if progress_window.cancelled:
                                return None
                            idx += 1
                            fn = os.path.basename(pdf)
                            # 用原本進度視窗回報階段（讓 6 階段圖示、整體/單檔進度條都會動）
                            progress_window.update_progress(idx, fn, "📄 提取文字內容", 2, "讀取 PDF...")
                            progress_window.update_progress(idx, fn, "🔍 OCR識別處理", 4, "辨識中...")
                            progress_window.update_progress(idx, fn, "📊 解析謄本結構", 5, "解析中...")
                            r = self.processor.process_multiple_pdfs([pdf])
                            progress_window.update_progress(idx, fn, "✨ 完善資料", 6, "完成")
                            if r:
                                gres.extend(r)
                                progress_window.add_history(f"✅ {fn} - 提取 {len(r)} 份謄本")
                            else:
                                progress_window.add_history(f"⚠️ {fn} - 未找到有效謄本")
                        return gres

                    if merge:
                        all_results = []
                        for folder, pdfs in groups.items():
                            gr = process_group(pdfs)
                            if gr is None:
                                break
                            all_results.extend(gr)
                        if all_results and not progress_window.cancelled:
                            out_files = export_results(all_results)
                            progress_window.add_history(f"📦 合併總表 → 輸出 {len(out_files)} 個檔案")
                    else:
                        inplace = (output_mode.get() == "group_inplace")
                        for folder, pdfs in groups.items():
                            gr = process_group(pdfs)
                            if gr is None:
                                break
                            if gr:
                                paths = export_results(gr, out_dir=folder if inplace else None)
                                out_files += paths
                                where = "原資料夾" if inplace else "output"
                                progress_window.add_history(
                                    f"📦 {os.path.basename(folder)} → {where}，輸出 {len(paths)} 個檔案")

                    if progress_window.cancelled:
                        rq.put(('cancelled', out_files))
                    else:
                        rq.put(('done', out_files))
                except Exception as e:
                    import traceback
                    traceback.print_exc()
                    rq.put(('error', str(e)))

            threading.Thread(target=worker, daemon=True).start()

            def check():
                try:
                    status, data = rq.get_nowait()
                    if status == 'done':
                        progress_window.show_completion(True, "", len(data))

                        def after_done():
                            try:
                                progress_window.close()
                            except Exception:
                                pass
                            proc_btn.config(state="normal")
                            scan_btn.config(state="normal")
                            outdir = os.path.dirname(data[0]) if data else getattr(self, 'output_directory', '(無輸出)')
                            if self.show_custom_messagebox(
                                "處理完成",
                                f"完成！共輸出 {len(data)} 個檔案\n\n儲存目錄：\n{outdir}\n\n是否開啟輸出資料夾？",
                                "yesno"):
                                try:
                                    os.startfile(outdir)
                                except Exception as _e:
                                    print(f"開啟資料夾失敗: {_e}", flush=True)
                        self.master.after(800, after_done)
                    elif status == 'cancelled':
                        progress_window.show_completion(False, "已取消", len(data))
                        proc_btn.config(state="normal")
                        scan_btn.config(state="normal")
                    elif status == 'error':
                        progress_window.show_completion(False, data, 0)
                        proc_btn.config(state="normal")
                        scan_btn.config(state="normal")
                        self.show_custom_messagebox("處理失敗", f"處理時發生錯誤：\n{data}", "error")
                    return
                except queue.Empty:
                    pass
                self.master.after(200, check)
            self.master.after(200, check)

    def batch_process(self):
        """批次處理 - 使用詳細進度視窗"""
        pdf_paths = self.get_selected_pdfs()
        if not pdf_paths:
            return
        
        # 🔧 新增：處理開始時更新狀態
        self.update_tab_title(is_processed=False)  # 先重置
        
        # 計算總檔案大小
        try:
            total_size = sum(os.path.getsize(path) / (1024 * 1024) for path in pdf_paths)
        except:
            total_size = 0
        
        result = self.show_custom_messagebox("確認批次處理", 
                                        f"即將批次處理 {len(pdf_paths)} 個PDF檔案\n"
                                        f"總大小: {total_size:.1f} MB\n"
                                        f"預估時間: {self.estimate_batch_time(total_size)}\n\n"
                                        f"⚠️ 目前僅支持電子謄本\n"
                                        f"💡 處理過程中可以暫停或取消\n\n"
                                        f"確定要繼續嗎？", "yesno")
        if not result:
            return
        
        self.progress_indicator.set_status(f"批次處理 {len(pdf_paths)} 個檔案...", "processing")
        
        # 🔧 使用新的詳細進度視窗
        progress_window = self.create_enhanced_batch_progress_window(pdf_paths, total_size)
        self.start_batch_processing_thread(pdf_paths, progress_window)

    def start_batch_processing_thread(self, pdf_paths, progress_window):
        """啟動詳細階段報告的批次處理執行緒 - 修正版本"""
        
        # 創建結果佇列
        result_queue = queue.Queue()
        
        def detailed_batch_processing_thread():
            """詳細階段報告的批次處理執行緒"""
            try:
                results = []
                
                for i, pdf_path in enumerate(pdf_paths, 1):
                    # 🔧 檢查是否被取消
                    if progress_window.cancelled:
                        print("批次處理被用戶取消", flush=True)
                        break
                    
                    # 🔧 檢查暫停狀態
                    while progress_window.paused and not progress_window.cancelled:
                        import time
                        time.sleep(0.5)
                    
                    if progress_window.cancelled:
                        break
                    
                    filename = os.path.basename(pdf_path)
                    print(f"開始處理第 {i} 個檔案: {filename}", flush=True)
                    
                    try:
                        # 🔧 階段1：載入PDF檔案
                        progress_window.update_progress(i, filename, "🔄 載入PDF檔案", 1, "正在讀取檔案...")
                        import time
                        time.sleep(0.2)  # 模擬處理時間
                        
                        # 🔧 階段2：提取文字內容
                        progress_window.update_progress(i, filename, "📄 提取文字內容", 2, "正在提取PDF文字...")
                        raw_text = extract_pdf_text_only(pdf_path)
                        if not raw_text:
                            progress_window.add_history(f"警告: {filename} - 無法提取文字內容")
                            continue
                        
                        # 檢查暫停
                        while progress_window.paused and not progress_window.cancelled:
                            time.sleep(0.5)
                        if progress_window.cancelled:
                            break
                        
                        # 🔧 階段3：清理跨頁資料
                        progress_window.update_progress(i, filename, "🧹 清理跨頁資料", 3, "正在處理頁面分割...")
                        time.sleep(0.1)
                        
                        # 🔧 階段4：OCR識別處理
                        progress_window.update_progress(i, filename, "🔍 OCR識別處理", 4, "正在進行文字識別...")
                        time.sleep(0.3)  # OCR通常比較耗時
                        
                        # 檢查暫停
                        while progress_window.paused and not progress_window.cancelled:
                            time.sleep(0.5)
                        if progress_window.cancelled:
                            break
                        
                        # 🔧 階段5：解析謄本結構
                        progress_window.update_progress(i, filename, "📊 解析謄本結構", 5, "正在分析謄本格式...")
                        
                        # 實際處理：呼叫原有的處理器
                        single_results = self.processor.process_multiple_pdfs([pdf_path])
                        
                        # 🔧 階段6：完善資料
                        progress_window.update_progress(i, filename, "✨ 完善資料", 6, "正在完善解析結果...")
                        time.sleep(0.2)
                        
                        if single_results:
                            results.extend(single_results)
                            progress_window.add_history(f"✅ 成功: {filename} - 提取 {len(single_results)} 份謄本")
                        else:
                            progress_window.add_history(f"⚠️ 警告: {filename} - 未找到有效謄本")
                        
                    except Exception as e:
                        progress_window.add_history(f"❌ 錯誤: {filename} - {str(e)}")
                        print(f"處理 {pdf_path} 失敗: {e}", flush=True)
                        continue
                
                # 將結果放入佇列
                if progress_window.cancelled:
                    result_queue.put(('cancelled', len(results)))
                else:
                    result_queue.put(('success', results))
                    
            except Exception as e:
                # 將錯誤放入佇列
                result_queue.put(('error', str(e)))
        
        def check_detailed_batch_result():
            """檢查詳細批次處理結果"""
            try:
                # 非阻塞檢查佇列
                status, data = result_queue.get_nowait()
                
                if status == 'success':
                    progress_window.show_completion(True, "", len(data))
                    
                    # 延遲處理結果顯示
                    def show_results():
                        progress_window.close()
                        if data:
                            self.display_results(data)
                            self.current_results = data
                            # 🔧 新增：批次處理完成後更新標籤狀態
                            self.update_tab_title(is_processed=True)                            
                            success_count = len(data)
                            self.progress_indicator.set_status(f"批次處理完成 - {success_count} 份", "success")
                            if hasattr(self, 'stats_label'):
                                self.stats_label.config(text=f"已處理 {success_count} 份謄本 (批次模式)")
                            else:
                                print(f"📊 統計：已處理 {success_count} 份謄本 (批次模式)", flush=True)
                            # 🔧 新增：調用批次完成處理
                            self.on_batch_process_complete(len(pdf_paths), success_count)
                        
                        else:
                            self.show_custom_messagebox("警告", "批次處理完成，但結果為空")
                    
                    # 當進度視窗關閉時顯示結果
                    self.master.after(1000, show_results)
                    
                elif status == 'cancelled':
                    partial_count = data
                    progress_window.show_completion(False, "已取消", partial_count)
                    self.progress_indicator.set_status("批次處理已取消", "ready")
                        
                elif status == 'error':
                    progress_window.show_completion(False, data, 0)
                    
                    def show_error():
                        self.show_custom_messagebox("錯誤", f"批次處理失敗:\n{data}")
                        self.progress_indicator.set_status("批次處理失敗", "error")
                    
                    self.master.after(2000, show_error)
                    
            except queue.Empty:
                # 佇列為空，繼續檢查
                if not progress_window.cancelled:
                    self.master.after(500, check_detailed_batch_result)
                else:
                    # 處理被取消，稍後再檢查一次結果
                    self.master.after(500, check_detailed_batch_result)
        
        # 啟動處理執行緒
        import threading
        thread = threading.Thread(target=detailed_batch_processing_thread, daemon=True)
        thread.start()
        
        # 開始檢查結果
        self.master.after(500, check_detailed_batch_result)

    # 🔧 確保使用正確的進度視窗類別
    def create_enhanced_batch_progress_window(self, pdf_paths, total_size):
        """創建修正版的詳細批次進度視窗"""
        
        class DetailedBatchProgressWindow:
            def __init__(self, parent, pdf_paths, total_size):
                self.parent = parent
                self.pdf_paths = pdf_paths
                self.total_size = total_size
                self.cancelled = False
                self.paused = False
                self.current_file_index = 0
                self.current_stage = ""
                self.total_stages = 6  # 每個檔案的處理階段數
                
                # 🔧 新增：記錄開始時間
                import time
                self.start_time = time.time()
                
                # 創建視窗
                self.window = tk.Toplevel(parent)
                self.window.title("🔄 批次處理進度")
                # 🔧 修正1：增加視窗高度，確保所有元素可見
                self.window.geometry("750x650")  # 從550增加到650
                self.window.configure(bg=ModernStyle.BG_LIGHT)
                self.window.transient(parent)
                self.window.grab_set()
                
                # 🔧 修正居中顯示
                parent.update_idletasks()
                parent_x = parent.winfo_rootx()
                parent_y = parent.winfo_rooty()
                parent_width = parent.winfo_width()
                parent_height = parent.winfo_height()

                # 計算中央位置
                window_width = 750
                window_height = 650
                pos_x = parent_x + (parent_width - window_width) // 2
                pos_y = parent_y + (parent_height - window_height) // 2

                self.window.geometry(f"{window_width}x{window_height}+{pos_x}+{pos_y}")
                
                # 🔧 修正2：允許調整大小，以防內容過多
                self.window.resizable(True, True)
                self.window.minsize(700, 600)
                
                # 禁止直接關閉視窗
                self.window.protocol("WM_DELETE_WINDOW", self.on_closing_attempt)
                
                self.create_detailed_ui()
                
            def create_detailed_ui(self):
                """創建修正版的詳細用戶界面"""
                # 主容器 - 使用 grid 確保按鈕固定在底部
                main_frame = tk.Frame(self.window, bg=ModernStyle.BG_LIGHT)
                main_frame.pack(fill="both", expand=True, padx=25, pady=25)
                
                # 配置 grid 權重
                main_frame.grid_rowconfigure(0, weight=1)  # 內容區域可擴展
                main_frame.grid_rowconfigure(1, weight=0)  # 按鈕區域固定
                main_frame.grid_columnconfigure(0, weight=1)
                
                # =========================
                # 內容區域（可滾動）
                # =========================
                content_frame = tk.Frame(main_frame, bg=ModernStyle.BG_LIGHT)
                content_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 15))
                
                # 標題區域
                title_frame = tk.Frame(content_frame, bg=ModernStyle.BG_LIGHT)
                title_frame.pack(fill="x", pady=(0, 15))
                
                title_icon = tk.Label(title_frame, text="⚙️", 
                                    font=("Segoe UI", 20),
                                    bg=ModernStyle.BG_LIGHT)
                title_icon.pack(side="left")
                
                self.title_label = tk.Label(title_frame, 
                                        text="批次處理進行中...",
                                        font=("微軟正黑體", 14, "bold"),
                                        bg=ModernStyle.BG_LIGHT,
                                        fg=ModernStyle.PRIMARY)
                self.title_label.pack(side="left", padx=(10, 0))
                
                # 檔案資訊區域
                info_frame = tk.Frame(content_frame, bg=ModernStyle.BG_SECONDARY, 
                                    relief="solid", borderwidth=1)
                info_frame.pack(fill="x", pady=(0, 15))
                
                info_text = (f"📁 總檔案數: {len(self.pdf_paths)} 個  |  "
                            f"💾 總大小: {self.total_size:.1f} MB  |  "
                            f"📊 總處理階段: {len(self.pdf_paths) * self.total_stages} 個")
                
                info_label = tk.Label(info_frame, text=info_text,
                                    font=("微軟正黑體", 10),
                                    bg=ModernStyle.BG_SECONDARY,
                                    fg=ModernStyle.TEXT_PRIMARY,
                                    justify="center")
                info_label.pack(padx=15, pady=10)

                # 🔧 新增：執行時間顯示
                self.runtime_label = tk.Label(info_frame, 
                                            text="⏱️ 執行時間：0.00 秒",
                                            font=("微軟正黑體", 10, "bold"),
                                            bg=ModernStyle.BG_SECONDARY,
                                            fg=ModernStyle.INFO)
                self.runtime_label.pack(padx=15, pady=(0, 10))

                # 🔧 啟動執行時間更新
                self.update_runtime()
                
                # 🔧 總進度區域 - 緊湊設計
                total_progress_frame = tk.Frame(content_frame, bg=ModernStyle.BG_LIGHT)
                total_progress_frame.pack(fill="x", pady=(0, 10))
                
                total_progress_label = tk.Label(total_progress_frame,
                                            text="整體進度:",
                                            font=("微軟正黑體", 11, "bold"),
                                            bg=ModernStyle.BG_LIGHT,
                                            fg=ModernStyle.TEXT_PRIMARY)
                total_progress_label.pack(anchor="w", pady=(0, 3))
                
                # 總進度條
                self.total_progress_var = tk.DoubleVar()
                total_max = len(self.pdf_paths) * self.total_stages
                self.total_progress_bar = ttk.Progressbar(total_progress_frame,
                                                        variable=self.total_progress_var,
                                                        maximum=total_max,
                                                        length=650,
                                                        mode='determinate')
                self.total_progress_bar.pack(pady=(0, 3))
                
                # 總進度資訊
                progress_info_frame = tk.Frame(total_progress_frame, bg=ModernStyle.BG_LIGHT)
                progress_info_frame.pack(fill="x")
                
                self.total_percentage_label = tk.Label(progress_info_frame,
                                                    text="0%",
                                                    font=("微軟正黑體", 10, "bold"),
                                                    bg=ModernStyle.BG_LIGHT,
                                                    fg=ModernStyle.SUCCESS)
                self.total_percentage_label.pack(side="left")
                
                self.total_count_label = tk.Label(progress_info_frame,
                                                text=f"0 / {len(self.pdf_paths)} 個檔案完成",
                                                font=("微軟正黑體", 10),
                                                bg=ModernStyle.BG_LIGHT,
                                                fg=ModernStyle.TEXT_SECONDARY)
                self.total_count_label.pack(side="right")
                
                # 🔧 當前檔案區域 - 包含詳細階段
                current_section = tk.LabelFrame(content_frame, 
                                            text="當前處理檔案",
                                            font=("微軟正黑體", 11, "bold"),
                                            bg=ModernStyle.BG_LIGHT,
                                            fg=ModernStyle.PRIMARY,
                                            padx=10, pady=10)
                current_section.pack(fill="x", pady=(0, 10))
                
                # 當前檔案名稱
                self.current_file_display = tk.Label(current_section,
                                                    text="準備開始處理...",
                                                    font=("微軟正黑體", 11),
                                                    bg=ModernStyle.BG_LIGHT,
                                                    fg=ModernStyle.TEXT_PRIMARY,
                                                    wraplength=650)
                self.current_file_display.pack(anchor="w", pady=(0, 8))
                
                # 當前檔案進度條
                self.current_progress_var = tk.DoubleVar()
                self.current_progress_bar = ttk.Progressbar(current_section,
                                                            variable=self.current_progress_var,
                                                            maximum=self.total_stages,
                                                            length=630,
                                                            mode='determinate')
                self.current_progress_bar.pack(pady=(0, 5))
                
                # 🔧 詳細階段資訊
                stage_detail_frame = tk.Frame(current_section, bg=ModernStyle.BG_LIGHT)
                stage_detail_frame.pack(fill="x", pady=(0, 5))
                
                self.current_stage_label = tk.Label(stage_detail_frame,
                                                text="等待開始...",
                                                font=("微軟正黑體", 10),
                                                bg=ModernStyle.BG_LIGHT,
                                                fg=ModernStyle.INFO)
                self.current_stage_label.pack(side="left")
                
                self.stage_progress_label = tk.Label(stage_detail_frame,
                                                    text="步驟 0/6",
                                                    font=("微軟正黑體", 10, "bold"),
                                                    bg=ModernStyle.BG_LIGHT,
                                                    fg=ModernStyle.SUCCESS)
                self.stage_progress_label.pack(side="right")

                # 🔧 階段狀態指示器
                stages_indicator_frame = tk.Frame(current_section, bg=ModernStyle.BG_LIGHT)
                stages_indicator_frame.pack(fill="x", pady=(5, 0))
                
                self.stage_indicators = []
                stage_names = [
                    "📥 載入", "📄 提取", "🧹 清理", 
                    "🔍 OCR", "📊 解析", "✨ 完善"
                ]
                
                for i, stage_name in enumerate(stage_names):
                    indicator = tk.Label(stages_indicator_frame,
                                    text=stage_name,
                                    font=("微軟正黑體", 9),
                                    bg=ModernStyle.BG_SECONDARY,
                                    fg=ModernStyle.TEXT_SECONDARY,
                                    relief="solid",
                                    borderwidth=1,
                                    padx=8, pady=3)
                    indicator.pack(side="left", padx=(0, 3))
                    self.stage_indicators.append(indicator)
                
                # 🔧 處理歷史區域 - 緊湊設計
                history_section = tk.LabelFrame(content_frame,
                                            text="處理記錄",
                                            font=("微軟正黑體", 11, "bold"),
                                            bg=ModernStyle.BG_LIGHT,
                                            fg=ModernStyle.SECONDARY,
                                            padx=6, pady=3)
                history_section.pack(fill="both", expand=True)
                
                # 滾動式歷史記錄
                self.history_text = scrolledtext.ScrolledText(history_section,
                                                            height=8,
                                                            width=80,
                                                            font=("Consolas", 9),
                                                            bg="#F8F9FA",
                                                            fg=ModernStyle.TEXT_PRIMARY,
                                                            relief="solid",
                                                            borderwidth=1,
                                                            state="disabled")
                self.history_text.pack(fill="both", expand=True)
                
                # =========================
                # 🔧 修正3：固定按鈕區域，避免重複按鈕
                # =========================
                button_frame = tk.Frame(main_frame, bg=ModernStyle.BG_LIGHT, 
                                    relief="solid", borderwidth=1, padx=20, pady=15)
                button_frame.grid(row=1, column=0, sticky="ew")
                
                # 按鈕容器
                btn_container = tk.Frame(button_frame, bg=ModernStyle.BG_LIGHT)
                btn_container.pack(fill="x")
                
                # 🔧 修正4：只創建兩個按鈕，完成時只改變其中一個
                self.cancel_btn = tk.Button(btn_container, 
                                        text="❌ 取消處理", 
                                        command=self.cancel_processing,
                                        bg=ModernStyle.DANGER,
                                        fg=ModernStyle.TEXT_LIGHT,
                                        font=("微軟正黑體", 12, "bold"),
                                        relief="flat",
                                        borderwidth=0,
                                        cursor="hand2",
                                        padx=25,
                                        pady=12,
                                        width=18)
                self.cancel_btn.pack(side="left")
                
                # 暫停按鈕
                self.pause_btn = tk.Button(btn_container, 
                                        text="⏸️ 暫停處理", 
                                        command=self.toggle_pause,
                                        bg=ModernStyle.WARNING,
                                        fg=ModernStyle.TEXT_PRIMARY,
                                        font=("微軟正黑體", 12),
                                        relief="flat",
                                        borderwidth=0,
                                        cursor="hand2",
                                        padx=25,
                                        pady=12,
                                        width=18)
                self.pause_btn.pack(side="right")
                
            def add_history(self, message, level="INFO"):
                """添加處理歷史記錄 - 帶顏色"""
                try:
                    self.history_text.config(state="normal")
                    from datetime import datetime
                    timestamp = datetime.now().strftime("%H:%M:%S")
                    
                    # 根據級別設置顏色標記
                    if level == "SUCCESS":
                        prefix = "✅"
                    elif level == "WARNING": 
                        prefix = "⚠️"
                    elif level == "ERROR":
                        prefix = "❌"
                    else:
                        prefix = "ℹ️"
                    
                    self.history_text.insert(tk.END, f"[{timestamp}] {prefix} {message}\n")
                    self.history_text.see(tk.END)
                    self.history_text.config(state="disabled")
                    self.window.update_idletasks()
                except Exception as e:
                    print(f"添加歷史記錄失敗: {e}", flush=True)
            
            def update_progress(self, file_index, filename, stage, stage_num, stage_description=""):
                """🔧 修正5：更新詳細進度，包含階段指示器"""
                if self.cancelled:
                    return
                
                try:
                    self.current_file_index = file_index
                    self.current_stage = stage
                    
                    # 計算總進度
                    total_completed_stages = (file_index - 1) * self.total_stages + stage_num
                    self.total_progress_var.set(total_completed_stages)
                    
                    total_max = len(self.pdf_paths) * self.total_stages
                    total_percentage = int((total_completed_stages / total_max) * 100)
                    self.total_percentage_label.config(text=f"{total_percentage}%")
                    self.total_count_label.config(text=f"{file_index-1} / {len(self.pdf_paths)} 個檔案完成")
                    
                    # 更新當前檔案資訊
                    short_filename = filename[:60] + "..." if len(filename) > 60 else filename
                    self.current_file_display.config(
                        text=f"📁 第 {file_index} 個檔案: {short_filename}"
                    )
                    
                    # 更新當前檔案進度
                    self.current_progress_var.set(stage_num)
                    
                    # 更新階段資訊
                    stage_text = f"{stage}"
                    if stage_description:
                        stage_text += f": {stage_description}"
                    self.current_stage_label.config(text=stage_text)
                    self.stage_progress_label.config(text=f"步驟 {stage_num}/{self.total_stages}")
                    
                    # 🔧 修正6：更新階段指示器
                    self.update_stage_indicators(stage_num)
                    
                    # 添加到歷史記錄
                    if stage_num == 1:  # 開始新檔案時
                        self.add_history(f"開始處理: {filename}", "INFO")
                    elif stage_num == self.total_stages:  # 完成檔案時
                        self.add_history(f"完成處理: {filename}", "SUCCESS")
                    
                    # 強制更新界面
                    self.window.update_idletasks()
                    
                except Exception as e:
                    print(f"更新進度時出錯: {e}", flush=True)
            
            def update_stage_indicators(self, current_stage):
                """更新階段指示器的視覺狀態"""
                try:
                    for i, indicator in enumerate(self.stage_indicators):
                        if i < current_stage:
                            # 已完成階段
                            indicator.config(bg=ModernStyle.SUCCESS, 
                                        fg=ModernStyle.TEXT_LIGHT,
                                        relief="solid")
                        elif i == current_stage - 1:
                            # 當前階段
                            indicator.config(bg=ModernStyle.WARNING, 
                                        fg=ModernStyle.TEXT_PRIMARY,
                                        relief="solid")
                        else:
                            # 未開始階段
                            indicator.config(bg=ModernStyle.BG_SECONDARY, 
                                        fg=ModernStyle.TEXT_SECONDARY,
                                        relief="solid")
                except Exception as e:
                    print(f"更新階段指示器失敗: {e}", flush=True)
            
            def update_runtime(self):
                """更新執行時間顯示"""
                if not self.cancelled and hasattr(self, 'runtime_label'):
                    import time
                    current_time = time.time()
                    elapsed = current_time - self.start_time
                    
                    # 格式化執行時間（保留2位小數）
                    if elapsed >= 60:
                        minutes = int(elapsed // 60)
                        seconds = elapsed % 60
                        time_str = f"{minutes}:{seconds:05.2f}"
                    else:
                        time_str = f"{elapsed:.2f} 秒"
                    
                    self.runtime_label.config(text=f"⏱️ 執行時間：{time_str}")
                    
                    # 每0.1秒更新一次
                    if hasattr(self, 'window'):
                        self.window.after(100, self.update_runtime)

            def close(self):
                """關閉進度視窗 - 停止計時"""
                self.cancelled = True  # 停止計時更新
                if hasattr(self, 'window'):
                    try:
                        self.window.destroy()
                    except:
                        pass
            
            
            def cancel_processing(self):
                """取消處理"""
                if self.cancelled:
                    return
                
                result = messagebox.askyesno(
                    "確認取消", 
                    f"確定要取消批次處理嗎？\n\n"
                    f"📊 當前進度: 第 {self.current_file_index} / {len(self.pdf_paths)} 個檔案\n"
                    f"⚠️ 已處理的結果可能會遺失",
                    parent=self.window
                )
                
                if result:
                    self.cancelled = True
                    self.title_label.config(text="正在取消批次處理...")
                    self.current_stage_label.config(text="❌ 正在取消處理...")
                    self.add_history("用戶手動取消了批次處理", "WARNING")
                    
                    # 禁用按鈕
                    self.cancel_btn.config(text="⏳ 取消中...", state="disabled")
                    self.pause_btn.config(state="disabled")
                    
                    self.window.update_idletasks()
            
            def toggle_pause(self):
                """暫停/繼續功能"""
                if self.cancelled:
                    return
                
                self.paused = not self.paused
                
                if self.paused:
                    self.pause_btn.config(text="▶️ 繼續處理", bg=ModernStyle.SUCCESS)
                    self.current_stage_label.config(text="⏸️ 處理已暫停 - 點擊繼續恢復")
                    self.title_label.config(text="批次處理已暫停")
                    self.add_history("處理已暫停", "WARNING")
                else:
                    self.pause_btn.config(text="⏸️ 暫停處理", bg=ModernStyle.WARNING)
                    self.current_stage_label.config(text="🔄 處理已恢復...")
                    self.title_label.config(text="批次處理進行中...")
                    self.add_history("處理已恢復", "INFO")
            
            def on_closing_attempt(self):
                """嘗試關閉視窗時的處理"""
                messagebox.showwarning(
                    "無法關閉", 
                    "請使用「❌ 取消處理」按鈕來停止批次處理，\n或等待處理完成後點擊確定關閉。",
                    parent=self.window
                )
            
            def show_completion(self, success=True, message="", results_count=0):
                """🔧 修正7：顯示完成狀態，只改變一個按鈕為確定"""
                if success and not self.cancelled:
                    self.total_progress_var.set(len(self.pdf_paths) * self.total_stages)
                    self.total_percentage_label.config(text="100%")
                    self.total_count_label.config(text=f"{len(self.pdf_paths)} / {len(self.pdf_paths)} 個檔案完成")
                    self.current_file_display.config(text="🎉 所有檔案處理完成！")
                    self.current_stage_label.config(text=f"✅ 成功處理 {results_count} 份謄本")
                    self.title_label.config(text="批次處理完成！")
                    self.current_progress_var.set(self.total_stages)
                    self.add_history(f"批次處理完成！成功處理 {results_count} 份謄本", "SUCCESS")
                    
                    # 更新所有階段指示器為完成狀態
                    self.update_stage_indicators(self.total_stages)
                    
                    # 🔧 修正：只改變一個按鈕，移除暫停按鈕
                    self.cancel_btn.config(text="✅ 確定關閉", 
                                        command=self.close,
                                        bg=ModernStyle.SUCCESS,
                                        state="normal")
                    self.pause_btn.pack_forget()  # 隱藏暫停按鈕
                    
                elif self.cancelled:
                    self.current_stage_label.config(text="❌ 處理已取消")
                    self.title_label.config(text="批次處理已取消")
                    self.cancel_btn.config(text="確定關閉", command=self.close, state="normal")
                    self.pause_btn.pack_forget()
                    self.add_history("批次處理已取消", "WARNING")
                    
                else:
                    self.current_stage_label.config(text=f"❌ 處理失敗: {message}")
                    self.title_label.config(text="批次處理失敗")
                    self.cancel_btn.config(text="確定關閉", command=self.close, state="normal")
                    self.pause_btn.pack_forget()
                    self.add_history(f"批次處理失敗: {message}", "ERROR")
            
            def close(self):
                """關閉進度視窗"""
                if hasattr(self, 'window'):
                    try:
                        self.window.destroy()
                    except:
                        pass
        
        return DetailedBatchProgressWindow(self.master, pdf_paths, total_size)
        
    def estimate_batch_time(self, total_size_mb):
        """估算批次處理時間"""
        if total_size_mb < 10:
            return "約 2-5 分鐘"
        elif total_size_mb < 50:
            return "約 5-15 分鐘" 
        elif total_size_mb < 100:
            return "約 15-30 分鐘"
        else:
            return "約 30 分鐘以上"
    
   
    def create_progress_window(self, total_files):
        """創建進度視窗 - 加上取消功能"""
        class ProgressWindow:
            def __init__(self, parent, total):
                self.window = tk.Toplevel(parent)
                self.window.title("🔄 批次處理進度")
                self.window.geometry("500x300")  # 增加高度以容納取消按鈕
                self.window.configure(bg=ModernStyle.BG_LIGHT)
                self.window.transient(parent)
                self.window.grab_set()
                
                # 居中顯示
                # 🔧 使用螢幕適應性工具計算最佳位置
                window_width = 600
                window_height = 400
                pos_x, pos_y = ScreenUtils.calculate_dialog_position(parent, window_width, window_height)

                self.window.geometry(f"{window_width}x{window_height}+{pos_x}+{pos_y}")
                
                # 禁止調整大小
                self.window.resizable(False, False)
                
                # 🔧 重要：禁止直接關閉視窗
                self.window.protocol("WM_DELETE_WINDOW", self.on_closing_attempt)
                
                self.cancelled = False
                self.total = total
                self.create_ui()
            
            def create_ui(self):
                # 主容器
                main_frame = tk.Frame(self.window, bg=ModernStyle.BG_LIGHT)
                main_frame.pack(fill="both", expand=True, padx=30, pady=30)
                
                # 標題
                self.title_label = tk.Label(main_frame, 
                                        text="正在批次處理謄本檔案...",
                                        font=ModernStyle.FONT_HEADER,
                                        bg=ModernStyle.BG_LIGHT,
                                        fg=ModernStyle.PRIMARY)
                self.title_label.pack(pady=(0, 15))
                
                # 檔案資訊
                info_frame = tk.Frame(main_frame, bg=ModernStyle.BG_SECONDARY, 
                                    relief="solid", borderwidth=1)
                info_frame.pack(fill="x", pady=(0, 20))
                
                info_text = f"📁 待處理檔案: {self.total} 個\n⏱️ 請耐心等待處理完成"
                info_label = tk.Label(info_frame, text=info_text,
                                    font=ModernStyle.FONT_MAIN,
                                    bg=ModernStyle.BG_SECONDARY,
                                    fg=ModernStyle.TEXT_PRIMARY,
                                    justify="left")
                info_label.pack(padx=15, pady=10)
                
                # 進度條
                progress_frame = tk.Frame(main_frame, bg=ModernStyle.BG_LIGHT)
                progress_frame.pack(fill="x", pady=(0, 20))
                
                self.progress_var = tk.DoubleVar()
                self.progress_bar = ttk.Progressbar(progress_frame, 
                                                variable=self.progress_var,
                                                maximum=self.total,
                                                length=400,
                                                mode='determinate')
                self.progress_bar.pack(pady=10)
                
                # 狀態標籤
                self.status_label = tk.Label(progress_frame,
                                            text=f"準備處理 {self.total} 個檔案...",
                                            font=ModernStyle.FONT_MAIN,
                                            bg=ModernStyle.BG_LIGHT,
                                            fg=ModernStyle.TEXT_SECONDARY)
                self.status_label.pack()
                
                # 🔧 新增：取消按鈕
                button_frame = tk.Frame(main_frame, bg=ModernStyle.BG_LIGHT)
                button_frame.pack(pady=(20, 0))
                
                self.cancel_btn = ModernButton(button_frame, "❌ 取消處理", 
                                            command=self.cancel_processing, 
                                            style="danger")
                self.cancel_btn.pack()
            
            def update_progress(self, current, filename):
                """更新進度"""
                if not self.cancelled:
                    self.progress_var.set(current)
                    self.status_label.config(
                        text=f"正在處理: {filename} ({current}/{self.total})"
                    )
                    self.window.update_idletasks()
            
            def cancel_processing(self):
                """🔧 新增：取消處理功能"""
                import tkinter.messagebox as msgbox
                result = msgbox.askyesno(
                    "確認取消", 
                    f"確定要取消批次處理嗎？\n\n"
                    f"📊 進度: {int(self.progress_var.get())}/{self.total}\n"
                    f"⚠️ 已處理的進度將會遺失。",
                    parent=self.window
                )
                if result:
                    self.cancelled = True
                    self.status_label.config(text="正在取消處理...")
                    self.title_label.config(text="正在取消批次處理...")
                    self.cancel_btn.config(text="取消中...", state="disabled")
                    self.window.update_idletasks()
            
            def on_closing_attempt(self):
                """🔧 新增：阻止直接關閉視窗"""
                import tkinter.messagebox as msgbox
                msgbox.showwarning(
                    "無法關閉", 
                    "請使用「❌ 取消處理」按鈕來停止批次處理",
                    parent=self.window
                )
            
            def close(self):
                """關閉進度視窗"""
                if hasattr(self, 'window'):
                    try:
                        self.window.destroy()
                    except:
                        pass
        
        return ProgressWindow(self.master, total_files)

    def display_results(self, results: List[Dict]):
        """顯示處理結果到結構化資料標籤頁"""
        self.struct_text.config(state="normal")
        self.struct_text.delete(1.0, tk.END)
        
        # 美化的結果顯示
        self.struct_text.insert(tk.END, f"🎯 處理結果報告\n")
        self.struct_text.insert(tk.END, f"{'='*40}\n")
        self.struct_text.insert(tk.END, f"📅 處理時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        self.struct_text.insert(tk.END, f"📊 總謄本數: {len(results)}\n")
        
        # 統計各類型謄本
        type_stats = {}
        for result in results:
            t_type = result.get('謄本類型', '未知類型')
            type_stats[t_type] = type_stats.get(t_type, 0) + 1
        
        self.struct_text.insert(tk.END, f"\n📈 謄本類型統計:\n")
        for t_type, count in type_stats.items():
            self.struct_text.insert(tk.END, f"   • {t_type}: {count} 份\n")
        
        self.struct_text.insert(tk.END, f"\n{'='*40}\n\n")
        
        # 詳細結果
        for i, result in enumerate(results, 1):
            self.struct_text.insert(tk.END, f"📄 謄本 {i}: {result.get('源檔案', '未知檔案')}\n")
            self.struct_text.insert(tk.END, f"{'-'*60}\n")
            
            # 遞歸顯示結果
            self.display_dict_content(result, indent=0)
            self.struct_text.insert(tk.END, "\n\n")
        
        self.struct_text.config(state="disabled")
        
        # 自動切換到結構化資料標籤頁
        self.switch_tab("struct")

    
    def display_dict_content(self, data: Any, indent: int = 0):
        """遞歸顯示字典內容（美化版）"""
        prefix = "  " * indent
        
        if isinstance(data, dict):
            for key, value in data.items():
                if isinstance(value, (dict, list)):
                    self.struct_text.insert(tk.END, f"{prefix}📁 {key}:\n")
                    self.display_dict_content(value, indent + 1)
                else:
                    # 根據欄位類型選擇圖示
                    icon = self.get_field_icon(key)
                    display_value = str(value)[:100] + "..." if len(str(value)) > 100 else str(value)
                    self.struct_text.insert(tk.END, f"{prefix}{icon} {key}: {display_value}\n")
        elif isinstance(data, list):
            for i, item in enumerate(data):
                self.struct_text.insert(tk.END, f"{prefix}📋 [{i+1}]:\n")
                self.display_dict_content(item, indent + 1)
        else:
            display_value = str(data)[:150] + "..." if len(str(data)) > 150 else str(data)
            self.struct_text.insert(tk.END, f"{prefix}📝 {display_value}\n")
    
    def get_field_icon(self, field_name):
        """根據欄位名稱返回對應圖示"""
        icon_map = {
            "謄本類型": "📋", "源檔案": "📄", "地號建號": "🏠", "列印時間": "⏰",
            "所有權人": "👤", "統一編號": "🆔", "住址": "🏠", "權利範圍": "⚖️",
            "登記日期": "📅", "登記原因": "📝", "建物門牌": "🏢", "主要用途": "🏛️",
            "面積": "📐", "總面積": "📐", "擔保債權總金額": "💰", "權利種類": "⚖️"
        }
        
        for key, icon in icon_map.items():
            if key in field_name:
                return icon
        return "•"

    def export_format(self, format_type: str, silent_mode: bool = False):
        """匯出指定格式 - 修復版本"""
        if not hasattr(self, 'current_results') or not self.current_results:
            self.show_custom_messagebox("提示", "請先進行處理以產生可匯出的資料")
            return None

        try:
            if format_type == "txt":
                filename = self.exporter._create_flexible_filename(self.current_results, "txt")
                filepath = self.exporter.export_to_txt_report(self.current_results, filename)

            elif format_type == "excel":
                filename = self.exporter._create_flexible_filename(self.current_results, "xlsx")
                filepath = self.exporter.export_to_excel(self.current_results, filename)

            elif format_type == "json":
                filename = self.exporter._create_flexible_filename(self.current_results, "json")
                filepath = self.exporter.export_to_json(self.current_results, filename)

            else:
                raise ValueError(f"不支援的格式: {format_type}")

            if not silent_mode:
                self.show_custom_messagebox("匯出成功", f"已匯出到:\n{filepath}")
            return filepath

            
        except Exception as e:
            print(f"❌ {format_type.upper()}匯出失敗: {e}")  # 🔥 總是 print 錯誤
            import traceback
            traceback.print_exc()  # 🔥 顯示完整 traceback
            if not silent_mode:
                messagebox.showerror("錯誤", f"{format_type.upper()}匯出失敗: {e}")
            self.progress_indicator.set_status("匯出失敗", "error")
            return None

    def export_all(self):
        """匯出所有格式 - 修復版本"""
        if not hasattr(self, 'current_results') or not self.current_results:
            self.show_custom_messagebox("提示", "請先進行處理以產生可匯出的資料")
            return
        
        try:
            self.progress_indicator.set_status("正在匯出所有格式...", "processing")
            
            # 🔧 修復：使用靜默模式匯出，收集所有檔案路徑
            exported_files = []
            
            # 依序匯出三種格式（靜默模式）
            txt_path = self.export_format("txt", silent_mode=True)
            if txt_path:
                exported_files.append(("TXT報告", txt_path))
            
            excel_path = self.export_format("excel", silent_mode=True)
            if excel_path:
                exported_files.append(("Excel表格", excel_path))
            
            json_path = self.export_format("json", silent_mode=True)
            if json_path:
                exported_files.append(("JSON資料", json_path))
            
            self.progress_indicator.set_status("所有格式匯出完成", "success")
            
            # 🔧 修復：統一顯示匯出結果和詢問是否開啟資料夾
            if exported_files:
                # 建立匯出成功清單
                file_list = "\n".join([f"• {name}: {os.path.basename(path)}" for name, path in exported_files])
                
                message = (f"✅ 全部格式匯出完成！\n\n"
                        f"📦 已成功匯出 {len(exported_files)} 個檔案:\n"
                        f"{file_list}\n\n"
                        f"💾 儲存目錄: {os.path.dirname(exported_files[0][1])}\n\n"
                        f"是否要開啟檔案所在資料夾？")
                
                result = self.show_custom_messagebox("批次匯出完成", message, "yesno")
                
                if result:
                    self.open_output_folder()
            else:
                self.show_custom_messagebox("錯誤", "所有格式匯出都失敗了")
                
        except Exception as e:
            messagebox.showerror("錯誤", f"批次匯出失敗: {e}")
            self.progress_indicator.set_status("匯出失敗", "error")

    # 🔧 額外改進：也可以為單獨匯出提供選擇
    def export_format_with_options(self, format_type: str):
        """匯出指定格式並提供選項 - 改進版本"""
        if not hasattr(self, 'current_results') or not self.current_results:
            self.show_custom_messagebox("提示", "請先進行處理以產生可匯出的資料")
            return
        
        # 先匯出檔案
        filepath = self.export_format(format_type, silent_mode=True)
        
        if filepath:
            # 顯示匯出完成對話框，提供多個選項
            self.show_export_completion_dialog(format_type, filepath)

    def show_export_completion_dialog(self, format_type, filepath):
        """顯示匯出完成對話框 - 提供多個選項"""
        dialog = tk.Toplevel(self.master)
        dialog.title("匯出完成")
        dialog.geometry("450x300")
        dialog.configure(bg=ModernStyle.BG_LIGHT)
        dialog.transient(self.master)
        dialog.grab_set()
        
        # 居中顯示
        dialog.geometry("+%d+%d" % (
            self.master.winfo_rootx() + 100, 
            self.master.winfo_rooty() + 100
        ))
        
        # 主要內容
        main_frame = tk.Frame(dialog, bg=ModernStyle.BG_LIGHT)
        main_frame.pack(fill="both", expand=True, padx=25, pady=25)
        
        # 成功圖示
        success_label = tk.Label(main_frame, 
                            text="✅",
                            font=("Segoe UI", 32),
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.SUCCESS)
        success_label.pack(pady=(0, 15))
        
        # 標題
        title_label = tk.Label(main_frame, 
                            text=f"{format_type.upper()} 檔案匯出完成！",
                            font=ModernStyle.FONT_HEADER,
                            bg=ModernStyle.BG_LIGHT,
                            fg=ModernStyle.PRIMARY)
        title_label.pack(pady=(0, 10))
        
        # 檔案資訊
        info_frame = tk.Frame(main_frame, bg=ModernStyle.BG_SECONDARY, 
                            relief="solid", borderwidth=1)
        info_frame.pack(fill="x", pady=(0, 20))
        
        info_text = (f"📁 檔案名稱: {os.path.basename(filepath)}\n"
                    f"💾 儲存位置: {os.path.dirname(filepath)}")
        
        info_label = tk.Label(info_frame, text=info_text,
                            font=ModernStyle.FONT_MAIN,
                            bg=ModernStyle.BG_SECONDARY,
                            fg=ModernStyle.TEXT_PRIMARY,
                            justify="left")
        info_label.pack(padx=15, pady=10)
        
        # 按鈕區域
        button_frame = tk.Frame(main_frame, bg=ModernStyle.BG_LIGHT)
        button_frame.pack(fill="x", pady=(10, 0))
        
        # 按鈕函數
        def open_folder():
            self.open_output_folder()
            dialog.destroy()
        
        def close_dialog():
            dialog.destroy()
        
        # 按鈕
        ModernButton(button_frame, "📂 開啟資料夾", 
                    command=open_folder, 
                    style="primary").pack(side="left", padx=(0, 10))
        
        ModernButton(button_frame, "確定", 
                    command=close_dialog, 
                    style="secondary").pack(side="right")

    def show_program_status(self):
        """顯示程式狀態資訊 - 完整版本"""
        # 🔧 創建美化的狀態視窗
        status_window = tk.Toplevel(self.master)
        status_window.title("程式狀態")
        status_window.geometry("650x550")  # 👈 稍微加高
        status_window.configure(bg="#ecf0f1")
        status_window.transient(self.master)
        status_window.grab_set()
        status_window.resizable(True, True)
        status_window.minsize(650, 550)  # 🔧 設定最小大小
        
        # 居中顯示
        status_window.update_idletasks()
        x = (status_window.winfo_screenwidth() // 2) - (325)
        y = (status_window.winfo_screenheight() // 2) - (275)
        status_window.geometry(f"650x500+{x}+{y}")
        
        # 主框架 - 加入邊框
        main_frame = tk.Frame(status_window, bg="#ecf0f1", relief="solid", borderwidth=1)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 標題
        title_label = tk.Label(main_frame, 
                            text="📊 程式狀態資訊", 
                            font=("微軟正黑體", 16, "bold"),
                            bg="#ecf0f1",
                            fg="#2c3e50")
        title_label.pack(pady=(15, 20))
        
        # 內容框架 - 白色背景
        content_frame = tk.Frame(main_frame, bg="white", relief="groove", borderwidth=2)
        content_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        # 🔧 收集完整的狀態資訊
        current_dir = getattr(self, 'current_folder', '未設定')
        output_dir = getattr(self, 'output_directory', '未設定')
        processing_state = getattr(self, 'processing_state', 'none')
        processed_count = len(getattr(self, 'processed_files', set()))
        
        # 處理狀態轉換
        status_text = {
            'none': '未開始',
            'processing': '處理中', 
            'completed': '已完成',
            'error': '發生錯誤'
        }.get(processing_state, processing_state)
        
        # 🔧 完整的狀態項目
        status_items = [
            ("🏠 當前瀏覽目錄", current_dir),
            ("📁 輸出目錄", output_dir),
            ("⚡ 處理狀態", status_text),
            ("📄 已處理檔案", f"{processed_count} 個"),
            ("💾 記憶體使用", self.get_memory_usage()),
            ("⏰ 程式執行時間", self.get_runtime())
        ]
        
        # 美化的狀態顯示
        for i, (label, value) in enumerate(status_items):
            item_frame = tk.Frame(content_frame, bg="white")
            item_frame.pack(fill="x", padx=15, pady=4)
            
            # 標籤
            label_text = tk.Label(item_frame, 
                                text=label + "：", 
                                font=("微軟正黑體", 11, "bold"),
                                bg="white",
                                fg="#34495e",
                                anchor="w")
            label_text.pack(anchor="w")
            
            # 值 - 用不同顏色的背景
            value_bg = "#f8f9fa" if i % 2 == 0 else "#ffffff"
            value_text = tk.Label(item_frame, 
                                text=value, 
                                font=("微軟正黑體", 10),
                                bg=value_bg,
                                fg="#2c3e50",
                                anchor="w",
                                relief="solid",
                                borderwidth=1,
                                padx=10, pady=4)
            value_text.pack(fill="x", pady=(2, 0))
        
        # 關閉按鈕
        button_frame = tk.Frame(main_frame, bg="#ecf0f1")
        button_frame.pack(side="bottom", fill="x", pady=(15, 10))
        
        close_btn = tk.Button(button_frame, 
                            text="✅ 確定",
                            command=status_window.destroy,
                            font=("微軟正黑體", 12, "bold"),
                            bg="#3498db", fg="white",
                            relief="flat", bd=0,
                            width=10, height=1)
        close_btn.pack()

    def get_memory_usage(self):
        """取得記憶體使用量"""
        try:
            import psutil
            process = psutil.Process()
            memory_mb = process.memory_info().rss / 1024 / 1024
            return f"{memory_mb:.1f} MB"
        except ImportError:
            return "未安裝 psutil"
        except Exception:
            return "無法取得"

    def get_runtime(self):
        """取得程式執行時間"""
        if not hasattr(self, 'start_time'):
            import time
            self.start_time = time.time()
            return "剛啟動"
        
        import time
        runtime = time.time() - self.start_time
        hours = int(runtime // 3600)
        minutes = int((runtime % 3600) // 60)
        seconds = int(runtime % 60)
        
        if hours > 0:
            return f"{hours}時{minutes}分{seconds}秒"
        elif minutes > 0:
            return f"{minutes}分{seconds}秒"
        else:
            return f"{seconds}秒"

    def get_memory_usage(self):
        """取得記憶體使用量"""
        try:
            import psutil
            process = psutil.Process()
            memory_mb = process.memory_info().rss / 1024 / 1024
            return f"{memory_mb:.1f} MB"
        except:
            return "無法取得"

    def get_runtime(self):
        """取得程式執行時間"""
        if not hasattr(self, 'start_time'):
            self.start_time = time.time()
            return "剛啟動"
        
        runtime = time.time() - self.start_time
        hours = int(runtime // 3600)
        minutes = int((runtime % 3600) // 60)
        seconds = int(runtime % 60)
        
        if hours > 0:
            return f"{hours}時{minutes}分{seconds}秒"
        elif minutes > 0:
            return f"{minutes}分{seconds}秒"
        else:
            return f"{seconds}秒"

    def open_output_folder(self):
        """開啟輸出資料夾 - 使用統一設定"""
        try:
            # 🔧 改用統一的輸出目錄
            output_dir = self.get_current_output_directory()
            
            if os.name == 'nt':  # Windows
                os.startfile(output_dir)
            elif os.name == 'posix':  # macOS and Linux
                os.system(f'open "{output_dir}"' if sys.platform == 'darwin' else f'xdg-open "{output_dir}"')
                
            print(f"📂 已打開輸出目錄：{output_dir}", flush=True)
            
        except Exception as e:
            print(f"❌ 打開輸出目錄失敗：{e}", flush=True)
    
    def show_settings(self):
        """顯示系統設定"""
        settings_window = tk.Toplevel(self.master)
        settings_window.title("⚙️ 系統設定")
        settings_window.geometry("500x500")
        # 🔧 居中顯示
        self.master.update_idletasks()
        main_x = self.master.winfo_rootx()
        main_y = self.master.winfo_rooty()
        main_width = self.master.winfo_width()
        main_height = self.master.winfo_height()

        pos_x = main_x + (main_width - 500) // 2
        pos_y = main_y + (main_height - 500) // 2
        settings_window.geometry(f"500x400+{pos_x}+{pos_y}")
        settings_window.configure(bg=ModernStyle.BG_LIGHT)
        settings_window.transient(self.master)
        
        # 設定內容
        settings_frame = ModernFrame(settings_window, title="系統設定")
        settings_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 處理器設定
        processor_frame = tk.LabelFrame(settings_frame, 
                                      text="🔧 處理器設定",
                                      font=ModernStyle.FONT_HEADER,
                                      bg=ModernStyle.BG_LIGHT,
                                      fg=ModernStyle.PRIMARY)
        processor_frame.pack(fill="x", pady=(0, 15))
        
        processor_status = "原程式處理器" if self.using_original else "增強版處理器"
        tk.Label(processor_frame, 
                text=f"當前使用: {processor_status}",
                font=ModernStyle.FONT_MAIN,
                bg=ModernStyle.BG_LIGHT).pack(anchor="w", padx=6, pady=3)
        
        # 路徑設定
        path_frame = tk.LabelFrame(settings_frame, 
                                 text="📁 路徑設定",
                                 font=ModernStyle.FONT_HEADER,
                                 bg=ModernStyle.BG_LIGHT,
                                 fg=ModernStyle.PRIMARY)
        path_frame.pack(fill="x", pady=(0, 15))
        
        tk.Label(path_frame, 
                text=f"PDF資料夾: {PDF_DIR}",
                font=ModernStyle.FONT_MAIN,
                bg=ModernStyle.BG_LIGHT).pack(anchor="w", padx=10, pady=2)
        
        tk.Label(path_frame, 
                text=f"輸出資料夾: {OUT_DIR}",
                font=ModernStyle.FONT_MAIN,
                bg=ModernStyle.BG_LIGHT).pack(anchor="w", padx=10, pady=2)
        
        # 關閉按鈕
        ModernButton(settings_frame, "關閉", 
                    command=settings_window.destroy, 
                    style="secondary").pack(pady=20)
    
    def select_template(self, template_name):
        """選擇匯出模板"""
        self.current_template = template_name
        self.update_template_buttons()
        
        # 更新顯示文字
        descriptions = {
            "完整": "(完整資料)",
            "核心": "(基本+所有權)", 
            "房產": "(基本+標示+所有權)",
            "權利": "(所有權+他項權利)"
        }
        
        self.current_selection_label.config(
            text=descriptions.get(template_name, "(自訂)")
        )
        
        # 更新狀態提示
        self.progress_indicator.set_status(f"已選擇 {template_name} 模板", "info")

    def update_template_buttons(self):
        """更新模板按鈕的視覺狀態"""
        for name, button in self.template_buttons.items():
            if name == self.current_template:
                # 選中狀態
                button.config(bg=ModernStyle.PRIMARY, 
                             fg=ModernStyle.TEXT_LIGHT,
                             relief="solid")
            else:
                # 未選中狀態
                button.config(bg=ModernStyle.BG_LIGHT, 
                             fg=ModernStyle.TEXT_PRIMARY,
                             relief="solid")

    def filter_data_by_template(self, data, template_type):
        """根據模板過濾資料 - 修正版本（所有權部最重要）"""
        if not data:
            return data
        
        if template_type == "all":
            # 完整資料：不過濾
            return data
        
        elif template_type == "core":
            # 核心模板：基本資訊 + 所有權部（最重要的資訊）
            filtered_data = []
            for item in data:
                filtered_item = {}
                
                # 保留基本欄位
                basic_fields = ["謄本類型", "源檔案", "檔案序號", "謄本序號"]
                for field in basic_fields:
                    if field in item:
                        filtered_item[field] = item[field]
                
                # 保留基本資訊
                if "基本資訊" in item:
                    filtered_item["基本資訊"] = item["基本資訊"]
                
                # 🔧 最重要：保留所有權部
                if "所有權部" in item:
                    filtered_item["所有權部"] = item["所有權部"]
                
                filtered_data.append(filtered_item)
            
            return filtered_data
        
        elif template_type == "property":
            # 房產模板：基本資訊 + 標示部 + 所有權部（房仲最需要的）
            filtered_data = []
            for item in data:
                filtered_item = {}
                
                # 保留基本欄位
                basic_fields = ["謄本類型", "源檔案", "檔案序號", "謄本序號"]
                for field in basic_fields:
                    if field in item:
                        filtered_item[field] = item[field]
                
                # 保留基本資訊
                if "基本資訊" in item:
                    filtered_item["基本資訊"] = item["基本資訊"]
                
                # 保留標示部（房屋基本資料）
                if "標示部" in item:
                    filtered_item["標示部"] = item["標示部"]
                
                # 🔧 重要：保留所有權部（誰擁有）
                if "所有權部" in item:
                    filtered_item["所有權部"] = item["所有權部"]
                
                filtered_data.append(filtered_item)
            
            return filtered_data
        
        elif template_type == "rights":
            # 權利模板：基本資訊 + 所有權部 + 他項權利部（法務完整版）
            filtered_data = []
            for item in data:
                filtered_item = {}
                
                # 保留基本欄位
                basic_fields = ["謄本類型", "源檔案", "檔案序號", "謄本序號"]
                for field in basic_fields:
                    if field in item:
                        filtered_item[field] = item[field]
                
                # 保留基本資訊
                if "基本資訊" in item:
                    filtered_item["基本資訊"] = item["基本資訊"]
                
                # 🔧 重要：保留所有權部（誰擁有）
                if "所有權部" in item:
                    filtered_item["所有權部"] = item["所有權部"]
                
                # 保留他項權利部（有什麼負擔）
                if "他項權利部" in item:
                    filtered_item["他項權利部"] = item["他項權利部"]
                
                filtered_data.append(filtered_item)
            
            return filtered_data
        
        else:
            # 預設返回完整資料
            return data

    # 🔧 新增：使用模板匯出的方法
    def export_with_template(self, format_type):
        """使用選中的模板匯出指定格式"""
        if not hasattr(self, 'current_results') or not self.current_results:
            self.show_custom_messagebox("提示", "請先進行處理以產生可匯出的資料")
            return None
        
        try:
            # 🔧 根據選中的模板過濾資料
            template_type = self.export_templates[self.current_template]
            filtered_data = self.filter_data_by_template(self.current_results, template_type)
            
            # 顯示過濾結果
            if template_type != "all":
                self.progress_indicator.set_status(
                    f"使用 {self.current_template} 模板匯出...", "processing"
                )
            
            # 🔧 使用 exporter 的友善中文檔名函數
            base_filename = self.exporter._create_flexible_filename(filtered_data, "base")
            # 移除副檔名以便添加模板後綴
            if "." in base_filename:
                name_without_ext = base_filename.rsplit(".", 1)[0]
            else:
                name_without_ext = base_filename

            # 添加模板後綴（如果不是"全部"模板）
            template_suffix = f"_{self.current_template}" if template_type != "all" else ""

            if format_type == "txt":
                filename = f"{name_without_ext}{template_suffix}.txt"
                filepath = self.exporter.export_to_txt_report(filtered_data, filename)

            elif format_type == "excel":
                filename = f"{name_without_ext}{template_suffix}.xlsx"
                filepath = self.exporter.export_to_excel(filtered_data, filename)

            elif format_type == "json":
                filename = f"{name_without_ext}{template_suffix}.json"
                filepath = self.exporter.export_to_json(filtered_data, filename)
            
            # 🔧 更新狀態提示
            if template_type != "all":
                status_msg = f"{format_type.upper()}檔案已匯出 ({self.current_template}模板)"
            else:
                status_msg = f"{format_type.upper()}檔案已匯出"
            
            self.progress_indicator.set_status(status_msg, "success")
            
            # 詢問是否開啟資料夾
            template_info = f" ({self.current_template}模板)" if template_type != "all" else ""
            message = (f"✅ {format_type.upper()} 檔案匯出完成！{template_info}\n\n"
                    f"📁 檔案位置: {os.path.basename(filepath)}\n"
                    f"💾 儲存目錄: {os.path.dirname(filepath)}\n\n"
                    f"是否要開啟檔案所在資料夾？")
            
            result = self.show_custom_messagebox("匯出完成", message, "yesno")
            
            if result:
                self.open_output_folder()
            
            return filepath
            
        except Exception as e:
            messagebox.showerror("錯誤", f"{format_type.upper()}匯出失敗: {e}")
            self.progress_indicator.set_status("匯出失敗", "error")
            return None

    # 🔧 新增：全部格式匯出也使用模板
    def export_all_with_template(self):
        """使用選中的模板匯出所有格式"""
        if not hasattr(self, 'current_results') or not self.current_results:
            self.show_custom_messagebox("提示", "請先進行處理以產生可匯出的資料")
            return
        
        try:
            template_name = self.current_template
            self.progress_indicator.set_status(f"正在使用 {template_name} 模板匯出所有格式...", "processing")
            
            # 依序匯出三種格式（靜默模式）
            exported_files = []
            
            txt_path = self.export_with_template_silent("txt")
            if txt_path:
                exported_files.append(("TXT報告", txt_path))
            
            excel_path = self.export_with_template_silent("excel")
            if excel_path:
                exported_files.append(("Excel表格", excel_path))
            
            json_path = self.export_with_template_silent("json")
            if json_path:
                exported_files.append(("JSON資料", json_path))
            
            self.progress_indicator.set_status(f"所有格式匯出完成 ({template_name})", "success")
            
            # 顯示匯出結果
            if exported_files:
                file_list = "\n".join([f"• {name}: {os.path.basename(path)}" for name, path in exported_files])
                template_info = f" ({template_name}模板)" if self.export_templates[template_name] != "all" else ""
                
                message = (f"✅ 全部格式匯出完成！{template_info}\n\n"
                        f"📦 已成功匯出 {len(exported_files)} 個檔案:\n"
                        f"{file_list}\n\n"
                        f"💾 儲存目錄: {os.path.dirname(exported_files[0][1])}\n\n"
                        f"是否要開啟檔案所在資料夾？")
                
                result = self.show_custom_messagebox("批次匯出完成", message, "yesno")
                
                if result:
                    self.open_output_folder()
            else:
                self.show_custom_messagebox("錯誤", "所有格式匯出都失敗了")
                
        except Exception as e:
            messagebox.showerror("錯誤", f"批次匯出失敗: {e}")
            self.progress_indicator.set_status("匯出失敗", "error")

    def export_with_template_silent(self, format_type):
        """靜默模式使用模板匯出（用於批次匯出）"""
        try:
            template_type = self.export_templates[self.current_template]
            filtered_data = self.filter_data_by_template(self.current_results, template_type)

            # 🔧 使用 exporter 的友善中文檔名函數
            base_filename = self.exporter._create_flexible_filename(filtered_data, "base")
            # 移除副檔名以便添加模板後綴
            if "." in base_filename:
                name_without_ext = base_filename.rsplit(".", 1)[0]
            else:
                name_without_ext = base_filename

            # 添加模板後綴（如果不是"全部"模板）
            template_suffix = f"_{self.current_template}" if template_type != "all" else ""

            if format_type == "txt":
                filename = f"{name_without_ext}{template_suffix}.txt"
                return self.exporter.export_to_txt_report(filtered_data, filename)

            elif format_type == "excel":
                filename = f"{name_without_ext}{template_suffix}.xlsx"
                return self.exporter.export_to_excel(filtered_data, filename)

            elif format_type == "json":
                filename = f"{name_without_ext}{template_suffix}.json"
                return self.exporter.export_to_json(filtered_data, filename)
                
        except Exception as e:
            print(f"靜默匯出 {format_type} 失敗: {e}", flush=True)
            return None

    # 🆕 新增這些方法到 ModernTranscriptGUI 類別中：
    def _execute_preview_for_files(self, file_paths):
        """執行檔案預覽 - 新版本"""
        try:
            if not file_paths:
                return
            
            # 取第一個檔案進行預覽
            first_file = file_paths[0]
            filename = os.path.basename(first_file)
            
            print(f"預覽檔案: {filename}", flush=True)
            
            # 執行原有的預覽邏輯
            self._perform_single_file_preview(first_file)
            
        except Exception as e:
            print(f"預覽執行錯誤: {e}", flush=True)
            self.show_custom_messagebox("錯誤", f"預覽失敗: {e}")
            
    def go_to_folder(self, new_folder, record_history=True):
        """統一導航入口：切到 new_folder。
        record_history=True 時把目前位置推入「返回前頁」歷史，並清空前進歷史。"""
        if record_history and self.current_folder and new_folder != self.current_folder:
            self.back_stack.append(self.current_folder)
            self.forward_stack = []  # 新的導航分支，前進歷史失效
        self.current_folder = new_folder
        self.load_folder_contents()
        self._update_navigation_buttons()

    def go_back(self):
        """返回前頁（瀏覽歷史的上一頁）"""
        if self.back_stack:
            self.forward_stack.append(self.current_folder)
            self.current_folder = self.back_stack.pop()
            self.load_folder_contents()
            self._update_navigation_buttons()

    def go_parent_folder(self):
        """返回上層資料夾（往上一層）"""
        # 🔧 已經在 "本機" 或 "網路" 視圖，不做任何事
        if self.current_folder in ("本機", "網路"):
            return
        # 🔧 UNC 電腦根（\\TABLE1）→ 回「網路」視圖
        if self._is_unc_computer_root(self.current_folder):
            self.go_to_folder("網路")
            return
        # 已在磁碟機根（例如 D:\）→ 上一層顯示「本機」
        parent_dir = os.path.dirname(self.current_folder)
        if parent_dir == self.current_folder:
            self.go_to_folder("本機")
        else:
            self.go_to_folder(parent_dir)

    def go_forward(self):
        """前進到下一頁（瀏覽歷史的下一頁）"""
        if self.forward_stack:
            self.back_stack.append(self.current_folder)
            self.current_folder = self.forward_stack.pop()
            self.load_folder_contents()
            self._update_navigation_buttons()

    def _update_navigation_buttons(self):
        """更新導航按鈕的啟用/停用狀態"""
        if hasattr(self, 'forward_btn'):
            self.forward_btn.config(state=tk.NORMAL if self.forward_stack else tk.DISABLED)
        if hasattr(self, 'back_btn'):
            self.back_btn.config(state=tk.NORMAL if self.back_stack else tk.DISABLED)

    def go_to_my_computer(self):
        """切到「本機」（顯示所有磁碟機）"""
        self.go_to_folder("本機")

    def go_to_network(self):
        """切到「網路」（顯示區網電腦）"""
        self.go_to_folder("網路")

    def _set_busy(self, busy):
        """切換忙碌游標並強制刷新畫面，讓「請稍候」提示能立刻顯示。"""
        try:
            cursor = "watch" if busy else ""
            self.master.config(cursor=cursor)
            self.file_tree.config(cursor=cursor)
            self.master.update_idletasks()
        except Exception:
            pass

    def _is_unc_computer_root(self, p):
        """是不是「電腦根」UNC（如 \\TABLE1，只有電腦名）。
        這種路徑 os.listdir 會 WinError 67，要改用 Shell COM 列共用。"""
        if not p or not p.startswith("\\\\"):
            return False
        parts = [x for x in p[2:].split("\\") if x]
        return len(parts) == 1

    def _net_list_computers(self):
        """列出區網電腦：優先 Windows Shell COM（檔案總管同一套），後備 net view。"""
        comps = []
        try:
            import win32com.client
            shell = win32com.client.Dispatch("Shell.Application")
            net = shell.NameSpace(18)  # ssfNETWORK
            if net:
                for item in net.Items():
                    try:
                        if item.IsFolder:
                            nm = str(item.Name).strip()
                            if nm:
                                comps.append("\\\\" + nm)
                    except Exception:
                        continue
        except Exception:
            pass
        if not comps:
            import subprocess
            try:
                r = subprocess.run(["net", "view"], capture_output=True, timeout=20,
                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                raw = r.stdout or b""
                text = ""
                for enc in ('cp950', 'utf-8', 'big5', 'cp437', 'mbcs'):
                    try:
                        text = raw.decode(enc); break
                    except Exception:
                        continue
                for line in text.splitlines():
                    line = line.strip()
                    if line.startswith("\\\\"):
                        comps.append(line.split()[0])
            except Exception:
                pass
        seen, uniq = set(), []
        for c in comps:
            if c not in seen:
                seen.add(c); uniq.append(c)
        return uniq

    def _net_list_shares(self, p):
        """列出一台電腦的共用資料夾：優先 Shell COM，後備 net view \\電腦。
        回傳 [(名稱, 完整路徑)]。"""
        items = []
        try:
            import win32com.client
            shell = win32com.client.Dispatch("Shell.Application")
            ns = shell.NameSpace(p)
            if ns:
                for it in ns.Items():
                    try:
                        if it.IsFolder:
                            nm = str(it.Name).strip()
                            if nm:
                                items.append((nm, p.rstrip("\\") + "\\" + nm))
                    except Exception:
                        continue
        except Exception:
            pass
        if not items:
            import subprocess
            try:
                r = subprocess.run(["net", "view", p], capture_output=True, timeout=20,
                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                raw = r.stdout or b""
                text = ""
                for enc in ('cp950', 'utf-8', 'big5', 'cp437', 'mbcs'):
                    try:
                        text = raw.decode(enc); break
                    except Exception:
                        continue
                for line in text.splitlines():
                    if ('Disk' in line) or ('磁碟' in line):
                        parts = line.split()
                        if parts:
                            items.append((parts[0], p.rstrip("\\") + "\\" + parts[0]))
            except Exception:
                pass
        return items

    def refresh_folder(self):
        """重新整理資料夾"""
        self.load_folder_contents()

    def toggle_view_mode(self):
        """切換檢視模式"""
        if self.view_mode == "details":
            self.view_mode = "list"
            self.file_tree["show"] = "tree"
        else:
            self.view_mode = "details"
            self.file_tree["show"] = "tree headings"
        self.load_folder_contents()

    def navigate_to_path(self, event=None):
        """導航到指定路徑"""
        new_path = self.path_var.get().strip()
        # 特殊位置：本機 / 網路 / UNC 電腦根（這些 os.path.isdir 會回 False）
        if new_path in ("本機", "網路") or self._is_unc_computer_root(new_path):
            self.go_to_folder(new_path)
            return
        if os.path.exists(new_path) and os.path.isdir(new_path):
            self.go_to_folder(new_path)
        else:
            messagebox.showerror("錯誤", "路徑不存在或不是資料夾")
            self.path_var.set(self.current_folder)

    def is_hidden_file(self, filepath):
        """檢測文件或資料夾是否為隱藏"""
        name = os.path.basename(filepath)

        # 1. 檢查是否以點號開頭（Unix/Linux/Mac 隱藏文件）
        if name.startswith('.'):
            return True

        # 2. Windows 系統：檢查文件屬性
        if os.name == 'nt':  # Windows
            try:
                import ctypes
                attrs = ctypes.windll.kernel32.GetFileAttributesW(filepath)
                # FILE_ATTRIBUTE_HIDDEN = 0x2
                if attrs != -1 and attrs & 0x2:
                    return True
            except:
                pass

        return False

    def load_folder_contents(self):
        """載入資料夾內容 - 加入排序功能和隱藏文件過濾"""
        # 清空樹狀檢視
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        # 更新路徑
        self.path_var.set(self.current_folder)

        try:
            items = []

            # 🔧 特殊處理：「網路」→ 列出區網電腦（較慢，先給等待提示）
            if self.current_folder == "網路":
                self.file_tree.insert("", "end",
                                      text="🔄 正在搜尋網路電腦，請稍候（約 5～20 秒）...",
                                      values=("-", "", "-"))
                if hasattr(self, 'file_count_label'):
                    self.file_count_label.config(text="搜尋網路中，請稍候...")
                self._set_busy(True)
                try:
                    comps = self._net_list_computers()
                finally:
                    self._set_busy(False)
                for _it in self.file_tree.get_children():  # 清掉等待提示
                    self.file_tree.delete(_it)
                for comp in comps:
                    self.file_tree.insert("", "end", text=f"💻 {comp}",
                                          values=("-", "電腦", "-"), tags=("netcomp", comp))
                if not comps:
                    self.file_tree.insert("", "end",
                                          text="（找不到網路電腦，可在位置列輸入 \\\\電腦名 後按前往）",
                                          values=("-", "", "-"))
                if hasattr(self, 'file_count_label'):
                    self.file_count_label.config(text=f"{len(comps)} 台電腦")
                return

            # 🔧 特殊處理：UNC 電腦根（\\TABLE1）→ 用 Shell COM 列共用資料夾
            #    （os.listdir 對電腦根會 WinError 67；連線也較慢，先給等待提示）
            if self._is_unc_computer_root(self.current_folder):
                self.file_tree.insert("", "end",
                                      text="🔄 正在連線並讀取共用資料夾，請稍候...",
                                      values=("-", "", "-"))
                if hasattr(self, 'file_count_label'):
                    self.file_count_label.config(text="連線網路電腦中，請稍候...")
                self._set_busy(True)
                try:
                    shares = self._net_list_shares(self.current_folder)
                finally:
                    self._set_busy(False)
                for _it in self.file_tree.get_children():  # 清掉等待提示
                    self.file_tree.delete(_it)
                for nm, _full in shares:
                    self.file_tree.insert("", "end", text=f"📁 {nm}",
                                          values=("-", "共用資料夾", "-"))
                if not shares:
                    self.file_tree.insert("", "end",
                                          text="（無法列出共用資料夾，可輸入 \\\\電腦\\共用名 後按前往）",
                                          values=("-", "", "-"))
                if hasattr(self, 'file_count_label'):
                    self.file_count_label.config(text=f"{len(shares)} 個共用資料夾")
                return

            # 🔧 特殊處理：如果是 "本機"，顯示所有磁碟機
            if self.current_folder == "本機":
                import string
                from ctypes import windll

                # 取得所有可用磁碟機
                drives = []
                bitmask = windll.kernel32.GetLogicalDrives()
                for letter in string.ascii_uppercase:
                    if bitmask & 1:
                        drive_path = f"{letter}:\\"
                        try:
                            # 嘗試取得磁碟機資訊
                            drive_type = windll.kernel32.GetDriveTypeW(drive_path)
                            # 2 = 可移除磁碟(USB), 3 = 本機磁碟, 4 = 網路磁碟, 5 = 光碟機
                            if drive_type in [2, 3, 4, 5]:
                                drives.append((letter, drive_path, drive_type))
                        except:
                            pass
                    bitmask >>= 1

                # 插入磁碟機到樹狀檢視
                for letter, drive_path, drive_type in drives:
                    if drive_type == 2:
                        drive_icon = "🔌"  # USB 隨身碟
                        drive_name = f"{drive_icon} 卸除式磁碟 ({letter}:)"
                    elif drive_type == 3:
                        drive_icon = "💾"  # 本機磁碟
                        drive_name = f"{drive_icon} 本機磁碟 ({letter}:)"
                    elif drive_type == 4:
                        drive_icon = "🌐"  # 網路磁碟
                        drive_name = f"{drive_icon} 網路磁碟 ({letter}:)"
                    else:  # drive_type == 5
                        drive_icon = "💿"  # 光碟機
                        drive_name = f"{drive_icon} 光碟機 ({letter}:)"

                    self.file_tree.insert("", "end", text=drive_name,
                                        values=("-", "磁碟機", "-"), tags=("drive", drive_path))

                # 更新狀態列
                if hasattr(self, 'file_count_label'):
                    self.file_count_label.config(text=f"{len(drives)} 個磁碟機")
                return

            # 載入資料夾
            for item in os.listdir(self.current_folder):
                item_path = os.path.join(self.current_folder, item)

                # 🔧 過濾隱藏文件和資料夾
                if self.is_hidden_file(item_path):
                    continue

                if os.path.isdir(item_path):
                    stat = os.stat(item_path)
                    modified = time.strftime('%Y-%m-%d %H:%M', time.localtime(stat.st_mtime))
                    items.append(("folder", item, item_path, "資料夾", "-", modified))

            # 載入PDF檔案
            for item in os.listdir(self.current_folder):
                item_path = os.path.join(self.current_folder, item)

                # 🔧 過濾隱藏文件
                if self.is_hidden_file(item_path):
                    continue

                if os.path.isfile(item_path) and item.lower().endswith('.pdf'):
                    stat = os.stat(item_path)
                    size = self.format_file_size(stat.st_size)
                    size_bytes = stat.st_size  # 用於排序的原始大小
                    modified = time.strftime('%Y-%m-%d %H:%M', time.localtime(stat.st_mtime))
                    items.append(("file", item, item_path, "PDF檔案", size, modified, size_bytes))
            
            # 🔧 根據選擇的欄位排序
            if hasattr(self, 'sort_column') and self.sort_column:
                if self.sort_column == "name":
                    items.sort(key=lambda x: x[1].lower(), reverse=self.sort_reverse)
                elif self.sort_column == "type":
                    items.sort(key=lambda x: (x[0] == "file", x[3]), reverse=self.sort_reverse)
                elif self.sort_column == "size":
                    items.sort(key=lambda x: x[6] if len(x) > 6 else 0, reverse=self.sort_reverse)
                elif self.sort_column == "modified":
                    items.sort(key=lambda x: x[5], reverse=self.sort_reverse)
            else:
                # 🔧 預設排序：按修改日期降序（最新的在最上面）
                items.sort(key=lambda x: x[5], reverse=True)

            # 插入到樹狀檢視
            # 🔧 修正：按新順序插入 values (修改日期, 類型, 大小)

            # 🔧 新增：添加「..」回上層目錄的項目（不是根目錄時）
            parent_path = os.path.dirname(self.current_folder)
            # 檢查是否有上層目錄（不是磁碟機根目錄）
            if parent_path and parent_path != self.current_folder:
                self.file_tree.insert("", "end", text="📂 ..",
                                    values=("-", "上層目錄", "-"), tags=("parent",))

            pdf_count = 0
            for item in items:
                if item[0] == "folder":
                    self.file_tree.insert("", "end", text=f"📁 {item[1]}",
                                        values=(item[5], "資料夾", "-"), tags=("folder",))
                else:  # file
                    self.file_tree.insert("", "end", text=f"📄 {item[1]}",
                                        values=(item[5], "PDF檔案", item[4]), tags=("file",))
                    pdf_count += 1
            
            # 更新狀態列
            folder_count = len([i for i in items if i[0] == "folder"])
            self.file_count_label.config(text=f"{folder_count} 個資料夾, {pdf_count} 個PDF檔案")
            
        except Exception as e:
            messagebox.showerror("錯誤", f"載入資料夾失敗: {e}")

    def format_file_size(self, size):
        """格式化檔案大小"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} TB"

    def _kb(self, func):
        """鍵盤快捷鍵共用：執行 func 並吃掉事件（避免 Treeview 預設行為）。"""
        try:
            func()
        except Exception as e:
            print(f"鍵盤操作失敗: {e}", flush=True)
        return "break"

    def on_tree_enter(self, event=None):
        """Enter：開啟選取項目（資料夾→進入；檔案→預覽），等同雙擊。"""
        self.on_file_double_click(event)
        return "break"

    def on_tree_backspace(self, event=None):
        """Backspace：回上一層資料夾。"""
        self.go_parent_folder()
        return "break"

    def on_file_double_click(self, event):
        """雙擊事件"""
        selection = self.file_tree.selection()
        if not selection:
            return

        item = selection[0]
        item_text = self.file_tree.item(item, 'text')
        item_tags = self.file_tree.item(item, 'tags')

        # 🔧 處理磁碟機點擊（從「本機」視圖進入磁碟機）
        if 'drive' in item_tags:
            self.go_to_folder(item_tags[1])  # tags=("drive", drive_path)
        # 🔧 處理網路電腦點擊（從「網路」視圖進入電腦看共用資料夾）
        elif 'netcomp' in item_tags:
            self.go_to_folder(item_tags[1])  # tags=("netcomp", \\電腦)
        # 🔧 處理「..」回上層目錄
        elif 'parent' in item_tags or item_text == "📂 ..":
            parent_path = os.path.dirname(self.current_folder)
            if parent_path and parent_path != self.current_folder:
                self.go_to_folder(parent_path)
        elif item_text.startswith("📁"):
            folder_name = item_text[2:]
            self.go_to_folder(os.path.join(self.current_folder, folder_name))
        elif item_text.startswith("📄"):
            self.on_file_single_click(event)

    def on_file_single_click(self, event):
        """單擊事件 - 預覽功能"""
        import time

        print("🖱️ 調試：檔案單擊事件觸發", flush=True)

        # 🔧 延遲50毫秒讓選擇狀態更新
        def delayed_action():
            selection = self.file_tree.selection()
            print(f"🔍 調試：延遲後選中項目數量: {len(selection)}", flush=True)

            if selection:
                item = selection[0]
                item_text = self.file_tree.item(item, 'text')
                print(f"🔍 調試：點擊項目文字: {item_text}", flush=True)

                # 正常的單擊預覽功能
                if item_text.startswith("📄"):  # 只處理PDF檔案
                    filename = item_text[2:]  # 移除 "📄 " 前綴
                    file_path = os.path.join(self.current_folder, filename)

                    print(f"🔍 調試：檔案名稱: {filename}", flush=True)
                    print(f"🔍 調試：完整路徑: {file_path}", flush=True)
                    print(f"🔍 調試：current_folder: {self.current_folder}", flush=True)

                    # 設定當前選擇
                    self.current_selection = [file_path]
                    print(f"✅ 調試：設定 current_selection: {self.current_selection}", flush=True)

                    # 呼叫預覽功能
                    print("🔍 調試：準備呼叫 smart_preview", flush=True)
                    self.smart_preview()

                # 更新選擇資訊
                if hasattr(self, 'selected_info_label'):
                    self.selected_info_label.config(text=f"已選擇: {len(selection)} 個項目")
            else:
                print("❌ 調試：延遲後仍沒有選中任何項目", flush=True)

            # 🔧 新增這行
            self.master.after(100, self.on_file_selection_changed)

        # 延遲執行
        self.master.after(50, delayed_action)

    def _show_column_menu(self, event):
        """在欄位標題上按右鍵：勾選要顯示的欄位（類型／大小）。"""
        menu = tk.Menu(self.master, tearoff=0, font=("微軟正黑體", 11),
                       bg="white", fg="black", activebackground="#E3F2FD")
        labels = {"modified": "📅 修改日期", "type": "📁 類型", "size": "📏 大小"}
        for col in self.all_value_columns:
            menu.add_checkbutton(label=labels.get(col, col),
                                 variable=self._col_vars[col],
                                 command=lambda c=col: self.toggle_column(c))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def toggle_column(self, col):
        """依固定順序重建顯示欄位。"""
        visible = [c for c in self.all_value_columns if self._col_vars[c].get()]
        self.visible_columns = visible
        try:
            self.file_tree.config(displaycolumns=tuple(visible))
        except Exception as e:
            print(f"切換欄位失敗: {e}", flush=True)

    def show_context_menu(self, event):
        """顯示右鍵選單（在欄位標題上→欄位顯示選單；在檔案上→操作選單）"""
        # 🔧 在欄位標題列按右鍵 → 顯示「顯示欄位」勾選選單
        try:
            region = self.file_tree.identify_region(event.x, event.y)
        except Exception:
            region = ""
        if region in ("heading", "separator"):
            self._show_column_menu(event)
            return

        context_menu = tk.Menu(self.master, tearoff=0)
        
        context_menu.configure(
            font=("微軟正黑體", 10),
            relief="solid",
            borderwidth=1,
            activeborderwidth=1,
            bg="white",
            fg="black",
            activebackground="#E3F2FD",
            activeforeground="black"
        )
        
        selection = self.file_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        item_text = self.file_tree.item(item, 'text')
        
        if item_text.startswith("📄"):  # PDF檔案
            context_menu.add_command(label="📖 預覽", command=self.preview_selected_file)
            context_menu.add_command(label="📋 複製", command=self.copy_selected_files)
            context_menu.add_command(label="📄 剪下", command=self.cut_selected_files)        # 👈 改用 📄
            context_menu.add_command(label="📝 重新命名", command=self.rename_selected_file)
            context_menu.add_command(label="❌ 刪除", command=self.delete_selected_files)      # 👈 改用 ❌
            context_menu.add_command(label="ℹ️ 內容", command=self.show_file_properties)
            
        elif item_text.startswith("📁"):  # 資料夾
            context_menu.add_command(label="📂 開啟", command=self.open_selected_folder)
            context_menu.add_command(label="📋 複製", command=self.copy_selected_files)
            context_menu.add_command(label="📁 剪下", command=self.cut_selected_files)        # 👈 改用 📁
            context_menu.add_command(label="📝 重新命名", command=self.rename_selected_file)
            context_menu.add_command(label="❌ 刪除", command=self.delete_selected_files)      # 👈 改用 ❌
            context_menu.add_command(label="ℹ️ 內容", command=self.show_file_properties)
        
        # 共同選單項目
        if hasattr(self, 'clipboard_files'):
            context_menu.add_command(label="📌 貼上", command=self.paste_files)
        
        # 顯示選單
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def preview_selected_file(self):
        """預覽選中的檔案"""
        self.smart_preview()

    def _push_undo(self, action):
        """把可復原的動作推入復原堆疊（最多保留 20 筆）。"""
        if not hasattr(self, 'undo_stack'):
            self.undo_stack = []
        self.undo_stack.append(action)
        if len(self.undo_stack) > 20:
            self.undo_stack.pop(0)

    def undo_last_action(self):
        """Ctrl+Z：復原上一個檔案動作（貼上／剪下／重新命名）。"""
        if not getattr(self, 'undo_stack', None):
            self.show_custom_messagebox("復原", "沒有可復原的動作")
            return
        action = self.undo_stack.pop()
        import shutil
        try:
            if action['type'] == 'rename':
                # 把新檔名改回舊檔名
                if os.path.exists(action['new']):
                    os.rename(action['new'], action['old'])
            elif action['type'] == 'paste':
                # 反向處理：複製的刪掉、剪下的搬回原處
                for op in reversed(action['ops']):
                    if op[0] == 'copy':
                        p = op[1]
                        if os.path.isdir(p):
                            shutil.rmtree(p, ignore_errors=True)
                        elif os.path.exists(p):
                            os.remove(p)
                    elif op[0] == 'move':
                        src, dst = op[1], op[2]
                        if os.path.exists(dst):
                            shutil.move(dst, src)
            elif action['type'] == 'delete':
                paths = action.get('paths', [])
                restored = self._restore_from_recycle_bin(paths)
                if restored > 0:
                    self.load_folder_contents()
                    self.show_custom_messagebox("復原", f"已從資源回收筒還原 {restored} 個項目")
                else:
                    if self.show_custom_messagebox(
                            "復原",
                            "無法自動還原（項目可能已不在資源回收筒）。\n\n是否開啟資源回收筒手動還原？",
                            "yesno"):
                        self.open_recycle_bin()
                return
            self.load_folder_contents()
            self.show_custom_messagebox("復原", "已復原上一個動作")
        except Exception as e:
            self.show_custom_messagebox("錯誤", f"復原失敗: {e}")

    def _restore_from_recycle_bin(self, original_paths):
        """用 Windows Shell COM 從資源回收筒還原指定原始路徑的項目，回傳還原數量。"""
        restored = 0
        try:
            import win32com.client
            shell = win32com.client.Dispatch("Shell.Application")
            recycle = shell.NameSpace(10)  # ssfBITBUCKET = 資源回收筒
            if not recycle:
                return 0
            targets_full = set(os.path.normcase(os.path.abspath(p)) for p in original_paths)
            targets_base = set(os.path.normcase(os.path.basename(p)) for p in original_paths)

            def do_restore(it):
                # 找「還原 / Restore」動作並執行
                for verb in it.Verbs():
                    try:
                        nm = str(verb.Name).replace('&', '')
                    except Exception:
                        continue
                    if ('還原' in nm) or ('復原' in nm) or (nm.lower() == 'restore'):
                        verb.DoIt()
                        return True
                return False

            # 先用「完整原始路徑」精準比對（GetDetailsOf 第 1 欄＝原始位置）
            for it in list(recycle.Items()):
                try:
                    name = str(it.Name)
                    try:
                        orig_folder = recycle.GetDetailsOf(it, 1)
                    except Exception:
                        orig_folder = ""
                    if orig_folder:
                        full = os.path.normcase(os.path.abspath(os.path.join(orig_folder, name)))
                        if full in targets_full and do_restore(it):
                            restored += 1
                except Exception:
                    continue
            # 後備：完整路徑都沒中時，改用「檔名」比對（只還原剛剛刪掉的那些名字）
            if restored == 0:
                for it in list(recycle.Items()):
                    try:
                        if os.path.normcase(str(it.Name)) in targets_base and do_restore(it):
                            restored += 1
                    except Exception:
                        continue
        except Exception as e:
            print(f"還原資源回收筒失敗: {e}", flush=True)
        return restored

    def open_recycle_bin(self):
        """開啟資源回收筒。"""
        try:
            os.startfile("shell:RecycleBinFolder")
        except Exception as e:
            self.show_custom_messagebox("錯誤", f"無法開啟資源回收筒：{e}")

    def copy_selected_files(self):
        """複製選中的檔案"""
        selection = self.file_tree.selection()
        if not selection:
            return
        
        self.clipboard_files = []
        self.clipboard_operation = "copy"
        
        for item in selection:
            item_text = self.file_tree.item(item, 'text')
            if item_text.startswith("📄") or item_text.startswith("📁"):
                filename = item_text[2:]  # 移除前綴
                file_path = os.path.join(self.current_folder, filename)
                self.clipboard_files.append(file_path)

        # 一般檔案總管複製是無聲的，不跳彈窗；用底部狀態列輕提示即可
        if hasattr(self, 'file_count_label'):
            self.file_count_label.config(text=f"已複製 {len(self.clipboard_files)} 個項目")

    def cut_selected_files(self):
        """剪下選中的檔案"""
        selection = self.file_tree.selection()
        if not selection:
            return
        
        self.clipboard_files = []
        self.clipboard_operation = "cut"
        
        for item in selection:
            item_text = self.file_tree.item(item, 'text')
            if item_text.startswith("📄") or item_text.startswith("📁"):
                filename = item_text[2:]  # 移除前綴
                file_path = os.path.join(self.current_folder, filename)
                self.clipboard_files.append(file_path)

        # 剪下也不跳彈窗，僅在底部狀態列輕提示
        if hasattr(self, 'file_count_label'):
            self.file_count_label.config(text=f"已剪下 {len(self.clipboard_files)} 個項目（貼上後生效）")

    def paste_files(self):
        """貼上檔案 - 加入完整的檔案處理邏輯"""
        if not hasattr(self, 'clipboard_files') or not self.clipboard_files:
            self.show_custom_messagebox("貼上", "沒有可貼上的項目")
            return
        
        import shutil

        undo_ops = []  # 記錄本次操作，供 Ctrl+Z 復原
        try:
            for file_path in self.clipboard_files:
                filename = os.path.basename(file_path)
                target_path = os.path.join(self.current_folder, filename)
                overwrote = False

                if os.path.exists(target_path):
                    # 🔧 檔案已存在，顯示處理選項
                    choice = self.show_file_conflict_dialog(filename)

                    if choice == "skip":
                        continue
                    elif choice == "rename":
                        target_path = self.get_unique_filename(target_path)
                    elif choice == "overwrite":
                        overwrote = True  # 覆蓋過的目標不納入復原（避免誤刪原檔）
                    elif choice == "cancel":
                        return  # 取消整個操作
                    else:
                        continue  # 預設跳過

                # 執行檔案操作
                if self.clipboard_operation == "copy":
                    if os.path.isdir(file_path):
                        shutil.copytree(file_path, target_path, dirs_exist_ok=True)
                    else:
                        shutil.copy2(file_path, target_path)
                    if not overwrote:
                        undo_ops.append(('copy', target_path))
                elif self.clipboard_operation == "cut":
                    shutil.move(file_path, target_path)
                    if not overwrote:
                        undo_ops.append(('move', file_path, target_path))

            if self.clipboard_operation == "cut":
                self.clipboard_files = []

            if undo_ops:
                self._push_undo({'type': 'paste', 'ops': undo_ops})

            self.load_folder_contents()
            self.show_custom_messagebox("完成", "貼上操作完成")
            
        except Exception as e:
            self.show_custom_messagebox("錯誤", f"貼上失敗: {e}")

    def show_file_conflict_dialog(self, filename):
        """顯示檔案衝突處理對話框"""
        dialog = tk.Toplevel(self.master)
        dialog.title("檔案已存在")
        _dw, _dh = 640, 460
        dialog.geometry(f"{_dw}x{_dh}")
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.resizable(False, False)

        # 居中顯示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (_dw // 2)
        y = (dialog.winfo_screenheight() // 2) - (_dh // 2)
        dialog.geometry(f"{_dw}x{_dh}+{x}+{y}")
        
        result = {"choice": "skip"}
        
        # 標題
        title_label = tk.Label(dialog,
                            text="檔案已存在",
                            font=("微軟正黑體", 20, "bold"),
                            fg="#e74c3c")
        title_label.pack(pady=(22, 14))

        # 檔案名稱
        file_label = tk.Label(dialog,
                            text=f"檔案名稱：{filename}",
                            font=("微軟正黑體", 15),
                            wraplength=580)
        file_label.pack(pady=(0, 18))

        # 詢問文字
        question_label = tk.Label(dialog,
                                text="目標資料夾中已經有同名檔案，請選擇處理方式：",
                                font=("微軟正黑體", 13))
        question_label.pack(pady=(0, 24))

        # 按鈕框架
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)

        # 🔧 四個選項按鈕
        def set_choice(choice):
            result["choice"] = choice
            dialog.destroy()

        _bfont = ("微軟正黑體", 14, "bold")
        # 覆蓋按鈕
        overwrite_btn = tk.Button(button_frame,
                                text="🔄 覆蓋檔案",
                                command=lambda: set_choice("overwrite"),
                                font=_bfont,
                                bg="#e74c3c", fg="white",
                                width=15, height=2, cursor="hand2")
        overwrite_btn.grid(row=0, column=0, padx=10, pady=10)

        # 重新命名按鈕
        rename_btn = tk.Button(button_frame,
                            text="📝 自動重新命名",
                            command=lambda: set_choice("rename"),
                            font=_bfont,
                            bg="#3498db", fg="white",
                            width=15, height=2, cursor="hand2")
        rename_btn.grid(row=0, column=1, padx=10, pady=10)

        # 跳過按鈕
        skip_btn = tk.Button(button_frame,
                            text="⏭️ 跳過此檔案",
                            command=lambda: set_choice("skip"),
                            font=_bfont,
                            bg="#95a5a6", fg="white",
                            width=15, height=2, cursor="hand2")
        skip_btn.grid(row=1, column=0, padx=10, pady=10)

        # 取消按鈕
        cancel_btn = tk.Button(button_frame,
                            text="❌ 取消操作",
                            command=lambda: set_choice("cancel"),
                            font=_bfont,
                            bg="#7f8c8d", fg="white",
                            width=15, height=2, cursor="hand2")
        cancel_btn.grid(row=1, column=1, padx=10, pady=10)

        # 預設選擇重新命名 + 鍵盤：Enter=重新命名，Esc=取消
        rename_btn.focus_set()
        dialog.bind('<Return>', lambda e: set_choice("rename"))
        dialog.bind('<Escape>', lambda e: set_choice("cancel"))

        # 處理關閉視窗
        dialog.protocol("WM_DELETE_WINDOW", lambda: set_choice("cancel"))
        
        # 等待用戶選擇
        dialog.wait_window()
        
        return result["choice"]

    def get_unique_filename(self, original_path):
        """產生唯一的檔案名稱"""
        directory = os.path.dirname(original_path)
        filename = os.path.basename(original_path)
        name, ext = os.path.splitext(filename)
        
        # 🔧 先嘗試加上 " - 複製"
        new_name = f"{name} - 複製{ext}"
        new_path = os.path.join(directory, new_name)
        
        # 如果還是存在，就加上數字
        counter = 1
        while os.path.exists(new_path):
            new_name = f"{name} - 複製({counter}){ext}"
            new_path = os.path.join(directory, new_name)
            counter += 1
        
        return new_path

    def rename_selected_file(self):
        """重新命名選中的檔案"""
        selection = self.file_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        item_text = self.file_tree.item(item, 'text')
        old_name = item_text[2:]  # 移除前綴
        
        # 創建重新命名對話框
        dialog = tk.Toplevel(self.master)
        dialog.title("重新命名")
        dialog.geometry("400x150")
        dialog.transient(self.master)
        dialog.grab_set()
        
        # 居中顯示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (150 // 2)
        dialog.geometry(f"400x150+{x}+{y}")
        
        tk.Label(dialog, text="新名稱:", font=("微軟正黑體", 12)).pack(pady=10)
        
        name_var = tk.StringVar(value=old_name)
        entry = tk.Entry(dialog, textvariable=name_var, font=("微軟正黑體", 12), width=40)
        entry.pack(pady=10)
        entry.select_range(0, tk.END)
        entry.focus()
        
        def do_rename():
            raw_name = name_var.get().strip()
            if not raw_name or raw_name == old_name:
                dialog.destroy()
                return
            
            # 🔧 完整的檔案名稱清理
            def clean_filename(filename):
                # 移除換行和回車
                filename = filename.replace('\n', ' ').replace('\r', ' ')
                
                # 移除檔案系統不允許的字符
                invalid_chars = ['<', '>', ':', '"', '|', '?', '*', '/', '\\']
                for char in invalid_chars:
                    filename = filename.replace(char, '')
                
                # 處理特殊字符
                filename = filename.replace(',', '-')  # 逗號改連字號
                filename = filename.replace(';', '-')  # 分號改連字號
                
                # 清理多餘空格
                filename = ' '.join(filename.split())
                
                # 移除開頭結尾的點和空格
                filename = filename.strip('. ')
                
                return filename
            
            new_name = clean_filename(raw_name)
            
            # 檢查清理後的名稱
            if not new_name or len(new_name) > 200:  # 留點空間給副檔名
                dialog.grab_release()
                dialog.destroy()
                self.show_custom_messagebox("錯誤", "檔案名稱無效、為空或過長")
                return
            
            # 確保有正確的副檔名
            if not new_name.lower().endswith('.pdf'):
                new_name += '.pdf'
            
            old_path = os.path.join(self.current_folder, old_name)
            new_path = os.path.join(self.current_folder, new_name)
            
            dialog.grab_release()
            
            try:
                os.rename(old_path, new_path)
                self._push_undo({'type': 'rename', 'old': old_path, 'new': new_path})
                dialog.destroy()
                self.load_folder_contents()
                self.show_custom_messagebox("完成", f"已重新命名為:\n{new_name}")
            except Exception as e:
                dialog.destroy()
                self.show_custom_messagebox("錯誤", f"重新命名失敗: {e}")
        
        def cancel_rename():
            """取消重新命名"""
            dialog.grab_release()
            dialog.destroy()

        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="確定", command=do_rename,
                font=("微軟正黑體", 12)).pack(side="left", padx=5)
        tk.Button(button_frame, text="取消", command=cancel_rename,
                font=("微軟正黑體", 12)).pack(side="left", padx=5)

        # 🔧 綁定鍵盤快捷鍵
        entry.bind('<Return>', lambda e: do_rename())
        entry.bind('<Escape>', lambda e: cancel_rename())
        dialog.bind('<Escape>', lambda e: cancel_rename())


    def delete_selected_files(self):
        """刪除選中的檔案"""
        selection = self.file_tree.selection()
        if not selection:
            return
        
        file_names = []
        file_paths = []
        
        for item in selection:
            item_text = self.file_tree.item(item, 'text')
            filename = item_text[2:]  # 移除前綴
            file_path = os.path.join(self.current_folder, filename)
            file_names.append(filename)
            file_paths.append(file_path)
        
        # 確認刪除
        result = self.show_custom_messagebox(
            "確認刪除", 
            f"確定要刪除這 {len(file_names)} 個項目嗎？\n\n" + "\n".join(file_names[:5]) + 
            ("..." if len(file_names) > 5 else ""),
            "yesno"
        )
        
        if not result:
            return
        
        try:
            for file_path in file_paths:
                if HAS_SEND2TRASH:
                    # 移到回收桶
                    send2trash.send2trash(file_path)
                else:
                    # 直接刪除
                    if os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                    else:
                        os.remove(file_path)
            
            if HAS_SEND2TRASH:
                self._push_undo({'type': 'delete', 'paths': file_paths})
            self.load_folder_contents()
            action = "移至回收桶" if HAS_SEND2TRASH else "永久刪除"
            self.show_custom_messagebox("完成", f"已{action} {len(file_paths)} 個項目")
            
        except Exception as e:
            self.show_custom_messagebox("錯誤", f"刪除失敗: {e}")

    def open_selected_folder(self):
        """開啟選中的資料夾"""
        selection = self.file_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        item_text = self.file_tree.item(item, 'text')
        
        if item_text.startswith("📁"):
            folder_name = item_text[2:]  # 移除📁前綴
            new_path = os.path.join(self.current_folder, folder_name)
            self.current_folder = new_path
            self.load_folder_contents()

    def show_file_properties(self):
        """顯示檔案屬性"""
        selection = self.file_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        item_text = self.file_tree.item(item, 'text')
        filename = item_text[2:]  # 移除前綴
        file_path = os.path.join(self.current_folder, filename)
        
        try:
            stat = os.stat(file_path)
            size = self.format_file_size(stat.st_size)
            created = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat.st_ctime))
            modified = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat.st_mtime))
            
            properties = f"""檔案名稱: {filename}
    檔案路徑: {file_path}
    檔案大小: {size}
    建立時間: {created}
    修改時間: {modified}
    檔案類型: {'資料夾' if os.path.isdir(file_path) else 'PDF檔案'}"""
            
            self.show_custom_messagebox("檔案屬性", properties)
            
        except Exception as e:
            self.show_custom_messagebox("錯誤", f"無法讀取檔案屬性: {e}")

    def show_help(self):
        """顯示使用說明（美化版）"""
        help_window = tk.Toplevel(self.master)
        help_window.title("📖 使用說明")
        help_window.geometry("700x600")
        help_window.configure(bg=ModernStyle.BG_LIGHT)
        help_window.transient(self.master)
        
        # 🔧 加入置中程式碼
        help_window.update_idletasks()  # 這行很重要
        screen_width = help_window.winfo_screenwidth()
        screen_height = help_window.winfo_screenheight()
        window_width = 700
        window_height = 600
        
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        help_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 創建捲動文本框
        help_frame = tk.Frame(help_window, bg=ModernStyle.BG_LIGHT)
        help_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        help_text = scrolledtext.ScrolledText(help_frame, 
                                            font=ModernStyle.FONT_MAIN, 
                                            wrap=tk.WORD,
                                            bg=ModernStyle.BG_LIGHT,
                                            fg=ModernStyle.TEXT_PRIMARY,
                                            relief="solid",
                                            borderwidth=1)
        help_text.pack(fill="both", expand=True)
        
        help_content = f"""

🏛️ 不動產登記謄本智慧處理系統 - 使用說明

{'='*52}

🎯 系統概述
本系統是專為處理不動產登記謄本設計的智慧化工具，支援第一類及第二類謄本，包括土地謄本和建物謄本的自動化解析與資料提取。

📊 當前狀態
處理器狀態: {'✅ 完整功能模式 (已導入原處理模組)' if self.using_original else '⚠️ 增強功能模式 (使用內建處理器)'}
支援格式: PDF
輸出格式: TXT報告、Excel表格、JSON資料

{'='*52}

🚀 快速開始

1️⃣ 檔案準備
   • 將PDF謄本檔案放入工作資料夾
   • 點擊「🔄 重新掃描」載入檔案清單
   • 或使用「📂 選擇資料夾」切換工作目錄

2️⃣ 預覽檢查  
   • 選擇要處理的PDF檔案
   • 點擊「🔍 快速預覽」查看內容
   • 檢查原始文本、清理後文本和謄本資訊

3️⃣ 選擇處理方式
   • ⚡ 簡略查詢: 適合基本資料提取（1-3個檔案）
   • 📋 批次處理: 適合任何數量的檔案處理（含單檔）

4️⃣ 匯出結果
   • 📄 TXT報告: 詳細的文字分析報告
   • 📈 Excel表格: 結構化資料表格
   • 🔗 JSON資料: 機器可讀的資料格式

{'='*52}

💡 使用技巧

預覽技巧:
- 先用快速預覽檢查檔案內容和格式
- 查看謄本資訊卡片確認類型識別正確
- 注意統計資訊了解文件複雜度

處理技巧:
- 批次處理支援單檔或多檔，操作更靈活
- 大量檔案可先抽樣簡略查詢測試
- 批次處理時注意觀察進度視窗

匯出技巧:
- TXT報告適合閱讀和列印
- Excel表格適合資料分析
- JSON格式適合程式化處理

{'='*52}

⚠️ 注意事項

檔案要求:
- 支援PDF格式的謄本檔案
- 建議檔案名稱包含識別資訊
- 確保PDF文字可正常提取

效能考量:
- 大檔案或大量檔案需要較長處理時間
- 建議先預覽再處理避免不必要等待
- 批次處理時避免同時執行其他高負載操作

品質提醒:
- {'完整功能模式提供最佳處理品質' if self.using_original else '當前為增強功能模式，如需完整功能請確認原模組導入'}
- 複雜版面或掃描品質較差的PDF可能影響識別準確度
- 建議處理後檢查關鍵資料的準確性

{'='*52}

🆘 常見問題

Q: 為什麼某些PDF無法處理？
A: 可能原因包括：PDF加密、掃描版PDF、版面過於複雜等

Q: 處理結果不準確怎麼辦？
A: 建議先使用預覽功能檢查，或嘗試不同處理模式

Q: 如何提高處理速度？
A: 選擇適當的處理模式，避免不必要的單檔處理

Q: 匯出的檔案在哪裡？
A: 點擊「📂 開啟輸出」按鈕或查看底部狀態列顯示的路徑

{'='*52}

📞 技術支援
如遇到技術問題，請檢查：
1. PDF檔案是否完整且可正常開啟
2. 系統設定中的路徑配置是否正確
3. 是否有足夠的磁碟空間用於輸出
4. 原處理模組是否正確導入

祝您使用愉快！ 🎉
        """
        
        help_text.insert("1.0", help_content)
        help_text.config(state="disabled")
        
        # 關閉按鈕
        ModernButton(help_window, "關閉說明", 
                    command=help_window.destroy, 
                    style="primary").pack(pady=10)

# ==================== 主程式入口 ====================
def _make_splash(root):
    """建立置中的「載入中」啟動畫面，回傳 (splash 視窗, 狀態文字 Label)。"""
    try:
        splash = tk.Toplevel(root)
        splash.overrideredirect(True)  # 無邊框
        sw, sh = 460, 200
        scrw = splash.winfo_screenwidth()
        scrh = splash.winfo_screenheight()
        splash.geometry(f"{sw}x{sh}+{(scrw - sw) // 2}+{(scrh - sh) // 2}")
        splash.configure(bg="#2C3E50")
        # 外框線
        border = tk.Frame(splash, bg="#1ABC9C")
        border.pack(fill="both", expand=True, padx=2, pady=2)
        inner = tk.Frame(border, bg="#2C3E50")
        inner.pack(fill="both", expand=True, padx=2, pady=2)

        tk.Label(inner, text="🏛  不動產登記謄本智慧處理工具",
                 font=("微軟正黑體", 16, "bold"),
                 bg="#2C3E50", fg="white").pack(pady=(38, 6))
        status = tk.Label(inner, text="正在啟動，請稍候...",
                          font=("微軟正黑體", 12),
                          bg="#2C3E50", fg="#ECF0F1")
        status.pack(pady=(6, 4))
        tk.Label(inner, text="（初次啟動載入辨識引擎較久，屬正常現象）",
                 font=("微軟正黑體", 9),
                 bg="#2C3E50", fg="#95A5A6").pack(pady=(2, 0))
        splash.update()
        return splash, status
    except Exception as e:
        print(f"啟動畫面建立失敗: {e}", flush=True)
        return None, None


def main():
    """主程式入口"""
    
    freeze_support()  # PyInstaller 打包支援
    print("🚀 啟動不動產登記謄本智慧處理系統（美化版GUI）", flush=True)
    print("="*60, flush=True)

    # 啟動美化版GUI
    try:
        # 🔧 修復工作列圖示問題：必須在建立視窗之前設定 AppUserModelID
        try:
            import ctypes
            myappid = 'RealEstateTool.Application.1.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        # 🔧 先建立主視窗但隱藏，立刻顯示「載入中」啟動畫面，給使用者即時回饋
        root = tk.Tk()
        root.withdraw()
        splash, splash_status = _make_splash(root)

        def _boot(msg):
            print(msg, flush=True)
            try:
                if splash_status:
                    splash_status.config(text=msg)
                    splash.update()
            except Exception:
                pass

        # 檢查必要模組
        try:
            import pandas as pd
            _boot("正在準備資料元件...")
        except ImportError:
            print("❌ pandas: 未安裝", flush=True)
            try:
                splash.destroy()
            except Exception:
                pass
            messagebox.showerror("缺少模組", "缺少 pandas 模組\n請執行: pip install pandas")
            return

        print(f"📁 工作目錄: {PDF_DIR}", flush=True)
        print(f"📁 輸出目錄: {OUT_DIR}", flush=True)

        # 🔧 載入最耗時的處理引擎（OCR / 解析模組）——在啟動畫面底下進行
        global ProcessorClass, ExporterClass
        _boot("正在載入謄本辨識與解析引擎，請稍候...")
        ProcessorClass, ExporterClass = import_original_modules()

        # 建立主介面
        _boot("正在建立操作介面...")
        app = ModernTranscriptGUI(root)

        # 🔧 載入完成 → 關閉啟動畫面、顯示主視窗
        try:
            splash.destroy()
        except Exception:
            pass
        root.deiconify()
        try:
            root.lift()
            root.focus_force()
        except Exception:
            pass

        print("🎯 美化版GUI介面已啟動", flush=True)
        print("🎨 享受全新的現代化使用體驗！", flush=True)
        print("="*60, flush=True)

        root.mainloop()

    except Exception as e:
        print(f"❌ GUI啟動失敗: {e}", flush=True)
        messagebox.showerror("啟動失敗", f"GUI介面啟動失敗:\n{e}")

if __name__ == "__main__":
    main()

# ==================== 程式完成 ====================
# 
# 使用說明：
# 1. 將以上9個部分的程式碼依序複製到同一個Python檔案中
# 2. 確保已安裝必要的套件：
#    pip install PyMuPDF pandas tkinter
# 3. 將PDF謄本檔案放在程式所在目錄
# 4. 執行程式：python enhanced_transcript_gui.py
# 
# 功能特色：
# - 🎨 現代化美觀介面
# - 📄 智慧PDF預覽功能  
# - ⚡ 簡略查詢（原快速處理）
# - 🎯 單檔精確處理
# - 📋 批次自動化處理
# - 📤 多格式資料匯出
# - 🛠️ 完整工具支援
# 
# 版本資訊：v1.5b - 修復附屬建物全形數字與無空白格式提取（NFKC + pattern3 fallback）
# ==========================================