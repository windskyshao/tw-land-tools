"""
基準目錄輔助模組
用於讓所有子程式都能正確找到 data.json 和工作資料夾的位置
"""
import argparse
import os
import sys

def get_base_dir():
    """
    取得基準目錄（執行檔所在目錄）

    🔥 關鍵修正：
    1. 優先使用 --base-dir 參數（主程式呼叫子程式時會傳遞）
    2. 如果沒有 --base-dir，才使用 sys.executable 推算
    3. 這樣可以確保子程式使用正確的主程式目錄
    """
    # 🔥 優先檢查 --base-dir 參數（無論打包或開發環境）
    parser = argparse.ArgumentParser()
    parser.add_argument('--base-dir', type=str, default=None, help='基準目錄')
    args, unknown = parser.parse_known_args()

    if args.base_dir:
        # 有傳遞 --base-dir 參數，直接使用
        return args.base_dir

    # 🔥 沒有 --base-dir 參數時，根據環境推算
    if hasattr(sys, '_MEIPASS'):  # PyInstaller 打包環境
        # 使用執行檔所在目錄，而不是 _MEIPASS 臨時目錄
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))

        # 如果路徑包含 python_embedded，往上一層找主程式目錄
        if 'python_embedded' in exe_dir:
            base_dir = os.path.dirname(exe_dir)
        # 如果路徑包含 _internal，往上一層
        elif '_internal' in exe_dir or exe_dir.endswith('_internal'):
            base_dir = os.path.dirname(exe_dir)
        else:
            base_dir = exe_dir

        return base_dir

    # 🔥 開發環境：使用腳本所在目錄
    import inspect
    frame = inspect.stack()[1]
    caller_file = frame[0].f_code.co_filename
    base_dir = os.path.dirname(os.path.abspath(caller_file))

    # 如果當前目錄是 _internal，則往上一層
    if os.path.basename(base_dir) == '_internal':
        base_dir = os.path.dirname(base_dir)

    return base_dir

# 全域變數：BASE_DIR
BASE_DIR = get_base_dir()

# 便利函數
def get_data_json_path():
    """取得 data.json 的完整路徑"""
    return os.path.join(BASE_DIR, 'data.json')

def get_work_folder(folder_name):
    """取得工作資料夾的完整路徑"""
    return os.path.join(BASE_DIR, folder_name)

def get_internal_dir():
    """
    取得 _internal 目錄的完整路徑

    在打包環境下：返回 sys._MEIPASS (臨時解壓目錄)
    在開發環境下：返回腳本所在目錄
    """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包環境：使用臨時解壓目錄
        return sys._MEIPASS
    else:
        # 開發環境：使用腳本所在目錄
        return os.path.dirname(os.path.abspath(__file__))

# 測試用
if __name__ == '__main__':
    print(f"基準目錄：{BASE_DIR}")
    print(f"data.json 路徑：{get_data_json_path()}")
    print(f"範例工作資料夾：{get_work_folder('大寮區赤崁段-2337')}")
