import tkinter as tk
from tkinter import simpledialog, messagebox
import json
import os

def load_config():
    """加載配置文件"""
    config_file = "config.json"
    if os.path.exists(config_file):
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                config = json.load(f)
            return config
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取配置文件時出錯: {e}")
    return {"username": "", "password": "", "card_cert": None}

def save_config(config, show_message=True):
    """保存配置

    Args:
        config: 配置字典
        show_message: 是否顯示成功訊息（預設 True，但在讀取JSON時會設為 False）
    """
    config_file = "config.json"
    try:
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

        # 設置文件權限（僅限 UNIX/Linux/Mac）
        if os.name != 'nt':
            try:
                os.chmod(config_file, 0o600)  # 只有擁有者可讀寫
            except Exception as e:
                if show_message:
                    messagebox.showwarning("警告", f"設置文件權限時出錯: {e}")

        # 🔥 只在需要時才顯示訊息（用於配置管理視窗）
        if show_message:
            # 使用自訂的大字體訊息框
            show_large_config_message("成功", "配置已保存")
        return True
    except Exception as e:
        if show_message:
            messagebox.showerror("錯誤", f"保存配置時出錯: {e}")
        return False

def show_large_config_message(title, message):
    """顯示大型訊息對話框（用於配置保存）"""
    msg_window = tk.Toplevel()
    msg_window.title(title)

    width = 400
    height = 250

    msg_window.geometry(f"{width}x{height}")
    msg_window.resizable(False, False)
    msg_window.grab_set()

    # 置中顯示
    msg_window.update_idletasks()
    screen_width = msg_window.winfo_screenwidth()
    screen_height = msg_window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    msg_window.geometry(f"{width}x{height}+{x}+{y}")

    # 標題
    title_label = tk.Label(msg_window, text=title,
                          font=("Microsoft JhengHei", 16, "bold"),
                          fg="#1976D2",
                          pady=15)
    title_label.pack()

    # 訊息內容
    content_frame = tk.Frame(msg_window)
    content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=10)

    msg_label = tk.Label(content_frame,
                        text=message,
                        font=("Microsoft JhengHei", 14),
                        justify=tk.LEFT,
                        wraplength=width-80)
    msg_label.pack(fill=tk.BOTH, expand=True)

    # 確定按鈕
    button_frame = tk.Frame(msg_window)
    button_frame.pack(side=tk.BOTTOM, pady=20)

    ok_btn = tk.Button(button_frame,
                      text="確定",
                      command=msg_window.destroy,
                      font=("Microsoft JhengHei", 12, "bold"),
                      bg='#4CAF50',
                      fg='white',
                      width=10,
                      height=2,
                      cursor="hand2")
    ok_btn.pack()

def open_config_manager(root, message_box, update_message):
    """打開配置管理視窗"""
    config = load_config()
    
    # 創建配置管理視窗
    config_window = tk.Toplevel(root)
    config_window.title("帳號配置管理")
    config_window.geometry("400x320")  # 適當的視窗大小
    config_window.resizable(False, False)
    
    # 確保這個視窗在主視窗上方
    config_window.transient(root)
    config_window.grab_set()
    
    # 配置字型設定
    title_font = ("Microsoft JhengHei", 16, "bold")  # 標題字型
    label_font = ("Microsoft JhengHei", 12)           # 標籤字型
    button_font = ("Microsoft JhengHei", 12)          # 按鈕字型
    
    # 顯示當前配置
    tk.Label(config_window, text="=== 配置管理 ===", font=title_font).pack(pady=10)
    
    # 主框架，包含所有內容
    main_frame = tk.Frame(config_window)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
    
    # 顯示用戶名
    user_frame = tk.Frame(main_frame)
    user_frame.pack(fill=tk.X, pady=5)
    tk.Label(user_frame, text=f"1. 用戶名: {config.get('username', '未設置')}", width=25, anchor="w", font=label_font).pack(side=tk.LEFT)
    tk.Button(user_frame, text="修改", command=lambda: modify_username(config_window, config, update_message), 
              font=button_font, width=6).pack(side=tk.RIGHT, padx=5)
    
    # 顯示密碼
    pass_frame = tk.Frame(main_frame)
    pass_frame.pack(fill=tk.X, pady=5)
    tk.Label(pass_frame, text=f"2. 密碼: {'******' if config.get('password') else '未設置'}", width=25, anchor="w", font=label_font).pack(side=tk.LEFT)
    tk.Button(pass_frame, text="修改", command=lambda: modify_password(config_window, config, update_message), 
              font=button_font, width=6).pack(side=tk.RIGHT, padx=5)
    
    # 顯示自然人憑證
    cert_frame = tk.Frame(main_frame)
    cert_frame.pack(fill=tk.X, pady=5)
    tk.Label(cert_frame, text=f"3. 自然人憑證: {'已設置' if config.get('card_cert') else '未設置'}", width=25, anchor="w", font=label_font).pack(side=tk.LEFT)
    tk.Button(cert_frame, text="修改", command=lambda: modify_certificate(config_window, config, update_message), 
              font=button_font, width=6).pack(side=tk.RIGHT, padx=5)
    
    # 清除配置按鈕
    clear_frame = tk.Frame(main_frame)
    clear_frame.pack(fill=tk.X, pady=5)
    tk.Label(clear_frame, text="4. 清除所有配置", width=25, anchor="w", font=label_font).pack(side=tk.LEFT)
    tk.Button(clear_frame, text="執行", command=lambda: clear_config(config_window, update_message), 
              font=button_font, width=6).pack(side=tk.RIGHT, padx=5)
    
    # 底部按鈕框架 - 獨立的框架用於居中放置關閉按鈕
    button_frame = tk.Frame(config_window)
    button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=15)  # 放在視窗底部
    
    # 關閉按鈕（置中）
    close_button = tk.Button(button_frame, text="關閉", command=config_window.destroy, 
                             font=button_font, width=6)
    # 使用pack居中放置按鈕
    close_button.pack(side=tk.TOP, anchor=tk.CENTER)
    
    # 更新消息框
    if update_message:
        update_message("已開啟配置管理視窗\n")

def modify_username(parent, config, update_message):
    """修改用戶名"""
    # 在這裡設定對話框字型
    dialog_font = ("Microsoft JhengHei", 12)
    
    # 使用自訂對話框以設定字型
    class CustomDialog(simpledialog.Dialog):
        def body(self, master):
            tk.Label(master, text="請輸入新的用戶名:", font=dialog_font).grid(row=0, column=0, padx=10, pady=5)
            self.entry = tk.Entry(master, font=dialog_font)
            self.entry.grid(row=0, column=1, padx=10, pady=5)
            self.entry.insert(0, config.get("username", ""))
            return self.entry
        
        def apply(self):
            self.result = self.entry.get()
    
    dialog = CustomDialog(parent, "輸入")
    new_username = dialog.result
    
    if new_username is not None:  # 用戶未取消
        config["username"] = new_username
        if save_config(config):
            if update_message:
                update_message(f"已更新用戶名為: {new_username}\n")
            # 重新加載視窗
            parent.destroy()
            open_config_manager(parent.master, None, update_message)

def modify_password(parent, config, update_message):
    """修改密碼"""
    # 在這裡設定對話框字型
    dialog_font = ("Microsoft JhengHei", 12)
    
    # 使用自訂對話框以設定字型
    class CustomDialog(simpledialog.Dialog):
        def body(self, master):
            tk.Label(master, text="請輸入新的密碼:", font=dialog_font).grid(row=0, column=0, padx=10, pady=5)
            self.entry = tk.Entry(master, font=dialog_font, show="*")
            self.entry.grid(row=0, column=1, padx=10, pady=5)
            return self.entry
        
        def apply(self):
            self.result = self.entry.get()
    
    dialog = CustomDialog(parent, "輸入")
    new_password = dialog.result
    
    if new_password is not None:  # 用戶未取消
        config["password"] = new_password
        if save_config(config):
            if update_message:
                update_message("已更新密碼\n")
            # 重新加載視窗
            parent.destroy()
            open_config_manager(parent.master, None, update_message)

def modify_certificate(parent, config, update_message):
    """修改自然人憑證資訊"""
    # 創建憑證資訊視窗
    cert_window = tk.Toplevel(parent)
    cert_window.title("設置自然人憑證")
    cert_window.geometry("320x180")
    cert_window.resizable(False, False)
    cert_window.transient(parent)
    cert_window.grab_set()
    
    # 設定字型
    label_font = ("Microsoft JhengHei", 12)
    entry_font = ("Microsoft JhengHei", 12)
    button_font = ("Microsoft JhengHei", 12)
    
    # 身分證號碼
    tk.Label(cert_window, text="身分證號碼:", font=label_font).grid(row=0, column=0, sticky="w", padx=10, pady=10)
    id_var = tk.StringVar()
    if config.get("card_cert") and config["card_cert"].get("id_no"):
        id_var.set(config["card_cert"]["id_no"])
    id_entry = tk.Entry(cert_window, textvariable=id_var, font=entry_font)
    id_entry.grid(row=0, column=1, padx=10, pady=10)
    
    # 憑證密碼
    tk.Label(cert_window, text="憑證密碼:", font=label_font).grid(row=1, column=0, sticky="w", padx=10, pady=10)
    pin_var = tk.StringVar()
    if config.get("card_cert") and config["card_cert"].get("pin"):
        pin_var.set(config["card_cert"]["pin"])
    pin_entry = tk.Entry(cert_window, textvariable=pin_var, show="*", font=entry_font)
    pin_entry.grid(row=1, column=1, padx=10, pady=10)
    
    # 確認按鈕
    def save_cert():
        id_no = id_var.get()
        pin = pin_var.get()
        
        if not id_no or not pin:
            messagebox.showwarning("警告", "身分證號碼和憑證密碼都不能為空")
            return
        
        config["card_cert"] = {
            "id_no": id_no,
            "pin": pin
        }
        
        if save_config(config):
            if update_message:
                update_message("已更新自然人憑證資訊\n")
            cert_window.destroy()
            # 重新加載視窗
            parent.destroy()
            open_config_manager(parent.master, None, update_message)
    
    # 取消按鈕
    def cancel():
        cert_window.destroy()
    
    button_frame = tk.Frame(cert_window)
    button_frame.grid(row=2, column=0, columnspan=2, pady=15)
    
    tk.Button(button_frame, text="確認", command=save_cert, font=button_font, width=6).pack(side=tk.LEFT, padx=10)
    tk.Button(button_frame, text="取消", command=cancel, font=button_font, width=6).pack(side=tk.LEFT, padx=10)

def clear_config(parent, update_message):
    """清除所有配置"""
    confirm = messagebox.askyesno("確認", "確定要清除所有配置嗎？", parent=parent)
    if confirm:
        config_file = "config.json"
        try:
            if os.path.exists(config_file):
                os.remove(config_file)
                messagebox.showinfo("成功", "配置已清除", parent=parent)
                if update_message:
                    update_message("已清除所有配置\n")
                parent.destroy()
        except Exception as e:
            messagebox.showerror("錯誤", f"清除配置時出錯: {e}", parent=parent)