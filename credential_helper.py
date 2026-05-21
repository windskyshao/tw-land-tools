"""credential_helper.py — 統一帳密管理（Windows Credential Manager）

透過 keyring 套件把帳密存進 Windows 認證管理員。帳密綁 Windows 使用者帳號 + 機器，
換電腦或別人拿到 exe 都讀不到。

主要 API：
    get_or_prompt_credentials(service, label) — 取得帳密（若無則跳對話框讓使用者輸入並儲存）
    get_credentials(service)                  — 純讀取，無資料回傳 None
    save_credentials(service, user, pwd)      — 儲存
    delete_credentials(service)               — 刪除
"""
import json
import keyring

SERVICE_PREFIX = "DiziDataQuery_"


def _full_service(service):
    return SERVICE_PREFIX + service


def get_credentials(service):
    """取得 (username, password)；無資料時回傳 (None, None)。"""
    try:
        data = keyring.get_password(_full_service(service), "credentials")
        if data:
            d = json.loads(data)
            return d.get("username"), d.get("password")
    except Exception as e:
        print(f"[credential_helper] 讀取 {service} 失敗：{e}", flush=True)
    return None, None


def save_credentials(service, username, password):
    """儲存帳密；回傳是否成功。"""
    try:
        data = json.dumps({"username": username, "password": password})
        keyring.set_password(_full_service(service), "credentials", data)
        return True
    except Exception as e:
        print(f"[credential_helper] 儲存 {service} 失敗：{e}", flush=True)
        return False


def delete_credentials(service):
    """刪除帳密（用於測試或重新設定）。"""
    try:
        keyring.delete_password(_full_service(service), "credentials")
        return True
    except Exception:
        return False


def prompt_credentials(label, default_username="", default_password=""):
    """跳出 tkinter 對話框讓使用者輸入帳密，回傳 (username, password) 或 None（取消）。

    說明：default_username 會自動填入帳號欄；default_password 為了資安**不**預填，
    每次修改都需重新輸入密碼。
    """
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()

    dlg = tk.Toplevel(root)
    dlg.title(f"設定帳密 — {label}")
    dlg.geometry("540x260")
    dlg.update_idletasks()
    x = (dlg.winfo_screenwidth() - 540) // 2
    y = (dlg.winfo_screenheight() - 260) // 2
    dlg.geometry(f"+{x}+{y}")

    tk.Label(dlg, text=f"請輸入「{label}」的帳號與密碼",
             font=("Microsoft JhengHei", 12, "bold"),
             padx=15, pady=10).pack(anchor='w')
    tk.Label(dlg, text="會用 Windows 認證管理員加密儲存（綁此電腦的 Windows 帳號）",
             font=("Microsoft JhengHei", 9), fg='#666',
             padx=15).pack(anchor='w')

    form = tk.Frame(dlg)
    form.pack(fill=tk.X, padx=20, pady=10)

    tk.Label(form, text="帳號：", font=("Microsoft JhengHei", 11)).grid(row=0, column=0, sticky='w', pady=4)
    user_entry = tk.Entry(form, font=("Microsoft JhengHei", 11), width=35)
    user_entry.grid(row=0, column=1, padx=5)
    # 🔥 明確 insert 預設值，比 StringVar 在某些情況下更穩定
    if default_username:
        user_entry.insert(0, default_username)

    tk.Label(form, text="密碼：", font=("Microsoft JhengHei", 11)).grid(row=1, column=0, sticky='w', pady=4)
    pwd_entry = tk.Entry(form, font=("Microsoft JhengHei", 11), width=35, show='*')
    pwd_entry.grid(row=1, column=1, padx=5)
    # 🔥 密碼不預填（資安考量，每次修改都要重新輸入）

    result = [None]

    def on_ok():
        # 🔥 直接從 Entry 讀（保險），避免 StringVar 在貼上+編輯+焦點切換時偶發失同步
        u = user_entry.get().strip()
        p = pwd_entry.get()  # 密碼不 strip，使用者真的可能用空白當密碼字元
        empty = []
        if not u:
            empty.append("帳號")
        if not p:
            empty.append("密碼")
        if empty:
            messagebox.showwarning("提示", f"{'、'.join(empty)} 不能空白", parent=dlg)
            return
        result[0] = (u, p)
        dlg.destroy()

    def on_cancel():
        dlg.destroy()

    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=12)
    tk.Button(btn_frame, text="儲存", font=("Microsoft JhengHei", 11),
              bg='#4CAF50', fg='white', width=10, command=on_ok).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="取消", font=("Microsoft JhengHei", 11),
              width=10, command=on_cancel).pack(side=tk.LEFT, padx=5)

    user_entry.focus_set()
    # 🔥 Enter 在帳號欄跳到密碼欄；在密碼欄送出（標準表單行為）
    user_entry.bind('<Return>', lambda e: pwd_entry.focus_set())
    pwd_entry.bind('<Return>', lambda e: on_ok())
    dlg.bind('<Escape>', lambda e: on_cancel())
    dlg.transient()
    try:
        dlg.grab_set()
    except Exception:
        pass
    root.wait_window(dlg)
    root.destroy()

    return result[0]


def get_or_prompt_credentials(service, label):
    """主介面：取得帳密；若無則跳對話框讓使用者輸入並儲存。回傳 (username, password) 或 (None, None)。"""
    user, pwd = get_credentials(service)
    if user and pwd:
        return user, pwd

    prompted = prompt_credentials(label)
    if prompted:
        save_credentials(service, prompted[0], prompted[1])
        return prompted
    return None, None


# 已知的服務清單（service_key, 顯示名稱）
# hinet 的自然人憑證另外用 _cert 後綴，username 欄存身分證號、password 欄存憑證 PIN
# 註：qpt_hinet 程式碼雖支援第一類謄本（含憑證），但實務上只用第二類，故未列入管理介面
KNOWN_SERVICES = [
    ("V523_KH", "V523 高雄市"),
    ("V523_PT", "V523 屏東縣"),
    ("104woo", "104實價網"),
    ("hinet", "全國地政電子謄本"),
    ("hinet_cert", "全國地政電子謄本（自然人憑證）"),
    ("qpt_hinet", "全功能地籍資料查詢"),
]


def open_credential_manager(parent=None):
    """開啟帳密管理對話框：列出所有服務，可設定/修改/清除"""
    import tkinter as tk
    from tkinter import messagebox

    if parent is None:
        root = tk.Tk()
        root.withdraw()
        dlg = tk.Toplevel(root)
        owns_root = True
    else:
        root = parent
        dlg = tk.Toplevel(parent)
        owns_root = False

    dlg.title("帳密管理")
    dlg.geometry("720x500")
    dlg.update_idletasks()
    x = (dlg.winfo_screenwidth() - 720) // 2
    y = (dlg.winfo_screenheight() - 500) // 2
    dlg.geometry(f"+{x}+{y}")

    tk.Label(dlg, text="帳密管理（Windows 認證管理員）",
             font=("Microsoft JhengHei", 13, "bold"),
             pady=8).pack()
    tk.Label(dlg, text="所有帳密以加密方式儲存於 Windows 認證管理員，僅此電腦此使用者可讀取",
             font=("Microsoft JhengHei", 9), fg='#666').pack()

    list_frame = tk.Frame(dlg)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

    # 表頭
    tk.Label(list_frame, text="服務", font=("Microsoft JhengHei", 10, "bold"),
             anchor='w').grid(row=0, column=0, sticky='w', pady=(0, 6))
    tk.Label(list_frame, text="目前帳號", font=("Microsoft JhengHei", 10, "bold"),
             anchor='w').grid(row=0, column=1, sticky='w', pady=(0, 6), padx=10)

    # 每一行的 widget 用 dict 儲存，方便 refresh
    rows = {}

    def _mask_for_display(svc, user):
        """個資遮蔽：自然人憑證（身分證號）等敏感欄位顯示時遮蔽中間段
        例：A123456789 → A12*****89
        """
        if not user:
            return user
        if svc.endswith("_cert") and len(user) >= 5:
            return user[:3] + "*" * (len(user) - 5) + user[-2:]
        return user

    def refresh_row(svc, label, row_idx):
        # 移除舊 widget
        if svc in rows:
            for w in rows[svc]:
                w.destroy()
        widgets = []

        u, _ = get_credentials(svc)
        status = _mask_for_display(svc, u) if u else "（未設定）"
        status_color = '#000' if u else '#999'

        lbl1 = tk.Label(list_frame, text=label, font=("Microsoft JhengHei", 11),
                        anchor='w', width=22)
        lbl1.grid(row=row_idx, column=0, sticky='w', pady=4)
        widgets.append(lbl1)

        lbl2 = tk.Label(list_frame, text=status, font=("Microsoft JhengHei", 11),
                        anchor='w', width=18, fg=status_color)
        lbl2.grid(row=row_idx, column=1, sticky='w', pady=4, padx=10)
        widgets.append(lbl2)

        def on_edit():
            u_now, p_now = get_credentials(svc)
            new = prompt_credentials(label,
                                     default_username=u_now or "",
                                     default_password=p_now or "")
            if new:
                save_credentials(svc, new[0], new[1])
                refresh_row(svc, label, row_idx)

        def on_delete():
            if messagebox.askyesno("確認", f"確定要清除「{label}」的帳密嗎？", parent=dlg):
                delete_credentials(svc)
                refresh_row(svc, label, row_idx)

        edit_btn = tk.Button(list_frame,
                             text="設定" if not u else "修改",
                             command=on_edit,
                             font=("Microsoft JhengHei", 10),
                             bg='#2196F3', fg='white', width=8)
        edit_btn.grid(row=row_idx, column=2, padx=4, pady=4)
        widgets.append(edit_btn)

        if u:
            del_btn = tk.Button(list_frame, text="清除",
                                command=on_delete,
                                font=("Microsoft JhengHei", 10),
                                bg='#F44336', fg='white', width=8)
            del_btn.grid(row=row_idx, column=3, padx=4, pady=4)
            widgets.append(del_btn)

        rows[svc] = widgets

    # 初始化所有行（從 row=1 開始，row=0 是表頭）
    for i, (svc, label) in enumerate(KNOWN_SERVICES, start=1):
        refresh_row(svc, label, i)

    tk.Button(dlg, text="關閉", font=("Microsoft JhengHei", 11),
              width=10, command=dlg.destroy).pack(pady=10)

    if parent:
        dlg.transient(parent)
    try:
        dlg.grab_set()
    except Exception:
        pass  # 另一個視窗已 grab 時不致命，跳過
    dlg.wait_window()
    if owns_root:
        root.destroy()
