# 智慧匯出封存功能（新版本）
# 這是 data_editor.py 中 export_archive() 的改良版本

def export_archive(force_reset=False):
    """匯出封存資料夾（智慧模式）

    Args:
        force_reset: 是否強制重新設定（Shift+點擊時為 True）
    """
    import os
    import json
    import shutil
    import zipfile
    import stat
    import time
    from datetime import datetime
    import tkinter as tk
    from tkinter import filedialog, messagebox

    try:
        update_message("=" * 50)
        update_message("開始智慧匯出封存...")

        # ========================================
        # 1. 檢查並讀取工作區資料
        # ========================================
        if not os.path.exists('data.json'):
            messagebox.showerror("錯誤", "找不到 data.json，無法確定資料夾名稱")
            return

        with open('data.json', 'r', encoding='utf-8') as f:
            data_list = json.load(f)

        if not data_list:
            messagebox.showerror("錯誤", "data.json 中沒有資料")
            return

        first_data = data_list[0]
        area = first_data.get('area', '')
        section = first_data.get('section', '')
        lot_number = first_data.get('lot_number', '')

        source_folder = f"{area}{section}-{lot_number}"

        if not os.path.exists(source_folder):
            messagebox.showerror("錯誤", f"找不到工作區資料夾：{source_folder}")
            return

        update_message(f"工作區資料夾：{source_folder}")

        # ========================================
        # 2. 檢查並建立必要的子資料夾
        # ========================================
        required_subfolders = [
            "0.謄本", "1.基本資料", "2.照片", "3.行情",
            "4.其他相關", "5.委託", "6.成交"
        ]

        for subfolder in required_subfolders:
            subfolder_path = os.path.join(source_folder, subfolder)
            if not os.path.exists(subfolder_path):
                os.makedirs(subfolder_path)
                update_message(f"[建立] {subfolder}")

        # ========================================
        # 3. 讀取建物門牌
        # ========================================
        data_final_path = get_data_final_path()
        門牌 = ""
        if os.path.exists(data_final_path):
            with open(data_final_path, 'r', encoding='utf-8') as f:
                data_final = json.load(f)
                門牌_原始 = data_final.get('建物門牌', '')
                if 門牌_原始:
                    門牌 = format_address(門牌_原始)
                    update_message(f"建物門牌：{門牌}")

        # ========================================
        # 4. 讀取上次匯出設定（智慧模式核心）
        # ========================================
        config = load_config()
        export_settings = config.get('export_settings', {})

        # 檢查是否為同一資料夾
        last_folder = export_settings.get('current_folder', '')
        is_same_folder = (last_folder == source_folder)
        has_previous_export = bool(export_settings.get('last_export_path'))

        # 決定使用模式
        use_smart_mode = (is_same_folder and has_previous_export and not force_reset)

        if use_smart_mode:
            update_message("[模式] 智慧更新模式")
        else:
            update_message("[模式] 首次匯出模式")
            export_settings = {}  # 清空設定

        # ========================================
        # 5. 智慧更新模式 - 檢查現有封存
        # ========================================
        if use_smart_mode:
            last_export_path = export_settings.get('last_export_path', '')
            last_export_format = export_settings.get('last_export_format', 'zip')
            last_archive_name = export_settings.get('last_archive_name', '')
            auto_export = export_settings.get('auto_export', False)  # 記住我的選擇

            # 構建完整路徑
            if last_export_format in ['folder', 'both']:
                last_full_path = os.path.join(last_export_path, last_archive_name)
            else:
                last_full_path = os.path.join(last_export_path, last_archive_name + '.zip')

            # 檢查檔案是否存在
            archive_exists = os.path.exists(last_full_path)

            if archive_exists:
                # 取得檔案大小
                if os.path.isdir(last_full_path):
                    total_size = sum(
                        os.path.getsize(os.path.join(dirpath, filename))
                        for dirpath, dirnames, filenames in os.walk(last_full_path)
                        for filename in filenames
                    )
                else:
                    total_size = os.path.getsize(last_full_path)

                size_mb = total_size / (1024 * 1024)

                # 取得檔案修改時間
                if os.path.isdir(last_full_path):
                    mod_time = os.path.getmtime(last_full_path)
                else:
                    mod_time = os.path.getmtime(last_full_path)

                mod_date = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M')

                update_message(f"[發現] 現有封存：{last_archive_name}")
                update_message(f"  位置：{last_export_path}")
                update_message(f"  大小：{size_mb:.2f} MB")
                update_message(f"  修改：{mod_date}")

                # 如果「記住我的選擇」，直接執行
                if auto_export:
                    update_message("[自動] 更新內容，保持日期")
                    update_mode = "keep_date"
                    selected_format = last_export_format
                    dest_parent = last_export_path
                    archive_name = last_archive_name
                else:
                    # 顯示智慧更新對話框
                    update_result = show_smart_update_dialog(
                        last_archive_name, last_export_path,
                        size_mb, mod_date, area, section, lot_number, 門牌
                    )

                    if update_result is None:
                        update_message("[取消] 使用者取消匯出")
                        return

                    update_mode, auto_export_choice = update_result

                    # 儲存「記住我的選擇」設定
                    if auto_export_choice:
                        export_settings['auto_export'] = True
                        config['export_settings'] = export_settings
                        save_config(config)
                        update_message("[設定] 已記住選擇，下次將自動執行")

                    selected_format = last_export_format
                    dest_parent = last_export_path

                    if update_mode == "keep_date":
                        # 保持原日期
                        archive_name = last_archive_name
                        update_message("[選擇] 更新內容，保持日期")
                    elif update_mode == "update_date":
                        # 更新為今天日期
                        today = datetime.now()
                        民國年 = today.year - 1911
                        日期 = f"{民國年}.{today.month:02d}.{today.day:02d}"
                        if 門牌:
                            archive_name = f"{日期}{area}{section}-{lot_number},{門牌}"
                        else:
                            archive_name = f"{日期}{area}{section}-{lot_number}"
                        update_message(f"[選擇] 更新日期為：{日期}")
                    elif update_mode == "create_new":
                        # 建立新封存（不覆蓋舊的）
                        today = datetime.now()
                        民國年 = today.year - 1911
                        日期 = f"{民國年}.{today.month:02d}.{today.day:02d}"
                        時間戳記 = today.strftime("%H%M%S")
                        if 門牌:
                            archive_name = f"{日期}{area}{section}-{lot_number},{門牌}_{時間戳記}"
                        else:
                            archive_name = f"{日期}{area}{section}-{lot_number}_{時間戳記}"
                        update_message(f"[選擇] 建立新封存（保留舊的）")
                    else:  # reset
                        # 重新選擇位置和格式
                        use_smart_mode = False
                        export_settings = {}
                        update_message("[選擇] 重新設定匯出選項")
            else:
                # 檔案不存在，使用上次設定但提示
                update_message("[警告] 上次封存檔案不存在，將重新建立")
                selected_format = last_export_format
                dest_parent = last_export_path

                # 使用今天日期
                today = datetime.now()
                民國年 = today.year - 1911
                日期 = f"{民國年}.{today.month:02d}.{today.day:02d}"
                if 門牌:
                    archive_name = f"{日期}{area}{section}-{lot_number},{門牌}"
                else:
                    archive_name = f"{日期}{area}{section}-{lot_number}"

        # ========================================
        # 6. 首次匯出模式 - 選擇格式和位置
        # ========================================
        if not use_smart_mode or update_mode == "reset":
            # 生成新的封存名稱
            today = datetime.now()
            民國年 = today.year - 1911
            日期 = f"{民國年}.{today.month:02d}.{today.day:02d}"
            if 門牌:
                archive_name = f"{日期}{area}{section}-{lot_number},{門牌}"
            else:
                archive_name = f"{日期}{area}{section}-{lot_number}"

            # 選擇匯出格式
            format_result = show_format_dialog()
            if format_result is None:
                update_message("[取消] 使用者取消選擇格式")
                return

            selected_format = format_result
            update_message(f"[選擇] 匯出格式：{selected_format}")

            # 選擇匯出位置
            dest_parent = filedialog.askdirectory(title="選擇封存位置")
            if not dest_parent:
                update_message("[取消] 使用者取消選擇位置")
                return

            update_message(f"[選擇] 匯出位置：{dest_parent}")

        # ========================================
        # 7. 執行匯出
        # ========================================
        update_message(f"封存名稱：{archive_name}")

        dest_path = os.path.join(dest_parent, archive_name)
        zip_path = dest_path + ".zip"

        # 移除舊檔案/資料夾
        def remove_readonly(func, path, excinfo):
            """移除唯讀屬性後重試刪除"""
            os.chmod(path, stat.S_IWRITE)
            func(path)

        if selected_format in ["folder", "both"] and os.path.exists(dest_path):
            update_message("[刪除] 正在刪除舊資料夾...")
            try:
                shutil.rmtree(dest_path, onerror=remove_readonly)
                time.sleep(0.3)
                update_message("[刪除] 已刪除舊資料夾")
            except Exception as e:
                update_message(f"[警告] 刪除舊資料夾失敗：{e}")

        if selected_format in ["zip", "both"] and os.path.exists(zip_path):
            update_message("[刪除] 正在刪除舊 ZIP...")
            try:
                os.remove(zip_path)
                update_message("[刪除] 已刪除舊 ZIP 檔案")
            except Exception as e:
                update_message(f"[警告] 刪除舊 ZIP 失敗：{e}")

        files_created = []

        # 複製資料夾
        if selected_format in ["folder", "both"]:
            update_message("開始複製資料夾...")
            shutil.copytree(source_folder, dest_path)
            update_message(f"[完成] 資料夾已匯出")
            files_created.append(("資料夾", dest_path))

        # 建立 ZIP
        if selected_format in ["zip", "both"]:
            update_message("開始建立 ZIP 壓縮檔...")
            try:
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root_dir, dirs, files in os.walk(source_folder):
                        for file in files:
                            file_path = os.path.join(root_dir, file)
                            arcname = os.path.join(
                                archive_name,
                                os.path.relpath(file_path, source_folder)
                            )
                            zipf.write(file_path, arcname)

                update_message(f"[完成] ZIP 壓縮檔已建立")
                files_created.append(("ZIP檔", zip_path))
            except Exception as e:
                update_message(f"[錯誤] 建立 ZIP 失敗：{e}")
                messagebox.showerror("錯誤", f"建立 ZIP 壓縮檔失敗：\n\n{e}")

        # ========================================
        # 8. 儲存匯出設定
        # ========================================
        export_settings = {
            'last_export_path': dest_parent,
            'last_export_format': selected_format,
            'last_archive_name': archive_name,
            'last_export_date': datetime.now().strftime('%Y-%m-%d'),
            'current_folder': source_folder,
            'auto_export': export_settings.get('auto_export', False)  # 保留原設定
        }

        config['export_settings'] = export_settings
        save_config(config)
        update_message("[設定] 已儲存匯出設定")

        # ========================================
        # 9. 顯示結果
        # ========================================
        update_message("=" * 50)
        update_message("匯出結果：")
        for file_type, file_path in files_created:
            update_message(f"  ✓ {file_type}：{os.path.basename(file_path)}")

        # 詢問是否開啟資料夾
        if files_created:
            result = show_large_yesno(
                "匯出完成",
                f"已成功匯出封存：\n\n{archive_name}\n\n位置：{dest_parent}\n\n是否要開啟資料夾？"
            )

            if result:
                try:
                    os.startfile(dest_parent)
                except Exception as e:
                    update_message(f"[錯誤] 開啟資料夾失敗：{e}")

    except Exception as e:
        update_message(f"[錯誤] 匯出封存失敗：{e}")
        import traceback
        update_message(traceback.format_exc())
        messagebox.showerror("錯誤", f"匯出封存失敗：\n\n{e}")


def show_format_dialog():
    """顯示匯出格式選擇對話框"""
    format_dialog = tk.Toplevel()
    format_dialog.title("選擇匯出格式")
    format_dialog.geometry("500x300")
    format_dialog.transient(root if 'root' in globals() else None)
    format_dialog.grab_set()

    # 置中顯示
    screen_width = format_dialog.winfo_screenwidth()
    screen_height = format_dialog.winfo_screenheight()
    x = (screen_width - 500) // 2
    y = (screen_height - 300) // 2
    format_dialog.geometry(f"500x300+{x}+{y}")

    tk.Label(format_dialog, text="請選擇匯出格式：",
            font=("Microsoft JhengHei", 14, "bold")).pack(pady=20)

    export_format = tk.StringVar(value="zip")

    tk.Radiobutton(format_dialog, text="資料夾（方便查閱）",
                  variable=export_format, value="folder",
                  font=("Microsoft JhengHei", 12)).pack(pady=10, anchor="w", padx=50)

    tk.Radiobutton(format_dialog, text="資料夾 + ZIP 壓縮檔（兼顧兩者）",
                  variable=export_format, value="both",
                  font=("Microsoft JhengHei", 12)).pack(pady=10, anchor="w", padx=50)

    tk.Radiobutton(format_dialog, text="只建立 ZIP 壓縮檔（節省空間）",
                  variable=export_format, value="zip",
                  font=("Microsoft JhengHei", 12)).pack(pady=10, anchor="w", padx=50)

    btn_frame = tk.Frame(format_dialog)
    btn_frame.pack(pady=20)

    result = [None]

    def on_confirm():
        result[0] = export_format.get()
        format_dialog.destroy()

    def on_cancel():
        format_dialog.destroy()

    tk.Button(btn_frame, text="確定", command=on_confirm,
             font=("Microsoft JhengHei", 12, "bold"),
             bg='#4CAF50', fg='white', width=10, height=2).pack(side=tk.LEFT, padx=10)

    tk.Button(btn_frame, text="取消", command=on_cancel,
             font=("Microsoft JhengHei", 12),
             bg='#F44336', fg='white', width=10, height=2).pack(side=tk.LEFT, padx=10)

    format_dialog.wait_window()
    return result[0]


def show_smart_update_dialog(archive_name, export_path, size_mb, mod_date,
                             area, section, lot_number, 門牌):
    """顯示智慧更新對話框

    Returns:
        (update_mode, auto_export): 更新模式和是否記住選擇
        update_mode: "keep_date", "update_date", "create_new", "reset", None
    """
    from datetime import datetime

    dialog = tk.Toplevel()
    dialog.title("智慧匯出 - 發現現有封存")
    dialog.geometry("600x500")
    dialog.transient(root if 'root' in globals() else None)
    dialog.grab_set()

    # 置中顯示
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width - 600) // 2
    y = (screen_height - 500) // 2
    dialog.geometry(f"600x500+{x}+{y}")

    # 標題
    tk.Label(dialog, text="📦 發現現有封存",
            font=("Microsoft JhengHei", 16, "bold"),
            fg='#FF9800').pack(pady=15)

    # 現有封存資訊框
    info_frame = tk.Frame(dialog, relief=tk.RIDGE, borderwidth=2, bg='#F5F5F5')
    info_frame.pack(padx=20, pady=10, fill=tk.BOTH)

    tk.Label(info_frame, text=f"位置：{export_path}",
            font=("Microsoft JhengHei", 10), bg='#F5F5F5',
            anchor='w').pack(padx=15, pady=3, fill=tk.X)

    tk.Label(info_frame, text=f"檔案：{archive_name}",
            font=("Microsoft JhengHei", 10, "bold"), bg='#F5F5F5',
            anchor='w').pack(padx=15, pady=3, fill=tk.X)

    tk.Label(info_frame, text=f"大小：{size_mb:.2f} MB",
            font=("Microsoft JhengHei", 10), bg='#F5F5F5',
            anchor='w').pack(padx=15, pady=3, fill=tk.X)

    tk.Label(info_frame, text=f"修改：{mod_date}",
            font=("Microsoft JhengHei", 10), bg='#F5F5F5',
            anchor='w').pack(padx=15, pady=3, fill=tk.X)

    # 分隔線
    tk.Frame(dialog, height=2, bg='#E0E0E0').pack(fill=tk.X, padx=20, pady=10)

    # 選擇更新方式
    tk.Label(dialog, text="請選擇更新方式：",
            font=("Microsoft JhengHei", 12, "bold")).pack(pady=10)

    update_mode = tk.StringVar(value="keep_date")

    # 生成今天日期
    today = datetime.now()
    民國年 = today.year - 1911
    今天日期 = f"{民國年}.{today.month:02d}.{today.day:02d}"

    tk.Radiobutton(dialog,
                  text=f"● 更新檔案內容，保持原日期（{archive_name.split(area)[0]}）",
                  variable=update_mode, value="keep_date",
                  font=("Microsoft JhengHei", 11),
                  fg='#2196F3', selectcolor='white').pack(pady=8, anchor="w", padx=50)

    tk.Radiobutton(dialog,
                  text=f"○ 更新檔案內容，並更新日期為今天（{今天日期}）",
                  variable=update_mode, value="update_date",
                  font=("Microsoft JhengHei", 11),
                  selectcolor='white').pack(pady=8, anchor="w", padx=50)

    tk.Radiobutton(dialog,
                  text="○ 建立新的封存（保留舊的，加上時間戳記）",
                  variable=update_mode, value="create_new",
                  font=("Microsoft JhengHei", 11),
                  selectcolor='white').pack(pady=8, anchor="w", padx=50)

    tk.Radiobutton(dialog,
                  text="○ 重新選擇位置和格式",
                  variable=update_mode, value="reset",
                  font=("Microsoft JhengHei", 11),
                  selectcolor='white').pack(pady=8, anchor="w", padx=50)

    # 分隔線
    tk.Frame(dialog, height=1, bg='#E0E0E0').pack(fill=tk.X, padx=50, pady=15)

    # 記住我的選擇
    auto_export_var = tk.BooleanVar(value=False)
    tk.Checkbutton(dialog,
                  text="☑ 記住我的選擇（下次直接更新，不再詢問）",
                  variable=auto_export_var,
                  font=("Microsoft JhengHei", 10),
                  fg='#666666').pack(pady=5)

    # 按鈕
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=20)

    result = [None]

    def on_confirm():
        result[0] = (update_mode.get(), auto_export_var.get())
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    tk.Button(btn_frame, text="開始匯出", command=on_confirm,
             font=("Microsoft JhengHei", 12, "bold"),
             bg='#4CAF50', fg='white', width=12, height=2).pack(side=tk.LEFT, padx=10)

    tk.Button(btn_frame, text="取消", command=on_cancel,
             font=("Microsoft JhengHei", 12),
             bg='#F44336', fg='white', width=12, height=2).pack(side=tk.LEFT, padx=10)

    dialog.wait_window()
    return result[0]
