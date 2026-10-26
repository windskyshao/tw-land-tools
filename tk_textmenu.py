"""輸入框/文字區的滑鼠右鍵選單(剪下、複製、貼上、刪除、全選、復原)。

Tk 的輸入框預設沒有右鍵選單，使用者只能用 Ctrl+V 貼上。install(root) 呼叫一次，
這個 Tk 程式裡所有的 Entry / ttk.Entry / Combobox / Spinbox / Text(含 ScrolledText) 都會有右鍵選單，
之後新開的視窗(Toplevel)也自動有。
"""
import tkinter as tk

_CLASSES = ("Entry", "TEntry", "TCombobox", "Spinbox", "TSpinbox", "Text")
_installed = set()


def _menu_font(widget):
    try:
        import tkinter.font as tkfont
        f = tkfont.nametofont("TkDefaultFont").actual()
        return (f.get("family", "Microsoft JhengHei"), max(11, abs(int(f.get("size", 11)))))
    except Exception:
        return ("Microsoft JhengHei", 11)


def _state(w):
    try:
        return str(w.cget("state"))
    except Exception:
        return "normal"


def _has_sel(w):
    try:
        if isinstance(w, tk.Text) or w.winfo_class() == "Text":
            return bool(w.tag_ranges("sel"))
        return bool(w.selection_present())
    except Exception:
        return False


def _popup(event):
    w = event.widget
    try:
        w.focus_set()
    except Exception:
        pass
    is_text = w.winfo_class() == "Text"
    st = _state(w)
    editable = st not in ("disabled", "readonly")
    sel = _has_sel(w)
    try:
        clip = bool(w.clipboard_get())
    except Exception:
        clip = False
    m = tk.Menu(w, tearoff=0, font=_menu_font(w))

    def ev(name):
        return lambda: w.event_generate(name)

    def select_all():
        try:
            if is_text:
                w.tag_add("sel", "1.0", "end-1c")
                w.mark_set("insert", "end-1c")
            else:
                w.selection_range(0, "end")
                w.icursor("end")
        except Exception:
            pass

    def delete_sel():
        try:
            if is_text:
                w.delete("sel.first", "sel.last")
            else:
                w.delete("sel.first", "sel.last")
        except Exception:
            pass

    m.add_command(label="剪下", accelerator="Ctrl+X", command=ev("<<Cut>>"),
                  state="normal" if (editable and sel) else "disabled")
    m.add_command(label="複製", accelerator="Ctrl+C", command=ev("<<Copy>>"),
                  state="normal" if sel else "disabled")
    m.add_command(label="貼上", accelerator="Ctrl+V", command=ev("<<Paste>>"),
                  state="normal" if (editable and clip) else "disabled")
    m.add_command(label="刪除", command=delete_sel, state="normal" if (editable and sel) else "disabled")
    m.add_separator()
    m.add_command(label="全選", accelerator="Ctrl+A", command=select_all)
    if is_text and editable:
        m.add_separator()
        m.add_command(label="復原", accelerator="Ctrl+Z", command=ev("<<Undo>>"))
        m.add_command(label="重做", accelerator="Ctrl+Y", command=ev("<<Redo>>"))
    try:
        m.tk_popup(event.x_root, event.y_root)
    finally:
        m.grab_release()
    return "break"


def install(root):
    """在這個 Tk 程式裡所有輸入框/文字區加上右鍵選單(重複呼叫沒關係)"""
    key = str(root)
    if key in _installed:
        return
    _installed.add(key)
    for cls in _CLASSES:
        try:
            # add="+"：不蓋掉程式原本綁在 class 上的右鍵功能(個別元件自己綁的優先)
            root.bind_class(cls, "<Button-3>", _popup, add="+")
        except Exception:
            pass
    # Entry 預設沒有 Ctrl+A 全選
    try:
        root.bind_class("Entry", "<Control-a>", lambda e: (e.widget.selection_range(0, "end"), "break")[1], add="+")
        root.bind_class("TEntry", "<Control-a>", lambda e: (e.widget.selection_range(0, "end"), "break")[1], add="+")
    except Exception:
        pass


def auto():
    """之後這個程式裡建立的每個 tk.Tk() 都自動 install()(子程式開頭呼叫一次就好)"""
    if getattr(tk.Tk, "_textmenu_auto", False):
        return
    _orig = tk.Tk.__init__

    def _init(self, *a, **k):
        _orig(self, *a, **k)
        try:
            install(self)
        except Exception:
            pass
    tk.Tk.__init__ = _init
    tk.Tk._textmenu_auto = True
