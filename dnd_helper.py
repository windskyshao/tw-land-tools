"""檔案拖曳(跟 Windows 檔案總管互通)：tkinterdnd2(tkdnd)的薄包裝，載入失敗就靜靜關掉這個功能。

    import dnd_helper
    if dnd_helper.enable(root):                       # 在既有的 tk.Tk() 上啟用
        dnd_helper.drop_files(widget, on_drop)        # 檔案總管 → 程式：on_drop(paths, event)
        dnd_helper.drag_files(widget, get_paths)      # 程式 → 檔案總管/桌面：get_paths() 回要拖出去的檔案路徑
"""
import os

_ok = {}


def enable(root):
    """在既有的 Tk 視窗載入 tkdnd；成功回 True(打包版要把 tkinterdnd2 的 tkdnd 資料夾一起打包)"""
    key = str(root)
    if key in _ok:
        return _ok[key]
    try:
        from tkinterdnd2 import TkinterDnD
        TkinterDnD._require(root)
        _ok[key] = True
    except Exception as e:
        print(f"[拖曳] 無法啟用檔案拖曳：{e}")
        _ok[key] = False
    return _ok[key]


def split_paths(widget, data):
    """tkdnd 給的 '{C:/有 空白/a.pdf} C:/b.pdf' → ['C:\\有 空白\\a.pdf', 'C:\\b.pdf']"""
    try:
        items = widget.tk.splitlist(data)
    except Exception:
        items = str(data).split()
    return [os.path.normpath(p) for p in items if p]


def drop_files(widget, on_drop, on_enter=None, on_leave=None):
    """讓 widget 接受從檔案總管拖進來的檔案/資料夾。on_drop(paths, event) 回傳 'copy'/'move'/None"""
    try:
        from tkinterdnd2 import DND_FILES, COPY
        widget.drop_target_register(DND_FILES)

        def _drop(event):
            paths = split_paths(widget, event.data)
            try:
                r = on_drop(paths, event)
            except Exception as e:
                print(f"[拖曳] 處理拖進來的檔案失敗：{e}")
                r = None
            if on_leave:
                on_leave(event)
            return r or COPY
        widget.dnd_bind('<<Drop>>', _drop)
        if on_enter:
            widget.dnd_bind('<<DropEnter>>', lambda e: (on_enter(e), e.action)[1])
            widget.dnd_bind('<<DropPosition>>', lambda e: (on_enter(e), e.action)[1])
        if on_leave:
            widget.dnd_bind('<<DropLeave>>', lambda e: (on_leave(e), e.action)[1])
        return True
    except Exception as e:
        print(f"[拖曳] 這個元件無法接受拖放：{e}")
        return False


def drag_files(widget, get_paths):
    """讓 widget 可以把檔案拖出去(到檔案總管、桌面、LINE、Email…)。get_paths() 回路徑清單(沒有就回空)"""
    try:
        from tkinterdnd2 import DND_FILES, COPY, MOVE, LINK
        widget.drag_source_register(1, DND_FILES)

        def _init(event):
            paths = [p for p in (get_paths() or []) if p and os.path.exists(p)]
            if not paths:
                return 'refuse_drop'
            return ((COPY, MOVE, LINK), DND_FILES, tuple(paths))
        widget.dnd_bind('<<DragInitCmd>>', _init)
        return True
    except Exception as e:
        print(f"[拖曳] 這個元件無法拖出檔案：{e}")
        return False
