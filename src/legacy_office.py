"""Read-only legacy Office conversion in an isolated process; never execute document code."""
from __future__ import annotations

from multiprocessing import Process, Queue
from pathlib import Path


def extract_legacy_readonly(path: Path, timeout_seconds: int) -> str:
    queue: Queue = Queue(); worker = Process(target=_worker, args=(str(path.resolve()), queue), daemon=True); worker.start(); worker.join(timeout_seconds)
    if worker.is_alive():
        worker.terminate(); worker.join(5); raise TimeoutError(f"旧版 Office 转换超时：{timeout_seconds} 秒")
    if queue.empty(): raise RuntimeError("旧版 Office 转换未返回结果")
    state, payload = queue.get()
    if state == "error": raise RuntimeError(payload)
    return payload


def _worker(filename: str, queue: Queue) -> None:
    app = document = None
    try:
        import win32com.client
        suffix = Path(filename).suffix.lower()
        if suffix == ".doc":
            app = win32com.client.DispatchEx("Word.Application"); app.Visible = False; app.DisplayAlerts = 0
            # msoAutomationSecurityForceDisable: do not run macros from the document.
            app.AutomationSecurity = 3; document = app.Documents.Open(filename, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False, NoEncodingDialog=True)
            queue.put(("ok", document.Content.Text))
        elif suffix == ".xls":
            app = win32com.client.DispatchEx("Excel.Application"); app.Visible = False; app.DisplayAlerts = False; app.AutomationSecurity = 3
            document = app.Workbooks.Open(filename, UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
            queue.put(("ok", "\n".join(" | ".join(str(item) for item in row if item is not None) for sheet in document.Worksheets for row in _rows(sheet.UsedRange.Value))))
        elif suffix == ".ppt":
            app = win32com.client.DispatchEx("PowerPoint.Application"); app.DisplayAlerts = 0
            document = app.Presentations.Open(filename, ReadOnly=True, Untitled=False, WithWindow=False)
            queue.put(("ok", "\n".join(shape.TextFrame.TextRange.Text for slide in document.Slides for shape in slide.Shapes if shape.HasTextFrame and shape.TextFrame.HasText)))
        else: raise ValueError("不支持的旧版 Office 类型")
    except Exception as exc: queue.put(("error", f"安全转换失败：{type(exc).__name__}: {str(exc)[:200]}"))
    finally:
        try:
            if document is not None: document.Close(False)
        finally:
            if app is not None: app.Quit()


def _rows(value):
    if not isinstance(value, tuple): return ((value,),)
    return value if value and isinstance(value[0], tuple) else (value,)
