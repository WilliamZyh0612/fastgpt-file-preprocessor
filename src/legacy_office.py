"""Read-only legacy Office conversion in an isolated process; never execute document code."""
from __future__ import annotations

from multiprocessing import Process, Queue
from pathlib import Path
import time


def extract_legacy_readonly(path: Path, timeout_seconds: int) -> str:
    queue: Queue = Queue(); worker = Process(target=_worker, args=(str(path.resolve()), queue), daemon=True); worker.start()
    office_pid: int | None = None; result: tuple[str, str] | None = None; deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline and result is None:
        try:
            state, payload = queue.get(timeout=min(.2, max(.01, deadline - time.monotonic())))
            if state == "office_pid": office_pid = int(payload)
            else: result = (state, payload)
        except Exception:
            if not worker.is_alive(): break
    worker.join(max(0, deadline - time.monotonic()))
    if worker.is_alive():
        # finally in the worker cannot run after terminate().  Only kill the PID
        # reported by this worker's DispatchEx instance; never enumerate or touch
        # Office processes that were already running for the interactive user.
        worker.terminate(); worker.join(5)
        if office_pid is not None: _terminate_owned_office_process(office_pid)
        raise TimeoutError(f"旧版 Office 转换超时：{timeout_seconds} 秒；已清理本任务 Office PID={office_pid or '未知'}")
    while result is None:
        try:
            state, payload = queue.get_nowait()
            if state == "office_pid": office_pid = int(payload)
            else: result = (state, payload)
        except Exception: break
    if result is None: raise RuntimeError("旧版 Office 转换未返回结果")
    state, payload = result
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
            app.AutomationSecurity = 3; app.Options.UpdateLinksAtOpen = False; _report_pid(app, queue)
            document = app.Documents.Open(filename, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False, NoEncodingDialog=True, UpdateLinks=0)
            queue.put(("ok", document.Content.Text))
        elif suffix == ".xls":
            app = win32com.client.DispatchEx("Excel.Application"); app.Visible = False; app.DisplayAlerts = False; app.AutomationSecurity = 3
            app.AskToUpdateLinks = False; app.EnableEvents = False; _report_pid(app, queue)
            document = app.Workbooks.Open(filename, UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
            queue.put(("ok", "\n".join(" | ".join(str(item) for item in row if item is not None) for sheet in document.Worksheets for row in _rows(sheet.UsedRange.Value))))
        elif suffix == ".ppt":
            app = win32com.client.DispatchEx("PowerPoint.Application"); app.DisplayAlerts = 0
            app.AutomationSecurity = 3; _report_pid(app, queue)
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


def _report_pid(app, queue: Queue) -> None:
    """Report only the Office process created by DispatchEx in this worker."""
    try:
        import ctypes
        hwnd = int(getattr(app, "Hwnd", 0) or getattr(app, "HWND", 0))
        pid = ctypes.c_ulong()
        if hwnd and ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid)):
            queue.put(("office_pid", str(pid.value)))
    except Exception:
        # PID collection is best effort; conversion remains read-only when Office
        # does not expose a window handle (for example, a headless PowerPoint).
        pass


def _terminate_owned_office_process(pid: int) -> bool:
    """Terminate a known task-owned Office PID, never an enumerated process."""
    try:
        import ctypes
        process = ctypes.windll.kernel32.OpenProcess(0x0001, False, pid)  # PROCESS_TERMINATE
        if not process: return False
        try: return bool(ctypes.windll.kernel32.TerminateProcess(process, 1))
        finally: ctypes.windll.kernel32.CloseHandle(process)
    except Exception:
        return False
