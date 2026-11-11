"""Эксплуатация: защита от двойного запуска и от спящего режима, версия робота."""
from __future__ import annotations

import json
import logging
import os
import sys
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from . import __version__

log = logging.getLogger("eis2nostroy")


class RobotBusy(RuntimeError):
    """Другая загрузка или пересборка уже идёт."""


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if sys.platform == "win32":
        import ctypes
        from ctypes import wintypes
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.OpenProcess.restype = wintypes.HANDLE
        handle = k32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
        if not handle:
            return False
        try:
            code = wintypes.DWORD()
            ok = k32.GetExitCodeProcess(handle, ctypes.byref(code))
            return bool(ok) and code.value == 259  # STILL_ACTIVE
        finally:
            k32.CloseHandle(handle)
    try:
        os.kill(pid, 0)  # на Windows так делать нельзя: там это завершает процесс
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def lock_info(path: Path) -> dict | None:
    """Кто держит замок, если он действующий; None — замка нет или он остался от упавшего запуска."""
    try:
        info = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return info if _pid_alive(int(info.get("pid", 0))) else None


@contextmanager
def run_lock(path: Path, what: str):
    """Один процесс загрузки или пересборки за раз. Замок упавшего процесса снимается сам."""
    path.parent.mkdir(parents=True, exist_ok=True)
    busy = lock_info(path)
    if busy and int(busy.get("pid", 0)) != os.getpid():
        raise RobotBusy(f"Уже идёт «{busy.get('what')}» с {busy.get('since')} (процесс {busy.get('pid')}). "
                        f"Дождитесь окончания или закройте то окно.")
    path.write_text(json.dumps({"pid": os.getpid(), "what": what,
                                "since": datetime.now().strftime("%d.%m.%Y %H:%M")}, ensure_ascii=False),
                    encoding="utf-8")
    try:
        yield
    finally:
        try:
            info = json.loads(path.read_text(encoding="utf-8"))
            if int(info.get("pid", 0)) == os.getpid():
                path.unlink()
        except (OSError, ValueError):
            pass


@contextmanager
def keep_awake():
    """Не давать Windows уснуть, пока идёт долгая операция (экран гаснуть может)."""
    flags_on = 0x80000000 | 0x00000001  # ES_CONTINUOUS | ES_SYSTEM_REQUIRED
    on = False
    if sys.platform == "win32":
        try:
            import ctypes
            on = bool(ctypes.windll.kernel32.SetThreadExecutionState(flags_on))
        except (OSError, AttributeError):
            on = False
    try:
        yield
    finally:
        if on:
            import ctypes
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)


def version_line() -> str:
    return f"Робот Flowlink eis2nostroy {__version__}"
