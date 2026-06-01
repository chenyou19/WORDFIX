from __future__ import annotations

import threading
from collections.abc import Callable

from .exceptions import ProcessStopped

class StopController:
    def __init__(self) -> None:
        self._event = threading.Event()
        self._lock = threading.Lock()
        self._stop_callbacks: list[Callable[[], None]] = []

    def stop(self) -> None:
        self._event.set()
        with self._lock:
            callbacks = list(self._stop_callbacks)
        for callback in callbacks:
            try:
                callback()
            except Exception:
                pass

    def clear(self) -> None:
        self._event.clear()

    def is_stopped(self) -> bool:
        return self._event.is_set()

    def check(self) -> None:
        if self._event.is_set():
            raise ProcessStopped("使用者已停止修改。")

    def register_stop_callback(self, callback: Callable[[], None]) -> None:
        with self._lock:
            self._stop_callbacks.append(callback)

    def unregister_stop_callback(self, callback: Callable[[], None]) -> None:
        with self._lock:
            try:
                self._stop_callbacks.remove(callback)
            except ValueError:
                pass
