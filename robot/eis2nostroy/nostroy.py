"""Передача сведений в НОСТРОЙ.

Спецификация API НОСТРОЙ для СРО на момент написания не опубликована.
Режим "file" готовит CSV/JSON для загрузки оператором. Режим "api"
оставлен как точка расширения: заполните send_api() по спецификации,
когда её выдаст техподдержка НОСТРОЙ.
"""
from __future__ import annotations

import logging
from pathlib import Path

from .export import write_nostroy_package
from .models import ReconciledRecord

log = logging.getLogger(__name__)


class NostroyTransport:
    def __init__(self, settings: dict, out_dir: Path):
        self.mode = settings.get("mode", "file")
        self.settings = settings
        self.out_dir = out_dir

    def send(self, records: list[ReconciledRecord]) -> str:
        if self.mode == "file":
            csv_path, json_path = write_nostroy_package(records, self.out_dir)
            return f"Пакет для загрузки в НОСТРОЙ: {csv_path.name}, {json_path.name}"
        if self.mode == "api":
            return self.send_api(records)
        raise ValueError(f"Неизвестный режим передачи: {self.mode}")

    def send_api(self, records: list[ReconciledRecord]) -> str:
        raise NotImplementedError(
            "Передача через API НОСТРОЙ появится после получения спецификации. "
            "Пока используйте mode = \"file\"."
        )
