"""Список членов СРО."""
from __future__ import annotations

from pathlib import Path

from .models import Member
from .tabular import pick, read_rows, to_date, to_inn


def load_members(path: Path) -> dict[str, Member]:
    """Читает членов СРО из XLSX/CSV.

    Нужен столбец с «ИНН». Остальные необязательны: «Наименование»,
    «Дата присвоения уровня по ОДО», «Дата вступления».
    Подходит лист «Члены» из шаблона сверки.
    """
    members: dict[str, Member] = {}
    for row in read_rows(path):
        inn = to_inn(pick(row, "инн"))
        if len(inn) not in (10, 12):
            continue
        members[inn] = Member(
            inn=inn,
            name=str(pick(row, "наименование", "член", "организация") or ""),
            odo_date=to_date(pick(row, "одо", "уровн")),
            joined=to_date(pick(row, "вступлен", "допуск")),
        )
    return members
