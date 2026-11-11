"""Договоры членов СРО, которые робот исключил, с причиной.

11.11: исключённые договоры не попадали ни в сверку, ни в файл решений, и оператор не мог
понять, почему договора нет. Теперь каждое исключение записывается в data/work/excluded.jsonl,
а сверка показывает их на листе «Исключены роботом».
Договоры, исключённые до появления журнала (разобранные раньше дни), в нём не появятся.
"""
from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

from .models import Contract


def _key(c: Contract) -> str:
    return f"{c.law}|{c.reg_number or c.raw_file}|{c.supplier_inn}"


class ExcludedLog:
    def __init__(self, path: Path):
        self.path = path
        self._known: dict[str, str] | None = None

    def _load_keys(self) -> dict[str, str]:
        if self._known is None:
            self._known = {}
            if self.path.exists():
                for line in self.path.read_text(encoding="utf-8").splitlines():
                    try:
                        d = json.loads(line)
                    except ValueError:
                        continue
                    self._known[d["key"]] = d.get("reason", "")
        return self._known

    def add(self, items: list[tuple[Contract, str]]) -> int:
        known = self._load_keys()
        new = []
        for c, reason in items:
            k = _key(c)
            if known.get(k) == reason:
                continue
            known[k] = reason
            new.append({"key": k, "reason": reason, "seen": datetime.now().isoformat(timespec="seconds"),
                        "contract": c.to_json()})
        if new:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with self.path.open("a", encoding="utf-8") as f:
                for d in new:
                    f.write(json.dumps(d, ensure_ascii=False) + "\n")
        return len(new)

    def load(self) -> list[tuple[Contract, str]]:
        """Последняя запись по каждому договору."""
        last: dict[str, dict] = {}
        if self.path.exists():
            for line in self.path.read_text(encoding="utf-8").splitlines():
                try:
                    d = json.loads(line)
                except ValueError:
                    continue
                last[d["key"]] = d
        return [(Contract.from_json(d["contract"]), d.get("reason", "")) for d in last.values()]


def for_report(log: ExcludedLog, members: dict, since: date, shown: set[str]) -> list[tuple[Contract, str]]:
    """Исключённые договоры нынешних членов СРО за период сверки, которых нет среди записей сверки."""
    out = []
    for c, reason in log.load():
        if c.supplier_inn not in members:
            continue
        if c.sign_date and c.sign_date < since:
            continue
        if c.reg_number and c.reg_number in shown:
            continue
        out.append((c, reason))
    out.sort(key=lambda x: (x[0].supplier_inn, x[0].sign_date or date.min), reverse=False)
    return out
