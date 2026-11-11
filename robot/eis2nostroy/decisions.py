"""Память робота: решения оператора и журнал переданного в НОСТРОЙ.

data/decisions.xlsx — очередь «Требуют решения». Робот сам дописывает в неё новые спорные
договоры; оператор заполняет колонку «Решение». При следующей сверке решение применяется,
и договор больше не попадает в очередь. Решение можно принять и по договору «К передаче»
(например, исключить) — добавьте строку с его реестровым номером.

data/peredano.csv — что уже загружено в НОСТРОЙ (заполняет команда mark-sent после загрузки
пакета). Такие договоры получают статус «Уже в НОСТРОЙ» и не попадают в следующий пакет,
пока у них не изменится цена.
"""
from __future__ import annotations

import csv
import logging
from datetime import date, datetime
from pathlib import Path

from .suggest import suggest
from .models import (CONFIRMED, DECLARED, EXCLUDED, IN_NOSTROY, MISMATCH, NEEDS_CHECK, PROBABLE,
                     Contract, ReconciledRecord)

log = logging.getLogger("eis2nostroy")

QUEUE = (NEEDS_CHECK, PROBABLE, MISMATCH, DECLARED)
CHOICES = ["Передать", "Исключить: не по ГрК", "Исключить: текущий ремонт", "Исключить: дубль или ошибка",
           "Ждать ответа члена"]
HEAD = ["Ключ (реестровый номер)", "ИНН члена", "Член СРО", "Закон", "Дата договора", "Цена, ₽", "Заказчик",
        "Предмет", "Статус робота при добавлении", "Замечание робота", "Решение", "Комментарий",
        "Кто решил", "Дата решения", "Статус робота сейчас", "Предложение робота", "Основание предложения"]
JOURNAL_HEAD = ["Ключ", "ИНН члена", "Реестровый номер", "Дата договора", "Цена, ₽", "Пакет", "Передан"]


def record_key(c: Contract) -> str:
    """Реестровый номер ЕИС, а если его нет — ИНН|извещение|дата."""
    if c.reg_number:
        return c.reg_number.strip()
    d = c.sign_date.isoformat() if c.sign_date else ""
    return f"{c.supplier_inn}|{c.notice_number}|{d}"


def kind_of(decision: str) -> str:
    s = (decision or "").strip().lower()
    if s.startswith("передат"):
        return "send"
    if s.startswith("исключ") or s.startswith("не передав"):
        return "exclude"
    if s.startswith("жд"):
        return "wait"
    return ""


def _s(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, date):
        return v.strftime("%d.%m.%Y")
    return str(v).strip()


# ---------------------------------------------------------------- решения
def load_decisions(path: Path) -> dict[str, dict]:
    if not path or not path.exists():
        return {}
    from openpyxl import load_workbook
    ws = load_workbook(path, data_only=True)["Решения"]
    rows = list(ws.iter_rows(values_only=True))
    hi = next((i for i, r in enumerate(rows) if r and _s(r[0]).startswith("Ключ")), None)
    if hi is None:
        return {}
    out = {}
    for r in rows[hi + 1:]:
        if not r or not _s(r[0]):
            continue
        d = dict(zip(HEAD, list(r) + [None] * (len(HEAD) - len(r))))
        out[_s(r[0])] = d
    return out


def apply_decisions(records: list[ReconciledRecord], decisions: dict[str, dict]) -> dict:
    stats = {"send": 0, "exclude": 0, "wait": 0}
    for r in records:
        r.robot_status = r.status
        d = decisions.get(record_key(r.primary))
        if not d:
            continue
        k = kind_of(_s(d.get("Решение")))
        if not k:
            continue
        who = ", ".join(x for x in (_s(d.get("Кто решил")), _s(d.get("Дата решения"))) if x)
        note = f"решение оператора: {_s(d.get('Решение'))}" + (f" ({who})" if who else "") + \
               (f" — {_s(d.get('Комментарий'))}" if _s(d.get("Комментарий")) else "")
        if k == "send":
            r.status = CONFIRMED
        elif k == "exclude":
            r.status = EXCLUDED
        r.issues.append(note)
        stats[k] += 1
    return stats


def sync_queue(path: Path, records: list[ReconciledRecord], decisions: dict[str, dict]) -> int:
    """Дописать в файл решений новые спорные договоры и обновить «Статус робота сейчас»."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    rows = {k: dict(v) for k, v in decisions.items()}
    now_status = {record_key(r.primary): r for r in records}
    added = 0
    for r in records:
        k = record_key(r.primary)
        if getattr(r, "robot_status", r.status) in QUEUE and k not in rows:
            c = r.primary
            rows[k] = {"Ключ (реестровый номер)": k, "ИНН члена": r.member.inn, "Член СРО": r.member.name,
                       "Закон": {"44": "44-ФЗ", "223": "223-ФЗ", "615": "ПП-615"}.get(c.law, c.law),
                       "Дата договора": c.sign_date, "Цена, ₽": c.price, "Заказчик": c.customer_name,
                       "Предмет": (c.subject or "")[:500], "Статус робота при добавлении": r.robot_status,
                       "Замечание робота": "; ".join(i for i in r.issues if not i.startswith("решение оператора"))[:500]}
            added += 1
    for k, row in rows.items():
        rec = now_status.get(k)
        row["Статус робота сейчас"] = getattr(rec, "robot_status", "") if rec else "нет в сверке"
        if rec:
            row["Заказчик"] = rec.primary.customer_name or row.get("Заказчик")
            if rec.primary.price is not None:
                row["Цена, ₽"] = rec.primary.price   # цена последней версии договора
            row["Предмет"] = (rec.primary.subject or "")[:500] or row.get("Предмет")
            row["Предложение робота"], row["Основание предложения"] = suggest(
                row.get("Предмет") or "", "; ".join(rec.issues), row["Статус робота сейчас"])
        else:
            row["Предложение робота"], row["Основание предложения"] = "", ""

    wb = Workbook()
    ws = wb.active
    ws.title = "Решения"
    ws.append(["Решения оператора по спорным договорам. Заполните «Решение» (список), «Комментарий», «Кто решил», "
               "«Дата решения». Новые спорные договоры робот дописывает сам; решённые больше не попадают в очередь."])
    ws.append([f"Обновлено {datetime.now():%d.%m.%Y %H:%M}. Варианты решения: " + "; ".join(CHOICES)])
    ws.append(HEAD)
    for j in range(1, len(HEAD) + 1):
        cell = ws.cell(3, j)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3A5F")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    order = sorted(rows.values(), key=lambda d: (bool(kind_of(_s(d.get("Решение")))), -(d.get("Цена, ₽") or 0)
                                                if isinstance(d.get("Цена, ₽"), (int, float)) else 0))
    for d in order:
        ws.append([d.get(h) for h in HEAD])
        rr = ws.max_row
        ws.cell(rr, 5).number_format = "DD.MM.YYYY"
        ws.cell(rr, 6).number_format = "#,##0.00"
        ws.cell(rr, 14).number_format = "DD.MM.YYYY"
        if not kind_of(_s(d.get("Решение"))):
            ws.cell(rr, 11).fill = PatternFill("solid", fgColor="FCEBC8")
    dv = DataValidation(type="list", formula1='"' + ",".join(CHOICES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"K4:K{max(ws.max_row, 4) + 500}")
    for col, w in zip("ABCDEFGHIJKLMNOPQ", (24, 13, 34, 9, 12, 16, 34, 50, 16, 40, 26, 30, 16, 13, 16, 26, 44)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "D4"
    ws.auto_filter.ref = f"A3:Q{max(ws.max_row, 3)}"
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(path)
    except PermissionError:
        log.warning("Файл решений %s открыт в Excel — новые спорные договоры не дописаны. Закройте его "
                    "и запустите сверку ещё раз.", path.name)
        return 0
    return added


# ---------------------------------------------------------------- переданное
def load_journal(path: Path) -> dict[str, dict]:
    if not path or not path.exists():
        return {}
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return {row["Ключ"]: row for row in csv.DictReader(fh, delimiter=";") if row.get("Ключ")}


def _to_float(v) -> float | None:
    try:
        return float(str(v).replace("\xa0", "").replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return None


def apply_journal(records: list[ReconciledRecord], journal: dict[str, dict], tol_pct: float = 1.0) -> dict:
    stats = {"sent": 0, "changed": 0}
    for r in records:
        if r.status != CONFIRMED:
            continue
        j = journal.get(record_key(r.primary))
        if not j:
            continue
        old, new = _to_float(j.get("Цена, ₽")), r.primary.price
        if old and new and abs(old - new) > abs(old) * tol_pct / 100:
            r.issues.append(f"передан {j.get('Передан')}, но цена изменилась с {old:,.2f} — передать изменение"
                            .replace(",", " "))
            stats["changed"] += 1
            continue
        r.status = IN_NOSTROY
        r.issues.append(f"передан в НОСТРОЙ {j.get('Передан')} (пакет {j.get('Пакет')})")
        stats["sent"] += 1
    return stats


def mark_sent(package_csv: Path, journal_path: Path, when: date | None = None) -> int:
    """Внести договоры загруженного пакета в журнал переданного."""
    when = when or date.today()
    journal = load_journal(journal_path)
    added = 0
    with open(package_csv, encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh, delimiter=";"))
    new_rows = []
    for row in rows:
        reg = (row.get("Реестровый номер договора в ЕИС") or "").strip()
        d = (row.get("Дата заключения") or "").strip()
        try:
            iso = datetime.strptime(d, "%d.%m.%Y").date().isoformat() if d else ""
        except ValueError:
            iso = d
        key = reg or f"{(row.get('ИНН члена СРО') or '').strip()}|{(row.get('Номер извещения') or '').strip()}|{iso}"
        item = {"Ключ": key, "ИНН члена": row.get("ИНН члена СРО", ""), "Реестровый номер": reg,
                "Дата договора": d, "Цена, ₽": row.get("Цена договора, ₽", ""), "Пакет": package_csv.name,
                "Передан": when.strftime("%d.%m.%Y")}
        if key not in journal:
            added += 1
        journal[key] = item
        new_rows.append(item)
    journal_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = journal_path.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=JOURNAL_HEAD, delimiter=";")
        w.writeheader()
        for item in journal.values():
            w.writerow({k: item.get(k, "") for k in JOURNAL_HEAD})
    tmp.replace(journal_path)
    return added
