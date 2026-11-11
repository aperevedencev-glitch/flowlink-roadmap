"""Точечная загрузка договоров 44-ФЗ по списку реестровых номеров.

Сначала каждый договор запрашивается по номеру (getDocsByReestrNumber, подсистема RGK): так
приходят и его последние версии. Если ЕИС не отдаёт договоры этим способом, робот ищет их
по региону заказчика (цифры 2–3 номера) и дням после даты заключения. Найденное в отборе
проходит те же правила, что и при обычной загрузке.
"""
from __future__ import annotations

import logging
import re
import zipfile
from datetime import date, timedelta
from pathlib import Path

from .config import Config
from .eis_client import EisClient, EisError
from .members import load_members
from .parser import scan_archive
from .roles import _xml_in
from .targeted import _date, plan
from .versions import FILE_REG

log = logging.getLogger("eis2nostroy")
TRIAL = 30   # столько первых номеров проверяет, отдаёт ли ЕИС договоры по номеру


def load_targets44(path: Path, law: str = "44") -> list[dict]:
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    out, seen = [], set()
    for ws in wb.worksheets:
        head = None
        for r in ws.iter_rows(values_only=True):
            cells = [str(c or "").strip() for c in r]
            if head is None:
                if any(c.startswith("Реестровый номер") for c in cells) and any(c.startswith("Дата заключ") for c in cells):
                    head = cells
                continue
            d = dict(zip(head, r))
            reg = re.sub(r"\D", "", str(next((d[k] for k in head if k.startswith("Реестровый номер")), "") or ""))
            sign = _date(next((d[k] for k in head if k.startswith("Дата заключ")), None))
            if len(reg) != (19 if law == "44" else 23) or not sign or reg in seen:
                continue
            seen.add(reg)
            out.append({"reg": reg, "sign": sign, "region": reg[1:3], "inn": str(d.get("ИНН члена СРО") or ""),
                        "member": str(d.get("Член СРО") or "")})
    return out


def regs_in(path: Path) -> set[str]:
    found = set()
    with zipfile.ZipFile(path) as zf:
        for _, info in _xml_in(zf):
            m = FILE_REG.search(info.filename)
            if m:
                found.add(m.group(1))
    return found


def cmd_targeted44(cfg: Config, path: Path, window: int, store, keep, client=None, law: str = "44") -> tuple[int, list[str]]:
    """law="223": договоры 223-ФЗ по номеру (подсистема RD223), без поиска по дням —
    его уже делает targeted223 (11.11: 4 договора не нашлись за окно дат)."""
    inns = set(load_members(cfg.path("members")))
    targets = load_targets44(path, law)
    wanted = {t["reg"] for t in targets}
    feed = next(f for f in cfg.feeds if f.law == law and f.kind != "protocol")
    client = client or EisClient(cfg)
    seen: set[str] = set()
    added = 0
    errors: list[str] = []

    def process(archive: Path) -> None:
        nonlocal added
        try:
            if law == "44":
                seen.update(regs_in(archive) & wanted)
            contracts, _ = scan_archive(archive, law, inns)
            seen.update({c.reg_number for c in contracts} & wanted)
        except (OSError, zipfile.BadZipFile) as exc:
            log.warning("архив %s пропущен: %s", archive.name, exc)
            return
        kept = keep(cfg, contracts)
        if kept:
            added += store.add(kept)

    log.info("Точечная загрузка %s-ФЗ по номерам: в списке %s договоров", law, len(targets))
    for n, t in enumerate(targets, 1):
        try:
            for archive in client.fetch_by_reestr(feed.subsystem, t["reg"], law):
                process(archive)
        except (EisError, OSError) as exc:
            errors.append(f"договор {t['reg']}: {exc}")
            if len(errors) <= 3:
                log.warning("договор %s: %s", t["reg"], str(exc)[:160])
        if law == "44" and n == min(TRIAL, len(targets)) and not seen:
            log.warning("ЕИС не отдала ни одного из первых %s договоров по номеру. Перехожу к поиску по регионам и дням.", n)
            break
        if n % 100 == 0 or n == len(targets):
            log.info("По номеру: %s из %s, найдено в ЕИС %s, новых записей %s, ошибок %s", n, len(targets), len(seen), added, len(errors))

    missing = [t for t in targets if t["reg"] not in seen]
    if missing and law == "44":
        days = plan(missing, window, date.today() - timedelta(days=1))
        log.info("Поиск по дням: договоров %s, запросов по регионам и дням %s", len(missing), sum(len(r) for r in days.values()))
        ordered = sorted(days.items(), reverse=True)
        for n, (day, regions) in enumerate(ordered, 1):
            client.unpurge_days(feed, sorted(regions), [day])   # архивы дня могли удалить после разбора
            client.prefetch_day(feed, sorted(regions), day, errors)
            for region in sorted(regions):
                try:
                    for archive in client.fetch_day(feed, region, day):
                        process(archive)
                except (EisError, OSError) as exc:
                    errors.append(f"{day} регион {region}: {exc}")
            if n % 25 == 0 or n == len(ordered):
                log.info("По дням: %s из %s, найдено в ЕИС %s из %s, новых записей %s, ошибок %s",
                         n, len(ordered), len(seen), len(targets), added, len(errors))
    store.compact()
    saved = {c.reg_number for c in store.load() if c.law == law}
    from .cli import _excluded
    excl = {c.reg_number: why for c, why in _excluded(cfg).load() if c.law == law}

    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Договоры"
    ws.append(["Реестровый номер", "Дата заключения", "ИНН члена СРО", "Член СРО", "Найден в ЕИС", "В рабочем файле", "Исключён роботом"])
    for t in targets:
        ws.append([t["reg"], t["sign"], t["inn"], t["member"], "да" if t["reg"] in seen else "нет",
                   "да" if t["reg"] in saved else "нет", excl.get(t["reg"], "")])
        ws.cell(ws.max_row, 2).number_format = "DD.MM.YYYY"
    for col, w in zip("ABCDEFG", (25, 14, 14, 50, 13, 15, 50)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    out = cfg.path("output_dir") or Path("data/out")
    out.mkdir(parents=True, exist_ok=True)
    report = out / (f"targeted44_{date.today():%Y%m%d}.xlsx" if law == "44" else f"targeted{law}_по_номерам_{date.today():%Y%m%d}.xlsx")
    wb.save(report)
    log.info(f"Точечная загрузка {law}-ФЗ завершена: "
             "из %s договоров найдено в ЕИС %s, в рабочем файле %s, новых записей %s, ошибок %s. Отчёт: %s",
             len(targets), len(seen), len(wanted & saved), added, len(errors), report.name)
    return added, errors
