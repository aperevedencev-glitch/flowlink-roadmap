"""Точечная загрузка договоров 223-ФЗ по известным реестровым номерам.

Список — любой Excel, где есть колонки «Реестровый номер» (23 цифры) и «Дата заключения»,
например «Договоры 223-ФЗ с 01.07.2017 — проверка роботом.xlsx».

Вместо обхода всех 86 регионов за каждый день робот запрашивает только регион заказчика
(цифры 2–3 реестрового номера — начало ИНН заказчика) и только дни от даты заключения
до даты + окно (договор размещают в ЕИС в течение нескольких рабочих дней). Протоколы
закупок запрашиваются только для договоров из списка. Это часы вместо суток.
"""
from __future__ import annotations

import io
import logging
import re
import zipfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from .config import Config
from .eis_client import EisClient, EisError
from .members import load_members
from .parser import scan_xml
from .protocols import WinIndex

log = logging.getLogger("eis2nostroy")
REG223 = re.compile(rb"(?<!\d)\d{23}(?!\d)")


def _date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v or "").strip()[:10]
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def load_targets(path: Path) -> list[dict]:
    """Договоры 223-ФЗ из Excel: реестровый номер, дата заключения, член СРО (если есть)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    out, seen = [], set()
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        head = None
        for r in rows:
            cells = [str(c or "").strip() for c in r]
            if head is None:
                if any(c.startswith("Реестровый номер") for c in cells) and any(c.startswith("Дата заключ") for c in cells):
                    head = cells
                continue
            d = dict(zip(head, r))
            reg = re.sub(r"\D", "", str(next((d[k] for k in head if k.startswith("Реестровый номер")), "") or ""))
            sign = _date(next((d[k] for k in head if k.startswith("Дата заключ")), None))
            if len(reg) != 23 or not sign or reg in seen:
                continue
            seen.add(reg)
            out.append({"reg": reg, "sign": sign, "region": reg[1:3],
                        "member": str(d.get("Член СРО") or ""), "inn": str(d.get("ИНН члена СРО") or "")})
    return out


def plan(targets: list[dict], window: int, last: date) -> dict[date, set[str]]:
    days: dict[date, set[str]] = defaultdict(set)
    for t in targets:
        for k in range(window + 1):
            d = t["sign"] + timedelta(days=k)
            if d <= last:
                days[d].add(t["region"])
    return days


def cmd_targeted223(cfg: Config, path: Path, window: int, store, wins_store, keep, make_lookup) -> tuple[int, list[str]]:
    members = load_members(cfg.path("members"))
    inns = set(members)
    have = {c.reg_number for c in store.load() if c.law == "223"}
    all_targets = load_targets(path)
    targets = [t for t in all_targets if t["reg"] not in have]
    wanted = {t["reg"] for t in targets}
    feed = next(f for f in cfg.feeds if f.law == "223" and f.kind != "protocol")
    days = plan(targets, window, date.today() - timedelta(days=1))
    pairs = sum(len(r) for r in days.values())
    log.info("Точечная загрузка 223-ФЗ: в списке %s договоров, уже есть %s, ищем %s. "
             "Запросов по регионам и дням: %s (вместо %s при полном обходе этих дней)",
             len(all_targets), len(all_targets) - len(targets), len(targets), pairs * len(feed.document_types), len(days) * 86 * len(feed.document_types))
    if not targets:
        return 0, []
    client = EisClient(cfg)
    stats: dict = {}
    errors: list[str] = []
    lookup = make_lookup(client, inns, stats, errors)
    index = WinIndex(wins_store.load())
    found_regs: set[str] = set()   # сохранены (договор члена под ГрК)
    seen_regs: set[str] = set()    # найдены в ЕИС
    added = 0
    failed_days: set[date] = set()

    def scan_day(day: date, regions: set[str]):
        nonlocal added
        client.unpurge_days(feed, sorted(regions), [day])   # архивы дня могли удалить после разбора
        client.prefetch_day(feed, sorted(regions), day, errors)
        for region in sorted(regions):
            try:
                archives = client.fetch_day(feed, region, day)
            except (EisError, OSError) as exc:
                errors.append(f"{day} {feed.subsystem} регион {region}: {exc}")
                failed_days.add(day)
                continue
            for archive in archives:
                contracts, wins = [], []
                try:
                    with zipfile.ZipFile(archive) as zf:
                        stack = [(zf, f"{archive.name}:")]
                        while stack:
                            z, prefix = stack.pop()
                            for info in z.infolist():
                                name = info.filename
                                if name.lower().endswith(".zip"):
                                    stack.append((zipfile.ZipFile(io.BytesIO(z.read(info))), f"{prefix}{name}/"))
                                elif name.lower().endswith(".xml"):
                                    data = z.read(info)
                                    hit = {m.decode() for m in REG223.findall(data)} & wanted
                                    seen_regs.update(hit)
                                    # протоколы запрашиваем только для договоров из списка
                                    c, w = scan_xml(data, "223", inns, index, f"{prefix}{name}", lookup if hit else None)
                                    contracts += c
                                    wins += w
                except (OSError, zipfile.BadZipFile) as exc:
                    log.warning("архив %s пропущен: %s", archive.name, exc)
                    continue
                kept = keep(cfg, contracts)
                if kept:
                    added += store.add(kept)
                    found_regs.update(c.reg_number for c in kept if c.reg_number in wanted)
                if wins:
                    wins_store.add(wins)

    ordered = sorted(days.items(), key=lambda x: x[0], reverse=True)
    for n, (day, regions) in enumerate(ordered, 1):
        scan_day(day, regions)
        if n % 25 == 0 or n == len(ordered):
            log.info("Точечная 223-ФЗ: дней %s из %s, найдено в ЕИС %s из %s, сохранено под ГрК %s, "
                     "проверено закупок %s, ошибок %s", n, len(ordered), len(seen_regs), len(targets),
                     len(found_regs), stats.get("lookups", 0), len(errors))
    if failed_days:
        log.info("Повторяю дни с ошибками ЕИС: %s", len(failed_days))
        retry, failed_days = sorted(failed_days), set()
        errors.clear()
        for day in retry:
            scan_day(day, days[day])

    # что не нашлось — в Excel, чтобы проверить вручную или окном пошире
    missing = [t for t in targets if t["reg"] not in seen_regs]
    if missing:
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Не найдены"
        ws.append(["Реестровый номер", "Дата заключения", "Регион заказчика", "ИНН члена СРО", "Член СРО", "Карточка в ЕИС"])
        for t in missing:
            ws.append([t["reg"], t["sign"], t["region"], t["inn"], t["member"],
                       f"https://zakupki.gov.ru/223/contract/public/contract/view/general-information.html?id=&regNumber={t['reg']}"])
            ws.cell(ws.max_row, 2).number_format = "DD.MM.YYYY"
        for col, w in zip("ABCDEF", (26, 14, 10, 14, 50, 60)):
            ws.column_dimensions[col].width = w
        out = cfg.path("output_dir") or Path("data/out")
        out.mkdir(parents=True, exist_ok=True)
        suffix = "" if window == 10 else f"_окно{window}"   # исходный файл при повторном поиске не трогаем
        p = out / f"targeted223_не_найдены{suffix}_{date.today():%Y%m%d}.xlsx"
        try:
            wb.save(p)
        except PermissionError:
            log.warning("Файл %s открыт в Excel — список ненайденных не сохранён", p.name)
        log.info("Не найдено в ЕИС за %s дн. после даты заключения: %s договоров — список в %s "
                 "(размещены позже окна или у заказчика другой регион).", window, len(missing), p.name)
    log.info("Точечная загрузка 223-ФЗ завершена: из %s договоров найдено в ЕИС %s, сохранено под ГрК %s "
             "(остальные — не строительный подряд), новых договоров членов всего %s, ошибок %s",
             len(targets), len(seen_regs), len(found_regs), added, len(errors))
    return added, errors
