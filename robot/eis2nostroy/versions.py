"""Проверка перед загрузкой: не устарела ли запись о договоре.

В рабочий файл попадает только версия договора, прошедшая отбор: поставщик — член СРО, работы
по ГрК. Если новая версия отбор не прошла, она не сохраняется, а прежняя запись остаётся и уходит
в пакет. Команда находит такие договоры: номер версии берётся из имён файлов в архивах, поэтому
архивы не распаковываются целиком. Только чтение, ЕИС не запрашивается, рабочий файл не меняется.
"""
from __future__ import annotations

import logging
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import date
from pathlib import Path

from .config import Config
from .grk import NOT_GRK, current_repair_reason
from .members import load_members
from .parser import SUBJECT, _first_text, grk_info
from .roles import SUPPLIER, _xml_in, inn_paths, last_package, real_suppliers, role_of
from .sources import ContractStore, file_version

log = logging.getLogger("eis2nostroy")

NOT_PARTY = "Член больше не поставщик"
REPAIR = "Стал текущим ремонтом или содержанием"
NOT_BUILD = "Стал не по ГрК"
HIDDEN = "Поставщик скрыт в новой версии"
MANUAL = "Проверить вручную"

FILE_REG = re.compile(r"(?:contract|contractCutted|pprf615Contract)_(\d+)_(\d+)_[0-9A-Fa-f]{32}\.xml$")


def verdict(data: bytes, inn: str, law: str) -> tuple[str, str]:
    """(вывод, пояснение) по новой версии договора, которая не вошла в рабочий файл."""
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return MANUAL, "XML не разобран"
    suppliers = real_suppliers(root)
    if suppliers and role_of(inn_paths(root, inn)) != SUPPLIER:
        return NOT_PARTY, "поставщик в новой версии: " + "; ".join(suppliers)
    subject = _first_text(root, SUBJECT)[:500]
    why = current_repair_reason(subject)
    if why:
        return REPAIR, f"{why}: {subject}"
    if law != "615":
        info = grk_info(root, subject)
        if info["grk"] == NOT_GRK:
            return NOT_BUILD, info["grk_note"]
    if not suppliers:
        return HIDDEN, "в XML нет ИНН поставщика: связать с членом можно только через протокол"
    return MANUAL, "новая версия не вошла в рабочий файл, причина не определена"


def newer_versions(cache: Path, rows: dict[tuple, object]) -> dict[str, dict]:
    """Договоры, у которых в кэше есть версия новее записанной. rows: (подсистема, номер) → договор."""
    subsystems = sorted({s for s, _ in rows})
    archives = [z for sub in subsystems for base in (cache / sub, cache / "byreestr" / sub) for z in sorted(base.rglob("*.zip"))]
    log.info("Архивов к просмотру: %s", len(archives))
    found: dict[str, dict] = {}
    for n, path in enumerate(archives, 1):
        parts = path.relative_to(cache).parts
        sub = parts[1] if parts[0] == "byreestr" else parts[0]
        try:
            with zipfile.ZipFile(path) as zf:
                for owner, info in _xml_in(zf):
                    m = FILE_REG.search(info.filename)
                    c = rows.get((sub, m.group(1))) if m else None
                    if c is None:
                        continue
                    version = int(m.group(2))
                    if version <= file_version(c.raw_file) or version <= found.get(c.reg_number, {"version": -1})["version"]:
                        continue
                    what, note = verdict(owner.read(info), c.supplier_inn, c.law)
                    found[c.reg_number] = {"version": version, "verdict": what, "note": note}
        except (OSError, zipfile.BadZipFile) as exc:
            log.warning("архив %s пропущен: %s", path.name, exc)
        if n % 5000 == 0 or n == len(archives):
            log.info("Просмотрено архивов %s из %s, договоров с новой версией %s", n, len(archives), len(found))
    return found


def cmd_versions(cfg: Config, store: ContractStore) -> tuple[Path, int]:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    subsystem = {f.law: f.subsystem for f in cfg.feeds if f.kind != "protocol"}
    contracts = [c for c in store.load() if c.reg_number and c.law in subsystem]
    rows = {(subsystem[c.law], c.reg_number): c for c in contracts}
    log.info("Договоров к проверке: %s. ЕИС не запрашивается.", len(rows))
    found = newer_versions(cfg.path("cache_dir") or Path("data/cache"), rows)
    out = cfg.path("output_dir") or Path("data/out")
    out.mkdir(parents=True, exist_ok=True)
    packages = sorted(out.glob("nostroy_*.csv"), key=lambda p: p.stat().st_mtime)
    newest = packages[-1].name if packages else ""
    in_package = last_package(out, set(found))
    members = load_members(cfg.path("members"))
    wb = Workbook()
    ws = wb.active
    ws.title = "К проверке"
    ws.append(["Вывод", "Пояснение", "Закон", "Реестровый номер", "ИНН члена СРО", "Член СРО", "Заказчик", "Цена в файле, ₽",
               "Предмет в файле", "Версия в файле", "Версия в кэше", "В последнем пакете"])
    totals: Counter = Counter()
    by_reg = {c.reg_number: c for c in contracts}
    order = sorted(found.items(), key=lambda kv: (kv[1]["verdict"], -(by_reg[kv[0]].price or 0)))
    for reg, x in order:
        c = by_reg[reg]
        totals[x["verdict"]] += 1
        member = members.get(c.supplier_inn)
        ws.append([x["verdict"], x["note"], c.law, reg, c.supplier_inn, member.name if member else "", c.customer_name, c.price,
                   c.subject, file_version(c.raw_file), x["version"], "да" if in_package.get(reg) == newest else "нет"])
        ws.cell(ws.max_row, 8).number_format = "#,##0.00"
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="14523B")
    for col, w in zip("ABCDEFGHIJKL", (34, 60, 8, 25, 14, 36, 36, 18, 60, 10, 10, 12)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    s = wb.create_sheet("Сводка", 0)
    s.append(["Договоров проверено", len(rows)])
    s.append(["С версией новее записанной", len(found)])
    s.append([])
    s.append(["Вывод", "Договоров"])
    for k, v in totals.most_common():
        s.append([k, v])
    s.column_dimensions["A"].width = 40
    s.column_dimensions["B"].width = 12
    path = out / f"versions_{date.today():%Y%m%d}.xlsx"
    wb.save(path)
    log.info("Проверка версий готова: %s. Проверено %s, к проверке %s. %s", path.name, len(rows), len(found),
             "; ".join(f"{k}: {v}" for k, v in totals.most_common()))
    return path, len(found)
