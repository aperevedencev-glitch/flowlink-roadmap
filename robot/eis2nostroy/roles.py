"""Какую роль член СРО играл в договорах, выпавших из рабочего файла при пересборке.

Запрашивать ЕИС не нужно: список берётся из резервной копии рабочего файла, а XML договоров —
из скачанных архивов. Команда только читает, рабочий файл и решения оператора не меняет.
Договор считается по последней версии XML, которая есть в кэше.
"""
from __future__ import annotations

import io
import json
import logging
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import date, timedelta
from pathlib import Path
from typing import Callable, Iterator

from .config import Config
from .decisions import load_journal
from .members import load_members
from .parser import SUPPLIER_BLOCKS, _blocks, _first_text, _local, _party_inns, parse_contract_xml

log = logging.getLogger("eis2nostroy")

SUPPLIER = "Поставщик"
SUBCONTRACTOR = "Субподрядчик"
CUSTOMER = "Заказчик"
OTHER = "Другая роль"
ABSENT = "ИНН в договоре нет"
NOT_CACHED = "Нет в кэше"

FILE_REG = re.compile(r"contract_(\d+)_(\d+)_[0-9A-Fa-f]{32}\.xml$")


def _is_sub(part: str) -> bool:
    return part.startswith("sub") and "subject" not in part


def inn_paths(root: ET.Element, inn: str) -> list[str]:
    """Все места XML, где стоит этот ИНН: путь из имён элементов без корня."""
    found: list[str] = []

    def walk(el: ET.Element, trail: list[str]) -> None:
        for child in el:
            name = child.tag.rsplit("}", 1)[-1]
            if name.lower() == "inn" and (child.text or "").strip() == inn:
                found.append("/".join(trail))
            walk(child, trail + [name])
    walk(root, [])
    return found


def role_of(paths: list[str]) -> str:
    """Сторона договора важнее субподряда: член может быть и тем, и другим."""
    kinds = set()
    for path in paths:
        parts = [p.lower() for p in path.split("/")]
        if any(_is_sub(p) for p in parts):
            kinds.add(SUBCONTRACTOR)
        elif any(k in p for p in parts for k in SUPPLIER_BLOCKS):
            kinds.add(SUPPLIER)
        elif any(k in p for p in parts for k in ("customer", "placer")):
            kinds.add(CUSTOMER)
        else:
            kinds.add(OTHER)
    for kind in (SUPPLIER, SUBCONTRACTOR, CUSTOMER, OTHER):
        if kind in kinds:
            return kind
    return ABSENT


def real_suppliers(root: ET.Element) -> list[str]:
    """ИНН и наименование сторон-поставщиков, без субподрядчиков."""
    out = []
    for block in _blocks(root, SUPPLIER_BLOCKS):
        if _is_sub(_local(block.tag)):
            continue
        name = _first_text(block, ("fullname", "shortname", "name"))
        for inn in dict.fromkeys(_party_inns(block)):
            out.append(f"{inn} {name}".strip())
    return list(dict.fromkeys(out))


def inspect(data: bytes, inn: str, name: str, keep: Callable[[list], list]) -> dict | None:
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return None
    paths = inn_paths(root, inn)
    role = role_of(paths)
    note = ""
    if role == SUPPLIER:
        found = parse_contract_xml(data, "44", {inn}, name)
        if not found:
            note = "разбор не считает его поставщиком: проверить XML вручную"
        elif keep(found):
            note = "договор проходит правила: его не должно было убрать, проверить вручную"
        else:
            note = "отсеян правилами ГрК: " + (found[0].extra.get("grk_note") or found[0].subject or "текущий ремонт")
    return {"role": role, "suppliers": real_suppliers(root), "paths": list(dict.fromkeys(paths))[:3], "note": note}


def _xml_in(archive: zipfile.ZipFile) -> Iterator[tuple[zipfile.ZipFile, zipfile.ZipInfo]]:
    for info in archive.infolist():
        low = info.filename.lower()
        if low.endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(archive.read(info))) as inner:
                yield from _xml_in(inner)
        elif low.endswith(".xml"):
            yield archive, info


def scan_cache(cache: Path, wanted: dict[str, dict], since: date, keep: Callable[[list], list]) -> dict[str, dict]:
    """Последняя версия XML каждого договора из wanted (реестровый номер → строка резервной копии)."""
    archives = [z for z in sorted((cache / "RGK").rglob("*.zip"))
                if z.parent.name == "contract" and z.parent.parent.name >= since.isoformat()]
    log.info("Архивов 44-ФЗ к просмотру: %s", len(archives))
    best: dict[str, dict] = {}
    for n, path in enumerate(archives, 1):
        try:
            with zipfile.ZipFile(path) as zf:
                for owner, info in _xml_in(zf):
                    m = FILE_REG.search(info.filename)
                    if not m or m.group(1) not in wanted:
                        continue
                    reg, version = m.group(1), int(m.group(2))
                    if reg in best and best[reg]["version"] > version:
                        continue
                    seen = inspect(owner.read(info), wanted[reg]["supplier_inn"], info.filename, keep)
                    if seen:
                        best[reg] = {**seen, "version": version}
        except (OSError, zipfile.BadZipFile) as exc:
            log.warning("архив %s пропущен: %s", path.name, exc)
        if n % 2000 == 0 or n == len(archives):
            log.info("Просмотрено архивов %s из %s, найдено договоров %s из %s", n, len(archives), len(best), len(wanted))
    return best


def load_rows(path: Path) -> list[dict]:
    with open(path, encoding="utf-8") as fh:
        return [json.loads(line) for line in fh if line.strip()]


def _reg(row: dict) -> str:
    m = FILE_REG.search(row.get("raw_file") or "")
    return m.group(1) if m else (row.get("reg_number") or "")


def gone_contracts(old: list[dict], new: list[dict], law: str = "44") -> dict[str, dict]:
    """Договоры закона law из резервной копии, которых нет в рабочем файле."""
    now = {(r["supplier_inn"], _reg(r)) for r in new}
    return {_reg(r): r for r in old if str(r.get("law")) == law and _reg(r) and (r["supplier_inn"], _reg(r)) not in now}


def pick_backup(store_path: Path, new: list[dict]) -> Path | None:
    """Самая свежая резервная копия, в которой есть договоры, отсутствующие в рабочем файле."""
    copies = sorted(store_path.parent.glob(f"{store_path.stem}*.bak.jsonl"), key=lambda p: p.stat().st_mtime, reverse=True)
    return next((p for p in copies if gone_contracts(load_rows(p), new)), None)


def last_package(out_dir: Path, regs: set[str]) -> dict[str, str]:
    """Самый поздний пакет НОСТРОЙ, куда входил договор."""
    res: dict[str, str] = {}
    for path in sorted(out_dir.glob("nostroy_*.csv"), key=lambda p: p.stat().st_mtime):
        data = path.read_bytes()
        for reg in regs:
            if reg.encode() in data:
                res[reg] = path.name
    return res


def cmd_roles44(cfg: Config, store_path: Path, journal_path: Path, keep: Callable[[list], list],
                backup: Path | None = None) -> Path | None:
    from openpyxl import Workbook
    new = load_rows(store_path)
    backup = backup or pick_backup(store_path, new)
    if not backup or not backup.exists():
        log.error("Нет резервной копии с договорами 44-ФЗ, которых нет в рабочем файле. Отчёт не построен.")
        return None
    gone = gone_contracts(load_rows(backup), new)
    log.info("Резервная копия %s: договоров 44-ФЗ вне рабочего файла %s. ЕИС не запрашивается.", backup.name, len(gone))
    dates = [date.fromisoformat(r["sign_date"]) for r in gone.values() if r.get("sign_date")]
    since = min(dates) - timedelta(days=3) if dates else date(2000, 1, 1)
    best = scan_cache(cfg.path("cache_dir") or Path("data/cache"), gone, since, keep)
    out = cfg.path("output_dir") or Path("data/out")
    out.mkdir(parents=True, exist_ok=True)
    packages = last_package(out, set(gone))
    journal = " ".join(load_journal(journal_path))
    members = load_members(cfg.path("members"))

    wb = Workbook()
    ws = wb.active
    ws.title = "Договоры"
    ws.append(["Роль члена СРО", "Пояснение", "Реестровый номер", "Дата заключения", "Цена, ₽", "ИНН члена СРО", "Член СРО",
               "Поставщик по договору", "Заказчик", "Предмет", "Версия XML", "Где стоит ИНН члена",
               "Последний пакет НОСТРОЙ", "В журнале переданных"])
    totals: Counter = Counter()
    money: Counter = Counter()
    ranked = sorted(gone.items(), key=lambda kv: (best.get(kv[0], {}).get("role", NOT_CACHED), -(kv[1].get("price") or 0)))
    for reg, row in ranked:
        x = best.get(reg)
        role = x["role"] if x else NOT_CACHED
        totals[role] += 1
        money[role] += row.get("price") or 0
        inn = row["supplier_inn"]
        member = members.get(inn)
        sign = date.fromisoformat(row["sign_date"]) if row.get("sign_date") else None
        ws.append([role, x["note"] if x else "XML договора нет в скачанных архивах", reg, sign, row.get("price"), inn,
                   member.name if member else "", "; ".join(x["suppliers"]) if x else "", row.get("customer_name", ""),
                   row.get("subject", ""), x["version"] if x else "", " | ".join(x["paths"]) if x else "",
                   packages.get(reg, ""), "да" if reg in journal else "нет"])
        ws.cell(ws.max_row, 4).number_format = "DD.MM.YYYY"
        ws.cell(ws.max_row, 5).number_format = "#,##0.00"
    for col, w in zip("ABCDEFGHIJKLMN", (20, 48, 23, 14, 18, 14, 40, 44, 44, 60, 9, 44, 24, 12)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    s = wb.create_sheet("Сводка", 0)
    s.append(["Резервная копия", backup.name])
    s.append([])
    s.append(["Роль члена СРО", "Договоров", "Сумма, ₽"])
    for role, count in totals.most_common():
        s.append([role, count, money[role]])
        s.cell(s.max_row, 3).number_format = "#,##0.00"
    s.column_dimensions["A"].width = 24
    s.column_dimensions["B"].width = 34
    s.column_dimensions["C"].width = 22
    path = out / f"roles44_{date.today():%Y%m%d}.xlsx"
    wb.save(path)
    log.info("Отчёт о ролях готов: %s. %s", path.name, "; ".join(f"{k}: {v}" for k, v in totals.most_common()))
    return path
