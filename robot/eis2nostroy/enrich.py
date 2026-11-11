"""Доработка договоров 223-ФЗ: предмет из извещения и отчёт по списку договоров.

В формате contractCutted предмета договора нет, поэтому сверка не может ни подтвердить
строительный подряд, ни отсеять текущий ремонт. Предмет есть в извещении о закупке:
purchaseNoticeData/name и lots/lot/lotData/subject (getDocsByReestrNumber, подсистема RI223).

subjects223 запрашивает извещения (с кэшем) и дописывает предмет в рабочий файл.
После rebuild предметы теряются: команду нужно повторить, из кэша это быстро.

audit223 ничего не запрашивает и не меняет: разбирает скачанные архивы и пишет,
что случилось с каждым договором из списка.
"""
from __future__ import annotations

import io
import json
import logging
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import date, datetime
from pathlib import Path

from .config import Config
from .eis_client import EisClient, EisError
from .grk import NOT_GRK, current_repair_reason
from .members import load_members
from .parser import grk_info, parse_contract_xml
from .targeted import REG223, load_targets, plan

log = logging.getLogger("eis2nostroy")


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _child_text(el: ET.Element, *path: str) -> str:
    for name in path:
        el = next((c for c in el if _local(c.tag) == name), None)
        if el is None:
            return ""
    return (el.text or "").strip()


def notice_subject(data: bytes) -> tuple[str, int]:
    """Предмет закупки из XML извещения и число лотов; ('', 0), если это не извещение."""
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return "", 0
    if not _local(root.tag).lower().startswith("purchasenotice"):
        return "", 0
    for item in root.iter():
        if not _local(item.tag).lower().startswith("purchasenotice") or not _local(item.tag).endswith("Data"):
            continue
        name = _child_text(item, "name")
        lots = [e for e in item.iter() if _local(e.tag) == "lotData"]
        if len(lots) == 1:
            return (_child_text(lots[0], "subject") or name)[:500], 1
        return name[:500], len(lots)
    return "", 0


def subject_from_archives(paths: list[Path]) -> tuple[str, int]:
    for path in paths:
        with zipfile.ZipFile(path) as zf:
            for name in sorted(zf.namelist()):
                if name.lower().endswith(".xml") and "notice" in name.lower():
                    subject, lots = notice_subject(zf.read(name))
                    if subject:
                        return subject, lots
    return "", 0


def cmd_subjects223(cfg: Config, store_path: Path) -> tuple[int, list[str]]:
    """Дописать предмет договорам 223-ФЗ из извещений. Возвращает (сколько дописано, ошибки)."""
    if not store_path.exists():
        log.error("Нет файла договоров %s", store_path)
        return 0, ["нет файла договоров"]
    with open(store_path, encoding="utf-8") as fh:
        rows = [json.loads(line) for line in fh if line.strip()]
    since = cfg.competitive_from.isoformat()
    todo = [r for r in rows if str(r.get("law")) == "223" and not (r.get("subject") or "").strip()
            and r.get("notice_number") and (r.get("sign_date") or "") >= since]
    notices = sorted({r["notice_number"] for r in todo})
    log.info("Договоров 223-ФЗ без предмета: %s, извещений к запросу: %s", len(todo), len(notices))
    client = EisClient(cfg)
    found: dict[str, tuple[str, int]] = {}
    errors: list[str] = []
    for n, notice in enumerate(notices, 1):
        try:
            subject, lots = subject_from_archives(client.fetch_by_reestr("RI223", notice))
        except (EisError, OSError, zipfile.BadZipFile) as exc:
            errors.append(f"закупка {notice}: {exc}")
            log.error("закупка %s: %s", notice, exc)
            continue
        if subject:
            found[notice] = (subject, lots)
        if n % 25 == 0 or n == len(notices):
            log.info("Извещения: %s из %s, предмет найден в %s, ошибок %s", n, len(notices), len(found), len(errors))
    filled = 0
    for r in todo:
        hit = found.get(r["notice_number"])
        if hit:
            r["subject"] = hit[0]
            r.setdefault("extra", {})["subject_source"] = "извещение" if hit[1] <= 1 else f"извещение, лотов {hit[1]}"
            filled += 1
    if filled:
        backup = store_path.with_name(f"{store_path.stem}_{datetime.now():%Y%m%d_%H%M%S}.bak.jsonl")
        shutil.copy2(store_path, backup)
        tmp = store_path.with_suffix(".tmp")
        with open(tmp, "w", encoding="utf-8") as fh:
            for r in rows:
                fh.write(json.dumps(r, ensure_ascii=False) + "\n")
        os.replace(tmp, store_path)
    multi = sum(1 for r in todo if "лотов" in (r.get("extra") or {}).get("subject_source", ""))
    log.info("Предмет дописан договорам: %s из %s (из них по общему названию многолотовой закупки: %s). "
             "Без извещения в ЕИС: %s", filled, len(todo), multi, len(todo) - filled)
    return filled, errors


def _xml_version(name: str) -> int:
    m = re.search(r"_(\d+)_[0-9A-Fa-f]{32}\.xml$", name)
    return int(m.group(1)) if m else 0


def _inspect(data: bytes, inns: set[str], name: str) -> dict | None:
    """Что сказано в XML договора: его номер, поставщик, ОКПД2 и вывод робота."""
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return None
    reg = next(((e.text or "").strip() for e in root.iter() if _local(e.tag) == "contractRegNumber"), "")
    suppliers = [(c.text or "").strip() for e in root.iter() if _local(e.tag) == "supplierInfo"
                 for c in e if _local(c.tag) == "inn"]
    okpd_names = [(c.text or "").strip() for e in root.iter() if _local(e.tag) == "okpd2"
                  for c in e if _local(c.tag) == "name"]
    found = parse_contract_xml(data, "223", inns, name)
    info = found[0].extra if found else grk_info(root)
    return {"reg": reg, "version": _xml_version(name), "suppliers": [s for s in suppliers if s],
            "member": bool(found), "okpd2": info.get("okpd2", []), "okpd_name": okpd_names[0] if okpd_names else "",
            "grk": info.get("grk", ""), "note": info.get("grk_note", ""),
            "subject": found[0].subject if found else ""}


def reason_of(x: dict | None, target_inn: str) -> tuple[str, str]:
    """(итог, пояснение) по договору из списка."""
    if x is None:
        return "Нет в скачанных архивах", "за окно дат в регионе заказчика договор не встретился"
    if not x["suppliers"]:
        return "В ЕИС поставщик скрыт", "в XML нет ИНН поставщика: связать с членом можно только через протокол"
    if not x["member"]:
        return "Поставщик не член СРО", f"в договоре ИНН {', '.join(x['suppliers'])}, в списке {target_inn}"
    if x["grk"] == NOT_GRK:
        return "Не строительные работы", f"ОКПД2 {', '.join(x['okpd2'][:3])}: {x['okpd_name']}"
    why = current_repair_reason(x["subject"])
    if why:
        return "Текущий ремонт или содержание", why
    return "Сохранён", x["grk"] or ""


def cmd_audit223(cfg: Config, path: Path, window: int, store) -> Path:
    """Отчёт по каждому договору списка. Только чтение кэша, ЕИС не запрашивается."""
    from openpyxl import Workbook
    inns = set(load_members(cfg.path("members")))
    targets = load_targets(path)
    wanted = {t["reg"] for t in targets}
    saved = {c.reg_number for c in store.load() if c.law == "223"}
    feed = next(f for f in cfg.feeds if f.law == "223" and f.kind != "protocol")
    cache = cfg.path("cache_dir") or Path("data/cache")
    days = plan(targets, window, date.today())
    best: dict[str, dict] = {}
    ordered = sorted(days.items())
    log.info("Отчёт по списку: договоров %s, дней %s. ЕИС не запрашивается.", len(targets), len(ordered))
    for n, (day, regions) in enumerate(ordered, 1):
        for region in sorted(regions):
            for doc_type in feed.document_types:
                for archive in sorted((cache / feed.subsystem / region / day.isoformat() / doc_type).glob("*.zip")):
                    try:
                        with zipfile.ZipFile(archive) as zf:
                            stack = [zf]
                            while stack:
                                z = stack.pop()
                                for info in z.infolist():
                                    name = info.filename
                                    if name.lower().endswith(".zip"):
                                        stack.append(zipfile.ZipFile(io.BytesIO(z.read(info))))
                                    elif name.lower().endswith(".xml"):
                                        data = z.read(info)
                                        hit = {m.decode() for m in REG223.findall(data)} & wanted
                                        if not hit:
                                            continue
                                        x = _inspect(data, inns, name)
                                        if x and x["reg"] in wanted and x["version"] >= best.get(x["reg"], {"version": -1})["version"]:
                                            best[x["reg"]] = x
                    except (OSError, zipfile.BadZipFile) as exc:
                        log.warning("архив %s пропущен: %s", archive.name, exc)
        if n % 50 == 0 or n == len(ordered):
            log.info("Отчёт: дней %s из %s, договоров списка найдено %s", n, len(ordered), len(best))
    wb = Workbook()
    ws = wb.active
    ws.title = "Договоры"
    ws.append(["Итог", "Пояснение", "Реестровый номер", "Дата заключения", "ИНН члена СРО", "Член СРО",
               "Поставщик в ЕИС", "Класс ГрК", "ОКПД2", "В рабочем файле"])
    totals: Counter = Counter()
    by_member: Counter = Counter()
    for t in targets:
        x = best.get(t["reg"])
        verdict, note = reason_of(x, t["inn"])
        if t["reg"] in saved and verdict != "Сохранён":
            note = f"{note}; в рабочем файле договор есть из другого источника"
        totals[verdict] += 1
        by_member[(verdict, t["member"])] += 1
        ws.append([verdict, note, t["reg"], t["sign"], t["inn"], t["member"],
                   ", ".join(x["suppliers"]) if x else "", x["grk"] if x else "",
                   ", ".join(x["okpd2"]) if x else "", "да" if t["reg"] in saved else "нет"])
        ws.cell(ws.max_row, 4).number_format = "DD.MM.YYYY"
    for col, w in zip("ABCDEFGHIJ", (30, 60, 26, 14, 14, 44, 16, 14, 24, 12)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    s = wb.create_sheet("Сводка", 0)
    s.append(["Итог", "Договоров"])
    for k, v in totals.most_common():
        s.append([k, v])
    s.append([])
    s.append(["Итог", "Член СРО", "Договоров"])
    for (verdict, member), v in sorted(by_member.items(), key=lambda kv: (kv[0][0], -kv[1])):
        s.append([verdict, member, v])
    s.column_dimensions["A"].width = 32
    s.column_dimensions["B"].width = 50
    out = cfg.path("output_dir") or Path("data/out")
    out.mkdir(parents=True, exist_ok=True)
    p = out / f"audit223_{date.today():%Y%m%d}.xlsx"
    wb.save(p)
    log.info("Отчёт по списку готов: %s. %s", p.name, "; ".join(f"{k}: {v}" for k, v in totals.most_common()))
    return p
