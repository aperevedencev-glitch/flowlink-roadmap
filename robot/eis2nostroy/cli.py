"""Командная строка.

    python -m eis2nostroy fetch --from 2026-10-01 --to 2026-10-15
    python -m eis2nostroy import-archives --law 44 data/archives/
    python -m eis2nostroy reconcile
    python -m eis2nostroy run --days 3          # ежедневный запуск
"""
from __future__ import annotations

import argparse
import logging
import sys
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path

from .alerts import notify
from .config import Config, load_config
from .eis_client import PURGE_FLAG, PURGED, EisClient, EisError, purge_folder, unpurge_folder
from .export import write_operator_workbook
from .members import load_members
from .models import CONFIRMED, DECLARED, EXCLUDED, IN_NOSTROY, MISMATCH, NEEDS_CHECK, PROBABLE
from .decisions import apply_decisions, apply_journal, load_decisions, load_journal, mark_sent, sync_queue
from .ops import RobotBusy, keep_awake, run_lock, version_line
from .nostroy import NostroyTransport
from .parser import scan_archive
import hashlib
import json
import os
import re
import time
import zipfile
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor

from .protocols import WinIndex, WinStore, parse_protocol_xml, protocol_types_from_nsi, unlinked_wins
from .reconcile import reconcile
from .sources import ContractStore, load_table_source

log = logging.getLogger("eis2nostroy")


def _store(cfg: Config) -> ContractStore:
    return ContractStore((cfg.path("work_dir") or Path("data/work")) / "eis_contracts.jsonl")


def _excluded(cfg: Config):
    from .excluded import ExcludedLog
    return ExcludedLog((cfg.path("work_dir") or Path("data/work")) / "excluded.jsonl")


def _wins(cfg: Config) -> WinStore:
    return WinStore((cfg.path("work_dir") or Path("data/work")) / "protocol_wins.jsonl")


def _data_dir(cfg: Config) -> Path:
    m = cfg.path("members")
    return m.parent if m else cfg.root / "data"


def _decisions_path(cfg: Config) -> Path:
    return cfg.path("decisions") or _data_dir(cfg) / "decisions.xlsx"


def _journal_path(cfg: Config) -> Path:
    return cfg.path("transferred") or _data_dir(cfg) / "peredano.csv"


def _lock_path(cfg: Config) -> Path:
    return (cfg.path("work_dir") or Path("data/work")) / "robot.lock"


def _types_file(cfg: Config) -> Path:
    return (cfg.path("work_dir") or Path("data/work")) / "protocol_types.json"


def _apply_protocol_types(cfg: Config) -> None:
    """Типы протоколов из справочника ЕИС (сохраняет команда probe)."""
    f = _types_file(cfg)
    if not f.exists():
        return
    types = json.loads(f.read_text(encoding="utf-8")).get("protocols") or []
    if types:
        for feed in cfg.feeds:
            if feed.kind == "protocol":
                feed.document_types = list(types)


def _load_nsi_protocol_types(cfg: Config, client: EisClient) -> list[str]:
    folder = (cfg.path("cache_dir") or Path("data/cache")) / "nsi"
    folder.mkdir(parents=True, exist_ok=True)
    for old in folder.glob("nsiDocumentTypes223_*.zip"):
        old.unlink()
    blobs = []
    for z in client.fetch_nsi("nsiDocumentTypes223", folder):
        with zipfile.ZipFile(z) as zf:
            blobs += [zf.read(n) for n in zf.namelist() if n.lower().endswith(".xml")]
    protocols, public = protocol_types_from_nsi(blobs)
    f = _types_file(cfg)
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps({"protocols": protocols, "public": public}, ensure_ascii=False, indent=1),
                 encoding="utf-8")
    return protocols


def _days(start: date, end: date):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def _grk_only(cfg: Config) -> bool:
    return bool(cfg.raw.get("grk", {}).get("only", True))


def _keep(cfg: Config, contracts: list) -> list:
    """Только договоры по Градостроительному кодексу (если [grk] only = true)."""
    if not _grk_only(cfg):
        return contracts
    from .grk import NOT_GRK, UNKNOWN, classify, current_repair_reason, resource_supply_reason, set_kvr_grk, set_kvr_not_grk
    set_kvr_grk(cfg.raw.get("grk", {}).get("kvr_grk"))
    set_kvr_not_grk(cfg.raw.get("grk", {}).get("kvr_not_grk"))
    kept, repair, dropped = [], 0, []
    for c in contracts:
        # 12.11: договоры 44-ФЗ с ИКЗ переклассифицируются по действующим правилам (список КВР СРО),
        # в том числе сохранённые раньше — архивы старых дней удалены, пересборка невозможна
        if c.law == "44" and c.extra.get("ikz") and c.extra.get("grk") != "ПП-615":
            cls, note = classify(c.extra["ikz"], c.extra.get("okpd2") or [], subject=c.subject)
            c.extra["grk"], c.extra["grk_note"] = cls, note
        if c.extra.get("grk") == NOT_GRK:
            dropped.append((c, "не работы по ГрК: " + (c.extra.get("grk_note") or "по коду ОКПД2 и предмету")))
            continue
        # договоры без кода ОКПД2, сохранённые до правила о ресурсоснабжении (10.11)
        if c.extra.get("grk") == UNKNOWN and resource_supply_reason(c.subject):
            repair += 1
            dropped.append((c, resource_supply_reason(c.subject) + ", кода ОКПД2 нет"))
            continue
        # текущий ремонт и содержание не сверяются с НОСТРОЙ, даже в записях,
        # сохранённых до появления этого правила
        why = current_repair_reason(c.subject) if cfg.raw.get("grk", {}).get("exclude_current_repair", True) else ""
        if why:
            repair += 1
            dropped.append((c, "текущий ремонт и содержание: " + why))
            continue
        kept.append(c)
    if dropped:
        try:
            _excluded(cfg).add(dropped)
        except OSError as exc:
            log.warning("журнал исключённых договоров не записан: %s", exc)
    if repair:
        log.info("Исключено договоров текущего ремонта, содержания и ресурсоснабжения: %s", repair)
    return kept


def _make_lookup(client: EisClient, inns: set[str], stats: dict, errors: list[str]):
    """Номер закупки → победы членов СРО в её протоколах (один запрос в ЕИС, с кэшем)."""
    memo: dict[str, list] = {}

    def lookup(notice: str) -> list:
        if notice in memo:
            return memo[notice]
        stats["lookups"] = stats.get("lookups", 0) + 1
        wins = []
        try:
            for archive in client.fetch_by_reestr("RI223", notice):
                with zipfile.ZipFile(archive) as zf:
                    for name in zf.namelist():
                        if name.lower().endswith(".xml"):
                            wins += parse_protocol_xml(zf.read(name), inns, f"{archive.name}:{name}")
        except (EisError, OSError, zipfile.BadZipFile) as exc:
            errors.append(f"закупка {notice}: {exc}")
            log.error("закупка %s: %s", notice, exc)
        memo[notice] = wins
        return wins

    return lookup


def _cached_day(cfg: Config, feed, region: str, day: date) -> list[Path]:
    cache = cfg.path("cache_dir") or Path("data/cache")
    out: list[Path] = []
    for doc_type in feed.document_types:
        folder = cache / feed.subsystem / region / day.isoformat() / doc_type
        if (folder / ".done").exists():
            out += sorted(folder.glob("*.zip"))
    return out


LAW_ORDER = {"223": 1, "44": 2, "615": 3}


def _ordered(feeds):
    """Порядок обработки: протоколы, затем 223-ФЗ, 44-ФЗ, 615-ПП."""
    return sorted(feeds, key=lambda f: 0 if f.kind == "protocol" else LAW_ORDER.get(f.law, 9))


# 07.11: разбор дня 44-ФЗ по всем регионам на компьютере СРО занимал 1,5–3 минуты, даже если
# день давно скачан и разобран. Разобранный день отмечается в scanned.json и больше не
# перебирается, пока не изменится список членов СРО или правила разбора (SCAN_VERSION).
SCAN_VERSION = "1"


def _scan_mark(cfg: Config, feed, inns: set[str]) -> str:
    h = hashlib.sha1()
    h.update(f"{SCAN_VERSION}|{feed.law}|{feed.subsystem}|{_grk_only(cfg)}|".encode())
    h.update(",".join(sorted(inns)).encode())
    return h.hexdigest()[:16]


def _keep_archives(cfg: Config) -> bool:
    """eis.keep_archives = false — удалять архивы дня сразу после разбора (см. purge_folder)."""
    return bool(cfg.eis.get("keep_archives", True))


def _scanned_file(cfg: Config, feed) -> Path:
    return (cfg.path("cache_dir") or Path("data/cache")) / feed.subsystem / "scanned.json"


def _load_scanned(cfg: Config, feed) -> dict:
    try:
        return json.loads(_scanned_file(cfg, feed).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_scanned(cfg: Config, feed, marks: dict) -> None:
    path = _scanned_file(cfg, feed)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(marks, ensure_ascii=False, sort_keys=True), encoding="utf-8")
    os.replace(tmp, path)


def _scan_job(job):
    """Разбор одного архива в отдельном процессе (для законов без проверки протоколов)."""
    path, law, inns = job
    try:
        return scan_archive(Path(path), law, inns)[0], None
    except (OSError, zipfile.BadZipFile) as exc:
        return [], f"{path}: {exc}"


def _scan_workers(cfg: Config) -> int:
    n = cfg.eis.get("scan_workers")
    if n is None:
        n = min(6, max(1, (os.cpu_count() or 2) - 1))
    return max(1, int(n))


def _fetch_one_day(cfg, client, feed, regions, day, new_folders, use_lookup,
                   inns, index, store, wins_store, errors, stats, counters, pool=None, marks=None, mark=None):
    """Разбор одного дня из кэша (архивы уже заказаны пачкой)."""
    before = (counters["found"], counters["wins"], stats.get("lookups", 0))
    key = day.isoformat()
    if marks is not None and not new_folders and marks.get(key) == mark and all(
            (((cfg.path("cache_dir") or Path("data/cache")) / feed.subsystem / r / key / t / ".done").exists())
            for r in regions for t in feed.document_types):
        counters["skipped"] = counters.get("skipped", 0) + 1
        return
    if mark is not None:
        # архивы дня удалены после разбора по другим правилам или другому списку членов — скачать заново
        client.unpurge_days(feed, regions, [day], mark)
    n_err, t0, n_arch = len(errors), time.monotonic(), 0
    # номера закупок для проверки протоколов — только для только что скачанных дней
    if use_lookup is not None and new_folders:
        notices: list[str] = []
        recorder = lambda n: notices.append(n) or []  # noqa: E731
        for region in regions:
            for archive in _cached_day(cfg, feed, region, day):
                try:
                    scan_archive(archive, feed.law, inns, index, recorder)
                except (OSError, zipfile.BadZipFile):
                    pass
        client.prefetch_reestr("RI223", sorted(set(notices)))
    archives: list[Path] = []
    cache = cfg.path("cache_dir") or Path("data/cache")
    failed = getattr(client, "failed_folders", set())
    for region in regions:
        if failed and any(cache / feed.subsystem / region / key / t in failed for t in feed.document_types):
            # 07.11: ЕИС не подготовила архивы за ~7 минут ожидания; повтор по одному региону
            # тратил ещё по 2 минуты на регион. Такой регион повторяется в конце загрузки.
            msg = f"{day} {feed.subsystem} регион {region}: архивы не готовы в ЕИС, повтор в конце загрузки"
            log.warning(msg)
            errors.append(msg)
            continue
        try:
            region_archives = client.fetch_day(feed, region, day)
        except (EisError, OSError) as exc:
            msg = f"{day} {feed.subsystem} регион {region}: {exc}"
            log.error(msg)
            errors.append(msg)
            continue
        if pool is not None and use_lookup is None:
            archives += region_archives
            continue
        for archive in region_archives:
            n_arch += 1
            try:
                contracts, wins = scan_archive(archive, feed.law, inns, index, use_lookup)
            except (OSError, zipfile.BadZipFile) as exc:
                msg = f"{day} {feed.subsystem} регион {region}: {exc}"
                log.error(msg)
                errors.append(msg)
                continue
            counters["found"] += store.add(_keep(cfg, contracts))
            counters["wins"] += wins_store.add(wins)
    if archives:
        # разбор архивов дня в нескольких процессах: на компьютере СРО это главное время
        n_arch += len(archives)
        for contracts, err in pool.map(_scan_job, [(str(a), feed.law, inns) for a in archives], chunksize=4):
            if err:
                msg = f"{day} {feed.subsystem}: {err}"
                log.error(msg)
                errors.append(msg)
            counters["found"] += store.add(_keep(cfg, contracts))
    log.info("%s %s: договоров членов +%s, побед +%s, проверено закупок +%s (всего договоров %s; архивов %s, разбор %.0f с)",
             feed.subsystem, f"{day:%d.%m.%Y}", counters["found"] - before[0], counters["wins"] - before[1],
             stats.get("lookups", 0) - before[2], counters["found"], n_arch, time.monotonic() - t0)
    day_ok = not any(m.startswith(key) for m in errors[n_err:])
    if marks is not None and day_ok:
        marks[key] = mark
        _save_scanned(cfg, feed, marks)
    if day_ok and mark is not None and not _keep_archives(cfg):
        # договоры дня уже записаны в data/work — архивы больше не нужны
        freed = 0
        for region in regions:
            for t in feed.document_types:
                folder = cache / feed.subsystem / region / key / t
                if (folder / ".done").exists():
                    freed += purge_folder(folder, mark)
        if freed:
            _set_purge_flag(cache)
            counters["purged"] = counters.get("purged", 0) + freed


def cmd_fetch(cfg: Config, start: date, end: date, newest_first: bool = True,
              laws: set[str] | None = None) -> tuple[int, list[str]]:
    """Договоры 223-ФЗ и протоколы закупок за период, от свежих дат к старым.

    Основной путь: договор без поставщика со строительными ОКПД2 → протоколы этой
    закупки по её номеру → победитель-член СРО. Поток протоколов по регионам
    (kind = "protocol"), если он задан, идёт первым.
    """
    members = load_members(cfg.path("members"))
    inns = set(members)
    client = EisClient(cfg)
    _apply_protocol_types(cfg)
    store, wins_store = _store(cfg), _wins(cfg)
    errors: list[str] = []
    counters = {"found": 0, "wins": 0}
    # протоколы, затем 223-ФЗ (главное для сверки), 44-ФЗ, 615-ПП
    feeds = _ordered(cfg.feeds)
    if laws:
        feeds = [f for f in feeds if f.law in laws]
        log.info("Загрузка только по закону: %s", ", ".join(sorted(laws)))
    days = list(_days(start, end))
    if newest_first:
        days.reverse()
    index = None
    win_regions: list[str] = []
    stats: dict = {}
    lookup = _make_lookup(client, inns, stats, errors)
    prefetch = bool(cfg.eis.get("prefetch", True))
    for feed in feeds:
        if feed.kind != "protocol" and index is None:
            all_wins = wins_store.load()
            index = WinIndex(all_wins)
            # регион заказчика ≈ первые две цифры его ИНН
            win_regions = sorted({w.customer_inn[:2] for w in all_wins if len(w.customer_inn) >= 10})
        regions = cfg.feed_regions(feed)
        if feed.regions and any(r.lower() == "wins" for r in feed.regions):
            regions = sorted(set(cfg.regions) | set(win_regions))
            log.info("%s: регионы заказчиков, где члены побеждали: %s", feed.subsystem, ", ".join(regions))
        use_lookup = lookup if (feed.lookup_protocols and feed.law == "223") else None
        # дни идут пачками (eis.days_batch, по умолчанию 7): архивы всей пачки заказываются
        # сразу, ЕИС готовит их параллельно — ожидание не повторяется для каждого дня
        batch = max(1, int(cfg.eis.get("days_batch", 7))) if prefetch else 1
        chunks = [days[i:i + batch] for i in range(0, len(days), batch)]
        marks = _load_scanned(cfg, feed) if use_lookup is None else None
        mark = _scan_mark(cfg, feed, inns)
        workers = _scan_workers(cfg) if use_lookup is None else 1
        pool = ProcessPoolExecutor(max_workers=workers) if workers > 1 else None
        skipped0 = counters.get("skipped", 0)

        def order(chunk):
            # уже разобранные дни в ЕИС не заказываем
            todo = [d for d in chunk if marks is None or marks.get(d.isoformat()) != mark]
            if todo:
                client.unpurge_days(feed, regions, todo, mark)
            return client.prefetch_days(feed, regions, todo, errors) if todo else {}
        try:
            # следующая пачка скачивается, пока разбирается текущая
            with ThreadPoolExecutor(max_workers=1) as bg:
                fut = bg.submit(order, chunks[0]) if prefetch and chunks else None
                for i, chunk in enumerate(chunks):
                    new_days = fut.result() if fut else {}
                    fut = bg.submit(order, chunks[i + 1]) if prefetch and i + 1 < len(chunks) else None
                    for day in chunk:
                        _fetch_one_day(cfg, client, feed, regions, day, new_days.get(day, 0), use_lookup,
                                       inns, index, store, wins_store, errors, stats, counters, pool, marks, mark)
        finally:
            if pool is not None:
                pool.shutdown(cancel_futures=True)
        if counters.get("skipped", 0) > skipped0:
            log.info("%s: дней, разобранных раньше, пропущено: %s", feed.subsystem, counters["skipped"] - skipped0)
    found, wins_found = counters["found"], counters["wins"]
    log.info("Новых договоров членов: %s, побед по протоколам: %s, ошибок: %s", found, wins_found, len(errors))
    if counters.get("purged"):
        log.info("Архивы разобранных дней удалены, освобождено %.1f ГБ", counters["purged"] / 1024 ** 3)
    store.compact()   # дашборды читают файл напрямую: в нём должна быть одна запись на договор
    return found, errors


def cmd_fetch_retry(cfg: Config, start: date, end: date, laws: set[str] | None = None):
    """Загрузка с автоматической догрузкой дней, на которых ЕИС ответила ошибкой (один повтор)."""
    found, errors = cmd_fetch(cfg, start, end, laws=laws)
    days = sorted({e[:10] for e in errors if len(e) >= 10 and e[4] == "-" and e[7] == "-"})
    if not days:
        return found, errors
    log.info("Догружаю дни с ошибками ЕИС: %s", ", ".join(days))
    left = [e for e in errors if not (len(e) >= 10 and e[:10] in days)]
    for d in days:
        day = date.fromisoformat(d)
        f2, e2 = cmd_fetch(cfg, day, day, laws=laws)
        found += f2
        left += e2
    log.info("После догрузки ошибок осталось: %s", len(left))
    return found, left


def cmd_etalon(cfg: Config) -> int:
    """Проверка правил на решениях оператора: где робот сейчас думает иначе, чем решил человек."""
    from openpyxl import Workbook
    decisions = load_decisions(_decisions_path(cfg))
    decided = {k: d for k, d in decisions.items() if d.get("Решение")}
    if not decided:
        log.info("В файле решений пока нет ни одного решения — эталон пуст.")
        return 0
    members = load_members(cfg.path("members"))
    contracts = _keep(cfg, _store(cfg).load())
    contracts.extend(unlinked_wins(_wins(cfg).load(), contracts))
    m = cfg.matching
    records = reconcile(members, contracts, cfg.competitive_from, cfg.noncompetitive_from,
                        int(m["date_tolerance_days"]), float(m["price_tolerance_pct"]))
    from .decisions import kind_of, record_key
    by_key = {record_key(r.primary): r for r in records}
    rows, agree, changed, cautious = [], 0, 0, 0
    for k, d in decided.items():
        r = by_key.get(k)
        robot = r.status if r else "нет в сверке"
        human = kind_of(str(d.get("Решение")))
        # ошибка робота — только уверенное «не то»: подтвердил то, что человек исключил,
        # или сам отбросил то, что человек передал. «Требует проверки» — осторожность, не ошибка.
        wrong = (human == "exclude" and robot == CONFIRMED) or \
                (human == "send" and robot not in (CONFIRMED, NEEDS_CHECK, PROBABLE, MISMATCH, DECLARED))
        ok = not wrong
        cautious += robot in (NEEDS_CHECK, PROBABLE, MISMATCH, DECLARED)
        agree += ok
        was = str(d.get("Статус робота при добавлении") or "")
        moved = bool(was) and was != robot
        changed += moved
        if not ok or moved:
            rows.append([k, d.get("Член СРО"), d.get("Решение"), was, robot,
                         "робот ошибается" if not ok else "правило изменило статус", "; ".join(r.issues) if r else ""])
    share = agree / len(decided)
    out = cfg.path("output_dir") or Path("data/out")
    wb = Workbook()
    ws = wb.active
    ws.title = "Эталон"
    ws.append([f"Проверка правил на {len(decided)} решениях оператора: робот не ошибся в {share:.0%}, "
               f"оставил человеку {cautious}, статус изменился у {changed}"])
    ws.append(["Ключ", "Член СРО", "Решение оператора", "Робот при добавлении", "Робот сейчас", "Вывод", "Замечания робота"])
    for row in rows:
        ws.append(row)
    for col, w in zip("ABCDEFG", (24, 34, 26, 18, 18, 22, 60)):
        ws.column_dimensions[col].width = w
    path = out / f"etalon_{date.today():%Y%m%d}.xlsx"
    out.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    (out / "etalon_last.json").write_text(json.dumps({"date": date.today().isoformat(), "decided": len(decided),
                                                      "agree": agree, "share": round(share, 4), "changed": changed,
                                                      "cautious": cautious},
                                                     ensure_ascii=False), encoding="utf-8")
    log.info("Эталон: решений %s, робот не ошибся в %s (%.0f%%), оставил человеку %s, статус изменился у %s. Файл %s",
             len(decided), agree, share * 100, cautious, changed, path.name)
    return 0 if share >= 0.95 else 1


def cmd_import_archives(cfg: Config, law: str, folder: Path) -> int:
    """Разбор архивов, скачанных заранее (например, история за прошлые годы).

    Протоколы и договоры можно класть в одну папку: модуль проходит её дважды —
    сначала собирает победы, потом связывает с ними договоры.
    """
    inns = set(load_members(cfg.path("members")))
    store, wins_store = _store(cfg), _wins(cfg)
    zips = sorted(Path(folder).rglob("*.zip"))
    wins_total = 0
    if law == "223":
        for z in zips:
            wins_total += wins_store.add(scan_archive(z, law, inns)[1])
    index = WinIndex(wins_store.load()) if law == "223" else None
    total = 0
    for z in zips:
        total += store.add(scan_archive(z, law, inns, index)[0])
    log.info("Из архивов добавлено договоров: %s, побед по протоколам: %s", total, wins_total)
    return total


def cmd_probe(cfg: Config, protocol_day: date, contract_day: date, region: str,
              expect_reg: str = "") -> bool:
    """Короткая проверка токена и настроек на одном дне и одном регионе.

    По умолчанию — закупка № 32414075550 ООО «ЛУКОЙЛ-ВОЛГОГРАДЭНЕРГО» (регион 34):
    итоговый протокол опубликован 28.10.2024, редакция договора — 24.12.2024,
    победитель — ООО «ВОЛГОГРАДМОНТАЖ».
    """
    members = load_members(cfg.path("members"))
    inns = set(members)
    client = EisClient(cfg)
    store, wins_store = _store(cfg), _wins(cfg)
    ok = True
    print(f"Членов СРО в списке: {len(inns)}")
    print(f"Сервис ЕИС: {cfg.endpoint('223')}")
    try:
        types = _load_nsi_protocol_types(cfg, client)
        print(f"Справочник типов документов 223-ФЗ получен, типов протоколов: {len(types)}")
        print("   " + ", ".join(types[:40]) + (" ..." if len(types) > 40 else ""))
    except (EisError, OSError, zipfile.BadZipFile) as exc:
        print(f"ОШИБКА справочника nsiDocumentTypes223 (токен или доступ): {exc}")
        ok = False
    _apply_protocol_types(cfg)
    for feed in sorted(cfg.feeds, key=lambda f: 0 if f.kind == "protocol" else 1):
        if feed.law != "223":
            continue
        day = protocol_day if feed.kind == "protocol" else contract_day
        index = None if feed.kind == "protocol" else WinIndex(wins_store.load())
        stats: dict = {}
        errs: list[str] = []
        lookup = _make_lookup(client, inns, stats, errs) if feed.lookup_protocols else None
        try:
            archives = client.fetch_day(feed, region, day)
        except (EisError, OSError) as exc:
            print(f"ОШИБКА {feed.subsystem} {feed.document_types} {day} регион {region}: {exc}")
            ok = False
            continue
        contracts, wins = [], []
        for a in archives:
            c, w = scan_archive(a, feed.law, inns, index, lookup)
            contracts += c
            wins += w
        if lookup is not None:
            print(f"   строительных закупок проверено по протоколам: {stats.get('lookups', 0)}")
            for e in errs[:5]:
                print(f"   ОШИБКА {e}")
            ok = ok and not errs
        store.add(contracts)
        wins_store.add(wins)
        print(f"{feed.subsystem} {', '.join(feed.document_types)} за {day:%d.%m.%Y}, регион {region}: "
              f"архивов {len(archives)}, побед членов {len(wins)}, договоров членов {len(contracts)}")
        for w in wins:
            print(f"   победа: {w.supplier_name} ИНН {w.supplier_inn}, протокол {w.protocol_number}, {w.price}")
        for c in contracts:
            print(f"   договор: {c.reg_number} ИНН {c.supplier_inn}, {c.price}, {c.sign_date}")
        if expect_reg and feed.kind != "protocol":
            hit = any(c.reg_number == expect_reg for c in contracts)
            print(("НАЙДЕН" if hit else "НЕ НАЙДЕН") + f" договор {expect_reg}")
            ok = ok and hit
    return ok


def cmd_rebuild(cfg: Config) -> int:
    """Пересобрать найденные договоры из уже скачанных архивов (без запросов в ЕИС).

    Нужен после изменения правил разбора. Безопасно: новый список пишется во временный файл
    и подменяет рабочий только в самом конце. Если прервать — рабочий файл не тронут,
    а повторный запуск продолжит с места остановки.
    """
    import os
    import shutil
    import time
    from datetime import datetime

    members = load_members(cfg.path("members"))
    inns = set(members)
    cache = cfg.path("cache_dir") or Path("data/cache")
    store_path = _store(cfg).path
    tmp_path = store_path.with_name(store_path.stem + ".rebuild.jsonl")
    prog_path = store_path.with_name("rebuild_progress.json")
    feeds = [f for f in _ordered(cfg.feeds) if f.kind != "protocol"]

    progress = None
    if tmp_path.exists() and prog_path.exists():
        try:
            progress = json.loads(prog_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            progress = None
    if progress:
        log.info("Продолжаю прерванную пересборку от %s", progress.get("started"))
    else:
        tmp_path.unlink(missing_ok=True)
        progress = {"started": datetime.now().strftime("%d.%m.%Y %H:%M"), "feeds": {}}
        for feed in feeds:
            archives = [z for z in sorted((cache / feed.subsystem).rglob("*.zip"))
                        if z.parent.name in feed.document_types]
            archives += sorted((cache / "byreestr" / feed.subsystem).rglob("*.zip"))   # точечная загрузка 44-ФЗ
            progress["feeds"][feed.subsystem] = {"archives": [str(z.relative_to(cache)) for z in archives], "done": 0}

    def save_progress():
        t = prog_path.with_suffix(".tmp")
        t.write_text(json.dumps(progress, ensure_ascii=False), encoding="utf-8")
        os.replace(t, prog_path)

    save_progress()
    store = ContractStore(tmp_path)
    index = WinIndex(_wins(cfg).load())

    def offline_lookup(notice: str) -> list:
        folder = cache / "byreestr" / "RI223" / notice
        if not (folder / ".done").exists():
            return []
        wins = []
        for archive in sorted(folder.glob("*.zip")):
            with zipfile.ZipFile(archive) as zf:
                for name in zf.namelist():
                    if name.lower().endswith(".xml"):
                        wins += parse_protocol_xml(zf.read(name), inns, f"{archive.name}:{name}")
        return wins

    total_all = sum(len(p["archives"]) for p in progress["feeds"].values())
    if not total_all:
        prog_path.unlink(missing_ok=True)
        tmp_path.unlink(missing_ok=True)
        log.error("В кэше %s нет скачанных архивов — пересобирать не из чего, рабочий файл не тронут.", cache)
        return 0
    done_all = sum(p["done"] for p in progress["feeds"].values())
    t0, d0 = time.monotonic(), done_all
    wins_store = _wins(cfg)
    for feed in feeds:
        p = progress["feeds"].get(feed.subsystem) or {"archives": [], "done": 0}
        archives = p["archives"]
        lookup = offline_lookup if feed.lookup_protocols else None
        log.info("%s: архивов %s, уже разобрано %s", feed.subsystem, len(archives), p["done"])
        for n in range(p["done"], len(archives)):
            z = cache / archives[n]
            if z.exists():
                try:
                    contracts, wins = scan_archive(z, feed.law, inns, index, lookup)
                except (OSError, zipfile.BadZipFile) as exc:
                    log.warning("архив %s пропущен: %s", z.name, exc)
                    contracts, wins = [], []
                kept = _keep(cfg, contracts)
                if kept:
                    store.add(kept)
                if wins:
                    wins_store.add(wins)
            p["done"] = n + 1
            done_all += 1
            if p["done"] % 200 == 0 or p["done"] == len(archives):
                save_progress()
                rate = (done_all - d0) / max(time.monotonic() - t0, 1)
                left = (total_all - done_all) / rate / 60 if rate else 0
                log.info("Пересборка: разобрано %s из %s (%s), осталось ≈ %s мин",
                         done_all, total_all, feed.subsystem, int(left))
    if (cache / PURGE_FLAG).exists() and store_path.exists():
        # архивы части дней удалены после разбора: их договоры берём из прежнего рабочего файла,
        # иначе пересборка потеряла бы всё, что найдено в этих днях
        rebuilt = {ContractStore._key(c) for c in store.load()}
        old = [c for c in ContractStore(store_path).load() if ContractStore._key(c) not in rebuilt]
        carried = store.add(_keep(cfg, old)) if old else 0
        log.warning("Архивы части дней удалены после разбора (eis.keep_archives = false). Договоры этих дней "
                    "(%s) перенесены из прежнего рабочего файла и проверены по текущему отбору по ГрК; "
                    "изменения в разборе архивов на них не действуют.", carried)
    total = store.compact()   # от каждого договора остаётся последняя версия
    old_total = 0
    if store_path.exists():
        with open(store_path, encoding="utf-8") as fh:
            old_total = sum(1 for line in fh if line.strip())
    tmp_path.touch()
    if old_total > 20 and total < old_total * 0.5:
        # вдвое меньше прежнего — скорее не тот кэш или сбой диска, а не новые правила
        keep = store_path.with_name(f"{store_path.stem}_rebuild_{datetime.now():%Y%m%d_%H%M%S}.jsonl")
        os.replace(tmp_path, keep)
        prog_path.unlink(missing_ok=True)
        log.error("Пересборка дала %s договоров вместо прежних %s — рабочий файл НЕ заменён. "
                  "Результат сохранён в %s; пришлите лог разработчику.", total, old_total, keep.name)
        return old_total
    if store_path.exists():
        shutil.copy2(store_path, store_path.with_name(f"{store_path.stem}_{datetime.now():%Y%m%d_%H%M%S}.bak.jsonl"))
    os.replace(tmp_path, store_path)
    prog_path.unlink(missing_ok=True)
    log.info("Пересобрано договоров членов: %s. Прежний файл сохранён рядом с пометкой .bak", total)
    return total


def _set_purge_flag(cache: Path) -> None:
    try:
        flag = cache / PURGE_FLAG
        if not flag.exists():
            flag.write_text("Архивы части дней удалены после разбора: см. .purged в папках дней\n", encoding="utf-8")
    except OSError:
        pass


_SCAN_LINE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})[,.]\d+ INFO (\S+) (\d{2})\.(\d{2})\.(\d{4}): договоров членов")


def _scanned_from_logs(cfg: Config) -> dict[tuple[str, str], datetime]:
    """(подсистема, день) → когда день последний раз разобран, по строкам журналов data/logs."""
    out: dict[tuple[str, str], datetime] = {}
    for lf in sorted((_data_dir(cfg) / "logs").glob("*.log")):
        try:
            with open(lf, encoding="utf-8", errors="replace") as fh:
                for line in fh:
                    m = _SCAN_LINE.match(line)
                    if not m:
                        continue
                    at = datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S")
                    k = (m.group(2), f"{m.group(5)}-{m.group(4)}-{m.group(3)}")
                    if k not in out or at > out[k]:
                        out[k] = at
        except OSError:
            continue
    return out


def cmd_purge_cache(cfg: Config, dry_run: bool = False) -> int:
    """Удалить архивы уже разобранных дней, оставив отметки .done (без запросов в ЕИС).

    День считается разобранным, если он отмечен в scanned.json по текущим правилам и списку
    членов или если в журнале есть строка его разбора, записанная позже скачивания папки.
    Дни без такого подтверждения не трогаются.
    """
    inns = set(load_members(cfg.path("members")))
    cache = cfg.path("cache_dir") or Path("data/cache")
    from_logs = _scanned_from_logs(cfg)
    total_freed, total_folders, total_kept = 0, 0, 0
    for feed in _ordered(cfg.feeds):
        if feed.kind == "protocol":
            continue
        base = cache / feed.subsystem
        if not base.is_dir():
            continue
        mark = _scan_mark(cfg, feed, inns)
        marked = {d for d, m in _load_scanned(cfg, feed).items() if m == mark}
        freed, folders, kept = 0, 0, 0
        for region_dir in sorted(base.iterdir()):
            if not region_dir.is_dir():
                continue
            for day_dir in sorted(region_dir.iterdir()):
                day = day_dir.name
                if not day_dir.is_dir() or len(day) != 10:
                    continue
                logged = from_logs.get((feed.subsystem, day))
                for t in feed.document_types:
                    folder = day_dir / t
                    done = folder / ".done"
                    if not done.exists() or (folder / PURGED).exists():
                        continue
                    zips = list(folder.glob("*.zip"))
                    if not zips:
                        continue
                    ok = day in marked
                    if not ok and logged is not None:
                        # строка разбора записана после скачивания этой папки
                        ok = datetime.fromtimestamp(done.stat().st_mtime) <= logged
                    if not ok:
                        kept += 1
                        continue
                    folders += 1
                    if dry_run:
                        freed += sum(z.stat().st_size for z in zips)
                    else:
                        freed += purge_folder(folder, mark)
                        _set_purge_flag(cache)
        log.info("%s: %s папок дней, %.1f ГБ%s; без подтверждения разбора оставлено папок: %s",
                 feed.subsystem, folders, freed / 1024 ** 3, " можно освободить" if dry_run else " освобождено", kept)
        total_freed += freed
        total_folders += folders
        total_kept += kept
    log.info("Итого: %s папок, %.1f ГБ%s. Оставлено без подтверждения разбора: %s папок.",
             total_folders, total_freed / 1024 ** 3, " можно освободить" if dry_run else " освобождено", total_kept)
    return 0


def cmd_reconcile(cfg: Config):
    members = load_members(cfg.path("members"))
    contracts = _keep(cfg, _store(cfg).load())
    pending = unlinked_wins(_wins(cfg).load(), contracts)
    if pending:
        log.info("Побед по протоколам без найденного договора: %s", len(pending))
        contracts.extend(pending)
    for key, source in (("finpod_import", "finpod"), ("member_notices", "notice"),
                        ("nostroy_registry", "nostroy")):
        p = cfg.path(key)
        if p and p.exists():
            items = load_table_source(p, source)
            log.info("%s: загружено %s записей", p.name, len(items))
            contracts.extend(items)
    m = cfg.matching
    records = reconcile(members, contracts, cfg.competitive_from, cfg.noncompetitive_from,
                        int(m["date_tolerance_days"]), float(m["price_tolerance_pct"]))
    # память робота: решения оператора и уже переданное в НОСТРОЙ
    decisions = load_decisions(_decisions_path(cfg))
    dstat = apply_decisions(records, decisions)
    jstat = apply_journal(records, load_journal(_journal_path(cfg)), float(m["price_tolerance_pct"]))
    added = sync_queue(_decisions_path(cfg), records, decisions)
    log.info("Решения оператора: передать %s, исключить %s, ждать %s; новых спорных в файле решений: %s. "
             "Уже в НОСТРОЙ: %s, изменились после передачи: %s",
             dstat["send"], dstat["exclude"], dstat["wait"], added, jstat["sent"], jstat["changed"])
    out = cfg.path("output_dir") or Path("data/out")
    from .excluded import for_report
    shown = {r.primary.reg_number for r in records if getattr(r, "primary", None) and r.primary.reg_number}
    excluded = for_report(_excluded(cfg), members, cfg.competitive_from, shown)
    log.info("Исключены роботом (лист в файле сверки): %s", len(excluded))
    book = write_operator_workbook(records, out / f"sverka_{date.today():%Y%m%d}.xlsx",
                                   cfg.raw.get("sro", {}).get("name", ""), excluded=excluded, members=members)
    sent = NostroyTransport(cfg.nostroy, out).send(records)
    counts = {s: sum(r.status == s for r in records)
              for s in (CONFIRMED, NEEDS_CHECK, PROBABLE, MISMATCH, DECLARED, IN_NOSTROY, EXCLUDED)}
    summary = (f"Членов: {len(members)}. Записей после сверки: {len(records)}.\n"
               f"К передаче: {counts[CONFIRMED]}. Требуют решения: "
               f"{counts[NEEDS_CHECK] + counts[PROBABLE] + counts[MISMATCH] + counts[DECLARED]} "
               f"(требуют проверки {counts[NEEDS_CHECK]}, вероятные {counts[PROBABLE]}, расхождения {counts[MISMATCH]}, заявлены членами {counts[DECLARED]}).\n"
               f"Уже в НОСТРОЙ: {counts[IN_NOSTROY]}. Исключены оператором: {counts[EXCLUDED]}.\n"
               f"Файл для оператора: {book.name}. {sent}")
    log.info(summary)
    return records, summary


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="eis2nostroy", description="Сведения о договорах членов СРО: ЕИС → НОСТРОЙ")
    ap.add_argument("-c", "--config", default="config.toml")
    ap.add_argument("-v", "--verbose", action="store_true")
    sub = ap.add_subparsers(dest="cmd", required=True)
    f = sub.add_parser("fetch", help="скачать договоры из ЕИС за период")
    f.add_argument("--from", dest="start", required=True, type=date.fromisoformat)
    f.add_argument("--to", dest="end", default=date.today(), type=date.fromisoformat)
    f.add_argument("--law", action="append", choices=["44", "223", "615"],
                   help="только этот закон (можно повторять)")
    i = sub.add_parser("import-archives", help="разобрать заранее скачанные архивы ЕИС")
    i.add_argument("--law", required=True, choices=["44", "223", "615"])
    i.add_argument("folder", type=Path)
    sub.add_parser("reconcile", help="сверить источники и подготовить выгрузки")
    ms = sub.add_parser("mark-sent", help="отметить пакет как загруженный в НОСТРОЙ")
    ms.add_argument("package", nargs="?", type=Path, help="CSV-пакет; по умолчанию последний nostroy_*.csv")
    sub.add_parser("etalon", help="проверить правила на решениях оператора")
    tg = sub.add_parser("targeted223", help="точечная загрузка 223-ФЗ по списку реестровых номеров")
    tg.add_argument("file", type=Path, help="Excel с колонками «Реестровый номер» и «Дата заключения»")
    tg.add_argument("--window", type=int, default=10, help="дней после даты заключения (по умолчанию 10)")
    t4 = sub.add_parser("targeted44", help="точечная загрузка 44-ФЗ по списку реестровых номеров")
    t4.add_argument("file", type=Path, help="Excel с колонками «Реестровый номер» и «Дата заключения»")
    t4.add_argument("--window", type=int, default=10, help="дней после даты заключения при поиске по дням")
    t4.add_argument("--law", choices=("44", "223"), default="44", help="223 — договоры 223-ФЗ по номеру, без поиска по дням")
    sub.add_parser("subjects223", help="дописать предмет договорам 223-ФЗ из извещений ЕИС")
    au = sub.add_parser("audit223", help="отчёт по каждому договору списка, без запросов в ЕИС")
    au.add_argument("file", type=Path, help="Excel со списком договоров")
    au.add_argument("--window", type=int, default=10, help="дней после даты заключения (по умолчанию 10)")
    ro = sub.add_parser("roles44", help="роль члена СРО в договорах 44-ФЗ, выпавших при пересборке, без запросов в ЕИС")
    ro.add_argument("--backup", type=Path, help="резервная копия рабочего файла; по умолчанию самая свежая подходящая")
    sub.add_parser("versions", help="найти договоры, чья новая версия не прошла отбор, без запросов в ЕИС")
    sub.add_parser("rebuild", help="пересобрать договоры из скачанных архивов без запросов в ЕИС")
    pc = sub.add_parser("purge-cache", help="удалить архивы уже разобранных дней, оставив отметки (без запросов в ЕИС)")
    pc.add_argument("--dry-run", action="store_true", help="только посчитать, сколько места освободится")
    pr = sub.add_parser("probe", help="проверить токен и настройки 223-ФЗ на одной закупке")
    pr.add_argument("--protocol-day", default="2024-10-28", type=date.fromisoformat)
    pr.add_argument("--contract-day", default="2024-12-24", type=date.fromisoformat)
    pr.add_argument("--region", default="34")
    pr.add_argument("--expect", default="53435098928240001840000")
    sub.add_parser("probe615", help="найти типы документов ПП-615, которые отдаёт ЕИС")
    r = sub.add_parser("run", help="ежедневный прогон: скачать последние дни и сверить")
    r.add_argument("--days", type=int, default=3)
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s")
    cfg = load_config(args.config)
    log.info(version_line())
    try:
        if args.cmd in ("fetch", "rebuild", "run", "targeted223", "targeted44", "subjects223", "purge-cache"):
            try:
                with run_lock(_lock_path(cfg), args.cmd), keep_awake():
                    return _main_locked(cfg, args)
            except RobotBusy as exc:
                log.error("%s", exc)
                print("\n" + str(exc))
                return 3
        if args.cmd == "mark-sent":
            out = cfg.path("output_dir") or Path("data/out")
            pkg = args.package or max(out.glob("nostroy_*.csv"), key=lambda p: p.stat().st_mtime, default=None)
            if not pkg:
                print("Пакетов nostroy_*.csv в папке data\\out нет.")
                return 1
            n = mark_sent(pkg, _journal_path(cfg))
            log.info("Пакет %s отмечен как загруженный в НОСТРОЙ: новых договоров в журнале %s", pkg.name, n)
            cmd_reconcile(cfg)
            return 0
        if args.cmd == "etalon":
            return cmd_etalon(cfg)
        return _main_locked(cfg, args)
    except Exception:  # noqa: BLE001 — любой сбой должен дойти до оператора
        notify(cfg.alerts, "Сверка договоров: сбой", traceback.format_exc()[-3000:])
        raise


def _main_locked(cfg: Config, args) -> int:
    if True:
        if args.cmd == "targeted223":
            from .targeted import cmd_targeted223
            _, errors = cmd_targeted223(cfg, args.file, args.window, _store(cfg), _wins(cfg), _keep, _make_lookup)
            return 1 if errors else 0
        if args.cmd == "targeted44":
            from .targeted44 import cmd_targeted44
            _, errors = cmd_targeted44(cfg, args.file, args.window, _store(cfg), _keep, law=args.law)
            return 1 if errors else 0
        if args.cmd == "subjects223":
            from .enrich import cmd_subjects223
            _, errors = cmd_subjects223(cfg, _store(cfg).path)
            return 1 if errors else 0
        if args.cmd == "audit223":
            from .enrich import cmd_audit223
            cmd_audit223(cfg, args.file, args.window, _store(cfg))
            return 0
        if args.cmd == "roles44":
            from .roles import cmd_roles44
            found = cmd_roles44(cfg, _store(cfg).path, _journal_path(cfg), lambda c: _keep(cfg, c), args.backup)
            return 0 if found else 1
        if args.cmd == "versions":
            from .versions import cmd_versions
            cmd_versions(cfg, _store(cfg))
            return 0
        if args.cmd == "fetch":
            _, errors = cmd_fetch_retry(cfg, args.start, args.end, laws=set(args.law or []))
            return 1 if errors else 0
        if args.cmd == "import-archives":
            cmd_import_archives(cfg, args.law, args.folder)
            return 0
        if args.cmd == "reconcile":
            cmd_reconcile(cfg)
            return 0
        if args.cmd == "purge-cache":
            return cmd_purge_cache(cfg, args.dry_run)
        if args.cmd == "rebuild":
            cmd_rebuild(cfg)
            cmd_reconcile(cfg)
            return 0
        if args.cmd == "probe615":
            from .probe615 import cmd_probe615
            return 0 if cmd_probe615(cfg) else 1
        if args.cmd == "probe":
            return 0 if cmd_probe(cfg, args.protocol_day, args.contract_day, args.region, args.expect) else 1
        if args.cmd == "run":
            end = date.today() - timedelta(days=1)
            _, errors = cmd_fetch_retry(cfg, end - timedelta(days=args.days - 1), end)
            _, summary = cmd_reconcile(cfg)
            if errors:
                summary += "\n\nОшибки загрузки из ЕИС:\n" + "\n".join(errors[:20])
            notify(cfg.alerts, "Сверка договоров: " + ("есть ошибки" if errors else "готово"), summary)
            return 1 if errors else 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
