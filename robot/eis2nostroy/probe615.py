"""Диагностика ПП РФ 615: какие типы документов ЕИС отдаёт по реестру договоров 615.

Запуск: «7 Проверка ПП-615.bat» (нужен токен ЕИС). Результат — в окне и в логе
data\\logs\\probe615_ГГГГММДД.log; пришлите этот лог.

1. Из описания сервиса (WSDL и схемы XSD) выписывает все допустимые значения
   documentType44, где есть «615» или «PPRF» — это настоящие имена типов документов.
2. Для подсистем RD615 (реестр договоров 615) и PPRF615 (извещения и протоколы 615)
   спрашивает ЕИС за несколько рабочих дней по Москве и Московской области
   с каждым из этих типов и показывает, где пришли архивы.
3. Первый найденный архив скачивает и печатает имена XML внутри и корневой элемент.
"""
from __future__ import annotations

import logging
import re
import zipfile
from datetime import date
from pathlib import Path
from urllib.parse import urljoin
import xml.etree.ElementTree as ET

from .config import Config, Feed
from .eis_client import EisClient, EisError

log = logging.getLogger("eis2nostroy")

DAYS = [date(2025, 6, 10), date(2025, 9, 16), date(2025, 11, 20), date(2026, 3, 17)]
REGIONS = ["77", "50"]
FALLBACK_TYPES = ["contract", "contractProcedure", "contract615", "contractPPRF615", "pprf615Contract"]


def schema_doc_types(client: EisClient, cfg: Config) -> list[str]:
    """Значения documentType44 из WSDL/XSD сервиса, связанные с 615."""
    start = cfg.endpoint("44") + "?wsdl"
    seen, todo, values = set(), [start], set()
    while todo and len(seen) < 40:
        url = todo.pop()
        if url in seen:
            continue
        seen.add(url)
        try:
            r = client.http.get(url, timeout=60)
            r.raise_for_status()
        except Exception as exc:  # noqa: BLE001 — диагностика, печатаем и идём дальше
            print(f"  не получено {url}: {exc}")
            continue
        text = r.text
        for loc in re.findall(r'schemaLocation="([^"]+)"', text):
            todo.append(urljoin(url, loc))
        for v in re.findall(r'enumeration value="([^"]+)"', text):
            if "615" in v or "pprf" in v.lower():
                values.add(v)
    print(f"Схем прочитано: {len(seen)}; типов документов с 615/PPRF: {len(values)}")
    for v in sorted(values):
        print("   ", v)
    return sorted(values)


def cmd_probe615(cfg: Config) -> bool:
    client = EisClient(cfg)
    print("Шаг 1. Типы документов ПП-615 из описания сервиса ЕИС")
    types = schema_doc_types(client, cfg) or FALLBACK_TYPES
    if types is FALLBACK_TYPES:
        print("  В схемах типов с 615 не найдено, пробую типовые имена:", ", ".join(types))
    print("\nШаг 2. Запросы к ЕИС по подсистемам RD615 и PPRF615")
    found = []
    for subsystem in ("RD615", "PPRF615"):
        for doc_type in types:
            feed = Feed(law="615", endpoint_law="44", subsystem=subsystem, document_types=[doc_type])
            hits = 0
            for region in REGIONS:
                for day in DAYS:
                    try:
                        urls = client.request_archives(feed, region, day, doc_type)
                    except (EisError, OSError, ET.ParseError) as exc:
                        msg = str(exc)[:160]
                        print(f"  {subsystem} {doc_type} {region} {day}: ошибка — {msg}")
                        break
                    hits += len(urls)
                    if urls and len(found) < 3:
                        found.append((subsystem, doc_type, region, day, urls[0]))
            print(f"  {subsystem:8} {doc_type:40} архивов: {hits}")
    print("\nШаг 3. Что внутри архивов")
    cache = (cfg.path("cache_dir") or Path("data/cache")) / "probe615"
    for subsystem, doc_type, region, day, url in found:
        try:
            z = client.download_archive(url, cache / f"{subsystem}_{doc_type}_{region}_{day}.zip")
            with zipfile.ZipFile(z) as zf:
                names = [n for n in zf.namelist() if n.lower().endswith(".xml")]
                print(f"  {subsystem} {doc_type} {region} {day}: файлов {len(names)}, например {names[:3]}")
                if names:
                    root = ET.fromstring(zf.read(names[0]))
                    print("    корневой элемент:", root.tag, "→", [c.tag.rsplit('}', 1)[-1] for c in root][:5])
        except Exception as exc:  # noqa: BLE001
            print(f"  архив не разобран: {exc}")
    if not found:
        print("  Ни по одному типу архивов нет. Пришлите лог — разберём ответ ЕИС.")
    return bool(found)
