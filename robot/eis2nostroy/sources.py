"""Дополнительные источники: файлы, которые оператор кладёт в папку.

* Выгрузка «Финподдержки» — записи основаны на протоколах закупок.
* Уведомления членов СРО о договорах (по регламенту, ст. 55.8 ГрК, приказ 655/пр).
* Уже переданные в НОСТРОЙ сведения (выгрузка из Единого реестра).

Столбцы ищутся по смыслу названия, точный формат не важен.
"""
from __future__ import annotations

import json
import logging
import os
import re
from pathlib import Path

from .models import Contract
from .tabular import pick, read_rows, to_date, to_float, to_inn

_VERSION = re.compile(r"_(\d+)_[0-9A-Fa-f]{32}\.xml$")


def _law(v) -> str:
    s = str(v or "")
    for law in ("223", "615", "44"):
        if law in s:
            return law
    return ""


def _competitive(v) -> bool | None:
    s = str(v or "").lower()
    if not s:
        return None
    if "неконкур" in s or "единств" in s:
        return False
    if "конкур" in s:
        return True
    return None


def load_table_source(path: Path, source: str) -> list[Contract]:
    out = []
    for row in read_rows(path):
        inn = to_inn(pick(row, "иннчлена", "иннпоставщика", "иннподрядчика", "иннисполнителя", "инн"))
        if len(inn) not in (10, 12):
            continue
        out.append(Contract(
            source=source,
            law=_law(pick(row, "закон", "фз", "основание")),
            supplier_inn=inn,
            reg_number=str(pick(row, "реестровый", "номерконтракта", "номердоговора", "реестр") or "").strip(),
            notice_number=str(pick(row, "извещен", "закупк") or "").strip(),
            customer_inn=to_inn(pick(row, "иннзаказчика")),
            customer_name=str(pick(row, "заказчик") or ""),
            sign_date=to_date(pick(row, "датадоговора", "датаконтракта", "датазаключ", "дата")),
            price=to_float(pick(row, "сумма", "цена", "стоимость")),
            subject=str(pick(row, "предмет", "объект") or "")[:500],
            competitive=_competitive(pick(row, "способ", "конкурент")),
            raw_file=path.name,
        ))
    return out


def file_version(raw_file: str) -> int:
    """Номер версии договора из имени XML: contract_<номер>_<версия>_<хеш>.xml; -1, если его нет."""
    m = _VERSION.search(raw_file or "")
    return int(m.group(1)) if m else -1


class ContractStore:
    """Найденные в ЕИС договоры членов: JSONL, без дублей, по одной записи на договор.

    У договора в ЕИС бывают десятки версий: допсоглашения меняют цену и предмет. В рабочем
    наборе остаётся последняя версия, в каком бы порядке ни шли архивы. Новая версия дописывается
    в конец файла, а при чтении и в compact() побеждает запись с большим номером версии.

    Ключи уже записанных договоров держатся в памяти: добавление не перечитывает весь файл.
    Если файл изменил другой процесс (размер или время изменения не совпадают с нашими),
    ключи перечитываются один раз.
    """

    def __init__(self, path: Path):
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)
        self._keys: dict | None = None
        self._stamp: tuple | None = None

    @staticmethod
    def _key(c: Contract) -> tuple:
        # дата может меняться от версии к версии, поэтому при реестровом номере в ключ не входит
        return (c.source, c.supplier_inn, c.reg_number or c.raw_file, None if c.reg_number else c.sign_date)

    def _file_stamp(self) -> tuple | None:
        try:
            st = self.path.stat()
        except FileNotFoundError:
            return None
        return (st.st_size, st.st_mtime_ns)

    def load(self) -> list[Contract]:
        if not self.path.exists():
            self._keys, self._stamp = {}, None
            return []
        stamp = self._file_stamp()
        best: dict[tuple, Contract] = {}
        with open(self.path, encoding="utf-8") as fh:
            for line in fh:
                if not line.strip():
                    continue
                try:
                    c = Contract.from_json(json.loads(line))
                except (ValueError, TypeError):
                    # недописанная строка: окно закрыли в момент записи — пропускаем её
                    logging.getLogger("eis2nostroy").warning("%s: пропущена повреждённая строка", self.path.name)
                    continue
                k = self._key(c)
                old = best.get(k)
                if old is None:
                    best[k] = c
                elif file_version(c.raw_file) > file_version(old.raw_file):
                    if not c.subject and old.subject:
                        # предмет 223-ФЗ дописан из извещения: в новой версии его снова нет
                        c.subject = old.subject
                        if "subject_source" in old.extra:
                            c.extra["subject_source"] = old.extra["subject_source"]
                    best[k] = c
        self._keys, self._stamp = {k: file_version(c.raw_file) for k, c in best.items()}, stamp
        return list(best.values())

    def _known(self) -> dict:
        if self._keys is None or self._file_stamp() != self._stamp:
            self.load()
        return self._keys

    def add(self, contracts: list[Contract]) -> int:
        keys = self._known()
        new = []
        for c in contracts:
            k, version = self._key(c), file_version(c.raw_file)
            if k not in keys or version > keys[k]:
                keys[k] = version
                new.append(c)
        if new:
            with open(self.path, "a", encoding="utf-8") as fh:
                if _ends_without_newline(self.path):
                    fh.write("\n")   # после оборванной строки — с новой строки, чтобы не склеить записи
                for c in new:
                    fh.write(json.dumps(c.to_json(), ensure_ascii=False) + "\n")
        self._stamp = self._file_stamp()
        return len(new)

    def compact(self) -> int:
        """Оставить в файле только последние версии. Возвращает число договоров."""
        rows = self.load()
        if not self.path.exists():
            return 0
        tmp = self.path.with_suffix(".compact.tmp")
        with open(tmp, "w", encoding="utf-8") as fh:
            for c in rows:
                fh.write(json.dumps(c.to_json(), ensure_ascii=False) + "\n")
        os.replace(tmp, self.path)
        self._stamp = self._file_stamp()
        return len(rows)


def _ends_without_newline(path: Path) -> bool:
    try:
        with open(path, "rb") as fh:
            fh.seek(-1, 2)
            return fh.read(1) != b"\n"
    except OSError:
        return False
