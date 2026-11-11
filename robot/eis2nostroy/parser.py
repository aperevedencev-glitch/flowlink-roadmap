"""Разбор архивов ЕИС: из XML договоров берём те, где поставщик — член СРО.

Схемы XML у 44-ФЗ, 223-ФЗ и 615-ПП разные и меняются с версиями ЕИС,
поэтому разбор устойчивый: ищем поля по локальным именам элементов, а не
по жёстким путям. ИНН поставщика ищется внутри блоков, в названии которых
есть supplier / participant / contractor / executor.
"""
from __future__ import annotations

import functools
import re
import io
import logging
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from .models import Contract, is_customer_name
from typing import Callable

from .grk import classify as grk_classify

from .protocols import (ProtocolWin, WinIndex, contract_candidate, contracts_from_223, is_protocol,
                        parse_protocol_xml)
from .tabular import to_date, to_float

log = logging.getLogger(__name__)

SUPPLIER_BLOCKS = ("supplier", "participant", "contractor", "executor", "winner")
CUSTOMER_BLOCKS = ("customer", "placer", "organization")
REG_NUM = ("regnum", "registrynum", "contractregnum", "contractregnumber", "registrationnumber", "regnumber")
NOTICE_NUM = ("notificationnumber", "purchasenumber", "purchasenoticenumber", "noticenumber")
SIGN_DATE = ("signdate", "contractdate", "conclusiondate", "concludecontractdate", "contractsigndate")
PRICE = ("price", "pricerur", "contractprice", "sum", "summ", "amount")
SUBJECT = ("contractsubject", "subject", "subjectcontract", "purchaseobjectinfo")
NONCOMPETITIVE_HINTS = ("singlesupplier", "единственн", "неконкурент", "noncompetitive")


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1].lower()


def _first_text(el: ET.Element, names: tuple[str, ...]) -> str:
    for e in el.iter():
        if _local(e.tag) in names and e.text and e.text.strip():
            return e.text.strip()
    return ""


# длина реестрового номера договора: 44-ФЗ — 19 цифр, 223-ФЗ — 23; у заказчика в 44-ФЗ код из 11 цифр
REG_LEN = {"44": 19, "223": 23, "615": 18}
OWNER_BLOCKS = ("customer", "placer", "organization", "supplier", "participant", "contractor", "executor")


def _reg_number(root: ET.Element, law: str) -> str:
    """Реестровый номер именно договора, а не код заказчика из вложенного блока."""
    # 223-ФЗ: contractRegNumber — номер договора; registrationNumber — номер карточки сведений,
    # он идёт в XML раньше и меняется с каждой версией (…0000 → …0001, …0003)
    if str(law).replace("-ФЗ", "").replace("fz", "") == "223":
        for e in root.iter():
            t = (e.text or "").strip()
            if _local(e.tag) == "contractregnumber" and len(t) == 23 and t.isdigit():
                return t
    parent = {c: p for p in root.iter() for c in p}
    cands = []
    for e in root.iter():
        if _local(e.tag) in REG_NUM and e.text and e.text.strip():
            t = e.text.strip()
            owner, x = False, parent.get(e)
            while x is not None:
                if any(k in _local(x.tag) for k in OWNER_BLOCKS):
                    owner = True
                    break
                x = parent.get(x)
            cands.append((t, owner))
    if not cands:
        return ""
    want = REG_LEN.get(str(law).replace("-ФЗ", "").replace("fz", ""))
    for pick in (lambda t, o: not o and want and len(t) == want and t.isdigit(),
                 lambda t, o: want and len(t) == want and t.isdigit(),
                 lambda t, o: not o):
        for t, o in cands:
            if pick(t, o):
                return t
    return max((t for t, _ in cands), key=len)


def _inns_in(el: ET.Element) -> list[str]:
    return [e.text.strip() for e in el.iter() if _local(e.tag) == "inn" and e.text and e.text.strip()]


def _party_inns(el: ET.Element) -> list[str]:
    """ИНН стороны договора без вложенных блоков субподрядчиков (subContractor… и т. п.)."""
    res = []

    def walk(x: ET.Element):
        for c in x:
            name = _local(c.tag)
            if name.startswith("sub") and "subject" not in name:
                continue
            if name == "inn" and c.text and c.text.strip():
                res.append(c.text.strip())
            walk(c)
    walk(el)
    return res


def _blocks(root: ET.Element, keys: tuple[str, ...]) -> list[ET.Element]:
    found = []
    for e in root.iter():
        name = _local(e.tag)
        if any(k in name for k in keys):
            # берём самые внешние блоки, без вложенных повторов
            if not any(e is f or _contains(f, e) for f in found):
                found.append(e)
    return found


def _contains(parent: ET.Element, child: ET.Element) -> bool:
    return any(x is child for x in parent.iter())


SKIP_CUSTOMER = ("single", "reason", "foundation")   # singleCustomer — основание закупки, а не заказчик


def _customer_name(root: ET.Element, blocks: list[ET.Element], customer_inn: str) -> str:
    """Наименование заказчика: из его блока, а не из основания закупки у единственного поставщика."""
    for b in blocks:
        if any(k in _local(b.tag) for k in SKIP_CUSTOMER):
            continue
        for names in (("fullname",), ("shortname",), ("name",)):
            t = _first_text(b, names)
            if is_customer_name(t):
                return t
    if customer_inn:
        # блок, где рядом лежат ИНН заказчика и его полное наименование
        for e in root.iter():
            if any(_local(c.tag) == "inn" and (c.text or "").strip() == customer_inn for c in e):
                for names in (("fullname",), ("shortname",)):
                    t = _first_text(e, names)
                    if is_customer_name(t):
                        return t
    return ""


def grk_info(root: ET.Element, subject: str = "") -> dict:
    """ИКЗ, коды ОКПД2 позиций и отнесение к ГрК (см. grk.py)."""
    ikz = ""
    for e in root.iter():
        if _local(e.tag) in ("ikz", "purchasecode", "identificationcode") and e.text:
            digits = "".join(ch for ch in e.text if ch.isdigit())
            if len(digits) == 36:
                ikz = digits
                break
    okpd = []
    for e in root.iter():
        if "okpd" in _local(e.tag):
            for c in e.iter():
                if _local(c.tag) in ("code", "okpdcode") and c.text and c.text.strip()[:2].isdigit():
                    okpd.append(c.text.strip())
    cls, note = grk_classify(ikz, okpd, subject=subject)
    return {"ikz": ikz, "okpd2": sorted(set(okpd))[:5], "grk": cls, "grk_note": note}


def parse_contract_xml(data: bytes, law: str, member_inns: set[str], raw_file: str = "") -> list[Contract]:
    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        log.warning("Не разобран XML %s: %s", raw_file, exc)
        return []
    if is_protocol(root):
        # протокол — не договор: победителей разбирает protocols.py
        return []
    # субподрядчики (subContractor…) — не стороны договора: член СРО-субподрядчик договор с заказчиком
    # не заключал и в реестр НОСТРОЙ по нему не отчитывается
    supplier_blocks = [b for b in _blocks(root, SUPPLIER_BLOCKS)
                       if not (_local(b.tag).startswith("sub") and "subject" not in _local(b.tag))]
    supplier_inns = {inn for b in supplier_blocks for inn in _party_inns(b)}
    hits = supplier_inns & member_inns
    if not hits:
        return []
    customer_blocks = _blocks(root, CUSTOMER_BLOCKS)
    customer_inn = next((i for b in customer_blocks for i in _inns_in(b) if i not in supplier_inns), "")
    customer_name = _customer_name(root, customer_blocks, customer_inn)
    method_text = " ".join(
        (e.text or "") + " " + _local(e.tag) for e in root.iter()
        if _local(e.tag) in ("placingway", "purchasemethod", "purchasecodename", "purchasemethodname",
                             "foundation", "singlecustomer", "singlesupplier", "competitive")
    ).lower()
    competitive = None
    if method_text.strip():
        competitive = not any(h in method_text for h in NONCOMPETITIVE_HINTS)
    price_el = None
    for e in root.iter():
        if _local(e.tag) in PRICE and e.text and to_float(e.text) is not None:
            price_el = e
            break
    subject = _first_text(root, SUBJECT)[:500]
    info = grk_info(root, subject)
    if str(law) == "615":
        # ПП-615 — капремонт общего имущества МКД: всегда ГрК, членство в СРО обязательно (ч. 7 ст. 166 ЖК)
        info.update(grk="ГрК", grk_note="ПП-615: капитальный ремонт общего имущества МКД (ч. 7 ст. 166 ЖК РФ)")
    result = []
    for inn in sorted(hits):
        result.append(Contract(
            source=f"eis_{law}",
            law=law,
            supplier_inn=inn,
            reg_number=_reg_number(root, law),
            notice_number=_first_text(root, NOTICE_NUM),
            customer_inn=customer_inn,
            customer_name=customer_name,
            sign_date=to_date(_first_text(root, SIGN_DATE)),
            price=to_float(price_el.text) if price_el is not None else None,
            subject=subject,
            competitive=competitive,
            raw_file=raw_file,
            extra=dict(info),
        ))
    return result


Lookup = Callable[[str], list[ProtocolWin]]   # номер закупки → победы членов в её протоколах


@functools.lru_cache(maxsize=8)
def _inn_pattern(inns: frozenset) -> "re.Pattern[bytes]":
    return re.compile(b"|".join(re.escape(i.encode()) for i in sorted(inns, key=len, reverse=True)))


def _mentions_member(data: bytes, member_inns: set[str]) -> bool:
    """Быстрая проверка: есть ли в XML хоть один ИНН члена СРО.

    07.11: полный разбор каждого XML занимал ~1,5 мин на день 44-ФЗ по всем регионам,
    хотя договоров членов среди них — единицы. Без ИНН члена договор члена быть не может."""
    if not member_inns:
        return False
    return _inn_pattern(frozenset(member_inns)).search(data) is not None


def scan_xml(data: bytes, law: str, member_inns: set[str], wins: WinIndex | None,
             raw_file: str = "", lookup: Lookup | None = None) -> tuple[list[Contract], list[ProtocolWin]]:
    """Один XML: договор члена СРО, победа в протоколе или ничего."""
    if law == "223" and b"urchaseProtocol" in data[:4000]:
        return [], parse_protocol_xml(data, member_inns, raw_file)
    found = parse_contract_xml(data, law, member_inns, raw_file) if _mentions_member(data, member_inns) else []
    if not found and law == "223" and wins:
        found = contracts_from_223(data, wins, raw_file)
    new_wins: list[ProtocolWin] = []
    if not found and law == "223" and lookup is not None:
        for cand in contract_candidate(data):
            got = lookup(cand["notice"])
            if got:
                new_wins += got
                found += contracts_from_223(data, WinIndex(got), raw_file)
    return found, new_wins


def scan_archive(path: Path, law: str, member_inns: set[str], wins: WinIndex | None = None,
                 lookup: Lookup | None = None) -> tuple[list[Contract], list[ProtocolWin]]:
    """Разбирает zip-архив ЕИС (в том числе вложенные zip)."""
    contracts: list[Contract] = []
    protocol_wins: list[ProtocolWin] = []

    def walk(zf: zipfile.ZipFile, prefix: str):
        for info in zf.infolist():
            name = info.filename
            if name.lower().endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(zf.read(info))) as inner:
                    walk(inner, f"{prefix}{name}/")
            elif name.lower().endswith(".xml"):
                c, w = scan_xml(zf.read(info), law, member_inns, wins, f"{prefix}{name}", lookup)
                contracts.extend(c)
                protocol_wins.extend(w)

    with zipfile.ZipFile(path) as zf:
        walk(zf, f"{path.name}:")
    return contracts, protocol_wins


def parse_archive(path: Path, law: str, member_inns: set[str], wins: WinIndex | None = None) -> list[Contract]:
    return scan_archive(path, law, member_inns, wins)[0]
