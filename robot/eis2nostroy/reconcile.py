"""Сверка источников и присвоение статусов."""
from __future__ import annotations

import re
from datetime import date

from .grk import GRK_PROBABLE, UNKNOWN
from .models import (CONFIRMED, DECLARED, IN_NOSTROY, MISMATCH, NEEDS_CHECK, OUT_OF_PERIOD, PROBABLE,
                     Contract, Member, ReconciledRecord, is_customer_name)

# Явно строительный предмет договора: строительство, реконструкция, капремонт, снос, СМР.
# Для 223-ФЗ нет КВР, поэтому по одному ОКПД2 робот не уверен; предмет снимает сомнение.
GRK_SUBJECT = re.compile(r"строительств|строительн\w*[- ]монтажн|реконструкц|капитальн\w*\s+ремонт|капремонт|"
                         r"\bснос|демонтаж\w*\s+(здани|сооружени|объект)|работы строительные|\bсмр\b", re.I)
NOT_BUILD_SUBJECT = re.compile(r"тепловой энерги|теплов\w+ энерги|электрическ\w+ энерги|^оказание услуг(?!.*(строительств|реконструкц|капитальн))|^поставк", re.I)

# Чем выше, тем надёжнее источник для «основной» записи
PRIORITY = {"eis_44": 5, "eis_223": 5, "eis_615": 5, "eis_223p": 4, "notice": 3, "finpod": 2, "nostroy": 1}


def _close(a: float | None, b: float | None, pct: float) -> bool:
    if a is None or b is None:
        return True
    if a == b:
        return True
    return abs(a - b) <= max(abs(a), abs(b)) * pct / 100


def _days(a: date | None, b: date | None) -> int | None:
    if a is None or b is None:
        return None
    return abs((a - b).days)


def same_contract(a: Contract, b: Contract, tol_days: int, tol_pct: float) -> bool:
    if a.supplier_inn != b.supplier_inn:
        return False
    if a.reg_number and b.reg_number:
        return a.reg_number == b.reg_number
    if a.notice_number and b.notice_number and a.notice_number == b.notice_number:
        return True
    if a.customer_inn and b.customer_inn and a.customer_inn != b.customer_inn:
        return False
    d = _days(a.sign_date, b.sign_date)
    if d is None or d > tol_days:
        return False
    # без номера считаем совпадением только при близкой цене
    return a.price is not None and b.price is not None and _close(a.price, b.price, tol_pct)


def cluster(contracts: list[Contract], tol_days: int, tol_pct: float) -> list[list[Contract]]:
    groups: list[list[Contract]] = []
    for c in contracts:
        for g in groups:
            if any(same_contract(c, x, tol_days, tol_pct) for x in g):
                g.append(c)
                break
        else:
            groups.append([c])
    return groups


def period_start(member: Member, competitive: bool | None, competitive_from: date,
                 noncompetitive_from: date) -> date:
    if competitive is False:
        return noncompetitive_from
    start = competitive_from
    # без даты уровня ОДО берём дату вступления в СРО
    anchor = member.odo_date or member.joined
    if anchor and anchor > start:
        start = anchor
    return start


def _family(source: str) -> str:
    if source == "eis_223p":
        return "protocol"       # победа по протоколу, договора в ЕИС нет
    return "eis" if source.startswith("eis_") else source


def _hold_doubtful(records: list[ReconciledRecord]) -> None:
    """Подтверждённые, но сомнительные договоры — оператору, а не сразу в пакет НОСТРОЙ.

    Цена 0 — обычно рамочное соглашение, а не договор подряда. Двойники — одна закупка, дата
    и цена у одного члена при разных реестровых номерах: договор могли зарегистрировать дважды.
    """
    groups: dict[tuple, list[ReconciledRecord]] = {}
    for r in records:
        if r.status != CONFIRMED:
            continue
        c = r.primary
        if c.price == 0:
            r.status = NEEDS_CHECK
            r.issues.append("цена договора 0 ₽: похоже на рамочное соглашение, а не на договор подряда")
        elif c.notice_number and c.sign_date and c.price is not None:
            groups.setdefault((r.member.inn, c.notice_number, c.sign_date, c.price), []).append(r)
    for same in groups.values():
        regs = sorted({r.primary.reg_number for r in same})
        if len(regs) < 2:
            continue
        for r in same:
            others = [x for x in regs if x != r.primary.reg_number]
            r.status = NEEDS_CHECK
            r.issues.append("возможный двойник: те же закупка, дата и цена у договора " + ", ".join(others))


def reconcile(members: dict[str, Member], contracts: list[Contract], competitive_from: date,
              noncompetitive_from: date, tol_days: int = 3, tol_pct: float = 1.0) -> list[ReconciledRecord]:
    # Наименование заказчика: в части записей вместо него попало основание закупки
    # («Пункт 25 части 1 статьи 93…»). Берём верное наименование того же ИНН из других записей.
    names: dict[str, str] = {}
    for c in contracts:
        if c.customer_inn and is_customer_name(c.customer_name):
            names.setdefault(c.customer_inn, c.customer_name)
    for c in contracts:
        if not is_customer_name(c.customer_name):
            c.customer_name = names.get(c.customer_inn, "")

    by_member: dict[str, list[Contract]] = {}
    for c in contracts:
        if c.supplier_inn in members:
            by_member.setdefault(c.supplier_inn, []).append(c)

    records: list[ReconciledRecord] = []
    for inn, items in by_member.items():
        member = members[inn]
        for group in cluster(items, tol_days, tol_pct):
            group.sort(key=lambda c: -PRIORITY.get(c.source, 0))
            primary = group[0]
            families = {_family(c.source) for c in group}
            issues = []
            prices = [c.price for c in group if c.price is not None]
            if prices and not all(_close(prices[0], p, tol_pct) for p in prices):
                issues.append("суммы в источниках различаются: " +
                              ", ".join(f"{c.source} {c.price:,.2f}".replace(",", " ") for c in group if c.price is not None))
            dates = [c.sign_date for c in group if c.sign_date]
            if dates and max(dates) != min(dates) and (max(dates) - min(dates)).days > tol_days:
                issues.append("даты в источниках различаются: " +
                              ", ".join(f"{c.source} {c.sign_date:%d.%m.%Y}" for c in group if c.sign_date))
            competitive = next((c.competitive for c in group if c.competitive is not None), None)
            start = period_start(member, competitive, competitive_from, noncompetitive_from)

            if primary.sign_date and primary.sign_date < start:
                status = OUT_OF_PERIOD
            elif issues:
                status = MISMATCH
            elif "nostroy" in families:
                status = IN_NOSTROY
            elif "eis" in families:
                status = CONFIRMED
                if primary.extra.get("grk") == GRK_PROBABLE:
                    subj = (primary.subject or "").strip().replace("ё", "е")
                    if GRK_SUBJECT.search(subj) and not NOT_BUILD_SUBJECT.search(subj):
                        pass  # строительные работы по ОКПД2 и строительный предмет — подтверждаем
                    else:
                        status = NEEDS_CHECK
                        why = "по предмету не строительный подряд" if NOT_BUILD_SUBJECT.search(subj) else "робот не уверен, что договор по ГрК"
                        issues.append(why + ": " + primary.extra.get("grk_note", ""))
                if primary.extra.get("grk") == UNKNOWN:
                    # 10.11: без кода ОКПД2 робот не знает, подряд ли это вообще — решает оператор
                    status = NEEDS_CHECK
                    issues.append("нет кода ОКПД2 ни в ИКЗ, ни в позициях — проверить по предмету, подряд ли это")
                if not primary.customer_name:
                    status = NEEDS_CHECK
                    issues.append("нет наименования заказчика — найти по ИНН " + (primary.customer_inn or "?"))
            elif "protocol" in families:
                status = PROBABLE
                issues.append("победа по протоколу закупки, договор в ЕИС пока не найден")
            elif "notice" in families:
                status = DECLARED
            elif "finpod" in families:
                status = PROBABLE
            else:
                status = PROBABLE

            records.append(ReconciledRecord(
                member=member, status=status, primary=primary,
                sources=sorted({c.source for c in group}), issues=issues, others=group[1:],
            ))
    _hold_doubtful(records)
    records.sort(key=lambda r: (r.member.inn, r.primary.sign_date or date.min))
    return records
