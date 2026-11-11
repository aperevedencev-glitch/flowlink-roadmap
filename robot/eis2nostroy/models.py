"""Общие структуры данных."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
import re
from datetime import date


@dataclass
class Member:
    inn: str
    name: str = ""
    odo_date: date | None = None   # дата присвоения уровня ответственности по ОДО
    joined: date | None = None     # дата вступления в СРО (если известна)


@dataclass
class Contract:
    """Договор (контракт) из любого источника, приведённый к одному виду."""
    source: str                     # eis_44, eis_223, eis_615, finpod, notice, nostroy
    law: str                        # 44, 223, 615 или "" если неизвестно
    supplier_inn: str
    reg_number: str = ""            # реестровый номер договора в ЕИС
    notice_number: str = ""         # номер извещения о закупке
    customer_inn: str = ""
    customer_name: str = ""
    sign_date: date | None = None
    price: float | None = None
    currency: str = "RUB"
    subject: str = ""
    competitive: bool | None = None # True — конкурентный способ, None — неизвестно
    status: str = ""                # статус исполнения в источнике
    raw_file: str = ""              # из какого файла архива взят
    extra: dict = field(default_factory=dict)

    def to_json(self) -> dict:
        d = asdict(self)
        d["sign_date"] = self.sign_date.isoformat() if self.sign_date else None
        return d

    @classmethod
    def from_json(cls, d: dict) -> "Contract":
        d = dict(d)
        if d.get("sign_date"):
            d["sign_date"] = date.fromisoformat(d["sign_date"])
        return cls(**d)


def is_customer_name(text: str) -> bool:
    """Похоже ли на наименование заказчика, а не на основание закупки или номер статьи."""
    t = (text or "").strip().lower()
    if not t:
        return False
    return not re.match(r"^(пункт|п\.\s*\d|часть|ч\.\s*\d|статья|ст\.\s*\d|закупка|в соответствии|на основании|\d)", t)


# Статусы записи после сверки (как в коммерческом предложении)
CONFIRMED = "Подтверждён"
PROBABLE = "Вероятный"
MISMATCH = "Расхождение"
DECLARED = "Заявлен членом"
IN_NOSTROY = "Уже в НОСТРОЙ"
OUT_OF_PERIOD = "Вне периода"
NEEDS_CHECK = "Требует проверки"   # договор из ЕИС, но не всё ясно: ГрК, заказчик
EXCLUDED = "Исключён оператором"  # решение оператора в data/decisions.xlsx

ACTIONS = {
    CONFIRMED: "Передать в НОСТРОЙ",
    PROBABLE: "Запросить подтверждение у члена СРО",
    MISMATCH: "Выбрать верное значение, при необходимости запросить члена",
    DECLARED: "Проверить приложенные документы",
    IN_NOSTROY: "Ничего",
    OUT_OF_PERIOD: "Не передавать",
    NEEDS_CHECK: "Проверить по замечанию: отнесение к ГрК, заказчик",
    EXCLUDED: "Не передавать (решение оператора)",
}


@dataclass
class ReconciledRecord:
    member: Member
    status: str
    primary: Contract                       # основная запись (самый надёжный источник)
    sources: list[str] = field(default_factory=list)
    issues: list[str] = field(default_factory=list)
    others: list[Contract] = field(default_factory=list)

    @property
    def action(self) -> str:
        return ACTIONS.get(self.status, "")
