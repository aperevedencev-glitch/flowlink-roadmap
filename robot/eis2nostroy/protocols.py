"""Договоры 223-ФЗ через протоколы закупок.

В карточках договоров 223-ФЗ (с 2019 года) поставщик в открытых данных ЕИС
не указан: договор нельзя найти по ИНН члена СРО. Победитель с ИНН есть
в протоколе закупки (purchaseProtocol и его разновидности), а договор
ссылается на протокол:

    протокол  guid ───────────────┐
              purchaseNoticeNumber ├─ договор  contractConfirmingDocs/contractDoc/protocolInfoGuid
              lot/guid ────────────┘           purchaseNoticeInfo/purchaseNoticeNumber + lotGuid

Модуль в два шага:
1. разбирает протоколы и запоминает победы членов СРО (protocol_wins.jsonl);
2. разбирает договоры 223-ФЗ и, если поставщика в договоре нет, ищет
   победу члена по guid протокола или по паре «извещение + лот».

Победа без найденного договора тоже попадает в сверку — со статусом
«Вероятный»: договор мог быть ещё не опубликован или не заключён.
"""
from __future__ import annotations

import json
import logging
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass, field
from datetime import date
from pathlib import Path

from .models import Contract
from .tabular import to_date, to_float

log = logging.getLogger(__name__)

PROTOCOL_SOURCE = "eis_223p"     # победа по протоколу, договор не найден
TRUE = {"true", "t", "1", "yes", "y", "да"}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _child(el: ET.Element | None, *names: str) -> ET.Element | None:
    """Первый прямой потомок с одним из локальных имён."""
    if el is None:
        return None
    for c in el:
        if _local(c.tag) in names:
            return c
    return None


def _text(el: ET.Element | None, *path: str) -> str:
    """Текст по цепочке локальных имён прямых потомков."""
    cur = el
    for name in path:
        cur = _child(cur, name)
        if cur is None:
            return ""
    return (cur.text or "").strip() if cur is not None else ""


def _find_all(el: ET.Element, name: str) -> list[ET.Element]:
    return [e for e in el.iter() if _local(e.tag) == name]


def _org(el: ET.Element | None) -> tuple[str, str]:
    main = _child(el, "mainInfo") if el is not None else None
    if main is None:
        return "", ""
    return _text(main, "inn"), _text(main, "fullName") or _text(main, "shortName")


def is_protocol(root: ET.Element) -> bool:
    return _local(root.tag).lower().startswith("purchaseprotocol")


def is_contract223(root: ET.Element) -> bool:
    return _local(root.tag).lower().startswith("contract") and bool(_find_all(root, "contractData"))


# ----------------------------------------------------------------- протоколы
@dataclass
class ProtocolWin:
    """Победа (или единственная допущенная заявка) члена СРО в протоколе."""
    supplier_inn: str
    supplier_name: str
    notice_number: str
    protocol_guid: str
    protocol_number: str
    protocol_type: str
    protocol_date: str          # ISO
    customer_inn: str
    customer_name: str
    lot_guid: str
    lot_number: str
    subject: str
    initial_sum: float | None
    price: float | None
    method: str
    reason: str                 # почему считаем победой
    raw_file: str = ""
    extra: dict = field(default_factory=dict)

    @property
    def key(self) -> tuple:
        return (self.supplier_inn, self.notice_number, self.lot_guid or self.lot_number)


def _winner_reason(app: ET.Element, apps_in_lot: int) -> str:
    place = _text(app, "commissionDecisionPlace") or _text(app, "applicationPlace") or _text(app, "place")
    if place == "1":
        return "первое место в решении комиссии"
    if _text(app, "contractSigned").lower() in TRUE:
        return "с участником заключён договор"
    for name in ("winner", "winnerIndication", "isWinner"):
        if _text(app, name).lower() in TRUE:
            return "признан победителем"
    accepted = _text(app, "accepted").lower() in TRUE
    if apps_in_lot == 1 and accepted:
        return "единственная допущенная заявка"
    return ""


def parse_protocol_xml(data: bytes, member_inns: set[str], raw_file: str = "") -> list[ProtocolWin]:
    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        log.warning("Не разобран протокол %s: %s", raw_file, exc)
        return []
    if not is_protocol(root):
        return []
    out: list[ProtocolWin] = []
    for pdata in _find_all(root, "purchaseProtocolData") or [root]:
        notice = _text(pdata, "purchaseInfo", "purchaseNoticeNumber")
        method = _text(pdata, "purchaseInfo", "purchaseCodeName") or _text(pdata, "purchaseInfo", "purchaseMethodName")
        cust_inn, cust_name = _org(_child(pdata, "customer"))
        if not cust_inn:
            cust_inn, cust_name = _org(_child(pdata, "placer"))
        pdate = to_date(_text(pdata, "protocolSignDate") or _text(pdata, "procedureDate")
                        or _text(pdata, "publicationDateTime")[:10])
        base = dict(
            notice_number=notice,
            protocol_guid=_text(pdata, "guid"),
            protocol_number=_text(pdata, "registrationNumber"),
            protocol_type=_text(pdata, "typeName"),
            protocol_date=pdate.isoformat() if pdate else "",
            customer_inn=cust_inn, customer_name=cust_name, method=method, raw_file=raw_file,
        )
        for lot_apps in _find_all(pdata, "protocolLotApplications"):
            lot = _child(lot_apps, "lot")
            apps = [a for a in lot_apps if _local(a.tag) == "application"]
            for app in apps:
                info = _child(app, "supplierInfo")
                inn = _text(info, "inn")
                if inn not in member_inns:
                    continue
                reason = _winner_reason(app, len(apps))
                if not reason:
                    continue
                out.append(ProtocolWin(
                    supplier_inn=inn,
                    supplier_name=_text(info, "name"),
                    lot_guid=_text(lot, "guid"),
                    lot_number=_text(lot, "ordinalNumber"),
                    subject=(_text(lot, "subject") or _text(pdata, "purchaseInfo", "name"))[:500],
                    initial_sum=to_float(_text(lot, "initialSum")),
                    price=to_float(_text(app, "price") or _text(app, "additionalPrice")),
                    reason=reason,
                    **base,
                ))
    return out


class WinStore:
    """Победы членов СРО по протоколам: JSONL, без дублей. Ключи — в памяти, как в ContractStore."""

    def __init__(self, path: Path):
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)
        self._keys: set | None = None
        self._stamp: tuple | None = None

    def _file_stamp(self) -> tuple | None:
        try:
            st = self.path.stat()
        except FileNotFoundError:
            return None
        return (st.st_size, st.st_mtime_ns)

    def load(self) -> list[ProtocolWin]:
        if not self.path.exists():
            self._keys, self._stamp = set(), None
            return []
        stamp = self._file_stamp()
        seen, out, keys = {}, [], set()
        with open(self.path, encoding="utf-8") as fh:
            for line in fh:
                if line.strip():
                    try:
                        w = ProtocolWin(**json.loads(line))
                    except (ValueError, TypeError):
                        log.warning("%s: пропущена повреждённая строка", self.path.name)
                        continue
                    keys.add((w.key, w.protocol_guid))
                    # более поздний протокол (итоговый) заменяет ранний
                    if w.key in seen:
                        out[seen[w.key]] = w
                    else:
                        seen[w.key] = len(out)
                        out.append(w)
        self._keys, self._stamp = keys, stamp
        return out

    def add(self, wins: list[ProtocolWin]) -> int:
        if self._keys is None or self._file_stamp() != self._stamp:
            self.load()
        keys = self._keys
        new = []
        for w in wins:
            k = (w.key, w.protocol_guid)
            if k not in keys:
                keys.add(k)
                new.append(w)
        if new:
            with open(self.path, "a", encoding="utf-8") as fh:
                if _ends_without_newline(self.path):
                    fh.write("\n")   # после оборванной строки — с новой строки, чтобы не склеить записи
                for w in new:
                    fh.write(json.dumps(asdict(w), ensure_ascii=False) + "\n")
        self._stamp = self._file_stamp()
        return len(new)


class WinIndex:
    def __init__(self, wins: list[ProtocolWin]):
        self.by_guid: dict[str, list[ProtocolWin]] = {}
        self.by_lot: dict[tuple, list[ProtocolWin]] = {}
        self.by_notice: dict[str, list[ProtocolWin]] = {}
        for w in wins:
            if w.protocol_guid:
                self.by_guid.setdefault(w.protocol_guid, []).append(w)
            if w.notice_number:
                self.by_notice.setdefault(w.notice_number, []).append(w)
                if w.lot_guid:
                    self.by_lot.setdefault((w.notice_number, w.lot_guid), []).append(w)

    def __bool__(self) -> bool:
        return bool(self.by_notice or self.by_guid)

    def match(self, protocol_guids: list[str], notice: str, lot_guid: str, lot_num: str) -> list[ProtocolWin]:
        for g in protocol_guids:
            if g in self.by_guid:
                hits = self.by_guid[g]
                return [w for w in hits if not lot_guid or not w.lot_guid or w.lot_guid == lot_guid] or hits
        if notice and lot_guid and (notice, lot_guid) in self.by_lot:
            return self.by_lot[(notice, lot_guid)]
        cand = self.by_notice.get(notice, [])
        if lot_num:
            cand = [w for w in cand if not w.lot_number or w.lot_number == lot_num]
        # без лота берём, только если член выиграл единственный лот закупки
        return cand if len({w.lot_guid or w.lot_number for w in cand}) == 1 else []


# ------------------------------------------------------------------ договоры
def contracts_from_223(data: bytes, index: WinIndex, raw_file: str = "") -> list[Contract]:
    """Договор 223-ФЗ без поставщика → договор члена СРО через протокол."""
    if not index:
        return []
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []
    if not is_contract223(root):
        return []
    out: list[Contract] = []
    for cd in _find_all(root, "contractData"):
        notice = _text(cd, "purchaseNoticeInfo", "purchaseNoticeNumber")
        guids = [(_text(d, "protocolInfoGuid")) for d in _find_all(cd, "contractDoc")]
        wins = index.match([g for g in guids if g], notice, _text(cd, "lotGuid"), _text(cd, "lotNum"))
        if not wins:
            continue
        cust_inn, cust_name = _org(_child(cd, "customer"))
        reg = _text(cd, "contractRegNumber") or _text(cd, "registrationNumber")
        end = to_date(_text(cd, "endExecutionDate"))
        codes = [(c.text or "").strip() for e in _find_all(cd, "okpd2") for c in e if _local(c.tag) == "code"]
        from .grk import classify as grk_classify
        grk_cls, grk_note = grk_classify("", codes)
        for w in wins:
            out.append(Contract(
                source="eis_223",
                law="223",
                supplier_inn=w.supplier_inn,
                reg_number=reg,
                notice_number=notice,
                customer_inn=cust_inn or w.customer_inn,
                customer_name=cust_name or w.customer_name,
                sign_date=to_date(_text(cd, "contractDate")),
                price=to_float(_text(cd, "price")),
                subject=(_text(cd, "subjectContract") or w.subject)[:500],
                competitive=True,
                status=("завершён по сроку" if end and end < date.today() else "исполнение"),
                raw_file=raw_file,
                extra={
                    "via": "протокол",
                    "protocol": w.protocol_number,
                    "protocol_guid": w.protocol_guid,
                    "protocol_date": w.protocol_date,
                    "version": _text(cd, "version"),
                    "contract_number": _text(cd, "name"),
                    "end_date": end.isoformat() if end else "",
                    "winner_reason": w.reason,
                    "okpd2": sorted(set(codes))[:5],
                    "grk": grk_cls,
                    "grk_note": grk_note,
                },
            ))
    return out


def unlinked_wins(wins: list[ProtocolWin], contracts: list[Contract]) -> list[Contract]:
    """Победы, для которых договор в ЕИС ещё не найден, — в сверку как «вероятные»."""
    linked_guids = {c.extra.get("protocol_guid") for c in contracts if c.extra.get("protocol_guid")}
    linked = {(c.supplier_inn, c.notice_number) for c in contracts if c.law == "223" and c.notice_number}
    out = []
    for w in wins:
        if w.protocol_guid in linked_guids or (w.supplier_inn, w.notice_number) in linked:
            continue
        out.append(Contract(
            source=PROTOCOL_SOURCE,
            law="223",
            supplier_inn=w.supplier_inn,
            notice_number=w.notice_number,
            customer_inn=w.customer_inn,
            customer_name=w.customer_name,
            sign_date=to_date(w.protocol_date),
            price=w.price,
            subject=w.subject,
            competitive=True,
            status="победа по протоколу, договор в ЕИС не найден",
            raw_file=w.raw_file,
            extra={"via": "протокол", "protocol": w.protocol_number, "protocol_guid": w.protocol_guid,
                   "winner_reason": w.reason},
        ))
    return out


# ---------------------------------------------------- справочник типов документов
CODE_NAMES = ("code", "docType", "documentType", "typeCode", "type")


def protocol_types_from_nsi(xml_blobs: list[bytes]) -> tuple[list[str], list[tuple[str, str]]]:
    """Из справочника nsiDocumentTypes223: типы протоколов, отдаваемые публично.

    Возвращает (коды протоколов, все публичные типы [(код, наименование)]).
    Разбор по локальным именам: запись — элемент с потомком distributedPublicly.
    """
    public: list[tuple[str, str]] = []
    for data in xml_blobs:
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        for rec in root.iter():
            kids = {_local(c.tag): (c.text or "").strip() for c in rec}
            if "distributedPublicly" not in kids:
                continue
            if kids["distributedPublicly"].lower() not in TRUE:
                continue
            code = next((kids[n] for n in CODE_NAMES if kids.get(n)), "")
            if code:
                public.append((code, kids.get("name") or kids.get("fullName") or ""))
    seen, uniq = set(), []
    for code, name in public:
        if code not in seen:
            seen.add(code)
            uniq.append((code, name))
    protocols = [c for c, n in uniq if "protocol" in c.lower() or "протокол" in n.lower()]
    return protocols, uniq


# ------------------------------------------- договор → протоколы по номеру закупки
from .grk import is_construction  # noqa: E402
CONSTRUCTION_OKPD2 = ("41.", "42.", "43.")


def contract_candidate(data: bytes, okpd_prefixes: tuple[str, ...] = CONSTRUCTION_OKPD2) -> list[dict]:
    """Договор 223-ФЗ без поставщика, но со строительными позициями ОКПД2.

    Возвращает для каждого такого договора номер закупки — по нему сервис ЕИС
    отдаёт протоколы (getDocsByReestrNumber, подсистема RI223).
    """
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []
    if not is_contract223(root):
        return []
    out = []
    for cd in _find_all(root, "contractData"):
        notice = _text(cd, "purchaseNoticeInfo", "purchaseNoticeNumber")
        if not notice:
            continue          # закупка без извещения — протокола нет
        codes = [(c.text or "").strip() for e in _find_all(cd, "okpd2") for c in e if _local(c.tag) == "code"]
        ok = (is_construction if okpd_prefixes == CONSTRUCTION_OKPD2
              else (lambda code: code.startswith(okpd_prefixes)))
        if okpd_prefixes and not any(ok(code) for code in codes):
            continue      # не строительные работы по ГрК (в т.ч. 41.1 — разработка проектов)
        out.append({"notice": notice, "reg": _text(cd, "contractRegNumber") or _text(cd, "registrationNumber"),
                    "okpd2": codes[:5]})
    return out


def _ends_without_newline(path: Path) -> bool:
    try:
        with open(path, "rb") as fh:
            fh.seek(-1, 2)
            return fh.read(1) != b"\n"
    except OSError:
        return False
