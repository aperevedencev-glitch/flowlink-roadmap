"""Сведения о членах СРО из открытого реестра НОСТРОЙ (reestr.nostroy.ru).

    python nostroy_reestr.py        — обновить data\\nostroy_members.json по всем ИНН из data\\members.xlsx

Берёт по каждому члену: статус членства, взносы в компенсационные фонды (КФ ВВ, КФ ОДО),
уровни ответственности, совокупный размер обязательств по договорам (как его посчитала СРО),
право на договоры с ОДО и дату взноса в КФ ОДО, договоры страхования (срок окончания),
последние проверки. Токен не нужен — реестр открытый. Около 5 секунд на члена.
"""
from __future__ import annotations

import json
import time
from datetime import datetime
from pathlib import Path

import openpyxl
import requests

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"
OUT_FILE = DATA / "nostroy_members.json"
API = "https://reestr.nostroy.ru/api"
SRO_REG = "СРО-С-130-21122009"   # «Стройкорпорация»
PAUSE = 0.5


def _inns() -> list[str]:
    ws = openpyxl.load_workbook(DATA / "members.xlsx", read_only=True, data_only=True).worksheets[0]
    return [str(r[0]).strip() for r in list(ws.iter_rows(values_only=True))[1:] if r and r[0]]


def _post(http: requests.Session, path: str, body: dict) -> dict:
    for attempt in range(4):
        try:
            r = http.post(f"{API}/{path}", json=body, timeout=40)
            r.raise_for_status()
            return r.json().get("data") or {}
        except (requests.RequestException, ValueError):
            if attempt == 3:
                raise
            time.sleep(5 * (attempt + 1))
    return {}


def _member(http: requests.Session, inn: str) -> dict:
    rows = (_post(http, "sro/all/member/list", {"filters": {"inn": inn}, "page": 1, "pageCount": 20}) or {}).get("data") or []
    ours = [m for m in rows if (m.get("sro") or {}).get("registration_number") == SRO_REG]
    m = next((x for x in ours if (x.get("member_status") or {}).get("code") == "1"), None) or (ours or rows or [None])[0]
    if not m:
        return {"inn": inn, "found": False}
    i = _post(http, f"member/{m['id']}/info", {})
    time.sleep(PAUSE)
    right = i.get("right") or {}
    return {
        "inn": inn, "found": True, "id": m["id"],
        "name": i.get("full_description") or m.get("full_description"),
        "status": (m.get("member_status") or {}).get("title"),
        "ours": bool(ours),
        "sro": (i.get("sro") or {}).get("short_description"),
        "kf_vv": i.get("compensation_fund_fee_vv"), "kf_odo": i.get("compensation_fund_fee_odo"),
        "lvl_vv": i.get("responsibility_level_vv"), "lvl_odo": i.get("responsibility_level_odo"),
        "liability": i.get("members_total_liability"),
        "liability_date": i.get("members_total_liability_last_count_dt_date_string"),
        "is_odo": right.get("is_odo"), "right_status": (right.get("right_status") or {}).get("title"),
        "odo_paid": right.get("compensation_fund_fee_odopayment_date"),
        "insurances": [{"begin": x.get("begin_date"), "end": x.get("end_date"), "sum": x.get("insurance_summ"),
                        "insurer": x.get("insurer"), "object": (x.get("object") or {}).get("short_title")}
                       for x in (i.get("insurances") or [])],
        "checks": [{"date": c.get("check_date"), "type": (c.get("member_check_type") or {}).get("title"),
                    "result": (c.get("member_check_result") or {}).get("title")} for c in (i.get("checks") or [])[:3]],
        "updated": i.get("last_updated_at_date_time_string"),
    }


def refresh(max_age_hours: float = 0) -> bool:
    """Обновить файл, если он старше max_age_hours. Возвращает True, если файл свежий."""
    if max_age_hours and OUT_FILE.exists():
        try:
            got = datetime.fromisoformat(json.loads(OUT_FILE.read_text(encoding="utf-8"))["fetched_at"])
            if (datetime.now() - got).total_seconds() < max_age_hours * 3600:
                return True
        except (OSError, ValueError, KeyError):
            pass
    http = requests.Session()
    http.headers["User-Agent"] = "Flowlink-SRO-robot"
    inns = _inns()
    print(f"Реестр НОСТРОЙ: обновляю сведения о {len(inns)} членах (около {len(inns) * 5 // 60 + 1} мин)…")
    res, errors = [], []
    for n, inn in enumerate(inns, 1):
        try:
            res.append(_member(http, inn))
        except Exception as exc:  # noqa: BLE001 — один член не должен ронять остальных
            errors.append(f"{inn}: {exc}")
        if n % 20 == 0:
            print(f"  {n} из {len(inns)}")
    if not res:
        print("Реестр НОСТРОЙ недоступен, оставляю прежние сведения.", *errors[:3], sep="\n  ")
        return False
    DATA.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_text(json.dumps({"fetched_at": datetime.now().isoformat(timespec="seconds"), "members": res,
                                    "errors": errors}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Реестр НОСТРОЙ: сведения о {len(res)} членах сохранены, ошибок {len(errors)}.")
    return True


if __name__ == "__main__":
    raise SystemExit(0 if refresh() else 1)
