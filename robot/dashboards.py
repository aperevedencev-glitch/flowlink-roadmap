"""Ежедневные HTML-дашборды Flowlink: руководителя СРО и ответственного сотрудника.

    python dashboards.py            — собрать оба дашборда из того, что уже есть в папке data

Читает: последнюю сверку data\\out\\sverka_*.xlsx, список договоров data\\work\\eis_contracts.jsonl,
членов data\\members.xlsx, сведения реестра НОСТРОЙ data\\nostroy_members.json (взносы в КФ,
уровни ответственности, страхование), сроки членских взносов data\\vznosy.xlsx и свежие логи data\\logs. Пишет в data\\out:
    Flowlink — дашборд руководителя.html
    Flowlink — дашборд сотрудника.html
и дописывает дневной срез в data\\out\\dash_history.json (из него строится динамика по дням).
Токен ЕИС не нужен, в интернет не ходит. Файлы открываются в любом браузере без интернета.
"""
from __future__ import annotations

import html
import json
import re
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"
OUT = DATA / "out"
SRO = "Ассоциация СРО «Стройкорпорация»"
PERIOD = date(2017, 7, 1)
LIMIT_SRO = 10_000_000
DUES_FILE = DATA / "vznosy.xlsx"
DUES_WARN_DAYS = 30      # за сколько дней до срока показывать взнос как «приближается»
DUES_URGENT_DAYS = 7     # срочно


# ---------- чтение данных ----------

def version_flags() -> list[dict]:
    """Договоры последнего пакета, у которых проверка версий нашла новую версию (versions_*.xlsx)."""
    reports = sorted(OUT.glob("versions_2*.xlsx"), key=lambda p: p.stat().st_mtime)
    packages = sorted(OUT.glob("nostroy_*.csv"), key=lambda p: p.stat().st_mtime)
    if not reports or not packages:
        return []
    try:
        ws = openpyxl.load_workbook(reports[-1], read_only=True, data_only=True)["К проверке"]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        text = packages[-1].read_text(encoding="utf-8-sig", errors="replace")
    except (OSError, KeyError, ValueError):
        return []
    return [{"verdict": r[0], "note": r[1], "reg": str(r[3]), "member": r[5], "price": r[7]}
            for r in rows if r and r[3] and str(r[3]) in text]


def latest_sverka() -> Path | None:
    files = sorted(OUT.glob("sverka_2*.xlsx"), key=lambda p: p.stat().st_mtime)
    files = [f for f in files if "baza" not in f.name]
    return files[-1] if files else None


def to_float(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace("\xa0", "").replace(" ", "").replace(",", "."))
    except ValueError:
        return None


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v or "").strip()[:10]
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def read_sverka(path: Path) -> tuple[list[dict], str]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Все записи"]
    rows = list(ws.iter_rows(values_only=True))
    stamp = str(rows[1][0] or "") if len(rows) > 1 else ""
    hdr_i = next(i for i, r in enumerate(rows) if r and r[0] == "Статус")
    hdr = [str(h or "") for h in rows[hdr_i]]
    out = []
    for r in rows[hdr_i + 1:]:
        if not r or not r[0]:
            continue
        d = dict(zip(hdr, r))
        out.append({
            "status": str(d.get("Статус") or ""),
            "action": str(d.get("Действие оператора") or ""),
            "inn": str(d.get("ИНН члена СРО") or ""),
            "member": str(d.get("Наименование члена СРО") or ""),
            "law": str(d.get("Основание (закон)") or ""),
            "reg": str(d.get("Реестровый номер договора в ЕИС") or ""),
            "notice": str(d.get("Номер извещения") or ""),
            "customer": str(d.get("Заказчик") or ""),
            "date": to_date(d.get("Дата заключения")),
            "price": to_float(d.get("Цена договора, ₽")),
            "subject": str(d.get("Предмет договора") or ""),
            "grk": str(d.get("Градостроительный кодекс") or ""),
            "note": str(d.get("Замечания") or ""),
        })
    m = re.search(r"(\d\d\.\d\d\.\d{4} \d\d:\d\d)", stamp)
    return out, (m.group(1) if m else datetime.fromtimestamp(path.stat().st_mtime).strftime("%d.%m.%Y %H:%M"))


def read_members() -> dict[str, str]:
    p = DATA / "members.xlsx"
    if not p.exists():
        return {}
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    res = {}
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r and r[0]:
            res[str(r[0]).strip()] = str(r[1] or "").strip()
    return res


def read_dues(today: date) -> tuple[list[dict], bool]:
    """Взносы членов из data\\vznosy.xlsx. Возвращает (неоплаченные со сроком, есть ли файл).

    Колонки: ИНН | Наименование | Вид взноса | Период | Срок оплаты | Сумма, ₽ | Дата оплаты | Примечание.
    Оплаченные (заполнена «Дата оплаты») и строки без срока не показываются.
    """
    if not DUES_FILE.exists():
        return [], False
    wb = openpyxl.load_workbook(DUES_FILE, read_only=True, data_only=True)
    ws = wb["Взносы"] if "Взносы" in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = next((i for i, r in enumerate(rows) if r and str(r[0] or "").strip() == "ИНН"), None)
    if hdr_i is None:
        return [], True
    hdr = [str(h or "").strip() for h in rows[hdr_i]]
    res = []
    for r in rows[hdr_i + 1:]:
        if not r or not r[0]:
            continue
        d = dict(zip(hdr, r))
        due = to_date(d.get("Срок оплаты"))
        if not due or d.get("Дата оплаты") not in (None, ""):
            continue
        days = (due - today).days
        state = ("просрочен" if days < 0 else "срочно" if days <= DUES_URGENT_DAYS
                 else "приближается" if days <= DUES_WARN_DAYS else "позже")
        res.append({"inn": str(d.get("ИНН") or "").strip(), "member": str(d.get("Наименование") or "").strip(),
                    "kind": str(d.get("Вид взноса") or "Взнос").strip(), "period": str(d.get("Период") or "").strip(),
                    "due": due, "days": days, "sum": to_float(d.get("Сумма, ₽")), "note": str(d.get("Примечание") or "").strip(),
                    "state": state})
    res.sort(key=lambda x: x["due"])
    return res, True


def dues_template(members: dict[str, str]) -> None:
    """Создать data\\vznosy.xlsx с членами СРО, если файла ещё нет (существующий не трогает)."""
    if DUES_FILE.exists() or not members:
        return
    from openpyxl.styles import Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Взносы"
    ws.append(["Сроки внесения взносов членами СРО. Заполните «Срок оплаты» и «Сумма»; после оплаты — «Дата оплаты». "
               "На каждый взнос и период — отдельная строка (строку можно копировать)."])
    ws.append([])
    hdr = ["ИНН", "Наименование", "Вид взноса", "Период", "Срок оплаты", "Сумма, ₽", "Дата оплаты", "Примечание"]
    ws.append(hdr)
    for c in ws[3]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="EFEBE3")
    for inn, name in members.items():
        ws.append([inn, name, "Членский взнос", "", None, None, None, ""])
    dv = DataValidation(type="list", allow_blank=True, formula1='"Членский взнос,Взнос в КФ ВВ,Взнос в КФ ОДО,'
                        'Отчисления в НОСТРОЙ,Целевой взнос,Вступительный взнос"')
    ws.add_data_validation(dv)
    dv.add(f"C4:C{3 + len(members) * 8}")
    for col, w in zip("ABCDEFGH", (14, 60, 22, 14, 14, 14, 14, 30)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=4, max_row=3 + len(members) * 8, min_col=5, max_col=7):
        for c in row:
            c.number_format = "DD.MM.YYYY" if c.column != 6 else "#,##0.00"
    ws.freeze_panes = "A4"
    DUES_FILE.parent.mkdir(parents=True, exist_ok=True)
    wb.save(DUES_FILE)
    print(f"Создан шаблон сроков взносов: {DUES_FILE} — заполните сроки, дашборды покажут приближающиеся.")


# Ч. 11, 13 ст. 55.16 ГрК: взнос в КФ ОДО и предельный размер обязательств по уровням ответственности
ODO_CAP = {"1": 90e6, "2": 500e6, "3": 3e9, "4": 10e9, "5": None}
ODO_FEE = {"1": 200_000, "2": 2_500_000, "3": 4_500_000, "4": 7_000_000, "5": 25_000_000}
ODO_NEAR = 0.8   # доля предела уровня, с которой предупреждать о доплате в КФ ОДО


def read_registry() -> tuple[list[dict], str]:
    p = DATA / "nostroy_members.json"
    if not p.exists():
        return [], ""
    try:
        d = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return [], ""
    when = d.get("fetched_at", "")
    try:
        when = datetime.fromisoformat(when).strftime("%d.%m.%Y %H:%M")
    except ValueError:
        pass
    return d.get("members") or [], when


def registry_items(reg: list[dict], confirmed: list[dict], today: date) -> list[dict]:
    """Сроки и взносы по данным реестра НОСТРОЙ: доплата в КФ ОДО, право ОДО, страхование."""
    big_recent = defaultdict(float)
    for r in confirmed:
        if r["date"] and (today - r["date"]).days <= 365 and (r["price"] or 0) > LIMIT_SRO and law_key(r["law"]) != "ПП-615":
            big_recent[r["inn"]] += r["price"] or 0
    items = []
    for m in reg:
        if not m.get("found") or m.get("status") != "Является членом":
            continue
        base = {"inn": m["inn"], "member": m.get("name") or "", "period": "", "src": "реестр НОСТРОЙ"}
        lvl = str(((m.get("lvl_odo") or {}).get("code")) or "")
        cap = ODO_CAP.get(lvl)
        liab = to_float(m.get("liability")) or 0
        paid = to_float(m.get("kf_odo")) or 0
        if m.get("is_odo") and cap and liab:
            share = liab / cap
            if share >= ODO_NEAR:
                nxt = str(int(lvl) + 1)
                extra = max(0.0, ODO_FEE.get(nxt, 0) - paid)
                items.append({**base, "kind": "Взнос в КФ ОДО (повышение уровня)", "due": None, "days": None,
                              "sum": extra or None, "state": "срочно" if share >= 1 else "приближается",
                              "pill": f"обязательства {share:.0%} уровня",
                              "note": (f"Обязательства {money(liab)} при пределе {money(cap)} ({lvl}-й уровень ОДО), "
                                       f"расчёт СРО на {m.get('liability_date') or '—'}. Сумма — разница взносов "
                                       f"{nxt}-го и текущего уровня, ориентировочно.")})
        elif not m.get("is_odo") and big_recent.get(m["inn"]):
            items.append({**base, "kind": "Взнос в КФ ОДО (право на конкурентные договоры)", "due": None, "days": None,
                          "sum": ODO_FEE["1"] - paid if paid < ODO_FEE["1"] else None, "state": "срочно",
                          "pill": "нет права ОДО",
                          "note": f"За 12 месяцев договоров свыше 10 млн ₽ на {money(big_recent[m['inn']])}, "
                                  f"а права заключать договоры с ОДО в реестре нет — проверить, конкурентные ли они."})
        if m.get("right_status") == "Приостановлено":
            items.append({**base, "kind": "Право выполнять работы приостановлено", "due": None, "days": None, "sum": None,
                          "state": "срочно", "pill": "право приостановлено",
                          "note": "По реестру право приостановлено — выяснить причину (страхование, взносы, нарушения) и срок устранения."})
        ends = [to_date(x.get("end")) for x in (m.get("insurances") or [])]
        ends = [e for e in ends if e]
        if ends:
            end = max(ends)
            days = (end - today).days
            if days <= DUES_WARN_DAYS:
                items.append({**base, "kind": "Договор страхования ответственности", "due": end, "days": days,
                              "sum": None, "state": "просрочен" if days < 0 else "срочно" if days <= DUES_URGENT_DAYS else "приближается",
                              "note": ("По реестру последний договор страхования закончился — проверить, есть ли новый полис, и внести его в реестр."
                                       if days < 0 else "Договор страхования по реестру заканчивается — нужен новый полис.")})
    return items


def read_store() -> list[dict]:
    p = DATA / "work" / "eis_contracts.jsonl"
    res = []
    if p.exists():
        for line in p.open(encoding="utf-8"):
            try:
                res.append(json.loads(line))
            except json.JSONDecodeError:
                pass
    return res


QUEUE_STATUSES = ("Требует проверки", "Вероятный", "Расхождение", "Заявлен членом")


def health(reg_when: str) -> dict:
    """Здоровье робота: версия, что идёт сейчас, последняя успешная загрузка, память, эталон."""
    import sys as _s
    _s.path.insert(0, str(HERE))
    h = {"version": "", "busy": "", "rebuild": "", "last_ok": "", "registry": reg_when or "нет",
         "decided": 0, "queue_total": 0, "etalon": ""}
    try:
        from eis2nostroy import __version__
        from eis2nostroy.ops import lock_info
        h["version"] = __version__
        b = lock_info(DATA / "work" / "robot.lock")
        if b:
            h["busy"] = f"{b.get('what')} с {b.get('since')}"
    except Exception:  # noqa: BLE001 — дашборд строится и без пакета
        pass
    prog = DATA / "work" / "rebuild_progress.json"
    if prog.exists():
        try:
            p = json.loads(prog.read_text(encoding="utf-8"))
            done = sum(f["done"] for f in p["feeds"].values())
            total = sum(len(f["archives"]) for f in p["feeds"].values()) or 1
            h["rebuild"] = f"{done / total:.0%} ({num(done)} из {num(total)}), начата {p.get('started')}"
        except (OSError, ValueError, KeyError):
            pass
    for lf in sorted((DATA / "logs").glob("*.log"), key=lambda x: x.stat().st_mtime, reverse=True):
        if not lf.name.startswith(("load", "ranges", "history", "daily")):
            continue
        text = lf.read_text(encoding="utf-8", errors="replace")
        if re.search(r"ошибок: 0\b", text):
            h["last_ok"] = datetime.fromtimestamp(lf.stat().st_mtime).strftime("%d.%m.%Y %H:%M")
            break
    dp = DATA / "decisions.xlsx"
    if dp.exists():
        try:
            ws = openpyxl.load_workbook(dp, read_only=True, data_only=True)["Решения"]
            rows = [r for r in ws.iter_rows(min_row=4, values_only=True) if r and r[0]]
            h["queue_total"] = len(rows)
            h["decided"] = sum(1 for r in rows if len(r) > 10 and r[10])
        except Exception:  # noqa: BLE001
            pass
    try:
        e = json.loads((OUT / "etalon_last.json").read_text(encoding="utf-8"))
        h["etalon"] = f"{e['share']:.0%} на {e['decided']} решениях ({datetime.fromisoformat(e['date']):%d.%m})"
    except (OSError, ValueError, KeyError):
        pass
    return h


def health_card(h: dict) -> str:
    rows = [("Версия робота", h["version"] or "—"),
            ("Сейчас идёт", h["busy"] or "ничего — робот свободен"),
            ("Незаконченная пересборка", h["rebuild"] or "нет"),
            ("Последняя загрузка без ошибок", h["last_ok"] or "—"),
            ("Реестр НОСТРОЙ обновлён", h["registry"]),
            ("Решения оператора", f"{num(h['decided'])} из {num(h['queue_total'])} в файле решений" if h["queue_total"] else "файла решений пока нет"),
            ("Эталон: робот не ошибся", h["etalon"] or "ещё не проверялся")]
    body = "".join(f'<tr><td class="muted">{esc(a)}</td><td>{esc(b)}</td></tr>' for a, b in rows)
    return f'<div class="card"><h2>Здоровье робота</h2><table style="min-width:0">{body}</table></div>'


def robot_status() -> dict:
    logs = sorted((DATA / "logs").glob("*.log"), key=lambda p: p.stat().st_mtime)
    if not logs:
        return {"file": "", "when": "", "errors": 0, "last": "", "running": False}
    last = logs[-1]
    text = last.read_text(encoding="utf-8", errors="replace").splitlines()
    tail = [l for l in text if re.match(r"\d{4}-\d\d-\d\d ", l)]
    summ = [int(m.group(1)) for l in text for m in [re.search(r"ошибок: (\d+)", l)] if m]
    errors = summ[-1] if summ else sum(1 for l in text if " ERROR " in l)
    finished = any("Записей после сверки" in l or "Готово" in l for l in text[-15:])
    progress = next((l for l in reversed(tail) if re.search(r"(RD223|RGK|RD615) \d\d\.\d\d\.\d{4}:", l)), "")
    return {
        "file": last.name,
        "when": datetime.fromtimestamp(last.stat().st_mtime).strftime("%d.%m.%Y %H:%M"),
        "errors": errors,
        "last": progress[24:] if progress else (tail[-1][24:] if tail else ""),
        "running": not finished,
    }


# ---------- вспомогательное ----------

def esc(s) -> str:
    return html.escape(str(s if s is not None else ""))


def money(v, unit="млн") -> str:
    if v is None:
        return "—"
    if unit == "млрд":
        return f"{v / 1e9:,.1f}".replace(",", " ").replace(".", ",") + " млрд ₽"
    return f"{v / 1e6:,.1f}".replace(",", " ").replace(".", ",") + " млн ₽"


def num(v) -> str:
    return f"{v:,}".replace(",", " ")


ABBR = [(r"федеральное государственное бюджетное учреждение", "ФГБУ"), (r"федеральное казенное учр(еждение|\.)", "ФКУ"),
        (r"государственное казенное учреждение", "ГКУ"), (r"государственное бюджетное учреждение", "ГБУ"),
        (r"муниципальное казенное учреждение", "МКУ"), (r"муниципальное бюджетное учреждение", "МБУ"),
        (r"федеральное управление автомобиль\w*( дорог)?", "ФУАД"), (r"управление автомобильной магистрали", "УАМ"),
        (r"федерального дорожного агентства", ""), (r"московской области", "МО"), (r"города москвы", "г. Москвы"),
        (r"администрация", "Адм."), (r"городского округа", "г.о."), (r"общество с ограниченной ответственностью", "ООО"),
        (r"публичное акционерное общество", "ПАО"), (r"открытое акционерное общество", "ОАО"),
        (r"акционерное общество", "АО"), (r"публично-правовая компания", "ППК")]


def short(name: str) -> str:
    s = name
    for pat, rep in ABBR:
        s = re.sub(pat, rep, s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"Общество с ограниченной ответственностью", "ООО", s, flags=re.I)
    s = re.sub(r"Акционерное общество", "АО", s, flags=re.I)
    s = re.sub(r"Публичное АО|Публичное акционерное общество", "ПАО", s, flags=re.I)
    s = re.sub(r"Открытое АО|Открытое акционерное общество", "ОАО", s, flags=re.I)
    s = re.sub(r"Индивидуальный предприниматель", "ИП", s, flags=re.I)
    s = re.sub(r'["«»]', "", s)
    return re.sub(r"\s+", " ", s).strip()


def bars(items: list[tuple[str, float, str]], fmt=lambda v: num(int(v)), width=520) -> str:
    """Горизонтальные столбцы: (подпись, значение, класс цвета)."""
    if not items:
        return '<p class="muted">Нет данных</p>'
    mx = max(v for _, v, _ in items) or 1
    row_h, label_w = 26, 215
    h = row_h * len(items)
    out = [f'<svg viewBox="0 0 {width} {h}" class="chart" role="img">']
    for i, (lab, v, cls) in enumerate(items):
        y = i * row_h
        w = max(2, (width - label_w - 90) * v / mx)
        t = lab if len(lab) <= 26 else lab[:25].rstrip() + "…"
        out.append(f'<text x="{label_w - 8}" y="{y + 17}" text-anchor="end" class="lbl">{esc(t)}</text>')
        out.append(f'<rect x="{label_w}" y="{y + 5}" width="{w:.1f}" height="16" rx="3" class="{cls}"><title>{esc(lab)}: {esc(fmt(v))}</title></rect>')
        out.append(f'<text x="{label_w + w + 6}" y="{y + 17}" class="val">{esc(fmt(v))}</text>')
    out.append("</svg>")
    return "".join(out)


def columns(series: list[tuple[str, dict]], keys: list[tuple[str, str]], fmt=lambda v: num(int(v)), height=200) -> str:
    """Столбики с накоплением по годам: series = [(год, {ключ: значение})]."""
    if not series:
        return '<p class="muted">Нет данных</p>'
    mx = max(sum(d.get(k, 0) for k, _ in keys) for _, d in series) or 1
    n = len(series)
    width = max(360, n * 54)
    bw = 34
    top = 18
    out = [f'<svg viewBox="0 0 {width} {height + 40 + top}" class="chart" role="img">']
    for i, (lab, d) in enumerate(series):
        x = 20 + i * (width - 40) / n
        y = height + top
        total = 0
        for k, cls in keys:
            v = d.get(k, 0)
            if v <= 0:
                continue
            hh = height * v / mx
            y -= hh
            total += v
            out.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw}" height="{hh:.1f}" class="{cls}"><title>{esc(lab)} · {esc(k)}: {esc(fmt(v))}</title></rect>')
        out.append(f'<text x="{x + bw / 2:.1f}" y="{height + top + 16}" text-anchor="middle" class="lbl">{esc(lab)}</text>')
        if total:
            out.append(f'<text x="{x + bw / 2:.1f}" y="{y - 4:.1f}" text-anchor="middle" class="val">{esc(fmt(total))}</text>')
    out.append("</svg>")
    return "".join(out)


def legend(keys: list[tuple[str, str]]) -> str:
    return '<div class="legend">' + "".join(f'<span><i class="{c}"></i>{esc(k)}</span>' for k, c in keys) + "</div>"


def eis_link(r: dict) -> str:
    reg = r["reg"]
    if not reg:
        return ""
    if r["law"].startswith("223"):
        u = f"https://zakupki.gov.ru/epz/contractfz223/card/contract-info.html?reestrNumber={reg}"
    elif "615" in r["law"]:
        u = f"https://zakupki.gov.ru/epz/capitalrepairs/card/general-info.html?reestr-number={reg}"
    else:
        u = f"https://zakupki.gov.ru/epz/contract/contractCard/common-info.html?reestrNumber={reg}"
    return f'<a href="{u}" target="_blank" rel="noopener">{esc(reg)}</a>'


CSS = """
:root{--bg:#F6F4EF;--card:#FFFFFF;--ink:#13202E;--muted:#5F6E7D;--line:#E3DFD6;--accent:#E87A2E;
--c44:#2F5D8A;--c223:#E87A2E;--c615:#3E8E6B;--ok:#3E8E6B;--warn:#D9A21B;--bad:#C0503A;--soft:#EFEBE3}
@media (prefers-color-scheme:dark){:root:not([data-theme=light]){--bg:#0F1822;--card:#16222F;--ink:#ECE8E0;--muted:#94A3B3;
--line:#2A3A4B;--soft:#1D2B3A;--c44:#6E9CCB}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.45 "IBM Plex Sans","Segoe UI",Arial,sans-serif}
header{background:#13202E;color:#F3EFE6;padding:22px 28px}header h1{margin:0;font-size:24px;font-weight:600}
header .sub{color:#B9C6D3;font-size:13px;margin-top:4px}header .tag{color:var(--accent);font:600 12px/1 "IBM Plex Mono",Consolas,monospace;letter-spacing:2px}
main{max-width:1240px;margin:0 auto;padding:20px 16px 48px}
.grid{display:grid;gap:14px}.k4{grid-template-columns:repeat(auto-fit,minmax(200px,1fr))}.k2{grid-template-columns:repeat(auto-fit,minmax(420px,1fr))}
@media (max-width:520px){.k2{grid-template-columns:1fr}header{padding:16px}}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px 18px;min-width:0}
.card h2{font-size:16px;margin:0 0 10px;font-weight:600}.card h3{font-size:14px;margin:12px 0 6px}
.kpi .v{font-size:28px;font-weight:600;line-height:1.1}.kpi .l{color:var(--muted);font-size:13px}.kpi .d{font-size:12px;margin-top:6px;color:var(--muted)}
.kpi.ok .v{color:var(--ok)}.kpi.warn .v{color:var(--warn)}.kpi.bad .v{color:var(--bad)}
.muted{color:var(--muted)}.small{font-size:12px}
.chart{width:100%;height:auto;display:block}.chart .lbl{font-size:11px;fill:var(--muted)}.chart .val{font-size:11px;fill:var(--ink)}
.c44{fill:var(--c44);background:var(--c44)}.c223{fill:var(--c223);background:var(--c223)}.c615{fill:var(--c615);background:var(--c615)}
.cok{fill:var(--ok);background:var(--ok)}.cwarn{fill:var(--warn);background:var(--warn)}.cbad{fill:var(--bad);background:var(--bad)}.cacc{fill:var(--accent);background:var(--accent)}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--muted);margin-top:6px}.legend i{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:5px;vertical-align:-1px}
table{width:100%;min-width:1320px;border-collapse:collapse;font-size:12.5px}td.nw{white-space:nowrap}td.sub{max-width:340px}td.note{min-width:220px;max-width:300px;color:var(--muted)}th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}
th{position:sticky;top:0;background:var(--soft);font-weight:600;cursor:pointer;white-space:nowrap}td.n,th.n{text-align:right;white-space:nowrap}
.scroll{max-height:560px;overflow:auto;border:1px solid var(--line);border-radius:8px}
.pill{display:inline-block;padding:1px 8px;border-radius:10px;font-size:12px;background:var(--soft);white-space:nowrap}
.pill.ok{background:rgba(62,142,107,.15);color:var(--ok)}.pill.warn{background:rgba(217,162,27,.18);color:#8a6400}.pill.bad{background:rgba(192,80,58,.15);color:var(--bad)}
input[type=search]{width:100%;max-width:420px;padding:8px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);font:inherit}
.tools{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:10px}.tools select{padding:7px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);font:inherit}
a{color:var(--c44)}footer{color:var(--muted);font-size:12px;margin-top:22px}
ol.steps{margin:6px 0 0 18px;padding:0}ol.steps li{margin:4px 0}
"""

JS = """
function sortTable(th){const t=th.closest('table'),i=[...th.parentNode.children].indexOf(th),b=t.tBodies[0];
const dir=th.dataset.dir==='asc'?'desc':'asc';th.dataset.dir=dir;const num=th.classList.contains('n');
[...b.rows].sort((a,c)=>{let x=a.cells[i].dataset.v??a.cells[i].innerText,y=c.cells[i].dataset.v??c.cells[i].innerText;
if(num){x=parseFloat(x)||0;y=parseFloat(y)||0}return (x>y?1:x<y?-1:0)*(dir==='asc'?1:-1)}).forEach(r=>b.appendChild(r))}
document.querySelectorAll('th').forEach(th=>th.addEventListener('click',()=>sortTable(th)));
function filt(id){const q=(document.getElementById(id+'-q')||{}).value||'';const s=document.getElementById(id+'-s');const sv=s?s.value:'';
const t=document.getElementById(id);let n=0;[...t.tBodies[0].rows].forEach(r=>{const ok=(!q||r.innerText.toLowerCase().includes(q.toLowerCase()))&&(!sv||r.dataset.k===sv);r.style.display=ok?'':'none';if(ok)n++});
const c=document.getElementById(id+'-n');if(c)c.textContent=n}
"""


def page(title: str, sub: str, body: str) -> str:
    return f"""<!doctype html><html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{esc(title)}</title><style>{CSS}</style></head>
<body><header><div class="tag">FLOWLINK · {esc(SRO.upper())}</div><h1>{esc(title)}</h1><div class="sub">{sub}</div></header>
<main>{body}<footer>Собрано автоматически роботом Flowlink {datetime.now():%d.%m.%Y %H:%M}. Источник — ЕИС (zakupki.gov.ru).
Файл обновляется после каждого прогона робота и раз в день по расписанию.</footer></main><script>{JS}</script></body></html>"""


# ---------- расчёты ----------

def law_key(law: str) -> str:
    if law.startswith("223"):
        return "223-ФЗ"
    if "615" in law:
        return "ПП-615"
    return "44-ФЗ"


def build():
    sv_path = latest_sverka()
    if sv_path is None:
        print("Нет файла сверки в data\\out — сначала запустите сверку робота.")
        return 1
    recs, stamp = read_sverka(sv_path)
    members = read_members()
    dues_template(members)
    store = read_store()
    rs = robot_status()

    confirmed = [r for r in recs if r["status"] == "Подтверждён"]          # новое, ещё не передано
    sent = [r for r in recs if r["status"] == "Уже в НОСТРОЙ"]
    excluded = [r for r in recs if r["status"] == "Исключён оператором"]
    check = [r for r in recs if r["status"] in QUEUE_STATUSES]
    grk_all = confirmed + sent                                              # всё под ГрК: для графиков
    to_transfer_sum = sum(r["price"] or 0 for r in confirmed)
    by_law = Counter(law_key(r["law"]) for r in confirmed)
    sum_by_law = defaultdict(float)
    for r in confirmed:
        sum_by_law[law_key(r["law"])] += r["price"] or 0
    no_sro = sum(1 for r in confirmed if (r["price"] or 0) <= LIMIT_SRO and law_key(r["law"]) != "ПП-615")
    members_with = {r["inn"] for r in grk_all}

    # динамика по годам заключения (к передаче)
    years = defaultdict(lambda: defaultdict(float))
    years_n = defaultdict(lambda: defaultdict(int))
    for r in grk_all:
        if r["date"]:
            years[str(r["date"].year)][law_key(r["law"])] += (r["price"] or 0) / 1e9
            years_n[str(r["date"].year)][law_key(r["law"])] += 1
    yser = sorted(years.items())
    ynser = sorted(years_n.items())

    # топ членов и заказчиков
    mem_sum = defaultdict(float)
    for r in grk_all:
        mem_sum[r["inn"]] += r["price"] or 0
    top_mem = sorted(mem_sum.items(), key=lambda x: -x[1])[:10]
    cust_sum = defaultdict(float)
    for r in grk_all:
        cust_sum[short(r["customer"]) or "Заказчик не указан"] += r["price"] or 0
    top_cust = sorted(cust_sum.items(), key=lambda x: -x[1])[:8]
    total = sum(mem_sum.values()) or 1
    top5_share = sum(v for _, v in top_mem[:5]) / total

    # причины проверки
    reasons = Counter()
    for r in check:
        n = r["note"]
        if "заказчик" in n:
            reasons["нет наименования заказчика"] += 1
        elif "не строительный" in n:
            reasons["по предмету не строительный подряд"] += 1
        elif "не уверен" in n:
            reasons["не ясно, работы ли по ГрК"] += 1
        elif r["status"].startswith("Вероят"):
            reasons["победа по протоколу, договора нет"] += 1
        elif "различ" in n:
            reasons["расхождение источников"] += 1
        else:
            reasons[r["status"] or "прочее"] += 1

    # новые договоры: появились в хранилище за 7 дней (по дате заключения)
    today = date.today()
    dues, dues_file = read_dues(today)
    reg, reg_when = read_registry()
    for d in dues:
        d.setdefault("src", "data\\vznosy.xlsx")
    dues = registry_items(reg, confirmed, today) + dues
    dues.sort(key=lambda d: ({"просрочен": 0, "срочно": 1, "приближается": 2}.get(d["state"], 3), d["due"] or date(1900, 1, 1)))
    dues_over = [d for d in dues if d["state"] == "просрочен"]
    dues_soon = [d for d in dues if d["state"] in ("срочно", "приближается")]
    dues_urgent = [d for d in dues if d["state"] == "срочно"]
    recent = sorted([r for r in confirmed if r["date"] and (today - r["date"]).days <= 30], key=lambda r: r["date"], reverse=True)

    # история по дням
    hist_p = OUT / "dash_history.json"
    try:
        hist = json.loads(hist_p.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        hist = []
    snap = {"date": today.isoformat(), "transfer": len(confirmed), "transfer_sum": round(to_transfer_sum),
            "check": len(check), "store": len(store), "members_with": len(members_with),
            "dues_over": len(dues_over), "dues_soon": len(dues_soon)}
    hist = [h for h in hist if h.get("date") != snap["date"]] + [snap]
    hist = hist[-120:]
    hist_p.write_text(json.dumps(hist, ensure_ascii=False, indent=1), encoding="utf-8")
    prev = hist[-2] if len(hist) > 1 else None

    def delta(key):
        if not prev:
            return ""
        d = snap[key] - prev.get(key, 0)
        return f"{'+' if d >= 0 else '−'}{num(abs(d))} с {datetime.fromisoformat(prev['date']):%d.%m}"

    keys = [("44-ФЗ", "c44"), ("223-ФЗ", "c223"), ("ПП-615", "c615")]
    hl = health(reg_when)
    status_line = (f"Сверка от {esc(stamp)} · членов СРО: {len(members) or 123} · "
                   f"робот {esc(hl['version'])}: {('идёт ' + esc(hl['busy'])) if hl['busy'] else 'свободен'}, лог {esc(rs['file'])} от {esc(rs['when'])}")

    # ---------- взносы ----------
    def dname(d):
        return short(d["member"] or members.get(d["inn"], d["inn"]))

    def rub(v):
        return f"{num(round(v))} ₽" if v else "—"

    def dpill(d):
        cls = {"просрочен": "bad", "срочно": "bad", "приближается": "warn"}.get(d["state"], "")
        txt = d.get("pill") or (f"просрочен на {-d['days']} дн." if d["days"] < 0 else "сегодня" if d["days"] == 0 else f"через {d['days']} дн.")
        return f'<span class="pill {cls}">{txt}</span>'

    if not reg:
        dues_note = ("Сведений из реестра НОСТРОЙ ещё нет — запустите «12 Обновить дашборды» "
                     "(при первом запуске робот скачает их, около 10 минут).")
    else:
        dues_note = ""
    reg_line = (f"Реестр НОСТРОЙ — сведения о {len(reg)} членах на {reg_when}" if reg else "Реестр НОСТРОЙ — сведений пока нет") + \
        ("; членские взносы — data\\vznosy.xlsx" if dues_file else "")
    dues_kpi = (f'<div class="card kpi {"bad" if dues_over or dues_urgent else "warn" if dues_soon else "ok"}">'
                f'<div class="l">Взносы и сроки: требуют внимания</div><div class="v">{num(len(dues_soon) + len(dues_over))}</div>'
                f'<div class="d">просрочено — {num(len(dues_over))}; срочно — {num(len(dues_urgent))}'
                f'{" · " + rub(sum(d["sum"] or 0 for d in dues_soon + dues_over)) if any(d["sum"] for d in dues_soon + dues_over) else ""}</div></div>')
    near = dues_over + dues_soon
    dues_list = ("<div style=\"overflow-x:auto\"><table style=\"min-width:560px\"><thead><tr><th>Член СРО</th><th>Взнос</th><th>Срок</th><th class=\"n\">Сумма</th><th></th></tr></thead><tbody>"
                 + "".join(f'<tr><td>{esc(dname(d))}</td><td>{esc(d["kind"])}{(" · " + esc(d["period"])) if d["period"] else ""}</td>'
                           f'<td class="nw">{d["due"].strftime("%d.%m.%Y") if d["due"] else "—"}</td><td class="n">{rub(d["sum"])}</td><td class="nw">{dpill(d)}</td></tr>'
                           for d in near[:12])
                 + "</tbody></table></div>") if near else f'<p class="muted">{esc(dues_note) or f"Взносов со сроком в ближайшие {DUES_WARN_DAYS} дней нет."}</p>'
    dues_card = (f'<div class="card"><h2>Взносы, страхование и право на работы: что требует внимания</h2>{dues_list}'
                 f'{f"<p class=small muted>Ещё {len(near) - 12} — на рабочем столе сотрудника.</p>" if len(near) > 12 else ""}'
                 f'<p class="small muted">{esc(reg_line)}. Показаны: доплата в КФ ОДО, когда обязательства члена дошли до {ODO_NEAR:.0%} предела уровня; договоры свыше 10 млн ₽ без права ОДО; страхование, истекающее в ближайшие {DUES_WARN_DAYS} дней; неоплаченные членские взносы.</p></div>')

    # ---------- руководитель ----------
    kpis = f"""<div class="grid k4">
<div class="card kpi ok"><div class="l">Новых договоров к передаче в НОСТРОЙ</div><div class="v">{num(len(confirmed))}</div><div class="d">{money(to_transfer_sum, 'млрд')}{(' · ' + delta('transfer')) if prev else ''} · уже передано {num(len(sent))}{(', исключено оператором ' + num(len(excluded))) if excluded else ''}</div></div>
<div class="card kpi warn"><div class="l">Требуют решения оператора</div><div class="v">{num(len(check))}</div><div class="d">{delta('check') or 'спорные случаи и пробелы в данных'}</div></div>
<div class="card kpi"><div class="l">Членов СРО с договорами под ГрК</div><div class="v">{num(len(members_with))}</div><div class="d">из {len(members) or 123}; без договоров — {max(0, (len(members) or 123) - len(members_with))}</div></div>
<div class="card kpi"><div class="l">Без обязательного членства в СРО</div><div class="v">{num(no_sro)}</div><div class="d">договоры до 10 млн ₽ (ч. 2.1 ст. 52 ГрК) — передаются в НОСТРОЙ, но СРО для них не обязательна</div></div>
{dues_kpi}
</div>"""
    law_cards = "".join(
        f'<div class="card kpi"><div class="l"><span class="pill">{k}</span> к передаче</div><div class="v">{num(by_law.get(k, 0))}</div><div class="d">{money(sum_by_law.get(k, 0), "млрд")}</div></div>'
        for k, _ in keys)
    trend = ""
    if len(hist) >= 2:
        trend = '<div class="card"><h2>Динамика по дням</h2>' + columns(
            [(datetime.fromisoformat(h["date"]).strftime("%d.%m"), {"к передаче": h["transfer"], "требуют решения": h["check"]}) for h in hist[-20:]],
            [("к передаче", "cok"), ("требуют решения", "cwarn")], height=160) + legend([("к передаче", "cok"), ("требуют решения", "cwarn")]) + "</div>"
    risks = []
    if dues_over:
        risks.append(f"Просрочены взносы у {len({d['inn'] for d in dues_over})} членов СРО ({len(dues_over)} начислений) — нужно напоминание или решение по членству.")
    if dues_urgent:
        risks.append(f"Срочно: {len(dues_urgent)} — превышен уровень ОДО, нет права ОДО при крупных договорах или страхование истекает в ближайшие {DUES_URGENT_DAYS} дней.")
    if rs["errors"]:
        risks.append(f"В последнем прогоне робота {rs['errors']} ошибок ЕИС — проверьте, догружены ли пропущенные дни.")
    if reasons.get("нет наименования заказчика"):
        risks.append(f"{reasons['нет наименования заказчика']} договоров без наименования заказчика — оператор дополняет по ИНН.")
    if top5_share > 0.6:
        risks.append(f"Пять крупнейших членов дают {top5_share:.0%} суммы договоров к передаче — при проверке НОСТРОЙ смотреть их в первую очередь.")
    if not by_law.get("ПП-615"):
        risks.append("Договоров ПП-615 в сверке пока нет — дождитесь окончания загрузки «10 Загрузка ПП-615 с 2017 года».")
    flags = version_flags()
    pkg_files = sorted(OUT.glob("nostroy_*.csv"), key=lambda p: p.stat().st_mtime)
    pkg = pkg_files[-1].name if pkg_files else "—"
    flags_sum = sum(to_float(f["price"]) or 0 for f in flags)
    flags_card = ""
    if flags:
        flag_rows = "".join(f'<tr><td class="nw">{esc(f["reg"])}</td><td>{esc(short(f["member"] or ""))}</td><td>{esc(f["verdict"])}</td>'
                            f'<td class="n">{num(round(to_float(f["price"]) or 0))} ₽</td><td>{esc(f["note"])}</td></tr>' for f in flags)
        flags_card = (f'<div class="card" style="margin-top:14px;border-left:4px solid var(--bad)"><h2>Перед загрузкой: устаревшие договоры в пакете</h2>'
                      f'<p>Проверка версий нашла в пакете <b>{esc(pkg)}</b> договоров с новой версией: {num(len(flags))} на {num(round(flags_sum))} ₽. '
                      f'Пакет не загружать, пока по ним не принято решение в data\\decisions.xlsx и не проведена новая сверка.</p>'
                      f'<div style="overflow-x:auto"><table style="min-width:640px"><thead><tr><th>Реестровый номер</th><th>Член СРО</th><th>Вывод</th><th class="n">Цена</th><th>Пояснение</th></tr></thead>'
                      f'<tbody>{flag_rows}</tbody></table></div></div>')
    load_step = (f'<li>Сначала принять решение по договорам из блока «Перед загрузкой», записать его в data\\decisions.xlsx и запустить сверку («11 Только сверка»). '
                 f'Пакет <b>{esc(pkg)}</b> не загружать: загрузить новый пакет после сверки и сразу отметить его загруженным.</li>') if flags else (
                 f'<li>Загрузить в НОСТРОЙ пакет <b>{esc(pkg)}</b> из папки data\\out (только новые договоры) и сразу отметить его загруженным: пульт → «5» (или «16 Отметить пакет загруженным»).</li>')
    lead_body = f"""{kpis}
{flags_card}
<div style="margin-top:14px">{dues_card}</div>
<div class="grid k4" style="margin-top:14px">{law_cards}</div>
<div class="grid k2" style="margin-top:14px">
<div class="card"><h2>Договоры к передаче по годам заключения, млрд ₽</h2>{columns(yser, keys, fmt=lambda v: f"{v:.1f}".replace(".", ","))}{legend(keys)}</div>
<div class="card"><h2>Договоры к передаче по годам, штук</h2>{columns(ynser, keys)}{legend(keys)}</div>
<div class="card"><h2>Крупнейшие члены СРО по сумме договоров под ГрК</h2>{bars([(short(members.get(i, i)), v, 'c44') for i, v in top_mem], fmt=lambda v: money(v, 'млрд'))}
<p class="small muted">Топ-5 — {top5_share:.0%} всей суммы.</p></div>
<div class="card"><h2>Крупнейшие заказчики</h2>{bars([(c, v, 'cacc') for c, v in top_cust], fmt=lambda v: money(v, 'млрд'))}</div>
</div>
<div class="grid k2" style="margin-top:14px">
<div class="card"><h2>Почему записи требуют решения</h2>{bars([(k, v, 'cwarn') for k, v in reasons.most_common()])}</div>
<div class="card"><h2>На что обратить внимание</h2>{'<ul>' + ''.join(f'<li>{esc(x)}</li>' for x in risks) + '</ul>' if risks else '<p class="muted">Замечаний нет.</p>'}
<h3>Последнее действие робота</h3><p class="small">{esc(rs['last']) or '—'}</p></div>
</div>
<div class="grid k2" style="margin-top:14px">{health_card(hl)}</div>{('<div style="margin-top:14px">' + trend + '</div>') if trend else ''}"""

    # ---------- сотрудник ----------
    def row(r, kind):
        cls = "ok" if r["status"] == "Подтверждён" else ("warn" if "провер" in r["status"].lower() else "bad")
        return (f'<tr data-k="{esc(kind)}"><td><span class="pill {cls}">{esc(r["status"])}</span></td><td class="nw">{esc(law_key(r["law"]))}</td>'
                f'<td>{esc(short(r["member"]))}</td><td>{eis_link(r)}</td><td class="nw" data-v="{r["date"].isoformat() if r["date"] else ""}">{r["date"].strftime("%d.%m.%Y") if r["date"] else ""}</td>'
                f'<td class="n" data-v="{r["price"] or 0}">{money(r["price"])}</td><td>{esc(short(r["customer"])[:70])}</td><td class="sub" title="{esc(r["subject"])}">{esc(r["subject"][:110] + ("…" if len(r["subject"]) > 110 else ""))}</td>'
                f'<td class="note small">{esc(r["note"][:140])}</td></tr>')

    def kind_of(r):
        n = r["note"]
        if "заказчик" in n:
            return "заказчик"
        if "не строительный" in n:
            return "не стройка"
        if "не уверен" in n:
            return "ГрК?"
        if r["status"].startswith("Вероят"):
            return "нет договора"
        return "прочее"

    qrows = "".join(row(r, kind_of(r)) for r in sorted(check, key=lambda r: -(r["price"] or 0)))
    trows = "".join(row(r, law_key(r["law"])) for r in sorted(confirmed, key=lambda r: (r["date"] or date(1900, 1, 1)), reverse=True))
    head = ('<thead><tr><th>Статус</th><th>Закон</th><th>Член СРО</th><th>Договор в ЕИС</th><th>Дата</th><th class="n">Цена</th>'
            '<th>Заказчик</th><th>Предмет</th><th>Замечание</th></tr></thead>')
    kinds = Counter(kind_of(r) for r in check)
    opts = "".join(f'<option value="{esc(k)}">{esc(k)} ({v})</option>' for k, v in kinds.most_common())
    lopts = "".join(f'<option value="{k}">{k} ({by_law.get(k, 0)})</option>' for k, _ in keys)
    files = sorted(OUT.glob("nostroy_*.csv"), key=lambda p: p.stat().st_mtime)
    pkg = files[-1].name if files else "—"
    drows = "".join(
        f'<tr data-k="{esc(d["state"])}"><td class="nw">{dpill(d)}</td><td>{esc(dname(d))}</td><td class="nw">{esc(d["inn"])}</td>'
        f'<td>{esc(d["kind"])}</td><td class="nw">{esc(d["period"])}</td><td class="nw" data-v="{d["due"].isoformat() if d["due"] else ""}">{d["due"].strftime("%d.%m.%Y") if d["due"] else "—"}</td>'
        f'<td class="n" data-v="{d["sum"] or 0}">{rub(d["sum"])}</td><td class="note small">{esc(d["note"])}</td></tr>'
        for d in near)
    dkinds = Counter(d["state"] for d in near)
    dopts = "".join(f'<option value="{k}">{k} ({dkinds[k]})</option>' for k in ("просрочен", "срочно", "приближается") if dkinds.get(k))
    dues_emp = (f'<div class="card" style="margin-top:14px"><h2>Взносы, страхование и право на работы</h2>'
                + (f'<div class="tools"><input type="search" id="d-q" placeholder="Поиск: член СРО, ИНН, вид взноса" oninput="filt(\'d\')">'
                   f'<select id="d-s" onchange="filt(\'d\')"><option value="">Все</option>{dopts}</select>'
                   f'<span class="small muted">Показано: <b id="d-n">{len(near)}</b></span></div>'
                   f'<div class="scroll" style="max-height:360px"><table id="d" style="min-width:900px"><thead><tr><th>Срок</th><th>Член СРО</th><th>ИНН</th><th>Вид взноса</th>'
                   f'<th>Период</th><th>Срок</th><th class="n">Сумма</th><th>Пояснение</th></tr></thead><tbody>{drows}</tbody></table></div>'
                   if near else f'<p class="muted">{esc(dues_note) or f"Взносов со сроком в ближайшие {DUES_WARN_DAYS} дней нет."}</p>')
                + f'<p class="small muted">{esc(reg_line)}. Сведения реестра обновляются раз в сутки при обновлении дашбордов; после оплаты членского взноса впишите дату в «Дата оплаты» файла data\\vznosy.xlsx.</p></div>')
    emp_body = f"""<div class="grid k4">
<div class="card kpi warn"><div class="l">В очереди на решение</div><div class="v">{num(len(check))}</div><div class="d">{delta('check')}</div></div>
<div class="card kpi ok"><div class="l">Новые к передаче</div><div class="v">{num(len(confirmed))}</div><div class="d">пакет {esc(pkg)} · уже передано {num(len(sent))}</div></div>
<div class="card kpi"><div class="l">Заключены за 30 дней</div><div class="v">{num(len(recent))}</div><div class="d">{money(sum(r['price'] or 0 for r in recent))}</div></div>
<div class="card kpi {'bad' if rs['errors'] else ''}"><div class="l">Ошибок ЕИС в последнем прогоне</div><div class="v">{num(rs['errors'])}</div><div class="d">{esc(rs['file'])} · {esc(rs['when'])}</div></div>
{dues_kpi}
</div>
{flags_card}
<div class="card" style="margin-top:14px"><h2>Порядок работы на сегодня</h2><ol class="steps">
<li>Разобрать очередь «Требуют решения» ниже: открыть договор в ЕИС, по смете и техническому заданию решить — капитальный это ремонт (ГрК) или текущий.</li>
<li>Для записей «заказчик» — найти наименование по ИНН заказчика.</li>
<li>Записать решения по очереди в файл <b>data\\decisions.xlsx</b> (колонка «Решение») — при следующей сверке они применятся, и эти договоры больше не вернутся в очередь.</li>
{load_step}
<li>Если в логе есть ошибки ЕИС — запустить догрузку пропущенных дней.</li>
<li>Напомнить членам СРО о взносах со сроком в ближайшие {DUES_URGENT_DAYS} дней и о просроченных (список ниже); отметить поступившие оплаты в data\\vznosy.xlsx.</li></ol>
<p class="small muted">Цена не основание исключать договор: до 10 млн ₽ договор передаётся, но членство в СРО для него не обязательно; для ПП-615 обязательно всегда.</p></div>
{dues_emp}
<div class="grid k2" style="margin-top:14px">{health_card(hl)}</div>
<div class="card" style="margin-top:14px"><h2>Очередь «Требуют решения»</h2>
<div class="tools"><input type="search" id="q-q" placeholder="Поиск: член СРО, заказчик, номер, предмет" oninput="filt('q')">
<select id="q-s" onchange="filt('q')"><option value="">Все причины</option>{opts}</select><span class="small muted">Показано: <b id="q-n">{len(check)}</b></span></div>
<div class="scroll"><table id="q">{head}<tbody>{qrows}</tbody></table></div></div>
<div class="card" style="margin-top:14px"><h2>К передаче в НОСТРОЙ</h2>
<div class="tools"><input type="search" id="t-q" placeholder="Поиск" oninput="filt('t')">
<select id="t-s" onchange="filt('t')"><option value="">Все законы</option>{lopts}</select><span class="small muted">Показано: <b id="t-n">{len(confirmed)}</b></span></div>
<div class="scroll"><table id="t">{head}<tbody>{trows}</tbody></table></div></div>"""

    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "Flowlink — дашборд руководителя.html").write_text(
        page("Дашборд руководителя СРО", status_line, lead_body), encoding="utf-8")
    (OUT / "Flowlink — дашборд сотрудника.html").write_text(
        page("Рабочий стол ответственного сотрудника", status_line, emp_body), encoding="utf-8")
    print(f"Дашборды обновлены: к передаче {len(confirmed)}, требуют решения {len(check)}. Папка data\\out")
    return 0


if __name__ == "__main__":
    raise SystemExit(build())
